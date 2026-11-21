import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import { execFile } from 'child_process';
import { OperationContext } from './syncOperation';
import { CatiaRequest, CatiaResult, CatiaStep, catiaSucceeded, safeFileName, validateCatiaResult } from './catiaTypes';

export interface ExecutionOutput { error?: (Error & { code?: string | number | null }) | null; stdout: string; stderr: string }
export const REQUEST_FILE_NAME = 'request.json';
export const RESULT_FILE_NAME = 'result.json';
/** Absolute paths passed to catia-run.exe (docs/dev/bridge-contract.md 3). request and result share the request directory. */
export interface CatiaInvocation { bridge: string; request: string; result: string }
export type CatiaExecutor = (invocation: CatiaInvocation, timeoutMs: number) => Promise<ExecutionOutput>;
/**
 * runnerPath: bin/catia-run.exe, used unless a custom executor is given.
 * bridgeMaster: the read-only master prepared by BridgeFiles (--bridge).
 */
export interface ExecutionOptions { execute?: CatiaExecutor; runnerPath?: string; bridgeMaster?: string; timeoutMs?: number; log?: (text: string) => void }
/** catia-run.exe exit codes (bridge-contract 6). */
export const EXIT = { invalidArguments: 2, catiaNotRunning: 10, noSystemService: 11, executeScriptFailed: 12, bridgeUnavailable: 13, vbaEditorClosed: 14 } as const;
/** Failures known to happen before anything is written in CATIA: the lock is released at once. */
const PRE_DISPATCH: (string | number)[] = [2, 10, 11, 13, 14, 'ENOENT', 'EACCES', 'EPERM'];
/** request.json is limited to 64 MiB by catia-run.exe. */
const MAX_REQUEST_BYTES = 64 * 1024 * 1024;
export class CatiaExecutionError extends Error {
    constructor(message: string, public readonly result?: CatiaResult, public readonly uncertain = false,
        /** Exit code of catia-run.exe (or the spawn error code) when known. */
        public readonly exitCode?: number | string) { super(message); }
}
export function createCatiaRequest(context: OperationContext, kind: CatiaStep, targetProject: string): CatiaRequest {
    const stepId = randomUUID();
    const directory = path.join(context.tempDir, stepId);
    fs.mkdirSync(directory);
    return { operationId: context.id, stepId, kind, targetProject, directory };
}
/** catia-run.exe --bridge <master> --request <dir>\request.json --result <dir>\result.json */
export function bridgeExecutor(runnerPath: string): CatiaExecutor {
    return (invocation, timeoutMs) => new Promise(resolve => {
        execFile(runnerPath, ['--bridge', invocation.bridge, '--request', invocation.request, '--result', invocation.result], {
            windowsHide: true, timeout: timeoutMs, maxBuffer: 10 * 1024 * 1024, encoding: 'utf8'
        }, (error, stdout, stderr) => resolve({ error, stdout, stderr }));
    });
}
/** request.json (bridge-contract 4). Throws before dispatch when the request breaks the contract. */
export function buildRequestJson(request: CatiaRequest): string {
    const push = request.kind === 'push';
    if ((request.kind === 'list') !== (request.targetProject === '')) { throw new Error('Invalid target project'); }
    if (!push && ((request.modules?.length ?? 0) || (request.deletions?.length ?? 0))) { throw new Error('Only push transfers modules'); }
    const names = new Set<string>();
    const unique = (name: string): void => {
        if (typeof name !== 'string' || !safeFileName(name) || names.has(name.toLowerCase())) { throw new Error(`Invalid transfer item: ${name}`); }
        names.add(name.toLowerCase());
    };
    for (const item of request.modules ?? []) {
        unique(item.name);
        if (![1, 2, 3].includes(item.type) || typeof item.text !== 'string') { throw new Error(`Invalid transfer item: ${item.name}`); }
    }
    for (const name of request.deletions ?? []) { unique(name); }
    const text = JSON.stringify({
        version: 2, operationId: request.operationId, stepId: request.stepId, kind: request.kind, targetProject: request.targetProject,
        modules: push ? (request.modules ?? []).map(({ name, type, text }) => ({ name, type, text })) : [],
        deletions: push ? [...(request.deletions ?? [])] : []
    });
    if (Buffer.byteLength(text, 'utf8') > MAX_REQUEST_BYTES) { throw new Error('Request exceeds 64 MiB'); }
    return text;
}
function readResult(resultPath: string, request: CatiaRequest): CatiaResult {
    const stat = fs.lstatSync(resultPath);
    if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 16 * 1024 * 1024) { throw new Error('Invalid result file'); }
    return validateCatiaResult(fs.readFileSync(resultPath, 'utf8'), request);
}
/** A late, validated terminal result can release a quarantined operation. */
function watchLateResult(context: OperationContext, request: CatiaRequest, resultPath: string, log?: (text: string) => void): void {
    const timer = setInterval(() => {
        // The lock was released explicitly; a late result must not touch a newer operation.
        if (context.isClosed()) { clearInterval(timer); return; }
        try {
            readResult(resultPath, request);
            clearInterval(timer);
            context.markSettled(); context.close();
            log?.(`[CATIA] Late terminal result received: ${request.operationId}`);
        } catch { /* Keep the lock until terminal evidence or explicit recovery. */ }
    }, 1000);
    timer.unref();
}
export async function runCatiaRequest(context: OperationContext, request: CatiaRequest, options: ExecutionOptions = {}): Promise<CatiaResult> {
    const relative = path.relative(context.tempDir, request.directory);
    if (!relative || relative.startsWith('..') || path.isAbsolute(relative) || request.operationId !== context.id) {
        throw new Error('Request does not belong to this operation');
    }
    const execute = options.execute ?? (options.runnerPath ? bridgeExecutor(options.runnerPath) : undefined);
    if (!execute) { throw new CatiaExecutionError('CATIA runner path is not configured'); }
    if (!options.bridgeMaster || !path.isAbsolute(options.bridgeMaster)) { throw new CatiaExecutionError('CATIA bridge is not prepared'); }
    const requestPath = path.join(request.directory, REQUEST_FILE_NAME);
    const resultPath = path.join(request.directory, RESULT_FILE_NAME);
    fs.writeFileSync(requestPath, buildRequestJson(request), { encoding: 'utf8', flag: 'wx' });
    const timeoutMs = options.timeoutMs ?? 60000;
    const deadline = Date.now() + timeoutMs;
    context.markExternalRunning(resultPath, request);
    let timeout: NodeJS.Timeout | undefined;
    let output: ExecutionOutput;
    try {
        output = await Promise.race([
            execute({ bridge: options.bridgeMaster, request: requestPath, result: resultPath }, timeoutMs),
            new Promise<never>((_, reject) => { timeout = setTimeout(() => reject(new Error('CATIA execution timeout')), timeoutMs); })
        ]);
    } catch (e) {
        watchLateResult(context, request, resultPath, options.log);
        throw new CatiaExecutionError(String(e), undefined, true);
    } finally { if (timeout) { clearTimeout(timeout); } }
    const exitCode = output.error ? (output.error.code ?? undefined) : 0;
    options.log?.(`[CATIA ${request.kind}] exit=${exitCode ?? output.error?.message}\n${output.stdout}\n${output.stderr}`);
    if (exitCode === EXIT.executeScriptFailed) { options.log?.(`[CATIA ${request.kind}] ExecuteScript failed (exit 12): ${output.stderr.trim()}`); }
    // Only errors known to precede dispatch prove that nothing was written in CATIA.
    const preDispatch = exitCode !== undefined && PRE_DISPATCH.includes(exitCode);
    let result: CatiaResult | undefined;
    let readError: unknown;
    do {
        try { result = readResult(resultPath, request); break; } catch (e) { readError = e; }
        if (output.error || fs.existsSync(resultPath) || Date.now() >= deadline) { break; }
        await new Promise(resolve => setTimeout(resolve, Math.min(100, Math.max(1, deadline - Date.now()))));
    } while (true);
    if (result || preDispatch) { context.markSettled(); }
    if (!result) {
        if (!preDispatch) { watchLateResult(context, request, resultPath, options.log); }
        throw new CatiaExecutionError(`CATIA completion not confirmed: ${output.error?.message ?? String(readError)}`, undefined, !preDispatch, exitCode);
    }
    options.log?.(`[CATIA result] ${JSON.stringify(result)}`);
    if (output.error || !catiaSucceeded(result)) {
        throw new CatiaExecutionError(output.error?.message ?? 'CATIA operation partially failed', result, false, exitCode);
    }
    if (request.kind === 'push' && result.expectedCount !== (request.modules?.length ?? 0) + (request.deletions?.length ?? 0)) {
        throw new CatiaExecutionError('CATIA result count differs from request', result);
    }
    for (const item of result.items) {
        const stat = fs.lstatSync(path.join(request.directory, item.file));
        if (!stat.isFile() || stat.isSymbolicLink()) { throw new CatiaExecutionError('Missing exported module', result); }
    }
    return result;
}
