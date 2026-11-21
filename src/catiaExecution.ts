import * as fs from 'fs';
import * as path from 'path';
import { randomUUID } from 'crypto';
import { execFile } from 'child_process';
import { OperationContext } from './syncOperation';
import { CatiaRequest, CatiaResult, CatiaStep, catiaSucceeded, validateCatiaResult } from './catiaTypes';
import { buildCatiaScript } from './catiaScripts';

export interface ExecutionOutput { error?: (Error & { code?: string | number | null }) | null; stdout: string; stderr: string }
/** Script that CATIA runs from the request directory (switch here to operation.CATScript). */
export const CATIA_SCRIPT_NAME = 'operation.catvbs';
export type CatiaExecutor = (directory: string, scriptName: string, timeoutMs: number) => Promise<ExecutionOutput>;
/** bridgePath: bin/catia-run.exe, used unless a custom executor is given. */
export interface ExecutionOptions { execute?: CatiaExecutor; bridgePath?: string; timeoutMs?: number; log?: (text: string) => void }
export class CatiaExecutionError extends Error {
    constructor(message: string, public readonly result?: CatiaResult, public readonly uncertain = false) { super(message); }
}
export function createCatiaRequest(context: OperationContext, kind: CatiaStep, targetProject: string, encoding: string): CatiaRequest {
    const stepId = randomUUID();
    const directory = path.join(context.tempDir, stepId);
    fs.mkdirSync(directory);
    return { operationId: context.id, stepId, kind, targetProject, encoding, directory };
}
/** catia-run.exe calls SystemService.ExecuteScript. Exit 10/11: CATIA/SystemService unavailable, 12: ExecuteScript failed. */
export function bridgeExecutor(bridgePath: string): CatiaExecutor {
    return (directory, scriptName, timeoutMs) => new Promise(resolve => {
        execFile(bridgePath, [directory, scriptName], {
            windowsHide: true, timeout: timeoutMs, maxBuffer: 10 * 1024 * 1024, encoding: 'utf8'
        }, (error, stdout, stderr) => resolve({ error, stdout, stderr }));
    });
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
    const execute = options.execute ?? (options.bridgePath ? bridgeExecutor(options.bridgePath) : undefined);
    if (!execute) { throw new CatiaExecutionError('CATIA bridge path is not configured'); }
    const resultPath = path.join(request.directory, 'result.json');
    fs.writeFileSync(path.join(request.directory, CATIA_SCRIPT_NAME), buildCatiaScript(request), 'utf8');
    const timeoutMs = options.timeoutMs ?? 60000;
    const deadline = Date.now() + timeoutMs;
    context.markExternalRunning(resultPath, request);
    let timeout: NodeJS.Timeout | undefined;
    let output: ExecutionOutput;
    try {
        output = await Promise.race([
            execute(request.directory, CATIA_SCRIPT_NAME, timeoutMs),
            new Promise<never>((_, reject) => { timeout = setTimeout(() => reject(new Error('CATIA execution timeout')), timeoutMs); })
        ]);
    } catch (e) {
        watchLateResult(context, request, resultPath, options.log);
        throw new CatiaExecutionError(String(e), undefined, true);
    } finally { if (timeout) { clearTimeout(timeout); } }
    options.log?.(`[CATIA ${request.kind}] ${output.stdout}\n${output.stderr}`);
    // Only errors known to precede dispatch prove that CATIA never started.
    const exitCode = (output.error as any)?.code;
    const preDispatch = [10, 11, 'ENOENT', 'EACCES', 'EPERM'].includes(exitCode);
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
        throw new CatiaExecutionError(`CATIA completion not confirmed: ${output.error?.message ?? String(readError)}`, undefined, !preDispatch);
    }
    options.log?.(`[CATIA result] ${JSON.stringify(result)}`);
    if (output.error || !catiaSucceeded(result)) {
        throw new CatiaExecutionError(output.error?.message ?? 'CATIA operation partially failed', result);
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
