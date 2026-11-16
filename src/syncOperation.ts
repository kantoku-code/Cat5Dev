import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { randomUUID } from 'crypto';
import { execFile } from 'child_process';
import { CatiaExpectation, stripBom, validateCatiaResult } from './catiaTypes';

type Phase = 'prepared' | 'external-running' | 'settled';
interface Owner {
    version: 1; id: string; pid: number; kind: string; createdAt: string;
    phase: Phase; tempDir: string; resultPath?: string; expected?: CatiaExpectation;
}
export interface OperationContext {
    readonly id: string; readonly tempDir: string;
    markExternalRunning(resultPath: string, expected: CatiaExpectation): void;
    markSettled(): void;
    close(): void;
    isClosed(): boolean;
}
export interface SyncOperationManagerOptions {
    lockRoot?: string; tempRoot?: string;
    isProcessAlive?: (pid: number) => boolean;
    isCatiaRunning?: () => Promise<boolean>;
}
export class SyncBusyError extends Error { constructor() { super('CATIA operation is busy or unsettled'); } }
const uuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
function code(e: unknown): string | undefined { return (e as NodeJS.ErrnoException)?.code; }
function alive(pid: number): boolean {
    try { process.kill(pid, 0); return true; } catch (e) { return code(e) !== 'ESRCH'; }
}
function catiaRunning(): Promise<boolean> {
    return new Promise((resolve, reject) => {
        execFile('tasklist.exe', ['/FO', 'CSV', '/NH'], {
            windowsHide: true, timeout: 10000, maxBuffer: 10 * 1024 * 1024, encoding: 'utf8'
        }, (error, stdout) => {
            if (error) { reject(error); return; }
            const lines = stripBom(stdout).split(/\r?\n/).map(x => x.trim()).filter(Boolean);
            if (!lines.length || lines.some(x => !/^"[^"]+","\d+"(?:,|$)/.test(x))) {
                reject(new Error('Cannot verify Windows process list')); return;
            }
            resolve(lines.some(x => /^"CNEXT\.exe",/i.test(x)));
        });
    });
}
function within(parent: string, candidate: string): boolean {
    const rel = path.relative(parent, candidate);
    return !!rel && rel !== '..' && !rel.startsWith(`..${path.sep}`) && !path.isAbsolute(rel);
}
function expectation(v: any): v is CatiaExpectation {
    return !!v && ['operationId', 'stepId', 'kind', 'targetProject'].every(k => typeof v[k] === 'string')
        && !!v.operationId && !!v.stepId && ['list', 'pull', 'check', 'push'].includes(v.kind);
}

export class SyncOperationManager {
    private readonly lockRoot: string;
    private readonly tempRoot: string;
    private readonly activePath: string;
    private readonly isProcessAlive: (pid: number) => boolean;
    private readonly isCatiaRunning: () => Promise<boolean>;
    /** This process's operation whose run ended while CATIA completion was still unconfirmed. */
    private quarantined?: { id: string; abandon: () => void };
    constructor(options: SyncOperationManagerOptions = {}) {
        const local = process.env.LOCALAPPDATA ?? path.join(os.homedir(), 'AppData', 'Local');
        const root = path.resolve(options.lockRoot ?? path.join(local, 'Cat5Dev', 'operation-lock'));
        const temp = path.resolve(options.tempRoot ?? os.tmpdir());
        fs.mkdirSync(root, { recursive: true }); fs.mkdirSync(temp, { recursive: true });
        this.lockRoot = fs.realpathSync(root); this.tempRoot = fs.realpathSync(temp);
        this.activePath = path.join(this.lockRoot, 'active');
        this.isProcessAlive = options.isProcessAlive ?? alive;
        this.isCatiaRunning = options.isCatiaRunning ?? catiaRunning;
    }
    acquire(kind: string): OperationContext {
        if (fs.existsSync(this.activePath)) { throw new SyncBusyError(); }
        const id = randomUUID();
        const candidate = path.join(this.lockRoot, `.candidate-${id}`);
        const tempDir = fs.mkdtempSync(path.join(this.tempRoot, `cat5dev-${id}-`));
        let owner: Owner = { version: 1, id, pid: process.pid, kind,
            createdAt: new Date().toISOString(), phase: 'prepared', tempDir };
        let acquired = false;
        try {
            fs.mkdirSync(candidate);
            fs.writeFileSync(path.join(candidate, 'owner.json'), JSON.stringify(owner), { flag: 'wx' });
            try { fs.renameSync(candidate, this.activePath); acquired = true; }
            catch (e) { if (fs.existsSync(this.activePath)) { throw new SyncBusyError(); } throw e; }
        } finally {
            if (!acquired) {
                fs.rmSync(candidate, { recursive: true, force: true }); this.removeTemp(owner);
            }
        }
        let closed = false;
        return {
            id, tempDir,
            markExternalRunning: (resultPath, expected) => {
                if (closed || !expectation(expected) || expected.operationId !== id
                    || !within(tempDir, path.resolve(resultPath))) { throw new Error('Invalid operation context'); }
                const next: Owner = { ...owner, phase: 'external-running', resultPath: path.resolve(resultPath), expected };
                this.replaceOwner(next); owner = next;
            },
            markSettled: () => {
                if (closed) { throw new Error('Operation already closed'); }
                const next: Owner = { ...owner, phase: 'settled' }; this.replaceOwner(next); owner = next;
            },
            close: () => {
                if (closed) { return; }
                if (owner.phase === 'external-running') {
                    // Keep the lock, but let this window release it explicitly once the user confirms CATIA is idle.
                    this.quarantined = { id, abandon: () => { closed = true; } };
                    return;
                }
                let cleanupError: unknown;
                try { this.removeTemp(owner); } catch (e) { cleanupError = e; }
                this.assertOwner(owner.id);
                const released = path.join(this.lockRoot, `.released-${owner.id}`);
                fs.renameSync(this.activePath, released);
                closed = true;
                if (this.quarantined?.id === id) { this.quarantined = undefined; }
                fs.rmSync(released, { recursive: true, force: true });
                if (cleanupError) { throw cleanupError; }
            },
            isClosed: () => closed
        };
    }
    /** True when the active lock is an unconfirmed operation left behind by this manager (not one still running). */
    hasQuarantine(): boolean {
        const quarantined = this.quarantined;
        if (!quarantined) { return false; }
        try {
            const owner = this.readOwner();
            if (owner && owner.id === quarantined.id && owner.pid === process.pid) { return true; }
        } catch { return false; }
        this.quarantined = undefined; // The lock was released or recovered elsewhere.
        return false;
    }
    /**
     * Force-release this manager's quarantined lock after user confirmation.
     * The temporary directory is kept for diagnostics, as with recover().
     */
    releaseQuarantined(): boolean {
        const quarantined = this.quarantined;
        if (!quarantined || !this.hasQuarantine()) { return false; }
        fs.renameSync(this.activePath, path.join(this.lockRoot, `.recovered-${quarantined.id}`));
        quarantined.abandon();
        this.quarantined = undefined;
        return true;
    }
    async recover(): Promise<boolean> {
        let owner: Owner | undefined;
        try { owner = this.readOwner(); } catch { return false; }
        if (!owner) { return true; }
        try { if (this.isProcessAlive(owner.pid)) { return false; } } catch { return false; }
        if (owner.phase === 'external-running' && !this.hasTerminalResult(owner)) {
            try { if (await this.isCatiaRunning()) { return false; } } catch { return false; }
        }
        // A fixed, nonempty tombstone prevents a delayed reclaimer from
        // renaming a newer owner's lock after another reclaimer succeeded.
        const retired = path.join(this.lockRoot, `.recovered-${owner.id}`);
        try { fs.renameSync(this.activePath, retired); return true; }
        catch { return !fs.existsSync(this.activePath); }
    }
    private readOwner(): Owner | undefined {
        let text: string;
        try { text = fs.readFileSync(path.join(this.activePath, 'owner.json'), 'utf8'); }
        catch (e) { if (code(e) === 'ENOENT' && !fs.existsSync(this.activePath)) { return undefined; } throw e; }
        const owner = JSON.parse(text) as Owner;
        if (!owner || owner.version !== 1 || typeof owner.id !== 'string' || !uuid.test(owner.id)
            || !Number.isSafeInteger(owner.pid) || owner.pid <= 0 || typeof owner.kind !== 'string'
            || !['prepared', 'external-running', 'settled'].includes(owner.phase)
            || typeof owner.tempDir !== 'string' || !this.managedTemp(owner)) { throw new Error('Invalid lock metadata'); }
        if (owner.phase === 'external-running' && (!expectation(owner.expected)
            || owner.expected.operationId !== owner.id || typeof owner.resultPath !== 'string'
            || !within(owner.tempDir, owner.resultPath))) { throw new Error('Invalid execution metadata'); }
        return owner;
    }
    private assertOwner(id: string): void {
        const current = this.readOwner();
        if (!current || current.id !== id || current.pid !== process.pid) { throw new Error('Lock ownership lost'); }
    }
    private replaceOwner(owner: Owner): void {
        this.assertOwner(owner.id);
        const temp = path.join(this.activePath, `.owner-${randomUUID()}.tmp`);
        try {
            fs.writeFileSync(temp, JSON.stringify(owner), { flag: 'wx' });
            fs.renameSync(temp, path.join(this.activePath, 'owner.json'));
        } finally { fs.rmSync(temp, { force: true }); }
    }
    private hasTerminalResult(owner: Owner): boolean {
        try {
            if (!owner.resultPath || !owner.expected) { return false; }
            const stat = fs.lstatSync(owner.resultPath);
            if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 16 * 1024 * 1024) { return false; }
            validateCatiaResult(fs.readFileSync(owner.resultPath, 'utf8'), owner.expected); return true;
        } catch { return false; }
    }
    private managedTemp(owner: Pick<Owner, 'id' | 'tempDir'>): boolean {
        return uuid.test(owner.id) && path.resolve(owner.tempDir) === owner.tempDir
            && path.dirname(owner.tempDir) === this.tempRoot
            && path.basename(owner.tempDir).startsWith(`cat5dev-${owner.id}-`);
    }
    private removeTemp(owner: Owner): void {
        if (!this.managedTemp(owner)) { throw new Error('Refusing to remove unowned temporary path'); }
        fs.rmSync(owner.tempDir, { recursive: true, force: true });
    }
}
