import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { createHash } from 'crypto';

/**
 * The bridge master (docs/dev/bridge-contract.md 2.1) and the cleanup of old links and copies (2.3).
 * Call prepare() only while holding the SyncOperationManager lock, before the first catia-run.exe of an operation.
 */

/** Master: C5D_Bridge-<sha8>.catvba. CATIA never loads this path; catia-run.exe loads a fresh link to it. */
export const MASTER_NAME = /^C5D_Bridge-([0-9a-f]{8})\.catvba$/i;
/** Links and read-only copies made by catia-run.exe: C5D_Bridge-<sha8>-<pid>-<yyyyMMddHHmmssfff>[-n].catvba */
export const LINK_NAME = /^C5D_Bridge-[0-9a-f]{8}-\d+-\d{17}(?:-\d{1,2})?\.catvba$/i;
/** Opening with share mode 0 fails with EBUSY while another process (CATIA) has the file open. */
const UV_FS_O_EXLOCK = 0x10000000;

export type LockState = 'unlocked' | 'locked' | 'unknown';
export interface BridgeFilesOptions {
    /** resources/bridge/C5D_Bridge.catvba of the installed extension. */
    bundledPath: string;
    /** Default: %LOCALAPPDATA%\Cat5Dev\bridge */
    bridgeDir?: string;
    /** Injected by tests; default opens the file with share mode 0. */
    lockState?: (file: string) => LockState;
}
export interface CleanupEntry { file: string; action: 'deleted' | 'kept (locked)' | 'kept (lock unknown)' | 'failed'; error?: string }
export interface BridgePreparation { master: string; action: 'reused' | 'copied' | 'replaced'; cleanup: CleanupEntry[] }
/** The bridge cannot be prepared; nothing was sent to CATIA. */
export class BridgeFileError extends Error { }

export function defaultBridgeDir(): string {
    const local = process.env.LOCALAPPDATA ?? path.join(os.homedir(), 'AppData', 'Local');
    return path.join(local, 'Cat5Dev', 'bridge');
}
export function fileLockState(file: string): LockState {
    if (process.platform !== 'win32') { return 'unknown'; }
    try { fs.closeSync(fs.openSync(file, fs.constants.O_RDONLY | UV_FS_O_EXLOCK)); return 'unlocked'; }
    catch (e) { return (e as NodeJS.ErrnoException).code === 'EBUSY' ? 'locked' : 'unknown'; }
}
function sha256(file: string): string {
    return createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}
function readOnly(file: string): boolean { return (fs.statSync(file).mode & 0o200) === 0; }
function describe(e: unknown): string {
    const error = e as NodeJS.ErrnoException;
    return `${error?.code ?? ''} ${error?.message ?? String(e)}`.trim();
}

export class BridgeFiles {
    readonly bridgeDir: string;
    private readonly lockState: (file: string) => LockState;
    private bundled?: { size: number; mtimeMs: number; sha256: string };
    constructor(private readonly options: BridgeFilesOptions) {
        this.bridgeDir = path.resolve(options.bridgeDir ?? defaultBridgeDir());
        this.lockState = options.lockState ?? fileLockState;
    }
    /** SHA-256 of the bundled bridge (cached while its size and time are unchanged). */
    bundledHash(): string {
        let stat: fs.Stats;
        try { stat = fs.statSync(this.options.bundledPath); } catch (e) { throw new BridgeFileError(`Bundled bridge not found: ${this.options.bundledPath} (${describe(e)})`); }
        if (!stat.isFile()) { throw new BridgeFileError(`Bundled bridge is not a file: ${this.options.bundledPath}`); }
        if (!this.bundled || this.bundled.size !== stat.size || this.bundled.mtimeMs !== stat.mtimeMs) {
            this.bundled = { size: stat.size, mtimeMs: stat.mtimeMs, sha256: sha256(this.options.bundledPath) };
        }
        return this.bundled.sha256;
    }
    /**
     * The bridge folder by its long, canonical name: fs.realpathSync.native expands 8.3 short names
     * (%LOCALAPPDATA% may be C:\Users\LONGUS~1\...), so --bridge never carries them. As configured while the folder does not exist.
     */
    directory(): string {
        try { return fs.realpathSync.native(this.bridgeDir); } catch { return this.bridgeDir; }
    }
    masterPath(sha8 = this.bundledHash().slice(0, 8)): string {
        return path.join(this.directory(), `C5D_Bridge-${sha8}.catvba`);
    }
    /** Cleans up (2.3), then makes sure the read-only master exists (2.1). Returns the absolute master path for --bridge. */
    prepare(): BridgePreparation {
        const hash = this.bundledHash();
        const sha8 = hash.slice(0, 8);
        try { fs.mkdirSync(this.bridgeDir, { recursive: true }); } catch (e) { throw new BridgeFileError(`Cannot create ${this.bridgeDir}: ${describe(e)}`); }
        const cleanup = this.cleanup(sha8);
        const { master, action } = this.ensureMaster(hash, sha8);
        return { master, action, cleanup };
    }
    /**
     * Deletes unlocked links/copies and old masters directly under the bridge folder; locked ones are kept for next time.
     * Hard links share the read-only attribute with their master, so files that remain get it back afterwards
     * (the current master is re-checked by ensureMaster). Names that do not match are never touched.
     */
    cleanup(sha8: string): CleanupEntry[] {
        let names: fs.Dirent[];
        const directory = this.directory();
        try { names = fs.readdirSync(directory, { withFileTypes: true }); } catch { return []; }
        const entries: CleanupEntry[] = [];
        for (const entry of names) {
            if (!entry.isFile()) { continue; }
            const master = MASTER_NAME.exec(entry.name);
            if (!LINK_NAME.test(entry.name) && !(master && master[1].toLowerCase() !== sha8)) { continue; }
            const file = path.join(directory, entry.name);
            const lock = this.lockState(file);
            if (lock !== 'unlocked') { entries.push({ file, action: lock === 'locked' ? 'kept (locked)' : 'kept (lock unknown)' }); continue; }
            try {
                fs.chmodSync(file, 0o666);
                fs.unlinkSync(file);
                entries.push({ file, action: 'deleted' });
            } catch (e) { entries.push({ file, action: 'failed', error: describe(e) }); }
        }
        for (const entry of entries) {
            if (entry.action === 'deleted') { continue; }
            try { fs.chmodSync(entry.file, 0o444); } catch (e) { entry.error = [entry.error, `read-only not restored: ${describe(e)}`].filter(Boolean).join('; '); }
        }
        return entries;
    }
    private ensureMaster(hash: string, sha8: string): { master: string; action: BridgePreparation['action'] } {
        const master = this.masterPath(sha8);
        let action: BridgePreparation['action'] = 'copied';
        let stat: fs.Stats | undefined;
        try { stat = fs.lstatSync(master); } catch (e) {
            if ((e as NodeJS.ErrnoException).code !== 'ENOENT') { throw new BridgeFileError(`Cannot inspect ${master}: ${describe(e)}`); }
        }
        if (stat) {
            if (!stat.isFile() || stat.isSymbolicLink()) { throw new BridgeFileError(`Bridge master is not a regular file: ${master}`); }
            let current: string;
            try { current = sha256(master); } catch (e) { throw new BridgeFileError(`Cannot read ${master}: ${describe(e)}`); }
            if (current === hash) { action = 'reused'; }
            else {
                // Damaged (e.g. an interrupted copy). A locked file cannot be replaced until CATIA exits.
                if (this.lockState(master) !== 'unlocked') { throw new BridgeFileError(`Bridge master is damaged and in use; restart CATIA: ${master}`); }
                try { fs.chmodSync(master, 0o666); fs.unlinkSync(master); } catch (e) { throw new BridgeFileError(`Cannot replace ${master}: ${describe(e)}`); }
                action = 'replaced';
            }
        }
        if (action !== 'reused') {
            // COPYFILE_EXCL: never overwrite a file that appeared meanwhile.
            try { fs.copyFileSync(this.options.bundledPath, master, fs.constants.COPYFILE_EXCL); } catch (e) { throw new BridgeFileError(`Cannot copy the bridge to ${master}: ${describe(e)}`); }
        }
        // Read-only before CATIA ever sees it: a link loaded read-only cannot be saved from the VBA editor.
        try { if (!readOnly(master)) { fs.chmodSync(master, 0o444); } } catch (e) { throw new BridgeFileError(`Cannot make ${master} read-only: ${describe(e)}`); }
        let verified: boolean;
        try { verified = readOnly(master) && (action === 'reused' || sha256(master) === hash); } catch (e) { throw new BridgeFileError(`Cannot verify ${master}: ${describe(e)}`); }
        if (!verified) { throw new BridgeFileError(`Bridge master is not read-only or differs from the bundled file: ${master}`); }
        return { master, action };
    }
}
