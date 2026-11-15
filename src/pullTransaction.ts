import * as fs from 'fs';
import * as path from 'path';
import { createHash } from 'crypto';
import { safeFileName } from './catiaTypes';

const managed = /\.(bas_utf|cls_utf|frm_utf)$/i;
export interface ModuleSnapshot {
    rootPath: string; modulesExisted: boolean;
    files: ReadonlyMap<string, { content: Buffer; hash: string }>;
}
export interface PullOptions {
    isDirty?: (absolutePath: string) => boolean;
    beforeMutation?: (stage: 'apply' | 'rollback', filename: string) => void;
}
export class PullConflictError extends Error { }
export class PullRecoveryError extends Error {
    constructor(public readonly recoveryPath: string, public readonly unrestored: string[], public readonly originalError: unknown) {
        super(`Pull recovery incomplete: ${recoveryPath}`);
    }
}
const hash = (data: Buffer): string => createHash('sha256').update(data).digest('hex');
const code = (e: unknown): string | undefined => (e as NodeJS.ErrnoException)?.code;
function regular(file: string): void {
    const stat = fs.lstatSync(file);
    if (!stat.isFile() || stat.isSymbolicLink()) { throw new PullConflictError(`Not a regular module: ${file}`); }
}
function fileName(name: string): void {
    if (!safeFileName(name) || !managed.test(name)) { throw new PullConflictError(`Invalid module filename: ${name}`); }
}
export function captureModules(rootPath: string): ModuleSnapshot {
    const root = path.resolve(rootPath);
    const directory = path.join(root, 'modules');
    const files = new Map<string, { content: Buffer; hash: string }>();
    let stat: fs.Stats;
    try { stat = fs.lstatSync(directory); }
    catch (e) { if (code(e) === 'ENOENT') { return { rootPath: root, modulesExisted: false, files }; } throw e; }
    if (!stat.isDirectory() || stat.isSymbolicLink()) { throw new PullConflictError('modules must be a real directory'); }
    const names = new Set<string>();
    for (const name of fs.readdirSync(directory).sort()) {
        if (!managed.test(name)) { continue; }
        fileName(name);
        if (names.has(name.toLowerCase())) { throw new PullConflictError(`Duplicate module: ${name}`); }
        names.add(name.toLowerCase());
        const fullPath = path.join(directory, name);
        regular(fullPath);
        const before = fs.statSync(fullPath);
        const content = fs.readFileSync(fullPath);
        regular(fullPath);
        const after = fs.statSync(fullPath);
        if (before.size !== after.size || before.mtimeMs !== after.mtimeMs || before.ctimeMs !== after.ctimeMs) {
            throw new PullConflictError(`Module changed while reading: ${name}`);
        }
        files.set(name, { content, hash: hash(content) });
    }
    return { rootPath: root, modulesExisted: true, files };
}
function incomingIndex(incoming: ReadonlyMap<string, string>): Map<string, { name: string; content: Buffer }> {
    const index = new Map<string, { name: string; content: Buffer }>();
    for (const [name, text] of incoming) {
        fileName(name);
        if (index.has(name.toLowerCase()) || typeof text !== 'string') { throw new PullConflictError(`Invalid incoming module: ${name}`); }
        index.set(name.toLowerCase(), { name, content: Buffer.from(text, 'utf8') });
    }
    return index;
}
function assertUnchanged(root: string, original: ModuleSnapshot): void {
    if (path.resolve(root) !== original.rootPath) { throw new PullConflictError('Snapshot workspace differs'); }
    const current = captureModules(root);
    if (current.modulesExisted !== original.modulesExisted || current.files.size !== original.files.size) {
        throw new PullConflictError('Local modules were added or removed during Pull');
    }
    for (const [name, entry] of original.files) {
        if (hash(entry.content) !== entry.hash || current.files.get(name)?.hash !== entry.hash) {
            throw new PullConflictError(`Local module changed during Pull: ${name}`);
        }
    }
}
export function listPullDeletions(snapshot: ModuleSnapshot, incoming: ReadonlyMap<string, string>): string[] {
    const index = incomingIndex(incoming);
    return [...snapshot.files.keys()].filter(name => !index.has(name.toLowerCase())).sort();
}
export function applyPull(rootPath: string, snapshot: ModuleSnapshot, incoming: ReadonlyMap<string, string>, deleteMissing: boolean,
    options: PullOptions = {}): { writtenCount: number; deletedCount: number } {
    const root = path.resolve(rootPath);
    const directory = path.join(root, 'modules');
    const originals = new Map([...snapshot.files.keys()].map(name => [name.toLowerCase(), name]));
    const received = incomingIndex(incoming);
    const prepared = [...received].map(([key, item]) => ({ ...item, name: originals.get(key) ?? item.name }));
    const writes = prepared.filter(item => snapshot.files.get(item.name)?.hash !== hash(item.content));
    const deletions = deleteMissing ? listPullDeletions(snapshot, incoming) : [];
    const assertReady = (): void => {
        assertUnchanged(root, snapshot);
        for (const name of [...prepared.map(item => item.name), ...deletions]) {
            if (options.isDirty?.(path.join(directory, name))) { throw new PullConflictError(`Unsaved module: ${name}`); }
        }
    };
    assertReady();
    if (!writes.length && !deletions.length) { return { writtenCount: 0, deletedCount: 0 }; }
    const transaction = fs.mkdtempSync(path.join(root, '.cat5dev-pull-'));
    const backup = path.join(transaction, 'backup');
    const stage = path.join(transaction, 'stage');
    const touched: string[] = [];
    let createdDirectory = false;
    const cleanup = (): void => {
        if (path.dirname(transaction) !== root || !path.basename(transaction).startsWith('.cat5dev-pull-')) { return; }
        try { fs.rmSync(transaction, { recursive: true, force: true }); } catch { /* Preserve backup if cleanup fails. */ }
    };
    try {
        fs.mkdirSync(backup); fs.mkdirSync(stage);
        for (const name of new Set([...writes.map(item => item.name), ...deletions])) {
            const original = snapshot.files.get(name);
            if (original) {
                regular(path.join(directory, name));
                fs.copyFileSync(path.join(directory, name), path.join(backup, name));
                if (hash(fs.readFileSync(path.join(backup, name))) !== original.hash) { throw new PullConflictError(`Backup changed: ${name}`); }
            }
        }
        for (const item of prepared) { fs.writeFileSync(path.join(stage, item.name), item.content, { flag: 'wx' }); }
        assertReady();
        fs.writeFileSync(path.join(transaction, 'recovery.json'), JSON.stringify({
            version: 1, rootPath: root, backup, modulesExisted: snapshot.modulesExisted,
            writes: writes.map(item => item.name), deletions,
            originalFiles: [...snapshot.files.keys()]
        }, null, 2));
        if (!snapshot.modulesExisted && writes.length) { fs.mkdirSync(directory); createdDirectory = true; }
        // Synchronous commit: do not yield to editor events between validation and writes.
        for (const item of writes) {
            const target = path.join(directory, item.name);
            options.beforeMutation?.('apply', item.name);
            const original = snapshot.files.get(item.name);
            if (original) {
                regular(target);
                if (hash(fs.readFileSync(target)) !== original.hash) { throw new PullConflictError(`Module changed: ${item.name}`); }
            } else {
                try { fs.lstatSync(target); throw new PullConflictError(`New path appeared: ${item.name}`); }
                catch (e) { if (code(e) !== 'ENOENT') { throw e; } }
            }
            touched.push(item.name);
            fs.copyFileSync(path.join(stage, item.name), target);
        }
        for (const name of deletions) {
            const target = path.join(directory, name);
            options.beforeMutation?.('apply', name); regular(target);
            if (hash(fs.readFileSync(target)) !== snapshot.files.get(name)!.hash) { throw new PullConflictError(`Module changed: ${name}`); }
            touched.push(name); fs.unlinkSync(target);
        }
    } catch (error) {
        const unrestored: string[] = [];
        for (const name of [...touched].reverse()) {
            try {
                options.beforeMutation?.('rollback', name);
                const target = path.join(directory, name);
                try { regular(target); } catch (e) { if (code(e) !== 'ENOENT') { throw e; } }
                const original = snapshot.files.get(name);
                if (original) {
                    fs.copyFileSync(path.join(backup, name), target);
                    if (hash(fs.readFileSync(target)) !== original.hash) { throw new Error('Restore hash mismatch'); }
                } else if (fs.existsSync(target)) { fs.unlinkSync(target); }
            } catch { unrestored.push(name); }
        }
        if (createdDirectory) {
            try { fs.rmdirSync(directory); } catch (e) { if (code(e) !== 'ENOENT') { unrestored.push('modules/'); } }
        }
        if (unrestored.length) {
            try { fs.writeFileSync(path.join(transaction, 'recovery.json'), JSON.stringify({ rootPath: root, backup, touched, unrestored, error: String(error) }, null, 2)); } catch { }
            throw new PullRecoveryError(transaction, unrestored, error);
        }
        cleanup(); throw error;
    }
    cleanup();
    return { writtenCount: writes.length, deletedCount: deletions.length };
}
