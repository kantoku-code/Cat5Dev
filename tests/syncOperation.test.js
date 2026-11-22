const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fork } = require('node:child_process');
const { once } = require('node:events');
const { fixture, result, shortPath } = require('./helpers');
const { SyncOperationManager, SyncBusyError } = require('../out/syncOperation');
function manager(root, options = {}) { return new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root, ...options }); }
test('operation isolation and immediate same-process exclusion', t => {
    const root = fixture(t), first = manager(root), second = manager(root);
    const context = first.acquire('pull');
    assert.throws(() => second.acquire('push'), SyncBusyError);
    const dir = path.join(context.tempDir, 'step'); fs.mkdirSync(dir);
    context.markExternalRunning(path.join(dir, 'result.json'), { operationId: context.id, stepId: 'step', kind: 'list', targetProject: '' });
    context.markSettled(); context.close();
    const next = second.acquire('push'); assert.notEqual(next.tempDir, context.tempDir); next.close();
    assert.ok(!fs.existsSync(context.tempDir));
});
test('live owner, corrupt metadata and process-query failure fail closed', async t => {
    const root = fixture(t), owner = manager(root).acquire('pull');
    assert.equal(await manager(root).recover(), false);
    owner.markExternalRunning(path.join(owner.tempDir, 'result.json'), { operationId: owner.id, stepId: 'step', kind: 'pull', targetProject: 'Review' }); owner.close();
    assert.equal(await manager(root, { isProcessAlive: () => false, isCatiaRunning: async () => { throw Error('denied'); } }).recover(), false);
    fs.writeFileSync(path.join(root, 'locks', 'active', 'owner.json'), '{}');
    assert.equal(await manager(root).recover(), false);
    assert.throws(() => manager(root).acquire('push'), SyncBusyError);
});
test('dead external owner requires terminal result or CATIA absence', async t => {
    for (const terminal of [false, true]) {
        const root = fixture(t), owner = manager(root).acquire('pull');
        const expected = { operationId: owner.id, stepId: 'step', kind: 'pull', targetProject: 'Review' };
        const resultPath = path.join(owner.tempDir, 'result.json'); owner.markExternalRunning(resultPath, expected); owner.close();
        const recovery = manager(root, { isProcessAlive: () => false, isCatiaRunning: async () => true });
        assert.equal(await recovery.recover(), false);
        if (terminal) fs.writeFileSync(resultPath, JSON.stringify(result(expected)));
        const safe = manager(root, { isProcessAlive: () => false, isCatiaRunning: async () => false });
        assert.equal(await (terminal ? recovery : safe).recover(), true);
        assert.ok(fs.existsSync(owner.tempDir)); // Diagnostic data survives recovery.
        const next = manager(root).acquire('push'); next.close();
    }
});
test('two stale reclaimers never retire a newly acquired lock', async t => {
    const root = fixture(t), owner = manager(root).acquire('pull');
    owner.markExternalRunning(path.join(owner.tempDir, 'result.json'), { operationId: owner.id, stepId: 'step', kind: 'pull', targetProject: 'Review' }); owner.close();
    let resolve1, resolve2;
    const first = manager(root, { isProcessAlive: () => false, isCatiaRunning: () => new Promise(r => { resolve1 = r; }) });
    const second = manager(root, { isProcessAlive: () => false, isCatiaRunning: () => new Promise(r => { resolve2 = r; }) });
    const a = first.recover(), b = second.recover();
    resolve1(false); assert.equal(await a, true);
    const next = manager(root).acquire('push');
    resolve2(false); assert.equal(await b, false);
    assert.equal(JSON.parse(fs.readFileSync(path.join(root, 'locks', 'active', 'owner.json'))).id, next.id); next.close();
});
test('only this manager\'s quarantined (not running) lock can be force-released', t => {
    const root = fixture(t), own = manager(root), other = manager(root);
    const context = own.acquire('push');
    context.markExternalRunning(path.join(context.tempDir, 'result.json'), { operationId: context.id, stepId: 'step', kind: 'push', targetProject: 'Review' });
    // Still running: the in-flight operation is never offered for release.
    assert.equal(own.hasQuarantine(), false); assert.equal(own.releaseQuarantined(), false);
    context.close();
    assert.equal(other.hasQuarantine(), false); assert.equal(other.releaseQuarantined(), false);
    assert.equal(own.hasQuarantine(), true);
    assert.equal(own.releaseQuarantined(), true);
    assert.ok(context.isClosed()); assert.equal(own.hasQuarantine(), false);
    assert.ok(fs.existsSync(context.tempDir)); // Diagnostic data survives release.
    assert.throws(() => context.markSettled());
    const next = other.acquire('push'); assert.equal(own.hasQuarantine(), false); next.close();
});
test('quarantine record is cleared once the lock is released by a late result', t => {
    const root = fixture(t), own = manager(root);
    const context = own.acquire('push');
    context.markExternalRunning(path.join(context.tempDir, 'result.json'), { operationId: context.id, stepId: 'step', kind: 'push', targetProject: 'Review' });
    context.close(); assert.equal(own.hasQuarantine(), true);
    context.markSettled(); context.close();
    assert.equal(own.hasQuarantine(), false); assert.equal(own.releaseQuarantined(), false);
    own.acquire('push').close();
});
test('orphaned lock is releasable even while CATIA runs, but a live owner\'s lock is not', t => {
    const root = fixture(t), owner = manager(root).acquire('push');
    owner.markExternalRunning(path.join(owner.tempDir, 'result.json'), { operationId: owner.id, stepId: 'step', kind: 'push', targetProject: 'Review' });
    // Simulate another process's lock: its owner pid is not this process.
    const ownerPath = path.join(root, 'locks', 'active', 'owner.json');
    fs.writeFileSync(ownerPath, JSON.stringify({ ...JSON.parse(fs.readFileSync(ownerPath, 'utf8')), pid: process.pid + 1 }));
    const live = manager(root, { isProcessAlive: () => true, isCatiaRunning: async () => true });
    assert.equal(live.hasQuarantine(), false); assert.equal(live.releaseQuarantined(), false);
    const failing = manager(root, { isProcessAlive: () => { throw Error('denied'); } });
    assert.equal(failing.hasQuarantine(), false);
    const orphan = manager(root, { isProcessAlive: () => false, isCatiaRunning: async () => true });
    assert.equal(orphan.hasQuarantine(), true);
    assert.equal(orphan.releaseQuarantined(), true);
    assert.ok(fs.existsSync(owner.tempDir)); // Diagnostic data survives release.
    const next = manager(root).acquire('push');
    // A delayed second release must not retire the newer lock.
    assert.equal(orphan.hasQuarantine(), false); assert.equal(orphan.releaseQuarantined(), false);
    next.close();
});
test('real child processes exclude each other across workspaces and recover after crash', async t => {
    const root = fixture(t), args = [path.join(root, 'locks'), root];
    const spawn = mode => fork(path.join(__dirname, 'lock-worker.js'), [...args, mode], { windowsHide: true, stdio: ['ignore', 'ignore', 'pipe', 'ipc'] });
    const first = spawn('hold'); t.after(() => first.kill()); const [a] = await once(first, 'message'); assert.equal(a.status, 'acquired');
    const second = spawn('hold'); const secondExit = once(second, 'exit'); const [b] = await once(second, 'message'); assert.equal(b.status, 'busy'); await secondExit;
    const firstExit = once(first, 'exit'); first.send('close'); await firstExit;
    const crash = spawn('crash'); const crashExit = once(crash, 'exit'); const [c] = await once(crash, 'message'); assert.equal(c.status, 'crashed'); await crashExit;
    assert.equal(await manager(root, { isCatiaRunning: async () => true }).recover(), false);
    assert.equal(await manager(root, { isCatiaRunning: async () => false }).recover(), true);
});
test('temporary folders use the long name of tempRoot (fs.realpathSync.native expands 8.3 short names)', t => {
    const root = fixture(t);
    const native = fs.realpathSync.native;
    const calls = [];
    fs.realpathSync.native = (file, ...rest) => { calls.push(file); return native(file, ...rest); };
    let plain;
    try { plain = manager(root); } finally { fs.realpathSync.native = native; }
    assert.ok(calls.includes(root), JSON.stringify(calls));
    const context = plain.acquire('pull');
    assert.equal(path.dirname(context.tempDir), native(root));
    context.close();
    const short = shortPath(root);
    if (!short) { t.diagnostic('no 8.3 short name on this volume; the short-path half is skipped'); return; }
    const viaShort = new SyncOperationManager({ lockRoot: path.join(short, 'locks'), tempRoot: short });
    const other = viaShort.acquire('pull');
    assert.equal(path.dirname(other.tempDir), native(root), `${short} -> ${other.tempDir}`);
    // A lock written with the short spelling (an older version) is still recognised and can be released.
    other.markExternalRunning(path.join(other.tempDir, 'result.json'), { operationId: other.id, stepId: 'step', kind: 'pull', targetProject: 'Review' });
    other.close();
    const ownerFile = path.join(native(root), 'locks', 'active', 'owner.json');
    const owner = JSON.parse(fs.readFileSync(ownerFile, 'utf8'));
    fs.writeFileSync(ownerFile, JSON.stringify({ ...owner, tempDir: path.join(short, path.basename(owner.tempDir)),
        resultPath: path.join(short, path.basename(owner.tempDir), 'result.json') }));
    assert.equal(viaShort.hasQuarantine(), true);
    assert.equal(viaShort.releaseQuarantined(), true);
    assert.ok(!fs.existsSync(path.join(native(root), 'locks', 'active')));
});
