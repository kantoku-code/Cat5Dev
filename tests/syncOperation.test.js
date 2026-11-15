const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fork } = require('node:child_process');
const { once } = require('node:events');
const { fixture, result } = require('./helpers');
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
