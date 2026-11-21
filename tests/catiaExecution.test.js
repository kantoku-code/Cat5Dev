const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fixture, result } = require('./helpers');
const { validateCatiaResult, catiaSucceeded, normalizeModuleText } = require('../out/catiaTypes');
const { createCatiaRequest, runCatiaRequest, CatiaExecutionError, CATIA_SCRIPT_NAME } = require('../out/catiaExecution');
const { SyncOperationManager } = require('../out/syncOperation');
function setup(t) {
    const root = fixture(t);
    const manager = new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root });
    const context = manager.acquire('test'); t.after(() => context.close());
    return { root, context, request: createCatiaRequest(context, 'pull', 'Review', 'shift_jis') };
}
test('result validation accepts BOM and confirmed partial failure but rejects malformed results', () => {
    const req = { operationId: 'op', stepId: 'step', kind: 'pull', targetProject: 'Review' };
    const good = result(req);
    assert.ok(catiaSucceeded(validateCatiaResult('\uFEFF' + JSON.stringify(good), req)));
    const partial = result(req, { expectedCount: 1, failures: [{ phase: 'export', target: 'A', code: 5, message: 'failed' }] });
    assert.ok(!catiaSucceeded(validateCatiaResult(JSON.stringify(partial), req)));
    for (const change of [{ operationId: 'other' }, { stepId: 'other' }, { targetProject: 'other' }, { expectedCount: -1 }, { expectedCount: 1 }, { failures: null }, { items: [{ name: 'A', type: 1, file: '../outside' }] }]) {
        assert.throws(() => validateCatiaResult(JSON.stringify({ ...good, ...change }), req));
    }
});
test('valid empty result settles the operation', async t => {
    const { context, request } = setup(t);
    const r = await runCatiaRequest(context, request, { execute: async () => {
        fs.writeFileSync(path.join(request.directory, 'result.json'), '\uFEFF' + JSON.stringify(result(request)));
        return { stdout: '', stderr: '' };
    } });
    assert.equal(r.succeededCount, 0);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('process errors, partial failures and mismatched request counts never succeed', async t => {
    for (const mode of ['process', 'partial', 'count']) {
        const { context, request } = setup(t); request.kind = 'push'; request.modules = [{ name: 'A', type: 1, file: 'a.txt' }];
        await assert.rejects(runCatiaRequest(context, request, { execute: async () => {
            const r = mode === 'partial' ? result(request, { expectedCount: 1, failures: [{ phase: 'push.delete', target: 'A', code: 5, message: 'denied' }] }) : result(request);
            fs.writeFileSync(path.join(request.directory, 'result.json'), JSON.stringify(r));
            return { error: mode === 'process' ? Object.assign(Error('exit'), { code: 12 }) : null, stdout: '', stderr: '' };
        } }), CatiaExecutionError);
        context.close(); assert.ok(!fs.existsSync(context.tempDir));
    }
});
test('missing, malformed or unrelated results quarantine lock and temporary files', async t => {
    for (const mode of ['missing', 'malformed', 'unrelated']) {
        const { context, request } = setup(t);
        await assert.rejects(runCatiaRequest(context, request, { timeoutMs: 15, execute: async () => {
            if (mode !== 'missing') fs.writeFileSync(path.join(request.directory, 'result.json'), mode === 'malformed' ? 'bad' : JSON.stringify(result(request, { operationId: 'other' })));
            return { stdout: '', stderr: '' };
        } }), e => e instanceof CatiaExecutionError && e.uncertain);
        context.close(); assert.ok(fs.existsSync(context.tempDir));
    }
});
test('known pre-dispatch failure releases lock without result', async t => {
    const { context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { execute: async () => ({ error: Object.assign(Error('no CATIA'), { code: 10 }), stdout: '', stderr: '' }) }), e => !e.uncertain);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('executor receives the request directory and script name; no runner script is written', async t => {
    const { context, request } = setup(t);
    let args;
    await runCatiaRequest(context, request, { execute: async (...received) => {
        args = received;
        fs.writeFileSync(path.join(request.directory, 'result.json'), JSON.stringify(result(request)));
        return { stdout: '', stderr: '' };
    } });
    assert.deepEqual(args.slice(0, 2), [request.directory, CATIA_SCRIPT_NAME]);
    assert.ok(fs.existsSync(path.join(request.directory, CATIA_SCRIPT_NAME)));
    assert.ok(!fs.readdirSync(request.directory).some(name => name.endsWith('.vbs')));
    context.close();
});
test('missing bridge executable is a known pre-dispatch failure', async t => {
    const { root, context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { bridgePath: path.join(root, 'missing', 'catia-run.exe') }), e => e instanceof CatiaExecutionError && !e.uncertain);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('without bridge path or executor nothing is dispatched', async t => {
    const { context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request), e => e instanceof CatiaExecutionError && !e.uncertain);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('export manifest requires all files', async t => {
    const { context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { execute: async () => {
        fs.writeFileSync(path.join(request.directory, 'result.json'), JSON.stringify(result(request, { expectedCount: 1, succeededCount: 1, items: [{ name: 'A', type: 1, file: 'missing.txt' }] })));
        return { stdout: '', stderr: '' };
    } }));
    context.close();
});
test('timeout retains lock until a late valid terminal result arrives', async t => {
    const { context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { timeoutMs: 10, execute: async () => new Promise(() => {}) }), e => e.uncertain);
    context.close(); assert.ok(fs.existsSync(context.tempDir));
    fs.writeFileSync(path.join(request.directory, 'result.json'), JSON.stringify(result(request)));
    await new Promise(resolve => setTimeout(resolve, 1150));
    assert.ok(!fs.existsSync(context.tempDir));
});
test('force-released timeout ignores a late result and keeps the newer lock', async t => {
    const root = fixture(t);
    const manager = new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root });
    const context = manager.acquire('push');
    const request = createCatiaRequest(context, 'push', 'Review', 'shift_jis');
    await assert.rejects(runCatiaRequest(context, request, { timeoutMs: 10, execute: async () => new Promise(() => {}) }), e => e.uncertain);
    context.close();
    assert.equal(manager.releaseQuarantined(), true);
    const next = manager.acquire('push');
    fs.writeFileSync(path.join(request.directory, 'result.json'), JSON.stringify(result(request)));
    await new Promise(resolve => setTimeout(resolve, 1150));
    assert.equal(JSON.parse(fs.readFileSync(path.join(root, 'locks', 'active', 'owner.json'))).id, next.id);
    assert.ok(fs.existsSync(context.tempDir)); // The abandoned watcher did not clean up.
    next.close();
});
test('module text normalization unifies CRLF, CR and LF and trims trailing whitespace', () => {
    assert.equal(normalizeModuleText('a\r\nb\rc\nd\r\n\r\n  '), 'a\nb\nc\nd');
    assert.equal(normalizeModuleText('a\r\nb'), normalizeModuleText('a\nb\n'));
    assert.equal(normalizeModuleText('  indented\r\n'), '  indented');
    assert.equal(normalizeModuleText(''), '');
    assert.notEqual(normalizeModuleText('a\r\nb'), normalizeModuleText('a\r\nc'));
});
