const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const childProcess = require('node:child_process');
const { fixture, result } = require('./helpers');
const { validateCatiaResult, catiaSucceeded, normalizeModuleText } = require('../out/catiaTypes');
const { createCatiaRequest, runCatiaRequest, bridgeExecutor, buildRequestJson, CatiaExecutionError, EXIT, REQUEST_FILE_NAME, RESULT_FILE_NAME } = require('../out/catiaExecution');
const { SyncOperationManager } = require('../out/syncOperation');
const bridgeMaster = 'C:\\Users\\Someone\\AppData\\Local\\Cat5Dev\\bridge\\C5D_Bridge-0123abcd.catvba';
function setup(t) {
    const root = fixture(t);
    const manager = new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root });
    const context = manager.acquire('test'); t.after(() => context.close());
    return { root, context, request: createCatiaRequest(context, 'pull', 'Review') };
}
const writeResult = (request, value) => fs.writeFileSync(path.join(request.directory, RESULT_FILE_NAME), typeof value === 'string' ? value : JSON.stringify(value));
const exit = (code, message = `exit ${code}`) => ({ error: Object.assign(Error(message), { code }), stdout: '', stderr: '' });
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
    const r = await runCatiaRequest(context, request, { bridgeMaster, execute: async () => {
        writeResult(request, '\uFEFF' + JSON.stringify(result(request)));
        return { stdout: '', stderr: '' };
    } });
    assert.equal(r.succeededCount, 0);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('process errors, partial failures and mismatched request counts never succeed', async t => {
    for (const mode of ['process', 'partial', 'count']) {
        const { context, request } = setup(t); request.kind = 'push'; request.modules = [{ name: 'A', type: 1, text: 'Sub A()\r\nEnd Sub' }];
        await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, execute: async () => {
            const r = mode === 'partial' ? result(request, { expectedCount: 1, failures: [{ phase: 'push.delete', target: 'A', code: 5, message: 'denied' }] }) : result(request);
            writeResult(request, r);
            return mode === 'process' ? exit(12) : { stdout: '', stderr: '' };
        } }), CatiaExecutionError);
        context.close(); assert.ok(!fs.existsSync(context.tempDir));
    }
});
test('missing, malformed or unrelated results quarantine lock and temporary files', async t => {
    for (const mode of ['missing', 'malformed', 'unrelated']) {
        const { context, request } = setup(t);
        await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, timeoutMs: 15, execute: async () => {
            if (mode !== 'missing') writeResult(request, mode === 'malformed' ? 'bad' : result(request, { operationId: 'other' }));
            return { stdout: '', stderr: '' };
        } }), e => e instanceof CatiaExecutionError && e.uncertain);
        context.close(); assert.ok(fs.existsSync(context.tempDir));
    }
});
test('exit codes 2, 10, 11, 13, 14 and 15 release the lock at once and keep the exit code', async t => {
    assert.equal(EXIT.runnerBusy, 15);
    for (const code of [2, 10, 11, 13, 14, 15]) {
        const { context, request } = setup(t);
        await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, execute: async () => exit(code) }),
            e => e instanceof CatiaExecutionError && !e.uncertain && e.exitCode === code && !e.result);
        context.close(); assert.ok(!fs.existsSync(context.tempDir), String(code));
    }
});
test('exit 12 with a result is a confirmed partial failure; without a result completion is unconfirmed', async t => {
    const logs = [];
    const withResult = setup(t);
    await assert.rejects(runCatiaRequest(withResult.context, withResult.request, { bridgeMaster, log: text => logs.push(text), execute: async () => {
        writeResult(withResult.request, result(withResult.request, { expectedCount: 1, failures: [{ phase: 'export.call', target: 'Review', code: -2147352567, message: 'boom' }] }));
        return { ...exit(12), stderr: '0x80020009: boom' };
    } }), e => e instanceof CatiaExecutionError && !e.uncertain && e.exitCode === 12 && e.result?.failures.length === 1);
    assert.ok(logs.some(line => line.includes('ExecuteScript failed (exit 12)') && line.includes('0x80020009: boom')), JSON.stringify(logs));
    withResult.context.close(); assert.ok(!fs.existsSync(withResult.context.tempDir));
    const without = setup(t);
    await assert.rejects(runCatiaRequest(without.context, without.request, { bridgeMaster, execute: async () => exit(12) }),
        e => e instanceof CatiaExecutionError && e.uncertain && e.exitCode === 12);
    without.context.close(); assert.ok(fs.existsSync(without.context.tempDir));
});
test('executor receives the bridge master and request/result paths; request.json follows the contract', async t => {
    const { context, request } = setup(t);
    request.kind = 'push';
    request.modules = [{ name: 'Module1', type: 1, text: 'Sub A()\r\n    MsgBox "日本語 🙂"\r\nEnd Sub' }, { name: 'Class1', type: 2, text: '' }];
    request.deletions = ['Old'];
    let received;
    await runCatiaRequest(context, request, { bridgeMaster, execute: async (invocation, timeoutMs) => {
        received = { invocation, timeoutMs, json: fs.readFileSync(invocation.request) };
        writeResult(request, result(request, { expectedCount: 3, succeededCount: 3, updatedCount: 2, deletedCount: 1 }));
        return { stdout: '', stderr: '' };
    } });
    assert.deepEqual(received.invocation, { bridge: bridgeMaster, request: path.join(request.directory, REQUEST_FILE_NAME), result: path.join(request.directory, RESULT_FILE_NAME) });
    assert.equal(received.timeoutMs, 60000);
    assert.notEqual(received.json[0], 0xEF, 'no BOM');
    const json = JSON.parse(received.json.toString('utf8'));
    assert.deepEqual(json, { version: 2, operationId: context.id, stepId: request.stepId, kind: 'push', targetProject: 'Review',
        modules: request.modules, deletions: ['Old'] });
    assert.ok(!('encoding' in json));
    assert.ok(!fs.readdirSync(request.directory).some(name => /\.(vbs|catvbs|catscript)$/i.test(name)));
    context.close();
});
test('non-push requests carry empty modules and deletions', async t => {
    const { context, request } = setup(t);
    let json;
    await runCatiaRequest(context, request, { bridgeMaster, execute: async invocation => {
        json = JSON.parse(fs.readFileSync(invocation.request, 'utf8'));
        writeResult(request, result(request));
        return { stdout: '', stderr: '' };
    } });
    assert.deepEqual(json, { version: 2, operationId: context.id, stepId: request.stepId, kind: 'pull', targetProject: 'Review', modules: [], deletions: [] });
    context.close();
});
test('requests that break the contract are refused before dispatch and the lock is released', async t => {
    const base = { operationId: 'op', stepId: 'step', directory: 'C:\\x' };
    assert.throws(() => buildRequestJson({ ...base, kind: 'list', targetProject: 'Review' }));
    assert.throws(() => buildRequestJson({ ...base, kind: 'pull', targetProject: '' }));
    assert.throws(() => buildRequestJson({ ...base, kind: 'pull', targetProject: 'Review', modules: [{ name: 'A', type: 1, text: '' }] }));
    assert.throws(() => buildRequestJson({ ...base, kind: 'push', targetProject: 'Review', modules: [{ name: 'A', type: 1, text: '' }], deletions: ['a'] }));
    assert.throws(() => buildRequestJson({ ...base, kind: 'push', targetProject: 'Review', modules: [{ name: 'A', type: 4, text: '' }] }));
    assert.throws(() => buildRequestJson({ ...base, kind: 'push', targetProject: 'Review', modules: [{ name: 'A', type: 1 }] }));
    assert.throws(() => buildRequestJson({ ...base, kind: 'push', targetProject: 'Review', deletions: ['a/b'] }));
    assert.doesNotThrow(() => buildRequestJson({ ...base, kind: 'list', targetProject: '' }));
    const { context, request } = setup(t);
    request.kind = 'push'; request.modules = [{ name: 'A', type: 1, text: 'x' }, { name: 'a', type: 1, text: 'y' }];
    let called = false;
    await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, execute: async () => { called = true; return { stdout: '', stderr: '' }; } }));
    assert.equal(called, false);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('bridgeExecutor runs catia-run.exe with --bridge, --request and --result', async () => {
    const original = childProcess.execFile;
    let call;
    childProcess.execFile = (file, args, options, callback) => { call = { file, args, options }; callback(null, 'load 1ms\n', ''); };
    try {
        const output = await bridgeExecutor('C:\\ext\\bin\\catia-run.exe')({ bridge: bridgeMaster, request: 'C:\\t\\request.json', result: 'C:\\t\\result.json' }, 1234);
        assert.equal(output.stdout, 'load 1ms\n');
    } finally { childProcess.execFile = original; }
    assert.equal(call.file, 'C:\\ext\\bin\\catia-run.exe');
    assert.deepEqual(call.args, ['--bridge', bridgeMaster, '--request', 'C:\\t\\request.json', '--result', 'C:\\t\\result.json']);
    assert.equal(call.options.timeout, 1234);
    assert.equal(call.options.windowsHide, true);
});
test('missing runner executable is a known pre-dispatch failure', async t => {
    const { root, context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, runnerPath: path.join(root, 'missing', 'catia-run.exe') }), e => e instanceof CatiaExecutionError && !e.uncertain);
    context.close(); assert.ok(!fs.existsSync(context.tempDir));
});
test('without runner path, executor or bridge master nothing is dispatched', async t => {
    for (const options of [{ bridgeMaster }, { execute: async () => { throw new Error('must not run'); } }, { execute: async () => { throw new Error('must not run'); }, bridgeMaster: 'relative.catvba' }]) {
        const { context, request } = setup(t);
        await assert.rejects(runCatiaRequest(context, request, options), e => e instanceof CatiaExecutionError && !e.uncertain);
        assert.ok(!fs.existsSync(path.join(request.directory, REQUEST_FILE_NAME)));
        context.close(); assert.ok(!fs.existsSync(context.tempDir));
    }
});
test('export manifest requires all files', async t => {
    const { context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, execute: async () => {
        writeResult(request, result(request, { expectedCount: 1, succeededCount: 1, items: [{ name: 'A', type: 1, file: 'missing.txt' }] }));
        return { stdout: '', stderr: '' };
    } }));
    context.close();
});
test('timeout retains lock until a late valid terminal result arrives', async t => {
    const { context, request } = setup(t);
    await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, timeoutMs: 10, execute: async () => new Promise(() => {}) }), e => e.uncertain);
    context.close(); assert.ok(fs.existsSync(context.tempDir));
    writeResult(request, result(request));
    await new Promise(resolve => setTimeout(resolve, 1150));
    assert.ok(!fs.existsSync(context.tempDir));
});
test('force-released timeout ignores a late result and keeps the newer lock', async t => {
    const root = fixture(t);
    const manager = new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root });
    const context = manager.acquire('push');
    const request = createCatiaRequest(context, 'push', 'Review');
    await assert.rejects(runCatiaRequest(context, request, { bridgeMaster, timeoutMs: 10, execute: async () => new Promise(() => {}) }), e => e.uncertain);
    context.close();
    assert.equal(manager.releaseQuarantined(), true);
    const next = manager.acquire('push');
    writeResult(request, result(request));
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
