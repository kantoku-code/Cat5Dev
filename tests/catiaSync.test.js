const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const { fixture, result } = require('./helpers');
const { SyncOperationManager } = require('../out/syncOperation');
let state;
const vscode = {
    env: { language: 'ja' }, ProgressLocation: { Notification: 1 },
    workspace: {
        get workspaceFolders() { return [{ uri: { fsPath: state.root } }]; },
        get textDocuments() { return state.documents || []; },
        getConfiguration() { return { get: (_, fallback) => fallback, inspect: () => ({}) }; }
    },
    commands: { executeCommand: async command => { state.events.push(['command', command]); } },
    window: {
        withProgress: async (_options, task) => task(),
        showInformationMessage: async message => { state.events.push(['info', message]); },
        showErrorMessage: async message => { state.events.push(['error', message]); },
        showWarningMessage: async (message, _options, ...actions) => {
            state.events.push(['warning', message]);
            return state.choice === 'delete' ? actions[0] : state.choice === 'keep' ? actions[1] : undefined;
        },
        showQuickPick: async () => state.selection
    }
};
const originalLoad = Module._load;
Module._load = function (id, ...args) { return id === 'vscode' ? vscode : originalLoad.call(this, id, ...args); };
const { CatiaSync } = require('../out/catiaSync');
Module._load = originalLoad;
function setup(t, options = {}) {
    const root = fixture(t); fs.mkdirSync(path.join(root, 'modules'));
    fs.writeFileSync(path.join(root, 'cat5dev.toml'), `[project]\ntarget_project = "${options.noTarget ? '' : 'Review'}"\nlanguage = "ja"\nencoding = "shift_jis"\n`);
    fs.writeFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'old');
    fs.writeFileSync(path.join(root, 'modules', 'LocalOnly.cls_utf'), 'local');
    state = { root, events: [], choice: options.choice, selection: options.selection, documents: options.documents };
    const steps = [];
    const manager = new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root });
    const output = { appendLine: () => {}, show: () => {} };
    const execute = async script => {
        const directory = path.dirname(script), source = fs.readFileSync(path.join(directory, 'operation.catvbs'), 'utf8');
        const get = name => new RegExp(`${name} = "([^"]*)"`).exec(source)[1];
        const request = { operationId: get('gOperationId'), stepId: get('gStepId'), kind: get('gKind'), targetProject: get('gTargetProject') };
        steps.push(request.kind);
        if (options.beforeResult) await options.beforeResult(request);
        let r;
        if (options.fail === request.kind) r = result(request, { failures: [{ phase: 'injected', target: '', code: 5, message: 'failure' }] });
        else if (request.kind === 'list') r = result(request, { projects: ['Review'], expectedCount: 1, succeededCount: 1 });
        else if (request.kind === 'pull' || request.kind === 'check') {
            const items = options.empty ? [] : [{ name: 'C5D_UserModule', type: 1, file: 'm0.txt' }];
            for (const item of items) fs.writeFileSync(path.join(directory, item.file), 'new source');
            r = result(request, { items, expectedCount: items.length, succeededCount: items.length });
        } else r = result(request, { expectedCount: 1, succeededCount: 1, updatedCount: 1 });
        fs.writeFileSync(path.join(directory, 'result.json'), JSON.stringify(r));
        return { stdout: '', stderr: '' };
    };
    return { root, steps, events: state.events, sync: new CatiaSync({}, output, manager, { execute, timeoutMs: 1000 }) };
}
test('Pull overwrite has no confirmation; only missing-file deletion is confirmed', async t => {
    for (const choice of ['keep', 'delete', undefined]) {
        const { root, sync, events } = setup(t, { choice });
        await sync.pull();
        const directory = path.join(root, 'modules');
        assert.equal(events.filter(e => e[0] === 'warning').length, 1);
        assert.equal(fs.readFileSync(path.join(directory, 'C5D_UserModule.bas_utf'), 'utf8'), choice ? 'new source\n' : 'old');
        assert.equal(fs.existsSync(path.join(directory, 'LocalOnly.cls_utf')), choice !== 'delete');
        assert.equal(events.some(e => e[0] === 'error'), false);
    }
});
test('target selection cancellation and Pull failure preserve local content', async t => {
    for (const options of [{ noTarget: true }, { fail: 'pull' }]) {
        const { root, sync, steps, events } = setup(t, options);
        await sync.pull();
        assert.equal(fs.readFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), 'old');
        assert.ok(!events.some(e => e[0] === 'info' && e[1].includes('正常にプル')));
        if (options.noTarget) assert.deepEqual(steps, ['list']);
    }
});
test('nested target selection uses the same operation and completes before returned Promise', async t => {
    const { sync, steps, root } = setup(t, { noTarget: true, selection: 'Review', choice: 'keep' });
    await sync.pull(); assert.deepEqual(steps, ['list', 'pull']);
    assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')));
    assert.equal(fs.readFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), 'new source\n');
});
test('edits saved during export abort local application', async t => {
    const { root, sync, events } = setup(t, { choice: 'keep', beforeResult: request => {
        if (request.kind === 'pull') fs.writeFileSync(path.join(state.root, 'modules', 'C5D_UserModule.bas_utf'), 'edited during export');
    } });
    await sync.pull();
    assert.equal(fs.readFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), 'edited during export');
    assert.ok(events.some(e => e[0] === 'warning' && e[1].includes('ローカル変更')));
});
test('dirty updated modules abort, while dirty kept local-only modules are preserved', async t => {
    for (const filename of ['C5D_UserModule.bas_utf', 'LocalOnly.cls_utf']) {
        const { sync, root, events } = setup(t, { choice: 'keep' });
        state.documents = [{ isDirty: true, uri: { fsPath: path.join(root, 'modules', filename) } }];
        await sync.pull();
        assert.equal(fs.readFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), filename.startsWith('C5D') ? 'old' : 'new source\n');
        assert.ok(!events.some(e => e[0] === 'error'));
    }
});
test('Push preflight failure prevents mutations and success notification', async t => {
    const { sync, steps, events } = setup(t, { fail: 'check' }); await sync.push();
    assert.deepEqual(steps, ['check']); assert.ok(events.some(e => e[0] === 'error'));
    assert.ok(!events.some(e => e[0] === 'info' && e[1].includes('プッシュしました')));
});
test('Push failure never reports success and always releases settled lock', async t => {
    const { sync, steps, events, root } = setup(t, { fail: 'push', choice: 'keep' }); await sync.push();
    assert.deepEqual(steps, ['check', 'push']); assert.ok(events.some(e => e[0] === 'error'));
    assert.ok(!events.some(e => e[0] === 'info' && e[1].includes('プッシュしました')));
    assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')));
});
