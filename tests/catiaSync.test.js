const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const { fixture, result } = require('./helpers');
const { SyncOperationManager } = require('../out/syncOperation');
const { BridgeFiles } = require('../out/bridgeFiles');
let state;
const vscode = {
    env: { language: 'ja' }, ProgressLocation: { Notification: 1 },
    workspace: {
        get workspaceFolders() { return [{ uri: { fsPath: state.root } }]; },
        get textDocuments() { return state.documents || []; },
        getConfiguration(section) { return { get: (key, fallback) => state?.config?.[`${section}.${key}`] ?? fallback, inspect: () => ({}) }; }
    },
    commands: { executeCommand: async command => { state.events.push(['command', command]); } },
    window: {
        withProgress: async (_options, task) => task(),
        // hangNotifications: VS Code はボタンなし通知の Promise をユーザーが閉じるまで resolve しない
        showInformationMessage: message => { state.events.push(['info', message]); return notification(); },
        showErrorMessage: message => { state.events.push(['error', message]); return notification(); },
        showWarningMessage: async (message, _options, ...actions) => {
            state.events.push(['warning', message]);
            if (!_options?.modal && !actions.length) return notification();
            if (state.choice === 'continue') return actions.find(action => action === '続行');
            return state.choice === 'delete' ? actions[0] : state.choice === 'keep' ? actions[1] : undefined;
        },
        showQuickPick: async () => state.selection
    }
};
function notification() { return state.hangNotifications ? new Promise(() => {}) : Promise.resolve(undefined); }
const originalLoad = Module._load;
Module._load = function (id, ...args) { return id === 'vscode' ? vscode : originalLoad.call(this, id, ...args); };
const { CatiaSync } = require('../out/catiaSync');
Module._load = originalLoad;
function setup(t, options = {}) {
    const root = fixture(t); fs.mkdirSync(path.join(root, 'modules'));
    fs.writeFileSync(path.join(root, 'cat5dev.toml'), `[project]\ntarget_project = "${options.noTarget ? '' : 'Review'}"\nlanguage = "ja"\nencoding = "shift_jis"\n`);
    const localFiles = options.localFiles || { 'C5D_UserModule.bas_utf': 'old', 'LocalOnly.cls_utf': 'local' };
    for (const [name, content] of Object.entries(localFiles)) fs.writeFileSync(path.join(root, 'modules', name), content);
    state = { root, events: [], choice: options.choice, selection: options.selection, documents: options.documents, hangNotifications: options.hangNotifications };
    const steps = [], applies = [], deletes = [];
    const manager = new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root });
    const logs = [];
    const output = { appendLine: line => logs.push(line), show: () => {} };
    const invocations = [];
    // Fake catia-run.exe (bridge-contract 12): reads request.json, writes m<i>.txt (UTF-8, no BOM) and result.json.
    const execute = async invocation => {
        invocations.push(invocation);
        const directory = path.dirname(invocation.request);
        const request = { ...JSON.parse(fs.readFileSync(invocation.request, 'utf8')), directory };
        steps.push(request.kind);
        if (options.beforeResult) await options.beforeResult(request);
        if (options.exitCode) return { error: Object.assign(Error(`exit ${options.exitCode}`), { code: options.exitCode }), stdout: '', stderr: '' };
        let r;
        if (options.fail === request.kind) r = result(request, { failures: [{ phase: 'injected', target: '', code: 5, message: 'failure' }] });
        else if (request.kind === 'list') {
            const projects = options.projects || ['Review'];
            r = result(request, { projects, expectedCount: projects.length, succeededCount: projects.length });
        } else if (request.kind === 'pull' || request.kind === 'check') {
            const remote = options.empty ? [] : options.remote || [{ name: 'C5D_UserModule', type: 1, text: options.remoteSource ?? 'new source' }];
            const items = remote.map((item, index) => ({ name: item.name, type: item.type, file: `m${index}.txt` }));
            remote.forEach((item, index) => fs.writeFileSync(path.join(directory, items[index].file), item.text));
            r = result(request, { items, expectedCount: items.length, succeededCount: items.length });
        } else {
            const applied = request.modules.map(m => m.name);
            const deleted = request.deletions;
            applies.push(...applied); deletes.push(...deleted);
            const total = applied.length + deleted.length;
            r = result(request, { expectedCount: total, succeededCount: total, updatedCount: applied.length, deletedCount: deleted.length });
        }
        fs.writeFileSync(invocation.result, JSON.stringify(r));
        return { stdout: '', stderr: '' };
    };
    const bundledPath = path.join(root, 'bundle', 'C5D_Bridge.catvba');
    fs.mkdirSync(path.dirname(bundledPath)); fs.writeFileSync(bundledPath, 'fake catvba');
    const bridge = new BridgeFiles({ bundledPath, bridgeDir: path.join(root, 'bridge') });
    return { root, steps, applies, deletes, logs, invocations, bridge, events: state.events,
        sync: new CatiaSync({}, output, manager, { execute, timeoutMs: 1000, bridge: options.noBridge ? undefined : bridge }) };
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
test('Pull normalizes CATIA CRLF output to LF with a single trailing newline', async t => {
    const { root, sync, events } = setup(t, { choice: 'keep', remoteSource: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n\r\n' });
    await sync.pull();
    assert.equal(fs.readFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), 'Sub A()\n    x = 1\nEnd Sub\n');
    assert.ok(!events.some(e => e[0] === 'error'));
});
test('Push treats line-ending-only differences as unchanged', async t => {
    for (const [local, remote] of [
        ['Sub A()\n    x = 1\nEnd Sub\n', 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n'],
        ['Sub A()\r\n    x = 1\r\nEnd Sub\r\n', 'Sub A()\n    x = 1\nEnd Sub\n']
    ]) {
        const { sync, steps, events } = setup(t, { localFiles: { 'C5D_UserModule.bas_utf': local }, remoteSource: Buffer.from(remote) });
        await sync.push();
        assert.deepEqual(steps, ['check']);
        assert.ok(events.some(e => e[0] === 'info' && e[1].includes('変更されていません')), JSON.stringify(events));
        assert.ok(!events.some(e => e[0] === 'error'));
    }
});
test('Push still transfers real changes under CRLF and sends the local text unchanged', async t => {
    let sent;
    const { sync, steps, events } = setup(t, {
        localFiles: { 'C5D_UserModule.bas_utf': 'Sub B()\r\n    x = 2\r\nEnd Sub\r\n' },
        remoteSource: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n',
        beforeResult: request => { if (request.kind === 'push') sent = request.modules[0].text; }
    });
    await sync.push();
    assert.deepEqual(steps, ['check', 'push']);
    assert.equal(sent, 'Sub B()\r\n    x = 2\r\nEnd Sub');
    assert.ok(!events.some(e => e[0] === 'error'));
});
const deletionRemote = [{ name: 'A', type: 1, text: 'Sub A()\r\nEnd Sub\r\n' }, { name: 'B', type: 1, text: 'Sub B()\r\nEnd Sub\r\n' }];
const has = (events, kind, text) => events.some(e => e[0] === kind && e[1].includes(text));
test('Push deletion-only: delete removes CATIA-only modules without reporting zero pushes', async t => {
    const { sync, steps, applies, deletes, events, root } = setup(t, { choice: 'delete', localFiles: { 'A.bas_utf': 'Sub A()\nEnd Sub\n' }, remote: deletionRemote });
    await sync.push();
    assert.deepEqual(steps, ['check', 'push']);
    assert.deepEqual(applies, []); assert.deepEqual(deletes, ['B']);
    assert.ok(has(events, 'warning', 'B'), JSON.stringify(events));
    assert.ok(has(events, 'info', '1 個のモジュールをCATIAから削除しました。'), JSON.stringify(events));
    assert.ok(!has(events, 'info', 'プッシュしました'));
    assert.ok(events.some(e => e[0] === 'command' && e[1] === 'cat5dev.refreshTree'));
    assert.ok(!events.some(e => e[0] === 'error'));
    assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')));
});
test('Push deletion-only: keep reports no changes and cancel aborts without pushing', async t => {
    const kept = setup(t, { choice: 'keep', localFiles: { 'A.bas_utf': 'Sub A()\nEnd Sub\n' }, remote: deletionRemote });
    await kept.sync.push();
    assert.deepEqual(kept.steps, ['check']);
    assert.ok(has(kept.events, 'warning', 'B'));
    assert.ok(has(kept.events, 'info', '変更されていません'), JSON.stringify(kept.events));
    assert.ok(!kept.events.some(e => e[0] === 'error'));
    const cancelled = setup(t, { localFiles: { 'A.bas_utf': 'Sub A()\nEnd Sub\n' }, remote: deletionRemote });
    await cancelled.sync.push();
    assert.deepEqual(cancelled.steps, ['check']);
    assert.ok(has(cancelled.events, 'info', 'プッシュを中止しました。'), JSON.stringify(cancelled.events));
    assert.ok(!has(cancelled.events, 'info', '削除しました'));
    assert.ok(!has(cancelled.events, 'info', '変更されていません'));
});
test('Push with only a new UserForm still syncs deletions, or stops when kept', async t => {
    const remote = [{ name: 'B', type: 1, text: 'Sub B()\r\nEnd Sub\r\n' }];
    const deleted = setup(t, { choice: 'delete', localFiles: { 'F.frm_utf': 'form' }, remote });
    await deleted.sync.push();
    assert.deepEqual(deleted.steps, ['check', 'push']);
    assert.deepEqual(deleted.applies, []); assert.deepEqual(deleted.deletes, ['B']);
    assert.ok(has(deleted.events, 'warning', 'UserForm'), JSON.stringify(deleted.events));
    assert.ok(has(deleted.events, 'info', '1 個のモジュールをCATIAから削除しました。'));
    const kept = setup(t, { choice: 'keep', localFiles: { 'F.frm_utf': 'form' }, remote });
    await kept.sync.push();
    assert.deepEqual(kept.steps, ['check']);
    assert.ok(has(kept.events, 'warning', 'UserForm'));
    assert.ok(has(kept.events, 'info', 'プッシュ対象のファイルがなくなった'), JSON.stringify(kept.events));
});
test('Push with changes and deletions applies both and mentions deletion sync', async t => {
    const { sync, steps, applies, deletes, events } = setup(t, { choice: 'delete', localFiles: { 'A.bas_utf': 'Sub A()\n    x = 1\nEnd Sub\n' }, remote: deletionRemote });
    await sync.push();
    assert.deepEqual(steps, ['check', 'push']);
    assert.deepEqual(applies, ['A']); assert.deepEqual(deletes, ['B']);
    assert.ok(has(events, 'info', '1 個のモジュールをプッシュしました。（削除同期を含む）'), JSON.stringify(events));
    assert.ok(!has(events, 'info', 'CATIAから削除しました'));
});
test('Push without changes or deletions reports no changes without confirmation', async t => {
    const { sync, steps, events } = setup(t, { choice: 'delete', localFiles: { 'A.bas_utf': 'Sub A()\nEnd Sub\n' }, remote: [deletionRemote[0]] });
    await sync.push();
    assert.deepEqual(steps, ['check']);
    assert.ok(!events.some(e => e[0] === 'warning'), JSON.stringify(events));
    assert.ok(has(events, 'info', '変更されていません'));
});
test('Push deletion-only failure reports error, no success, and releases lock', async t => {
    const { sync, steps, events, root } = setup(t, { fail: 'push', choice: 'delete', localFiles: { 'A.bas_utf': 'Sub A()\nEnd Sub\n' }, remote: deletionRemote });
    await sync.push();
    assert.deepEqual(steps, ['check', 'push']);
    assert.ok(events.some(e => e[0] === 'error'), JSON.stringify(events));
    assert.ok(!has(events, 'info', '削除しました'));
    assert.ok(!has(events, 'info', 'プッシュしました'));
    assert.ok(!events.some(e => e[0] === 'command' && e[1] === 'cat5dev.refreshTree'));
    assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')));
});
const unsavedWarning = '未保存の変更があります';
const dirtyDoc = fsPath => ({ isDirty: true, uri: { fsPath } });
test('Push with unsaved module files asks first and cancels before contacting CATIA', async t => {
    const { sync, steps, events, root } = setup(t);
    state.documents = [dirtyDoc(path.join(root, 'modules', 'C5D_UserModule.bas_utf'))];
    await sync.push();
    assert.deepEqual(steps, []);
    assert.ok(events.some(e => e[0] === 'warning' && e[1].includes(unsavedWarning) && e[1].includes('C5D_UserModule.bas_utf')), JSON.stringify(events));
    assert.ok(!events.some(e => e[0] === 'warning' && e[1].includes('LocalOnly.cls_utf')));
    assert.ok(has(events, 'info', 'プッシュを中止しました。'), JSON.stringify(events));
    assert.ok(!events.some(e => e[0] === 'error'));
    assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')));
});
test('Push with unsaved module files continues with the content saved on disk', async t => {
    let sent;
    const { sync, steps, events, root } = setup(t, {
        choice: 'continue',
        beforeResult: request => { if (request.kind === 'push') sent = request.modules[0].text; }
    });
    state.documents = [dirtyDoc(path.join(root, 'modules', 'C5D_UserModule.bas_utf'))];
    await sync.push();
    assert.ok(has(events, 'warning', unsavedWarning), JSON.stringify(events));
    assert.deepEqual(steps, ['check', 'push']);
    assert.equal(sent, 'old');
    assert.ok(!has(events, 'info', 'プッシュを中止しました。'));
    assert.ok(!events.some(e => e[0] === 'error'));
});
test('Push without unsaved module files shows no unsaved warning', async t => {
    const { sync, steps, events } = setup(t, { choice: 'continue' });
    await sync.push();
    assert.ok(!has(events, 'warning', unsavedWarning), JSON.stringify(events));
    assert.deepEqual(steps, ['check', 'push']);
});
test('Push ignores dirty documents that are not Push targets', async t => {
    const { sync, steps, events, root } = setup(t, { choice: 'continue' });
    state.documents = [
        dirtyDoc(path.join(root, 'cat5dev.toml')),
        dirtyDoc('Untitled-1'),
        dirtyDoc(path.join(root, 'modules', 'x.bas'))
    ];
    await sync.push();
    assert.ok(!has(events, 'warning', unsavedWarning), JSON.stringify(events));
    assert.deepEqual(steps, ['check', 'push']);
});
test('Push detects unsaved module files regardless of path case', async t => {
    const { sync, steps, events, root } = setup(t);
    state.documents = [dirtyDoc(path.join(root, 'modules', 'C5D_UserModule.bas_utf').toUpperCase())];
    await sync.push();
    assert.ok(has(events, 'warning', unsavedWarning), JSON.stringify(events));
    assert.deepEqual(steps, []);
});
test('Push unsaved warning comes before project selection', async t => {
    const { sync, steps, events, root } = setup(t, { noTarget: true, selection: 'Review' });
    state.documents = [dirtyDoc(path.join(root, 'modules', 'C5D_UserModule.bas_utf'))];
    await sync.push();
    assert.deepEqual(steps, []);
    assert.ok(has(events, 'warning', unsavedWarning), JSON.stringify(events));
    assert.ok(has(events, 'info', 'プッシュを中止しました。'));
});
test('Push strips a UTF-8 BOM before comparing and sending', async t => {
    const same = setup(t, { localFiles: { 'C5D_UserModule.bas_utf': '﻿Sub A()\r\n    x = 1\r\nEnd Sub\r\n' }, remoteSource: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n' });
    await same.sync.push();
    assert.deepEqual(same.steps, ['check']);
    assert.ok(has(same.events, 'info', '変更されていません'), JSON.stringify(same.events));
    let sent;
    const changed = setup(t, {
        localFiles: { 'C5D_UserModule.bas_utf': '﻿Sub B()\r\n    x = 2\r\nEnd Sub\r\n' },
        remoteSource: 'Sub A()\r\n    x = 1\r\nEnd Sub\r\n',
        beforeResult: request => { if (request.kind === 'push') sent = request.modules[0].text; }
    });
    await changed.sync.push();
    assert.deepEqual(changed.steps, ['check', 'push']);
    assert.equal(sent, 'Sub B()\r\n    x = 2\r\nEnd Sub');
    assert.ok(!changed.events.some(e => e[0] === 'error'), JSON.stringify(changed.events));
});
const settles = (promise, label) => Promise.race([promise, new Promise((_, reject) => setTimeout(() => reject(new Error(`${label} did not finish while notifications stay open`)), 2000).unref())]);
test('Pull/Push finish and release the lock without waiting for notifications to be dismissed', async t => {
    const cases = [
        { label: 'pull success', op: 'pull', options: { choice: 'keep' }, expect: ['info', '正常にプル'] },
        { label: 'pull deletion cancelled', op: 'pull', options: {}, expect: ['info', 'プル'] },
        { label: 'pull failure', op: 'pull', options: { fail: 'pull' }, expect: ['error', ''] },
        { label: 'pull local change conflict', op: 'pull', options: { choice: 'keep', beforeResult: request => {
            if (request.kind === 'pull') fs.writeFileSync(path.join(state.root, 'modules', 'C5D_UserModule.bas_utf'), 'edited');
        } }, expect: ['warning', 'ローカル変更'] },
        { label: 'push success', op: 'push', options: { choice: 'keep' }, expect: ['info', 'プッシュしました'] },
        { label: 'push no changes', op: 'push', options: { localFiles: { 'A.bas_utf': 'Sub A()\nEnd Sub\n' }, remote: [deletionRemote[0]] }, expect: ['info', '変更されていません'] },
        { label: 'push failure', op: 'push', options: { fail: 'check' }, expect: ['error', ''] },
    ];
    for (const c of cases) {
        const { sync, events, root } = setup(t, { ...c.options, hangNotifications: true });
        await settles(sync[c.op](), c.label);
        assert.ok(has(events, ...c.expect), `${c.label}: ${JSON.stringify(events)}`);
        assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')), c.label);
    }
});
test('format on pull follows [formatter] format_on_pull, falling back to the legacy VSCode setting only when the key is absent', async t => {
    const cases = [
        { formatter: '', legacy: true, expected: false },
        { formatter: '[formatter]\nenabled = false\nformat_on_pull = true\n', legacy: true, expected: false },
        { formatter: '[formatter]\nenabled = true\nformat_on_pull = true\n', legacy: false, expected: true },
        { formatter: '[formatter]\nenabled = true\nformat_on_pull = false\n', legacy: true, expected: false },
        { formatter: '[formatter]\nenabled = true\n', legacy: true, expected: true },
        { formatter: '[formatter]\nenabled = true\n', legacy: undefined, expected: false },
    ];
    for (const c of cases) {
        const { root, sync } = setup(t);
        fs.appendFileSync(path.join(root, 'cat5dev.toml'), c.formatter);
        state.config = { 'cat5dev.formatter.formatOnPull': c.legacy };
        assert.equal(sync.formatOnPull(root), c.expected, JSON.stringify(c));
    }
});
test('Pull and check read m<i>.txt as UTF-8; the cat5dev.toml encoding is ignored', async t => {
    const source = 'Sub A()\r\n    MsgBox "日本語 🙂 中文"\r\nEnd Sub\r\n';
    const pulled = setup(t, { choice: 'keep', remoteSource: Buffer.from(source, 'utf8') });
    await pulled.sync.pull();
    assert.equal(fs.readFileSync(path.join(pulled.root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), 'Sub A()\n    MsgBox "日本語 🙂 中文"\nEnd Sub\n');
    assert.ok(!pulled.events.some(e => e[0] === 'error'), JSON.stringify(pulled.events));
    const unchanged = setup(t, { localFiles: { 'C5D_UserModule.bas_utf': 'Sub A()\n    MsgBox "日本語 🙂 中文"\nEnd Sub\n' }, remoteSource: Buffer.from(source, 'utf8') });
    await unchanged.sync.push();
    assert.deepEqual(unchanged.steps, ['check']);
    assert.ok(has(unchanged.events, 'info', '変更されていません'), JSON.stringify(unchanged.events));
});
test('Push sends the local text unchanged in request.json, including characters outside Shift_JIS', async t => {
    let request;
    const { sync, steps, events } = setup(t, {
        localFiles: { 'C5D_UserModule.bas_utf': 'Sub B()\r\n    s = "🙂 中文"\r\nEnd Sub\r\n' },
        beforeResult: r => { if (r.kind === 'push') request = { ...r, files: fs.readdirSync(r.directory) }; }
    });
    await sync.push();
    assert.deepEqual(steps, ['check', 'push']);
    assert.deepEqual(request.modules, [{ name: 'C5D_UserModule', type: 1, text: 'Sub B()\r\n    s = "🙂 中文"\r\nEnd Sub' }]);
    assert.deepEqual(request.deletions, []);
    assert.equal(request.version, 2); assert.ok(!('encoding' in request));
    assert.deepEqual(request.files, ['request.json'], 'Push writes only request.json (no m<i>.txt)');
    assert.ok(!events.some(e => e[0] === 'error'), JSON.stringify(events));
});
test('Select never offers the C5D_Bridge project', async t => {
    let offered;
    const { sync, events } = setup(t, { noTarget: true, projects: ['C5D_Bridge', 'Review'] });
    vscode.window.showQuickPick = async items => { offered = items; return 'Review'; };
    try { await sync.selectProject(); } finally { vscode.window.showQuickPick = async () => state.selection; }
    assert.deepEqual(offered, ['Review']);
    const only = setup(t, { noTarget: true, projects: ['C5D_Bridge'] });
    await only.sync.selectProject();
    assert.ok(has(only.events, 'info', 'VBAプロジェクトが見つかりませんでした'), JSON.stringify(only.events));
    assert.ok(!events.some(e => e[0] === 'error'));
});
test('exit codes 10, 13 and 14 show dedicated messages and release the lock; 2 and 11 are general errors', async t => {
    const cases = [[10, 'CATIA が起動していません'], [13, '仲介プロジェクト'], [14, 'VBA エディタ'], [2, 'プルに失敗しました'], [11, 'プルに失敗しました']];
    for (const [exitCode, text] of cases) {
        const { sync, events, root, steps } = setup(t, { exitCode });
        await sync.pull();
        assert.deepEqual(steps, ['pull'], String(exitCode));
        assert.ok(has(events, 'error', text), `${exitCode}: ${JSON.stringify(events)}`);
        assert.equal(events.filter(e => e[0] === 'error').length, 1, String(exitCode));
        assert.ok(!fs.existsSync(path.join(root, 'locks', 'active')), String(exitCode));
        assert.equal(fs.readFileSync(path.join(root, 'modules', 'C5D_UserModule.bas_utf'), 'utf8'), 'old');
    }
});
test('the bridge master is prepared once per operation, read-only, and passed as --bridge', async t => {
    const { sync, invocations, bridge, steps, logs } = setup(t, { noTarget: true, selection: 'Review', choice: 'keep' });
    await sync.pull();
    assert.deepEqual(steps, ['list', 'pull']);
    const master = bridge.masterPath();
    assert.match(path.basename(master), /^C5D_Bridge-[0-9a-f]{8}\.catvba$/);
    assert.deepEqual(invocations.map(i => i.bridge), [master, master]);
    assert.equal(fs.readFileSync(master, 'utf8'), 'fake catvba');
    assert.equal(fs.statSync(master).mode & 0o200, 0, 'read-only');
    assert.equal(logs.filter(line => line.startsWith('[bridge] master')).length, 1, JSON.stringify(logs));
    for (const i of invocations) {
        assert.equal(path.dirname(i.request), path.dirname(i.result));
        assert.equal(path.basename(i.request), 'request.json'); assert.equal(path.basename(i.result), 'result.json');
    }
});
test('a bridge that cannot be prepared stops before catia-run.exe with a general error', async t => {
    const missing = setup(t, { choice: 'keep' });
    fs.unlinkSync(path.join(missing.root, 'bundle', 'C5D_Bridge.catvba'));
    await missing.sync.pull();
    assert.deepEqual(missing.steps, []);
    assert.ok(has(missing.events, 'error', 'プルに失敗しました'), JSON.stringify(missing.events));
    assert.ok(!fs.existsSync(path.join(missing.root, 'locks', 'active')));
    const unconfigured = setup(t, { choice: 'keep', noBridge: true });
    await unconfigured.sync.pull();
    assert.deepEqual(unconfigured.steps, []);
    assert.ok(has(unconfigured.events, 'error', 'プルに失敗しました'));
});
