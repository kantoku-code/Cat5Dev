const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const { fixture } = require('./helpers');
let state;
const handlers = new Map();
class WorkspaceEdit {
    constructor() { this.renames = []; }
    renameFile(oldUri, newUri, options) { this.renames.push({ oldUri, newUri, options }); }
}
const vscode = {
    env: { language: 'ja', clipboard: { writeText: async text => { state.clipboard = text; } } },
    Uri: { file: fsPath => ({ scheme: 'file', fsPath: path.resolve(fsPath) }) },
    WorkspaceEdit,
    workspace: {
        get workspaceFolders() { return [{ uri: { fsPath: state.root } }]; },
        get textDocuments() { return state.documents || []; },
        getConfiguration() { return { get: (_, fallback) => fallback, inspect: () => ({}) }; },
        fs: {
            delete: async (uri, options) => {
                state.events.push(['delete', uri.fsPath, options]);
                if (state.deleteError) throw state.deleteError;
                fs.unlinkSync(uri.fsPath);
            }
        },
        applyEdit: async edit => {
            state.events.push(['applyEdit', edit.renames.length]);
            if (state.applyEditResult !== undefined) return state.applyEditResult;
            for (const { oldUri, newUri, options } of edit.renames) {
                const caseOnly = oldUri.fsPath.toLowerCase() === newUri.fsPath.toLowerCase();
                if (!options?.overwrite && !caseOnly && fs.existsSync(newUri.fsPath)) return false;
                fs.renameSync(oldUri.fsPath, newUri.fsPath);
            }
            return true;
        }
    },
    commands: {
        registerCommand: (id, handler) => { handlers.set(id, handler); return { dispose() { } }; },
        executeCommand: async command => { state.events.push(['command', command]); }
    },
    window: {
        showInformationMessage: async message => { state.events.push(['info', message]); },
        showErrorMessage: async message => { state.events.push(['error', message]); },
        showWarningMessage: async (message, options, ...actions) => {
            state.events.push(['warning', message]);
            if (state.onWarning) state.onWarning(message);
            if (!options?.modal) return undefined;
            return state.confirm ? actions[0] : undefined;
        },
        showInputBox: async options => {
            state.inputBox = options;
            state.events.push(['input', options.value]);
            return typeof state.input === 'function' ? state.input(options) : state.input;
        }
    }
};
const originalLoad = Module._load;
Module._load = function (id, ...args) { return id === 'vscode' ? vscode : originalLoad.call(this, id, ...args); };
const { registerFileCommands, resolveModulePath, validateModuleName } = require('../out/fileCommands');
Module._load = originalLoad;
registerFileCommands({ subscriptions: [] });

function setup(t, options = {}) {
    const root = fixture(t);
    const modules = path.join(root, 'modules');
    fs.mkdirSync(modules);
    fs.writeFileSync(path.join(root, 'cat5dev.toml'), '[project]\ntarget_project = "Review"\nlanguage = "ja"\nencoding = "shift_jis"\n');
    const files = options.files || { 'Main.bas_utf': 'main', 'Helper.cls_utf': 'helper', 'Form1.frm_utf': 'form', 'notes.txt': 'txt' };
    for (const [name, content] of Object.entries(files)) fs.writeFileSync(path.join(modules, name), content);
    state = { root, events: [], ...options };
    const file = name => path.join(modules, name);
    const node = name => ({ label: name, nodeType: 'module', contextValue: 'catiaModule', filePath: file(name) });
    const uri = name => vscode.Uri.file(file(name));
    return { root, modules, file, node, uri, events: state.events, list: () => fs.readdirSync(modules).sort() };
}
const run = (id, arg) => handlers.get(id)(arg);
const errors = () => state.events.filter(e => e[0] === 'error');
const refreshed = () => state.events.some(e => e[0] === 'command' && e[1] === 'cat5dev.refreshTree');

test('commands are registered', () => {
    assert.deepEqual([...handlers.keys()].sort(), ['cat5dev.copyPath', 'cat5dev.deleteFile', 'cat5dev.renameFile']);
});

test('resolveModulePath accepts tree nodes and file Uris directly under modules', t => {
    const { file, node, uri, root } = setup(t);
    assert.equal(resolveModulePath(node('Main.bas_utf')), file('Main.bas_utf'));
    assert.equal(resolveModulePath(uri('Helper.cls_utf')), file('Helper.cls_utf'));
    assert.equal(resolveModulePath({ filePath: file('Form1.frm_utf') }), file('Form1.frm_utf'));
    assert.equal(errors().length, 0);

    assert.equal(resolveModulePath(undefined), undefined);
    assert.ok(errors().at(-1)[1].includes('対象のモジュールが指定されていません'));
    assert.equal(resolveModulePath({ scheme: 'untitled', fsPath: file('Main.bas_utf') }), undefined);
    assert.ok(errors().at(-1)[1].includes('対象のモジュールが指定されていません'));
    assert.equal(resolveModulePath({ label: 'project', nodeType: 'project' }), undefined);

    fs.writeFileSync(path.join(root, 'Outside.bas_utf'), 'x');
    fs.mkdirSync(path.join(root, 'modules', 'sub'));
    fs.writeFileSync(path.join(root, 'modules', 'sub', 'Nested.bas_utf'), 'x');
    fs.mkdirSync(path.join(root, 'modules', 'Dir.bas_utf'));
    for (const bad of [
        path.join(root, 'Outside.bas_utf'),
        path.join(root, 'modules', 'sub', 'Nested.bas_utf'),
        path.join(root, 'modules', '..', 'Outside.bas_utf'),
        file('notes.txt'),
        file('Missing.bas_utf'),
        file('Dir.bas_utf')
    ]) {
        const before = errors().length;
        assert.equal(resolveModulePath({ filePath: bad }), undefined, bad);
        assert.equal(errors().length, before + 1);
        assert.ok(errors().at(-1)[1].includes('modules フォルダ内のVBAファイルではありません'));
    }
});

test('copyPath copies the module path for nodes and Uris, and reports a missing target', async t => {
    const { file, node, uri } = setup(t);
    await run('cat5dev.copyPath', node('Main.bas_utf'));
    assert.equal(state.clipboard, file('Main.bas_utf'));
    assert.ok(state.events.some(e => e[0] === 'info' && e[1].includes(file('Main.bas_utf'))));
    await run('cat5dev.copyPath', uri('Helper.cls_utf'));
    assert.equal(state.clipboard, file('Helper.cls_utf'));
    state.clipboard = undefined;
    await run('cat5dev.copyPath', undefined);
    assert.equal(state.clipboard, undefined);
    assert.equal(errors().length, 1);
    assert.ok(errors()[0][1].includes('ツリーから右クリック'));
});

test('delete removes the file to trash after confirmation and refreshes the tree', async t => {
    for (const target of ['node', 'uri']) {
        const ctx = setup(t, { confirm: true });
        await run('cat5dev.deleteFile', ctx[target]('Main.bas_utf'));
        assert.equal(fs.existsSync(ctx.file('Main.bas_utf')), false);
        const del = ctx.events.find(e => e[0] === 'delete');
        assert.deepEqual(del[2], { useTrash: true });
        assert.ok(ctx.events.some(e => e[0] === 'warning' && e[1].includes('Main.bas_utf を削除しますか')));
        assert.ok(refreshed());
        assert.equal(errors().length, 0);
    }
});

test('delete cancellation keeps the file', async t => {
    const { node, file, events } = setup(t, { confirm: false });
    await run('cat5dev.deleteFile', node('Main.bas_utf'));
    assert.equal(fs.readFileSync(file('Main.bas_utf'), 'utf8'), 'main');
    assert.ok(!events.some(e => e[0] === 'delete'));
});

test('delete rejects missing and invalid targets without touching files', async t => {
    const { root, file, list } = setup(t, { confirm: true });
    fs.writeFileSync(path.join(root, 'Outside.bas_utf'), 'x');
    const before = list();
    for (const arg of [undefined, { filePath: path.join(root, 'Outside.bas_utf') }, { filePath: file('notes.txt') }, vscode.Uri.file(file('notes.txt'))]) {
        const count = errors().length;
        await run('cat5dev.deleteFile', arg);
        assert.equal(errors().length, count + 1);
    }
    assert.deepEqual(list(), before);
    assert.ok(fs.existsSync(path.join(root, 'Outside.bas_utf')));
    assert.ok(!state.events.some(e => e[0] === 'delete' || e[0] === 'warning'));
});

test('delete of a file that disappeared before deletion only refreshes', async t => {
    const { node, file } = setup(t, { confirm: true, onWarning: () => fs.unlinkSync(file('Main.bas_utf')) });
    await run('cat5dev.deleteFile', node('Main.bas_utf'));
    assert.equal(errors().length, 0);
    assert.ok(refreshed());
});

test('delete reports other failures and keeps the file', async t => {
    const err = Object.assign(new Error('access denied'), { code: 'EPERM' });
    const { node, file } = setup(t, { confirm: true, deleteError: err });
    await run('cat5dev.deleteFile', node('Main.bas_utf'));
    assert.ok(fs.existsSync(file('Main.bas_utf')));
    assert.equal(errors().length, 1);
    assert.ok(errors()[0][1].includes('削除に失敗しました (Main.bas_utf): access denied'));
});

test('delete is refused when the module has unsaved changes', async t => {
    const { node, file, events } = setup(t, { confirm: true });
    state.documents = [{ isDirty: true, uri: { fsPath: file('Main.bas_utf').toUpperCase() } }];
    await run('cat5dev.deleteFile', node('Main.bas_utf'));
    assert.ok(fs.existsSync(file('Main.bas_utf')));
    assert.ok(events.some(e => e[0] === 'warning' && e[1].includes('未保存の変更')));
    assert.ok(!events.some(e => e[0] === 'delete'));
});

test('rename shows the module name without extension and keeps the extension', async t => {
    for (const [target, name, ext] of [['node', 'Main', '.bas_utf'], ['uri', 'Helper', '.cls_utf']]) {
        const ctx = setup(t, { input: 'Renamed' });
        await run('cat5dev.renameFile', ctx[target](name + ext));
        assert.equal(state.inputBox.value, name);
        assert.equal(fs.existsSync(ctx.file(name + ext)), false);
        assert.equal(fs.readFileSync(ctx.file('Renamed' + ext), 'utf8'), name === 'Main' ? 'main' : 'helper');
        assert.ok(refreshed());
        assert.equal(errors().length, 0);
    }
});

test('rename rejects names of other existing modules and leaves both files unchanged', async t => {
    for (const input of ['helper', 'HELPER', 'Form1', 'form1']) {
        const { node, file, list } = setup(t, { input });
        const before = list();
        assert.ok(validateModuleName(input, file('Main.bas_utf')).includes('同名のモジュールが既に存在します'));
        await run('cat5dev.renameFile', node('Main.bas_utf'));
        assert.ok(state.inputBox.validateInput(input).includes('同名のモジュールが既に存在します'));
        assert.deepEqual(list(), before);
        assert.equal(fs.readFileSync(file('Main.bas_utf'), 'utf8'), 'main');
        assert.equal(fs.readFileSync(file('Helper.cls_utf'), 'utf8'), 'helper');
        assert.ok(!state.events.some(e => e[0] === 'applyEdit'));
    }
});

test('rename allows a case-only change', async t => {
    const { node, list } = setup(t, { input: 'MAIN' });
    await run('cat5dev.renameFile', node('Main.bas_utf'));
    assert.equal(state.inputBox.validateInput('MAIN'), undefined);
    assert.equal(state.inputBox.validateInput('Main'), undefined);
    assert.ok(list().includes('MAIN.bas_utf'));
    assert.ok(!list().includes('Main.bas_utf'));
    assert.equal(errors().length, 0);
});

test('rename rejects invalid module names', async t => {
    for (const input of ['..\\x', 'a/b', 'CON', '1abc', '', 'a b', 'a.b', 'nul']) {
        const { node, list } = setup(t, { input });
        const before = list();
        await run('cat5dev.renameFile', node('Main.bas_utf'));
        assert.ok(state.inputBox.validateInput(input), `validateInput should reject ${JSON.stringify(input)}`);
        assert.deepEqual(list(), before);
        assert.ok(!state.events.some(e => e[0] === 'applyEdit'));
    }
    const { file } = setup(t);
    assert.equal(validateModuleName('Valid_Name1', file('Main.bas_utf')), undefined);
    assert.equal(validateModuleName('日本語モジュール', file('Main.bas_utf')), undefined);
});

test('rename is refused when the module has unsaved changes', async t => {
    const { node, file, list } = setup(t, { input: 'Renamed' });
    state.documents = [{ isDirty: true, uri: { fsPath: file('Main.bas_utf') } }];
    const before = list();
    await run('cat5dev.renameFile', node('Main.bas_utf'));
    assert.deepEqual(list(), before);
    assert.ok(state.events.some(e => e[0] === 'warning' && e[1].includes('未保存の変更')));
});

test('UserForm rename warns and cancellation keeps the form', async t => {
    const { node, list, events } = setup(t, { input: 'Form2', confirm: false });
    const before = list();
    await run('cat5dev.renameFile', node('Form1.frm_utf'));
    assert.ok(events.some(e => e[0] === 'warning' && e[1].includes('UserForm')));
    assert.deepEqual(list(), before);
    assert.ok(!events.some(e => e[0] === 'applyEdit'));
});

test('UserForm rename proceeds after confirmation', async t => {
    const { node, file } = setup(t, { input: 'Form2', confirm: true });
    await run('cat5dev.renameFile', node('Form1.frm_utf'));
    assert.ok(fs.existsSync(file('Form2.frm_utf')));
    assert.ok(!fs.existsSync(file('Form1.frm_utf')));
});

test('InputBox cancellation and unchanged names do nothing', async t => {
    for (const input of [undefined, 'Main']) {
        const { node, list, events } = setup(t, { input });
        const before = list();
        await run('cat5dev.renameFile', node('Main.bas_utf'));
        assert.deepEqual(list(), before);
        assert.ok(!events.some(e => e[0] === 'applyEdit'));
        assert.equal(errors().length, 0);
    }
});

test('rename reports rejected workspace edits and missing targets', async t => {
    const { node, list } = setup(t, { input: 'Renamed', applyEditResult: false });
    const before = list();
    await run('cat5dev.renameFile', node('Main.bas_utf'));
    assert.deepEqual(list(), before);
    assert.ok(errors().some(e => e[1].includes('名前変更に失敗しました')));
    const count = errors().length;
    await run('cat5dev.renameFile', undefined);
    assert.equal(errors().length, count + 1);
});
