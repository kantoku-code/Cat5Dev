// VbaLinter.lintNow (src/linter.ts) with a fake vscode and a fake /lint: a result that arrives after the document changed,
// after a newer lint, or after the document was closed must not overwrite the diagnostics.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const { fixture } = require('./helpers');

let root;
const collections = [];
class Range { constructor(sl, sc, el, ec) { Object.assign(this, { sl, sc, el, ec }); } }
class Diagnostic { constructor(range, message, severity) { Object.assign(this, { range, message, severity }); } }
const vscode = {
    env: { language: 'en' },
    Range, Diagnostic, DiagnosticSeverity: { Error: 0, Warning: 1 },
    workspace: {
        get workspaceFolders() { return [{ uri: { fsPath: root } }]; },
        textDocuments: [],
        getConfiguration() { return { get: (_key, fallback) => fallback, inspect: () => ({}) }; }
    },
    languages: {
        createDiagnosticCollection() {
            const map = new Map();
            const collection = {
                map,
                get: uri => map.get(uri.toString()),
                set: (uri, diagnostics) => { map.set(uri.toString(), diagnostics); },
                delete: uri => { map.delete(uri.toString()); },
                dispose: () => map.clear()
            };
            collections.push(collection);
            return collection;
        }
    }
};
const originalLoad = Module._load;
Module._load = function (id, ...args) { return id === 'vscode' ? vscode : originalLoad.call(this, id, ...args); };
const vbaServer = require('../out/vbaServer');
const { VbaLinter } = require('../out/linter');
Module._load = originalLoad;

/** A fake /lint: every request waits until the test answers it (a diagnostic with that message, or a failure). */
function fakeLint(t) {
    const pending = [];
    const original = vbaServer.httpPost;
    vbaServer.httpPost = () => new Promise((resolve, reject) => pending.push({ resolve, reject }));
    t.after(() => { vbaServer.httpPost = original; });
    const answer = (index, message) => pending[index].resolve(JSON.stringify({ error: '',
        diagnostics: [{ line: 0, col: 0, end_line: 0, end_col: 1, severity: 'warning', code: 'VBA001', message }] }));
    const fail = index => pending[index].reject(new Error('lint server down'));
    return { pending, answer, fail };
}
/** texts: the text of each version (the last one is used for later versions). */
function document(texts = ['Sub A()\nEnd Sub']) {
    const uri = { toString: () => 'file:///c%3A/w/modules/Module1.bas_utf' };
    const doc = { uri, version: 1, isClosed: false, lineCount: 2, lineAt: () => ({ text: 'Sub A()' }) };
    doc.getText = () => texts[Math.min(doc.version, texts.length) - 1];
    return doc;
}
function setup(t, { nonAnsi = false } = {}) {
    root = fixture(t);
    // VBA010 is off (unless asked for, then with code page 932) so that the fake /lint decides the diagnostics.
    fs.writeFileSync(path.join(root, 'cat5dev.toml'), '[project]\nlanguage = "en"\n\n[lint]\nenabled = true\n\n[lint.rules]\n'
        + (nonAnsi ? 'non_ansi_character = true\nansi_code_page = 932\n' : 'non_ansi_character = false\n'));
    const lint = fakeLint(t);
    const output = { appendLine: () => {} };
    const linter = new VbaLinter({ getBaseUrl: () => 'http://127.0.0.1:1' }, output);
    t.after(() => linter.dispose());
    const collection = collections[collections.length - 1];
    return { lint, linter, collection };
}
const until = async (condition) => { for (let i = 0; i < 100 && !condition(); i++) await new Promise(resolve => setImmediate(resolve)); };
const messages = (collection, doc) => (collection.get(doc.uri) ?? []).map(d => d.message);

test('a slower lint of an older version does not overwrite the newer result', async t => {
    const { lint, linter, collection } = setup(t);
    const doc = document();
    const first = linter.lintNow(doc);
    await until(() => lint.pending.length === 1);
    doc.version = 2;
    const second = linter.lintNow(doc);
    await until(() => lint.pending.length === 2);
    lint.answer(1, 'new');
    await second;
    assert.deepEqual(messages(collection, doc), ['new']);
    lint.answer(0, 'old');
    await first;
    assert.deepEqual(messages(collection, doc), ['new'], 'the late result of version 1 is dropped');
});

test('of two lints of the same version, the one started last wins even when it finishes first', async t => {
    const { lint, linter, collection } = setup(t);
    const doc = document();
    const first = linter.lintNow(doc);
    await until(() => lint.pending.length === 1);
    const second = linter.lintNow(doc);
    await until(() => lint.pending.length === 2);
    lint.answer(1, 'latest');
    await second;
    lint.answer(0, 'stale');
    await first;
    assert.deepEqual(messages(collection, doc), ['latest']);
});

test('a lint that finishes after the document changed or closed sets nothing', async t => {
    const { lint, linter, collection } = setup(t);
    const edited = document();
    const pendingEdit = linter.lintNow(edited);
    await until(() => lint.pending.length === 1);
    edited.version = 2; // edited; the debounced lint has not started yet
    lint.answer(0, 'for version 1');
    await pendingEdit;
    assert.equal(collection.get(edited.uri), undefined);

    const closed = document();
    const pendingClose = linter.lintNow(closed);
    await until(() => lint.pending.length === 2);
    closed.isClosed = true;
    linter.clearDocument(closed.uri);
    lint.answer(1, 'after close');
    await pendingClose;
    assert.equal(collection.get(closed.uri), undefined, 'a closed document gets no diagnostics back');
});

test('a failed Go lint keeps the previous Go diagnostics only when no newer lint has started', async t => {
    const { lint, linter, collection } = setup(t, { nonAnsi: true });
    // Version 1 has a character outside CP932 (VBA010); version 2 has none.
    const doc = document(['Sub A() \u{1F642}\nEnd Sub', 'Sub A()\nEnd Sub']);
    const first = linter.lintNow(doc);
    await until(() => lint.pending.length === 1);
    lint.answer(0, 'kept');
    await first;
    const failing = linter.lintNow(doc);
    await until(() => lint.pending.length === 2);
    lint.fail(1);
    await failing;
    assert.deepEqual(collection.get(doc.uri).map(d => d.code), ['VBA001', 'VBA010'], 'the previous Go diagnostics stay, VBA010 is updated');
    assert.equal(collection.get(doc.uri)[0].message, 'kept');

    const stale = linter.lintNow(doc);
    await until(() => lint.pending.length === 3);
    doc.version = 2;
    const fresh = linter.lintNow(doc);
    await until(() => lint.pending.length === 4);
    lint.answer(3, 'fresh');
    await fresh;
    lint.fail(2);
    await stale;
    assert.deepEqual(messages(collection, doc), ['fresh'], 'the stale failure does not write back the old VBA010 with what it read');
});
