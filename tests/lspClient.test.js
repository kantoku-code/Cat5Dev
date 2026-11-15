const { test } = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');

// vscode と vscode-languageclient/node を差し替えて、クライアント設定とファイル監視の登録だけを確かめる
const watchers = [];
const vscode = {
    workspace: {
        createFileSystemWatcher(glob) {
            const w = { glob, dispose() {} };
            watchers.push(w);
            return w;
        },
    },
};
const clients = [];
class LanguageClient {
    constructor(id, name, serverOptions, clientOptions) {
        Object.assign(this, { id, name, serverOptions, clientOptions });
        clients.push(this);
    }
    start() { return Promise.resolve(); }
    stop() { return Promise.resolve(); }
}
const languageclient = { LanguageClient, TransportKind: { stdio: 0 } };

const originalLoad = Module._load;
Module._load = function (id, ...args) {
    if (id === 'vscode') { return vscode; }
    if (id === 'vscode-languageclient/node') { return languageclient; }
    return originalLoad.call(this, id, ...args);
};
const lsp = require('../out/lspClient');
Module._load = originalLoad;

test('client options forward file events and only accept file: documents', () => {
    const watcher = { glob: 'x' };
    const opts = lsp.buildClientOptions(watcher);
    assert.equal(opts.synchronize.fileEvents, watcher);
    assert.deepEqual(opts.documentSelector, [
        { scheme: 'file', language: 'vb', pattern: '**/*.bas_utf' },
        { scheme: 'file', language: 'vb', pattern: '**/*.cls_utf' },
        { scheme: 'file', language: 'vb', pattern: '**/*.frm_utf' },
    ]);
});

test('startLspClient registers a module watcher and passes it to the client', async () => {
    const saved = process.env.CAT5DEV_VBA_LS;
    process.env.CAT5DEV_VBA_LS = process.execPath; // 存在する任意のファイル
    try {
        const ctx = { extensionPath: __dirname, subscriptions: [] };
        lsp.startLspClient(ctx);
        assert.equal(watchers.length, 1);
        assert.equal(watchers[0].glob, '**/*.{bas_utf,cls_utf,frm_utf}');
        assert.ok(ctx.subscriptions.includes(watchers[0]));
        assert.equal(clients.length, 1);
        assert.equal(clients[0].clientOptions.synchronize.fileEvents, watchers[0]);
        await lsp.stopLspClient();
    } finally {
        if (saved === undefined) {
            delete process.env.CAT5DEV_VBA_LS;
        } else {
            process.env.CAT5DEV_VBA_LS = saved;
        }
    }
});
