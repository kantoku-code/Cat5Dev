// Run only in an isolated VSCode Extension Development Host.
const vscode = require('vscode');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
exports.run = async () => {
    const root = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    assert.ok(root && path.basename(root) === 'workspace', 'isolated workspace required');
    const source = 'Sub Test()\nDim value_ As Long\nDebug.Print value_\nMsgBox Prompt:="hello"\nEnd Sub\n';
    const document = await vscode.workspace.openTextDocument(vscode.Uri.file(path.join(root, 'modules', 'Smoke.bas_utf')));
    const extensionRoot = process.env.CAT5DEV_SMOKE_EXTENSION_ROOT || path.resolve(__dirname, '..');
    const extension = vscode.extensions.all.find(e => path.resolve(e.extensionPath).toLowerCase() === path.resolve(extensionRoot).toLowerCase());
    assert.ok(extension, 'development extension loaded'); await extension.activate();
    const commands = await vscode.commands.getCommands();
    assert.ok(commands.includes('cat5dev.recoverSyncLock'));
    let edits = [];
    for (let attempt = 0; attempt < 40; attempt++) {
        edits = await vscode.commands.executeCommand('vscode.executeFormatDocumentProvider', document.uri, { insertSpaces: true, tabSize: 4 }) || [];
        if (edits.length) break;
        await new Promise(resolve => setTimeout(resolve, 250));
    }
    assert.ok(edits.length, 'formatter server became ready');
    const edit = new vscode.WorkspaceEdit(); edit.set(document.uri, edits); assert.ok(await vscode.workspace.applyEdit(edit));
    const formatted = document.getText();
    assert.ok(formatted.includes('Debug.Print value_'));
    assert.ok(formatted.includes('Prompt:="hello"'));
    assert.ok(!formatted.includes('value _') && !formatted.includes(': ='));
    assert.match(formatted, /\r?\nEnd Sub\r?\n$/);
    assert.notEqual(formatted, source);
    fs.writeFileSync(path.join(root, '..', 'smoke-result.json'), JSON.stringify({ passed: true, extension: extension.id, formatted }));
    console.log('CAT5DEV_VSCODE_SMOKE_PASSED');
};
