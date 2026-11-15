const { test } = require('node:test');
const assert = require('node:assert/strict');
const Module = require('node:module');

class Range {
    constructor(startLine, startCharacter, endLine, endCharacter) {
        Object.assign(this, { startLine, startCharacter, endLine, endCharacter });
    }
}
class DocumentSymbol {
    constructor(name, detail, kind, range, selectionRange) {
        Object.assign(this, { name, detail, kind, range, selectionRange });
    }
}
const vscode = { Range, DocumentSymbol, SymbolKind: { Function: 11, Property: 6, Struct: 22, Enum: 9 } };
const originalLoad = Module._load;
Module._load = function (id, ...args) { return id === 'vscode' ? vscode : originalLoad.call(this, id, ...args); };
const { VbaDocumentSymbolProvider } = require('../out/symbolProvider');
Module._load = originalLoad;

function documentOf(text) {
    const lines = text.split('\n');
    return { lineCount: lines.length, lineAt: i => ({ text: lines[i] }) };
}

test('document symbols include Japanese procedure, property, type and enum names', () => {
    const source = [
        'Sub 処理()', 'End Sub',
        'Private Function 計算1(x)', 'End Function',
        'Property Get 名前()', 'End Property',
        'Type 点', 'End Type',
        'Enum 色', 'End Enum'
    ].join('\n');
    const symbols = new VbaDocumentSymbolProvider().provideDocumentSymbols(documentOf(source));
    assert.deepEqual(symbols.map(s => s.name), ['処理', '計算1', 'Property Get 名前', '点', '色']);
    assert.deepEqual(symbols.map(s => s.kind), [11, 11, 6, 22, 9]);
    assert.deepEqual(symbols.map(s => [s.range.startLine, s.range.endLine]), [[0, 1], [2, 3], [4, 5], [6, 7], [8, 9]]);
});

test('document symbols still recognize ASCII names and full-width indentation', () => {
    const source = ['Public Sub Main()', 'End Sub', '　Function 計算１()', '　End Function'].join('\n');
    const symbols = new VbaDocumentSymbolProvider().provideDocumentSymbols(documentOf(source));
    assert.deepEqual(symbols.map(s => s.name), ['Main', '計算１']);
});
