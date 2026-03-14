import * as vscode from 'vscode';

const START_PATTERN = /^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:(Sub|Function)\s+(\w+)\s*\(|(Property\s+(?:Get|Let|Set))\s+(\w+)\s*\(|(Type)\s+(\w+)|(Enum)\s+(\w+))/i;
const END_PATTERN = /^\s*End\s+(Sub|Function|Property|Type|Enum)\b/i;

export class VbaDocumentSymbolProvider implements vscode.DocumentSymbolProvider {
    provideDocumentSymbols(document: vscode.TextDocument): vscode.DocumentSymbol[] {
        const symbols: vscode.DocumentSymbol[] = [];

        for (let i = 0; i < document.lineCount; i++) {
            const line = document.lineAt(i);
            const match = START_PATTERN.exec(line.text);
            if (!match) { continue; }

            let name: string;
            let kind: vscode.SymbolKind;

            if (match[1] && match[2]) {
                name = match[2];
                kind = vscode.SymbolKind.Function;
            } else if (match[3] && match[4]) {
                name = `${match[3].replace(/\s+/g, ' ')} ${match[4]}`;
                kind = vscode.SymbolKind.Property;
            } else if (match[5] && match[6]) {
                name = match[6];
                kind = vscode.SymbolKind.Struct;
            } else if (match[7] && match[8]) {
                name = match[8];
                kind = vscode.SymbolKind.Enum;
            } else {
                continue;
            }

            // Find the matching End Sub/Function/Property/Type/Enum
            let endLine = i;
            for (let j = i + 1; j < document.lineCount; j++) {
                if (END_PATTERN.test(document.lineAt(j).text)) {
                    endLine = j;
                    break;
                }
            }

            const startPos = new vscode.Range(i, 0, i, line.text.length);
            const fullRange = new vscode.Range(i, 0, endLine, document.lineAt(endLine).text.length);
            symbols.push(new vscode.DocumentSymbol(name, '', kind, fullRange, startPos));
        }

        return symbols;
    }
}
