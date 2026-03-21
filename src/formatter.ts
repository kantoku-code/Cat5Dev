import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { spawn } from 'child_process';

const FORMATTER_TIMEOUT_MS = 10000;

/** VBA ファイルをフォーマットする (vbafmt バイナリを呼び出す) */
export async function formatVbaDocument(
    text: string,
    context: vscode.ExtensionContext,
    outputChannel: vscode.OutputChannel
): Promise<string | null> {
    const binaryName = process.platform === 'win32' ? 'vbafmt.exe' : 'vbafmt';
    const binaryPath = path.join(context.extensionPath, 'bin', binaryName);

    if (!fs.existsSync(binaryPath)) {
        outputChannel.appendLine(`[vbafmt] バイナリが見つかりません: ${binaryPath}`);
        outputChannel.appendLine('[vbafmt] npm run compile:go を実行してバイナリをビルドしてください');
        return null;
    }

    const config = vscode.workspace.getConfiguration('cat5dev.formatter');
    const indentSize: number = config.get('indentSize', 4);
    const capitalizeKeywords: boolean = config.get('capitalizeKeywords', true);
    const fixIndentation: boolean = config.get('fixIndentation', true);

    const args: string[] = [
        `--indent-size=${indentSize}`,
        `--capitalize-keywords=${capitalizeKeywords}`,
        `--fix-indentation=${fixIndentation}`,
        '--line-endings=CRLF',
    ];

    return new Promise((resolve) => {
        const proc = spawn(binaryPath, args);
        let stdout = '';
        let stderr = '';

        const timer = setTimeout(() => {
            proc.kill();
            outputChannel.appendLine('[vbafmt] タイムアウト (10秒)');
            resolve(null);
        }, FORMATTER_TIMEOUT_MS);

        proc.stdout.on('data', (data: Buffer) => {
            stdout += data.toString('utf8');
        });
        proc.stderr.on('data', (data: Buffer) => {
            stderr += data.toString('utf8');
        });
        proc.on('close', (code: number) => {
            clearTimeout(timer);
            if (code !== 0 || stderr) {
                outputChannel.appendLine(`[vbafmt] エラー (exit ${code}): ${stderr}`);
                resolve(null);
                return;
            }
            resolve(stdout);
        });
        proc.on('error', (err: Error) => {
            clearTimeout(timer);
            outputChannel.appendLine(`[vbafmt] プロセス起動エラー: ${err.message}`);
            resolve(null);
        });

        proc.stdin.write(text, 'utf8');
        proc.stdin.end();
    });
}

/** DocumentFormattingEditProvider 実装 */
export class VbaDocumentFormatter implements vscode.DocumentFormattingEditProvider {
    constructor(
        private readonly context: vscode.ExtensionContext,
        private readonly outputChannel: vscode.OutputChannel
    ) {}

    async provideDocumentFormattingEdits(
        document: vscode.TextDocument,
        _options: vscode.FormattingOptions,
        _token: vscode.CancellationToken
    ): Promise<vscode.TextEdit[]> {
        const formatted = await formatVbaDocument(
            document.getText(),
            this.context,
            this.outputChannel
        );
        if (formatted === null) {
            return [];
        }
        const fullRange = new vscode.Range(
            document.positionAt(0),
            document.positionAt(document.getText().length)
        );
        return [vscode.TextEdit.replace(fullRange, formatted)];
    }
}

/** formatOnSave ハンドラを登録する */
export function registerFormatOnSave(
    context: vscode.ExtensionContext,
    outputChannel: vscode.OutputChannel,
    selector: vscode.DocumentSelector
): vscode.Disposable {
    return vscode.workspace.onWillSaveTextDocument((event) => {
        const config = vscode.workspace.getConfiguration('cat5dev.formatter');
        if (!config.get<boolean>('formatOnSave', false)) {
            return;
        }

        // selector に一致するドキュメントのみ
        if (!vscode.languages.match(selector, event.document)) {
            return;
        }

        const formatPromise = formatVbaDocument(
            event.document.getText(),
            context,
            outputChannel
        ).then((formatted) => {
            if (formatted === null) {
                return [];
            }
            const fullRange = new vscode.Range(
                event.document.positionAt(0),
                event.document.positionAt(event.document.getText().length)
            );
            return [vscode.TextEdit.replace(fullRange, formatted)];
        });

        event.waitUntil(formatPromise);
    });
}
