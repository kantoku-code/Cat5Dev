import * as vscode from 'vscode';
import { VbaServer, httpPost } from './vbaServer';

const FORMATTER_TIMEOUT_MS = 10000;

/** VBA ファイルをフォーマットする (vbafmt HTTP サーバーを呼び出す) */
export async function formatVbaDocument(
    text: string,
    server: VbaServer,
    outputChannel: vscode.OutputChannel
): Promise<string | null> {
    const baseUrl = server.getBaseUrl();
    if (baseUrl === null) {
        outputChannel.appendLine('[vbafmt] サーバーが起動していません');
        return null;
    }

    const config = vscode.workspace.getConfiguration('cat5dev.formatter');

    const requestBody = JSON.stringify({
        code: text,
        options: {
            indent_size: config.get<number>('indentSize', 4),
            capitalize_keywords: config.get<boolean>('capitalizeKeywords', true),
            fix_indentation: config.get<boolean>('fixIndentation', true),
            line_endings: 'CRLF',
            trim_trailing_space: config.get<boolean>('trimTrailingSpace', true),
            ensure_continuation_space: config.get<boolean>('ensureContinuationSpace', true),
            indent_continuation_lines: config.get<boolean>('indentContinuationLines', true),
            max_blank_lines: config.get<number>('maxBlankLines', 2),
            normalize_operator_spacing: config.get<boolean>('normalizeOperatorSpacing', false),
            normalize_comma_spacing: config.get<boolean>('normalizeCommaSpacing', false),
            normalize_comment_space: config.get<boolean>('normalizeCommentSpace', false),
            expand_type_suffixes: config.get<boolean>('expandTypeSuffixes', false),
            split_colon_statements: false,
            normalize_then_placement: false,
            normalize_on_error: false,
        }
    });

    try {
        const responseText = await httpPost(`${baseUrl}/format`, requestBody, FORMATTER_TIMEOUT_MS);
        const response = JSON.parse(responseText) as { result: string; error: string };
        if (response.error) {
            outputChannel.appendLine(`[vbafmt] エラー: ${response.error}`);
            return null;
        }
        return response.result;
    } catch (err) {
        outputChannel.appendLine(`[vbafmt] リクエストエラー: ${err}`);
        return null;
    }
}

/** DocumentFormattingEditProvider 実装 */
export class VbaDocumentFormatter implements vscode.DocumentFormattingEditProvider {
    constructor(
        private readonly server: VbaServer,
        private readonly outputChannel: vscode.OutputChannel
    ) {}

    async provideDocumentFormattingEdits(
        document: vscode.TextDocument,
        _options: vscode.FormattingOptions,
        _token: vscode.CancellationToken
    ): Promise<vscode.TextEdit[]> {
        const formatted = await formatVbaDocument(
            document.getText(),
            this.server,
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
    server: VbaServer,
    outputChannel: vscode.OutputChannel,
    selector: vscode.DocumentSelector
): vscode.Disposable {
    return vscode.workspace.onWillSaveTextDocument((event) => {
        const config = vscode.workspace.getConfiguration('cat5dev.formatter');
        if (!config.get<boolean>('formatOnSave', false)) {
            return;
        }

        if (!vscode.languages.match(selector, event.document)) {
            return;
        }

        const formatPromise = formatVbaDocument(
            event.document.getText(),
            server,
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
