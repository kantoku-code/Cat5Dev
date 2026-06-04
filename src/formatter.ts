import * as vscode from 'vscode';
import { VbaServer, httpPost } from './vbaServer';
import { readFormatterOptions } from './lintConfig';
import { t } from './i18n';

const FORMATTER_TIMEOUT_MS = 10000;

/** VBA ファイルをフォーマットする (vbafmt HTTP サーバーを呼び出す) */
export async function formatVbaDocument(
    text: string,
    server: VbaServer,
    outputChannel: vscode.OutputChannel,
    workspaceRoot: string
): Promise<string | null> {
    const baseUrl = server.getBaseUrl();
    if (baseUrl === null) {
        outputChannel.appendLine(t('log.vbafmt.serverNotRunning'));
        return null;
    }

    const opts = readFormatterOptions(workspaceRoot);
    if (!opts.enabled) {
        return null;
    }

    const requestBody = JSON.stringify({
        code: text,
        options: {
            indent_size: opts.indent_size,
            capitalize_keywords: opts.capitalize_keywords,
            fix_indentation: opts.fix_indentation,
            line_endings: 'CRLF',
            trim_trailing_space: opts.trim_trailing_space,
            ensure_continuation_space: opts.ensure_continuation_space,
            indent_continuation_lines: opts.indent_continuation_lines,
            max_blank_lines: opts.max_blank_lines,
            normalize_operator_spacing: opts.normalize_operator_spacing,
            normalize_comma_spacing: opts.normalize_comma_spacing,
            normalize_comment_space: opts.normalize_comment_space,
            expand_type_suffixes: opts.expand_type_suffixes,
            split_colon_statements: false,
            normalize_then_placement: false,
            normalize_on_error: false,
        }
    });

    try {
        const responseText = await httpPost(`${baseUrl}/format`, requestBody, FORMATTER_TIMEOUT_MS);
        const response = JSON.parse(responseText) as { result: string; error: string };
        if (response.error) {
            outputChannel.appendLine(t('log.vbafmt.error', response.error));
            return null;
        }
        return response.result;
    } catch (err) {
        outputChannel.appendLine(t('log.vbafmt.requestError', String(err)));
        return null;
    }
}

/** DocumentFormattingEditProvider 実装 */
export class VbaDocumentFormatter implements vscode.DocumentFormattingEditProvider {
    constructor(
        private readonly server: VbaServer,
        private readonly outputChannel: vscode.OutputChannel,
        private readonly workspaceRoot: string
    ) {}

    async provideDocumentFormattingEdits(
        document: vscode.TextDocument,
        _options: vscode.FormattingOptions,
        _token: vscode.CancellationToken
    ): Promise<vscode.TextEdit[]> {
        const formatted = await formatVbaDocument(
            document.getText(),
            this.server,
            this.outputChannel,
            this.workspaceRoot
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
    selector: vscode.DocumentSelector,
    workspaceRoot: string
): vscode.Disposable {
    return vscode.workspace.onWillSaveTextDocument((event) => {
        const opts = readFormatterOptions(workspaceRoot);
        if (!opts.format_on_save) {
            return;
        }

        if (!vscode.languages.match(selector, event.document)) {
            return;
        }

        const formatPromise = formatVbaDocument(
            event.document.getText(),
            server,
            outputChannel,
            workspaceRoot
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
