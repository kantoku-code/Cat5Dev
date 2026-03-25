import * as vscode from 'vscode';
import { VbaServer, httpPost } from './vbaServer';
import { readLintOptions, LintOptions } from './lintConfig';

const LINT_TIMEOUT_MS = 10000;
const DEBOUNCE_MS = 500;

interface GoDiagnostic {
    line: number;
    col: number;
    end_line: number;
    end_col: number;
    severity: string;
    code: string;
    message: string;
}

interface LintResponse {
    diagnostics: GoDiagnostic[];
    error: string;
}

/** Go サーバーの /lint を呼び出して診断結果を返す */
async function callLint(
    text: string,
    options: LintOptions,
    server: VbaServer,
    outputChannel: vscode.OutputChannel
): Promise<GoDiagnostic[] | null> {
    const baseUrl = server.getBaseUrl();
    if (baseUrl === null) {
        return null;
    }

    const body = JSON.stringify({ code: text, options });
    try {
        const raw = await httpPost(`${baseUrl}/lint`, body, LINT_TIMEOUT_MS);
        const resp = JSON.parse(raw) as LintResponse;
        if (resp.error) {
            outputChannel.appendLine(`[vba-lint] エラー: ${resp.error}`);
            return null;
        }
        return resp.diagnostics ?? [];
    } catch (err) {
        outputChannel.appendLine(`[vba-lint] リクエストエラー: ${err}`);
        return null;
    }
}

/** GoDiagnostic を VSCode の Diagnostic に変換する */
function toVscodeDiagnostic(d: GoDiagnostic, document: vscode.TextDocument): vscode.Diagnostic {
    const lineCount = document.lineCount;
    const startLine = Math.min(d.line, lineCount - 1);
    const endLine   = Math.min(d.end_line, lineCount - 1);

    const startChar = d.col >= 0 ? d.col : 0;
    const endChar   = d.end_col >= 0
        ? d.end_col
        : document.lineAt(endLine).text.length;

    const range = new vscode.Range(startLine, startChar, endLine, endChar);

    const severity = d.severity === 'error'
        ? vscode.DiagnosticSeverity.Error
        : vscode.DiagnosticSeverity.Warning;

    const diag = new vscode.Diagnostic(range, d.message, severity);
    diag.code = d.code;
    diag.source = 'cat5dev';
    return diag;
}

/** Linter クラス: DiagnosticCollection を管理する */
export class VbaLinter implements vscode.Disposable {
    private readonly collection: vscode.DiagnosticCollection;
    private readonly timers = new Map<string, NodeJS.Timeout>();
    private workspaceRoot: string | undefined;

    constructor(
        private readonly server: VbaServer,
        private readonly outputChannel: vscode.OutputChannel
    ) {
        this.collection = vscode.languages.createDiagnosticCollection('cat5dev-lint');
        this.workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    }

    /** ドキュメントを即座に Lint する */
    async lintNow(document: vscode.TextDocument): Promise<void> {
        const options = this.workspaceRoot
            ? readLintOptions(this.workspaceRoot)
            : undefined;

        if (!options) {
            return;
        }

        const diags = await callLint(document.getText(), options, this.server, this.outputChannel);
        if (diags === null) {
            return;
        }

        const vsDiags = diags.map(d => toVscodeDiagnostic(d, document));
        this.collection.set(document.uri, vsDiags);
    }

    /** デバウンス付き Lint（編集中のリアルタイム実行用） */
    lintDebounced(document: vscode.TextDocument): void {
        const key = document.uri.toString();
        const existing = this.timers.get(key);
        if (existing) {
            clearTimeout(existing);
        }
        const timer = setTimeout(() => {
            this.timers.delete(key);
            this.lintNow(document);
        }, DEBOUNCE_MS);
        this.timers.set(key, timer);
    }

    /** すべての開いている VB ドキュメントを再 Lint する */
    async relintAll(selector: vscode.DocumentSelector): Promise<void> {
        for (const doc of vscode.workspace.textDocuments) {
            if (vscode.languages.match(selector, doc)) {
                await this.lintNow(doc);
            }
        }
    }

    /** ドキュメントの診断をクリアする */
    clearDocument(uri: vscode.Uri): void {
        this.collection.delete(uri);
    }

    dispose(): void {
        for (const t of this.timers.values()) {
            clearTimeout(t);
        }
        this.timers.clear();
        this.collection.dispose();
    }
}

/** Linter を登録して Disposable のリストを返す */
export function registerLinter(
    server: VbaServer,
    outputChannel: vscode.OutputChannel,
    selector: vscode.DocumentSelector
): vscode.Disposable[] {
    const linter = new VbaLinter(server, outputChannel);
    const disposables: vscode.Disposable[] = [linter];

    // ファイルを開いたとき
    disposables.push(
        vscode.workspace.onDidOpenTextDocument((doc) => {
            if (vscode.languages.match(selector, doc)) {
                linter.lintNow(doc);
            }
        })
    );

    // 編集中（デバウンス付き）
    disposables.push(
        vscode.workspace.onDidChangeTextDocument((e) => {
            if (vscode.languages.match(selector, e.document)) {
                linter.lintDebounced(e.document);
            }
        })
    );

    // 保存時（cat5dev.toml の変更を検知して全ファイル再 Lint）
    disposables.push(
        vscode.workspace.onDidSaveTextDocument((doc) => {
            if (doc.fileName.endsWith('cat5dev.toml')) {
                linter.relintAll(selector);
                return;
            }
            if (vscode.languages.match(selector, doc)) {
                linter.lintNow(doc);
            }
        })
    );

    // ファイルを閉じたとき診断をクリア
    disposables.push(
        vscode.workspace.onDidCloseTextDocument((doc) => {
            linter.clearDocument(doc.uri);
        })
    );

    // 起動時に既に開いているドキュメントを Lint
    for (const doc of vscode.workspace.textDocuments) {
        if (vscode.languages.match(selector, doc)) {
            linter.lintNow(doc);
        }
    }

    return disposables;
}
