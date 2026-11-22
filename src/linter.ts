import * as vscode from 'vscode';
import { VbaServer, httpPost } from './vbaServer';
import { readLintOptions, LintOptions } from './lintConfig';
import { t } from './i18n';
import {
    NON_ANSI_DIAGNOSTIC_CODE, detectSystemAnsiCodePage, findNonAnsiRanges, getAnsiCharChecker,
} from './ansiCodePage';

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

    // VBA010（non_ansi_character）は拡張機能側で判定するので、Go には送らない
    const { non_ansi_character: _nonAnsi, ansi_code_page: _codePage, ...goOptions } = options;
    const body = JSON.stringify({ code: text, options: goOptions });
    try {
        const raw = await httpPost(`${baseUrl}/lint`, body, LINT_TIMEOUT_MS);
        const resp = JSON.parse(raw) as LintResponse;
        if (resp.error) {
            outputChannel.appendLine(t('log.vbalint.error', resp.error));
            return null;
        }
        return resp.diagnostics ?? [];
    } catch (err) {
        outputChannel.appendLine(t('log.vbalint.requestError', String(err)));
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
    /** 文書ごとの Lint の通し番号。最後に始めた Lint の結果だけを反映する */
    private readonly generations = new Map<string, number>();
    private workspaceRoot: string | undefined;
    /** システムの ANSI コードページ。起動時に 1 回だけ非同期で検出する */
    private readonly systemCodePage: Promise<number>;
    /** 対応外として一度ログに出したコードページ（同じログを繰り返さない） */
    private readonly reportedUnsupported = new Set<number>();

    constructor(
        private readonly server: VbaServer,
        private readonly outputChannel: vscode.OutputChannel
    ) {
        this.collection = vscode.languages.createDiagnosticCollection('cat5dev-lint');
        this.workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
        this.systemCodePage = detectSystemAnsiCodePage(vscode.env.language).then(r => {
            this.outputChannel.appendLine(r.source === 'registry'
                ? t('log.vbalint.ansiCodePageDetected', String(r.codePage))
                : t('log.vbalint.ansiCodePageFallback', String(r.codePage), r.reason ?? ''));
            return r.codePage;
        });
    }

    /** VBA010: CATIA の VBA（ANSI コードページ）で往復できない文字を警告する */
    private async nonAnsiDiagnostics(text: string, options: LintOptions): Promise<vscode.Diagnostic[]> {
        if (!options.non_ansi_character) {
            return [];
        }
        const codePage = options.ansi_code_page > 0 ? options.ansi_code_page : await this.systemCodePage;
        const checker = getAnsiCharChecker(codePage);
        if (!checker) {
            if (!this.reportedUnsupported.has(codePage)) {
                this.reportedUnsupported.add(codePage);
                this.outputChannel.appendLine(t('log.vbalint.ansiCodePageUnsupported', String(codePage)));
            }
            return [];
        }
        return findNonAnsiRanges(text, checker).map(r => {
            const diag = new vscode.Diagnostic(
                new vscode.Range(r.line, r.startCol, r.line, r.endCol),
                t('lint.nonAnsiCharacter', String(codePage), r.text),
                vscode.DiagnosticSeverity.Warning,
            );
            diag.code = NON_ANSI_DIAGNOSTIC_CODE;
            diag.source = 'cat5dev';
            return diag;
        });
    }

    /** ドキュメントを即座に Lint する */
    async lintNow(document: vscode.TextDocument): Promise<void> {
        const options = this.workspaceRoot
            ? readLintOptions(this.workspaceRoot)
            : undefined;

        if (!options) {
            return;
        }

        // await の間に文書が編集・Lint し直し・クローズされたら、古い結果で上書きしない
        const key = document.uri.toString();
        const version = document.version;
        const generation = (this.generations.get(key) ?? 0) + 1;
        this.generations.set(key, generation);

        const text = document.getText();
        const [diags, nonAnsi] = await Promise.all([
            callLint(text, options, this.server, this.outputChannel),
            this.nonAnsiDiagnostics(text, options),
        ]);
        if (document.isClosed || document.version !== version || this.generations.get(key) !== generation) {
            return;
        }

        // Go の Lint に失敗したときは、前回の Go の診断を残して VBA010 だけ更新する
        const goDiags = diags === null
            ? (this.collection.get(document.uri) ?? []).filter(d => d.code !== NON_ANSI_DIAGNOSTIC_CODE)
            : diags.map(d => toVscodeDiagnostic(d, document));
        this.collection.set(document.uri, [...goDiags, ...nonAnsi]);
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
        const key = uri.toString();
        const timer = this.timers.get(key);
        if (timer) {
            clearTimeout(timer);
            this.timers.delete(key);
        }
        // 消さずに進める: 実行中の Lint は捨て、開き直した後の Lint と番号が重ならない
        this.generations.set(key, (this.generations.get(key) ?? 0) + 1);
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
