import * as path from 'path';
import * as fs from 'fs';
import * as vscode from 'vscode';
import {
    LanguageClient,
    LanguageClientOptions,
    ServerOptions,
    TransportKind,
} from 'vscode-languageclient/node';

let client: LanguageClient | undefined;

/** サーバーの索引対象（標準モジュール・クラス・フォーム）。 */
export const MODULE_GLOB = '**/*.{bas_utf,cls_utf,frm_utf}';

/** クライアント設定を作る（テスト用に切り出し）。
 *  scheme を file に限るのは、git の差分ビューなどのバッファをモジュールとして扱わないため。 */
export function buildClientOptions(watcher: vscode.FileSystemWatcher): LanguageClientOptions {
    return {
        documentSelector: [
            { scheme: 'file', language: 'vb', pattern: '**/*.bas_utf' },
            { scheme: 'file', language: 'vb', pattern: '**/*.cls_utf' },
            { scheme: 'file', language: 'vb', pattern: '**/*.frm_utf' },
        ],
        synchronize: { fileEvents: watcher },
    };
}

/** vba-ls.exe のパスを解決する。
 *  優先順位: 環境変数 CAT5DEV_VBA_LS → 拡張の bin/ ディレクトリ */
function resolveServerPath(context: vscode.ExtensionContext): string | undefined {
    const envPath = process.env['CAT5DEV_VBA_LS'];
    if (envPath && fs.existsSync(envPath)) {
        return envPath;
    }
    const bundled = path.join(context.extensionPath, 'bin', 'vba-ls.exe');
    if (fs.existsSync(bundled)) {
        return bundled;
    }
    return undefined;
}

/** LSP クライアントを起動する。既に起動中の場合は何もしない。 */
export function startLspClient(context: vscode.ExtensionContext): void {
    if (client) {
        return;
    }

    const serverExe = resolveServerPath(context);
    if (!serverExe) {
        // vba-ls.exe が見つからない場合は LSP 機能を無効化（フォーマッタ・Lint には影響しない）
        return;
    }

    const serverOptions: ServerOptions = {
        run:   { command: serverExe, transport: TransportKind.stdio },
        debug: { command: serverExe, transport: TransportKind.stdio },
    };

    // モジュールの作成・削除・リネーム（Pull による書き換えを含む）をサーバーの索引に通知する
    const watcher = vscode.workspace.createFileSystemWatcher(MODULE_GLOB);
    context.subscriptions.push(watcher);
    const clientOptions = buildClientOptions(watcher);

    client = new LanguageClient(
        'vba-ls',
        'VBA Language Server',
        serverOptions,
        clientOptions
    );

    client.start();
    context.subscriptions.push({ dispose: () => stopLspClient() });
}

/** LSP クライアントを停止する。 */
export async function stopLspClient(): Promise<void> {
    if (!client) {
        return;
    }
    await client.stop();
    client = undefined;
}
