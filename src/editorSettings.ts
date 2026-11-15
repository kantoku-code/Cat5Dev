import * as vscode from 'vscode';
import { readBreadcrumbsSort } from './lintConfig';

/**
 * cat5dev.toml の [editor] セクションを VSCode のワークスペース設定へ反映する。
 *
 * breadcrumbs.symbolSortOrder は言語別オーバーライド可能な設定だが、
 * VSCode のパンくずドロップダウンはリソース無しで値を読むため [vb] 指定が効かない。
 * そのためワークスペース単位（.vscode/settings.json）に書き込んで適用する。
 * キーが無い場合は既存のワークスペース設定に触れない。
 */
export function registerEditorSettingsSync(outputChannel: vscode.OutputChannel): vscode.Disposable {
    const folder = vscode.workspace.workspaceFolders?.[0];
    if (!folder) { return new vscode.Disposable(() => { /* no workspace */ }); }

    const apply = async () => {
        const order = readBreadcrumbsSort(folder.uri.fsPath);
        if (!order) { return; }

        const cfg = vscode.workspace.getConfiguration('breadcrumbs', folder.uri);
        if (cfg.inspect<string>('symbolSortOrder')?.workspaceValue === order) { return; }

        try {
            await cfg.update('symbolSortOrder', order, vscode.ConfigurationTarget.Workspace);
            outputChannel.appendLine(`[editor] breadcrumbs.symbolSortOrder = "${order}" (cat5dev.toml)`);
        } catch (e) {
            outputChannel.appendLine(`[editor] breadcrumbs.symbolSortOrder の更新に失敗: ${e}`);
        }
    };

    const watcher = vscode.workspace.createFileSystemWatcher(
        new vscode.RelativePattern(folder, 'cat5dev.toml')
    );
    watcher.onDidChange(apply);
    watcher.onDidCreate(apply);

    void apply();
    return watcher;
}
