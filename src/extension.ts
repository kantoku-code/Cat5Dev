import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { CatiaVbaTreeProvider } from './treeView';
import { VbaDocumentSymbolProvider } from './symbolProvider';
import { VbaDocumentFormatter, registerFormatOnSave } from './formatter';
import { VbaServer } from './vbaServer';
import { t, getLanguage, setLanguage, Language, messages } from './i18n';
import { registerLinter } from './linter';
import { tomlTemplate, gitignoreTemplate, readProjectSettings, writeTomlProjectKey } from './lintConfig';
import { startLspClient } from './lspClient';
import { CatiaSync } from './catiaSync';
import { registerFileCommands } from './fileCommands';

const outputChannel = vscode.window.createOutputChannel('CATIA VBA Sync');


export function activate(context: vscode.ExtensionContext) {
    const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '';
    const vbaServer = new VbaServer(context, outputChannel);
    context.subscriptions.push(vbaServer);
    vbaServer.start();
    const sync = new CatiaSync(vbaServer, outputChannel);

    let pullCmd = vscode.commands.registerCommand('cat5dev.pullFromCatia', () => {
        return sync.pull();
    });

    let pushCmd = vscode.commands.registerCommand('cat5dev.pushToCatia', () => {
        return sync.push();
    });

    let selectCmd = vscode.commands.registerCommand('cat5dev.selectProject', () => {
        return sync.selectProject();
    });

    const recoverCmd = vscode.commands.registerCommand('cat5dev.recoverSyncLock', () => sync.recoverLock());
    context.subscriptions.push(recoverCmd);

    let switchLanguageCmd = vscode.commands.registerCommand('cat5dev.switchLanguage', async () => {
        const currentLang = getLanguage();
        const selected = await vscode.window.showQuickPick(
            [
                { label: t('language.japanese'), description: currentLang === 'ja' ? t('language.description') : '', value: 'ja' },
                { label: t('language.english'), description: currentLang === 'en' ? t('language.description') : '', value: 'en' },
                { label: t('language.chinese'), description: currentLang === 'zh' ? t('language.description') : '', value: 'zh' }
            ],
            { placeHolder: t('language.title') }
        );
        if (selected && selected.value !== currentLang) {
            const nextLang = selected.value as Language;
            await setLanguage(nextLang);
            const msg = (messages[nextLang] as any)['language.reload'] || messages.ja['language.reload'];
            vscode.window.showInformationMessage(msg);
        }
    });

    registerFileCommands(context);

    const VBA_SELECTOR = ['bas_utf', 'cls_utf', 'frm_utf'].map(l => ({ language: 'vb', pattern: `**/*.${l}` }));
    context.subscriptions.push(
        vscode.languages.registerDocumentSymbolProvider(VBA_SELECTOR, new VbaDocumentSymbolProvider())
    );
    context.subscriptions.push(
        vscode.languages.registerDocumentFormattingEditProvider(VBA_SELECTOR, new VbaDocumentFormatter(vbaServer, outputChannel, workspaceRoot))
    );
    context.subscriptions.push(
        registerFormatOnSave(vbaServer, outputChannel, VBA_SELECTOR, workspaceRoot)
    );

    const treeProvider = new CatiaVbaTreeProvider(context);
    const treeView = vscode.window.createTreeView('catiaVbaModules', {
        treeDataProvider: treeProvider
    });
    vscode.commands.registerCommand('cat5dev.refreshTree', () => treeProvider.refresh());

    // Auto-refresh tree when sidebar becomes visible
    treeView.onDidChangeVisibility(e => {
        if (e.visible) treeProvider.refresh();
    });

    // Auto-refresh tree when cat5dev.toml changes (e.g. after selectProject)
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (workspaceFolders) {
        const configWatcher = vscode.workspace.createFileSystemWatcher(
            new vscode.RelativePattern(workspaceFolders[0], 'cat5dev.toml')
        );
        configWatcher.onDidChange(() => treeProvider.refresh());
        configWatcher.onDidCreate(() => treeProvider.refresh());
        context.subscriptions.push(configWatcher);

        // Auto-refresh tree when modules folder changes
        const modulesWatcher = vscode.workspace.createFileSystemWatcher(
            new vscode.RelativePattern(workspaceFolders[0], 'modules/**')
        );
        modulesWatcher.onDidChange(() => treeProvider.refresh());
        modulesWatcher.onDidCreate(() => treeProvider.refresh());
        modulesWatcher.onDidDelete(() => treeProvider.refresh());
        context.subscriptions.push(modulesWatcher);
    }

    context.subscriptions.push(treeView);

    // Linter 登録
    const linterDisposables = registerLinter(vbaServer, outputChannel, VBA_SELECTOR);
    context.subscriptions.push(...linterDisposables);

    // VBA Language Server（補完・ホバー・診断）を起動
    startLspClient(context);

    // cat5dev.init コマンド
    const initCmd = vscode.commands.registerCommand('cat5dev.init', async () => {
        const workspaceFolders = vscode.workspace.workspaceFolders;
        if (!workspaceFolders) {
            vscode.window.showErrorMessage(t('error.noWorkspace'));
            return;
        }
        const rootPath = workspaceFolders[0].uri.fsPath;
        const tomlPath = path.join(rootPath, 'cat5dev.toml');

        const overwriteBtn = t('init.overwrite');
        let existingTargetProject = '';
        if (fs.existsSync(tomlPath)) {
            existingTargetProject = readProjectSettings(rootPath).targetProject;
            const answer = await vscode.window.showWarningMessage(
                t('init.tomlExists'),
                { modal: true },
                overwriteBtn
            );
            if (answer !== overwriteBtn) {
                return;
            }
        }

        const currentLang = getLanguage();
        fs.writeFileSync(tomlPath, tomlTemplate(currentLang), 'utf-8');
        if (existingTargetProject) {
            writeTomlProjectKey(rootPath, 'target_project', existingTargetProject);
        }

        const gitignorePath = path.join(rootPath, '.gitignore');
        if (fs.existsSync(gitignorePath)) {
            const ans = await vscode.window.showWarningMessage(
                t('init.gitignoreExists'),
                { modal: true },
                overwriteBtn
            );
            if (ans === overwriteBtn) {
                fs.writeFileSync(gitignorePath, gitignoreTemplate(), 'utf-8');
            }
        } else {
            fs.writeFileSync(gitignorePath, gitignoreTemplate(), 'utf-8');
        }

        const doc = await vscode.workspace.openTextDocument(tomlPath);
        await vscode.window.showTextDocument(doc);
        vscode.window.showInformationMessage(t('init.success'));
    });

    context.subscriptions.push(pullCmd, pushCmd, selectCmd, switchLanguageCmd, initCmd);
}

export function deactivate() { }
