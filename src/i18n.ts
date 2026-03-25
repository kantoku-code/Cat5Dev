import * as vscode from 'vscode';

export type Language = 'ja' | 'en';

export function getLanguage(): Language {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) return 'ja';

    const settingsPath = require('path').join(workspaceFolders[0].uri.fsPath, '.vscode', 'settings.json');
    if (require('fs').existsSync(settingsPath)) {
        try {
            const settings = JSON.parse(require('fs').readFileSync(settingsPath, 'utf-8'));
            const lang = settings.language;
            if (lang === 'ja' || lang === 'en') {
                return lang;
            }
        } catch (e) { }
    }
    return 'ja';
}

export function setLanguage(lang: Language): void {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) return;

    const path = require('path');
    const fs = require('fs');
    const rootPath = workspaceFolders[0].uri.fsPath;
    const vscodePath = path.join(rootPath, '.vscode');
    const settingsPath = path.join(vscodePath, 'settings.json');

    if (!fs.existsSync(vscodePath)) {
        fs.mkdirSync(vscodePath, { recursive: true });
    }

    let settings: any = {};
    if (fs.existsSync(settingsPath)) {
        try {
            settings = JSON.parse(fs.readFileSync(settingsPath, 'utf-8'));
        } catch (e) { }
    }

    settings.language = lang;
    fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 4), 'utf-8');
}

export const messages = {
    ja: {
        // Sidebar & TreeView
        'sidebar.title': 'CATIA V5 VBA',
        'sidebar.modules': 'Modules',
        'treeview.modules': 'Modules',
        'treeview.classModules': 'Class Modules',
        'treeview.forms': 'Forms',
        'treeview.targetProject': 'Target CATIA VBA Project',

        // Commands
        'command.pull': 'CATIA: Pull VBA Modules',
        'command.push': 'CATIA: Push VBA Modules',
        'command.select': 'CATIA: Select Target Project',
        'command.refresh': 'Refresh Modules',
        'command.switchLanguage': 'CATIA: Switch Language',

        // Error messages
        'error.noWorkspace': 'CATIA VBA同期設定を行うには、ワークスペースフォルダを開いてください。',
        'error.pullFailed': 'プルに失敗しました。詳細 Output を確認してください。',
        'error.pushFailed': 'プッシュに失敗しました。詳細 Output を確認してください。',
        'error.selectFailed': 'CATIAからVBAプロジェクトを取得できませんでした。\n詳細 Output を確認してください。',
        'error.noModulesDir': 'プッシュ対象の modules ディレクトリがワークスペース内に見つかりません。',
        'error.noModuleFiles': 'ワークスペース内にプッシュ対象のVBAファイル (.bas_utf, .cls_utf, .frm_utf) が見つかりませんでした。',
        'error.checkComponentsFailed': '[Check Components Error]',

        // Information messages
        'info.projectNotFound': 'CATIA内にVBAプロジェクトが見つかりませんでした。',
        'info.projectSelected': 'ターゲットVBAプロジェクトを {0} に設定しました。',
        'info.pullSuccess': 'CATIAから {0} 個のモジュールを正常にプルしました。',
        'info.pushSuccess': '{0} 個のモジュールをプッシュしました。{1}{2}',
        'info.noChanges': 'すべてのモジュール ({0} 個) は前回のプッシュから変更されていません。プッシュをスキップしました。',
        'info.pullCancelled': 'プルをキャンセルしました。',
        'info.pushCancelled': 'プッシュを中止しました。',

        // Warning messages
        'warning.deleteModules': '以下のモジュールはCATIA側に存在しますが、VSCode側には存在しません:\n{0}\n\nこれらをCATIAから削除して完全に同期しますか？',
        'warning.newUserForms': '以下のUserFormはCATIA側に存在しないため新規作成できません。CATIA側で空の同名UserFormを事前に作成してください。これらのファイルはスキップされます:\n{0}',
        'warning.noMoreFiles': 'プッシュ対象のファイルがなくなったため、処理を終了します。',
        'warning.longModuleNames': '以下のモジュール名は31文字を超えています（VBAエディタの上限）:\n{0}\n\nこのままプッシュを続行しますか？',

        // Dialog options
        'dialog.delete': 'はい（削除する）',
        'dialog.keep': 'いいえ（残す）',
        'dialog.continue': '続行',

        // Progress titles
        'progress.pull': 'CATIAからVBAをプルしています ({0})...',
        'progress.push': 'CATIAへVBAをプッシュしています ({0})...',

        // Select project
        'select.placeholder': '同期対象のCATIA VBAプロジェクトを選択してください',

        // Language switch
        'language.title': '言語を選択してください',
        'language.japanese': '日本語',
        'language.english': 'English',

        // File operations
        'file.rename': '名前変更',
        'file.delete': '削除',
        'file.copy': 'パスをコピー',
    },
    en: {
        // Sidebar & TreeView
        'sidebar.title': 'CATIA V5 VBA',
        'sidebar.modules': 'Modules',
        'treeview.modules': 'Modules',
        'treeview.classModules': 'Class Modules',
        'treeview.forms': 'Forms',
        'treeview.targetProject': 'Target CATIA VBA Project',

        // Commands
        'command.pull': 'CATIA: Pull VBA Modules',
        'command.push': 'CATIA: Push VBA Modules',
        'command.select': 'CATIA: Select Target Project',
        'command.refresh': 'Refresh Modules',
        'command.switchLanguage': 'CATIA: Switch Language',

        // Error messages
        'error.noWorkspace': 'Open a workspace folder to configure CATIA VBA sync.',
        'error.pullFailed': 'Pull failed. Check the Output panel for details.',
        'error.pushFailed': 'Push failed. Check the Output panel for details.',
        'error.selectFailed': 'Failed to retrieve VBA projects from CATIA.\nCheck the Output panel for details.',
        'error.noModulesDir': 'The modules directory not found in workspace.',
        'error.noModuleFiles': 'No VBA files (.bas_utf, .cls_utf, .frm_utf) found in workspace.',
        'error.checkComponentsFailed': '[Check Components Error]',

        // Information messages
        'info.projectNotFound': 'No VBA projects found in CATIA.',
        'info.projectSelected': 'Target VBA project set to {0}.',
        'info.pullSuccess': 'Successfully pulled {0} modules from CATIA.',
        'info.pushSuccess': 'Pushed {0} modules.{1}{2}',
        'info.noChanges': 'All modules ({0}) are unchanged since the last push. Skipped.',
        'info.pullCancelled': 'Pull cancelled.',
        'info.pushCancelled': 'Push cancelled.',

        // Warning messages
        'warning.deleteModules': 'The following modules exist in CATIA but not in VSCode:\n{0}\n\nDelete them from CATIA to fully sync?',
        'warning.newUserForms': 'The following UserForms do not exist in CATIA and cannot be created. Create empty UserForms with these names in CATIA first. These files will be skipped:\n{0}',
        'warning.noMoreFiles': 'No more files to push. Finishing.',
        'warning.longModuleNames': 'The following module names exceed 31 characters (VBA editor limit):\n{0}\n\nProceed with push anyway?',

        // Dialog options
        'dialog.delete': 'Yes (Delete)',
        'dialog.keep': 'No (Keep)',
        'dialog.continue': 'Proceed',

        // Progress titles
        'progress.pull': 'Pulling VBA from CATIA ({0})...',
        'progress.push': 'Pushing VBA to CATIA ({0})...',

        // Select project
        'select.placeholder': 'Select a CATIA VBA project to sync',

        // Language switch
        'language.title': 'Select a language',
        'language.japanese': '日本語',
        'language.english': 'English',

        // File operations
        'file.rename': 'Rename',
        'file.delete': 'Delete',
        'file.copy': 'Copy Path',
    }
};

export function t(key: keyof typeof messages.ja, ...args: string[]): string {
    const lang = getLanguage();
    let text = messages[lang][key as keyof typeof messages.ja] || messages.ja[key];

    // Simple string interpolation
    args.forEach((arg, index) => {
        text = text.replace(`{${index}}`, arg);
    });

    return text;
}
