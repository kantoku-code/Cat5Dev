import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { readProjectSettings, writeTomlProjectKey } from './lintConfig';

export type Language = 'ja' | 'en' | 'zh';

export function getLanguage(): Language {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (workspaceFolders) {
        const tomlPath = path.join(workspaceFolders[0].uri.fsPath, 'cat5dev.toml');
        if (fs.existsSync(tomlPath)) {
            const { language } = readProjectSettings(workspaceFolders[0].uri.fsPath);
            if (language === 'ja' || language === 'en' || language === 'zh') { return language; }
        }
    }

    // Fallback 1: VSCode config setting
    const configLang = vscode.workspace.getConfiguration('cat5dev').get<string>('language');
    if (configLang === 'ja' || configLang === 'en' || configLang === 'zh') { return configLang; }

    // Fallback 2: VSCode display language
    const envLang = vscode.env.language.toLowerCase();
    if (envLang.startsWith('zh')) { return 'zh'; }
    if (envLang.startsWith('ja')) { return 'ja'; }

    return 'en'; // Default fallback
}

export async function setLanguage(lang: Language): Promise<void> {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) { return; }

    const tomlPath = path.join(workspaceFolders[0].uri.fsPath, 'cat5dev.toml');
    if (fs.existsSync(tomlPath)) {
        writeTomlProjectKey(workspaceFolders[0].uri.fsPath, 'language', lang);
    } else {
        // Fallback to updating VSCode configuration if toml doesn't exist
        await vscode.workspace.getConfiguration('cat5dev').update('language', lang, vscode.ConfigurationTarget.Workspace);
    }
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
        'language.chinese': '简体中文',

        // File operations
        'file.rename': '名前変更',
        'file.delete': '削除',
        'file.copy': 'パスをコピー',
        'file.deleteConfirm': '{0} を削除しますか？',
        'file.deleteButton': '削除',
        'file.copySuccess': 'パスをコピーしました: {0}',

        // Language switch
        'language.description': '(現在の言語)',
        'language.reload': '言語変更を反映するにはVSCodeを再読み込みしてください。',

        // Push
        'push.deleteSync': '（削除同期を含む）',

        // Init command
        'init.tomlExists': 'cat5dev.toml は既に存在します。上書きしますか？',
        'init.overwrite': '上書き',
        'init.gitignoreExists': '.gitignore は既に存在します。上書きしますか？',
        'init.success': 'cat5dev.toml を作成しました。',
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
        'language.chinese': '简体中文',

        // File operations
        'file.rename': 'Rename',
        'file.delete': 'Delete',
        'file.copy': 'Copy Path',
        'file.deleteConfirm': 'Delete {0}?',
        'file.deleteButton': 'Delete',
        'file.copySuccess': 'Path copied: {0}',

        // Language switch
        'language.description': '(Current language)',
        'language.reload': 'Please reload VSCode to apply language changes.',

        // Push
        'push.deleteSync': ' (includes deletion sync)',

        // Init command
        'init.tomlExists': 'cat5dev.toml already exists. Overwrite?',
        'init.overwrite': 'Overwrite',
        'init.gitignoreExists': '.gitignore already exists. Overwrite?',
        'init.success': 'cat5dev.toml has been created.',
    },
    zh: {
        // Sidebar & TreeView
        'sidebar.title': 'CATIA V5 VBA',
        'sidebar.modules': '模块',
        'treeview.modules': '标准模块',
        'treeview.classModules': '类模块',
        'treeview.forms': '用户窗体',
        'treeview.targetProject': '目标 CATIA VBA 项目',

        // Commands
        'command.pull': 'CATIA: 拉取 VBA 模块',
        'command.push': 'CATIA: 推送 VBA 模块',
        'command.select': 'CATIA: 选择目标项目',
        'command.refresh': '刷新模块列表',
        'command.switchLanguage': 'CATIA: 切换语言',

        // Error messages
        'error.noWorkspace': '请打开一个工作区文件夹以配置 CATIA VBA 同步。',
        'error.pullFailed': '拉取失败。详情请检查输出面板。',
        'error.pushFailed': '推送失败。详情请检查输出面板。',
        'error.selectFailed': '无法从 CATIA 获取 VBA 项目。\n详情请检查输出面板。',
        'error.noModulesDir': '工作区中未找到 modules 目录。',
        'error.noModuleFiles': '工作区中未找到可推送的 VBA 文件 (.bas_utf, .cls_utf, .frm_utf)。',
        'error.checkComponentsFailed': '[检查组件错误]',

        // Information messages
        'info.projectNotFound': '在 CATIA 中未找到 VBA 项目。',
        'info.projectSelected': '已将目标 VBA 项目设置为 {0}。',
        'info.pullSuccess': '成功从 CATIA 拉取了 {0} 个模块。',
        'info.pushSuccess': '已推送 {0} 个模块。{1}{2}',
        'info.noChanges': '自上次推送以来，所有模块 ({0} 个) 均未更改。已跳过推送。',
        'info.pullCancelled': '拉取已取消。',
        'info.pushCancelled': '推送已取消。',

        // Warning messages
        'warning.deleteModules': '以下模块存在于 CATIA 中，但不存在于 VSCode 中:\n{0}\n\n是否从 CATIA 中删除它们以实现完全同步？',
        'warning.newUserForms': '以下用户窗体 (UserForm) 在 CATIA 中不存在，无法新建。请先在 CATIA 中创建同名的空用户窗体。这些文件将被跳过:\n{0}',
        'warning.noMoreFiles': '没有更多可推送的文件。处理结束。',
        'warning.longModuleNames': '以下模块名称超过了 31 个字符（VBA 编辑器的限制）:\n{0}\n\n是否仍要继续推送？',

        // Dialog options
        'dialog.delete': '是 (删除)',
        'dialog.keep': '否 (保留)',
        'dialog.continue': '继续',

        // Progress titles
        'progress.pull': '正在从 CATIA 拉取 VBA ({0})...',
        'progress.push': '正在向 CATIA 推送 VBA ({0})...',

        // Select project
        'select.placeholder': '选择要同步的 CATIA VBA 项目',

        // Language switch
        'language.title': '请选择语言',
        'language.japanese': '日本語',
        'language.english': 'English',
        'language.chinese': '简体中文',

        // File operations
        'file.rename': '重命名',
        'file.delete': '删除',
        'file.copy': '复制路径',
        'file.deleteConfirm': '确定要删除 {0} 吗？',
        'file.deleteButton': '删除',
        'file.copySuccess': '已复制路径: {0}',

        // Language switch
        'language.description': '(当前语言)',
        'language.reload': '请重新加载 VSCode 以应用语言更改。',

        // Push
        'push.deleteSync': '（包含删除同步）',

        // Init command
        'init.tomlExists': 'cat5dev.toml 已存在。是否覆盖？',
        'init.overwrite': '覆盖',
        'init.gitignoreExists': '.gitignore 已存在。是否覆盖？',
        'init.success': '已创建 cat5dev.toml。',
    }
};

export function t(key: keyof typeof messages.ja, ...args: string[]): string {
    const lang = getLanguage();
    let text = (messages[lang] as any)[key] || messages.ja[key];

    // Simple string interpolation
    args.forEach((arg, index) => {
        text = text.replace(`{${index}}`, arg);
    });

    return text;
}
