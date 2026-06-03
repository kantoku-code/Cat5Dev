import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import { exec } from 'child_process';
import * as iconv from 'iconv-lite';
import { CatiaVbaTreeProvider } from './treeView';
import { VbaDocumentSymbolProvider } from './symbolProvider';
import { VbaDocumentFormatter, registerFormatOnSave, formatVbaDocument } from './formatter';
import { VbaServer } from './vbaServer';
import { t, getLanguage, setLanguage, Language, messages } from './i18n';
import { registerLinter } from './linter';
import { tomlTemplate, gitignoreTemplate, readProjectSettings, writeTomlProjectKey } from './lintConfig';
import { startLspClient } from './lspClient';

const outputChannel = vscode.window.createOutputChannel('CATIA VBA Sync');


/** CATScriptのログファイルを読み込み、OutputChannelに出力する。エラーがあればtrueを返す */
function flushCatScriptErrors(tempDir: string): boolean {
    const errLogPath = path.join(tempDir, 'c5d_err.log');
    let hasErrors = false;
    try {
        const content = fs.readFileSync(errLogPath, 'utf-8').trim();
        outputChannel.appendLine(`[CATScript Log]\n${content || '(空)'}`);
        outputChannel.show(true);
        hasErrors = /\[Push\.(Fatal|Add|DeleteLines|AddFromString)\]/.test(content);
    } catch {
        // ファイルが存在しない場合はスキップ
    }
    try { fs.unlinkSync(errLogPath); } catch (e) { }
    return hasErrors;
}

export function activate(context: vscode.ExtensionContext) {
    const workspaceRoot = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath ?? '';
    const vbaServer = new VbaServer(context, outputChannel);
    context.subscriptions.push(vbaServer);
    vbaServer.start();

    let pullCmd = vscode.commands.registerCommand('cat5dev.pullFromCatia', () => {
        executeCatiaPull(context, vbaServer);
    });

    let pushCmd = vscode.commands.registerCommand('cat5dev.pushToCatia', () => {
        executeCatiaPush(context);
    });

    let selectCmd = vscode.commands.registerCommand('cat5dev.selectProject', () => {
        executeSelectProject(context);
    });

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

    let renameFileCmd = vscode.commands.registerCommand('cat5dev.renameFile', async (fileUri: vscode.Uri) => {
        const filePath = fileUri.fsPath;
        const fileName = path.basename(filePath);
        const newName = await vscode.window.showInputBox({
            prompt: t('file.rename'),
            value: fileName
        });
        if (newName && newName !== fileName) {
            const newPath = path.join(path.dirname(filePath), newName);
            fs.renameSync(filePath, newPath);
            vscode.commands.executeCommand('cat5dev.refreshTree');
        }
    });

    let deleteFileCmd = vscode.commands.registerCommand('cat5dev.deleteFile', async (fileUri: vscode.Uri) => {
        const filePath = fileUri.fsPath;
        const fileName = path.basename(filePath);
        const deleteBtn = t('file.deleteButton');
        const confirmed = await vscode.window.showWarningMessage(
            t('file.deleteConfirm', fileName),
            { modal: true },
            deleteBtn
        );
        if (confirmed === deleteBtn) {
            fs.unlinkSync(filePath);
            vscode.commands.executeCommand('cat5dev.refreshTree');
        }
    });

    let copyPathCmd = vscode.commands.registerCommand('cat5dev.copyPath', async (fileUri: vscode.Uri) => {
        const filePath = fileUri.fsPath;
        await vscode.env.clipboard.writeText(filePath);
        vscode.window.showInformationMessage(t('file.copySuccess', filePath));
    });

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

    context.subscriptions.push(pullCmd, pushCmd, selectCmd, switchLanguageCmd, renameFileCmd, deleteFileCmd, copyPathCmd, initCmd);
}

export function deactivate() { }

async function getTargetProject(context: vscode.ExtensionContext, rootPath: string): Promise<string | undefined> {
    const { targetProject } = readProjectSettings(rootPath);
    if (targetProject) { return targetProject; }

    // Auto-prompt if empty
    return await executeSelectProject(context, rootPath) as string | undefined;
}

async function executeSelectProject(_context: vscode.ExtensionContext, rootPath?: string): Promise<string | undefined> {
    if (!rootPath) {
        const workspaceFolders = vscode.workspace.workspaceFolders;
        if (!workspaceFolders) {
            vscode.window.showErrorMessage(t('error.noWorkspace'));
            return undefined;
        }
        rootPath = workspaceFolders[0].uri.fsPath;
    }

    const tempDir = path.join(os.tmpdir(), 'cat5dev');
    if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
    }
    fs.mkdirSync(tempDir);

    const catScriptPath = path.join(tempDir, 'c5d_list.catvbs');
    const vbsScriptPath = path.join(tempDir, 'c5d_run_list.vbs');
    const outTxtPath = path.join(tempDir, 'projects.txt');

    const catScriptContent = `
Sub CATMain()
    On Error Resume Next
    Dim ag_fso, ag_apc, ag_vbe, ag_i, ag_outStr, ag_tmpProj, ag_px, ag_jx, ag_cx, ag_cName
    Dim ag_errPath, ag_ef

    Set ag_fso = CreateObject("Scripting.FileSystemObject")
    ag_errPath = "${tempDir}\\c5d_err.log"

    Set ag_apc = CreateObject("MSAPC.Apc.7.1")
    If ag_apc Is Nothing Then Set ag_apc = CreateObject("MSAPC.Apc")
    Set ag_vbe = ag_apc.VBE

    If Err.Number <> 0 Then Exit Sub

    Set ag_outStr = CreateObject("ADODB.Stream")
    ag_outStr.Type = 2
    ag_outStr.Charset = "shift_jis"
    ag_outStr.Open

    For ag_i = 1 To ag_vbe.VBProjects.Count
        ag_outStr.WriteText ag_vbe.VBProjects.Item(ag_i).Name & vbCrLf
    Next

    ag_outStr.SaveToFile "${outTxtPath}", 2
    ag_outStr.Close

    ' --- CLEANUP INJECTED MACROS ---
    For ag_px = 1 To ag_vbe.VBProjects.Count
        Set ag_tmpProj = ag_vbe.VBProjects.Item(ag_px)
        For ag_cx = ag_tmpProj.VBComponents.Count To 1 Step -1
            ag_cName = UCase(ag_tmpProj.VBComponents.Item(ag_cx).Name)
            If Left(ag_cName, 4) = "C5D_" Then
                On Error Resume Next
                ag_tmpProj.VBComponents.Remove ag_tmpProj.VBComponents.Item(ag_cx)
                If Err.Number <> 0 Then
                    Set ag_ef = ag_fso.OpenTextFile(ag_errPath, 8, True)
                    ag_ef.WriteLine "[List.Cleanup] Remove '" & ag_cName & "' Err=" & Err.Number & ": " & Err.Description
                    ag_ef.Close
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next
    Next
    ' --------------------------------
End Sub
`;
    fs.writeFileSync(catScriptPath, catScriptContent, 'utf-8');

    const vbsContent = `
On Error Resume Next
Dim ag_catia, ag_sys, ag_args()
Set ag_catia = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then WScript.Quit 1
Set ag_sys = ag_catia.SystemService
If Err.Number <> 0 Or ag_sys Is Nothing Then WScript.Quit 1
ag_sys.ExecuteScript "${tempDir}", 1, "c5d_list.catvbs", "CATMain", ag_args
`;
    fs.writeFileSync(vbsScriptPath, vbsContent, 'utf-8');

    return new Promise<string | undefined>((resolve, reject) => {
        exec(`%SystemRoot%\\SysWOW64\\cscript.exe //nologo "${vbsScriptPath}"`, async (error, stdout, stderr) => {
            if (fs.existsSync(vbsScriptPath)) fs.unlinkSync(vbsScriptPath);
            if (fs.existsSync(catScriptPath)) fs.unlinkSync(catScriptPath);

            flushCatScriptErrors(tempDir);

            if (error || !fs.existsSync(outTxtPath)) {
                const detail = stdout || stderr || (error ? error.message : 'Unknown error');
                outputChannel.appendLine(`[Select Target Project Error]`);
                outputChannel.appendLine(`Error: ${error?.message || 'None'}`);
                outputChannel.appendLine(`STDOUT: ${stdout}`);
                outputChannel.appendLine(`STDERR: ${stderr}`);
                outputChannel.show(true);

                vscode.window.showErrorMessage(t('error.selectFailed'));
                return resolve(undefined);
            }

            const buffer = fs.readFileSync(outTxtPath);
            const text = iconv.decode(buffer, 'shift_jis');
            const projects = text.split('\n').map(p => p.replace(/\r/g, '').trim()).filter(p => p.length > 0);

            fs.unlinkSync(outTxtPath);

            if (projects.length === 0) {
                vscode.window.showInformationMessage(t('info.projectNotFound'));
                return resolve(undefined);
            }

            const selected = await vscode.window.showQuickPick(projects, {
                placeHolder: t('select.placeholder')
            });

            if (selected) {
                writeTomlProjectKey(rootPath!, 'target_project', selected);
                vscode.window.showInformationMessage(t('info.projectSelected', selected));
            }
            resolve(selected);
        });
    });
}

async function executeCatiaPull(context: vscode.ExtensionContext, vbaServer: VbaServer) {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) {
        vscode.window.showErrorMessage(t('error.noWorkspace'));
        return;
    }
    const rootPath = workspaceFolders[0].uri.fsPath;

    const modulesDir = path.join(rootPath, 'modules');
    if (!fs.existsSync(modulesDir)) {
        fs.mkdirSync(modulesDir);
    } else {
        // Clean existing modules before pull
        const existingModules = fs.readdirSync(modulesDir);
        for (const f of existingModules) {
            if (f.endsWith('.bas_utf') || f.endsWith('.cls_utf') || f.endsWith('.frm_utf')) {
                fs.unlinkSync(path.join(modulesDir, f));
            }
        }
    }

    const targetProject = await getTargetProject(context, rootPath);
    if (!targetProject) return;

    const tempDir = path.join(os.tmpdir(), 'cat5dev');

    if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
    }
    fs.mkdirSync(tempDir);

    const catScriptPath = path.join(tempDir, 'c5d_pull.catvbs');
    const vbsScriptPath = path.join(tempDir, 'c5d_run_pull.vbs');

    // CATScript to extract all modules and write their code to temp files
    const catScriptContent = `
Sub CATMain()
    On Error Resume Next
    Dim ag_fso, ag_apc, ag_vbe, ag_i, ag_j, ag_proj, ag_comp, ag_codeMod, ag_lineCount, ag_outPath, ag_outStr, ag_devProj, ag_tmpProj, ag_px, ag_cx, ag_cName
    Dim ag_errPath, ag_ef

    Dim targetProjName
    targetProjName = "${targetProject}"

    Set ag_fso = CreateObject("Scripting.FileSystemObject")
    ag_errPath = "${tempDir}\\c5d_err.log"

    Set ag_apc = CreateObject("MSAPC.Apc.7.1")
    If ag_apc Is Nothing Then Set ag_apc = CreateObject("MSAPC.Apc")
    Set ag_vbe = ag_apc.VBE

    If Err.Number <> 0 Then
        ' Write error log
        Set ag_outStr = CreateObject("ADODB.Stream")
        ag_outStr.Type = 2
        ag_outStr.Charset = "shift_jis"
        ag_outStr.Open
        ag_outStr.WriteText "ERROR: VBE access failed"
        ag_outStr.SaveToFile "${tempDir}\\_error.log", 2
        ag_outStr.Close
        Exit Sub
    End If

    Set ag_devProj = Nothing
    For ag_i = 1 To ag_vbe.VBProjects.Count
        Set ag_proj = ag_vbe.VBProjects.Item(ag_i)
        If ag_proj.Name = targetProjName Then
            Set ag_devProj = ag_proj
            Exit For
        End If
    Next

    If ag_devProj Is Nothing Then Exit Sub

    For ag_j = 1 To ag_devProj.VBComponents.Count
        Set ag_comp = ag_devProj.VBComponents.Item(ag_j)
        Set ag_codeMod = ag_comp.CodeModule
        ag_lineCount = ag_codeMod.CountOfLines

        If ag_lineCount > 0 Then
            ag_outPath = "${tempDir}\\" & ag_comp.Name & "_TYPE_" & ag_comp.Type & ".txt"
            Set ag_outStr = CreateObject("ADODB.Stream")
            ag_outStr.Type = 2
            ag_outStr.Charset = "shift_jis"
            ag_outStr.Open
            ag_outStr.WriteText ag_codeMod.Lines(1, ag_lineCount)
            ag_outStr.SaveToFile ag_outPath, 2
            ag_outStr.Close
        End If
    Next

    ' --- CLEANUP INJECTED MACROS ---
    For ag_px = 1 To ag_vbe.VBProjects.Count
        Set ag_tmpProj = ag_vbe.VBProjects.Item(ag_px)
        For ag_cx = ag_tmpProj.VBComponents.Count To 1 Step -1
            ag_cName = UCase(ag_tmpProj.VBComponents.Item(ag_cx).Name)
            If Left(ag_cName, 4) = "C5D_" Then
                On Error Resume Next
                ag_tmpProj.VBComponents.Remove ag_tmpProj.VBComponents.Item(ag_cx)
                If Err.Number <> 0 Then
                    Set ag_ef = ag_fso.OpenTextFile(ag_errPath, 8, True)
                    ag_ef.WriteLine "[Pull.Cleanup] Remove '" & ag_cName & "' Err=" & Err.Number & ": " & Err.Description
                    ag_ef.Close
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next
    Next
    ' --------------------------------
End Sub
`;
    fs.writeFileSync(catScriptPath, catScriptContent, 'utf-8');

    const vbsContent = `
On Error Resume Next
Dim catia, sys, args()
Set catia = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then WScript.Quit 1
Set sys = catia.SystemService
If Err.Number <> 0 Or sys Is Nothing Then WScript.Quit 1
sys.ExecuteScript "${tempDir}", 1, "c5d_pull.catvbs", "CATMain", args
`;
    fs.writeFileSync(vbsScriptPath, vbsContent, 'utf-8');

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('progress.pull', targetProject),
        cancellable: false
    }, async (progress) => {
        return new Promise<void>((resolve, reject) => {
            exec(`%SystemRoot%\\SysWOW64\\cscript.exe //nologo "${vbsScriptPath}"`, async (error, stdout, stderr) => {
                if (fs.existsSync(vbsScriptPath)) fs.unlinkSync(vbsScriptPath);
                if (fs.existsSync(catScriptPath)) fs.unlinkSync(catScriptPath);

                flushCatScriptErrors(tempDir);

                if (error) {
                    outputChannel.appendLine(`[Pull Error]`);
                    outputChannel.appendLine(`Error: ${error.message}`);
                    outputChannel.appendLine(`STDOUT: ${stdout}`);
                    outputChannel.appendLine(`STDERR: ${stderr}`);
                    outputChannel.show(true);
                    vscode.window.showErrorMessage(t('error.pullFailed'));
                    return reject();
                }

                // Process output text files in TempDir
                const files = fs.readdirSync(tempDir);
                let count = 0;
                for (const file of files) {
                    if (file.endsWith('.txt') && file.includes('_TYPE_')) {
                        const parts = file.replace('.txt', '').split('_TYPE_');
                        const compType = parts.pop();
                        const compName = parts.join('_TYPE_');

                        let ext = '.bas_utf'; // Standard module (1) or default
                        if (compType === '2') ext = '.cls_utf'; // Class module
                        else if (compType === '3') ext = '.frm_utf'; // Userform

                        const shiftJisBuffer = fs.readFileSync(path.join(tempDir, file));
                        const utf8String = iconv.decode(shiftJisBuffer, 'shift_jis');

                        // Normalize newlines: Remove all trailing newlines/spaces and ensure exactly one LF
                        const normalized = utf8String.replace(/\r/g, '').trimEnd() + '\n';

                        // formatOnPull が有効な場合はフォーマットしてから保存
                        const fmtConfig = vscode.workspace.getConfiguration('cat5dev.formatter');
                        let saveContent = normalized;
                        if (fmtConfig.get<boolean>('formatOnPull', false)) {
                            const formatted = await formatVbaDocument(normalized, vbaServer, outputChannel, rootPath);
                            if (formatted !== null) {
                                saveContent = formatted;
                            }
                        }

                        // Save to modules directory
                        fs.writeFileSync(path.join(modulesDir, compName + ext), saveContent, 'utf-8');
                        count++;

                        // Cleanup
                        fs.unlinkSync(path.join(tempDir, file));
                    }
                }

                vscode.window.showInformationMessage(t('info.pullSuccess', String(count)));
                vscode.commands.executeCommand('cat5dev.refreshTree');
                resolve();
            });
        });
    });
}

async function executeCatiaPush(context: vscode.ExtensionContext) {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) {
        vscode.window.showErrorMessage(t('error.noWorkspace'));
        return;
    }
    const rootPath = workspaceFolders[0].uri.fsPath;

    const modulesDir = path.join(rootPath, 'modules');
    if (!fs.existsSync(modulesDir)) {
        vscode.window.showInformationMessage(t('error.noModulesDir'));
        return;
    }

    const targetProject = await getTargetProject(context, rootPath);
    if (!targetProject) return;

    const tempDir = path.join(os.tmpdir(), 'cat5dev');

    if (fs.existsSync(tempDir)) {
        fs.rmSync(tempDir, { recursive: true, force: true });
    }
    fs.mkdirSync(tempDir);

    // 1. Prepare local files into tempDir and collect local component names
    const files = fs.readdirSync(modulesDir);
    let count = 0;
    let skippedCount = 0;
    const localCompNames: string[] = [];
    const localContents: Record<string, string> = {};
    const localCompTypes: Record<string, string> = {};
    const longNames: string[] = [];

    for (const file of files) {
        if (file.endsWith('.bas_utf') || file.endsWith('.cls_utf') || file.endsWith('.frm_utf')) {
            const compName = file.substring(0, file.lastIndexOf('.')); // Strip extension
            localCompNames.push(compName);

            if (compName.length > 31) { longNames.push(compName); }

            let compType = '1';
            if (file.endsWith('.cls_utf')) compType = '2';
            else if (file.endsWith('.frm_utf')) compType = '3';

            const utf8Buffer = fs.readFileSync(path.join(modulesDir, file));

            // Decode from utf-8 and re-encode to Shift-JIS for CATIA
            const utf8String = utf8Buffer.toString('utf-8');

            // Normalize for CATIA: Remove trailing newlines to prevent multiplication via AddFromString
            const trimmed = utf8String.trimEnd();

            localContents[compName] = trimmed;
            localCompTypes[compName] = compType;

            const shiftJisBuffer = iconv.encode(trimmed, 'shift_jis');

            const tempFilePath = path.join(tempDir, `${compName}_TYPE_${compType}.txt`);
            fs.writeFileSync(tempFilePath, shiftJisBuffer);
            count++;
        }
    }

    // B. 31文字超モジュール名の警告
    if (longNames.length > 0) {
        const longResp = await vscode.window.showWarningMessage(
            t('warning.longModuleNames', longNames.join(', ')),
            { modal: true },
            t('dialog.continue')
        );
        if (longResp !== t('dialog.continue')) {
            return;
        }
    }

    if (count === 0) {
        vscode.window.showInformationMessage(t('error.noModuleFiles'));
        return;
    }

    // 2. Fetch remote components to compute diff for deletion
    const remoteCompsFile = path.join(tempDir, 'remote_comps.txt');
    const checkCatScript = `
Sub CATMain()
    On Error Resume Next
    Dim ag_fso, ag_apc, ag_vbe, ag_i, ag_j, ag_comp, ag_outStr, ag_codeStr, ag_devProj, ag_tmpProj, ag_px, ag_cx, ag_cName
    Dim ag_errPath, ag_ef, ag_lineCount
    Set ag_fso = CreateObject("Scripting.FileSystemObject")
    ag_errPath = "${tempDir}\\c5d_err.log"

    Set ag_apc = CreateObject("MSAPC.Apc.7.1")
    If ag_apc Is Nothing Then Set ag_apc = CreateObject("MSAPC.Apc")
    Set ag_vbe = ag_apc.VBE

    If Err.Number <> 0 Then
        Set ag_ef = ag_fso.OpenTextFile(ag_errPath, 8, True)
        ag_ef.WriteLine "[Check] VBE access failed. Err=" & Err.Number & ": " & Err.Description
        ag_ef.Close
        Exit Sub
    End If

    Set ag_devProj = Nothing
    For ag_i = 1 To ag_vbe.VBProjects.Count
        If ag_vbe.VBProjects.Item(ag_i).Name = "${targetProject}" Then
            Set ag_devProj = ag_vbe.VBProjects.Item(ag_i)
            Exit For
        End If
    Next
    If ag_devProj Is Nothing Then
        Set ag_ef = ag_fso.OpenTextFile(ag_errPath, 8, True)
        ag_ef.WriteLine "[Check] Project '${targetProject}' not found in CATIA. Total VBProjects: " & ag_vbe.VBProjects.Count
        ag_ef.Close
        Exit Sub
    End If

    Set ag_outStr = CreateObject("ADODB.Stream")
    ag_outStr.Type = 2
    ag_outStr.Charset = "shift_jis"
    ag_outStr.Open

    For ag_j = 1 To ag_devProj.VBComponents.Count
        Set ag_comp = ag_devProj.VBComponents.Item(ag_j)
        If ag_comp.Type = 1 Or ag_comp.Type = 2 Or ag_comp.Type = 3 Then
            ag_outStr.WriteText ag_comp.Name & vbCrLf
            ag_lineCount = ag_comp.CodeModule.CountOfLines
            If ag_lineCount > 0 Then
                Set ag_codeStr = CreateObject("ADODB.Stream")
                ag_codeStr.Type = 2
                ag_codeStr.Charset = "shift_jis"
                ag_codeStr.Open
                ag_codeStr.WriteText ag_comp.CodeModule.Lines(1, ag_lineCount)
                ag_codeStr.SaveToFile "${tempDir}\\" & ag_comp.Name & "_REMOTE.txt", 2
                ag_codeStr.Close
            End If
        End If
    Next

    ag_outStr.SaveToFile "${remoteCompsFile}", 2
    ag_outStr.Close

    ' --- CLEANUP INJECTED MACROS ---
    For ag_px = 1 To ag_vbe.VBProjects.Count
        Set ag_tmpProj = ag_vbe.VBProjects.Item(ag_px)
        For ag_cx = ag_tmpProj.VBComponents.Count To 1 Step -1
            ag_cName = UCase(ag_tmpProj.VBComponents.Item(ag_cx).Name)
            If Left(ag_cName, 4) = "C5D_" Then
                On Error Resume Next
                ag_tmpProj.VBComponents.Remove ag_tmpProj.VBComponents.Item(ag_cx)
                If Err.Number <> 0 Then
                    Set ag_ef = ag_fso.OpenTextFile(ag_errPath, 8, True)
                    ag_ef.WriteLine "[Check.Cleanup] Remove '" & ag_cName & "' Err=" & Err.Number & ": " & Err.Description
                    ag_ef.Close
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        Next
    Next
    ' --------------------------------
End Sub
`;
    fs.writeFileSync(path.join(tempDir, 'c5d_check.catvbs'), checkCatScript, 'utf-8');

    const checkVbsPath = path.join(tempDir, 'c5d_run_check.vbs');
    const checkVbsScript = `
On Error Resume Next
Dim ag_catia, ag_sys, ag_args()
Set ag_catia = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then WScript.Quit 1
Set ag_sys = ag_catia.SystemService
If Err.Number <> 0 Or ag_sys Is Nothing Then WScript.Quit 1
ag_sys.ExecuteScript "${tempDir}", 1, "c5d_check.catvbs", "CATMain", ag_args
`;
    fs.writeFileSync(checkVbsPath, checkVbsScript, 'utf-8');

    // Run check synchronously within an await
    await new Promise<void>((resolve) => {
        exec(`%SystemRoot%\\SysWOW64\\cscript.exe //nologo "${checkVbsPath}"`, (error, stdout, stderr) => {
            if (error) {
                outputChannel.appendLine(`[Check Components Error]`);
                outputChannel.appendLine(`Error: ${error.message}`);
                outputChannel.appendLine(`STDOUT: ${stdout}`);
                outputChannel.appendLine(`STDERR: ${stderr}`);
                outputChannel.show(true);
            }
            if (fs.existsSync(checkVbsPath)) fs.unlinkSync(checkVbsPath);
            if (fs.existsSync(path.join(tempDir, 'c5d_check.catvbs'))) fs.unlinkSync(path.join(tempDir, 'c5d_check.catvbs'));

            flushCatScriptErrors(tempDir);
            resolve();
        });
    });

    let remoteCompNames: string[] = [];
    if (fs.existsSync(remoteCompsFile)) {
        const buffer = fs.readFileSync(remoteCompsFile);
        const text = iconv.decode(buffer, 'shift_jis');
        remoteCompNames = text.split('\n').map(p => p.replace(/\r/g, '').trim()).filter(p => p.length > 0);
        fs.unlinkSync(remoteCompsFile);
    }

    // C4. リモート内容と一致するモジュールをスキップ
    for (const compName of localCompNames) {
        const remoteFilePath = path.join(tempDir, `${compName}_REMOTE.txt`);
        if (fs.existsSync(remoteFilePath)) {
            const remoteBuf = fs.readFileSync(remoteFilePath);
            const remoteText = iconv.decode(remoteBuf, 'shift_jis').trimEnd();
            fs.unlinkSync(remoteFilePath);

            if (localContents[compName] === remoteText) {
                const typeFilePath = path.join(tempDir, `${compName}_TYPE_${localCompTypes[compName]}.txt`);
                if (fs.existsSync(typeFilePath)) {
                    fs.unlinkSync(typeFilePath);
                    count--;
                    skippedCount++;
                }
            }
        }
    }
    // 残留 _REMOTE.txt を削除
    for (const f of fs.readdirSync(tempDir)) {
        if (f.endsWith('_REMOTE.txt')) { fs.unlinkSync(path.join(tempDir, f)); }
    }

    if (count === 0) {
        vscode.window.showInformationMessage(t('info.noChanges', String(skippedCount)));
        return;
    }

    // 3. Prompt user if there are files in CATIA missing locally
    const toDelete = remoteCompNames.filter(r => !localCompNames.includes(r));
    let performDelete = false;

    if (toDelete.length > 0) {
        const resp = await vscode.window.showWarningMessage(
            t('warning.deleteModules', toDelete.join(', ')),
            { modal: true },
            t('dialog.delete'), t('dialog.keep')
        );
        if (resp === undefined) {
            // Abort push if they cancelled modal
            vscode.window.showInformationMessage(t('info.pushCancelled'));
            return;
        }
        if (resp === t('dialog.delete')) {
            performDelete = true;
            const delListShiftJis = iconv.encode(toDelete.join('\r\n'), 'shift_jis');
            fs.writeFileSync(path.join(tempDir, 'delete_list.txt'), delListShiftJis);
        }
    }

    // 3.5 Check for new UserForms (cannot be created via Push)
    const newForms = files.filter(f => f.endsWith('.frm_utf')).map(f => f.substring(0, f.lastIndexOf('.'))).filter(name => !remoteCompNames.includes(name));
    if (newForms.length > 0) {
        vscode.window.showWarningMessage(
            t('warning.newUserForms', newForms.join(', '))
        );
        // Remove from tempDir so they are not pushed
        for (const name of newForms) {
            const pattern = `${name}_TYPE_3.txt`;
            const p = path.join(tempDir, pattern);
            if (fs.existsSync(p)) {
                fs.unlinkSync(p);
                count--;
            }
        }
        if (count === 0) {
            vscode.window.showInformationMessage(t('warning.noMoreFiles'));
            return;
        }
    }

    // 4. Execute Push
    const catScriptPath = path.join(tempDir, 'c5d_push.catvbs');
    const pushVbsPath = path.join(tempDir, 'c5d_run_push.vbs');

    // CATScript to read txt files and push them into target project modules
    // ログ書き込みはADODB.Streamを使用（FSO書き込みがCATIA環境で動作しない場合があるため）
    const catScriptContent = `
Sub CATMain()
    On Error Resume Next
    Dim fso, apc, vbe, i, j, proj, comp, codeMod, inStr, newContent, devProj
    Dim targetProjName, folder, fileItem, parts, compName, compType
    Dim errPath, donePath
    Dim logCount, errCount
    Dim logStream

    targetProjName = "${targetProject}"
    errPath = "${tempDir}\\c5d_err.log"
    donePath = "${tempDir}\\c5d_push_done.txt"
    logCount = 0
    errCount = 0

    ' ログをADODB.Streamで蓄積し、最後にまとめて書き出す
    Set logStream = CreateObject("ADODB.Stream")
    logStream.Type = 2
    logStream.Charset = "utf-8"
    logStream.Open
    logStream.WriteText "[Push.Start] target=" & targetProjName & " time=" & Now & vbCrLf

    Set fso = CreateObject("Scripting.FileSystemObject")
    logStream.WriteText "[Push.Debug] fso=" & Not (fso Is Nothing) & vbCrLf

    Set apc = CreateObject("MSAPC.Apc.7.1")
    logStream.WriteText "[Push.Debug] apc.7.1 isNull=" & (apc Is Nothing) & " Err=" & Err.Number & vbCrLf
    Err.Clear
    If apc Is Nothing Then
        Set apc = CreateObject("MSAPC.Apc.6.2")
        logStream.WriteText "[Push.Debug] apc.6.2 isNull=" & (apc Is Nothing) & " Err=" & Err.Number & vbCrLf
        Err.Clear
    End If
    If apc Is Nothing Then
        Set apc = CreateObject("MSAPC.Apc")
        logStream.WriteText "[Push.Debug] apc isNull=" & (apc Is Nothing) & " Err=" & Err.Number & vbCrLf
        Err.Clear
    End If
    If apc Is Nothing Then
        logStream.WriteText "[Push.Fatal] APC取得失敗 Err=" & Err.Number & ": " & Err.Description & vbCrLf
        logStream.SaveToFile errPath, 2
        logStream.Close
        Exit Sub
    End If
    Set vbe = apc.VBE
    logStream.WriteText "[Push.Debug] vbe isNull=" & (vbe Is Nothing) & " Err=" & Err.Number & vbCrLf
    If Err.Number <> 0 Then
        logStream.WriteText "[Push.Fatal] VBE取得失敗 Err=" & Err.Number & ": " & Err.Description & vbCrLf
        logStream.SaveToFile errPath, 2
        logStream.Close
        Exit Sub
    End If
    Err.Clear

    Set devProj = Nothing
    For i = 1 To vbe.VBProjects.Count
        Set proj = vbe.VBProjects.Item(i)
        If proj.Name = targetProjName Then
            Set devProj = proj
            Exit For
        End If
    Next
    logStream.WriteText "[Push.Debug] devProj=" & Not (devProj Is Nothing) & " projects=" & vbe.VBProjects.Count & vbCrLf

    If devProj Is Nothing Then
        logStream.WriteText "[Push.Fatal] プロジェクト '" & targetProjName & "' が見つかりません" & vbCrLf
        logStream.SaveToFile errPath, 2
        logStream.Close
        Exit Sub
    End If

    ' Perform Deletions
    If fso.FileExists("${tempDir}\\delete_list.txt") Then
        Set inStr = CreateObject("ADODB.Stream")
        inStr.Type = 2
        inStr.Charset = "shift_jis"
        inStr.Open
        inStr.LoadFromFile "${tempDir}\\delete_list.txt"

        Dim delNames, d, k
        delNames = Split(inStr.ReadText, vbCrLf)
        inStr.Close

        For Each d In delNames
            If Trim(d) <> "" Then
                Set comp = Nothing
                For k = 1 To devProj.VBComponents.Count
                    If UCase(devProj.VBComponents.Item(k).Name) = UCase(Trim(d)) Then
                        Set comp = devProj.VBComponents.Item(k)
                        Exit For
                    End If
                Next
                If Not comp Is Nothing Then
                    On Error Resume Next
                    devProj.VBComponents.Remove comp
                    If Err.Number <> 0 Then
                        logStream.WriteText "[Push.Delete] Remove '" & Trim(d) & "' Err=" & Err.Number & ": " & Err.Description & vbCrLf
                        errCount = errCount + 1
                        Err.Clear
                    Else
                        logStream.WriteText "[Push.Delete] Remove '" & Trim(d) & "' -> OK" & vbCrLf
                    End If
                    On Error Resume Next
                End If
            End If
        Next
        fso.DeleteFile "${tempDir}\\delete_list.txt"
    End If

    Set folder = fso.GetFolder("${tempDir}")
    logStream.WriteText "[Push.Debug] folder=" & Not (folder Is Nothing) & vbCrLf

    ' ループ中に削除するとコレクションが狂うため、先にパスを配列収集
    Dim filePaths(), fileCount, fi, fp
    fileCount = 0
    If Not (folder Is Nothing) Then
        ReDim filePaths(folder.Files.Count - 1)
        logStream.WriteText "[Push.Debug] tempDir files=" & folder.Files.Count & vbCrLf
        For Each fileItem In folder.Files
            filePaths(fileCount) = fileItem.Path
            fileCount = fileCount + 1
        Next
    End If
    logStream.WriteText "[Push.Debug] collected=" & fileCount & vbCrLf

    For fi = 0 To fileCount - 1
        fp = filePaths(fi)
        If fso.FileExists(fp) Then
            If UCase(fso.GetExtensionName(fp)) = "TXT" And InStr(fso.GetFileName(fp), "_TYPE_") > 0 Then
                parts = Split(fso.GetBaseName(fp), "_TYPE_")
                compName = parts(0)
                compType = CInt(parts(1))

                ' C5D_内部ファイルや無効なTypeはスキップして削除
                If Left(UCase(compName), 4) = "C5D_" Or compType < 1 Or compType > 3 Then
                    fso.DeleteFile fp
                Else
                    logStream.WriteText "[Push.Process] " & compName & " (Type=" & compType & ")" & vbCrLf

                    Set comp = Nothing
                    For k = 1 To devProj.VBComponents.Count
                        If UCase(devProj.VBComponents.Item(k).Name) = UCase(compName) Then
                            Set comp = devProj.VBComponents.Item(k)
                            Exit For
                        End If
                    Next

                    If comp Is Nothing Then
                        On Error Resume Next
                        Set comp = devProj.VBComponents.Add(compType)
                        If Err.Number <> 0 Then
                            logStream.WriteText "[Push.Add] Add '" & compName & "' Err=" & Err.Number & ": " & Err.Description & vbCrLf
                            errCount = errCount + 1
                            Err.Clear
                        End If
                        On Error Resume Next
                        If Not comp Is Nothing Then comp.Name = compName
                    End If

                    If comp Is Nothing Then
                        logStream.WriteText "[Push.Add] Component '" & compName & "' could not be created or found. Skipped." & vbCrLf
                        errCount = errCount + 1
                    Else
                        Set inStr = CreateObject("ADODB.Stream")
                        inStr.Type = 2
                        inStr.Charset = "shift_jis"
                        inStr.Open
                        inStr.LoadFromFile fp
                        newContent = inStr.ReadText
                        inStr.Close

                        Set codeMod = comp.CodeModule
                        On Error Resume Next
                        If codeMod.CountOfLines > 0 Then
                            codeMod.DeleteLines 1, codeMod.CountOfLines
                            If Err.Number <> 0 Then
                                logStream.WriteText "[Push.DeleteLines] '" & compName & "' Err=" & Err.Number & ": " & Err.Description & vbCrLf
                                errCount = errCount + 1
                                Err.Clear
                            End If
                        End If

                        codeMod.AddFromString newContent
                        If Err.Number <> 0 Then
                            logStream.WriteText "[Push.AddFromString] '" & compName & "' Err=" & Err.Number & ": " & Err.Description & vbCrLf
                            errCount = errCount + 1
                            Err.Clear
                        Else
                            logCount = logCount + 1
                        End If
                        On Error Resume Next
                    End If

                    fso.DeleteFile fp
                End If
            End If
        End If
    Next

    logStream.WriteText "[Push.Done] processed=" & logCount & " errors=" & errCount & " time=" & Now & vbCrLf
    logStream.SaveToFile errPath, 2
    logStream.Close

    ' 完了フラグ（ADODB.Streamで書き込み）
    Dim doneStream
    Set doneStream = CreateObject("ADODB.Stream")
    doneStream.Type = 2
    doneStream.Charset = "utf-8"
    doneStream.Open
    doneStream.WriteText "done" & vbCrLf
    doneStream.SaveToFile donePath, 2
    doneStream.Close

End Sub
`;
    fs.writeFileSync(catScriptPath, catScriptContent, 'utf-8');

    outputChannel.appendLine(`[Push] tempDir: ${tempDir}`);
    outputChannel.appendLine(`[Push] 対象プロジェクト: ${targetProject}`);
    outputChannel.appendLine(`[Push] 転送ファイル数: ${count}`);
    outputChannel.appendLine(`[Push] CATScript生成: ${catScriptPath}`);
    outputChannel.show(true);

    const doneFlagPath = path.join(tempDir, 'c5d_push_done.txt');
    // 既存の完了フラグがあれば削除しておく
    if (fs.existsSync(doneFlagPath)) { fs.unlinkSync(doneFlagPath); }

    const pushVbsScript = `
On Error Resume Next
Dim ag_catia, ag_sys, ag_args()
WScript.Echo "VBS: Start"
Set ag_catia = GetObject(, "CATIA.Application")
WScript.Echo "VBS: CATIA Err=" & Err.Number
If Err.Number <> 0 Then WScript.Quit 1
Err.Clear
Set ag_sys = ag_catia.SystemService
WScript.Echo "VBS: SystemService Err=" & Err.Number
If Err.Number <> 0 Or ag_sys Is Nothing Then WScript.Quit 1
Err.Clear
WScript.Echo "VBS: Calling ExecuteScript..."
ag_sys.ExecuteScript "${tempDir}", 1, "c5d_push.catvbs", "CATMain", ag_args
WScript.Echo "VBS: ExecuteScript returned Err=" & Err.Number
Dim fso2, doneFile, startTime
Set fso2 = CreateObject("Scripting.FileSystemObject")
doneFile = "${tempDir}\\c5d_push_done.txt"
WScript.Echo "VBS: doneFile exists=" & fso2.FileExists(doneFile)
startTime = Now
Do While Not fso2.FileExists(doneFile)
    WScript.Sleep 500
    If DateDiff("s", startTime, Now) > 60 Then
        WScript.Echo "VBS: Polling timeout"
        Exit Do
    End If
Loop
WScript.Echo "VBS: Poll done, doneFile exists=" & fso2.FileExists(doneFile)
If fso2.FileExists(doneFile) Then fso2.DeleteFile doneFile
WScript.Echo "VBS: End"
`;
    fs.writeFileSync(pushVbsPath, pushVbsScript, 'utf-8');
    outputChannel.appendLine(`[Push] VBS生成: ${pushVbsPath}`);
    outputChannel.appendLine(`[Push] cscript実行開始...`);
    outputChannel.show(true);

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('progress.push', targetProject),
        cancellable: false
    }, async (progress) => {
        return new Promise<void>((resolve, reject) => {
            exec(`%SystemRoot%\\SysWOW64\\cscript.exe //nologo "${pushVbsPath}"`, { maxBuffer: 1024 * 1024 * 10 }, (error, stdout, stderr) => {
                outputChannel.appendLine(`[Push] cscript終了 error=${error?.code ?? 'null'} stdout="${stdout.trim()}" stderr="${stderr.trim()}"`);
                outputChannel.show(true);
                if (fs.existsSync(pushVbsPath)) fs.unlinkSync(pushVbsPath);
                if (fs.existsSync(catScriptPath)) fs.unlinkSync(catScriptPath);

                const hasErrors = flushCatScriptErrors(tempDir);

                if (error) {
                    outputChannel.appendLine(`[Push Error]`);
                    outputChannel.appendLine(`Error: ${error.message}`);
                    outputChannel.appendLine(`STDOUT: ${stdout}`);
                    outputChannel.appendLine(`STDERR: ${stderr}`);
                    outputChannel.show(true);
                    vscode.window.showErrorMessage(t('error.pushFailed'));
                    return reject();
                }
                if (hasErrors) {
                    outputChannel.show(true);
                }
                const deleteMsg = performDelete ? t('push.deleteSync') : '';
                vscode.window.showInformationMessage(t('info.pushSuccess', String(count), '', deleteMsg));
                vscode.commands.executeCommand('cat5dev.refreshTree');
                resolve();
            });
        });
    });
}
