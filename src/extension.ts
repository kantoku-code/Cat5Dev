import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import * as crypto from 'crypto';
import { exec } from 'child_process';
import * as iconv from 'iconv-lite';
import { CatiaVbaTreeProvider } from './treeView';
import { VbaDocumentSymbolProvider } from './symbolProvider';
import { t, getLanguage, setLanguage } from './i18n';

const outputChannel = vscode.window.createOutputChannel('CATIA VBA Sync');

function ensureGitignore(rootPath: string): void {
    const gitignorePath = path.join(rootPath, '.gitignore');
    const cacheEntry = '.catia-vba-push-cache.json';

    if (!fs.existsSync(gitignorePath)) {
        fs.writeFileSync(gitignorePath, `${cacheEntry}\n`, 'utf-8');
        return;
    }

    const content = fs.readFileSync(gitignorePath, 'utf-8');
    if (!content.includes(cacheEntry)) {
        const sep = content.endsWith('\n') ? '' : '\n';
        fs.appendFileSync(gitignorePath, `${sep}${cacheEntry}\n`, 'utf-8');
    }
}

/** CATScriptのエラーログファイルを読み込み、OutputChannelに出力する */
function flushCatScriptErrors(tempDir: string): void {
    const errLogPath = path.join(tempDir, 'c5d_err.log');
    if (!fs.existsSync(errLogPath)) { return; }
    try {
        const content = fs.readFileSync(errLogPath, 'utf-8').trim();
        if (content) {
            outputChannel.appendLine(`[CATScript Errors]\n${content}`);
            outputChannel.show(true);
        }
    } catch (e) { }
    try { fs.unlinkSync(errLogPath); } catch (e) { }
}

export function activate(context: vscode.ExtensionContext) {
    let pullCmd = vscode.commands.registerCommand('cat5dev.pullFromCatia', () => {
        executeCatiaPull(context);
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
                { label: t('language.japanese'), description: 'Current language', value: 'ja' },
                { label: t('language.english'), description: 'Current language', value: 'en' }
            ],
            { placeHolder: t('language.title') }
        );
        if (selected && selected.value !== currentLang) {
            setLanguage(selected.value as 'ja' | 'en');
            vscode.window.showInformationMessage('Please reload VSCode to apply language changes.');
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
        const confirmed = await vscode.window.showWarningMessage(
            `Delete ${fileName}?`,
            { modal: true },
            'Delete'
        );
        if (confirmed === 'Delete') {
            fs.unlinkSync(filePath);
            vscode.commands.executeCommand('cat5dev.refreshTree');
        }
    });

    let copyPathCmd = vscode.commands.registerCommand('cat5dev.copyPath', async (fileUri: vscode.Uri) => {
        const filePath = fileUri.fsPath;
        await vscode.env.clipboard.writeText(filePath);
        vscode.window.showInformationMessage(`Path copied: ${filePath}`);
    });

    const VBA_SELECTOR = ['bas_utf', 'cls_utf', 'frm_utf'].map(l => ({ language: 'vb', pattern: `**/*.${l}` }));
    context.subscriptions.push(
        vscode.languages.registerDocumentSymbolProvider(VBA_SELECTOR, new VbaDocumentSymbolProvider())
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

    // Auto-refresh tree when .vscode/settings.json changes (e.g. after selectProject)
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (workspaceFolders) {
        const configWatcher = vscode.workspace.createFileSystemWatcher(
            new vscode.RelativePattern(workspaceFolders[0], '.vscode/settings.json')
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

    context.subscriptions.push(pullCmd, pushCmd, selectCmd, switchLanguageCmd, renameFileCmd, deleteFileCmd, copyPathCmd);
}

export function deactivate() { }

async function getTargetProject(context: vscode.ExtensionContext, rootPath: string): Promise<string | undefined> {
    const settingsPath = path.join(rootPath, '.vscode', 'settings.json');
    if (fs.existsSync(settingsPath)) {
        try {
            const settings = JSON.parse(fs.readFileSync(settingsPath, 'utf-8'));
            if (settings.targetProject) {
                return settings.targetProject;
            }
        } catch (e) { }
    }

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

    ensureGitignore(rootPath);

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
                const vscodePath = path.join(rootPath!, '.vscode');
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

                settings.targetProject = selected;
                fs.writeFileSync(settingsPath, JSON.stringify(settings, null, 4), 'utf-8');
                vscode.window.showInformationMessage(t('info.projectSelected', selected));
            }
            resolve(selected);
        });
    });
}

async function executeCatiaPull(context: vscode.ExtensionContext) {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) {
        vscode.window.showErrorMessage(t('error.noWorkspace'));
        return;
    }
    const rootPath = workspaceFolders[0].uri.fsPath;

    ensureGitignore(rootPath);

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
            exec(`%SystemRoot%\\SysWOW64\\cscript.exe //nologo "${vbsScriptPath}"`, (error, stdout, stderr) => {
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

                        // Save to modules directory
                        fs.writeFileSync(path.join(modulesDir, compName + ext), normalized, 'utf-8');
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

    ensureGitignore(rootPath);

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

    // Load push cache (stores SHA-256 hash of last successfully pushed content per module, keyed by project)
    const cachePath = path.join(rootPath, '.catia-vba-push-cache.json');
    let allProjectsCache: Record<string, Record<string, string>> = {};
    if (fs.existsSync(cachePath)) {
        try {
            allProjectsCache = JSON.parse(fs.readFileSync(cachePath, 'utf-8'));
        } catch (e) { allProjectsCache = {}; }
    }
    let pushCache: Record<string, string> = allProjectsCache[targetProject] ?? {};
    // Track hashes of modules being pushed this session (for cache update on success)
    const pendingCacheUpdates: Record<string, string> = {};

    for (const file of files) {
        if (file.endsWith('.bas_utf') || file.endsWith('.cls_utf') || file.endsWith('.frm_utf')) {
            const compName = file.substring(0, file.lastIndexOf('.')); // Strip extension
            localCompNames.push(compName);

            let compType = '1';
            if (file.endsWith('.cls_utf')) compType = '2';
            else if (file.endsWith('.frm_utf')) compType = '3';

            const utf8Buffer = fs.readFileSync(path.join(modulesDir, file));

            // Decode from utf-8 and re-encode to Shift-JIS for CATIA
            const utf8String = utf8Buffer.toString('utf-8');

            // Normalize for CATIA: Remove trailing newlines to prevent multiplication via AddFromString
            const trimmed = utf8String.trimEnd();

            // Compute hash of normalized content to detect changes
            const hash = crypto.createHash('sha256').update(trimmed, 'utf-8').digest('hex');
            pendingCacheUpdates[compName] = hash;

            // Skip modules whose content hasn't changed since last successful push
            if (pushCache[compName] === hash) {
                skippedCount++;
                continue;
            }

            const shiftJisBuffer = iconv.encode(trimmed, 'shift_jis');

            const tempFilePath = path.join(tempDir, `${compName}_TYPE_${compType}.txt`);
            fs.writeFileSync(tempFilePath, shiftJisBuffer);
            count++;
        }
    }

    if (count === 0 && skippedCount > 0) {
        vscode.window.showInformationMessage(t('info.noChanges', String(skippedCount)));
        return;
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
    Dim ag_fso, ag_apc, ag_vbe, ag_i, ag_j, ag_comp, ag_outStr, ag_devProj, ag_tmpProj, ag_px, ag_cx, ag_cName
    Dim ag_errPath, ag_ef
    Set ag_fso = CreateObject("Scripting.FileSystemObject")
    ag_errPath = "${tempDir}\\c5d_err.log"

    Set ag_apc = CreateObject("MSAPC.Apc.7.1")
    If ag_apc Is Nothing Then Set ag_apc = CreateObject("MSAPC.Apc")
    Set ag_vbe = ag_apc.VBE

    If Err.Number <> 0 Then Exit Sub

    Set ag_devProj = Nothing
    For ag_i = 1 To ag_vbe.VBProjects.Count
        If ag_vbe.VBProjects.Item(ag_i).Name = "${targetProject}" Then
            Set ag_devProj = ag_vbe.VBProjects.Item(ag_i)
            Exit For
        End If
    Next
    If ag_devProj Is Nothing Then Exit Sub

    Set ag_outStr = CreateObject("ADODB.Stream")
    ag_outStr.Type = 2
    ag_outStr.Charset = "shift_jis"
    ag_outStr.Open

    For ag_j = 1 To ag_devProj.VBComponents.Count
        Set ag_comp = ag_devProj.VBComponents.Item(ag_j)
        If ag_comp.Type = 1 Or ag_comp.Type = 2 Or ag_comp.Type = 3 Then
            ag_outStr.WriteText ag_comp.Name & vbCrLf
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
    const catScriptContent = `
Sub CATMain()
    On Error Resume Next
    Dim fso, apc, vbe, i, j, proj, comp, codeMod, inPath, inStr, newContent, devProj
    Dim targetProjName, folder, fileItem, parts, compName, compType
    Dim errPath, ef

    targetProjName = "${targetProject}"

    Set fso = CreateObject("Scripting.FileSystemObject")
    errPath = "${tempDir}\\c5d_err.log"

    Set apc = CreateObject("MSAPC.Apc.7.1")
    If apc Is Nothing Then Set apc = CreateObject("MSAPC.Apc")
    Set vbe = apc.VBE

    If Err.Number <> 0 Then
        Exit Sub
    End If

    Set devProj = Nothing
    For i = 1 To vbe.VBProjects.Count
        Set proj = vbe.VBProjects.Item(i)
        If proj.Name = targetProjName Then
            Set devProj = proj
            Exit For
        End If
    Next

    If devProj Is Nothing Then Exit Sub

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
                        Set ef = fso.OpenTextFile(errPath, 8, True)
                        ef.WriteLine "[Push.Delete] Remove '" & Trim(d) & "' Err=" & Err.Number & ": " & Err.Description
                        ef.Close
                        Err.Clear
                    End If
                    On Error GoTo 0
                End If
            End If
        Next
        fso.DeleteFile "${tempDir}\\delete_list.txt"
    End If

    Set folder = fso.GetFolder("${tempDir}")

    For Each fileItem In folder.Files
        If UCase(fso.GetExtensionName(fileItem.Path)) = "TXT" And InStr(fileItem.Name, "_TYPE_") > 0 Then
            parts = Split(fso.GetBaseName(fileItem.Path), "_TYPE_")
            compName = parts(0)
            compType = CInt(parts(1))

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
                    Set ef = fso.OpenTextFile(errPath, 8, True)
                    ef.WriteLine "[Push.Add] Add '" & compName & "' (Type=" & compType & ") Err=" & Err.Number & ": " & Err.Description
                    ef.Close
                    Err.Clear
                End If
                On Error GoTo 0
                If Not comp Is Nothing Then comp.Name = compName
            End If

            If comp Is Nothing Then
                Set ef = fso.OpenTextFile(errPath, 8, True)
                ef.WriteLine "[Push.Add] Component '" & compName & "' could not be created or found. Skipped."
                ef.Close
            Else
                Set inStr = CreateObject("ADODB.Stream")
                inStr.Type = 2
                inStr.Charset = "shift_jis"
                inStr.Open
                inStr.LoadFromFile fileItem.Path
                newContent = inStr.ReadText
                inStr.Close

                Set codeMod = comp.CodeModule
                On Error Resume Next
                If codeMod.CountOfLines > 0 Then
                    codeMod.DeleteLines 1, codeMod.CountOfLines
                    If Err.Number <> 0 Then
                        Set ef = fso.OpenTextFile(errPath, 8, True)
                        ef.WriteLine "[Push.DeleteLines] '" & compName & "' Err=" & Err.Number & ": " & Err.Description
                        ef.Close
                        Err.Clear
                    End If
                End If

                codeMod.AddFromString newContent
                If Err.Number <> 0 Then
                    Set ef = fso.OpenTextFile(errPath, 8, True)
                    ef.WriteLine "[Push.AddFromString] '" & compName & "' Err=" & Err.Number & ": " & Err.Description
                    ef.Close
                    Err.Clear
                End If
                On Error GoTo 0
            End If

            fso.DeleteFile fileItem.Path
        End If
    Next

    ' --- CLEANUP INJECTED MACROS ---
    Dim px, cx, cName, tmpProj
    For px = 1 To vbe.VBProjects.Count
        Set tmpProj = vbe.VBProjects.Item(px)
        For cx = tmpProj.VBComponents.Count To 1 Step -1
            cName = UCase(tmpProj.VBComponents.Item(cx).Name)
            If Left(cName, 4) = "C5D_" Then
                On Error Resume Next
                tmpProj.VBComponents.Remove tmpProj.VBComponents.Item(cx)
                If Err.Number <> 0 Then
                    Set ef = fso.OpenTextFile(errPath, 8, True)
                    ef.WriteLine "[Push.Cleanup] Remove '" & cName & "' Err=" & Err.Number & ": " & Err.Description
                    ef.Close
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

    const pushVbsScript = `
On Error Resume Next
Dim ag_catia, ag_sys, ag_args()
Set ag_catia = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then WScript.Quit 1
Set ag_sys = ag_catia.SystemService
If Err.Number <> 0 Or ag_sys Is Nothing Then WScript.Quit 1
ag_sys.ExecuteScript "${tempDir}", 1, "c5d_push.catvbs", "CATMain", ag_args
`;
    fs.writeFileSync(pushVbsPath, pushVbsScript, 'utf-8');

    vscode.window.withProgress({
        location: vscode.ProgressLocation.Notification,
        title: t('progress.push', targetProject),
        cancellable: false
    }, async (progress) => {
        return new Promise<void>((resolve, reject) => {
            exec(`%SystemRoot%\\SysWOW64\\cscript.exe //nologo "${pushVbsPath}"`, (error, stdout, stderr) => {
                if (fs.existsSync(pushVbsPath)) fs.unlinkSync(pushVbsPath);
                if (fs.existsSync(catScriptPath)) fs.unlinkSync(catScriptPath);

                flushCatScriptErrors(tempDir);

                if (error) {
                    outputChannel.appendLine(`[Push Error]`);
                    outputChannel.appendLine(`Error: ${error.message}`);
                    outputChannel.appendLine(`STDOUT: ${stdout}`);
                    outputChannel.appendLine(`STDERR: ${stderr}`);
                    outputChannel.show(true);
                    vscode.window.showErrorMessage(t('error.pushFailed'));
                    return reject();
                }
                // Update push cache with successfully pushed modules (scoped to targetProject)
                for (const [modName, hash] of Object.entries(pendingCacheUpdates)) {
                    pushCache[modName] = hash;
                }
                // Remove deleted modules from cache
                if (performDelete) {
                    for (const delName of toDelete) {
                        delete pushCache[delName];
                    }
                }
                allProjectsCache[targetProject] = pushCache;
                try {
                    fs.writeFileSync(cachePath, JSON.stringify(allProjectsCache, null, 2), 'utf-8');
                    outputChannel.appendLine(`[Push Cache] Saved to: ${cachePath}`);
                } catch (e) {
                    outputChannel.appendLine(`[Push Cache] Failed to save cache: ${e}`);
                }

                const skipMsg = skippedCount > 0 ? `（変更なしスキップ: ${skippedCount} 個）` : '';
                const deleteMsg = performDelete ? '（削除同期を含む）' : '';
                vscode.window.showInformationMessage(t('info.pushSuccess', String(count), skipMsg, deleteMsg));
                vscode.commands.executeCommand('cat5dev.refreshTree');
                resolve();
            });
        });
    });
}
