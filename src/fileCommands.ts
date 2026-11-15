import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { t } from './i18n';
import { safeFileName } from './catiaTypes';

const moduleExtension = /\.(bas|cls|frm)_utf$/i;
const vbaIdentifier = /^\p{L}[\p{L}\p{N}_]*$/u;

function modulesDirectory(): string | undefined {
    const root = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    return root ? path.join(path.resolve(root), 'modules') : undefined;
}

function samePath(a: string, b: string): boolean {
    return path.resolve(a).toLowerCase() === path.resolve(b).toLowerCase();
}

function errorText(e: unknown): string {
    return e instanceof Error ? e.message : String(e);
}

function isDirty(filePath: string): boolean {
    return vscode.workspace.textDocuments.some(doc => doc.isDirty && samePath(doc.uri.fsPath, filePath));
}

async function refreshTree(): Promise<void> {
    await vscode.commands.executeCommand('cat5dev.refreshTree');
}

/**
 * Resolves the command argument (a tree node with `filePath`, or a file Uri) to a module file
 * directly under the workspace's modules folder. Shows an error and returns undefined otherwise.
 */
export function resolveModulePath(arg: unknown): string | undefined {
    const value = arg as { filePath?: unknown; scheme?: unknown; fsPath?: unknown } | undefined;
    let candidate: string | undefined;
    if (typeof value?.filePath === 'string') {
        candidate = value.filePath;
    } else if (value?.scheme === 'file' && typeof value.fsPath === 'string') {
        candidate = value.fsPath;
    }
    if (candidate === undefined) {
        vscode.window.showErrorMessage(t('file.noTarget'));
        return undefined;
    }
    const resolved = path.resolve(candidate);
    const directory = modulesDirectory();
    let valid = !!directory && samePath(path.dirname(resolved), directory) && moduleExtension.test(resolved);
    if (valid) {
        try {
            const stat = fs.lstatSync(resolved);
            valid = stat.isFile() && !stat.isSymbolicLink();
        } catch {
            valid = false;
        }
    }
    if (!valid) {
        vscode.window.showErrorMessage(t('file.invalidTarget', candidate));
        return undefined;
    }
    return resolved;
}

/** Returns an error message for an unusable new module name, or undefined when it is acceptable. */
export function validateModuleName(name: string, filePath: string): string | undefined {
    const ext = moduleExtension.exec(filePath)?.[0] ?? '';
    if (!safeFileName(name) || !safeFileName(name + ext) || !vbaIdentifier.test(name)) {
        return t('file.invalidName', name);
    }
    const self = path.basename(filePath).toLowerCase();
    const wanted = name.toLowerCase();
    let entries: string[] = [];
    try { entries = fs.readdirSync(path.dirname(filePath)); } catch { entries = []; }
    for (const entry of entries) {
        if (entry.toLowerCase() === self || !moduleExtension.test(entry)) { continue; }
        if (entry.replace(moduleExtension, '').toLowerCase() === wanted) {
            return t('file.exists', name);
        }
    }
    return undefined;
}

async function copyPath(arg: unknown): Promise<void> {
    const filePath = resolveModulePath(arg);
    if (!filePath) { return; }
    await vscode.env.clipboard.writeText(filePath);
    vscode.window.showInformationMessage(t('file.copySuccess', filePath));
}

async function deleteFile(arg: unknown): Promise<void> {
    const filePath = resolveModulePath(arg);
    if (!filePath) { return; }
    const fileName = path.basename(filePath);
    if (isDirty(filePath)) {
        vscode.window.showWarningMessage(t('file.dirty', fileName));
        return;
    }
    const deleteBtn = t('file.deleteButton');
    const confirmed = await vscode.window.showWarningMessage(t('file.deleteConfirm', fileName), { modal: true }, deleteBtn);
    if (confirmed !== deleteBtn) { return; }
    try {
        await vscode.workspace.fs.delete(vscode.Uri.file(filePath), { useTrash: true });
    } catch (e) {
        const code = (e as { code?: unknown })?.code;
        const missing = code === 'ENOENT' || code === 'FileNotFound' || !fs.existsSync(filePath);
        if (!missing) {
            vscode.window.showErrorMessage(t('file.deleteFailed', fileName, errorText(e)));
        }
    }
    await refreshTree();
}

async function renameFile(arg: unknown): Promise<void> {
    const filePath = resolveModulePath(arg);
    if (!filePath) { return; }
    const fileName = path.basename(filePath);
    if (isDirty(filePath)) {
        vscode.window.showWarningMessage(t('file.dirty', fileName));
        return;
    }
    const ext = moduleExtension.exec(fileName)![0];
    const baseName = fileName.slice(0, -ext.length);
    const newName = await vscode.window.showInputBox({
        prompt: t('file.rename'),
        value: baseName,
        validateInput: value => validateModuleName(value, filePath)
    });
    if (newName === undefined || newName === baseName) { return; }
    const invalid = validateModuleName(newName, filePath);
    if (invalid) {
        vscode.window.showErrorMessage(invalid);
        return;
    }
    if (/\.frm_utf$/i.test(ext)) {
        const proceed = t('dialog.continue');
        if (await vscode.window.showWarningMessage(t('file.formRenameWarning'), { modal: true }, proceed) !== proceed) { return; }
    }
    if (isDirty(filePath)) {
        vscode.window.showWarningMessage(t('file.dirty', fileName));
        return;
    }
    const newPath = path.join(path.dirname(filePath), newName + ext);
    const edit = new vscode.WorkspaceEdit();
    edit.renameFile(vscode.Uri.file(filePath), vscode.Uri.file(newPath), { overwrite: false });
    let reason = '';
    let applied = false;
    try {
        applied = await vscode.workspace.applyEdit(edit);
    } catch (e) {
        reason = errorText(e);
    }
    if (!applied) {
        vscode.window.showErrorMessage(t('file.renameFailed', `${fileName} -> ${newName}${ext}`, reason || 'applyEdit'));
    }
    await refreshTree();
}

export function registerFileCommands(context: vscode.ExtensionContext): void {
    context.subscriptions.push(
        vscode.commands.registerCommand('cat5dev.renameFile', renameFile),
        vscode.commands.registerCommand('cat5dev.deleteFile', deleteFile),
        vscode.commands.registerCommand('cat5dev.copyPath', copyPath)
    );
}
