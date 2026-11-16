import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as iconv from 'iconv-lite';
import { VbaServer } from './vbaServer';
import { formatVbaDocument } from './formatter';
import { t, getLanguage } from './i18n';
import { defaultEncodingFor, readFormatterOptions, readProjectSettings, writeTomlProjectKey } from './lintConfig';
import { OperationContext, SyncBusyError, SyncOperationManager } from './syncOperation';
import { CatiaExecutionError, ExecutionOptions, createCatiaRequest, runCatiaRequest } from './catiaExecution';
import { CatiaItem, CatiaRequest, CatiaResult, normalizeModuleText, safeFileName, stripBom } from './catiaTypes';
import { PullConflictError, PullRecoveryError, applyPull, captureModules, listPullDeletions } from './pullTransaction';

const moduleExtension = /\.(bas_utf|cls_utf|frm_utf)$/i;
const extensionForType = { 1: '.bas_utf', 2: '.cls_utf', 3: '.frm_utf' };

/** UI orchestration; filesystem and process behavior remain independently testable. */
export class CatiaSync {
    constructor(private readonly server: VbaServer, private readonly output: vscode.OutputChannel,
        private readonly manager: SyncOperationManager = new SyncOperationManager(), private readonly execution: ExecutionOptions = {}) { }

    private root(): string | undefined {
        const root = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
        if (!root) { void vscode.window.showErrorMessage(t('error.noWorkspace')); }
        return root;
    }
    private encoding(root: string): string {
        return readProjectSettings(root).encoding || defaultEncodingFor(getLanguage());
    }
    /** cat5dev.toml の [formatter] format_on_pull。未記載なら旧 VSCode 設定 cat5dev.formatter.formatOnPull を読む（互換） */
    private formatOnPull(root: string): boolean {
        const opts = readFormatterOptions(root);
        if (!opts.enabled) { return false; }
        return opts.format_on_pull ?? vscode.workspace.getConfiguration('cat5dev.formatter').get<boolean>('formatOnPull', false);
    }
    private dirty(filePath: string): boolean {
        return vscode.workspace.textDocuments.some(doc => doc.isDirty && path.resolve(doc.uri.fsPath).toLowerCase() === path.resolve(filePath).toLowerCase());
    }
    private async run<T>(kind: string, task: (context: OperationContext, root: string) => Promise<T>): Promise<T | undefined> {
        const root = this.root(); if (!root) { return undefined; }
        let context: OperationContext | undefined;
        try {
            context = this.manager.acquire(kind);
            return await task(context, root);
        } catch (error) {
            this.output.appendLine(`[${kind}] ${error instanceof Error ? error.stack : String(error)}`);
            if (error instanceof SyncBusyError) { await vscode.window.showWarningMessage(t('sync.busy')); }
            else if (error instanceof PullRecoveryError) {
                await vscode.window.showErrorMessage(t('sync.recoveryFailed', error.recoveryPath, error.unrestored.join(', ')));
                this.output.show(true);
            } else if (error instanceof PullConflictError) { await vscode.window.showWarningMessage(t('sync.localChanged')); }
            else if (error instanceof CatiaExecutionError && error.uncertain) {
                this.output.show(true); await vscode.window.showErrorMessage(t('sync.unconfirmed'));
            } else if (error instanceof CatiaExecutionError && error.result?.failures.length) {
                this.output.show(true);
                await vscode.window.showErrorMessage(t('sync.partialFailure', String(error.result.succeededCount), String(error.result.failures.length)));
            } else {
                this.output.show(true);
                await vscode.window.showErrorMessage(t(kind === 'pull' ? 'error.pullFailed' : kind === 'push' ? 'error.pushFailed' : 'error.selectFailed'));
            }
            return undefined;
        } finally {
            try { context?.close(); } catch (error) { this.output.appendLine(`[cleanup] ${String(error)}`); }
        }
    }
    private step(context: OperationContext, request: CatiaRequest): Promise<CatiaResult> {
        return runCatiaRequest(context, request, { ...this.execution, log: text => this.output.appendLine(text) });
    }
    private async select(context: OperationContext, root: string): Promise<string | undefined> {
        const result = await this.step(context, createCatiaRequest(context, 'list', '', this.encoding(root)));
        if (!result.projects.length) { await vscode.window.showInformationMessage(t('info.projectNotFound')); return undefined; }
        const selected = await vscode.window.showQuickPick(result.projects, { placeHolder: t('select.placeholder') });
        if (selected) {
            writeTomlProjectKey(root, 'target_project', selected);
            await vscode.window.showInformationMessage(t('info.projectSelected', selected));
            await vscode.commands.executeCommand('cat5dev.refreshTree');
        }
        return selected;
    }
    private async target(context: OperationContext, root: string): Promise<string | undefined> {
        return readProjectSettings(root).targetProject || await this.select(context, root);
    }
    selectProject(): Promise<string | undefined> { return this.run('list', (context, root) => this.select(context, root)); }

    pull(): Promise<void | undefined> {
        return this.run('pull', async (context, root) => {
            const project = await this.target(context, root); if (!project) { return; }
            const snapshot = captureModules(root);
            await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: t('progress.pull', project), cancellable: false }, async () => {
                const request = createCatiaRequest(context, 'pull', project, this.encoding(root));
                const result = await this.step(context, request);
                const incoming = new Map<string, string>();
                const formatOnPull = this.formatOnPull(root);
                for (const item of result.items) {
                    let text = normalizeModuleText(iconv.decode(fs.readFileSync(path.join(request.directory, item.file)), request.encoding)) + '\n';
                    if (formatOnPull) {
                        text = await formatVbaDocument(text, this.server, this.output, root) ?? text;
                    }
                    incoming.set(item.name + extensionForType[item.type], text);
                }
                const deletions = listPullDeletions(snapshot, incoming);
                let deleteMissing = false;
                if (deletions.length) {
                    const answer = await vscode.window.showWarningMessage(t('sync.pullDeletions', deletions.join('\n')), { modal: true }, t('dialog.delete'), t('dialog.keep'));
                    if (answer === undefined) { await vscode.window.showInformationMessage(t('info.pullCancelled')); return; }
                    deleteMissing = answer === t('dialog.delete');
                }
                const applied = applyPull(root, snapshot, incoming, deleteMissing, { isDirty: file => this.dirty(file) });
                this.output.appendLine(`[Pull] received=${result.succeededCount} written=${applied.writtenCount} deleted=${applied.deletedCount}`);
                await vscode.window.showInformationMessage(t('info.pullSuccess', String(result.succeededCount)));
                await vscode.commands.executeCommand('cat5dev.refreshTree');
            });
        });
    }
    push(): Promise<void | undefined> {
        return this.run('push', async (context, root) => {
            const directory = path.join(root, 'modules');
            if (!fs.existsSync(directory)) { await vscode.window.showInformationMessage(t('error.noModulesDir')); return; }
            const files = fs.readdirSync(directory).filter(name => moduleExtension.test(name));
            if (!files.length) { await vscode.window.showInformationMessage(t('error.noModuleFiles')); return; }
            // Push sends the content saved on disk; warn before selecting a project or contacting CATIA.
            const unsaved = files.filter(file => this.dirty(path.join(directory, file)));
            if (unsaved.length && await vscode.window.showWarningMessage(t('warning.unsavedModules', unsaved.join('\n')), { modal: true }, t('dialog.continue')) !== t('dialog.continue')) {
                await vscode.window.showInformationMessage(t('info.pushCancelled')); return;
            }
            const project = await this.target(context, root); if (!project) { return; }
            const encoding = this.encoding(root);
            // text is sent as-is; comparable ignores line-ending differences (CATIA exports CRLF, Pull writes LF).
            const local: { name: string; type: 1 | 2 | 3; text: string; comparable: string }[] = [];
            const seen = new Set<string>();
            for (const file of files) {
                const name = file.slice(0, file.lastIndexOf('.'));
                if (!safeFileName(name) || seen.has(name.toLowerCase())) { throw new Error(`Invalid or duplicate module: ${name}`); }
                seen.add(name.toLowerCase());
                const fullPath = path.join(directory, file);
                const stat = fs.lstatSync(fullPath);
                if (!stat.isFile() || stat.isSymbolicLink()) { throw new Error(`Not a regular module: ${file}`); }
                // Editors may save UTF-8 with BOM; strip it so it is neither compared nor sent to CATIA.
                const text = stripBom(fs.readFileSync(fullPath, 'utf8')).trimEnd();
                local.push({ name, type: /\.cls_utf$/i.test(file) ? 2 : /\.frm_utf$/i.test(file) ? 3 : 1,
                    text, comparable: normalizeModuleText(text) });
            }
            const longNames = local.filter(item => item.name.length > 31).map(item => item.name);
            if (longNames.length && await vscode.window.showWarningMessage(t('warning.longModuleNames', longNames.join(', ')), { modal: true }, t('dialog.continue')) !== t('dialog.continue')) { return; }
            await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: t('progress.push', project), cancellable: false }, async () => {
                const check = createCatiaRequest(context, 'check', project, encoding);
                const remote = await this.step(context, check);
                const remoteByName = new Map(remote.items.map(item => [item.name.toLowerCase(), item]));
                const changed = local.filter(item => {
                    const other = remoteByName.get(item.name.toLowerCase());
                    return !other || other.type !== item.type || item.comparable !== normalizeModuleText(iconv.decode(fs.readFileSync(path.join(check.directory, other.file)), encoding));
                });
                const missing = remote.items.filter(item => !seen.has(item.name.toLowerCase())).map(item => item.name);
                if (!changed.length && !missing.length) { await vscode.window.showInformationMessage(t('info.noChanges', String(local.length))); return; }
                let deletions: string[] = [];
                if (missing.length) {
                    const answer = await vscode.window.showWarningMessage(t('warning.deleteModules', missing.join(', ')), { modal: true }, t('dialog.delete'), t('dialog.keep'));
                    if (answer === undefined) { await vscode.window.showInformationMessage(t('info.pushCancelled')); return; }
                    if (answer === t('dialog.delete')) { deletions = missing; }
                }
                const newForms = changed.filter(item => item.type === 3 && !remoteByName.has(item.name.toLowerCase()));
                if (newForms.length) { await vscode.window.showWarningMessage(t('warning.newUserForms', newForms.map(item => item.name).join(', '))); }
                const transferable = changed.filter(item => !newForms.includes(item));
                if (!transferable.length && !deletions.length) {
                    await vscode.window.showInformationMessage(changed.length ? t('warning.noMoreFiles') : t('info.noChanges', String(local.length)));
                    return;
                }
                const request = createCatiaRequest(context, 'push', project, encoding);
                request.modules = transferable.map((item, index): CatiaItem => {
                    const file = `m${index}.txt`;
                    fs.writeFileSync(path.join(request.directory, file), iconv.encode(item.text, encoding));
                    return { name: item.name, type: item.type, file };
                });
                request.deletions = deletions;
                const result = await this.step(context, request);
                await vscode.window.showInformationMessage(result.updatedCount === 0 && result.deletedCount > 0
                    ? t('info.pushDeleteOnly', String(result.deletedCount))
                    : t('info.pushSuccess', String(result.updatedCount), '', result.deletedCount ? t('push.deleteSync') : ''));
                await vscode.commands.executeCommand('cat5dev.refreshTree');
            });
        });
    }
    async recoverLock(): Promise<void> {
        try {
            const recovered = await this.manager.recover();
            if (recovered) { await vscode.window.showInformationMessage(t('sync.lockRecovered')); }
            else { await vscode.window.showWarningMessage(t('sync.lockRecoveryBlocked')); }
        } catch (error) { this.output.appendLine(`[recovery] ${String(error)}`); await vscode.window.showWarningMessage(t('sync.lockRecoveryBlocked')); }
    }
}
