import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import { VbaServer } from './vbaServer';
import { formatVbaDocument } from './formatter';
import { t } from './i18n';
import { readFormatterOptions, readProjectSettings, writeTomlProjectKey } from './lintConfig';
import { OperationContext, SyncBusyError, SyncOperationManager } from './syncOperation';
import { CatiaExecutionError, EXIT, ExecutionOptions, createCatiaRequest, runCatiaRequest } from './catiaExecution';
import { BRIDGE_PROJECT_NAME, CatiaModuleText, CatiaRequest, CatiaResult, normalizeModuleText, safeFileName, stripBom } from './catiaTypes';
import { BridgeFileError, BridgeFiles } from './bridgeFiles';
import { PullConflictError, PullRecoveryError, applyPull, captureModules, listPullDeletions } from './pullTransaction';

const moduleExtension = /\.(bas_utf|cls_utf|frm_utf)$/i;
const extensionForType = { 1: '.bas_utf', 2: '.cls_utf', 3: '.frm_utf' };

/** How CatiaSync runs catia-run.exe; `bridge` prepares the bridge master once per operation. */
export interface SyncExecutionOptions extends Omit<ExecutionOptions, 'bridgeMaster' | 'log'> { bridge?: BridgeFiles }

/** UI orchestration; filesystem and process behavior remain independently testable. */
export class CatiaSync {
    /** Bridge master per operation: prepared (with cleanup) under the lock, before the first catia-run.exe. */
    private readonly masters = new WeakMap<OperationContext, string>();
    constructor(private readonly server: VbaServer, private readonly output: vscode.OutputChannel,
        private readonly manager: SyncOperationManager = new SyncOperationManager(), private readonly execution: SyncExecutionOptions = {}) { }

    private root(): string | undefined {
        const root = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
        if (!root) { void vscode.window.showErrorMessage(t('error.noWorkspace')); }
        return root;
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
            const exitCode = error instanceof CatiaExecutionError ? error.exitCode : undefined;
            if (error instanceof SyncBusyError) { this.showBusy(); }
            else if (exitCode === EXIT.runnerBusy) { void vscode.window.showWarningMessage(t('sync.runnerBusy')); }
            else if (exitCode === EXIT.catiaNotRunning) { void vscode.window.showErrorMessage(t('sync.catiaNotRunning')); }
            else if (exitCode === EXIT.vbaEditorClosed) { void vscode.window.showErrorMessage(t('sync.vbaEditorClosed')); }
            // The bridge master could not be prepared (BridgeFiles.prepare): nothing was sent to CATIA; the details are in Output.
            else if (exitCode === EXIT.bridgeUnavailable || error instanceof BridgeFileError) {
                this.output.show(true); void vscode.window.showErrorMessage(t('sync.bridgeUnavailable'));
            }
            else if (error instanceof PullRecoveryError) {
                void vscode.window.showErrorMessage(t('sync.recoveryFailed', error.recoveryPath, error.unrestored.join(', ')));
                this.output.show(true);
            } else if (error instanceof PullConflictError) { void vscode.window.showWarningMessage(t('sync.localChanged')); }
            else if (error instanceof CatiaExecutionError && error.uncertain) {
                this.output.show(true); void vscode.window.showErrorMessage(t('sync.unconfirmed'));
            } else if (error instanceof CatiaExecutionError && error.result?.failures.length) {
                this.output.show(true);
                void vscode.window.showErrorMessage(t('sync.partialFailure', String(error.result.succeededCount), String(error.result.failures.length)));
            } else {
                this.output.show(true);
                void vscode.window.showErrorMessage(t(kind === 'pull' ? 'error.pullFailed' : kind === 'push' ? 'error.pushFailed' : 'error.selectFailed'));
            }
            return undefined;
        } finally {
            try { context?.close(); } catch (error) { this.output.appendLine(`[cleanup] ${String(error)}`); }
        }
    }
    private bridgeMaster(context: OperationContext): string {
        const known = this.masters.get(context);
        if (known) { return known; }
        if (!this.execution.bridge) { throw new Error('CATIA bridge files are not configured'); }
        const prepared = this.execution.bridge.prepare();
        const counts = new Map<string, number>();
        for (const entry of prepared.cleanup) {
            counts.set(entry.action, (counts.get(entry.action) ?? 0) + 1);
            if (entry.action !== 'deleted' || entry.error) { this.output.appendLine(`[bridge] ${entry.action}: ${entry.file}${entry.error ? ` (${entry.error})` : ''}`); }
        }
        const summary = [...counts].map(([action, count]) => `${action}=${count}`).join(', ') || 'nothing';
        this.output.appendLine(`[bridge] master ${prepared.action}: ${prepared.master}; cleanup: ${summary}`);
        this.masters.set(context, prepared.master);
        return prepared.master;
    }
    private step(context: OperationContext, request: CatiaRequest): Promise<CatiaResult> {
        const { execute, runnerPath, timeoutMs } = this.execution;
        return runCatiaRequest(context, request, { execute, runnerPath, timeoutMs, bridgeMaster: this.bridgeMaster(context), log: text => this.output.appendLine(text) });
    }
    private async select(context: OperationContext, root: string): Promise<string | undefined> {
        const result = await this.step(context, createCatiaRequest(context, 'list', ''));
        // catia-run.exe already leaves the bridge out; filter again so it can never become the target project.
        const projects = result.projects.filter(name => name !== BRIDGE_PROJECT_NAME);
        if (!projects.length) { void vscode.window.showInformationMessage(t('info.projectNotFound')); return undefined; }
        const selected = await vscode.window.showQuickPick(projects, { placeHolder: t('select.placeholder') });
        if (selected) {
            writeTomlProjectKey(root, 'target_project', selected);
            void vscode.window.showInformationMessage(t('info.projectSelected', selected));
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
                const request = createCatiaRequest(context, 'pull', project);
                const result = await this.step(context, request);
                const incoming = new Map<string, string>();
                const formatOnPull = this.formatOnPull(root);
                for (const item of result.items) {
                    // catia-run.exe writes m<i>.txt as UTF-8 without BOM (bridge-contract 5.1).
                    let text = normalizeModuleText(fs.readFileSync(path.join(request.directory, item.file), 'utf8')) + '\n';
                    if (formatOnPull) {
                        text = await formatVbaDocument(text, this.server, this.output, root) ?? text;
                    }
                    incoming.set(item.name + extensionForType[item.type], text);
                }
                const deletions = listPullDeletions(snapshot, incoming);
                let deleteMissing = false;
                if (deletions.length) {
                    const answer = await vscode.window.showWarningMessage(t('sync.pullDeletions', deletions.join('\n')), { modal: true }, t('dialog.delete'), t('dialog.keep'));
                    if (answer === undefined) { void vscode.window.showInformationMessage(t('info.pullCancelled')); return; }
                    deleteMissing = answer === t('dialog.delete');
                }
                const applied = applyPull(root, snapshot, incoming, deleteMissing, { isDirty: file => this.dirty(file) });
                this.output.appendLine(`[Pull] received=${result.succeededCount} written=${applied.writtenCount} deleted=${applied.deletedCount}`);
                void vscode.window.showInformationMessage(t('info.pullSuccess', String(result.succeededCount)));
                await vscode.commands.executeCommand('cat5dev.refreshTree');
            });
        });
    }
    push(): Promise<void | undefined> {
        return this.run('push', async (context, root) => {
            const directory = path.join(root, 'modules');
            if (!fs.existsSync(directory)) { void vscode.window.showInformationMessage(t('error.noModulesDir')); return; }
            const files = fs.readdirSync(directory).filter(name => moduleExtension.test(name));
            if (!files.length) { void vscode.window.showInformationMessage(t('error.noModuleFiles')); return; }
            // Push sends the content saved on disk; warn before selecting a project or contacting CATIA.
            const unsaved = files.filter(file => this.dirty(path.join(directory, file)));
            if (unsaved.length && await vscode.window.showWarningMessage(t('warning.unsavedModules', unsaved.join('\n')), { modal: true }, t('dialog.continue')) !== t('dialog.continue')) {
                void vscode.window.showInformationMessage(t('info.pushCancelled')); return;
            }
            const project = await this.target(context, root); if (!project) { return; }
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
                const check = createCatiaRequest(context, 'check', project);
                const remote = await this.step(context, check);
                const remoteByName = new Map(remote.items.map(item => [item.name.toLowerCase(), item]));
                const changed = local.filter(item => {
                    const other = remoteByName.get(item.name.toLowerCase());
                    return !other || other.type !== item.type || item.comparable !== normalizeModuleText(fs.readFileSync(path.join(check.directory, other.file), 'utf8'));
                });
                const missing = remote.items.filter(item => !seen.has(item.name.toLowerCase())).map(item => item.name);
                if (!changed.length && !missing.length) { void vscode.window.showInformationMessage(t('info.noChanges', String(local.length))); return; }
                let deletions: string[] = [];
                if (missing.length) {
                    const answer = await vscode.window.showWarningMessage(t('warning.deleteModules', missing.join(', ')), { modal: true }, t('dialog.delete'), t('dialog.keep'));
                    if (answer === undefined) { void vscode.window.showInformationMessage(t('info.pushCancelled')); return; }
                    if (answer === t('dialog.delete')) { deletions = missing; }
                }
                const newForms = changed.filter(item => item.type === 3 && !remoteByName.has(item.name.toLowerCase()));
                if (newForms.length) { void vscode.window.showWarningMessage(t('warning.newUserForms', newForms.map(item => item.name).join(', '))); }
                const transferable = changed.filter(item => !newForms.includes(item));
                if (!transferable.length && !deletions.length) {
                    void vscode.window.showInformationMessage(changed.length ? t('warning.noMoreFiles') : t('info.noChanges', String(local.length)));
                    return;
                }
                const request = createCatiaRequest(context, 'push', project);
                // The text travels in request.json unchanged (no encoding conversion; cat5dev.toml encoding is ignored).
                request.modules = transferable.map((item): CatiaModuleText => ({ name: item.name, type: item.type, text: item.text }));
                request.deletions = deletions;
                const result = await this.step(context, request);
                void vscode.window.showInformationMessage(result.updatedCount === 0 && result.deletedCount > 0
                    ? t('info.pushDeleteOnly', String(result.deletedCount))
                    : t('info.pushSuccess', String(result.updatedCount), '', result.deletedCount ? t('push.deleteSync') : ''));
                await vscode.commands.executeCommand('cat5dev.refreshTree');
            });
        });
    }
    /** A lock left by an unconfirmed operation or an exited window gets a release action; a running one does not. */
    private showBusy(): void {
        if (!this.manager.hasQuarantine()) { void vscode.window.showWarningMessage(t('sync.busy')); return; }
        void vscode.window.showWarningMessage(t('sync.busyQuarantined'), t('dialog.releaseLock')).then(answer => {
            if (answer === t('dialog.releaseLock')) { return this.recoverLock(); }
        });
    }
    async recoverLock(): Promise<void> {
        try {
            if (this.manager.hasQuarantine()) {
                const answer = await vscode.window.showWarningMessage(t('sync.releaseConfirm'), { modal: true }, t('dialog.releaseLock'));
                if (answer !== t('dialog.releaseLock')) { return; }
                if (this.manager.releaseQuarantined()) {
                    this.output.appendLine('[recovery] Released an unconfirmed or orphaned operation lock by user confirmation.');
                    void vscode.window.showInformationMessage(t('sync.lockRecovered'));
                } else { void vscode.window.showWarningMessage(t('sync.lockRecoveryBlocked')); }
                return;
            }
            const recovered = await this.manager.recover();
            if (recovered) { void vscode.window.showInformationMessage(t('sync.lockRecovered')); }
            else { void vscode.window.showWarningMessage(t('sync.lockRecoveryBlocked')); }
        } catch (error) { this.output.appendLine(`[recovery] ${String(error)}`); void vscode.window.showWarningMessage(t('sync.lockRecoveryBlocked')); }
    }
}
