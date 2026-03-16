import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { t } from './i18n';

export class CatiaTreeNode extends vscode.TreeItem {
    constructor(
        public readonly label: string,
        public readonly nodeType: 'project' | 'category' | 'module',
        public readonly compType: string,
        public readonly collapsibleState: vscode.TreeItemCollapsibleState,
        public readonly filePath?: string
    ) {
        super(label, collapsibleState);

        if (nodeType === 'project') {
            this.contextValue = 'catiaProject';
            this.iconPath = new vscode.ThemeIcon('folder-library');
            this.tooltip = t('treeview.targetProject');
        } else if (nodeType === 'category') {
            this.contextValue = 'catiaCategory';
            this.iconPath = new vscode.ThemeIcon('folder');
        } else {
            this.contextValue = 'catiaModule';
            if (compType === '1') {
                this.iconPath = new vscode.ThemeIcon('symbol-file');
            } else if (compType === '2') {
                this.iconPath = new vscode.ThemeIcon('symbol-class');
            } else if (compType === '3') {
                this.iconPath = new vscode.ThemeIcon('window');
            } else {
                this.iconPath = new vscode.ThemeIcon('file');
            }
            if (filePath) {
                this.command = {
                    command: 'vscode.open',
                    title: 'Open File',
                    arguments: [vscode.Uri.file(filePath)],
                };
            }
        }
    }
}

function getCategoryLabels(): Record<string, string> {
    return {
        '1': t('treeview.modules'),
        '2': t('treeview.classModules'),
        '3': t('treeview.forms'),
    };
}

const CATEGORY_ORDER: Record<string, number> = {
    '3': 0,
    '1': 1,
    '2': 2,
};

export class CatiaVbaTreeProvider implements vscode.TreeDataProvider<CatiaTreeNode> {
    private _onDidChangeTreeData: vscode.EventEmitter<CatiaTreeNode | undefined | void> = new vscode.EventEmitter<CatiaTreeNode | undefined | void>();
    readonly onDidChangeTreeData: vscode.Event<CatiaTreeNode | undefined | void> = this._onDidChangeTreeData.event;

    private cachedModules: CatiaTreeNode[] | undefined;
    private cachedProject: string | undefined;

    constructor(_context: vscode.ExtensionContext) { }

    refresh(): void {
        this.cachedModules = undefined;
        this.cachedProject = undefined;
        this._onDidChangeTreeData.fire();
    }

    getTreeItem(element: CatiaTreeNode): vscode.TreeItem {
        return element;
    }

    async getChildren(element?: CatiaTreeNode): Promise<CatiaTreeNode[]> {
        if (element?.nodeType === 'module') {
            return [];
        }

        const workspaceFolders = vscode.workspace.workspaceFolders;
        if (!workspaceFolders) return [];
        const rootPath = workspaceFolders[0].uri.fsPath;

        const settingsPath = path.join(rootPath, '.vscode', 'settings.json');
        let targetProject = '';
        if (fs.existsSync(settingsPath)) {
            try {
                const settings = JSON.parse(fs.readFileSync(settingsPath, 'utf-8'));
                if (settings.targetProject) { targetProject = settings.targetProject; }
            } catch (e) { }
        }

        if (!targetProject) { return []; }

        if (!element) {
            // Root: project node
            return [new CatiaTreeNode(targetProject, 'project', '', vscode.TreeItemCollapsibleState.Expanded)];
        }

        if (element.nodeType === 'project') {
            // Load modules from local filesystem
            if (!this.cachedModules || this.cachedProject !== targetProject) {
                this.cachedModules = this.loadLocalModules(rootPath);
                this.cachedProject = targetProject;
            }
            const types = [...new Set(this.cachedModules.map(m => m.compType))].sort((a, b) => (CATEGORY_ORDER[a] ?? 99) - (CATEGORY_ORDER[b] ?? 99));
            const categoryLabels = getCategoryLabels();
            return types.map(t => new CatiaTreeNode(
                categoryLabels[t] ?? `Type ${t}`,
                'category',
                t,
                vscode.TreeItemCollapsibleState.Expanded
            ));
        }

        if (element.nodeType === 'category') {
            return (this.cachedModules ?? []).filter(m => m.compType === element.compType);
        }

        return [];
    }

    private loadLocalModules(rootPath: string): CatiaTreeNode[] {
        const modulesDir = path.join(rootPath, 'modules');
        if (!fs.existsSync(modulesDir)) { return []; }

        const extToType: Record<string, string> = {
            '.bas_utf': '1',
            '.cls_utf': '2',
            '.frm_utf': '3',
        };

        const nodes: CatiaTreeNode[] = [];
        for (const file of fs.readdirSync(modulesDir)) {
            const ext = Object.keys(extToType).find(e => file.endsWith(e));
            if (ext) {
                const name = file.slice(0, -ext.length);
                const filePath = path.join(modulesDir, file);
                nodes.push(new CatiaTreeNode(name, 'module', extToType[ext], vscode.TreeItemCollapsibleState.None, filePath));
            }
        }
        nodes.sort((a, b) => a.label.localeCompare(b.label));
        return nodes;
    }
}
