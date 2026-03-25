import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import * as http from 'http';
import { spawn, ChildProcess } from 'child_process';

const SERVER_START_TIMEOUT_MS = 5000;
const MAX_RESTART_COUNT = 3;

/** vbafmt HTTP サーバーのプロセス管理 */
export class VbaServer implements vscode.Disposable {
    private process: ChildProcess | null = null;
    private port: number | null = null;
    private restartCount = 0;
    private disposed = false;
    private readonly binaryPath: string;

    constructor(
        private readonly context: vscode.ExtensionContext,
        private readonly outputChannel: vscode.OutputChannel
    ) {
        const binaryName = process.platform === 'win32' ? 'vbafmt.exe' : 'vbafmt';
        this.binaryPath = path.join(context.extensionPath, 'bin', binaryName);
    }

    async start(): Promise<void> {
        if (!fs.existsSync(this.binaryPath)) {
            this.outputChannel.appendLine(`[vbafmt] バイナリが見つかりません: ${this.binaryPath}`);
            this.outputChannel.appendLine('[vbafmt] npm run compile:go を実行してバイナリをビルドしてください');
            return;
        }

        try {
            this.port = await this.spawnServer();
            this.outputChannel.appendLine(`[vbafmt] サーバー起動完了 (port: ${this.port})`);
        } catch (err) {
            this.outputChannel.appendLine(`[vbafmt] サーバー起動失敗: ${err}`);
        }
    }

    private spawnServer(): Promise<number> {
        return new Promise((resolve, reject) => {
            const proc = spawn(this.binaryPath, ['--server', '--port=0'], {
                stdio: ['ignore', 'pipe', 'pipe']
            });
            this.process = proc;

            const timer = setTimeout(() => {
                proc.kill();
                reject(new Error('サーバー起動タイムアウト'));
            }, SERVER_START_TIMEOUT_MS);

            let portResolved = false;
            proc.stdout!.on('data', (data: Buffer) => {
                const text = data.toString('utf8');
                if (!portResolved) {
                    const match = text.match(/PORT=(\d+)/);
                    if (match) {
                        portResolved = true;
                        clearTimeout(timer);
                        resolve(parseInt(match[1], 10));
                    }
                }
                // stdout の残りはログとして流す
                this.outputChannel.append(`[vbafmt-server] ${text}`);
            });

            proc.stderr!.on('data', (data: Buffer) => {
                this.outputChannel.append(`[vbafmt-server] ${data.toString('utf8')}`);
            });

            proc.on('error', (err: Error) => {
                clearTimeout(timer);
                reject(err);
            });

            proc.on('exit', (code) => {
                clearTimeout(timer);
                if (!portResolved) {
                    reject(new Error(`プロセスが終了しました (exit ${code})`));
                    return;
                }
                this.port = null;
                this.process = null;
                if (!this.disposed) {
                    this.scheduleRestart();
                }
            });
        });
    }

    private scheduleRestart(): void {
        if (this.restartCount >= MAX_RESTART_COUNT) {
            this.outputChannel.appendLine('[vbafmt] サーバー再起動の上限に達しました。手動で再起動してください。');
            return;
        }
        const delay = Math.pow(2, this.restartCount) * 1000;
        this.restartCount++;
        this.outputChannel.appendLine(`[vbafmt] ${delay / 1000}秒後に再起動します... (${this.restartCount}/${MAX_RESTART_COUNT})`);
        setTimeout(() => {
            if (!this.disposed) {
                this.start();
            }
        }, delay);
    }

    getBaseUrl(): string | null {
        if (this.port === null) {
            return null;
        }
        return `http://127.0.0.1:${this.port}`;
    }

    dispose(): void {
        this.disposed = true;
        if (this.process) {
            this.process.kill();
            this.process = null;
        }
        this.port = null;
    }
}

/** HTTP POST でフォーマットリクエストを送信する */
export function httpPost(url: string, body: string, timeoutMs: number): Promise<string> {
    return new Promise((resolve, reject) => {
        const parsed = new URL(url);
        const options: http.RequestOptions = {
            hostname: parsed.hostname,
            port: parsed.port,
            path: parsed.pathname,
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Content-Length': Buffer.byteLength(body, 'utf8'),
            },
        };

        const req = http.request(options, (res) => {
            let data = '';
            res.on('data', (chunk: Buffer) => { data += chunk.toString('utf8'); });
            res.on('end', () => resolve(data));
        });

        req.setTimeout(timeoutMs, () => {
            req.destroy();
            reject(new Error('HTTP request timeout'));
        });

        req.on('error', reject);
        req.write(body, 'utf8');
        req.end();
    });
}
