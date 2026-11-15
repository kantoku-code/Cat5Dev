const { SyncOperationManager } = require('../out/syncOperation');
const [lockRoot, tempRoot, mode] = process.argv.slice(2);
const manager = new SyncOperationManager({ lockRoot, tempRoot });
try {
    const context = manager.acquire('worker');
    if (mode === 'crash') {
        context.markExternalRunning(require('node:path').join(context.tempDir, 'result.json'), { operationId: context.id, stepId: 'step', kind: 'pull', targetProject: 'Review' });
        process.send({ status: 'crashed', id: context.id, tempDir: context.tempDir }); process.exit(0);
    }
    process.send({ status: 'acquired', id: context.id, tempDir: context.tempDir });
    process.on('message', () => { context.close(); process.exit(0); });
} catch (e) { process.send({ status: 'busy', error: String(e) }); process.exit(0); }
