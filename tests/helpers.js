const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
function fixture(t) {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), 'cat5dev-tests-'));
    t.after(() => { if (path.dirname(root) === fs.realpathSync(os.tmpdir()) || path.dirname(root) === os.tmpdir()) fs.rmSync(root, { recursive: true, force: true }); });
    return root;
}
function result(request, overrides = {}) {
    return { version: 1, operationId: request.operationId, stepId: request.stepId, kind: request.kind,
        targetProject: request.targetProject, expectedCount: 0, succeededCount: 0,
        failures: [], items: [], projects: [], updatedCount: 0, deletedCount: 0, ...overrides };
}
module.exports = { fixture, result };
