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
/** The 8.3 short form of an existing path (cmd's %~sI), or null when there is none (8.3 names disabled, not Windows, ...). */
function shortPath(file) {
    if (process.platform !== 'win32' || /["%^&|<>!]/.test(file)) return null;
    try {
        const short = require('node:child_process').execSync(`for %I in ("${file}") do @echo %~sI`, { shell: 'cmd.exe', encoding: 'utf8', windowsHide: true }).trim();
        return short && short.toLowerCase() !== file.toLowerCase() && fs.existsSync(short) ? short : null;
    } catch { return null; }
}
module.exports = { fixture, result, shortPath };
