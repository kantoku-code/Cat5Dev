// Examples generated from the same source used by the extension at runtime.
const fs = require('node:fs');
const path = require('node:path');
const { buildCatiaScript, buildRunnerScript } = require('../out/catiaScripts');
const request = {
    operationId: '{{OPERATION_ID}}', stepId: '{{STEP_ID}}', kind: 'push',
    targetProject: '{{TARGET_PROJECT}}', encoding: '{{ENCODING}}', directory: '{{TEMP_DIR}}',
    modules: [{ name: '{{MODULE_NAME}}', type: 1, file: 'module.txt' }], deletions: []
};
fs.writeFileSync(path.join(__dirname, 'c5d_push.catvbs.bas_utf'), buildCatiaScript(request));
fs.writeFileSync(path.join(__dirname, 'c5d_run_push.vbs.bas_utf'), buildRunnerScript(request));
