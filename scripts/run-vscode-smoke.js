const fs = require('node:fs');
const path = require('node:path');
const { execFile, execFileSync } = require('node:child_process');
const root = path.resolve(__dirname, '..');
fs.mkdirSync(path.join(root, 'temp'), { recursive: true });
const smokeRoot = fs.mkdtempSync(path.join(root, 'temp', 'vscode-smoke-'));
let extensionRoot = root;
const vsixIndex = process.argv.indexOf('--vsix');
if (vsixIndex !== -1) {
    const vsix = path.resolve(process.argv[vsixIndex + 1]);
    const unpack = path.join(smokeRoot, 'package'); fs.mkdirSync(unpack);
    execFileSync('python', ['-c', 'import zipfile,sys,pathlib; root=pathlib.Path(sys.argv[2]).resolve(); z=zipfile.ZipFile(sys.argv[1]); names=[n for n in z.namelist() if n.startswith("extension/")]; assert all(root in (root/n).resolve().parents for n in names); z.extractall(root,names)', vsix, unpack], { windowsHide: true });
    extensionRoot = path.join(unpack, 'extension');
}
const workspace = path.join(smokeRoot, 'workspace'); fs.mkdirSync(workspace);
fs.mkdirSync(path.join(workspace, 'modules'));
fs.writeFileSync(path.join(workspace, 'cat5dev.toml'), '[project]\nlanguage = "en"\ntarget_project = ""\n[formatter]\nenabled = true\nformat_on_save = false\nnormalize_operator_spacing = true\n');
fs.writeFileSync(path.join(workspace, 'modules', 'Smoke.bas_utf'), 'Sub Test()\nDim value_ As Long\nDebug.Print value_\nMsgBox Prompt:="hello"\nEnd Sub\n');
const executable = process.env.CAT5DEV_CODE_EXE || path.join(process.env.LOCALAPPDATA, 'Programs', 'Microsoft VS Code', 'Code.exe');
const args = ['--user-data-dir', path.join(smokeRoot, 'profile'), '--extensions-dir', path.join(smokeRoot, 'extensions'),
    '--extensionDevelopmentPath=' + extensionRoot, '--extensionTestsPath=' + path.join(root, 'tests', 'vscodeSmoke.js'),
    '--skip-welcome', '--skip-release-notes', '--disable-workspace-trust', '--disable-extensions', '--disable-gpu', workspace];
const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE;
env.CAT5DEV_SMOKE_EXTENSION_ROOT = extensionRoot;
const testDatabase = path.join(smokeRoot, 'catia_api.db');
fs.copyFileSync(path.join(extensionRoot, 'bin', 'catia_api.db'), testDatabase);
env.CAT5DEV_TLB_DB = testDatabase;
execFile(executable, args, { env, windowsHide: true, timeout: 90000, maxBuffer: 10 * 1024 * 1024 }, (error, stdout, stderr) => {
    const resultPath = path.join(smokeRoot, 'smoke-result.json');
    console.log('Smoke artifacts: ' + smokeRoot);
    if (fs.existsSync(resultPath)) console.log(fs.readFileSync(resultPath, 'utf8'));
    else console.error(stdout + '\n' + stderr);
    if (error || !fs.existsSync(resultPath)) { console.error(String(error || 'Missing smoke result')); process.exitCode = 1; }
});
