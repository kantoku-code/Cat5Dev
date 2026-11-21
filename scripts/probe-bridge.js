// Bridge probe (developer tool, not shipped): can the VBA project C5D_Bridge (resources/bridge/C5D_Bridge.catvba) be used as a
// mediator for Push without VBScript? Everything goes through CATIA.SystemService.ExecuteScript(<full path of .catvba>,
// 2 = catScriptLibraryTypeVBAProject, <module>, <function>, <args>) via temp/eval-probe.exe (tools/catia-run/EvalProbe.cs, step vbacall).
// No .catvbs/.CATScript is ever run, so vbscript.dll can be judged on a freshly restarted CATIA.
// Usage: node scripts/probe-bridge.js [<step> ...]     (default: all steps in this order)
//        node scripts/probe-bridge.js --check-only     (CATIA is not asked to do anything: DLLs, CNEXT.exe pid/start time, bridge files)
// Steps (C5D_Bridge_Push / C5D_ProbeModule are standard modules of the project C5D_Bridge):
//   version     C5D_Bridge_Push.C5D_Version                              -> "c5d-bridge:1"
//   install     C5D_Bridge_Push.C5D_Install("C5D_ProbeModule", <source>) -> "ok"      (writes PROBE_SOURCE into C5D_Bridge, in memory)
//   echo        C5D_ProbeModule.C5D_ProbeEcho(<Japanese text + U+2713>) -> "probe:Long:<same text>" (Unicode round trip, real VBA)
//   projects    C5D_ProbeModule.C5D_ProbeProjects                        -> "count=N;err=0"
//   self-guard  C5D_Bridge_Push.C5D_Install("C5D_Bridge_Push", "x")      -> "error:self"
// Before the steps, resources/bridge/C5D_Bridge.catvba is copied to %LOCALAPPDATA%\Cat5Dev\bridge\C5D_Bridge.catvba (overwritten;
// aborts when the target is locked) and that full path is the library of every call. CNEXT.exe modules are recorded before the first
// step and after each step. Each step has a 60 s timeout (CATIA may be stopped at a dialog); the run stops there.
// Every string handed to eval-probe.exe is passed as b64:<UTF-8 base64> so non-ASCII text, CR/LF and quotes survive the command line.
// Each run writes temp/probe-bridge-<yyyyMMdd-HHmmss>.json and prints a summary to stdout.
// A read-only copy (left by --unload-experiment) is made writable for the copy and read-only again afterwards.
//
// Unload experiment (can C5D_Bridge be removed from CATIA after use, and does a read-only .catvba still work?):
//        node scripts/probe-bridge.js --unload-experiment   (does NOT copy; uses the copy CATIA already has loaded)
//        node scripts/probe-bridge.js --unload-cleanup      (unloads C5D_Bridge again; the read-only attribute stays)
//   Steps of --unload-experiment (UNLOAD_STEPS): pre, install-unloader (C5D_ProbeUnload), apc-info, unload (C5D_UnloadSelf),
//   after-unload-probe (C5D_ProbeModule must be gone), reload-version, install-unloader-2, unload-2, set-readonly (local: lock probe +
//   chmod 0o444), readonly-version, readonly-install (C5D_ProbeReadonly), readonly-call. It stops there with C5D_Bridge loaded from
//   the read-only file so the user can try to save it in the VBA editor. A timeout (60 s, dialog?), an unreachable CATIA or a vanished
//   CNEXT.exe ("CATIA crashed after step X") aborts the run; a COM exception of a step does not.
//   Reports: temp/probe-bridge-unload-<time>.json / temp/probe-bridge-unload-cleanup-<time>.json.
//
// Read-only from the start (does a .catvba that is read-only BEFORE CATIA loads it stay unsaved? Setting the attribute after the
// load did not help: CATIA kept a writable handle and the VBA editor saved over it):
//        node scripts/probe-bridge.js --readonly-fresh        (right after a CATIA start, VBA editor NOT opened yet)
//        node scripts/probe-bridge.js --readonly-fresh-check  (no CATIA call; after the user tried to save in the VBA editor)
//   --readonly-fresh puts the bridge at %LOCALAPPDATA%\Cat5Dev\bridge\C5D_Bridge-<first 8 hex of sha256>.catvba (reused when the hash
//   matches, else replaced; never the plain C5D_Bridge.catvba above), sets and verifies the read-only attribute, records CNEXT.exe and
//   vbscript.dll, then runs READONLY_FRESH_STEPS (version, install C5D_ProbeReadonly, projects, call) one --no-retry call at a time
//   with that path as the library, and records time/size/read-only/lock of the file again. A failed step does not stop the run.
//   --readonly-fresh-check records the same file facts plus the hash (vs the source) and compares them with the latest run's report.
//   Reports: temp/probe-bridge-readonly-fresh-<time>.json / temp/probe-bridge-readonly-fresh-check-<time>.json.
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { execFile, execFileSync } = require('node:child_process');
const { moduleSnapshot, KEY_DLLS, timestamp, addedModules, powershellPath } = require('./probe-catscript');
const { buildProbe, parseProbeOutput } = require('./probe-evaluate');

const ROOT = path.resolve(__dirname, '..');
const TEMP = path.join(ROOT, 'temp');
const BRIDGE_FILE = 'C5D_Bridge.catvba';
const BRIDGE_PROJECT = 'C5D_Bridge';
const BRIDGE_MODULE = 'C5D_Bridge_Push';
const PROBE_MODULE = 'C5D_ProbeModule';
const BRIDGE_VERSION = 'c5d-bridge:1';
const STEP_TIMEOUT_MS = 60000;
/** Japanese text + U+2713 (CHECK MARK, not in Shift-JIS): must come back unchanged through ExecuteScript. */
const ECHO_ARGUMENT = '日本語テスト✓';
const LOCK_CODES = ['EBUSY', 'EPERM', 'EACCES'];

/** VBA written into C5D_ProbeModule by C5D_Install (ASCII only, CRLF, no Attribute line). */
const PROBE_SOURCE = [
    'Option Explicit',
    'Public Function C5D_ProbeEcho(ByVal a As Variant) As Variant',
    '    Dim n As Long',
    '    n = 7',
    '    C5D_ProbeEcho = "probe:" & TypeName(n) & ":" & CStr(a)',
    'End Function',
    'Public Function C5D_ProbeProjects() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim projects As Object',
    '    Dim n As Long',
    '    Set apc = CreateObject("MSAPC.Apc.7.1")',
    '    If apc Is Nothing Then',
    '        Err.Clear',
    '        Set apc = CreateObject("MSAPC.Apc")',
    '    End If',
    '    If apc Is Nothing Then',
    '        C5D_ProbeProjects = "count=-1;err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Set projects = apc.VBE.VBProjects',
    '    If projects Is Nothing Then',
    '        C5D_ProbeProjects = "count=-1;err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    n = -1',
    '    n = projects.Count',
    '    C5D_ProbeProjects = "count=" & CStr(n) & ";err=" & CStr(Err.Number)',
    'End Function',
    ''
].join('\r\n');

const UNLOAD_MODULE = 'C5D_ProbeUnload';
const READONLY_MODULE = 'C5D_ProbeReadonly';
const UNLOAD_RESULT = 'removed;err=0';
/** Argument of after-unload-probe; C5D_ProbeEcho answering "probe:Long:x" means C5D_Bridge was NOT unloaded. */
const UNLOADED_ECHO_ARGUMENT = 'x';
/** VBA lines that set apc to MSAPC.Apc.7.1 (else MSAPC.Apc) or return "<fn> = <prefix>:<err>:<desc>" when neither exists. */
const apcLines = (fn, prefix) => [
    '    Set apc = CreateObject("MSAPC.Apc.7.1")',
    '    If apc Is Nothing Then',
    '        Err.Clear',
    '        Set apc = CreateObject("MSAPC.Apc")',
    '    End If',
    '    If apc Is Nothing Then',
    `        ${fn} = "${prefix}" & CStr(Err.Number) & ":" & Err.Description`,
    '        Exit Function',
    '    End If'
];
/**
 * VBA written into C5D_ProbeUnload (ASCII only, CRLF, On Error Resume Next, no MsgBox). Property values are read into variables
 * before any If, because with Resume Next an error inside an If condition runs the Then branch.
 *   C5D_UnloadSelf()       VBE.VBProjects.Remove <project named C5D_Bridge> -> "removed;err=0" | "error:remove:<n>:<desc>" | "error:..."
 *   C5D_ApcProjectsInfo()  apc.Projects (ApcProjects): "apcProjects=N|<Name>:<TypeName>:vb=<VBProject.Name>|...;err=N" (no unload)
 */
const UNLOAD_SOURCE = [
    'Option Explicit',
    'Public Function C5D_UnloadSelf() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim vbps As Object',
    '    Dim proj As Object',
    '    Dim found As Object',
    '    Dim nm As String',
    '    Dim i As Long',
    '    Dim n As Long',
    ...apcLines('C5D_UnloadSelf', 'error:apc:'),
    '    Err.Clear',
    '    Set vbps = apc.VBE.VBProjects',
    '    If vbps Is Nothing Then',
    '        C5D_UnloadSelf = "error:vbprojects:" & CStr(Err.Number) & ":" & Err.Description',
    '        Exit Function',
    '    End If',
    '    n = -1',
    '    n = vbps.Count',
    '    For i = 1 To n',
    '        Set proj = Nothing',
    '        Set proj = vbps.Item(i)',
    '        nm = ""',
    '        If Not proj Is Nothing Then',
    '            nm = CStr(proj.Name)',
    '        End If',
    `        If nm = "${BRIDGE_PROJECT}" Then`,
    '            Set found = proj',
    '            Exit For',
    '        End If',
    '    Next',
    '    If found Is Nothing Then',
    '        C5D_UnloadSelf = "error:notfound:count=" & CStr(n) & ";err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    vbps.Remove found',
    '    If Err.Number = 0 Then',
    `        C5D_UnloadSelf = "${UNLOAD_RESULT}"`,
    '    Else',
    '        C5D_UnloadSelf = "error:remove:" & CStr(Err.Number) & ":" & Err.Description',
    '    End If',
    'End Function',
    'Public Function C5D_ApcProjectsInfo() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim s As String',
    '    Dim nm As String',
    '    Dim tn As String',
    '    Dim vn As String',
    '    Dim i As Long',
    '    Dim n As Long',
    ...apcLines('C5D_ApcProjectsInfo', 'apcProjects=-1;error:apc:'),
    '    Err.Clear',
    '    Set aps = apc.Projects',
    '    If aps Is Nothing Then',
    '        C5D_ApcProjectsInfo = "apcProjects=-1;err=" & CStr(Err.Number) & ":" & Err.Description',
    '        Exit Function',
    '    End If',
    '    n = -1',
    '    n = aps.Count',
    '    s = "apcProjects=" & CStr(n)',
    '    If Err.Number <> 0 Then',
    '        s = s & ";countErr=" & CStr(Err.Number) & ":" & Err.Description',
    '        Err.Clear',
    '    End If',
    '    For i = 1 To n',
    '        Set ap = Nothing',
    '        Set ap = aps.Item(i)',
    '        If ap Is Nothing Then',
    '            s = s & "|#" & CStr(i) & ":err=" & CStr(Err.Number)',
    '            Err.Clear',
    '        Else',
    '            nm = "?"',
    '            tn = "?"',
    '            vn = "?"',
    '            nm = CStr(ap.Name)',
    '            tn = TypeName(ap)',
    '            vn = CStr(ap.VBProject.Name)',
    '            s = s & "|" & nm & ":" & tn & ":vb=" & vn',
    '        End If',
    '    Next',
    '    C5D_ApcProjectsInfo = s & ";err=" & CStr(Err.Number)',
    'End Function',
    ''
].join('\r\n');
/** VBA written into C5D_ProbeReadonly while C5D_Bridge is loaded from the read-only file. */
const READONLY_SOURCE = [
    'Option Explicit',
    'Public Function C5D_ProbeReadonlyEcho(ByVal a As Variant) As Variant',
    '    C5D_ProbeReadonlyEcho = "readonly:" & CStr(a)',
    'End Function',
    ''
].join('\r\n');

/** Steps in run order. requires: skipped unless these steps (when selected in this run) matched. */
const STEPS = [
    { name: 'version', module: BRIDGE_MODULE, func: 'C5D_Version', args: [], expected: BRIDGE_VERSION },
    { name: 'install', module: BRIDGE_MODULE, func: 'C5D_Install', args: [PROBE_MODULE, PROBE_SOURCE], expected: 'ok', requires: ['version'] },
    { name: 'echo', module: PROBE_MODULE, func: 'C5D_ProbeEcho', args: [ECHO_ARGUMENT], expected: 'probe:Long:' + ECHO_ARGUMENT, requires: ['version', 'install'] },
    { name: 'projects', module: PROBE_MODULE, func: 'C5D_ProbeProjects', args: [], pattern: /^count=\d+;err=0$/, requires: ['version', 'install'] },
    { name: 'self-guard', module: BRIDGE_MODULE, func: 'C5D_Install', args: [BRIDGE_MODULE, 'x'], expected: 'error:self', requires: ['version'] }
];
const STEP_NAMES = STEPS.map(step => step.name);

const versionStep = name => ({ name, module: BRIDGE_MODULE, func: 'C5D_Version', args: [], expected: BRIDGE_VERSION });
const installUnloaderStep = (name, requires) => ({ name, module: BRIDGE_MODULE, func: 'C5D_Install', args: [UNLOAD_MODULE, UNLOAD_SOURCE], expected: 'ok', requires });
const unloadSelfStep = (name, requires) => ({ name, module: UNLOAD_MODULE, func: 'C5D_UnloadSelf', args: [], expected: UNLOAD_RESULT, requires });
/**
 * --unload-experiment, in run order. local: run in node (no CATIA call). expectFailure: the step matches unless the call
 * returns exactly this value. A COM exception does not stop the run (unload in particular); requires skips dependent steps.
 */
const UNLOAD_STEPS = [
    versionStep('pre'),
    installUnloaderStep('install-unloader'),
    { name: 'apc-info', module: UNLOAD_MODULE, func: 'C5D_ApcProjectsInfo', args: [], pattern: /^apcProjects=\d+/, requires: ['install-unloader'] },
    unloadSelfStep('unload', ['install-unloader']),
    { name: 'after-unload-probe', module: PROBE_MODULE, func: 'C5D_ProbeEcho', args: [UNLOADED_ECHO_ARGUMENT], expectFailure: 'probe:Long:' + UNLOADED_ECHO_ARGUMENT },
    versionStep('reload-version'),
    installUnloaderStep('install-unloader-2', ['reload-version']),
    unloadSelfStep('unload-2', ['install-unloader-2']),
    { name: 'set-readonly', local: 'set-readonly' },
    versionStep('readonly-version'),
    { name: 'readonly-install', module: BRIDGE_MODULE, func: 'C5D_Install', args: [READONLY_MODULE, READONLY_SOURCE], expected: 'ok' },
    { name: 'readonly-call', module: READONLY_MODULE, func: 'C5D_ProbeReadonlyEcho', args: [UNLOADED_ECHO_ARGUMENT], expected: 'readonly:' + UNLOADED_ECHO_ARGUMENT, requires: ['readonly-install'] }
];
/** --unload-cleanup: unloads C5D_Bridge (auto-loading it first if needed); the read-only attribute is left as it is. */
const CLEANUP_STEPS = [installUnloaderStep('cleanup-install'), unloadSelfStep('cleanup-unload', ['cleanup-install'])];

const READONLY_FRESH_RESULT = 'readonly-fresh:ok';
/**
 * VBA written into C5D_ProbeReadonly by --readonly-fresh: C5D_Probe plus C5D_ProbeProjects taken over from PROBE_SOURCE
 * (C5D_Bridge itself has no VBProjects function, so the projects step needs this module first).
 */
const READONLY_FRESH_SOURCE = [
    'Option Explicit',
    'Public Function C5D_Probe() As Variant',
    `    C5D_Probe = "${READONLY_FRESH_RESULT}"`,
    'End Function',
    PROBE_SOURCE.slice(PROBE_SOURCE.indexOf('Public Function C5D_ProbeProjects'))
].join('\r\n');
/**
 * --readonly-fresh, in run order (right after a CATIA start, VBA editor not opened; library = the read-only hash-named copy).
 * No requires: every step runs and is recorded even when an earlier one failed (only a timeout / crash / unreachable CATIA aborts).
 */
const READONLY_FRESH_STEPS = [
    versionStep('version'),
    { name: 'install', module: BRIDGE_MODULE, func: 'C5D_Install', args: [READONLY_MODULE, READONLY_FRESH_SOURCE], expected: 'ok' },
    { name: 'projects', module: READONLY_MODULE, func: 'C5D_ProbeProjects', args: [], pattern: /^count=\d+;err=0$/ },
    { name: 'call', module: READONLY_MODULE, func: 'C5D_Probe', args: [], expected: READONLY_FRESH_RESULT }
];

const USAGE = 'Usage: node scripts/probe-bridge.js [<step> ...]\n'
    + '       node scripts/probe-bridge.js --check-only\n'
    + '       node scripts/probe-bridge.js --unload-experiment | --unload-cleanup\n'
    + '       node scripts/probe-bridge.js --readonly-fresh | --readonly-fresh-check\n'
    + `  steps (default: all, always run in this order): ${STEP_NAMES.join(' ')}\n`
    + `  --unload-experiment: ${UNLOAD_STEPS.map(s => s.name).join(' ')} (no copy; stops with C5D_Bridge loaded read-only)\n`
    + `  --unload-cleanup: ${CLEANUP_STEPS.map(s => s.name).join(' ')} (read-only attribute stays)\n`
    + `  --readonly-fresh: read-only hash-named copy (${BRIDGE_PROJECT}-<sha256 8>.catvba), then ${READONLY_FRESH_STEPS.map(s => s.name).join(' ')}\n`
    + '    (run right after a CATIA start, before the VBA editor is opened)\n'
    + '  --readonly-fresh-check: only records the hash-named copy (time, size, read-only, lock, hash); CATIA is not called\n'
    + '  Never runs a .catvbs/.CATScript. --check-only only reads CNEXT.exe modules and the bridge files.';

const MODE_OPTIONS = { '--check-only': 'check-only', '--unload-experiment': 'unload-experiment', '--unload-cleanup': 'unload-cleanup',
    '--readonly-fresh': 'readonly-fresh', '--readonly-fresh-check': 'readonly-fresh-check' };
/** { mode: 'run' | 'check-only' | 'unload-experiment' | 'unload-cleanup' | 'readonly-fresh' | 'readonly-fresh-check' | 'help', steps }. Throws on bad arguments. */
function parseArgs(args) {
    if (args.some(arg => arg === '--help' || arg === '-h')) return { mode: 'help', steps: [] };
    const modes = new Set();
    const steps = new Set();
    for (const arg of args) {
        if (MODE_OPTIONS[arg]) modes.add(arg);
        else if (STEP_NAMES.includes(arg)) steps.add(arg);
        else throw new Error(`Unknown option or step: ${arg}\n${USAGE}`);
    }
    if (modes.size > 1) throw new Error(`${[...modes].join(' and ')} cannot be combined\n${USAGE}`);
    const [option] = modes;
    if (option && steps.size) throw new Error(`${option} takes no steps\n${USAGE}`);
    if (option === '--unload-experiment') return { mode: 'unload-experiment', steps: UNLOAD_STEPS.map(s => s.name) };
    if (option === '--unload-cleanup') return { mode: 'unload-cleanup', steps: CLEANUP_STEPS.map(s => s.name) };
    if (option === '--readonly-fresh') return { mode: 'readonly-fresh', steps: READONLY_FRESH_STEPS.map(s => s.name) };
    if (option) return { mode: MODE_OPTIONS[option], steps: [] };
    return { mode: 'run', steps: steps.size ? STEP_NAMES.filter(name => steps.has(name)) : [...STEP_NAMES] };
}

/** One command-line token for eval-probe.exe: b64:<UTF-8 base64> (survives any code page, CR/LF and quotes). */
const encodeToken = text => 'b64:' + Buffer.from(String(text), 'utf8').toString('base64');
/**
 * eval-probe.exe arguments of a step: [--no-retry] --label <name> vbacall <library> <module> <function> [args...].
 * noRetry: exactly one ExecuteScript call (no ByRef retry after an exception), for calls with side effects such as unloading.
 */
function callArgs(step, library, options = {}) {
    return [...(options.noRetry ? ['--no-retry'] : []), '--label', step.name, 'vbacall', ...[library, step.module, step.func, ...step.args].map(encodeToken)];
}
/** Whether a step's VBA return value is the expected one. */
function matchStep(step, result) {
    if (typeof result !== 'string') return false;
    return step.pattern ? step.pattern.test(result) : result === step.expected;
}
/** Whether an eval-probe.exe result line matches the step (expectFailure: anything but exactly that return value, incl. an exception). */
function matchCall(step, line) {
    if (!line) return false;
    if (step.expectFailure !== undefined) return !(line.ok === true && line.result === step.expectFailure);
    return line.ok === true && matchStep(step, line.result);
}
/** "U+XXXX" of every non-ASCII code point (to see how a Unicode text came back). */
function codepoints(text) {
    if (typeof text !== 'string') return '';
    return [...text].filter(ch => ch.codePointAt(0) > 0x7e).map(ch => 'U+' + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')).join(' ');
}

function bridgePaths(env = process.env, root = ROOT) {
    if (!env.LOCALAPPDATA) throw new Error('%LOCALAPPDATA% is not set');
    return { source: path.join(root, 'resources', 'bridge', BRIDGE_FILE), target: path.join(env.LOCALAPPDATA, 'Cat5Dev', 'bridge', BRIDGE_FILE) };
}
const sha256 = file => createHash('sha256').update(fs.readFileSync(file)).digest('hex');
/** { path, exists, size, modifiedAt, readOnly, sha256, error } without modifying the file (readOnly = Windows read-only attribute). */
function fileState(file) {
    const state = { path: file, exists: fs.existsSync(file), size: null, modifiedAt: null, readOnly: null, sha256: null, error: null };
    if (!state.exists) return state;
    try {
        const stat = fs.statSync(file);
        Object.assign(state, { size: stat.size, modifiedAt: stat.mtime.toISOString(), readOnly: (stat.mode & 0o200) === 0 });
        state.sha256 = sha256(file);
    } catch (error) { state.error = `${error.code ?? ''} ${error.message}`.trim(); }
    return state;
}
const errorText = error => `${error.code ?? ''} ${error.message}`.trim();
/**
 * Copies the bridge to the per-user folder (overwrites). A locked target is reported (locked: true), not thrown.
 * A read-only target (left by --unload-experiment) is made writable for the copy and read-only again afterwards (also when the copy fails).
 * Returns { ok, locked, code, message, wasReadOnly, readOnlyRestored, source, before, after }.
 */
function copyBridge(source, target, copy = fs.copyFileSync, chmod = fs.chmodSync) {
    if (!fs.existsSync(source)) throw new Error(`${source} does not exist (save the VBA project ${BRIDGE_PROJECT} there first)`);
    const result = { ok: false, locked: false, code: null, message: null, wasReadOnly: false, readOnlyRestored: null,
        source: fileState(source), before: fileState(target), after: null };
    fs.mkdirSync(path.dirname(target), { recursive: true });
    result.wasReadOnly = result.before.readOnly === true;
    const restore = () => {
        if (!result.wasReadOnly) return;
        try { chmod(target, 0o444); result.readOnlyRestored = true; } catch (error) {
            result.readOnlyRestored = false;
            result.message = [result.message, `read-only attribute not restored: ${errorText(error)}`].filter(Boolean).join('; ');
        }
    };
    try {
        if (result.wasReadOnly) chmod(target, 0o666);
        copy(source, target);
    } catch (error) {
        if (!LOCK_CODES.includes(error.code)) { restore(); throw error; }
        Object.assign(result, { locked: true, code: error.code, message: error.message });
        restore();
        result.after = fileState(target);
        return result;
    }
    restore();
    result.after = fileState(target);
    result.ok = result.after.sha256 !== null && result.after.sha256 === result.source.sha256;
    if (!result.ok) result.message = [result.message, 'hash of the copy differs from the source'].filter(Boolean).join('; ');
    return result;
}

/** C5D_Bridge-<first 8 hex of the sha256>.catvba (the production-like name of the --readonly-fresh copy). */
function hashedBridgeName(sourceSha256) {
    if (!/^[0-9a-f]{64}$/.test(String(sourceSha256))) throw new Error(`not a sha256: ${sourceSha256}`);
    return `${BRIDGE_PROJECT}-${sourceSha256.slice(0, 8)}.catvba`;
}
/** %LOCALAPPDATA%\Cat5Dev\bridge\C5D_Bridge-<8 hex>.catvba (never the plain C5D_Bridge.catvba of the other modes). */
function hashedBridgePath(sourceSha256, env = process.env) {
    if (!env.LOCALAPPDATA) throw new Error('%LOCALAPPDATA% is not set');
    return path.join(env.LOCALAPPDATA, 'Cat5Dev', 'bridge', hashedBridgeName(sourceSha256));
}
/** libuv's UV_FS_O_EXLOCK on Windows (share mode 0); not in fs.constants of node 22. */
const UV_FS_O_EXLOCK = 0x10000000;
/**
 * Whether another process has the file open: opens it for reading with share mode 0 and closes it (nothing written, works on a
 * read-only file). { inUse: false } | { inUse: true, code: 'EBUSY' } | { inUse: null, code } (other error) | { skipped }.
 */
function lockProbe(file, open = fs.openSync, close = fs.closeSync, platform = process.platform) {
    if (platform !== 'win32') return { skipped: 'not Windows' };
    if (!fs.existsSync(file)) return { skipped: 'missing' };
    try { close(open(file, fs.constants.O_RDONLY | UV_FS_O_EXLOCK)); return { inUse: false, code: null, message: null }; } catch (error) {
        return { inUse: error.code === 'EBUSY' ? true : null, code: error.code ?? null, message: error.message };
    }
}
/**
 * --readonly-fresh: puts the bridge at the hash-named target and sets the read-only attribute (before any CATIA call).
 * Same hash -> reused (only the attribute is ensured); missing -> copied; different -> replaced (read-only cleared first).
 * A locked target is reported (locked: true), not thrown. ok = target hash equals the source and the attribute is verified set.
 * Returns { ok, action, locked, code, message, source, before, readOnly (setReadOnly result), after }.
 */
function prepareHashedBridge(source, target, copy = fs.copyFileSync, chmod = fs.chmodSync) {
    if (!fs.existsSync(source)) throw new Error(`${source} does not exist (save the VBA project ${BRIDGE_PROJECT} there first)`);
    const result = { ok: false, action: null, locked: false, code: null, message: null, source: fileState(source), before: fileState(target), readOnly: null, after: null };
    fs.mkdirSync(path.dirname(target), { recursive: true });
    if (result.before.exists && result.before.sha256 === result.source.sha256) result.action = 'reused';
    else {
        result.action = result.before.exists ? 'replaced' : 'copied';
        try {
            if (result.before.readOnly) chmod(target, 0o666);
            copy(source, target);
        } catch (error) {
            if (!LOCK_CODES.includes(error.code)) throw error;
            Object.assign(result, { locked: true, code: error.code, message: error.message });
            if (result.before.readOnly) { try { chmod(target, 0o444); } catch { /* reported by after.readOnly */ } }
            result.after = fileState(target);
            return result;
        }
    }
    result.readOnly = setReadOnly(target, { chmod });
    result.after = result.readOnly.after ?? fileState(target);
    const sameHash = result.after.sha256 !== null && result.after.sha256 === result.source.sha256;
    result.ok = sameHash && result.after.readOnly === true;
    result.message = [!sameHash && 'hash of the target differs from the source', result.after.readOnly !== true && `read-only attribute not set (${result.readOnly.message})`]
        .filter(Boolean).join('; ') || null;
    return result;
}

/** CNEXT.exe processes with their start time (Win32_Process): [{ pid, startedAt, path }] or { error }. */
function cnextProcesses() {
    const script = [
        "$ErrorActionPreference = 'Stop'",
        '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false',
        '$items = @(Get-CimInstance -ClassName Win32_Process -Filter "Name = \'CNEXT.exe\'" | ForEach-Object {',
        "    $s = $null; if ($_.CreationDate) { $s = $_.CreationDate.ToString('o') }",
        '    @{ pid = [int]$_.ProcessId; startedAt = $s; path = $_.ExecutablePath }',
        '})',
        'ConvertTo-Json -InputObject @{ processes = $items } -Depth 4 -Compress'
    ].join('\n');
    try {
        const stdout = execFileSync(powershellPath(), ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand',
            Buffer.from(script, 'utf16le').toString('base64')], { windowsHide: true, timeout: 60000, stdio: ['ignore', 'pipe', 'pipe'] });
        const parsed = JSON.parse(stdout.toString('utf8').replace(/^﻿/, '').trim());
        const list = parsed.processes == null ? [] : Array.isArray(parsed.processes) ? parsed.processes : [parsed.processes];
        return { processes: list.map(p => ({ pid: p.pid, startedAt: p.startedAt ?? null, path: p.path ?? null })), error: null };
    } catch (error) {
        return { processes: [], error: String(error.stderr ? error.stderr.toString().trim() || error : error) };
    }
}

/** Runs one step in its own eval-probe.exe process. A timeout most likely means a modal dialog in CATIA. */
function runCallStep(exe, step, library, options = {}) {
    const args = callArgs(step, library, options);
    const started = Date.now();
    return new Promise(resolve => {
        execFile(exe, args, { windowsHide: true, timeout: STEP_TIMEOUT_MS, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }, (error, stdout, stderr) => {
            const parsed = parseProbeOutput(stdout || '');
            const timedOut = Boolean(error && error.killed);
            const result = parsed.lines.find(line => line.step === step.name) ?? parsed.lines[0] ?? null;
            resolve({
                step: step.name, call: `${step.module}.${step.func}`, exitCode: error ? (typeof error.code === 'number' ? error.code : null) : 0, timedOut,
                error: error && !timedOut && typeof error.code !== 'number' ? String(error) : null,
                result, match: matchCall(step, result),
                // what eval-probe.exe actually passed to ExecuteScript, compared with what was meant (node -> exe command line)
                argumentsIntact: result ? result.library === library && JSON.stringify(result.arguments ?? null) === JSON.stringify(step.args) : null,
                stdoutOther: parsed.other, stderr: (stderr || '').trim(), durationMs: Date.now() - started
            });
        });
    });
}

/** How a step ended, for verdicts. */
function describeOutcome(e) {
    return e.skipped ? `not run (${e.skipped})` : e.timedOut ? 'timed out (CATIA may be stopped at a dialog)'
        : e.ok ? `returned ${JSON.stringify(e.result ?? null)}` : `failed (${e.hresult ?? 'no result'})`;
}
/** "vbscript.dll ..." part of a verdict. */
function describeVbscript(beforeLoaded, entries) {
    if (beforeLoaded) return 'vbscript.dll undetermined (already loaded before the first step; restart CATIA, run nothing else, and re-run)';
    const first = entries.find(entry => entry.loadedAfter);
    return first ? `vbscript.dll loaded (first seen after step ${first.step})` : 'vbscript.dll not loaded';
}

/**
 * Verdict of an --unload-experiment / --unload-cleanup run. entries: [{ step, ok, match, result, hresult, message, timedOut, skipped,
 * crashed, loadedAfter }] in run order; set-readonly also has { readOnlyAfter, lockProbe: { writable, code } | { skipped } , code }.
 */
function judgeUnload(beforeLoaded, entries) {
    const get = name => entries.find(entry => entry.step === name);
    const ran = e => e && !e.skipped;
    const parts = [];
    const pre = get('pre');
    if (pre) parts.push(pre.match ? 'pre: bridge responded' : `pre: C5D_Version ${describeOutcome(pre)}`);
    for (const name of ['install-unloader', 'install-unloader-2', 'cleanup-install', 'readonly-install']) {
        const e = get(name);
        if (e && !e.match) parts.push(`${name}: C5D_Install ${describeOutcome(e)}`);
    }
    const info = get('apc-info');
    if (ran(info)) parts.push(info.ok ? `apc-info: ${info.result}` : `apc-info: ${describeOutcome(info)}`);
    for (const name of ['unload', 'unload-2', 'cleanup-unload']) {
        const e = get(name);
        if (!e) continue;
        parts.push(e.match ? `${name}: removed` : `${name}: C5D_UnloadSelf ${describeOutcome(e)}${e.message ? ' ' + e.message : ''}`);
    }
    const probe = get('after-unload-probe');
    if (ran(probe) && !probe.timedOut) {
        parts.push(probe.match ? `after-unload-probe: C5D_ProbeModule is gone (${probe.ok ? `returned ${JSON.stringify(probe.result)}` : `${probe.hresult} ${probe.message ?? ''}`.trim()})`
            : 'after-unload-probe: C5D_ProbeModule still answers -> NOT unloaded');
    }
    const reload = get('reload-version');
    if (ran(reload)) parts.push(reload.match ? 'reload-version: re-loaded by ExecuteScript' : `reload-version: C5D_Version ${describeOutcome(reload)}`);
    const ro = get('set-readonly');
    if (ran(ro)) {
        const lock = !ro.lockProbe ? 'not probed' : ro.lockProbe.skipped ? ro.lockProbe.skipped
            : ro.lockProbe.writable ? 'file not locked' : `file LOCKED (${ro.lockProbe.code})`;
        parts.push(ro.match ? `set-readonly: read-only set (${lock})` : `set-readonly: FAILED (${ro.code ?? 'attribute not set'}; ${lock})`);
    }
    const roVersion = get('readonly-version');
    if (ran(roVersion)) parts.push(roVersion.match ? 'readonly-version: read-only .catvba loads and answers' : `readonly-version: C5D_Version ${describeOutcome(roVersion)}`);
    const roCall = get('readonly-call');
    if (ran(roCall)) parts.push(roCall.match ? 'readonly-call: module installed into the read-only project works' : `readonly-call: ${describeOutcome(roCall)}`);
    const crashed = entries.find(e => e.crashed);
    if (crashed) parts.push(`CATIA crashed after step ${crashed.step}`);
    const timedOut = entries.find(e => e.timedOut);
    if (timedOut) parts.push(`timed out at step ${timedOut.step} (CATIA may be stopped at a dialog)`);
    parts.push(describeVbscript(beforeLoaded, entries));
    return parts.join(', ');
}

/**
 * Verdict of a run. beforeLoaded: vbscript.dll before the first step.
 * entries: [{ step, ok, match, result, hresult, timedOut, skipped, loadedAfter }] of the selected steps in run order
 * (skipped = reason when the step was not run).
 */
function judgeBridge(beforeLoaded, entries) {
    const get = name => entries.find(entry => entry.step === name);
    const outcome = describeOutcome;
    const parts = [];
    const version = get('version'), install = get('install'), echo = get('echo'), projects = get('projects'), guard = get('self-guard');
    if (version) parts.push(version.match ? 'bridge auto-loaded' : `bridge not reachable: C5D_Version ${outcome(version)}`);
    if (install) {
        if (!install.match) parts.push(`self-install failed: C5D_Install ${outcome(install)}`);
        else if (!echo) parts.push('self-install returned ok (echo not run)');
        else if (echo.match) parts.push('self-install works');
        else if (echo.ok && typeof echo.result === 'string' && echo.result.startsWith('probe:Long:')) {
            parts.push(`self-install works but Unicode did not round-trip (echo ${outcome(echo)})`);
        } else parts.push(`self-install returned ok but C5D_ProbeEcho ${outcome(echo)}`);
    } else if (echo) parts.push(echo.match ? 'C5D_ProbeEcho works (Unicode round-trip)' : `C5D_ProbeEcho ${outcome(echo)}`);
    if (projects && !projects.match) parts.push(`C5D_ProbeProjects ${outcome(projects)}`);
    if (guard) parts.push(guard.match ? 'self-guard works' : `self-guard FAILED: C5D_Install(${BRIDGE_MODULE}) ${outcome(guard)}`);
    parts.push(describeVbscript(beforeLoaded, entries));
    return parts.join(', ');
}

function snapshotter(report) {
    return label => (report.snapshots[label] = moduleSnapshot(label));
}
function checkCnext(report, first) {
    if (!first.cnextPids.length) throw new Error('CNEXT.exe is not running');
    if (first.cnextPids.length > 1) report.notes.push(`More than one CNEXT.exe is running (${first.cnextPids.join(', ')}); a DLL counts as loaded if any of them has it.`);
    if (first.powershell.error) report.notes.push(`Process.Modules could not be read (${first.powershell.error}); judged by tasklist /M only.`);
}

function runCheckOnly(report) {
    const paths = bridgePaths();
    report.bridge = { source: fileState(paths.source), target: fileState(paths.target) };
    report.cnext = cnextProcesses();
    const first = snapshotter(report)('current');
    checkCnext(report, first);
    if (first.loaded) report.notes.push('vbscript.dll is already loaded; restart CATIA (open only the VBA editor) before the run.');
    if (!report.bridge.source.exists) report.notes.push(`${paths.source} does not exist yet; the run would abort.`);
    report.summary = {
        dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, { current: first.dlls[dll].loaded }])),
        verdict: `check-only: vbscript.dll ${first.loaded ? 'LOADED' : 'not loaded'} in CNEXT.exe (${first.cnextPids.join(', ')})`
    };
}

async function runBridge(report, options) {
    const paths = bridgePaths();
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    const copy = copyBridge(paths.source, paths.target);
    report.bridge = { ...copy, library: paths.target };
    if (copy.locked) {
        const same = copy.after?.sha256 && copy.after.sha256 === copy.source.sha256;
        throw new Error(`${paths.target} is locked or read-only (${copy.code}); CATIA may have it loaded. `
            + `Close CATIA (or remove it from the macro libraries) and re-run. The locked file is ${same ? 'identical to' : 'DIFFERENT from'} the source.`);
    }
    if (!copy.ok) throw new Error(`Copy of the bridge failed: ${copy.message}`);
    report.library = paths.target;
    report.exe = buildProbe(ROOT);
    const snap = snapshotter(report);
    const first = snap('before');
    checkCnext(report, first);
    if (first.loaded) report.notes.push('vbscript.dll was already loaded before the first step; restart CATIA (and run nothing else) to judge the dependency. Step results are still valid.');
    report.runs = [];
    const entries = [];
    let aborted = null;
    for (const name of options.steps) {
        const step = STEPS.find(s => s.name === name);
        const failed = (step.requires ?? []).filter(dep => options.steps.includes(dep) && !entries.find(e => e.step === dep)?.match);
        if (aborted || failed.length) {
            const skipped = aborted ?? `requires ${failed.join(', ')}`;
            entries.push({ step: name, skipped });
            report.runs.push({ step: name, call: `${step.module}.${step.func}`, skipped });
            continue;
        }
        const run = await runCallStep(report.exe, step, paths.target);
        report.runs.push(run);
        const after = snap('after:' + name);
        const r = run.result;
        entries.push({ step: name, ok: r?.ok === true, match: run.match, result: r?.result ?? null, hresult: r?.hresult ?? null, timedOut: run.timedOut, loadedAfter: after.loaded });
        if (run.timedOut) {
            report.notes.push(`Step ${name} timed out after ${STEP_TIMEOUT_MS / 1000}s; CATIA may be stopped at a dialog (check CATIA). Remaining steps were not run.`);
            aborted = `aborted after ${name} timed out`;
        } else if (run.exitCode === 10 || run.exitCode === 11) {
            report.notes.push(`Step ${name}: could not reach CATIA (exit ${run.exitCode}); remaining steps were not run.`);
            aborted = `aborted: CATIA not reachable in ${name}`;
        }
        if (run.argumentsIntact === false) report.notes.push(`Step ${name}: the arguments eval-probe.exe passed differ from the intended ones (command-line encoding problem).`);
    }
    report.bridge.targetAtEnd = fileState(paths.target);
    if (report.bridge.targetAtEnd.sha256 && report.bridge.targetAtEnd.sha256 !== copy.after.sha256) {
        report.notes.push(`${paths.target} changed during the run (CATIA saved the project?).`);
    }
    if (entries.some(e => e.step === 'install' && e.match)) {
        report.notes.push(`${BRIDGE_PROJECT} in CATIA now has the unsaved module ${PROBE_MODULE}; do not save the project (quit CATIA without saving it).`);
    }
    const snapshots = Object.values(report.snapshots);
    const echoRun = report.runs.find(run => run.step === 'echo' && run.result);
    report.summary = {
        dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, Object.fromEntries(snapshots.map(s => [s.label, s.dlls[dll].loaded]))])),
        modulesAdded: snapshots.length > 1 ? addedModules(snapshots[0], snapshots[snapshots.length - 1]) : undefined,
        steps: report.runs.map(run => ({ step: run.step, call: run.call, skipped: run.skipped ?? null, timedOut: run.timedOut ?? false, exitCode: run.exitCode ?? null,
            ok: run.result?.ok ?? null, match: run.skipped ? null : run.match, result: run.result?.result ?? null, resultType: run.result?.resultType ?? null,
            hresult: run.result?.hresult ?? null, message: run.result?.message ?? null, paramMode: run.result?.paramMode ?? null, argumentsIntact: run.argumentsIntact ?? null })),
        echoCodepoints: echoRun ? { expected: codepoints(ECHO_ARGUMENT), returned: codepoints(echoRun.result.result) } : undefined,
        verdict: judgeBridge(first.loaded, entries)
    };
    if (entries.some(e => e.skipped || !e.match) || report.runs.some(run => run.exitCode !== undefined && run.exitCode !== 0)) process.exitCode = 1;
}

/**
 * set-readonly (local step): lock probe (open r+ and close, nothing written) and then the read-only attribute (chmod 0o444).
 * Returns { ok, code, message, before, lockProbe: { writable, code, message } | { skipped }, after, durationMs }.
 */
function setReadOnly(file, ops = {}) {
    const { chmod = fs.chmodSync, open = fs.openSync, close = fs.closeSync } = ops;
    const started = Date.now();
    const result = { ok: false, code: null, message: null, before: fileState(file), lockProbe: null, after: null, durationMs: 0 };
    if (!result.before.exists) {
        result.message = `${file} does not exist`;
    } else {
        if (result.before.readOnly) result.lockProbe = { skipped: 'already read-only' };
        else {
            try { close(open(file, 'r+')); result.lockProbe = { writable: true, code: null, message: null }; } catch (error) {
                result.lockProbe = { writable: false, code: error.code ?? null, message: error.message };
            }
        }
        try { chmod(file, 0o444); } catch (error) { Object.assign(result, { code: error.code ?? null, message: error.message }); }
        result.after = fileState(file);
        result.ok = result.after.readOnly === true;
        if (!result.ok && !result.message) result.message = 'read-only attribute is not set after chmod';
    }
    result.durationMs = Date.now() - started;
    return result;
}

/** Module snapshot that never throws: { snapshot } or { error }. */
function trySnapshot(report, label) {
    try { return { snapshot: (report.snapshots[label] = moduleSnapshot(label)) }; } catch (error) {
        report.snapshotErrors = [...(report.snapshotErrors ?? []), { label, error: String(error) }];
        return { error: String(error) };
    }
}
/** Whether any of the CNEXT.exe pids seen before the first step still runs (null = could not tell). */
function catiaAlive(pids, snapshot) {
    let current = snapshot?.cnextPids;
    if (!current) {
        const listed = cnextProcesses();
        if (listed.error) return null;
        current = listed.processes.map(p => p.pid);
    }
    return pids.some(pid => current.includes(pid));
}

/** --unload-experiment / --unload-cleanup. Uses the copy in %LOCALAPPDATA% as it is (no copy). */
async function runUnload(report, options) {
    const steps = options.mode === 'unload-cleanup' ? CLEANUP_STEPS : UNLOAD_STEPS;
    const paths = bridgePaths();
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    report.bridge = { source: fileState(paths.source), target: fileState(paths.target), targetAtEnd: null };
    if (!report.bridge.target.exists) throw new Error(`${paths.target} does not exist; run node scripts/probe-bridge.js (default mode) first`);
    report.library = paths.target;
    report.exe = buildProbe(ROOT);
    const first = trySnapshot(report, 'before');
    if (first.error) throw new Error(`module snapshot failed: ${first.error}`);
    checkCnext(report, first.snapshot);
    const pids = first.snapshot.cnextPids;
    if (first.snapshot.loaded) report.notes.push('vbscript.dll was already loaded before the first step.');
    const { entries, aborted } = await runSequence(report, steps, paths.target, pids, (step, earlier) => {
        const unload2 = earlier.find(e => e.step === 'unload-2');
        if (!unload2?.match) report.notes.push('unload-2 did not remove C5D_Bridge; it may still be loaded, so readonly-version may answer from the loaded project instead of auto-loading the read-only file.');
        const ro = setReadOnly(paths.target);
        return { run: { local: true, ...ro, match: ro.ok },
            entry: { ok: ro.ok, match: ro.ok, code: ro.code, message: ro.message, lockProbe: ro.lockProbe, readOnlyAfter: ro.after?.readOnly ?? null } };
    });
    report.bridge.targetAtEnd = fileState(paths.target);
    const startHash = report.bridge.target.sha256, endHash = report.bridge.targetAtEnd.sha256;
    if (startHash && endHash && startHash !== endHash) report.notes.push(`${paths.target} changed during the run (CATIA saved the project?).`);
    if (report.bridge.target.sha256 && report.bridge.source.sha256 && report.bridge.target.sha256 !== report.bridge.source.sha256) {
        report.notes.push(`${paths.target} differs from ${paths.source} at the start (saved from the VBA editor?).`);
    }
    report.stoppedForUser = options.mode === 'unload-experiment' && !aborted;
    if (report.stoppedForUser) {
        report.notes.push(`C5D_Bridge is now loaded from the read-only ${paths.target} with the unsaved module ${READONLY_MODULE}. `
            + 'Ask the user to open the VBA editor, select C5D_Bridge and try to save it (Save button / File > Save C5D_Bridge) and to note what happens '
            + '(error message, Save As dialog, or saved). Then run: node scripts/probe-bridge.js --unload-cleanup');
    }
    if (options.mode === 'unload-cleanup' && report.bridge.targetAtEnd.readOnly === false) report.notes.push(`${paths.target} is not read-only.`);
    report.summary = { ...sequenceSummary(report, entries), verdict: judgeUnload(first.snapshot.loaded, entries) };
    if (aborted) process.exitCode = 1;
}

/**
 * Runs steps one CATIA call (--no-retry) at a time, with a module snapshot and a CATIA-alive check after each. A timeout, a vanished
 * CNEXT.exe or an unreachable CATIA aborts the rest; a COM exception does not (requires skips dependent steps).
 * local(step, entries) runs a local step and returns { run, entry } (step/call are added here). Fills report.runs; returns { entries, aborted }.
 */
async function runSequence(report, steps, library, pids, local) {
    report.runs = [];
    const entries = [];
    let aborted = null;
    for (const step of steps) {
        const call = step.local ? `(local) ${step.local}` : `${step.module}.${step.func}`;
        const failed = (step.requires ?? []).filter(dep => !entries.find(e => e.step === dep)?.match);
        if (aborted || failed.length) {
            const skipped = aborted ?? `requires ${failed.join(', ')}`;
            entries.push({ step: step.name, skipped });
            report.runs.push({ step: step.name, call, skipped });
            continue;
        }
        let run, entry;
        if (step.local) {
            ({ run, entry } = local(step, entries));
            run = { step: step.name, call, ...run };
            entry = { step: step.name, ...entry };
        } else {
            // one ExecuteScript call per step: a ByRef retry after an exception would auto-load C5D_Bridge again behind an unload
            run = await runCallStep(report.exe, step, library, { noRetry: true });
            const r = run.result;
            entry = { step: step.name, ok: r?.ok === true, match: run.match, result: r?.result ?? null, hresult: r?.hresult ?? null, message: r?.message ?? null, timedOut: run.timedOut };
        }
        report.runs.push(run);
        const after = trySnapshot(report, 'after:' + step.name);
        entry.loadedAfter = after.snapshot?.loaded ?? null;
        run.catiaAlive = entry.alive = catiaAlive(pids, after.snapshot);
        entries.push(entry);
        if (entry.alive === false) {
            entry.crashed = run.crashed = true;
            report.notes.push(`CATIA crashed after step ${step.name} (CNEXT.exe ${pids.join(', ')} is gone). Remaining steps were not run.`);
            aborted = `aborted: CATIA crashed after step ${step.name}`;
        } else if (run.timedOut) {
            report.notes.push(`Step ${step.name} timed out after ${STEP_TIMEOUT_MS / 1000}s; CATIA may be stopped at a dialog (check CATIA). Remaining steps were not run.`);
            aborted = `aborted after ${step.name} timed out`;
        } else if (run.exitCode === 10 || run.exitCode === 11) {
            report.notes.push(`Step ${step.name}: could not reach CATIA (exit ${run.exitCode}); remaining steps were not run.`);
            aborted = `aborted: CATIA not reachable in ${step.name}`;
        }
        if (run.argumentsIntact === false) report.notes.push(`Step ${step.name}: the arguments eval-probe.exe passed differ from the intended ones (command-line encoding problem).`);
    }
    return { entries, aborted };
}
/** dlls / modulesAdded / steps part of report.summary after runSequence. */
function sequenceSummary(report, entries) {
    const snapshots = Object.values(report.snapshots);
    return {
        dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, Object.fromEntries(snapshots.map(s => [s.label, s.dlls[dll].loaded]))])),
        modulesAdded: snapshots.length > 1 ? addedModules(snapshots[0], snapshots[snapshots.length - 1]) : undefined,
        steps: report.runs.map(run => ({ step: run.step, call: run.call, skipped: run.skipped ?? null, timedOut: run.timedOut ?? false,
            exitCode: run.exitCode ?? null, ok: run.local ? run.ok : run.result?.ok ?? null, match: run.skipped ? null : run.match,
            result: run.result?.result ?? null, hresult: run.result?.hresult ?? null, message: run.local ? run.message : run.result?.message ?? null,
            elapsedMs: run.local ? run.durationMs : run.result?.elapsedMs ?? null, processMs: run.durationMs ?? null,
            catiaAlive: run.catiaAlive ?? null, vbscriptLoaded: entries.find(e => e.step === run.step)?.loadedAfter ?? null }))
    };
}

/** Verdict of --readonly-fresh. entries as in judgeUnload. */
function judgeReadonlyFresh(beforeLoaded, entries) {
    const get = name => entries.find(entry => entry.step === name);
    const parts = [];
    const version = get('version'), install = get('install'), projects = get('projects'), call = get('call');
    if (version) parts.push(version.match ? 'version: read-only .catvba auto-loaded and answers' : `version: C5D_Version ${describeOutcome(version)}`);
    if (install) parts.push(install.match ? 'install: ok' : `install: C5D_Install ${describeOutcome(install)}`);
    if (projects && !projects.skipped) parts.push(projects.ok ? `projects: ${projects.result}` : `projects: C5D_ProbeProjects ${describeOutcome(projects)}`);
    if (call) parts.push(call.match ? `call: ${READONLY_MODULE}.C5D_Probe works` : `call: C5D_Probe ${describeOutcome(call)}`);
    const crashed = entries.find(e => e.crashed);
    if (crashed) parts.push(`CATIA crashed after step ${crashed.step}`);
    const timedOut = entries.find(e => e.timedOut);
    if (timedOut) parts.push(`timed out at step ${timedOut.step} (CATIA may be stopped at a dialog)`);
    parts.push(describeVbscript(beforeLoaded, entries));
    return parts.join(', ');
}

/** The source bridge and its hash-named target path (throws when the source is missing or unreadable). */
function hashedPaths(env = process.env, root = ROOT) {
    const { source } = bridgePaths(env, root);
    if (!fs.existsSync(source)) throw new Error(`${source} does not exist (save the VBA project ${BRIDGE_PROJECT} there first)`);
    return { source, target: hashedBridgePath(sha256(source), env) };
}

/**
 * --readonly-fresh: hash-named read-only copy (attribute set and verified before any CATIA call), CNEXT.exe pid and vbscript.dll,
 * then READONLY_FRESH_STEPS, then the file facts again. Stops for the user (VBA editor save test) when no step aborted.
 */
async function runReadonlyFresh(report) {
    const paths = hashedPaths();
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    const prepared = prepareHashedBridge(paths.source, paths.target);
    report.bridge = { action: prepared.action, ok: prepared.ok, locked: prepared.locked, code: prepared.code, message: prepared.message,
        source: prepared.source, before: prepared.before, readOnly: prepared.readOnly, target: prepared.after, lock: lockProbe(paths.target),
        targetAtEnd: null, lockAtEnd: null };
    report.library = paths.target;
    if (prepared.locked) {
        throw new Error(`${paths.target} exists, differs from the source and cannot be replaced (${prepared.code}); CATIA may have it loaded. `
            + 'Close CATIA and re-run.');
    }
    if (!prepared.ok) throw new Error(`Preparing ${paths.target} failed: ${prepared.message}`);
    if (report.bridge.lock.inUse) report.notes.push(`${paths.target} is already open in another process before the first step (CATIA has it loaded?).`);
    report.exe = buildProbe(ROOT);
    const first = trySnapshot(report, 'before');
    if (first.error) throw new Error(`module snapshot failed: ${first.error}`);
    checkCnext(report, first.snapshot);
    const pids = first.snapshot.cnextPids;
    report.cnextPids = pids;
    report.vbscriptBefore = first.snapshot.loaded;
    if (first.snapshot.loaded) report.notes.push('vbscript.dll was already loaded before the first step (CATIA was not freshly started, or something ran a script).');
    const { entries, aborted } = await runSequence(report, READONLY_FRESH_STEPS, paths.target, pids);
    report.bridge.targetAtEnd = fileState(paths.target);
    report.bridge.lockAtEnd = lockProbe(paths.target);
    const startHash = report.bridge.target.sha256, endHash = report.bridge.targetAtEnd.sha256;
    if (startHash && endHash && startHash !== endHash) report.notes.push(`${paths.target} changed during the run (CATIA saved the project?).`);
    if (report.bridge.targetAtEnd.readOnly !== true) report.notes.push(`${paths.target} is NOT read-only at the end.`);
    report.stoppedForUser = !aborted;
    if (report.stoppedForUser) {
        report.notes.push(`C5D_Bridge should now be loaded from the read-only ${paths.target} with the unsaved module ${READONLY_MODULE}. `
            + 'Ask the user to open the VBA editor (Alt+F11), select C5D_Bridge in the project explorer, and report whether the Save button '
            + '(and File > Save C5D_Bridge) is grayed out, and what happens when saving (error message, Save As dialog, or saved without a message). '
            + 'Then run: node scripts/probe-bridge.js --readonly-fresh-check');
    }
    report.summary = { ...sequenceSummary(report, entries), verdict: judgeReadonlyFresh(first.snapshot.loaded, entries) };
    if (aborted) process.exitCode = 1;
}

/** Newest temp/probe-bridge-readonly-fresh-<time>.json (not a -check report) as { path, report }, or null. */
function latestReadonlyFreshReport(directory = TEMP) {
    let names;
    try { names = fs.readdirSync(directory); } catch { return null; }
    const latest = names.filter(name => /^probe-bridge-readonly-fresh-\d{8}-\d{6}(-\d+)?\.json$/.test(name))
        .sort((a, b) => a.localeCompare(b, 'en', { numeric: true })).pop();
    if (!latest) return null;
    const file = path.join(directory, latest);
    try { return { path: file, report: JSON.parse(fs.readFileSync(file, 'utf8')) }; } catch (error) { return { path: file, error: errorText(error) }; }
}
/**
 * --readonly-fresh-check (no CATIA call): the hash-named copy's time, size, read-only attribute, lock and hash (vs the source),
 * compared with the end of the latest --readonly-fresh report. deps: { env, root, cnext, reportsDir } for tests.
 */
function runReadonlyFreshCheck(report, deps = {}) {
    const { env = process.env, root = ROOT, cnext = cnextProcesses, reportsDir = TEMP } = deps;
    const paths = hashedPaths(env, root);
    const target = fileState(paths.target);
    report.library = paths.target;
    report.bridge = { source: fileState(paths.source), target, lock: lockProbe(paths.target), sameAsSource: null, previous: null };
    report.bridge.sameAsSource = target.sha256 ? target.sha256 === report.bridge.source.sha256 : null;
    report.cnext = cnext();
    if (!target.exists) report.notes.push(`${paths.target} does not exist; run node scripts/probe-bridge.js --readonly-fresh first.`);
    else if (!target.sha256) report.notes.push(`${paths.target} could not be read (${target.error}).`);
    else if (!report.bridge.sameAsSource) report.notes.push(`${paths.target} DIFFERS from the source (saved from the VBA editor?).`);
    if (target.exists && target.readOnly !== true) report.notes.push(`${paths.target} is NOT read-only.`);
    const latest = latestReadonlyFreshReport(reportsDir);
    const atEnd = latest?.report?.bridge?.targetAtEnd;
    if (latest) {
        report.bridge.previous = { report: latest.path, error: latest.error ?? null, targetAtEnd: atEnd ?? null,
            changed: atEnd && atEnd.path === paths.target ? atEnd.modifiedAt !== target.modifiedAt || atEnd.size !== target.size || atEnd.sha256 !== target.sha256 : null };
    }
    const changed = report.bridge.previous?.changed;
    report.summary = {
        verdict: !target.exists ? 'check: target missing'
            : `check: ${report.bridge.sameAsSource === true ? 'same as source' : report.bridge.sameAsSource === false ? 'DIFFERS from source' : 'hash unreadable'}, `
                + `read-only ${yn(target.readOnly)}, ${describeLock(report.bridge.lock)}`
                + `${changed === true ? ', CHANGED since the --readonly-fresh run' : changed === false ? ', unchanged since the --readonly-fresh run' : ''}`
    };
}
const describeLock = lock => !lock ? 'lock not probed' : lock.skipped ? `lock probe skipped (${lock.skipped})`
    : lock.inUse === true ? 'open in another process' : lock.inUse === false ? 'not open elsewhere' : `lock probe failed (${lock.code})`;

/** temp/<prefix>-<time>.json; never overwrites an earlier report. */
function uniqueReportPath(directory, prefix = 'probe-bridge') {
    const base = path.join(directory, `${prefix}-${timestamp()}`);
    let candidate = base + '.json';
    for (let i = 2; fs.existsSync(candidate); i++) candidate = `${base}-${i}.json`;
    return candidate;
}
function describeStart(startedAt, now = Date.now()) {
    if (!startedAt) return 'start time unknown';
    const minutes = Math.round((now - Date.parse(startedAt)) / 60000);
    return `started ${startedAt} (${minutes} min ago)`;
}
function formatSummary(report) {
    const hash = state => !state ? '-' : !state.exists ? '(missing)' : state.sha256 ? state.sha256.slice(0, 16) : `(unreadable: ${state.error})`;
    const lines = [`mode: ${report.mode}`, `report: ${report.reportPath}`];
    if (report.bridge) {
        const b = report.bridge;
        lines.push(`bridge source: ${b.source.path} sha256=${hash(b.source)}${b.source.size !== null ? ` ${b.source.size} bytes` : ''}`);
        if (report.mode === 'check-only') lines.push(`bridge target: ${b.target.path} sha256=${hash(b.target)}`);
        else {
            lines.push(`bridge target: ${b.after?.path ?? b.before.path} before=${hash(b.before)} after=${hash(b.after)}${b.locked ? ` LOCKED (${b.code})` : b.ok ? ' (copied, hash ok)' : ''}`);
            if (b.wasReadOnly) lines.push(`  target was read-only: made writable for the copy, read-only ${b.readOnlyRestored ? 'restored' : 'NOT restored'}`);
            if (b.targetAtEnd) lines.push(`  at end: ${hash(b.targetAtEnd)}${b.targetAtEnd.sha256 === b.after?.sha256 ? ' (unchanged)' : ' (CHANGED)'}`);
        }
    }
    if (report.error) lines.push(`error: ${report.error}`);
    if (report.cnext) {
        if (report.cnext.error) lines.push(`CNEXT.exe start time: ERROR ${report.cnext.error}`);
        for (const p of report.cnext.processes) lines.push(`CNEXT.exe pid ${p.pid}: ${describeStart(p.startedAt)}`);
    }
    formatSnapshots(lines, report);
    if (report.runs?.length) {
        lines.push('', `steps (library ${report.library}):`);
        for (const run of report.runs) {
            const r = run.result;
            if (run.skipped) lines.push(`  ${run.step}: SKIPPED (${run.skipped})`);
            else if (run.timedOut) lines.push(`  ${run.step}: TIMEOUT after ${STEP_TIMEOUT_MS / 1000}s - CATIA may be stopped at a dialog (check CATIA; later steps were skipped)`);
            else if (!r) lines.push(`  ${run.step}: no JSON output (exitCode=${run.exitCode}${run.error ? ', ' + run.error : ''}${run.stderr ? ', stderr: ' + run.stderr : ''})`);
            else {
                lines.push(`  ${run.step}: ok=${yn(r.ok)} match=${yn(run.match)} result=${JSON.stringify(r.result)} (${r.resultType}) paramMode=${r.paramMode} ${r.elapsedMs}ms`);
                lines.push(`    ${run.call}  hresult=${r.hresult}${r.message ? ' message=' + r.message : ''}  arguments intact=${yn(run.argumentsIntact)}`);
                if (r.attempts?.length > 1) lines.push(`    attempts: ${r.attempts.join(' / ')}`);
            }
        }
    }
    const s = report.summary;
    if (s?.echoCodepoints) lines.push('', `echo non-ASCII code points: expected ${s.echoCodepoints.expected} / returned ${s.echoCodepoints.returned || '(none)'}`);
    if (s?.verdict) lines.push('', `verdict: ${s.verdict}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    return lines.join('\n');
}
const yn = value => value === true ? 'yes' : value === false ? 'no' : '-';
/** Process.Modules line, the KEY_DLLS table of every snapshot and the modules added (appended to lines). */
function formatSnapshots(lines, report) {
    const snapshots = Object.values(report.snapshots);
    if (snapshots.length) {
        const first = snapshots[0], last = snapshots[snapshots.length - 1];
        const readable = first.processes.filter(p => !p.error && p.moduleCount > 0);
        lines.push(`Process.Modules: ${first.powershell.error ? 'ERROR ' + first.powershell.error
            : readable.length ? readable.map(p => `pid ${p.pid}: ${p.moduleCount} modules`).join(', ') : 'no readable process'}`);
        const short = label => label.replace(/^after:/, '+');
        const width = Math.max(...snapshots.map(s => short(s.label).length), 6) + 2;
        lines.push('', 'dll'.padEnd(14) + snapshots.map(s => short(s.label).padEnd(width)).join(''));
        for (const dll of KEY_DLLS) {
            lines.push(dll.padEnd(14) + snapshots.map(s => {
                const d = s.dlls[dll];
                const detail = d.byModules === null ? '(T)' : d.byModules !== d.byTasklist ? `(M:${yn(d.byModules)},T:${yn(d.byTasklist)})` : '';
                return (yn(d.loaded) + detail).padEnd(width);
            }).join(''));
        }
        lines.push('  (+x = after step x; yes/no = loaded in CNEXT.exe by Process.Modules or tasklist /M; (T) = tasklist only)');
        if (snapshots.length > 1) lines.push(`modules added ${first.label} -> ${last.label}: ${addedModules(first, last).join(', ') || '(none)'}`);
    }
}

/** Summary of --unload-experiment / --unload-cleanup. */
function formatUnloadSummary(report) {
    const hash = state => !state ? '-' : !state.exists ? '(missing)' : state.sha256 ? state.sha256.slice(0, 16) : `(unreadable: ${state.error})`;
    const fileLine = state => `sha256=${hash(state)} readOnly=${yn(state?.readOnly)}`;
    const lines = [`mode: ${report.mode}`, `report: ${report.reportPath}`];
    const b = report.bridge;
    if (b) {
        lines.push(`bridge source: ${b.source.path} ${fileLine(b.source)}`);
        lines.push(`bridge target: ${b.target.path} start: ${fileLine(b.target)}${b.target.sha256 && b.target.sha256 === b.source.sha256 ? ' (= source)' : ' (DIFFERS from source)'}`);
        if (b.action) {
            lines.push(`  prepared: ${b.action}${b.locked ? ` LOCKED (${b.code})` : ''}, read-only before CATIA: ${yn(b.target.readOnly)}, `
                + `${describeLock(b.lock)}${b.message ? `, ${b.message}` : ''}`);
            lines.push(`  ${factsLine(b.target)}`);
        }
        if (b.targetAtEnd) lines.push(`  at end: ${fileLine(b.targetAtEnd)}${b.targetAtEnd.sha256 === b.target.sha256 ? ' (unchanged)' : ' (CHANGED)'}`);
        if (b.targetAtEnd && b.action) lines.push(`  ${factsLine(b.targetAtEnd)}, ${describeLock(b.lockAtEnd)}`);
    }
    if (report.error) lines.push(`error: ${report.error}`);
    if (report.cnext) {
        if (report.cnext.error) lines.push(`CNEXT.exe start time: ERROR ${report.cnext.error}`);
        for (const p of report.cnext.processes) lines.push(`CNEXT.exe pid ${p.pid}: ${describeStart(p.startedAt)}`);
    }
    formatSnapshots(lines, report);
    for (const s of report.snapshotErrors ?? []) lines.push(`snapshot ${s.label} FAILED: ${s.error}`);
    if (report.runs?.length) {
        lines.push('', `steps (library ${report.library}):`);
        for (const run of report.runs) {
            const tail = `catia=${run.catiaAlive === true ? 'alive' : run.catiaAlive === false ? 'GONE' : '?'} vbscript.dll=${yn(report.snapshots['after:' + run.step]?.loaded)}`;
            const r = run.result;
            if (run.skipped) lines.push(`  ${run.step}: SKIPPED (${run.skipped})`);
            else if (run.local) {
                const lock = !run.lockProbe ? '-' : run.lockProbe.skipped ? run.lockProbe.skipped
                    : run.lockProbe.writable ? 'writable (not locked)' : `LOCKED (${run.lockProbe.code}: ${run.lockProbe.message})`;
                lines.push(`  ${run.step}: ok=${yn(run.ok)} readOnly ${yn(run.before?.readOnly)} -> ${yn(run.after?.readOnly)}  lock probe: ${lock}  ${run.durationMs}ms  ${tail}`);
                if (run.message) lines.push(`    message=${run.code ? run.code + ' ' : ''}${run.message}`);
            } else if (run.timedOut) lines.push(`  ${run.step}: TIMEOUT after ${STEP_TIMEOUT_MS / 1000}s - CATIA may be stopped at a dialog (check CATIA; later steps were skipped)  ${tail}`);
            else if (!r) lines.push(`  ${run.step}: no JSON output (exitCode=${run.exitCode}${run.error ? ', ' + run.error : ''}${run.stderr ? ', stderr: ' + run.stderr : ''})  ${tail}`);
            else {
                lines.push(`  ${run.step}: ok=${yn(r.ok)} match=${yn(run.match)} result=${JSON.stringify(r.result)} elapsed=${r.elapsedMs}ms (process ${run.durationMs}ms)  ${tail}`);
                lines.push(`    ${run.call}  hresult=${r.hresult}${r.message ? ' message=' + r.message : ''}${run.argumentsIntact === false ? '  ARGUMENTS NOT INTACT' : ''}`);
            }
        }
    }
    if (report.summary?.verdict) lines.push('', `verdict: ${report.summary.verdict}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    if (report.stoppedForUser && report.mode === 'readonly-fresh') {
        lines.push('', 'NEXT: ask the user to open the VBA editor, select C5D_Bridge, report whether the Save button is grayed out',
            '      and what happens on save, then run: node scripts/probe-bridge.js --readonly-fresh-check');
    } else if (report.stoppedForUser) {
        lines.push('', 'NEXT: ask the user to try saving C5D_Bridge in the VBA editor (Save button / File > Save C5D_Bridge),',
            '      record what happens, then run: node scripts/probe-bridge.js --unload-cleanup');
    }
    return lines.join('\n');
}
/** "modified=<iso> size=<n> readOnly=<yes/no>" of a fileState. */
const factsLine = state => `modified=${state?.modifiedAt ?? '-'} size=${state?.size ?? '-'} readOnly=${yn(state?.readOnly)}`;

/** Summary of --readonly-fresh-check. */
function formatReadonlyCheck(report) {
    const lines = [`mode: ${report.mode}`, `report: ${report.reportPath}`];
    const b = report.bridge;
    if (b) {
        const hash = state => !state.exists ? '(missing)' : state.sha256 ? state.sha256.slice(0, 16) : `(unreadable: ${state.error})`;
        lines.push(`bridge source: ${b.source.path} sha256=${hash(b.source)}`);
        lines.push(`bridge target: ${b.target.path} sha256=${hash(b.target)}${b.sameAsSource === true ? ' (= source)' : b.sameAsSource === false ? ' (DIFFERS from source)' : ''}`);
        lines.push(`  ${factsLine(b.target)}, ${describeLock(b.lock)}`);
        const p = b.previous;
        if (p?.targetAtEnd) lines.push(`  at end of ${p.report}: ${factsLine(p.targetAtEnd)}${p.changed ? ' (CHANGED since)' : p.changed === false ? ' (unchanged since)' : ''}`);
        else if (p) lines.push(`  previous report ${p.report}: ${p.error ?? 'no targetAtEnd'}`);
    }
    if (report.error) lines.push(`error: ${report.error}`);
    if (report.cnext) {
        if (report.cnext.error) lines.push(`CNEXT.exe start time: ERROR ${report.cnext.error}`);
        for (const p of report.cnext.processes) lines.push(`CNEXT.exe pid ${p.pid}: ${describeStart(p.startedAt)}`);
    }
    if (report.summary?.verdict) lines.push('', `verdict: ${report.summary.verdict}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    return lines.join('\n');
}

const REPORT_PREFIXES = { 'unload-experiment': 'probe-bridge-unload', 'unload-cleanup': 'probe-bridge-unload-cleanup',
    'readonly-fresh': 'probe-bridge-readonly-fresh', 'readonly-fresh-check': 'probe-bridge-readonly-fresh-check' };
/** Runner, summary formatter and report prefix of a parsed mode (not 'help'). */
function dispatch(options) {
    switch (options.mode) {
        case 'check-only': return { run: runCheckOnly, format: formatSummary, prefix: 'probe-bridge' };
        case 'unload-experiment': case 'unload-cleanup': return { run: runUnload, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'readonly-fresh': return { run: runReadonlyFresh, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'readonly-fresh-check': return { run: runReadonlyFreshCheck, format: formatReadonlyCheck, prefix: REPORT_PREFIXES[options.mode] };
        case 'run': return { run: runBridge, format: formatSummary, prefix: 'probe-bridge' };
        default: throw new Error(`no runner for mode ${options.mode}`);
    }
}
async function main() {
    let options;
    try { options = parseArgs(process.argv.slice(2)); } catch (error) { console.error(error.message); process.exitCode = 2; return; }
    if (options.mode === 'help') { console.log(USAGE); return; }
    const { run, format, prefix } = dispatch(options);
    fs.mkdirSync(TEMP, { recursive: true });
    const reportPath = uniqueReportPath(TEMP, prefix);
    const report = { mode: options.mode, steps: options.steps, startedAt: new Date().toISOString(), reportPath, project: BRIDGE_PROJECT,
        notes: [], snapshots: {} };
    try {
        await (options.mode === 'check-only' || options.mode === 'readonly-fresh-check' ? run(report) : run(report, options));
    } catch (error) {
        report.error = String(error); process.exitCode = 1;
    }
    report.finishedAt = new Date().toISOString();
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(format(report));
}
if (require.main === module) main();
module.exports = { parseArgs, encodeToken, callArgs, matchStep, matchCall, codepoints, bridgePaths, fileState, copyBridge, setReadOnly, judgeBridge,
    judgeUnload, describeStart, formatUnloadSummary, STEPS, STEP_NAMES, UNLOAD_STEPS, CLEANUP_STEPS, PROBE_SOURCE, UNLOAD_SOURCE, READONLY_SOURCE, ECHO_ARGUMENT,
    BRIDGE_MODULE, PROBE_MODULE, UNLOAD_MODULE, READONLY_MODULE, UNLOAD_RESULT,
    hashedBridgeName, hashedBridgePath, lockProbe, prepareHashedBridge, judgeReadonlyFresh, runReadonlyFresh, runReadonlyFreshCheck, formatReadonlyCheck,
    dispatch, READONLY_FRESH_STEPS, READONLY_FRESH_SOURCE, READONLY_FRESH_RESULT };
