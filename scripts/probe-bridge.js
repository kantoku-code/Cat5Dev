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
//
// Remove experiment (can C5D_Bridge be taken out of CATIA's VBA editor without quitting CATIA? VBProjects.Remove from inside gave 440):
//        node scripts/probe-bridge.js --remove-experiment         (freshly started CATIA; leaves CATIA changed -> restart it afterwards)
//        node scripts/probe-bridge.js --remove-experiment-risky   (CATIA MAY CRASH; freshly started CATIA, nothing unsaved)
//   Library A = the read-only hash-named copy of --readonly-fresh; --remove-experiment also prepares the read-only copy B
//   %LOCALAPPDATA%\Cat5Dev\bridge\C5D_Bridge-probe-b.catvba (same project name). C5D_ProbeRemove (REMOVE_SOURCE) is written into A;
//   "self" = the project that has that module. Steps (REMOVE_STEPS, danger low -> medium):
//     info: version install alive info-apc info-vbe (apc.Projects / VBE.VBProjects with Mode, Locked, Flags, Type, FileName, storage ...)
//     objects back to C#: obj-apc obj-apcproject obj-vbproject obj-array (eval-probe.exe vbacall-obj: the VBA function returns the
//       object, eval-probe.exe reads properties through IDispatch after ExecuteScript has returned)
//     attempts, each followed by <name>-check (C5D_RmAlive must be gone) and <name>-reload (C5D_Version answers again); once one
//       removes the project the later ones are skipped (unless): inside-remove (VBProjects.Remove from VBA, known 440), hold + held-remove
//       (object kept in a module variable, removed in a later call), outside-remove (C#: VBProject.Collection.Remove(VBProject)),
//       outside-close (C#: ApcProject.Close(False)), outside-design-close (Mode = 0 first), outside-end-close (Apc.End() first)
//     reinstall after-info-apc after-info-vbe, then (d) b-load b-info-apc b-install b-module-in-b b-module-in-a b-remove b-remove-info
//       b-close b-close-info b-reload final-info-apc (a second C5D_Bridge: does it load, where does C5D_Install write, can A remove/close B)
//   --remove-experiment-risky (REMOVE_RISKY_STEPS): inside-close (ApcProject.Close from its own code), then C5D_ProbeTimer
//     (TIMER_SOURCE: PtrSafe SetTimer + AddressOf): timer-remove / timer-close schedule the Remove / Close for 500 ms after
//     ExecuteScript returns, wait 3 s, collect C5D_RmTimerResult, check, reload; read-timer-log reads remove-timer.log.
//   Reports: temp/probe-bridge-remove-<time>.json / temp/probe-bridge-remove-risky-<time>.json (summary.verdict, summary.infos).
//   NOTE: the after-unload-probe of --unload-experiment called C5D_ProbeModule, which that mode never installs, so its "gone" was
//   no evidence of an unload; the remove steps check their own module C5D_ProbeRemove (installed and answered by step alive).
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

// ---- --remove-experiment / --remove-experiment-risky: can C5D_Bridge be taken out of CATIA's VBA editor without quitting CATIA?
// Type library findings (node scripts/dump-typelib.js -> temp/typelib-apc71-<time>.txt):
//   MSAPC.IApcProject.Close([opt] SaveChanges As Boolean = False)  DISPID 0x60020027  "Closes the project (does not remove the project from disk)"
//   MSAPC.IApcProjects.Close()  DISPID 0x60020009  closes ALL projects (never used here: it would close the user's projects too)
//   MSAPC.IApcProject.Mode (AxProjectMode: 0 design / 1 run, get+let), Locked, Flags, Dirty, Compiled, Storage(.Name = file name)
//   MSAPC.IApc.ExecutingProject, Running, End() ("Ends execution")
//   VBIDE._VBProjects.Remove(lpc As VBProject)  DISPID 0x8A (failed with 440 from inside); VBIDE._VBProject.Type (100 host / 101 stand-alone)
const REMOVE_MODULE = 'C5D_ProbeRemove';
const REMOVE_ALIVE = 'alive';
const B_MODULE = 'C5D_ProbeB';
const B_ALIVE = 'b-alive';
const TIMER_MODULE = 'C5D_ProbeTimer';
/** Second copy of the bridge for the (d) steps (same project name C5D_Bridge, different file). Not a C5D_Bridge-<8 hex> name. */
const COPY_B_FILE = 'C5D_Bridge-probe-b.catvba';
/** Log the SetTimer callback (risky mode) appends to, next to the bridge copies. */
const TIMER_LOG_FILE = 'remove-timer.log';
/** Placeholder in step args, replaced by the log path at run time (materializeSteps). */
const TIMER_LOG_TOKEN = '{timer-log}';
const INFO_APC_PATTERN = /^apc=[^;|]*;.*;count=\d+;err=0/;
const INFO_VBE_PATTERN = /^vbe=[^;|]*;.*;count=\d+;err=0/;

/** VBA: v = <expr>, then ";<label>=<value>" (or "=!<err>") appended to s. Properties are never read inside an If condition. */
const propLines = (label, expr, indent = '    ') => [
    `${indent}Err.Clear`, `${indent}v = "?"`, `${indent}v = ${expr}`, `${indent}e = Err.Number`, `${indent}s = s & ";${label}=" & RmFmt(v, e)`
];
/** VBA: s = <prefix> & error text, Exit Function when obj Is Nothing. */
const nothingGuard = (obj, fn, text) => [
    `    If ${obj} Is Nothing Then`,
    `        ${fn} = "${text}" & CStr(Err.Number) & ":" & RmClean(Err.Description)`,
    '        Exit Function',
    '    End If'
];
/** VBA: "<ok>" when Err.Number = 0 after the operation, else "error:<what>:<n>:<desc>" (+ suffix). */
const resultLines = (fn, ok, what, suffix = '""') => [
    '    e = Err.Number',
    '    d = Err.Description',
    '    If e = 0 Then',
    `        ${fn} = "${ok}" & ${suffix}`,
    '    Else',
    `        ${fn} = "error:${what}:" & CStr(e) & ":" & RmClean(d) & ${suffix}`,
    '    End If'
];
/**
 * VBA written into C5D_ProbeRemove (ASCII only, CRLF, On Error Resume Next everywhere, no MsgBox, no Declare).
 * "Self" = the project that has the module C5D_ProbeRemove (works with two projects named C5D_Bridge).
 *   C5D_RmAlive()               "alive" (removal check: after a removal + auto-reload this module is gone)
 *   C5D_RmApcInfo()             "apc=<TypeName>;apcVersion=..;vbaVersion=..;running=..;exec=..;count=N;err=E|name=..;display=..;...|..."
 *   C5D_RmVbeInfo()             "vbe=<TypeName>;version=..;count=N;err=E|name=..;type=..;mode=..;...|..."
 *   C5D_RmApc() / C5D_RmSelfApcProject() / C5D_RmSelfVbProject()   the object itself (Set) for vbacall-obj, else "error:..."
 *   C5D_RmSelfPair()            Array(VBE.VBProjects, self VBProject)
 *   C5D_RmRemoveSelfInside()    VBE.VBProjects.Remove self -> "removed;err=0" | "error:remove:<n>:<desc>"   (known: 440)
 *   C5D_RmHold() / C5D_RmRemoveHeld()  keep self VBProject in a module variable, remove it in a later ExecuteScript call
 *   C5D_RmRemoveOther(key) / C5D_RmCloseOther(key)   the other C5D_Bridge (storage/file name contains key, else the first
 *                               C5D_Bridge without C5D_ProbeRemove): VBProjects.Remove / ApcProject.Close False
 *                               -> "removed;via=key|fallback;err=0" / "closed;via=..;err=0" | "error:...;via=.."
 *   C5D_RmCloseSelfInside()     ApcProject(self).Close False from its own running code (risky mode only)
 */
const REMOVE_SOURCE = [
    'Option Explicit',
    `Private Const RM_MARKER As String = "${REMOVE_MODULE}"`,
    `Private Const RM_BMODULE As String = "${B_MODULE}"`,
    `Private Const RM_PROJECT As String = "${BRIDGE_PROJECT}"`,
    'Private mHeld As Object',
    `Public Function C5D_RmAlive() As Variant`,
    `    C5D_RmAlive = "${REMOVE_ALIVE}"`,
    'End Function',
    'Private Function RmClean(ByVal t As String) As String',
    '    t = Replace(t, "|", "/")',
    '    t = Replace(t, ";", ",")',
    '    t = Replace(t, vbCr, " ")',
    '    t = Replace(t, vbLf, " ")',
    '    RmClean = t',
    'End Function',
    'Private Function RmFmt(ByVal v As Variant, ByVal e As Long) As String',
    '    On Error Resume Next',
    '    Dim t As String',
    '    If e <> 0 Then',
    '        RmFmt = "!" & CStr(e)',
    '        Exit Function',
    '    End If',
    '    t = "?"',
    '    t = CStr(v)',
    '    RmFmt = RmClean(t)',
    'End Function',
    'Private Function RmApc() As Object',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Set apc = Nothing',
    '    Set apc = CreateObject("MSAPC.Apc.7.1")',
    '    If apc Is Nothing Then',
    '        Err.Clear',
    '        Set apc = CreateObject("MSAPC.Apc")',
    '    End If',
    '    Set RmApc = apc',
    'End Function',
    'Private Function RmHasModule(ByVal vbp As Object, ByVal moduleName As String) As Boolean',
    '    On Error Resume Next',
    '    Dim c As Object',
    '    Set c = Nothing',
    '    Set c = vbp.VBComponents.Item(moduleName)',
    '    RmHasModule = Not (c Is Nothing)',
    '    Err.Clear',
    'End Function',
    'Private Function RmSelfApc(ByVal apc As Object) As Object',
    '    On Error Resume Next',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim vbp As Object',
    '    Dim hit As Boolean',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Set RmSelfApc = Nothing',
    '    Set aps = Nothing',
    '    Set aps = apc.Projects',
    '    n = -1',
    '    n = aps.Count',
    '    For i = 1 To n',
    '        Set ap = Nothing',
    '        Set ap = aps.Item(i)',
    '        Set vbp = Nothing',
    '        Set vbp = ap.VBProject',
    '        hit = RmHasModule(vbp, RM_MARKER)',
    '        If hit Then',
    '            Set RmSelfApc = ap',
    '            Exit Function',
    '        End If',
    '    Next',
    'End Function',
    'Private Function RmSelfVbp(ByVal apc As Object) As Object',
    '    On Error Resume Next',
    '    Dim vbps As Object',
    '    Dim vbp As Object',
    '    Dim hit As Boolean',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Set RmSelfVbp = Nothing',
    '    Set vbps = Nothing',
    '    Set vbps = apc.VBE.VBProjects',
    '    n = -1',
    '    n = vbps.Count',
    '    For i = 1 To n',
    '        Set vbp = Nothing',
    '        Set vbp = vbps.Item(i)',
    '        hit = RmHasModule(vbp, RM_MARKER)',
    '        If hit Then',
    '            Set RmSelfVbp = vbp',
    '            Exit Function',
    '        End If',
    '    Next',
    'End Function',
    "' storage name chain of an ApcProject, innermost first: <name><<parent name>... (Parent is a Storage or the ApcGlobal object)",
    'Private Function RmStorage(ByVal ap As Object) As String',
    '    On Error Resume Next',
    '    Dim st As Object',
    '    Dim p As Object',
    '    Dim s As String',
    '    Dim nm As String',
    '    Dim tn As String',
    '    Dim i As Long',
    '    Set st = Nothing',
    '    Set st = ap.Storage',
    '    For i = 1 To 4',
    '        If st Is Nothing Then Exit For',
    '        nm = "?"',
    '        nm = CStr(st.Name)',
    '        s = s & "<" & nm',
    '        Set p = Nothing',
    '        Set p = st.Parent',
    '        Set st = Nothing',
    '        tn = ""',
    '        tn = TypeName(p)',
    '        If InStr(1, tn, "Storage", vbTextCompare) > 0 Then Set st = p',
    '    Next',
    '    Err.Clear',
    '    RmStorage = Mid(s, 2)',
    'End Function',
    'Private Function RmProjInfo(ByVal ap As Object, ByVal ex As Object) As String',
    '    On Error Resume Next',
    '    Dim s As String',
    '    Dim v As Variant',
    '    Dim e As Long',
    '    Dim vbp As Object',
    '    Dim same As Boolean',
    '    s = ""',
    ...propLines('name', 'ap.Name'),
    ...propLines('display', 'ap.DisplayName'),
    ...propLines('tn', 'TypeName(ap)'),
    ...propLines('mode', 'ap.Mode'),
    ...propLines('dirty', 'ap.Dirty'),
    ...propLines('locked', 'ap.Locked'),
    ...propLines('flags', 'ap.Flags'),
    ...propLines('compiled', 'ap.Compiled'),
    ...propLines('hostProt', 'ap.HostProtection'),
    ...propLines('userProt', 'ap.UserProtection'),
    ...propLines('events', 'ap.EventsFrozen'),
    '    Set vbp = Nothing',
    '    Set vbp = ap.VBProject',
    ...propLines('vbName', 'vbp.Name'),
    ...propLines('vbType', 'vbp.Type'),
    ...propLines('vbMode', 'vbp.Mode'),
    ...propLines('vbProt', 'vbp.Protection'),
    ...propLines('vbFile', 'vbp.FileName'),
    ...propLines('vbSaved', 'vbp.Saved'),
    '    s = s & ";marker=" & CStr(RmHasModule(vbp, RM_MARKER))',
    '    s = s & ";bModule=" & CStr(RmHasModule(vbp, RM_BMODULE))',
    '    same = False',
    '    same = (ap Is ex)',
    '    s = s & ";exec=" & CStr(same)',
    '    s = s & ";storage=" & RmClean(RmStorage(ap))',
    '    Err.Clear',
    '    RmProjInfo = Mid(s, 2)',
    'End Function',
    'Private Function RmVbpInfo(ByVal vbp As Object) As String',
    '    On Error Resume Next',
    '    Dim s As String',
    '    Dim v As Variant',
    '    Dim e As Long',
    '    s = ""',
    ...propLines('name', 'vbp.Name'),
    ...propLines('tn', 'TypeName(vbp)'),
    ...propLines('type', 'vbp.Type'),
    ...propLines('mode', 'vbp.Mode'),
    ...propLines('prot', 'vbp.Protection'),
    ...propLines('file', 'vbp.FileName'),
    ...propLines('saved', 'vbp.Saved'),
    '    s = s & ";marker=" & CStr(RmHasModule(vbp, RM_MARKER))',
    '    s = s & ";bModule=" & CStr(RmHasModule(vbp, RM_BMODULE))',
    '    Err.Clear',
    '    RmVbpInfo = Mid(s, 2)',
    'End Function',
    'Public Function C5D_RmApcInfo() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim ex As Object',
    '    Dim s As String',
    '    Dim v As Variant',
    '    Dim e As Long',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmApcInfo', 'apc=Nothing;err='),
    '    s = "apc=" & TypeName(apc)',
    ...propLines('apcVersion', 'apc.APCVersion'),
    ...propLines('vbaVersion', 'apc.VBAVersion'),
    ...propLines('running', 'apc.Running'),
    '    Set ex = Nothing',
    '    Set ex = apc.ExecutingProject',
    ...propLines('exec', 'ex.Name'),
    '    Err.Clear',
    '    Set aps = Nothing',
    '    Set aps = apc.Projects',
    '    n = -1',
    '    n = aps.Count',
    '    s = s & ";count=" & CStr(n) & ";err=" & CStr(Err.Number)',
    '    For i = 1 To n',
    '        Set ap = Nothing',
    '        Err.Clear',
    '        Set ap = aps.Item(i)',
    '        If ap Is Nothing Then',
    '            s = s & "|#" & CStr(i) & ":err=" & CStr(Err.Number)',
    '        Else',
    '            s = s & "|" & RmProjInfo(ap, ex)',
    '        End If',
    '    Next',
    '    C5D_RmApcInfo = s',
    'End Function',
    'Public Function C5D_RmVbeInfo() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ide As Object',
    '    Dim vbps As Object',
    '    Dim vbp As Object',
    '    Dim s As String',
    '    Dim v As Variant',
    '    Dim e As Long',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmVbeInfo', 'vbe=Nothing;apcErr='),
    '    Set ide = Nothing',
    '    Set ide = apc.VBE',
    ...nothingGuard('ide', 'C5D_RmVbeInfo', 'vbe=Nothing;err='),
    '    s = "vbe=" & TypeName(ide)',
    ...propLines('version', 'ide.Version'),
    '    Err.Clear',
    '    Set vbps = Nothing',
    '    Set vbps = ide.VBProjects',
    '    n = -1',
    '    n = vbps.Count',
    '    s = s & ";count=" & CStr(n) & ";err=" & CStr(Err.Number)',
    '    For i = 1 To n',
    '        Set vbp = Nothing',
    '        Err.Clear',
    '        Set vbp = vbps.Item(i)',
    '        If vbp Is Nothing Then',
    '            s = s & "|#" & CStr(i) & ":err=" & CStr(Err.Number)',
    '        Else',
    '            s = s & "|" & RmVbpInfo(vbp)',
    '        End If',
    '    Next',
    '    C5D_RmVbeInfo = s',
    'End Function',
    'Public Function C5D_RmApc() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmApc', 'error:apc:'),
    '    Set C5D_RmApc = apc',
    'End Function',
    'Public Function C5D_RmSelfApcProject() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ap As Object',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmSelfApcProject', 'error:apc:'),
    '    Set ap = RmSelfApc(apc)',
    ...nothingGuard('ap', 'C5D_RmSelfApcProject', 'error:notfound:'),
    '    Set C5D_RmSelfApcProject = ap',
    'End Function',
    'Public Function C5D_RmSelfVbProject() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim vbp As Object',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmSelfVbProject', 'error:apc:'),
    '    Set vbp = RmSelfVbp(apc)',
    ...nothingGuard('vbp', 'C5D_RmSelfVbProject', 'error:notfound:'),
    '    Set C5D_RmSelfVbProject = vbp',
    'End Function',
    'Public Function C5D_RmSelfPair() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim vbp As Object',
    '    Dim vbps As Object',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmSelfPair', 'error:apc:'),
    '    Set vbp = RmSelfVbp(apc)',
    ...nothingGuard('vbp', 'C5D_RmSelfPair', 'error:notfound:'),
    '    Set vbps = apc.VBE.VBProjects',
    '    C5D_RmSelfPair = Array(vbps, vbp)',
    'End Function',
    'Public Function C5D_RmRemoveSelfInside() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim vbp As Object',
    '    Dim vbps As Object',
    '    Dim e As Long',
    '    Dim d As String',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmRemoveSelfInside', 'error:apc:'),
    '    Set vbp = RmSelfVbp(apc)',
    ...nothingGuard('vbp', 'C5D_RmRemoveSelfInside', 'error:notfound:'),
    '    Set vbps = apc.VBE.VBProjects',
    '    Err.Clear',
    '    vbps.Remove vbp',
    ...resultLines('C5D_RmRemoveSelfInside', 'removed;err=0', 'remove'),
    'End Function',
    'Public Function C5D_RmHold() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmHold', 'error:apc:'),
    '    Set mHeld = Nothing',
    '    Set mHeld = RmSelfVbp(apc)',
    ...nothingGuard('mHeld', 'C5D_RmHold', 'error:notfound:'),
    '    C5D_RmHold = "held;name=" & CStr(mHeld.Name)',
    'End Function',
    'Public Function C5D_RmRemoveHeld() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim vbps As Object',
    '    Dim e As Long',
    '    Dim d As String',
    ...nothingGuard('mHeld', 'C5D_RmRemoveHeld', 'error:notheld:'),
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmRemoveHeld', 'error:apc:'),
    '    Set vbps = apc.VBE.VBProjects',
    '    Err.Clear',
    '    vbps.Remove mHeld',
    ...resultLines('C5D_RmRemoveHeld', 'removed;err=0', 'remove'),
    '    Set mHeld = Nothing',
    'End Function',
    'Private Function RmFindOther(ByVal apc As Object, ByVal key As String, ByRef via As String) As Object',
    '    On Error Resume Next',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim vbp As Object',
    '    Dim fallback As Object',
    '    Dim hit As Boolean',
    '    Dim nm As String',
    '    Dim st As String',
    '    Dim fl As String',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Set RmFindOther = Nothing',
    '    Set fallback = Nothing',
    '    via = "none"',
    '    Set aps = Nothing',
    '    Set aps = apc.Projects',
    '    n = -1',
    '    n = aps.Count',
    '    For i = 1 To n',
    '        Set ap = Nothing',
    '        Set ap = aps.Item(i)',
    '        Set vbp = Nothing',
    '        Set vbp = ap.VBProject',
    '        nm = ""',
    '        nm = CStr(ap.Name)',
    '        hit = RmHasModule(vbp, RM_MARKER)',
    '        If (Not hit) And nm = RM_PROJECT Then',
    '            st = ""',
    '            st = RmStorage(ap)',
    '            fl = ""',
    '            fl = CStr(vbp.FileName)',
    '            If InStr(1, st, key, vbTextCompare) > 0 Or InStr(1, fl, key, vbTextCompare) > 0 Then',
    '                via = "key"',
    '                Set RmFindOther = ap',
    '                Exit Function',
    '            End If',
    '            If fallback Is Nothing Then Set fallback = ap',
    '        End If',
    '    Next',
    '    If Not fallback Is Nothing Then',
    '        via = "fallback"',
    '        Set RmFindOther = fallback',
    '    End If',
    'End Function',
    'Public Function C5D_RmRemoveOther(ByVal key As Variant) As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ap As Object',
    '    Dim vbp As Object',
    '    Dim vbps As Object',
    '    Dim via As String',
    '    Dim e As Long',
    '    Dim d As String',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmRemoveOther', 'error:apc:'),
    '    Set ap = RmFindOther(apc, CStr(key), via)',
    ...nothingGuard('ap', 'C5D_RmRemoveOther', 'error:notfound:'),
    '    Set vbp = Nothing',
    '    Set vbp = ap.VBProject',
    '    Set vbps = apc.VBE.VBProjects',
    '    Err.Clear',
    '    vbps.Remove vbp',
    ...resultLines('C5D_RmRemoveOther', 'removed;err=0', 'remove', '";via=" & via'),
    'End Function',
    'Public Function C5D_RmCloseOther(ByVal key As Variant) As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ap As Object',
    '    Dim via As String',
    '    Dim e As Long',
    '    Dim d As String',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmCloseOther', 'error:apc:'),
    '    Set ap = RmFindOther(apc, CStr(key), via)',
    ...nothingGuard('ap', 'C5D_RmCloseOther', 'error:notfound:'),
    '    Err.Clear',
    '    ap.Close False',
    ...resultLines('C5D_RmCloseOther', 'closed;err=0', 'close', '";via=" & via'),
    'End Function',
    'Public Function C5D_RmCloseSelfInside() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ap As Object',
    '    Dim e As Long',
    '    Dim d As String',
    '    Set apc = RmApc()',
    ...nothingGuard('apc', 'C5D_RmCloseSelfInside', 'error:apc:'),
    '    Set ap = RmSelfApc(apc)',
    ...nothingGuard('ap', 'C5D_RmCloseSelfInside', 'error:notfound:'),
    '    Err.Clear',
    '    ap.Close False',
    ...resultLines('C5D_RmCloseSelfInside', 'closed;err=0', 'close'),
    'End Function',
    ''
].join('\r\n');
/** VBA written into C5D_ProbeB through the copy B (shows which of two projects named C5D_Bridge C5D_Install writes into). */
const B_SOURCE = [
    'Option Explicit',
    'Public Function C5D_BAlive() As Variant',
    `    C5D_BAlive = "${B_ALIVE}"`,
    'End Function',
    ''
].join('\r\n');
/**
 * VBA written into C5D_ProbeTimer (risky mode only; a separate module so the Declare lines never touch the safe mode).
 *   C5D_RmSchedule(mode, logPath, delayMs)  SetTimer(0, 0, delayMs, AddressOf C5D_RmTimerProc) -> "scheduled;mode=..;timer=<id>;err=0"
 *   C5D_RmTimerProc  (WM_TIMER, after ExecuteScript has returned) KillTimer, then mode "remove": VBE.VBProjects.Remove self /
 *                    mode "close": ApcProject(self).Close False; result kept in a module variable and appended to logPath
 *   C5D_RmTimerResult()  "result=<scheduled:..|remove:ok|close:ok|error:<mode>:<n>:<desc>|...>;pending=<True|False>"
 * Self = the project that has the module C5D_ProbeTimer.
 */
const TIMER_SOURCE = [
    'Option Explicit',
    'Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr',
    'Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long',
    `Private Const TM_MARKER As String = "${TIMER_MODULE}"`,
    'Private mTimer As LongPtr',
    'Private mMode As String',
    'Private mLog As String',
    'Private mResult As String',
    'Private Sub TmLog(ByVal t As String)',
    '    On Error Resume Next',
    '    Dim f As Integer',
    '    If Len(mLog) = 0 Then Exit Sub',
    '    f = FreeFile',
    '    Open mLog For Append As #f',
    '    Print #f, Format(Now, "yyyy-mm-dd hh:nn:ss") & " " & t',
    '    Close #f',
    'End Sub',
    'Private Function TmSelf(ByVal apc As Object) As Object',
    '    On Error Resume Next',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim c As Object',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Set TmSelf = Nothing',
    '    Set aps = Nothing',
    '    Set aps = apc.Projects',
    '    n = -1',
    '    n = aps.Count',
    '    For i = 1 To n',
    '        Set ap = Nothing',
    '        Set ap = aps.Item(i)',
    '        Set c = Nothing',
    '        Set c = ap.VBProject.VBComponents.Item(TM_MARKER)',
    '        If Not c Is Nothing Then',
    '            Set TmSelf = ap',
    '            Exit Function',
    '        End If',
    '    Next',
    'End Function',
    'Public Function C5D_RmSchedule(ByVal mode As Variant, ByVal logPath As Variant, ByVal delayMs As Variant) As Variant',
    '    On Error Resume Next',
    '    mMode = CStr(mode)',
    '    mLog = CStr(logPath)',
    '    mResult = "scheduled:" & mMode',
    '    TmLog "schedule:" & mMode',
    '    If mTimer <> 0 Then',
    '        KillTimer 0, mTimer',
    '        mTimer = 0',
    '    End If',
    '    Err.Clear',
    '    mTimer = SetTimer(0, 0, CLng(delayMs), AddressOf C5D_RmTimerProc)',
    '    C5D_RmSchedule = "scheduled;mode=" & mMode & ";timer=" & CStr(mTimer) & ";err=" & CStr(Err.Number)',
    'End Function',
    'Public Sub C5D_RmTimerProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ap As Object',
    '    Dim vbp As Object',
    '    Dim vbps As Object',
    '    Dim e As Long',
    '    Dim d As String',
    '    KillTimer 0, idEvent',
    '    mTimer = 0',
    '    TmLog "fired:" & mMode',
    '    Set apc = Nothing',
    '    Set apc = CreateObject("MSAPC.Apc.7.1")',
    '    If apc Is Nothing Then',
    '        Err.Clear',
    '        Set apc = CreateObject("MSAPC.Apc")',
    '    End If',
    '    If apc Is Nothing Then',
    '        mResult = "error:apc:" & CStr(Err.Number)',
    '        TmLog mResult',
    '        Exit Sub',
    '    End If',
    '    Set ap = TmSelf(apc)',
    '    If ap Is Nothing Then',
    '        mResult = "error:notfound"',
    '        TmLog mResult',
    '        Exit Sub',
    '    End If',
    '    TmLog "before:" & mMode',
    '    Err.Clear',
    '    If mMode = "close" Then',
    '        ap.Close False',
    '    Else',
    '        Set vbp = ap.VBProject',
    '        Set vbps = apc.VBE.VBProjects',
    '        Err.Clear',
    '        vbps.Remove vbp',
    '    End If',
    '    e = Err.Number',
    '    d = Err.Description',
    '    If e = 0 Then',
    '        mResult = mMode & ":ok"',
    '    Else',
    '        mResult = "error:" & mMode & ":" & CStr(e) & ":" & Replace(Replace(d, ";", ","), "|", "/")',
    '    End If',
    '    TmLog "after:" & mResult',
    'End Sub',
    'Public Function C5D_RmTimerResult() As Variant',
    '    C5D_RmTimerResult = "result=" & mResult & ";pending=" & CStr(mTimer <> 0)',
    'End Function',
    ''
].join('\r\n');

const rmStep = (name, func, extra = {}) => ({ name, module: REMOVE_MODULE, func, args: [], ...extra });
const installRemoveStep = (name, requires) => ({ name, module: BRIDGE_MODULE, func: 'C5D_Install', args: [REMOVE_MODULE, REMOVE_SOURCE], expected: 'ok', requires });
const infoApcStep = (name, requires, extra = {}) => rmStep(name, 'C5D_RmApcInfo', { pattern: INFO_APC_PATTERN, requires, info: 'apc', ...extra });
const infoVbeStep = (name, requires) => rmStep(name, 'C5D_RmVbeInfo', { pattern: INFO_VBE_PATTERN, requires, info: 'vbe' });
/**
 * vbacall-obj step: the VBA function returns an object; eval-probe.exe runs actions on it after ExecuteScript returned.
 * match: the call returned resultType (default ComObject) and every action index in mustSucceed is ok.
 */
const objStep = (name, func, actions, extra = {}) => rmStep(name, func, { actions, mustSucceed: [], resultType: 'ComObject', ...extra });
/** After a removal attempt: C5D_RmAlive must NOT answer "alive" (module gone = project removed and auto-reloaded from disk, or CATIA lost it). */
const removalCheck = attempt => rmStep(`${attempt}-check`, 'C5D_RmAlive', { expectFailure: REMOVE_ALIVE, after: [attempt], checkOf: attempt });
/** Only when the check says removed: does ExecuteScript still load the bridge (C5D_Version)? */
const reloadAfter = attempt => ({ ...versionStep(`${attempt}-reload`), requires: [`${attempt}-check`], reloadOf: attempt });
/** attempt + its check + its reload. */
const attemptSteps = step => [{ ...step, attempt: true }, removalCheck(step.name), reloadAfter(step.name)];
/** Checks of the attempts listed (an attempt is skipped once an earlier one removed the project). */
const checksOf = (...attempts) => attempts.map(name => `${name}-check`);

/**
 * --remove-experiment, in run order (danger: low < medium < high). Library = the read-only hash-named copy (A) unless library: 'b'
 * (the second copy COPY_B_FILE). requires: skipped unless those matched; unless: skipped when any of those matched (an earlier
 * attempt removed the project); after: skipped when that step was skipped. Never closes or removes a project other than C5D_Bridge.
 */
const REMOVE_STEPS = [
    versionStep('version'),
    installRemoveStep('install', ['version']),
    rmStep('alive', 'C5D_RmAlive', { expected: REMOVE_ALIVE, requires: ['install'] }),
    infoApcStep('info-apc', ['alive']),
    infoVbeStep('info-vbe', ['alive']),
    objStep('obj-apc', 'C5D_RmApc', ['type:$', 'get:$.APCVersion', 'get:$.VBAVersion', 'get:$.Running', 'get:$.Projects.Count', 'type:$.Projects',
        'type:$.VBE', 'get:$.VBE.VBProjects.Count', 'type:$.ExecutingProject'], { requires: ['alive'] }),
    objStep('obj-apcproject', 'C5D_RmSelfApcProject', ['type:$', 'get:$.Name', 'get:$.Mode', 'get:$.Locked', 'get:$.Flags', 'get:$.Dirty',
        'get:$.Compiled', 'get:$.Storage.Name', 'type:$.VBProject', 'get:$.VBProject.Name', 'get:$.VBProject.Type', 'get:$.VBProject.Mode',
        'get:$.Apc.Running', 'get:$.Apc.Projects.Count'], { requires: ['alive'] }),
    objStep('obj-vbproject', 'C5D_RmSelfVbProject', ['type:$', 'get:$.Name', 'get:$.Type', 'get:$.Mode', 'get:$.Protection', 'get:$.FileName',
        'get:$.Saved', 'type:$.Collection', 'get:$.Collection.Count'], { requires: ['alive'] }),
    objStep('obj-array', 'C5D_RmSelfPair', ['type:$0', 'type:$1', 'get:$0.Count', 'get:$1.Name'], { resultType: 'Object[]', requires: ['alive'] }),
    ...attemptSteps(rmStep('inside-remove', 'C5D_RmRemoveSelfInside', { expected: 'removed;err=0', requires: ['alive'], danger: 'low (known: 440)' })),
    rmStep('hold', 'C5D_RmHold', { pattern: /^held;/, requires: ['alive'], unless: checksOf('inside-remove') }),
    ...attemptSteps(rmStep('held-remove', 'C5D_RmRemoveHeld', { expected: 'removed;err=0', requires: ['hold'], unless: checksOf('inside-remove'), danger: 'low' })),
    ...attemptSteps(objStep('outside-remove', 'C5D_RmSelfVbProject', ['get:$.Name', 'get:$.Type', 'get:$.Collection.Count', 'call:$.Collection.Remove($)',
        'sleep:300', 'get:$.VBE.VBProjects.Count', 'get:$.Name'], { mustSucceed: [3], requires: ['obj-vbproject'],
        unless: checksOf('inside-remove', 'held-remove'), danger: 'low' })),
    ...attemptSteps(objStep('outside-close', 'C5D_RmSelfApcProject', ['get:$.Name', 'get:$.Mode', 'get:$.Apc.Projects.Count', 'call:$.Close(false)',
        'sleep:300', 'get:$.Apc.Projects.Count', 'get:$.Name'], { mustSucceed: [3], requires: ['obj-apcproject'],
        unless: checksOf('inside-remove', 'held-remove', 'outside-remove'), danger: 'medium' })),
    ...attemptSteps(objStep('outside-design-close', 'C5D_RmSelfApcProject', ['get:$.Mode', 'let:$.Mode=0', 'get:$.Mode', 'call:$.Close(false)',
        'sleep:300', 'get:$.Apc.Projects.Count', 'get:$.Name'], { mustSucceed: [3], requires: ['obj-apcproject'],
        unless: checksOf('inside-remove', 'held-remove', 'outside-remove', 'outside-close'), danger: 'medium' })),
    ...attemptSteps(objStep('outside-end-close', 'C5D_RmSelfApcProject', ['call:$.Apc.End()', 'sleep:300', 'get:$.Mode', 'call:$.Close(false)',
        'sleep:300', 'get:$.Apc.Projects.Count', 'get:$.Name'], { mustSucceed: [3], requires: ['obj-apcproject'],
        unless: checksOf('inside-remove', 'held-remove', 'outside-remove', 'outside-close', 'outside-design-close'),
        danger: 'medium (Apc.End resets module-level variables of every VBA project)' })),
    installRemoveStep('reinstall'),
    infoApcStep('after-info-apc', ['reinstall']),
    infoVbeStep('after-info-vbe', ['reinstall']),
    // (d) a second project with the same name C5D_Bridge (copy B): what CATIA does, which one C5D_Install writes into, and whether A can take B out
    { ...versionStep('b-load'), library: 'b' },
    infoApcStep('b-info-apc', ['reinstall', 'b-load']),
    { name: 'b-install', library: 'b', module: BRIDGE_MODULE, func: 'C5D_Install', args: [B_MODULE, B_SOURCE], expected: 'ok', requires: ['b-load'] },
    { name: 'b-module-in-b', library: 'b', module: B_MODULE, func: 'C5D_BAlive', args: [], expected: B_ALIVE, requires: ['b-install'] },
    { name: 'b-module-in-a', module: B_MODULE, func: 'C5D_BAlive', args: [], expected: B_ALIVE, requires: ['b-install'] },
    rmStep('b-remove', 'C5D_RmRemoveOther', { args: [COPY_B_FILE], pattern: /^removed;err=0;via=/, requires: ['reinstall', 'b-load'], danger: 'low' }),
    infoApcStep('b-remove-info', ['reinstall'], { after: ['b-remove'] }),
    rmStep('b-close', 'C5D_RmCloseOther', { args: [COPY_B_FILE], pattern: /^closed;err=0;via=/, requires: ['reinstall', 'b-load'], unless: ['b-remove'], danger: 'medium' }),
    infoApcStep('b-close-info', ['reinstall'], { after: ['b-close'] }),
    { ...versionStep('b-reload'), library: 'b', requires: ['b-load'] },
    infoApcStep('final-info-apc', ['reinstall'])
];

/**
 * --remove-experiment-risky (CATIA may crash; run on a freshly started CATIA, after --remove-experiment found nothing):
 * (c) ApcProject.Close from the project's own running code, then (e) SetTimer + AddressOf: Remove / Close run from a WM_TIMER
 * callback after ExecuteScript has returned. The callback's result stays in a module variable (collected by C5D_RmTimerResult)
 * and is appended to TIMER_LOG_FILE (read by the local step read-timer-log).
 */
const timerAttempt = (mode, unless) => [
    { name: `timer-${mode}`, module: TIMER_MODULE, func: 'C5D_RmSchedule', args: [mode, TIMER_LOG_TOKEN, '500'], pattern: /^scheduled;.*;timer=[1-9]\d*;err=0$/,
        requires: ['timer-alive'], unless, attempt: true, danger: mode === 'close' ? 'highest' : 'high' },
    { name: `timer-${mode}-wait`, local: 'wait', ms: 3000, after: [`timer-${mode}`] },
    // informational: answers only while C5D_ProbeTimer still exists (i.e. the callback's Remove / Close did NOT take effect);
    // after a real removal it fails with "macro program unknown" and the callback's outcome is only in the timer log
    { name: `timer-${mode}-result`, module: TIMER_MODULE, func: 'C5D_RmTimerResult', args: [], pattern: /^result=/, after: [`timer-${mode}`] },
    removalCheck(`timer-${mode}`),
    reloadAfter(`timer-${mode}`)
];
const REMOVE_RISKY_STEPS = [
    versionStep('version'),
    installRemoveStep('install', ['version']),
    rmStep('alive', 'C5D_RmAlive', { expected: REMOVE_ALIVE, requires: ['install'] }),
    { name: 'install-timer', module: BRIDGE_MODULE, func: 'C5D_Install', args: [TIMER_MODULE, TIMER_SOURCE], expected: 'ok', requires: ['alive'] },
    { name: 'timer-alive', module: TIMER_MODULE, func: 'C5D_RmTimerResult', args: [], pattern: /^result=;pending=False$/, requires: ['install-timer'] },
    infoApcStep('info-apc', ['alive']),
    ...attemptSteps(rmStep('inside-close', 'C5D_RmCloseSelfInside', { expected: 'closed;err=0', requires: ['alive'], danger: 'high' })),
    ...timerAttempt('remove', checksOf('inside-close')),
    ...timerAttempt('close', checksOf('inside-close', 'timer-remove')),
    { name: 'read-timer-log', local: 'read-timer-log' },
    installRemoveStep('reinstall'),
    infoApcStep('final-info-apc', ['reinstall'])
];
/** Steps with TIMER_LOG_TOKEN in args replaced by the log path. */
function materializeSteps(steps, values) {
    return steps.map(step => step.args?.includes(TIMER_LOG_TOKEN) ? { ...step, args: step.args.map(arg => arg === TIMER_LOG_TOKEN ? values.timerLog : arg) } : step);
}

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
    + '       node scripts/probe-bridge.js --remove-experiment | --remove-experiment-risky\n'
    + `  --remove-experiment: ${REMOVE_STEPS.map(s => s.name).join(' ')}\n`
    + '    (read-only hash-named copy A + read-only copy B; run on a freshly started CATIA; leaves CATIA changed, restart it afterwards)\n'
    + `  --remove-experiment-risky: ${REMOVE_RISKY_STEPS.map(s => s.name).join(' ')}\n`
    + '    (CATIA MAY CRASH: Close from its own code and SetTimer callbacks; only on a freshly started CATIA with nothing unsaved)\n'
    + '  Never runs a .catvbs/.CATScript. --check-only only reads CNEXT.exe modules and the bridge files.';

const MODE_OPTIONS = { '--check-only': 'check-only', '--unload-experiment': 'unload-experiment', '--unload-cleanup': 'unload-cleanup',
    '--readonly-fresh': 'readonly-fresh', '--readonly-fresh-check': 'readonly-fresh-check',
    '--remove-experiment': 'remove-experiment', '--remove-experiment-risky': 'remove-experiment-risky' };
/**
 * { mode: 'run' | 'check-only' | 'unload-experiment' | 'unload-cleanup' | 'readonly-fresh' | 'readonly-fresh-check' | 'remove-experiment'
 *   | 'remove-experiment-risky' | 'help', steps }. Throws on bad arguments.
 */
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
    if (option === '--remove-experiment') return { mode: 'remove-experiment', steps: REMOVE_STEPS.map(s => s.name) };
    if (option === '--remove-experiment-risky') return { mode: 'remove-experiment-risky', steps: REMOVE_RISKY_STEPS.map(s => s.name) };
    if (option) return { mode: MODE_OPTIONS[option], steps: [] };
    return { mode: 'run', steps: steps.size ? STEP_NAMES.filter(name => steps.has(name)) : [...STEP_NAMES] };
}

/** One command-line token for eval-probe.exe: b64:<UTF-8 base64> (survives any code page, CR/LF and quotes). */
const encodeToken = text => 'b64:' + Buffer.from(String(text), 'utf8').toString('base64');
/**
 * eval-probe.exe arguments of a step: [--no-retry] --label <name> vbacall <library> <module> <function> [args...].
 * noRetry: exactly one ExecuteScript call (no ByRef retry after an exception), for calls with side effects such as unloading.
 * A step with actions (objStep) uses vbacall-obj <library> <module> <function> [actions...] instead (always one call, no VBA args).
 */
function callArgs(step, library, options = {}) {
    if (step.actions) return ['--label', step.name, 'vbacall-obj', ...[library, step.module, step.func, ...step.actions].map(encodeToken)];
    return [...(options.noRetry ? ['--no-retry'] : []), '--label', step.name, 'vbacall', ...[library, step.module, step.func, ...step.args].map(encodeToken)];
}
/** Whether a step's VBA return value is the expected one. */
function matchStep(step, result) {
    if (typeof result !== 'string') return false;
    return step.pattern ? step.pattern.test(result) : result === step.expected;
}
/**
 * Whether an eval-probe.exe result line matches the step (expectFailure: anything but exactly that return value, incl. an exception).
 * Steps with actions: the call returned step.resultType (ComObject = an object came back) and every action in mustSucceed is ok.
 */
function matchCall(step, line, actions = []) {
    if (!line) return false;
    if (step.expectFailure !== undefined) return !(line.ok === true && line.result === step.expectFailure);
    if (step.actions) {
        return line.ok === true && line.resultType === step.resultType
            && (step.mustSucceed ?? []).every(index => actions.find(a => a.action === index)?.ok === true);
    }
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
            const result = parsed.lines.find(line => line.step === step.name && line.action === undefined) ?? parsed.lines.find(line => line.action === undefined) ?? null;
            const actions = step.actions ? parsed.lines.filter(line => line.step === step.name && line.action !== undefined) : undefined;
            resolve({
                step: step.name, call: `${step.module}.${step.func}`, library, exitCode: error ? (typeof error.code === 'number' ? error.code : null) : 0, timedOut,
                error: error && !timedOut && typeof error.code !== 'number' ? String(error) : null,
                result, ...(actions ? { actions } : {}), match: matchCall(step, result, actions),
                // what eval-probe.exe actually passed to ExecuteScript, compared with what was meant (node -> exe command line)
                argumentsIntact: result ? result.library === library && JSON.stringify(result.arguments ?? null) === JSON.stringify(step.args)
                    && (!step.actions || JSON.stringify(result.actions ?? null) === JSON.stringify(step.actions)) : null,
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
 * Why a step is not run (null = run it): an earlier abort; requires (a listed step did not match); unless (a listed step matched,
 * e.g. an earlier attempt already removed the project); after (a listed step was skipped or is not in this run).
 */
function skipReason(step, entries, aborted) {
    if (aborted) return aborted;
    const find = name => entries.find(e => e.step === name);
    const failed = (step.requires ?? []).filter(dep => !find(dep)?.match);
    if (failed.length) return `requires ${failed.join(', ')}`;
    const done = (step.unless ?? []).filter(dep => find(dep)?.match);
    if (done.length) return `not needed: ${done.join(', ')} matched`;
    const notRun = (step.after ?? []).filter(dep => !find(dep) || find(dep).skipped);
    if (notRun.length) return `${notRun.join(', ')} not run`;
    return null;
}

/**
 * Runs steps one CATIA call (--no-retry) at a time, with a module snapshot and a CATIA-alive check after each. A timeout, a vanished
 * CNEXT.exe or an unreachable CATIA aborts the rest; a COM exception does not (requires / unless / after skip dependent steps).
 * library: the .catvba path, or a function(step) returning it. local(step, entries) runs a local step (may be async) and returns
 * { run, entry } (step/call are added here). Fills report.runs; returns { entries, aborted }.
 */
async function runSequence(report, steps, library, pids, local) {
    report.runs = [];
    const entries = [];
    let aborted = null;
    for (const step of steps) {
        const call = step.local ? `(local) ${step.local}` : `${step.module}.${step.func}`;
        const skipped = skipReason(step, entries, aborted);
        if (skipped) {
            entries.push({ step: step.name, skipped });
            report.runs.push({ step: step.name, call, skipped });
            continue;
        }
        let run, entry;
        if (step.local) {
            ({ run, entry } = await local(step, entries));
            run = { step: step.name, call, ...run };
            entry = { step: step.name, ...entry };
        } else {
            // one ExecuteScript call per step: a ByRef retry after an exception would auto-load C5D_Bridge again behind an unload
            run = await runCallStep(report.exe, step, typeof library === 'function' ? library(step) : library, { noRetry: true });
            const r = run.result;
            entry = { step: step.name, ok: r?.ok === true, match: run.match, result: r?.result ?? null, hresult: r?.hresult ?? null, message: r?.message ?? null, timedOut: run.timedOut };
            if (run.actions) entry.actions = run.actions;
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
            catiaAlive: run.catiaAlive ?? null, vbscriptLoaded: entries.find(e => e.step === run.step)?.loadedAfter ?? null,
            ...(run.actions ? { resultType: run.result?.resultType ?? null, isDispatch: run.result?.isDispatch ?? null, typeInfo: run.result?.typeInfo ?? null,
                actions: run.actions.map(a => ({ action: a.action, op: a.op, ok: a.ok, result: a.result, resultType: a.resultType, typeInfo: a.typeInfo,
                    hresult: a.hresult, message: a.message })) } : {}) }))
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

/**
 * C5D_RmApcInfo / C5D_RmVbeInfo text -> { header: { key: value }, count, projects: [{ key: value }] } (null for a non-string).
 * Format: "<k>=<v>;<k>=<v>...|<project fields>|...". Values never contain | or ; (RmClean); the first = splits key and value.
 */
function parseInfo(text) {
    if (typeof text !== 'string') return null;
    const fields = part => Object.fromEntries(part.split(';').filter(Boolean).map(kv => {
        const i = kv.indexOf('=');
        return i < 0 ? [kv, ''] : [kv.slice(0, i), kv.slice(i + 1)];
    }));
    const [head, ...rest] = text.split('|');
    const header = fields(head);
    const count = /^\d+$/.test(header.count ?? '') ? Number(header.count) : null;
    return { header, count, projects: rest.map(fields) };
}
/** Projects named C5D_Bridge in a parsed info. */
const bridgeProjects = info => (info?.projects ?? []).filter(p => p.name === BRIDGE_PROJECT);
/** "N projects, C5D_Bridge x M [<storage file>, marker/B-module flags]" for verdicts. */
function describeInfo(entry) {
    if (!entry || entry.skipped) return entry ? `not run (${entry.skipped})` : 'not run';
    if (!entry.ok) return describeOutcome(entry);
    const info = parseInfo(entry.result);
    if (!info || info.count === null) return `unexpected ${JSON.stringify(entry.result)}`;
    const bridges = bridgeProjects(info).map(p => {
        const file = (p.storage || p.vbFile || p.file || '').split('<').pop().split(/[\\/]/).pop() || '?';
        const flags = [p.marker === 'True' && 'probe', p.bModule === 'True' && 'B-module', p.exec === 'True' && 'executing',
            p.vbType !== undefined && `type=${p.vbType}`, p.type !== undefined && `type=${p.type}`, p.mode !== undefined && `mode=${p.mode}`].filter(Boolean);
        return `${file}${flags.length ? ` (${flags.join(',')})` : ''}`;
    });
    return `${info.count} projects, ${BRIDGE_PROJECT} x${bridges.length}${bridges.length ? ` [${bridges.join('; ')}]` : ''}`;
}
/** "op=value" of the get/type actions of an object step (error -> "op=ERR <hresult>"). */
function describeActions(actions, kinds = ['get', 'type']) {
    return (actions ?? []).filter(a => kinds.includes(a.kind ?? a.op.split(':')[0])).map(a => {
        const op = a.op.replace(/^(get|type):/, (m, k) => k === 'type' ? 'type ' : '');
        return `${op}=${a.ok ? a.result ?? 'Empty' : `ERR ${a.hresult}${a.message ? ' ' + a.message.split(/\r?\n/)[0] : ''}`}`;
    }).join(', ');
}
/** Outcome of one removal attempt (+ its -check / -reload steps). */
function describeAttempt(step, get) {
    const e = get(step.name);
    if (!e) return `${step.name}: not in this run`;
    if (e.skipped) return `${step.name}: not run (${e.skipped})`;
    let outcome;
    if (e.timedOut) outcome = 'timed out (CATIA may be stopped at a dialog)';
    else if (step.actions) {
        if (!e.ok) outcome = `ExecuteScript failed (${e.hresult} ${e.message ?? ''})`.trim();
        else if (e.resultType !== step.resultType) outcome = `no object came back (resultType ${e.resultType ?? '?'}, result ${JSON.stringify(e.result)})`;
        else {
            const target = (e.actions ?? []).find(a => a.action === step.mustSucceed[0]);
            outcome = !target ? 'destructive action not run' : target.ok ? `${target.op} ok` : `${target.op} FAILED ${target.hresult} ${target.message ?? ''}`.trim();
            const counts = (e.actions ?? []).filter(a => /Count$/.test(a.op) || /\.Name$/.test(a.op)).map(a => `${a.op.replace(/^get:/, '')}=${a.ok ? a.result : 'ERR ' + a.hresult}`);
            if (counts.length) outcome += ` [${counts.join(', ')}]`;
        }
    } else outcome = describeOutcome(e);
    const check = get(`${step.name}-check`);
    const verdict = !check || check.skipped ? 'removal not checked' : check.timedOut ? 'check timed out'
        : check.match ? `REMOVED (C5D_RmAlive ${check.ok ? `returned ${JSON.stringify(check.result)}` : 'not found'})` : 'still loaded';
    const reload = get(`${step.name}-reload`);
    const reloaded = !reload || reload.skipped ? '' : reload.match ? ', bridge re-loads afterwards' : `, bridge does NOT re-load (${describeOutcome(reload)})`;
    return `${step.name}: ${outcome} -> ${verdict}${reloaded}`;
}

/**
 * Verdict of --remove-experiment / --remove-experiment-risky. entries as in judgeUnload (object steps also have resultType,
 * typeInfo, actions); steps = the step definitions of the run (attempt / actions / mustSucceed / resultType).
 */
function judgeRemove(beforeLoaded, entries, steps) {
    const get = name => entries.find(entry => entry.step === name);
    const ran = e => e && !e.skipped;
    const parts = [];
    for (const name of ['version', 'install', 'alive', 'install-timer', 'timer-alive', 'reinstall']) {
        const e = get(name);
        if (ran(e) && !e.match) parts.push(`${name}: ${describeOutcome(e)}`);
    }
    for (const name of ['info-apc', 'info-vbe']) {
        if (ran(get(name))) parts.push(`${name}: ${describeInfo(get(name))}`);
    }
    for (const name of ['obj-apc', 'obj-apcproject', 'obj-vbproject', 'obj-array']) {
        const e = get(name);
        if (!ran(e)) continue;
        if (e.timedOut || !e.ok) { parts.push(`${name}: ${describeOutcome(e)}`); continue; }
        const what = e.resultType === 'ComObject' ? `object came back (${e.typeInfo ?? 'no type info'}${e.isDispatch === false ? ', not IDispatch' : ''})`
            : `no object (resultType ${e.resultType}, result ${JSON.stringify(e.result)})`;
        parts.push(`${name}: ${what}; ${describeActions(e.actions)}`);
    }
    const attempts = steps.filter(step => step.attempt);
    for (const step of attempts) parts.push(describeAttempt(step, get));
    const removed = attempts.filter(step => get(`${step.name}-check`)?.match);
    parts.push(removed.length ? `C5D_Bridge REMOVED by ${removed.map(s => s.name).join(', ')}` : 'C5D_Bridge was not removed by any attempt');
    for (const name of ['timer-remove-result', 'timer-close-result']) {
        const e = get(name);
        if (ran(e)) parts.push(`${name}: ${describeOutcome(e)}`);
    }
    const log = get('read-timer-log');
    if (ran(log)) parts.push(`timer log: ${log.log === null ? 'missing' : JSON.stringify(log.log.trim().split(/\r?\n/).map(line => line.replace(/^\S+ \S+ /, '')).join(' / '))}`);
    for (const name of ['after-info-apc', 'after-info-vbe']) {
        if (ran(get(name))) parts.push(`${name}: ${describeInfo(get(name))}`);
    }
    const bLoad = get('b-load');
    if (ran(bLoad)) {
        parts.push(bLoad.match ? 'b-load: copy B auto-loaded and answers' : `b-load: ${describeOutcome(bLoad)}`);
        if (ran(get('b-info-apc'))) parts.push(`b-info-apc: ${describeInfo(get('b-info-apc'))}`);
        const inB = get('b-module-in-b'), inA = get('b-module-in-a');
        if (ran(inB) || ran(inA)) {
            const where = [inA?.match && 'A (first C5D_Bridge)', inB?.match && 'B'].filter(Boolean);
            parts.push(`C5D_Install through copy B wrote into: ${where.length ? where.join(' and ') : 'neither (or not callable)'}`);
        }
        const bRemove = get('b-remove'), bClose = get('b-close');
        if (bRemove) parts.push(`b-remove (A removes B): ${bRemove.skipped ? `not run (${bRemove.skipped})` : describeOutcome(bRemove)}; after: ${describeInfo(get('b-remove-info'))}`);
        if (bClose) parts.push(`b-close (A closes B): ${bClose.skipped ? `not run (${bClose.skipped})` : describeOutcome(bClose)}; after: ${describeInfo(get('b-close-info'))}`);
        const bReload = get('b-reload');
        if (ran(bReload)) parts.push(bReload.match ? 'b-reload: copy B answers again' : `b-reload: ${describeOutcome(bReload)}`);
    }
    if (ran(get('final-info-apc'))) parts.push(`final-info-apc: ${describeInfo(get('final-info-apc'))}`);
    const crashed = entries.find(e => e.crashed);
    if (crashed) parts.push(`CATIA crashed after step ${crashed.step}`);
    const timedOut = entries.find(e => e.timedOut);
    if (timedOut) parts.push(`timed out at step ${timedOut.step} (CATIA may be stopped at a dialog)`);
    parts.push(describeVbscript(beforeLoaded, entries));
    return parts.join(' | ');
}

/** report.bridge-shaped facts of a prepareHashedBridge result (as in --readonly-fresh). */
function preparedReport(prepared, target) {
    return { action: prepared.action, ok: prepared.ok, locked: prepared.locked, code: prepared.code, message: prepared.message, source: prepared.source,
        before: prepared.before, readOnly: prepared.readOnly, target: prepared.after, lock: lockProbe(target), targetAtEnd: null, lockAtEnd: null };
}
const readTextOrNull = file => { try { return fs.readFileSync(file, 'utf8'); } catch { return null; } };

/**
 * --remove-experiment / --remove-experiment-risky. A = the read-only hash-named copy (as --readonly-fresh, reused when the hash
 * matches); safe mode also prepares copy B (COPY_B_FILE, read-only); risky mode empties TIMER_LOG_FILE. Then REMOVE_STEPS /
 * REMOVE_RISKY_STEPS one --no-retry call at a time (a timeout, a vanished CNEXT.exe or an unreachable CATIA aborts the rest).
 */
async function runRemove(report, options) {
    const risky = options.mode === 'remove-experiment-risky';
    const paths = hashedPaths();
    const folder = path.dirname(paths.target);
    const copyB = path.join(folder, COPY_B_FILE);
    const timerLog = path.join(folder, TIMER_LOG_FILE);
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    const prepared = prepareHashedBridge(paths.source, paths.target);
    report.bridge = preparedReport(prepared, paths.target);
    report.library = paths.target;
    if (prepared.locked) throw new Error(`${paths.target} exists, differs from the source and cannot be replaced (${prepared.code}); CATIA may have it loaded. Close CATIA and re-run.`);
    if (!prepared.ok) throw new Error(`Preparing ${paths.target} failed: ${prepared.message}`);
    if (report.bridge.lock.inUse) report.notes.push(`${paths.target} is already open in another process (CATIA has C5D_Bridge loaded from an earlier run?); restart CATIA for a clean run.`);
    if (!risky) {
        const preparedB = prepareHashedBridge(paths.source, copyB);
        report.bridgeB = preparedReport(preparedB, copyB);
        report.libraryB = copyB;
        if (preparedB.locked) throw new Error(`${copyB} exists, differs from the source and cannot be replaced (${preparedB.code}); restart CATIA and re-run.`);
        if (!preparedB.ok) throw new Error(`Preparing ${copyB} failed: ${preparedB.message}`);
        if (report.bridgeB.lock.inUse) report.notes.push(`${copyB} is already open in another process (loaded by an earlier run?); restart CATIA for a clean run.`);
    } else {
        fs.mkdirSync(folder, { recursive: true });
        fs.writeFileSync(timerLog, '');
        report.timerLog = timerLog;
    }
    report.exe = buildProbe(ROOT);
    const first = trySnapshot(report, 'before');
    if (first.error) throw new Error(`module snapshot failed: ${first.error}`);
    checkCnext(report, first.snapshot);
    const pids = first.snapshot.cnextPids;
    report.cnextPids = pids;
    report.vbscriptBefore = first.snapshot.loaded;
    if (first.snapshot.loaded) report.notes.push('vbscript.dll was already loaded before the first step.');
    const steps = materializeSteps(risky ? REMOVE_RISKY_STEPS : REMOVE_STEPS, { timerLog });
    const { entries, aborted } = await runSequence(report, steps, step => step.library === 'b' ? copyB : paths.target, pids, async step => {
        if (step.local === 'wait') {
            await new Promise(resolve => setTimeout(resolve, step.ms));
            return { run: { local: true, kind: 'wait', ok: true, match: true, durationMs: step.ms }, entry: { ok: true, match: true } };
        }
        if (step.local === 'read-timer-log') {
            const log = readTextOrNull(timerLog);
            return { run: { local: true, kind: 'read-timer-log', ok: log !== null, match: log !== null, log, durationMs: 0, message: log === null ? `${timerLog} missing` : null },
                entry: { ok: log !== null, match: log !== null, log } };
        }
        throw new Error(`unknown local step ${step.local}`);
    });
    for (const entry of entries) {
        const run = report.runs.find(r => r.step === entry.step);
        if (run?.actions) Object.assign(entry, { resultType: run.result?.resultType ?? null, typeInfo: run.result?.typeInfo ?? null, isDispatch: run.result?.isDispatch ?? null });
    }
    report.bridge.targetAtEnd = fileState(paths.target);
    report.bridge.lockAtEnd = lockProbe(paths.target);
    if (report.bridgeB) { report.bridgeB.targetAtEnd = fileState(copyB); report.bridgeB.lockAtEnd = lockProbe(copyB); }
    for (const b of [report.bridge, report.bridgeB].filter(Boolean)) {
        if (b.target?.sha256 && b.targetAtEnd.sha256 && b.target.sha256 !== b.targetAtEnd.sha256) report.notes.push(`${b.targetAtEnd.path} changed during the run (CATIA saved the project?).`);
        if (b.targetAtEnd.exists && b.targetAtEnd.readOnly !== true) report.notes.push(`${b.targetAtEnd.path} is NOT read-only at the end.`);
    }
    report.notes.push('CATIA was changed by this run (probe modules in C5D_Bridge, maybe a second C5D_Bridge from copy B or a closed project). '
        + 'Quit CATIA without saving C5D_Bridge and restart it before any other probe run.');
    report.summary = { ...sequenceSummary(report, entries),
        infos: Object.fromEntries(steps.filter(s => s.info).map(s => [s.name, parseInfo(entries.find(e => e.step === s.name)?.result)])),
        verdict: judgeRemove(first.snapshot.loaded, entries, steps) };
    if (aborted) process.exitCode = 1;
}

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
    const bb = report.bridgeB;
    if (bb) {
        lines.push(`bridge copy B: ${bb.target?.path ?? bb.before.path} prepared: ${bb.action}${bb.locked ? ` LOCKED (${bb.code})` : ''}, `
            + `read-only before CATIA: ${yn(bb.target?.readOnly)}, ${describeLock(bb.lock)}${bb.message ? `, ${bb.message}` : ''}`);
        if (bb.targetAtEnd) lines.push(`  at end: ${fileLine(bb.targetAtEnd)}${bb.targetAtEnd.sha256 === bb.target?.sha256 ? ' (unchanged)' : ' (CHANGED)'}, ${describeLock(bb.lockAtEnd)}`);
    }
    if (report.timerLog) lines.push(`timer log: ${report.timerLog}`);
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
            else if (run.kind === 'wait') lines.push(`  ${run.step}: waited ${run.durationMs}ms  ${tail}`);
            else if (run.kind === 'read-timer-log') {
                lines.push(`  ${run.step}: ${run.log === null ? 'log MISSING' : `${run.log.split(/\r?\n/).filter(Boolean).length} lines`}  ${tail}`);
                for (const line of (run.log ?? '').split(/\r?\n/).filter(Boolean)) lines.push(`    | ${line}`);
            } else if (run.local) {
                const lock = !run.lockProbe ? '-' : run.lockProbe.skipped ? run.lockProbe.skipped
                    : run.lockProbe.writable ? 'writable (not locked)' : `LOCKED (${run.lockProbe.code}: ${run.lockProbe.message})`;
                lines.push(`  ${run.step}: ok=${yn(run.ok)} readOnly ${yn(run.before?.readOnly)} -> ${yn(run.after?.readOnly)}  lock probe: ${lock}  ${run.durationMs}ms  ${tail}`);
                if (run.message) lines.push(`    message=${run.code ? run.code + ' ' : ''}${run.message}`);
            } else if (run.timedOut) lines.push(`  ${run.step}: TIMEOUT after ${STEP_TIMEOUT_MS / 1000}s - CATIA may be stopped at a dialog (check CATIA; later steps were skipped)  ${tail}`);
            else if (!r) lines.push(`  ${run.step}: no JSON output (exitCode=${run.exitCode}${run.error ? ', ' + run.error : ''}${run.stderr ? ', stderr: ' + run.stderr : ''})  ${tail}`);
            else {
                lines.push(`  ${run.step}: ok=${yn(r.ok)} match=${yn(run.match)} result=${JSON.stringify(r.result)} elapsed=${r.elapsedMs}ms (process ${run.durationMs}ms)  ${tail}`);
                lines.push(`    ${run.call}  hresult=${r.hresult}${r.message ? ' message=' + r.message : ''}${run.argumentsIntact === false ? '  ARGUMENTS NOT INTACT' : ''}`);
                if (run.actions) {
                    lines.push(`    returned ${r.resultType}${r.typeInfo ? ' ' + r.typeInfo : ''}${r.isDispatch !== undefined && r.isDispatch !== null ? ` isDispatch=${yn(r.isDispatch)}` : ''}`
                        + `${r.arrayLength !== undefined ? ` array[${r.arrayLength}] ${(r.elements ?? []).join(', ')}` : ''}`);
                    for (const a of run.actions) {
                        lines.push(`      [${a.action}] ${a.op} -> ${a.skipped ? 'not run' : a.ok ? `ok ${JSON.stringify(a.result)} (${a.resultType}${a.typeInfo ? ' ' + a.typeInfo : ''})`
                            : `FAILED ${a.hresult} ${(a.message ?? '').split(/\r?\n/)[0]}${a.exception ? ` [${a.exception}]` : ''}`} ${a.elapsedMs}ms`);
                    }
                }
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
    'readonly-fresh': 'probe-bridge-readonly-fresh', 'readonly-fresh-check': 'probe-bridge-readonly-fresh-check',
    'remove-experiment': 'probe-bridge-remove', 'remove-experiment-risky': 'probe-bridge-remove-risky' };
/** Runner, summary formatter and report prefix of a parsed mode (not 'help'). */
function dispatch(options) {
    switch (options.mode) {
        case 'check-only': return { run: runCheckOnly, format: formatSummary, prefix: 'probe-bridge' };
        case 'unload-experiment': case 'unload-cleanup': return { run: runUnload, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'readonly-fresh': return { run: runReadonlyFresh, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'readonly-fresh-check': return { run: runReadonlyFreshCheck, format: formatReadonlyCheck, prefix: REPORT_PREFIXES[options.mode] };
        case 'remove-experiment': case 'remove-experiment-risky': return { run: runRemove, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
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
    dispatch, READONLY_FRESH_STEPS, READONLY_FRESH_SOURCE, READONLY_FRESH_RESULT,
    REMOVE_STEPS, REMOVE_RISKY_STEPS, REMOVE_SOURCE, B_SOURCE, TIMER_SOURCE, REMOVE_MODULE, B_MODULE, TIMER_MODULE, COPY_B_FILE, TIMER_LOG_FILE,
    TIMER_LOG_TOKEN, materializeSteps, skipReason, parseInfo, describeInfo, judgeRemove, runRemove, runSequence };
