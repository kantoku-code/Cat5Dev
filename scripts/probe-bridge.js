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
//
// Reload experiment (after an outside Close, the same path string gives 0x80004005 "macro program unknown": how can Push load the
// bridge again?). Continues in the SAME CATIA the remove experiment left (A closed, copy B loaded); does not restart anything:
//        node scripts/probe-bridge.js --reload-experiment
//   Local preparation (no CATIA call): facts + lock of A and B (never modified), read-only copies
//   C5D_Bridge-probe-<yyyyMMddHHmmss>-{c1,c2,c3,z}.catvba (new names every run, so a re-run never meets a "known" path), a hard link
//   C5D_Bridge-<8 hex>-hl-<time>.catvba of A, the 8.3 name of A (Scripting.FileSystemObject in PowerShell).
//   Each project gets its own module C5D_Rl_<tag> (reloadSource: C5D_RlAlive_/C5D_RlInfo_/C5D_RlVbe_/C5D_RlFind_<tag>; unique public
//   names). Info (apc.Projects with DisplayName, VBProject.FileName, C5D_* modules) and Close go through B: C5D_RlFind_b(<path>) returns
//   the C5D_Bridge project of that path and eval-probe.exe calls its Close(false) after ExecuteScript returned (vbacall-obj + --obj-arg).
//   Steps (buildReloadSteps): b-alive b-install info-apc info-vbe a-call locks-start;
//     rounds c1..c3: <c>-load (auto-load, ms) <c>-install <c>-alive <c>-info <c>-close (Close ms) <c>-gone <c>-recall (same path after
//       Close) <c>-lock (local: lock right after Close);
//     info-vbe-mid; path variants of A (pathVariants: lower upper dot dotdot slash short hl ext): <v>-load <v>-install <v>-alive <v>-info
//       <v>-close <v>-gone (a loaded variant is closed before the next one);
//     info-vbe-end b-close b-recall, then z (0 -> 1 -> 0 C5D_Bridge): z-load z-install z-info z-vbe z-close z-recall; locks-end.
//   Info steps never require anything; install/alive/close require their load. Only a timeout, an unreachable CATIA or a vanished
//   CNEXT.exe stops the run. At the end the probe copies and the hard link are deleted when not locked (else listed in notes).
//   Report: temp/probe-bridge-reload-<time>.json (summary.verdictLines: one line per round / variant / lock state; summary.reload as data).
//   NOTE: every <x>-close of that run failed with 0x80020005 DISP_E_TYPEMISMATCH before VBA ran: eval-probe.exe passed the --obj-arg
//   values as string[] (VT_ARRAY|VT_BSTR) instead of object[] (VT_ARRAY|VT_VARIANT). Fixed in EvalProbe.cs (VariantArray).
//   Also C5D_Install always wrote into B (the first C5D_Bridge), so "alive via own path" failed for every copy.
//
// Self-close experiment (with several projects named C5D_Bridge loaded: can a copy identify ITSELF, be closed from outside, and does
// the production idea "a new hard link of the master per Push, auto-load, work, close itself from outside" repeat cleanly?).
// Continues in the SAME CATIA the reload experiment left (B first, the other copies loaded); does not restart anything:
//        node scripts/probe-bridge.js --self-close-experiment
//   Modules: C5D_SxB (SXB_SOURCE, written into B by C5D_Install through B's path): C5D_SxPut(key, source) writes the standard module
//   C5D_Sx into the C5D_Bridge whose DisplayName (exact, then ignoring case) or VBProject.FileName (ignoring case) is key - exactly one
//   hit or an error text; only projects named C5D_Bridge, only the module C5D_Sx. C5D_SxBList / C5D_SxList: every IApcProject and
//   VBProject (name, DisplayName, FileName, module count, C5D_* modules, Saved). C5D_Sx (SX_SOURCE): C5D_SxSelf returns
//   apc.ExecutingProject (the IApcProject running the code; taken while running, from C# it is Empty afterwards), C5D_SxWho its text.
//   Steps (the copies t1..tN come from the start list, so phases 2-4 are built after phase 1):
//     1 sc-b-alive sc-b-install (C5D_SxB) sc-start-list sc-b-put (C5D_Sx into B) sc-b-who
//     2 per copy t<n> (every C5D_Bridge except B): <t>-put (through B) <t>-who (own path: ExecutingProject = itself?) <t>-close
//       (own path: C5D_SxSelf; eval-probe.exe guards Name = C5D_Bridge, DisplayName = the path, VBProject.FileName = the FileName of the
//       start list, and only then Close(false)) <t>-gone (list through B) <t>-recall (same path: expected 0x80004005)
//     3 b-list (only B left?) b-close (same guarded self-close, only when b-list shows B alone) b-recall
//     4 rounds p1 p2 p3 and the final check pz, each: <r>-link (local: new hard link C5D_Bridge-<8 hex>-p<time>-<n>.catvba of A, ms)
//       <r>-load (auto-load, ms) <r>-install (C5D_Install C5D_Sx: the only C5D_Bridge = itself, ms) <r>-who <r>-list (C5D_Bridge x1 =
//       this link, user projects + 1) <r>-close (guarded, ms) <r>-recall <r>-lock; a round starts only after the previous close.
//     locks-end (A, B, links, the copies' files). Then the hard links are deleted when not locked (A read-only again).
//   Info steps never require anything; a close requires its who (self-identified). Only a timeout, an unreachable CATIA or a
//   vanished CNEXT.exe stops the run. User projects are only read (start list vs last list compared).
//   Report: temp/probe-bridge-selfclose-<time>.json (summary.verdictLines: one fact per line; summary.selfClose / summary.lists as data).
//   RESULT of that run: t1-t4 (independent copies, lower-case path) closed, but the put into t5 (A as \.\) gave RPC_E_SERVERFAULT and
//   from then on every call into B gave 0x800ADF09 Unknown VBA error: never load one real file under several path strings at once.
//
// Cycle experiment (the production idea "one Push = one NEW path, loaded alone, self-identified, closed from outside", repeated, on a
// freshly RESTARTED CATIA with the VBA editor opened once):
//        node scripts/probe-bridge.js --cycle-experiment [--cycle-method hardlink|copy|both] [--cycle-rounds N]   (default both, N = 5)
//   Local first (no CATIA call): every C5D_Bridge*.catvba in %LOCALAPPDATA%\Cat5Dev\bridge must be free (lock probe), else the run
//   stops at once ("CATIA を再起動してください": a closed C5D_Bridge keeps its file until CATIA exits). Then the old files of earlier
//   probes (all but A = C5D_Bridge-<8 hex>.catvba and the old name C5D_Bridge.catvba) are deleted (read-only cleared; for a hard link of
//   A the read-only attribute of A is set again) = production's cleanup at the next start; A is prepared read-only (--readonly-fresh).
//   A is never loaded through its own path. Rounds (buildCycleRound), each on a NEW file next to A:
//     copy: C5D_Bridge-<8 hex>-c<yyyyMMddHHmmss>-<n>.catvba (read-only copy of A), hardlink: ...-h<yyyyMMddHHmmss>-<n>.catvba (fs.linkSync)
//     <r>-prepare (local, ms) <r>-load (C5D_Version, auto-load) <r>-install (C5D_Install C5D_Cy = CY_SOURCE: C5D_CySelf / C5D_CyWho /
//     C5D_CyList / C5D_CyRead) <r>-who (ExecutingProject = this path?) <r>-list (C5D_Bridge x1 = this path, user projects as at the
//     start) <r>-read (toQI.C5D_BridgeTest: CountOfLines and Lines(1, 1), READ ONLY) <r>-close (C5D_CySelf; guards Name / DisplayName,
//     then Close(false)) <r>-recall (same path again: 0x80004005 expected) <r>-lock (local).
//   Order: start round s (copy; its list = user projects at the start; another C5D_Bridge or no user project stops), the rounds of each
//   method (both: copy c1..cN first, then hardlink h1..hN), final round z (copy; the list at the end). No module snapshot between the
//   calls of a round (only at its lock step), so the calls follow each other as in a Push. Any failed round stops all later rounds
//   (a C5D_Bridge may be left loaded and the next C5D_Install would write into it); 0x800ADF09 / RPC_E_SERVERFAULT / an RPC failure /
//   a timeout / a vanished CNEXT.exe stop at once (fatal: CATIA may be unstable). The reason is in the verdict.
//   At the end the files of this run that are not locked are deleted (normally all are locked; the next run deletes them).
//   Report: temp/probe-bridge-cycle-<time>.json (summary.verdictLines: one fact per line; summary.cycle: per round facts).
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
    + '       node scripts/probe-bridge.js --reload-experiment\n'
    + '  --reload-experiment: right after --remove-experiment in the SAME CATIA (A closed, B loaded; do not restart): how to load the bridge\n'
    + '    again after a Close (path strings of A, a hard link, a new file name per round x3, then B closed and a last copy z: 0 -> 1 -> 0)\n'
    + '       node scripts/probe-bridge.js --self-close-experiment\n'
    + '  --self-close-experiment: right after --reload-experiment in the SAME CATIA (B first, the other copies loaded; do not restart):\n'
    + '    C5D_SxB into B, C5D_Sx into each copy by its path, self-identification by apc.ExecutingProject through the own path,\n'
    + '    guarded Close from outside (Name/DisplayName/FileName must match) of every copy, then B; then the production idea x3 (+1 final\n'
    + '    check): a new hard link of A per round, auto-load, C5D_Install, self-identify, close from outside. User projects are never written.\n'
    + '       node scripts/probe-bridge.js --cycle-experiment [--cycle-method hardlink|copy|both] [--cycle-rounds N]\n'
    + '  --cycle-experiment: right after a CATIA RESTART (VBA editor opened once): stops at once when a C5D_Bridge*.catvba of the bridge\n'
    + '    folder is in use; deletes the old ones, then per Push one NEW path (read-only copy of A or hard link of A; default both = copy\n'
    + '    first, N = 5 each) loaded alone: load install who list read (toQI.C5D_BridgeTest, read only) close recall lock; start round s\n'
    + '    and final round z (copies) list the user projects. 0x800ADF09 / RPC_E_SERVERFAULT / timeout / CATIA gone stop everything.\n'
    + '  Never runs a .catvbs/.CATScript. --check-only only reads CNEXT.exe modules and the bridge files.';

const MODE_OPTIONS = { '--check-only': 'check-only', '--unload-experiment': 'unload-experiment', '--unload-cleanup': 'unload-cleanup',
    '--readonly-fresh': 'readonly-fresh', '--readonly-fresh-check': 'readonly-fresh-check',
    '--remove-experiment': 'remove-experiment', '--remove-experiment-risky': 'remove-experiment-risky', '--reload-experiment': 'reload-experiment',
    '--self-close-experiment': 'self-close-experiment', '--cycle-experiment': 'cycle-experiment' };
/** Options that take a value (only with --cycle-experiment). */
const VALUE_OPTIONS = ['--cycle-method', '--cycle-rounds'];
/**
 * { mode: 'run' | 'check-only' | 'unload-experiment' | 'unload-cleanup' | 'readonly-fresh' | 'readonly-fresh-check' | 'remove-experiment'
 *   | 'remove-experiment-risky' | 'reload-experiment' | 'self-close-experiment' | 'cycle-experiment' | 'help', steps }. Throws on bad arguments.
 * (--self-close-experiment: steps = [], its step names depend on the copies found in CATIA and are recorded in the report.
 *  --cycle-experiment: steps = [] and cycle = { methods, rounds }.)
 */
function parseArgs(args) {
    if (args.some(arg => arg === '--help' || arg === '-h')) return { mode: 'help', steps: [] };
    const modes = new Set();
    const steps = new Set();
    const values = {};
    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        if (VALUE_OPTIONS.includes(arg)) {
            if (i + 1 >= args.length) throw new Error(`${arg} needs a value\n${USAGE}`);
            if (values[arg] !== undefined) throw new Error(`${arg} given twice\n${USAGE}`);
            values[arg] = args[++i];
        } else if (MODE_OPTIONS[arg]) modes.add(arg);
        else if (STEP_NAMES.includes(arg)) steps.add(arg);
        else throw new Error(`Unknown option or step: ${arg}\n${USAGE}`);
    }
    if (modes.size > 1) throw new Error(`${[...modes].join(' and ')} cannot be combined\n${USAGE}`);
    const [option] = modes;
    if (option && steps.size) throw new Error(`${option} takes no steps\n${USAGE}`);
    const given = Object.keys(values);
    if (given.length && option !== '--cycle-experiment') throw new Error(`${given.join(', ')} only with --cycle-experiment\n${USAGE}`);
    if (option === '--cycle-experiment') return { mode: 'cycle-experiment', steps: [], cycle: parseCycleOptions(values) };
    if (option === '--unload-experiment') return { mode: 'unload-experiment', steps: UNLOAD_STEPS.map(s => s.name) };
    if (option === '--unload-cleanup') return { mode: 'unload-cleanup', steps: CLEANUP_STEPS.map(s => s.name) };
    if (option === '--readonly-fresh') return { mode: 'readonly-fresh', steps: READONLY_FRESH_STEPS.map(s => s.name) };
    if (option === '--remove-experiment') return { mode: 'remove-experiment', steps: REMOVE_STEPS.map(s => s.name) };
    if (option === '--remove-experiment-risky') return { mode: 'remove-experiment-risky', steps: REMOVE_RISKY_STEPS.map(s => s.name) };
    if (option === '--reload-experiment') return { mode: 'reload-experiment', steps: reloadStepNames() };
    if (option) return { mode: MODE_OPTIONS[option], steps: [] };
    return { mode: 'run', steps: steps.size ? STEP_NAMES.filter(name => steps.has(name)) : [...STEP_NAMES] };
}

/** One command-line token for eval-probe.exe: b64:<UTF-8 base64> (survives any code page, CR/LF and quotes). */
const encodeToken = text => 'b64:' + Buffer.from(String(text), 'utf8').toString('base64');
/**
 * eval-probe.exe arguments of a step: [--no-retry] --label <name> vbacall <library> <module> <function> [args...].
 * noRetry: exactly one ExecuteScript call (no ByRef retry after an exception), for calls with side effects such as unloading.
 * A step with actions (objStep) uses [--obj-arg <arg> ...] vbacall-obj <library> <module> <function> [actions...] instead (always one
 * call; step.args become --obj-arg tokens, i.e. ByVal string arguments of the VBA function).
 */
function callArgs(step, library, options = {}) {
    if (step.actions) {
        return [...(step.args ?? []).flatMap(arg => ['--obj-arg', encodeToken(arg)]), '--label', step.name, 'vbacall-obj',
            ...[library, step.module, step.func, ...step.actions].map(encodeToken)];
    }
    return [...(options.noRetry ? ['--no-retry'] : []), '--label', step.name, 'vbacall', ...[library, step.module, step.func, ...step.args].map(encodeToken)];
}
/** Whether a step's VBA return value is the expected one. */
function matchStep(step, result) {
    if (typeof result !== 'string') return false;
    if (step.check) return step.check(result) === true;
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
 * Why a step is not run (null = run it): an earlier abort; disabled (a reason set when the steps were built, e.g. no 8.3 name); requires (a listed step did not match); unless (a listed step matched,
 * e.g. an earlier attempt already removed the project); after (a listed step was skipped or is not in this run).
 */
function skipReason(step, entries, aborted) {
    if (aborted) return aborted;
    if (step.disabled) return step.disabled;
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
 * prior: { entries, aborted } of an earlier runSequence of the same report to continue (report.runs is kept, requires / after may
 * name its steps; used when later steps are built from the results of earlier ones).
 * options.snapshot(step): false = no module snapshot / CATIA-alive check after this step (loadedAfter / alive stay null; keeps
 * back-to-back calls close together as in production). options.fatal(entry, step): a reason text aborts the rest like a timeout
 * (entry.fatal is set), e.g. 0x800ADF09 / RPC_E_SERVERFAULT in --cycle-experiment.
 */
async function runSequence(report, steps, library, pids, local, prior = null, options = {}) {
    if (!prior) report.runs = [];
    const entries = prior ? prior.entries : [];
    let aborted = prior?.aborted ?? null;
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
        if (options.snapshot && !options.snapshot(step)) {
            entry.loadedAfter = null;
            run.catiaAlive = entry.alive = null;
        } else {
            const after = trySnapshot(report, 'after:' + step.name);
            entry.loadedAfter = after.snapshot?.loaded ?? null;
            run.catiaAlive = entry.alive = catiaAlive(pids, after.snapshot);
        }
        entries.push(entry);
        const fatal = options.fatal ? options.fatal(entry, step) : null;
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
        } else if (fatal) {
            entry.fatal = run.fatal = fatal;
            report.notes.push(`Step ${step.name}: ${fatal}; CATIA may be unstable. Remaining steps were not run.`);
            aborted = `aborted: ${fatal} in ${step.name}`;
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

// ---- --reload-experiment: after C5D_Bridge was closed from outside (IApcProject.Close(False) from C#, found by --remove-experiment),
// how can Push load the bridge again? ExecuteScript with the same path string then fails with 0x80004005 "macro program unknown"
// (CATIA keeps the path as a known library). Tried here: other strings for the same file (case, 8.3, /, \\?\, .\, ..\, hard link)
// and a new file name per Push (the production idea), each loaded, marked, called, closed from outside and checked.
// Every project-specific module is C5D_Rl_<tag> with public functions C5D_Rl<Kind>_<tag> (unique names, so two of them in one project
// never make an ambiguous name). Info and Close go through copy B (loaded by the remove experiment) until B itself is closed.
const RL_PREFIX = 'C5D_Rl_';
const RL_TAG = /^[a-z][a-z0-9]{0,7}$/;
const rlModule = tag => `${RL_PREFIX}${tag}`;
const rlFunc = (kind, tag) => `C5D_Rl${kind}_${tag}`;
const rlAliveValue = tag => `rl-alive:${tag}`;
const RL_APC_PATTERN = /^apc=[^;|]*;count=\d+;err=0(\||$)/;
const RL_VBE_PATTERN = /^vbe=[^;|]*;count=\d+;err=0(\||$)/;
/** Actions of a close step: index 2 is the Close (timed by eval-probe.exe), DisplayName shows which project was found. */
const RL_CLOSE_ACTIONS = ['get:$.Name', 'get:$.DisplayName', 'call:$.Close(false)', 'get:$.Name'];
const RL_CLOSE_ACTION = 2;
/** Repeat rounds (a new file name each time) and the final 0 -> 1 -> 0 copy. */
const RL_ROUNDS = ['c1', 'c2', 'c3'];
const RL_FINAL = 'z';
/** VBA: v = <expr>, then ";<label>=<value>" (or "=!<err>") appended to s. */
const rlProp = (label, expr) => ['    Err.Clear', '    v = "?"', `    v = ${expr}`, '    e = Err.Number', `    s = s & ";${label}=" & RlFmt(v, e)`];

/**
 * VBA written into C5D_Rl_<tag> (ASCII only, CRLF, On Error Resume Next, no MsgBox, no Declare, never closes or removes anything:
 * Close is done by eval-probe.exe on the object C5D_RlFind_<tag> returns, after ExecuteScript has returned).
 *   C5D_RlAlive_<tag>()     "rl-alive:<tag>"
 *   C5D_RlInfo_<tag>()      "apc=<TypeName>;count=N;err=E|name=..;display=..;mode=..;dirty=..;vbType=..;vbFile=..;mods=<C5D_* modules>;self=..|..."
 *   C5D_RlVbe_<tag>()       "vbe=<TypeName>;count=N;err=E|name=..;type=..;mode=..;file=..;saved=..;mods=..;self=..|..."
 *   C5D_RlFind_<tag>(key)   the IApcProject named C5D_Bridge whose DisplayName or VBProject.FileName is key: pass 1 exact, pass 2 ignoring
 *                           case, pass 3 same file name (ignoring case). Exactly one hit per pass, else "error:find:ambiguous:..|notfound:.."
 * self = the project has the module C5D_Rl_<tag>.
 */
function reloadSource(tag) {
    if (!RL_TAG.test(tag)) throw new Error(`bad reload tag: ${tag}`);
    const fn = kind => rlFunc(kind, tag);
    const listHeader = (name, guardObj, guardText) => [
        '    If ' + guardObj + ' Is Nothing Then',
        `        ${name} = "${guardText}" & CStr(Err.Number)`,
        '        Exit Function',
        '    End If'
    ];
    return [
        'Option Explicit',
        `Private Const RL_MARKER As String = "${rlModule(tag)}"`,
        `Private Const RL_PROJECT As String = "${BRIDGE_PROJECT}"`,
        `Public Function ${fn('Alive')}() As Variant`,
        `    ${fn('Alive')} = "${rlAliveValue(tag)}"`,
        'End Function',
        'Private Function RlClean(ByVal t As String) As String',
        '    t = Replace(t, "|", "/")',
        '    t = Replace(t, ";", ",")',
        '    t = Replace(t, vbCr, " ")',
        '    t = Replace(t, vbLf, " ")',
        '    RlClean = t',
        'End Function',
        'Private Function RlFmt(ByVal v As Variant, ByVal e As Long) As String',
        '    On Error Resume Next',
        '    Dim t As String',
        '    If e <> 0 Then',
        '        RlFmt = "!" & CStr(e)',
        '        Exit Function',
        '    End If',
        '    t = "?"',
        '    t = CStr(v)',
        '    RlFmt = RlClean(t)',
        'End Function',
        'Private Function RlApc() As Object',
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Set apc = Nothing',
        '    Set apc = CreateObject("MSAPC.Apc.7.1")',
        '    If apc Is Nothing Then',
        '        Err.Clear',
        '        Set apc = CreateObject("MSAPC.Apc")',
        '    End If',
        '    Set RlApc = apc',
        'End Function',
        "' file name part of a path (after the last \\ or /)",
        'Private Function RlBase(ByVal p As String) As String',
        '    Dim i As Long',
        '    Dim c As String',
        '    For i = Len(p) To 1 Step -1',
        '        c = Mid(p, i, 1)',
        '        If c = "\\" Or c = "/" Then Exit For',
        '    Next',
        '    RlBase = Mid(p, i + 1)',
        'End Function',
        "' names of the components whose name starts with C5D_, comma separated",
        'Private Function RlMods(ByVal vbp As Object) As String',
        '    On Error Resume Next',
        '    Dim comps As Object',
        '    Dim c As Object',
        '    Dim nm As String',
        '    Dim s As String',
        '    Dim i As Long',
        '    Dim n As Long',
        '    Set comps = Nothing',
        '    Set comps = vbp.VBComponents',
        '    n = 0',
        '    n = comps.Count',
        '    For i = 1 To n',
        '        Set c = Nothing',
        '        Set c = comps.Item(i)',
        '        nm = ""',
        '        nm = CStr(c.Name)',
        '        If Left(nm, 4) = "C5D_" Then s = s & "," & nm',
        '    Next',
        '    Err.Clear',
        '    RlMods = Mid(s, 2)',
        'End Function',
        'Private Function RlTail(ByVal vbp As Object) As String',
        '    On Error Resume Next',
        '    Dim mods As String',
        '    mods = ""',
        '    mods = RlMods(vbp)',
        '    RlTail = ";mods=" & RlClean(mods) & ";self=" & CStr(InStr(1, "," & mods & ",", "," & RL_MARKER & ",", vbBinaryCompare) > 0)',
        '    Err.Clear',
        'End Function',
        'Private Function RlProj(ByVal ap As Object) As String',
        '    On Error Resume Next',
        '    Dim s As String',
        '    Dim v As Variant',
        '    Dim e As Long',
        '    Dim vbp As Object',
        '    s = ""',
        ...rlProp('name', 'ap.Name'),
        ...rlProp('display', 'ap.DisplayName'),
        ...rlProp('mode', 'ap.Mode'),
        ...rlProp('dirty', 'ap.Dirty'),
        '    Set vbp = Nothing',
        '    Set vbp = ap.VBProject',
        ...rlProp('vbType', 'vbp.Type'),
        ...rlProp('vbFile', 'vbp.FileName'),
        '    s = s & RlTail(vbp)',
        '    Err.Clear',
        '    RlProj = Mid(s, 2)',
        'End Function',
        'Private Function RlVbp(ByVal vbp As Object) As String',
        '    On Error Resume Next',
        '    Dim s As String',
        '    Dim v As Variant',
        '    Dim e As Long',
        '    s = ""',
        ...rlProp('name', 'vbp.Name'),
        ...rlProp('type', 'vbp.Type'),
        ...rlProp('mode', 'vbp.Mode'),
        ...rlProp('file', 'vbp.FileName'),
        ...rlProp('saved', 'vbp.Saved'),
        '    s = s & RlTail(vbp)',
        '    Err.Clear',
        '    RlVbp = Mid(s, 2)',
        'End Function',
        `Public Function ${fn('Info')}() As Variant`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Dim aps As Object',
        '    Dim ap As Object',
        '    Dim s As String',
        '    Dim i As Long',
        '    Dim n As Long',
        '    Set apc = RlApc()',
        ...listHeader(fn('Info'), 'apc', 'apc=Nothing;err='),
        '    s = "apc=" & TypeName(apc)',
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
        '            s = s & "|" & RlProj(ap)',
        '        End If',
        '    Next',
        `    ${fn('Info')} = s`,
        'End Function',
        `Public Function ${fn('Vbe')}() As Variant`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Dim ide As Object',
        '    Dim vbps As Object',
        '    Dim vbp As Object',
        '    Dim s As String',
        '    Dim i As Long',
        '    Dim n As Long',
        '    Set apc = RlApc()',
        ...listHeader(fn('Vbe'), 'apc', 'vbe=Nothing;apcErr='),
        '    Set ide = Nothing',
        '    Set ide = apc.VBE',
        ...listHeader(fn('Vbe'), 'ide', 'vbe=Nothing;err='),
        '    s = "vbe=" & TypeName(ide)',
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
        '            s = s & "|" & RlVbp(vbp)',
        '        End If',
        '    Next',
        `    ${fn('Vbe')} = s`,
        'End Function',
        'Private Function RlFind(ByVal apc As Object, ByVal key As String, ByRef why As String) As Object',
        '    On Error Resume Next',
        '    Dim aps As Object',
        '    Dim ap As Object',
        '    Dim hit As Object',
        '    Dim nm As String',
        '    Dim dn As String',
        '    Dim fl As String',
        '    Dim kb As String',
        '    Dim ok As Boolean',
        '    Dim pass As Long',
        '    Dim hits As Long',
        '    Dim i As Long',
        '    Dim n As Long',
        '    Set RlFind = Nothing',
        '    Set aps = Nothing',
        '    Set aps = apc.Projects',
        '    n = -1',
        '    n = aps.Count',
        '    kb = RlBase(key)',
        '    For pass = 1 To 3',
        '        hits = 0',
        '        Set hit = Nothing',
        '        For i = 1 To n',
        '            Set ap = Nothing',
        '            Set ap = aps.Item(i)',
        '            nm = ""',
        '            nm = CStr(ap.Name)',
        '            dn = ""',
        '            dn = CStr(ap.DisplayName)',
        '            fl = ""',
        '            fl = CStr(ap.VBProject.FileName)',
        '            ok = False',
        '            If nm = RL_PROJECT Then',
        '                If pass = 1 Then',
        '                    ok = (dn = key) Or (fl = key)',
        '                ElseIf pass = 2 Then',
        '                    ok = (StrComp(dn, key, vbTextCompare) = 0) Or (StrComp(fl, key, vbTextCompare) = 0)',
        '                Else',
        '                    ok = (StrComp(RlBase(dn), kb, vbTextCompare) = 0) Or (StrComp(RlBase(fl), kb, vbTextCompare) = 0)',
        '                End If',
        '            End If',
        '            If ok Then',
        '                hits = hits + 1',
        '                Set hit = ap',
        '            End If',
        '        Next',
        '        If hits = 1 Then',
        '            why = "pass" & CStr(pass)',
        '            Set RlFind = hit',
        '            Exit Function',
        '        End If',
        '        If hits > 1 Then',
        '            why = "ambiguous:pass" & CStr(pass) & ":" & CStr(hits)',
        '            Exit Function',
        '        End If',
        '    Next',
        '    why = "notfound:count=" & CStr(n)',
        'End Function',
        `Public Function ${fn('Find')}(ByVal key As Variant) As Variant`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Dim ap As Object',
        '    Dim why As String',
        '    Set apc = RlApc()',
        ...listHeader(fn('Find'), 'apc', 'error:apc:'),
        '    why = ""',
        '    Set ap = RlFind(apc, CStr(key), why)',
        '    If ap Is Nothing Then',
        `        ${fn('Find')} = "error:find:" & why`,
        '        Exit Function',
        '    End If',
        `    Set ${fn('Find')} = ap`,
        'End Function',
        ''
    ].join('\r\n');
}

const rlVersion = (name, library, extra = {}) => ({ ...versionStep(name), library, ...extra });
const rlInstall = (name, tag, library, extra = {}) => ({ name, library, module: BRIDGE_MODULE, func: 'C5D_Install', args: [rlModule(tag), reloadSource(tag)], expected: 'ok', ...extra });
const rlAlive = (name, tag, library, extra = {}) => ({ name, library, module: rlModule(tag), func: rlFunc('Alive', tag), args: [], expected: rlAliveValue(tag), ...extra });
/** Info through the project that has C5D_Rl_<via> (kind apc = apc.Projects, vbe = VBE.VBProjects). Never requires anything. */
const rlInfo = (name, via, library, kind = 'apc', extra = {}) => ({ name, library, module: rlModule(via), func: rlFunc(kind === 'apc' ? 'Info' : 'Vbe', via), args: [],
    pattern: kind === 'apc' ? RL_APC_PATTERN : RL_VBE_PATTERN, info: kind, ...extra });
/** C5D_RlFind_<via>(target) returns the project of target; eval-probe.exe closes it after ExecuteScript returned. */
const rlClose = (name, via, library, target, extra = {}) => ({ name, library, module: rlModule(via), func: rlFunc('Find', via), args: [target],
    actions: RL_CLOSE_ACTIONS, mustSucceed: [RL_CLOSE_ACTION], resultType: 'ComObject', closeOf: target, ...extra });
/** C5D_Version through a closed library: matches unless it answers (expected: 0x80004005 "macro program unknown"). */
const rlRecall = (name, library, extra = {}) => ({ name, library, module: BRIDGE_MODULE, func: 'C5D_Version', args: [], expectFailure: BRIDGE_VERSION, ...extra });
const rlLocks = (name, files) => ({ name, local: 'locks', files });

/**
 * Steps of --reload-experiment. p: { a, b, z, copies: [{ tag, path }], variants: [{ tag, library, disabled }] }.
 * Info steps never require anything (each step may fail on its own; only a timeout / crash / unreachable CATIA stops the run);
 * calls that only make sense after a load require that load. disabled: reason set at build time (no 8.3 name, no hard link).
 * Order: state through B, the repeat rounds (the production idea, before the exotic path strings), the path variants of A,
 * then B is closed and a last copy z shows 0 -> 1 -> 0 C5D_Bridge.
 */
function buildReloadSteps(p) {
    const steps = [
        { name: 'b-alive', library: p.b, module: B_MODULE, func: 'C5D_BAlive', args: [], expected: B_ALIVE },
        rlInstall('b-install', 'b', p.b),
        rlInfo('info-apc', 'b', p.b),
        rlInfo('info-vbe', 'b', p.b, 'vbe'),
        rlRecall('a-call', p.a),
        rlLocks('locks-start', ['a', 'b'])
    ];
    for (const { tag, path: file } of p.copies) {
        const loaded = { requires: [`${tag}-load`] };
        steps.push(
            rlVersion(`${tag}-load`, file, { round: tag }),
            rlInstall(`${tag}-install`, tag, file, loaded),
            rlAlive(`${tag}-alive`, tag, file, loaded),
            rlInfo(`${tag}-info`, 'b', p.b),
            rlClose(`${tag}-close`, 'b', p.b, file, loaded),
            rlInfo(`${tag}-gone`, 'b', p.b),
            rlRecall(`${tag}-recall`, file),
            rlLocks(`${tag}-lock`, [tag]));
    }
    steps.push(rlInfo('info-vbe-mid', 'b', p.b, 'vbe'));
    for (const v of p.variants) {
        const loaded = { requires: [`${v.tag}-load`] };
        steps.push(
            rlVersion(`${v.tag}-load`, v.library, { variant: v.tag, ...(v.disabled ? { disabled: v.disabled } : {}) }),
            rlInstall(`${v.tag}-install`, v.tag, v.library, loaded),
            rlAlive(`${v.tag}-alive`, v.tag, v.library, loaded),
            rlInfo(`${v.tag}-info`, 'b', p.b, 'apc', { after: [`${v.tag}-load`] }),
            rlClose(`${v.tag}-close`, 'b', p.b, v.library, loaded),
            rlInfo(`${v.tag}-gone`, 'b', p.b, 'apc', { after: [`${v.tag}-close`] }));
    }
    const z = p.z;
    steps.push(
        rlInfo('info-vbe-end', 'b', p.b, 'vbe'),
        rlClose('b-close', 'b', p.b, p.b),
        rlRecall('b-recall', p.b),
        rlVersion('z-load', z),
        rlInstall('z-install', RL_FINAL, z, { requires: ['z-load'] }),
        rlInfo('z-info', RL_FINAL, z),
        rlInfo('z-vbe', RL_FINAL, z, 'vbe'),
        rlClose('z-close', RL_FINAL, z, z, { requires: ['z-load'] }),
        rlRecall('z-recall', z),
        rlLocks('locks-end', ['a', 'b', ...p.copies.map(c => c.tag), RL_FINAL, 'hl']));
    return steps;
}
/** Step names of --reload-experiment (independent of the paths). */
function reloadStepNames() {
    const x = 'C:\\x\\C5D_Bridge-x.catvba';
    return buildReloadSteps({ a: x, b: x, z: x, copies: RL_ROUNDS.map(tag => ({ tag, path: x })), variants: pathVariants(x, { short: { path: x }, hardlink: { ok: true, path: x } }) })
        .map(s => s.name);
}

/**
 * Other strings for the same file A, in run order (\\?\ last: the most likely to be refused). short: { path, error } of the 8.3
 * name; hardlink: { ok, path, error }. A variant is disabled (with the reason) when it cannot be made or equals A / an earlier one.
 */
function pathVariants(a, { short, hardlink } = {}) {
    const dir = path.win32.dirname(a), base = path.win32.basename(a);
    const list = [
        { tag: 'lower', label: '(i) lower case', library: a.toLowerCase() },
        { tag: 'upper', label: '(ii) upper case', library: a.toUpperCase() },
        { tag: 'dot', label: '(vi) .\\ before the file name', library: `${dir}\\.\\${base}` },
        { tag: 'dotdot', label: '..\\<folder> before the file name', library: `${dir}\\..\\${path.win32.basename(dir)}\\${base}` },
        { tag: 'slash', label: '(iv) / separators', library: a.replace(/\\/g, '/') },
        { tag: 'short', label: '(iii) 8.3 short path', library: short?.path ?? null,
            disabled: !short ? 'no short path computed' : short.error ? `no short path (${short.error})`
                : !String(short.path ?? '').includes('~') ? 'no 8.3 short name (8.3 name generation off on this volume?)' : undefined },
        { tag: 'hl', label: '(vii) hard link', library: hardlink?.path ?? null,
            disabled: !hardlink ? 'no hard link made' : !hardlink.ok ? `hard link not created (${hardlink.error})` : undefined },
        { tag: 'ext', label: '(v) \\\\?\\ prefix', library: `\\\\?\\${a}` }
    ];
    const seen = new Map([[a, 'A itself']]);
    for (const v of list) {
        if (!v.disabled && seen.has(v.library)) v.disabled = `same string as ${seen.get(v.library)}`;
        if (!v.disabled) seen.set(v.library, v.tag);
        if (!v.disabled) delete v.disabled;
    }
    return list;
}

/** 8.3 short path of an existing file through Scripting.FileSystemObject in PowerShell (never in CATIA): { path, error }. */
function shortPathOf(file, run = execFileSync) {
    const script = [
        "$ErrorActionPreference = 'Stop'",
        '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false',
        '(New-Object -ComObject Scripting.FileSystemObject).GetFile($env:C5D_SHORT_PATH_OF).ShortPath'
    ].join('\n');
    try {
        const out = run(powershellPath(), ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')],
            { windowsHide: true, timeout: 60000, stdio: ['ignore', 'pipe', 'pipe'], env: { ...process.env, C5D_SHORT_PATH_OF: file } });
        const text = out.toString('utf8').replace(/^\uFEFF/, '').trim();
        return text ? { path: text, error: null } : { path: null, error: 'empty output' };
    } catch (error) {
        return { path: null, error: String(error.stderr ? error.stderr.toString().trim() || error : error).split(/\r?\n/)[0] };
    }
}
/** Hard link of a at link (same volume; the read-only attribute is shared): { ok, path, error, sameFile, links }. */
function makeHardlink(a, link, ops = {}) {
    const { linkSync = fs.linkSync, statSync = fs.statSync } = ops;
    const result = { ok: false, path: link, error: null, sameFile: null, links: null };
    try {
        linkSync(a, link);
        const sa = statSync(a, { bigint: true }), sl = statSync(link, { bigint: true });
        result.sameFile = sa.ino === sl.ino && sa.dev === sl.dev;
        result.links = Number(sl.nlink);
        result.ok = result.sameFile !== false;
        if (!result.ok) result.error = 'the link is not the same file';
    } catch (error) { result.error = errorText(error); }
    return result;
}
/**
 * Deletes the probe files that are not locked (read-only cleared first). items: [{ key, path, linkOf }] (linkOf: the file a hard link
 * shares its attributes with; its read-only attribute is set again afterwards). Returns [{ key, path, action, error, lock, linkOfReadOnly }].
 */
function cleanupProbeFiles(items, ops = {}) {
    const { chmod = fs.chmodSync, unlink = fs.unlinkSync, open, close, platform } = ops;
    return items.map(({ key, path: file, linkOf }) => {
        const out = { key, path: file, action: null, error: null, lock: null };
        if (!fs.existsSync(file)) { out.action = 'missing'; return out; }
        out.lock = lockProbe(file, open, close, platform);
        if (out.lock.inUse !== false) { out.action = out.lock.inUse ? 'kept (locked)' : `kept (lock unknown: ${out.lock.code ?? out.lock.skipped})`; return out; }
        try {
            chmod(file, 0o666);
            unlink(file);
            out.action = 'deleted';
        } catch (error) {
            out.action = 'failed';
            out.error = errorText(error);
            try { if (fs.existsSync(file)) chmod(file, 0o444); } catch { /* reported by the file state */ }
        }
        if (linkOf) {
            try { chmod(linkOf, 0o444); } catch (error) { out.error = [out.error, `read-only of ${linkOf} not restored: ${errorText(error)}`].filter(Boolean).join('; '); }
            out.linkOfReadOnly = fileState(linkOf).readOnly;
        }
        return out;
    });
}

// -- judging --
const baseName = p => String(p ?? '').split(/[\\/]/).pop();
const sameText = (x, y) => String(x).toLowerCase() === String(y).toLowerCase();
/** C5D_Bridge projects of a parsed C5D_RlInfo / C5D_RlVbe text: [{ display, file, mods, self }]. */
function rlBridges(info) {
    return (info?.projects ?? []).filter(p => p.name === BRIDGE_PROJECT)
        .map(p => ({ display: p.display ?? '', file: p.vbFile ?? p.file ?? '', mods: String(p.mods ?? '').split(',').filter(Boolean), self: p.self === 'True' }));
}
const rlPathOf = bridge => bridge.display || bridge.file;
/** The C5D_Bridge projects whose path has the file name of file (ignoring case). */
const bridgesOfFile = (info, file) => rlBridges(info).filter(b => sameText(baseName(rlPathOf(b)), baseName(file)) || sameText(baseName(b.file), baseName(file)));
/** ExecuteScript failure text without the "ExecuteScript(...)" line eval-probe.exe adds. */
const shortMessage = message => String(message ?? '').split(/\r?\n/).map(l => l.trim()).filter(l => l && !l.startsWith('ExecuteScript(')).join(' / ');
function rlOutcome(e) {
    if (!e) return 'not in this run';
    if (e.skipped) return `not run (${e.skipped})`;
    if (e.timedOut) return 'timed out (CATIA may be stopped at a dialog)';
    if (!e.ok) return `failed (${[e.hresult ?? 'no result', shortMessage(e.message)].filter(Boolean).join(' ')})`;
    return `returned ${JSON.stringify(e.result ?? null)}`;
}
/** "N projects, C5D_Bridge xM [<file> {<probe modules>}; ...]" of an info step. */
function describeRlInfo(entry) {
    if (!entry || entry.skipped || !entry.ok) return rlOutcome(entry);
    const info = parseInfo(entry.result);
    if (!info || info.count === null) return `unexpected ${JSON.stringify(entry.result)}`;
    const bridges = rlBridges(info).map(b => {
        const mods = b.mods.filter(m => m !== BRIDGE_MODULE);
        return `${baseName(rlPathOf(b)) || '?'}${mods.length ? ` {${mods.join(',')}}` : ''}`;
    });
    return `${info.count} projects, ${BRIDGE_PROJECT} x${bridges.length}${bridges.length ? ` [${bridges.join('; ')}]` : ''}`;
}
const parsedOk = e => (e && !e.skipped && e.ok ? parseInfo(e.result) : null);
const bridgeCount = e => { const info = parsedOk(e); return info && info.count !== null ? rlBridges(info).length : null; };
/** The last apc info entry (ok) before the step called name. */
function apcInfoBefore(name, entries, steps) {
    const index = entries.findIndex(e => e.step === name);
    for (let i = index - 1; i >= 0; i--) {
        const step = steps.find(s => s.name === entries[i].step);
        if (step?.info === 'apc' && parsedOk(entries[i])?.count != null) return entries[i];
    }
    return null;
}
/** { ok, ms, found, error } of a close step (ms = the Close call itself, timed by eval-probe.exe). */
function closeFacts(e) {
    if (!e || e.skipped) return { ok: null, ms: null, found: null, error: e ? e.skipped : 'not in this run' };
    const action = (e.actions ?? []).find(a => a.action === RL_CLOSE_ACTION);
    const found = (e.actions ?? []).find(a => a.op === 'get:$.DisplayName');
    if (!action) return { ok: false, ms: null, found: null, error: e.resultType && e.resultType !== 'ComObject' ? `no project (${JSON.stringify(e.result)})` : rlOutcome(e) };
    return { ok: action.ok === true, ms: action.elapsedMs ?? null, found: found?.ok ? found.result : null,
        error: action.ok ? null : `${action.hresult} ${String(action.message ?? '').split(/\r?\n/)[0]}`.trim() };
}
const lockWord = l => !l ? 'not probed' : l.skipped ? l.skipped : l.inUse === true ? 'LOCKED' : l.inUse === false ? 'free' : `probe failed (${l.code})`;
const recallWord = e => !e || e.skipped ? rlOutcome(e) : e.timedOut ? rlOutcome(e) : e.match
    ? `fails as expected (${[e.hresult, shortMessage(e.message)].filter(Boolean).join(' ')})` : `ANSWERS ${JSON.stringify(e.result)} (loaded again)`;

/** Facts of one repeat round c<n> (or z). */
function judgeRound(tag, file, get, entries, steps) {
    const load = get(`${tag}-load`), info = get(`${tag}-info`), gone = get(`${tag}-gone`);
    const infoParsed = parsedOk(info), goneParsed = parsedOk(gone);
    const holders = rlBridges(infoParsed).filter(b => b.mods.includes(rlModule(tag)));
    return {
        round: tag, file,
        load: { ok: load?.skipped ? null : load?.match ?? null, ms: load?.elapsedMs ?? null, outcome: rlOutcome(load) },
        install: get(`${tag}-install`)?.skipped ? null : get(`${tag}-install`)?.match ?? null,
        alive: get(`${tag}-alive`)?.skipped ? null : get(`${tag}-alive`)?.match ?? null,
        loaded: infoParsed ? bridgesOfFile(infoParsed, file).length > 0 : null,
        moduleIn: infoParsed ? holders.map(b => baseName(rlPathOf(b))) : null,
        counts: [bridgeCount(apcInfoBefore(`${tag}-load`, entries, steps)), bridgeCount(info), bridgeCount(gone)],
        close: closeFacts(get(`${tag}-close`)),
        gone: goneParsed && goneParsed.count !== null ? bridgesOfFile(goneParsed, file).length === 0 : null,
        recall: get(`${tag}-recall`) ? recallWord(get(`${tag}-recall`)) : null,
        lockAfterClose: get(`${tag}-lock`)?.locks?.[tag] ?? null
    };
}
function roundLine(r) {
    const counts = r.counts.map(n => n === null ? '?' : n).join(' -> ');
    const close = r.close.ok === null ? `close ${r.close.error}` : r.close.ok ? `close ok ${r.close.ms}ms${r.close.found ? ` (found ${baseName(r.close.found)})` : ''}`
        : `close FAILED ${r.close.error}`;
    return `round ${r.round} ${r.file}: load ${r.load.ok ? `ok ${r.load.ms}ms` : r.load.outcome}, install ${yn(r.install)}, alive via own path ${yn(r.alive)}, `
        + `module went into ${r.moduleIn === null ? '?' : r.moduleIn.length ? r.moduleIn.join(' + ') : 'no C5D_Bridge'}, C5D_Bridge ${counts}, ${close}, `
        + `gone ${yn(r.gone)}, same path after close ${r.recall ?? '-'}, lock right after close ${lockWord(r.lockAfterClose)}`;
}
/** Facts of one path variant of A. */
function judgeVariant(v, get, entries, steps) {
    const load = get(`${v.tag}-load`), info = get(`${v.tag}-info`);
    const before = parsedOk(apcInfoBefore(`${v.tag}-load`, entries, steps));
    const now = parsedOk(info);
    const known = new Set(rlBridges(before).map(rlPathOf));
    const added = now && before ? rlBridges(now).filter(b => !known.has(rlPathOf(b))).map(rlPathOf) : null;
    const holders = rlBridges(now).filter(b => b.mods.includes(rlModule(v.tag))).map(rlPathOf);
    const gone = parsedOk(get(`${v.tag}-gone`));
    // gone: the project(s) that appeared with this variant are no longer listed (the display may differ from the string passed,
    // e.g. an 8.3 name shown long); without such a project, by the file name of the variant
    const left = !gone || gone.count === null ? null : added?.length ? rlBridges(gone).filter(b => added.includes(rlPathOf(b)))
        : bridgesOfFile(gone, v.library);
    return {
        variant: v.tag, label: v.label, library: v.library, disabled: v.disabled ?? null,
        load: { ok: load?.skipped ? null : load?.match ?? null, ms: load?.elapsedMs ?? null, outcome: rlOutcome(load) },
        newProjects: added, moduleIn: now ? holders : null,
        alive: get(`${v.tag}-alive`)?.skipped ? null : get(`${v.tag}-alive`)?.match ?? null,
        close: closeFacts(get(`${v.tag}-close`)),
        gone: left === null ? null : left.length === 0
    };
}
function variantLine(v) {
    const head = `variant ${v.variant} ${v.label} ${JSON.stringify(v.library)}: `;
    if (v.disabled) return head + `skipped (${v.disabled})`;
    if (!v.load.ok) {
        const extra = v.newProjects?.length ? ` but a new C5D_Bridge appeared: ${v.newProjects.join(', ')}` : '';
        return head + `NOT loaded: ${v.load.outcome}${extra}`;
    }
    const where = v.newProjects === null ? 'no info' : v.newProjects.length ? `new project ${v.newProjects.map(p => JSON.stringify(p)).join(', ')}`
        : 'NO new project (answered by an already loaded one?)';
    const close = v.close.ok === null ? `close ${v.close.error}` : v.close.ok ? `close ok ${v.close.ms}ms` : `close FAILED ${v.close.error}`;
    return head + `LOADED ${v.load.ms}ms (${where}), module went into ${v.moduleIn?.length ? v.moduleIn.join(' + ') : 'no C5D_Bridge'}, alive ${yn(v.alive)}, ${close}, gone ${yn(v.gone)}`;
}

/**
 * Verdict of --reload-experiment: { lines, reload } (lines: one fact per line; reload: the same as data).
 * ctx: { files: { key: path }, copies: [{ tag, path }], variants, cleanup }.
 */
function judgeReload(beforeLoaded, entries, steps, ctx) {
    const get = name => entries.find(e => e.step === name);
    const lines = [];
    const bAlive = get('b-alive');
    lines.push(`start: B ${bAlive?.match ? 'is the B of the remove run (C5D_ProbeB answers)' : `C5D_ProbeB ${rlOutcome(bAlive)} (B freshly loaded? CATIA restarted?)`}, `
        + `b-install ${get('b-install')?.match ? 'ok' : rlOutcome(get('b-install'))}`);
    lines.push(`info-apc: ${describeRlInfo(get('info-apc'))}`);
    lines.push(`info-vbe: ${describeRlInfo(get('info-vbe'))}`);
    lines.push(`A path ${baseName(ctx.files.a)}: ${recallWord(get('a-call'))}`);
    const locks = name => get(name)?.locks ?? {};
    lines.push(`locks at start: ${Object.entries(locks('locks-start')).map(([k, l]) => `${k} ${lockWord(l)}`).join(', ') || 'not probed'}`);
    const rounds = ctx.copies.map(c => judgeRound(c.tag, baseName(c.path), get, entries, steps));
    for (const r of rounds) lines.push(roundLine(r));
    lines.push(`info-vbe-mid: ${describeRlInfo(get('info-vbe-mid'))}`);
    const variants = ctx.variants.map(v => judgeVariant(v, get, entries, steps));
    for (const v of variants) lines.push(variantLine(v));
    lines.push(`info-vbe-end: ${describeRlInfo(get('info-vbe-end'))}`);
    const bClose = closeFacts(get('b-close'));
    lines.push(`b-close: ${bClose.ok === null ? bClose.error : bClose.ok ? `ok ${bClose.ms}ms` : `FAILED ${bClose.error}`}; B path after close ${recallWord(get('b-recall'))}`);
    const z = judgeRound(RL_FINAL, baseName(ctx.files[RL_FINAL]), get, entries, steps);
    const zInfo = parsedOk(get('z-info'));
    const bStill = zInfo ? bridgesOfFile(zInfo, ctx.files.b).length > 0 : null;
    lines.push(`final ${z.file} (0 -> 1 -> 0): load ${z.load.ok ? `ok ${z.load.ms}ms` : z.load.outcome}, install ${yn(z.install)}, z-info ${describeRlInfo(get('z-info'))}`
        + ` (B still listed: ${yn(bStill)}), close ${z.close.ok === null ? z.close.error : z.close.ok ? `ok ${z.close.ms}ms` : `FAILED ${z.close.error}`}, `
        + `same path after close ${z.recall ?? '-'}`);
    lines.push(`z-vbe: ${describeRlInfo(get('z-vbe'))}`);
    lines.push(`locks at end: ${Object.entries(locks('locks-end')).map(([k, l]) => `${k} ${lockWord(l)}`).join(', ') || 'not probed'}`);
    if (ctx.cleanup) lines.push(`cleanup: ${ctx.cleanup.map(c => `${c.key} ${c.action}${c.error ? ` (${c.error})` : ''}`).join(', ') || 'nothing'}`);
    const crashed = entries.find(e => e.crashed);
    if (crashed) lines.push(`CATIA crashed after step ${crashed.step}`);
    const timedOut = entries.find(e => e.timedOut);
    if (timedOut) lines.push(`timed out at step ${timedOut.step} (CATIA may be stopped at a dialog)`);
    lines.push(describeVbscript(beforeLoaded, entries));
    return { lines, reload: { rounds, variants, bClose: { ...bClose, recall: recallWord(get('b-recall')) }, final: { ...z, bStillListed: bStill },
        locks: { start: locks('locks-start'), end: locks('locks-end') }, cleanup: ctx.cleanup ?? null } };
}

/**
 * --reload-experiment (CATIA as the remove experiment left it: A closed, B loaded; NOT restarted). Local preparation before any CATIA
 * call: facts and locks of A and B (never modified), read-only copies c1..c3 and z with per-run names, a hard link of A, the 8.3 name
 * of A. Then buildReloadSteps one --no-retry call at a time; at the end the probe copies / hard link are deleted when not locked.
 */
async function runReload(report) {
    const paths = hashedPaths();
    const a = paths.target;
    const folder = path.dirname(a);
    const b = path.join(folder, COPY_B_FILE);
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    if (!fs.existsSync(a)) throw new Error(`${a} does not exist; this mode continues after --remove-experiment (A closed, B loaded)`);
    if (!fs.existsSync(b)) throw new Error(`${b} does not exist; this mode continues after --remove-experiment (A closed, B loaded)`);
    const runTag = timestamp().replace('-', '');
    const files = { a, b };
    report.files = { a: { role: 'A (closed by --remove-experiment; never modified)', state: fileState(a), lock: lockProbe(a) },
        b: { role: 'B (loaded by --remove-experiment; never modified)', state: fileState(b), lock: lockProbe(b) } };
    if (report.files.a.state.sha256 !== sha256(paths.source)) report.notes.push(`${a} differs from ${paths.source}.`);
    const copies = [];
    for (const tag of [...RL_ROUNDS, RL_FINAL]) {
        const file = path.join(folder, `${BRIDGE_PROJECT}-probe-${runTag}-${tag}.catvba`);
        const prepared = prepareHashedBridge(paths.source, file);
        report.files[tag] = { role: tag === RL_FINAL ? 'final copy (after B is closed)' : `repeat round ${tag}`, prepared: { action: prepared.action, ok: prepared.ok, message: prepared.message },
            state: prepared.after, lock: lockProbe(file) };
        if (!prepared.ok) throw new Error(`Preparing ${file} failed: ${prepared.message ?? prepared.code}`);
        files[tag] = file;
        if (tag !== RL_FINAL) copies.push({ tag, path: file });
    }
    const hardlink = makeHardlink(a, path.join(folder, `${path.basename(a, '.catvba')}-hl-${runTag}.catvba`));
    report.files.hl = { role: 'hard link of A', hardlink, state: hardlink.ok ? fileState(hardlink.path) : null, lock: hardlink.ok ? lockProbe(hardlink.path) : null };
    if (hardlink.ok) files.hl = hardlink.path;
    if (fileState(a).readOnly !== true) report.notes.push(`${a} is NOT read-only after making the hard link.`);
    const short = shortPathOf(a);
    report.shortPath = short;
    const variants = pathVariants(a, { short, hardlink }).map(v => ({ ...v, exists: v.library ? fs.existsSync(v.library) : null }));
    report.variants = variants;
    report.library = '(per step, see each run)';
    report.libraryPerStep = true;
    report.exe = buildProbe(ROOT);
    const first = trySnapshot(report, 'before');
    if (first.error) throw new Error(`module snapshot failed: ${first.error}`);
    checkCnext(report, first.snapshot);
    const pids = first.snapshot.cnextPids;
    report.cnextPids = pids;
    report.vbscriptBefore = first.snapshot.loaded;
    if (first.snapshot.loaded) report.notes.push('vbscript.dll was already loaded before the first step.');
    const steps = buildReloadSteps({ a, b, z: files[RL_FINAL], copies, variants });
    const { entries, aborted } = await runSequence(report, steps, step => step.library, pids, async step => {
        if (step.local !== 'locks') throw new Error(`unknown local step ${step.local}`);
        const started = Date.now();
        const locks = Object.fromEntries(step.files.filter(key => files[key]).map(key => [key, lockProbe(files[key])]));
        return { run: { local: true, kind: 'locks', ok: true, match: true, locks, durationMs: Date.now() - started }, entry: { ok: true, match: true, locks } };
    });
    for (const entry of entries) {
        const run = report.runs.find(r => r.step === entry.step);
        if (run?.result) Object.assign(entry, { elapsedMs: run.result.elapsedMs ?? null, processMs: run.durationMs ?? null });
        if (run?.actions) Object.assign(entry, { resultType: run.result?.resultType ?? null, typeInfo: run.result?.typeInfo ?? null });
    }
    for (const [key, file] of Object.entries(files)) Object.assign(report.files[key], { stateAtEnd: fileState(file), lockAtEnd: lockProbe(file) });
    for (const key of ['a', 'b']) {
        const f = report.files[key];
        if (f.state.sha256 && f.stateAtEnd.sha256 && f.state.sha256 !== f.stateAtEnd.sha256) report.notes.push(`${files[key]} changed during the run.`);
        if (f.stateAtEnd.readOnly !== true) report.notes.push(`${files[key]} is NOT read-only at the end.`);
    }
    report.cleanup = cleanupProbeFiles([...[...RL_ROUNDS, RL_FINAL].map(key => ({ key, path: files[key] })),
        ...(files.hl ? [{ key: 'hl', path: files.hl, linkOf: a }] : [])]);
    const kept = report.cleanup.filter(c => c.action !== 'deleted' && c.action !== 'missing');
    if (kept.length) report.notes.push(`Probe files left in ${folder} (${kept.map(c => `${path.basename(c.path)}: ${c.action}`).join(', ')}); delete them after CATIA is restarted.`);
    if (fileState(a).readOnly !== true) report.notes.push(`${a} is NOT read-only after the cleanup.`);
    report.notes.push('CATIA was changed by this run (C5D_Bridge copies loaded and closed, B closed). Restart CATIA before a probe run that needs a fresh CATIA.');
    const { lines, reload } = judgeReload(first.snapshot.loaded, entries, steps, { files, copies, variants, cleanup: report.cleanup });
    report.summary = { ...sequenceSummary(report, entries),
        infos: Object.fromEntries(steps.filter(s => s.info).map(s => [s.name, parseInfo(entries.find(e => e.step === s.name)?.result)])),
        reload, verdictLines: lines, verdict: lines.join('\n') };
    if (aborted) process.exitCode = 1;
}

// ---- --self-close-experiment: C5D_Install always writes into the FIRST project named C5D_Bridge, so with several copies loaded a
// copy cannot mark itself. Here (1) a writer in B (C5D_SxB.C5D_SxPut) puts the module C5D_Sx into the copy with a given path,
// (2) C5D_Sx asks apc.ExecutingProject who is running (C5D_SxWho: text, C5D_SxSelf: the IApcProject itself), (3) eval-probe.exe
// closes the returned project only after guard actions confirmed its Name / DisplayName / VBProject.FileName, (4) the production
// idea is repeated: a new hard link of the master A per Push, auto-load through that path, C5D_Install (the only C5D_Bridge = itself),
// self-identify, close from outside.
const SX_MODULE = 'C5D_Sx';
const SXB_MODULE = 'C5D_SxB';
const SX_ROUNDS = ['p1', 'p2', 'p3'];
const SX_FINAL = 'pz';
/** Never more than this many copies are touched (a list with more C5D_Bridge projects is not what this mode expects). */
const SX_MAX_TARGETS = 20;
const SX_PUT_PATTERN = /^ok;/;
const SX_LIST_PATTERN = /^list=[^;|]*;.*;count=\d+;vbeCount=\d+;err=0(\||$)/;
/** VBA: v = <expr>, then ";<label>=<value>" (or "=!<err>") appended to s (fmt = the module's own formatting function). */
const sxProp = (fmt, label, expr) => ['    Err.Clear', '    v = "?"', `    v = ${expr}`, '    e = Err.Number', `    s = s & ";${label}=" & ${fmt}(v, e)`];
/** Private helpers and the list function of one module (p = prefix of the private names: Sx in C5D_Sx, SxB in C5D_SxB). */
function sxCommonLines(p) {
    const prop = (label, expr) => sxProp(`${p}Fmt`, label, expr);
    return [
        `Private Function ${p}Clean(ByVal t As String) As String`,
        '    t = Replace(t, "|", "/")',
        '    t = Replace(t, ";", ",")',
        '    t = Replace(t, vbCr, " ")',
        '    t = Replace(t, vbLf, " ")',
        `    ${p}Clean = t`,
        'End Function',
        `Private Function ${p}Fmt(ByVal v As Variant, ByVal e As Long) As String`,
        '    On Error Resume Next',
        '    Dim t As String',
        '    If e <> 0 Then',
        `        ${p}Fmt = "!" & CStr(e)`,
        '        Exit Function',
        '    End If',
        '    t = "?"',
        '    t = CStr(v)',
        `    ${p}Fmt = ${p}Clean(t)`,
        'End Function',
        `Private Function ${p}Apc() As Object`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Set apc = Nothing',
        '    Set apc = CreateObject("MSAPC.Apc.7.1")',
        '    If apc Is Nothing Then',
        '        Err.Clear',
        '        Set apc = CreateObject("MSAPC.Apc")',
        '    End If',
        `    Set ${p}Apc = apc`,
        'End Function',
        "' \";comps=<VBComponents.Count>;c5d=<names of the components starting with C5D_>\"",
        `Private Function ${p}Comps(ByVal vbp As Object) As String`,
        '    On Error Resume Next',
        '    Dim comps As Object',
        '    Dim c As Object',
        '    Dim nm As String',
        '    Dim s As String',
        '    Dim i As Long',
        '    Dim n As Long',
        '    Set comps = Nothing',
        '    Set comps = vbp.VBComponents',
        '    n = -1',
        '    n = comps.Count',
        '    For i = 1 To n',
        '        Set c = Nothing',
        '        Set c = comps.Item(i)',
        '        nm = ""',
        '        nm = CStr(c.Name)',
        '        If Left(nm, 4) = "C5D_" Then s = s & "," & nm',
        '    Next',
        '    Err.Clear',
        `    ${p}Comps = ";comps=" & CStr(n) & ";c5d=" & ${p}Clean(Mid(s, 2))`,
        'End Function',
        `Private Function ${p}ApcProj(ByVal ap As Object, ByVal ex As Object) As String`,
        '    On Error Resume Next',
        '    Dim s As String',
        '    Dim v As Variant',
        '    Dim e As Long',
        '    Dim vbp As Object',
        '    Dim same As Boolean',
        '    s = ""',
        ...prop('name', 'ap.Name'),
        ...prop('display', 'ap.DisplayName'),
        ...prop('mode', 'ap.Mode'),
        '    Set vbp = Nothing',
        '    Set vbp = ap.VBProject',
        ...prop('file', 'vbp.FileName'),
        `    s = s & ${p}Comps(vbp)`,
        '    same = False',
        '    same = (ap Is ex)',
        '    s = s & ";exec=" & CStr(same)',
        '    Err.Clear',
        `    ${p}ApcProj = s`,
        'End Function',
        `Private Function ${p}Vbp(ByVal vbp As Object) As String`,
        '    On Error Resume Next',
        '    Dim s As String',
        '    Dim v As Variant',
        '    Dim e As Long',
        '    s = ""',
        ...prop('name', 'vbp.Name'),
        ...prop('type', 'vbp.Type'),
        ...prop('mode', 'vbp.Mode'),
        ...prop('file', 'vbp.FileName'),
        ...prop('saved', 'vbp.Saved'),
        `    s = s & ${p}Comps(vbp)`,
        '    Err.Clear',
        `    ${p}Vbp = s`,
        'End Function',
        "' \"list=<TypeName(apc)>;running=..;exec=<DisplayName of apc.ExecutingProject>;count=<apc.Projects.Count>;vbeCount=<VBProjects.Count>;err=E\"",
        "' then \"|src=apc;i=..;name=..;display=..;mode=..;file=..;comps=..;c5d=..;exec=..\" per IApcProject and \"|src=vbe;i=..;name=..;type=..;...\" per VBProject",
        `Private Function ${p}ListText() As String`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Dim aps As Object',
        '    Dim ap As Object',
        '    Dim ex As Object',
        '    Dim ide As Object',
        '    Dim vbps As Object',
        '    Dim vbp As Object',
        '    Dim s As String',
        '    Dim v As Variant',
        '    Dim e As Long',
        '    Dim i As Long',
        '    Dim n As Long',
        '    Dim m As Long',
        `    Set apc = ${p}Apc()`,
        '    If apc Is Nothing Then',
        `        ${p}ListText = "list=Nothing;err=" & CStr(Err.Number)`,
        '        Exit Function',
        '    End If',
        '    s = "list=" & TypeName(apc)',
        ...prop('running', 'apc.Running'),
        '    Set ex = Nothing',
        '    Set ex = apc.ExecutingProject',
        ...prop('exec', 'ex.DisplayName'),
        '    Err.Clear',
        '    Set aps = Nothing',
        '    Set aps = apc.Projects',
        '    n = -1',
        '    n = aps.Count',
        '    Set ide = Nothing',
        '    Set ide = apc.VBE',
        '    Set vbps = Nothing',
        '    Set vbps = ide.VBProjects',
        '    m = -1',
        '    m = vbps.Count',
        '    s = s & ";count=" & CStr(n) & ";vbeCount=" & CStr(m) & ";err=" & CStr(Err.Number)',
        '    For i = 1 To n',
        '        Set ap = Nothing',
        '        Err.Clear',
        '        Set ap = aps.Item(i)',
        '        If ap Is Nothing Then',
        '            s = s & "|src=apc;i=" & CStr(i) & ";err=" & CStr(Err.Number)',
        '        Else',
        `            s = s & "|src=apc;i=" & CStr(i) & ${p}ApcProj(ap, ex)`,
        '        End If',
        '    Next',
        '    For i = 1 To m',
        '        Set vbp = Nothing',
        '        Err.Clear',
        '        Set vbp = vbps.Item(i)',
        '        If vbp Is Nothing Then',
        '            s = s & "|src=vbe;i=" & CStr(i) & ";err=" & CStr(Err.Number)',
        '        Else',
        `            s = s & "|src=vbe;i=" & CStr(i) & ${p}Vbp(vbp)`,
        '        End If',
        '    Next',
        `    ${p}ListText = s`,
        'End Function'
    ];
}

/**
 * VBA of C5D_Sx, put into every copy that is to be closed (ASCII, CRLF, On Error Resume Next, no MsgBox, never closes/removes/saves):
 *   C5D_SxSelf()  apc.ExecutingProject (the IApcProject running this code, i.e. the copy ExecuteScript resolved the path to) as an
 *                 object, or "error:apc:<n>" / "error:noexec:<n>". eval-probe.exe reads it after ExecuteScript returned (then Running
 *                 is False and ExecutingProject is Empty, so it must be taken here, while running).
 *   C5D_SxWho()   "who=ok;running=..;name=..;display=<DisplayName>;mode=..;vbName=..;file=<VBProject.FileName>;comps=..;c5d=.."
 *   C5D_SxList()  the list of SxListText (all IApcProjects and VBProjects; read-only)
 */
/** VBA of C5D_<p>Self (apc.ExecutingProject as an object) and C5D_<p>Who (its text), using the private helpers of sxCommonLines(p). */
function sxSelfWhoLines(p) {
    const self = `C5D_${p}Self`, who = `C5D_${p}Who`;
    return [
        `Public Function ${self}() As Variant`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Dim ex As Object',
        `    Set apc = ${p}Apc()`,
        '    If apc Is Nothing Then',
        `        ${self} = "error:apc:" & CStr(Err.Number)`,
        '        Exit Function',
        '    End If',
        '    Err.Clear',
        '    Set ex = Nothing',
        '    Set ex = apc.ExecutingProject',
        '    If ex Is Nothing Then',
        `        ${self} = "error:noexec:" & CStr(Err.Number)`,
        '        Exit Function',
        '    End If',
        `    Set ${self} = ex`,
        'End Function',
        `Public Function ${who}() As Variant`,
        '    On Error Resume Next',
        '    Dim apc As Object',
        '    Dim ex As Object',
        '    Dim vbp As Object',
        '    Dim s As String',
        '    Dim v As Variant',
        '    Dim e As Long',
        `    Set apc = ${p}Apc()`,
        '    If apc Is Nothing Then',
        `        ${who} = "error:apc:" & CStr(Err.Number)`,
        '        Exit Function',
        '    End If',
        '    s = ""',
        ...sxProp(`${p}Fmt`, 'running', 'apc.Running'),
        '    Err.Clear',
        '    Set ex = Nothing',
        '    Set ex = apc.ExecutingProject',
        '    If ex Is Nothing Then',
        `        ${who} = "error:noexec:" & CStr(Err.Number) & s`,
        '        Exit Function',
        '    End If',
        ...sxProp(`${p}Fmt`, 'name', 'ex.Name'),
        ...sxProp(`${p}Fmt`, 'display', 'ex.DisplayName'),
        ...sxProp(`${p}Fmt`, 'mode', 'ex.Mode'),
        '    Set vbp = Nothing',
        '    Set vbp = ex.VBProject',
        ...sxProp(`${p}Fmt`, 'vbName', 'vbp.Name'),
        ...sxProp(`${p}Fmt`, 'file', 'vbp.FileName'),
        `    s = s & ${p}Comps(vbp)`,
        '    Err.Clear',
        `    ${who} = "who=ok" & s`,
        'End Function'
    ];
}
const SX_SOURCE = [
    'Option Explicit',
    ...sxCommonLines('Sx'),
    ...sxSelfWhoLines('Sx'),
    'Public Function C5D_SxList() As Variant',
    '    C5D_SxList = SxListText()',
    'End Function',
    ''
].join('\r\n');

/**
 * VBA of C5D_SxB, written into B by C5D_Install (B is the first C5D_Bridge, so C5D_Install through B's path writes into B):
 *   C5D_SxPut(key, source)  among the IApcProjects NAMED C5D_Bridge only (never a user project): pass 1 DisplayName = key exactly,
 *     pass 2 DisplayName ignoring case, pass 3 VBProject.FileName ignoring case (dot/dotdot/slash variants share one FileName, so
 *     FileName alone is ambiguous; DisplayName keeps the string ExecuteScript was given). Exactly one hit in a pass -> the standard
 *     module C5D_Sx of that project gets source (added, or its lines replaced). Only C5D_Sx is ever written (never C5D_Bridge_Push).
 *     -> "ok;pass=N;existed=<True|False>;lines=N;display=..;file=.." | "error:ambiguous:passN:<hits>" | "error:notfound:count=N" | "error:..."
 *   C5D_SxBList()  the same list as C5D_SxList (read-only)
 */
const SXB_SOURCE = [
    'Option Explicit',
    `Private Const SX_PROJECT As String = "${BRIDGE_PROJECT}"`,
    `Private Const SX_MODULE As String = "${SX_MODULE}"`,
    ...sxCommonLines('SxB'),
    'Public Function C5D_SxPut(ByVal key As Variant, ByVal source As Variant) As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim hit As Object',
    '    Dim vbp As Object',
    '    Dim comps As Object',
    '    Dim comp As Object',
    '    Dim cm As Object',
    '    Dim k As String',
    '    Dim nm As String',
    '    Dim dn As String',
    '    Dim fl As String',
    '    Dim ok As Boolean',
    '    Dim existed As Boolean',
    '    Dim pass As Long',
    '    Dim found As Long',
    '    Dim hits As Long',
    '    Dim i As Long',
    '    Dim n As Long',
    '    Dim cnt As Long',
    '    Dim e As Long',
    '    Set apc = SxBApc()',
    '    If apc Is Nothing Then',
    '        C5D_SxPut = "error:apc:" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    k = ""',
    '    k = CStr(key)',
    '    If Len(k) = 0 Then',
    '        C5D_SxPut = "error:key:empty"',
    '        Exit Function',
    '    End If',
    '    Set aps = Nothing',
    '    Set aps = apc.Projects',
    '    n = -1',
    '    n = aps.Count',
    '    found = 0',
    '    For pass = 1 To 3',
    '        hits = 0',
    '        Set hit = Nothing',
    '        For i = 1 To n',
    '            Set ap = Nothing',
    '            Set ap = aps.Item(i)',
    '            nm = ""',
    '            nm = CStr(ap.Name)',
    '            dn = ""',
    '            dn = CStr(ap.DisplayName)',
    '            fl = ""',
    '            fl = CStr(ap.VBProject.FileName)',
    '            ok = False',
    '            If nm = SX_PROJECT Then',
    '                If pass = 1 Then',
    '                    ok = (StrComp(dn, k, vbBinaryCompare) = 0)',
    '                ElseIf pass = 2 Then',
    '                    ok = (StrComp(dn, k, vbTextCompare) = 0)',
    '                Else',
    '                    ok = (StrComp(fl, k, vbTextCompare) = 0)',
    '                End If',
    '            End If',
    '            If ok Then',
    '                hits = hits + 1',
    '                Set hit = ap',
    '            End If',
    '        Next',
    '        If hits > 1 Then',
    '            C5D_SxPut = "error:ambiguous:pass" & CStr(pass) & ":" & CStr(hits)',
    '            Exit Function',
    '        End If',
    '        If hits = 1 Then',
    '            found = pass',
    '            Exit For',
    '        End If',
    '    Next',
    '    If found = 0 Then',
    '        C5D_SxPut = "error:notfound:count=" & CStr(n)',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    Set vbp = Nothing',
    '    Set vbp = hit.VBProject',
    '    If vbp Is Nothing Then',
    '        C5D_SxPut = "error:vbproject:" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Set comps = Nothing',
    '    Set comps = vbp.VBComponents',
    '    If comps Is Nothing Then',
    '        C5D_SxPut = "error:components:" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Set comp = Nothing',
    '    Set comp = comps.Item(SX_MODULE)',
    '    Err.Clear',
    '    existed = Not (comp Is Nothing)',
    '    If Not existed Then',
    '        Set comp = comps.Add(1)',
    '        If comp Is Nothing Then',
    '            C5D_SxPut = "error:add:" & CStr(Err.Number) & ":" & SxBClean(Err.Description)',
    '            Exit Function',
    '        End If',
    '        Err.Clear',
    '        comp.Name = SX_MODULE',
    '        e = Err.Number',
    '        If e <> 0 Then',
    '            C5D_SxPut = "error:name:" & CStr(e) & ":" & SxBClean(Err.Description)',
    '            Exit Function',
    '        End If',
    '    End If',
    '    Set cm = Nothing',
    '    Set cm = comp.CodeModule',
    '    If cm Is Nothing Then',
    '        C5D_SxPut = "error:codemodule:" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    cnt = 0',
    '    cnt = cm.CountOfLines',
    '    Err.Clear',
    '    If cnt > 0 Then cm.DeleteLines 1, cnt',
    '    e = Err.Number',
    '    If e <> 0 Then',
    '        C5D_SxPut = "error:delete:" & CStr(e) & ":" & SxBClean(Err.Description)',
    '        Exit Function',
    '    End If',
    '    cm.AddFromString CStr(source)',
    '    e = Err.Number',
    '    If e <> 0 Then',
    '        C5D_SxPut = "error:addfromstring:" & CStr(e) & ":" & SxBClean(Err.Description)',
    '        Exit Function',
    '    End If',
    '    cnt = -1',
    '    cnt = cm.CountOfLines',
    '    dn = ""',
    '    dn = CStr(hit.DisplayName)',
    '    fl = ""',
    '    fl = CStr(vbp.FileName)',
    '    Err.Clear',
    '    C5D_SxPut = "ok;pass=" & CStr(found) & ";existed=" & CStr(existed) & ";lines=" & CStr(cnt) & ";display=" & SxBClean(dn) & ";file=" & SxBClean(fl)',
    'End Function',
    'Public Function C5D_SxBList() As Variant',
    '    C5D_SxBList = SxBListText()',
    'End Function',
    ''
].join('\r\n');

/** "k=v;k=v" -> { k: v } (the first = splits; values never contain ; or |). */
const sxFields = text => Object.fromEntries(String(text ?? '').split(';').filter(Boolean).map(kv => {
    const i = kv.indexOf('=');
    return i < 0 ? [kv, ''] : [kv.slice(0, i), kv.slice(i + 1)];
}));
/** C5D_SxList / C5D_SxBList text -> { header, count, vbeCount, apc: [{..}], vbe: [{..}] } or null (not a list). */
function parseSxList(text) {
    if (typeof text !== 'string' || !text.startsWith('list=')) return null;
    const [head, ...rest] = text.split('|');
    const header = sxFields(head);
    const num = value => /^\d+$/.test(value ?? '') ? Number(value) : null;
    const items = rest.map(sxFields);
    return { header, count: num(header.count), vbeCount: num(header.vbeCount), err: header.err ?? null,
        apc: items.filter(p => p.src === 'apc'), vbe: items.filter(p => p.src === 'vbe') };
}
/** C5D_SxWho text -> fields ({ who: 'ok', display, file, ... }), { error: <text> } for "error:..." or null. */
function parseWho(text) {
    if (typeof text !== 'string') return null;
    if (text.startsWith('who=')) return sxFields(text);
    return text.startsWith('error:') ? { error: text } : null;
}
/** The C5D_Bridge entries (apc part) of a parsed list. */
const sxBridges = list => (list?.apc ?? []).filter(p => p.name === BRIDGE_PROJECT);
/** The user projects (everything not named C5D_Bridge) of a parsed list: { apc: [{ name, file, comps }], vbe: [{ name, file, comps, saved }] }. */
function sxUserProjects(list) {
    if (!list) return null;
    const user = p => p.name !== BRIDGE_PROJECT;
    return { apc: list.apc.filter(user).map(p => ({ name: p.name, file: p.file, comps: p.comps })),
        vbe: list.vbe.filter(user).map(p => ({ name: p.name, file: p.file, comps: p.comps, saved: p.saved })) };
}
/** Whether a C5D_SxWho result says the running project is the C5D_Bridge shown as library (DisplayName, ignoring case). */
function whoIsSelf(result, library) {
    const who = parseWho(result);
    return Boolean(who && who.who === 'ok' && who.name === BRIDGE_PROJECT && sameText(who.display, library));
}
/** VBA-style quoted literal for a guard action ('' for '). */
const sxLiteral = text => `'${String(text).replace(/'/g, "''")}'`;
/**
 * Actions of a self-close step: read Name / DisplayName / VBProject.FileName of the returned project, then three guards (Name, the
 * library string, the expected FileName). Only if all three hold, Close(false) (index SX_CLOSE_ACTION) and a Name read that should
 * fail (disconnected) afterwards.
 */
const sxCloseActions = (display, file) => ['get:$.Name', 'get:$.DisplayName', 'get:$.VBProject.FileName', `guard:$.Name=${sxLiteral(BRIDGE_PROJECT)}`,
    `guard:$.DisplayName=${sxLiteral(display)}`, `guard:$.VBProject.FileName=${sxLiteral(file)}`, 'call:$.Close(false)', 'get:$.Name'];
const SX_CLOSE_ACTION = 6;

const sxCall = (name, library, module, func, args = [], extra = {}) => ({ name, library, module, func, args, ...extra });
/** List through B (C5D_SxB) or through the copy itself (C5D_Sx). Info: never requires anything (check, if given, only sets match). */
const sxListStep = (name, library, via, extra = {}) => sxCall(name, library, via === 'b' ? SXB_MODULE : SX_MODULE, via === 'b' ? 'C5D_SxBList' : 'C5D_SxList', [],
    { pattern: SX_LIST_PATTERN, info: 'sx', ...extra });
const sxWhoStep = (name, library, extra = {}) => sxCall(name, library, SX_MODULE, 'C5D_SxWho', [], { check: result => whoIsSelf(result, library), self: library, ...extra });
const sxCloseStep = (name, library, file, extra = {}) => sxCall(name, library, SX_MODULE, 'C5D_SxSelf', [],
    { actions: sxCloseActions(library, file), mustSucceed: [SX_CLOSE_ACTION], resultType: 'ComObject', closeOf: library, ...extra });
const sxRecallStep = (name, library, extra = {}) => sxCall(name, library, BRIDGE_MODULE, 'C5D_Version', [], { expectFailure: BRIDGE_VERSION, ...extra });

/** Phase 1 (through B): identify B, install the writer, list, put C5D_Sx into B, B identifies itself. */
function buildSelfCloseStart(b) {
    return [
        sxCall('sc-b-alive', b, B_MODULE, 'C5D_BAlive', [], { expected: B_ALIVE }),
        sxCall('sc-b-install', b, BRIDGE_MODULE, 'C5D_Install', [SXB_MODULE, SXB_SOURCE], { expected: 'ok' }),
        sxListStep('sc-start-list', b, 'b'),
        sxCall('sc-b-put', b, SXB_MODULE, 'C5D_SxPut', [b, SX_SOURCE], { pattern: SX_PUT_PATTERN, requires: ['sc-b-install'], putOf: b }),
        sxWhoStep('sc-b-who', b, { requires: ['sc-b-put'] })
    ];
}
/** Phase 2: every other copy t<n> = { tag, display, file }: put, who (own path), guarded self-close, list through B, same path again. */
function buildSelfCloseTargets(b, targets) {
    const steps = [];
    for (const t of targets) {
        steps.push(
            sxCall(`${t.tag}-put`, b, SXB_MODULE, 'C5D_SxPut', [t.display, SX_SOURCE], { pattern: SX_PUT_PATTERN, requires: ['sc-b-install'], target: t.tag, putOf: t.display }),
            sxWhoStep(`${t.tag}-who`, t.display, { requires: [`${t.tag}-put`], target: t.tag }),
            sxCloseStep(`${t.tag}-close`, t.display, t.file, { requires: [`${t.tag}-who`], target: t.tag }),
            sxListStep(`${t.tag}-gone`, b, 'b', { target: t.tag }),
            sxRecallStep(`${t.tag}-recall`, t.display, { after: [`${t.tag}-close`], target: t.tag }));
    }
    return steps;
}
/** Phase 3: B closes itself, but only when the list through B shows no other C5D_Bridge. */
function buildSelfCloseB(b, bFile = b) {
    return [
        sxListStep('b-list', b, 'b', { check: result => { const l = parseSxList(result); const br = sxBridges(l); return SX_LIST_PATTERN.test(result) && br.length === 1 && sameText(br[0].display, b); } }),
        sxCloseStep('b-close', b, bFile, { requires: ['sc-b-who', 'b-list'] }),
        sxRecallStep('b-recall', b, { after: ['b-close'] })
    ];
}
/**
 * Phase 4: the production idea per round r = { tag, path } (a new hard link of A): link, auto-load, C5D_Install (writes into the
 * first C5D_Bridge = this one), who, list (C5D_Bridge x1 = this link, userCount + 1 projects), guarded self-close, same path again,
 * lock. Round n starts only after round n-1 (or B) was closed (else C5D_Install would write into the left-over project).
 */
function buildSelfCloseRounds(rounds, userCount = null, first = 'b-close') {
    const steps = [];
    let previous = first;
    for (const r of rounds) {
        const p = r.path;
        const listCheck = result => {
            const l = parseSxList(result), br = sxBridges(l);
            return Boolean(l && SX_LIST_PATTERN.test(result) && br.length === 1 && sameText(br[0].display, p)
                && (userCount === null || (l.count === userCount + 1 && l.vbeCount === userCount + 1)));
        };
        steps.push(
            { name: `${r.tag}-link`, local: 'link', path: p, round: r.tag, requires: [previous] },
            sxCall(`${r.tag}-load`, p, BRIDGE_MODULE, 'C5D_Version', [], { expected: BRIDGE_VERSION, requires: [`${r.tag}-link`], round: r.tag }),
            sxCall(`${r.tag}-install`, p, BRIDGE_MODULE, 'C5D_Install', [SX_MODULE, SX_SOURCE], { expected: 'ok', requires: [`${r.tag}-load`], round: r.tag }),
            sxWhoStep(`${r.tag}-who`, p, { requires: [`${r.tag}-install`], round: r.tag }),
            sxListStep(`${r.tag}-list`, p, 'self', { after: [`${r.tag}-load`], check: listCheck, round: r.tag }),
            sxCloseStep(`${r.tag}-close`, p, p, { requires: [`${r.tag}-who`], round: r.tag }),
            sxRecallStep(`${r.tag}-recall`, p, { after: [`${r.tag}-close`], round: r.tag }),
            { name: `${r.tag}-lock`, local: 'locks', files: [r.tag], after: [`${r.tag}-link`], round: r.tag });
        previous = `${r.tag}-close`;
    }
    return steps;
}
/** Hard link names of the rounds: C5D_Bridge-<8 hex>-p<yyyyMMddHHmmss>-<n>.catvba next to A (n = 1..4; 4 = the final check pz). */
function selfCloseRounds(a, runTag) {
    const dir = path.win32.dirname(a), base = path.win32.basename(a, '.catvba');
    return [...SX_ROUNDS, SX_FINAL].map((tag, i) => ({ tag, path: path.win32.join(dir, `${base}-p${runTag}-${i + 1}.catvba`) }));
}
/**
 * The copies to close: every C5D_Bridge of the start list (apc part) except B, in list order, as { tag: t<n>, display, file, label }.
 * label = the tag of the --reload-experiment path it equals (expected: [{ tag, path }]), else null. Throws above SX_MAX_TARGETS.
 */
function selfCloseTargets(list, b, expected = []) {
    const others = sxBridges(list).filter(p => !sameText(p.display, b));
    if (others.length > SX_MAX_TARGETS) throw new Error(`${others.length} C5D_Bridge copies loaded (more than ${SX_MAX_TARGETS}); not what this mode expects`);
    return others.map((p, i) => ({ tag: `t${i + 1}`, display: p.display, file: p.file,
        label: (expected.find(x => x.path === p.display) ?? expected.find(x => sameText(x.path, p.display)))?.tag ?? null }));
}
/** The C5D_Bridge paths the latest --reload-experiment left loaded: [{ tag, path }] (copies c1..c3, z and the loaded path variants), or null. */
function reloadLeftovers(directory = TEMP) {
    let names;
    try { names = fs.readdirSync(directory); } catch { return null; }
    const latest = names.filter(name => /^probe-bridge-reload-\d{8}-\d{6}(-\d+)?\.json$/.test(name)).sort((x, y) => x.localeCompare(y, 'en', { numeric: true })).pop();
    if (!latest) return null;
    try {
        const r = JSON.parse(fs.readFileSync(path.join(directory, latest), 'utf8'));
        const loaded = name => r.summary?.steps?.find(s => s.step === name)?.match === true;
        const list = [...RL_ROUNDS, RL_FINAL].filter(tag => loaded(`${tag}-load`) && r.files?.[tag]?.state?.path).map(tag => ({ tag, path: r.files[tag].state.path }));
        for (const v of r.variants ?? []) if (!v.disabled && loaded(`${v.tag}-load`)) list.push({ tag: v.tag, path: v.library });
        return { report: path.join(directory, latest), list };
    } catch (error) { return { report: path.join(directory, latest), error: errorText(error), list: [] }; }
}
/** All step names of --self-close-experiment for n copies (the targets come from the live list, so the names depend on n). */
function selfCloseStepNames(n = 0) {
    const x = 'C:\\x\\C5D_Bridge-x.catvba';
    const targets = Array.from({ length: n }, (_, i) => ({ tag: `t${i + 1}`, display: x, file: x }));
    return [...buildSelfCloseStart(x), ...buildSelfCloseTargets(x, targets), ...buildSelfCloseB(x), ...buildSelfCloseRounds(selfCloseRounds(x, '0')), { name: 'locks-end' }]
        .map(s => s.name);
}

// -- judging
const actionOf = (e, index) => (e?.actions ?? []).find(a => a.action === index);
/** { ok, ms, display, file, guard, error } of a self-close step (guard: the first failed guard, if any; closeIndex: the Close action). */
function sxCloseFacts(e, closeIndex = SX_CLOSE_ACTION) {
    if (!e || e.skipped) return { ok: null, ms: null, display: null, file: null, guard: null, error: e ? e.skipped : 'not in this run' };
    const get = i => { const a = actionOf(e, i); return a?.ok ? a.result : null; };
    const failedGuard = (e.actions ?? []).find(a => String(a.op).startsWith('guard:') && a.ok === false && !a.skipped);
    const close = actionOf(e, closeIndex);
    const base = { display: get(1), file: get(2), guard: failedGuard ? `${failedGuard.op} -> ${failedGuard.message ?? failedGuard.hresult}` : null };
    if (!close) return { ok: false, ms: null, ...base, error: e.ok ? `no object (${JSON.stringify(e.result)})` : rlOutcome(e) };
    if (close.skipped) return { ok: false, ms: null, ...base, error: base.guard ? `guard failed: ${base.guard}` : close.message ?? rlOutcome(e) };
    return { ok: close.ok === true, ms: close.elapsedMs ?? null, ...base, error: close.ok ? null : `${close.hresult} ${String(close.message ?? '').split(/\r?\n/)[0]}`.trim() };
}
const closeWord = c => c.ok === null ? `close not run (${c.error})` : c.ok ? `closed ${c.ms}ms` : `NOT closed (${c.error})`;
const whoWord = (e, library) => {
    if (!e || e.skipped || !e.ok) return `who ${rlOutcome(e)}`;
    const who = parseWho(e.result);
    if (!who || who.who !== 'ok') return `who ${JSON.stringify(e.result)}`;
    const exact = who.display === library;
    return `ExecutingProject = ${JSON.stringify(who.display)} file ${JSON.stringify(who.file)} -> self ${whoIsSelf(e.result, library) ? (exact ? 'yes (exact)' : 'yes (ignoring case)') : 'NO'}`;
};
const putWord = e => !e || e.skipped || !e.ok ? `put ${rlOutcome(e)}` : SX_PUT_PATTERN.test(e.result)
    ? `put ok (pass ${sxFields(e.result).pass}, existed ${sxFields(e.result).existed})` : `put ${JSON.stringify(e.result)}`;
const listOf = e => (e && !e.skipped && e.ok ? parseSxList(e.result) : null);
const bridgeNames = list => sxBridges(list).map(p => baseName(p.display) + (p.display.includes('\\.\\') ? ' (\\.\\)' : p.display.includes('\\..\\') ? ' (\\..\\)'
    : p.display.includes('/') ? ' (/)' : p.display.startsWith('\\\\?\\') ? ' (\\\\?\\)' : ''));
const describeSxList = e => {
    if (!e || e.skipped || !e.ok) return rlOutcome(e);
    const l = parseSxList(e.result);
    if (!l) return `unexpected ${JSON.stringify(e.result)}`;
    const br = bridgeNames(l);
    return `apc ${l.count} / vbe ${l.vbeCount} projects, ${BRIDGE_PROJECT} x${br.length}${br.length ? ` [${br.join('; ')}]` : ''}${l.err !== '0' ? ` err=${l.err}` : ''}`;
};
/** Differences of the user projects between two lists (name, FileName, module count; VBE also Saved), [] = identical. */
function userDiff(start, end) {
    if (!start || !end) return null;
    const diffs = [];
    for (const kind of ['apc', 'vbe']) {
        const a = start[kind], b = end[kind];
        if (a.length !== b.length) diffs.push(`${kind}: ${a.length} -> ${b.length} projects`);
        for (let i = 0; i < Math.max(a.length, b.length); i++) {
            const x = a[i], y = b[i];
            if (JSON.stringify(x) !== JSON.stringify(y)) diffs.push(`${kind} #${i + 1}: ${JSON.stringify(x ?? null)} -> ${JSON.stringify(y ?? null)}`);
        }
    }
    return diffs;
}

/**
 * Verdict of --self-close-experiment: { lines, data }. ctx: { b, a, targets, rounds, leftovers, cleanup, userCount }.
 * entries carry elapsedMs (ExecuteScript time measured by eval-probe.exe) and, for local link steps, ms.
 */
function judgeSelfClose(beforeLoaded, entries, ctx) {
    const get = name => entries.find(e => e.step === name);
    const lines = [];
    const startList = listOf(get('sc-start-list'));
    lines.push(`start: B ${baseName(ctx.b)} ${get('sc-b-alive')?.match ? 'answers C5D_ProbeB (the B of the earlier runs)' : `C5D_ProbeB ${rlOutcome(get('sc-b-alive'))}`}, `
        + `C5D_SxB via C5D_Install ${get('sc-b-install')?.match ? 'ok' : rlOutcome(get('sc-b-install'))}`);
    lines.push(`start list (through B): ${describeSxList(get('sc-start-list'))}`);
    const users = sxUserProjects(startList);
    if (users) lines.push(`user projects at start: ${users.vbe.length} (${users.vbe.map(p => `${p.name}[${p.comps}]`).join(', ')})`);
    if (ctx.leftovers) {
        const live = ctx.targets.map(t => t.display);
        const expected = ctx.leftovers.list.map(x => x.path);
        const missing = expected.filter(p => !live.some(l => l === p));
        const extra = ctx.targets.filter(t => !t.label).map(t => t.display);
        lines.push(`copies other than B: ${ctx.targets.length} live vs ${expected.length} left by ${baseName(ctx.leftovers.report)}`
            + `${missing.length ? `; not loaded any more: ${missing.map(p => JSON.stringify(p)).join(', ')}` : ''}${extra.length ? `; not from that run: ${extra.map(p => JSON.stringify(p)).join(', ')}` : ''}`);
    } else lines.push(`copies other than B: ${ctx.targets.length} (no --reload-experiment report to compare with)`);
    lines.push(`B: ${putWord(get('sc-b-put'))}, ${whoWord(get('sc-b-who'), ctx.b)}`);
    let previous = startList ? sxBridges(startList).length : null;
    const targets = ctx.targets.map(t => {
        const close = sxCloseFacts(get(`${t.tag}-close`));
        const gone = listOf(get(`${t.tag}-gone`));
        const now = gone ? sxBridges(gone).length : null;
        const left = gone ? sxBridges(gone).some(p => p.display === t.display) : null;
        const fact = { tag: t.tag, label: t.label, display: t.display, file: t.file, put: get(`${t.tag}-put`)?.result ?? null,
            self: get(`${t.tag}-who`)?.match ?? null, who: parseWho(get(`${t.tag}-who`)?.result), close,
            counts: [previous, now], gone: left === null ? null : !left, recall: get(`${t.tag}-recall`) ? recallWord(get(`${t.tag}-recall`)) : null };
        lines.push(`target ${t.tag}${t.label ? ` (${t.label})` : ''} ${JSON.stringify(t.display)}: ${putWord(get(`${t.tag}-put`))}, ${whoWord(get(`${t.tag}-who`), t.display)}, `
            + `${closeWord(close)}, ${BRIDGE_PROJECT} ${previous ?? '?'} -> ${now ?? '?'}, gone ${yn(fact.gone)}, same path after close ${fact.recall ?? '-'}`);
        if (now !== null) previous = now;
        return fact;
    });
    const bClose = sxCloseFacts(get('b-close'));
    lines.push(`B before its close: ${describeSxList(get('b-list'))} (only B left: ${yn(get('b-list')?.skipped ? null : get('b-list')?.match ?? null)}); `
        + `B ${closeWord(bClose)}, B path after close ${get('b-recall') ? recallWord(get('b-recall')) : '-'}`);
    const rounds = ctx.rounds.map(r => {
        const link = get(`${r.tag}-link`), load = get(`${r.tag}-load`), install = get(`${r.tag}-install`), list = get(`${r.tag}-list`);
        const close = sxCloseFacts(get(`${r.tag}-close`));
        const lock = get(`${r.tag}-lock`)?.locks?.[r.tag] ?? null;
        const fact = { tag: r.tag, path: r.path, link: link?.skipped ? null : link?.match ?? null, linkMs: link?.ms ?? null,
            load: load?.skipped ? null : load?.match ?? null, loadMs: load?.elapsedMs ?? null,
            install: install?.skipped ? null : install?.match ?? null, installMs: install?.elapsedMs ?? null,
            self: get(`${r.tag}-who`)?.skipped ? null : get(`${r.tag}-who`)?.match ?? null, listOk: list?.skipped ? null : list?.match ?? null,
            counts: listOf(list) ? { apc: listOf(list).count, vbe: listOf(list).vbeCount, bridges: sxBridges(listOf(list)).length } : null,
            close, recall: get(`${r.tag}-recall`) ? recallWord(get(`${r.tag}-recall`)) : null, lockAfterClose: lock };
        const head = `round ${r.tag}${r.tag === SX_FINAL ? ' (final check)' : ''} ${baseName(r.path)}: `;
        if (!link || link.skipped) lines.push(head + `not run (${link?.skipped ?? 'not in this run'})`);
        else {
            lines.push(head + `link ${fact.link ? `ok ${fact.linkMs}ms` : `FAILED ${link.message ?? ''}`}, load ${fact.load ? `ok ${fact.loadMs}ms` : rlOutcome(load)}, `
                + `install ${fact.install ? `ok ${fact.installMs}ms` : rlOutcome(install)}, ${whoWord(get(`${r.tag}-who`), r.path)}, `
                + `list ${describeSxList(list)} (C5D_Bridge x1 = this link and user ${ctx.userCount ?? '?'} + 1: ${yn(fact.listOk)}), `
                + `${closeWord(close)}, same path after close ${fact.recall ?? '-'}, lock after close ${lockWord(lock)}`);
        }
        return fact;
    });
    const doneRounds = rounds.filter(r => r.close.ok);
    if (rounds.length) {
        const ms = key => doneRounds.map(r => key(r)).filter(n => typeof n === 'number');
        const avg = list => list.length ? Math.round(list.reduce((x, y) => x + y, 0) / list.length) : '-';
        lines.push(`rounds closed: ${doneRounds.length}/${rounds.length}; average load ${avg(ms(r => r.loadMs))}ms, install ${avg(ms(r => r.installMs))}ms, `
            + `close ${avg(ms(r => r.close.ms))}ms`);
    }
    const all = [...targets.map(t => ({ self: t.self, close: t.close })), { self: get('sc-b-who')?.match ?? null, close: bClose },
        ...rounds.map(r => ({ self: r.self, close: r.close }))];
    const tried = all.filter(x => x.self !== null);
    lines.push(`self-identification by apc.ExecutingProject through the own path: ${tried.filter(x => x.self).length}/${tried.length}; `
        + `guarded Close: ${all.filter(x => x.close.ok).length}/${all.filter(x => x.close.ok !== null).length} closed`
        + `${all.some(x => x.close.guard) ? `, a guard stopped ${all.filter(x => x.close.guard).length}` : ''}`);
    const lastList = [...ctx.rounds].reverse().map(r => listOf(get(`${r.tag}-list`))).find(Boolean) ?? listOf(get('b-list'));
    const diff = userDiff(users, sxUserProjects(lastList));
    lines.push(diff === null ? 'user projects start vs end: not compared (a list is missing)'
        : diff.length ? `user projects start vs end: DIFFER: ${diff.join('; ')}` : `user projects start vs end: identical (${users.vbe.length} projects: name, FileName, module count, Saved)`);
    const locks = get('locks-end')?.locks ?? {};
    lines.push(`locks at end: ${Object.entries(locks).map(([k, l]) => `${k} ${lockWord(l)}`).join(', ') || 'not probed'}`);
    if (ctx.cleanup) lines.push(`cleanup: ${ctx.cleanup.map(c => `${c.key} ${c.action}${c.error ? ` (${c.error})` : ''}`).join(', ') || 'nothing'}`
        + `${ctx.aReadOnly === undefined ? '' : `; A read-only afterwards ${yn(ctx.aReadOnly)}`}`);
    const crashed = entries.find(e => e.crashed);
    if (crashed) lines.push(`CATIA crashed after step ${crashed.step}`);
    const timedOut = entries.find(e => e.timedOut);
    if (timedOut) lines.push(`timed out at step ${timedOut.step} (CATIA may be stopped at a dialog)`);
    lines.push(describeVbscript(beforeLoaded, entries));
    return { lines, data: { targets, bClose, rounds, userDiff: diff, locks, cleanup: ctx.cleanup ?? null } };
}

/**
 * --self-close-experiment (the SAME CATIA the reload experiment left: B first, other copies loaded; NOT restarted). Phase 1 through
 * B; the copies to close come from its start list; then phases 2-4 (buildSelfCloseTargets / buildSelfCloseB / buildSelfCloseRounds)
 * one --no-retry call at a time. Only a timeout, an unreachable CATIA or a vanished CNEXT.exe stops the run. At the end the hard
 * links are deleted when not locked (A read-only again).
 */
async function runSelfClose(report) {
    const paths = hashedPaths();
    const a = paths.target;
    const b = path.join(path.dirname(a), COPY_B_FILE);
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    for (const f of [a, b]) if (!fs.existsSync(f)) throw new Error(`${f} does not exist; this mode continues after --reload-experiment (B and the copies loaded)`);
    const runTag = timestamp().replace('-', '');
    const rounds = selfCloseRounds(a, runTag);
    const files = { a, b, ...Object.fromEntries(rounds.map(r => [r.tag, r.path])) };
    report.files = { a: { role: 'A (master, read-only; hard links are made of it; never modified)', state: fileState(a), lock: lockProbe(a) },
        b: { role: 'B (first C5D_Bridge; never modified on disk)', state: fileState(b), lock: lockProbe(b) } };
    if (report.files.a.state.readOnly !== true) report.notes.push(`${a} is NOT read-only at the start.`);
    if (report.files.a.state.sha256 !== sha256(paths.source)) report.notes.push(`${a} differs from ${paths.source}.`);
    for (const r of rounds) if (fs.existsSync(r.path)) throw new Error(`${r.path} already exists`);
    const leftovers = reloadLeftovers();
    report.leftovers = leftovers;
    report.library = '(per step, see each run)';
    report.libraryPerStep = true;
    report.exe = buildProbe(ROOT);
    const first = trySnapshot(report, 'before');
    if (first.error) throw new Error(`module snapshot failed: ${first.error}`);
    checkCnext(report, first.snapshot);
    const pids = first.snapshot.cnextPids;
    report.cnextPids = pids;
    report.vbscriptBefore = first.snapshot.loaded;
    if (first.snapshot.loaded) report.notes.push('vbscript.dll was already loaded before the first step.');
    const local = async step => {
        const started = Date.now();
        if (step.local === 'link') {
            const link = makeHardlink(a, step.path);
            const ms = Date.now() - started;
            report.files[step.round] = { role: `hard link of A for round ${step.round}`, hardlink: link, state: link.ok ? fileState(step.path) : null, lock: link.ok ? lockProbe(step.path) : null };
            const aReadOnly = fileState(a).readOnly;
            return { run: { local: true, kind: 'link', ok: link.ok, match: link.ok, ...link, aReadOnly, durationMs: ms, message: link.error },
                entry: { ok: link.ok, match: link.ok, ms, path: step.path, message: link.error, aReadOnly } };
        }
        if (step.local !== 'locks') throw new Error(`unknown local step ${step.local}`);
        const locks = Object.fromEntries(step.files.filter(key => files[key] && fs.existsSync(files[key])).map(key => [key, lockProbe(files[key])]));
        return { run: { local: true, kind: 'locks', ok: true, match: true, locks, durationMs: Date.now() - started }, entry: { ok: true, match: true, locks } };
    };
    // phase 1, then the copies from its list
    const start = buildSelfCloseStart(b);
    let state = await runSequence(report, start, step => step.library, pids, local);
    const startList = listOf(state.entries.find(e => e.step === 'sc-start-list'));
    if (!startList) report.notes.push('The start list through B failed; no copy is closed (B and the rounds are still tried).');
    const targets = selfCloseTargets(startList, b, leftovers?.list ?? []);
    report.targets = targets;
    const userCount = startList ? sxUserProjects(startList).vbe.length : null;
    const bFile = sxBridges(startList).find(p => sameText(p.display, b))?.file || b;
    // lock probes at the end: A, B, the hard links and the copies (the files behind their path strings; nothing is deleted)
    targets.forEach(t => { files[t.tag] = t.display; });
    const rest = [...buildSelfCloseTargets(b, targets), ...buildSelfCloseB(b, bFile), ...buildSelfCloseRounds(rounds, userCount),
        { name: 'locks-end', local: 'locks', files: Object.keys(files) }];
    state = await runSequence(report, rest, step => step.library, pids, local, state);
    const { entries, aborted } = state;
    const steps = [...start, ...rest];
    report.steps = steps.map(s => s.name);
    for (const entry of entries) {
        const run = report.runs.find(r => r.step === entry.step);
        if (run?.result) Object.assign(entry, { elapsedMs: run.result.elapsedMs ?? null, processMs: run.durationMs ?? null });
        if (run?.actions) Object.assign(entry, { resultType: run.result?.resultType ?? null, typeInfo: run.result?.typeInfo ?? null });
    }
    for (const [key, f] of Object.entries(report.files)) {
        const file = files[key];
        if (file) Object.assign(f, { stateAtEnd: fileState(file), lockAtEnd: lockProbe(file) });
    }
    for (const key of ['a', 'b']) {
        const f = report.files[key];
        if (f.state.sha256 && f.stateAtEnd.sha256 && f.state.sha256 !== f.stateAtEnd.sha256) report.notes.push(`${files[key]} changed during the run.`);
    }
    report.cleanup = cleanupProbeFiles(rounds.filter(r => fs.existsSync(r.path)).map(r => ({ key: r.tag, path: r.path, linkOf: a })));
    const aReadOnly = fileState(a).readOnly;
    if (aReadOnly !== true) report.notes.push(`${a} is NOT read-only after the cleanup.`);
    const kept = report.cleanup.filter(c => c.action !== 'deleted' && c.action !== 'missing');
    if (kept.length) report.notes.push(`Hard links left in ${path.dirname(a)} (${kept.map(c => `${path.basename(c.path)}: ${c.action}`).join(', ')}); `
        + 'a hard link shares the lock of A (one file), so they stay until CATIA releases A; delete them after CATIA is restarted.');
    report.notes.push('CATIA was changed by this run (C5D_Bridge copies closed, hard links loaded and closed). Restart CATIA before a probe run that needs a fresh CATIA.');
    const { lines, data } = judgeSelfClose(first.snapshot.loaded, entries, { b, a, targets, rounds, leftovers, cleanup: report.cleanup, userCount, aReadOnly });
    report.summary = { ...sequenceSummary(report, entries),
        lists: Object.fromEntries(steps.filter(s => s.info === 'sx').map(s => [s.name, parseSxList(entries.find(e => e.step === s.name)?.result)])),
        selfClose: data, verdictLines: lines, verdict: lines.join('\n') };
    if (aborted) process.exitCode = 1;
}
/** Summary of --self-close-experiment: the copies found, then as formatReloadSummary. */
function formatSelfCloseSummary(report) {
    const text = formatReloadSummary(report).split('\n');
    const lines = (report.targets ?? []).map(t => `target ${t.tag}${t.label ? ` (${t.label})` : ''}: ${JSON.stringify(t.display)} file ${JSON.stringify(t.file)}`);
    if (report.leftovers) lines.push(`left by ${report.leftovers.report}: ${report.leftovers.list.length} copies${report.leftovers.error ? ` (${report.leftovers.error})` : ''}`);
    return [...text.slice(0, 2), ...lines, ...text.slice(2)].join('\n');
}

// ---- --cycle-experiment: the production idea on a freshly restarted CATIA. One Push = a NEW path (a read-only copy or a hard link
// of the master A), loaded alone, C5D_Install C5D_Cy (the only C5D_Bridge = itself), self-identification, the Push/Pull-like read of
// a user project, guarded Close from outside. Repeated per method (copy first, then hardlink: the hard link is the riskier one, since
// the same real file under several path strings once broke CATIA's VBA with 0x800ADF09). A is never loaded through its own path.
const CY_MODULE = 'C5D_Cy';
/** The user project and module C5D_CyRead reads (never writes): the CATIA test target of this repository. */
const CY_READ_PROJECT = 'toQI';
const CY_READ_MODULE = 'C5D_BridgeTest';
/** Run order of --cycle-method both: the safer copy first. */
const CY_METHODS = ['copy', 'hardlink'];
const CY_DEFAULT_ROUNDS = 5;
const CY_MAX_ROUNDS = 20;
const CY_START = 's';
const CY_FINAL = 'z';
const CY_READ_PATTERN = /^read=ok;/;
/** Index of call:$.Close(false) in cyCloseActions. */
const CY_CLOSE_ACTION = 5;
/** Errors after which CATIA's VBA is considered unstable: the run stops (no further call into CATIA). */
const CY_FATAL = [[/0x800ADF09/i, 'Unknown VBA error 0x800ADF09'], [/0x80010105|RPC_E_SERVERFAULT/i, 'RPC_E_SERVERFAULT 0x80010105'],
    [/0x800706BA|0x800706BE|0x80010108|RPC_E_DISCONNECTED/i, 'RPC failure (CATIA gone?)']];
/** Old-name bridge (the plain copy of the default mode); never deleted by the cleanup. */
const CY_KEEP = [BRIDGE_FILE];
const CY_LEFTOVER = /^C5D_Bridge.*\.catvba$/i;

/**
 * VBA of C5D_Cy (ASCII, CRLF, On Error Resume Next, no MsgBox / Declare, no property read in an If condition; never closes, removes,
 * saves or writes anything; C5D_Install writes it into the only loaded C5D_Bridge = the project of the round):
 *   C5D_CySelf()  apc.ExecutingProject as an object (eval-probe.exe guards Name / DisplayName and closes it after ExecuteScript returned)
 *   C5D_CyWho()   "who=ok;running=..;name=..;display=<DisplayName>;mode=..;vbName=..;file=<VBProject.FileName>;comps=..;c5d=.."
 *   C5D_CyList()  every IApcProject and VBProject (as C5D_SxList)
 *   C5D_CyRead(projName, modName)  READS the first VBProject named projName: "read=ok;project=..;hits=N;file=..;saved=..;module=..;
 *                 lines=<CountOfLines>;first=<Lines(1, 1)>;savedAfter=.." | "error:noproject;count=N;err=E" | "error:nomodule;..."
 */
const CY_SOURCE = [
    'Option Explicit',
    ...sxCommonLines('Cy'),
    ...sxSelfWhoLines('Cy'),
    'Public Function C5D_CyList() As Variant',
    '    C5D_CyList = CyListText()',
    'End Function',
    'Public Function C5D_CyRead(ByVal projName As Variant, ByVal modName As Variant) As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim ide As Object',
    '    Dim vbps As Object',
    '    Dim vbp As Object',
    '    Dim hit As Object',
    '    Dim comps As Object',
    '    Dim comp As Object',
    '    Dim cm As Object',
    '    Dim pn As String',
    '    Dim mn As String',
    '    Dim nm As String',
    '    Dim s As String',
    '    Dim v As Variant',
    '    Dim e As Long',
    '    Dim hits As Long',
    '    Dim cnt As Long',
    '    Dim i As Long',
    '    Dim n As Long',
    '    pn = ""',
    '    pn = CStr(projName)',
    '    mn = ""',
    '    mn = CStr(modName)',
    '    Set apc = CyApc()',
    '    If apc Is Nothing Then',
    '        C5D_CyRead = "error:apc:" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Set ide = Nothing',
    '    Set ide = apc.VBE',
    '    Set vbps = Nothing',
    '    Set vbps = ide.VBProjects',
    '    n = -1',
    '    n = vbps.Count',
    '    hits = 0',
    '    Set hit = Nothing',
    '    For i = 1 To n',
    '        Set vbp = Nothing',
    '        Set vbp = vbps.Item(i)',
    '        nm = ""',
    '        nm = CStr(vbp.Name)',
    '        If StrComp(nm, pn, vbBinaryCompare) = 0 Then',
    '            hits = hits + 1',
    '            If hits = 1 Then Set hit = vbp',
    '        End If',
    '    Next',
    '    If hit Is Nothing Then',
    '        C5D_CyRead = "error:noproject;count=" & CStr(n) & ";err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    s = ";project=" & CyClean(pn) & ";hits=" & CStr(hits)',
    ...sxProp('CyFmt', 'file', 'hit.FileName'),
    ...sxProp('CyFmt', 'saved', 'hit.Saved'),
    '    s = s & ";module=" & CyClean(mn)',
    '    Err.Clear',
    '    Set comps = Nothing',
    '    Set comps = hit.VBComponents',
    '    Set comp = Nothing',
    '    Set comp = comps.Item(mn)',
    '    If comp Is Nothing Then',
    '        C5D_CyRead = "error:nomodule" & s & ";err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    Set cm = Nothing',
    '    Set cm = comp.CodeModule',
    '    If cm Is Nothing Then',
    '        C5D_CyRead = "error:codemodule" & s & ";err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    cnt = -1',
    '    cnt = cm.CountOfLines',
    '    s = s & ";lines=" & CStr(cnt)',
    '    Err.Clear',
    '    v = ""',
    '    If cnt > 0 Then',
    '        v = cm.Lines(1, 1)',
    '    End If',
    '    e = Err.Number',
    '    s = s & ";first=" & CyFmt(v, e)',
    ...sxProp('CyFmt', 'savedAfter', 'hit.Saved'),
    '    Err.Clear',
    '    C5D_CyRead = "read=ok" & s',
    'End Function',
    ''
].join('\r\n');

/** { methods, rounds } of --cycle-method / --cycle-rounds (values as given on the command line; missing = defaults). Throws on bad values. */
function parseCycleOptions(values = {}) {
    const method = values['--cycle-method'] ?? 'both';
    if (!['both', ...CY_METHODS].includes(method)) throw new Error(`--cycle-method must be hardlink, copy or both (got ${method})`);
    const text = values['--cycle-rounds'] ?? String(CY_DEFAULT_ROUNDS);
    const rounds = /^\d+$/.test(text) ? Number(text) : NaN;
    if (!(rounds >= 1 && rounds <= CY_MAX_ROUNDS)) throw new Error(`--cycle-rounds must be 1..${CY_MAX_ROUNDS} (got ${text})`);
    return { methods: method === 'both' ? [...CY_METHODS] : [method], rounds };
}

/**
 * The files of one run, next to A: start round s and final round z are read-only copies, the rounds are c1..cN (copy) and/or
 * h1..hN (hard link): C5D_Bridge-<8 hex>-c<yyyyMMddHHmmss>-<s|n|z>.catvba / C5D_Bridge-<8 hex>-h<yyyyMMddHHmmss>-<n>.catvba.
 */
function cycleRounds(a, runTag, methods = CY_METHODS, n = CY_DEFAULT_ROUNDS) {
    const dir = path.win32.dirname(a), base = path.win32.basename(a, '.catvba');
    const file = (letter, suffix) => path.win32.join(dir, `${base}-${letter}${runTag}-${suffix}.catvba`);
    return {
        start: { tag: CY_START, method: 'copy', path: file('c', CY_START) },
        rounds: methods.flatMap(method => Array.from({ length: n }, (_, i) => ({ tag: `${method === 'copy' ? 'c' : 'h'}${i + 1}`, method,
            path: file(method === 'copy' ? 'c' : 'h', String(i + 1)) }))),
        final: { tag: CY_FINAL, method: 'copy', path: file('c', CY_FINAL) }
    };
}

/** Actions of a cycle close: read Name / DisplayName / FileName, guard Name and DisplayName, then Close(false) and a Name read that should fail. */
const cyCloseActions = display => ['get:$.Name', 'get:$.DisplayName', 'get:$.VBProject.FileName', `guard:$.Name=${sxLiteral(BRIDGE_PROJECT)}`,
    `guard:$.DisplayName=${sxLiteral(display)}`, 'call:$.Close(false)', 'get:$.Name'];

/** Whether a parsed C5D_CyList shows exactly one C5D_Bridge (this path) and, when start is given, as many user projects as at the start. */
function cycleListOk(list, library, start = null) {
    if (!list || list.err !== '0' || list.count === null || list.vbeCount === null) return false;
    const apc = sxBridges(list), vbe = list.vbe.filter(p => p.name === BRIDGE_PROJECT);
    if (apc.length !== 1 || vbe.length !== 1 || !sameText(apc[0].display, library)) return false;
    if (!start) return true;
    const users = sxUserProjects(list);
    return users.apc.length === start.apc && users.vbe.length === start.vbe;
}

/**
 * Steps of one round r = { tag, method, path }: prepare (local: copy or hard link) load install who list read close recall lock (local).
 * start = { apc, vbe } user project counts of the start round (null for the start round itself). who / list / read / close require
 * the own who (a mismatch means C5D_Install may have written elsewhere: nothing more runs through that path).
 */
function buildCycleRound(r, start = null) {
    const p = r.path, t = r.tag;
    const call = (name, module, func, args, extra) => ({ name: `${t}-${name}`, library: p, module, func, args, round: t, ...extra });
    return [
        { name: `${t}-prepare`, local: 'prepare', method: r.method, path: p, round: t },
        call('load', BRIDGE_MODULE, 'C5D_Version', [], { expected: BRIDGE_VERSION, requires: [`${t}-prepare`] }),
        call('install', BRIDGE_MODULE, 'C5D_Install', [CY_MODULE, CY_SOURCE], { expected: 'ok', requires: [`${t}-load`] }),
        call('who', CY_MODULE, 'C5D_CyWho', [], { check: result => whoIsSelf(result, p), self: p, requires: [`${t}-install`] }),
        call('list', CY_MODULE, 'C5D_CyList', [], { check: result => cycleListOk(parseSxList(result), p, start), info: 'cy', requires: [`${t}-who`] }),
        call('read', CY_MODULE, 'C5D_CyRead', [CY_READ_PROJECT, CY_READ_MODULE], { pattern: CY_READ_PATTERN, requires: [`${t}-who`] }),
        call('close', CY_MODULE, 'C5D_CySelf', [], { actions: cyCloseActions(p), mustSucceed: [CY_CLOSE_ACTION], resultType: 'ComObject',
            closeAction: CY_CLOSE_ACTION, closeOf: p, requires: [`${t}-who`] }),
        call('recall', BRIDGE_MODULE, 'C5D_Version', [], { expectFailure: BRIDGE_VERSION, requires: [`${t}-close`] }),
        { name: `${t}-lock`, local: 'lock', path: p, round: t, after: [`${t}-prepare`] }
    ];
}
/** All step names of --cycle-experiment (start round, rounds, final round) for the given methods and round count. */
function cycleStepNames(methods = CY_METHODS, n = CY_DEFAULT_ROUNDS) {
    const plan = cycleRounds('C:\\x\\C5D_Bridge-x.catvba', '0', methods, n);
    return [plan.start, ...plan.rounds, plan.final].flatMap(r => buildCycleRound(r).map(s => s.name));
}

/** Reason text when an entry shows a CATIA-is-unstable error (the actions after the Close are ignored: "disconnected" is expected there). */
function cycleFatal(entry, step = {}) {
    if (!entry || entry.skipped) return null;
    const texts = [entry.hresult, entry.message];
    for (const a of entry.actions ?? []) if (step.closeAction === undefined || a.action <= step.closeAction) texts.push(a.hresult, a.message);
    const text = texts.filter(Boolean).join(' ');
    return CY_FATAL.find(([pattern]) => pattern.test(text))?.[1] ?? null;
}

const readOnlyOf = file => { try { return (fs.statSync(file).mode & 0o200) === 0; } catch { return null; } };
const linksOf = file => { try { return Number(fs.statSync(file, { bigint: true }).nlink); } catch { return null; } };
/** Whether x and y are the same file (hard links): same volume and file id. null when either cannot be read. */
function sameFileOf(x, y) {
    try {
        const sx = fs.statSync(x, { bigint: true }), sy = fs.statSync(y, { bigint: true });
        return sx.ino === sy.ino && sx.dev === sy.dev;
    } catch { return null; }
}

/**
 * The local file of a round (before any CATIA call through it): copy = A copied (never overwriting) and made read-only, hash = A;
 * hardlink = a new hard link of A (shares A's read-only attribute). Returns { ok, method, path, ms, error, readOnly, sameFile, links, sha256Same }.
 */
function prepareCycleFile(method, a, file, ops = {}) {
    const { copyFile = fs.copyFileSync, chmod = fs.chmodSync, linkSync = fs.linkSync, aSha256 = null } = ops;
    const out = { ok: false, method, path: file, ms: null, error: null, readOnly: null, sameFile: null, links: null, sha256Same: null };
    const started = Date.now();
    if (method === 'copy') {
        try {
            copyFile(a, file, fs.constants.COPYFILE_EXCL);
            chmod(file, 0o444);
        } catch (error) { out.error = errorText(error); }
        out.ms = Date.now() - started;
        if (!out.error) {
            out.readOnly = readOnlyOf(file);
            out.sameFile = sameFileOf(a, file);
            out.sha256Same = sha256(file) === (aSha256 ?? sha256(a));
            out.ok = out.readOnly === true && out.sha256Same && out.sameFile === false;
            if (!out.ok) out.error = [out.readOnly !== true && 'read-only not set', !out.sha256Same && 'hash differs from A', out.sameFile !== false && 'not a separate file'].filter(Boolean).join('; ');
        }
    } else if (method === 'hardlink') {
        const link = makeHardlink(a, file, { linkSync });
        out.ms = Date.now() - started;
        Object.assign(out, { error: link.error, sameFile: link.sameFile, links: link.links });
        if (link.ok) {
            out.readOnly = readOnlyOf(file);
            out.ok = link.sameFile === true && out.readOnly === true;
            if (!out.ok) out.error = [link.sameFile !== true && 'not the same file as A', out.readOnly !== true && 'read-only not set (A is not read-only?)'].filter(Boolean).join('; ');
        }
    } else throw new Error(`unknown method ${method}`);
    return out;
}

/**
 * Every C5D_Bridge*.catvba in the bridge folder with its lock state (before any CATIA call): { dir, files: [{ name, path, size, role,
 * hardlinkOfA, lock }] }. role: 'master' (A), 'old-name' (C5D_Bridge.catvba, kept) or 'leftover' (an earlier probe's file, to delete).
 */
function scanBridgeFolder(dir, a, ops = {}) {
    const { readdir = fs.readdirSync, lock = file => lockProbe(file) } = ops;
    let names;
    try { names = readdir(dir); } catch (error) { return { dir, files: [], error: errorText(error) }; }
    const aName = path.basename(a);
    const files = names.filter(name => CY_LEFTOVER.test(name)).sort((x, y) => x.localeCompare(y, 'en')).map(name => {
        const file = path.join(dir, name);
        const role = sameText(name, aName) ? 'master' : CY_KEEP.some(k => sameText(k, name)) ? 'old-name' : 'leftover';
        let size = null;
        try { size = fs.statSync(file).size; } catch { /* reported as null */ }
        return { name, path: file, size, role, hardlinkOfA: role === 'leftover' ? sameFileOf(a, file) === true : null, lock: lock(file) };
    });
    return { dir, files, error: null };
}
/**
 * Production's "clean up at the next start": deletes the leftovers of a scan that are not locked (read-only cleared first; a hard link of
 * A gets A's read-only attribute back, see cleanupProbeFiles). Returns { items, deleted, deletedBytes, hardlinks, kept, failed, aReadOnly, aLinks }.
 */
function cleanupBridgeLeftovers(scan, a, ops = {}) {
    const leftovers = scan.files.filter(f => f.role === 'leftover');
    const done = cleanupProbeFiles(leftovers.map(f => ({ key: f.name, path: f.path, linkOf: f.hardlinkOfA ? a : undefined })), ops);
    const items = done.map((c, i) => ({ ...c, size: leftovers[i].size, hardlinkOfA: leftovers[i].hardlinkOfA }));
    const deleted = items.filter(c => c.action === 'deleted');
    return { items, deleted: deleted.length, deletedBytes: deleted.reduce((sum, c) => sum + (c.size ?? 0), 0),
        hardlinks: deleted.filter(c => c.hardlinkOfA).length, kept: items.filter(c => c.action.startsWith('kept')).map(c => c.key),
        failed: items.filter(c => c.action === 'failed').map(c => `${c.key}: ${c.error}`), aReadOnly: readOnlyOf(a), aLinks: linksOf(a) };
}

// -- judging
const avgMs = list => { const n = list.filter(x => typeof x === 'number'); return n.length ? Math.round(n.reduce((x, y) => x + y, 0) / n.length) : null; };
const msText = value => value === null || value === undefined ? '-' : `${value}ms`;
const mb = bytes => `${(bytes / 1048576).toFixed(1)} MB`;
const flagOf = e => (!e || e.skipped ? null : e.match === true);
function readWord(e) {
    if (!e || e.skipped || !e.ok) return `read ${rlOutcome(e)}`;
    if (!CY_READ_PATTERN.test(e.result ?? '')) return `read ${JSON.stringify(e.result)}`;
    const f = sxFields(e.result);
    return `read ${f.project}.${f.module} lines=${f.lines} first=${JSON.stringify(f.first ?? '')} saved ${f.saved}->${f.savedAfter}${f.hits !== '1' ? ` (${f.hits} projects named ${f.project})` : ''}`;
}
/** Facts of one round from the entries (null fields = not run). users = sxUserProjects of its list. */
function cycleRoundFacts(r, entries, startUsers = null) {
    const get = k => entries.find(e => e.step === `${r.tag}-${k}`);
    const prepare = get('prepare'), load = get('load'), install = get('install'), who = get('who'), list = get('list'), read = get('read');
    const closeEntry = get('close'), recall = get('recall'), lock = get('lock');
    const parsed = listOf(list);
    const users = sxUserProjects(parsed);
    const close = sxCloseFacts(closeEntry, CY_CLOSE_ACTION);
    const whoFields = who && !who.skipped && who.ok ? parseWho(who.result) : null;
    const mine = entries.filter(e => String(e.step).startsWith(`${r.tag}-`));
    const fatal = mine.find(e => e.fatal || e.timedOut || e.crashed);
    const f = { tag: r.tag, method: r.method, path: r.path, ran: Boolean(prepare && !prepare.skipped),
        prepare: flagOf(prepare), prepareMs: prepare?.ms ?? null, prepareError: prepare?.message ?? null, prepareInfo: prepare?.prepare ?? null,
        load: flagOf(load), loadMs: load?.elapsedMs ?? null, install: flagOf(install), installMs: install?.elapsedMs ?? null,
        who: flagOf(who), whoMs: who?.elapsedMs ?? null, whoExact: whoFields ? whoFields.display === r.path : null, whoDisplay: whoFields?.display ?? null, whoFile: whoFields?.file ?? null,
        list: flagOf(list), listMs: list?.elapsedMs ?? null, counts: parsed ? { apc: parsed.count, vbe: parsed.vbeCount, bridges: sxBridges(parsed).length } : null,
        users, userDiff: startUsers && users ? userDiff(startUsers, users) : null,
        read: flagOf(read), readMs: read?.elapsedMs ?? null, readResult: read?.result ?? null,
        close, closeCallMs: closeEntry?.elapsedMs ?? null, recall: flagOf(recall), lock: lock?.locks?.[r.tag] ?? null,
        fatal: fatal ? fatal.fatal ?? (fatal.timedOut ? `timed out in ${fatal.step}` : `CATIA crashed after ${fatal.step}`) : null,
        words: { load: rlOutcome(load), install: rlOutcome(install), who: whoWord(who, r.path), list: describeSxList(list), read: readWord(read),
            recall: recall ? recallWord(recall) : '-' } };
    f.ok = f.prepare === true && f.load === true && f.install === true && f.who === true && f.list === true && f.close.ok === true && f.recall === true && !f.fatal;
    return f;
}
/**
 * Why the run stops after a round (null = go on): { kind, reason }. kind: 'fatal' (CATIA may be unstable: aborted / 0x800ADF09 /
 * RPC_E_SERVERFAULT / timeout / crash), 'restart' (another C5D_Bridge loaded), 'vbe' (no user project visible), 'round' (this round failed).
 */
function cycleStopReason(f, { start = false, aborted = null } = {}) {
    if (aborted || f.fatal) return { kind: 'fatal', reason: f.fatal ?? String(aborted).replace(/^aborted: /, '') };
    if (f.prepare !== true) return { kind: 'round', reason: `prepare failed (${f.prepareError ?? 'not run'})` };
    if (f.load !== true) return { kind: 'round', reason: `load: C5D_Version ${f.words.load}` };
    if (f.install !== true) return { kind: 'round', reason: `install: C5D_Install ${f.words.install}` };
    if (f.who !== true) {
        return { kind: start ? 'restart' : 'round', reason: `${f.words.who}: the code did not run in this path's project, so C5D_Install may have written into `
            + `another C5D_Bridge${start ? ' (one loaded before this run? CATIA を再起動してください / restart CATIA)' : ''}` };
    }
    if (start && f.counts && f.counts.bridges > 1) return { kind: 'restart', reason: `${BRIDGE_PROJECT} x${f.counts.bridges} loaded at the start: CATIA was not restarted. CATIA を再起動してください (restart CATIA)` };
    if (start && f.users && f.users.vbe.length === 0) {
        return { kind: 'vbe', reason: 'no user project visible (apc.Projects shows only C5D_Bridge before the VBA editor was opened): '
            + 'restart CATIA, open the VBA editor once (Alt+F11), then re-run' };
    }
    if (f.list !== true) return { kind: 'round', reason: `list: ${f.words.list} (expected ${BRIDGE_PROJECT} x1 = this path${start ? '' : ' and the user project count of the start'})` };
    if (f.close.ok !== true) return { kind: 'round', reason: `close: ${closeWord(f.close)} (this ${BRIDGE_PROJECT} stays loaded; the next C5D_Install would write into it)` };
    if (f.recall !== true) return { kind: 'round', reason: `same path after close: ${f.words.recall} (re-loaded: a ${BRIDGE_PROJECT} stays loaded)` };
    return null;
}
function cycleRoundLine(f, label = '') {
    const head = `round ${f.tag} (${f.method}${label}) ${baseName(f.path)}: `;
    if (!f.ran) return head + `not run (${f.notRun ?? 'not in this run'})`;
    const p = f.prepareInfo;
    const prep = f.prepare ? `${f.method === 'copy' ? 'copy' : 'hard link'} ok ${msText(f.prepareMs)}${f.method === 'hardlink' && p?.links ? ` (links ${p.links})` : ''}`
        : `prepare FAILED (${f.prepareError})`;
    const diff = f.tag === CY_START ? 'user projects listed = the start reference' : f.userDiff === null ? 'user projects not compared'
        : f.userDiff.length ? `user projects DIFFER: ${f.userDiff.join('; ')}` : 'user projects = start';
    return head + [prep, `load ${f.load ? `ok ${msText(f.loadMs)}` : f.words.load}`, `install ${f.install ? `ok ${msText(f.installMs)}` : f.words.install}`, f.words.who,
        `list ${f.words.list} (ok: ${yn(f.list)})`, f.words.read, `${closeWord(f.close)}${f.closeCallMs !== null ? ` (ExecuteScript ${f.closeCallMs}ms)` : ''}`,
        `same path after close ${f.words.recall}`, `lock after close ${lockWord(f.lock)}`, diff].join(', ') + (f.fatal ? `; FATAL: ${f.fatal}` : '');
}
function cycleMethodLine(method, facts, planned) {
    const ran = facts.filter(f => f.method === method && f.ran);
    const ok = ran.filter(f => f.ok);
    const avg = key => msText(avgMs(ok.map(key)));
    const compared = ran.filter(f => f.userDiff !== null);
    const differ = compared.filter(f => f.userDiff.length);
    const probed = ran.filter(f => f.lock && !f.lock.skipped);
    return `method ${method}: ${ok.length}/${planned} rounds ok (${ran.length} run); average of the ok rounds: prepare ${avg(f => f.prepareMs)}, load ${avg(f => f.loadMs)}, `
        + `install ${avg(f => f.installMs)}, who ${avg(f => f.whoMs)}, list ${avg(f => f.listMs)}, read ${avg(f => f.readMs)}, close ${avg(f => f.close.ms)} `
        + `(ExecuteScript ${avg(f => f.closeCallMs)}); read ok ${ran.filter(f => f.read).length}/${ran.length}; `
        + `locked after close ${probed.filter(f => f.lock.inUse === true).length}/${probed.length} probed; user projects = start in ${compared.length - differ.length}/${ran.length}`
        + `${differ.length ? `, DIFFER in ${differ.map(f => f.tag).join(', ')}` : ''}`;
}
/**
 * Verdict of --cycle-experiment: { lines, data }. ctx: { a, methods, rounds (planned per method), preCheck, preCleanup, master, start,
 * facts (rounds in run order, incl. not run ones), final, stop, cleanup, aEnd }.
 */
function judgeCycle(beforeLoaded, entries, ctx) {
    const lines = [];
    const pc = ctx.preCheck;
    if (pc) {
        const locked = pc.files.filter(f => f.lock?.inUse === true);
        lines.push(`precheck: ${pc.files.length} ${BRIDGE_PROJECT}*.catvba in ${pc.dir}${pc.error ? ` (folder not read: ${pc.error})` : ''}, `
            + `${locked.length ? `IN USE: ${locked.map(f => f.name).join(', ')}` : 'none in use'}`);
    }
    const cl = ctx.preCleanup;
    if (cl) {
        lines.push(`cleanup before run: deleted ${cl.deleted}/${cl.items.length} old files (${mb(cl.deletedBytes)}; hard links of A ${cl.hardlinks})`
            + `${cl.kept.length ? `, kept ${cl.kept.join(', ')}` : ''}${cl.failed.length ? `, FAILED ${cl.failed.join(', ')}` : ''}; A read-only afterwards ${yn(cl.aReadOnly)}, A links ${cl.aLinks ?? '-'}`);
    }
    const m = ctx.master;
    if (m) lines.push(`master A ${baseName(m.path)}: ${m.action}, read-only ${yn(m.readOnly)}, = source ${yn(m.sameAsSource)}, links ${m.links ?? '-'} (never loaded through its own path)`);
    if (ctx.start) lines.push(cycleRoundLine(ctx.start, ', start list'));
    const startUsers = ctx.start?.users ?? null;
    if (startUsers) lines.push(`user projects at start: ${startUsers.vbe.length} (${startUsers.vbe.map(p => `${p.name}[${p.comps}] saved=${p.saved}`).join(', ')})`);
    for (const f of ctx.facts ?? []) lines.push(cycleRoundLine(f));
    if (ctx.facts) for (const method of ctx.methods ?? []) lines.push(cycleMethodLine(method, ctx.facts, ctx.rounds));
    if (ctx.final) lines.push(cycleRoundLine(ctx.final, ', final list'));
    // user projects at the end: the final list, else the last list that was taken
    const lists = [ctx.final, ...[...(ctx.facts ?? [])].reverse()].filter(f => f?.users);
    const last = lists[0];
    if (startUsers && last) {
        const diff = userDiff(startUsers, last.users);
        lines.push(`user projects start vs end (list of round ${last.tag}${last === ctx.final ? ', final' : ''}): `
            + (diff.length ? `DIFFER: ${diff.join('; ')}` : `identical (${startUsers.vbe.length} projects: name, FileName, module count, Saved)`));
    } else if (ctx.start) lines.push('user projects start vs end: not compared (a list is missing)');
    const s = ctx.stop;
    lines.push(!s ? 'stop: none (every planned round and the final list ran)'
        : `stop (${s.kind}${s.kind === 'fatal' ? ': CATIA may be unstable, restart it before anything else' : ''}) ${s.after ? `after round ${s.after}` : 'before the first CATIA call'}: `
        + `${s.reason}${s.notRun?.length ? `; not run: ${s.notRun.join(', ')}` : ''}`);
    if (ctx.cleanup) {
        const kept = ctx.cleanup.filter(c => c.action !== 'deleted' && c.action !== 'missing');
        lines.push(`cleanup after run: deleted ${ctx.cleanup.filter(c => c.action === 'deleted').length}/${ctx.cleanup.length}`
            + `${kept.length ? `, kept ${kept.map(c => `${baseName(c.path)} (${c.action}${c.error ? `: ${c.error}` : ''})`).join(', ')}` : ''}`);
    }
    if (ctx.aEnd) lines.push(`A at end: unchanged ${yn(ctx.aEnd.unchanged)}, read-only ${yn(ctx.aEnd.readOnly)}, links ${ctx.aEnd.links ?? '-'}, ${lockWord(ctx.aEnd.lock)}`);
    const crashed = entries.find(e => e.crashed);
    if (crashed) lines.push(`CATIA crashed after step ${crashed.step}`);
    const timedOut = entries.find(e => e.timedOut);
    if (timedOut) lines.push(`timed out at step ${timedOut.step} (CATIA may be stopped at a dialog)`);
    if (entries.length) lines.push(describeVbscript(beforeLoaded, entries));
    return { lines, data: { start: ctx.start ?? null, rounds: ctx.facts ?? [], final: ctx.final ?? null, stop: ctx.stop ?? null } };
}

/**
 * --cycle-experiment (right after a CATIA restart, VBA editor opened once). Local first: every C5D_Bridge*.catvba of the bridge folder
 * must be free (else nothing is done: restart CATIA), the old ones are deleted, A is prepared read-only. Then the start round s (copy;
 * its list = the user projects at the start; another C5D_Bridge there stops the run), the rounds of each method (copy before
 * hardlink), the final round z (copy; the list at the end). Each round is one runSequence; no module snapshot between the CATIA
 * calls of a round (only at its lock step). Any stop reason (cycleStopReason) ends the run; a fatal one skips the final round too.
 */
async function runCycle(report, options = {}) {
    const { methods, rounds: n } = options.cycle ?? parseCycleOptions();
    const paths = hashedPaths();
    const a = paths.target, dir = path.dirname(a);
    report.cycle = { methods, rounds: n, readProject: CY_READ_PROJECT, readModule: CY_READ_MODULE };
    report.cnext = cnextProcesses();
    if (!report.cnext.error && !report.cnext.processes.length) throw new Error('CNEXT.exe is not running');
    const finish = (ctx, entries = [], beforeLoaded = false) => {
        const { lines, data } = judgeCycle(beforeLoaded, entries, { a, methods, rounds: n, ...ctx });
        report.summary = { ...(report.runs ? sequenceSummary(report, entries) : {}), cycle: data, verdictLines: lines, verdict: lines.join('\n') };
    };
    // 1. nothing in the bridge folder may be open (a closed C5D_Bridge keeps its file until CATIA exits)
    fs.mkdirSync(dir, { recursive: true });
    const preCheck = scanBridgeFolder(dir, a);
    report.preCheck = preCheck;
    const inUse = preCheck.files.filter(f => f.lock?.inUse === true);
    if (preCheck.error || inUse.length) {
        const stop = { kind: 'restart', after: null, reason: preCheck.error ? `bridge folder not readable (${preCheck.error})`
            : `${inUse.map(f => f.name).join(', ')} in use: CATIA still has a ${BRIDGE_PROJECT} of an earlier run. CATIA を再起動してください (restart CATIA; nothing was done)` };
        report.stop = stop;
        report.notes.push(stop.reason);
        finish({ preCheck, stop });
        process.exitCode = 1;
        return;
    }
    // 2. production's cleanup at the next start, 3. the master
    const preCleanup = cleanupBridgeLeftovers(preCheck, a);
    report.preCleanup = preCleanup;
    if (preCleanup.failed.length) report.notes.push(`Cleanup before the run failed for ${preCleanup.failed.join(', ')}.`);
    const prepared = prepareHashedBridge(paths.source, a);
    report.bridge = { action: prepared.action, ok: prepared.ok, locked: prepared.locked, code: prepared.code, message: prepared.message,
        source: prepared.source, before: prepared.before, readOnly: prepared.readOnly, target: prepared.after, lock: lockProbe(a), targetAtEnd: null, lockAtEnd: null };
    if (prepared.locked) throw new Error(`${a} differs from the source and cannot be replaced (${prepared.code}); restart CATIA and re-run.`);
    if (!prepared.ok) throw new Error(`Preparing ${a} failed: ${prepared.message}`);
    const aSha256 = prepared.after.sha256;
    const master = { path: a, action: prepared.action, readOnly: prepared.after.readOnly, sameAsSource: aSha256 === prepared.source.sha256, links: linksOf(a) };
    const runTag = timestamp().replace('-', '');
    const plan = cycleRounds(a, runTag, methods, n);
    const all = [plan.start, ...plan.rounds, plan.final];
    for (const r of all) if (fs.existsSync(r.path)) throw new Error(`${r.path} already exists`);
    report.plan = all;
    report.files = { a: { role: 'A (master, read-only; copied / hard-linked, never loaded through its own path)', state: prepared.after, lock: report.bridge.lock } };
    report.library = '(per step, see each run)';
    report.libraryPerStep = true;
    report.exe = buildProbe(ROOT);
    const first = trySnapshot(report, 'before');
    if (first.error) throw new Error(`module snapshot failed: ${first.error}`);
    checkCnext(report, first.snapshot);
    const pids = first.snapshot.cnextPids;
    report.cnextPids = pids;
    report.vbscriptBefore = first.snapshot.loaded;
    if (first.snapshot.loaded) report.notes.push('vbscript.dll was already loaded before the first step (CATIA not freshly restarted?).');
    const local = async step => {
        if (step.local === 'prepare') {
            const p = prepareCycleFile(step.method, a, step.path, { aSha256 });
            report.files[step.round] = { role: `${step.method === 'copy' ? 'read-only copy' : 'hard link'} of A for round ${step.round}`, prepare: p,
                state: p.ok ? fileState(step.path) : null, lock: p.ok ? lockProbe(step.path) : null };
            return { run: { local: true, kind: step.method === 'copy' ? 'copy' : 'link', ...p, match: p.ok, aReadOnly: readOnlyOf(a), durationMs: p.ms, message: p.error },
                entry: { ok: p.ok, match: p.ok, ms: p.ms, path: step.path, message: p.error, prepare: p } };
        }
        if (step.local !== 'lock') throw new Error(`unknown local step ${step.local}`);
        const started = Date.now();
        const lock = lockProbe(step.path);
        if (report.files[step.round]) report.files[step.round].lockAfterClose = lock;
        return { run: { local: true, kind: 'locks', ok: true, match: true, locks: { [step.round]: lock }, durationMs: Date.now() - started },
            entry: { ok: true, match: true, locks: { [step.round]: lock } } };
    };
    const seqOptions = { snapshot: step => step.local === 'lock', fatal: cycleFatal };
    const steps = [];
    let state = null;
    const runRound = async (r, startCounts, startUsers) => {
        const roundSteps = buildCycleRound(r, startCounts);
        steps.push(...roundSteps);
        state = await runSequence(report, roundSteps, s => s.library, pids, local, state, seqOptions);
        for (const entry of state.entries) {
            const run = report.runs.find(x => x.step === entry.step);
            if (run?.result && entry.elapsedMs === undefined) Object.assign(entry, { elapsedMs: run.result.elapsedMs ?? null, processMs: run.durationMs ?? null });
        }
        return cycleRoundFacts(r, state.entries, startUsers);
    };
    let stop = null;
    const notRun = (r, why) => ({ ...cycleRoundFacts(r, []), ran: false, notRun: why });
    // 4. start round
    const start = await runRound(plan.start, null, null);
    const startStop = cycleStopReason(start, { start: true, aborted: state.aborted });
    if (startStop) stop = { ...startStop, after: plan.start.tag };
    const startUsers = start.users;
    const startCounts = startUsers ? { apc: startUsers.apc.length, vbe: startUsers.vbe.length } : null;
    // 5. rounds (copy first), 6. final list
    const facts = [];
    for (const r of plan.rounds) {
        if (stop) { facts.push(notRun(r, `stopped after round ${stop.after}`)); continue; }
        const f = await runRound(r, startCounts, startUsers);
        facts.push(f);
        const why = cycleStopReason(f, { aborted: state.aborted });
        if (why) stop = { ...why, after: r.tag };
    }
    let final;
    if (stop) final = notRun(plan.final, `stopped after round ${stop.after}`);
    else {
        final = await runRound(plan.final, startCounts, startUsers);
        const why = cycleStopReason(final, { aborted: state.aborted });
        if (why) stop = { ...why, after: plan.final.tag };
    }
    if (stop) {
        stop.notRun = [...facts, final].filter(f => !f.ran).map(f => f.tag);
        report.notes.push(`Stopped (${stop.kind}) after round ${stop.after}: ${stop.reason}`);
        if (stop.kind === 'fatal') report.notes.push('CATIA may be unstable: restart CATIA before anything else (do not save the C5D_Bridge projects).');
    }
    report.stop = stop;
    report.steps = steps.map(s => s.name);
    const { entries } = state;
    // 7. end: file facts, then the files of this run that CATIA does not hold any more are deleted (normally none: closed files stay locked)
    for (const [key, f] of Object.entries(report.files)) {
        const file = key === 'a' ? a : all.find(r => r.tag === key)?.path;
        if (file) Object.assign(f, { stateAtEnd: fileState(file), lockAtEnd: lockProbe(file) });
    }
    report.bridge.targetAtEnd = report.files.a.stateAtEnd;
    report.bridge.lockAtEnd = report.files.a.lockAtEnd;
    report.cleanup = cleanupProbeFiles(all.filter(r => fs.existsSync(r.path)).map(r => ({ key: r.tag, path: r.path, linkOf: r.method === 'hardlink' ? a : undefined })));
    const kept = report.cleanup.filter(c => c.action !== 'deleted' && c.action !== 'missing');
    if (kept.length) report.notes.push(`Files of this run left in ${dir} (${kept.map(c => `${path.basename(c.path)}: ${c.action}`).join(', ')}); `
        + 'CATIA holds them until it exits (a hard link holds A). The next --cycle-experiment deletes them after a CATIA restart.');
    const aEnd = { unchanged: fileState(a).sha256 === aSha256, readOnly: readOnlyOf(a), links: linksOf(a), lock: lockProbe(a) };
    if (!aEnd.unchanged) report.notes.push(`${a} changed during the run.`);
    if (aEnd.readOnly !== true) report.notes.push(`${a} is NOT read-only at the end.`);
    report.notes.push('CATIA was changed by this run (C5D_Bridge copies loaded and closed). Restart CATIA before the next probe run.');
    finish({ preCheck, preCleanup, master, start, facts, final, stop, cleanup: report.cleanup, aEnd }, entries, first.snapshot.loaded);
    report.summary.lists = Object.fromEntries(steps.filter(s => s.info === 'cy').map(s => [s.name, parseSxList(entries.find(e => e.step === s.name)?.result)]));
    if (stop || state.aborted) process.exitCode = 1;
}
/** Summary of --cycle-experiment: pre-check and cleanup before the run, then as formatReloadSummary. */
function formatCycleSummary(report) {
    const text = formatReloadSummary(report).split('\n');
    const lines = [];
    if (report.cycle) lines.push(`cycle: methods ${report.cycle.methods.join(' -> ')}, ${report.cycle.rounds} rounds each, read ${report.cycle.readProject}.${report.cycle.readModule}`);
    for (const f of report.preCheck?.files ?? []) lines.push(`precheck ${f.role} ${f.name}${f.hardlinkOfA ? ' (hard link of A)' : ''}: ${lockWord(f.lock)}`);
    for (const c of report.preCleanup?.items ?? []) lines.push(`cleanup before run ${c.key}: ${c.action}${c.error ? ` (${c.error})` : ''}`);
    return [...text.slice(0, 2), ...lines, ...text.slice(2)].join('\n');
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
            } else if (run.kind === 'copy') {
                lines.push(`  ${run.step}: copy ${run.ok ? 'ok' : `FAILED (${run.error})`} readOnly=${yn(run.readOnly)} sha256=A ${yn(run.sha256Same)} `
                    + `separate file=${yn(run.sameFile === null ? null : !run.sameFile)} ${run.durationMs}ms  ${tail}`);
            } else if (run.kind === 'link') {
                lines.push(`  ${run.step}: hard link ${run.ok ? 'ok' : `FAILED (${run.error})`} sameFile=${yn(run.sameFile)} links=${run.links ?? '-'} `
                    + `A readOnly=${yn(run.aReadOnly)} ${run.durationMs}ms  ${tail}`);
            } else if (run.kind === 'locks') {
                lines.push(`  ${run.step}: ${Object.entries(run.locks ?? {}).map(([key, l]) => `${key} ${lockWord(l)}`).join(', ') || '(no files)'}  ${tail}`);
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
                if (report.libraryPerStep) lines.push(`    library ${run.library}`);
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
    if (report.summary?.verdictLines) lines.push('', 'verdict:', ...report.summary.verdictLines.map(line => `  ${line}`));
    else if (report.summary?.verdict) lines.push('', `verdict: ${report.summary.verdict}`);
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

/** Summary of --reload-experiment: the probe files and path variants, then the steps and verdict as in formatUnloadSummary. */
function formatReloadSummary(report) {
    const lines = [];
    for (const [key, f] of Object.entries(report.files ?? {})) {
        lines.push(`file ${key} (${f.role}): ${f.state?.path ?? f.hardlink?.path ?? '-'}`);
        if (f.hardlink && !f.hardlink.ok) lines.push(`  hard link NOT created: ${f.hardlink.error}`);
        if (f.state) lines.push(`  start: ${factsLine(f.state)}, ${lockWord(f.lock)}${f.prepared ? `, prepared ${f.prepared.action}` : ''}`);
        if (f.stateAtEnd) lines.push(`  end:   ${f.stateAtEnd.exists ? factsLine(f.stateAtEnd) : 'missing'}, ${lockWord(f.lockAtEnd)}`);
    }
    if (report.shortPath) lines.push(`8.3 short path of A: ${report.shortPath.path ?? `none (${report.shortPath.error})`}`);
    for (const v of report.variants ?? []) {
        lines.push(`variant ${v.tag} ${v.label}: ${JSON.stringify(v.library)}${v.disabled ? ` SKIPPED (${v.disabled})` : ''}${v.exists === false ? ' (node: does not exist)' : ''}`);
    }
    for (const c of report.cleanup ?? []) lines.push(`cleanup ${c.key}: ${c.action}${c.error ? ` (${c.error})` : ''} ${c.path}`);
    const text = formatUnloadSummary(report).split('\n');
    // mode / report lines first, then the file facts
    return [...text.slice(0, 2), ...lines, ...text.slice(2)].join('\n');
}

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
    'remove-experiment': 'probe-bridge-remove', 'remove-experiment-risky': 'probe-bridge-remove-risky', 'reload-experiment': 'probe-bridge-reload',
    'self-close-experiment': 'probe-bridge-selfclose', 'cycle-experiment': 'probe-bridge-cycle' };
/** Runner, summary formatter and report prefix of a parsed mode (not 'help'). */
function dispatch(options) {
    switch (options.mode) {
        case 'check-only': return { run: runCheckOnly, format: formatSummary, prefix: 'probe-bridge' };
        case 'unload-experiment': case 'unload-cleanup': return { run: runUnload, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'readonly-fresh': return { run: runReadonlyFresh, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'readonly-fresh-check': return { run: runReadonlyFreshCheck, format: formatReadonlyCheck, prefix: REPORT_PREFIXES[options.mode] };
        case 'remove-experiment': case 'remove-experiment-risky': return { run: runRemove, format: formatUnloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'reload-experiment': return { run: runReload, format: formatReloadSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'self-close-experiment': return { run: runSelfClose, format: formatSelfCloseSummary, prefix: REPORT_PREFIXES[options.mode] };
        case 'cycle-experiment': return { run: runCycle, format: formatCycleSummary, prefix: REPORT_PREFIXES[options.mode] };
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
    TIMER_LOG_TOKEN, materializeSteps, skipReason, parseInfo, describeInfo, judgeRemove, runRemove, runSequence,
    reloadSource, buildReloadSteps, reloadStepNames, pathVariants, shortPathOf, makeHardlink, cleanupProbeFiles, rlBridges, describeRlInfo,
    judgeReload, runReload, formatReloadSummary, RL_ROUNDS, RL_FINAL, RL_CLOSE_ACTIONS, rlModule, rlFunc,
    SX_MODULE, SXB_MODULE, SX_SOURCE, SXB_SOURCE, SX_ROUNDS, SX_FINAL, SX_CLOSE_ACTION, SX_MAX_TARGETS, sxCloseActions, sxLiteral, parseSxList, parseWho,
    whoIsSelf, sxUserProjects, userDiff, buildSelfCloseStart, buildSelfCloseTargets, buildSelfCloseB, buildSelfCloseRounds, selfCloseRounds,
    selfCloseTargets, selfCloseStepNames, reloadLeftovers, sxCloseFacts, judgeSelfClose, runSelfClose, formatSelfCloseSummary,
    CY_MODULE, CY_SOURCE, CY_READ_PROJECT, CY_READ_MODULE, CY_METHODS, CY_DEFAULT_ROUNDS, CY_MAX_ROUNDS, CY_START, CY_FINAL, CY_CLOSE_ACTION,
    parseCycleOptions, cycleRounds, cyCloseActions, cycleListOk, buildCycleRound, cycleStepNames, cycleFatal, prepareCycleFile, scanBridgeFolder,
    cleanupBridgeLeftovers, cycleRoundFacts, cycleStopReason, judgeCycle, runCycle, formatCycleSummary };
