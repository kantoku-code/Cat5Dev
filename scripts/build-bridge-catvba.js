// Rebuilds resources/bridge/C5D_Bridge.catvba as bridge v2 in a running CATIA (developer tool, not shipped).
// Contract and manual fallback: docs/dev/bridge-contract.md (section 11). Every CATIA call is one
// SystemService.ExecuteScript(<full path>, 2, <module>, <function>, <args>) through temp/eval-probe.exe
// (tools/catia-run/EvalProbe.cs: vbacall / vbacall-obj, b64 tokens, 60 s timeout per call, see scripts/probe-bridge.js).
//
// Usage:
//   node scripts/build-bridge-catvba.js [--dry-run] [--base <file.catvba>] [--save-method save|close] [--no-install] [--allow-orphans]
//       Automatic build (default base: resources/bridge/C5D_Bridge.catvba, v1 or v2):
//         precheck   no C5D_Bridge*.catvba of %LOCALAPPDATA%\Cat5Dev\bridge / bridge-build (and the base / target) may be in use when
//                    the base is v1 (its C5D_Install writes into the FIRST project named C5D_Bridge), else CATIA must be restarted
//         prepare    writable copy of the base: %LOCALAPPDATA%\Cat5Dev\bridge-build\C5D_Bridge-build-<stamp>-w.catvba (never the base)
//         load / install-builder / builder-who   auto-load, write the temporary module C5D_Builder (via the base's C5D_Install), check
//                    that ExecutingProject is this copy and that it is the only project named C5D_Bridge
//         put-push / put-runtime   C5D_Builder replaces C5D_Bridge_Push with resources/bridge/C5D_Bridge_Push.bas (v2) and writes
//                    C5D_Bridge_Runtime (resources/bridge/C5D_Bridge_Runtime.bas) into ExecutingProject.VBProject
//         version-v2 / runtime-version   the rewritten modules answer
//         remove-builder   from outside (after ExecuteScript returned): VBComponents.Remove(C5D_Builder) on C5D_Self (guarded)
//         who-v2     the components are exactly C5D_Bridge_Push and C5D_Bridge_Runtime
//         save       IApcProject.Save(True) from outside (guarded)   [--save-method close: skipped, the close uses Close(True)]
//         close      IApcProject.Close(False) from outside (guarded; also runs after a failure, unless CATIA is gone or at a dialog)
//         collect    copy of the saved work file to temp/bridge-build-<stamp>/C5D_Bridge.catvba (a file CATIA still holds may not be
//                    readable: then exit 3 and "--verify <work file> --install" after quitting CATIA)
//         inspect    offline (scripts/lib/catvba-read.js): project C5D_Bridge, modules exactly C5D_Bridge_Push = the .bas and
//                    C5D_Bridge_Runtime = the .bas, and no orphan storages (leftover form storages of other projects)
//         verify     a NEW read-only copy C5D_Bridge-build-<stamp>-v.catvba is loaded: C5D_Version = c5d-bridge:2, C5D_Who, C5D_List,
//                    C5D_RT_Version, C5D_RT_ListProjects, C5D_Install + a call of the installed probe module (Unicode round trip),
//                    C5D_Install(C5D_Bridge_Push) = "error:self", then a guarded Close(False)
//         install    resources/bridge/C5D_Bridge.catvba is copied to temp/C5D_Bridge.catvba.bak-<stamp> and replaced by the collected
//                    file (skipped with --no-install; refused while the inspect found orphan storages, unless --allow-orphans)
//   node scripts/build-bridge-catvba.js --verify <file.catvba> [--dry-run] [--install] [--allow-orphans]
//       inspect + verify (+ install) of a file made by hand (docs/dev/bridge-contract.md 11.3); never writes into <file>
//   node scripts/build-bridge-catvba.js --inspect <file.catvba>     offline only (no CATIA); exit 0 = a valid v2 bridge without orphans
//   node scripts/build-bridge-catvba.js --export-import <dir>       writes <dir>/C5D_Bridge_Push.bas and C5D_Bridge_Runtime.bas with an
//       "Attribute VB_Name" line, for File > Import in the VBA editor (the sources in resources/bridge have no Attribute line)
// --dry-run prints the plan (paths, every call with its decoded arguments and eval-probe.exe tokens) as JSON: no CATIA call, no file
// written, eval-probe.exe not built. Each real run writes temp/build-bridge-catvba-<yyyyMMdd-HHmmss>.json and prints a summary.
// CATIA state for a build: running, freshly restarted when the base is v1 (no Push since the start), nothing unsaved that matters, no
// dialog open. The VBA editor does not need to be open. User projects are never written (only read by C5D_RT_ListProjects).
// Exit codes: 0 = done, 1 = a step failed (see the report), 2 = bad arguments, 3 = stopped, resumable (see the printed command).
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { readCatvba, codeOf } = require('./lib/catvba-read');

const ROOT = path.resolve(__dirname, '..');
const BRIDGE_PROJECT = 'C5D_Bridge';
const PUSH_MODULE = 'C5D_Bridge_Push';
const RUNTIME_MODULE = 'C5D_Bridge_Runtime';
const BUILDER_MODULE = 'C5D_Builder';
const PROBE_MODULE = 'C5D_BuildProbe';
const V1 = 'c5d-bridge:1';
const V2 = 'c5d-bridge:2';
/** Japanese, U+2713 and U+1F642 (a surrogate pair, not in Shift-JIS): must come back unchanged. */
const PROBE_ARGUMENT = '日本語テスト✓🙂';
const LEFTOVER = /^C5D_Bridge.*\.catvba$/i;
/** Storages / streams of a VBA project storage that are not modules (MS-OVBA 2.2). */
const PROJECT_ENTRIES = new Set(['VBA', 'PROJECT', 'PROJECTwm', 'PROJECTlk']);

/**
 * VBA of the temporary module C5D_Builder (same rules as the bridge sources: ASCII, CRLF, no Attribute line, On Error Resume Next,
 * no MsgBox / Declare, no property read in an If / loop condition; never closes, saves or removes anything):
 *   C5D_BuildSelf()  apc.ExecutingProject (for the guarded Close when the rewrite of C5D_Bridge_Push failed)
 *   C5D_BuildWho()   "ok" LF <Name> TAB <DisplayName> TAB <VBProject.FileName> TAB <IApcProjects named C5D_Bridge> TAB <apc.Projects.Count>
 *   C5D_BuildPut(moduleName, source)  writes the standard module moduleName of ExecutingProject.VBProject (Name must be C5D_Bridge;
 *                    created when missing; an existing one must be a standard module) -> "ok:<CountOfLines>" | "error:self" | "error:..."
 */
const BUILDER_SOURCE = [
    'Option Explicit',
    "' C5D_Builder - temporary module of scripts/build-bridge-catvba.js (developer tool); removed before the project is saved.",
    '',
    'Private Const BRIDGE_PROJECT As String = "C5D_Bridge"',
    'Private Const BUILDER_MODULE As String = "C5D_Builder"',
    '',
    'Private Function BdErr(ByVal kind As String, ByVal number As Long, ByVal description As String) As String',
    '    On Error Resume Next',
    '    BdErr = "error:" & kind & ":" & CStr(number) & ":" & description',
    'End Function',
    '',
    'Private Function BdEsc(ByVal t As String) As String',
    '    On Error Resume Next',
    '    t = Replace(t, "%", "%25")',
    '    t = Replace(t, vbTab, "%09")',
    '    t = Replace(t, vbLf, "%0A")',
    '    t = Replace(t, vbCr, "%0D")',
    '    BdEsc = t',
    'End Function',
    '',
    'Private Function BdExec() As Object',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Set BdExec = Nothing',
    '    Set apc = Nothing',
    '    Set apc = CreateObject("MSAPC.Apc.7.1")',
    '    If apc Is Nothing Then',
    '        Err.Clear',
    '        Set apc = CreateObject("MSAPC.Apc")',
    '    End If',
    '    If apc Is Nothing Then',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    Set BdExec = apc.ExecutingProject',
    'End Function',
    '',
    'Public Function C5D_BuildSelf() As Variant',
    '    On Error Resume Next',
    '    Dim ex As Object',
    '    Set ex = BdExec()',
    '    If ex Is Nothing Then',
    '        C5D_BuildSelf = BdErr("self.exec", 424, "No executing project")',
    '        Exit Function',
    '    End If',
    '    Set C5D_BuildSelf = ex',
    'End Function',
    '',
    'Public Function C5D_BuildWho() As Variant',
    '    On Error Resume Next',
    '    Dim ex As Object',
    '    Dim aps As Object',
    '    Dim ap As Object',
    '    Dim pn As String',
    '    Dim dn As String',
    '    Dim fn As String',
    '    Dim nm As String',
    '    Dim e As Long',
    '    Dim d As String',
    '    Dim n As Long',
    '    Dim i As Long',
    '    Dim hits As Long',
    '    Set ex = BdExec()',
    '    If ex Is Nothing Then',
    '        C5D_BuildWho = BdErr("who.exec", 424, "No executing project")',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    pn = ""',
    '    pn = CStr(ex.Name)',
    '    dn = ""',
    '    dn = CStr(ex.DisplayName)',
    '    fn = ""',
    '    fn = CStr(ex.VBProject.FileName)',
    '    Set aps = Nothing',
    '    Set aps = ex.Apc.Projects',
    '    n = -1',
    '    n = aps.Count',
    '    e = Err.Number',
    '    d = Err.Description',
    '    If e <> 0 Then',
    '        C5D_BuildWho = BdErr("who.project", e, d)',
    '        Exit Function',
    '    End If',
    '    hits = 0',
    '    For i = 1 To n',
    '        Err.Clear',
    '        Set ap = Nothing',
    '        Set ap = aps.Item(i)',
    '        nm = ""',
    '        nm = CStr(ap.Name)',
    '        If StrComp(nm, BRIDGE_PROJECT, vbBinaryCompare) = 0 Then hits = hits + 1',
    '    Next',
    '    Err.Clear',
    '    C5D_BuildWho = "ok" & vbLf & BdEsc(pn) & vbTab & BdEsc(dn) & vbTab & BdEsc(fn) & vbTab & CStr(hits) & vbTab & CStr(n)',
    'End Function',
    '',
    'Public Function C5D_BuildPut(ByVal moduleName As Variant, ByVal source As Variant) As Variant',
    '    On Error Resume Next',
    '    Dim ex As Object',
    '    Dim comps As Object',
    '    Dim comp As Object',
    '    Dim hit As Object',
    '    Dim cm As Object',
    '    Dim mn As String',
    '    Dim src As String',
    '    Dim pn As String',
    '    Dim cn As String',
    '    Dim e As Long',
    '    Dim d As String',
    '    Dim t As Long',
    '    Dim n As Long',
    '    Dim i As Long',
    '    mn = ""',
    '    mn = CStr(moduleName)',
    '    src = ""',
    '    src = CStr(source)',
    '    If StrComp(mn, BUILDER_MODULE, vbTextCompare) = 0 Then',
    '        C5D_BuildPut = "error:self"',
    '        Exit Function',
    '    End If',
    '    Set ex = BdExec()',
    '    If ex Is Nothing Then',
    '        C5D_BuildPut = BdErr("put.exec", 424, "No executing project")',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    pn = ""',
    '    pn = CStr(ex.Name)',
    '    If StrComp(pn, BRIDGE_PROJECT, vbBinaryCompare) <> 0 Then',
    '        C5D_BuildPut = BdErr("put.project", 5, "Executing project is not " & BRIDGE_PROJECT)',
    '        Exit Function',
    '    End If',
    '    Err.Clear',
    '    Set comps = Nothing',
    '    Set comps = ex.VBProject.VBComponents',
    '    n = -1',
    '    n = comps.Count',
    '    e = Err.Number',
    '    d = Err.Description',
    '    If e <> 0 Then',
    '        C5D_BuildPut = BdErr("put.components", e, d)',
    '        Exit Function',
    '    End If',
    '    Set hit = Nothing',
    '    For i = 1 To n',
    '        Err.Clear',
    '        Set comp = Nothing',
    '        Set comp = comps.Item(i)',
    '        cn = ""',
    '        cn = CStr(comp.Name)',
    '        If StrComp(cn, mn, vbTextCompare) = 0 Then',
    '            Set hit = comp',
    '            Exit For',
    '        End If',
    '    Next',
    '    If hit Is Nothing Then',
    '        Err.Clear',
    '        Set hit = comps.Add(1)',
    '        hit.Name = mn',
    '        e = Err.Number',
    '        d = Err.Description',
    '        If e <> 0 Then',
    '            C5D_BuildPut = BdErr("put.add", e, d)',
    '            Exit Function',
    '        End If',
    '    Else',
    '        Err.Clear',
    '        t = -1',
    '        t = hit.Type',
    '        If t <> 1 Then',
    '            C5D_BuildPut = BdErr("put.type", 5, "Existing module is not a standard module")',
    '            Exit Function',
    '        End If',
    '    End If',
    '    Err.Clear',
    '    Set cm = Nothing',
    '    Set cm = hit.CodeModule',
    '    n = -1',
    '    n = cm.CountOfLines',
    '    If n > 0 Then',
    '        cm.DeleteLines 1, n',
    '    End If',
    '    If Len(src) > 0 Then',
    '        cm.AddFromString src',
    '    End If',
    '    e = Err.Number',
    '    d = Err.Description',
    '    If e <> 0 Then',
    '        C5D_BuildPut = BdErr("put.write", e, d)',
    '        Exit Function',
    '    End If',
    '    n = -1',
    '    n = cm.CountOfLines',
    '    C5D_BuildPut = "ok:" & CStr(n)',
    'End Function',
    ''
].join('\r\n');

/** VBA of the probe module the verify writes through the new C5D_Install (proves the install target is ExecutingProject). */
const PROBE_SOURCE = [
    'Option Explicit',
    'Public Function C5D_BuildProbeEcho(ByVal a As Variant) As Variant',
    '    On Error Resume Next',
    '    C5D_BuildProbeEcho = "build-probe:" & CStr(a)',
    'End Function',
    ''
].join('\r\n');

const USAGE = [
    'Usage: node scripts/build-bridge-catvba.js [--dry-run] [--base <file.catvba>] [--save-method save|close] [--no-install] [--allow-orphans]',
    '       node scripts/build-bridge-catvba.js --verify <file.catvba> [--dry-run] [--install] [--allow-orphans]',
    '       node scripts/build-bridge-catvba.js --inspect <file.catvba>',
    '       node scripts/build-bridge-catvba.js --export-import <dir>',
    '  See the header of this file and docs/dev/bridge-contract.md (section 11).'
].join('\n');

/**
 * { mode: 'build' | 'verify' | 'inspect' | 'export-import' | 'help', dryRun, base, file, dir, saveMethod, install, allowOrphans }.
 * Throws on bad arguments.
 */
function parseArgs(args) {
    if (args.some(a => a === '--help' || a === '-h')) return { mode: 'help' };
    const out = { mode: 'build', dryRun: false, base: null, file: null, dir: null, saveMethod: 'save', install: null, allowOrphans: false };
    const seen = new Set();
    const value = (name, i) => {
        if (i + 1 >= args.length || args[i + 1].startsWith('--')) throw new Error(`${name} needs a value\n${USAGE}`);
        return args[i + 1];
    };
    let mode = null;
    for (let i = 0; i < args.length; i++) {
        const a = args[i];
        if (seen.has(a)) throw new Error(`${a} given twice\n${USAGE}`);
        seen.add(a);
        if (a === '--dry-run') out.dryRun = true;
        else if (a === '--no-install') out.install = false;
        else if (a === '--install') out.install = true;
        else if (a === '--allow-orphans') out.allowOrphans = true;
        else if (a === '--base') out.base = value(a, i++);
        else if (a === '--save-method') {
            out.saveMethod = value(a, i++);
            if (!['save', 'close'].includes(out.saveMethod)) throw new Error(`--save-method must be save or close (got ${out.saveMethod})`);
        } else if (['--verify', '--inspect', '--export-import'].includes(a)) {
            if (mode) throw new Error(`only one of --verify / --inspect / --export-import\n${USAGE}`);
            mode = a.slice(2);
            if (mode === 'export-import') out.dir = value(a, i++); else out.file = value(a, i++);
        } else throw new Error(`Unknown option: ${a}\n${USAGE}`);
    }
    out.mode = mode ?? 'build';
    if (seen.has('--no-install') && seen.has('--install')) throw new Error('--install and --no-install exclude each other');
    if (out.mode !== 'build' && (out.base || seen.has('--save-method') || seen.has('--no-install'))) {
        throw new Error(`--base / --save-method / --no-install only for the build\n${USAGE}`);
    }
    if (out.mode === 'build' && seen.has('--install')) throw new Error('the build installs by default (use --no-install to skip)');
    if (['inspect', 'export-import'].includes(out.mode) && (out.dryRun || out.allowOrphans || seen.has('--install'))) {
        throw new Error(`--${out.mode} takes no other option\n${USAGE}`);
    }
    out.install = out.mode === 'build' ? out.install !== false : out.mode === 'verify' ? out.install === true : false;
    return out;
}

/** The two bridge sources and their versions (resources/bridge/*.bas). */
function readSources(root = ROOT) {
    const read = name => fs.readFileSync(path.join(root, 'resources', 'bridge', name), 'utf8');
    const push = read(`${PUSH_MODULE}.bas`), runtime = read(`${RUNTIME_MODULE}.bas`);
    const pushVersion = /Private Const BRIDGE_VERSION As String = "([^"]+)"/.exec(push)?.[1] ?? null;
    const runtimeVersion = /Private Const RT_VERSION As String = "([^"]+)"/.exec(runtime)?.[1] ?? null;
    if (pushVersion !== V2) throw new Error(`${PUSH_MODULE}.bas must declare BRIDGE_VERSION = "${V2}" (got ${pushVersion})`);
    if (!/^c5d-runtime:\d+$/.test(runtimeVersion ?? '')) throw new Error(`${RUNTIME_MODULE}.bas must declare RT_VERSION = "c5d-runtime:<n>"`);
    return { push, runtime, pushVersion, runtimeVersion };
}

const pad = n => String(n).padStart(2, '0');
const stampOf = (d = new Date()) => `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;

/** Every path of one run (nothing is created here). */
function createPaths({ env = process.env, root = ROOT, stamp = stampOf(), base = null, file = null } = {}) {
    if (!env.LOCALAPPDATA) throw new Error('%LOCALAPPDATA% is not set');
    const local = path.join(env.LOCALAPPDATA, 'Cat5Dev');
    const buildDir = path.join(local, 'bridge-build');
    const temp = path.join(root, 'temp');
    const target = path.join(root, 'resources', 'bridge', 'C5D_Bridge.catvba');
    return {
        root, stamp, target, bridgeDir: path.join(local, 'bridge'), buildDir,
        base: path.resolve(base ?? target), file: file ? path.resolve(file) : null,
        work: path.join(buildDir, `C5D_Bridge-build-${stamp}-w.catvba`),
        verify: path.join(buildDir, `C5D_Bridge-build-${stamp}-v.catvba`),
        collectDir: path.join(temp, `bridge-build-${stamp}`),
        collected: path.join(temp, `bridge-build-${stamp}`, 'C5D_Bridge.catvba'),
        backup: path.join(temp, `C5D_Bridge.catvba.bak-${stamp}`),
        report: path.join(temp, `build-bridge-catvba-${stamp.slice(0, 8)}-${stamp.slice(8)}.json`)
    };
}

/** Lines of a module as compared here: trailing blanks of each line and trailing empty lines ignored (VBE may drop them). */
const comparable = text => String(text).replace(/\r\n?/g, '\n').split('\n').map(l => l.replace(/[ \t]+$/, '')).join('\n').replace(/\n+$/, '');

/**
 * Offline facts of a .catvba (no CATIA): { ok, problems, projectName, modules: [{ name, type, chars, same }], version, runtimeVersion,
 * orphans: [storage names under the VBA project storage that belong to no module] }. sources (readSources) adds the comparison.
 */
function inspectCatvba(file, sources = null) {
    const facts = { file, ok: false, problems: [], projectName: null, modules: [], version: null, runtimeVersion: null, orphans: [] };
    let r;
    try { r = readCatvba(file); } catch (error) { facts.problems.push(String(error.message ?? error)); return facts; }
    facts.projectName = r.projectName;
    const push = r.modules.find(m => m.name === PUSH_MODULE), runtime = r.modules.find(m => m.name === RUNTIME_MODULE);
    facts.version = push ? /C5D_Version = "([^"]+)"|BRIDGE_VERSION As String = "([^"]+)"/.exec(push.source)?.slice(1).find(Boolean) ?? null : null;
    facts.runtimeVersion = runtime ? /RT_VERSION As String = "([^"]+)"/.exec(runtime.source)?.[1] ?? null : null;
    const expected = sources ? { [PUSH_MODULE]: sources.push, [RUNTIME_MODULE]: sources.runtime } : null;
    facts.modules = r.modules.map(m => ({ name: m.name, type: m.type, chars: m.source.length,
        same: expected && expected[m.name] !== undefined ? comparable(codeOf(m.source)) === comparable(expected[m.name]) : null }));
    // storages directly under the project storage (".../_VBA_Project/<name>/...") that are neither VBA/PROJECT* nor a module
    const projectPrefix = r.streams.find(n => /\/VBA\/dir$/i.test(n))?.replace(/VBA\/dir$/i, '') ?? '';
    const children = new Set(r.streams.filter(n => n.startsWith(projectPrefix)).map(n => n.slice(projectPrefix.length).split('/')[0]));
    const moduleNames = new Set(r.modules.map(m => m.name));
    facts.orphans = [...children].filter(c => !PROJECT_ENTRIES.has(c) && !moduleNames.has(c)).sort();
    if (r.projectName !== BRIDGE_PROJECT) facts.problems.push(`project name is ${r.projectName}, expected ${BRIDGE_PROJECT}`);
    const names = r.modules.map(m => m.name).sort();
    if (JSON.stringify(names) !== JSON.stringify([PUSH_MODULE, RUNTIME_MODULE])) facts.problems.push(`modules are [${names.join(', ')}], expected [${PUSH_MODULE}, ${RUNTIME_MODULE}]`);
    for (const m of facts.modules) if (m.type !== 'procedural') facts.problems.push(`${m.name} is not a standard module`);
    if (facts.version !== V2) facts.problems.push(`C5D_Bridge_Push version is ${facts.version}, expected ${V2}`);
    if (sources && facts.runtimeVersion !== sources.runtimeVersion) facts.problems.push(`runtime version is ${facts.runtimeVersion}, expected ${sources.runtimeVersion}`);
    for (const m of facts.modules) if (m.same === false) facts.problems.push(`${m.name} differs from resources/bridge/${m.name}.bas`);
    facts.ok = facts.problems.length === 0;
    return facts;
}

/** Records of a C5D_Who / C5D_List / C5D_BuildWho / C5D_RT_List* text: { ok, rows: string[][], error } (contract section 8.1). */
function parseRecords(text) {
    if (typeof text !== 'string') return { ok: false, rows: [], error: 'not a string' };
    if (text.startsWith('error:')) return { ok: false, rows: [], error: text };
    const lines = text.split('\n');
    if (lines[0] !== 'ok') return { ok: false, rows: [], error: `unexpected: ${text.slice(0, 80)}` };
    const unescape = f => f.replace(/%(09|0A|0D|25)/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
    return { ok: true, rows: lines.slice(1).map(l => l.split('\t').map(unescape)), error: null };
}
const sameText = (x, y) => String(x).toLowerCase() === String(y).toLowerCase();
/** Whether C5D_BuildWho shows this copy as ExecutingProject and it is the only project named C5D_Bridge. */
function builderWhoOk(result, library) {
    const r = parseRecords(result);
    const [name, display, file, hits] = r.rows[0] ?? [];
    return r.ok && name === BRIDGE_PROJECT && sameText(display, library) && sameText(file, library) && hits === '1';
}
/** Whether C5D_Who shows this copy with exactly the given standard modules. */
function whoOk(result, library, modules = [PUSH_MODULE, RUNTIME_MODULE]) {
    const r = parseRecords(result);
    if (!r.ok || !r.rows.length) return false;
    const [name, display, file] = r.rows[0];
    const comps = r.rows.slice(1).map(([n, t]) => `${n}:${t}`).sort();
    return name === BRIDGE_PROJECT && sameText(display, library) && sameText(file, library)
        && JSON.stringify(comps) === JSON.stringify(modules.map(m => `${m}:1`).sort());
}
/** Whether C5D_List shows exactly one project named C5D_Bridge: this copy, flagged as the executing one. */
function listOk(result, library) {
    const r = parseRecords(result);
    const bridges = r.rows.filter(([name]) => name === BRIDGE_PROJECT);
    return r.ok && bridges.length === 1 && sameText(bridges[0][1], library) && bridges[0][2] === '1';
}

const literal = text => `'${String(text).replace(/'/g, "''")}'`;
const guards = library => [`guard:$.Name=${literal(BRIDGE_PROJECT)}`, `guard:$.DisplayName=${literal(library)}`];
const call = (name, library, module, func, args, extra = {}) => ({ name, library, module, func, args, ...extra });
const objCall = (name, library, module, func, actions, mustSucceed, extra = {}) =>
    call(name, library, module, func, [], { actions, mustSucceed, resultType: 'ComObject', ...extra });

/** Steps of the automatic build (before the verify). */
function buildSteps(p, sources, options = {}) {
    const w = p.work, close = options.saveMethod === 'close';
    return [
        { name: 'precheck', local: 'precheck' },
        { name: 'prepare', local: 'prepare' },
        call('load', w, PUSH_MODULE, 'C5D_Version', [], { check: r => r === V1 || r === V2 }),
        call('install-builder', w, PUSH_MODULE, 'C5D_Install', [BUILDER_MODULE, BUILDER_SOURCE], { expected: 'ok' }),
        call('builder-who', w, BUILDER_MODULE, 'C5D_BuildWho', [], { check: r => builderWhoOk(r, w) }),
        call('put-push', w, BUILDER_MODULE, 'C5D_BuildPut', [PUSH_MODULE, sources.push], { pattern: /^ok:\d+$/ }),
        call('put-runtime', w, BUILDER_MODULE, 'C5D_BuildPut', [RUNTIME_MODULE, sources.runtime], { pattern: /^ok:\d+$/ }),
        call('version-v2', w, PUSH_MODULE, 'C5D_Version', [], { expected: V2 }),
        call('runtime-version', w, RUNTIME_MODULE, 'C5D_RT_Version', [], { expected: sources.runtimeVersion }),
        objCall('remove-builder', w, PUSH_MODULE, 'C5D_Self',
            [...guards(w), `call:$.VBProject.VBComponents.Remove($.VBProject.VBComponents.Item(${literal(BUILDER_MODULE)}))`], [2]),
        call('who-v2', w, PUSH_MODULE, 'C5D_Who', [], { check: r => whoOk(r, w) }),
        ...(close ? [] : [objCall('save', w, PUSH_MODULE, 'C5D_Self', [...guards(w), 'call:$.Save(true)', 'get:$.Dirty'], [2])]),
        // module/func/actions of the close are chosen when it runs (closeStep): C5D_Self of v2, else C5D_BuildSelf, else nothing
        { name: 'close', library: w, cleanup: true, closeOf: w, saveOnClose: close },
        { name: 'collect', local: 'collect' },
        { name: 'inspect', local: 'inspect' }
    ];
}
/** Steps of the verify of file (a new read-only copy p.verify is loaded; file itself is never loaded). */
function verifySteps(p, sources) {
    const v = p.verify;
    return [
        { name: 'verify-prepare', local: 'verify-prepare' },
        call('v-load', v, PUSH_MODULE, 'C5D_Version', [], { expected: V2 }),
        call('v-who', v, PUSH_MODULE, 'C5D_Who', [], { check: r => whoOk(r, v) }),
        call('v-list', v, PUSH_MODULE, 'C5D_List', [], { check: r => listOk(r, v) }),
        call('v-runtime', v, RUNTIME_MODULE, 'C5D_RT_Version', [], { expected: sources.runtimeVersion }),
        call('v-projects', v, RUNTIME_MODULE, 'C5D_RT_ListProjects', [], { check: r => parseRecords(r).ok }),
        call('v-install', v, PUSH_MODULE, 'C5D_Install', [PROBE_MODULE, PROBE_SOURCE], { expected: 'ok' }),
        call('v-echo', v, PROBE_MODULE, 'C5D_BuildProbeEcho', [PROBE_ARGUMENT], { expected: `build-probe:${PROBE_ARGUMENT}` }),
        call('v-self-guard', v, PUSH_MODULE, 'C5D_Install', [PUSH_MODULE, 'x'], { expected: 'error:self' }),
        { ...objCall('v-close', v, PUSH_MODULE, 'C5D_Self', [...guards(v), 'call:$.Close(false)'], [2]), cleanup: true, closeOf: v }
    ];
}
/** The whole plan of a mode. */
function planSteps(mode, p, sources, options = {}) {
    const install = options.install ? [{ name: 'install', local: 'install' }] : [];
    if (mode === 'verify') return [{ name: 'inspect', local: 'inspect' }, ...verifySteps(p, sources), ...install];
    return [...buildSteps(p, sources, options), ...verifySteps(p, sources), ...install];
}

/**
 * The concrete call of the build's close, from what ran before: v2 C5D_Self when version-v2 matched, else C5D_BuildSelf when the
 * builder identified this copy, else null (nothing identifies the copy: it is left loaded, CATIA must be restarted).
 * Close(True) only for --save-method close when every step before matched.
 */
function closeStep(step, entries) {
    const matched = name => entries.find(e => e.step === name)?.match === true;
    const all = entries.every(e => e.match === true);
    const self = matched('version-v2') ? [PUSH_MODULE, 'C5D_Self'] : matched('builder-who') ? [BUILDER_MODULE, 'C5D_BuildSelf'] : null;
    if (!self) return null;
    const save = step.saveOnClose && all;
    return objCall(step.name, step.library, self[0], self[1], [...guards(step.library), `call:$.Close(${save ? 'true' : 'false'})`], [2],
        { cleanup: true, closeOf: step.library, savedByClose: save });
}

/** Plan as printed by --dry-run: every step with decoded arguments (long sources shortened) and the eval-probe.exe tokens. */
function describePlan(steps, callArgs) {
    const short = a => (typeof a === 'string' && a.length > 120 ? `${a.slice(0, 60)}... (${a.length} chars)` : a);
    return steps.map(s => {
        if (s.local) return { step: s.name, local: s.local };
        if (!s.module) return { step: s.name, call: '(chosen at run time: C5D_Self / C5D_BuildSelf + guarded Close)', library: s.library, cleanup: true };
        return { step: s.name, call: `${s.module}.${s.func}`, library: s.library, args: s.args.map(short), actions: s.actions ?? null,
            cleanup: s.cleanup === true, tokens: callArgs(s, s.library) };
    });
}

/** Lock state of every C5D_Bridge*.catvba of the bridge folders plus the given files: [{ path, inUse }]. */
function scanLocks(p, lockProbe, extra = []) {
    const files = [];
    for (const dir of [p.bridgeDir, p.buildDir]) {
        let names = [];
        try { names = fs.readdirSync(dir); } catch { /* missing folder: nothing loaded from it */ }
        for (const n of names) if (LEFTOVER.test(n)) files.push(path.join(dir, n));
    }
    for (const f of extra) if (f && fs.existsSync(f) && !files.some(x => sameText(x, f))) files.push(f);
    return files.map(f => ({ path: f, inUse: lockProbe(f).inUse ?? null }));
}
/** Deletes old, unused build files of earlier runs (read-only cleared first). Returns the deleted paths. */
function cleanupBuildDir(p, lockProbe) {
    const deleted = [];
    let names = [];
    try { names = fs.readdirSync(p.buildDir); } catch { return deleted; }
    for (const n of names) {
        const f = path.join(p.buildDir, n);
        if (!/^C5D_Bridge-build-\d{14}-[wv]\.catvba$/i.test(n) || f === p.work || f === p.verify || lockProbe(f).inUse !== false) continue;
        try { fs.chmodSync(f, 0o666); fs.unlinkSync(f); deleted.push(f); } catch { /* kept for the next run */ }
    }
    return deleted;
}

/** Local steps. ctx: { p, sources, options, deps, facts }. Returns { ok, message, stop?, data? }. */
function runLocal(step, ctx) {
    const { p, deps, options } = ctx;
    switch (step.local) {
        case 'precheck': {
            const baseFacts = deps.inspect(p.base);
            ctx.facts.base = baseFacts;
            if (![V1, V2].includes(baseFacts.version)) return { ok: false, message: `base ${p.base} has no C5D_Bridge_Push of v1/v2 (${baseFacts.problems.join('; ')})` };
            const extra = baseFacts.modules.filter(m => ![PUSH_MODULE, RUNTIME_MODULE].includes(m.name));
            if (extra.length) return { ok: false, message: `base has other modules: ${extra.map(m => m.name).join(', ')} (only ${PUSH_MODULE} / ${RUNTIME_MODULE} are allowed)` };
            const deleted = cleanupBuildDir(p, deps.lockProbe);
            const locks = scanLocks(p, deps.lockProbe, [p.base, p.target]);
            const busy = locks.filter(l => l.inUse !== false);
            ctx.facts.locks = locks; ctx.facts.deletedOldBuildFiles = deleted;
            if (busy.length && baseFacts.version === V1) {
                return { ok: false, stop: 'restart', message: `in use (or not checkable) while the base is v1: ${busy.map(l => path.basename(l.path)).join(', ')} -> restart CATIA (no Push before the build) and run again` };
            }
            return { ok: true, message: `base ${baseFacts.version}${busy.length ? `; in use (v2 base, continuing; builder-who checks): ${busy.map(l => path.basename(l.path)).join(', ')}` : '; no bridge file in use'}` };
        }
        case 'prepare':
            fs.mkdirSync(p.buildDir, { recursive: true });
            fs.copyFileSync(p.base, p.work, fs.constants.COPYFILE_EXCL);
            fs.chmodSync(p.work, 0o666);
            return { ok: true, message: `writable copy ${p.work}` };
        case 'collect':
            try {
                fs.mkdirSync(p.collectDir, { recursive: true });
                fs.copyFileSync(p.work, p.collected, fs.constants.COPYFILE_EXCL);
                ctx.fileToVerify = p.collected;
                return { ok: true, message: `copied to ${p.collected}` };
            } catch (error) {
                return { ok: false, stop: 'resume', message: `cannot read the saved work file (${error.code ?? error.message}); quit CATIA, then run: node scripts/build-bridge-catvba.js --verify "${p.work}" --install` };
            }
        case 'inspect': {
            const file = ctx.fileToVerify;
            const facts = deps.inspect(file, ctx.sources);
            ctx.facts.inspect = facts;
            if (!facts.ok) return { ok: false, message: `${file}: ${facts.problems.join('; ')}` };
            return { ok: true, message: `${file}: v2, modules = sources${facts.orphans.length ? `; ${facts.orphans.length} orphan storages: ${facts.orphans.join(', ')}` : '; no orphan storages'}` };
        }
        case 'verify-prepare':
            fs.mkdirSync(p.buildDir, { recursive: true });
            fs.copyFileSync(ctx.fileToVerify, p.verify, fs.constants.COPYFILE_EXCL);
            fs.chmodSync(p.verify, 0o444);
            return { ok: true, message: `read-only copy ${p.verify}` };
        case 'install': {
            const orphans = ctx.facts.inspect?.orphans ?? [];
            if (orphans.length && !options.allowOrphans) {
                return { ok: false, message: `not installed: ${orphans.length} orphan storages (data of other projects) would ship in the VSIX; make the bridge by hand (contract 11.3) or pass --allow-orphans` };
            }
            const before = deps.sha256(p.target);
            fs.mkdirSync(path.dirname(p.backup), { recursive: true });
            fs.copyFileSync(p.target, p.backup, fs.constants.COPYFILE_EXCL);
            fs.copyFileSync(ctx.fileToVerify, p.target);
            const after = deps.sha256(p.target);
            ctx.facts.installed = { before, after, backup: p.backup };
            return { ok: true, message: `resources/bridge/C5D_Bridge.catvba replaced (sha256 ${before.slice(0, 8)} -> ${after.slice(0, 8)}; old file: ${p.backup})` };
        }
        default: throw new Error(`unknown local step ${step.local}`);
    }
}

/**
 * Runs the plan one step at a time. A failed step stops the rest except the cleanup steps (the guarded closes), which still run unless
 * CATIA is gone, unreachable, at a dialog (timeout) or reported a fatal VBA error. deps (injectable for tests): callStep(step) ->
 * runCallStep result, alive() -> true / false / null, fatal(entry, step) -> reason | null, lockProbe, inspect, sha256, log.
 * Returns { entries, stopped, failed }.
 */
async function runPlan(steps, ctx) {
    const { deps } = ctx;
    const entries = [];
    let failed = null, hard = null, stopped = null;
    for (const planned of steps) {
        let step = planned, reason = null;
        if (hard) reason = hard;
        else if (!planned.cleanup && failed) reason = `not run: ${failed} failed`;
        else if (planned.cleanup) {
            // a close runs only when ExecuteScript of the load returned (the copy is loaded), whatever the value was
            const load = planned.closeOf === ctx.p.verify ? 'v-load' : 'load';
            if (entries.find(e => e.step === load)?.ok !== true) reason = 'not run: nothing was loaded';
            else if (!planned.module) {
                step = closeStep(planned, entries);
                if (!step) { reason = 'not run: the copy cannot identify itself (no C5D_Self / C5D_BuildSelf); it stays loaded, restart CATIA'; failed = failed ?? planned.name; }
            }
        }
        if (reason) { entries.push({ step: planned.name, skipped: reason, match: false }); continue; }
        if (step.local) {
            let r;
            try { r = runLocal(step, ctx); } catch (error) { r = { ok: false, message: String(error.message ?? error) }; }
            entries.push({ step: step.name, local: step.local, match: r.ok, message: r.message });
            deps.log?.(`${r.ok ? 'ok  ' : 'FAIL'} ${step.name}: ${r.message}`);
            if (!r.ok) { failed = failed ?? step.name; if (r.stop) { stopped = r.stop; hard = `not run: stopped at ${step.name}`; } }
            continue;
        }
        const run = await deps.callStep(step);
        const r = run.result;
        const entry = { step: step.name, call: `${step.module}.${step.func}`, match: run.match === true, ok: r?.ok === true, result: r?.result ?? null,
            hresult: r?.hresult ?? null, message: r?.message ?? null, timedOut: run.timedOut === true, exitCode: run.exitCode ?? null,
            ...(run.actions ? { actions: run.actions.map(a => ({ action: a.action, op: a.op, ok: a.ok, result: a.result ?? null, hresult: a.hresult ?? null, message: a.message ?? null })) } : {}),
            ...(step.savedByClose ? { savedByClose: true } : {}) };
        entries.push(entry);
        const alive = deps.alive();
        const fatal = deps.fatal(entry, step);
        deps.log?.(`${entry.match ? 'ok  ' : 'FAIL'} ${step.name} (${entry.call}): ${JSON.stringify(entry.result)}${entry.hresult ? ` ${entry.hresult} ${String(entry.message ?? '').split('\n')[0]}` : ''}`);
        if (!entry.match) failed = failed ?? step.name;
        if (alive === false) hard = `not run: CATIA (CNEXT.exe) is gone after ${step.name}`;
        else if (run.timedOut) hard = `not run: ${step.name} timed out (CATIA may be at a dialog)`;
        else if (run.exitCode === 10 || run.exitCode === 11) hard = `not run: CATIA not reachable in ${step.name}`;
        else if (fatal) hard = `not run: ${fatal} in ${step.name}`;
        if (hard) entry.stoppedRun = hard;
    }
    return { entries, stopped, failed };
}

/** Report summary lines: one per step plus the verdict. */
function summaryLines(result, ctx) {
    const lines = result.entries.map(e => e.skipped ? `-    ${e.step}: ${e.skipped}` : `${e.match ? 'ok  ' : 'FAIL'} ${e.step}${e.call ? ` (${e.call})` : ''}: ${e.local ? e.message : JSON.stringify(e.result)}${e.hresult ? ` [${e.hresult}]` : ''}`);
    const verdict = result.failed ? (result.stopped === 'resume' ? `STOPPED at ${result.failed} (resumable, see above)` : `FAILED at ${result.failed}`)
        : ctx.options.install ? 'DONE: resources/bridge/C5D_Bridge.catvba is bridge v2 (commit it; the master name sha8 changes)' : 'DONE: verified (not installed)';
    return [...lines, verdict];
}

async function main(argv = process.argv.slice(2)) {
    let options;
    try { options = parseArgs(argv); } catch (error) { console.error(error.message); return 2; }
    if (options.mode === 'help') { console.log(USAGE); return 0; }
    if (options.mode === 'export-import') {
        const sources = readSources();
        fs.mkdirSync(options.dir, { recursive: true });
        for (const [name, text] of [[PUSH_MODULE, sources.push], [RUNTIME_MODULE, sources.runtime]]) {
            const file = path.join(options.dir, `${name}.bas`);
            fs.writeFileSync(file, `Attribute VB_Name = "${name}"\r\n${text}`, { flag: 'wx' });
            console.log(file);
        }
        return 0;
    }
    if (options.mode === 'inspect') {
        const facts = inspectCatvba(path.resolve(options.file), readSources());
        console.log(JSON.stringify(facts, null, 2));
        return facts.ok && !facts.orphans.length ? 0 : 1;
    }
    const sources = readSources();
    const p = createPaths({ base: options.base, file: options.file });
    const steps = planSteps(options.mode, p, sources, options);
    const probe = require('./probe-bridge');
    if (options.dryRun) {
        console.log(JSON.stringify({ mode: options.mode, dryRun: true, options, paths: p, versions: { push: sources.pushVersion, runtime: sources.runtimeVersion },
            steps: describePlan(steps, probe.callArgs) }, null, 2));
        return 0;
    }
    const exe = require('./probe-evaluate').buildProbe(ROOT);
    const listed = probe.cnextProcesses();
    const pids = listed.processes.map(x => x.pid);
    if (!pids.length) { console.error(`CATIA (CNEXT.exe) is not running${listed.error ? `: ${listed.error}` : ''}`); return 1; }
    const ctx = {
        p, sources, options, facts: {}, fileToVerify: options.mode === 'verify' ? p.file : null,
        deps: {
            callStep: step => probe.runCallStep(exe, step, step.library, { noRetry: true }),
            alive: () => probe.catiaAlive(pids),
            fatal: (entry, step) => probe.cycleFatal(entry, { closeAction: step.actions ? step.actions.length - 1 : undefined }),
            lockProbe: probe.lockProbe, inspect: inspectCatvba, sha256: probe.sha256, log: line => console.log(line)
        }
    };
    const report = { mode: options.mode, startedAt: new Date().toISOString(), options, paths: p, cnextPids: pids, exe,
        versions: { push: sources.pushVersion, runtime: sources.runtimeVersion } };
    const result = await runPlan(steps, ctx);
    Object.assign(report, { entries: result.entries, facts: ctx.facts, failed: result.failed, stopped: result.stopped, finishedAt: new Date().toISOString() });
    report.summary = summaryLines(result, ctx);
    fs.mkdirSync(path.dirname(p.report), { recursive: true });
    fs.writeFileSync(p.report, JSON.stringify(report, null, 2));
    console.log(['', ...report.summary, `report: ${p.report}`].join('\n'));
    return result.failed ? (result.stopped === 'resume' ? 3 : 1) : 0;
}

if (require.main === module) main().then(code => { process.exitCode = code; }, error => { console.error(error); process.exitCode = 1; });

module.exports = { parseArgs, readSources, createPaths, inspectCatvba, parseRecords, builderWhoOk, whoOk, listOk, buildSteps, verifySteps, planSteps,
    closeStep, describePlan, runPlan, runLocal, summaryLines, comparable, main, BUILDER_SOURCE, PROBE_SOURCE, PROBE_ARGUMENT, BRIDGE_PROJECT, PUSH_MODULE,
    RUNTIME_MODULE, BUILDER_MODULE, PROBE_MODULE, V1, V2 };
