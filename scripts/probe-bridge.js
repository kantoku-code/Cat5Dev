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

/** Steps in run order. requires: skipped unless these steps (when selected in this run) matched. */
const STEPS = [
    { name: 'version', module: BRIDGE_MODULE, func: 'C5D_Version', args: [], expected: BRIDGE_VERSION },
    { name: 'install', module: BRIDGE_MODULE, func: 'C5D_Install', args: [PROBE_MODULE, PROBE_SOURCE], expected: 'ok', requires: ['version'] },
    { name: 'echo', module: PROBE_MODULE, func: 'C5D_ProbeEcho', args: [ECHO_ARGUMENT], expected: 'probe:Long:' + ECHO_ARGUMENT, requires: ['version', 'install'] },
    { name: 'projects', module: PROBE_MODULE, func: 'C5D_ProbeProjects', args: [], pattern: /^count=\d+;err=0$/, requires: ['version', 'install'] },
    { name: 'self-guard', module: BRIDGE_MODULE, func: 'C5D_Install', args: [BRIDGE_MODULE, 'x'], expected: 'error:self', requires: ['version'] }
];
const STEP_NAMES = STEPS.map(step => step.name);
const USAGE = 'Usage: node scripts/probe-bridge.js [<step> ...]\n'
    + '       node scripts/probe-bridge.js --check-only\n'
    + `  steps (default: all, always run in this order): ${STEP_NAMES.join(' ')}\n`
    + '  Never runs a .catvbs/.CATScript. --check-only only reads CNEXT.exe modules and the bridge files.';

/** { mode: 'run' | 'check-only' | 'help', steps }. Throws on bad arguments. */
function parseArgs(args) {
    if (args.some(arg => arg === '--help' || arg === '-h')) return { mode: 'help', steps: [] };
    let checkOnly = false;
    const steps = new Set();
    for (const arg of args) {
        if (arg === '--check-only') checkOnly = true;
        else if (STEP_NAMES.includes(arg)) steps.add(arg);
        else throw new Error(`Unknown option or step: ${arg}\n${USAGE}`);
    }
    if (checkOnly && steps.size) throw new Error(`--check-only takes no steps\n${USAGE}`);
    if (checkOnly) return { mode: 'check-only', steps: [] };
    return { mode: 'run', steps: steps.size ? STEP_NAMES.filter(name => steps.has(name)) : [...STEP_NAMES] };
}

/** One command-line token for eval-probe.exe: b64:<UTF-8 base64> (survives any code page, CR/LF and quotes). */
const encodeToken = text => 'b64:' + Buffer.from(String(text), 'utf8').toString('base64');
/** eval-probe.exe arguments of a step: --label <name> vbacall <library> <module> <function> [args...]. */
function callArgs(step, library) {
    return ['--label', step.name, 'vbacall', ...[library, step.module, step.func, ...step.args].map(encodeToken)];
}
/** Whether a step's VBA return value is the expected one. */
function matchStep(step, result) {
    if (typeof result !== 'string') return false;
    return step.pattern ? step.pattern.test(result) : result === step.expected;
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
/** { path, exists, size, modifiedAt, sha256, error } without modifying the file. */
function fileState(file) {
    const state = { path: file, exists: fs.existsSync(file), size: null, modifiedAt: null, sha256: null, error: null };
    if (!state.exists) return state;
    try {
        const stat = fs.statSync(file);
        Object.assign(state, { size: stat.size, modifiedAt: stat.mtime.toISOString(), sha256: sha256(file) });
    } catch (error) { state.error = `${error.code ?? ''} ${error.message}`.trim(); }
    return state;
}
/**
 * Copies the bridge to the per-user folder (overwrites). A locked / read-only target is reported (locked: true), not thrown.
 * Returns { ok, locked, code, message, source, before, after }.
 */
function copyBridge(source, target, copy = fs.copyFileSync) {
    if (!fs.existsSync(source)) throw new Error(`${source} does not exist (save the VBA project ${BRIDGE_PROJECT} there first)`);
    const result = { ok: false, locked: false, code: null, message: null, source: fileState(source), before: fileState(target), after: null };
    fs.mkdirSync(path.dirname(target), { recursive: true });
    try { copy(source, target); } catch (error) {
        if (!LOCK_CODES.includes(error.code)) throw error;
        Object.assign(result, { locked: true, code: error.code, message: error.message, after: fileState(target) });
        return result;
    }
    result.after = fileState(target);
    result.ok = result.after.sha256 !== null && result.after.sha256 === result.source.sha256;
    if (!result.ok) result.message = 'hash of the copy differs from the source';
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
function runCallStep(exe, step, library) {
    const args = callArgs(step, library);
    const started = Date.now();
    return new Promise(resolve => {
        execFile(exe, args, { windowsHide: true, timeout: STEP_TIMEOUT_MS, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }, (error, stdout, stderr) => {
            const parsed = parseProbeOutput(stdout || '');
            const timedOut = Boolean(error && error.killed);
            const result = parsed.lines.find(line => line.step === step.name) ?? parsed.lines[0] ?? null;
            resolve({
                step: step.name, call: `${step.module}.${step.func}`, exitCode: error ? (typeof error.code === 'number' ? error.code : null) : 0, timedOut,
                error: error && !timedOut && typeof error.code !== 'number' ? String(error) : null,
                result, match: result?.ok === true && matchStep(step, result.result),
                // what eval-probe.exe actually passed to ExecuteScript, compared with what was meant (node -> exe command line)
                argumentsIntact: result ? result.library === library && JSON.stringify(result.arguments ?? null) === JSON.stringify(step.args) : null,
                stdoutOther: parsed.other, stderr: (stderr || '').trim(), durationMs: Date.now() - started
            });
        });
    });
}

/**
 * Verdict of a run. beforeLoaded: vbscript.dll before the first step.
 * entries: [{ step, ok, match, result, hresult, timedOut, skipped, loadedAfter }] of the selected steps in run order
 * (skipped = reason when the step was not run).
 */
function judgeBridge(beforeLoaded, entries) {
    const get = name => entries.find(entry => entry.step === name);
    const outcome = e => e.skipped ? `not run (${e.skipped})` : e.timedOut ? 'timed out (CATIA may be stopped at a dialog)'
        : e.ok ? `returned ${JSON.stringify(e.result ?? null)}` : `failed (${e.hresult ?? 'no result'})`;
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
    if (beforeLoaded) parts.push('vbscript.dll undetermined (already loaded before the first step; restart CATIA, run nothing else, and re-run)');
    else {
        const first = entries.find(entry => entry.loadedAfter);
        parts.push(first ? `vbscript.dll loaded (first seen after step ${first.step})` : 'vbscript.dll not loaded');
    }
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

/** temp/probe-bridge-<time>.json; never overwrites an earlier report. */
function uniqueReportPath(directory) {
    const base = path.join(directory, `probe-bridge-${timestamp()}`);
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
    const yn = value => value === true ? 'yes' : value === false ? 'no' : '-';
    const hash = state => !state ? '-' : !state.exists ? '(missing)' : state.sha256 ? state.sha256.slice(0, 16) : `(unreadable: ${state.error})`;
    const lines = [`mode: ${report.mode}`, `report: ${report.reportPath}`];
    if (report.bridge) {
        const b = report.bridge;
        lines.push(`bridge source: ${b.source.path} sha256=${hash(b.source)}${b.source.size !== null ? ` ${b.source.size} bytes` : ''}`);
        if (report.mode === 'check-only') lines.push(`bridge target: ${b.target.path} sha256=${hash(b.target)}`);
        else {
            lines.push(`bridge target: ${b.after?.path ?? b.before.path} before=${hash(b.before)} after=${hash(b.after)}${b.locked ? ` LOCKED (${b.code})` : b.ok ? ' (copied, hash ok)' : ''}`);
            if (b.targetAtEnd) lines.push(`  at end: ${hash(b.targetAtEnd)}${b.targetAtEnd.sha256 === b.after?.sha256 ? ' (unchanged)' : ' (CHANGED)'}`);
        }
    }
    if (report.error) lines.push(`error: ${report.error}`);
    if (report.cnext) {
        if (report.cnext.error) lines.push(`CNEXT.exe start time: ERROR ${report.cnext.error}`);
        for (const p of report.cnext.processes) lines.push(`CNEXT.exe pid ${p.pid}: ${describeStart(p.startedAt)}`);
    }
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

async function main() {
    let options;
    try { options = parseArgs(process.argv.slice(2)); } catch (error) { console.error(error.message); process.exitCode = 2; return; }
    if (options.mode === 'help') { console.log(USAGE); return; }
    fs.mkdirSync(TEMP, { recursive: true });
    const reportPath = uniqueReportPath(TEMP);
    const report = { mode: options.mode, steps: options.steps, startedAt: new Date().toISOString(), reportPath, project: BRIDGE_PROJECT,
        notes: [], snapshots: {} };
    try {
        if (options.mode === 'check-only') runCheckOnly(report);
        else await runBridge(report, options);
    } catch (error) {
        report.error = String(error); process.exitCode = 1;
    }
    report.finishedAt = new Date().toISOString();
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(formatSummary(report));
}
if (require.main === module) main();
module.exports = { parseArgs, encodeToken, callArgs, matchStep, codepoints, bridgePaths, fileState, copyBridge, judgeBridge, describeStart,
    STEPS, STEP_NAMES, PROBE_SOURCE, ECHO_ARGUMENT, BRIDGE_MODULE, PROBE_MODULE };
