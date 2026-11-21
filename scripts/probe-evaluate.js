// Evaluate probe (developer tool, not shipped): can CATIA.SystemService.Evaluate run VBA code (iLanguage = 1)
// without loading vbscript.dll, and can that VBA code reach CATIA's own VBE (VBProjects)?
// Usage: node scripts/probe-evaluate.js --check-only
//        node scripts/probe-evaluate.js [--byref] basic sub typed params apc apc2
//        node scripts/probe-evaluate.js jscript jscript-params jscript-fso jscript-apc jscript-read
// Steps (each runs as its own temp/eval-probe.exe process, 60 s timeout; see tools/catia-run/EvalProbe.cs):
//   check   connect + SystemService only          basic   Function CATMain() returning "ok:2"
//   sub     Sub CATMain (no return value)          typed   `Dim n As Long` (VBA-only syntax) returning "vba:Long:3"
//   params  one string parameter (ByVal, then ByRef if needed)
//   apc     CreateObject("MSAPC.Apc.7.1"/"6.2"/"MSAPC.Apc").VBE.VBProjects count + names
//   apc2    = apc2-getobject (GetObject(, "MSAPC.Apc")), apc2-catia (CATIA.VBE), apc2-application (Application.VBE)
//   jscript         language 4 (loads jscript.dll) returning "js:2"
//   jscript-params  one string parameter to function CATMain(a)
//   jscript-fso     new ActiveXObject("ADODB.Stream") / ("Scripting.FileSystemObject") (create only, no file written)
//   jscript-apc     new ActiveXObject("MSAPC.Apc.7.1"/"6.2"/"MSAPC.Apc").VBE.VBProjects -> "count=N|name1|...;<create log>"
//   jscript-read    first 3 projects: components (Name, Type, CountOfLines) + CodeModule.Lines(1, min(3, n)) of the
//                   first non-empty module (JSON; via = lines-call / lines-perline / none)
//   JScript steps (language 4) always run after the other steps, in the order given.
// Read-only for CATIA: never adds/changes modules or projects. The exe is built into temp/ (git-ignored) on every run.
// Modules of CNEXT.exe are recorded before the first step and after each step (scripts/probe-catscript.js moduleSnapshot).
// Each run writes temp/probe-evaluate-<yyyyMMdd-HHmmss>.json and prints a summary to stdout.
const fs = require('node:fs');
const path = require('node:path');
const { execFile, execFileSync } = require('node:child_process');
const { moduleSnapshot, KEY_DLLS, timestamp, addedModules } = require('./probe-catscript');

const STEP_TIMEOUT_MS = 60000;
/** Step name -> CATScriptLanguage (1 = VBA, 4 = JScript, null = no Evaluate). Order = recommended order. */
const STEPS = { check: null, basic: 1, sub: 1, typed: 1, params: 1, apc: 1, 'apc2-getobject': 1, 'apc2-catia': 1, 'apc2-application': 1,
    jscript: 4, 'jscript-params': 4, 'jscript-fso': 4, 'jscript-apc': 4, 'jscript-read': 4 };
const ALIASES = { apc2: ['apc2-getobject', 'apc2-catia', 'apc2-application'] };
const JSCRIPT = 4;
const STEP_LIST = Object.keys(STEPS).flatMap(name => name === ALIASES.apc2[0] ? ['apc2', name] : [name]);
const USAGE = 'Usage: node scripts/probe-evaluate.js --check-only\n'
    + '       node scripts/probe-evaluate.js [--byref] <step> [<step> ...]\n'
    + `  steps: ${STEP_LIST.join(' ')}\n`
    + '  recommended (CATIA freshly restarted): --check-only, then basic sub params apc apc2, then the jscript steps last;\n'
    + '  JScript only (CATIA freshly restarted): --check-only, then jscript jscript-params jscript-fso jscript-apc jscript-read';

/** { mode: 'help' | 'check-only' | 'run', steps, byref, notes }. Throws on bad arguments. */
function parseArgs(args) {
    if (args.some(arg => arg === '--help' || arg === '-h')) return { mode: 'help', steps: [], byref: false, notes: [] };
    const checkOnly = args.includes('--check-only');
    const byref = args.includes('--byref');
    const names = args.filter(arg => arg !== '--check-only' && arg !== '--byref');
    const unknown = names.filter(name => !Object.hasOwn(STEPS, name) && !Object.hasOwn(ALIASES, name));
    if (unknown.length) throw new Error(`Unknown step or option: ${unknown.join(' ')}\n${USAGE}`);
    if (checkOnly && names.length) throw new Error(`--check-only runs no step: ${names.join(' ')}\n${USAGE}`);
    if (checkOnly) return { mode: 'check-only', steps: [], byref, notes: [] };
    if (!names.length) throw new Error(`No step given.\n${USAGE}`);
    const steps = [...new Set(names.flatMap(name => ALIASES[name] ?? [name]))];
    const notes = [];
    // JScript steps load jscript.dll: run them after every other step (keeping their order) so the VBA steps are judged first.
    const ordered = [...steps.filter(name => STEPS[name] !== JSCRIPT), ...steps.filter(name => STEPS[name] === JSCRIPT)];
    if (ordered.some((name, i) => name !== steps[i])) {
        notes.push(`JScript steps were moved to the end (now: ${ordered.join(' ')}); they load jscript.dll, so the VBA steps are judged before them.`);
    }
    return { mode: 'run', steps: ordered, byref, notes };
}

/**
 * Verdict about vbscript.dll. beforeLoaded: state before the first step (or now, for check-only).
 * entries: [{ step, language, ok, loadedAfter }] in run order (loadedAfter = vbscript.dll loaded after that step).
 */
function judge(mode, beforeLoaded, entries = []) {
    if (mode === 'check-only') {
        return beforeLoaded ? 'vbscript.dll is already loaded in CNEXT.exe (no step was run)'
            : 'vbscript.dll is not loaded in CNEXT.exe (no step was run)';
    }
    if (beforeLoaded) return 'undetermined: vbscript.dll was already loaded before the first step (restart CATIA and re-run)';
    const vba = entries.filter(entry => entry.language === 1);
    const js = entries.filter(entry => entry.language === JSCRIPT);
    const firstIndex = entries.findIndex(entry => entry.loadedAfter);
    const first = entries[firstIndex];
    let verdict;
    if (first && first.language === 1) verdict = `Evaluate(VBA) loaded vbscript.dll (first seen after step ${first.step})`;
    else if (first && first.language === JSCRIPT) verdict = `${vba.length ? 'Evaluate(VBA) did not load vbscript.dll; ' : ''}Evaluate(JScript) loaded vbscript.dll (first seen after step ${first.step})`;
    else if (first) verdict = `vbscript.dll was loaded after step ${first.step} (no Evaluate in that step)`;
    else {
        verdict = [vba.length ? 'Evaluate(VBA) did not load vbscript.dll' : '', js.length ? 'Evaluate(JScript) did not load vbscript.dll' : '']
            .filter(Boolean).join('; ') || 'no Evaluate step was run; vbscript.dll not loaded';
    }
    if (first && first.language !== JSCRIPT && entries.slice(firstIndex + 1).some(entry => entry.language === JSCRIPT)) {
        verdict += ' (JScript steps ran after vbscript.dll was loaded, so they cannot be judged; restart CATIA and run only the jscript steps)';
    }
    if (vba.length && !vba.some(entry => entry.ok)) verdict += ' (warning: no VBA step succeeded, so VBA code may never have run)';
    if (js.length && !js.some(entry => entry.ok)) verdict += ' (warning: no JScript step succeeded, so JScript code may never have run)';
    return verdict;
}

/** A step result that is itself JSON (jscript-read) -> parsed value; otherwise undefined. */
function parseJsonResult(result) {
    if (typeof result !== 'string' || !result.startsWith('{')) return undefined;
    try { return JSON.parse(result); } catch { return undefined; }
}

/** JSON lines printed by eval-probe.exe; non-JSON lines are kept as text. */
function parseProbeOutput(stdout) {
    const lines = [], other = [];
    for (const line of stdout.split(/\r?\n/).map(text => text.trim()).filter(Boolean)) {
        try { lines.push(JSON.parse(line)); } catch { other.push(line); }
    }
    return { lines, other };
}

function buildProbe(root) {
    const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
    if (!fs.existsSync(csc)) throw new Error(`csc.exe not found: ${csc}`);
    const exe = path.join(root, 'temp', 'eval-probe.exe');
    fs.mkdirSync(path.dirname(exe), { recursive: true });
    try {
        execFileSync(csc, ['/nologo', '/optimize+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
            path.join(root, 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
    } catch (error) {
        throw new Error(`csc failed: ${String(error.stdout || '') + String(error.stderr || '') || error}`);
    }
    return exe;
}

/** Runs one step in its own eval-probe.exe process. A timeout most likely means a modal dialog in CATIA. */
function runStep(exe, step, byref) {
    const args = [...(byref ? ['--byref'] : []), step];
    const started = Date.now();
    return new Promise(resolve => {
        execFile(exe, args, { windowsHide: true, timeout: STEP_TIMEOUT_MS, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024 }, (error, stdout, stderr) => {
            const parsed = parseProbeOutput(stdout || '');
            const timedOut = Boolean(error && error.killed);
            resolve({
                step, args, exitCode: error ? (typeof error.code === 'number' ? error.code : null) : 0, timedOut,
                error: error && !timedOut && typeof error.code !== 'number' ? String(error) : null,
                result: parsed.lines.find(line => line.step === step) ?? parsed.lines[0] ?? null,
                stdoutOther: parsed.other, stderr: (stderr || '').trim(), durationMs: Date.now() - started
            });
        });
    });
}

/** temp/probe-evaluate-<time>.json; never overwrites an earlier report. */
function uniqueReportPath(directory) {
    const base = path.join(directory, `probe-evaluate-${timestamp()}`);
    let candidate = base + '.json';
    for (let i = 2; fs.existsSync(candidate); i++) candidate = `${base}-${i}.json`;
    return candidate;
}

function formatSummary(report) {
    const yn = value => value === true ? 'yes' : value === false ? 'no' : '-';
    const lines = [`mode: ${report.mode}${report.byref ? ' (--byref)' : ''}`, `report: ${report.reportPath}`];
    if (report.error) lines.push(`error: ${report.error}`);
    const snapshots = Object.values(report.snapshots);
    if (snapshots.length) {
        const first = snapshots[0], last = snapshots[snapshots.length - 1];
        const readable = first.processes.filter(p => !p.error && p.moduleCount > 0);
        lines.push(`CNEXT.exe pids: ${first.cnextPids.join(', ') || '(none)'}`);
        lines.push(`Process.Modules: ${first.powershell.error ? 'ERROR ' + first.powershell.error
            : readable.length ? readable.map(p => `pid ${p.pid}: ${p.moduleCount} modules`).join(', ') : 'no readable process'}`);
        const width = Math.max(...snapshots.map(s => s.label.length), 6) + 2;
        lines.push('', 'dll'.padEnd(14) + snapshots.map(s => s.label.padEnd(width)).join(''));
        for (const dll of KEY_DLLS) {
            lines.push(dll.padEnd(14) + snapshots.map(s => {
                const d = s.dlls[dll];
                const detail = d.byModules === null ? '(T)' : d.byModules !== d.byTasklist ? `(M:${yn(d.byModules)},T:${yn(d.byTasklist)})` : '';
                return (yn(d.loaded) + detail).padEnd(width);
            }).join(''));
        }
        lines.push('  (yes/no = loaded in CNEXT.exe by Process.Modules or tasklist /M; (T) = tasklist only)');
        if (snapshots.length > 1) lines.push(`modules added ${first.label} -> ${last.label}: ${addedModules(first, last).join(', ') || '(none)'}`);
    }
    if (report.runs.length) {
        lines.push('', 'steps:');
        for (const run of report.runs) {
            const r = run.result;
            if (run.timedOut) lines.push(`  ${run.step}: TIMEOUT after ${STEP_TIMEOUT_MS / 1000}s - CATIA may be stopped at a dialog (check CATIA; later steps were skipped)`);
            else if (!r) lines.push(`  ${run.step}: no JSON output (exitCode=${run.exitCode}${run.error ? ', ' + run.error : ''}${run.stderr ? ', stderr: ' + run.stderr : ''})`);
            else {
                lines.push(`  ${run.step}: ok=${yn(r.ok)} match=${yn(r.match)} result=${JSON.stringify(r.result)} (${r.resultType}) ${r.elapsedMs}ms`
                    + (r.paramMode && (run.step === 'params' || run.step === 'jscript-params') ? ` paramMode=${r.paramMode}` : ''));
                if (!r.ok || r.hresult !== '0x00000000') lines.push(`    hresult=${r.hresult} message=${r.message}`);
                if (r.attempts?.length > 1) lines.push(`    attempts: ${r.attempts.join(' / ')}`);
                const parsed = parseJsonResult(r.result);
                if (parsed !== undefined) lines.push('    result (parsed):', ...JSON.stringify(parsed, null, 2).split('\n').map(line => '      ' + line));
            }
        }
    }
    if (report.summary?.verdict) lines.push('', `verdict: ${report.summary.verdict}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    return lines.join('\n');
}

async function main() {
    let options;
    try { options = parseArgs(process.argv.slice(2)); } catch (error) { console.error(error.message); process.exitCode = 2; return; }
    if (options.mode === 'help') { console.log(USAGE); return; }
    const root = path.resolve(__dirname, '..');
    const reportPath = uniqueReportPath(path.join(root, 'temp'));
    const report = { mode: options.mode, byref: options.byref, steps: options.steps, startedAt: new Date().toISOString(), reportPath,
        exe: null, notes: [...options.notes], snapshots: {}, runs: [] };
    const snap = label => (report.snapshots[label] = moduleSnapshot(label));
    try {
        report.exe = buildProbe(root);
        const first = snap(options.mode === 'check-only' ? 'current' : 'before');
        if (!first.cnextPids.length) throw new Error('CNEXT.exe is not running');
        if (first.cnextPids.length > 1) report.notes.push(`More than one CNEXT.exe is running (${first.cnextPids.join(', ')}); a DLL counts as loaded if any of them has it.`);
        if (first.powershell.error) report.notes.push(`Process.Modules could not be read (${first.powershell.error}); judged by tasklist /M only.`);
        if (options.mode === 'run' && first.loaded) report.notes.push('vbscript.dll was already loaded before the first step; restart CATIA to judge the dependency. Step results are still valid.');
        const entries = [];
        for (const step of options.steps) {
            const run = await runStep(report.exe, step, options.byref);
            report.runs.push(run);
            const after = snap('after:' + step);
            entries.push({ step, language: STEPS[step], ok: run.result?.ok === true, loadedAfter: after.loaded });
            if (run.timedOut) {
                report.notes.push(`Step ${step} timed out after ${STEP_TIMEOUT_MS / 1000}s; CATIA may be stopped at a dialog. Remaining steps were not run.`);
                break;
            }
            if (run.exitCode === 10 || run.exitCode === 11) {
                report.notes.push(`Step ${step}: could not reach CATIA (exit ${run.exitCode}); remaining steps were not run.`);
                break;
            }
        }
        const snapshots = Object.values(report.snapshots);
        report.summary = {
            dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, Object.fromEntries(snapshots.map(s => [s.label, s.dlls[dll].loaded]))])),
            modulesAdded: snapshots.length > 1 ? addedModules(snapshots[0], snapshots[snapshots.length - 1]) : undefined,
            steps: report.runs.map(run => ({ step: run.step, timedOut: run.timedOut, exitCode: run.exitCode, ok: run.result?.ok ?? null,
                match: run.result?.match ?? null, result: run.result?.result ?? null, hresult: run.result?.hresult ?? null, paramMode: run.result?.paramMode ?? null })),
            verdict: judge(options.mode, first.loaded, entries)
        };
        if (report.runs.some(run => run.timedOut || !run.result || run.exitCode !== 0)) process.exitCode = 1;
    } catch (error) {
        report.error = String(error); process.exitCode = 1;
    }
    report.finishedAt = new Date().toISOString();
    fs.mkdirSync(path.dirname(reportPath), { recursive: true });
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(formatSummary(report));
}
if (require.main === module) main();
module.exports = { parseArgs, judge, parseProbeOutput, parseJsonResult, buildProbe, STEPS, ALIASES };
