// VBAProject probe (developer tool, not shipped): can CATIA.SystemService.ExecuteScript with
// iType = 2 (catScriptLibraryTypeVBAProject) call a procedure of a loaded VBA project as real VBA (VBE7.DLL),
// without loading vbscript.dll?
// Usage: node scripts/probe-vbaproject.js --call [--library <path.catvba>] [<step> ...]
//        node scripts/probe-vbaproject.js --write-bas
//        node scripts/probe-vbaproject.js --info | --install | --remove     (need `npm run compile` and bin/catia-run.exe)
// Modes:
//   --call       runs temp/eval-probe.exe (built from tools/catia-run/EvalProbe.cs) once per step:
//                  vbaproj-ping vbaproj-sub vbaproj-projects vbaproj-filename-basename vbaproj-filename-empty vbaproj-filename-project
//                (vbaproj-filename = the three -filename-* steps). Never runs a .catvbs/.CATScript, so it can judge vbscript.dll
//                on a freshly restarted CATIA. --library defaults to the "library" of temp/probe-vbaproject-info.json (--info).
//                CNEXT.exe modules (vbscript/jscript/scrrun/vbe7/apc71) are recorded before the first step and after each step.
//                Each step has a 60 s timeout (CATIA may be stopped at a dialog); the run stops there.
//   --write-bas  writes temp/C5D_BridgeTest.bas (ASCII, CRLF, no Attribute line) to paste into a new standard module of toQI.
//   --info       (operation.catvbs; loads vbscript.dll) reads toQI's FileName/Protection/components -> temp/probe-vbaproject-info.json.
//   --install    (operation.catvbs) adds the standard module C5D_BridgeTest to toQI; aborts when it already exists.
//   --remove     (operation.catvbs) deletes C5D_BridgeTest only when its text has the owner marker; verifies it is gone.
// The VBA project is never saved. Only the module C5D_BridgeTest is added/removed; other modules are never written.
// Each run writes temp/probe-vbaproject-<mode>-<yyyyMMdd-HHmmss>.json and prints a summary to stdout.
const fs = require('node:fs');
const path = require('node:path');
const { execFile } = require('node:child_process');
const { moduleSnapshot, KEY_DLLS, timestamp, addedModules } = require('./probe-catscript');
const { buildProbe, parseProbeOutput } = require('./probe-evaluate');

const ROOT = path.resolve(__dirname, '..');
const TEMP = path.join(ROOT, 'temp');
const INFO_PATH = path.join(TEMP, 'probe-vbaproject-info.json');
const BASELINE_PATH = path.join(TEMP, 'probe-vbaproject-baseline.json');
const BAS_PATH = path.join(TEMP, 'C5D_BridgeTest.bas');
const PROJECT = 'toQI';
const MODULE = 'C5D_BridgeTest';
const MARKER = "' Cat5Dev-bridge-test-owner";
const ENCODING = 'shift_jis';
const STEP_TIMEOUT_MS = 60000;
const PING_EXPECTED = 'vba:Long:cat5dev-param';
const CALL_STEPS = ['vbaproj-ping', 'vbaproj-sub', 'vbaproj-projects', 'vbaproj-filename-basename', 'vbaproj-filename-empty', 'vbaproj-filename-project'];
const CALL_ALIASES = { 'vbaproj-filename': ['vbaproj-filename-basename', 'vbaproj-filename-empty', 'vbaproj-filename-project'] };
const PING_STEPS = ['vbaproj-ping', 'vbaproj-filename-basename', 'vbaproj-filename-empty', 'vbaproj-filename-project'];
const MODES = { '--info': 'info', '--install': 'install', '--call': 'call', '--remove': 'remove', '--write-bas': 'write-bas' };
const USAGE = 'Usage: node scripts/probe-vbaproject.js --call [--library <path.catvba>] [<step> ...]\n'
    + '       node scripts/probe-vbaproject.js --write-bas | --info | --install | --remove\n'
    + `  call steps (default: all, in this order): ${CALL_STEPS.join(' ')} (alias: vbaproj-filename)\n`
    + '  --info/--install/--remove run operation.catvbs (loads vbscript.dll); --call and --write-bas never do.';

/** VBA of the module C5D_BridgeTest (ASCII only, CRLF, no Attribute line). Same text for --write-bas and --install. */
const BRIDGE_BODY = [
    MARKER,
    'Option Explicit',
    'Public Function C5D_Ping(ByVal a As Variant) As Variant',
    '    Dim n As Long',
    '    n = 3',
    '    C5D_Ping = "vba:" & TypeName(n) & ":" & CStr(a)',
    'End Function',
    'Public Sub C5D_Noop()',
    'End Sub',
    'Public Function C5D_Projects() As Variant',
    '    On Error Resume Next',
    '    Dim apc As Object',
    '    Dim projects As Object',
    '    Dim s As String',
    '    Dim i As Long',
    '    Set apc = CreateObject("MSAPC.Apc.7.1")',
    '    If apc Is Nothing Then',
    '        Err.Clear',
    '        Set apc = CreateObject("MSAPC.Apc")',
    '    End If',
    '    If apc Is Nothing Then',
    '        C5D_Projects = "apc=nothing;err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    Set projects = apc.VBE.VBProjects',
    '    If projects Is Nothing Then',
    '        C5D_Projects = "projects=nothing;err=" & CStr(Err.Number)',
    '        Exit Function',
    '    End If',
    '    s = "count=" & CStr(projects.Count)',
    '    For i = 1 To projects.Count',
    '        s = s & "|" & projects.Item(i).Name',
    '    Next',
    '    C5D_Projects = s & ";err=" & CStr(Err.Number)',
    'End Function',
    ''
].join('\r\n');

/** { mode, library, steps }. Throws on bad arguments. */
function parseArgs(args) {
    if (args.some(arg => arg === '--help' || arg === '-h')) return { mode: 'help', library: null, steps: [] };
    const modes = [], steps = [];
    let library = null;
    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        if (Object.hasOwn(MODES, arg)) modes.push(MODES[arg]);
        else if (arg === '--library') {
            if (i + 1 >= args.length || !args[i + 1] || args[i + 1].startsWith('--')) throw new Error(`--library requires a path\n${USAGE}`);
            library = args[++i];
        } else if (CALL_STEPS.includes(arg) || Object.hasOwn(CALL_ALIASES, arg)) steps.push(arg);
        else throw new Error(`Unknown option or step: ${arg}\n${USAGE}`);
    }
    if (modes.length !== 1) throw new Error(`Choose exactly one mode (${Object.keys(MODES).join(' ')})\n${USAGE}`);
    const mode = modes[0];
    if (mode !== 'call' && (library !== null || steps.length)) throw new Error(`--library and step names are only for --call\n${USAGE}`);
    const expanded = [...new Set(steps.flatMap(step => CALL_ALIASES[step] ?? [step]))];
    return { mode, library, steps: mode === 'call' ? (expanded.length ? expanded : [...CALL_STEPS]) : [] };
}

/** "count=N|name1|...;err=N" / "apc=nothing;err=N" / "projects=nothing;err=N" -> object (null when the text does not match). */
function parseProjectsResult(text) {
    if (typeof text !== 'string') return null;
    let m = /^count=(-?\d+)((?:\|[^;]*)?);err=(-?\d+)$/.exec(text);
    if (m) return { count: Number(m[1]), names: m[2] ? m[2].slice(1).split('|') : [], err: Number(m[3]) };
    m = /^(apc|projects)=nothing;err=(-?\d+)$/.exec(text);
    if (m) return { count: null, names: [], failedAt: m[1], err: Number(m[2]) };
    return null;
}

/**
 * Verdict of --call. beforeLoaded: vbscript.dll before the first step.
 * entries: [{ step, ok, match, result, hresult, timedOut, loadedAfter }] in run order.
 */
function judgeCall(beforeLoaded, entries) {
    const ping = entries.find(entry => entry.step === 'vbaproj-ping');
    let functional;
    if (!ping) functional = 'ExecuteScript(VBAProject): vbaproj-ping was not run';
    else if (ping.timedOut) functional = 'ExecuteScript(VBAProject) timed out in vbaproj-ping (CATIA may be stopped at a dialog)';
    else if (ping.match) functional = 'ExecuteScript(VBAProject) ran real VBA (Long)';
    else if (ping.ok && typeof ping.result === 'string' && /^vba:Integer:/.test(ping.result)) {
        functional = 'ExecuteScript(VBAProject) ran the code but As Long was ignored (Integer: not real VBA)';
    } else if (ping.ok) functional = `ExecuteScript(VBAProject) returned an unexpected result ${JSON.stringify(ping.result ?? null)}`;
    else functional = `ExecuteScript(VBAProject) failed to call C5D_Ping (${ping.hresult ?? 'no result'})`;
    if (beforeLoaded) {
        return `${functional}; vbscript.dll undetermined: it was already loaded before the first step (restart CATIA, run nothing else, and re-run --call)`;
    }
    const first = entries.find(entry => entry.loadedAfter);
    return first ? `${functional} and loaded vbscript.dll (first seen after step ${first.step})` : `${functional} and did not load vbscript.dll`;
}

/** Library given with --library, else the "library" of temp/probe-vbaproject-info.json. */
function resolveLibrary(given, infoPath = INFO_PATH) {
    if (given) return { library: given, source: '--library' };
    if (!fs.existsSync(infoPath)) throw new Error(`No --library given and ${infoPath} does not exist (run --info, or pass --library <path.catvba>)`);
    const info = JSON.parse(fs.readFileSync(infoPath, 'utf8'));
    if (typeof info.library !== 'string' || !info.library) throw new Error(`${infoPath} has no "library" (FileName of ${PROJECT}); pass --library <path.catvba>`);
    return { library: info.library, source: infoPath };
}

/** Runs one step in its own eval-probe.exe process. A timeout most likely means a modal dialog in CATIA. */
function runCallStep(exe, step, library, project = PROJECT) {
    const args = [step, library, '--project', project];
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

async function runCall(report, options) {
    const { library, source } = resolveLibrary(options.library);
    Object.assign(report, { library, librarySource: source, libraryExists: fs.existsSync(library), steps: options.steps, runs: [] });
    if (!report.libraryExists) report.notes.push(`Library file not found on this PC: ${library} (the steps still run; CATIA decides how to resolve it).`);
    report.exe = buildProbe(ROOT);
    const snap = label => (report.snapshots[label] = moduleSnapshot(label));
    const first = snap('before');
    if (!first.cnextPids.length) throw new Error('CNEXT.exe is not running');
    if (first.cnextPids.length > 1) report.notes.push(`More than one CNEXT.exe is running (${first.cnextPids.join(', ')}); a DLL counts as loaded if any of them has it.`);
    if (first.powershell.error) report.notes.push(`Process.Modules could not be read (${first.powershell.error}); judged by tasklist /M only.`);
    if (first.loaded) report.notes.push('vbscript.dll was already loaded before the first step; restart CATIA (and run nothing else) to judge the dependency. Step results are still valid.');
    const entries = [];
    for (const step of options.steps) {
        const run = await runCallStep(report.exe, step, library);
        report.runs.push(run);
        const after = snap('after:' + step);
        const r = run.result;
        entries.push({ step, ok: r?.ok === true, match: r?.match ?? null, result: r?.result ?? null, hresult: r?.hresult ?? null, timedOut: run.timedOut, loadedAfter: after.loaded });
        if (run.timedOut) {
            report.notes.push(`Step ${step} timed out after ${STEP_TIMEOUT_MS / 1000}s; CATIA may be stopped at a dialog (check CATIA). Remaining steps were not run.`);
            break;
        }
        if (run.exitCode === 10 || run.exitCode === 11) {
            report.notes.push(`Step ${step}: could not reach CATIA (exit ${run.exitCode}); remaining steps were not run.`);
            break;
        }
    }
    const snapshots = Object.values(report.snapshots);
    const projectsRun = report.runs.find(run => run.step === 'vbaproj-projects');
    report.summary = {
        dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, Object.fromEntries(snapshots.map(s => [s.label, s.dlls[dll].loaded]))])),
        modulesAdded: snapshots.length > 1 ? addedModules(snapshots[0], snapshots[snapshots.length - 1]) : undefined,
        steps: report.runs.map(run => ({ step: run.step, library: run.result?.library ?? run.args[1], timedOut: run.timedOut, exitCode: run.exitCode,
            ok: run.result?.ok ?? null, match: run.result?.match ?? null, result: run.result?.result ?? null, hresult: run.result?.hresult ?? null,
            message: run.result?.message ?? null, paramMode: run.result?.paramMode ?? null })),
        libraryResolvedBy: entries.filter(entry => PING_STEPS.includes(entry.step) && entry.match).map(entry => entry.step),
        projects: projectsRun ? parseProjectsResult(projectsRun.result?.result) : undefined,
        verdict: judgeCall(first.loaded, entries)
    };
    if (report.runs.some(run => run.timedOut || !run.result || run.exitCode !== 0)) process.exitCode = 1;
}

// ---- --info / --install / --remove: the existing catvbs route (src/catiaExecution.ts, compiled to out/).

/** One VBScript statement whose value (or Err) is appended to the result's "projects" as "key=value" / "key!err=N text". */
function infoProperty(key, expression) {
    return `
    Err.Clear
    value = Empty
    value = ${expression}
    errN = Err.Number
    errText = Err.Description
    Err.Clear
    C5DInfo "${key}", value, errN, errText`;
}
/** Sub appended to operation.catvbs (kind 'check'): project properties and every component (all types) of gProject. */
const INFO_SUB = `
Sub C5DProjectInfo()
    On Error Resume Next
    Dim value, errN, errText, total, i, component, componentName, componentType
    ${infoProperty('name', 'gProject.Name')}
    ${infoProperty('fileName', 'gProject.FileName')}
    ${infoProperty('protection', 'gProject.Protection')}
    ${infoProperty('saved', 'gProject.Saved')}
    ${infoProperty('type', 'gProject.Type')}
    ${infoProperty('componentCount', 'gComponents.Count')}
    total = value
    If errN <> 0 Then Exit Sub
    For i = 1 To total
        Err.Clear
        Set component = Nothing
        componentName = ""
        componentType = -1
        Set component = gComponents.Item(i)
        componentName = component.Name
        componentType = component.Type
        errN = Err.Number
        errText = Err.Description
        Err.Clear
        C5DInfo "component", componentName & ":" & CStr(componentType), errN, errText
    Next
End Sub

Sub C5DInfo(key, value, errN, errText)
    On Error Resume Next
    If errN = 0 Then
        AppendJson gProjectNames, JsonQuote(key & "=" & CStr(value))
    Else
        AppendJson gProjectNames, JsonQuote(key & "!err=" & CStr(errN) & " " & errText)
    End If
End Sub
`.replace(/\r?\n/g, '\r\n');
const EXPORT_CALL = '\r\n    ExportModules\r\n    FinishResult\r\nEnd Sub\r\n';
/** operation.catvbs of a 'check' request with C5DProjectInfo called after ExportModules. */
function withProjectInfo(source) {
    if (!source.includes(EXPORT_CALL)) throw new Error('operation.catvbs layout changed: ExportModules call not found');
    return source.replace(EXPORT_CALL, '\r\n    ExportModules\r\n    C5DProjectInfo\r\n    FinishResult\r\nEnd Sub\r\n') + INFO_SUB;
}
/** "projects" of the C5DProjectInfo result -> { properties, errors, components: [{ name, type }] }. */
function parseInfo(entries) {
    const properties = {}, errors = {}, components = [];
    for (const entry of entries) {
        const m = /^([A-Za-z]+)(=|!err=)([\s\S]*)$/.exec(entry);
        if (!m) continue;
        if (m[2] === '!err=') { errors[m[1]] = m[3]; if (m[1] === 'component') components.push({ name: null, type: null, error: m[3] }); continue; }
        if (m[1] === 'component') {
            const at = m[3].lastIndexOf(':');
            components.push({ name: m[3].slice(0, at), type: Number(m[3].slice(at + 1)) });
        } else properties[m[1]] = m[3];
    }
    return { properties, errors, components };
}
const isBridge = name => typeof name === 'string' && name.toLowerCase() === MODULE.toLowerCase();

/** Runs requests of one sync operation (the extension's lock); call(kind, configure?, override?) -> { request, result }. */
async function withOperation(report, body) {
    const bridgePath = path.join(ROOT, 'bin', 'catia-run.exe');
    if (!fs.existsSync(bridgePath)) throw new Error(`Missing ${bridgePath}; run npm run compile:bridge`);
    let modules;
    try {
        modules = { ...require('../out/syncOperation'), ...require('../out/catiaExecution') };
    } catch (error) { throw new Error(`out/ is missing or stale; run npm run compile (${error.message})`); }
    const iconv = require('iconv-lite');
    const context = new modules.SyncOperationManager().acquire('probe-vbaproject');
    report.requests = [];
    const call = async (kind, project, configure, override) => {
        const request = modules.createCatiaRequest(context, kind, project, ENCODING);
        if (configure) configure(request);
        const execute = override ? (directory, scriptName, timeoutMs) => {
            const file = path.join(directory, scriptName);
            fs.writeFileSync(file, override(fs.readFileSync(file, 'utf8')));
            return modules.bridgeExecutor(bridgePath)(directory, scriptName, timeoutMs);
        } : undefined;
        const entry = { kind, project, modules: request.modules, deletions: request.deletions };
        report.requests.push(entry);
        try {
            const result = await modules.runCatiaRequest(context, request, { bridgePath, execute });
            Object.assign(entry, { expected: result.expectedCount, succeeded: result.succeededCount, failures: result.failures });
            return { request, result };
        } catch (error) {
            Object.assign(entry, { error: String(error), failures: error.result?.failures });
            throw error;
        }
    };
    /** 'check' of toQI + C5DProjectInfo: { info, items: [{ name, type, text }] }. */
    const inspect = async () => {
        const list = await call('list', '');
        if (!list.result.projects.includes(PROJECT)) throw new Error(`Project ${PROJECT} not found (projects: ${list.result.projects.join(', ')})`);
        const { request, result } = await call('check', PROJECT, undefined, withProjectInfo);
        const items = result.items.map(item => ({ name: item.name, type: item.type,
            text: iconv.decode(fs.readFileSync(path.join(request.directory, item.file)), ENCODING) }));
        return { projects: list.result.projects, info: parseInfo(result.projects), items };
    };
    try { return await body({ call, inspect, iconv }); } finally {
        try { context.close(); } catch (error) { report.notes.push(`Lock release failed: ${error}`); }
    }
}
function describeInspection(inspection) {
    const bridgeComponent = inspection.info.components.find(c => isBridge(c.name));
    const bridgeItem = inspection.items.find(item => isBridge(item.name));
    return {
        projects: inspection.projects,
        project: PROJECT,
        library: inspection.info.properties.fileName ?? null,
        properties: inspection.info.properties,
        errors: inspection.info.errors,
        components: inspection.info.components,
        bridgePresent: Boolean(bridgeComponent || bridgeItem),
        bridgeType: bridgeComponent?.type ?? bridgeItem?.type ?? null,
        bridgeOwned: bridgeItem ? bridgeItem.text.includes(MARKER) : null
    };
}
const { createHash } = require('node:crypto');
const hashes = items => Object.fromEntries(items.filter(item => !isBridge(item.name))
    .map(item => [item.name, createHash('sha256').update(item.text).digest('hex')]).sort(([a], [b]) => a.localeCompare(b)));
/** Names of modules whose text differs from (or is missing in / new to) the baseline. */
function changedModules(baseline, current) {
    return [...new Set([...Object.keys(baseline), ...Object.keys(current)])].filter(name => baseline[name] !== current[name]).sort();
}
const normalize = text => text.replace(/\r\n?/g, '\n').trimEnd();

async function runInfo(report) {
    await withOperation(report, async ({ inspect }) => {
        const info = describeInspection(await inspect());
        report.info = info;
        fs.writeFileSync(INFO_PATH, JSON.stringify({ ...info, collectedAt: new Date().toISOString() }, null, 2));
        report.infoPath = INFO_PATH;
        if (String(info.properties.protection) !== '0') report.notes.push(`Protection=${info.properties.protection}: the project may be locked (1 = vbext_pp_locked).`);
        report.summary = { verdict: `${PROJECT}: library=${info.library ?? '(unknown)'}, ${MODULE} ${info.bridgePresent ? 'present' : 'absent'}` };
    });
}
async function runInstall(report) {
    await withOperation(report, async ({ call, inspect, iconv }) => {
        const before = await inspect();
        const info = describeInspection(before);
        report.before = info;
        if (info.bridgePresent) throw new Error(`${MODULE} already exists in ${PROJECT} (owner marker: ${info.bridgeOwned}); not installing`);
        if (String(info.properties.protection) !== '0') throw new Error(`${PROJECT} Protection=${info.properties.protection}; not installing`);
        fs.writeFileSync(BASELINE_PATH, JSON.stringify({ project: PROJECT, at: new Date().toISOString(), hashes: hashes(before.items) }, null, 2));
        report.baselinePath = BASELINE_PATH;
        await call('push', PROJECT, request => {
            fs.writeFileSync(path.join(request.directory, 'm0.txt'), iconv.encode(BRIDGE_BODY, ENCODING));
            request.modules = [{ name: MODULE, type: 1, file: 'm0.txt' }];
        });
        const after = await inspect();
        const item = after.items.find(i => isBridge(i.name));
        report.installed = Boolean(item);
        report.bodyMatches = item ? normalize(item.text) === normalize(BRIDGE_BODY) : false;
        report.otherModulesChanged = changedModules(hashes(before.items), hashes(after.items));
        if (item && !report.bodyMatches) report.installedText = item.text;
        fs.writeFileSync(INFO_PATH, JSON.stringify({ ...describeInspection(after), collectedAt: new Date().toISOString() }, null, 2));
        report.summary = { verdict: report.installed ? `${MODULE} added to ${PROJECT}${report.bodyMatches ? '' : ' (WARNING: text differs from the source)'}`
            : `${MODULE} not found after push` };
        if (!report.installed || !report.bodyMatches) process.exitCode = 1;
    });
}
async function runRemove(report) {
    await withOperation(report, async ({ call, inspect }) => {
        const before = await inspect();
        const info = describeInspection(before);
        report.before = info;
        if (!info.bridgePresent) { report.summary = { verdict: `${MODULE} is not in ${PROJECT}; nothing removed` }; return; }
        if (info.bridgeType !== 1) throw new Error(`${MODULE} is not a standard module (type ${info.bridgeType}); refusing to remove`);
        if (!info.bridgeOwned) throw new Error(`${MODULE} has no owner marker (${MARKER}); refusing to remove`);
        await call('push', PROJECT, request => { request.modules = []; request.deletions = [MODULE]; });
        const after = await inspect();
        report.removed = !describeInspection(after).bridgePresent;
        if (fs.existsSync(BASELINE_PATH)) {
            const baseline = JSON.parse(fs.readFileSync(BASELINE_PATH, 'utf8'));
            report.otherModulesChangedSinceInstall = changedModules(baseline.hashes, hashes(after.items));
        } else report.notes.push(`${BASELINE_PATH} not found (module was not added by --install); other modules were not compared.`);
        report.otherModulesChanged = changedModules(hashes(before.items), hashes(after.items));
        report.summary = { verdict: report.removed ? `${MODULE} removed from ${PROJECT} (project not saved)` : `${MODULE} still present after deletion` };
        if (!report.removed) process.exitCode = 1;
    });
}
function runWriteBas(report) {
    fs.writeFileSync(BAS_PATH, BRIDGE_BODY, 'ascii');
    report.basPath = BAS_PATH;
    report.summary = { verdict: `wrote ${BAS_PATH} (${BRIDGE_BODY.split('\r\n').length - 1} lines, ASCII)` };
    report.notes.push(`Paste it into a new standard module of ${PROJECT} named ${MODULE} (replace everything, incl. an auto-inserted Option Explicit). Do not save the project.`);
}

/** temp/probe-vbaproject-<mode>-<time>.json; never overwrites an earlier report. */
function uniqueReportPath(directory, mode) {
    const base = path.join(directory, `probe-vbaproject-${mode}-${timestamp()}`);
    let candidate = base + '.json';
    for (let i = 2; fs.existsSync(candidate); i++) candidate = `${base}-${i}.json`;
    return candidate;
}
function formatSummary(report) {
    const yn = value => value === true ? 'yes' : value === false ? 'no' : '-';
    const lines = [`mode: ${report.mode}`, `report: ${report.reportPath}`];
    if (report.library) lines.push(`library: ${report.library} (from ${report.librarySource}${report.libraryExists ? '' : '; NOT FOUND on disk'})`);
    if (report.error) lines.push(`error: ${report.error}`);
    const snapshots = Object.values(report.snapshots);
    if (snapshots.length) {
        const first = snapshots[0], last = snapshots[snapshots.length - 1];
        const readable = first.processes.filter(p => !p.error && p.moduleCount > 0);
        lines.push(`CNEXT.exe pids: ${first.cnextPids.join(', ') || '(none)'}`);
        lines.push(`Process.Modules: ${first.powershell.error ? 'ERROR ' + first.powershell.error
            : readable.length ? readable.map(p => `pid ${p.pid}: ${p.moduleCount} modules`).join(', ') : 'no readable process'}`);
        const short = label => label.replace(/^after:vbaproj-/, '+');
        const width = Math.max(...snapshots.map(s => short(s.label).length), 6) + 2;
        lines.push('', 'dll'.padEnd(14) + snapshots.map(s => short(s.label).padEnd(width)).join(''));
        for (const dll of KEY_DLLS) {
            lines.push(dll.padEnd(14) + snapshots.map(s => {
                const d = s.dlls[dll];
                const detail = d.byModules === null ? '(T)' : d.byModules !== d.byTasklist ? `(M:${yn(d.byModules)},T:${yn(d.byTasklist)})` : '';
                return (yn(d.loaded) + detail).padEnd(width);
            }).join(''));
        }
        lines.push('  (+x = after step vbaproj-x; yes/no = loaded in CNEXT.exe by Process.Modules or tasklist /M; (T) = tasklist only)');
        if (snapshots.length > 1) lines.push(`modules added ${first.label} -> ${last.label}: ${addedModules(first, last).join(', ') || '(none)'}`);
    }
    if (report.runs?.length) {
        lines.push('', 'steps:');
        for (const run of report.runs) {
            const r = run.result;
            if (run.timedOut) lines.push(`  ${run.step}: TIMEOUT after ${STEP_TIMEOUT_MS / 1000}s - CATIA may be stopped at a dialog (check CATIA; later steps were skipped)`);
            else if (!r) lines.push(`  ${run.step}: no JSON output (exitCode=${run.exitCode}${run.error ? ', ' + run.error : ''}${run.stderr ? ', stderr: ' + run.stderr : ''})`);
            else {
                lines.push(`  ${run.step}: ok=${yn(r.ok)} match=${yn(r.match)} result=${JSON.stringify(r.result)} (${r.resultType}) paramMode=${r.paramMode} ${r.elapsedMs}ms`);
                lines.push(`    ExecuteScript(${JSON.stringify(r.library)}, ${r.libraryType}, ${JSON.stringify(r.program)}, ${JSON.stringify(r.function)})`);
                if (!r.ok || r.hresult !== '0x00000000') lines.push(`    hresult=${r.hresult} message=${r.message}`);
                if (r.attempts?.length > 1) lines.push(`    attempts: ${r.attempts.join(' / ')}`);
            }
        }
    }
    const s = report.summary;
    if (s?.libraryResolvedBy) lines.push('', `library forms that reached C5D_Ping: ${s.libraryResolvedBy.join(', ') || '(none)'}`);
    if (s?.projects) lines.push(`C5D_Projects: ${JSON.stringify(s.projects)}`);
    const info = report.info ?? report.before;
    if (info) {
        lines.push('', `${info.project}: FileName=${info.library ?? '?'} Protection=${info.properties.protection ?? '?'} Saved=${info.properties.saved ?? '?'}`);
        lines.push(`  components: ${info.components.map(c => c.name ? `${c.name}(${c.type})` : `!${c.error}`).join(', ') || '(none)'}`);
        lines.push(`  ${MODULE}: ${info.bridgePresent ? `present (type ${info.bridgeType}, owner marker ${yn(info.bridgeOwned)})` : 'absent'}`);
        if (Object.keys(info.errors).length) lines.push(`  errors: ${JSON.stringify(info.errors)}`);
    }
    if (report.otherModulesChanged) lines.push(`other modules changed by this run: ${report.otherModulesChanged.join(', ') || '(none)'}`);
    if (report.otherModulesChangedSinceInstall) lines.push(`other modules changed since --install: ${report.otherModulesChangedSinceInstall.join(', ') || '(none)'}`);
    for (const request of report.requests ?? []) {
        if (request.error) lines.push(`request ${request.kind} failed: ${request.error}${request.failures?.length ? ' ' + JSON.stringify(request.failures) : ''}`);
    }
    if (s?.verdict) lines.push('', `verdict: ${s.verdict}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    return lines.join('\n');
}

async function main() {
    let options;
    try { options = parseArgs(process.argv.slice(2)); } catch (error) { console.error(error.message); process.exitCode = 2; return; }
    if (options.mode === 'help') { console.log(USAGE); return; }
    fs.mkdirSync(TEMP, { recursive: true });
    const reportPath = uniqueReportPath(TEMP, options.mode);
    const report = { mode: options.mode, startedAt: new Date().toISOString(), reportPath, project: PROJECT, module: MODULE, notes: [], snapshots: {} };
    try {
        if (options.mode === 'call') await runCall(report, options);
        else if (options.mode === 'write-bas') runWriteBas(report);
        else if (options.mode === 'info') await runInfo(report);
        else if (options.mode === 'install') await runInstall(report);
        else await runRemove(report);
    } catch (error) {
        report.error = String(error); process.exitCode = 1;
    }
    report.finishedAt = new Date().toISOString();
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(formatSummary(report));
}
if (require.main === module) main();
module.exports = { parseArgs, judgeCall, parseProjectsResult, resolveLibrary, withProjectInfo, parseInfo, changedModules,
    BRIDGE_BODY, MARKER, CALL_STEPS, INFO_SUB };
