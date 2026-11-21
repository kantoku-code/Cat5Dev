// B0 probe (developer tool, not shipped): does CATScript need vbscript.dll, and which syntax does it accept?
// Usage: npm run compile && npm run compile:bridge && node scripts/probe-catscript.js [mode]
// Read-only for CATIA: only reads VBProjects; never adds/changes modules or projects.
// Modes (staged isolation; the .catvbs control loads vbscript.dll for sure, so it only runs in the default mode):
//   --check-only      run nothing; record which script DLLs CNEXT.exe has loaded now
//   --no-apc          probe.CATScript without the apc/vbe sections; no generated list, no .catvbs
//   --apc             probe.CATScript with apc/vbe + generated list as .CATScript; no .catvbs (alias: --full-catscript)
//   (none)            legacy: probe.CATScript + generated list + probe.catvbs control
// Modules are read with 64-bit PowerShell (Process.Modules); `tasklist /M <dll>` is a secondary check
// (the full `tasklist /M` text listing was seen to miss modules). Each run writes
// temp/probe-catscript-<mode>-<yyyyMMdd-HHmmss>.json and prints a summary to stdout.
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const RESULT_NAME = 'probe-result.txt';
const q = text => `"${text.replace(/"/g, '""')}"`;
/** Same error-capture style as buildCatiaScript: Err.Clear, statement, copy Err, Err.Clear. */
function probe(statement, name, value = '""') {
    return `
    Err.Clear
    ${statement}
    errN = Err.Number
    errText = Err.Description
    Err.Clear
    If errN = 0 Then
        Note ${q(name)}, 0, ${value}
    Else
        Note ${q(name)}, errN, errText
    End If`;
}
// Each section is one Sub; deps are prepended when a section is run alone.
const SECTIONS = {
    fso: { deps: [], code: `
Sub ProbeFso()
    On Error Resume Next
    Dim errN, errText, joined
    ${probe('Set gFSO = CreateObject("Scripting.FileSystemObject")', 'fso.create')}
    If gFSO Is Nothing Then Exit Sub
    ${probe('joined = gFSO.BuildPath(gDirectory, "x.txt")', 'fso.buildpath', 'joined')}
End Sub` },
    stream: { deps: ['fso'], code: `
Sub ProbeStream()
    On Error Resume Next
    Dim errN, errText, stream, temporaryPath, finalPath, text, readBack
    If gFSO Is Nothing Then
        Note "stream", 91, "FSO unavailable"
        Exit Sub
    End If
    text = "probe-" & ChrW(26085) & ChrW(26412) & ChrW(-24691)
    temporaryPath = gFSO.BuildPath(gDirectory, "stream.tmp")
    finalPath = gFSO.BuildPath(gDirectory, "stream.txt")
    ${probe('Set stream = CreateObject("ADODB.Stream")', 'stream.create')}
    ${probe('stream.Type = 2', 'stream.type')}
    ${probe('stream.Charset = "utf-8"', 'stream.charset')}
    ${probe('stream.Open', 'stream.open')}
    ${probe('stream.WriteText text', 'stream.write')}
    ${probe('stream.SaveToFile temporaryPath, 2', 'stream.save')}
    ${probe('stream.Close', 'stream.close')}
    ${probe('gFSO.MoveFile temporaryPath, finalPath', 'fso.movefile')}
    ${probe('Set stream = CreateObject("ADODB.Stream")', 'stream.read.create')}
    ${probe('stream.Type = 2', 'stream.read.type')}
    ${probe('stream.Charset = "shift_jis"', 'stream.read.charset.sjis')}
    ${probe('stream.Charset = "utf-8"', 'stream.read.charset')}
    ${probe('stream.Open', 'stream.read.open')}
    ${probe('stream.LoadFromFile finalPath', 'stream.read.load')}
    ${probe('readBack = stream.ReadText(-1)', 'stream.read.text', 'CStr(Len(readBack))')}
    ${probe('stream.Close', 'stream.read.close')}
    If readBack = text Then
        Note "stream.roundtrip", 0, "equal"
    Else
        Note "stream.roundtrip", 1, "different"
    End If
End Sub` },
    apc: { deps: [], code: `
Sub ProbeApc()
    On Error Resume Next
    Set gApc = Nothing
    TryApc "MSAPC.Apc.7.1"
    If gApc Is Nothing Then TryApc "MSAPC.Apc.6.2"
    If gApc Is Nothing Then TryApc "MSAPC.Apc"
    If gApc Is Nothing Then Note "apc", 429, "No APC object"
End Sub

Sub TryApc(progId)
    On Error Resume Next
    Dim errN, errText
    Err.Clear
    Set gApc = CreateObject(progId)
    errN = Err.Number
    errText = Err.Description
    Err.Clear
    If errN <> 0 Then Set gApc = Nothing
    Note "apc.create." & progId, errN, errText
End Sub` },
    vbe: { deps: ['apc'], code: `
Sub ProbeVbe()
    On Error Resume Next
    Dim errN, errText, projects, total, candidate, candidateName
    If gApc Is Nothing Then
        Note "vbe", 91, "APC unavailable"
        Exit Sub
    End If
    total = -1
    ${probe('total = gApc.VBE.VBProjects.Count', 'vbe.chain.count', 'CStr(total)')}
    ${probe('Set projects = gApc.VBE.VBProjects', 'vbe.projects')}
    ${probe('total = projects.Count', 'vbe.count', 'CStr(total)')}
    If total > 0 Then
        ${probe('Set candidate = projects.Item(1)', 'vbe.item')}
        ${probe('candidateName = candidate.Name', 'vbe.item.name', 'CStr(Len(candidateName))')}
    End If
End Sub` },
    objfunc: { deps: ['fso'], code: `
Function ProbeReturnObject(value)
    On Error Resume Next
    Set ProbeReturnObject = Nothing
    If value Is Nothing Then Exit Function
    Set ProbeReturnObject = value
End Function

Function ProbeReturnBoolean()
    ProbeReturnBoolean = False
    ProbeReturnBoolean = True
End Function

Sub ProbeObjectFunction()
    On Error Resume Next
    Dim errN, errText, found, flag
    Set found = Nothing
    ${probe('Set found = ProbeReturnObject(gFSO)', 'objfunc.object', 'CStr(Not (found Is Nothing))')}
    ${probe('Set found = ProbeReturnObject(Nothing)', 'objfunc.nothing', 'CStr(found Is Nothing)')}
    ${probe('flag = ProbeReturnBoolean()', 'func.boolean', 'CStr(flag)')}
    flag = False
    ${probe('If Not ProbeReturnBoolean() Then Exit Sub\n    flag = True', 'func.ifnot', 'CStr(flag)')}
End Sub` },
    byref: { deps: [], code: `
Sub ProbeAppend(ByRef buffer, value)
    If Len(buffer) > 0 Then buffer = buffer & ","
    buffer = buffer & value
End Sub

Function ProbeLookup(moduleName, ByRef lookupSucceeded)
    On Error Resume Next
    lookupSucceeded = False
    Set ProbeLookup = Nothing
    lookupSucceeded = True
End Function

Function ProbeRead(fileName, ByRef sourceText)
    ProbeRead = False
    sourceText = "read:" & fileName
    ProbeRead = True
End Function

Sub ProbeByRef()
    On Error Resume Next
    Dim errN, errText, buffer, lookupSucceeded, found, sourceText
    buffer = ""
    gBuffer = ""
    ${probe('ProbeAppend buffer, "a"\n    ProbeAppend buffer, "b"', 'byref.local', 'buffer')}
    ${probe('ProbeAppend gBuffer, "a"\n    ProbeAppend gBuffer, "b"', 'byref.global', 'gBuffer')}
    lookupSucceeded = False
    ${probe('Set found = ProbeLookup("m", lookupSucceeded)', 'byref.function', 'CStr(lookupSucceeded) & "/" & CStr(found Is Nothing)')}
    sourceText = ""
    ${probe('If Not ProbeRead("f", sourceText) Then Exit Sub', 'byref.ifnot', 'sourceText')}
End Sub` },
    builtins: { deps: [], code: `
Sub ProbeBuiltins()
    On Error Resume Next
    Dim errN, errText, value, code, i
    ${probe('value = ChrW(-24691)\n    code = AscW(value)', 'builtin.chrw.negative', 'CStr(code)')}
    ${probe('code = AscW(ChrW(26085))', 'builtin.ascw', 'CStr(code)')}
    ${probe('value = Right("0000" & Hex(255), 4)', 'builtin.hex.right', 'value')}
    ${probe('value = Mid("abc", 2, 1) & CStr(Len("abc")) & Chr(92)', 'builtin.mid.len.chr', 'value')}
    ${probe('code = StrComp("a", "A", vbBinaryCompare)', 'builtin.strcomp.binary', 'CStr(code)')}
    ${probe('code = StrComp("a", "A", vbTextCompare)', 'builtin.strcomp.text', 'CStr(code)')}
    ${probe(`code = 92
    Select Case code
        Case 34
            value = "quote"
        Case 92
            value = "backslash"
        Case Else
            value = "other"
    End Select`, 'builtin.selectcase', 'value')}
    ${probe('value = "a" & _\n        "b"', 'builtin.continuation', 'value')}
    ${probe('value = ""\n    For i = 1 To 3\n        value = value & CStr(i)\n    Next', 'builtin.for', 'value')}
End Sub` },
    err: { deps: [], code: `
Sub ProbeErr()
    On Error Resume Next
    Dim errN, errText, missing
    ${probe('Set missing = CreateObject("Cat5Dev.NoSuchProgId")', 'err.createobject')}
    Err.Clear
    Note "err.clear", Err.Number, Err.Description
End Sub` }
};
const ENTRY = { fso: 'ProbeFso', stream: 'ProbeStream', apc: 'ProbeApc', vbe: 'ProbeVbe', objfunc: 'ProbeObjectFunction', byref: 'ProbeByRef', builtins: 'ProbeBuiltins', err: 'ProbeErr' };
function expand(names) {
    const out = [];
    const add = name => { for (const dep of SECTIONS[name].deps) add(dep); if (!out.includes(name)) out.push(name); };
    names.forEach(add); return out;
}
/** Writes results line by line (name TAB Err.Number TAB detail) with ADODB.Stream as UTF-8, FSO as fallback. */
function buildProbeScript(directory, names = Object.keys(SECTIONS)) {
    const { vbString } = require('../out/catiaScripts');
    const sections = expand(names);
    return `Option Explicit
Dim gFSO, gApc, gLog, gDirectory, gResultPath, gBuffer

Sub CATMain()
    On Error Resume Next
    gLog = ""
    Set gFSO = Nothing
    Set gApc = Nothing
    gDirectory = ${vbString(directory)}
    gResultPath = ${vbString(path.join(directory, RESULT_NAME))}
    Note "start", 0, ${q(sections.join(','))}
    ${sections.map(name => ENTRY[name]).join('\n    ')}
    FinishProbe
End Sub

Sub Note(stepName, errorNumber, detail)
    gLog = gLog & stepName & Chr(9) & CStr(errorNumber) & Chr(9) & CStr(detail) & Chr(13) & Chr(10)
End Sub
${sections.map(name => SECTIONS[name].code).join('\n')}

Sub FinishProbe()
    On Error Resume Next
    Dim stream, file, errN, errText
    Note "finish", 0, ""
    Err.Clear
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText gLog
    stream.SaveToFile gResultPath, 2
    stream.Close
    If Err.Number = 0 Then Exit Sub
    errN = Err.Number
    errText = Err.Description
    Err.Clear
    If gFSO Is Nothing Then Set gFSO = CreateObject("Scripting.FileSystemObject")
    Set file = gFSO.CreateTextFile(gResultPath & ".fso.txt", True, True)
    file.Write gLog & "finish.stream" & Chr(9) & CStr(errN) & Chr(9) & errText
    file.Close
End Sub
`.replace(/\r?\n/g, '\r\n');
}
function parseSteps(text) {
    const steps = [];
    for (const line of text.replace(/^﻿/, '').split(/\r?\n/)) {
        const m = /^([^\t]+)\t(-?\d+)\t(.*)$/.exec(line);
        if (m) steps.push({ name: m[1], err: Number(m[2]), detail: m[3] });
        else if (line && steps.length) steps[steps.length - 1].detail += '\n' + line;
    }
    return steps;
}
function decode(buffer) {
    const text = buffer.toString('utf8');
    if (!text.includes('�')) return text;
    try {
        const iconv = require('iconv-lite');
        const cp = /(\d+)\D*$/.exec(execFileSync('cmd', ['/c', 'chcp'], { encoding: 'latin1' }));
        const name = cp ? 'cp' + cp[1] : 'cp932';
        if (iconv.encodingExists(name)) return iconv.decode(buffer, name);
    } catch { /* fall through */ }
    return buffer.toString('latin1');
}
/** DLLs judged in every snapshot (lower case). vbscript.dll drives the verdict. */
const KEY_DLLS = ['vbscript.dll', 'jscript.dll', 'scrrun.dll', 'vbe7.dll', 'apc71.dll'];
/** Other modules recorded with their paths for reference (CATScript engine, VBA/APC runtime, WSH ...). */
const RELATED = /script|vbs|vba|vbe|apc|msvbvm|scrrun|wsh/i;
/** `tasklist /FI "IMAGENAME eq CNEXT.exe" [/M <dll>]`; pids of the matching CNEXT.exe rows. */
function tasklist(moduleName) {
    const args = ['/FI', 'IMAGENAME eq CNEXT.exe', ...(moduleName ? ['/M', moduleName] : []), '/FO', 'CSV', '/NH'];
    const raw = decode(execFileSync('tasklist', args, { windowsHide: true })).trim();
    const pids = raw.split(/\r?\n/).filter(line => /^"CNEXT\.exe"/i.test(line)).map(line => Number(line.split('","')[1]));
    return { command: 'tasklist ' + args.map(a => a.includes(' ') ? `"${a}"` : a).join(' '), raw, pids };
}
/** 64-bit Windows PowerShell even from a 32-bit node (Sysnative is visible only to WOW64 processes). */
function powershellPath() {
    const root = process.env.SystemRoot || 'C:\\Windows';
    const sysnative = path.join(root, 'Sysnative', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
    return fs.existsSync(sysnative) ? sysnative : path.join(root, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
}
const asArray = value => value == null ? [] : Array.isArray(value) ? value : [value];
/** (Get-Process -Name <name>).Modules for every matching process: { exe, is64, processes: [{ pid, path, error, modules: [{ name, path }] }], error }. */
function listModules(processName) {
    const exe = powershellPath();
    const script = [
        "$ErrorActionPreference = 'Stop'",
        '[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false',
        `$processes = @(Get-Process -Name '${processName.replace(/'/g, "''")}' -ErrorAction SilentlyContinue)`,
        '$items = @()',
        'foreach ($p in $processes) {',
        '    $item = @{ pid = $p.Id; path = $null; error = $null; modules = @() }',
        '    try { $item.path = $p.Path } catch { }',
        '    try { $item.modules = @($p.Modules | ForEach-Object { @{ name = $_.ModuleName; path = $_.FileName } }) } catch { $item.error = $_.Exception.Message }',
        '    $items += $item',
        '}',
        'ConvertTo-Json -InputObject @{ is64 = [Environment]::Is64BitProcess; processes = $items } -Depth 6 -Compress'
    ].join('\n');
    try {
        const stdout = execFileSync(exe, ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')],
            { windowsHide: true, timeout: 60000, maxBuffer: 64 * 1024 * 1024, stdio: ['ignore', 'pipe', 'pipe'] });
        const parsed = JSON.parse(stdout.toString('utf8').replace(/^\uFEFF/, '').trim());
        const processes = asArray(parsed.processes).map(p => ({ pid: p.pid, path: p.path ?? null, error: p.error ?? null, modules: asArray(p.modules) }));
        return { exe, is64: parsed.is64 === true, processes, error: parsed.is64 === true ? null : 'PowerShell is not 64-bit; modules of 64-bit CNEXT.exe may be incomplete' };
    } catch (error) {
        return { exe, is64: null, processes: [], error: String(error.stderr ? decode(error.stderr).trim() || error : error) };
    }
}
/** Which KEY_DLLS CNEXT.exe has loaded: Process.Modules first, `tasklist /M <dll>` as a second source (either one => loaded). */
function moduleSnapshot(label) {
    const all = tasklist();
    const ps = listModules('CNEXT');
    const processes = ps.processes.map(p => ({
        pid: p.pid, path: p.path, error: p.error, moduleCount: p.modules.length,
        related: p.modules.filter(m => RELATED.test(m.name)).map(m => ({ name: m.name, path: m.path })),
        moduleNames: p.modules.map(m => m.name)
    }));
    const readable = processes.filter(p => !p.error && p.moduleCount > 0);
    const dlls = {};
    for (const dll of KEY_DLLS) {
        const listed = tasklist(dll);
        const byModules = readable.length ? readable.some(p => p.moduleNames.some(name => name.toLowerCase() === dll)) : null;
        const byTasklist = listed.pids.length > 0;
        dlls[dll] = { loaded: byModules === true || byTasklist, byModules, byTasklist, tasklistPids: listed.pids, tasklistCommand: listed.command };
    }
    const cnextPids = [...new Set([...all.pids, ...processes.map(p => p.pid)])];
    return { label, at: new Date().toISOString(), cnextPids, powershell: { exe: ps.exe, is64: ps.is64, error: ps.error }, processes, dlls, loaded: dlls['vbscript.dll'].loaded };
}
async function runScript(bridge, directory, scriptName, source, timeoutMs = 60000) {
    const { bridgeExecutor } = require('../out/catiaExecution');
    fs.mkdirSync(directory, { recursive: true });
    fs.writeFileSync(path.join(directory, scriptName), source, 'utf8');
    const started = Date.now();
    const output = await bridgeExecutor(bridge)(directory, scriptName, timeoutMs);
    const run = { script: path.join(directory, scriptName), exitCode: output.error ? output.error.code ?? null : 0,
        error: output.error ? output.error.message : null, stdout: output.stdout.trim(), stderr: output.stderr.trim(), durationMs: Date.now() - started };
    for (const name of [RESULT_NAME, RESULT_NAME + '.fso.txt']) {
        const file = path.join(directory, name);
        if (fs.existsSync(file)) { run.resultFile = name; run.steps = parseSteps(fs.readFileSync(file, name === RESULT_NAME ? 'utf8' : 'utf16le')); break; }
    }
    run.resultFound = Boolean(run.resultFile);
    const streamFile = path.join(directory, 'stream.txt');
    if (fs.existsSync(streamFile)) run.streamFile = fs.readFileSync(streamFile).toString('hex');
    return run;
}
const MODE_FLAGS = { '--check-only': 'check-only', '--no-apc': 'no-apc', '--apc': 'apc', '--full-catscript': 'apc' };
const USAGE = 'Usage: node scripts/probe-catscript.js [--check-only | --no-apc | --apc | --full-catscript]\n'
    + '  (no option: legacy run = probe.CATScript + generated list + probe.catvbs control)';
const NO_APC_SECTIONS = Object.keys(SECTIONS).filter(name => !['apc', 'vbe'].includes(name));
/** Steps that fail on purpose (name -> expected Err.Number); not reported as failures. */
const EXPECTED_ERRORS = { 'err.createobject': 429 };
function parseMode(args) {
    if (args.some(arg => arg === '--help' || arg === '-h')) return 'help';
    const unknown = args.filter(arg => !Object.hasOwn(MODE_FLAGS, arg));
    if (unknown.length) throw new Error(`Unknown option: ${unknown.join(' ')}\n${USAGE}`);
    const modes = [...new Set(args.map(arg => MODE_FLAGS[arg]))];
    if (modes.length > 1) throw new Error(`Choose only one mode: ${args.join(' ')}\n${USAGE}`);
    return modes[0] ?? 'all';
}
/** Verdict about vbscript.dll from the per-snapshot load state (label -> boolean). */
function judge(mode, loaded) {
    if (mode === 'check-only') {
        return loaded.current ? 'vbscript.dll is already loaded in CNEXT.exe (no script was run)'
            : 'vbscript.dll is not loaded in CNEXT.exe (no script was run)';
    }
    if (loaded.before) return 'undetermined: vbscript.dll was already loaded before the run (restart CATIA and re-run)';
    if (mode === 'no-apc') {
        return loaded.afterCatScript ? 'CATScript engine (without APC/VBE) loaded vbscript.dll'
            : 'CATScript engine did not load vbscript.dll';
    }
    if (mode === 'apc') {
        return loaded.afterProbe ? 'CATScript with APC/VBE access loaded vbscript.dll'
            : loaded.afterGeneratedList ? 'generated list CATScript loaded vbscript.dll (the APC/VBE probe did not)'
                : 'CATScript with APC/VBE access (incl. generated list) did not load vbscript.dll';
    }
    return loaded.afterCatScript ? 'CATScript loads vbscript.dll (B2)'
        : loaded.afterCatvbs ? 'CATScript does not load vbscript.dll; catvbs does (B1 viable)'
            : 'inconclusive: control catvbs did not load vbscript.dll';
}
function summarizeRun(name, run) {
    const failedSteps = (run.steps || []).filter(step => step.err !== 0 && EXPECTED_ERRORS[step.name] !== step.err).map(step => `${step.name}(${step.err})`);
    const summary = { name, script: path.basename(run.script), exitCode: run.exitCode, error: run.error, resultFound: run.resultFound, failedSteps };
    if (run.listResult) summary.listFailures = run.listResult.failures.map(f => `${f.phase}(${f.code}) ${f.target}: ${f.message}`);
    if (run.listResultError) summary.listResultError = run.listResultError;
    return summary;
}
/** probe.CATScript with the given sections; when it writes no result, each section is run alone. */
async function runProbe(report, bridge, work, names) {
    const directory = path.join(work, 'catscript-full');
    const full = await runScript(bridge, directory, 'probe.CATScript', buildProbeScript(directory, names));
    report.catscript.full = full;
    if (!full.resultFound) {
        report.notes.push('Full CATScript probe produced no result; each section was run alone to isolate the failing syntax.');
        report.catscript.sections = {};
        for (const name of names) {
            const sectionDirectory = path.join(work, 'catscript-' + name);
            report.catscript.sections[name] = await runScript(bridge, sectionDirectory, 'probe.CATScript', buildProbeScript(sectionDirectory, [name]));
        }
    }
    return full;
}
/** The real generated script (read-only 'list') renamed to .CATScript. */
async function runGeneratedList(report, bridge, work) {
    const { buildCatiaScript } = require('../out/catiaScripts');
    const { validateCatiaResult } = require('../out/catiaTypes');
    const listDirectory = path.join(work, 'catscript-list');
    fs.mkdirSync(listDirectory);
    const request = { operationId: 'probe', stepId: 'list', kind: 'list', targetProject: '', encoding: 'shift_jis', directory: listDirectory };
    const list = await runScript(bridge, listDirectory, 'operation.CATScript', buildCatiaScript(request));
    try {
        const result = validateCatiaResult(fs.readFileSync(path.join(listDirectory, 'result.json'), 'utf8'), request);
        list.listResult = { expectedCount: result.expectedCount, succeededCount: result.succeededCount, failures: result.failures, projectCount: result.projects.length };
    } catch (error) { list.listResultError = String(error); }
    report.catscript.generatedList = list;
    return list;
}
function timestamp(date = new Date()) {
    const p = n => String(n).padStart(2, '0');
    return `${date.getFullYear()}${p(date.getMonth() + 1)}${p(date.getDate())}-${p(date.getHours())}${p(date.getMinutes())}${p(date.getSeconds())}`;
}
/** temp/probe-catscript-<mode>-<time>.json; never overwrites an earlier report. */
function uniqueReportPath(directory, mode) {
    const base = path.join(directory, `probe-catscript-${mode}-${timestamp()}`);
    let candidate = base + '.json';
    for (let i = 2; fs.existsSync(candidate); i++) candidate = `${base}-${i}.json`;
    return candidate;
}
/** Module names (lower case) present in the last snapshot but not in the first one. */
function addedModules(first, last) {
    const names = snapshot => new Set(snapshot.processes.flatMap(p => p.moduleNames.map(name => name.toLowerCase())));
    const before = names(first);
    return [...names(last)].filter(name => !before.has(name)).sort();
}
function formatSummary(report) {
    const yn = value => value === true ? 'yes' : value === false ? 'no' : '?';
    const lines = [`mode: ${report.mode}`, `report: ${report.reportPath}`];
    if (report.error) lines.push(`error: ${report.error}`);
    const snapshots = Object.values(report.snapshots);
    if (snapshots.length) {
        const first = snapshots[0], last = snapshots[snapshots.length - 1];
        const ps = first.powershell;
        const readable = first.processes.filter(p => !p.error && p.moduleCount > 0);
        lines.push(`CNEXT.exe pids: ${first.cnextPids.join(', ') || '(none)'}`);
        lines.push(`Process.Modules: ${ps.error ? 'ERROR ' + ps.error
            : readable.length ? `ok (64-bit=${yn(ps.is64)}; ${readable.map(p => `pid ${p.pid}: ${p.moduleCount} modules`).join(', ')})` : 'no readable process'}`);
        for (const p of first.processes.filter(p => p.error)) lines.push(`  pid ${p.pid}: ${p.error}`);
        const width = Math.max(...snapshots.map(s => s.label.length), 12) + 2;
        lines.push('', 'dll'.padEnd(14) + snapshots.map(s => s.label.padEnd(width)).join(''));
        for (const dll of KEY_DLLS) {
            const cells = snapshots.map(s => {
                const d = s.dlls[dll];
                const detail = d.byModules === null ? ' (T only)' : d.byModules !== d.byTasklist ? ` (M:${yn(d.byModules)} T:${yn(d.byTasklist)})` : '';
                return (yn(d.loaded) + detail).padEnd(width);
            });
            lines.push(dll.padEnd(14) + cells.join(''));
        }
        lines.push('  (M = Process.Modules, T = tasklist /M <dll>; loaded if either says yes)');
        const related = [...new Set(last.processes.flatMap(p => p.related.map(m => m.name)))];
        lines.push(`related modules at ${last.label}: ${related.join(', ') || '(none)'}`);
        if (report.summary?.modulesAdded) lines.push(`modules added ${first.label} -> ${last.label}: ${report.summary.modulesAdded.join(', ') || '(none)'}`);
    }
    if (report.summary?.runs?.length) {
        lines.push('', 'runs:');
        for (const run of report.summary.runs) {
            lines.push(`  ${run.name} (${run.script}): exitCode=${run.exitCode} result=${run.resultFound ? 'yes' : 'NO'} failedSteps=${run.failedSteps.join(', ') || 'none'}`);
            if (run.error) lines.push(`    error: ${run.error.trim().replace(/\s*\r?\n\s*/g, ' / ')}`);
            if (run.listFailures?.length) lines.push(`    list failures: ${run.listFailures.join(' / ')}`);
            if (run.listResultError) lines.push(`    list result: ${run.listResultError}`);
        }
    }
    if (report.summary?.verdict) lines.push('', `verdict: ${report.summary.verdict}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    return lines.join('\n');
}
async function main() {
    let mode;
    try { mode = parseMode(process.argv.slice(2)); } catch (error) { console.error(error.message); process.exitCode = 2; return; }
    if (mode === 'help') { console.log(USAGE); return; }
    const root = path.resolve(__dirname, '..');
    const bridge = path.join(root, 'bin', 'catia-run.exe');
    const reportPath = uniqueReportPath(path.join(root, 'temp'), mode);
    const report = { mode, startedAt: new Date().toISOString(), reportPath, bridge, workDirectory: null, notes: [], snapshots: {}, catscript: {}, catvbs: {} };
    const snap = label => (report.snapshots[label] = moduleSnapshot(label));
    try {
        if (mode !== 'check-only' && !fs.existsSync(bridge)) throw new Error(`Missing ${bridge}; run npm run compile:bridge`);
        const first = snap(mode === 'check-only' ? 'current' : 'before');
        if (!first.cnextPids.length) throw new Error('CNEXT.exe is not running');
        if (first.cnextPids.length > 1) report.notes.push(`More than one CNEXT.exe is running (${first.cnextPids.join(', ')}); a DLL counts as loaded if any of them has it.`);
        if (first.powershell.error) report.notes.push(`Process.Modules could not be read (${first.powershell.error}); judged by tasklist /M only, which may miss modules.`);
        if (mode !== 'check-only') {
            const work = report.workDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'cat5dev-probe-'));
            if (first.loaded) report.notes.push('vbscript.dll was already loaded in CNEXT.exe before the probe; dependency cannot be judged (restart CATIA and re-run). Syntax results below are still valid.');
            if (mode === 'no-apc') {
                await runProbe(report, bridge, work, NO_APC_SECTIONS);
                snap('afterCatScript');
            } else if (mode === 'apc') {
                await runProbe(report, bridge, work, Object.keys(SECTIONS));
                snap('afterProbe');
                await runGeneratedList(report, bridge, work);
                snap('afterGeneratedList');
            } else {
                await runProbe(report, bridge, work, Object.keys(SECTIONS));
                await runGeneratedList(report, bridge, work);
                snap('afterCatScript');
                report.catvbs.full = await runScript(bridge, path.join(work, 'catvbs-full'), 'probe.catvbs', buildProbeScript(path.join(work, 'catvbs-full')));
                snap('afterCatvbs');
            }
        }
        const snapshots = Object.values(report.snapshots);
        const runs = [
            ...(report.catscript.full ? [summarizeRun('catscript.full', report.catscript.full)] : []),
            ...Object.entries(report.catscript.sections || {}).map(([name, run]) => summarizeRun('catscript.section.' + name, run)),
            ...(report.catscript.generatedList ? [summarizeRun('catscript.generatedList', report.catscript.generatedList)] : []),
            ...(report.catvbs.full ? [summarizeRun('catvbs.full', report.catvbs.full)] : [])
        ];
        let verdict = judge(mode, Object.fromEntries(snapshots.map(s => [s.label, s.loaded])));
        if (report.catscript.full && !report.catscript.full.resultFound) verdict += ' (warning: probe.CATScript wrote no result; check exitCode/error)';
        report.summary = {
            mode,
            dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, Object.fromEntries(snapshots.map(s => [s.label, s.dlls[dll].loaded]))])),
            modulesAdded: snapshots.length > 1 ? addedModules(snapshots[0], snapshots[snapshots.length - 1]) : undefined,
            runs,
            verdict
        };
    } catch (error) {
        report.error = String(error); process.exitCode = 1;
    }
    report.finishedAt = new Date().toISOString();
    fs.mkdirSync(path.dirname(reportPath), { recursive: true });
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(formatSummary(report));
}
if (require.main === module) main();
module.exports = { buildProbeScript, SECTIONS, NO_APC_SECTIONS, parseMode, judge, listModules, moduleSnapshot, KEY_DLLS, timestamp, addedModules };
