// B0 probe (developer tool, not shipped): does CATScript need vbscript.dll, and which syntax does it accept?
// Usage: npm run compile && npm run compile:bridge && node scripts/probe-catscript.js
// Read-only for CATIA: only reads VBProjects; never adds/changes modules or projects.
// Order: tasklist -> probe.CATScript (+ per-section runs if it fails) -> generated list script as .CATScript
//        -> tasklist -> probe.catvbs (control) -> tasklist. Result: temp/probe-catscript-result.json and stdout.
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
function tasklist(moduleFilter) {
    const args = ['/FI', 'IMAGENAME eq CNEXT.exe', ...(moduleFilter ? ['/M', 'vbscript.dll'] : []), '/FO', 'CSV', '/NH'];
    const raw = decode(execFileSync('tasklist', args, { windowsHide: true })).trim();
    const pids = raw.split(/\r?\n/).filter(line => /^"CNEXT\.exe"/i.test(line)).map(line => Number(line.split('","')[1]));
    return { command: 'tasklist ' + args.map(a => a.includes(' ') ? `"${a}"` : a).join(' '), raw, pids };
}
function vbscriptState(label) {
    const all = tasklist(false), loaded = tasklist(true);
    return { label, at: new Date().toISOString(), cnextPids: all.pids, vbscriptLoadedPids: loaded.pids, loaded: loaded.pids.length > 0, command: loaded.command, raw: loaded.raw };
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
async function main() {
    const root = path.resolve(__dirname, '..');
    const bridge = path.join(root, 'bin', 'catia-run.exe');
    const reportPath = path.join(root, 'temp', 'probe-catscript-result.json');
    const work = fs.mkdtempSync(path.join(os.tmpdir(), 'cat5dev-probe-'));
    const report = { startedAt: new Date().toISOString(), bridge, workDirectory: work, notes: [], tasklist: {}, catscript: {}, catvbs: {} };
    try {
        if (!fs.existsSync(bridge)) throw new Error(`Missing ${bridge}; run npm run compile:bridge`);
        report.tasklist.before = vbscriptState('before');
        if (!report.tasklist.before.cnextPids.length) throw new Error('CNEXT.exe is not running');
        if (report.tasklist.before.loaded) {
            report.notes.push('vbscript.dll was already loaded in CNEXT.exe before the probe; dependency cannot be judged (restart CATIA and re-run). Syntax results below are still valid.');
        }
        const full = await runScript(bridge, path.join(work, 'catscript-full'), 'probe.CATScript', buildProbeScript(path.join(work, 'catscript-full')));
        report.catscript.full = full;
        if (!full.resultFound) {
            report.notes.push('Full CATScript probe produced no result; each section was run alone to isolate the failing syntax.');
            report.catscript.sections = {};
            for (const name of Object.keys(SECTIONS)) {
                const directory = path.join(work, 'catscript-' + name);
                report.catscript.sections[name] = await runScript(bridge, directory, 'probe.CATScript', buildProbeScript(directory, [name]));
            }
        }
        // The real generated script (read-only 'list') renamed to .CATScript.
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
        report.tasklist.afterCatScript = vbscriptState('afterCatScript');
        report.catvbs.full = await runScript(bridge, path.join(work, 'catvbs-full'), 'probe.catvbs', buildProbeScript(path.join(work, 'catvbs-full')));
        report.tasklist.afterCatvbs = vbscriptState('afterCatvbs');
        const { before, afterCatScript, afterCatvbs } = report.tasklist;
        report.summary = {
            vbscriptPreloaded: before.loaded,
            catscriptProbeCompleted: full.resultFound,
            catscriptGeneratedListOk: Boolean(list.listResult && !list.listResult.failures.length),
            vbscriptLoadedAfterCatScript: afterCatScript.loaded,
            vbscriptLoadedAfterCatvbs: afterCatvbs.loaded,
            verdict: before.loaded ? 'undetermined: vbscript.dll already loaded before probe'
                : afterCatScript.loaded ? 'CATScript loads vbscript.dll (B2)'
                    : afterCatvbs.loaded ? 'CATScript does not load vbscript.dll; catvbs does (B1 viable)'
                        : 'inconclusive: control catvbs did not load vbscript.dll'
        };
    } catch (error) {
        report.error = String(error); process.exitCode = 1;
    }
    fs.mkdirSync(path.dirname(reportPath), { recursive: true });
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2));
    console.log(JSON.stringify(report, null, 2));
}
if (require.main === module) main();
module.exports = { buildProbeScript, SECTIONS };
