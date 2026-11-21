// Static checks of the VBA that runs inside CATIA (no CATIA): resources/bridge/C5D_Bridge_Push.bas (v2), C5D_Bridge_Runtime.bas and the
// temporary modules of scripts/build-bridge-catvba.js. Rules (docs/dev/bridge-contract.md section 9): ASCII only, CRLF, no Attribute line,
// On Error Resume Next first in every procedure, no MsgBox / Declare / GoTo / file or shell access, no property read in an If / loop
// condition, no module-level variables, the exact public signatures of the contract.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { BUILDER_SOURCE, PROBE_SOURCE, parseRecords } = require('../scripts/build-bridge-catvba');

const dir = path.join(__dirname, '..', 'resources', 'bridge');
const PUSH = fs.readFileSync(path.join(dir, 'C5D_Bridge_Push.bas'), 'utf8');
const RUNTIME = fs.readFileSync(path.join(dir, 'C5D_Bridge_Runtime.bas'), 'utf8');

/** The code part of a line: string literals blanked ("" kept as a quote pair) and the comment removed. */
function codeOf(line) {
    let out = '', inString = false;
    for (let i = 0; i < line.length; i++) {
        const c = line[i];
        if (inString) {
            if (c === '"' && line[i + 1] === '"') { out += '__'; i++; continue; }
            if (c === '"') { inString = false; out += '"'; continue; }
            out += '_';
        } else {
            if (c === "'") break;
            if (c === '"') inString = true;
            out += c;
        }
    }
    return out.replace(/\s+$/, '');
}
/** Procedures: [{ scope, kind, name, params, returns, body: [lines] }]; module-level code lines outside procedures. */
function parseModule(source) {
    const lines = source.split('\r\n');
    const procs = [], moduleLevel = [];
    let current = null;
    for (const raw of lines) {
        const code = codeOf(raw);
        const head = /^(Public|Private) (Function|Sub) (\w+)\((.*)\)(?: As (\w+))?$/.exec(code);
        if (head) {
            assert.equal(current, null, `nested procedure ${head[3]}`);
            current = { scope: head[1], kind: head[2], name: head[3], params: head[4], returns: head[5] ?? null, body: [], raw: [] };
            continue;
        }
        if (/^End (Function|Sub)$/.test(code)) { assert.ok(current, 'End without procedure'); procs.push(current); current = null; continue; }
        if (current) { current.body.push(code); current.raw.push(raw); } else if (code.trim()) moduleLevel.push(code);
    }
    assert.equal(current, null, 'unterminated procedure');
    return { lines, procs, moduleLevel };
}
const MEMBER = /\b[A-Za-z_]\w*\s*\.\s*[A-Za-z_]/;
/** Every rule of the contract for one VBA source; returns the parsed module. */
function checkRules(source, label) {
    assert.ok(/^[\x20-\x7e\r\n]*$/.test(source), `${label}: ASCII printable only (no tab, no non-ASCII)`);
    assert.ok(!/\r(?!\n)|(?<!\r)\n/.test(source), `${label}: CRLF only`);
    assert.ok(source.endsWith('\r\n'), `${label}: ends with CRLF`);
    assert.ok(!/^Attribute /m.test(source), `${label}: no Attribute line`);
    const m = parseModule(source);
    assert.equal(m.lines[0], 'Option Explicit', `${label}: Option Explicit first`);
    for (const line of m.lines) assert.ok(line.length <= 1023, `${label}: line longer than VBA allows`);
    for (const line of m.moduleLevel) {
        assert.ok(/^Option Explicit$|^Private Const \w+ As String = "_*"$/.test(line), `${label}: module-level code other than Option Explicit / Private Const: ${line}`);
    }
    for (const p of m.procs) {
        assert.equal(p.kind, 'Function', `${label}.${p.name}: functions only`);
        assert.equal(p.body[0], '    On Error Resume Next', `${label}.${p.name}: On Error Resume Next first`);
        for (const line of p.body) {
            const where = `${label}.${p.name}: ${line.trim()}`;
            assert.ok(!/\b(MsgBox|InputBox|Declare|GoTo|GoSub|Resume|Stop|Shell|SendKeys|Kill|FileCopy|Environ|DoEvents|Open|Close|Print|Write|Put|Get|ExecuteLine|Run|Save|SaveAs|Unload)\b/.test(line.replace(/On Error Resume Next/, '')), `forbidden statement: ${where}`);
            assert.ok(!/^\s*End\s*$/.test(line), `bare End: ${where}`);
            const cond = /^\s*(?:If|ElseIf)\s+(.*?)\s+Then\b/.exec(line)?.[1] ?? /^\s*(?:Do|Loop)\s+(?:While|Until)\s+(.*)$/.exec(line)?.[1]
                ?? /^\s*While\s+(.*)$/.exec(line)?.[1] ?? /^\s*Select Case\s+(.*)$/.exec(line)?.[1] ?? /^\s*For\s+\w+\s*=\s*(.*)$/.exec(line)?.[1] ?? null;
            if (cond !== null) assert.ok(!MEMBER.test(cond), `property read in a condition: ${where}`);
            if (/CreateObject\(/.test(line)) assert.ok(/CreateObject\("_+"\)/.test(line), `CreateObject with a literal only: ${where}`);
        }
    }
    for (const raw of m.lines.filter(l => /CreateObject\(/.test(codeOf(l)))) {
        assert.match(raw, /CreateObject\("MSAPC\.Apc(?:\.7\.1|\.6\.2)?"\)/, `${label}: CreateObject only for MSAPC.Apc`);
    }
    return m;
}
const publicSignatures = m => m.procs.filter(p => p.scope === 'Public').map(p => `${p.name}(${p.params}) As ${p.returns}`);

test('C5D_Bridge_Push.bas v2 follows the VBA rules and has exactly the contract functions', () => {
    const m = checkRules(PUSH, 'C5D_Bridge_Push');
    assert.deepEqual(publicSignatures(m), [
        'C5D_Version() As Variant',
        'C5D_Self() As Variant',
        'C5D_Who() As Variant',
        'C5D_List() As Variant',
        'C5D_Install(ByVal moduleName As Variant, ByVal source As Variant) As Variant'
    ]);
    assert.match(PUSH, /^Private Const BRIDGE_VERSION As String = "c5d-bridge:2"\r$/m);
    assert.match(PUSH, /^Private Const BRIDGE_PROJECT As String = "C5D_Bridge"\r$/m);
    assert.match(PUSH, /^Private Const BRIDGE_MODULE As String = "C5D_Bridge_Push"\r$/m);
    // writes only into ExecutingProject.VBProject (never "the first project named C5D_Bridge"), refuses itself, never removes / closes / saves
    const install = m.procs.find(p => p.name === 'C5D_Install').raw.join('\n');
    assert.match(install, /Set ex = apc\.ExecutingProject[\s\S]*Set vbp = ex\.VBProject[\s\S]*Set comps = vbp\.VBComponents/);
    assert.doesNotMatch(install, /VBE|VBProjects/);
    assert.match(install, /StrComp\(mn, BRIDGE_MODULE, vbTextCompare\) = 0 Then\n\s+C5D_Install = "error:self"/);
    assert.match(install, /comps\.Add\(1\)/);
    assert.doesNotMatch(PUSH, /\.Remove\b|\.Close\b|\.Save\b|VBProjects/);
    assert.match(m.procs.find(p => p.name === 'C5D_Self').raw.join('\n'), /Set C5D_Self = ex$/m);
});

test('C5D_Bridge_Runtime.bas follows the VBA rules and has exactly the contract functions', () => {
    const m = checkRules(RUNTIME, 'C5D_Bridge_Runtime');
    assert.deepEqual(publicSignatures(m), [
        'C5D_RT_Version() As Variant',
        'C5D_RT_ListProjects() As Variant',
        'C5D_RT_ListModules(ByVal projectName As Variant) As Variant',
        'C5D_RT_ReadModule(ByVal projectName As Variant, ByVal moduleName As Variant) As Variant',
        'C5D_RT_WriteModule(ByVal projectName As Variant, ByVal moduleName As Variant, ByVal moduleType As Variant, ByVal moduleText As Variant) As Variant',
        'C5D_RT_DeleteModule(ByVal projectName As Variant, ByVal moduleName As Variant) As Variant'
    ]);
    assert.match(RUNTIME, /^Private Const RT_VERSION As String = "c5d-runtime:\d+"\r$/m);
    assert.doesNotMatch(RUNTIME, /ExecutingProject|\.Close\b|\.Save\b|ADODB|Scripting\.|Charset/);
    // the failure phases and messages of the old operation.catvbs (src/catiaScripts.ts) are kept
    for (const phase of ['connect.apc', 'connect.projects', 'project.lookup', 'project.components', 'list.count', 'list.project', 'list.name',
        'export.count', 'export.component', 'export.type', 'export.name', 'export.codemodule', 'export.linecount', 'export.lines',
        'push.add', 'push.name', 'push.type', 'push.codemodule', 'push.linecount', 'push.deletelines', 'push.addfromstring', 'push.delete']) {
        assert.ok(RUNTIME.includes(`"${phase}"`) || RUNTIME.includes(`prefix & ".${phase.split('.').pop()}"`), `phase ${phase}`);
    }
    for (const message of ['Target project not found', 'New UserForms require a designer', 'Module type does not match', 'Module to delete not found', 'No APC object']) {
        assert.ok(RUNTIME.includes(`"${message}"`), message);
    }
    // the projects named C5D_Bridge are never listed or matched; modules are matched ignoring case, projects exactly
    const body = name => m.procs.find(p => p.name === name).raw.join('\n');
    assert.match(body('RtFindProject'), /StrComp\(nm, BRIDGE_PROJECT, vbBinaryCompare\) <> 0 Then\n\s+If StrComp\(nm, projectName, vbBinaryCompare\) = 0/);
    assert.match(body('C5D_RT_ListProjects'), /StrComp\(nm, BRIDGE_PROJECT, vbBinaryCompare\) <> 0/);
    assert.match(body('RtFindModule'), /StrComp\(nm, moduleName, vbTextCompare\) = 0/);
    // the APC fallbacks of the old ConnectProjects: 7.1, then 6.2, then MSAPC.Apc
    assert.match(body('RtProjects'), /"MSAPC\.Apc\.7\.1"[\s\S]*"MSAPC\.Apc\.6\.2"[\s\S]*"MSAPC\.Apc"/);
    // write: a new UserForm is refused before Add; type checked before the lines are touched; DeleteLines before AddFromString
    const write = body('C5D_RT_WriteModule');
    assert.ok(write.indexOf('New UserForms require a designer') < write.indexOf('comps.Add(mt)'));
    assert.ok(write.indexOf('Module type does not match') < write.indexOf('cm.DeleteLines'));
    assert.ok(write.indexOf('cm.DeleteLines 1, n') < write.indexOf('cm.AddFromString src'));
    assert.match(write, /If Len\(src\) > 0 Then\n\s+Err\.Clear\n\s+cm\.AddFromString src/);
    assert.match(body('C5D_RT_DeleteModule'), /comps\.Remove comp/);
    assert.match(body('C5D_RT_ReadModule'), /C5D_RT_ReadModule = "ok:" & src$/m);
});

test('temporary modules of the build script follow the same rules', () => {
    const builder = checkRules(BUILDER_SOURCE, 'C5D_Builder');
    assert.deepEqual(publicSignatures(builder), ['C5D_BuildSelf() As Variant', 'C5D_BuildWho() As Variant',
        'C5D_BuildPut(ByVal moduleName As Variant, ByVal source As Variant) As Variant']);
    assert.doesNotMatch(BUILDER_SOURCE, /\.Remove\b|\.Close\b|\.Save\b|VBProjects/);
    checkRules(PROBE_SOURCE, 'C5D_BuildProbe');
});

test('the rule checker rejects what the rules forbid', () => {
    const wrap = body => ['Option Explicit', 'Public Function F() As Variant', '    On Error Resume Next', ...body, 'End Function', ''].join('\r\n');
    assert.doesNotThrow(() => checkRules(wrap(['    x = a.b', '    If x Then F = 1', "    If y = 1 Then F = \"a.b\" ' c.d"]), 'ok'));
    for (const [bad, why] of [[wrap(['    If a.b Then F = 1']), /condition/], [wrap(['    If Err.Number <> 0 Then F = 1']), /condition/],
        [wrap(['    For i = 1 To c.Count']), /condition/], [wrap(['    MsgBox "x"']), /forbidden/], [wrap(['    a.Save']), /forbidden/],
        [wrap(['    x = "\u00e9"']), /ASCII/], [wrap(['    x = 1']).replace(/\r\n/g, '\n'), /CRLF/], [`Attribute VB_Name = "x"\r\n${wrap([])}`, /Attribute|Option Explicit/],
        [wrap([]).replace('    On Error Resume Next\r\n', ''), /On Error Resume Next/], [`Option Explicit\r\nDim g As Long\r\n${wrap([]).slice(17)}`, /module-level/],
        [wrap(['    Set o = CreateObject("Scripting.FileSystemObject")']), /MSAPC/]]) {
        assert.throws(() => checkRules(bad, 'bad'), why);
    }
});

test('record format: escaping of the VBA side and parseRecords are inverse', () => {
    // BrEsc / RtEsc: % first, then TAB, LF, CR
    const esc = t => t.replace(/%/g, '%25').replace(/\t/g, '%09').replace(/\n/g, '%0A').replace(/\r/g, '%0D');
    const fields = ['C:\\a%09b\\x.catvba', 'tab\there', 'line\r\nbreak', '%25%', '日本語🙂', ''];
    const parsed = parseRecords(`ok\n${fields.map(esc).join('\t')}\n${esc('C5D_Bridge_Push')}\t1`);
    assert.equal(parsed.ok, true);
    assert.deepEqual(parsed.rows, [fields, ['C5D_Bridge_Push', '1']]);
    assert.deepEqual(parseRecords('ok'), { ok: true, rows: [], error: null });
    assert.equal(parseRecords('error:list.apc:429:No APC object').ok, false);
    assert.equal(parseRecords(null).ok, false);
    assert.equal(parseRecords('okay').ok, false);
});
