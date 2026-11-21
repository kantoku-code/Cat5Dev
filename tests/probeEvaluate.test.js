const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, judge, parseProbeOutput, parseJsonResult } = require('../scripts/probe-evaluate');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
const JSCRIPT_STEPS = ['jscript-params', 'jscript-fso', 'jscript-apc', 'jscript-read'];

function buildProbe(t) {
    const out = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${out}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    return out;
}

test('Evaluate probe parses steps and options', () => {
    assert.equal(parseArgs(['--help']).mode, 'help');
    assert.deepEqual(parseArgs(['--check-only']), { mode: 'check-only', steps: [], byref: false, notes: [] });
    assert.deepEqual(parseArgs(['basic', 'sub', 'basic']).steps, ['basic', 'sub']);
    assert.deepEqual(parseArgs(['apc2']).steps, ['apc2-getobject', 'apc2-catia', 'apc2-application']);
    assert.equal(parseArgs(['--byref', 'params']).byref, true);
    const moved = parseArgs(['jscript', 'basic']);
    assert.deepEqual(moved.steps, ['basic', 'jscript']);
    assert.equal(moved.notes.length, 1);
    assert.throws(() => parseArgs([]), /No step/);
    assert.throws(() => parseArgs(['--check-only', 'basic']), /runs no step/);
    assert.throws(() => parseArgs(['vbscript']), /Unknown step/);
});
test('Evaluate probe keeps the order of JScript-only runs and moves JScript steps after VBA steps', () => {
    const only = parseArgs(['jscript', 'jscript-params', 'jscript-fso', 'jscript-apc', 'jscript-read']);
    assert.deepEqual(only.steps, ['jscript', 'jscript-params', 'jscript-fso', 'jscript-apc', 'jscript-read']);
    assert.deepEqual(only.notes, []);
    assert.deepEqual(parseArgs(['jscript-read', 'jscript']).steps, ['jscript-read', 'jscript']);
    const mixed = parseArgs(['jscript-apc', 'basic', 'jscript', 'apc']);
    assert.deepEqual(mixed.steps, ['basic', 'apc', 'jscript-apc', 'jscript']);
    assert.equal(mixed.notes.length, 1);
    assert.deepEqual(parseArgs(['basic', 'jscript', 'jscript-read']).notes, []);
});
test('Evaluate probe verdicts', () => {
    assert.match(judge('check-only', false), /not loaded/);
    assert.match(judge('run', true, []), /^undetermined/);
    const ok = { language: 1, ok: true, loadedAfter: false };
    assert.equal(judge('run', false, [{ ...ok, step: 'basic' }, { ...ok, step: 'apc' }]), 'Evaluate(VBA) did not load vbscript.dll');
    assert.match(judge('run', false, [{ ...ok, step: 'basic' }, { ...ok, step: 'apc', loadedAfter: true }]), /^Evaluate\(VBA\) loaded vbscript\.dll .*apc/);
    assert.match(judge('run', false, [{ ...ok, step: 'basic' }, { step: 'jscript', language: 4, ok: true, loadedAfter: true }]), /^Evaluate\(VBA\) did not load .*JScript/);
    assert.match(judge('run', false, [{ ...ok, step: 'basic', ok: false }]), /warning: no VBA step succeeded/);
    assert.equal(judge('run', false, [{ step: 'jscript', language: 4, ok: true, loadedAfter: false }]), 'Evaluate(JScript) did not load vbscript.dll');
});
test('Evaluate probe verdicts for JScript steps', () => {
    const js = step => ({ step, language: 4, ok: true, loadedAfter: false });
    const all = ['jscript', ...JSCRIPT_STEPS].map(js);
    assert.equal(judge('run', false, all), 'Evaluate(JScript) did not load vbscript.dll');
    const loaded = all.map((entry, i) => ({ ...entry, loadedAfter: i >= 3 }));
    assert.equal(judge('run', false, loaded), 'Evaluate(JScript) loaded vbscript.dll (first seen after step jscript-apc)');
    assert.equal(judge('run', false, [{ step: 'basic', language: 1, ok: true, loadedAfter: false }, ...all]),
        'Evaluate(VBA) did not load vbscript.dll; Evaluate(JScript) did not load vbscript.dll');
    assert.match(judge('run', false, [{ step: 'basic', language: 1, ok: true, loadedAfter: true }, ...all.map(e => ({ ...e, loadedAfter: true }))]),
        /^Evaluate\(VBA\) loaded vbscript\.dll .*JScript steps ran after/);
    assert.match(judge('run', false, all.map(e => ({ ...e, ok: false }))), /warning: no JScript step succeeded/);
});
test('Evaluate probe output lines are parsed as JSON', () => {
    const { lines, other } = parseProbeOutput('{"step":"basic","ok":true,"result":"ok:2"}\r\nnoise\r\n\r\n');
    assert.deepEqual(lines, [{ step: 'basic', ok: true, result: 'ok:2' }]);
    assert.deepEqual(other, ['noise']);
    assert.deepEqual(parseJsonResult('{"ok":true}'), { ok: true });
    assert.equal(parseJsonResult('count=0;MSAPC.Apc.7.1=ok'), undefined);
    assert.equal(parseJsonResult('{broken'), undefined);
});
test('EvalProbe.cs compiles with the .NET Framework csc (C# 5)', { skip: !fs.existsSync(csc) }, t => {
    assert.ok(fs.existsSync(buildProbe(t)));
});

// ---- JScript code of the jscript-* steps, printed by `eval-probe.exe --print-code` (no CATIA connection) and run in
// Node with a fake ActiveXObject. Node is a superset of ES3, so ES5+ features are rejected by a lint below.
const ES5_OR_LATER = /\b(let|const|class)\s|=>|`|\.(forEach|map|filter|reduce|some|every|indexOf|lastIndexOf|trim|bind)\(|\bJSON\.|\bObject\.(keys|create|defineProperty)|\bArray\.isArray|\bDate\.now|,\s*[\]}]/;
function comError(number, message) { return Object.assign(new Error(message), { number }); }
function fakeProjects(projects) {
    return { Count: projects.length, Item: i => {
        const project = projects[i - 1];
        if (project.fail) throw comError(0x800A0409 | 0, 'project failure');
        const components = project.components;
        return { Name: project.name, VBComponents: { Count: components.length, Item: j => {
            const c = components[j - 1];
            return { Name: c.name, Type: c.type, CodeModule: { CountOfLines: c.lines.length, Lines: (start, count) => {
                if (c.linesFail && c.linesFail(start, count)) throw comError(0x800A01B6 | 0, 'Lines failure');
                return c.lines.slice(start - 1, start - 1 + count).join('\r\n');
            } } };
        } } };
    } };
}
function runJScript(code, { apc = null, failIds = ['MSAPC.Apc.7.1'], args = [] } = {}) {
    const created = [];
    function ActiveXObject(id) {
        created.push(id);
        if (failIds.includes(id)) throw comError(0x800A01AD | 0, 'cannot create object');
        if (/^MSAPC/.test(id)) return { VBE: { VBProjects: apc } };
        return {};
    }
    const catMain = new Function('ActiveXObject', `${code}\nreturn CATMain;`)(ActiveXObject);
    return { result: catMain(...args), created };
}

test('JScript steps are ES3 and behave with a fake APC', { skip: !fs.existsSync(csc) }, t => {
    const exe = buildProbe(t);
    const code = Object.fromEntries(JSCRIPT_STEPS.map(step =>
        [step, execFileSync(exe, ['--print-code', step], { windowsHide: true, encoding: 'utf8', timeout: 10000 })]));
    for (const [step, source] of Object.entries(code)) {
        assert.doesNotMatch(source, ES5_OR_LATER, step);
        assert.doesNotMatch(source, /\.(Add|Remove|DeleteLines|AddFromString|InsertLines|ReplaceLine|Export|Import|Write|SaveToFile|CreateTextFile|OpenTextFile)\b/, step);
        assert.match(source, /function CATMain\(/, step);
    }

    assert.equal(runJScript(code['jscript-params'], { args: ['cat5dev-param'] }).result, 'param:cat5dev-param');
    assert.equal(runJScript(code['jscript-fso']).result, 'ADODB.Stream=ok(object)|Scripting.FileSystemObject=ok(object)');
    assert.match(runJScript(code['jscript-fso'], { failIds: ['ADODB.Stream'] }).result,
        /^ADODB\.Stream=err=0x800A01AD cannot create object\|Scripting\.FileSystemObject=ok\(object\)$/);

    const projects = fakeProjects([
        { name: 'Empty', components: [{ name: 'Blank', type: 1, lines: [] }] },
        { name: 'Review', components: [
            { name: 'ThisDocument', type: 100, lines: [] },
            { name: 'C5D_UserModule', type: 1, lines: ["' 日本語コメント \"quoted\" \\ path", 'Option Explicit', 'Sub CATMain()', 'End Sub'] }] },
        { name: 'Locked', fail: true, components: [] },
        { name: 'Fourth', components: [{ name: 'Never', type: 1, lines: ['x'] }] }]);
    const apc = runJScript(code['jscript-apc'], { apc: projects });
    assert.equal(apc.result, 'count=4|Empty|Review|#3 err=0x800A0409 project failure|Fourth;MSAPC.Apc.7.1=err=0x800A01AD cannot create object;MSAPC.Apc.6.2=ok');
    assert.deepEqual(apc.created, ['MSAPC.Apc.7.1', 'MSAPC.Apc.6.2']);
    assert.match(runJScript(code['jscript-apc'], { failIds: ['MSAPC.Apc.7.1', 'MSAPC.Apc.6.2', 'MSAPC.Apc'] }).result, /^count=-1\|apc=null;.*MSAPC\.Apc=err=/);
    assert.match(runJScript(code['jscript-apc'], { apc: fakeProjects([]) }).result, /^count=0;/);

    const read = JSON.parse(runJScript(code['jscript-read'], { apc: projects }).result);
    assert.equal(read.ok, true); assert.equal(read.projectCount, 4); assert.equal(read.projects.length, 3);
    assert.deepEqual(read.projects[1].components.map(c => [c.name, c.type, c.countOfLines]), [['ThisDocument', 100, 0], ['C5D_UserModule', 1, 4]]);
    assert.match(read.projects[2].error, /^Item err=0x800A0409/);
    const sample = read.sample;
    assert.equal(sample.project, 'Review'); assert.equal(sample.component, 'C5D_UserModule');
    assert.equal(sample.requested, 3); assert.equal(sample.via, 'lines-call'); assert.equal(sample.perlineEqual, true);
    const expected = ["' 日本語コメント \"quoted\" \\ path", 'Option Explicit', 'Sub CATMain()'].join('\r\n');
    assert.equal(sample.length, expected.length); assert.equal(sample.head, expected.slice(0, 40));
    assert.equal(sample.nonAscii, 7); assert.match(sample.nonAsciiCodes, /^U\+65E5 U\+672C U\+8A9E/); assert.equal(sample.replacementChars, 0);
    assert.match(runJScript(code['jscript-read'], { apc: projects }).result, /\\u000D\\u000A/);

    const perline = fakeProjects([{ name: 'P', components: [{ name: 'M', type: 1, lines: ['a', 'b'], linesFail: (start, count) => count > 1 }] }]);
    const fallback = JSON.parse(runJScript(code['jscript-read'], { apc: perline }).result).sample;
    assert.equal(fallback.via, 'lines-perline'); assert.equal(fallback.head, 'a\r\nb'); assert.match(fallback.attempts, /^lines-call err=0x800A01B6/);
    const never = fakeProjects([{ name: 'P', components: [{ name: 'M', type: 1, lines: ['a'], linesFail: () => true }] }]);
    const none = JSON.parse(runJScript(code['jscript-read'], { apc: never }).result).sample;
    assert.equal(none.via, 'none'); assert.equal(none.head, undefined);
    assert.equal(JSON.parse(runJScript(code['jscript-read'], { apc: fakeProjects([]) }).result).sample, null);
    assert.equal(JSON.parse(runJScript(code['jscript-read'], { failIds: ['MSAPC.Apc.7.1', 'MSAPC.Apc.6.2', 'MSAPC.Apc'] }).result).ok, false);
});
