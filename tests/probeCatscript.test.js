// B0 probe (scripts/probe-catscript.js, developer tool). Moved from catiaScripts.test.js when the generated
// operation.catvbs and its cscript harness were removed (Push/Pull now run through C5D_Bridge.catvba).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const cscript = path.join(process.env.SystemRoot || 'C:\\Windows', 'SysWOW64', 'cscript.exe');

// The B0 probe .vbs (developer tool) is sometimes quarantined by Windows Defender as Trojan:Script/ObfusScript.A!ml
// (a machine-learning false positive): the file vanishes or becomes unreadable, or AMSI blocks cscript.
function blockedByAntivirus(file, error) {
    if (['EACCES', 'EPERM', 'EBUSY'].includes(error?.code)) return true;
    try { fs.readFileSync(file); } catch { return true; }
    const output = [error?.stdout, error?.stderr].filter(Boolean).map(buffer => {
        try { return new TextDecoder('shift_jis').decode(buffer); } catch { return String(buffer); }
    }).join('\n');
    return /malicious|antivirus|悪意のある|ウイルス対策/i.test(output);
}
/** Writes and runs a probe .vbs; false (test skipped) when the antivirus removed or blocked it. */
function runProbeVbs(t, file, source) {
    try {
        fs.writeFileSync(file, source);
        execFileSync(cscript, ['//nologo', file], { windowsHide: true, timeout: 10000 });
        return true;
    } catch (error) {
        if (!blockedByAntivirus(file, error)) throw error;
        t.skip(`probe .vbs was removed or blocked by the antivirus (likely a Defender false positive): ${error.code ?? error.message}`);
        return false;
    }
}
test('B0 probe scripts compile and CATIA-independent sections run under VBScript', { skip: !fs.existsSync(cscript) }, t => {
    const { buildProbeScript, SECTIONS } = require('../scripts/probe-catscript');
    const directory = fixture(t), file = path.join(directory, 'probe.vbs');
    // Compile only: top level has no calls, so APC/VBE sections are never executed.
    for (const names of [undefined, ...Object.keys(SECTIONS).map(name => [name])]) {
        if (!runProbeVbs(t, file, buildProbeScript(directory, names))) return;
    }
    if (!runProbeVbs(t, file, buildProbeScript(directory, ['stream', 'objfunc', 'byref', 'builtins', 'err']) + 'CATMain\r\n')) return;
    const steps = Object.fromEntries(fs.readFileSync(path.join(directory, 'probe-result.txt'), 'utf8').replace(/^﻿/, '').trim().split('\r\n')
        .map(line => line.split('\t')).map(([name, err, detail]) => [name, { err: Number(err), detail }]));
    for (const [name, step] of Object.entries(steps)) if (name !== 'err.createobject') assert.equal(step.err, 0, name);
    assert.equal(steps['err.createobject'].err, 429);
    assert.equal(steps['stream.roundtrip'].detail, 'equal');
    assert.equal(steps['byref.global'].detail, 'a,b'); assert.equal(steps['byref.function'].detail, 'True/True');
    assert.equal(steps['builtin.chrw.negative'].detail, '-24691'); assert.equal(steps['builtin.selectcase'].detail, 'backslash');
    assert.equal(steps['objfunc.object'].detail, 'True'); assert.equal(steps['func.ifnot'].detail, 'True');
});
test('B0 probe modes select sections and verdicts', () => {
    const { buildProbeScript, NO_APC_SECTIONS, parseMode, judge } = require('../scripts/probe-catscript');
    assert.equal(parseMode([]), 'all');
    assert.equal(parseMode(['--check-only']), 'check-only');
    assert.equal(parseMode(['--no-apc']), 'no-apc');
    assert.equal(parseMode(['--apc']), 'apc');
    assert.equal(parseMode(['--full-catscript']), 'apc');
    assert.equal(parseMode(['--apc', '--full-catscript']), 'apc');
    assert.throws(() => parseMode(['--no-apc', '--apc']), /only one mode/);
    assert.throws(() => parseMode(['--catvbs']), /Unknown option/);
    const noApc = buildProbeScript('C:\\probe', NO_APC_SECTIONS);
    assert.doesNotMatch(noApc, /ProbeApc|ProbeVbe|MSAPC|\.VBE/);
    assert.match(noApc, /ProbeFso[\s\S]*ProbeStream/);
    assert.match(judge('check-only', { current: false }), /not loaded/);
    assert.match(judge('no-apc', { before: true, afterCatScript: true }), /^undetermined/);
    assert.equal(judge('no-apc', { before: false, afterCatScript: false }), 'CATScript engine did not load vbscript.dll');
    assert.match(judge('apc', { before: false, afterProbe: false, afterGeneratedList: true }), /^generated list/);
    assert.match(judge('all', { before: false, afterCatScript: false, afterCatvbs: true }), /B1 viable/);
});
test('B0 probe reads process modules with 64-bit PowerShell', { skip: process.platform !== 'win32' }, () => {
    const { listModules } = require('../scripts/probe-catscript');
    const result = listModules('node');
    assert.equal(result.error, null);
    assert.equal(result.is64, true);
    const self = result.processes.find(p => p.pid === process.pid);
    assert.ok(self, 'current node process listed');
    assert.ok(self.modules.some(m => m.name.toLowerCase() === 'ntdll.dll' && /ntdll\.dll$/i.test(m.path)));
});
