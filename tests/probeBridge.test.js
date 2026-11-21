const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, encodeToken, callArgs, matchStep, codepoints, bridgePaths, fileState, copyBridge, judgeBridge, describeStart,
    STEPS, STEP_NAMES, PROBE_SOURCE, ECHO_ARGUMENT, BRIDGE_MODULE, PROBE_MODULE } = require('../scripts/probe-bridge');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
const step = name => STEPS.find(s => s.name === name);

test('bridge probe parses --check-only and step names', () => {
    assert.equal(parseArgs(['-h']).mode, 'help');
    assert.deepEqual(parseArgs([]), { mode: 'run', steps: ['version', 'install', 'echo', 'projects', 'self-guard'] });
    assert.deepEqual(parseArgs(['self-guard', 'version']), { mode: 'run', steps: ['version', 'self-guard'] });
    assert.deepEqual(parseArgs(['--check-only']), { mode: 'check-only', steps: [] });
    assert.throws(() => parseArgs(['--check-only', 'version']), /takes no steps/);
    assert.throws(() => parseArgs(['--call']), /Unknown/);
    assert.throws(() => parseArgs(['ping']), /Unknown/);
    assert.deepEqual(STEP_NAMES, ['version', 'install', 'echo', 'projects', 'self-guard']);
});
test('bridge probe steps call the expected procedures', () => {
    assert.deepEqual(STEPS.map(s => [s.name, s.module, s.func, s.args.length]), [
        ['version', BRIDGE_MODULE, 'C5D_Version', 0],
        ['install', BRIDGE_MODULE, 'C5D_Install', 2],
        ['echo', PROBE_MODULE, 'C5D_ProbeEcho', 1],
        ['projects', PROBE_MODULE, 'C5D_ProbeProjects', 0],
        ['self-guard', BRIDGE_MODULE, 'C5D_Install', 2]
    ]);
    assert.deepEqual(step('install').args, ['C5D_ProbeModule', PROBE_SOURCE]);
    assert.deepEqual(step('self-guard').args, ['C5D_Bridge_Push', 'x']);
    assert.equal(ECHO_ARGUMENT, '日本語テスト✓');
    assert.ok(matchStep(step('version'), 'c5d-bridge:1'));
    assert.ok(!matchStep(step('version'), 'c5d-bridge:2'));
    assert.ok(matchStep(step('echo'), 'probe:Long:日本語テスト✓'));
    assert.ok(!matchStep(step('echo'), 'probe:Long:日本語テスト?'));
    assert.ok(matchStep(step('projects'), 'count=3;err=0'));
    assert.ok(!matchStep(step('projects'), 'count=-1;err=429'));
    assert.ok(matchStep(step('self-guard'), 'error:self'));
    assert.ok(!matchStep(step('install'), null));
    assert.equal(codepoints('a日✓?'), 'U+65E5 U+2713');
});
test('C5D_ProbeModule source is ASCII VBA with CRLF and the two functions', () => {
    assert.ok(/^[\x20-\x7e\r\n]*$/.test(PROBE_SOURCE));
    assert.ok(!/[^\r]\n/.test(PROBE_SOURCE));
    assert.ok(PROBE_SOURCE.startsWith('Option Explicit\r\n'));
    assert.ok(!/^Attribute /m.test(PROBE_SOURCE));
    assert.ok(!/\bMsgBox\b/i.test(PROBE_SOURCE));
    assert.ok(PROBE_SOURCE.includes('Public Function C5D_ProbeEcho(ByVal a As Variant) As Variant\r\n    Dim n As Long\r\n    n = 7\r\n'));
    assert.ok(PROBE_SOURCE.includes('C5D_ProbeEcho = "probe:" & TypeName(n) & ":" & CStr(a)'));
    assert.ok(PROBE_SOURCE.includes('Public Function C5D_ProbeProjects() As Variant'));
    assert.ok(PROBE_SOURCE.indexOf('"MSAPC.Apc.7.1"') < PROBE_SOURCE.indexOf('"MSAPC.Apc"'));
    assert.ok(PROBE_SOURCE.includes('"count=" & CStr(n) & ";err=" & CStr(Err.Number)'));
    assert.ok(!/\.Name\b/.test(PROBE_SOURCE), 'project names are not returned');
});
test('bridge probe verdicts', () => {
    const ok = (name, result) => ({ step: name, ok: true, match: true, result, loadedAfter: false });
    const all = [ok('version', 'c5d-bridge:1'), ok('install', 'ok'), ok('echo', 'probe:Long:' + ECHO_ARGUMENT), ok('projects', 'count=2;err=0'), ok('self-guard', 'error:self')];
    assert.equal(judgeBridge(false, all), 'bridge auto-loaded, self-install works, self-guard works, vbscript.dll not loaded');
    assert.equal(judgeBridge(false, all.map(e => e.step === 'install' ? { ...e, loadedAfter: true } : e)),
        'bridge auto-loaded, self-install works, self-guard works, vbscript.dll loaded (first seen after step install)');
    assert.match(judgeBridge(true, all), /vbscript\.dll undetermined/);
    const versionFailed = [{ step: 'version', ok: false, match: false, result: null, hresult: '0x80020009' },
        { step: 'install', skipped: 'requires version' }, { step: 'self-guard', skipped: 'requires version' }];
    assert.equal(judgeBridge(false, versionFailed), 'bridge not reachable: C5D_Version failed (0x80020009), '
        + 'self-install failed: C5D_Install not run (requires version), self-guard FAILED: C5D_Install(C5D_Bridge_Push) not run (requires version), vbscript.dll not loaded');
    assert.match(judgeBridge(false, [all[0], { ...all[1], match: false, result: 'error:apc' }]), /self-install failed: C5D_Install returned "error:apc"/);
    assert.match(judgeBridge(false, [all[0], all[1], { ...all[2], match: false, result: 'probe:Long:???????' }]), /self-install works but Unicode did not round-trip/);
    assert.match(judgeBridge(false, [all[0], all[1], { ...all[2], ok: false, match: false, result: null, hresult: '0x80004005' }]),
        /self-install returned ok but C5D_ProbeEcho failed \(0x80004005\)/);
    assert.match(judgeBridge(false, [all[0], { ...all[1], match: false, timedOut: true }]), /C5D_Install timed out/);
    assert.match(judgeBridge(false, [all[0], { ...all[4], match: false, result: 'ok' }]), /self-guard FAILED: C5D_Install\(C5D_Bridge_Push\) returned "ok"/);
    assert.match(judgeBridge(false, [all[0], all[1], all[2], { ...all[3], match: false, result: 'count=-1;err=429' }]), /C5D_ProbeProjects returned "count=-1;err=429"/);
    assert.equal(judgeBridge(false, [all[2]]), 'C5D_ProbeEcho works (Unicode round-trip), vbscript.dll not loaded');
});
test('bridge paths and copy (overwrite, hash, locked target)', t => {
    const dir = fixture(t);
    assert.deepEqual(bridgePaths({ LOCALAPPDATA: 'C:\\Users\\u\\AppData\\Local' }, 'C:\\repo'), {
        source: path.join('C:\\repo', 'resources', 'bridge', 'C5D_Bridge.catvba'),
        target: path.join('C:\\Users\\u\\AppData\\Local', 'Cat5Dev', 'bridge', 'C5D_Bridge.catvba')
    });
    assert.throws(() => bridgePaths({}), /LOCALAPPDATA/);
    const source = path.join(dir, 'src', 'C5D_Bridge.catvba'), target = path.join(dir, 'dst', 'bridge', 'C5D_Bridge.catvba');
    assert.throws(() => copyBridge(source, target), /does not exist/);
    fs.mkdirSync(path.dirname(source), { recursive: true });
    fs.writeFileSync(source, Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 1, 2, 3]));
    const first = copyBridge(source, target);
    assert.equal(first.ok, true);
    assert.equal(first.before.exists, false);
    assert.equal(first.after.sha256, first.source.sha256);
    assert.match(first.source.sha256, /^[0-9a-f]{64}$/);
    fs.writeFileSync(target, 'old');
    const second = copyBridge(source, target);
    assert.equal(second.ok, true);
    assert.equal(second.before.size, 3);
    assert.deepEqual(fs.readFileSync(target), fs.readFileSync(source));
    const locked = copyBridge(source, target, () => { throw Object.assign(new Error('resource busy or locked'), { code: 'EBUSY' }); });
    assert.equal(locked.locked, true);
    assert.equal(locked.ok, false);
    assert.equal(locked.code, 'EBUSY');
    assert.throws(() => copyBridge(source, target, () => { throw Object.assign(new Error('no space'), { code: 'ENOSPC' }); }), /no space/);
    assert.equal(fileState(path.join(dir, 'none')).exists, false);
    assert.match(describeStart('2026-11-21T10:00:00.0000000+09:00', Date.parse('2026-11-21T10:30:00+09:00')), /\(30 min ago\)/);
    assert.equal(describeStart(null), 'start time unknown');
});
test('vbacall passes Unicode, CR/LF and quotes from node to eval-probe.exe intact', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const library = 'C:\\Users\\ユーザー ✓\\AppData\\Local\\Cat5Dev\\bridge\\C5D_Bridge.catvba';
    for (const s of STEPS) {
        const args = callArgs(s, library);
        assert.deepEqual(args.slice(0, 3), ['--label', s.name, 'vbacall']);
        assert.ok(args.slice(3).every(arg => /^b64:[A-Za-z0-9+/=]*$/.test(arg)));
        const line = JSON.parse(execFileSync(exe, ['--dry-run', ...args], { windowsHide: true, encoding: 'utf8', timeout: 10000 }).trim());
        assert.deepEqual([line.step, line.method, line.library, line.libraryType, line.program, line.function, line.retryByRef, line.expected],
            [s.name, 'ExecuteScript', library, 2, s.module, s.func, true, null]);
        assert.deepEqual(line.arguments, s.args);
    }
    const tricky = 'a "quoted" \\ path\\\r\nline2 ✓ 𠮷';
    const line = JSON.parse(execFileSync(exe, ['--dry-run', 'vbacall', 'L', 'M', 'F', encodeToken(tricky), '', '--literal'],
        { windowsHide: true, encoding: 'utf8', timeout: 10000 }).trim());
    assert.equal(line.step, 'vbacall');
    assert.deepEqual(line.arguments, [tricky, '', '--literal']);
    assert.equal(encodeToken('✓'), 'b64:4pyT');
    assert.throws(() => execFileSync(exe, ['vbacall', 'L', 'M'], { windowsHide: true, stdio: 'pipe' }), error => error.status === 2);
    assert.throws(() => execFileSync(exe, ['--dry-run', 'vbacall', 'L', 'M', 'F', 'b64:@@'], { windowsHide: true, stdio: 'pipe' }), error => error.status === 2);
    assert.throws(() => execFileSync(exe, ['--label'], { windowsHide: true, stdio: 'pipe' }), error => error.status === 2);
});
