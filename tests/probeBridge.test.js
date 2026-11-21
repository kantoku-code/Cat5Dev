const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, encodeToken, callArgs, matchStep, matchCall, codepoints, bridgePaths, fileState, copyBridge, setReadOnly, judgeBridge, judgeUnload,
    describeStart, formatUnloadSummary, STEPS, STEP_NAMES, UNLOAD_STEPS, CLEANUP_STEPS, PROBE_SOURCE, UNLOAD_SOURCE, READONLY_SOURCE, ECHO_ARGUMENT, BRIDGE_MODULE,
    PROBE_MODULE, UNLOAD_MODULE, READONLY_MODULE, UNLOAD_RESULT, hashedBridgeName, hashedBridgePath, lockProbe, prepareHashedBridge, judgeReadonlyFresh,
    runReadonlyFresh, runReadonlyFreshCheck, formatReadonlyCheck, dispatch, READONLY_FRESH_STEPS, READONLY_FRESH_SOURCE } = require('../scripts/probe-bridge');
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

test('bridge probe parses --unload-experiment / --unload-cleanup', () => {
    assert.deepEqual(parseArgs(['--unload-experiment']), { mode: 'unload-experiment', steps: UNLOAD_STEPS.map(s => s.name) });
    assert.deepEqual(parseArgs(['--unload-cleanup']), { mode: 'unload-cleanup', steps: ['cleanup-install', 'cleanup-unload'] });
    assert.throws(() => parseArgs(['--unload-experiment', 'version']), /takes no steps/);
    assert.throws(() => parseArgs(['--unload-experiment', '--unload-cleanup']), /cannot be combined/);
    assert.throws(() => parseArgs(['--check-only', '--unload-cleanup']), /cannot be combined/);
    assert.deepEqual(parseArgs([]).steps, STEP_NAMES, 'default mode is unchanged');
});
test('unload experiment steps run in the agreed order', () => {
    const shape = s => [s.name, s.local ? 'local' : `${s.module}.${s.func}`, s.requires ?? []];
    assert.deepEqual(UNLOAD_STEPS.map(shape), [
        ['pre', 'C5D_Bridge_Push.C5D_Version', []],
        ['install-unloader', 'C5D_Bridge_Push.C5D_Install', []],
        ['apc-info', 'C5D_ProbeUnload.C5D_ApcProjectsInfo', ['install-unloader']],
        ['unload', 'C5D_ProbeUnload.C5D_UnloadSelf', ['install-unloader']],
        ['after-unload-probe', 'C5D_ProbeModule.C5D_ProbeEcho', []],
        ['reload-version', 'C5D_Bridge_Push.C5D_Version', []],
        ['install-unloader-2', 'C5D_Bridge_Push.C5D_Install', ['reload-version']],
        ['unload-2', 'C5D_ProbeUnload.C5D_UnloadSelf', ['install-unloader-2']],
        ['set-readonly', 'local', []],
        ['readonly-version', 'C5D_Bridge_Push.C5D_Version', []],
        ['readonly-install', 'C5D_Bridge_Push.C5D_Install', []],
        ['readonly-call', 'C5D_ProbeReadonly.C5D_ProbeReadonlyEcho', ['readonly-install']]
    ]);
    assert.deepEqual(CLEANUP_STEPS.map(shape), [
        ['cleanup-install', 'C5D_Bridge_Push.C5D_Install', []],
        ['cleanup-unload', 'C5D_ProbeUnload.C5D_UnloadSelf', ['cleanup-install']]
    ]);
    const get = name => UNLOAD_STEPS.find(s => s.name === name);
    for (const name of ['install-unloader', 'install-unloader-2']) assert.deepEqual(get(name).args, [UNLOAD_MODULE, UNLOAD_SOURCE]);
    assert.deepEqual(CLEANUP_STEPS[0].args, [UNLOAD_MODULE, UNLOAD_SOURCE]);
    assert.deepEqual(get('readonly-install').args, [READONLY_MODULE, READONLY_SOURCE]);
    assert.equal(get('unload').expected, 'removed;err=0');
    assert.equal(UNLOAD_RESULT, 'removed;err=0');
    // after-unload-probe matches unless C5D_ProbeModule still answers
    const probe = get('after-unload-probe');
    assert.deepEqual(probe.args, ['x']);
    assert.equal(matchCall(probe, { ok: false, result: null, hresult: '0x80020006' }), true);
    assert.equal(matchCall(probe, { ok: true, result: '' }), true);
    assert.equal(matchCall(probe, { ok: true, result: 'probe:Long:x' }), false);
    assert.equal(matchCall(probe, null), false);
    assert.equal(matchCall(get('unload'), { ok: true, result: 'removed;err=0' }), true);
    assert.equal(matchCall(get('unload'), { ok: false, result: 'removed;err=0' }), false);
    assert.equal(matchCall(get('apc-info'), { ok: true, result: 'apcProjects=2|a:ApcProject:vb=a;err=0' }), true);
    assert.equal(matchCall(get('readonly-call'), { ok: true, result: 'readonly:x' }), true);
});
test('C5D_ProbeUnload / C5D_ProbeReadonly sources are safe ASCII VBA', () => {
    for (const source of [UNLOAD_SOURCE, READONLY_SOURCE]) {
        assert.ok(/^[\x20-\x7e\r\n]*$/.test(source));
        assert.ok(!/[^\r]\n/.test(source));
        assert.ok(source.startsWith('Option Explicit\r\n'));
        assert.ok(!/^Attribute /m.test(source));
        assert.ok(!/\bMsgBox\b|\bInputBox\b|^\s*Stop\s*$|^\s*End\s*$/im.test(source));
    }
    const functions = UNLOAD_SOURCE.split('Public Function ').slice(1);
    assert.deepEqual(functions.map(f => f.slice(0, f.indexOf('('))), ['C5D_UnloadSelf', 'C5D_ApcProjectsInfo']);
    for (const f of functions) assert.ok(f.includes(') As Variant\r\n    On Error Resume Next\r\n'), 'On Error first');
    const [unload, info] = functions;
    assert.ok(unload.indexOf('"MSAPC.Apc.7.1"') < unload.indexOf('"MSAPC.Apc"'));
    assert.ok(unload.includes('Set vbps = apc.VBE.VBProjects'));
    assert.ok(unload.includes('nm = CStr(proj.Name)\r\n'));
    assert.ok(unload.includes('If nm = "C5D_Bridge" Then'));
    assert.ok(unload.includes('    Err.Clear\r\n    vbps.Remove found\r\n    If Err.Number = 0 Then\r\n        C5D_UnloadSelf = "removed;err=0"'));
    assert.ok(unload.includes('"error:remove:" & CStr(Err.Number) & ":" & Err.Description'));
    assert.ok(info.includes('Set aps = apc.Projects'));
    assert.ok(info.includes('s = "apcProjects=" & CStr(n)'));
    assert.ok(info.includes('tn = TypeName(ap)'));
    assert.ok(!/\.Remove\b/.test(info), 'apc-info does not unload');
    // with Resume Next an error inside an If condition runs the Then branch: conditions only use variables / Is Nothing / Err.Number
    for (const line of UNLOAD_SOURCE.split('\r\n').filter(l => /^\s*If /.test(l))) {
        assert.match(line, /^\s*If ((Not )?\w+ Is Nothing|nm = "C5D_Bridge"|Err\.Number (= 0|<> 0)) Then$/, line);
    }
    assert.ok(READONLY_SOURCE.includes('C5D_ProbeReadonlyEcho = "readonly:" & CStr(a)'));
});
test('copyBridge clears and restores a read-only target; setReadOnly sets it and probes the lock', t => {
    const dir = fixture(t);
    const source = path.join(dir, 'src.catvba'), target = path.join(dir, 'bridge', 'C5D_Bridge.catvba');
    fs.writeFileSync(source, 'new');
    fs.mkdirSync(path.dirname(target));
    fs.writeFileSync(target, 'old');
    t.after(() => { try { fs.chmodSync(target, 0o666); } catch { /* already removed */ } });
    const ro = setReadOnly(target);
    assert.equal(ro.ok, true);
    assert.deepEqual([ro.before.readOnly, ro.after.readOnly], [false, true]);
    assert.deepEqual(ro.lockProbe, { writable: true, code: null, message: null });
    assert.equal(fs.readFileSync(target, 'utf8'), 'old', 'the lock probe writes nothing');
    assert.deepEqual(setReadOnly(target).lockProbe, { skipped: 'already read-only' });
    const copied = copyBridge(source, target);
    assert.equal(copied.ok, true);
    assert.deepEqual([copied.wasReadOnly, copied.readOnlyRestored, copied.after.readOnly], [true, true, true]);
    assert.equal(fs.readFileSync(target, 'utf8'), 'new');
    const busy = copyBridge(source, target, () => { throw Object.assign(new Error('busy'), { code: 'EBUSY' }); });
    assert.deepEqual([busy.locked, busy.wasReadOnly, busy.readOnlyRestored, fileState(target).readOnly], [true, true, true, true]);
    fs.chmodSync(target, 0o666);
    const plain = copyBridge(source, target);
    assert.deepEqual([plain.wasReadOnly, plain.readOnlyRestored, plain.after.readOnly], [false, null, false]);
    const locked = setReadOnly(target, { open: () => { throw Object.assign(new Error('busy'), { code: 'EBUSY' }); },
        chmod: () => { throw Object.assign(new Error('denied'), { code: 'EPERM' }); } });
    assert.deepEqual([locked.ok, locked.code, locked.lockProbe.writable, locked.lockProbe.code], [false, 'EPERM', false, 'EBUSY']);
    assert.equal(setReadOnly(path.join(dir, 'none')).ok, false);
});
test('unload experiment verdicts', () => {
    const ok = (step, result) => ({ step, ok: true, match: true, result, loadedAfter: false });
    const all = [ok('pre', 'c5d-bridge:1'), ok('install-unloader', 'ok'), ok('apc-info', 'apcProjects=1|C5D_Bridge:ApcProject:vb=C5D_Bridge;err=0'),
        ok('unload', 'removed;err=0'), { step: 'after-unload-probe', ok: false, match: true, result: null, hresult: '0x80020006', message: 'unknown' },
        ok('reload-version', 'c5d-bridge:1'), ok('install-unloader-2', 'ok'), ok('unload-2', 'removed;err=0'),
        { step: 'set-readonly', ok: true, match: true, lockProbe: { writable: true } }, ok('readonly-version', 'c5d-bridge:1'),
        ok('readonly-install', 'ok'), ok('readonly-call', 'readonly:x')];
    assert.equal(judgeUnload(false, all), 'pre: bridge responded, apc-info: apcProjects=1|C5D_Bridge:ApcProject:vb=C5D_Bridge;err=0, unload: removed, '
        + 'unload-2: removed, after-unload-probe: C5D_ProbeModule is gone (0x80020006 unknown), reload-version: re-loaded by ExecuteScript, '
        + 'set-readonly: read-only set (file not locked), readonly-version: read-only .catvba loads and answers, '
        + 'readonly-call: module installed into the read-only project works, vbscript.dll not loaded');
    const notUnloaded = all.map(e => e.step === 'after-unload-probe' ? { ...ok(e.step, 'probe:Long:x'), match: false } : e);
    assert.match(judgeUnload(false, notUnloaded), /NOT unloaded/);
    const crashed = [all[0], all[1], all[2], { step: 'unload', ok: false, match: false, result: null, hresult: '0x800706BE', crashed: true },
        { step: 'after-unload-probe', skipped: 'aborted: CATIA crashed after step unload' }];
    assert.match(judgeUnload(false, crashed), /unload: C5D_UnloadSelf failed \(0x800706BE\).*CATIA crashed after step unload/);
    const locked = all.map(e => e.step === 'set-readonly' ? { ...e, lockProbe: { writable: false, code: 'EBUSY' } } : e);
    assert.match(judgeUnload(false, locked), /set-readonly: read-only set \(file LOCKED \(EBUSY\)\)/);
    const dialog = all.slice(0, 9).concat([{ step: 'readonly-version', timedOut: true, match: false }]);
    assert.match(judgeUnload(true, dialog), /readonly-version: C5D_Version timed out.*timed out at step readonly-version.*vbscript\.dll undetermined/);
    assert.equal(judgeUnload(false, [ok('cleanup-install', 'ok'), ok('cleanup-unload', 'removed;err=0')]), 'cleanup-unload: removed, vbscript.dll not loaded');
});
test('unload summary shows each step with CATIA alive and vbscript.dll', () => {
    const { KEY_DLLS } = require('../scripts/probe-catscript');
    const snapshot = (label, vbs) => ({ label, cnextPids: [10], powershell: { error: null }, loaded: vbs,
        processes: [{ pid: 10, error: null, moduleCount: 2, moduleNames: ['CNEXT.exe', ...(vbs ? ['vbscript.dll'] : [])] }],
        dlls: Object.fromEntries(KEY_DLLS.map(dll => [dll, { loaded: dll === 'vbscript.dll' && vbs, byModules: dll === 'vbscript.dll' && vbs, byTasklist: dll === 'vbscript.dll' && vbs }])) });
    const file = (readOnly, sha) => ({ path: 'C:\\b\\C5D_Bridge.catvba', exists: true, readOnly, sha256: sha });
    const text = formatUnloadSummary({ mode: 'unload-experiment', reportPath: 'r.json', notes: ['n1'], library: 'C:\\b\\C5D_Bridge.catvba',
        stoppedForUser: true, cnext: { error: null, processes: [{ pid: 10, startedAt: null }] },
        bridge: { source: file(false, 'a'.repeat(64)), target: file(false, 'a'.repeat(64)), targetAtEnd: file(true, 'a'.repeat(64)) },
        snapshots: { before: snapshot('before', false), 'after:unload': snapshot('after:unload', false), 'after:set-readonly': snapshot('after:set-readonly', true) },
        runs: [
            { step: 'unload', call: 'C5D_ProbeUnload.C5D_UnloadSelf', match: false, durationMs: 900, catiaAlive: true, argumentsIntact: true,
                result: { ok: false, result: null, hresult: '0x80020009', message: 'boom', elapsedMs: 800 } },
            { step: 'set-readonly', call: '(local) set-readonly', local: true, ok: true, match: true, durationMs: 3, catiaAlive: true,
                before: { readOnly: false }, after: { readOnly: true }, lockProbe: { writable: false, code: 'EBUSY', message: 'busy' } },
            { step: 'readonly-version', call: 'C5D_Bridge_Push.C5D_Version', timedOut: true, catiaAlive: null },
            { step: 'readonly-install', call: 'C5D_Bridge_Push.C5D_Install', skipped: 'aborted after readonly-version timed out' }
        ],
        summary: { verdict: 'v' } });
    assert.match(text, /unload: ok=no match=no result=null elapsed=800ms \(process 900ms\) {2}catia=alive vbscript\.dll=no/);
    assert.match(text, /hresult=0x80020009 message=boom/);
    assert.match(text, /set-readonly: ok=yes readOnly no -> yes {2}lock probe: LOCKED \(EBUSY: busy\) {2}3ms {2}catia=alive vbscript\.dll=yes/);
    assert.match(text, /readonly-version: TIMEOUT after 60s.*catia=\? vbscript\.dll=-/);
    assert.match(text, /readonly-install: SKIPPED/);
    assert.match(text, /start: sha256=aaaaaaaaaaaaaaaa readOnly=no \(= source\)/);
    assert.match(text, /at end: sha256=aaaaaaaaaaaaaaaa readOnly=yes \(unchanged\)/);
    assert.match(text, /verdict: v\nnote: n1\n\nNEXT: .*--unload-cleanup/s);
});
test('vbacall arguments of the unload / readonly-fresh steps survive the command line', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const library = 'C:\\Users\\u\\AppData\\Local\\Cat5Dev\\bridge\\C5D_Bridge.catvba';
    for (const s of [...UNLOAD_STEPS, ...CLEANUP_STEPS, ...READONLY_FRESH_STEPS].filter(s => !s.local)) {
        const args = callArgs(s, library, { noRetry: true });
        assert.deepEqual(args.slice(0, 4), ['--no-retry', '--label', s.name, 'vbacall']);
        const line = JSON.parse(execFileSync(exe, ['--dry-run', ...args], { windowsHide: true, encoding: 'utf8', timeout: 10000 }).trim());
        assert.deepEqual([line.step, line.library, line.libraryType, line.program, line.function, line.retryByRef], [s.name, library, 2, s.module, s.func, false]);
        assert.deepEqual(line.arguments, s.args);
    }
    const retry = JSON.parse(execFileSync(exe, ['--dry-run', ...callArgs(UNLOAD_STEPS[0], library)], { windowsHide: true, encoding: 'utf8', timeout: 10000 }).trim());
    assert.equal(retry.retryByRef, true, 'without --no-retry the ByRef retry stays');
});

test('bridge probe parses --readonly-fresh / --readonly-fresh-check and dispatches without CATIA', () => {
    assert.deepEqual(parseArgs(['--readonly-fresh']), { mode: 'readonly-fresh', steps: ['version', 'install', 'projects', 'call'] });
    assert.deepEqual(parseArgs(['--readonly-fresh-check']), { mode: 'readonly-fresh-check', steps: [] });
    assert.throws(() => parseArgs(['--readonly-fresh', 'version']), /takes no steps/);
    assert.throws(() => parseArgs(['--readonly-fresh-check', 'call']), /Unknown/);
    assert.throws(() => parseArgs(['--readonly-fresh', '--readonly-fresh-check']), /cannot be combined/);
    assert.throws(() => parseArgs(['--readonly-fresh', '--unload-experiment']), /cannot be combined/);
    const d = mode => dispatch({ mode });
    assert.deepEqual([d('readonly-fresh').run, d('readonly-fresh').format, d('readonly-fresh').prefix],
        [runReadonlyFresh, formatUnloadSummary, 'probe-bridge-readonly-fresh']);
    assert.deepEqual([d('readonly-fresh-check').run, d('readonly-fresh-check').format, d('readonly-fresh-check').prefix],
        [runReadonlyFreshCheck, formatReadonlyCheck, 'probe-bridge-readonly-fresh-check']);
    assert.equal(d('unload-cleanup').prefix, 'probe-bridge-unload-cleanup');
    assert.equal(d('run').prefix, 'probe-bridge');
    assert.equal(d('check-only').prefix, 'probe-bridge');
    assert.throws(() => d('help'), /no runner/);
});
test('readonly-fresh uses a hash-named copy, not the plain C5D_Bridge.catvba', () => {
    const sha = '0123abcd' + 'e'.repeat(56);
    assert.equal(hashedBridgeName(sha), 'C5D_Bridge-0123abcd.catvba');
    const env = { LOCALAPPDATA: 'C:\\Users\\u\\AppData\\Local' };
    assert.equal(hashedBridgePath(sha, env), path.join('C:\\Users\\u\\AppData\\Local', 'Cat5Dev', 'bridge', 'C5D_Bridge-0123abcd.catvba'));
    assert.notEqual(hashedBridgePath(sha, env), bridgePaths(env, 'C:\\repo').target);
    assert.throws(() => hashedBridgeName('0123abcd'), /not a sha256/);
    assert.throws(() => hashedBridgeName(sha.toUpperCase()), /not a sha256/);
    assert.throws(() => hashedBridgePath(sha, {}), /LOCALAPPDATA/);
});
test('readonly-fresh steps: version, install C5D_ProbeReadonly, projects, call (no requires)', () => {
    assert.deepEqual(READONLY_FRESH_STEPS.map(s => [s.name, `${s.module}.${s.func}`, s.requires ?? []]), [
        ['version', 'C5D_Bridge_Push.C5D_Version', []],
        ['install', 'C5D_Bridge_Push.C5D_Install', []],
        ['projects', 'C5D_ProbeReadonly.C5D_ProbeProjects', []],
        ['call', 'C5D_ProbeReadonly.C5D_Probe', []]
    ]);
    const get = name => READONLY_FRESH_STEPS.find(s => s.name === name);
    assert.deepEqual(get('install').args, [READONLY_MODULE, READONLY_FRESH_SOURCE]);
    assert.equal(READONLY_MODULE, 'C5D_ProbeReadonly');
    assert.equal(matchCall(get('version'), { ok: true, result: 'c5d-bridge:1' }), true);
    assert.equal(matchCall(get('call'), { ok: true, result: 'readonly-fresh:ok' }), true);
    assert.equal(matchCall(get('call'), { ok: false, result: null, hresult: '0x80020006' }), false);
    assert.equal(matchCall(get('projects'), { ok: true, result: 'count=2;err=0' }), true);
    assert.equal(matchCall(get('projects'), { ok: true, result: 'count=-1;err=429' }), false);
    const source = READONLY_FRESH_SOURCE;
    assert.ok(/^[\x20-\x7e\r\n]*$/.test(source));
    assert.ok(!/[^\r]\n/.test(source));
    assert.ok(source.startsWith('Option Explicit\r\nPublic Function C5D_Probe() As Variant\r\n    C5D_Probe = "readonly-fresh:ok"\r\nEnd Function\r\n'));
    assert.ok(source.endsWith(PROBE_SOURCE.slice(PROBE_SOURCE.indexOf('Public Function C5D_ProbeProjects'))), 'C5D_ProbeProjects as in C5D_ProbeModule');
    assert.equal(source.split('Option Explicit').length, 2);
    assert.ok(!/\bMsgBox\b|\bInputBox\b|^\s*Stop\s*$|^\s*End\s*$|\.Remove\b/im.test(source));
});
test('prepareHashedBridge copies, reuses or replaces and sets read-only before CATIA; a locked target is reported', t => {
    const dir = fixture(t);
    const source = path.join(dir, 'src.catvba'), target = path.join(dir, 'bridge', 'C5D_Bridge-00000000.catvba');
    t.after(() => { try { fs.chmodSync(target, 0o666); } catch { /* removed */ } });
    assert.throws(() => prepareHashedBridge(source, target), /does not exist/);
    fs.writeFileSync(source, 'new');
    const copied = prepareHashedBridge(source, target);
    assert.deepEqual([copied.ok, copied.action, copied.before.exists, copied.after.readOnly, copied.after.sha256 === copied.source.sha256],
        [true, 'copied', false, true, true]);
    assert.deepEqual(copied.readOnly.lockProbe, { writable: true, code: null, message: null });
    const reused = prepareHashedBridge(source, target, () => assert.fail('no copy when the hash matches'));
    assert.deepEqual([reused.ok, reused.action, reused.after.readOnly], [true, 'reused', true]);
    fs.chmodSync(target, 0o666);
    assert.equal(prepareHashedBridge(source, target, () => assert.fail('no copy')).after.readOnly, true, 'reused copy is made read-only');
    fs.chmodSync(target, 0o666);
    fs.writeFileSync(target, 'saved by the VBA editor');
    fs.chmodSync(target, 0o444);
    const replaced = prepareHashedBridge(source, target);
    assert.deepEqual([replaced.ok, replaced.action, replaced.before.readOnly, replaced.after.readOnly], [true, 'replaced', true, true]);
    assert.equal(fs.readFileSync(target, 'utf8'), 'new');
    fs.chmodSync(target, 0o666);
    fs.writeFileSync(target, 'other');
    fs.chmodSync(target, 0o444);
    const locked = prepareHashedBridge(source, target, () => { throw Object.assign(new Error('busy'), { code: 'EBUSY' }); });
    assert.deepEqual([locked.ok, locked.locked, locked.code, locked.action, locked.after.readOnly], [false, true, 'EBUSY', 'replaced', true]);
    assert.equal(fs.readFileSync(target, 'utf8'), 'other');
    assert.throws(() => prepareHashedBridge(source, target, () => { throw Object.assign(new Error('no space'), { code: 'ENOSPC' }); }), /no space/);
    fs.chmodSync(target, 0o666);
    fs.rmSync(target);
    const notReadOnly = prepareHashedBridge(source, target, fs.copyFileSync, () => { /* attribute not set */ });
    assert.equal(notReadOnly.ok, false);
    assert.match(notReadOnly.message, /read-only attribute not set/);
});
test('lockProbe detects a file opened by another handle without writing', { skip: process.platform !== 'win32' }, t => {
    const file = path.join(fixture(t), 'x.catvba');
    assert.deepEqual(lockProbe(file), { skipped: 'missing' });
    assert.deepEqual(lockProbe(file, undefined, undefined, 'linux'), { skipped: 'not Windows' });
    fs.writeFileSync(file, 'abc');
    fs.chmodSync(file, 0o444);
    t.after(() => { try { fs.chmodSync(file, 0o666); } catch { /* removed */ } });
    assert.deepEqual(lockProbe(file), { inUse: false, code: null, message: null }, 'works on a read-only file');
    const fd = fs.openSync(file, 'r');
    try { assert.deepEqual([lockProbe(file).inUse, lockProbe(file).code], [true, 'EBUSY']); } finally { fs.closeSync(fd); }
    assert.equal(lockProbe(file).inUse, false);
    assert.equal(lockProbe(file, () => { throw Object.assign(new Error('denied'), { code: 'EACCES' }); }).inUse, null);
    assert.equal(fs.readFileSync(file, 'utf8'), 'abc');
});
test('readonly-fresh verdicts', () => {
    const ok = (step, result) => ({ step, ok: true, match: true, result, loadedAfter: false });
    const all = [ok('version', 'c5d-bridge:1'), ok('install', 'ok'), ok('projects', 'count=1;err=0'), ok('call', 'readonly-fresh:ok')];
    assert.equal(judgeReadonlyFresh(false, all), 'version: read-only .catvba auto-loaded and answers, install: ok, projects: count=1;err=0, '
        + 'call: C5D_ProbeReadonly.C5D_Probe works, vbscript.dll not loaded');
    const failed = [{ step: 'version', ok: false, match: false, result: null, hresult: '0x80020009' }, { ...ok('install', 'error:vbe'), match: false },
        { ...ok('projects', 'count=-1;err=429'), match: false }, { step: 'call', ok: false, match: false, result: null, hresult: '0x80020006', loadedAfter: true }];
    assert.equal(judgeReadonlyFresh(false, failed), 'version: C5D_Version failed (0x80020009), install: C5D_Install returned "error:vbe", '
        + 'projects: count=-1;err=429, call: C5D_Probe failed (0x80020006), vbscript.dll loaded (first seen after step call)');
    const dialog = [all[0], { step: 'install', timedOut: true, match: false }, { step: 'projects', skipped: 'aborted after install timed out' },
        { step: 'call', skipped: 'aborted after install timed out' }];
    assert.match(judgeReadonlyFresh(true, dialog), /install: C5D_Install timed out.*call: C5D_Probe not run.*timed out at step install.*vbscript\.dll undetermined/);
});
test('--readonly-fresh-check records the hash-named copy without CATIA and compares it with the last run', t => {
    const dir = fixture(t);
    const root = path.join(dir, 'repo'), env = { LOCALAPPDATA: path.join(dir, 'local') }, reportsDir = path.join(dir, 'temp');
    const cnext = () => ({ processes: [{ pid: 42, startedAt: null, path: null }], error: null });
    const newReport = () => ({ mode: 'readonly-fresh-check', reportPath: 'r.json', notes: [], snapshots: {} });
    assert.throws(() => runReadonlyFreshCheck(newReport(), { env, root, cnext, reportsDir }), /does not exist/);
    const source = path.join(root, 'resources', 'bridge', 'C5D_Bridge.catvba');
    fs.mkdirSync(path.dirname(source), { recursive: true });
    fs.writeFileSync(source, Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 9]));
    const missing = newReport();
    runReadonlyFreshCheck(missing, { env, root, cnext, reportsDir });
    assert.equal(missing.bridge.target.exists, false);
    assert.match(missing.notes.join('\n'), /--readonly-fresh first/);
    assert.equal(missing.summary.verdict, 'check: target missing');
    const prepared = prepareHashedBridge(source, hashedBridgePath(fileState(source).sha256, env));
    const target = prepared.after.path;
    t.after(() => { try { fs.chmodSync(target, 0o666); } catch { /* removed */ } });
    assert.match(path.basename(target), /^C5D_Bridge-[0-9a-f]{8}\.catvba$/);
    fs.mkdirSync(reportsDir);
    fs.writeFileSync(path.join(reportsDir, 'probe-bridge-readonly-fresh-20261121-100000.json'), JSON.stringify({ bridge: { targetAtEnd: { path: 'old' } } }));
    fs.writeFileSync(path.join(reportsDir, 'probe-bridge-readonly-fresh-20261121-110000.json'), JSON.stringify({ bridge: { targetAtEnd: prepared.after } }));
    fs.writeFileSync(path.join(reportsDir, 'probe-bridge-readonly-fresh-check-20261121-120000.json'), '{}');
    const same = newReport();
    runReadonlyFreshCheck(same, { env, root, cnext, reportsDir });
    assert.deepEqual([same.library, same.bridge.sameAsSource, same.bridge.target.readOnly, same.bridge.previous.changed], [target, true, true, false]);
    if (process.platform === 'win32') assert.equal(same.bridge.lock.inUse, false);
    assert.equal(path.basename(same.bridge.previous.report), 'probe-bridge-readonly-fresh-20261121-110000.json');
    assert.deepEqual(same.cnext.processes.map(p => p.pid), [42]);
    assert.deepEqual(same.notes, []);
    assert.match(same.summary.verdict, /^check: same as source, read-only yes, .*unchanged since the --readonly-fresh run$/);
    fs.chmodSync(target, 0o666);
    fs.writeFileSync(target, 'saved');
    const saved = newReport();
    runReadonlyFreshCheck(saved, { env, root, cnext, reportsDir });
    assert.deepEqual([saved.bridge.sameAsSource, saved.bridge.target.readOnly, saved.bridge.previous.changed], [false, false, true]);
    assert.match(saved.notes.join('\n'), /DIFFERS from the source.*\n.*is NOT read-only/);
    assert.match(saved.summary.verdict, /DIFFERS from source, read-only no, .*CHANGED since the --readonly-fresh run/);
    const text = formatReadonlyCheck(saved);
    assert.match(text, /bridge target: .*C5D_Bridge-[0-9a-f]{8}\.catvba sha256=[0-9a-f]{16} \(DIFFERS from source\)/);
    assert.match(text, /\n {2}modified=\S+ size=5 readOnly=no, /);
    assert.match(text, /at end of .*-110000\.json: modified=\S+ size=5 readOnly=yes \(CHANGED since\)/);
    assert.match(text, /CNEXT\.exe pid 42: start time unknown/);
});
