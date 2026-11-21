// --remove-experiment / --remove-experiment-risky of scripts/probe-bridge.js and the vbacall-obj step of EvalProbe.cs (no CATIA).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, callArgs, matchCall, dispatch, formatUnloadSummary, BRIDGE_MODULE, REMOVE_STEPS, REMOVE_RISKY_STEPS, REMOVE_SOURCE, B_SOURCE,
    TIMER_SOURCE, REMOVE_MODULE, B_MODULE, TIMER_MODULE, COPY_B_FILE, TIMER_LOG_TOKEN, materializeSteps, skipReason, parseInfo, describeInfo,
    judgeRemove, runRemove } = require('../scripts/probe-bridge');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');
const names = steps => steps.map(s => s.name);
const stepOf = (steps, name) => steps.find(s => s.name === name);

test('bridge probe parses --remove-experiment / --remove-experiment-risky and dispatches to runRemove', () => {
    assert.deepEqual(parseArgs(['--remove-experiment']), { mode: 'remove-experiment', steps: names(REMOVE_STEPS) });
    assert.deepEqual(parseArgs(['--remove-experiment-risky']), { mode: 'remove-experiment-risky', steps: names(REMOVE_RISKY_STEPS) });
    assert.throws(() => parseArgs(['--remove-experiment', 'version']), /takes no steps/);
    assert.throws(() => parseArgs(['--remove-experiment', '--remove-experiment-risky']), /cannot be combined/);
    assert.throws(() => parseArgs(['--remove-experiment', '--unload-experiment']), /cannot be combined/);
    for (const [mode, prefix] of [['remove-experiment', 'probe-bridge-remove'], ['remove-experiment-risky', 'probe-bridge-remove-risky']]) {
        const d = dispatch({ mode });
        assert.deepEqual([d.run, d.format, d.prefix], [runRemove, formatUnloadSummary, prefix]);
    }
});

test('remove steps: unique names, dependencies on earlier steps, every attempt has a check and a reload', () => {
    for (const steps of [REMOVE_STEPS, REMOVE_RISKY_STEPS]) {
        const list = names(steps);
        assert.equal(new Set(list).size, list.length);
        steps.forEach((s, i) => {
            for (const dep of [...(s.requires ?? []), ...(s.unless ?? []), ...(s.after ?? [])]) {
                assert.ok(list.indexOf(dep) >= 0 && list.indexOf(dep) < i, `${s.name} -> ${dep}`);
            }
        });
        const attempts = steps.filter(s => s.attempt);
        assert.ok(attempts.length >= 2);
        attempts.forEach((a, k) => {
            assert.ok(a.danger, `${a.name} has a danger level`);
            const check = stepOf(steps, `${a.name}-check`), reload = stepOf(steps, `${a.name}-reload`);
            assert.deepEqual([check.module, check.func, check.expectFailure, check.after], [REMOVE_MODULE, 'C5D_RmAlive', 'alive', [a.name]]);
            assert.deepEqual([reload.module, reload.func, reload.requires], [BRIDGE_MODULE, 'C5D_Version', [`${a.name}-check`]]);
            assert.ok(list.indexOf(check.name) > list.indexOf(a.name) && list.indexOf(reload.name) === list.indexOf(check.name) + 1);
            // once an attempt removed the project the later attempts are skipped
            for (const earlier of attempts.slice(0, k)) assert.ok((a.unless ?? []).includes(`${earlier.name}-check`), `${a.name} unless ${earlier.name}`);
        });
    }
    assert.deepEqual(REMOVE_STEPS.filter(s => s.attempt).map(s => s.name),
        ['inside-remove', 'held-remove', 'outside-remove', 'outside-close', 'outside-design-close', 'outside-end-close']);
    assert.deepEqual(REMOVE_RISKY_STEPS.filter(s => s.attempt).map(s => s.name), ['inside-close', 'timer-remove', 'timer-close']);
    assert.deepEqual(REMOVE_STEPS.filter(s => s.library === 'b').map(s => s.name), ['b-load', 'b-install', 'b-module-in-b', 'b-reload']);
    assert.ok(!REMOVE_RISKY_STEPS.some(s => s.library === 'b'));
    assert.ok(!REMOVE_STEPS.some(s => s.module === TIMER_MODULE || (s.args ?? []).includes(TIMER_SOURCE)), 'the safe mode never installs the Declare module');
    assert.deepEqual(stepOf(REMOVE_STEPS, 'install').args, [REMOVE_MODULE, REMOVE_SOURCE]);
    assert.deepEqual(stepOf(REMOVE_STEPS, 'b-install').args, [B_MODULE, B_SOURCE]);
    assert.deepEqual(stepOf(REMOVE_STEPS, 'b-remove').args, [COPY_B_FILE]);
    assert.equal(COPY_B_FILE, 'C5D_Bridge-probe-b.catvba');
    assert.ok(!/^C5D_Bridge-[0-9a-f]{8}\.catvba$/.test(COPY_B_FILE), 'copy B never looks like a production hash-named copy');
});

test('remove steps never close all projects or add projects', () => {
    const actions = [...REMOVE_STEPS, ...REMOVE_RISKY_STEPS].flatMap(s => s.actions ?? []);
    assert.ok(actions.length > 20);
    assert.ok(!actions.some(a => /Projects\.Close|\.Add\(|\.Open\(|\.Save(As|CopyAs)?\(/.test(a)), 'no IApcProjects.Close (closes every project), Add, Open or Save');
    for (const source of [REMOVE_SOURCE, TIMER_SOURCE]) {
        assert.ok(!/Projects\.Close|\.Close True|\.Save\b|VBProjects\.Add|\.Open\b/.test(source));
        for (const line of source.split('\r\n').filter(l => /\.Close\b/.test(l) && !/^\s*Close #f$/.test(l))) assert.match(line, /^\s*ap\.Close False$/);
    }
    // destructive actions of the outside attempts
    const destructive = REMOVE_STEPS.filter(s => s.attempt && s.actions).map(s => [s.name, s.actions[s.mustSucceed[0]]]);
    assert.deepEqual(destructive, [['outside-remove', 'call:$.Collection.Remove($)'], ['outside-close', 'call:$.Close(false)'],
        ['outside-design-close', 'call:$.Close(false)'], ['outside-end-close', 'call:$.Close(false)']]);
});

test('C5D_ProbeRemove / C5D_ProbeB / C5D_ProbeTimer sources are safe ASCII VBA', () => {
    for (const source of [REMOVE_SOURCE, B_SOURCE, TIMER_SOURCE]) {
        assert.ok(/^[\x20-\x7e\r\n]*$/.test(source), 'ASCII only');
        assert.ok(!/[^\r]\n/.test(source), 'CRLF');
        assert.ok(source.startsWith('Option Explicit\r\n'));
        assert.equal(source.split('Option Explicit').length, 2);
        assert.ok(!/^Attribute /m.test(source));
        assert.ok(!/\bMsgBox\b|\bInputBox\b|^\s*Stop\s*$|^\s*End\s*$/im.test(source));
        // with On Error Resume Next an error inside an If condition runs the Then branch: never read a property there
        for (const line of source.split('\r\n').filter(l => /^\s*(Else)?If .* Then/.test(l))) {
            const condition = /If (.*) Then/.exec(line)[1];
            assert.ok(!/[A-Za-z_)]\.[A-Za-z]/.test(condition), `property read in an If condition: ${line.trim()}`);
        }
        // every procedure with an object call handles errors
        for (const body of source.split(/\r\n(?=(?:Public|Private) (?:Function|Sub) )/).slice(1)) {
            if (/\.\w/.test(body.split('\r\n').slice(1).join('\n'))) assert.match(body, /\r\n {4}On Error Resume Next\r\n/, body.split('\r\n')[0]);
        }
    }
    assert.ok(!/Declare/.test(REMOVE_SOURCE + B_SOURCE));
    const declares = TIMER_SOURCE.split('\r\n').filter(l => /Declare/.test(l));
    assert.deepEqual(declares.map(l => /^Private Declare PtrSafe Function (\w+) Lib "user32"/.exec(l)?.[1]), ['SetTimer', 'KillTimer']);
    assert.ok(/AddressOf C5D_RmTimerProc/.test(TIMER_SOURCE));
    assert.ok(/Public Sub C5D_RmTimerProc\(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long\)/.test(TIMER_SOURCE));
    assert.ok(/\r\n {4}KillTimer 0, idEvent\r\n/.test(TIMER_SOURCE), 'the callback kills its timer first (fires once)');
    for (const fn of ['C5D_RmAlive', 'C5D_RmApcInfo', 'C5D_RmVbeInfo', 'C5D_RmApc', 'C5D_RmSelfApcProject', 'C5D_RmSelfVbProject', 'C5D_RmSelfPair',
        'C5D_RmRemoveSelfInside', 'C5D_RmHold', 'C5D_RmRemoveHeld', 'C5D_RmRemoveOther', 'C5D_RmCloseOther', 'C5D_RmCloseSelfInside']) {
        assert.ok(REMOVE_SOURCE.includes(`\r\nPublic Function ${fn}(`), fn);
    }
    const called = [...REMOVE_STEPS, ...REMOVE_RISKY_STEPS].filter(s => s.module === REMOVE_MODULE).map(s => s.func);
    for (const fn of called) assert.ok(REMOVE_SOURCE.includes(`Public Function ${fn}(`), `${fn} exists`);
    assert.match(REMOVE_SOURCE, /Set C5D_RmSelfApcProject = ap\r\n/);
    assert.match(REMOVE_SOURCE, /C5D_RmSelfPair = Array\(vbps, vbp\)\r\n/);
});

test('materializeSteps puts the timer log path into the schedule calls only', () => {
    const steps = materializeSteps(REMOVE_RISKY_STEPS, { timerLog: 'C:\\x\\remove-timer.log' });
    assert.deepEqual(steps.filter(s => s.func === 'C5D_RmSchedule').map(s => s.args), [['remove', 'C:\\x\\remove-timer.log', '500'], ['close', 'C:\\x\\remove-timer.log', '500']]);
    assert.ok(!JSON.stringify(steps).includes(TIMER_LOG_TOKEN));
    assert.ok(REMOVE_RISKY_STEPS.some(s => (s.args ?? []).includes(TIMER_LOG_TOKEN)), 'the definitions are not modified');
    assert.equal(materializeSteps(REMOVE_STEPS, { timerLog: 'x' })[1], REMOVE_STEPS[1]);
});

test('skipReason: abort, requires, unless, after', () => {
    const entries = [{ step: 'a', match: true }, { step: 'b', match: false }, { step: 'c', skipped: 'requires b' }];
    assert.equal(skipReason({ name: 'x' }, entries, null), null);
    assert.equal(skipReason({ name: 'x' }, entries, 'aborted after a timed out'), 'aborted after a timed out');
    assert.equal(skipReason({ name: 'x', requires: ['a'] }, entries, null), null);
    assert.equal(skipReason({ name: 'x', requires: ['a', 'b'] }, entries, null), 'requires b');
    assert.equal(skipReason({ name: 'x', unless: ['b'] }, entries, null), null);
    assert.equal(skipReason({ name: 'x', unless: ['a', 'b'] }, entries, null), 'not needed: a matched');
    assert.equal(skipReason({ name: 'x', after: ['a', 'b'] }, entries, null), null);
    assert.equal(skipReason({ name: 'x', after: ['c'] }, entries, null), 'c not run');
    assert.equal(skipReason({ name: 'x', after: ['zz'] }, entries, null), 'zz not run');
});

test('matchCall of object steps: an object must come back and the destructive action must succeed', () => {
    const close = stepOf(REMOVE_STEPS, 'outside-close');
    const line = { ok: true, resultType: 'ComObject', result: '(COM object)' };
    const actions = close.actions.map((op, action) => ({ action, op, ok: true }));
    assert.equal(matchCall(close, line, actions), true);
    assert.equal(matchCall(close, line, actions.map(a => a.action === 3 ? { ...a, ok: false } : a)), false);
    assert.equal(matchCall(close, line, actions.map(a => a.action === 6 ? { ...a, ok: false } : a)), true, 'reading Name after Close may fail');
    assert.equal(matchCall(close, { ...line, resultType: 'String', result: 'error:notfound:0:' }, actions), false);
    assert.equal(matchCall(close, { ...line, ok: false }, actions), false);
    const info = stepOf(REMOVE_STEPS, 'obj-apcproject');
    assert.equal(matchCall(info, line, []), true, 'info steps only need the object');
    assert.equal(matchCall(stepOf(REMOVE_STEPS, 'obj-array'), { ok: true, resultType: 'Object[]' }, []), true);
    assert.deepEqual(callArgs(close, 'L').slice(0, 3), ['--label', 'outside-close', 'vbacall-obj']);
    assert.deepEqual(callArgs(close, 'L', { noRetry: true }).slice(0, 3), ['--label', 'outside-close', 'vbacall-obj'], 'vbacall-obj never retries');
});

const APC_INFO = 'apc=IApc;apcVersion=7.01.1105;vbaVersion=7.01.1105;running=True;exec=C5D_Bridge;count=3;err=0'
    + '|name=toQI;display=toQI;tn=IApcProject;mode=0;vbType=100;marker=False;bModule=False;exec=False;storage=The VBA Project<C:\\p\\toQI.catvba'
    + '|name=C5D_Bridge;tn=IApcProject;mode=1;vbType=100;marker=True;bModule=False;exec=True;storage=apc<C:\\L\\C5D_Bridge-9d6a56f7.catvba'
    + '|name=C5D_Bridge;tn=IApcProject;mode=0;vbType=100;vbFile=!438;marker=False;bModule=True;exec=False;storage=apc<C:\\L\\C5D_Bridge-probe-b.catvba';

test('parseInfo / describeInfo read the info strings of C5D_RmApcInfo and C5D_RmVbeInfo', () => {
    const info = parseInfo(APC_INFO);
    assert.equal(info.count, 3);
    assert.deepEqual(info.header, { apc: 'IApc', apcVersion: '7.01.1105', vbaVersion: '7.01.1105', running: 'True', exec: 'C5D_Bridge', count: '3', err: '0' });
    assert.deepEqual(info.projects.map(p => [p.name, p.marker, p.storage]), [['toQI', 'False', 'The VBA Project<C:\\p\\toQI.catvba'],
        ['C5D_Bridge', 'True', 'apc<C:\\L\\C5D_Bridge-9d6a56f7.catvba'], ['C5D_Bridge', 'False', 'apc<C:\\L\\C5D_Bridge-probe-b.catvba']]);
    assert.equal(info.projects[2].vbFile, '!438');
    assert.equal(parseInfo(null), null);
    assert.equal(parseInfo('apc=Nothing;err=429:x').count, null);
    assert.equal(describeInfo({ step: 'i', ok: true, result: APC_INFO }), '3 projects, C5D_Bridge x2 [C5D_Bridge-9d6a56f7.catvba (probe,executing,type=100,mode=1); '
        + 'C5D_Bridge-probe-b.catvba (B-module,type=100,mode=0)]');
    assert.equal(describeInfo({ step: 'i', skipped: 'requires x' }), 'not run (requires x)');
    assert.equal(describeInfo({ step: 'i', ok: false, hresult: '0x80004005' }), 'failed (0x80004005)');
    assert.match(describeInfo({ step: 'i', ok: true, result: 'vbe=VBE;version=7.01;count=1;err=0|name=C5D_Bridge;type=100;mode=2;file=;marker=True' }),
        /^1 projects, C5D_Bridge x1 \[\? \(probe,type=100,mode=2\)\]$/);
});

test('remove verdicts: nothing removed / removed by outside-close / copy B', () => {
    const ok = (step, result, extra = {}) => ({ step, ok: true, match: true, result, loadedAfter: false, ...extra });
    const notAlive = step => ({ step, ok: true, match: false, result: 'alive', loadedAfter: false });
    const objOk = (step, typeInfo, actions) => ok(step, '(COM object)', { resultType: 'ComObject', typeInfo, isDispatch: true, actions });
    const act = (action, op, result, ok2 = true, extra = {}) => ({ action, op, kind: op.split(':')[0], ok: ok2, result, hresult: ok2 ? '0x00000000' : '0x800A01B6', message: ok2 ? null : 'Method Close failed', ...extra });
    const base = [ok('version', 'c5d-bridge:1'), ok('install', 'ok'), ok('alive', 'alive'), ok('info-apc', APC_INFO),
        objOk('obj-apcproject', 'IApcProject', [act(0, 'type:$', 'IApcProject'), act(1, 'get:$.Name', 'C5D_Bridge'), act(2, 'get:$.Mode', '0')]),
        ok('obj-array', '[(COM object), (COM object)]', { resultType: 'Object[]' }),
        { step: 'inside-remove', ok: true, match: false, result: 'error:remove:440:Method Remove of object _VBProjects failed', loadedAfter: false },
        notAlive('inside-remove-check'), { step: 'inside-remove-reload', skipped: 'requires inside-remove-check' }];
    const failedClose = [...base, { ...objOk('outside-close', 'IApcProject', [act(0, 'get:$.Name', 'C5D_Bridge'), act(1, 'get:$.Mode', '0'),
        act(2, 'get:$.Apc.Projects.Count', '3'), act(3, 'call:$.Close(false)', null, false), act(5, 'get:$.Apc.Projects.Count', '3')]), match: false },
    notAlive('outside-close-check')];
    const verdict = judgeRemove(false, failedClose, REMOVE_STEPS);
    assert.match(verdict, /info-apc: 3 projects, C5D_Bridge x2/);
    assert.match(verdict, /obj-apcproject: object came back \(IApcProject\); type \$=IApcProject, \$\.Name=C5D_Bridge, \$\.Mode=0/);
    assert.match(verdict, /obj-array: no object \(resultType Object\[\]/);
    assert.match(verdict, /inside-remove: returned "error:remove:440:[^"]*" -> still loaded/);
    assert.match(verdict, /outside-close: call:\$\.Close\(false\) FAILED 0x800A01B6 Method Close failed \[\$\.Name=C5D_Bridge, \$\.Apc\.Projects\.Count=3, \$\.Apc\.Projects\.Count=3\] -> still loaded/);
    assert.match(verdict, /held-remove: not in this run/);
    assert.match(verdict, /C5D_Bridge was not removed by any attempt/);
    assert.match(verdict, /vbscript\.dll not loaded$/);
    const removed = [...base, objOk('outside-close', 'IApcProject', [act(3, 'call:$.Close(false)', null), act(5, 'get:$.Apc.Projects.Count', '2'),
        act(6, 'get:$.Name', null, false)]),
    { step: 'outside-close-check', ok: false, match: true, result: null, hresult: '0x80004005', loadedAfter: false }, ok('outside-close-reload', 'c5d-bridge:1'),
    { step: 'outside-design-close', skipped: 'not needed: outside-close-check matched' },
    ok('b-load', 'c5d-bridge:1'), ok('b-info-apc', APC_INFO), ok('b-install', 'ok'), { step: 'b-module-in-b', ok: false, match: false, hresult: '0x80004005' },
    ok('b-module-in-a', 'b-alive'), ok('b-remove', 'removed;err=0;via=key'), ok('b-remove-info', APC_INFO.split('|').slice(0, 3).join('|').replace('count=3', 'count=2')),
    { step: 'b-close', skipped: 'not needed: b-remove matched' }, ok('b-reload', 'c5d-bridge:1')];
    const v2 = judgeRemove(false, removed, REMOVE_STEPS);
    assert.match(v2, /outside-close: call:\$\.Close\(false\) ok \[\$\.Apc\.Projects\.Count=2, \$\.Name=ERR 0x800A01B6\] -> REMOVED \(C5D_RmAlive not found\), bridge re-loads afterwards/);
    assert.match(v2, /outside-design-close: not run \(not needed: outside-close-check matched\)/);
    assert.match(v2, /C5D_Bridge REMOVED by outside-close/);
    assert.match(v2, /b-load: copy B auto-loaded and answers \| b-info-apc: 3 projects, C5D_Bridge x2/);
    assert.match(v2, /C5D_Install through copy B wrote into: A \(first C5D_Bridge\)/);
    assert.match(v2, /b-remove \(A removes B\): returned "removed;err=0;via=key"; after: 2 projects, C5D_Bridge x1/);
    assert.match(v2, /b-close \(A closes B\): not run \(not needed: b-remove matched\)/);
    assert.match(v2, /b-reload: copy B answers again/);
    const crashed = [...base.slice(0, 6), { step: 'inside-remove', crashed: true, ok: false, match: false, timedOut: false, hresult: '0x800706BA', loadedAfter: null }];
    assert.match(judgeRemove(false, crashed, REMOVE_STEPS), /CATIA crashed after step inside-remove/);
    const timer = [ok('version', 'c5d-bridge:1'), ok('timer-remove', 'scheduled;mode=remove;timer=12345;err=0'),
        { step: 'timer-remove-result', ok: false, match: false, hresult: '0x80004005' }, { step: 'timer-remove-check', ok: false, match: true, hresult: '0x80004005' },
        ok('read-timer-log', null, { log: '2026-11-21 10:00:00 schedule:remove\r\n2026-11-21 10:00:01 fired:remove\r\n2026-11-21 10:00:01 after:remove:ok\r\n' })];
    const v3 = judgeRemove(false, timer, REMOVE_RISKY_STEPS);
    assert.match(v3, /timer-remove: returned "scheduled;mode=remove;timer=12345;err=0" -> REMOVED/);
    assert.match(v3, /timer log: "schedule:remove \/ fired:remove \/ after:remove:ok"/);
});

test('remove summary prints the object actions and copy B', () => {
    const report = { mode: 'remove-experiment', reportPath: 'r.json', notes: ['n'], snapshots: {}, library: 'A.catvba',
        bridge: { action: 'reused', source: { path: 's', exists: true, sha256: 'a'.repeat(64), readOnly: false }, target: { path: 'A.catvba', exists: true, sha256: 'a'.repeat(64), readOnly: true },
            lock: { inUse: false }, targetAtEnd: null },
        bridgeB: { action: 'copied', before: { path: 'B.catvba' }, target: { path: 'B.catvba', exists: true, sha256: 'a'.repeat(64), readOnly: true }, lock: { inUse: false } },
        runs: [{ step: 'obj-vbproject', call: 'C5D_ProbeRemove.C5D_RmSelfVbProject', durationMs: 50, catiaAlive: true, match: true,
            result: { ok: true, result: '(COM object)', resultType: 'ComObject', typeInfo: '_VBProject', isDispatch: true, hresult: '0x00000000', elapsedMs: 9 },
            actions: [{ action: 0, op: 'get:$.Name', ok: true, result: 'C5D_Bridge', resultType: 'String', elapsedMs: 1 },
                { action: 1, op: 'call:$.Collection.Remove($)', ok: false, hresult: '0x800A01B8', message: 'Method failed\r\nmore', exception: 'COMException', elapsedMs: 2 }] },
        { step: 'timer-remove-wait', call: '(local) wait', local: true, kind: 'wait', ok: true, durationMs: 3000, catiaAlive: true },
        { step: 'read-timer-log', call: '(local) read-timer-log', local: true, kind: 'read-timer-log', ok: true, log: 'l1\r\nl2\r\n', durationMs: 0 }],
        summary: { verdict: 'v' } };
    const text = formatUnloadSummary(report);
    assert.match(text, /bridge copy B: B\.catvba prepared: copied, read-only before CATIA: yes, not open elsewhere/);
    assert.match(text, /returned ComObject _VBProject isDispatch=yes/);
    assert.match(text, /\[0\] get:\$\.Name -> ok "C5D_Bridge" \(String\) 1ms/);
    assert.match(text, /\[1\] call:\$\.Collection\.Remove\(\$\) -> FAILED 0x800A01B8 Method failed \[COMException\] 2ms/);
    assert.match(text, /timer-remove-wait: waited 3000ms/);
    assert.match(text, /read-timer-log: 2 lines.*\n {4}\| l1\n {4}\| l2/);
});

test('eval-probe.exe: vbacall-obj arguments of every remove step survive the command line; actions evaluate on a local COM object', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const run = args => execFileSync(exe, args, { windowsHide: true, encoding: 'utf8', timeout: 20000 });
    const library = 'C:\\Users\\u\\AppData\\Local\\Cat5Dev\\bridge\\C5D_Bridge-9d6a56f7.catvba';
    for (const s of materializeSteps([...REMOVE_STEPS, ...REMOVE_RISKY_STEPS], { timerLog: 'C:\\Users\\u\\remove-timer.log' }).filter(s => !s.local)) {
        const line = JSON.parse(run(['--dry-run', ...callArgs(s, library, { noRetry: true })]).trim());
        assert.deepEqual([line.step, line.library, line.libraryType, line.program, line.function, line.retryByRef], [s.name, library, 2, s.module, s.func, false]);
        assert.deepEqual(line.arguments, s.args);
        if (s.actions) assert.deepEqual([line.actions, line.parsed], [s.actions, s.actions], `${s.name}: actions are parsed back to the same text`);
        else assert.equal(line.actions, undefined);
    }
    const bad = args => assert.throws(() => execFileSync(exe, args, { windowsHide: true, stdio: 'pipe' }), error => error.status === 2, args.join(' '));
    bad(['--dry-run', 'vbacall-obj', 'L', 'M']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'get:$.']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'call:$.Name']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'let:$.Close(1)=2']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'drop:$']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', "get:$.Item('x)"]);
    const parsed = JSON.parse(run(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', "call:$.X( 'a''b' ,-3,missing , true,$0.Y )", 'sleep:5']).trim());
    assert.deepEqual(parsed.parsed, ["call:$.X('a''b', -3, missing, true, $0.Y)", 'sleep:5']);
    if (process.platform !== 'win32') return;
    // the evaluator itself, on Scripting.Dictionary (no CATIA): property get, method call with a COM object argument, let, type, errors
    let out;
    try {
        out = run(['--obj-selftest', 'Scripting.Dictionary', 'type:$', "call:$.Add('k', 7)", 'get:$.Count', "get:$.Item('k')", "call:$.Add('self', $)",
            "type:$.Item('self')", 'let:$.CompareMode=1', 'get:$.Nope', 'get:$1.Count']);
    } catch (error) { out = String(error.stdout); }
    const lines = out.trim().split(/\r?\n/).map(l => JSON.parse(l));
    assert.deepEqual([lines[0].resultType, lines[0].isDispatch, lines[0].typeInfo], ['ComObject', true, 'IDictionary']);
    const a = lines.slice(1);
    assert.deepEqual(a.map(x => [x.op, x.ok]), [['type:$', true], ["call:$.Add('k', 7)", true], ['get:$.Count', true], ["get:$.Item('k')", true],
        ["call:$.Add('self', $)", true], ["type:$.Item('self')", true], ['let:$.CompareMode=1', false], ['get:$.Nope', false], ['get:$1.Count', false]]);
    assert.deepEqual([a[0].result, a[2].result, a[3].result, a[5].result], ['IDictionary', '1', '7', 'IDictionary']);
    assert.equal(a[7].hresult, '0x80020006', 'DISP_E_UNKNOWNNAME is recorded, not thrown');
    assert.match(a[8].message, /needs an array result/);
});
