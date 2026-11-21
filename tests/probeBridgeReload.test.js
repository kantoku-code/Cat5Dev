// --reload-experiment of scripts/probe-bridge.js and --obj-arg of EvalProbe.cs vbacall-obj (no CATIA).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, callArgs, matchCall, dispatch, skipReason, BRIDGE_MODULE, B_MODULE, reloadSource, buildReloadSteps, reloadStepNames, pathVariants,
    makeHardlink, cleanupProbeFiles, rlBridges, describeRlInfo, parseInfo, judgeReload, runReload, formatReloadSummary, RL_ROUNDS, RL_FINAL,
    RL_CLOSE_ACTIONS, rlModule, rlFunc } = require('../scripts/probe-bridge');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');

const DIR = 'C:\\Users\\u\\AppData\\Local\\Cat5Dev\\bridge';
const A = `${DIR}\\C5D_Bridge-9d6a56f7.catvba`;
const B = `${DIR}\\C5D_Bridge-probe-b.catvba`;
const copyPath = tag => `${DIR}\\C5D_Bridge-probe-20261121140000-${tag}.catvba`;
const HL = `${DIR}\\C5D_Bridge-9d6a56f7-hl-20261121140000.catvba`;
const SHORT = `${DIR}\\C5D_BR~2.CAT`;
const variants = (extra = {}) => pathVariants(A, { short: { path: SHORT, error: null }, hardlink: { ok: true, path: HL }, ...extra });
const paths = (v = variants()) => ({ a: A, b: B, z: copyPath(RL_FINAL), copies: RL_ROUNDS.map(tag => ({ tag, path: copyPath(tag) })), variants: v });
const stepOf = (steps, name) => steps.find(s => s.name === name);

test('bridge probe parses --reload-experiment and dispatches to runReload', () => {
    const names = reloadStepNames();
    assert.deepEqual(parseArgs(['--reload-experiment']), { mode: 'reload-experiment', steps: names });
    assert.deepEqual(names, buildReloadSteps(paths()).map(s => s.name), 'the step names do not depend on the paths');
    assert.throws(() => parseArgs(['--reload-experiment', 'version']), /takes no steps/);
    assert.throws(() => parseArgs(['--reload-experiment', '--remove-experiment']), /cannot be combined/);
    const d = dispatch({ mode: 'reload-experiment' });
    assert.deepEqual([d.run, d.format, d.prefix], [runReload, formatReloadSummary, 'probe-bridge-reload']);
});

test('pathVariants: other strings for the same file, \\\\?\\ last; unusable ones are disabled with the reason', () => {
    const list = variants();
    assert.deepEqual(list.map(v => [v.tag, v.library, v.disabled]), [
        ['lower', A.toLowerCase(), undefined],
        ['upper', A.toUpperCase(), undefined],
        ['dot', `${DIR}\\.\\C5D_Bridge-9d6a56f7.catvba`, undefined],
        ['dotdot', `${DIR}\\..\\bridge\\C5D_Bridge-9d6a56f7.catvba`, undefined],
        ['slash', 'C:/Users/u/AppData/Local/Cat5Dev/bridge/C5D_Bridge-9d6a56f7.catvba', undefined],
        ['short', SHORT, undefined],
        ['hl', HL, undefined],
        ['ext', `\\\\?\\${A}`, undefined]]);
    const off = pathVariants(A, { short: { path: A, error: null }, hardlink: { ok: false, path: HL, error: 'EBUSY resource busy' } });
    assert.match(off.find(v => v.tag === 'short').disabled, /no 8\.3 short name/);
    assert.equal(off.find(v => v.tag === 'hl').disabled, 'hard link not created (EBUSY resource busy)');
    assert.match(pathVariants(A, { short: { path: null, error: 'boom' } }).find(v => v.tag === 'short').disabled, /no short path \(boom\)/);
    assert.equal(pathVariants(A).find(v => v.tag === 'hl').disabled, 'no hard link made');
    // a variant that equals A or an earlier variant is not tried twice
    const lowerA = 'c:\\x\\c5d_bridge-9d6a56f7.catvba';
    const dup = pathVariants(lowerA, { short: { path: 'c:\\x\\C5D_BR~1.CAT' }, hardlink: { ok: true, path: 'c:\\x\\h.catvba' } });
    assert.equal(dup.find(v => v.tag === 'lower').disabled, 'same string as A itself');
    assert.equal(dup.find(v => v.tag === 'upper').disabled, undefined);
});

test('reload steps: unique names, dependencies on earlier steps, info steps never depend on success, libraries per phase', () => {
    const steps = buildReloadSteps(paths());
    const names = steps.map(s => s.name);
    assert.equal(new Set(names).size, names.length);
    steps.forEach((s, i) => {
        for (const dep of [...(s.requires ?? []), ...(s.unless ?? []), ...(s.after ?? [])]) assert.ok(names.indexOf(dep) >= 0 && names.indexOf(dep) < i, `${s.name} -> ${dep}`);
        if (s.info) assert.deepEqual([s.requires, s.unless], [undefined, undefined], `${s.name}: an info step is always run`);
        if (s.local) assert.deepEqual([s.local, s.requires, s.after], ['locks', undefined, undefined]);
    });
    assert.equal(steps.filter(s => s.info).length, 28);
    const bClose = names.indexOf('b-close');
    steps.forEach((s, i) => {
        if (s.local) return;
        if (s.info || s.actions) {
            // info and Close go through B until B is closed, then through z
            assert.equal(s.library, i <= bClose ? B : copyPath(RL_FINAL), s.name);
            assert.equal(s.module, rlModule(i <= bClose ? 'b' : RL_FINAL), s.name);
        }
    });
    for (const tag of RL_ROUNDS) {
        const file = copyPath(tag);
        assert.deepEqual(['load', 'install', 'alive', 'info', 'close', 'gone', 'recall', 'lock'].map(k => stepOf(steps, `${tag}-${k}`) && k), ['load', 'install', 'alive', 'info', 'close', 'gone', 'recall', 'lock']);
        assert.deepEqual([stepOf(steps, `${tag}-load`).library, stepOf(steps, `${tag}-load`).func], [file, 'C5D_Version']);
        assert.deepEqual(stepOf(steps, `${tag}-install`).args, [rlModule(tag), reloadSource(tag)]);
        for (const k of ['install', 'alive', 'close']) assert.deepEqual(stepOf(steps, `${tag}-${k}`).requires, [`${tag}-load`]);
        assert.deepEqual([stepOf(steps, `${tag}-alive`).library, stepOf(steps, `${tag}-alive`).module, stepOf(steps, `${tag}-alive`).expected], [file, rlModule(tag), `rl-alive:${tag}`]);
        assert.deepEqual([stepOf(steps, `${tag}-close`).args, stepOf(steps, `${tag}-close`).actions], [[file], RL_CLOSE_ACTIONS]);
        assert.deepEqual([stepOf(steps, `${tag}-recall`).library, stepOf(steps, `${tag}-recall`).expectFailure], [file, 'c5d-bridge:1']);
        assert.deepEqual(stepOf(steps, `${tag}-lock`).files, [tag]);
    }
    for (const v of variants()) {
        assert.equal(stepOf(steps, `${v.tag}-load`).library, v.library);
        assert.deepEqual(stepOf(steps, `${v.tag}-close`).args, [v.library]);
        assert.deepEqual(stepOf(steps, `${v.tag}-info`).after, [`${v.tag}-load`]);
        assert.deepEqual(stepOf(steps, `${v.tag}-gone`).after, [`${v.tag}-close`]);
    }
    // the repeat rounds (production idea) run before the exotic path strings
    assert.ok(names.indexOf('c3-lock') < names.indexOf('lower-load'));
    assert.ok(names.indexOf('ext-gone') < bClose && names.indexOf('ext-load') > names.indexOf('hl-gone'));
    assert.deepEqual(stepOf(steps, 'b-close').args, [B]);
    assert.deepEqual([stepOf(steps, 'z-close').args, stepOf(steps, 'z-close').library], [[copyPath(RL_FINAL)], copyPath(RL_FINAL)]);
    assert.deepEqual(stepOf(steps, 'locks-end').files, ['a', 'b', 'c1', 'c2', 'c3', 'z', 'hl']);
    // A itself is only called (a-call), never closed or written to
    assert.deepEqual(steps.filter(s => s.library === A).map(s => s.name), ['a-call']);
    assert.ok(!steps.some(s => (s.args ?? []).includes(A)));
    // closing only through C5D_RlFind (C5D_Bridge projects only), one Close per step
    for (const s of steps.filter(s => s.actions)) {
        assert.equal(s.func, rlFunc('Find', s.module.slice(rlModule('').length)));
        assert.deepEqual([s.actions.filter(a => /Close/.test(a)), s.mustSucceed], [['call:$.Close(false)'], [2]]);
    }
});

test('disabled variants skip their load; dependent steps follow through requires / after', () => {
    const steps = buildReloadSteps(paths(pathVariants(A, { short: { path: A }, hardlink: { ok: false, error: 'EBUSY' } })));
    const load = stepOf(steps, 'short-load');
    assert.match(load.disabled, /no 8\.3 short name/);
    assert.equal(skipReason(load, [], null), load.disabled);
    assert.equal(skipReason(load, [], 'aborted after x timed out'), 'aborted after x timed out');
    const entries = [{ step: 'short-load', skipped: load.disabled }];
    assert.equal(skipReason(stepOf(steps, 'short-install'), entries, null), 'requires short-load');
    assert.equal(skipReason(stepOf(steps, 'short-info'), entries, null), 'short-load not run');
    assert.equal(skipReason(stepOf(steps, 'short-gone'), [...entries, { step: 'short-close', skipped: 'requires short-load' }], null), 'short-close not run');
    assert.equal(stepOf(steps, 'hl-load').disabled, 'hard link not created (EBUSY)');
    assert.equal(stepOf(steps, 'lower-load').disabled, undefined);
});

test('C5D_Rl_<tag> source: safe ASCII VBA, unique public names per tag, never closes / removes / saves', () => {
    assert.throws(() => reloadSource('C1'), /bad reload tag/);
    assert.throws(() => reloadSource('a-b'), /bad reload tag/);
    assert.throws(() => reloadSource('toolongtag'), /bad reload tag/);
    for (const tag of ['b', 'c1', RL_FINAL, 'dotdot', 'short', 'hl', 'ext']) {
        const source = reloadSource(tag);
        assert.ok(/^[\x20-\x7e\r\n]*$/.test(source), 'ASCII only');
        assert.ok(!/[^\r]\n/.test(source), 'CRLF');
        assert.ok(source.startsWith('Option Explicit\r\n'));
        assert.ok(!/^Attribute /m.test(source));
        assert.ok(!/\bMsgBox\b|\bInputBox\b|^\s*Stop\s*$|^\s*End\s*$|Declare|\.Close\b|\.Remove\b|\.Save(As)?\b|\.Add\b|\.Open\b|\.End\b/im.test(source));
        for (const line of source.split('\r\n').filter(l => /^\s*(Else)?If .* Then/.test(l))) {
            assert.ok(!/[A-Za-z_)]\.[A-Za-z]/.test(/If (.*) Then/.exec(line)[1]), `property read in an If condition: ${line.trim()}`);
        }
        for (const body of source.split(/\r\n(?=(?:Public|Private) (?:Function|Sub) )/).slice(1)) {
            if (/\.\w/.test(body.split('\r\n').slice(1).join('\n'))) assert.match(body, /\r\n {4}On Error Resume Next\r\n/, body.split('\r\n')[0]);
        }
        const publics = [...source.matchAll(/\r\nPublic Function (\w+)\(/g)].map(m => m[1]);
        assert.deepEqual(publics, ['Alive', 'Info', 'Vbe', 'Find'].map(kind => rlFunc(kind, tag)));
        assert.ok(publics.every(name => name.length <= 31));
        assert.ok(source.includes(`Private Const RL_MARKER As String = "${rlModule(tag)}"\r\n`));
        assert.ok(source.includes(`    ${rlFunc('Alive', tag)} = "rl-alive:${tag}"\r\n`));
        assert.ok(source.includes(`    Set ${rlFunc('Find', tag)} = ap\r\n`));
    }
    const source = reloadSource('b');
    assert.ok(source.includes('        If c = "\\" Or c = "/" Then Exit For\r\n'), 'RlBase splits at \\ and /');
    assert.ok(source.includes('            If nm = RL_PROJECT Then\r\n'), 'RlFind only returns projects named C5D_Bridge');
    assert.match(source, /ok = \(dn = key\) Or \(fl = key\)\r\n[\s\S]*vbTextCompare\) = 0\) Or[\s\S]*StrComp\(RlBase\(dn\), kb, vbTextCompare\)/);
    assert.match(source, /why = "ambiguous:pass" & CStr\(pass\)/);
    // the steps call exactly these functions
    for (const s of buildReloadSteps(paths()).filter(s => s.module?.startsWith('C5D_Rl_'))) {
        assert.ok(reloadSource(s.module.slice(7)).includes(`\r\nPublic Function ${s.func}(`), `${s.name}: ${s.module}.${s.func}`);
    }
    assert.ok(!reloadSource('c1').includes('C5D_RlInfo_b'), 'one module never calls another tag');
});

test('vbacall-obj passes step args as --obj-arg tokens before the call', () => {
    const close = stepOf(buildReloadSteps(paths()), 'c1-close');
    const args = callArgs(close, B, { noRetry: true });
    assert.deepEqual(args.slice(0, 5), ['--obj-arg', 'b64:' + Buffer.from(copyPath('c1')).toString('base64'), '--label', 'c1-close', 'vbacall-obj']);
    assert.equal(args.length, 5 + 3 + RL_CLOSE_ACTIONS.length);
    const line = { ok: true, resultType: 'ComObject', result: '(COM object)' };
    const acts = RL_CLOSE_ACTIONS.map((op, action) => ({ action, op, ok: true }));
    assert.equal(matchCall(close, line, acts), true);
    assert.equal(matchCall(close, line, acts.map(a => a.action === 2 ? { ...a, ok: false } : a)), false);
    assert.equal(matchCall(close, line, acts.map(a => a.action === 3 ? { ...a, ok: false } : a)), true, 'Name after Close may fail (disconnected)');
    assert.equal(matchCall(close, { ...line, resultType: 'String', result: 'error:find:notfound:count=10' }, acts), false);
});

// --- verdict
const proj = (name, display, mods = [BRIDGE_MODULE], extra = '') => `|name=${name};display=${display};mode=1;dirty=False;vbType=100;vbFile=${display};mods=${mods.join(',')};self=False${extra}`;
const apcInfo = (...bridges) => `apc=IApc;count=${bridges.length + 1};err=0${proj('toQI', 'C:\\w\\toQI.catvba', [])}${bridges.join('')}`;
const bProj = (more = []) => proj('C5D_Bridge', B, [BRIDGE_MODULE, B_MODULE, rlModule('b'), ...more]);
const ok = (step, result, extra = {}) => ({ step, ok: true, match: true, result, loadedAfter: false, alive: true, ...extra });
const fail = (step, extra = {}) => ({ step, ok: false, match: false, result: null, hresult: '0x80004005', loadedAfter: false, alive: true,
    message: `ExecuteScript(x, C5D_Bridge_Push, C5D_Version)\nﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push "\n`, ...extra });
const recalled = step => ({ ...fail(step), match: true });
const closed = (step, display, ms = 470) => ok(step, '(COM object)', { resultType: 'ComObject', actions: [
    { action: 0, op: 'get:$.Name', ok: true, result: 'C5D_Bridge' }, { action: 1, op: 'get:$.DisplayName', ok: true, result: display },
    { action: 2, op: 'call:$.Close(false)', ok: true, result: null, elapsedMs: ms },
    { action: 3, op: 'get:$.Name', ok: false, hresult: '0x80040501', message: 'Object is disconnected.' }] });

test('reload verdict: one line per round / variant / lock state, plus the same facts as data', () => {
    const v = variants();
    const steps = buildReloadSteps(paths(v));
    const c1 = copyPath('c1');
    const entries = [
        ok('b-alive', 'b-alive'), ok('b-install', 'ok'), ok('info-apc', apcInfo(bProj())), ok('info-vbe', 'vbe=VBE;count=2;err=0|name=C5D_Bridge;type=100;mode=0;file=' + B + ';saved=False;mods=C5D_Bridge_Push;self=True'),
        recalled('a-call'), ok('locks-start', null, { locks: { a: { inUse: true, code: 'EBUSY' }, b: { inUse: true, code: 'EBUSY' } } }),
        ok('c1-load', 'c5d-bridge:1', { elapsedMs: 393 }), ok('c1-install', 'ok'), ok('c1-alive', 'rl-alive:c1'),
        ok('c1-info', apcInfo(bProj(), proj('C5D_Bridge', c1, [BRIDGE_MODULE, rlModule('c1')]))), closed('c1-close', c1, 469),
        ok('c1-gone', apcInfo(bProj())), recalled('c1-recall'), ok('c1-lock', null, { locks: { c1: { inUse: true, code: 'EBUSY' } } }),
        // round c2: C5D_Install wrote into B (first C5D_Bridge), the module is not reachable through the c2 path
        ok('c2-load', 'c5d-bridge:1', { elapsedMs: 380 }), ok('c2-install', 'ok'), fail('c2-alive'),
        ok('c2-info', apcInfo(bProj([rlModule('c2')]), proj('C5D_Bridge', copyPath('c2')))), closed('c2-close', copyPath('c2'), 12),
        ok('c2-gone', apcInfo(bProj([rlModule('c2')]))), recalled('c2-recall'), ok('c2-lock', null, { locks: { c2: { inUse: false } } }),
        // lower: same path for CATIA (known) -> not loaded
        fail('lower-load'), { step: 'lower-install', skipped: 'requires lower-load' }, ok('lower-info', apcInfo(bProj())),
        { step: 'lower-close', skipped: 'requires lower-load' }, { step: 'lower-gone', skipped: 'lower-close not run' },
        // short: loaded, shown with the long name
        ok('short-load', 'c5d-bridge:1', { elapsedMs: 410 }), ok('short-install', 'ok'), ok('short-alive', 'rl-alive:short'),
        ok('short-info', apcInfo(bProj(), proj('C5D_Bridge', A, [BRIDGE_MODULE, rlModule('short')]))), closed('short-close', A, 20), ok('short-gone', apcInfo(bProj())),
        closed('b-close', B, 30), recalled('b-recall'),
        ok('z-load', 'c5d-bridge:1', { elapsedMs: 500 }), ok('z-install', 'ok'), ok('z-info', apcInfo(proj('C5D_Bridge', copyPath('z'), [BRIDGE_MODULE, rlModule('z')]))),
        closed('z-close', copyPath('z'), 25), recalled('z-recall'),
        ok('locks-end', null, { locks: { a: { inUse: true }, b: { inUse: false }, c1: { inUse: true }, z: { inUse: false } } })];
    const cleanup = [{ key: 'c1', action: 'kept (locked)' }, { key: 'z', action: 'deleted' }];
    const files = { a: A, b: B, z: copyPath('z'), c1, c2: copyPath('c2'), c3: copyPath('c3') };
    const { lines, reload } = judgeReload(false, entries, steps, { files, copies: paths(v).copies, variants: v, cleanup });
    const line = re => lines.find(l => re.test(l));
    assert.match(lines[0], /^start: B is the B of the remove run \(C5D_ProbeB answers\), b-install ok$/);
    assert.match(line(/^info-apc/), /^info-apc: 2 projects, C5D_Bridge x1 \[C5D_Bridge-probe-b\.catvba \{C5D_ProbeB,C5D_Rl_b\}\]$/);
    assert.match(line(/^A path/), /^A path C5D_Bridge-9d6a56f7\.catvba: fails as expected \(0x80004005 ﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push "\)$/);
    assert.equal(line(/^locks at start/), 'locks at start: a LOCKED, b LOCKED');
    assert.equal(line(/^round c1/), `round c1 ${path.win32.basename(c1)}: load ok 393ms, install yes, alive via own path yes, module went into ${path.win32.basename(c1)}, `
        + `C5D_Bridge 1 -> 2 -> 1, close ok 469ms (found ${path.win32.basename(c1)}), gone yes, same path after close fails as expected (0x80004005 ﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push "), `
        + 'lock right after close LOCKED');
    assert.match(line(/^round c2/), /alive via own path no, module went into C5D_Bridge-probe-b\.catvba, C5D_Bridge 1 -> 2 -> 1, close ok 12ms .*lock right after close free$/);
    assert.match(line(/^round c3/), /^round c3 \S+: load not in this run, install -, .*C5D_Bridge \? -> \? -> \?, close not in this run/);
    assert.match(line(/^variant lower/), /^variant lower \(i\) lower case "c:\\\\users.*": NOT loaded: failed \(0x80004005 ﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push "\)$/);
    assert.match(line(/^variant short/), /^variant short \(iii\) 8\.3 short path "[^"]*C5D_BR~2\.CAT": LOADED 410ms \(new project "C:\\\\Users[^"]*C5D_Bridge-9d6a56f7\.catvba"\), module went into [^,]*9d6a56f7\.catvba, alive yes, close ok 20ms, gone yes$/);
    assert.match(line(/^variant ext/), /^variant ext \(v\) \\\\\?\\ prefix .*: NOT loaded: not in this run$/);
    assert.equal(line(/^b-close/), 'b-close: ok 30ms; B path after close fails as expected (0x80004005 ﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push ")');
    assert.match(line(/^final/), /^final C5D_Bridge-probe-20261121140000-z\.catvba \(0 -> 1 -> 0\): load ok 500ms, install yes, z-info 2 projects, C5D_Bridge x1 \[C5D_Bridge-probe-20261121140000-z\.catvba \{C5D_Rl_z\}\] \(B still listed: no\), close ok 25ms, same path after close fails as expected/);
    assert.equal(line(/^locks at end/), 'locks at end: a LOCKED, b free, c1 LOCKED, z free');
    assert.equal(line(/^cleanup/), 'cleanup: c1 kept (locked), z deleted');
    assert.equal(lines[lines.length - 1], 'vbscript.dll not loaded');
    assert.deepEqual(reload.rounds.map(r => [r.round, r.load.ok, r.load.ms, r.close.ok, r.close.ms, r.gone, r.lockAfterClose?.inUse ?? null]),
        [['c1', true, 393, true, 469, true, true], ['c2', true, 380, true, 12, true, false], ['c3', null, null, null, null, null, null]]);
    assert.deepEqual(reload.variants.filter(x => x.variant === 'short' || x.variant === 'lower').map(x => [x.variant, x.load.ok, x.newProjects, x.gone]),
        [['lower', false, [], null], ['short', true, [A], true]]);
    assert.deepEqual([reload.final.bStillListed, reload.bClose.ok, reload.bClose.ms], [false, true, 30]);
    // crash / timeout lines
    const crashed = judgeReload(false, [...entries.slice(0, 7), { ...fail('c1-install'), crashed: true, alive: false }], steps, { files, copies: paths(v).copies, variants: v });
    assert.ok(crashed.lines.includes('CATIA crashed after step c1-install'));
    const loaded = judgeReload(false, [{ ...ok('b-alive', 'b-alive'), loadedAfter: true }], steps, { files, copies: [], variants: [] });
    assert.equal(loaded.lines[loaded.lines.length - 1], 'vbscript.dll loaded (first seen after step b-alive)');
});

test('rlBridges / describeRlInfo read C5D_RlInfo and C5D_RlVbe texts', () => {
    const info = parseInfo(apcInfo(bProj(), proj('C5D_Bridge', 'C:\\x\\y.catvba')));
    assert.deepEqual(rlBridges(info).map(b => [b.display, b.mods]), [[B, [BRIDGE_MODULE, B_MODULE, rlModule('b')]], ['C:\\x\\y.catvba', [BRIDGE_MODULE]]]);
    assert.equal(describeRlInfo(ok('i', apcInfo())), '1 projects, C5D_Bridge x0');
    assert.equal(describeRlInfo({ step: 'i', skipped: 'x not run' }), 'not run (x not run)');
    assert.equal(describeRlInfo(fail('i')), 'failed (0x80004005 ﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push ")');
    assert.equal(describeRlInfo(ok('i', 'vbe=VBE;count=1;err=0|name=C5D_Bridge;type=100;file=C:\\q\\z.catvba;mods=C5D_Bridge_Push,C5D_Rl_z;self=True')),
        '1 projects, C5D_Bridge x1 [z.catvba {C5D_Rl_z}]');
});

test('makeHardlink / cleanupProbeFiles: free probe files are deleted, locked ones kept, the hard link source stays read-only', t => {
    const dir = fixture(t);
    const a = path.join(dir, 'A.catvba');
    fs.writeFileSync(a, 'x');
    fs.chmodSync(a, 0o444);
    const hl = makeHardlink(a, path.join(dir, 'A-hl.catvba'));
    assert.deepEqual([hl.ok, hl.sameFile, hl.links, hl.error], [true, true, 2, null]);
    assert.equal(fs.statSync(hl.path).mode & 0o200, 0, 'the link shares the read-only attribute');
    const again = makeHardlink(a, hl.path);
    assert.equal(again.ok, false);
    assert.match(again.error, /^EEXIST/);
    const free = path.join(dir, 'free.catvba'), locked = path.join(dir, 'locked.catvba');
    for (const f of [free, locked]) { fs.writeFileSync(f, 'y'); fs.chmodSync(f, 0o444); }
    const open = (file, flags) => { if (file === locked) throw Object.assign(new Error('EBUSY: resource busy or locked'), { code: 'EBUSY' }); return fs.openSync(file, 'r'); };
    const out = cleanupProbeFiles([{ key: 'c1', path: free }, { key: 'c2', path: locked }, { key: 'c3', path: path.join(dir, 'gone.catvba') },
        { key: 'hl', path: hl.path, linkOf: a }], { open, platform: 'win32' });
    assert.deepEqual(out.map(o => [o.key, o.action]), [['c1', 'deleted'], ['c2', 'kept (locked)'], ['c3', 'missing'], ['hl', 'deleted']]);
    assert.deepEqual([fs.existsSync(free), fs.existsSync(locked), fs.existsSync(hl.path), fs.existsSync(a)], [false, true, false, true]);
    assert.equal(out[3].linkOfReadOnly, true);
    assert.equal(fs.statSync(a).mode & 0o200, 0, 'A is read-only again');
    assert.equal(fs.statSync(locked).mode & 0o200, 0, 'a kept file stays read-only');
    const unknown = cleanupProbeFiles([{ key: 'c2', path: locked }], { platform: 'linux' });
    assert.equal(unknown[0].action, 'kept (lock unknown: not Windows)');
    fs.chmodSync(locked, 0o666);
});

test('reload summary prints the files, lock steps, per-step library and the verdict lines', () => {
    const report = { mode: 'reload-experiment', reportPath: 'r.json', notes: ['n1'], snapshots: {}, library: '(per step, see each run)', libraryPerStep: true,
        files: { a: { role: 'A', state: { path: A, exists: true, modifiedAt: 'm', size: 5, readOnly: true }, lock: { inUse: true, code: 'EBUSY' },
            stateAtEnd: { path: A, exists: true, modifiedAt: 'm', size: 5, readOnly: true }, lockAtEnd: { inUse: true } },
        hl: { role: 'hard link of A', hardlink: { ok: false, path: HL, error: 'EBUSY x' }, state: null, lock: null } },
        shortPath: { path: SHORT, error: null }, variants: variants().slice(5, 6), cleanup: [{ key: 'c1', action: 'deleted', path: copyPath('c1') }],
        runs: [{ step: 'locks-start', call: '(local) locks', local: true, kind: 'locks', ok: true, locks: { a: { inUse: true }, b: { inUse: false } }, durationMs: 1, catiaAlive: true },
            { step: 'c1-load', call: 'C5D_Bridge_Push.C5D_Version', library: copyPath('c1'), durationMs: 500, catiaAlive: true, match: true,
                result: { ok: true, result: 'c5d-bridge:1', hresult: '0x00000000', elapsedMs: 390 } }],
        summary: { verdictLines: ['line one', 'line two'], verdict: 'line one\nline two' } };
    const text = formatReloadSummary(report);
    assert.match(text, /^mode: reload-experiment\nreport: r\.json\nfile a \(A\): C:\\Users/);
    assert.match(text, /\n {2}start: modified=m size=5 readOnly=yes, LOCKED\n {2}end: {3}modified=m size=5 readOnly=yes, LOCKED\n/);
    assert.match(text, /file hl \(hard link of A\): .*-hl-20261121140000\.catvba\n {2}hard link NOT created: EBUSY x/);
    assert.match(text, /8\.3 short path of A: .*C5D_BR~2\.CAT/);
    assert.match(text, /variant short \(iii\) 8\.3 short path: ".*C5D_BR~2\.CAT"/);
    assert.match(text, /cleanup c1: deleted /);
    assert.match(text, /locks-start: a LOCKED, b free {2}catia=alive/);
    assert.match(text, /c1-load: ok=yes match=yes result="c5d-bridge:1" elapsed=390ms .*\n.*\n {4}library .*-c1\.catvba/);
    assert.match(text, /\nverdict:\n {2}line one\n {2}line two\nnote: n1/);
});

test('eval-probe.exe: --obj-arg reaches vbacall-obj as ByVal arguments; every reload step survives the command line', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const run = args => JSON.parse(execFileSync(exe, args, { windowsHide: true, encoding: 'utf8', timeout: 20000 }).trim());
    for (const s of buildReloadSteps(paths()).filter(s => !s.local && !s.disabled)) {
        const line = run(['--dry-run', ...callArgs(s, s.library, { noRetry: true })]);
        assert.deepEqual([line.step, line.library, line.libraryType, line.program, line.function, line.retryByRef], [s.name, s.library, 2, s.module, s.func, false]);
        assert.deepEqual(line.arguments, s.args, s.name);
        if (s.actions) assert.deepEqual([line.actions, line.parsed], [s.actions, s.actions]);
    }
    const japanese = run(['--dry-run', '--obj-arg', 'b64:' + Buffer.from('C:\\ü\\日本.catvba').toString('base64'), '--obj-arg', 'plain', 'vbacall-obj', 'L', 'M', 'F', 'get:$.Name']);
    assert.deepEqual([japanese.arguments, japanese.retryByRef], [['C:\\ü\\日本.catvba', 'plain'], false]);
    const bad = args => assert.throws(() => execFileSync(exe, args, { windowsHide: true, stdio: 'pipe' }), error => error.status === 2, args.join(' '));
    bad(['--dry-run', '--obj-arg', 'x', 'vbacall', 'L', 'M', 'F']);
    bad(['--dry-run', '--obj-arg', 'x', 'check']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', '--obj-arg']);
    bad(['--dry-run', '--obj-arg']);
    bad(['--dry-run', '--obj-arg', 'b64:***', 'vbacall-obj', 'L', 'M', 'F']);
});
