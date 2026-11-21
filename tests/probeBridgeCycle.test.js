// --cycle-experiment of scripts/probe-bridge.js (no CATIA): arguments, C5D_Cy source, steps, stop rules, local file handling, verdict.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, dispatch, callArgs, matchCall, runSequence, sxLiteral, BRIDGE_MODULE, CY_MODULE, CY_SOURCE, CY_READ_PROJECT, CY_READ_MODULE,
    CY_METHODS, CY_DEFAULT_ROUNDS, CY_MAX_ROUNDS, CY_START, CY_FINAL, CY_CLOSE_ACTION, parseCycleOptions, cycleRounds, cyCloseActions, cycleListOk,
    buildCycleRound, cycleStepNames, cycleFatal, prepareCycleFile, scanBridgeFolder, cleanupBridgeLeftovers, cycleRoundFacts, cycleStopReason,
    judgeCycle, runCycle, formatCycleSummary, parseSxList } = require('../scripts/probe-bridge');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');

const DIR = 'C:\\Users\\u\\AppData\\Local\\Cat5Dev\\bridge';
const A = `${DIR}\\C5D_Bridge-9d6a56f7.catvba`;
const PLAN = cycleRounds(A, '20261121150000', CY_METHODS, 2);
const USERS = [['Develop', 'C:\\Work\\Dev\\Develop.catvba', 56], ['toQI', 'C:\\Work\\共同研究\\vba\\toQI.catvba', 6]];

/** C5D_CyList text: users first, then the given C5D_Bridge paths. */
function listText(bridges, { users = USERS, saved = 'True', err = '0' } = {}) {
    const all = [...users.map(([name, file, comps]) => ({ name, display: file, file, comps })), ...bridges.map(display => ({ name: 'C5D_Bridge', display, file: display, comps: 3 }))];
    const apc = all.map((p, i) => `|src=apc;i=${i + 1};name=${p.name};display=${p.display};mode=1;file=${p.file};comps=${p.comps};c5d=;exec=False`).join('');
    const vbe = all.map((p, i) => `|src=vbe;i=${i + 1};name=${p.name};type=100;mode=0;file=${p.file};saved=${saved};comps=${p.comps};c5d=`).join('');
    return `list=IApc;running=True;exec=;count=${all.length};vbeCount=${all.length};err=${err}${apc}${vbe}`;
}
const who = (display, name = 'C5D_Bridge') => `who=ok;running=True;name=${name};display=${display};mode=1;vbName=${name};file=${display};comps=3;c5d=C5D_Bridge_Push,C5D_Cy`;
const READ = 'read=ok;project=toQI;hits=1;file=C:\\Work\\toQI.catvba;saved=True;module=C5D_BridgeTest;lines=42;first=Option Explicit;savedAfter=True';
const ok = (step, result, extra = {}) => ({ step, ok: true, match: true, result, loadedAfter: null, alive: null, ...extra });
const unknown = step => ({ step, ok: false, match: true, result: null, hresult: '0x80004005', alive: null,
    message: 'ExecuteScript(x, C5D_Bridge_Push, C5D_Version)\nﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push "\n' });
const ADF09 = 'ExecuteScript(x, C5D_Cy, C5D_CyList)\n実行ｴﾗｰ \n(MS VBA のｽｸﾘﾌﾟﾄ ｴﾝｼﾞﾝが次のｴﾗｰを報告しました: \nAPC.ApcHost.7 : Unknown VBA error; error code = 0x800ADF09.)\n';
function closeActions(display, { ms = 590 } = {}) {
    return cyCloseActions(display).map((op, action) => {
        if (action === CY_CLOSE_ACTION) return { action, op, ok: true, skipped: false, result: null, elapsedMs: ms };
        if (action === CY_CLOSE_ACTION + 1) return { action, op, ok: false, skipped: false, hresult: '0x80040501', message: 'Object is disconnected.' };
        return { action, op, ok: true, skipped: false, result: ['C5D_Bridge', display, display][action] ?? 'C5D_Bridge' };
    });
}
/** Entries of a complete, clean round. */
function roundEntries(r, { users = USERS, ms = {} } = {}) {
    const t = r.tag, p = r.path;
    return [ok(`${t}-prepare`, undefined, { ms: ms.prepare ?? 12, prepare: { ok: true, links: r.method === 'hardlink' ? 2 : 1 } }),
        ok(`${t}-load`, 'c5d-bridge:1', { elapsedMs: ms.load ?? 470 }), ok(`${t}-install`, 'ok', { elapsedMs: 30 }), ok(`${t}-who`, who(p), { elapsedMs: 4 }),
        ok(`${t}-list`, listText([p], { users }), { elapsedMs: 9 }), ok(`${t}-read`, READ, { elapsedMs: 5 }),
        ok(`${t}-close`, '(COM object)', { elapsedMs: 10, resultType: 'ComObject', actions: closeActions(p) }), unknown(`${t}-recall`),
        ok(`${t}-lock`, undefined, { locks: { [t]: { inUse: true, code: 'EBUSY' } } })];
}

test('bridge probe parses --cycle-experiment with --cycle-method / --cycle-rounds and dispatches to runCycle', () => {
    assert.deepEqual(parseArgs(['--cycle-experiment']), { mode: 'cycle-experiment', steps: [], cycle: { methods: ['copy', 'hardlink'], rounds: CY_DEFAULT_ROUNDS } });
    assert.deepEqual(parseArgs(['--cycle-rounds', '3', '--cycle-experiment', '--cycle-method', 'hardlink']).cycle, { methods: ['hardlink'], rounds: 3 });
    assert.deepEqual(parseArgs(['--cycle-experiment', '--cycle-method', 'copy']).cycle.methods, ['copy']);
    assert.deepEqual(parseArgs(['--cycle-experiment', '--cycle-method', 'both']).cycle.methods, ['copy', 'hardlink'], 'both: copy first');
    assert.equal(parseArgs(['--cycle-experiment', '--cycle-rounds', String(CY_MAX_ROUNDS)]).cycle.rounds, CY_MAX_ROUNDS);
    assert.throws(() => parseArgs(['--cycle-experiment', '--cycle-method', 'move']), /--cycle-method must be hardlink, copy or both/);
    for (const bad of ['0', '21', 'x', '2.5', '-1']) assert.throws(() => parseArgs(['--cycle-experiment', '--cycle-rounds', bad]), /--cycle-rounds must be 1\.\.20/, bad);
    assert.throws(() => parseArgs(['--cycle-experiment', '--cycle-rounds']), /needs a value/);
    assert.throws(() => parseArgs(['--cycle-experiment', '--cycle-rounds', '2', '--cycle-rounds', '3']), /given twice/);
    assert.throws(() => parseArgs(['--cycle-method', 'copy']), /only with --cycle-experiment/);
    assert.throws(() => parseArgs(['--self-close-experiment', '--cycle-rounds', '2']), /only with --cycle-experiment/);
    assert.throws(() => parseArgs(['--cycle-experiment', '--self-close-experiment']), /cannot be combined/);
    assert.throws(() => parseArgs(['--cycle-experiment', 'version']), /takes no steps/);
    assert.deepEqual(parseCycleOptions(), { methods: CY_METHODS, rounds: 5 });
    const d = dispatch({ mode: 'cycle-experiment' });
    assert.deepEqual([d.run, d.format, d.prefix], [runCycle, formatCycleSummary, 'probe-bridge-cycle']);
    // the other modes are unchanged
    assert.deepEqual(parseArgs(['--self-close-experiment']), { mode: 'self-close-experiment', steps: [] });
});

test('C5D_Cy source: safe ASCII VBA, reads only (no write / close / remove / save), Self / Who / List / Read', () => {
    const source = CY_SOURCE;
    assert.ok(/^[\x20-\x7e\r\n]*$/.test(source), 'ASCII only');
    assert.ok(!/[^\r]\n/.test(source), 'CRLF');
    assert.ok(source.startsWith('Option Explicit\r\n'));
    assert.ok(source.endsWith('End Function\r\n'));
    assert.ok(!/^Attribute /m.test(source));
    assert.ok(source.split('\r\n').every(line => line.length <= 1023), 'VBA line length');
    assert.ok(!/\bMsgBox\b|\bInputBox\b|^\s*Stop\s*$|^\s*End\s*$|Declare|\.Close\b|\.Remove\b|\.Save(As)?\b|\.Open\b|\.End\b|Projects\.Add/im.test(source));
    assert.ok(!/AddFromString|AddFromFile|DeleteLines|InsertLines|ReplaceLine|\.Import\b|\.Export\b|\.Add\(|\.Name = |\.Saved = |Kill |SendKeys|Shell\(/i.test(source), 'writes nothing');
    assert.ok(!source.includes(BRIDGE_MODULE), 'never names C5D_Bridge_Push');
    for (const line of source.split('\r\n').filter(l => /^\s*(Else)?If .* Then/.test(l))) {
        assert.ok(!/[A-Za-z_)]\.[A-Za-z]/.test(/If (.*) Then/.exec(line)[1]), `property read in an If condition: ${line.trim()}`);
    }
    for (const body of source.split(/\r\n(?=(?:Public|Private) (?:Function|Sub) )/).slice(1)) {
        if (/\.\w/.test(body.split('\r\n').slice(1).join('\n'))) assert.match(body, /\r\n {4}On Error Resume Next\r\n/, body.split('\r\n')[0]);
    }
    const publics = [...source.matchAll(/\r\nPublic Function (\w+)\(/g)].map(m => m[1]);
    assert.deepEqual(publics, ['C5D_CySelf', 'C5D_CyWho', 'C5D_CyList', 'C5D_CyRead']);
    assert.ok(!publics.some(name => ['C5D_Version', 'C5D_Install'].includes(name)));
    const privates = [...source.matchAll(/\r\nPrivate Function (\w+)\(/g)].map(m => m[1]);
    assert.ok(privates.every(name => name.startsWith('Cy')), privates.join(','));
    assert.match(source, /Public Function C5D_CySelf\(\)[\s\S]*?Set ex = apc\.ExecutingProject\r\n[\s\S]*?Set C5D_CySelf = ex\r\nEnd Function/);
    assert.match(source, /C5D_CyWho = "who=ok" & s\r\n/);
    // the read: the first VBProject with exactly that name, CountOfLines and Lines(1, 1) only, Saved before and after
    assert.match(source, /Public Function C5D_CyRead\(ByVal projName As Variant, ByVal modName As Variant\) As Variant\r\n/);
    assert.ok(source.includes('        If StrComp(nm, pn, vbBinaryCompare) = 0 Then\r\n'));
    assert.ok(source.includes('            If hits = 1 Then Set hit = vbp\r\n'));
    assert.ok(source.includes('    Set comp = comps.Item(mn)\r\n'));
    assert.ok(source.includes('    cnt = cm.CountOfLines\r\n'));
    assert.ok(source.includes('        v = cm.Lines(1, 1)\r\n'));
    assert.match(source, /v = hit\.Saved\r\n[^\r]*\r\n {4}s = s & ";saved=" & CyFmt\(v, e\)[\s\S]*v = cm\.Lines\(1, 1\)[\s\S]*v = hit\.Saved\r\n[^\r]*\r\n {4}s = s & ";savedAfter=" & CyFmt\(v, e\)/);
    assert.ok(source.includes('    C5D_CyRead = "read=ok" & s\r\n'));
});

test('cycle plan: new paths next to A (never A itself), copy before hardlink, start s and final z are copies', () => {
    const all = [PLAN.start, ...PLAN.rounds, PLAN.final];
    assert.deepEqual(all.map(r => [r.tag, r.method, path.win32.basename(r.path)]), [
        ['s', 'copy', 'C5D_Bridge-9d6a56f7-c20261121150000-s.catvba'],
        ['c1', 'copy', 'C5D_Bridge-9d6a56f7-c20261121150000-1.catvba'], ['c2', 'copy', 'C5D_Bridge-9d6a56f7-c20261121150000-2.catvba'],
        ['h1', 'hardlink', 'C5D_Bridge-9d6a56f7-h20261121150000-1.catvba'], ['h2', 'hardlink', 'C5D_Bridge-9d6a56f7-h20261121150000-2.catvba'],
        ['z', 'copy', 'C5D_Bridge-9d6a56f7-c20261121150000-z.catvba']]);
    assert.equal(new Set(all.map(r => r.path.toLowerCase())).size, all.length);
    assert.ok(all.every(r => path.win32.dirname(r.path) === DIR && r.path.toLowerCase() !== A.toLowerCase()));
    assert.deepEqual([PLAN.start.tag, PLAN.final.tag], [CY_START, CY_FINAL]);
    assert.deepEqual(cycleRounds(A, '1', ['hardlink'], 3).rounds.map(r => r.tag), ['h1', 'h2', 'h3']);
});

test('cycle round steps: one path per round, who before list / read / close, guarded close, read-only read of toQI', () => {
    const r = PLAN.rounds[2];
    const steps = buildCycleRound(r, { apc: 2, vbe: 2 });
    const names = steps.map(s => s.name);
    assert.deepEqual(names, ['prepare', 'load', 'install', 'who', 'list', 'read', 'close', 'recall', 'lock'].map(k => `h1-${k}`));
    steps.forEach((s, i) => {
        for (const dep of [...(s.requires ?? []), ...(s.after ?? [])]) assert.ok(names.indexOf(dep) >= 0 && names.indexOf(dep) < i, `${s.name} -> ${dep}`);
        if (!s.local) assert.equal(s.library, r.path, `${s.name}: library is the round's own path`);
    });
    const step = k => steps.find(s => s.name === `h1-${k}`);
    assert.deepEqual([step('prepare').local, step('prepare').method, step('prepare').path], ['prepare', 'hardlink', r.path]);
    assert.deepEqual([step('load').module, step('load').func, step('load').expected], [BRIDGE_MODULE, 'C5D_Version', 'c5d-bridge:1']);
    assert.deepEqual([step('install').func, step('install').args, step('install').requires], ['C5D_Install', [CY_MODULE, CY_SOURCE], ['h1-load']]);
    for (const k of ['list', 'read', 'close']) assert.deepEqual(step(k).requires, ['h1-who'], k);
    assert.deepEqual(step('who').requires, ['h1-install']);
    assert.deepEqual([step('read').module, step('read').func, step('read').args], [CY_MODULE, 'C5D_CyRead', [CY_READ_PROJECT, CY_READ_MODULE]]);
    assert.deepEqual([CY_READ_PROJECT, CY_READ_MODULE], ['toQI', 'C5D_BridgeTest']);
    const close = step('close');
    assert.deepEqual([close.func, close.args, close.resultType, close.mustSucceed, close.closeAction], ['C5D_CySelf', [], 'ComObject', [CY_CLOSE_ACTION], CY_CLOSE_ACTION]);
    assert.equal(close.actions[CY_CLOSE_ACTION], 'call:$.Close(false)');
    assert.deepEqual(close.actions.slice(3, CY_CLOSE_ACTION), [`guard:$.Name=${sxLiteral('C5D_Bridge')}`, `guard:$.DisplayName=${sxLiteral(r.path)}`]);
    assert.deepEqual([step('recall').expectFailure, step('recall').requires], ['c5d-bridge:1', ['h1-close']]);
    assert.deepEqual([step('lock').local, step('lock').after], ['lock', ['h1-prepare']]);
    // who: ExecutingProject must be this path; list: only this C5D_Bridge and the user projects of the start
    assert.equal(matchCall(step('who'), { ok: true, result: who(r.path) }), true);
    assert.equal(matchCall(step('who'), { ok: true, result: who(PLAN.rounds[0].path) }), false);
    assert.equal(matchCall(step('who'), { ok: true, result: who(r.path, 'toQI') }), false);
    assert.equal(matchCall(step('list'), { ok: true, result: listText([r.path]) }), true);
    assert.equal(matchCall(step('list'), { ok: true, result: listText([r.path, PLAN.rounds[0].path]) }), false, 'a second C5D_Bridge');
    assert.equal(matchCall(step('list'), { ok: true, result: listText([r.path], { users: USERS.slice(0, 1) }) }), false, 'user project count changed');
    assert.equal(matchCall(step('list'), { ok: true, result: listText([r.path], { err: '91' }) }), false);
    assert.equal(cycleListOk(parseSxList(listText([PLAN.start.path], { users: [] })), PLAN.start.path), true, 'the start round only checks C5D_Bridge x1');
    assert.equal(matchCall(step('read'), { ok: true, result: READ }), true);
    assert.equal(matchCall(step('read'), { ok: true, result: 'error:noproject;count=10;err=0' }), false);
    assert.equal(matchCall(step('recall'), unknown('x')), true);
    assert.equal(matchCall(step('recall'), { ok: true, result: 'c5d-bridge:1' }), false, 're-loaded');
    // whole run: unique names, master A never a library
    const names2 = cycleStepNames(CY_METHODS, CY_DEFAULT_ROUNDS);
    assert.equal(new Set(names2).size, names2.length);
    assert.equal(names2.length, (2 + 2 * CY_DEFAULT_ROUNDS) * 9);
    assert.deepEqual([names2[0], names2[9], names2[9 * 6], names2[names2.length - 1]], ['s-prepare', 'c1-prepare', 'h1-prepare', 'z-lock']);
    const every = [PLAN.start, ...PLAN.rounds, PLAN.final].flatMap(x => buildCycleRound(x));
    assert.ok(every.every(s => s.local || s.library.toLowerCase() !== A.toLowerCase()));
});

test('cycleFatal: 0x800ADF09, RPC_E_SERVERFAULT and RPC failures are fatal; "macro unknown" and the disconnected read after Close are not', () => {
    const close = buildCycleRound(PLAN.rounds[0]).find(s => s.name.endsWith('-close'));
    assert.equal(cycleFatal({ step: 'x', ok: false, hresult: '0x80004005', message: ADF09 }), 'Unknown VBA error 0x800ADF09');
    assert.equal(cycleFatal({ step: 'x', ok: false, hresult: '0x80010105', message: 'サーバーによって例外が返されました。 (RPC_E_SERVERFAULT)' }), 'RPC_E_SERVERFAULT 0x80010105');
    assert.equal(cycleFatal({ step: 'x', ok: false, hresult: '0x800706BA', message: 'RPC サーバーを利用できません。' }), 'RPC failure (CATIA gone?)');
    assert.equal(cycleFatal(unknown('x')), null);
    assert.equal(cycleFatal({ step: 'x', skipped: 'requires y' }), null);
    assert.equal(cycleFatal(ok('x', 'ok')), null);
    const actions = closeActions(PLAN.rounds[0].path);
    assert.equal(cycleFatal(ok('c1-close', '(COM object)', { actions }), close), null);
    actions[CY_CLOSE_ACTION + 1] = { ...actions[CY_CLOSE_ACTION + 1], hresult: '0x80010108', message: 'RPC_E_DISCONNECTED' };
    assert.equal(cycleFatal(ok('c1-close', '(COM object)', { actions }), close), null, 'after the Close: expected');
    actions[CY_CLOSE_ACTION] = { ...actions[CY_CLOSE_ACTION], ok: false, hresult: '0x80010105', message: 'RPC_E_SERVERFAULT' };
    assert.equal(cycleFatal(ok('c1-close', '(COM object)', { actions }), close), 'RPC_E_SERVERFAULT 0x80010105', 'the Close itself');
});

test('runSequence options: fatal aborts the rest like a timeout, snapshot false skips the module snapshot', async () => {
    const report = { notes: [], snapshots: {} };
    const steps = ['a', 'b', 'c'].map(name => ({ name, local: 'x' }));
    const local = async step => ({ run: { local: true, ok: true, match: true }, entry: { ok: true, match: true, hresult: step.name === 'b' ? '0x80010105' : null } });
    const { entries, aborted } = await runSequence(report, steps, 'L', [1], local, null, { snapshot: () => false, fatal: cycleFatal });
    assert.deepEqual(entries.map(e => [e.step, e.skipped ?? null, e.fatal ?? null]), [['a', null, null], ['b', null, 'RPC_E_SERVERFAULT 0x80010105'],
        ['c', 'aborted: RPC_E_SERVERFAULT 0x80010105 in b', null]]);
    assert.equal(aborted, 'aborted: RPC_E_SERVERFAULT 0x80010105 in b');
    assert.deepEqual([entries[0].alive, entries[0].loadedAfter, Object.keys(report.snapshots).length], [null, null, 0]);
    assert.match(report.notes[0], /^Step b: RPC_E_SERVERFAULT 0x80010105; CATIA may be unstable/);
    assert.equal(report.runs.find(r => r.step === 'b').fatal, 'RPC_E_SERVERFAULT 0x80010105');
});

test('prepareCycleFile: read-only copy (separate file, same hash) and hard link (same file, shared read-only); never overwrites', t => {
    const dir = fixture(t);
    const a = path.join(dir, 'C5D_Bridge-9d6a56f7.catvba');
    fs.writeFileSync(a, 'bridge');
    fs.chmodSync(a, 0o444);
    const copy = prepareCycleFile('copy', a, path.join(dir, 'c1.catvba'));
    assert.deepEqual([copy.ok, copy.readOnly, copy.sameFile, copy.sha256Same, copy.error], [true, true, false, true, null]);
    assert.equal(typeof copy.ms, 'number');
    const again = prepareCycleFile('copy', a, copy.path);
    assert.equal(again.ok, false);
    assert.match(again.error, /^EEXIST/);
    const link = prepareCycleFile('hardlink', a, path.join(dir, 'h1.catvba'));
    assert.deepEqual([link.ok, link.readOnly, link.sameFile, link.links, link.error], [true, true, true, 2, null]);
    assert.match(prepareCycleFile('hardlink', a, link.path).error, /^EEXIST/);
    fs.chmodSync(a, 0o666);
    const writable = prepareCycleFile('hardlink', a, path.join(dir, 'h2.catvba'));
    assert.equal(writable.ok, false);
    assert.match(writable.error, /read-only not set/);
    assert.throws(() => prepareCycleFile('move', a, path.join(dir, 'x.catvba')), /unknown method/);
    for (const f of fs.readdirSync(dir)) fs.chmodSync(path.join(dir, f), 0o666);
});

test('scanBridgeFolder / cleanupBridgeLeftovers: old probe files deleted (hard link of A too, A read-only again), A / old name / locked kept', t => {
    const dir = fixture(t);
    const file = name => path.join(dir, name);
    const a = file('C5D_Bridge-9d6a56f7.catvba');
    fs.writeFileSync(a, 'AAAA');
    fs.chmodSync(a, 0o444);
    fs.writeFileSync(file('C5D_Bridge.catvba'), 'old');
    for (const name of ['C5D_Bridge-probe-b.catvba', 'C5D_Bridge-probe-20261121140718-c1.catvba', 'C5D_Bridge-9d6a56f7-c20261121150000-1.catvba', 'c5d_bridge-locked.CATVBA']) {
        fs.writeFileSync(file(name), 'xx');
        fs.chmodSync(file(name), 0o444);
    }
    fs.linkSync(a, file('C5D_Bridge-9d6a56f7-hl-20261121140718.catvba'));
    fs.writeFileSync(file('remove-timer.log'), 'log');
    const locked = file('c5d_bridge-locked.CATVBA');
    const busy = Object.assign(new Error('EBUSY: resource busy or locked'), { code: 'EBUSY' });
    const open = f => { if (f === locked) throw busy; return fs.openSync(f, 'r'); };
    const scan = scanBridgeFolder(dir, a, { lock: f => (f === locked ? { inUse: true, code: 'EBUSY' } : { inUse: false, code: null }) });
    assert.equal(scan.error, null);
    const role = name => scan.files.find(f => f.name === name);
    assert.equal(scan.files.length, 7, 'remove-timer.log is not a bridge');
    assert.deepEqual([role('C5D_Bridge-9d6a56f7.catvba').role, role('C5D_Bridge.catvba').role, role('c5d_bridge-locked.CATVBA').role], ['master', 'old-name', 'leftover']);
    assert.deepEqual([role('C5D_Bridge-9d6a56f7-hl-20261121140718.catvba').hardlinkOfA, role('C5D_Bridge-probe-b.catvba').hardlinkOfA], [true, false]);
    assert.equal(role('c5d_bridge-locked.CATVBA').lock.inUse, true);
    const out = cleanupBridgeLeftovers(scan, a, { open, platform: 'win32' });
    assert.deepEqual([out.deleted, out.items.length, out.hardlinks, out.kept, out.failed], [4, 5, 1, ['c5d_bridge-locked.CATVBA'], []]);
    assert.equal(out.deletedBytes, 2 + 2 + 2 + 4);
    assert.deepEqual([out.aReadOnly, out.aLinks], [true, 1]);
    assert.deepEqual(fs.readdirSync(dir).sort(), ['C5D_Bridge-9d6a56f7.catvba', 'C5D_Bridge.catvba', 'c5d_bridge-locked.CATVBA', 'remove-timer.log'].sort());
    assert.equal(fs.readFileSync(a, 'utf8'), 'AAAA');
    assert.deepEqual(scanBridgeFolder(path.join(dir, 'missing'), a).files, []);
    for (const f of fs.readdirSync(dir)) fs.chmodSync(path.join(dir, f), 0o666);
});

test('cycle round facts and stop rules: clean round, who mismatch, list mismatch, failed close, re-load, fatal, start checks', () => {
    const r = PLAN.rounds[0];
    const startUsers = cycleRoundFacts(PLAN.start, roundEntries(PLAN.start)).users;
    const clean = cycleRoundFacts(r, roundEntries(r), startUsers);
    assert.deepEqual([clean.ok, clean.ran, clean.loadMs, clean.installMs, clean.close.ok, clean.close.ms, clean.closeCallMs, clean.read, clean.lock.inUse, clean.whoExact],
        [true, true, 470, 30, true, 590, 10, true, true, true]);
    assert.deepEqual(clean.userDiff, []);
    assert.equal(cycleStopReason(clean), null);
    const replace = (name, entry) => roundEntries(r).map(e => (e.step === name ? entry : e));
    // who: another project ran the code
    const whoBad = cycleRoundFacts(r, replace('c1-who', ok('c1-who', who(PLAN.rounds[1].path), { match: false })).map(e => (['c1-list', 'c1-read', 'c1-close', 'c1-recall'].includes(e.step)
        ? { step: e.step, skipped: 'requires c1-who' } : e)), startUsers);
    assert.deepEqual(cycleStopReason(whoBad).kind, 'round');
    assert.match(cycleStopReason(whoBad).reason, /self NO: the code did not run in this path's project, so C5D_Install may have written into another C5D_Bridge$/);
    assert.equal(cycleStopReason(whoBad, { start: true }).kind, 'restart');
    // list: a second C5D_Bridge; the close of itself still ran
    const listBad = cycleRoundFacts(r, replace('c1-list', ok('c1-list', listText([r.path, 'C:\\x.catvba']), { match: false })), startUsers);
    assert.match(cycleStopReason(listBad).reason, /^list: apc 4 \/ vbe 4 projects, C5D_Bridge x2/);
    // close failed / same path answered again
    const closeBad = cycleRoundFacts(r, replace('c1-close', ok('c1-close', 'error:noexec:91', { match: false, resultType: 'String', actions: [] })), startUsers);
    assert.match(cycleStopReason(closeBad).reason, /^close: NOT closed \(no object \("error:noexec:91"\)\) \(this C5D_Bridge stays loaded/);
    const reloaded = cycleRoundFacts(r, replace('c1-recall', ok('c1-recall', 'c5d-bridge:1', { match: false })), startUsers);
    assert.match(cycleStopReason(reloaded).reason, /^same path after close: .*\(re-loaded/);
    // read is recorded, not a stop
    const readBad = cycleRoundFacts(r, replace('c1-read', ok('c1-read', 'error:noproject;count=4;err=0', { match: false })), startUsers);
    assert.deepEqual([readBad.ok, readBad.read, cycleStopReason(readBad)], [true, false, null]);
    // fatal
    const fatal = cycleRoundFacts(r, replace('c1-install', { step: 'c1-install', ok: false, match: false, hresult: '0x80004005', message: ADF09, fatal: 'Unknown VBA error 0x800ADF09' }), startUsers);
    assert.deepEqual(cycleStopReason(fatal, { aborted: 'aborted: Unknown VBA error 0x800ADF09 in c1-install' }), { kind: 'fatal', reason: 'Unknown VBA error 0x800ADF09' });
    assert.deepEqual(cycleStopReason(clean, { aborted: 'aborted after c1-load timed out' }), { kind: 'fatal', reason: 'aborted after c1-load timed out' });
    const prepBad = cycleRoundFacts(r, [ok('c1-prepare', undefined, { ok: false, match: false, message: 'EEXIST: file already exists' })]);
    assert.match(cycleStopReason(prepBad).reason, /^prepare failed \(EEXIST/);
    // start: another C5D_Bridge, no user project (VBA editor never opened)
    const s = PLAN.start;
    const two = cycleRoundFacts(s, roundEntries(s).map(e => (e.step === 's-list' ? ok('s-list', listText([s.path, 'C:\\old\\C5D_Bridge.catvba']), { match: false }) : e)));
    assert.deepEqual(cycleStopReason(two, { start: true }).kind, 'restart');
    assert.match(cycleStopReason(two, { start: true }).reason, /C5D_Bridge x2 loaded at the start: CATIA was not restarted\. CATIA を再起動してください/);
    const none = cycleRoundFacts(s, roundEntries(s, { users: [] }));
    assert.deepEqual(cycleStopReason(none, { start: true }).kind, 'vbe');
    assert.equal(cycleStopReason(cycleRoundFacts(s, roundEntries(s)), { start: true }), null);
});

test('cycle verdict: precheck, cleanup, master, per round, per method, final, user projects, stop, A at end', () => {
    const startFacts = cycleRoundFacts(PLAN.start, roundEntries(PLAN.start));
    const users = startFacts.users;
    const [c1, c2, h1, h2] = PLAN.rounds;
    const entries = [...roundEntries(PLAN.start), ...roundEntries(c1, { ms: { load: 400 } }), ...roundEntries(c2, { ms: { load: 500 } }),
        ...roundEntries(h1).map(e => (e.step === 'h1-install' ? { step: 'h1-install', ok: false, match: false, hresult: '0x80010105', message: 'RPC_E_SERVERFAULT', fatal: 'RPC_E_SERVERFAULT 0x80010105' }
            : e.step === 'h1-prepare' || e.step === 'h1-load' ? e : { step: e.step, skipped: 'aborted: RPC_E_SERVERFAULT 0x80010105 in h1-install' }))];
    const facts = [c1, c2, h1].map(r => cycleRoundFacts(r, entries, users));
    facts.push({ ...cycleRoundFacts(h2, []), ran: false, notRun: 'stopped after round h1' });
    const final = { ...cycleRoundFacts(PLAN.final, []), ran: false, notRun: 'stopped after round h1' };
    const stop = { ...cycleStopReason(facts[2], { aborted: 'aborted: RPC_E_SERVERFAULT 0x80010105 in h1-install' }), after: 'h1', notRun: ['h2', 'z'] };
    const preCheck = { dir: DIR, files: [{ name: 'C5D_Bridge-9d6a56f7.catvba', role: 'master', lock: { inUse: false } }, { name: 'C5D_Bridge-probe-b.catvba', role: 'leftover', lock: { inUse: false } }] };
    const preCleanup = { items: [{ key: 'C5D_Bridge-probe-b.catvba', action: 'deleted' }, { key: 'C5D_Bridge-9d6a56f7-hl-1.catvba', action: 'deleted' }],
        deleted: 2, deletedBytes: 2 * 5529088, hardlinks: 1, kept: [], failed: [], aReadOnly: true, aLinks: 1 };
    const { lines, data } = judgeCycle(false, entries, { a: A, methods: CY_METHODS, rounds: 2, preCheck, preCleanup,
        master: { path: A, action: 'reused', readOnly: true, sameAsSource: true, links: 1 }, start: startFacts, facts, final, stop,
        cleanup: [{ key: 's', path: PLAN.start.path, action: 'kept (locked)' }], aEnd: { unchanged: true, readOnly: true, links: 2, lock: { inUse: true } } });
    const line = prefix => lines.find(l => l.startsWith(prefix));
    assert.match(line('precheck'), /^precheck: 2 C5D_Bridge\*\.catvba in .*, none in use$/);
    assert.match(line('cleanup before run'), /^cleanup before run: deleted 2\/2 old files \(10\.5 MB; hard links of A 1\); A read-only afterwards yes, A links 1$/);
    assert.match(line('master A'), /^master A C5D_Bridge-9d6a56f7\.catvba: reused, read-only yes, = source yes, links 1 \(never loaded through its own path\)$/);
    assert.match(line('round s (copy, start list)'), /copy ok 12ms, load ok 470ms, install ok 30ms, ExecutingProject = .* self yes \(exact\), list apc 3 \/ vbe 3 projects, C5D_Bridge x1 .*\(ok: yes\), read toQI\.C5D_BridgeTest lines=42 first="Option Explicit" saved True->True, closed 590ms \(ExecuteScript 10ms\), same path after close fails as expected .*lock after close LOCKED/);
    assert.match(line('user projects at start'), /^user projects at start: 2 \(Develop\[56\] saved=True, toQI\[6\] saved=True\)$/);
    assert.match(line('round c1 (copy)'), /load ok 400ms.*user projects = start$/);
    assert.match(line('round h1 (hardlink)'), /hard link ok 12ms \(links 2\), load ok 470ms, install failed \(0x80010105 RPC_E_SERVERFAULT\).*; FATAL: RPC_E_SERVERFAULT 0x80010105$/);
    assert.match(line('round h2'), /not run \(stopped after round h1\)/);
    assert.match(line('method copy'), /^method copy: 2\/2 rounds ok \(2 run\); average of the ok rounds: prepare 12ms, load 450ms, install 30ms, who 4ms, list 9ms, read 5ms, close 590ms \(ExecuteScript 10ms\); read ok 2\/2; locked after close 2\/2 probed; user projects = start in 2\/2$/);
    assert.match(line('method hardlink'), /^method hardlink: 0\/2 rounds ok \(1 run\);.*load -,/);
    assert.match(line('round z (copy, final list)'), /not run \(stopped after round h1\)/);
    assert.match(line('user projects start vs end'), /^user projects start vs end \(list of round c2\): identical \(2 projects/);
    assert.match(line('stop'), /^stop \(fatal: CATIA may be unstable, restart it before anything else\) after round h1: RPC_E_SERVERFAULT 0x80010105; not run: h2, z$/);
    assert.match(line('cleanup after run'), /^cleanup after run: deleted 0\/1, kept C5D_Bridge-9d6a56f7-c20261121150000-s\.catvba \(kept \(locked\)\)$/);
    assert.match(line('A at end'), /^A at end: unchanged yes, read-only yes, links 2, LOCKED$/);
    assert.equal(lines[lines.length - 1], 'vbscript.dll not loaded');
    assert.deepEqual([data.rounds.length, data.stop.after, data.final.ran], [4, 'h1', false]);
    // a user project that changed (Saved False after the read) is reported
    const dirty = cycleRoundFacts(c1, roundEntries(c1).map(e => (e.step === 'c1-list' ? ok('c1-list', listText([c1.path], { saved: 'False' })) : e)), users);
    assert.equal(dirty.userDiff.length, 2);
    const v2 = judgeCycle(false, [], { a: A, methods: ['copy'], rounds: 1, start: startFacts, facts: [dirty], final: null, stop: null });
    assert.match(v2.lines.find(l => l.startsWith('user projects start vs end')), /DIFFER: vbe #1/);
    assert.equal(v2.lines.find(l => l.startsWith('stop')), 'stop: none (every planned round and the final list ran)');
    // precheck stop: nothing else
    const v3 = judgeCycle(false, [], { a: A, methods: CY_METHODS, rounds: 5,
        preCheck: { dir: DIR, files: [{ name: 'C5D_Bridge-9d6a56f7.catvba', role: 'master', lock: { inUse: true } }] },
        stop: { kind: 'restart', after: null, reason: 'C5D_Bridge-9d6a56f7.catvba in use: ... CATIA を再起動してください (restart CATIA; nothing was done)' } });
    assert.deepEqual(v3.lines, ['precheck: 1 C5D_Bridge*.catvba in C:\\Users\\u\\AppData\\Local\\Cat5Dev\\bridge, IN USE: C5D_Bridge-9d6a56f7.catvba',
        'stop (restart) before the first CATIA call: C5D_Bridge-9d6a56f7.catvba in use: ... CATIA を再起動してください (restart CATIA; nothing was done)']);
});

test('cycle summary prints the plan, precheck, cleanup, copy / link / lock steps and the verdict', () => {
    const report = { mode: 'cycle-experiment', reportPath: 'r.json', notes: ['n1'], snapshots: {}, library: '(per step, see each run)', libraryPerStep: true,
        cycle: { methods: ['copy', 'hardlink'], rounds: 2, readProject: 'toQI', readModule: 'C5D_BridgeTest' },
        preCheck: { dir: DIR, files: [{ name: 'C5D_Bridge-probe-b.catvba', role: 'leftover', hardlinkOfA: false, lock: { inUse: false } },
            { name: 'C5D_Bridge-9d6a56f7-hl-1.catvba', role: 'leftover', hardlinkOfA: true, lock: { inUse: false } }] },
        preCleanup: { items: [{ key: 'C5D_Bridge-probe-b.catvba', action: 'deleted' }] },
        runs: [{ step: 'c1-prepare', call: '(local) prepare', local: true, kind: 'copy', ok: true, readOnly: true, sha256Same: true, sameFile: false, durationMs: 15, catiaAlive: null },
            { step: 'h1-prepare', call: '(local) prepare', local: true, kind: 'link', ok: true, sameFile: true, links: 2, aReadOnly: true, durationMs: 1, catiaAlive: null },
            { step: 'h1-lock', call: '(local) lock', local: true, kind: 'locks', ok: true, locks: { h1: { inUse: true } }, durationMs: 1, catiaAlive: true }],
        summary: { verdictLines: ['v1', 'v2'] } };
    const text = formatCycleSummary(report);
    assert.match(text, /^mode: cycle-experiment\nreport: r\.json\ncycle: methods copy -> hardlink, 2 rounds each, read toQI\.C5D_BridgeTest\n/);
    assert.match(text, /precheck leftover C5D_Bridge-9d6a56f7-hl-1\.catvba \(hard link of A\): free/);
    assert.match(text, /cleanup before run C5D_Bridge-probe-b\.catvba: deleted/);
    assert.match(text, /c1-prepare: copy ok readOnly=yes sha256=A yes separate file=yes 15ms {2}catia=\?/);
    assert.match(text, /h1-prepare: hard link ok sameFile=yes links=2 A readOnly=yes 1ms/);
    assert.match(text, /h1-lock: h1 LOCKED {2}catia=alive/);
    assert.match(text, /\nverdict:\n {2}v1\n {2}v2\nnote: n1$/);
});

test('eval-probe.exe --dry-run: every cycle call survives the command line (paths, the C5D_Cy source, guards)', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const run = args => execFileSync(exe, args, { windowsHide: true, encoding: 'utf8', timeout: 20000 }).trim().split(/\r?\n/).map(l => JSON.parse(l));
    const jp = cycleRounds('C:\\作業\\it\'s\\C5D_Bridge-9d6a56f7.catvba', '20261121150000', ['hardlink'], 1);
    for (const s of [...buildCycleRound(PLAN.start), ...buildCycleRound(PLAN.rounds[3], { apc: 9, vbe: 9 }), ...buildCycleRound(jp.rounds[0])].filter(s => !s.local)) {
        const [line] = run(['--dry-run', ...callArgs(s, s.library, { noRetry: true })]);
        assert.deepEqual([line.step, line.library, line.program, line.function, line.parametersType], [s.name, s.library, s.module, s.func, 'System.Object[]']);
        assert.deepEqual(line.arguments, s.args, s.name);
        if (s.actions) assert.deepEqual([line.actions, line.parsed], [s.actions, s.actions]);
        else assert.equal(line.retryByRef, false, s.name);
    }
});
