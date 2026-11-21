// --self-close-experiment of scripts/probe-bridge.js, guard actions and the iParameters array type of EvalProbe.cs (no CATIA).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, dispatch, callArgs, matchCall, BRIDGE_MODULE, B_MODULE, SX_MODULE, SXB_MODULE, SX_SOURCE, SXB_SOURCE, SX_ROUNDS, SX_FINAL,
    SX_CLOSE_ACTION, SX_MAX_TARGETS, sxCloseActions, sxLiteral, parseSxList, parseWho, whoIsSelf, sxUserProjects, userDiff, buildSelfCloseStart,
    buildSelfCloseTargets, buildSelfCloseB, buildSelfCloseRounds, selfCloseRounds, selfCloseTargets, selfCloseStepNames, reloadLeftovers,
    sxCloseFacts, judgeSelfClose, runSelfClose, formatSelfCloseSummary, buildReloadSteps, pathVariants, RL_ROUNDS, RL_FINAL } = require('../scripts/probe-bridge');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');

const DIR = 'C:\\Users\\u\\AppData\\Local\\Cat5Dev\\bridge';
const A = `${DIR}\\C5D_Bridge-9d6a56f7.catvba`;
const B = `${DIR}\\C5D_Bridge-probe-b.catvba`;
const C1 = `${DIR}\\C5D_Bridge-probe-20261121140718-c1.catvba`;
const DOT = `${DIR}\\.\\C5D_Bridge-9d6a56f7.catvba`;
const DOTDOT = `${DIR}\\..\\bridge\\C5D_Bridge-9d6a56f7.catvba`;
const LOWER = A.toLowerCase();
const JP = 'C:\\作業\\it\'s\\C5D_Bridge-日本.catvba';
const USERS = [['Develop', 'C:\\Work\\Dev\\Develop.catvba', 3], ['toQI', 'C:\\Work\\共同研究\\vba\\toQI.catvba', 5]];
const ROUNDS = selfCloseRounds(A, '20261121150000');

// C5D_SxList text: users first, then the given C5D_Bridge projects [{ display, file }]
function listText(bridges, { users = USERS, exec = '' } = {}) {
    const all = [...users.map(([name, file, comps]) => ({ name, display: file, file, comps })), ...bridges.map(b => ({ name: 'C5D_Bridge', file: b.display, comps: 2, ...b }))];
    const apc = all.map((p, i) => `|src=apc;i=${i + 1};name=${p.name};display=${p.display};mode=1;file=${p.file};comps=${p.comps};c5d=;exec=False`).join('');
    const vbe = all.map((p, i) => `|src=vbe;i=${i + 1};name=${p.name};type=100;mode=0;file=${p.file};saved=True;comps=${p.comps};c5d=`).join('');
    return `list=IApc;running=True;exec=${exec};count=${all.length};vbeCount=${all.length};err=0${apc}${vbe}`;
}
const who = (display, file = display, name = 'C5D_Bridge') => `who=ok;running=True;name=${name};display=${display};mode=1;vbName=${name};file=${file};comps=2;c5d=C5D_Bridge_Push,C5D_Sx`;
const ok = (step, result, extra = {}) => ({ step, ok: true, match: true, result, loadedAfter: false, alive: true, ...extra });
const unknown = step => ({ step, ok: false, match: true, result: null, hresult: '0x80004005', loadedAfter: false, alive: true,
    message: 'ExecuteScript(x, C5D_Bridge_Push, C5D_Version)\nﾏｸﾛ ﾌﾟﾛｸﾞﾗﾑが不明です "C5D_Bridge_Push "\n' });
/** Actions of a self-close step as eval-probe.exe reports them (guardFail: index of the guard that fails). */
function closeActions(display, file, { guardFail = null, ms = 450 } = {}) {
    return sxCloseActions(display, file).map((op, action) => {
        const skipped = guardFail !== null && action > guardFail;
        const kind = op.split(':')[0];
        if (skipped) return { action, op, kind, ok: false, skipped: true, message: `not run: guard ${guardFail} failed` };
        if (action === guardFail) return { action, op, kind, ok: false, skipped: false, message: 'guard failed: got \'x\', expected \'y\' (ignoring case)' };
        if (action === SX_CLOSE_ACTION) return { action, op, kind, ok: true, skipped: false, result: null, elapsedMs: ms };
        if (action === SX_CLOSE_ACTION + 1) return { action, op, kind, ok: false, skipped: false, hresult: '0x80040501', message: 'disconnected' };
        return { action, op, kind, ok: true, skipped: false, result: ['C5D_Bridge', display, file][action] ?? 'C5D_Bridge' };
    });
}
const closed = (step, display, file = display, opts = {}) => ok(step, '(COM object)', { resultType: 'ComObject', actions: closeActions(display, file, opts),
    match: opts.guardFail === undefined || opts.guardFail === null });

test('bridge probe parses --self-close-experiment and dispatches to runSelfClose', () => {
    assert.deepEqual(parseArgs(['--self-close-experiment']), { mode: 'self-close-experiment', steps: [] });
    assert.throws(() => parseArgs(['--self-close-experiment', 'version']), /takes no steps/);
    assert.throws(() => parseArgs(['--self-close-experiment', '--reload-experiment']), /cannot be combined/);
    const d = dispatch({ mode: 'self-close-experiment' });
    assert.deepEqual([d.run, d.format, d.prefix], [runSelfClose, formatSelfCloseSummary, 'probe-bridge-selfclose']);
});

test('C5D_Sx / C5D_SxB sources: safe ASCII VBA, never closes / removes / saves, writes only C5D_Sx into projects named C5D_Bridge', () => {
    for (const source of [SX_SOURCE, SXB_SOURCE]) {
        assert.ok(/^[\x20-\x7e\r\n]*$/.test(source), 'ASCII only');
        assert.ok(!/[^\r]\n/.test(source), 'CRLF');
        assert.ok(source.startsWith('Option Explicit\r\n'));
        assert.ok(!/^Attribute /m.test(source));
        assert.ok(source.split('\r\n').every(line => line.length <= 1023), 'VBA line length');
        assert.ok(!/\bMsgBox\b|\bInputBox\b|^\s*Stop\s*$|^\s*End\s*$|Declare|\.Close\b|\.Remove\b|\.Save(As)?\b|\.Open\b|\.End\b|Projects\.Add/im.test(source));
        assert.ok(!source.includes(BRIDGE_MODULE), 'never names C5D_Bridge_Push');
        for (const line of source.split('\r\n').filter(l => /^\s*(Else)?If .* Then/.test(l))) {
            assert.ok(!/[A-Za-z_)]\.[A-Za-z]/.test(/If (.*) Then/.exec(line)[1]), `property read in an If condition: ${line.trim()}`);
        }
        for (const body of source.split(/\r\n(?=(?:Public|Private) (?:Function|Sub) )/).slice(1)) {
            if (/\.\w/.test(body.split('\r\n').slice(1).join('\n'))) assert.match(body, /\r\n {4}On Error Resume Next\r\n/, body.split('\r\n')[0]);
        }
    }
    const publics = source => [...source.matchAll(/\r\nPublic Function (\w+)\(/g)].map(m => m[1]);
    assert.deepEqual(publics(SX_SOURCE), ['C5D_SxSelf', 'C5D_SxWho', 'C5D_SxList']);
    assert.deepEqual(publics(SXB_SOURCE), ['C5D_SxPut', 'C5D_SxBList']);
    // B holds both modules plus C5D_Bridge_Push / C5D_ProbeB / C5D_Rl_*: no public name may repeat inside one project
    const all = [...publics(SX_SOURCE), ...publics(SXB_SOURCE), 'C5D_Version', 'C5D_Install', 'C5D_BAlive'];
    assert.equal(new Set(all).size, all.length);
    assert.ok(!all.some(name => name.startsWith('C5D_Rl')));
    // self: apc.ExecutingProject, taken while the VBA runs
    assert.match(SX_SOURCE, /Public Function C5D_SxSelf\(\)[\s\S]*?Set ex = apc\.ExecutingProject\r\n[\s\S]*?Set C5D_SxSelf = ex\r\nEnd Function/);
    assert.match(SX_SOURCE, /C5D_SxWho = "who=ok" & s\r\n/);
    // the writer: only projects named C5D_Bridge, DisplayName exact -> ignoring case -> FileName ignoring case, exactly one hit
    assert.ok(SXB_SOURCE.includes('Private Const SX_PROJECT As String = "C5D_Bridge"\r\n'));
    assert.ok(SXB_SOURCE.includes(`Private Const SX_MODULE As String = "${SX_MODULE}"\r\n`));
    assert.ok(SXB_SOURCE.includes('            If nm = SX_PROJECT Then\r\n'));
    assert.match(SXB_SOURCE, /If pass = 1 Then\r\n\s+ok = \(StrComp\(dn, k, vbBinaryCompare\) = 0\)\r\n\s+ElseIf pass = 2 Then\r\n\s+ok = \(StrComp\(dn, k, vbTextCompare\) = 0\)\r\n\s+Else\r\n\s+ok = \(StrComp\(fl, k, vbTextCompare\) = 0\)/);
    assert.match(SXB_SOURCE, /If hits > 1 Then\r\n\s+C5D_SxPut = "error:ambiguous:pass"/);
    assert.deepEqual([...SXB_SOURCE.matchAll(/\.Add\(/g)].length, 1);
    assert.ok(SXB_SOURCE.includes('        Set comp = comps.Add(1)\r\n'));
    assert.ok(SXB_SOURCE.includes('    Set comp = comps.Item(SX_MODULE)\r\n'));
    assert.ok(SXB_SOURCE.includes('        comp.Name = SX_MODULE\r\n'));
    assert.ok(!/\.Name = (?!SX_MODULE)/.test(SXB_SOURCE), 'renames nothing else');
});

test('self-close steps: unique names, dependencies on earlier steps, info never depends, close only after own who, guarded close', () => {
    const targets = [{ tag: 't1', display: C1, file: C1 }, { tag: 't2', display: DOT, file: A }, { tag: 't3', display: JP, file: JP }];
    const steps = [...buildSelfCloseStart(B), ...buildSelfCloseTargets(B, targets), ...buildSelfCloseB(B), ...buildSelfCloseRounds(ROUNDS, 9)];
    const names = steps.map(s => s.name);
    assert.equal(new Set(names).size, names.length);
    assert.deepEqual(names, selfCloseStepNames(3).slice(0, -1), 'selfCloseStepNames matches (without locks-end)');
    steps.forEach((s, i) => {
        for (const dep of [...(s.requires ?? []), ...(s.unless ?? []), ...(s.after ?? [])]) assert.ok(names.indexOf(dep) >= 0 && names.indexOf(dep) < i, `${s.name} -> ${dep}`);
        if (s.info) assert.equal(s.requires, undefined, `${s.name}: an info step is always run`);
        if (!s.local) assert.ok(typeof s.library === 'string' && s.library.length > 0, `${s.name}: library`);
    });
    const step = name => steps.find(s => s.name === name);
    // phase 1 through B
    for (const name of ['sc-b-alive', 'sc-b-install', 'sc-start-list', 'sc-b-put', 'sc-b-who']) assert.equal(step(name).library, B, name);
    assert.deepEqual([step('sc-b-install').module, step('sc-b-install').func, step('sc-b-install').args], [BRIDGE_MODULE, 'C5D_Install', [SXB_MODULE, SXB_SOURCE]]);
    assert.deepEqual([step('sc-b-put').module, step('sc-b-put').func, step('sc-b-put').args], [SXB_MODULE, 'C5D_SxPut', [B, SX_SOURCE]]);
    assert.equal(step('sc-b-alive').module, B_MODULE);
    // phase 2: put and the list through B, who / close / recall through the copy's own path
    for (const t of targets) {
        assert.deepEqual([step(`${t.tag}-put`).library, step(`${t.tag}-put`).args], [B, [t.display, SX_SOURCE]]);
        assert.deepEqual([step(`${t.tag}-gone`).library, step(`${t.tag}-gone`).func], [B, 'C5D_SxBList']);
        for (const k of ['who', 'close', 'recall']) assert.equal(step(`${t.tag}-${k}`).library, t.display, `${t.tag}-${k}`);
        assert.deepEqual(step(`${t.tag}-close`).requires, [`${t.tag}-who`]);
        assert.deepEqual(step(`${t.tag}-close`).actions, sxCloseActions(t.display, t.file));
    }
    // every close: a ComObject, the Close is action SX_CLOSE_ACTION and comes after the three guards
    for (const s of steps.filter(x => x.actions)) {
        assert.equal(s.func, 'C5D_SxSelf');
        assert.deepEqual([s.mustSucceed, s.resultType, s.args], [[SX_CLOSE_ACTION], 'ComObject', []]);
        assert.equal(s.actions[SX_CLOSE_ACTION], 'call:$.Close(false)');
        assert.deepEqual(s.actions.slice(3, 6).map(a => a.split('=')[0]), ['guard:$.Name', 'guard:$.DisplayName', 'guard:$.VBProject.FileName']);
        assert.equal(s.actions[4], `guard:$.DisplayName=${sxLiteral(s.library)}`);
        assert.ok(s.requires.some(r => r.endsWith('-who')), `${s.name} requires a who`);
    }
    assert.equal(sxLiteral(JP), '\'C:\\作業\\it\'\'s\\C5D_Bridge-日本.catvba\'');
    // phase 3: B only when alone
    assert.deepEqual(step('b-close').requires, ['sc-b-who', 'b-list']);
    assert.equal(step('b-list').check(listText([{ display: B }])), true);
    assert.equal(step('b-list').check(listText([{ display: B }, { display: C1 }])), false);
    // phase 4: link (local) -> load -> install -> who -> list -> close; each round after the previous close
    let previous = 'b-close';
    for (const r of ROUNDS) {
        assert.deepEqual([step(`${r.tag}-link`).local, step(`${r.tag}-link`).path, step(`${r.tag}-link`).requires], ['link', r.path, [previous]]);
        for (const k of ['load', 'install', 'who', 'list', 'close', 'recall']) assert.equal(step(`${r.tag}-${k}`).library, r.path);
        assert.deepEqual(step(`${r.tag}-install`).args, [SX_MODULE, SX_SOURCE]);
        assert.deepEqual(step(`${r.tag}-close`).actions, sxCloseActions(r.path, r.path));
        assert.equal(step(`${r.tag}-list`).check(listText([{ display: r.path }], { users: [...USERS, ...Array.from({ length: 7 }, (_, i) => [`u${i}`, `C:\\u${i}.catvba`, 1])] })), true);
        assert.equal(step(`${r.tag}-list`).check(listText([{ display: r.path }])), false, 'user count must be 9');
        previous = `${r.tag}-close`;
    }
    assert.deepEqual(ROUNDS.map(r => [r.tag, path.win32.basename(r.path)]), [...SX_ROUNDS, SX_FINAL].map((tag, i) => [tag, `C5D_Bridge-9d6a56f7-p20261121150000-${i + 1}.catvba`]));
});

test('self-close who / list parsing, targets from the start list, reload leftovers', t => {
    const w = parseWho(who(DOT, A));
    assert.deepEqual([w.who, w.display, w.file, w.name], ['ok', DOT, A, 'C5D_Bridge']);
    assert.deepEqual(parseWho('error:noexec:91;running=False'), { error: 'error:noexec:91;running=False' });
    assert.equal(parseWho(null), null);
    assert.equal(whoIsSelf(who(DOT, A), DOT), true);
    assert.equal(whoIsSelf(who(LOWER), A.toUpperCase()), true, 'ignoring case');
    assert.equal(whoIsSelf(who(DOTDOT, A), DOT), false);
    assert.equal(whoIsSelf(who(DOT, A, 'toQI'), DOT), false);
    assert.equal(whoIsSelf('error:noexec:91', DOT), false);
    const step = buildSelfCloseTargets(B, [{ tag: 't1', display: DOT, file: A }]).find(s => s.name === 't1-who');
    assert.equal(matchCall(step, { ok: true, result: who(DOT, A) }), true);
    assert.equal(matchCall(step, { ok: true, result: who(B) }), false);

    const list = parseSxList(listText([{ display: B }, { display: C1 }, { display: DOT, file: A }, { display: LOWER }]));
    assert.deepEqual([list.count, list.vbeCount, list.err, list.apc.length, list.vbe.length], [6, 6, '0', 6, 6]);
    assert.deepEqual(sxUserProjects(list).vbe, USERS.map(([name, file, comps]) => ({ name, file, comps: String(comps), saved: 'True' })));
    const leftovers = [{ tag: 'c1', path: C1 }, { tag: 'dot', path: DOT }, { tag: 'lower', path: LOWER }];
    assert.deepEqual(selfCloseTargets(list, B.toLowerCase(), leftovers), [
        { tag: 't1', display: C1, file: C1, label: 'c1' }, { tag: 't2', display: DOT, file: A, label: 'dot' }, { tag: 't3', display: LOWER, file: LOWER, label: 'lower' }]);
    assert.deepEqual(selfCloseTargets(null, B), []);
    const many = parseSxList(listText(Array.from({ length: SX_MAX_TARGETS + 2 }, (_, i) => ({ display: `C:\\x${i}.catvba` }))));
    assert.throws(() => selfCloseTargets(many, B), /more than 20/);
    assert.equal(parseSxList('apc=IApc;count=1'), null);

    const start = sxUserProjects(list);
    assert.deepEqual(userDiff(start, sxUserProjects(parseSxList(listText([])))), []);
    const changed = sxUserProjects(parseSxList(listText([], { users: [USERS[0], [USERS[1][0], USERS[1][1], 6]] })));
    assert.equal(userDiff(start, changed).length, 2, 'apc and vbe entry of toQI differ');
    assert.equal(userDiff(null, start), null);

    const dir = fixture(t);
    const reload = { files: { c1: { state: { path: C1 } }, c2: { state: { path: 'C:\\c2' } }, z: { state: { path: 'C:\\z' } } },
        variants: pathVariants(A, { short: { path: 'C:\\S~1.CAT' }, hardlink: { ok: true, path: 'C:\\hl' } }),
        summary: { steps: [{ step: 'c1-load', match: true }, { step: 'c2-load', match: false }, { step: 'z-load', match: true }, { step: 'dot-load', match: true },
            { step: 'upper-load', match: false }] } };
    fs.writeFileSync(path.join(dir, 'probe-bridge-reload-20261121-140718.json'), JSON.stringify(reload));
    fs.writeFileSync(path.join(dir, 'probe-bridge-reload-20261121-120000.json'), '{}');
    const left = reloadLeftovers(dir);
    assert.equal(path.basename(left.report), 'probe-bridge-reload-20261121-140718.json');
    assert.deepEqual(left.list, [{ tag: 'c1', path: C1 }, { tag: 'z', path: 'C:\\z' }, { tag: 'dot', path: DOT }]);
    assert.equal(reloadLeftovers(path.join(dir, 'missing')), null);
    // the reload run's steps still build (its rounds / variants unchanged by this mode)
    assert.ok(buildReloadSteps({ a: A, b: B, z: 'C:\\z', copies: RL_ROUNDS.map(tag => ({ tag, path: `C:\\${tag}` })), variants: pathVariants(A) }).some(s => s.name === `${RL_FINAL}-close`));
});

test('sxCloseFacts: closed, stopped by a guard, no object, not run', () => {
    assert.deepEqual(sxCloseFacts(closed('x', DOT, A)), { ok: true, ms: 450, display: DOT, file: A, guard: null, error: null });
    const stopped = sxCloseFacts(closed('x', DOT, A, { guardFail: 4 }));
    assert.equal(stopped.ok, false);
    assert.match(stopped.error, /^guard failed: guard:\$\.DisplayName=.* -> guard failed: got 'x'/);
    assert.deepEqual(sxCloseFacts(ok('x', 'error:noexec:91', { resultType: 'String', actions: [] })).error, 'no object ("error:noexec:91")');
    assert.deepEqual(sxCloseFacts({ step: 'x', skipped: 'requires x-who' }).ok, null);
});

test('self-close verdict: one line per copy, B and round, self-identification count, user projects compared', () => {
    const targets = [{ tag: 't1', display: C1, file: C1, label: 'c1' }, { tag: 't2', display: DOT, file: A, label: 'dot' }];
    const users9 = [...USERS, ...Array.from({ length: 7 }, (_, i) => [`u${i}`, `C:\\u${i}.catvba`, 1])];
    const L = (bridges) => listText(bridges.map(display => ({ display })), { users: users9 });
    const entries = [
        ok('sc-b-alive', 'b-alive'), ok('sc-b-install', 'ok'), ok('sc-start-list', L([B, C1, DOT])),
        ok('sc-b-put', `ok;pass=1;existed=False;lines=80;display=${B};file=${B}`), ok('sc-b-who', who(B)),
        ok('t1-put', `ok;pass=1;existed=False;lines=80;display=${C1};file=${C1}`), ok('t1-who', who(C1)), closed('t1-close', C1),
        ok('t1-gone', L([B, DOT])), unknown('t1-recall'),
        ok('t2-put', `ok;pass=1;existed=False;lines=80;display=${DOT};file=${A}`), ok('t2-who', who(DOTDOT, A), { match: false }),
        { step: 't2-close', skipped: 'requires t2-who' }, ok('t2-gone', L([B, DOT])), { step: 't2-recall', skipped: 't2-close not run' },
        ok('b-list', L([B, DOT]), { match: false }), { step: 'b-close', skipped: 'requires b-list' }, { step: 'b-recall', skipped: 'b-close not run' },
        ...ROUNDS.flatMap(r => [{ step: `${r.tag}-link`, skipped: 'requires b-close' }]),
        { step: 'locks-end', ok: true, match: true, locks: { a: { inUse: true }, b: { inUse: true } } }];
    const { lines, data } = judgeSelfClose(false, entries, { b: B, a: A, targets, rounds: ROUNDS, userCount: 9,
        leftovers: { report: 'temp\\probe-bridge-reload-20261121-140718.json', list: [{ tag: 'c1', path: C1 }, { tag: 'dot', path: DOT }, { tag: 'z', path: 'C:\\z' }] },
        cleanup: [], aReadOnly: true });
    const line = prefix => lines.find(l => l.startsWith(prefix));
    assert.match(line('start: '), /answers C5D_ProbeB .* C5D_SxB via C5D_Install ok/);
    assert.match(line('start list'), /apc 12 \/ vbe 12 projects, C5D_Bridge x3 \[C5D_Bridge-probe-b\.catvba; C5D_Bridge-probe-20261121140718-c1\.catvba; C5D_Bridge-9d6a56f7\.catvba \(\\\.\\\)\]/);
    assert.match(line('user projects at start'), /^user projects at start: 9 \(Develop\[3\], toQI\[5\], u0\[1\]/);
    assert.match(line('copies other than B'), /2 live vs 3 left by probe-bridge-reload-20261121-140718\.json; not loaded any more: "C:\\\\z"/);
    assert.match(line('B: '), /put ok \(pass 1, existed False\), ExecutingProject = .*-> self yes \(exact\)/);
    assert.match(line('target t1 (c1)'), /put ok .* self yes \(exact\), closed 450ms, C5D_Bridge 3 -> 2, gone yes, same path after close fails as expected \(0x80004005/);
    assert.match(line('target t2 (dot)'), /ExecutingProject = ".*\\\\\.\.\\\\bridge.*" .*-> self NO, close not run \(requires t2-who\), C5D_Bridge 2 -> 2, gone no/);
    assert.match(line('B before its close'), /C5D_Bridge x2 .*\(only B left: no\); B close not run \(requires b-list\)/);
    assert.match(line('round p1 '), /not run \(requires b-close\)/);
    assert.match(line('rounds closed'), /^rounds closed: 0\/4/);
    assert.match(line('self-identification'), /through the own path: 2\/3; guarded Close: 1\/1 closed/);
    assert.match(line('user projects start vs end'), /identical \(9 projects/);
    assert.match(line('locks at end'), /a LOCKED, b LOCKED/);
    assert.match(line('cleanup'), /nothing; A read-only afterwards yes/);
    assert.equal(lines[lines.length - 1], 'vbscript.dll not loaded');
    assert.deepEqual([data.targets[0].gone, data.targets[1].self, data.bClose.ok], [true, false, null]);

    // a full round
    const r = ROUNDS[0];
    const roundEntries = [ok('sc-start-list', L([B])), ok('b-close', '(COM object)'),
        { step: `${r.tag}-link`, ok: true, match: true, ms: 3, path: r.path }, ok(`${r.tag}-load`, 'c5d-bridge:1', { elapsedMs: 380 }),
        ok(`${r.tag}-install`, 'ok', { elapsedMs: 120 }), ok(`${r.tag}-who`, who(r.path)), ok(`${r.tag}-list`, L([r.path])), closed(`${r.tag}-close`, r.path, r.path, { ms: 460 }),
        unknown(`${r.tag}-recall`), { step: `${r.tag}-lock`, ok: true, match: true, locks: { [r.tag]: { inUse: true } } }];
    const full = judgeSelfClose(false, roundEntries, { b: B, a: A, targets: [], rounds: [r], userCount: 9 });
    assert.match(full.lines.find(l => l.startsWith('round p1')), /link ok 3ms, load ok 380ms, install ok 120ms, ExecutingProject = .* self yes \(exact\), list apc 10 \/ vbe 10 projects, C5D_Bridge x1 .*user 9 \+ 1: yes\), closed 460ms, same path after close fails as expected .*lock after close LOCKED/);
    assert.match(full.lines.find(l => l.startsWith('rounds closed')), /1\/1; average load 380ms, install 120ms, close 460ms/);
});

test('self-close summary prints the copies, link steps and the verdict', () => {
    const report = { mode: 'self-close-experiment', reportPath: 'r.json', notes: [], snapshots: {}, library: '(per step, see each run)', libraryPerStep: true,
        targets: [{ tag: 't1', display: DOT, file: A, label: 'dot' }], leftovers: { report: 'x.json', list: [{}] },
        files: { a: { role: 'A', state: { path: A, exists: true, readOnly: true }, lock: { inUse: true } } },
        runs: [{ step: 'p1-link', call: '(local) link', local: true, kind: 'link', ok: true, sameFile: true, links: 2, aReadOnly: true, durationMs: 4, catiaAlive: true }],
        summary: { verdictLines: ['v1'] } };
    const text = formatSelfCloseSummary(report);
    assert.match(text, /^mode: self-close-experiment\nreport: r\.json\ntarget t1 \(dot\): ".*\\\\\.\\\\C5D_Bridge-9d6a56f7\.catvba" file ".*"\nleft by x\.json: 1 copies\nfile a \(A\)/);
    assert.match(text, /p1-link: hard link ok sameFile=yes links=2 A readOnly=yes 4ms/);
    assert.match(text, /\nverdict:\n {2}v1/);
});

test('eval-probe.exe: iParameters is object[] (0x80020005 fix), guard actions parse and stop the later actions', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const run = args => execFileSync(exe, args, { windowsHide: true, encoding: 'utf8', timeout: 20000 }).trim().split(/\r?\n/).map(l => JSON.parse(l));
    // the --reload-experiment close (one --obj-arg) failed because the array was string[]
    const reloadClose = buildReloadSteps({ a: A, b: B, z: 'C:\\z', copies: [{ tag: 'c1', path: C1 }], variants: [] }).find(s => s.name === 'c1-close');
    const [rl] = run(['--dry-run', ...callArgs(reloadClose, B, { noRetry: true })]);
    assert.deepEqual([rl.parametersType, rl.arguments], ['System.Object[]', [C1]]);
    const [plain] = run(['--dry-run', 'vbacall', 'L', 'M', 'F', 'a', 'b']);
    assert.equal(plain.parametersType, 'System.Object[]');
    // every self-close call survives the command line (Japanese path and a quote in a guard)
    const targets = [{ tag: 't1', display: JP, file: JP }, { tag: 't2', display: DOT, file: A }];
    for (const s of [...buildSelfCloseStart(B), ...buildSelfCloseTargets(B, targets), ...buildSelfCloseB(B), ...buildSelfCloseRounds(ROUNDS, 9)].filter(s => !s.local)) {
        const [line] = run(['--dry-run', ...callArgs(s, s.library, { noRetry: true })]);
        assert.deepEqual([line.step, line.library, line.program, line.function, line.retryByRef, line.parametersType], [s.name, s.library, s.module, s.func, false, 'System.Object[]']);
        assert.deepEqual(line.arguments, s.args, s.name);
        if (s.actions) assert.deepEqual([line.actions, line.parsed], [s.actions, s.actions]);
    }
    // guard semantics on a local COM object (no CATIA)
    const selftest = (...actions) => {
        try { return { status: 0, lines: run(['--obj-selftest', 'Scripting.Dictionary', ...actions]) }; } catch (error) {
            return { status: error.status, lines: String(error.stdout).trim().split(/\r?\n/).filter(Boolean).map(l => JSON.parse(l)) };
        }
    };
    const pass = selftest('guard:$.Count=0', "guard:$.CompareMode='0'", 'get:$.Count');
    assert.deepEqual([pass.status, pass.lines.slice(1).map(l => l.ok)], [0, [true, true, true]]);
    const stop = selftest('get:$.Count', 'guard:$.Count=1', 'call:$.RemoveAll()', 'get:$.Count');
    assert.equal(stop.status, 1);
    assert.deepEqual(stop.lines.slice(1).map(l => [l.op, l.ok, l.skipped]), [['get:$.Count', true, false], ['guard:$.Count=1', false, false],
        ['call:$.RemoveAll()', false, true], ['get:$.Count', false, true]]);
    assert.equal(stop.lines[2].message, "guard failed: got '0', expected '1' (ignoring case)");
    assert.equal(stop.lines[3].message, 'not run: guard 1 failed');
    const bad = args => assert.throws(() => execFileSync(exe, args, { windowsHide: true, stdio: 'pipe' }), error => error.status === 2, args.join(' '));
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'guard:$.Name=$.Name']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'guard:$.Name=missing']);
    bad(['--dry-run', 'vbacall-obj', 'L', 'M', 'F', 'guard:$.Name']);
});
