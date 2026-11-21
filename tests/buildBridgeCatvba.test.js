// scripts/build-bridge-catvba.js without CATIA: arguments, plan, --dry-run (no file, no exe, no CATIA), offline inspection of a .catvba,
// and runPlan with a fake ExecuteScript (local file steps run for real in a temporary folder).
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const b = require('../scripts/build-bridge-catvba');
const { matchCall } = require('../scripts/probe-bridge');
const { readCatvba, codeOf } = require('../scripts/lib/catvba-read');

const SCRIPT = path.join(__dirname, '..', 'scripts', 'build-bridge-catvba.js');
const SHIPPED = path.join(__dirname, '..', 'resources', 'bridge', 'C5D_Bridge.catvba');
const decode = token => Buffer.from(token.slice(4), 'base64').toString('utf8');

test('arguments: build by default, verify / inspect / export-import, and their option rules', () => {
    assert.deepEqual(b.parseArgs([]), { mode: 'build', dryRun: false, base: null, file: null, dir: null, saveMethod: 'save', install: true, allowOrphans: false });
    assert.equal(b.parseArgs(['--no-install']).install, false);
    assert.equal(b.parseArgs(['--save-method', 'close', '--dry-run']).saveMethod, 'close');
    assert.equal(b.parseArgs(['--base', 'x.catvba']).base, 'x.catvba');
    assert.deepEqual([b.parseArgs(['--verify', 'f.catvba']).install, b.parseArgs(['--verify', 'f.catvba', '--install']).install], [false, true]);
    assert.equal(b.parseArgs(['--verify', 'f.catvba', '--dry-run', '--allow-orphans']).allowOrphans, true);
    assert.equal(b.parseArgs(['--inspect', 'f.catvba']).mode, 'inspect');
    assert.equal(b.parseArgs(['--export-import', 'd']).dir, 'd');
    assert.equal(b.parseArgs(['-h']).mode, 'help');
    for (const [args, why] of [[['--save-method', 'x'], /save or close/], [['--base'], /needs a value/], [['--verify', '--dry-run'], /needs a value/],
        [['--verify', 'a', '--inspect', 'b'], /only one of/], [['--dry-run', '--dry-run'], /given twice/], [['--install'], /installs by default/],
        [['--verify', 'a', '--base', 'b'], /only for the build/], [['--inspect', 'a', '--dry-run'], /takes no other option/],
        [['--verify', 'a', '--install', '--no-install'], /exclude each other/], [['--bogus'], /Unknown option/]]) {
        assert.throws(() => b.parseArgs(args), why, args.join(' '));
    }
});

test('paths: every file of a run is new, next to the bridge folder (never the master folder) and under temp/', () => {
    const p = b.createPaths({ env: { LOCALAPPDATA: 'C:\\L' }, root: 'C:\\repo', stamp: '20261121150000' });
    assert.equal(p.work, 'C:\\L\\Cat5Dev\\bridge-build\\C5D_Bridge-build-20261121150000-w.catvba');
    assert.equal(p.verify, 'C:\\L\\Cat5Dev\\bridge-build\\C5D_Bridge-build-20261121150000-v.catvba');
    assert.equal(p.bridgeDir, 'C:\\L\\Cat5Dev\\bridge');
    assert.equal(p.base, 'C:\\repo\\resources\\bridge\\C5D_Bridge.catvba');
    assert.equal(p.target, p.base);
    assert.equal(p.collected, 'C:\\repo\\temp\\bridge-build-20261121150000\\C5D_Bridge.catvba');
    assert.equal(p.backup, 'C:\\repo\\temp\\C5D_Bridge.catvba.bak-20261121150000');
    assert.equal(p.report, 'C:\\repo\\temp\\build-bridge-catvba-20261121-150000.json');
    assert.throws(() => b.createPaths({ env: {} }), /LOCALAPPDATA/);
});

test('plan: builder first, Builder removed before the save, guarded closes, verify on a new copy, install last', () => {
    const p = b.createPaths({ env: { LOCALAPPDATA: 'C:\\L' }, root: 'C:\\repo', stamp: '20261121150000' });
    const sources = b.readSources();
    const names = steps => steps.map(s => s.name);
    const build = b.planSteps('build', p, sources, { saveMethod: 'save', install: true });
    assert.deepEqual(names(build), ['precheck', 'prepare', 'load', 'install-builder', 'builder-who', 'put-push', 'put-runtime', 'version-v2', 'runtime-version',
        'remove-builder', 'who-v2', 'save', 'close', 'collect', 'inspect', 'verify-prepare', 'v-load', 'v-who', 'v-list', 'v-runtime', 'v-projects',
        'v-install', 'v-echo', 'v-self-guard', 'v-close', 'install']);
    assert.ok(!names(b.planSteps('build', p, sources, { saveMethod: 'close', install: false })).some(n => n === 'save' || n === 'install'));
    assert.deepEqual(names(b.planSteps('verify', p, sources, { install: false })).slice(0, 2), ['inspect', 'verify-prepare']);
    const step = name => build.find(s => s.name === name);
    for (const s of build.filter(x => x.library)) assert.ok(s.library === p.work || s.library === p.verify, s.name);
    assert.deepEqual(step('put-push').args, [b.PUSH_MODULE, sources.push]);
    assert.deepEqual(step('put-runtime').args, [b.RUNTIME_MODULE, sources.runtime]);
    assert.deepEqual(step('install-builder').args, [b.BUILDER_MODULE, b.BUILDER_SOURCE]);
    // every action on a project object is behind the two guards (Name, DisplayName = this copy)
    for (const s of build.filter(x => x.actions)) {
        assert.deepEqual(s.actions.slice(0, 2), [`guard:$.Name='C5D_Bridge'`, `guard:$.DisplayName='${s.library}'`], s.name);
    }
    assert.equal(step('remove-builder').actions[2], "call:$.VBProject.VBComponents.Remove($.VBProject.VBComponents.Item('C5D_Builder'))");
    assert.equal(step('save').actions[2], 'call:$.Save(true)');
    assert.equal(step('v-close').actions[2], 'call:$.Close(false)');
    assert.equal(step('v-echo').expected, `build-probe:${b.PROBE_ARGUMENT}`);
    assert.equal(step('close').cleanup, true);
    assert.equal(step('v-close').cleanup, true);
});

test('checks of the who / list texts identify exactly this copy', () => {
    const w = 'C:\\L\\Cat5Dev\\bridge-build\\C5D_Bridge-build-1-w.catvba';
    assert.ok(b.builderWhoOk(`ok\nC5D_Bridge\t${w}\t${w}\t1\t10`, w));
    assert.ok(b.builderWhoOk(`ok\nC5D_Bridge\t${w.toLowerCase()}\t${w}\t1\t10`, w), 'case-insensitive path');
    assert.ok(!b.builderWhoOk(`ok\nC5D_Bridge\t${w}\t${w}\t2\t11`, w), 'a second C5D_Bridge');
    assert.ok(!b.builderWhoOk(`ok\nC5D_Bridge\tC:\\other.catvba\tC:\\other.catvba\t1\t10`, w), 'another project');
    assert.ok(b.whoOk(`ok\nC5D_Bridge\t${w}\t${w}\nC5D_Bridge_Runtime\t1\nC5D_Bridge_Push\t1`, w));
    assert.ok(!b.whoOk(`ok\nC5D_Bridge\t${w}\t${w}\nC5D_Bridge_Push\t1\nC5D_Bridge_Runtime\t1\nC5D_Builder\t1`, w), 'Builder still there');
    assert.ok(!b.whoOk('error:who.exec:424:No executing project', w));
    assert.ok(b.listOk(`ok\ntoQI\tC:\\toQI.catvba\t0\nC5D_Bridge\t${w}\t1`, w));
    assert.ok(!b.listOk(`ok\nC5D_Bridge\tC:\\old.catvba\t0\nC5D_Bridge\t${w}\t1`, w), 'leftover bridge');
    assert.ok(!b.listOk(`ok\nC5D_Bridge\t${w}\t0`, w), 'not the executing one');
});

test('--dry-run prints the plan without CATIA, without building eval-probe.exe and without writing a file', t => {
    const local = fixture(t);
    const stdout = execFileSync(process.execPath, [SCRIPT, '--dry-run'], { env: { ...process.env, LOCALAPPDATA: local }, encoding: 'utf8', windowsHide: true });
    const plan = JSON.parse(stdout);
    assert.equal(plan.dryRun, true);
    assert.deepEqual(plan.versions, { push: 'c5d-bridge:2', runtime: b.readSources().runtimeVersion });
    assert.deepEqual(fs.readdirSync(local), [], 'nothing created under %LOCALAPPDATA%');
    assert.ok(!fs.existsSync(plan.paths.collectDir), 'no collect folder');
    const call = plan.steps.find(s => s.step === 'put-push');
    const tokens = call.tokens;
    assert.deepEqual(tokens.slice(0, 4), ['--label', 'put-push', 'vbacall', tokens[3]]);
    assert.deepEqual(tokens.slice(3).map(decode), [plan.paths.work, b.BUILDER_MODULE, 'C5D_BuildPut', b.PUSH_MODULE, b.readSources().push]);
    const save = plan.steps.find(s => s.step === 'save');
    assert.equal(save.tokens[2], 'vbacall-obj');
    assert.deepEqual(save.tokens.slice(3).map(decode).slice(0, 3), [plan.paths.work, b.PUSH_MODULE, 'C5D_Self']);
    const echo = plan.steps.find(s => s.step === 'v-echo');
    assert.equal(decode(echo.tokens.at(-1)), b.PROBE_ARGUMENT, 'Unicode argument survives as b64');
    const verify = JSON.parse(execFileSync(process.execPath, [SCRIPT, '--verify', path.join(local, 'none.catvba'), '--dry-run'],
        { env: { ...process.env, LOCALAPPDATA: local }, encoding: 'utf8', windowsHide: true }));
    assert.equal(verify.mode, 'verify');
    assert.ok(!verify.steps.some(s => s.step === 'install'));
    assert.deepEqual(fs.readdirSync(local), []);
});

test('offline inspection: the shipped .catvba is bridge v2, its modules match resources/bridge/*.bas and it has no orphan storage', () => {
    const sources = b.readSources();
    const facts = b.inspectCatvba(SHIPPED, sources);
    assert.deepEqual(facts.problems, []);
    assert.equal(facts.ok, true);
    assert.equal(facts.projectName, 'C5D_Bridge');
    assert.equal(facts.version, 'c5d-bridge:2');
    assert.equal(facts.runtimeVersion, sources.runtimeVersion);
    assert.deepEqual(facts.modules.map(m => [m.name, m.type, m.same]).sort(),
        [['C5D_Bridge_Push', 'procedural', true], ['C5D_Bridge_Runtime', 'procedural', true]]);
    assert.deepEqual(facts.orphans, []);
    for (const module of readCatvba(SHIPPED).modules) {
        assert.match(module.source, new RegExp(`^Attribute VB_Name = "${module.name}"\\r\\n`));
        assert.ok(!codeOf(module.source).startsWith('Attribute'), module.name);
    }
    assert.match(b.inspectCatvba(__filename).problems[0], /catvba: not an OLE compound file/);
    assert.equal(b.comparable('a  \r\nb\r\n\r\n'), b.comparable('a\nb'));
});

/** A fake ExecuteScript that answers like a healthy CATIA (overrides: step name -> result / { result, ok, timedOut, actions }). */
function fakeCatia(p, overrides = {}) {
    const calls = [];
    const who = lib => `ok\nC5D_Bridge\t${lib}\t${lib}\nC5D_Bridge_Push\t1\nC5D_Bridge_Runtime\t1`;
    const answers = {
        load: 'c5d-bridge:1', 'install-builder': 'ok', 'builder-who': `ok\nC5D_Bridge\t${p.work}\t${p.work}\t1\t10`, 'put-push': 'ok:340',
        'put-runtime': 'ok:530', 'version-v2': 'c5d-bridge:2', 'runtime-version': b.readSources().runtimeVersion, 'who-v2': who(p.work),
        'v-load': 'c5d-bridge:2', 'v-who': who(p.verify), 'v-list': `ok\ntoQI\tC:\\toQI.catvba\t0\nC5D_Bridge\t${p.verify}\t1`,
        'v-runtime': b.readSources().runtimeVersion, 'v-projects': 'ok\ntoQI', 'v-install': 'ok', 'v-echo': `build-probe:${b.PROBE_ARGUMENT}`, 'v-self-guard': 'error:self'
    };
    const callStep = async step => {
        calls.push(step);
        const o = overrides[step.name];
        const spec = o && typeof o === 'object' ? o : { result: o ?? (step.actions ? '(COM object)' : answers[step.name]) };
        const line = { step: step.name, ok: spec.ok ?? true, result: spec.result, resultType: step.actions ? 'ComObject' : 'String', hresult: spec.hresult ?? null };
        const actions = step.actions ? spec.actions ?? step.actions.map((op, action) => ({ action, op, ok: true })) : undefined;
        return { result: line, actions, timedOut: spec.timedOut === true, exitCode: spec.timedOut ? null : 0, match: !spec.timedOut && matchCall(step, line, actions) };
    };
    return { calls, callStep };
}
/** A run context in a temporary repository + %LOCALAPPDATA%, with a small fake base file. */
function context(t, options = {}, inspectOverride = null) {
    const root = fixture(t);
    fs.mkdirSync(path.join(root, 'resources', 'bridge'), { recursive: true });
    fs.writeFileSync(path.join(root, 'resources', 'bridge', 'C5D_Bridge.catvba'), 'old-bridge');
    const p = b.createPaths({ env: { LOCALAPPDATA: path.join(root, 'local') }, root, stamp: '20261121150000' });
    const sources = b.readSources();
    const inspect = inspectOverride ?? (file => file === p.base
        ? { ok: false, version: 'c5d-bridge:1', modules: [{ name: 'C5D_Bridge_Push' }], problems: ['v1'], orphans: [] }
        : { ok: true, version: 'c5d-bridge:2', modules: [], problems: [], orphans: [] });
    const fake = fakeCatia(p, options.overrides);
    const ctx = { p, sources, options: { saveMethod: 'save', install: true, allowOrphans: false, ...options.run }, facts: {}, fileToVerify: null,
        deps: { callStep: fake.callStep, alive: () => true, fatal: () => null, lockProbe: () => ({ inUse: false }), inspect, sha256: f => require('node:crypto').createHash('sha256').update(fs.readFileSync(f)).digest('hex') } };
    return { p, ctx, fake, steps: b.planSteps('build', p, sources, ctx.options) };
}
const entry = (result, name) => result.entries.find(e => e.step === name);

test('fake run: a healthy build saves, closes through C5D_Self, verifies a new copy and installs with a backup', async t => {
    const { p, ctx, fake, steps } = context(t);
    const result = await b.runPlan(steps, ctx);
    assert.equal(result.failed, null, JSON.stringify(result.entries.filter(e => !e.match)));
    assert.ok(result.entries.every(e => e.match), 'every step matched');
    const close = fake.calls.find(s => s.name === 'close');
    assert.deepEqual([close.module, close.func, close.actions.at(-1)], [b.PUSH_MODULE, 'C5D_Self', 'call:$.Close(false)']);
    assert.ok(fake.calls.findIndex(s => s.name === 'remove-builder') < fake.calls.findIndex(s => s.name === 'save'));
    assert.equal(fs.readFileSync(p.target, 'utf8'), 'old-bridge', 'the collected copy of the fake base (same bytes)');
    assert.equal(fs.readFileSync(p.backup, 'utf8'), 'old-bridge');
    assert.equal((fs.statSync(p.verify).mode & 0o200), 0, 'verify copy is read-only');
    assert.notEqual(p.work, p.base);
    assert.match(b.summaryLines(result, ctx).at(-1), /^DONE: resources\/bridge\/C5D_Bridge\.catvba is bridge v2/);
});

test('fake run: --save-method close closes with Close(True) only when everything before matched', async t => {
    const ok = context(t, { run: { saveMethod: 'close' } });
    await b.runPlan(ok.steps, ok.ctx);
    assert.equal(ok.fake.calls.find(s => s.name === 'close').actions.at(-1), 'call:$.Close(true)');
    const bad = context(t, { run: { saveMethod: 'close' }, overrides: { 'who-v2': 'ok\nC5D_Bridge\tx\tx' } });
    const result = await b.runPlan(bad.steps, bad.ctx);
    assert.equal(result.failed, 'who-v2');
    assert.equal(bad.fake.calls.find(s => s.name === 'close').actions.at(-1), 'call:$.Close(false)', 'a broken state is never saved');
});

test('fake run: a second C5D_Bridge stops the build before anything is rewritten, and the copy cannot be closed', async t => {
    const { p, ctx, fake, steps } = context(t, { overrides: { 'builder-who': `ok\nC5D_Bridge\t${'C:\\x.catvba'}\tC:\\x.catvba\t2\t11` } });
    const result = await b.runPlan(steps, ctx);
    assert.equal(result.failed, 'builder-who');
    assert.ok(!fake.calls.some(s => ['put-push', 'save'].includes(s.name)), 'no write after the failed check');
    assert.match(entry(result, 'close').skipped, /cannot identify itself.*restart CATIA/);
    assert.match(entry(result, 'v-close').skipped, /nothing was loaded/);
    assert.equal(fs.readFileSync(p.target, 'utf8'), 'old-bridge');
    assert.ok(!fs.existsSync(p.backup));
});

test('fake run: when the rewrite of C5D_Bridge_Push fails, the close goes through C5D_BuildSelf', async t => {
    const { ctx, fake, steps } = context(t, { overrides: { 'put-push': 'error:put.write:5:x' } });
    const result = await b.runPlan(steps, ctx);
    assert.equal(result.failed, 'put-push');
    const close = fake.calls.find(s => s.name === 'close');
    assert.deepEqual([close.module, close.func], [b.BUILDER_MODULE, 'C5D_BuildSelf']);
});

test('fake run: a timeout (dialog) or a vanished CATIA stops every later call, also the closes', async t => {
    const { ctx, fake, steps } = context(t, { overrides: { save: { timedOut: true, ok: false, result: null } } });
    const result = await b.runPlan(steps, ctx);
    assert.equal(result.failed, 'save');
    assert.match(entry(result, 'close').skipped, /timed out/);
    assert.equal(fake.calls.at(-1).name, 'save');
    const gone = context(t);
    let n = 0;
    gone.ctx.deps.alive = () => (++n < 3);
    const r2 = await b.runPlan(gone.steps, gone.ctx);
    assert.match(entry(r2, 'close').skipped, /CNEXT\.exe\) is gone/);
});

test('fake run: an unreadable saved file stops resumably; orphan storages block the install', async t => {
    const locked = context(t);
    const realCopy = fs.copyFileSync;
    t.mock.method(fs, 'copyFileSync', (src, dst, mode) => {
        if (src === locked.p.work && dst === locked.p.collected) throw Object.assign(new Error('busy'), { code: 'EBUSY' });
        return realCopy(src, dst, mode);
    });
    const result = await b.runPlan(locked.steps, locked.ctx);
    t.mock.restoreAll();
    assert.equal(result.stopped, 'resume');
    assert.match(entry(result, 'collect').message, /--verify ".*-w\.catvba" --install/);
    assert.match(b.summaryLines(result, locked.ctx).at(-1), /^STOPPED at collect/);
    const orphan = context(t, {}, file => (file.endsWith('C5D_Bridge.catvba') && !file.includes('bridge-build-')
        ? { ok: false, version: 'c5d-bridge:1', modules: [{ name: 'C5D_Bridge_Push' }], problems: [], orphans: [] }
        : { ok: true, version: 'c5d-bridge:2', modules: [], problems: [], orphans: ['frmMain'] }));
    const r2 = await b.runPlan(orphan.steps, orphan.ctx);
    assert.equal(r2.failed, 'install');
    assert.match(entry(r2, 'install').message, /orphan storages/);
    assert.equal(fs.readFileSync(orphan.p.target, 'utf8'), 'old-bridge');
    const allowed = context(t, { run: { allowOrphans: true } }, orphan.ctx.deps.inspect);
    assert.equal((await b.runPlan(allowed.steps, allowed.ctx)).failed, null);
});

test('fake run: a v1 base with a bridge file in use stops at the precheck (CATIA must be restarted)', async t => {
    const { ctx, fake, steps } = context(t);
    ctx.deps.lockProbe = file => ({ inUse: file === ctx.p.target });
    const result = await b.runPlan(steps, ctx);
    assert.equal(result.failed, 'precheck');
    assert.equal(result.stopped, 'restart');
    assert.match(entry(result, 'precheck').message, /restart CATIA/);
    assert.equal(fake.calls.length, 0, 'no CATIA call');
});

test('--export-import writes importable .bas files with an Attribute VB_Name line', t => {
    const dir = path.join(fixture(t), 'import');
    execFileSync(process.execPath, [SCRIPT, '--export-import', dir], { encoding: 'utf8', windowsHide: true });
    const push = fs.readFileSync(path.join(dir, 'C5D_Bridge_Push.bas'), 'utf8');
    assert.ok(push.startsWith('Attribute VB_Name = "C5D_Bridge_Push"\r\nOption Explicit\r\n'));
    assert.equal(push.slice(push.indexOf('\r\n') + 2), b.readSources().push);
    assert.ok(fs.readFileSync(path.join(dir, 'C5D_Bridge_Runtime.bas'), 'utf8').startsWith('Attribute VB_Name = "C5D_Bridge_Runtime"\r\n'));
});
