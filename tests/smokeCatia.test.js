const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { fixture, result } = require('./helpers');
const { SyncOperationManager } = require('../out/syncOperation');
const { BridgeFiles } = require('../out/bridgeFiles');
const { run, parseArgs, withSmokeLine, otherModules, diffModules, refuseReason, judge, SMOKE_LINE, OWNER_MARKER, PROJECT, MODULE }
    = require('../scripts/smoke-catia');

const ORIGINAL = `${OWNER_MARKER}\r\nOption Explicit\r\nPublic Sub C5D_Noop()\r\nEnd Sub`;
/** A catia-run.exe stand-in that follows docs/dev/bridge-contract.md 4-5 on an in-memory toQI. */
function fakeCatia(options = {}) {
    const state = { calls: [], modules: new Map([['Module1', { type: 1, text: 'Sub A()\r\nEnd Sub' }], ['Form1', { type: 3, text: '' }],
        [MODULE, { type: 1, text: options.original ?? ORIGINAL }]]) };
    state.executor = async invocation => {
        const request = JSON.parse(fs.readFileSync(invocation.request, 'utf8'));
        const directory = path.dirname(invocation.request);
        state.calls.push({ kind: request.kind, modules: request.modules });
        let value;
        if (request.kind === 'list') value = result(request, { expectedCount: 2, succeededCount: 2, projects: ['Other', PROJECT] });
        else if (request.kind === 'pull') {
            const items = [...state.modules].map(([name, mod], i) => {
                fs.writeFileSync(path.join(directory, `m${i}.txt`), mod.text, 'utf8');
                return { name, type: mod.type, file: `m${i}.txt` };
            });
            value = result(request, { expectedCount: items.length, succeededCount: items.length, items });
        } else {
            for (const mod of request.modules) state.modules.set(mod.name, { type: mod.type, text: options.mangle ? mod.text.replace(/[^\x00-\x7f]/gu, '?') : mod.text });
            const n = request.modules.length;
            value = result(request, { expectedCount: n, succeededCount: n, updatedCount: n });
        }
        fs.writeFileSync(invocation.result, JSON.stringify(value));
        return { stdout: `${request.kind} ok`, stderr: '' };
    };
    return state;
}
function deps(t, catia, loaded = { before: false, after: false }) {
    const root = fixture(t);
    const bundledPath = path.join(root, 'C5D_Bridge.catvba');
    fs.writeFileSync(bundledPath, 'bridge');
    return {
        executor: catia.executor,
        manager: new SyncOperationManager({ lockRoot: path.join(root, 'locks'), tempRoot: root }),
        bridge: new BridgeFiles({ bundledPath, bridgeDir: path.join(root, 'bridge') }),
        inspectBridge: () => ({ ok: true, problems: [], version: 'c5d-bridge:2', runtimeVersion: 'c5d-runtime:1', orphans: 0 }),
        snapshot: label => ({ label, cnextPids: [1], powershell: { error: null }, dlls: { 'vbscript.dll': { loaded: loaded[label] } }, loaded: loaded[label] })
    };
}
const newReport = () => ({ steps: [], checks: {}, snapshots: {}, notes: [] });

test('smoke runs list, pull, push, pull, push (restore), pull, list and restores C5D_BridgeTest', async t => {
    const catia = fakeCatia();
    const report = newReport();
    await run(report, { timeoutMs: 5000 }, deps(t, catia));
    assert.deepEqual(report.steps.map(s => [s.name, s.exitCode]), [['list', 0], ['pull', 0], ['push', 0], ['pull (verify)', 0],
        ['push (restore)', 0], ['pull (verify restore)', 0], ['list (after)', 0]]);
    assert.ok(report.steps.every(s => Number.isInteger(s.ms)));
    assert.deepEqual(report.checks, { targetListed: true, bridgeHidden: true, roundTrip: true, unicodeKept: true, restored: true,
        otherModulesUnchanged: true, projectsUnchanged: true, allExitZero: true });
    const pushes = catia.calls.filter(c => c.kind === 'push');
    assert.deepEqual(pushes.map(c => c.modules), [[{ name: MODULE, type: 1, text: withSmokeLine(ORIGINAL) }], [{ name: MODULE, type: 1, text: ORIGINAL }]]);
    assert.equal(catia.modules.get(MODULE).text, ORIGINAL);
    assert.deepEqual(judge(report), { passed: true, exitCode: 0, verdict: 'PASSED: Push / Pull round trip without vbscript.dll' });
    assert.equal(judge({ ...report, snapshots: { before: { loaded: true }, after: { loaded: true } } }).exitCode, 3);
    assert.equal(judge({ ...report, snapshots: { before: { loaded: false }, after: { loaded: true } } }).exitCode, 1);
});

test('a mangled round trip fails but the original text is still pushed back', async t => {
    const catia = fakeCatia({ mangle: true });
    const report = newReport();
    await run(report, { timeoutMs: 5000 }, deps(t, catia));
    assert.equal(report.checks.roundTrip, false);
    assert.equal(report.checks.unicodeKept, false);
    assert.equal(report.checks.restored, true);
    assert.equal(catia.modules.get(MODULE).text, ORIGINAL);
    assert.equal(judge(report).exitCode, 1);
});

test('the smoke refuses to write a module without the owner marker or with a leftover smoke line', async t => {
    for (const original of ['Sub Other()\r\nEnd Sub', withSmokeLine(ORIGINAL)]) {
        const catia = fakeCatia({ original });
        const report = newReport();
        await assert.rejects(run(report, { timeoutMs: 5000 }, deps(t, catia)), /owner marker|already contains/);
        assert.ok(!catia.calls.some(c => c.kind === 'push'));
        assert.equal(catia.modules.get(MODULE).text, original);
    }
    assert.match(refuseReason(undefined), /not in toQI/);
    assert.match(refuseReason({ type: 2, text: ORIGINAL }), /not a standard module/);
    assert.equal(refuseReason({ type: 1, text: ORIGINAL }), null);
});

test('a bundled bridge that is not v2 stops the smoke before the lock, the snapshot and CATIA', async t => {
    const catia = fakeCatia();
    const d = deps(t, catia);
    let snapshots = 0;
    d.snapshot = () => { snapshots++; return {}; };
    d.inspectBridge = () => ({ ok: false, problems: ['C5D_Bridge_Push version is c5d-bridge:1, expected c5d-bridge:2'], orphans: 25 });
    const report = newReport();
    await assert.rejects(run(report, { timeoutMs: 5000 }, d), e => e.exitCode === 2 && /not a valid v2 bridge.*c5d-bridge:1/.test(e.message));
    assert.equal(catia.calls.length, 0);
    assert.equal(snapshots, 0);
    assert.equal(report.steps.length, 0);
});

test('smoke helpers', () => {
    assert.deepEqual(parseArgs([]), { dryRun: false, timeoutMs: 60000 });
    assert.deepEqual(parseArgs(['--dry-run', '--timeout', '90000']), { dryRun: true, timeoutMs: 90000 });
    assert.deepEqual(parseArgs(['-h']), { help: true });
    assert.throws(() => parseArgs(['--timeout', 'x']), /--timeout/);
    assert.throws(() => parseArgs(['--target', 'toQI']), /Unknown option/);
    assert.equal(withSmokeLine('a\r\nb\r\n'), 'a\r\nb\r\n' + SMOKE_LINE);
    const before = otherModules([{ name: 'B', type: 1, text: 'x\r\n' }, { name: 'A', type: 2, text: 'y' }, { name: MODULE.toLowerCase(), type: 1, text: 'z' }]);
    assert.deepEqual(before.map(m => m.name), ['A', 'B']);
    assert.deepEqual(diffModules(before, otherModules([{ name: 'B', type: 1, text: 'x' }, { name: 'A', type: 2, text: 'y2' }, { name: 'C', type: 1, text: '' }])), ['A', 'C']);
});
