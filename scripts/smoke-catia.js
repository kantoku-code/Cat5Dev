// Real-CATIA smoke test of Push / Pull through bin/catia-run.exe and the C5D_Bridge (developer tool, not shipped).
// Usage: node scripts/smoke-catia.js [--dry-run] [--timeout <ms>]
// Before: npm run compile && npm run compile:bridge. CATIA is running with no dialog open, the VBA editor has been
//   opened once (Alt+F11), and the project toQI has the standard module C5D_BridgeTest whose first line is
//   "' Cat5Dev-bridge-test-owner" (node scripts/probe-vbaproject.js --write-bas writes a body to paste).
// Only toQI/C5D_BridgeTest is written, and the project is never saved:
//   list -> pull -> push (original + "' smoke 日本語 🙂") -> pull (round trip) -> push (original) -> pull (restored) -> list
// Each step is one catia-run.exe (same request.json / result.json route as the extension, under the extension's lock).
// Checks: exit code 0 for every step, the round trip of Japanese and 🙂, the restore, the other modules of toQI and the
// project list unchanged, and vbscript.dll not loaded in CNEXT.exe (Process.Modules / tasklist /M, before and after).
// The bundled resources/bridge/C5D_Bridge.catvba must be bridge v2 (checked offline first; a v1 bridge cannot be closed again
// and would stay loaded in CATIA until it restarts). Rebuild it with docs/dev/bridge-contract.md 11 when the check fails.
// Writes temp/catia-smoke-<yyyyMMdd-HHmmss>.json. Exit: 0 passed, 1 failed, 2 bad arguments, missing build or bundled bridge not v2,
//   3 everything passed but vbscript.dll was already loaded before the first step (restart CATIA and re-run).
// --dry-run checks the build and prints the plan; it neither takes the lock nor contacts CATIA nor writes files.
const fs = require('node:fs');
const path = require('node:path');
const { createHash } = require('node:crypto');
const { moduleSnapshot, timestamp } = require('./probe-catscript');

const ROOT = path.resolve(__dirname, '..');
const TEMP = path.join(ROOT, 'temp');
const RUNNER = path.join(ROOT, 'bin', 'catia-run.exe');
const BUNDLED_BRIDGE = path.join(ROOT, 'resources', 'bridge', 'C5D_Bridge.catvba');
const PROJECT = 'toQI';
const MODULE = 'C5D_BridgeTest';
/** First line of C5D_BridgeTest (scripts/probe-vbaproject.js BRIDGE_BODY); nothing else is ever written. */
const OWNER_MARKER = "' Cat5Dev-bridge-test-owner";
const SMOKE_LINE = "' smoke 日本語 🙂";
const BRIDGE_PROJECT = 'C5D_Bridge';
const DEFAULT_TIMEOUT_MS = 60000;
const USAGE = 'Usage: node scripts/smoke-catia.js [--dry-run] [--timeout <ms>]\n'
    + `  Writes only ${PROJECT}/${MODULE} (never saved) and pushes its original text back at the end.`;

/** { dryRun, timeoutMs } or { help: true }. Throws on bad arguments. */
function parseArgs(args) {
    const options = { dryRun: false, timeoutMs: DEFAULT_TIMEOUT_MS };
    for (let i = 0; i < args.length; i++) {
        const arg = args[i];
        if (arg === '--help' || arg === '-h') return { help: true };
        if (arg === '--dry-run') options.dryRun = true;
        else if (arg === '--timeout') {
            const value = Number(args[++i]);
            if (!Number.isInteger(value) || value < 1000) throw new Error(`--timeout requires milliseconds (>= 1000)\n${USAGE}`);
            options.timeoutMs = value;
        } else throw new Error(`Unknown option: ${arg}\n${USAGE}`);
    }
    return options;
}
const normalize = text => text.replace(/\r\n?/g, '\n').trimEnd();
/** The pushed text: the original (as pulled) plus the smoke line, CRLF like the VBE. */
const withSmokeLine = original => original.trimEnd() + '\r\n' + SMOKE_LINE;
/** Modules other than C5D_BridgeTest: [{ name, type, sha256 of the normalized text }] sorted by name. */
function otherModules(modules) {
    return modules.filter(m => m.name.toLowerCase() !== MODULE.toLowerCase())
        .map(m => ({ name: m.name, type: m.type, sha256: createHash('sha256').update(normalize(m.text)).digest('hex') }))
        .sort((a, b) => a.name < b.name ? -1 : a.name > b.name ? 1 : 0);
}
/** Names whose type or text differs between two otherModules() lists (or that exist in only one). */
function diffModules(before, after) {
    const key = list => new Map(list.map(m => [m.name, `${m.type}:${m.sha256}`]));
    const a = key(before), b = key(after);
    return [...new Set([...a.keys(), ...b.keys()])].filter(name => a.get(name) !== b.get(name)).sort();
}
/** Why the original text of C5D_BridgeTest must not be touched, or null. */
function refuseReason(found) {
    if (!found) return `${MODULE} is not in ${PROJECT}; add it by hand (node scripts/probe-vbaproject.js --write-bas) and re-run`;
    if (found.type !== 1) return `${MODULE} is not a standard module (type ${found.type})`;
    if (!found.text.startsWith(OWNER_MARKER)) return `${MODULE} does not start with the owner marker ${JSON.stringify(OWNER_MARKER)}`;
    if (found.text.includes(SMOKE_LINE)) return `${MODULE} already contains ${JSON.stringify(SMOKE_LINE)} (left by an interrupted run); remove that line in the VBA editor and re-run`;
    return null;
}
/** Overall verdict from report.checks and the vbscript.dll snapshots: { passed, exitCode, verdict }. */
function judge(report) {
    const failed = Object.entries(report.checks).filter(([, ok]) => ok !== true).map(([name]) => name);
    const before = report.snapshots.before?.loaded, after = report.snapshots.after?.loaded;
    if (report.error || failed.length) return { passed: false, exitCode: 1, verdict: `FAILED: ${[report.error, failed.length ? 'checks ' + failed.join(', ') : ''].filter(Boolean).join('; ')}` };
    if (after === true && before === false) return { passed: false, exitCode: 1, verdict: 'FAILED: vbscript.dll was loaded into CNEXT.exe during the run' };
    if (before === true) return { passed: false, exitCode: 3, verdict: 'Push / Pull passed; vbscript.dll undetermined (already loaded before the first step: restart CATIA, run nothing else, re-run)' };
    if (after !== false) return { passed: false, exitCode: 1, verdict: 'FAILED: vbscript.dll could not be checked' };
    return { passed: true, exitCode: 0, verdict: 'PASSED: Push / Pull round trip without vbscript.dll' };
}
/** temp/catia-smoke-<time>.json; never overwrites an earlier report. */
function uniqueReportPath(directory) {
    const base = path.join(directory, `catia-smoke-${timestamp()}`);
    let candidate = base + '.json';
    for (let i = 2; fs.existsSync(candidate); i++) candidate = `${base}-${i}.json`;
    return candidate;
}
function loadModules() {
    try {
        return {
            ...require('../out/syncOperation'), ...require('../out/catiaExecution'),
            ...require('../out/bridgeFiles'), ...require('../out/catiaTypes')
        };
    } catch (error) { throw Object.assign(new Error(`out/ is missing or stale; run npm run compile (${error.message})`), { exitCode: 2 }); }
}
function checkBuild() {
    const missing = [RUNNER, BUNDLED_BRIDGE].filter(file => !fs.existsSync(file));
    if (missing.length) throw Object.assign(new Error(`Missing ${missing.join(', ')}; run npm run compile:bridge`), { exitCode: 2 });
}
/**
 * Offline check of the bundled .catvba (no CATIA): a v1 bridge has no C5D_Self, so catia-run.exe could not close it and it would
 * stay loaded until CATIA restarts. The smoke never contacts CATIA unless this is a valid v2 bridge matching resources/bridge/*.bas.
 */
function inspectBundledBridge(file = BUNDLED_BRIDGE) {
    const { inspectCatvba, readSources } = require('./build-bridge-catvba');
    const facts = inspectCatvba(file, readSources(ROOT));
    return { ok: facts.ok, problems: facts.problems, version: facts.version, runtimeVersion: facts.runtimeVersion, orphans: facts.orphans.length };
}
/** Compact copy of a probe-catscript moduleSnapshot for the summary (the full one stays in the report). */
const dllState = snap => snap && Object.fromEntries(Object.entries(snap.dlls).map(([dll, d]) => [dll, d.loaded]));

/**
 * Runs the smoke into `report`. deps (tests only) replace the CATIA-facing parts:
 * { snapshot(label), executor (a CatiaExecutor), manager (SyncOperationManager), bridge (BridgeFiles) }.
 */
async function run(report, options, deps = {}) {
    const m = loadModules();
    if (!deps.executor) checkBuild();
    const bridge = deps.bridge ?? new m.BridgeFiles({ bundledPath: BUNDLED_BRIDGE });
    const snapshot = deps.snapshot ?? moduleSnapshot;
    const executor = deps.executor ?? m.bridgeExecutor(RUNNER);
    const bundled = report.bundledBridge = (deps.inspectBridge ?? inspectBundledBridge)();
    if (!bundled.ok) {
        throw Object.assign(new Error(`${BUNDLED_BRIDGE} is not a valid v2 bridge (${bundled.problems.join('; ')}); `
            + 'rebuild it first (docs/dev/bridge-contract.md 11). CATIA was not contacted.'), { exitCode: 2 });
    }
    if (bundled.orphans) report.notes.push(`The bundled bridge has ${bundled.orphans} orphan storage(s) (bridge-contract 11.4); do not ship it.`);
    if (options.dryRun) {
        report.plan = {
            runner: RUNNER, bundledBridge: BUNDLED_BRIDGE, bundledSha256: bridge.bundledHash(), master: bridge.masterPath(),
            project: PROJECT, module: MODULE, smokeLine: SMOKE_LINE, timeoutMs: options.timeoutMs,
            steps: ['list', 'pull', 'push (+smoke line)', 'pull (verify)', 'push (original)', 'pull (verify restore)', 'list'],
            sampleRequest: JSON.parse(m.buildRequestJson({ operationId: '<operation>', stepId: '<step>', kind: 'push', targetProject: PROJECT,
                directory: '', modules: [{ name: MODULE, type: 1, text: withSmokeLine(OWNER_MARKER) }] }))
        };
        report.checks = {};
        return;
    }
    const first = report.snapshots.before = snapshot('before');
    if (!first.cnextPids.length) throw new Error('CNEXT.exe is not running; start CATIA');
    if (first.cnextPids.length > 1) report.notes.push(`More than one CNEXT.exe is running (${first.cnextPids.join(', ')})`);
    if (first.powershell.error) report.notes.push(`Process.Modules could not be read (${first.powershell.error}); judged by tasklist /M only`);

    const context = (deps.manager ?? new m.SyncOperationManager()).acquire('smoke-catia');
    const checks = report.checks;
    try {
        const prepared = bridge.prepare();
        report.bridge = { master: prepared.master, action: prepared.action, cleanup: prepared.cleanup };
        const step = async (name, kind, configure) => {
            const entry = { name, kind, exitCode: null, ms: null };
            report.steps.push(entry);
            const request = m.createCatiaRequest(context, kind, kind === 'list' ? '' : PROJECT);
            if (configure) configure(request);
            const started = Date.now();
            const execute = async (invocation, timeoutMs) => {
                const output = await executor(invocation, timeoutMs);
                entry.exitCode = output.error ? (output.error.code ?? (output.error.killed ? 'killed' : String(output.error))) : 0;
                entry.stdout = output.stdout.trim();
                if (output.stderr.trim()) entry.stderr = output.stderr.trim();
                return output;
            };
            try {
                const result = await m.runCatiaRequest(context, request, { execute, bridgeMaster: prepared.master, timeoutMs: options.timeoutMs });
                Object.assign(entry, { ms: Date.now() - started, expected: result.expectedCount, succeeded: result.succeededCount });
                return { request, result };
            } catch (error) {
                Object.assign(entry, { ms: Date.now() - started, error: String(error), uncertain: error.uncertain === true,
                    failures: error.result?.failures, exitCode: entry.exitCode ?? error.exitCode ?? null });
                throw error;
            }
        };
        const list = async name => (await step(name, 'list')).result.projects;
        const pull = async name => {
            const { request, result } = await step(name, 'pull');
            return result.items.map(item => ({ name: item.name, type: item.type,
                text: fs.readFileSync(path.join(request.directory, item.file), 'utf8') }));
        };
        const push = (name, text) => step(name, 'push', request => { request.modules = [{ name: MODULE, type: 1, text }]; request.deletions = []; });
        const testModule = modules => modules.find(mod => mod.name.toLowerCase() === MODULE.toLowerCase());

        const projects = await list('list');
        report.projects = projects;
        checks.targetListed = projects.includes(PROJECT);
        checks.bridgeHidden = !projects.includes(BRIDGE_PROJECT);
        if (!checks.targetListed) throw new Error(`Project ${PROJECT} not found (projects: ${projects.join(', ') || 'none'})`);

        const before = await pull('pull');
        const original = testModule(before);
        const refused = refuseReason(original);
        if (refused) throw new Error(refused);
        report.original = { text: original.text, lines: original.text.split(/\r\n|\n/).length };
        const others = otherModules(before);
        report.otherModules = others.map(mod => `${mod.name}(${mod.type})`);

        const modified = withSmokeLine(original.text);
        let pushed = false;
        try {
            pushed = true;
            await push('push', modified);
            const roundTrip = await pull('pull (verify)');
            const got = testModule(roundTrip);
            checks.roundTrip = Boolean(got) && normalize(got.text) === normalize(modified);
            checks.unicodeKept = Boolean(got) && got.text.includes(SMOKE_LINE);
            if (!checks.roundTrip) report.roundTripText = got ? got.text : null;
            report.otherModulesChangedByPush = diffModules(others, otherModules(roundTrip));
        } finally {
            if (pushed) {
                try {
                    await push('push (restore)', original.text);
                    const after = await pull('pull (verify restore)');
                    const restored = testModule(after);
                    checks.restored = Boolean(restored) && normalize(restored.text) === normalize(original.text);
                    if (!checks.restored) report.restoredText = restored ? restored.text : null;
                    report.otherModulesChanged = diffModules(others, otherModules(after));
                    checks.otherModulesUnchanged = report.otherModulesChanged.length === 0 && (report.otherModulesChangedByPush ?? []).length === 0;
                } catch (error) {
                    checks.restored = false;
                    report.notes.push(`Restore failed: ${error}. ${PROJECT}/${MODULE} may still end with ${JSON.stringify(SMOKE_LINE)}; `
                        + `its original text is in "original.text" of this report (remove the line by hand, do not save).`);
                    throw error;
                }
            }
        }
        const projectsAfter = await list('list (after)');
        checks.projectsUnchanged = JSON.stringify(projectsAfter) === JSON.stringify(projects);
        if (!checks.projectsUnchanged) report.projectsAfter = projectsAfter;
    } finally {
        checks.allExitZero = report.steps.length > 0 && report.steps.every(entry => entry.exitCode === 0);
        try { context.close(); } catch (error) { report.notes.push(`Lock release failed: ${error}`); }
        if (report.steps.some(entry => entry.uncertain)) {
            report.notes.push('A step did not confirm completion (e.g. timeout); the Cat5Dev lock may be kept. Check CATIA for a dialog, '
                + 'then release it from VSCode (CATIA: Recover Synchronization Lock).');
        }
        report.snapshots.after = snapshot('after');
    }
}
function formatSummary(report) {
    const lines = report.reportPath ? [`report: ${report.reportPath}`] : [];
    if (report.plan) lines.push(JSON.stringify(report.plan, null, 2));
    if (report.bridge) lines.push(`bridge master: ${report.bridge.master} (${report.bridge.action})`);
    for (const entry of report.steps) {
        lines.push(`  ${entry.name.padEnd(22)} exit=${String(entry.exitCode).padEnd(4)} ${String(entry.ms ?? '-').padStart(6)}ms`
            + (entry.expected !== undefined ? `  ${entry.succeeded}/${entry.expected}` : '') + (entry.error ? `  ${entry.error}` : ''));
    }
    for (const [name, ok] of Object.entries(report.checks ?? {})) lines.push(`  check ${name}: ${ok === true ? 'ok' : ok === false ? 'FAIL' : '-'}`);
    for (const label of ['before', 'after']) {
        const state = dllState(report.snapshots[label]);
        if (state) lines.push(`  ${label}: ${Object.entries(state).map(([dll, loaded]) => `${dll}=${loaded ? 'loaded' : 'no'}`).join(' ')}`);
    }
    if (report.error) lines.push(`error: ${report.error}`);
    for (const note of report.notes) lines.push(`note: ${note}`);
    lines.push(`verdict: ${report.verdict}`);
    return lines.join('\n');
}
async function main() {
    let options;
    try { options = parseArgs(process.argv.slice(2)); } catch (error) { console.error(error.message); process.exitCode = 2; return; }
    if (options.help) { console.log(USAGE); return; }
    const report = { startedAt: new Date().toISOString(), dryRun: options.dryRun, project: PROJECT, module: MODULE,
        timeoutMs: options.timeoutMs, steps: [], checks: {}, snapshots: {}, notes: [] };
    let exitCode;
    try { await run(report, options); } catch (error) {
        report.error = String(error);
        if (error.exitCode === 2) exitCode = 2;
    }
    report.finishedAt = new Date().toISOString();
    if (options.dryRun) {
        Object.assign(report, { passed: !report.error, verdict: report.error ? `dry run failed: ${report.error}` : 'dry run: build found, CATIA not contacted' });
        process.exitCode = report.error ? (exitCode ?? 1) : 0;
        console.log(formatSummary(report));
        return;
    }
    const outcome = judge(report);
    Object.assign(report, { passed: outcome.passed, verdict: outcome.verdict });
    if (exitCode !== 2) {
        fs.mkdirSync(TEMP, { recursive: true });
        report.reportPath = uniqueReportPath(TEMP);
        fs.writeFileSync(report.reportPath, JSON.stringify(report, null, 2));
    }
    process.exitCode = exitCode ?? outcome.exitCode;
    console.log(formatSummary(report));
}
if (require.main === module) main();
module.exports = { run, parseArgs, withSmokeLine, otherModules, diffModules, refuseReason, judge, SMOKE_LINE, OWNER_MARKER, PROJECT, MODULE };
