const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { parseArgs, judgeCall, parseProjectsResult, resolveLibrary, BRIDGE_BODY, MARKER, CALL_STEPS }
    = require('../scripts/probe-vbaproject');
const csc = path.join(process.env.SystemRoot || 'C:\\Windows', 'Microsoft.NET', 'Framework64', 'v4.0.30319', 'csc.exe');

test('VBAProject probe parses modes, library and steps', () => {
    assert.equal(parseArgs(['--help']).mode, 'help');
    assert.deepEqual(parseArgs(['--call']), { mode: 'call', library: null, steps: CALL_STEPS });
    assert.deepEqual(parseArgs(['--call', '--library', 'C:\\x\\toQI.catvba', 'vbaproj-ping']),
        { mode: 'call', library: 'C:\\x\\toQI.catvba', steps: ['vbaproj-ping'] });
    assert.deepEqual(parseArgs(['--call', 'vbaproj-filename', 'vbaproj-filename-empty']).steps,
        ['vbaproj-filename-basename', 'vbaproj-filename-empty', 'vbaproj-filename-project']);
    assert.deepEqual(parseArgs(['--write-bas']), { mode: 'write-bas', library: null, steps: [] });
    assert.throws(() => parseArgs([]), /exactly one mode/);
    assert.throws(() => parseArgs(['--write-bas', '--call']), /exactly one mode/);
    for (const removed of ['--info', '--install', '--remove']) {
        assert.throws(() => parseArgs([removed]), new RegExp(`${removed} no longer works.*smoke-catia`), removed);
    }
    assert.throws(() => parseArgs(['--call', '--library']), /requires a path/);
    assert.throws(() => parseArgs(['--write-bas', '--library', 'x']), /only for --call/);
    assert.throws(() => parseArgs(['--call', 'basic']), /Unknown/);
});
test('VBAProject probe resolves the library from --library or the info file', t => {
    const dir = fixture(t), info = path.join(dir, 'info.json');
    assert.deepEqual(resolveLibrary('C:\\given.catvba', info), { library: 'C:\\given.catvba', source: '--library' });
    assert.throws(() => resolveLibrary(null, info), /does not exist/);
    fs.writeFileSync(info, JSON.stringify({ library: '' }));
    assert.throws(() => resolveLibrary(null, info), /no "library"/);
    fs.writeFileSync(info, JSON.stringify({ library: 'D:\\p\\toQI.catvba' }));
    assert.deepEqual(resolveLibrary(null, info), { library: 'D:\\p\\toQI.catvba', source: info });
});
test('VBAProject probe verdicts', () => {
    const ping = { step: 'vbaproj-ping', ok: true, match: true, result: 'vba:Long:cat5dev-param', loadedAfter: false };
    assert.equal(judgeCall(false, [ping, { step: 'vbaproj-sub', ok: true, loadedAfter: false }]),
        'ExecuteScript(VBAProject) ran real VBA (Long) and did not load vbscript.dll');
    assert.equal(judgeCall(false, [ping, { step: 'vbaproj-projects', ok: true, loadedAfter: true }]),
        'ExecuteScript(VBAProject) ran real VBA (Long) and loaded vbscript.dll (first seen after step vbaproj-projects)');
    assert.match(judgeCall(true, [ping]), /^ExecuteScript\(VBAProject\) ran real VBA \(Long\); vbscript\.dll undetermined/);
    assert.match(judgeCall(false, [{ ...ping, match: false, result: 'vba:Integer:cat5dev-param' }]), /As Long was ignored/);
    assert.match(judgeCall(false, [{ ...ping, ok: false, match: false, result: null, hresult: '0x80004005' }]), /failed to call C5D_Ping \(0x80004005\) and did not load/);
    assert.match(judgeCall(false, [{ ...ping, ok: false, timedOut: true }]), /timed out/);
    assert.match(judgeCall(false, []), /not run/);
});
test('C5D_Projects results are parsed', () => {
    assert.deepEqual(parseProjectsResult('count=2|toQI|Other;err=0'), { count: 2, names: ['toQI', 'Other'], err: 0 });
    assert.deepEqual(parseProjectsResult('count=0;err=0'), { count: 0, names: [], err: 0 });
    assert.deepEqual(parseProjectsResult('apc=nothing;err=429'), { count: null, names: [], failedAt: 'apc', err: 429 });
    assert.deepEqual(parseProjectsResult('projects=nothing;err=91'), { count: null, names: [], failedAt: 'projects', err: 91 });
    assert.equal(parseProjectsResult('other'), null);
    assert.equal(parseProjectsResult(null), null);
});
test('C5D_BridgeTest body is ASCII VBA with the owner marker and no Attribute line', () => {
    assert.ok(/^[\x20-\x7e\r\n]*$/.test(BRIDGE_BODY));
    assert.ok(BRIDGE_BODY.startsWith(MARKER + '\r\n'));
    assert.ok(!/^Attribute /m.test(BRIDGE_BODY));
    assert.ok(!/\bMsgBox\b/i.test(BRIDGE_BODY));
    assert.ok(!/[^\r]\n/.test(BRIDGE_BODY));
    for (const name of ['Function C5D_Ping(ByVal a As Variant) As Variant', 'Sub C5D_Noop()', 'Function C5D_Projects() As Variant']) {
        assert.ok(BRIDGE_BODY.includes('Public ' + name), name);
    }
    assert.ok(BRIDGE_BODY.includes('Dim n As Long'));
});
test('EvalProbe vbaproj steps take the library path and call ExecuteScript with type 2', { skip: !fs.existsSync(csc) }, t => {
    const exe = path.join(fixture(t), 'eval-probe.exe');
    execFileSync(csc, ['/nologo', '/warnaserror+', '/target:exe', '/platform:anycpu', `/out:${exe}`,
        path.join(__dirname, '..', 'tools', 'catia-run', 'EvalProbe.cs')], { windowsHide: true, stdio: 'pipe' });
    const library = 'C:\\a b\\toQI.catvba';
    const lines = execFileSync(exe, ['--dry-run', 'vbaproj-ping', library, 'vbaproj-sub', library, 'vbaproj-projects', library, 'vbaproj-filename', library, '--project', 'toQI'],
        { windowsHide: true, encoding: 'utf8', timeout: 10000 }).trim().split(/\r?\n/).map(line => JSON.parse(line));
    assert.deepEqual(lines.map(l => [l.step, l.method, l.library, l.libraryType, l.program, l.function]), [
        ['vbaproj-ping', 'ExecuteScript', library, 2, 'C5D_BridgeTest', 'C5D_Ping'],
        ['vbaproj-sub', 'ExecuteScript', library, 2, 'C5D_BridgeTest', 'C5D_Noop'],
        ['vbaproj-projects', 'ExecuteScript', library, 2, 'C5D_BridgeTest', 'C5D_Projects'],
        ['vbaproj-filename-basename', 'ExecuteScript', 'toQI.catvba', 2, 'C5D_BridgeTest', 'C5D_Ping'],
        ['vbaproj-filename-empty', 'ExecuteScript', '', 2, 'C5D_BridgeTest', 'C5D_Ping'],
        ['vbaproj-filename-project', 'ExecuteScript', 'toQI', 2, 'C5D_BridgeTest', 'C5D_Ping']
    ]);
    assert.equal(lines[0].expected, 'vba:Long:cat5dev-param');
    assert.equal(lines[0].parameters, '[cat5dev-param]');
    assert.ok(lines.every(l => l.retryByRef === true));
    assert.throws(() => execFileSync(exe, ['vbaproj-ping'], { windowsHide: true, stdio: 'pipe' }), error => error.status === 2);
    assert.throws(() => execFileSync(exe, ['vbaproj-ping', '--byref'], { windowsHide: true, stdio: 'pipe' }), error => error.status === 2);
});
