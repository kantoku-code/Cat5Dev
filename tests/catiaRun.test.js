// catia-run.exe (tools/catia-run/CatiaRun*.cs) without CATIA: the build (runtime embedded), the command line and request.json
// checks through --validate-only, and the bridge procedure (docs/dev/bridge-contract.md 5-7) against a fake ExecuteScript host.
// No test here connects to CATIA: every exe run is either --validate-only or has arguments that fail before the connection,
// and the harness exe has its own Main that never calls GetActiveObject.
const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync, spawnSync } = require('node:child_process');
const { findCsc, sourceFiles, cscArguments, RESOURCE_NAME } = require('../scripts/build-bridge');

const root = path.resolve(__dirname, '..');
const csc = process.platform === 'win32' ? findCsc() : null;
const skip = !csc;
const RUNTIME_BAS = path.join(root, 'resources', 'bridge', 'C5D_Bridge_Runtime.bas');
const RUNTIME_TEXT = fs.readFileSync(RUNTIME_BAS, 'latin1');
const RT_VERSION = /^Private Const RT_VERSION As String = "(c5d-runtime:\d+)"\r?$/m.exec(RUNTIME_TEXT)[1];
const PUSH = 'C5D_Bridge_Push';
/** validateCatiaResult of the extension (out/ after npm run compile), when it is there: result.json must pass it. */
const validateCatiaResult = (() => {
    try { return require('../out/catiaTypes').validateCatiaResult; } catch { return undefined; }
})();
const RT = 'C5D_Bridge_Runtime';

// ---- the harness: BridgeSession with a scripted IScriptHost (compiled with the production sources, /main:CatiaRunHarness)
const HARNESS = String.raw`
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

public sealed class FakeProject { public string FileName { get; set; } }

public sealed class FakeSelf
{
    readonly CatiaRunHarness owner;
    readonly string closeError;
    public FakeSelf(CatiaRunHarness owner, string name, string displayName, string fileName, string closeError)
    {
        this.owner = owner; Name = name; DisplayName = displayName; VBProject = new FakeProject { FileName = fileName }; this.closeError = closeError;
    }
    public string Name { get; set; }
    public string DisplayName { get; set; }
    public FakeProject VBProject { get; set; }
    public void Close(bool save)
    {
        if (closeError != null) { throw new COMException(closeError, unchecked((int)0x80004005)); }
        owner.Closes.Add(DisplayName + (save ? " (save)" : ""));
    }
}

public sealed class CatiaRunHarness : IScriptHost
{
    readonly List<Dictionary<string, object>> replies;
    int next;
    public readonly List<string> Calls = new List<string>();
    public readonly List<string> Closes = new List<string>();
    public string Mismatch;

    CatiaRunHarness(List<Dictionary<string, object>> replies) { this.replies = replies; }

    public object Execute(string library, string module, string function, object[] args)
    {
        StringBuilder call = new StringBuilder("{\"library\":" + JsonText.Quote(library) + ",\"module\":" + JsonText.Quote(module)
            + ",\"function\":" + JsonText.Quote(function) + ",\"args\":[");
        for (int i = 0; i < args.Length; i++) { if (i > 0) { call.Append(','); } call.Append(JsonText.Quote((string)args[i])); }
        Calls.Add(call.Append("]}").ToString());
        if (next >= replies.Count) { Mismatch = Mismatch ?? "no reply left for " + module + "." + function; throw new InvalidOperationException(Mismatch); }
        Dictionary<string, object> reply = replies[next++];
        object expected;
        if ((string)reply["module"] != module || (string)reply["function"] != function
            || (reply.TryGetValue("library", out expected) && (string)expected != library))
        {
            Mismatch = Mismatch ?? "call " + Calls.Count + " is " + library + " " + module + "." + function + ", expected "
                + (reply.ContainsKey("library") ? (string)reply["library"] + " " : "") + reply["module"] + "." + reply["function"];
            throw new InvalidOperationException(Mismatch);
        }
        object value;
        if (reply.TryGetValue("string", out value)) { return (string)value; }
        if (reply.TryGetValue("int", out value)) { return (int)(double)value; }
        if (reply.TryGetValue("empty", out value)) { return null; }
        if (reply.TryGetValue("throw", out value))
        {
            Dictionary<string, object> t = (Dictionary<string, object>)value;
            throw new System.Reflection.TargetInvocationException(new COMException((string)t["message"], (int)(double)t["hresult"]));
        }
        Dictionary<string, object> s = (Dictionary<string, object>)reply["self"];
        object closeError;
        s.TryGetValue("closeError", out closeError);
        return new FakeSelf(this, (string)s["name"], (string)s["displayName"], (string)s["fileName"], closeError as string);
    }

    static int Main(string[] args)
    {
        Dictionary<string, object> scenario = (Dictionary<string, object>)JsonParser.Parse(File.ReadAllText(args[0], Encoding.UTF8));
        List<Dictionary<string, object>> replies = new List<Dictionary<string, object>>();
        foreach (object r in (List<object>)scenario["replies"]) { replies.Add((Dictionary<string, object>)r); }
        CatiaRunHarness host = new CatiaRunHarness(replies);
        string error;
        RuntimeSource runtime = RuntimeSource.Load(out error);
        if (runtime == null) { Console.Error.WriteLine(error); return 99; }
        BridgeRequest request = BridgeRequest.Read((string)scenario["request"]);
        StringWriter log = new StringWriter(), err = new StringWriter();
        BridgeSession session = new BridgeSession(host, (string)scenario["link"], (string)scenario["master"], runtime, request,
            (string)scenario["result"], log, err);
        int code = session.Run();
        StringBuilder sb = new StringBuilder("{\"exitCode\":" + code + ",\"calls\":[" + string.Join(",", host.Calls.ToArray()) + "],\"closes\":[");
        for (int i = 0; i < host.Closes.Count; i++) { if (i > 0) { sb.Append(','); } sb.Append(JsonText.Quote(host.Closes[i])); }
        sb.Append("],\"remaining\":" + (replies.Count - host.next) + ",\"mismatch\":" + (host.Mismatch == null ? "null" : JsonText.Quote(host.Mismatch)));
        sb.Append(",\"log\":" + JsonText.Quote(log.ToString()) + ",\"err\":" + JsonText.Quote(err.ToString()) + "}");
        byte[] bytes = new UTF8Encoding(false).GetBytes(sb.ToString());
        using (Stream stdout = Console.OpenStandardOutput()) { stdout.Write(bytes, 0, bytes.Length); }
        return 0;
    }
}
`;

let buildDir;
let exe;
let harnessExe;

before(() => {
    if (skip) return;
    buildDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cat5dev-catiarun-'));
    exe = path.join(buildDir, 'catia-run.exe');
    execFileSync(process.execPath, [path.join(root, 'scripts', 'build-bridge.js'), '--out', exe], { windowsHide: true, stdio: 'pipe' });
    const harnessSource = path.join(buildDir, 'CatiaRunHarness.cs');
    fs.writeFileSync(harnessSource, HARNESS, 'utf8');
    harnessExe = path.join(buildDir, 'catia-run-harness.exe');
    execFileSync(csc, [...cscArguments(harnessExe, [...sourceFiles(), harnessSource]), '/main:CatiaRunHarness'], { windowsHide: true, stdio: 'pipe' });
});

after(() => { if (buildDir) fs.rmSync(buildDir, { recursive: true, force: true }); });

let counter = 0;
/** A fresh work folder with a dummy master (C5D_Bridge-0123abcd.catvba) and request.json; returns the paths. */
function workspace(t, request) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'cat5dev-catiarun-ws-'));
    t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
    const bridgeDir = path.join(dir, 'bridge');
    const work = path.join(dir, `step${counter++}`);
    fs.mkdirSync(bridgeDir);
    fs.mkdirSync(work);
    const master = path.join(bridgeDir, 'C5D_Bridge-0123abcd.catvba');
    fs.writeFileSync(master, 'dummy');
    const requestPath = path.join(work, 'request.json');
    if (request !== undefined) fs.writeFileSync(requestPath, typeof request === 'string' || Buffer.isBuffer(request) ? request : JSON.stringify(request));
    return { dir, bridgeDir, work, master, request: requestPath, result: path.join(work, 'result.json') };
}

function run(args) {
    const r = spawnSync(exe, args, { windowsHide: true, encoding: 'utf8', timeout: 20000 });
    return { status: r.status, stdout: r.stdout, stderr: r.stderr };
}

function validate(ws, extra = []) {
    return run(['--validate-only', '--bridge', ws.master, '--request', ws.request, '--result', ws.result, ...extra]);
}

const base = (kind, more = {}) => ({ version: 2, operationId: 'op-1', stepId: 'step-1', kind, targetProject: kind === 'list' ? '' : 'toQI', ...more });

test('build-bridge.js compiles CatiaRun*.cs only, with the runtime embedded', { skip }, () => {
    const names = sourceFiles().map(file => path.basename(file));
    assert.equal(names[0], 'CatiaRun.cs');
    assert.ok(names.length > 1 && names.every(name => /^CatiaRun(\..+)?\.cs$/.test(name)), names.join(', '));
    assert.ok(!names.includes('EvalProbe.cs') && !names.includes('TypeLibDump.cs'));
    const args = cscArguments('C:\\out\\x.exe');
    assert.ok(args.includes('/warnaserror+'));
    assert.ok(args.includes(`/resource:${RUNTIME_BAS},${RESOURCE_NAME}`));
    const image = fs.readFileSync(exe);
    assert.ok(image.includes(fs.readFileSync(RUNTIME_BAS)), 'the runtime text is embedded byte for byte');
    assert.ok(image.includes(Buffer.from(RESOURCE_NAME, 'utf8')), 'manifest resource name');
});

test('--validate-only accepts every kind and writes nothing', { skip }, t => {
    const requests = [
        base('list'),
        base('pull'),
        base('check', { modules: [], deletions: [] }),
        base('push', { modules: [{ name: 'Module1', type: 1, text: 'Option Explicit\r\n\' 日本語 🙂\n' }, { name: 'Class1', type: 2, text: '' },
            { name: 'Form1', type: 3, text: 'x' }], deletions: ['Old'] }),
        base('push'),
        base('pull', { encoding: 'shift_jis', extra: { any: [1, null, true] } }), // unknown fields are ignored
        base('push', { modules: [{ name: 'M', type: 1, text: '\ud800 lone surrogate' }] })
    ];
    for (const request of requests) {
        const ws = workspace(t, request);
        const r = validate(ws);
        assert.equal(r.status, 0, `${JSON.stringify(request)}\n${r.stderr}`);
        assert.match(r.stdout, new RegExp(`^validate ${request.kind} \\d+ms\\r?\\nexit 0 \\d+ms`));
        assert.deepEqual(fs.readdirSync(ws.work), ['request.json']);
    }
    const bom = workspace(t, Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(JSON.stringify(base('list')))]));
    assert.equal(validate(bom).status, 0, 'a UTF-8 BOM is accepted');
    const spaced = workspace(t, ` \r\n${JSON.stringify(base('pull'), null, 2)}\n`);
    assert.equal(validate(spaced).status, 0, 'JSON whitespace is accepted');
});

test('request.json that breaks the contract exits 2', { skip }, t => {
    const module = (more = {}) => ({ name: 'Module1', type: 1, text: '', ...more });
    const cases = {
        'not JSON': '{"version":2,',
        'trailing text': `${JSON.stringify(base('list'))} x`,
        'duplicate key': '{"version":2,"version":2,"operationId":"a","stepId":"b","kind":"list","targetProject":""}',
        'array root': '[]',
        'version 1': base('list', { version: 1 }),
        'version string': base('list', { version: '2' }),
        'no version': (() => { const r = base('list'); delete r.version; return r; })(),
        'empty operationId': base('list', { operationId: '' }),
        'numeric stepId': base('list', { stepId: 1 }),
        'bad kind': base('export'),
        'list with target': base('list', { targetProject: 'toQI' }),
        'pull without target': base('pull', { targetProject: '' }),
        'no targetProject': (() => { const r = base('pull'); delete r.targetProject; return r; })(),
        'modules for pull': base('pull', { modules: [module()] }),
        'deletions for check': base('check', { deletions: ['A'] }),
        'modules not an array': base('push', { modules: {} }),
        'module not an object': base('push', { modules: ['Module1'] }),
        'empty name': base('push', { modules: [module({ name: '' })] }),
        'name with colon': base('push', { modules: [module({ name: 'a:b' })] }),
        'name with control char': base('push', { modules: [module({ name: 'a\tb' })] }),
        'deletion with slash': base('push', { deletions: ['a/b'] }),
        'deletion not a string': base('push', { deletions: [1] }),
        'type 0': base('push', { modules: [module({ type: 0 })] }),
        'type 100': base('push', { modules: [module({ type: 100 })] }),
        'type 1.5': base('push', { modules: [module({ type: 1.5 })] }),
        'type string': base('push', { modules: [module({ type: '1' })] }),
        'text missing': base('push', { modules: [{ name: 'M', type: 1 }] }),
        'text null': base('push', { modules: [module({ text: null })] }),
        'duplicate module': base('push', { modules: [module({ name: 'Module1' }), module({ name: 'MODULE1' })] }),
        'module also deleted': base('push', { modules: [module({ name: 'Module1' })], deletions: ['module1'] }),
        'duplicate deletion': base('push', { deletions: ['A', 'a'] }),
        'invalid UTF-8': Buffer.from([0x7b, 0x22, 0xff, 0x22, 0x7d]),
        'raw control char in string': '{"version":2,"operationId":"a\u0001","stepId":"b","kind":"list","targetProject":""}',
        'bad escape': '{"version":2,"operationId":"\\x","stepId":"b","kind":"list","targetProject":""}',
        'leading zero': '{"version":02,"operationId":"a","stepId":"b","kind":"list","targetProject":""}',
        'deep nesting': `{"version":2,"operationId":"a","stepId":"b","kind":"list","targetProject":"","x":${'['.repeat(100)}${']'.repeat(100)}}`
    };
    for (const [label, request] of Object.entries(cases)) {
        const ws = workspace(t, request);
        const r = validate(ws);
        assert.equal(r.status, 2, `${label}: ${r.stdout}${r.stderr}`);
        assert.match(r.stderr, /^invalid arguments or request: /, label);
        assert.deepEqual(fs.readdirSync(ws.work), ['request.json'], label);
    }
    const large = workspace(t, '');
    fs.truncateSync(large.request, 64 * 1024 * 1024 + 1);
    const r = validate(large);
    assert.equal(r.status, 2);
    assert.match(r.stderr, /larger than 64 MiB/);
});

test('bad command lines exit 2 before CATIA is touched', { skip }, t => {
    const ws = workspace(t, base('list'));
    const all = ['--bridge', ws.master, '--request', ws.request, '--result', ws.result];
    // Without --validate-only only for arguments that cannot pass the checks (so the exe never reaches GetActiveObject).
    assert.equal(run([]).status, 2);
    assert.equal(run(['--help']).status, 2);
    assert.equal(run(['--bridge']).status, 2);
    const cases = {
        'missing --result': ['--bridge', ws.master, '--request', ws.request],
        'missing --bridge': ['--request', ws.request, '--result', ws.result],
        'duplicate --request': [...all, '--request', ws.request],
        'duplicate --validate-only': [...all, '--validate-only'],
        'unknown option': [...all, '--runtime', ws.master],
        'positional argument': [...all, 'extra'],
        'old usage': [ws.work, 'operation.catvbs'],
        'relative path': ['--bridge', ws.master, '--request', 'request.json', '--result', ws.result],
        'forward slashes': ['--bridge', ws.master.replace(/\\/g, '/'), '--request', ws.request, '--result', ws.result],
        'dot segment': ['--bridge', `${ws.bridgeDir}\\x\\..\\C5D_Bridge-0123abcd.catvba`, '--request', ws.request, '--result', ws.result],
        'trailing space': ['--bridge', ws.master, '--request', ws.request, '--result', `${ws.result} `],
        'different folders': ['--bridge', ws.master, '--request', ws.request, '--result', path.join(ws.dir, 'result.json')],
        'bridge missing': ['--bridge', path.join(ws.bridgeDir, 'C5D_Bridge-89abcdef.catvba'), '--request', ws.request, '--result', ws.result],
        'bridge upper-case sha': ['--bridge', path.join(ws.bridgeDir, 'C5D_Bridge-0123ABCD.catvba'), '--request', ws.request, '--result', ws.result],
        'bridge without sha': ['--bridge', path.join(ws.bridgeDir, 'C5D_Bridge.catvba'), '--request', ws.request, '--result', ws.result],
        'bridge is a link name': ['--bridge', path.join(ws.bridgeDir, 'C5D_Bridge-0123abcd-1-20261121151114123.catvba'), '--request', ws.request, '--result', ws.result],
        'request missing': ['--bridge', ws.master, '--request', path.join(ws.work, 'none.json'), '--result', ws.result]
    };
    fs.writeFileSync(path.join(ws.bridgeDir, 'C5D_Bridge.catvba'), 'x');
    fs.writeFileSync(path.join(ws.bridgeDir, 'C5D_Bridge-0123abcd-1-20261121151114123.catvba'), 'x');
    for (const [label, args] of Object.entries(cases)) {
        const r = run(['--validate-only', ...args]);
        assert.equal(r.status, 2, `${label}: ${r.stdout}${r.stderr}`);
    }
    fs.writeFileSync(ws.result, '{}');
    const exists = run(['--validate-only', ...all]);
    assert.equal(exists.status, 2);
    assert.match(exists.stderr, /--result already exists/);
    assert.equal(fs.readFileSync(ws.result, 'utf8'), '{}', 'an existing result is never overwritten');
    assert.equal(run(all).status, 2, 'also without --validate-only');
    assert.equal(fs.readFileSync(ws.result, 'utf8'), '{}');
});

// ---- bridge procedure against the fake host

const linkName = 'C5D_Bridge-0123abcd-0-20261121151114123.catvba';
const selfReply = (link, more = {}) => ({ module: PUSH, function: 'C5D_Self', library: link, self: { name: 'C5D_Bridge', displayName: link, fileName: link, ...more } });
const listText = (link, extra = []) => ['ok', 'Develop\tC:\\vba\\Develop.catvba\t0', '\t\t0', ...extra, `C5D_Bridge\t${link}\t1`].join('\n');
/** load, who (runtime present), list (only self), runtime version: the usual preparation. */
const prepare = (link, { who, list } = {}) => [
    { module: PUSH, function: 'C5D_Version', library: link, string: 'c5d-bridge:2' },
    { module: PUSH, function: 'C5D_Who', library: link, string: who ?? `ok\nC5D_Bridge\t${link}\t${link}\nC5D_Bridge_Push\t1\nC5D_Bridge_Runtime\t1` },
    { module: PUSH, function: 'C5D_List', library: link, string: list ?? listText(link) },
    { module: RT, function: 'C5D_RT_Version', library: link, string: RT_VERSION }
];
const projectsReply = text => ({ module: RT, function: 'C5D_RT_ListProjects', string: text });

function session(t, request, replies) {
    const ws = workspace(t, request);
    const link = path.join(ws.bridgeDir, linkName);
    const list = typeof replies === 'function' ? replies(link, ws) : replies;
    const scenario = path.join(ws.dir, 'scenario.json');
    fs.writeFileSync(scenario, JSON.stringify({ link, master: ws.master, request: ws.request, result: ws.result, replies: list }));
    const out = JSON.parse(execFileSync(harnessExe, [scenario], { windowsHide: true, encoding: 'utf8', timeout: 20000 }));
    assert.equal(out.mismatch, null, out.err);
    assert.equal(out.remaining, 0, `unused replies\n${out.err}`);
    const text = fs.existsSync(ws.result) ? fs.readFileSync(ws.result, 'utf8') : null;
    const result = text === null ? null : JSON.parse(text);
    if (text !== null && typeof validateCatiaResult === 'function') validateCatiaResult(text, request);
    return { ...out, ws, link, result, files: fs.readdirSync(ws.work).sort() };
}

const RESULT_KEYS = ['version', 'operationId', 'stepId', 'kind', 'targetProject', 'expectedCount', 'succeededCount', 'failures', 'items',
    'projects', 'updatedCount', 'deletedCount'];

test('pull: installs the runtime, exports type 1-3 modules as UTF-8 m<i>.txt and closes itself', { skip }, t => {
    const body = 'Option Explicit\r\n\' 日本語 🙂 "q" \\ tab\there\r\nSub A()\r\nEnd Sub';
    const s = session(t, base('pull'), link => [
        { module: PUSH, function: 'C5D_Version', library: link, string: 'c5d-bridge:2' },
        { module: PUSH, function: 'C5D_Who', string: `ok\nC5D_Bridge\t${link}\t${link}\nC5D_Bridge_Push\t1` },
        { module: PUSH, function: 'C5D_List', string: listText(link) },
        { module: PUSH, function: 'C5D_Install', string: 'ok' },
        { module: RT, function: 'C5D_RT_Version', string: RT_VERSION },
        projectsReply('ok\nDevelop\ntoQI'),
        { module: RT, function: 'C5D_RT_ListModules', string: 'ok\nModule1\t1\nThisDocument\t100\nClass1\t2\nForm1\t3' },
        { module: RT, function: 'C5D_RT_ReadModule', string: `ok:${body}` },
        { module: RT, function: 'C5D_RT_ReadModule', string: 'error:export.lines:9:Subscript out of range\r\nline 2' },
        { module: RT, function: 'C5D_RT_ReadModule', string: 'ok:' },
        selfReply(link)
    ]);
    assert.equal(s.exitCode, 0, s.err);
    const install = s.calls.find(call => call.function === 'C5D_Install');
    assert.deepEqual(install.args, [RT, RUNTIME_TEXT], 'the embedded runtime text is installed unchanged');
    assert.ok(s.calls.every(call => call.library === s.link), 'every call uses the link path');
    assert.deepEqual(s.calls.filter(call => call.function === 'C5D_RT_ReadModule').map(call => call.args), [['toQI', 'Module1'], ['toQI', 'Class1'], ['toQI', 'Form1']]);
    assert.deepEqual(Object.keys(s.result), RESULT_KEYS);
    assert.deepEqual(s.result, { version: 1, operationId: 'op-1', stepId: 'step-1', kind: 'pull', targetProject: 'toQI', expectedCount: 3, succeededCount: 2,
        failures: [{ phase: 'export.lines', target: 'Class1', code: 9, message: 'Subscript out of range\r\nline 2' }],
        items: [{ name: 'Module1', type: 1, file: 'm0.txt' }, { name: 'Form1', type: 3, file: 'm2.txt' }], projects: [], updatedCount: 0, deletedCount: 0 });
    assert.deepEqual(s.files, ['m0.txt', 'm2.txt', 'request.json', 'result.json'], 'no result.json.tmp is left');
    assert.deepEqual(fs.readFileSync(path.join(s.ws.work, 'm0.txt')), Buffer.from(body, 'utf8'), 'UTF-8 without BOM');
    assert.equal(fs.statSync(path.join(s.ws.work, 'm2.txt')).size, 0);
    assert.deepEqual(s.closes, [s.link]);
    for (const stage of ['load', 'who exact', 'list 0 wreck\\(s\\) closed', `runtime installed ${RT_VERSION}`, 'projects 2', 'pull 2/3', 'result', 'close ok']) {
        assert.match(s.log, new RegExp(`^${stage} \\d+ms\\r?$`, 'm'), stage);
    }
});

test('check: a project failure stops the export; result is still written with exit 0', { skip }, t => {
    const s = session(t, base('check'), link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: RT, function: 'C5D_RT_ListModules', string: 'ok\nModule1\t1\nModule2\t1' },
        { module: RT, function: 'C5D_RT_ReadModule', string: 'error:project.lookup:5:Target project not found' },
        selfReply(link)]);
    assert.equal(s.exitCode, 0);
    assert.deepEqual(s.result.failures, [{ phase: 'project.lookup', target: 'toQI', code: 5, message: 'Target project not found' }]);
    assert.equal(s.result.expectedCount, 2);
    assert.equal(s.result.succeededCount, 0);
    assert.match(s.log, /^runtime ok c5d-runtime:\d+ \d+ms/m);
});

test('push: writes, then deletes; per-module failures continue, connect./project. failures stop', { skip }, t => {
    const request = base('push', { modules: [{ name: 'Module1', type: 1, text: 'A\n🙂' }, { name: 'Class1', type: 2, text: '' }],
        deletions: ['Old1', 'Old2'] });
    const s = session(t, request, link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: RT, function: 'C5D_RT_WriteModule', string: 'ok' },
        { module: RT, function: 'C5D_RT_WriteModule', string: 'error:push.type:5:Module type does not match' },
        { module: RT, function: 'C5D_RT_DeleteModule', string: 'error:connect.apc:429:No APC object' },
        selfReply(link)]);
    assert.equal(s.exitCode, 0);
    assert.deepEqual(s.calls.filter(call => call.module === RT && call.function !== 'C5D_RT_Version' && call.function !== 'C5D_RT_ListProjects').map(call => call.args),
        [['toQI', 'Module1', '1', 'A\n🙂'], ['toQI', 'Class1', '2', ''], ['toQI', 'Old1']]);
    assert.deepEqual(s.result, { version: 1, operationId: 'op-1', stepId: 'step-1', kind: 'push', targetProject: 'toQI', expectedCount: 4, succeededCount: 1,
        failures: [{ phase: 'push.type', target: 'Class1', code: 5, message: 'Module type does not match' },
            { phase: 'connect.apc', target: 'toQI', code: 429, message: 'No APC object' }],
        items: [], projects: [], updatedCount: 1, deletedCount: 0 });
});

test('push: an ExecuteScript exception or a non-string return stops with exit 12 and a *.call failure', { skip }, t => {
    const request = base('push', { modules: [{ name: 'Module1', type: 1, text: 'x' }, { name: 'Module2', type: 1, text: 'y' }], deletions: ['Old'] });
    const thrown = session(t, request, link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: RT, function: 'C5D_RT_WriteModule', string: 'ok' },
        { module: RT, function: 'C5D_RT_WriteModule', throw: { hresult: -2147418111, message: 'Call was rejected by callee.' } },
        selfReply(link)]);
    assert.equal(thrown.exitCode, 12);
    assert.deepEqual(thrown.result.failures, [{ phase: 'push.call', target: 'Module2', code: -2147418111, message: 'Call was rejected by callee.' }]);
    assert.equal(thrown.result.expectedCount, 3, 'expectedCount stays modules + deletions');
    assert.equal(thrown.result.updatedCount, 1);
    assert.match(thrown.err, /C5D_RT_WriteModule: 0x80010001: Call was rejected by callee\./);
    assert.deepEqual(thrown.closes, [thrown.link]);

    const odd = session(t, request, link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: RT, function: 'C5D_RT_WriteModule', int: 5 }, selfReply(link)]);
    assert.equal(odd.exitCode, 12);
    assert.deepEqual(odd.result.failures, [{ phase: 'push.call', target: 'Module1', code: 13, message: 'Unexpected return value: Int32' }]);

    const text = session(t, request, link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: RT, function: 'C5D_RT_WriteModule', string: 'done' }, selfReply(link)]);
    assert.equal(text.exitCode, 12);
    assert.equal(text.result.failures[0].code, 13);
    assert.match(text.result.failures[0].message, /^Unexpected return value: done$/);
});

test('pull: an exception while reading is export.call (exit 12)', { skip }, t => {
    const s = session(t, base('pull'), link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: RT, function: 'C5D_RT_ListModules', string: 'ok\nModule1\t1\nModule2\t1' },
        { module: RT, function: 'C5D_RT_ReadModule', throw: { hresult: -2147467259, message: 'boom' } },
        selfReply(link)]);
    assert.equal(s.exitCode, 12);
    assert.deepEqual(s.result.failures, [{ phase: 'export.call', target: 'Module1', code: -2147467259, message: 'boom' }]);
});

test('list: projects as returned (escapes undone), C5D_Bridge excluded by the runtime', { skip }, t => {
    const s = session(t, base('list'), link => [...prepare(link), projectsReply('ok\nDevelop\nA%09B%25\ntoQI\ntoQI'), selfReply(link)]);
    assert.equal(s.exitCode, 0);
    assert.deepEqual(s.result.projects, ['Develop', 'A\tB%', 'toQI', 'toQI']);
    assert.equal(s.result.expectedCount, 4);
    assert.equal(s.result.succeededCount, 4);
});

test('list: an error from C5D_RT_ListProjects is one failure (exit 0); an exception is list.call (exit 12)', { skip }, t => {
    const error = session(t, base('pull'), link => [...prepare(link), projectsReply('error:list.count:91:Object variable not set'), selfReply(link)]);
    assert.equal(error.exitCode, 0);
    assert.deepEqual(error.result.failures, [{ phase: 'list.count', target: 'toQI', code: 91, message: 'Object variable not set' }]);
    assert.equal(error.calls.filter(call => call.function === 'C5D_RT_ListModules').length, 0);
    const thrown = session(t, base('list'), link => [...prepare(link),
        { module: RT, function: 'C5D_RT_ListProjects', throw: { hresult: -2147467259, message: 'x' } }, selfReply(link)]);
    assert.equal(thrown.exitCode, 12);
    assert.deepEqual(thrown.result.failures, [{ phase: 'list.call', target: '', code: -2147467259, message: 'x' }]);
});

test('no project in the VBA editor exits 14 without result.json, after closing itself', { skip }, t => {
    const s = session(t, base('pull'), link => [...prepare(link), projectsReply('ok'), selfReply(link)]);
    assert.equal(s.exitCode, 14);
    assert.equal(s.result, null);
    assert.deepEqual(s.closes, [s.link]);
});

test('bridge preparation failures exit 13 and close only an identified bridge', { skip }, t => {
    const v1 = session(t, base('list'), link => [{ module: PUSH, function: 'C5D_Version', library: link, string: 'c5d-bridge:1' }]);
    assert.equal(v1.exitCode, 13);
    assert.match(v1.err, /bridge v1 is not supported; restart CATIA/);
    assert.deepEqual(v1.closes, []);

    const loadFails = session(t, base('list'), link => [{ module: PUSH, function: 'C5D_Version', throw: { hresult: -2147467259, message: 'not found' } }]);
    assert.equal(loadFails.exitCode, 13);
    assert.match(loadFails.err, /load: C5D_Bridge_Push\.C5D_Version failed: 0x80004005: not found/);

    const other = session(t, base('list'), link => [{ module: PUSH, function: 'C5D_Version', string: 'c5d-bridge:2' },
        { module: PUSH, function: 'C5D_Who', string: `ok\nC5D_Bridge\t${link.replace(linkName, 'C5D_Bridge-0123abcd.catvba')}\t${link}\nC5D_Bridge_Push\t1` }]);
    assert.equal(other.exitCode, 13, 'executing project is another path: not identified, not closed');
    assert.equal(other.calls.length, 2);

    const caseOnly = session(t, base('list'), link => prepare(link, { who: `ok\nC5D_Bridge\t${link.toUpperCase()}\t${link}\nC5D_Bridge_Push\t1\nC5D_Bridge_Runtime\t1`,
        list: listText(link.toLowerCase()) }).concat([projectsReply('ok\ntoQI'), selfReply(link)]));
    assert.equal(caseOnly.exitCode, 0, 'paths are compared ignoring case');
    assert.match(caseOnly.log, /^who ignore-case/m);

    const twoSelves = session(t, base('list'), link => [...prepare(link, { list: listText(link, [`C5D_Bridge\t${link}\t0`]) }).slice(0, 3), selfReply(link)]);
    assert.equal(twoSelves.exitCode, 13, 'another row with our own path');
    assert.deepEqual(twoSelves.closes, [twoSelves.link]);

    const installFails = session(t, base('list'), link => [...prepare(link).slice(0, 3),
        { module: RT, function: 'C5D_RT_Version', string: 'c5d-runtime:0' },
        { module: PUSH, function: 'C5D_Install', string: 'error:install.type:5:Existing module is not a standard module' },
        selfReply(link)]);
    assert.equal(installFails.exitCode, 13);
    assert.match(installFails.err, /C5D_Install failed: error:install\.type/);
    assert.deepEqual(installFails.closes, [installFails.link]);

    const wrongVersion = session(t, base('list'), link => [...prepare(link).slice(0, 2),
        { module: PUSH, function: 'C5D_List', string: listText(link) },
        { module: RT, function: 'C5D_RT_Version', throw: { hresult: -2147467259, message: 'compile error' } },
        { module: PUSH, function: 'C5D_Install', string: 'ok' },
        { module: RT, function: 'C5D_RT_Version', string: 'c5d-runtime:0' },
        selfReply(link)]);
    assert.equal(wrongVersion.exitCode, 13, 'the version after install must be the embedded one');
});

test('wrecks in the master folder are closed through their own path before anything is written', { skip }, t => {
    const s = session(t, base('list'), (link, ws) => {
        const wreck = path.join(ws.bridgeDir, 'C5D_Bridge-0123abcd-0-20261121150000000.catvba');
        const old = path.join(ws.bridgeDir, 'c5d_bridge-89abcdef.CATVBA');
        return [
            { module: PUSH, function: 'C5D_Version', library: link, string: 'c5d-bridge:2' },
            { module: PUSH, function: 'C5D_Who', string: `ok\nC5D_Bridge\t${link}\t${link}\nC5D_Bridge_Push\t1\nC5D_Bridge_Runtime\t1` },
            { module: PUSH, function: 'C5D_List', string: listText(link, [`C5D_Bridge\t${wreck}\t0`, `C5D_Bridge\t${old}\t0`]) },
            { module: PUSH, function: 'C5D_Self', library: wreck, self: { name: 'C5D_Bridge', displayName: wreck, fileName: wreck } },
            { module: PUSH, function: 'C5D_Self', library: old, self: { name: 'C5D_Bridge', displayName: old, fileName: old } },
            { module: PUSH, function: 'C5D_List', library: link, string: listText(link) },
            { module: RT, function: 'C5D_RT_Version', string: RT_VERSION },
            projectsReply('ok\ntoQI'),
            selfReply(link)
        ];
    });
    assert.equal(s.exitCode, 0, s.err);
    assert.equal(s.closes.length, 3);
    assert.equal(s.closes[2], s.link);
    assert.match(s.log, /^list 2 wreck\(s\) closed/m);
});

test('wrecks that cannot be closed safely exit 13 without touching them', { skip }, t => {
    const outside = session(t, base('list'), (link, ws) => [...prepare(link, { list: listText(link, [`C5D_Bridge\t${path.join(ws.dir, 'C5D_Bridge.catvba')}\t0`]) }).slice(0, 3),
        selfReply(link)]);
    assert.equal(outside.exitCode, 13);
    assert.match(outside.err, /cannot close a C5D_Bridge that is not ours/);
    assert.deepEqual(outside.closes, [outside.link], 'only our own bridge is closed');

    const unnamed = session(t, base('list'), link => [...prepare(link, { list: listText(link, ['C5D_Bridge\t\t0']) }).slice(0, 3), selfReply(link)]);
    assert.equal(unnamed.exitCode, 13, 'a C5D_Bridge without DisplayName');

    const many = session(t, base('list'), (link, ws) => [...prepare(link, {
        list: listText(link, [1, 2, 3, 4, 5, 6].map(i => `C5D_Bridge\t${path.join(ws.bridgeDir, `C5D_Bridge-0123abcd-0-2026112115000000${i}.catvba`)}\t0`))
    }).slice(0, 3), selfReply(link)]);
    assert.equal(many.exitCode, 13, 'more than 5 wrecks');
    assert.match(many.err, /too many C5D_Bridge projects/);

    const v1wreck = session(t, base('list'), (link, ws) => {
        const wreck = path.join(ws.bridgeDir, 'C5D_Bridge-0123abcd-0-20261121150000000.catvba');
        return [...prepare(link, { list: listText(link, [`C5D_Bridge\t${wreck}\t0`]) }).slice(0, 3),
            { module: PUSH, function: 'C5D_Self', library: wreck, throw: { hresult: -2147467259, message: 'macro not found' } }, selfReply(link)];
    });
    assert.equal(v1wreck.exitCode, 13, 'a v1 wreck has no C5D_Self');

    const guard = session(t, base('list'), (link, ws) => {
        const wreck = path.join(ws.bridgeDir, 'C5D_Bridge-0123abcd-0-20261121150000000.catvba');
        return [...prepare(link, { list: listText(link, [`C5D_Bridge\t${wreck}\t0`]) }).slice(0, 3),
            { module: PUSH, function: 'C5D_Self', library: wreck, self: { name: 'C5D_Bridge', displayName: link, fileName: link } }, selfReply(link)];
    });
    assert.equal(guard.exitCode, 13, 'C5D_Self of the wreck path returned another project');
    assert.deepEqual(guard.closes, [guard.link]);

    const stays = session(t, base('list'), (link, ws) => {
        const wreck = path.join(ws.bridgeDir, 'C5D_Bridge-0123abcd-0-20261121150000000.catvba');
        return [...prepare(link, { list: listText(link, [`C5D_Bridge\t${wreck}\t0`]) }).slice(0, 3),
            { module: PUSH, function: 'C5D_Self', library: wreck, self: { name: 'C5D_Bridge', displayName: wreck, fileName: wreck } },
            { module: PUSH, function: 'C5D_List', string: listText(link, [`C5D_Bridge\t${wreck}\t0`]) }, selfReply(link)];
    });
    assert.equal(stays.exitCode, 13, 'the wreck is still listed after Close');
});

test('the final Close is guarded; a failed or skipped Close does not change the exit code', { skip }, t => {
    const mismatch = session(t, base('list'), link => [...prepare(link), projectsReply('ok\ntoQI'), selfReply(link, { fileName: link.replace(linkName, 'other.catvba') })]);
    assert.equal(mismatch.exitCode, 0);
    assert.deepEqual(mismatch.closes, []);
    assert.match(mismatch.err, /^warning: bridge close failed: guard mismatch/m);
    assert.ok(mismatch.result, 'result.json was written before the close');

    const name = session(t, base('list'), link => [...prepare(link), projectsReply('ok\ntoQI'), selfReply(link, { name: 'c5d_bridge' })]);
    assert.deepEqual(name.closes, [], 'Name is compared ordinally');

    const throws = session(t, base('list'), link => [...prepare(link), projectsReply('ok\ntoQI'), selfReply(link, { closeError: 'busy' })]);
    assert.equal(throws.exitCode, 0);
    assert.match(throws.err, /^warning: bridge close failed: 0x80004005: busy/m);
    assert.match(throws.log, /^close failed \d+ms/m);

    const text = session(t, base('list'), link => [...prepare(link), projectsReply('ok\ntoQI'),
        { module: PUSH, function: 'C5D_Self', string: 'error:self.exec:0:No executing project' }]);
    assert.equal(text.exitCode, 0);
    assert.match(text.err, /^warning: bridge close failed: C5D_Self returned error:self\.exec/m);
});
