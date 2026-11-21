// Explicit integration test. Existing projects are mutated only with --target <project>,
// and then only through the dedicated C5D_BridgeTest module (never saved).
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const { createHash } = require('node:crypto');
const iconv = require('iconv-lite');
const { SyncOperationManager } = require('../out/syncOperation');
const { createCatiaRequest, runCatiaRequest, bridgeExecutor } = require('../out/catiaExecution');
const { vbString } = require('../out/catiaScripts');
const { normalizeModuleText } = require('../out/catiaTypes');
const { captureModules, applyPull } = require('../out/pullTransaction');
const reportPath = path.resolve(__dirname, '../temp/catia-smoke-result.json');
const targetIndex = process.argv.indexOf('--target');
const targetProject = targetIndex >= 0 ? process.argv[targetIndex + 1] : undefined;
const report = { passed: false, disposableRequested: process.argv.includes('--disposable'), targetMode: targetIndex >= 0, steps: [] };
const manager = new SyncOperationManager();
const context = manager.acquire('integration-test');
const projectName = 'C5D_Test_' + context.id.replace(/-/g, '').slice(0, 16);
const marker = 'Cat5Dev-test-owner:' + context.id;
let created = false;
const testModule = 'C5D_BridgeTest';
let testModuleOwned = false;
// Built by npm run compile:bridge.
const bridgePath = path.resolve(__dirname, '../bin/catia-run.exe');
function executeOverride(transform) {
    return (directory, scriptName, timeoutMs) => {
        const source = path.join(directory, scriptName);
        fs.writeFileSync(source, transform(fs.readFileSync(source, 'utf8')));
        return bridgeExecutor(bridgePath)(directory, scriptName, timeoutMs);
    };
}
async function step(kind, project = '', configure, options) {
    const request = createCatiaRequest(context, kind, project, 'shift_jis');
    if (configure) configure(request);
    const result = await runCatiaRequest(context, request, { bridgePath, ...options });
    report.steps.push({ kind, project, expected: result.expectedCount, succeeded: result.succeededCount });
    return { request, result };
}
async function snapshot(projects) {
    const values = {};
    for (const project of projects) {
        const { request, result } = await step('check', project);
        values[project] = result.items.map(item => ({ name: item.name, type: item.type,
            hash: createHash('sha256').update(fs.readFileSync(path.join(request.directory, item.file))).digest('hex') }))
            .sort((a, b) => a.name.localeCompare(b.name));
    }
    return values;
}
// Identifiers are unique so VBE case normalization cannot touch other modules.
const testBodies = [
    [`' ${marker}`, "' 日本語テスト：漢字かなカナ", 'Sub C5D_BridgeTest_Main()', '    Dim c5dBridgeText As String',
        '    c5dBridgeText = "初回の本文です"', '    Debug.Print c5dBridgeText', 'End Sub'],
    [`' ${marker}`, "' 日本語テスト：更新後の本文", 'Sub C5D_BridgeTest_Main()', '    Debug.Print C5D_BridgeTest_Value("更新")',
        'End Sub', '', 'Function C5D_BridgeTest_Value(c5dBridgeArg)', '    C5D_BridgeTest_Value = c5dBridgeArg & "：漢字かなカナ"', 'End Function']
].map(lines => lines.join('\r\n'));
const isTestModule = item => item.name.toLowerCase() === testModule.toLowerCase();
function pushTestModule(request, text) {
    fs.writeFileSync(path.join(request.directory, 'm0.txt'), iconv.encode(text, request.encoding));
    request.modules = [{ name: testModule, type: 1, file: 'm0.txt' }];
}
async function findTestModule(kind) {
    const { request, result } = await step(kind, targetProject);
    const item = result.items.find(isTestModule);
    return item && { type: item.type, text: iconv.decode(fs.readFileSync(path.join(request.directory, item.file)), request.encoding) };
}
async function verifyTestModule(kind, text) {
    const found = await findTestModule(kind);
    assert.ok(found, `${testModule} missing after push`);
    assert.equal(found.type, 1);
    assert.equal(normalizeModuleText(found.text), normalizeModuleText(text));
}
async function deleteTestModule() {
    await step('push', targetProject, request => { request.modules = []; request.deletions = [testModule]; });
    report.testModuleRemoved = !(await findTestModule('check'));
    if (report.testModuleRemoved) testModuleOwned = false;
}
async function targetSmoke(projects) {
    assert.ok(targetProject && !targetProject.startsWith('--'), '--target requires a project name');
    assert.ok(!report.disposableRequested, '--target and --disposable cannot be combined');
    report.targetProject = targetProject;
    assert.ok(projects.includes(targetProject), `Project not found: ${targetProject}`);
    const baseline = (await snapshot([targetProject]))[targetProject];
    report.moduleCount = baseline.length;
    assert.ok(!baseline.some(isTestModule), `${testModule} already exists in ${targetProject}; refusing to touch it`);
    testModuleOwned = true;
    await step('push', targetProject, request => pushTestModule(request, testBodies[0]));
    await verifyTestModule('pull', testBodies[0]);
    await step('push', targetProject, request => pushTestModule(request, testBodies[1]));
    await verifyTestModule('check', testBodies[1]);
    await deleteTestModule();
    assert.ok(report.testModuleRemoved, `${testModule} still present after deletion`);
    assert.deepEqual((await snapshot([targetProject]))[targetProject], baseline);
    report.existingModulesUnchanged = true;
    report.passed = true;
}
const createProject = `
Sub CreateReviewProject()
    On Error Resume Next
    Dim p, ownership, errN, errText
    gExpected = 1
    Err.Clear
    Set p = gProjects.Add(1)
    If Err.Number <> 0 Then
        errN = Err.Number
        errText = Err.Description
        Err.Clear
        AddFailure "test.create", ${vbString(projectName)}, errN, errText
        Exit Sub
    End If
    Err.Clear
    p.Name = ${vbString(projectName)}
    If Err.Number <> 0 Then
        errN = Err.Number
        errText = Err.Description
        Err.Clear
        gProjects.Remove p
        AddFailure "test.name", ${vbString(projectName)}, errN, errText
        Exit Sub
    End If
    Err.Clear
    Set ownership = p.VBComponents.Add(1)
    ownership.Name = "C5D_Ownership"
    ownership.CodeModule.AddFromString ${vbString("' " + marker)}
    If Err.Number <> 0 Then
        errN = Err.Number
        errText = Err.Description
        Err.Clear
        gProjects.Remove p
        AddFailure "test.marker", ${vbString(projectName)}, errN, errText
        Exit Sub
    End If
    AppendJson gProjectNames, JsonQuote(p.Name)
    gSucceeded = 1
End Sub
`;
const removeProject = `
Sub RemoveReviewProject()
    On Error Resume Next
    Dim p, component, source, errN, errText
    gExpected = 1
    Set p = FindProject()
    If p Is Nothing Then Exit Sub
    Err.Clear
    Set component = p.VBComponents.Item("C5D_Ownership")
    source = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)
    If Err.Number <> 0 Or InStr(source, ${vbString(marker)}) = 0 Then
        Err.Clear
        AddFailure "test.ownership", gTargetProject, 5, "Ownership marker missing; refusing removal"
        Exit Sub
    End If
    Err.Clear
    gProjects.Remove p
    If Err.Number <> 0 Then
        errN = Err.Number
        errText = Err.Description
        Err.Clear
        AddFailure "test.remove", gTargetProject, errN, errText
        Exit Sub
    End If
    AppendJson gProjectNames, JsonQuote(gTargetProject)
    gSucceeded = 1
End Sub
`;
(async () => {
    const initial = await step('list'); report.projectCount = initial.result.projects.length;
    if (report.targetMode) { await targetSmoke(initial.result.projects); return; }
    if (!report.disposableRequested) { report.passed = true; return; }
    const baseline = await snapshot(initial.result.projects);
    assert.ok(!initial.result.projects.includes(projectName));
    const made = await step('list', '', undefined, { execute: executeOverride(source => source.replace(/\r\n    ListProjects\r\n/, '\r\n    CreateReviewProject\r\n') + createProject) });
    assert.deepEqual(made.result.projects, [projectName]); created = true;
    await step('push', projectName, request => {
        const modules = [
            { name: 'C5D_Ownership', type: 1, text: "' " + marker },
            { name: 'C5D_UserModule', type: 1, text: 'Sub Test()\r\nDebug.Print 1\r\nEnd Sub' },
            { name: 'ReviewClass', type: 2, text: '' }
        ];
        request.modules = modules.map((m, i) => { const file = `m${i}.txt`; fs.writeFileSync(path.join(request.directory, file), m.text); return { name: m.name, type: m.type, file }; });
    });
    const pulled = await step('pull', projectName);
    assert.ok(pulled.result.items.some(item => item.name === 'C5D_UserModule'));
    assert.ok(pulled.result.items.some(item => item.name === 'ReviewClass' && item.type === 2));
    const local = fs.mkdtempSync(path.resolve(__dirname, '../temp/catia-pull-smoke-'));
    fs.mkdirSync(path.join(local, 'modules')); fs.writeFileSync(path.join(local, 'modules', 'LocalOnly.bas_utf'), 'keep');
    const incoming = new Map(pulled.result.items.map(item => [item.name + ({ 1: '.bas_utf', 2: '.cls_utf', 3: '.frm_utf' })[item.type], fs.readFileSync(path.join(pulled.request.directory, item.file), 'utf8')]));
    applyPull(local, captureModules(local), incoming, false); assert.ok(fs.existsSync(path.join(local, 'modules', 'LocalOnly.bas_utf')));
    applyPull(local, captureModules(local), incoming, true); assert.ok(!fs.existsSync(path.join(local, 'modules', 'LocalOnly.bas_utf')));
    await step('push', projectName, request => { request.modules = []; request.deletions = ['C5D_UserModule']; });
    const afterDeletion = await step('check', projectName); assert.ok(!afterDeletion.result.items.some(item => item.name === 'C5D_UserModule'));
    const after = await snapshot(initial.result.projects); assert.deepEqual(after, baseline);
    report.existingProjectsUnchanged = true;
    report.passed = true;
})().catch(error => {
    report.error = String(error); if (error.result) report.failureResult = error.result;
    process.exitCode = 1;
}).finally(async () => {
    if (created) {
        try {
            await step('list', projectName, undefined, { execute: executeOverride(source => source.replace(/\r\n    ListProjects\r\n/, '\r\n    RemoveReviewProject\r\n') + removeProject) });
            report.disposableRemoved = true;
        } catch (error) { report.disposableRemoved = false; report.cleanupError = String(error); report.disposableProject = projectName; report.passed = false; process.exitCode = 1; }
    }
    if (testModuleOwned) {
        // Only reached when this run created the module (it was absent beforehand).
        try {
            if (await findTestModule('check')) await deleteTestModule(); else { report.testModuleRemoved = true; testModuleOwned = false; }
        } catch (error) { report.testModuleRemoved = false; report.cleanupError = String(error); }
        if (testModuleOwned) { report.leftoverModule = `${targetProject}/${testModule}`; report.passed = false; process.exitCode = 1; }
    }
    try { context.close(); } catch (error) { report.closeError = String(error); }
    fs.writeFileSync(reportPath, JSON.stringify(report, null, 2)); console.log(JSON.stringify(report));
});
