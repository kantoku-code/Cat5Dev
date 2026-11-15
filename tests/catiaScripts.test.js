const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { fixture } = require('./helpers');
const { buildCatiaScript, buildRunnerScript, vbString } = require('../out/catiaScripts');
const { validateCatiaResult, catiaSucceeded } = require('../out/catiaTypes');
const cscript = path.join(process.env.SystemRoot || 'C:\\Windows', 'SysWOW64', 'cscript.exe');

// Real WSH + ADODB execute the generated scripts; CATIA interfaces are fake.
// No GetObject/real APC connection is made by this harness.
const mocks = String.raw`
Dim hFailure, hSource, hCalls
Class FakeProjects
    Public Property Get Count
        Count = 1
    End Property
    Public Function Item(index)
        Set Item = New FakeProject
    End Function
End Class
Class FakeProject
    Public Property Get Name
        If Len(gTargetProject) > 0 Then
            Name = gTargetProject
        Else
            Name = "Review"
        End If
    End Property
    Public Property Get VBComponents
        Set VBComponents = New FakeComponents
    End Property
End Class
Class FakeComponents
    Public Property Get Count
        Count = 1
    End Property
    Public Function Item(index)
        Set Item = New FakeComponent
    End Function
    Public Function Add(moduleType)
        Dim component
        hCalls = hCalls & "add;"
        If hFailure = "push.add" Then Err.Raise 5, , "injected add failure"
        Set component = New FakeComponent
        component.ModuleType = moduleType
        Set Add = component
    End Function
    Public Sub Remove(component)
        hCalls = hCalls & "remove;"
        If hFailure = "push.delete" Then Err.Raise 5, , "injected delete failure"
    End Sub
End Class
Class FakeComponent
    Private moduleName
    Public ModuleType
    Private Sub Class_Initialize()
        moduleName = "C5D_UserModule"
        ModuleType = 1
    End Sub
    Public Property Get Name
        Name = moduleName
    End Property
    Public Property Let Name(value)
        hCalls = hCalls & "name;"
        If hFailure = "push.name" Then Err.Raise 5, , "injected name failure"
        moduleName = value
    End Property
    Public Property Get [Type]
        [Type] = ModuleType
    End Property
    Public Property Get CodeModule
        If hFailure = "push.codemodule" Then Err.Raise 5, , "injected module failure"
        Set CodeModule = New FakeCode
    End Property
End Class
Class FakeCode
    Public Property Get CountOfLines
        If hFailure = "push.linecount" Then Err.Raise 5, , "injected linecount failure"
        If Len(hSource) > 0 Then
            CountOfLines = 1
        Else
            CountOfLines = 0
        End If
    End Property
    Public Function Lines(start, count)
        Lines = hSource
    End Function
    Public Sub DeleteLines(start, count)
        hCalls = hCalls & "deletelines;"
        If hFailure = "push.deletelines" Then Err.Raise 5, , "injected deletelines failure"
        hSource = ""
    End Sub
    Public Sub AddFromString(source)
        hCalls = hCalls & "addfromstring;"
        If hFailure = "push.addfromstring" Then Err.Raise 5, , "injected insertion failure"
        hSource = source
    End Sub
End Class
`;
function run(t, kind, { failure = '', source = 'old', modules = [], deletions = [], project = 'Review' } = {}) {
    const directory = fixture(t);
    const request = { operationId: 'review', stepId: 'step', kind, targetProject: kind === 'list' ? '' : project,
        encoding: 'shift_jis', directory, modules, deletions };
    for (const module of modules) if (failure !== 'push.read.load') fs.writeFileSync(path.join(directory, module.file), 'new source');
    const script = buildCatiaScript(request).replace(/Function ConnectProjects\(\)[\s\S]*?End Function/, `Function ConnectProjects()\r\n    Set gProjects = New FakeProjects\r\n    ConnectProjects = True\r\nEnd Function`)
        + mocks + `\r\nhFailure = ${vbString(failure)}\r\nhSource = ${vbString(source)}\r\nhCalls = ""\r\nCATMain\r\nWScript.Echo hCalls\r\n`;
    const scriptPath = path.join(directory, 'harness.vbs'); fs.writeFileSync(scriptPath, script);
    const calls = execFileSync(cscript, ['//nologo', scriptPath], { encoding: 'utf8', windowsHide: true, timeout: 10000 }).trim();
    const result = validateCatiaResult(fs.readFileSync(path.join(directory, 'result.json'), 'utf8'), request);
    return { result, calls, request };
}
test('generated source and runner are ASCII and never clean C5D_ prefixes', () => {
    const req = { operationId: 'op', stepId: 'step', kind: 'push', targetProject: '日本語"\\\n项目', encoding: 'gbk', directory: 'C:\\日本語', modules: [{ name: 'C5D_UserModule', type: 1, file: 'm0.txt' }] };
    const script = buildCatiaScript(req);
    assert.ok(/^[\x00-\x7f]*$/.test(script)); assert.ok(/^[\x00-\x7f]*$/.test(buildRunnerScript(req)));
    assert.match(script, /ApplyModule "C5D_UserModule"/);
    assert.doesNotMatch(script, /CLEANUP INJECTED|Left\(.*C5D_|Polling timeout/);
    assert.doesNotMatch(buildRunnerScript(req), /DeleteFile|Sleep/);
});
test('real VBScript list and export preserve C5D_ modules including empty code', { skip: !fs.existsSync(cscript) }, t => {
    assert.deepEqual(run(t, 'list').result.projects, ['Review']);
    for (const kind of ['pull', 'check']) {
        const { result, request, calls } = run(t, kind, { source: '', project: '日本語"\\\n项目' });
        assert.ok(catiaSucceeded(result)); assert.equal(result.items[0].name, 'C5D_UserModule');
        assert.equal(fs.readFileSync(path.join(request.directory, result.items[0].file), 'utf8'), '');
        assert.equal(calls, '');
    }
});
test('real VBScript Push supports existing C5D_ code and new classes', { skip: !fs.existsSync(cscript) }, t => {
    const first = run(t, 'push', { modules: [{ name: 'C5D_UserModule', type: 1, file: 'm0.txt' }] });
    assert.ok(catiaSucceeded(first.result)); assert.equal(first.result.updatedCount, 1); assert.match(first.calls, /deletelines;addfromstring;/);
    assert.ok(!first.calls.includes('remove;'));
    const cls = run(t, 'push', { modules: [{ name: 'NewClass', type: 2, file: 'm0.txt' }] });
    assert.ok(catiaSucceeded(cls.result)); assert.match(cls.calls, /add;name;/);
});
test('real VBScript Push succeeds with deletions only and no modules', { skip: !fs.existsSync(cscript) }, t => {
    const { result, calls } = run(t, 'push', { modules: [], deletions: ['C5D_UserModule'] });
    assert.ok(catiaSucceeded(result));
    assert.equal(result.expectedCount, 1); assert.equal(result.updatedCount, 0); assert.equal(result.deletedCount, 1);
    assert.equal(calls, 'remove;');
});
test('real VBScript errors abort later steps of the failed module', { skip: !fs.existsSync(cscript) }, t => {
    for (const failure of ['push.read.load', 'push.add', 'push.name', 'push.codemodule', 'push.linecount', 'push.deletelines', 'push.addfromstring', 'push.delete']) {
        const deleting = failure === 'push.delete';
        const { result, calls } = run(t, 'push', { failure,
            modules: deleting ? [] : [{ name: ['push.add', 'push.name'].includes(failure) ? 'NewModule' : 'C5D_UserModule', type: 1, file: 'm0.txt' }],
            deletions: deleting ? ['C5D_UserModule'] : [] });
        assert.ok(!catiaSucceeded(result), failure); assert.equal(result.succeededCount, 0, failure);
        assert.equal(result.failures[0].phase, failure);
        if (['push.read.load', 'push.add', 'push.name', 'push.codemodule', 'push.linecount'].includes(failure)) assert.ok(!calls.includes('deletelines;'), failure);
        if (failure === 'push.deletelines') assert.ok(!calls.includes('addfromstring;'));
    }
});
test('reference templates are generated from the runtime implementation', () => {
    const request = { operationId: '{{OPERATION_ID}}', stepId: '{{STEP_ID}}', kind: 'push', targetProject: '{{TARGET_PROJECT}}', encoding: '{{ENCODING}}', directory: '{{TEMP_DIR}}', modules: [{ name: '{{MODULE_NAME}}', type: 1, file: 'module.txt' }], deletions: [] };
    assert.equal(fs.readFileSync(path.join(__dirname, '../scripts/c5d_push.catvbs.bas_utf'), 'utf8'), buildCatiaScript(request));
    assert.equal(fs.readFileSync(path.join(__dirname, '../scripts/c5d_run_push.vbs.bas_utf'), 'utf8'), buildRunnerScript(request));
});
