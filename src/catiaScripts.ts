import { CatiaRequest, safeFileName } from './catiaTypes';

/** ASCII source avoids the different CATScript/WSH source code pages. */
export function vbString(value: string): string {
    const parts: string[] = [];
    let text = '';
    const flush = (): void => { if (text) { parts.push(`"${text.replace(/"/g, '""')}"`); text = ''; } };
    for (let i = 0; i < value.length; i++) {
        const code = value.charCodeAt(i);
        if (code >= 32 && code <= 126) { text += value[i]; }
        else { flush(); parts.push(`ChrW(${code > 32767 ? code - 65536 : code})`); }
    }
    flush(); return parts.length ? parts.join(' & ') : '""';
}
function checked(statement: string, phase: string, target = 'gTargetProject', exit = 'Exit Sub'): string {
    return `
    Err.Clear
    ${statement}
    If Err.Number <> 0 Then
        errN = Err.Number
        errText = Err.Description
        Err.Clear
        AddFailure ${vbString(phase)}, ${target}, errN, errText
        ${exit}
    End If`;
}
const crlf = (source: string): string => source.trimStart().replace(/[ \t]+$/gm, '').replace(/\r?\n/g, '\r\n');
export function buildRunnerScript(request: CatiaRequest): string {
    return crlf(`
Option Explicit
Dim app, service, args
On Error Resume Next
Err.Clear
Set app = GetObject(, "CATIA.Application")
If Err.Number <> 0 Then WScript.Quit 10
If app Is Nothing Then WScript.Quit 10
Err.Clear
Set service = app.SystemService
If Err.Number <> 0 Then WScript.Quit 11
If service Is Nothing Then WScript.Quit 11
args = Array()
Err.Clear
Call service.ExecuteScript(${vbString(request.directory)}, 1, "operation.catvbs", "CATMain", args)
If Err.Number <> 0 Then
    WScript.StdErr.WriteLine CStr(Err.Number) & ": " & Err.Description
    WScript.Quit 12
End If
WScript.Quit 0
`);
}
export function buildCatiaScript(request: CatiaRequest): string {
    for (const item of request.modules ?? []) {
        if (!safeFileName(item.name) || !safeFileName(item.file) || ![1, 2, 3].includes(item.type)) { throw new Error('Invalid transfer item'); }
    }
    for (const name of request.deletions ?? []) { if (!safeFileName(name)) { throw new Error('Invalid deletion target'); } }
    const count = request.kind === 'push' ? (request.modules?.length ?? 0) + (request.deletions?.length ?? 0) : 0;
    const work = request.kind === 'list' ? 'ListProjects' : request.kind === 'push'
        ? [...(request.modules ?? []).map(m => `ApplyModule ${vbString(m.name)}, ${m.type}, ${vbString(m.file)}`),
            ...(request.deletions ?? []).map(name => `DeleteModule ${vbString(name)}`)].join('\n    ')
        : 'ExportModules';
    return crlf(`
Option Explicit
Dim gOperationId, gStepId, gKind, gTargetProject, gEncoding, gDirectory
Dim gFSO, gApc, gProjects, gProject, gComponents
Dim gFailures, gItems, gProjectNames, gExpected, gSucceeded, gUpdated, gDeleted

Sub CATMain()
    On Error Resume Next
    Dim errN, errText
    gOperationId = ${vbString(request.operationId)}
    gStepId = ${vbString(request.stepId)}
    gKind = ${vbString(request.kind)}
    gTargetProject = ${vbString(request.targetProject)}
    gEncoding = ${vbString(request.encoding)}
    gDirectory = ${vbString(request.directory)}
    gFailures = ""
    gItems = ""
    gProjectNames = ""
    gExpected = ${count}
    gSucceeded = 0
    gUpdated = 0
    gDeleted = 0
    Set gFSO = Nothing
    Set gApc = Nothing
    Set gProjects = Nothing
    Set gProject = Nothing
    Set gComponents = Nothing
    ${checked('Set gFSO = CreateObject("Scripting.FileSystemObject")', 'init.filesystem', 'gTargetProject', 'FinishResult\n        Exit Sub')}
    If Not ConnectProjects() Then
        FinishResult
        Exit Sub
    End If
    If gKind <> "list" Then
        Set gProject = FindProject()
        If gProject Is Nothing Then
            FinishResult
            Exit Sub
        End If
        ${checked('Set gComponents = gProject.VBComponents', 'project.components', 'gTargetProject', 'FinishResult\n        Exit Sub')}
    End If
    ${work}
    FinishResult
End Sub

Function ConnectProjects()
    On Error Resume Next
    Dim errN, errText
    ConnectProjects = False
    Err.Clear
    Set gApc = CreateObject("MSAPC.Apc.7.1")
    If Err.Number <> 0 Then
        Set gApc = Nothing
        Err.Clear
    End If
    If gApc Is Nothing Then
        Err.Clear
        Set gApc = CreateObject("MSAPC.Apc.6.2")
        If Err.Number <> 0 Then
            Set gApc = Nothing
            Err.Clear
        End If
    End If
    If gApc Is Nothing Then
        ${checked('Set gApc = CreateObject("MSAPC.Apc")', 'connect.apc', 'gTargetProject', 'Exit Function')}
    End If
    If gApc Is Nothing Then
        AddFailure "connect.apc", gTargetProject, 429, "No APC object"
        Exit Function
    End If
    ${checked('Set gProjects = gApc.VBE.VBProjects', 'connect.projects', 'gTargetProject', 'Exit Function')}
    ConnectProjects = True
End Function

Function FindProject()
    On Error Resume Next
    Dim total, i, candidate, candidateName, errN, errText
    Set FindProject = Nothing
    ${checked('total = gProjects.Count', 'project.count', 'gTargetProject', 'Exit Function')}
    For i = 1 To total
        ${checked('Set candidate = gProjects.Item(i)', 'project.lookup', 'gTargetProject', 'Exit Function')}
        ${checked('candidateName = candidate.Name', 'project.name', 'gTargetProject', 'Exit Function')}
        If StrComp(candidateName, gTargetProject, vbBinaryCompare) = 0 Then
            Set FindProject = candidate
            Exit Function
        End If
    Next
    AddFailure "project.lookup", gTargetProject, 5, "Target project not found"
End Function

Sub ListProjects()
    On Error Resume Next
    Dim total, i, candidate, candidateName, errN, errText
    ${checked('total = gProjects.Count', 'list.count')}
    gExpected = total
    For i = 1 To total
        ${checked('Set candidate = gProjects.Item(i)', 'list.project', 'CStr(i)')}
        ${checked('candidateName = candidate.Name', 'list.name', 'CStr(i)')}
        AppendJson gProjectNames, JsonQuote(candidateName)
        gSucceeded = gSucceeded + 1
    Next
End Sub

Sub ExportModules()
    On Error Resume Next
    Dim total, i, component, componentType, errN, errText
    ${checked('total = gComponents.Count', 'export.count')}
    For i = 1 To total
        ${checked('Set component = gComponents.Item(i)', 'export.component', 'CStr(i)')}
        ${checked('componentType = component.Type', 'export.type', 'CStr(i)')}
        If componentType = 1 Or componentType = 2 Or componentType = 3 Then
            gExpected = gExpected + 1
            ExportOne component, componentType, gExpected - 1
        End If
    Next
End Sub

Sub ExportOne(component, componentType, ordinal)
    On Error Resume Next
    Dim componentName, codeModule, lineCount, sourceText, fileName, errN, errText
    ${checked('componentName = component.Name', 'export.name', 'CStr(ordinal)')}
    ${checked('Set codeModule = component.CodeModule', 'export.codemodule', 'componentName')}
    ${checked('lineCount = codeModule.CountOfLines', 'export.linecount', 'componentName')}
    sourceText = ""
    If lineCount > 0 Then
        ${checked('sourceText = codeModule.Lines(1, lineCount)', 'export.lines', 'componentName')}
    End If
    fileName = "m" & CStr(ordinal) & ".txt"
    If Not WriteModuleFile(fileName, sourceText, componentName) Then Exit Sub
    AppendJson gItems, "{""name"":" & JsonQuote(componentName) & _
        ",""type"":" & CStr(componentType) & ",""file"":" & JsonQuote(fileName) & "}"
    gSucceeded = gSucceeded + 1
End Sub

Function WriteModuleFile(fileName, sourceText, moduleName)
    On Error Resume Next
    Dim stream, filePath, errN, errText
    WriteModuleFile = False
    ${checked('filePath = gFSO.BuildPath(gDirectory, fileName)', 'export.path', 'moduleName', 'Exit Function')}
    ${checked('Set stream = CreateObject("ADODB.Stream")', 'export.stream', 'moduleName', 'Exit Function')}
    ${checked('stream.Type = 2', 'export.streamtype', 'moduleName', 'Exit Function')}
    ${checked('stream.Charset = gEncoding', 'export.encoding', 'moduleName', 'Exit Function')}
    ${checked('stream.Open', 'export.open', 'moduleName', 'Exit Function')}
    ${checked('stream.WriteText sourceText', 'export.write', 'moduleName', 'Exit Function')}
    ${checked('stream.SaveToFile filePath, 2', 'export.save', 'moduleName', 'Exit Function')}
    ${checked('stream.Close', 'export.close', 'moduleName', 'Exit Function')}
    WriteModuleFile = True
End Function

Function ReadModuleFile(fileName, moduleName, ByRef sourceText)
    On Error Resume Next
    Dim stream, filePath, errN, errText
    ReadModuleFile = False
    sourceText = ""
    ${checked('filePath = gFSO.BuildPath(gDirectory, fileName)', 'push.read.path', 'moduleName', 'Exit Function')}
    ${checked('Set stream = CreateObject("ADODB.Stream")', 'push.read.stream', 'moduleName', 'Exit Function')}
    ${checked('stream.Type = 2', 'push.read.streamtype', 'moduleName', 'Exit Function')}
    ${checked('stream.Charset = gEncoding', 'push.read.encoding', 'moduleName', 'Exit Function')}
    ${checked('stream.Open', 'push.read.open', 'moduleName', 'Exit Function')}
    ${checked('stream.LoadFromFile filePath', 'push.read.load', 'moduleName', 'Exit Function')}
    ${checked('sourceText = stream.ReadText(-1)', 'push.read.text', 'moduleName', 'Exit Function')}
    ${checked('stream.Close', 'push.read.close', 'moduleName', 'Exit Function')}
    ReadModuleFile = True
End Function

Function FindModule(moduleName, ByRef lookupSucceeded)
    On Error Resume Next
    Dim total, i, candidate, candidateName, errN, errText
    lookupSucceeded = False
    Set FindModule = Nothing
    ${checked('total = gComponents.Count', 'push.lookup.count', 'moduleName', 'Exit Function')}
    For i = 1 To total
        ${checked('Set candidate = gComponents.Item(i)', 'push.lookup.component', 'moduleName', 'Exit Function')}
        ${checked('candidateName = candidate.Name', 'push.lookup.name', 'moduleName', 'Exit Function')}
        If StrComp(candidateName, moduleName, vbTextCompare) = 0 Then
            Set FindModule = candidate
            lookupSucceeded = True
            Exit Function
        End If
    Next
    lookupSucceeded = True
End Function

Sub ApplyModule(moduleName, moduleType, fileName)
    On Error Resume Next
    Dim sourceText, component, lookupSucceeded, existingType, codeModule, lineCount, errN, errText
    If Not ReadModuleFile(fileName, moduleName, sourceText) Then Exit Sub
    Set component = FindModule(moduleName, lookupSucceeded)
    If Not lookupSucceeded Then Exit Sub
    If component Is Nothing Then
        If moduleType = 3 Then
            AddFailure "push.add", moduleName, 5, "New UserForms require a designer"
            Exit Sub
        End If
        ${checked('Set component = gComponents.Add(moduleType)', 'push.add', 'moduleName')}
        ${checked('component.Name = moduleName', 'push.name', 'moduleName')}
    Else
        ${checked('existingType = component.Type', 'push.type', 'moduleName')}
        If existingType <> moduleType Then
            AddFailure "push.type", moduleName, 5, "Module type does not match"
            Exit Sub
        End If
    End If
    ${checked('Set codeModule = component.CodeModule', 'push.codemodule', 'moduleName')}
    ${checked('lineCount = codeModule.CountOfLines', 'push.linecount', 'moduleName')}
    If lineCount > 0 Then
        ${checked('codeModule.DeleteLines 1, lineCount', 'push.deletelines', 'moduleName')}
    End If
    If Len(sourceText) > 0 Then
        ${checked('codeModule.AddFromString sourceText', 'push.addfromstring', 'moduleName')}
    End If
    gSucceeded = gSucceeded + 1
    gUpdated = gUpdated + 1
End Sub

Sub DeleteModule(moduleName)
    On Error Resume Next
    Dim component, lookupSucceeded, errN, errText
    Set component = FindModule(moduleName, lookupSucceeded)
    If Not lookupSucceeded Then Exit Sub
    If component Is Nothing Then
        AddFailure "push.delete", moduleName, 5, "Module to delete not found"
        Exit Sub
    End If
    ${checked('gComponents.Remove component', 'push.delete', 'moduleName')}
    gSucceeded = gSucceeded + 1
    gDeleted = gDeleted + 1
End Sub

${JSON_RUNTIME}

Sub FinishResult()
    On Error Resume Next
    Dim result, stream, temporaryPath, finalPath, errN, errText
    If gFSO Is Nothing Then Exit Sub
    result = "{""version"":1"
    result = result & ",""operationId"":" & JsonQuote(gOperationId)
    result = result & ",""stepId"":" & JsonQuote(gStepId)
    result = result & ",""kind"":" & JsonQuote(gKind)
    result = result & ",""targetProject"":" & JsonQuote(gTargetProject)
    result = result & ",""expectedCount"":" & CStr(gExpected)
    result = result & ",""succeededCount"":" & CStr(gSucceeded)
    result = result & ",""failures"":" & "[" & gFailures & "]"
    result = result & ",""items"":" & "[" & gItems & "]"
    result = result & ",""projects"":" & "[" & gProjectNames & "]"
    result = result & ",""updatedCount"":" & CStr(gUpdated)
    result = result & ",""deletedCount"":" & CStr(gDeleted) & "}"
    ${checked('temporaryPath = gFSO.BuildPath(gDirectory, "result.tmp")', 'result.temppath')}
    ${checked('finalPath = gFSO.BuildPath(gDirectory, "result.json")', 'result.path')}
    ${checked('Set stream = CreateObject("ADODB.Stream")', 'result.stream')}
    ${checked('stream.Type = 2', 'result.streamtype')}
    ${checked('stream.Charset = "utf-8"', 'result.encoding')}
    ${checked('stream.Open', 'result.open')}
    ${checked('stream.WriteText result', 'result.write')}
    ${checked('stream.SaveToFile temporaryPath, 2', 'result.save')}
    ${checked('stream.Close', 'result.close')}
    ${checked('gFSO.MoveFile temporaryPath, finalPath', 'result.publish')}
End Sub
`);
}

const JSON_RUNTIME = String.raw`
Sub AddFailure(phase, target, errorCode, message)
    AppendJson gFailures, "{""phase"":" & JsonQuote(phase) & _
        ",""target"":" & JsonQuote(target) & ",""code"":" & CStr(errorCode) & _
        ",""message"":" & JsonQuote(message) & "}"
End Sub
Sub AppendJson(ByRef buffer, value)
    If Len(buffer) > 0 Then buffer = buffer & ","
    buffer = buffer & value
End Sub
Function JsonQuote(value)
    Dim text, output, i, character, code
    text = CStr(value)
    output = ""
    For i = 1 To Len(text)
        character = Mid(text, i, 1)
        code = AscW(character)
        If code < 0 Then code = code + 65536
        Select Case code
            Case 34
                output = output & Chr(92) & Chr(34)
            Case 92
                output = output & Chr(92) & Chr(92)
            Case Else
                If code < 32 Or code > 126 Then
                    output = output & Chr(92) & "u" & Right("0000" & Hex(code), 4)
                Else
                    output = output & character
                End If
        End Select
    Next
    JsonQuote = Chr(34) & output & Chr(34)
End Function
`;
