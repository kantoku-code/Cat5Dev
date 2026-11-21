Option Explicit
' C5D_Bridge_Runtime - Cat5Dev runtime module, installed into C5D_Bridge by catia-run.exe. Do not edit by hand.
' Contract: docs/dev/bridge-contract.md (section "C5D_Bridge_Runtime"). Rules: ASCII only, CRLF, no Attribute line,
' On Error Resume Next in every procedure, no MsgBox / Declare, no property read in an If / loop condition, no module state.
' Ported from the operation.catvbs that src/catiaScripts.ts generated (ConnectProjects / FindProject / ListProjects /
' ExportModules / ApplyModule / DeleteModule): same lookups, same checks, same failure phases and messages.
' Every function returns a String: "ok..." on success, "error:<phase>:<number>:<description>" on failure.
' A project named C5D_Bridge is never listed and never matched as a target.

Private Const RT_VERSION As String = "c5d-runtime:1"
Private Const BRIDGE_PROJECT As String = "C5D_Bridge"

Private Function RtErr(ByVal phase As String, ByVal number As Long, ByVal description As String) As String
    On Error Resume Next
    RtErr = "error:" & phase & ":" & CStr(number) & ":" & description
End Function

Private Function RtEsc(ByVal t As String) As String
    On Error Resume Next
    t = Replace(t, "%", "%25")
    t = Replace(t, vbTab, "%09")
    t = Replace(t, vbLf, "%0A")
    t = Replace(t, vbCr, "%0D")
    RtEsc = t
End Function

' VBE.VBProjects of CATIA's APC (MSAPC.Apc.7.1, then 6.2, then MSAPC.Apc) into vbps; "" or an error text.
Private Function RtProjects(ByRef vbps As Object) As String
    On Error Resume Next
    Dim apc As Object
    Dim ide As Object
    Dim e As Long
    Dim d As String
    Set vbps = Nothing
    Set apc = Nothing
    Err.Clear
    Set apc = CreateObject("MSAPC.Apc.7.1")
    If apc Is Nothing Then
        Err.Clear
        Set apc = CreateObject("MSAPC.Apc.6.2")
    End If
    If apc Is Nothing Then
        Err.Clear
        Set apc = CreateObject("MSAPC.Apc")
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            RtProjects = RtErr("connect.apc", e, d)
            Exit Function
        End If
    End If
    If apc Is Nothing Then
        RtProjects = RtErr("connect.apc", 429, "No APC object")
        Exit Function
    End If
    Err.Clear
    Set ide = Nothing
    Set ide = apc.VBE
    Set vbps = ide.VBProjects
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        Set vbps = Nothing
        RtProjects = RtErr("connect.projects", e, d)
        Exit Function
    End If
    If vbps Is Nothing Then
        RtProjects = RtErr("connect.projects", 424, "No VBProjects")
        Exit Function
    End If
    RtProjects = ""
End Function

' The first VBProject whose Name equals projectName (binary compare; C5D_Bridge never matches) into vbp; "" or an error text.
Private Function RtFindProject(ByVal projectName As String, ByRef vbp As Object) As String
    On Error Resume Next
    Dim vbps As Object
    Dim cand As Object
    Dim r As String
    Dim nm As String
    Dim e As Long
    Dim d As String
    Dim n As Long
    Dim i As Long
    Set vbp = Nothing
    r = RtProjects(vbps)
    If Len(r) > 0 Then
        RtFindProject = r
        Exit Function
    End If
    Err.Clear
    n = -1
    n = vbps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        RtFindProject = RtErr("project.count", e, d)
        Exit Function
    End If
    For i = 1 To n
        Err.Clear
        Set cand = Nothing
        Set cand = vbps.Item(i)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            RtFindProject = RtErr("project.lookup", e, d)
            Exit Function
        End If
        nm = ""
        nm = CStr(cand.Name)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            RtFindProject = RtErr("project.name", e, d)
            Exit Function
        End If
        If StrComp(nm, BRIDGE_PROJECT, vbBinaryCompare) <> 0 Then
            If StrComp(nm, projectName, vbBinaryCompare) = 0 Then
                Set vbp = cand
                RtFindProject = ""
                Exit Function
            End If
        End If
    Next
    RtFindProject = RtErr("project.lookup", 5, "Target project not found")
End Function

' VBComponents of the target project into comps; "" or an error text.
Private Function RtComponents(ByVal projectName As String, ByRef comps As Object) As String
    On Error Resume Next
    Dim vbp As Object
    Dim r As String
    Dim e As Long
    Dim d As String
    Set comps = Nothing
    r = RtFindProject(projectName, vbp)
    If Len(r) > 0 Then
        RtComponents = r
        Exit Function
    End If
    Err.Clear
    Set comps = vbp.VBComponents
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        Set comps = Nothing
        RtComponents = RtErr("project.components", e, d)
        Exit Function
    End If
    If comps Is Nothing Then
        RtComponents = RtErr("project.components", 424, "No VBComponents")
        Exit Function
    End If
    RtComponents = ""
End Function

' The component named moduleName (text compare) into comp, Nothing when there is none; "" or an error text
' (phases <prefix>.count / <prefix>.component / <prefix>.name).
Private Function RtFindModule(ByVal comps As Object, ByVal moduleName As String, ByVal prefix As String, ByRef comp As Object) As String
    On Error Resume Next
    Dim cand As Object
    Dim nm As String
    Dim e As Long
    Dim d As String
    Dim n As Long
    Dim i As Long
    Set comp = Nothing
    Err.Clear
    n = -1
    n = comps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        RtFindModule = RtErr(prefix & ".count", e, d)
        Exit Function
    End If
    For i = 1 To n
        Err.Clear
        Set cand = Nothing
        Set cand = comps.Item(i)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            RtFindModule = RtErr(prefix & ".component", e, d)
            Exit Function
        End If
        nm = ""
        nm = CStr(cand.Name)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            RtFindModule = RtErr(prefix & ".name", e, d)
            Exit Function
        End If
        If StrComp(nm, moduleName, vbTextCompare) = 0 Then
            Set comp = cand
            RtFindModule = ""
            Exit Function
        End If
    Next
    RtFindModule = ""
End Function

Public Function C5D_RT_Version() As Variant
    On Error Resume Next
    C5D_RT_Version = RT_VERSION
End Function

' "ok" then LF + <project name> per VBProject in VBE order, without the projects named C5D_Bridge.
Public Function C5D_RT_ListProjects() As Variant
    On Error Resume Next
    Dim vbps As Object
    Dim cand As Object
    Dim r As String
    Dim s As String
    Dim nm As String
    Dim e As Long
    Dim d As String
    Dim n As Long
    Dim i As Long
    r = RtProjects(vbps)
    If Len(r) > 0 Then
        C5D_RT_ListProjects = r
        Exit Function
    End If
    Err.Clear
    n = -1
    n = vbps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_ListProjects = RtErr("list.count", e, d)
        Exit Function
    End If
    s = "ok"
    For i = 1 To n
        Err.Clear
        Set cand = Nothing
        Set cand = vbps.Item(i)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_ListProjects = RtErr("list.project", e, d)
            Exit Function
        End If
        nm = ""
        nm = CStr(cand.Name)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_ListProjects = RtErr("list.name", e, d)
            Exit Function
        End If
        If StrComp(nm, BRIDGE_PROJECT, vbBinaryCompare) <> 0 Then
            s = s & vbLf & RtEsc(nm)
        End If
    Next
    C5D_RT_ListProjects = s
End Function

' "ok" then LF + <component name> TAB <component type> per VBComponent in VBComponents order (all types, also 100).
Public Function C5D_RT_ListModules(ByVal projectName As Variant) As Variant
    On Error Resume Next
    Dim comps As Object
    Dim comp As Object
    Dim r As String
    Dim s As String
    Dim nm As String
    Dim e As Long
    Dim d As String
    Dim t As Long
    Dim n As Long
    Dim i As Long
    r = RtComponents(CStr(projectName), comps)
    If Len(r) > 0 Then
        C5D_RT_ListModules = r
        Exit Function
    End If
    Err.Clear
    n = -1
    n = comps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_ListModules = RtErr("export.count", e, d)
        Exit Function
    End If
    s = "ok"
    For i = 1 To n
        Err.Clear
        Set comp = Nothing
        Set comp = comps.Item(i)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_ListModules = RtErr("export.component", e, d)
            Exit Function
        End If
        t = -1
        t = comp.Type
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_ListModules = RtErr("export.type", e, d)
            Exit Function
        End If
        nm = ""
        nm = CStr(comp.Name)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_ListModules = RtErr("export.name", e, d)
            Exit Function
        End If
        s = s & vbLf & RtEsc(nm) & vbTab & CStr(t)
    Next
    C5D_RT_ListModules = s
End Function

' "ok:" & CodeModule.Lines(1, CountOfLines) (the text as VBE returns it: no Attribute line, CRLF between lines,
' no CRLF after the last line; "ok:" alone for an empty module).
Public Function C5D_RT_ReadModule(ByVal projectName As Variant, ByVal moduleName As Variant) As Variant
    On Error Resume Next
    Dim comps As Object
    Dim comp As Object
    Dim cm As Object
    Dim r As String
    Dim mn As String
    Dim src As String
    Dim e As Long
    Dim d As String
    Dim n As Long
    mn = ""
    mn = CStr(moduleName)
    r = RtComponents(CStr(projectName), comps)
    If Len(r) > 0 Then
        C5D_RT_ReadModule = r
        Exit Function
    End If
    r = RtFindModule(comps, mn, "export", comp)
    If Len(r) > 0 Then
        C5D_RT_ReadModule = r
        Exit Function
    End If
    If comp Is Nothing Then
        C5D_RT_ReadModule = RtErr("export.component", 5, "Module not found")
        Exit Function
    End If
    Err.Clear
    Set cm = Nothing
    Set cm = comp.CodeModule
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_ReadModule = RtErr("export.codemodule", e, d)
        Exit Function
    End If
    n = -1
    n = cm.CountOfLines
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_ReadModule = RtErr("export.linecount", e, d)
        Exit Function
    End If
    src = ""
    If n > 0 Then
        Err.Clear
        src = cm.Lines(1, n)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_ReadModule = RtErr("export.lines", e, d)
            Exit Function
        End If
    End If
    C5D_RT_ReadModule = "ok:" & src
End Function

' Creates (types 1 and 2 only) or overwrites a module: DeleteLines 1, CountOfLines then AddFromString moduleText (unchanged).
' "ok" or an error text. A new UserForm (type 3) and a type mismatch are refused before anything is changed.
Public Function C5D_RT_WriteModule(ByVal projectName As Variant, ByVal moduleName As Variant, ByVal moduleType As Variant, ByVal moduleText As Variant) As Variant
    On Error Resume Next
    Dim comps As Object
    Dim comp As Object
    Dim cm As Object
    Dim r As String
    Dim mn As String
    Dim src As String
    Dim e As Long
    Dim d As String
    Dim mt As Long
    Dim et As Long
    Dim n As Long
    mn = ""
    mn = CStr(moduleName)
    src = ""
    src = CStr(moduleText)
    Err.Clear
    mt = -1
    mt = CLng(moduleType)
    e = Err.Number
    If e <> 0 Then
        C5D_RT_WriteModule = RtErr("push.type", 5, "Invalid module type")
        Exit Function
    End If
    If mt < 1 Or mt > 3 Then
        C5D_RT_WriteModule = RtErr("push.type", 5, "Invalid module type")
        Exit Function
    End If
    r = RtComponents(CStr(projectName), comps)
    If Len(r) > 0 Then
        C5D_RT_WriteModule = r
        Exit Function
    End If
    r = RtFindModule(comps, mn, "push.lookup", comp)
    If Len(r) > 0 Then
        C5D_RT_WriteModule = r
        Exit Function
    End If
    If comp Is Nothing Then
        If mt = 3 Then
            C5D_RT_WriteModule = RtErr("push.add", 5, "New UserForms require a designer")
            Exit Function
        End If
        Err.Clear
        Set comp = comps.Add(mt)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_WriteModule = RtErr("push.add", e, d)
            Exit Function
        End If
        Err.Clear
        comp.Name = mn
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_WriteModule = RtErr("push.name", e, d)
            Exit Function
        End If
    Else
        Err.Clear
        et = -1
        et = comp.Type
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_WriteModule = RtErr("push.type", e, d)
            Exit Function
        End If
        If et <> mt Then
            C5D_RT_WriteModule = RtErr("push.type", 5, "Module type does not match")
            Exit Function
        End If
    End If
    Err.Clear
    Set cm = Nothing
    Set cm = comp.CodeModule
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_WriteModule = RtErr("push.codemodule", e, d)
        Exit Function
    End If
    n = -1
    n = cm.CountOfLines
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_WriteModule = RtErr("push.linecount", e, d)
        Exit Function
    End If
    If n > 0 Then
        Err.Clear
        cm.DeleteLines 1, n
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_WriteModule = RtErr("push.deletelines", e, d)
            Exit Function
        End If
    End If
    If Len(src) > 0 Then
        Err.Clear
        cm.AddFromString src
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_RT_WriteModule = RtErr("push.addfromstring", e, d)
            Exit Function
        End If
    End If
    C5D_RT_WriteModule = "ok"
End Function

' VBComponents.Remove of the module named moduleName (text compare). "ok" or an error text.
Public Function C5D_RT_DeleteModule(ByVal projectName As Variant, ByVal moduleName As Variant) As Variant
    On Error Resume Next
    Dim comps As Object
    Dim comp As Object
    Dim r As String
    Dim mn As String
    Dim e As Long
    Dim d As String
    mn = ""
    mn = CStr(moduleName)
    r = RtComponents(CStr(projectName), comps)
    If Len(r) > 0 Then
        C5D_RT_DeleteModule = r
        Exit Function
    End If
    r = RtFindModule(comps, mn, "push.lookup", comp)
    If Len(r) > 0 Then
        C5D_RT_DeleteModule = r
        Exit Function
    End If
    If comp Is Nothing Then
        C5D_RT_DeleteModule = RtErr("push.delete", 5, "Module to delete not found")
        Exit Function
    End If
    Err.Clear
    comps.Remove comp
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_RT_DeleteModule = RtErr("push.delete", e, d)
        Exit Function
    End If
    C5D_RT_DeleteModule = "ok"
End Function
