Option Explicit
' C5D_Bridge_Push v2 - Cat5Dev bridge bootstrap. Do not edit by hand.
' Contract: docs/dev/bridge-contract.md (section "C5D_Bridge_Push"). Rules: ASCII only, CRLF, no Attribute line,
' On Error Resume Next in every procedure, no MsgBox / Declare, no property read in an If / loop condition.
' This module never closes, saves or removes anything; catia-run.exe closes the project from outside.

Private Const BRIDGE_VERSION As String = "c5d-bridge:2"
Private Const BRIDGE_PROJECT As String = "C5D_Bridge"
Private Const BRIDGE_MODULE As String = "C5D_Bridge_Push"

Private Function BrErr(ByVal kind As String, ByVal number As Long, ByVal description As String) As String
    On Error Resume Next
    BrErr = "error:" & kind & ":" & CStr(number) & ":" & description
End Function

Private Function BrEsc(ByVal t As String) As String
    On Error Resume Next
    t = Replace(t, "%", "%25")
    t = Replace(t, vbTab, "%09")
    t = Replace(t, vbLf, "%0A")
    t = Replace(t, vbCr, "%0D")
    BrEsc = t
End Function

Private Function BrApc() As Object
    On Error Resume Next
    Dim apc As Object
    Set apc = Nothing
    Set apc = CreateObject("MSAPC.Apc.7.1")
    If apc Is Nothing Then
        Err.Clear
        Set apc = CreateObject("MSAPC.Apc")
    End If
    Err.Clear
    Set BrApc = apc
End Function

Public Function C5D_Version() As Variant
    On Error Resume Next
    C5D_Version = BRIDGE_VERSION
End Function

Public Function C5D_Self() As Variant
    On Error Resume Next
    Dim apc As Object
    Dim ex As Object
    Dim e As Long
    Set apc = BrApc()
    If apc Is Nothing Then
        C5D_Self = BrErr("self.apc", 429, "No APC object")
        Exit Function
    End If
    Err.Clear
    Set ex = Nothing
    Set ex = apc.ExecutingProject
    e = Err.Number
    If ex Is Nothing Then
        C5D_Self = BrErr("self.exec", e, "No executing project")
        Exit Function
    End If
    Set C5D_Self = ex
End Function

Public Function C5D_Who() As Variant
    On Error Resume Next
    Dim apc As Object
    Dim ex As Object
    Dim vbp As Object
    Dim comps As Object
    Dim comp As Object
    Dim pn As String
    Dim dn As String
    Dim fn As String
    Dim cn As String
    Dim s As String
    Dim e As Long
    Dim d As String
    Dim t As Long
    Dim n As Long
    Dim i As Long
    Set apc = BrApc()
    If apc Is Nothing Then
        C5D_Who = BrErr("who.apc", 429, "No APC object")
        Exit Function
    End If
    Err.Clear
    Set ex = Nothing
    Set ex = apc.ExecutingProject
    e = Err.Number
    If ex Is Nothing Then
        C5D_Who = BrErr("who.exec", e, "No executing project")
        Exit Function
    End If
    Err.Clear
    pn = ""
    pn = CStr(ex.Name)
    dn = ""
    dn = CStr(ex.DisplayName)
    Set vbp = Nothing
    Set vbp = ex.VBProject
    fn = ""
    fn = CStr(vbp.FileName)
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_Who = BrErr("who.project", e, d)
        Exit Function
    End If
    s = "ok" & vbLf & BrEsc(pn) & vbTab & BrEsc(dn) & vbTab & BrEsc(fn)
    Err.Clear
    Set comps = Nothing
    Set comps = vbp.VBComponents
    n = -1
    n = comps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_Who = BrErr("who.components", e, d)
        Exit Function
    End If
    For i = 1 To n
        Err.Clear
        Set comp = Nothing
        Set comp = comps.Item(i)
        cn = ""
        cn = CStr(comp.Name)
        t = -1
        t = comp.Type
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Who = BrErr("who.component", e, d)
            Exit Function
        End If
        s = s & vbLf & BrEsc(cn) & vbTab & CStr(t)
    Next
    C5D_Who = s
End Function

Public Function C5D_List() As Variant
    On Error Resume Next
    Dim apc As Object
    Dim ex As Object
    Dim aps As Object
    Dim ap As Object
    Dim pn As String
    Dim dn As String
    Dim me1 As String
    Dim same As Boolean
    Dim s As String
    Dim e As Long
    Dim d As String
    Dim n As Long
    Dim i As Long
    Set apc = BrApc()
    If apc Is Nothing Then
        C5D_List = BrErr("list.apc", 429, "No APC object")
        Exit Function
    End If
    Err.Clear
    Set ex = Nothing
    Set ex = apc.ExecutingProject
    Set aps = Nothing
    Set aps = apc.Projects
    n = -1
    n = aps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_List = BrErr("list.projects", e, d)
        Exit Function
    End If
    s = "ok"
    For i = 1 To n
        Err.Clear
        Set ap = Nothing
        Set ap = aps.Item(i)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_List = BrErr("list.item", e, d)
            Exit Function
        End If
        pn = ""
        pn = CStr(ap.Name)
        dn = ""
        dn = CStr(ap.DisplayName)
        same = False
        same = (ap Is ex)
        me1 = "0"
        If same Then me1 = "1"
        Err.Clear
        s = s & vbLf & BrEsc(pn) & vbTab & BrEsc(dn) & vbTab & me1
    Next
    C5D_List = s
End Function

Public Function C5D_Install(ByVal moduleName As Variant, ByVal source As Variant) As Variant
    On Error Resume Next
    Dim apc As Object
    Dim ex As Object
    Dim vbp As Object
    Dim comps As Object
    Dim comp As Object
    Dim hit As Object
    Dim cm As Object
    Dim mn As String
    Dim src As String
    Dim pn As String
    Dim cn As String
    Dim e As Long
    Dim d As String
    Dim t As Long
    Dim n As Long
    Dim i As Long
    mn = ""
    mn = CStr(moduleName)
    src = ""
    src = CStr(source)
    If StrComp(mn, BRIDGE_MODULE, vbTextCompare) = 0 Then
        C5D_Install = "error:self"
        Exit Function
    End If
    If Len(mn) = 0 Then
        C5D_Install = BrErr("install.name", 5, "Empty module name")
        Exit Function
    End If
    Set apc = BrApc()
    If apc Is Nothing Then
        C5D_Install = BrErr("install.apc", 429, "No APC object")
        Exit Function
    End If
    Err.Clear
    Set ex = Nothing
    Set ex = apc.ExecutingProject
    e = Err.Number
    If ex Is Nothing Then
        C5D_Install = BrErr("install.exec", e, "No executing project")
        Exit Function
    End If
    Err.Clear
    pn = ""
    pn = CStr(ex.Name)
    If StrComp(pn, BRIDGE_PROJECT, vbBinaryCompare) <> 0 Then
        C5D_Install = BrErr("install.project", 5, "Executing project is not " & BRIDGE_PROJECT)
        Exit Function
    End If
    Err.Clear
    Set vbp = Nothing
    Set vbp = ex.VBProject
    Set comps = Nothing
    Set comps = vbp.VBComponents
    n = -1
    n = comps.Count
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_Install = BrErr("install.components", e, d)
        Exit Function
    End If
    Set hit = Nothing
    For i = 1 To n
        Err.Clear
        Set comp = Nothing
        Set comp = comps.Item(i)
        cn = ""
        cn = CStr(comp.Name)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Install = BrErr("install.lookup", e, d)
            Exit Function
        End If
        If StrComp(cn, mn, vbTextCompare) = 0 Then
            Set hit = comp
            Exit For
        End If
    Next
    If hit Is Nothing Then
        Err.Clear
        Set hit = comps.Add(1)
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Install = BrErr("install.add", e, d)
            Exit Function
        End If
        Err.Clear
        hit.Name = mn
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Install = BrErr("install.rename", e, d)
            Exit Function
        End If
    Else
        Err.Clear
        t = -1
        t = hit.Type
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Install = BrErr("install.type", e, d)
            Exit Function
        End If
        If t <> 1 Then
            C5D_Install = BrErr("install.type", 5, "Existing module is not a standard module")
            Exit Function
        End If
    End If
    Err.Clear
    Set cm = Nothing
    Set cm = hit.CodeModule
    n = -1
    n = cm.CountOfLines
    e = Err.Number
    d = Err.Description
    If e <> 0 Then
        C5D_Install = BrErr("install.codemodule", e, d)
        Exit Function
    End If
    If n > 0 Then
        Err.Clear
        cm.DeleteLines 1, n
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Install = BrErr("install.deletelines", e, d)
            Exit Function
        End If
    End If
    If Len(src) > 0 Then
        Err.Clear
        cm.AddFromString src
        e = Err.Number
        d = Err.Description
        If e <> 0 Then
            C5D_Install = BrErr("install.addfromstring", e, d)
            Exit Function
        End If
    End If
    C5D_Install = "ok"
End Function
