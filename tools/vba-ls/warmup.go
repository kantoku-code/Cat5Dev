package main

import (
	"fmt"
	"os"
	"time"
)

// warmupSource は ANTLR の DFA を温めるための合成モジュール。よく使う構文を一通り含める。
// 解析結果は捨てる（共有 DFA に予測結果が残ることだけが目的）。
const warmupSource = `Attribute VB_Name = "Warmup"
Option Explicit

' 説明
Private Const MAX_COUNT As Long = 10
Public g_doc As PartDocument
Private m_items() As String, m_n&
Private WithEvents m_app As Application

Private Type TInfo
    Name As String
    Values(2) As Double
End Type

Public Enum EKind
    kA = 1
    kB
End Enum

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Public Property Get Count() As Long
    Count = m_n
End Property

Public Property Let Count(ByVal value As Long)
    m_n = value
End Property

Public Property Set Doc(ByVal oDoc As PartDocument)
    Set g_doc = oDoc
End Property

Public Function Run(ByVal doc As PartDocument, Optional ByVal n As Long = 1, ParamArray rest()) As Boolean
    Dim oPart As Part, i As Long, j As Long, s As String
    Dim dic As New Scripting.Dictionary
    Dim col As Collection
    Static calls As Long
    On Error GoTo ErrHandler
    Set oPart = doc.Part
    Set col = New Collection
    Set dic = CreateObject("Scripting.Dictionary")
    With oPart.Bodies
        For i = 1 To .Count Step 1
            If .Item(i).Name <> "" And Not IsEmpty(i) Then
                Debug.Print "body: " & .Item(i).Name, n * 2 + i; s
            ElseIf i Mod 2 = 0 Or i < 0 Then
                Call col.Add(.Item(i), CStr(i))
            Else
                s = Left$(s, 1) & Mid(s, 2, 3) & String$(2, " ")
            End If
        Next i
    End With
    For Each v In col
        For j = 0 To UBound(m_items)
            dic(CStr(j)) = v
        Next j, v
    Select Case n
        Case 1, 2 To 3: Run = True
        Case Is > 10
            Run = False
        Case Else: Run = False
    End Select
    Do While i > 0
        i = i - 1
        If i = 5 Then Exit Do
    Loop
    Do
        i = i + 1
    Loop Until i >= MAX_COUNT
    While False
    Wend
    ReDim Preserve m_items(n)
    x = Foo(1, , b:=2) + _
        Bar(, 1)
    oPart.Update
    MsgBox "done", vbOKOnly, "title"
    Exit Function
ErrHandler:
    Err.Raise Err.Number, "Run", Err.Description
End Function

Private Sub Class_Initialize()
    If m_n = 0 Then m_n = 1 Else m_n = 2
    Me.Count = 1
End Sub
`

// warmUpParser は合成モジュールを 1 回解析して DFA を温める。解析は parseMu で直列化され、
// 最初の didOpen の解析とは同時に走らない。出力は stderr だけに出す（stdout は LSP の通信路）。
func warmUpParser() {
	parseMu.Lock()
	defer parseMu.Unlock()
	start := time.Now()
	st := ParseSymbols(warmupSource).Stats
	fmt.Fprintf(os.Stderr, "vba-ls: パーサのウォームアップ %v（構文エラー %d 件）\n", time.Since(start).Round(time.Millisecond), st.SyntaxErrors)
}
