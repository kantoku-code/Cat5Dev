package main

import (
	"strings"
	"testing"
)

func lintCode(code string, diags []Diagnostic) []Diagnostic {
	var out []Diagnostic
	for _, d := range diags {
		if d.Code == code {
			out = append(out, d)
		}
	}
	return out
}

func vba009(src string) []Diagnostic {
	return checkUnmatchedBlocks(splitLines(src), DefaultLintOptions())
}

func vba006(src string, max int) []Diagnostic {
	opts := DefaultLintOptions()
	opts.MaxNestingDepth = max
	return checkNestingDepth(splitLines(src), opts)
}

func assertNoVBA009(t *testing.T, src string) {
	t.Helper()
	if d := vba009(src); len(d) != 0 {
		t.Fatalf("VBA009 が出ないはず: %+v", d)
	}
}

// DumpVariable（入れ子の Select / If の中に For ...: Next が 1 行で書かれている）
func TestVBA009_ColonSeparatedForNextInNestedBlocks(t *testing.T) {
	src := `Function DumpVariable( _
        ByVal x As Variant) As String
    Dim ty As String: ty = TypeName(x)
    Select Case ty
        Case "Boolean":     DumpVariable = x
        #If Win64 Then
        Case "LongLong":    DumpVariable = x & "^"
        #End If
        Case "Date":
            If Abs(x) >= 1 Then d = Month(x) & "/" & Day(x)
        Case "String"
            If StrPtr(x) = 0 Then
                DumpVariable = "vbNullString"
            Else
                DumpVariable = """" & Replace(x, """", """""") & """"
            End If
        Case Else
            If IsObject(x) Then
                Select Case ty
                    Case "Dictionary"
                        For j = 0 To x.count - 1: ar2(j) = DumpVariable(keys(j)) & ":" & DumpVariable(x(keys(j))): Next
                    Case "Collection"
                        For k = 1 To x.count: ar3(k - 1) = DumpVariable(x.Item(k)): Next
                    Case Else
                        DumpVariable = ToStr(x)
                End Select
            ElseIf IsArray(x) Then
                With x
                    For i = lb To xb: ar(i) = DumpVariable(x(i)): Next
                End With
            Else
                Err.Raise 51
        End If
    End Select
End Function
`
	assertNoVBA009(t, src)
}

func TestVBA009_SingleLineWhileWend(t *testing.T) {
	src := `Private Function ArrRank( _
        ByVal arr As Variant) As Integer
    If Not IsArray(arr) Then Err.Raise 13
    Dim i As Integer: i = 0
    On Error Resume Next
    While Err.Number = 0: x = UBound(arr, IncrPre(i)): Wend
    ArrRank = i - 1
End Function
`
	assertNoVBA009(t, src)
}

func TestVBA009_ConditionalCompilationOnlyFirstBranch(t *testing.T) {
	src := `#If VBA7 Then
Private Function GetPtr() As LongPtr
#ElseIf Win32 Then
Private Function GetPtr() As Long
#Else
Private Function GetPtr() As Integer
#End If
    GetPtr = 0
End Function
`
	assertNoVBA009(t, src)
}

func TestVBA009_ContinuationAndThenInComment(t *testing.T) {
	src := `Sub A()
    If a = 1 And _
       b = 2 Then
        c = 1
    End If
    If x Then ' then を含むコメント
        y = 1
    End If
    If s = "a Then b" Then
        y = 2
    End If
    ' コメントの続き _
      If z Then
    Rem 行全体のコメント
End Sub
`
	assertNoVBA009(t, src)
}

func TestVBA009_NextWithMultipleCounters(t *testing.T) {
	src := `Sub A()
    For i = 1 To 3
        For j = 1 To 3
            x = i * j
    Next j, i
End Sub
`
	assertNoVBA009(t, src)
}

func TestVBA009_BlocksInsideSingleLineIf(t *testing.T) {
	src := `Sub A()
    If a Then For i = 1 To 3: x = i: Next
    If b Then x = 1: y = 2 Else z = 3
    Do: x = x + 1: Loop While x < 10
    With obj: .Value = 1: End With
End Sub
`
	assertNoVBA009(t, src)
}

func TestVBA009_ModifiersAndLabels(t *testing.T) {
	src := `Public Static Function F() As Long
    Static n As Long
ErrHandler:
    F = n
End Function
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Type T
    a As Long
End Type
Private Enum E
    e1
End Enum
Public Property Get P() As Long
End Property
`
	assertNoVBA009(t, src)
}

// 本当の End If 忘れは開始行に 1 件だけ。後ろのプロシージャには連鎖しない
func TestVBA009_MissingEndIfReportedOnce(t *testing.T) {
	src := `Sub A()
    If x Then
        For i = 1 To 3
            y = i
        Next
End Sub

Sub B()
    If x Then
    End If
End Sub
`
	d := vba009(src)
	if len(d) != 1 {
		t.Fatalf("1 件のはず: %+v", d)
	}
	if d[0].Line != 1 || d[0].Message != "'End If' がありません（開始: 2 行目）" {
		t.Fatalf("想定外の診断: %+v", d[0])
	}
}

func TestVBA009_MissingEndSubBeforeNextProcedure(t *testing.T) {
	src := `Sub A()
    x = 1

Sub B()
End Sub
`
	d := vba009(src)
	if len(d) != 1 || d[0].Line != 0 || d[0].Message != "'End Sub' がありません（開始: 1 行目）" {
		t.Fatalf("想定外の診断: %+v", d)
	}
}

func TestVBA009_ExtraEndIf(t *testing.T) {
	src := `Sub A()
    If x Then
    End If
    End If
End Sub
`
	d := vba009(src)
	if len(d) != 1 || d[0].Line != 3 || d[0].Message != "対応する開始ブロックのない 'End If' です" {
		t.Fatalf("想定外の診断: %+v", d)
	}
}

func TestVBA009_ExtraNext(t *testing.T) {
	src := `Sub A()
    Next
End Sub
`
	d := vba009(src)
	if len(d) != 1 || d[0].Line != 1 || d[0].Message != "対応する For のない Next です" {
		t.Fatalf("想定外の診断: %+v", d)
	}
}

func TestVBA009_MissingNextInsideSelect(t *testing.T) {
	src := `Sub A()
    Select Case x
        Case 1
            For i = 1 To 3
                y = i
    End Select
End Sub
`
	d := vba009(src)
	if len(d) != 1 || d[0].Line != 3 || d[0].Message != "'Next' がありません（開始: 4 行目）" {
		t.Fatalf("想定外の診断: %+v", d)
	}
}

func TestVBA006_ColonSeparatedForNextDoesNotNest(t *testing.T) {
	var b strings.Builder
	b.WriteString("Sub A()\n")
	for i := 0; i < 5; i++ {
		b.WriteString("    For i = 1 To 3: x = i: Next\n")
	}
	b.WriteString("    If a Then\n        y = 1\n    End If\nEnd Sub\n")
	if d := vba006(b.String(), 2); len(d) != 0 {
		t.Fatalf("VBA006 が出ないはず: %+v", d)
	}
	if d := lintCode("VBA009", Lint(b.String(), DefaultLintOptions())); len(d) != 0 {
		t.Fatalf("VBA009 が出ないはず: %+v", d)
	}
}

func TestVBA006_StillReportsDeepNesting(t *testing.T) {
	src := `Sub A()
    If a Then
        If b Then
            y = 1
        End If
    End If
End Sub
`
	d := vba006(src, 2)
	if len(d) != 1 || d[0].Line != 2 {
		t.Fatalf("3 行目に 1 件のはず: %+v", d)
	}
}

// Rem コメントの中のコロンで文を分けない
func TestVBA009_RemCommentWithColon(t *testing.T) {
	src := `Sub A()
    Rem TODO: For each item, check it
    x = 1: Rem note: If later Then
    rem lower: Do
    REM upper: While x
ErrH: Rem label: With obj
    Rem say "hi": For i = 1 To 3
End Sub
`
	assertNoVBA009(t, src)
	if d := vba006(src, 1); len(d) != 0 {
		t.Fatalf("VBA006 が出ないはず: %+v", d)
	}
	// 文字列の中の Rem はコメントではない
	assertNoVBA009(t, `Sub A()
    x = "Rem: For": For i = 1 To 2: Next
End Sub
`)
}

// Remove や RemX は Rem ではなくコードとして読む
func TestVBA009_RemPrefixedIdentifierIsCode(t *testing.T) {
	assertNoVBA009(t, `Sub A()
    Remove = 1: For i = 1 To 2: Next
    RemX = 2: If a Then: End If
End Sub
`)

	src := `Sub A()
    Remove = 1: For i = 1 To 2
End Sub
`
	d := vba009(src)
	if len(d) != 1 || d[0].Line != 1 || d[0].Message != "'Next' がありません（開始: 2 行目）" {
		t.Fatalf("想定外の診断: %+v", d)
	}
}

// Rem ... _ は次の行もコメントの続き
func TestVBA009_RemCommentContinuation(t *testing.T) {
	src := `Sub A()
    Rem comment continues _
        If x Then
    x = 1: Rem also continues: _
        For i = 1 To 3
    y = 2
End Sub
`
	assertNoVBA009(t, src)
}

// #EndIf（空白なし）でも条件付きコンパイルを閉じる。後ろの本当の End If 忘れは報告する
func TestVBA009_HashEndIfWithoutSpace(t *testing.T) {
	src := `#If VBA7 Then
Private Function GetPtr() As LongPtr
#Else
Private Function GetPtr() As Long
#EndIf
    GetPtr = 0
End Function

Sub B()
    If x Then
        y = 1
End Sub
`
	d := vba009(src)
	if len(d) != 1 || d[0].Line != 9 || d[0].Message != "'End If' がありません（開始: 10 行目）" {
		t.Fatalf("想定外の診断: %+v", d)
	}
}

// EndIf（空白なし）は End If と同じ
func TestVBA009_EndIfWithoutSpace(t *testing.T) {
	src := `Sub A()
    If x Then
        y = 1
    EndIf
    If z Then
        y = 2
    endif
End Sub
`
	assertNoVBA009(t, src)
	if d := vba006(src, 2); len(d) != 0 {
		t.Fatalf("VBA006 が出ないはず: %+v", d)
	}
}

// 修飾子が複数あるプロシージャ（Public Static Function）も 1 段として数える
func TestVBA006_MultipleModifiers(t *testing.T) {
	src := `Public Static Function F() As Long
    Static n As Long
    If a Then
        If b Then
            F = n
        End If
    End If
End Function

Sub G()
    If a Then
    End If
End Sub
`
	d := vba006(src, 2)
	if len(d) != 1 || d[0].Line != 3 || d[0].Message != "ネストの深さが 2 を超えています（現在 3）" {
		t.Fatalf("4 行目に 1 件のはず: %+v", d)
	}
	assertNoVBA009(t, src)
}

// Static x As Long（静的ローカル変数）や Private Declare は深さに影響しない
func TestVBA006_StaticLocalAndDeclareDoNotNest(t *testing.T) {
	src := `Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetTickCount2 Lib "kernel32" () As Long

Sub A()
    Static x As Long
    Static y As Long
    If a Then
        x = 1
    End If
End Sub
`
	if d := vba006(src, 2); len(d) != 0 {
		t.Fatalf("VBA006 が出ないはず: %+v", d)
	}
	assertNoVBA009(t, src)
}
