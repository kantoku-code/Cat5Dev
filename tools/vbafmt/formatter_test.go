package main

import (
	"os"
	"strings"
	"testing"
)

func TestSplitLines(t *testing.T) {
	tests := []struct {
		name  string
		input string
		want  []string
	}{
		{"CRLF", "a\r\nb\r\nc", []string{"a", "b", "c"}},
		{"LF", "a\nb\nc", []string{"a", "b", "c"}},
		{"CR", "a\rb\rc", []string{"a", "b", "c"}},
		{"mixed", "a\r\nb\nc\r", []string{"a", "b", "c"}},
		{"empty", "", []string{}},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := splitLines(tt.input)
			if len(got) != len(tt.want) {
				t.Fatalf("len=%d want %d", len(got), len(tt.want))
			}
			for i := range got {
				if got[i] != tt.want[i] {
					t.Errorf("line[%d]=%q want %q", i, got[i], tt.want[i])
				}
			}
		})
	}
}

func TestCapitalizeLine(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		// 基本的なキーワード大文字化
		{"dim x as integer", "Dim x As Integer"},
		{"public sub MySub()", "Public Sub MySub()"},
		{"end if", "End If"},
		{"end sub", "End Sub"},
		{"select case x", "Select Case x"},
		{"for each item in col", "For Each item In col"},
		{"option explicit", "Option Explicit"},
		// 文字列内は変換しない
		{`x = "dim y as string"`, `x = "dim y as string"`},
		// コメントは変換しない
		{"' dim x as integer", "' dim x as integer"},
		// コメントと混在
		{"dim x as integer ' comment dim", "Dim x As Integer ' comment dim"},
		// Property
		{"property get Name() as string", "Property Get Name() As String"},
		// Exit
		{"exit sub", "Exit Sub"},
		{"exit function", "Exit Function"},
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := capitalizeLine(tt.input)
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestFixIndentation(t *testing.T) {
	opts := DefaultOptions()
	opts.CapitalizeKeywords = false

	tests := []struct {
		name  string
		input string
		want  string
	}{
		{
			name: "Sub/End Sub",
			input: joinLF(
				"Public Sub MySub()",
				"x = 1",
				"End Sub",
			),
			want: joinLF(
				"Public Sub MySub()",
				"    x = 1",
				"End Sub",
			),
		},
		{
			name: "If/Else/End If",
			input: joinLF(
				"Sub Test()",
				"If x > 0 Then",
				"y = 1",
				"Else",
				"y = 0",
				"End If",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    If x > 0 Then",
				"        y = 1",
				"    Else",
				"        y = 0",
				"    End If",
				"End Sub",
			),
		},
		{
			name: "Select Case",
			input: joinLF(
				"Sub Test()",
				"Select Case x",
				"Case 1",
				"y = 1",
				"Case 2",
				"y = 2",
				"Case Else",
				"y = 0",
				"End Select",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    Select Case x",
				"        Case 1",
				"            y = 1",
				"        Case 2",
				"            y = 2",
				"        Case Else",
				"            y = 0",
				"    End Select",
				"End Sub",
			),
		},
		{
			name: "For/Next",
			input: joinLF(
				"Sub Test()",
				"For i = 1 To 10",
				"x = x + i",
				"Next i",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    For i = 1 To 10",
				"        x = x + i",
				"    Next i",
				"End Sub",
			),
		},
		{
			name: "Attribute at column 0",
			input: joinLF(
				"Attribute VB_Name = \"Module1\"",
				"Sub Test()",
				"x = 1",
				"End Sub",
			),
			want: joinLF(
				"Attribute VB_Name = \"Module1\"",
				"Sub Test()",
				"    x = 1",
				"End Sub",
			),
		},
		{
			name: "single-line If (no indent change)",
			input: joinLF(
				"Sub Test()",
				"If x > 0 Then y = 1",
				"z = 2",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    If x > 0 Then y = 1",
				"    z = 2",
				"End Sub",
			),
		},
		{
			name: "With/End With",
			input: joinLF(
				"Sub Test()",
				"With obj",
				".Name = \"test\"",
				".Value = 1",
				"End With",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    With obj",
				"        .Name = \"test\"",
				"        .Value = 1",
				"    End With",
				"End Sub",
			),
		},
		{
			name: "Type block",
			input: joinLF(
				"Type MyType",
				"x As Integer",
				"y As Long",
				"End Type",
			),
			want: joinLF(
				"Type MyType",
				"    x As Integer",
				"    y As Long",
				"End Type",
			),
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			lines := splitLines(tt.input)
			got := strings.Join(fixIndentation(lines, opts.IndentSize, false), "\n")
			want := tt.want
			if got != want {
				t.Errorf("\ngot:\n%s\nwant:\n%s", got, want)
			}
		})
	}
}

func TestFormatIntegration(t *testing.T) {
	input, err := os.ReadFile("testdata/input.bas_utf")
	if err != nil {
		t.Skip("testdata/input.bas_utf が見つかりません")
	}
	expected, err := os.ReadFile("testdata/expected.bas_utf")
	if err != nil {
		t.Skip("testdata/expected.bas_utf が見つかりません")
	}

	opts := DefaultOptions()
	opts.LineEndings = "LF" // テスト用: expected ファイルは LF
	got := Format(string(input), opts)

	if got != string(expected) {
		// 差分を表示
		gotLines := strings.Split(got, "\n")
		expLines := strings.Split(string(expected), "\n")
		for i := 0; i < len(gotLines) && i < len(expLines); i++ {
			if gotLines[i] != expLines[i] {
				t.Errorf("line %d:\n  got:  %q\n  want: %q", i+1, gotLines[i], expLines[i])
			}
		}
		if len(gotLines) != len(expLines) {
			t.Errorf("行数が異なります: got %d, want %d", len(gotLines), len(expLines))
		}
	}
}

func TestEnsureContinuationSpace(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"Call foo(_", "Call foo( _"},          // _ 直前にスペースなし → 追加
		{"Call foo( _", "Call foo( _"},         // 既にスペースあり → そのまま
		{"x = 1 _", "x = 1 _"},                // 既にスペースあり → そのまま
		{"x = 1_", "x = 1 _"},                 // スペースなし → 追加
		{`x = "abc_"`, `x = "abc_"`},          // 文字列内の _ はスキップ
		{"x = 1 ' abc_", "x = 1 ' abc_"},      // コメント内の _ はスキップ
		{"x = 1", "x = 1"},                    // 継続行でない → そのまま
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := ensureContinuationSpaceLine(tt.input)
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestTrimTrailingSpace(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"x = 1   ", "x = 1"},
		{"x = 1\t\t", "x = 1"},
		{"x = 1", "x = 1"},
		{"", ""},
		{"x = 1 _", "x = 1 _"}, // 継続行の _ は保持
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := trimTrailingSpace([]string{tt.input})
			if got[0] != tt.want {
				t.Errorf("got %q want %q", got[0], tt.want)
			}
		})
	}
}

func TestNormalizeCommaSpacing(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"foo(a,b,c)", "foo(a, b, c)"},
		{"foo(a, b, c)", "foo(a, b, c)"},
		{`foo("a,b",c)`, `foo("a,b", c)`}, // 文字列内のカンマは変換しない
		{"foo(a,b) ' x,y", "foo(a, b) ' x,y"}, // コメント内は変換しない
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := normalizeCommaLine(tt.input)
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestNormalizeCommentSpace(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"'comment", "' comment"},
		{"' comment", "' comment"}, // 既にスペースあり
		{"x = 1 'note", "x = 1 ' note"},
		{"''special", "''special"}, // '' は変換しない
		{`x = "'" & y 'test`, `x = "'" & y ' test`}, // 文字列内の ' はコメントでない
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := normalizeCommentSpaceLine(tt.input)
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestNormalizeOperatorSpacing(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"x=1", "x = 1"},
		{"x=1+2", "x = 1 + 2"},
		{"x = 1 + 2", "x = 1 + 2"}, // 既にスペースあり
		{"x=-1", "x = -1"},          // 単項マイナス
		{"x=(-1)", "x = (-1)"},      // 括弧内単項マイナス
		{"If x<>0 Then", "If x <> 0 Then"},
		{"If x<=10 Then", "If x <= 10 Then"},
		{"If x>=10 Then", "If x >= 10 Then"},
		{`x = &H1F`, `x = &H1F`},                      // 16進リテラルの & はスキップ
		{`x = "a=b" & y`, `x = "a=b" & y`},            // 文字列内は変換しない（& は演算子）
		{"x=1 ' a=b", "x = 1 ' a=b"},                  // コメント内は変換しない
		{"1E+30", "1E+30"},                             // 指数表記: + はスペースなし
		{"1E-10", "1E-10"},                             // 指数表記: - はスペースなし
		{"x = 2.5E+6 + 1", "x = 2.5E+6 + 1"},         // 指数表記と通常演算子の混在
		{"dMin1 = 1E+30", "dMin1 = 1E+30"},            // ユーザー報告のケース
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := normalizeOperatorLine(tt.input)
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestExpandTypeSuffixes(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"Dim x%", "Dim x As Integer"},
		{"Dim x&", "Dim x As Long"},
		{"Dim x!", "Dim x As Single"},
		{"Dim x#", "Dim x As Double"},
		{"Dim x@", "Dim x As Currency"},
		{"Dim x$", "Dim x As String"},
		{"Dim x%, y&", "Dim x As Integer, y As Long"},
		{"Dim x As Integer", "Dim x As Integer"}, // 既に As Type あり
		{"Dim x As Integer, y&", "Dim x As Integer, y As Long"}, // 混在
		{"x% = 10", "x% = 10"}, // Dim 文脈外は変換しない
		{"Private x%", "Private x As Integer"},
		{"Dim x%(10)", "Dim x(10) As Integer"}, // 配列
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := expandTypeSuffixLine(tt.input)
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestNormalizeBlankLines(t *testing.T) {
	tests := []struct {
		name  string
		max   int
		input string
		want  string
	}{
		{
			name:  "連続3空行→2行",
			max:   2,
			input: joinLF("a", "", "", "", "b"),
			want:  joinLF("a", "", "", "b"),
		},
		{
			name:  "プロシージャ間空行保証",
			max:   2,
			input: joinLF("End Sub", "Sub Foo()"),
			want:  joinLF("End Sub", "", "Sub Foo()"),
		},
		{
			name:  "プロシージャ間に既に空行あり",
			max:   2,
			input: joinLF("End Sub", "", "Sub Foo()"),
			want:  joinLF("End Sub", "", "Sub Foo()"),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			lines := splitLines(tt.input)
			got := strings.Join(normalizeBlankLines(lines, tt.max), "\n")
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestSplitColonStatements(t *testing.T) {
	tests := []struct {
		input string
		want  []string
	}{
		{"Dim x: x = 1", []string{"Dim x", "x = 1"}},
		{"    Dim x: x = 1", []string{"    Dim x", "    x = 1"}}, // インデント維持
		{"Label:", []string{"Label:"}},                            // ラベルは分割しない
		{`x = "a:b": y = 1`, []string{`x = "a:b"`, "y = 1"}},    // 文字列内コロンは分割しない
		{"x = 1", []string{"x = 1"}},                             // コロンなし
	}
	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			got := splitColonLine(tt.input)
			if len(got) != len(tt.want) {
				t.Errorf("len=%d want %d: %v", len(got), len(tt.want), got)
				return
			}
			for i := range got {
				if got[i] != tt.want[i] {
					t.Errorf("[%d] got %q want %q", i, got[i], tt.want[i])
				}
			}
		})
	}
}

func TestNormalizeThenPlacement(t *testing.T) {
	tests := []struct {
		name  string
		input string
		want  string
	}{
		{
			name:  "継続行Then同行化",
			input: joinLF("If x > 0 _", "    Then"),
			want:  joinLF("If x > 0 Then"),
		},
		{
			name:  "通常のIfは変換しない",
			input: joinLF("If x > 0 Then", "    y = 1"),
			want:  joinLF("If x > 0 Then", "    y = 1"),
		},
		{
			name:  "継続行でもThenでなければ変換しない",
			input: joinLF("x = 1 _", "    + 2"),
			want:  joinLF("x = 1 _", "    + 2"),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			lines := splitLines(tt.input)
			got := strings.Join(normalizeThenPlacement(lines), "\n")
			if got != tt.want {
				t.Errorf("got %q want %q", got, tt.want)
			}
		})
	}
}

func TestIndentContinuationLines(t *testing.T) {
	tests := []struct {
		name  string
		input string
		want  string
	}{
		{
			name: "継続行後の閉じ括弧はインデントしない",
			input: joinLF(
				"Sub Test()",
				"Call foo(_",
				"vbModeless_",
				")",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    Call foo(_",
				"        vbModeless_",
				"    )",
				"End Sub",
			),
		},
		{
			name: "継続行の次行が+1インデント",
			input: joinLF(
				"Sub Test()",
				"x = 1 _",
				"+ 2",
				"End Sub",
			),
			want: joinLF(
				"Sub Test()",
				"    x = 1 _",
				"        + 2",
				"End Sub",
			),
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			lines := splitLines(tt.input)
			got := strings.Join(fixIndentation(lines, 4, true), "\n")
			if got != tt.want {
				t.Errorf("got:\n%s\nwant:\n%s", got, tt.want)
			}
		})
	}
}

// joinLF は行を LF で結合する (テスト用)
func joinLF(lines ...string) string {
	return strings.Join(lines, "\n")
}
