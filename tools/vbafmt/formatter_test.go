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
			got := strings.Join(fixIndentation(lines, opts.IndentSize), "\n")
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

// joinLF は行を LF で結合する (テスト用)
func joinLF(lines ...string) string {
	return strings.Join(lines, "\n")
}
