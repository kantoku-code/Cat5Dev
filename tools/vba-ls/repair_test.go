package main

import (
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

// repairTemplate は Sub A の途中の 1 行（{{LINE}}）だけを差し替えて解析するためのテンプレート。
const repairTemplate = "Option Explicit\n" +
	"Private mA As Long\n" +
	"Sub A()\n" +
	"    Dim d As PartDocument\n" +
	"    Dim n As Long\n" +
	"    Dim x, o\n" +
	"{{LINE}}\n" +
	"    Dim after As Long\n" +
	"End Sub\n" +
	"Sub B(ByVal p As Part)\n" +
	"    Dim z As Long\n" +
	"End Sub\n" +
	"Private mZ As String\n"

func withLine(line string) string { return strings.Replace(repairTemplate, "{{LINE}}", line, 1) }

// 入力途中の行があっても、後ろのプロシージャ・引数・モジュール変数・後続のローカルが落ちない。
func TestRepairKeepsFollowingSymbols(t *testing.T) {
	for _, line := range []string{
		"    d.",
		"    x = d.",
		"    x = d.Part.",
		"    Call Foo(d.",
		"    If n = 1 Then x = d.",
		"    MsgBox \"a\", ",
		"    x = Foo(1, ",
		"    x = Mid(\"a\", 1",
		"    x = n +",
		"    MsgBox \"abc",
		"    x = Foo(\"a",
		"    Set o = New",
		"    Set o = New ",
		"    x = n And",
		"    x = n &",
		"    Foo a:=",
		"    With d.",
		"    x = Foo(1, ' メモ",
		"    .",
		"    x = (n + 1) * (",
	} {
		lfAndCRLF(t, withLine(line), func(t *testing.T, src string) {
			table := ParseSymbols(src)
			if table.Stats.Repaired != 1 {
				t.Errorf("%q: Repaired = %d, want 1", line, table.Stats.Repaired)
			}
			for _, w := range []struct{ name, scope string }{
				{"B", "module"}, {"p", "Sub:B"}, {"z", "Sub:B"}, {"mZ", "module"}, {"mA", "module"}, {"d", "Sub:A"},
			} {
				if table.FindByName(w.name, w.scope) == nil {
					t.Errorf("%q: %s（%s）が落ちた（構文エラー %d 件）", line, w.name, w.scope, table.Stats.SyntaxErrors)
				}
			}
			if !strings.HasPrefix(strings.TrimSpace(line), "With") {
				if table.FindByName("after", "Sub:A") == nil {
					t.Errorf("%q: after が落ちた（構文エラー %d 件）", line, table.Stats.SyntaxErrors)
				}
			}
		})
	}
}

// 補修で前方のトークンの位置は変わらない: 入力途中の行を空行にしたソースと、シンボルの行・列・型名の位置が一致する。
func TestRepairKeepsPositions(t *testing.T) {
	for _, line := range []string{"    d.", "    MsgBox \"あいう", "    x = Foo(\"😀\", ' 日本語のコメント", "    x = 部品 +"} {
		lfAndCRLF(t, withLine(line), func(t *testing.T, src string) {
			got := ParseSymbols(src)
			base := ParseSymbols(strings.Replace(src, line, "", 1))
			pos := func(st *SymbolTable) map[string][4]int {
				m := map[string][4]int{}
				for _, s := range st.Symbols {
					m[s.Scope+"/"+s.Name] = [4]int{s.Line, s.Col, s.TypeNameLine, s.TypeNameCol}
				}
				return m
			}
			gp, bp := pos(got), pos(base)
			for k, v := range bp {
				if gp[k] != v {
					t.Errorf("%q: %s の位置 = %v, want %v", line, k, gp[k], v)
				}
			}
		})
	}
	// 補修後のテキスト: 置換は同じバイト長、追記は物理行の末尾だけ
	src := "Sub A()\r\n    x = Foo(\"a, _\r\n    y\r\n    z = Bar(1, ' あ\r\nEnd Sub\r\n"
	lines := splitLines(src)
	fixed, n := repairSource(src, lines)
	if n != 2 {
		t.Fatalf("補修した論理行 = %d, want 2: %q", n, fixed)
	}
	fl := strings.Split(fixed, "\n")
	ol := strings.Split(src, "\n")
	for i := range ol[:len(ol)-1] {
		if !strings.HasSuffix(fl[i], "\r") {
			t.Errorf("行 %d の CR が失われた: %q", i, fl[i])
		}
	}
	if fl[1] != "    x = Foo(\"a, _\r" {
		t.Errorf("文字列の中の行は変えない: %q", fl[1])
	}
	if fl[2] != "    y\")\r" {
		t.Errorf("閉じていない文字列と括弧は行末に追記: %q", fl[2])
	}
	if want := "    z = Bar(1  " + strings.Repeat(" ", len("' あ")) + ")\r"; fl[3] != want {
		t.Errorf("行末コメントは空白にしてから追記: %q, want %q", fl[3], want)
	}
}

// 正しいコードは補修しない。
func TestRepairLeavesValidCode(t *testing.T) {
	src := "Option Explicit\n" +
		"Private Const MAX_N& = &H10&\n" +
		"Sub A(ByVal n&, ByVal x^)\n" +
		"    Dim s As String * 10\n" +
		"    Dim y As Double\n" +
		"    y = 1.\n" +
		"    Debug.Print n;\n" +
		"    Debug.Print n,\n" +
		"    Print #1, s;\n" +
		"    Print #1, s,\n" +
		"    Debug.Print \"(\" & \"\"\"\"\n" +
		"    y = Foo(1, _\n" +
		"        2)\n" +
		"    y = n + _\n" +
		"        1\n" +
		"    Rem コメント (\n" +
		"    y = 1: Rem (\n" +
		"    ' x = Foo(\n" +
		"Label1:\n" +
		"    If n = 1 Then\n" +
		"    ElseIf n = 2 Then y = 2 Else y = 3\n" +
		"    End If\n" +
		"    Select Case n\n" +
		"        Case Is > 1, 2 To 3\n" +
		"    End Select\n" +
		"    For n = 1 To 3 Step 1: Next\n" +
		"    y = x.New.Step\n" +
		"    GoTo Label1\n" +
		"End Sub\n"
	lfAndCRLF(t, src, func(t *testing.T, src string) {
		fixed, n := repairSource(src, splitLines(src))
		if n != 0 || fixed != src {
			t.Errorf("正しいコードを補修した（%d 行）: %q", n, fixed)
		}
	})
}

// 実サンプルでは補修が 1 行も変更せず、シンボル表は補修なしと完全に一致する。
func TestRepairSampleDirUnchanged(t *testing.T) {
	files := sampleFiles(t)
	for _, f := range files {
		src := stripBOM(readSample(t, f))
		lines := splitLines(src)
		if fixed, n := repairSource(src, lines); n != 0 || fixed != src {
			t.Errorf("%s: %d 行を補修した", filepath.Base(f), n)
			fl, ol := splitLines(fixed), lines
			for i := range ol {
				if i < len(fl) && fl[i] != ol[i] {
					t.Logf("  %d: %q → %q", i+1, ol[i], fl[i])
				}
			}
			continue
		}
		with := parseModuleOpt(src, lines, parseAuto, true)
		without := parseModuleOpt(src, lines, parseAuto, false)
		if with.Stats.Repaired != 0 {
			t.Errorf("%s: Repaired = %d", filepath.Base(f), with.Stats.Repaired)
		}
		if !reflect.DeepEqual(comparableTable(with), comparableTable(without)) {
			t.Errorf("%s: 補修ありと補修なしでシンボル表が異なる", filepath.Base(f))
		}
	}
	t.Logf("%d ファイル: 補修 0 行、シンボル表は補修なしと一致", len(files))
}
