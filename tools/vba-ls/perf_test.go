package main

import (
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"testing"
	"time"
)

// sampleFiles は CAT5DEV_SAMPLE_DIR 配下の VBA モジュール（*.bas_utf / *.cls_utf / *.frm_utf）を返す。
// 環境変数が未設定ならテストを Skip する。
func sampleFiles(t *testing.T) []string {
	t.Helper()
	dir := os.Getenv("CAT5DEV_SAMPLE_DIR")
	if dir == "" {
		t.Skip("CAT5DEV_SAMPLE_DIR が未設定")
	}
	var files []string
	for _, pat := range []string{"*.bas_utf", "*.cls_utf", "*.frm_utf"} {
		m, err := filepath.Glob(filepath.Join(dir, pat))
		if err != nil {
			t.Fatal(err)
		}
		files = append(files, m...)
	}
	if len(files) == 0 {
		t.Skipf("%s に VBA モジュールがない", dir)
	}
	sort.Strings(files)
	return files
}

func readSample(t *testing.T, path string) string {
	t.Helper()
	b, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	return string(b)
}

// TestParseSampleDir は実プロジェクトのモジュールをすべて解析し、時間と LL へのフォールバック数を記録する。
// CAT5DEV_PARSE_BUDGET_SEC が設定されていれば、合計時間がそれを超えたら失敗にする。
func TestParseSampleDir(t *testing.T) {
	files := sampleFiles(t)
	var total time.Duration
	var maxDur time.Duration
	maxFile, fallbacks, lines := "", 0, 0
	for _, f := range files {
		src := readSample(t, f)
		table := ParseSymbols(src)
		st := table.Stats
		total += st.Duration
		lines += strings.Count(src, "\n") + 1
		if st.Duration > maxDur {
			maxDur, maxFile = st.Duration, filepath.Base(f)
		}
		if st.UsedLL {
			fallbacks++
			t.Logf("LL フォールバック: %s（SLL 失敗 %d 行目、LL の構文エラー %d 件、%v）", filepath.Base(f), st.SLLErrorLine, st.SyntaxErrors, st.Duration)
		}
	}
	t.Logf("%d ファイル・%d 行: 合計 %v、最大 %s（%v）、LL フォールバック %d 件", len(files), lines, total, maxFile, maxDur, fallbacks)
	if b := os.Getenv("CAT5DEV_PARSE_BUDGET_SEC"); b != "" {
		sec, err := strconv.ParseFloat(b, 64)
		if err != nil {
			t.Fatalf("CAT5DEV_PARSE_BUDGET_SEC=%q: %v", b, err)
		}
		if total.Seconds() > sec {
			t.Fatalf("合計 %v が予算 %vs を超えた", total, sec)
		}
	}
}

// comparableTable は Stats（時間・モード）を除いた比較用のコピーを返す。
func comparableTable(t *SymbolTable) SymbolTable {
	c := *t
	c.Stats = ParseStats{}
	return c
}

// TestSLLMatchesLL は二段階パース（SLL → LL）と LL のみのパースで同じシンボル表になることを確かめる。
func TestSLLMatchesLL(t *testing.T) {
	files := sampleFiles(t)
	var autoDur, llDur time.Duration
	for _, f := range files {
		src := readSample(t, f)
		lines := splitLines(src)
		auto := parseModule(src, lines, parseAuto)
		ll := parseModule(src, lines, parseLLOnly)
		autoDur += auto.Stats.Duration
		llDur += ll.Stats.Duration
		if a, b := comparableTable(auto), comparableTable(ll); !reflect.DeepEqual(a, b) {
			t.Errorf("%s: SLL と LL の結果が異なる（UsedLL=%v）", filepath.Base(f), auto.Stats.UsedLL)
		}
	}
	t.Logf("二段階 %v / LL のみ %v", autoDur, llDur)
}

// syntheticModule はベンチマーク用の合成モジュール（n 個のプロシージャ）を作る。
func syntheticModule(n int) string {
	var sb strings.Builder
	sb.WriteString("Option Explicit\n\nPrivate m_doc As PartDocument\nPublic Const MAX_COUNT As Long = 10\n\n")
	for i := 0; i < n; i++ {
		fmt.Fprintf(&sb, `' 処理 %[1]d
Public Function Proc%[1]d(ByVal doc As PartDocument, Optional ByVal n As Long = 1) As Boolean
    Dim part1 As Part, i As Long
    Dim names() As String
    Set part1 = doc.Part
    With part1.Bodies
        For i = 1 To .Count
            If .Item(i).Name <> "" Then
                Debug.Print "body: " & .Item(i).Name, n * 2 + i
            End If
        Next i
    End With
    Select Case n
        Case 1: Proc%[1]d = True
        Case Else: Proc%[1]d = False
    End Select
End Function

`, i)
	}
	return sb.String()
}

func BenchmarkParseSymbols(b *testing.B) {
	src := syntheticModule(50)
	b.SetBytes(int64(len(src)))
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		ParseSymbols(src)
	}
}

func BenchmarkParseSymbolsLLOnly(b *testing.B) {
	src := syntheticModule(50)
	lines := splitLines(src)
	b.SetBytes(int64(len(src)))
	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		parseModule(src, lines, parseLLOnly)
	}
}

// 合成モジュールは構文エラーなしで SLL だけで解析でき、LL と同じ結果になる。
func TestSyntheticModuleUsesSLL(t *testing.T) {
	src := syntheticModule(3)
	lines := splitLines(src)
	auto := parseModule(src, lines, parseAuto)
	if auto.Stats.UsedLL {
		t.Fatalf("合成モジュールで LL にフォールバックした: %+v（%q）", auto.Stats, lines[auto.Stats.SLLErrorLine-1])
	}
	ll := parseModule(src, lines, parseLLOnly)
	if ll.Stats.SyntaxErrors != 0 {
		t.Fatalf("合成モジュールに構文エラー %d 件", ll.Stats.SyntaxErrors)
	}
	if !reflect.DeepEqual(comparableTable(auto), comparableTable(ll)) {
		t.Fatal("SLL と LL の結果が異なる")
	}
}

// 構文エラーのあるソースは LL にフォールバックし、stdout には何も出さずに解析を続ける。
func TestSyntaxErrorFallsBackToLL(t *testing.T) {
	src := "Sub M()\n    Dim doc As PartDocument\n    If Then Else (\nEnd Sub\nSub N()\n    Dim p As Part\nEnd Sub\n"
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	orig := os.Stdout
	os.Stdout = w
	table := ParseSymbols(src)
	os.Stdout = orig
	w.Close()
	var buf [256]byte
	n, _ := r.Read(buf[:])
	r.Close()
	if n > 0 {
		t.Errorf("stdout に出力された: %q", buf[:n])
	}
	if !table.Stats.UsedLL || table.Stats.SyntaxErrors == 0 {
		t.Errorf("Stats = %+v, want UsedLL かつ構文エラーあり", table.Stats)
	}
	if table.FindByName("doc", "Sub:M") == nil {
		t.Errorf("エラー前のシンボルが取れない: %+v", table.Symbols)
	}
}

// 頻出の構文は SLL だけで解析できる（文法の穴で LL にフォールバックしない）。
// いずれも修正前の文法では SLL で構文エラーになり、LL でやり直していた。
func TestCommonConstructsParseWithSLL(t *testing.T) {
	for name, body := range map[string]string{
		"Call x.y":           "    Call oLogger.Export\n    Call Main.Dump(\"a\")\n",
		"obj.Method 引数":      "    cFields.Add sField\n    Err.Clear\n    Me.Delete\n",
		"Debug.Print":        "    Debug.Print sMsg\n    Debug.Print \"a\" & .Item(i).Name, n * 2 + i\n",
		"With 内の Call .x":    "    With d\n        Call .Add(\"a\", 1)\n    End With\n",
		"空白行の後のラベル":          "    Exit Sub\n    \nErrHandler:\n    x = 1\n",
		"コメント行の後のラベル":        "    Exit Sub\n    ' c\nErrHandler:\n    x = 1\n",
		"コメント行の後の ElseIf":    "    If a Then\n        ' c\n    ElseIf b Then\n        x = 1\n    End If\n",
		"行末コメントの後の ElseIf":   "    If a Then\n        x = 1 ' c\n    ElseIf b Then\n        x = 1\n    End If\n",
		"空白行の後の If":          "    x = 1\n    \n    If a Then\n        x = 1\n    End If\n",
		"String$()":          "    s = String$((n + 1) * 2, \" \")\n",
		"Dim As Lib.Type":    "    Dim d As Scripting.Dictionary\n",
		"Set = New Lib.Type": "    Set d = New Scripting.Dictionary\n",
	} {
		lfAndCRLF(t, "Sub M()\n"+body+"End Sub\n", func(t *testing.T, src string) {
			table := ParseSymbols(src)
			if table.Stats.UsedLL {
				t.Errorf("%s: LL にフォールバックした（SLL 失敗 %d 行目、LL の構文エラー %d 件）", name, table.Stats.SLLErrorLine, table.Stats.SyntaxErrors)
			}
		})
	}
}

// String$( ) の後の宣言も失われない（修正前は構文エラーでプロシージャ以降のシンボルが欠けていた）。
func TestTypedReservedNameKeepsFollowingSymbols(t *testing.T) {
	src := "Function F() As String\n    F = String$(3, \" \")\n    Dim d As PartDocument\nEnd Function\n"
	table := ParseSymbols(src)
	if table.Stats.SyntaxErrors != 0 || table.FindByName("d", "Function:F") == nil {
		t.Fatalf("Stats=%+v Symbols=%+v", table.Stats, table.Symbols)
	}
}
