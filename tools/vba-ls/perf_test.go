package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"runtime"
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

// For・引数リストの文法の書き換え（SLL の先読みを減らすため）で、受け付ける言語が変わっていない。
func TestRewrittenRulesAcceptSameLanguage(t *testing.T) {
	for _, tc := range []struct {
		name string
		body string
		ok   bool
	}{
		{"Next", "    For i = 1 To 3\n        x = i\n    Next\n", true},
		{"Next i", "    For i = 1 To 3\n        x = i\n    Next i\n", true},
		{"Next j, i", "    For i = 1 To 3\n        For j = 1 To 2\n            x = j\n        Next j, i\n", true},
		{"Next k, j, i", "    For i = 1 To 3\n        For Each j In c\n            For k = 1 To 2\n            Next k, j, i\n", true},
		{"For Each の Next j, i", "    For Each i In c\n        For Each j In i\n        Next j, i\n", true},
		{"内側の Next に変数がない", "    For i = 1 To 3\n        For j = 1 To 2\n        Next, i\n", false},
		{"最後の文が For でない", "    For i = 1 To 3\n        x = 1, i\n", false},
		{"本体がない", "    For i = 1 To 3, i\n", false},
		{"空の実引数", "    Call Foo(, 1)\n    Call Foo(1, , 2)\n    Call Foo(1, ,)\n", false}, // 最後の実引数は省略できない
		{"省略した実引数", "    Call Foo(, 1)\n    Call Foo(1, , 2)\n    Foo , 1\n", true},
		{"名前付き実引数", "    Call Foo(1, , b:=2, c:=3)\n    x = Bar(a:=1)\n", true},
		{"引数の宣言", "End Sub\nSub P(a, ByVal b As Long, Optional c = 1, ParamArray d())\n", true},
		{"Optional の後の位置引数", "End Sub\nSub P(Optional a = 1, b)\n", true}, // 元の invalidParameterList も受け付ける
		{"Property Let の引数", "End Sub\nProperty Let V(ByVal i As Long, Optional j = 1, ByVal v As Long)\nEnd Property\nSub Q()\n", true},
	} {
		lfAndCRLF(t, "Sub M()\n"+tc.body+"End Sub\n", func(t *testing.T, src string) {
			st := ParseSymbols(src).Stats
			if got := st.SyntaxErrors == 0; got != tc.ok {
				t.Errorf("%s: 構文エラー %d 件（UsedLL=%v）, want ok=%v", tc.name, st.SyntaxErrors, st.UsedLL, tc.ok)
			}
			if tc.ok && st.UsedLL {
				t.Errorf("%s: LL にフォールバックした（SLL 失敗 %d 行目）", tc.name, st.SLLErrorLine)
			}
		})
	}
}

// ウォームアップ用の合成モジュールは構文エラーなしで SLL だけで解析でき、stdout には何も出さない。
func TestWarmUpParser(t *testing.T) {
	st := ParseSymbols(warmupSource).Stats
	if st.UsedLL || st.SyntaxErrors != 0 {
		t.Fatalf("ウォームアップ用ソース: %+v", st)
	}
	r, w, err := os.Pipe()
	if err != nil {
		t.Fatal(err)
	}
	orig := os.Stdout
	os.Stdout = w
	warmUpParser()
	os.Stdout = orig
	w.Close()
	var buf [256]byte
	if n, _ := r.Read(buf[:]); n > 0 {
		t.Errorf("stdout に出力された: %q", buf[:n])
	}
	r.Close()
}

// String$( ) の後の宣言も失われない（修正前は構文エラーでプロシージャ以降のシンボルが欠けていた）。
func TestTypedReservedNameKeepsFollowingSymbols(t *testing.T) {
	src := "Function F() As String\n    F = String$(3, \" \")\n    Dim d As PartDocument\nEnd Function\n"
	table := ParseSymbols(src)
	if table.Stats.SyntaxErrors != 0 || table.FindByName("d", "Function:F") == nil {
		t.Fatalf("Stats=%+v Symbols=%+v", table.Stats, table.Symbols)
	}
}

// TestWorkspaceIndexSampleDir は実プロジェクトを索引し、時間・件数・メモリと、ワークスペースの型による診断の増分を記録する。
// CAT5DEV_SAMPLE_DIR が modules を指していれば、その親をワークスペースのルートにして modules の自動検出も確かめる。
// CAT5DEV_WS_DIAG_BUDGET が設定されていれば、ワークスペース由来の診断の増分がそれを超えたら失敗にする。
func TestWorkspaceIndexSampleDir(t *testing.T) {
	files := sampleFiles(t)
	dir := os.Getenv("CAT5DEV_SAMPLE_DIR")
	root := dir
	if strings.EqualFold(filepath.Base(filepath.Clean(dir)), "modules") {
		root = filepath.Dir(filepath.Clean(dir))
	}
	// 1 回目: 初回（コールド）の索引時間。ヒープの増分には ANTLR の共有 DFA の構築分も含まれる
	var m0, m1, m2, m3 runtime.MemStats
	runtime.GC()
	runtime.ReadMemStats(&m0)
	ws := NewWorkspace()
	ws.SetRoots([]string{root})
	st := ws.IndexAll(context.Background())
	runtime.GC()
	runtime.ReadMemStats(&m1)
	// 2 回目: DFA が温まった状態で別の索引を作り、索引そのものが保持するメモリを測る
	ws2 := NewWorkspace()
	ws2.SetRoots([]string{root})
	runtime.GC()
	runtime.ReadMemStats(&m2)
	st2 := ws2.IndexAll(context.Background())
	runtime.GC()
	runtime.ReadMemStats(&m3)
	retained := float64(int64(m3.HeapAlloc)-int64(m2.HeapAlloc)) / (1 << 20)
	t.Logf("索引の時間: 初回 %v（ヒープ増分 %.1f MB、DFA を含む）/ 2 回目 %v（索引が保持するヒープ %.1f MB）",
		st.Duration, float64(int64(m1.HeapAlloc)-int64(m0.HeapAlloc))/(1<<20), st2.Duration, retained)
	runtime.KeepAlive(ws2)
	if retained > 20 {
		t.Logf("注意: 索引が保持するヒープが 20MB を超えた（%.1f MB）。プロシージャ内のシンボルを落とす縮約を検討する", retained)
	}
	mods := ws.Modules()
	pub := 0
	type cnt struct {
		name string
		n    int
	}
	var classes []cnt
	var withErrors []string
	for _, e := range mods {
		pub += len(e.public.list)
		if e.Table.Stats.SyntaxErrors > 0 {
			withErrors = append(withErrors, e.Name)
		}
		if e.Kind == ModuleClass {
			classes = append(classes, cnt{e.Name, len(e.public.list)})
		}
	}
	sort.Slice(classes, func(i, j int) bool { return classes[i].n > classes[j].n })
	if len(classes) > 5 {
		classes = classes[:5]
	}
	t.Logf("索引 %d 件（std %d / class %d / form %d、失敗 %d、スキップ %d）、公開メンバー %d 件、上位クラス %v",
		st.Parsed, st.ByKind[ModuleStd], st.ByKind[ModuleClass], st.ByKind[ModuleForm], st.Failed, st.Skipped, pub, classes)
	t.Logf("構文エラーのあるモジュール（Complete にしない）: %d 件 %v", len(withErrors), withErrors)
	if st.Parsed != len(files) {
		t.Errorf("索引した件数 %d がサンプルの件数 %d と違う", st.Parsed, len(files))
	}

	// 抜き取り検査（サンプルに該当のモジュールがあるときだけ）
	db := sampleTLB(t)
	types := NewDefaultResolver(db, ws.TypesProvider(), ws.ModulesProvider())
	if e := ws.Entry(filepath.Join(dir, "Main.bas_utf")); e != nil {
		for _, name := range []string{"Dump", "ShowAlert", "DEBUG_MODE"} {
			if _, ok := types.Member(TypeRef{Lib: moduleLib, Name: "Main"}, name); !ok {
				t.Errorf("Main.%s が引けない", name)
			}
		}
	}
	if e := ws.Entry(filepath.Join(dir, "frmMain.frm_utf")); e != nil {
		self := types.WithLocal(newSelfProvider(e))
		if _, ok := self.Member(e.selfRef(), "Hide"); !ok {
			t.Error("frmMain の Me.Hide が引けない")
		}
	}
	if e := ws.Entry(filepath.Join(dir, "clsPcHoleInfo.cls_utf")); e != nil {
		if _, ok := types.WithLocal(newSelfProvider(e)).Member(e.selfRef(), "Delete"); !ok {
			t.Error("clsPcHoleInfo の Me.Delete が引けない")
		}
	}
	for _, e := range mods {
		if strings.EqualFold(filepath.Base(filepath.Dir(e.Path)), "temp") {
			t.Errorf("modules 以外のファイルが索引された: %s", e.Path)
		}
	}

	// 診断の増分: ワークスペースあり（自モジュール・Me. を含む）となし（2-0 相当）の差
	base := NewDefaultResolver(db)
	var extra []string
	totalWS, totalBase := 0, 0
	for _, f := range files {
		e := ws.Entry(f)
		if e == nil {
			t.Errorf("索引に無い: %s", f)
			continue
		}
		text := stripBOM(readSample(t, f))
		lines := splitLines(text)
		with := BuildDiagnostics(newDocumentFull(e.URI, text, lines, e.Table, e, types))
		without := BuildDiagnostics(newDocumentFull("file:///x.bas", text, lines, e.Table, nil, base))
		totalWS += len(with)
		totalBase += len(without)
		seen := map[string]int{}
		for _, d := range without {
			seen[fmt.Sprintf("%d:%d:%s", d.Range.Start.Line, d.Range.Start.Character, d.Message)]++
		}
		for _, d := range with {
			k := fmt.Sprintf("%d:%d:%s", d.Range.Start.Line, d.Range.Start.Character, d.Message)
			if seen[k] > 0 {
				seen[k]--
				continue
			}
			extra = append(extra, fmt.Sprintf("%s:%d: %s | %s", filepath.Base(f), d.Range.Start.Line+1, d.Message, strings.TrimSpace(lines[d.Range.Start.Line])))
		}
	}
	t.Logf("診断: ワークスペースなし %d 件 / あり %d 件、ワークスペース由来の増分 %d 件", totalBase, totalWS, len(extra))
	for i, s := range extra {
		if i >= 20 {
			t.Logf("  …ほか %d 件", len(extra)-20)
			break
		}
		t.Logf("  %s", s)
	}
	if b := os.Getenv("CAT5DEV_WS_DIAG_BUDGET"); b != "" {
		n, err := strconv.Atoi(b)
		if err != nil {
			t.Fatalf("CAT5DEV_WS_DIAG_BUDGET=%q: %v", b, err)
		}
		if len(extra) > n {
			t.Fatalf("ワークスペース由来の診断 %d 件が予算 %d 件を超えた", len(extra), n)
		}
	}
	runtime.KeepAlive(ws)
}

// sampleTLB は同梱の bin/catia_api.db を読み取り専用で開く（無ければ nil）。
func sampleTLB(t *testing.T) *TLBDatabase {
	t.Helper()
	p := filepath.Join("..", "..", "bin", "catia_api.db")
	if _, err := os.Stat(p); err != nil {
		return nil
	}
	db, err := LoadTLBDatabase(p)
	if err != nil {
		t.Logf("TLB を開けない: %v", err)
		return nil
	}
	t.Cleanup(func() { db.Close() })
	return db
}
