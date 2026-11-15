package main

import (
	"strings"
	"testing"
)

// identCase は識別子・型補完の表駆動のケース。
type identCase struct {
	name    string
	uri     string // 空なら file:///test.bas
	src     string // "|" がカーソル
	want    []string
	notWant []string
	none    bool                          // 候補なし（nil）を期待
	order   []string                      // この順に並ぶこと（sortText の比較）
	kinds   map[string]CompletionItemKind // 候補の種類
	detail  map[string]string             // 候補の Detail
	exactly int                           // > 0 なら同じラベルの候補がこの数だけあること（want[0] について）
}

func runIdentCases(t *testing.T, types *CompositeResolver, cases []identCase) {
	t.Helper()
	for _, tc := range cases {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
				r := types
				if r == nil {
					r = newTestResolver(t)
				}
				uri := tc.uri
				if uri == "" {
					uri = "file:///test.bas"
				}
				clean, line, char := caretPos(t, src)
				items := BuildCompletions(NewDocument(uri, clean, ParseSymbols(clean), r), line, char)
				checkItems(t, items, tc)
			})
		})
	}
}

func checkItems(t *testing.T, items []CompletionItem, tc identCase) {
	t.Helper()
	if tc.none {
		if items != nil {
			t.Fatalf("候補なしのはずが %d 件: %v", len(items), labelList(items, 20))
		}
		return
	}
	byLabel := map[string]CompletionItem{}
	count := map[string]int{}
	for _, it := range items {
		if _, ok := byLabel[it.Label]; !ok {
			byLabel[it.Label] = it
		}
		count[it.Label]++
		if it.SortText == "" {
			t.Errorf("%s に sortText がない", it.Label)
		}
	}
	for _, w := range tc.want {
		if _, ok := byLabel[w]; !ok {
			t.Errorf("候補に %s がない: %v", w, labelList(items, 40))
		}
	}
	for _, n := range tc.notWant {
		if _, ok := byLabel[n]; ok {
			t.Errorf("候補に %s が含まれている", n)
		}
	}
	for i := 1; i < len(tc.order); i++ {
		a, b := byLabel[tc.order[i-1]], byLabel[tc.order[i]]
		if a.SortText == "" || b.SortText == "" || a.SortText >= b.SortText {
			t.Errorf("並び順 %s(%q) < %s(%q) になっていない", tc.order[i-1], a.SortText, tc.order[i], b.SortText)
		}
	}
	for l, k := range tc.kinds {
		if it, ok := byLabel[l]; !ok || it.Kind != k {
			t.Errorf("%s の Kind = %v, want %v", l, it.Kind, k)
		}
	}
	for l, d := range tc.detail {
		if it := byLabel[l]; it.Detail != d {
			t.Errorf("%s の Detail = %q, want %q", l, it.Detail, d)
		}
	}
	if tc.exactly > 0 && count[tc.want[0]] != tc.exactly {
		t.Errorf("%s の件数 = %d, want %d", tc.want[0], count[tc.want[0]], tc.exactly)
	}
}

func labelList(items []CompletionItem, max int) []string {
	var out []string
	for i, it := range items {
		if i >= max {
			out = append(out, "…")
			break
		}
		out = append(out, it.Label)
	}
	return out
}

func TestIdentifierCompletion(t *testing.T) {
	runIdentCases(t, nil, []identCase{
		{name: "ローカルがモジュール変数を隠す", src: "Dim x As Long\nSub A()\n    Dim x As String\n    x|\nEnd Sub\n",
			want: []string{"x"}, exactly: 1, detail: map[string]string{"x": "Dim x As String"}},
		{name: "ローカル → モジュール → 組み込み → キーワード", src: "Dim mB\nSub A()\n    Dim lA\n    |\nEnd Sub\n",
			want: []string{"lA", "mB", "MsgBox", "Dim"}, order: []string{"lA", "mB", "MsgBox", "Dim"}},
		{name: "スコープ外のローカルは出ない", src: "Sub A()\n    Dim onlyA\nEnd Sub\nSub B()\n    |\nEnd Sub\n",
			want: []string{"A", "B"}, notWant: []string{"onlyA"}},
		{name: "Private Sub の中", src: "Private Sub A()\n    Dim v\n    |\nEnd Sub\nPrivate Function F()\nEnd Function\n",
			want: []string{"v", "A", "F"}, kinds: map[string]CompletionItemKind{"v": CompletionItemKindVariable, "F": CompletionItemKindFunction}},
		{name: "引数", src: "Sub A(ByVal p As Long)\n    |\nEnd Sub\n",
			want: []string{"p"}, kinds: map[string]CompletionItemKind{"p": CompletionItemKindVariable}},
		{name: "モジュールの Const・Enum・Enum メンバー", src: "Const K = 1\nEnum E\n    eA\nEnd Enum\nSub A()\n    x = |\nEnd Sub\n",
			want: []string{"K", "E", "eA"}, kinds: map[string]CompletionItemKind{"K": CompletionItemKindConstant, "E": CompletionItemKindEnum, "eA": CompletionItemKindEnumMember}},
		{name: "Property は 1 件", src: "Property Get P() As Long\nEnd Property\nProperty Let P(v As Long)\nEnd Property\nSub A()\n    x = |\nEnd Sub\n",
			want: []string{"P"}, exactly: 1, kinds: map[string]CompletionItemKind{"P": CompletionItemKindProperty}},
		{name: "組み込み関数・定数（式の位置）", src: "Sub A()\n    x = |\nEnd Sub\n",
			want:    []string{"MsgBox", "InStr", "Mid", "UBound", "CStr", "Format", "vbCrLf", "vbYes", "Err", "Debug", "Nothing", "True", "Not"},
			notWant: []string{"Dim", "Sub", "If", "Me"},
			kinds: map[string]CompletionItemKind{"MsgBox": CompletionItemKindFunction, "vbCrLf": CompletionItemKindConstant,
				"vbYes": CompletionItemKindEnumMember, "Err": CompletionItemKindVariable, "Nothing": CompletionItemKindKeyword},
			detail: map[string]string{"vbYes": "VbMsgBoxResult.vbYes = 6", "Mid": "Mid(String, Start As Long, [Length]) As Variant"}},
		{name: "文頭のキーワード", src: "Sub A()\n    |\nEnd Sub\n",
			want: []string{"Dim", "If", "For", "Set", "With", "Call", "MsgBox"}, notWant: []string{"Then", "Nothing"}},
		{name: "Call の後", src: "Const K = 1\nSub A()\n    Call |\nEnd Sub\n",
			want: []string{"A", "MsgBox"}, notWant: []string{"vbCrLf", "Dim", "K", "vbYes"}},
		{name: "文字列内", src: "Sub A()\n    x = \"ab|\"\nEnd Sub\n", none: true},
		{name: "コメント内", src: "Sub A()\n    ' ab|\nEnd Sub\n", none: true},
		{name: "Dim の名前", src: "Sub A()\n    Dim |\nEnd Sub\n", none: true},
		{name: "Dim の 2 つ目の名前", src: "Sub A()\n    Dim a As Long, |\nEnd Sub\n", none: true},
		{name: "Sub の名前", src: "Sub |", none: true},
		{name: "引数名", src: "Sub A(ByVal |", none: true},
		{name: "Optional の後の引数名", src: "Sub A(Optional |", none: true},
		{name: "引数の既定値は式", src: "Sub A(Optional n As Long = |", want: []string{"vbYes"}},
		{name: "Const の値", src: "Const X = |", want: []string{"vbCrLf", "MsgBox"}},
		{name: "ReDim は既存の名前", src: "Sub A()\n    Dim arr()\n    ReDim |\nEnd Sub\n", want: []string{"arr"}},
		{name: "Type の本体", src: "Type T\n    |\nEnd Type\n", none: true},
		{name: "GoTo の後", src: "Sub A()\n    On Error GoTo |\nEnd Sub\n", none: true},
		{name: "Exit の後", src: "Sub A()\n    Exit |\nEnd Sub\n", want: []string{"Do", "For", "Function", "Property", "Sub"}, notWant: []string{"MsgBox", "If"}},
		{name: "End の後", src: "Sub A()\n    End |\nEnd Sub\n", want: []string{"If", "Sub", "With", "Select"}, notWant: []string{"MsgBox"}},
		{name: "日本語の識別子", src: "Sub A()\n    Dim 部品数 As Long\n    部|\nEnd Sub\n", want: []string{"部品数"}},
		{name: "単一行 If の Then の後", src: "Sub A()\n    If x Then |\nEnd Sub\n", want: []string{"Set", "Call", "MsgBox"}, notWant: []string{"Dim", "Sub"}},
		{name: "標準モジュールに Me は無い", src: "Sub A()\n    x = |\nEnd Sub\n", notWant: []string{"Me"}},
		{name: "クラスの Me", uri: "file:///C.cls_utf", src: "Sub A()\n    x = |\nEnd Sub\n", want: []string{"Me"}},
		{name: "数値リテラル", src: "Sub A()\n    x = 1|\nEnd Sub\n", none: true},
		{name: "Event・UDT は識別子に出さない", src: "Public Event Changed()\nType T\n    a As Long\nEnd Type\nSub A()\n    x = |\nEnd Sub\n", notWant: []string{"Changed", "T", "a"}},
	})
}

func TestTypeCompletion(t *testing.T) {
	runIdentCases(t, nil, []identCase{
		{name: "As の後", src: "Type TLocal\n    a As Long\nEnd Type\nEnum ELocal\n    eA\nEnd Enum\nSub A()\n    Dim x As |\nEnd Sub\n",
			want:    []string{"Long", "String", "Variant", "Collection", "Dictionary", "PartDocument", "TLocal", "ELocal", "VbMsgBoxResult"},
			notWant: []string{"MsgBox", "UserForm", "x", "A", "IUnknown"},
			kinds: map[string]CompletionItemKind{"Long": CompletionItemKindKeyword, "Collection": CompletionItemKindClass,
				"PartDocument": CompletionItemKindInterface, "TLocal": CompletionItemKindStruct, "ELocal": CompletionItemKindEnum},
			order: []string{"TLocal", "Collection"}},
		{name: "As の入力途中（サーバーは絞らない）", src: "Sub A()\n    Dim x As Lo|\nEnd Sub\n", want: []string{"Long", "String", "PartDocument"}},
		{name: "修飾付き", src: "Sub A()\n    Dim x As Scripting.|\nEnd Sub\n", want: []string{"Dictionary", "FileSystemObject"}, notWant: []string{"Collection", "Long", "PartDocument"}},
		{name: "VBA 修飾", src: "Sub A()\n    Dim x As VBA.|\nEnd Sub\n", want: []string{"Collection", "Long"}, notWant: []string{"Dictionary"}},
		{name: "New の後", src: "Type TLocal\n    a As Long\nEnd Type\nSub A()\n    Set c = New |\nEnd Sub\n",
			want: []string{"Collection", "Dictionary", "FileSystemObject"}, notWant: []string{"PartDocument", "Long", "TLocal", "File", "ErrObject"}},
		{name: "As New の後", src: "Sub A()\n    Dim c As New |\nEnd Sub\n", want: []string{"Collection", "Dictionary"}, notWant: []string{"PartDocument", "Long"}},
		{name: "New 修飾付き", src: "Sub A()\n    Set c = New Scripting.|\nEnd Sub\n", want: []string{"Dictionary", "FileSystemObject"}, notWant: []string{"File", "TextStream"}},
		{name: "Function の戻り値", src: "Function F() As |", want: []string{"Long", "PartDocument"}},
		{name: "引数の型", src: "Sub A(ByVal p As |", want: []string{"Long", "PartDocument"}},
		{name: "Implements", src: "Implements |", want: []string{"PartDocument"}},
		{name: "TypeOf Is", src: "Sub A()\n    If TypeOf o Is |\nEnd Sub\n", want: []string{"PartDocument", "Collection"}},
		{name: "Open … As", src: "Sub A()\n    Open f For Input As |\nEnd Sub\n", none: true},
	})
}

// 空白で起動したときは型・キーワードの文脈だけに応答する。
func TestCompletionSpaceTrigger(t *testing.T) {
	r := newTestResolver(t)
	for _, tc := range []struct {
		src  string
		want bool
	}{
		{"Sub A()\n    x = 1 |\nEnd Sub\n", false},
		{"Sub A()\n    MsgBox |\nEnd Sub\n", false},
		{"Sub A()\n    |\nEnd Sub\n", false},
		{"Sub A()\n    Dim x As |\nEnd Sub\n", true},
		{"Sub A()\n    Set c = New |\nEnd Sub\n", true},
		{"Sub A()\n    Exit |\nEnd Sub\n", true},
	} {
		clean, line, char := caretPos(t, tc.src)
		doc := NewDocument("file:///test.bas", clean, ParseSymbols(clean), r)
		got := BuildCompletionsTrigger(doc, line, char, " ")
		if (got != nil) != tc.want {
			t.Errorf("%q: 候補 %d 件, want 応答=%v", tc.src, len(got), tc.want)
		}
	}
}

// handler: 起動文字に空白を広告し、context の triggerCharacter を見る。
func TestCompletionHandlerTrigger(t *testing.T) {
	h := newLSPHandler(newTestResolver(t))
	init := h.initializeResult()
	if tc := init.Capabilities.CompletionProvider.TriggerCharacters; strings.Join(tc, "|") != ".| " {
		t.Errorf("triggerCharacters = %q", tc)
	}
	const uri = "file:///a.bas"
	src := "Sub M()\n    x = 1 \n    Dim d As \nEnd Sub\n"
	h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: src}})
	h.waitParsed(uri, 60e9)
	p := completionParams(uri, 1, len("    x = 1 "))
	p.Context = &CompletionContext{TriggerKind: CompletionTriggerCharacter, TriggerCharacter: " "}
	if items := h.completion(p); items != nil {
		t.Errorf("空白の起動で式の位置: %d 件", len(items))
	}
	p = completionParams(uri, 2, len("    Dim d As "))
	p.Context = &CompletionContext{TriggerKind: CompletionTriggerCharacter, TriggerCharacter: " "}
	if !labelsOf(h.completion(p))["PartDocument"] {
		t.Error("空白の起動で As の後に型が出ない")
	}
	p = completionParams(uri, 1, len("    x = 1 "))
	p.Context = &CompletionContext{TriggerKind: 1}
	if !labelsOf(h.completion(p))["MsgBox"] {
		t.Error("手動の起動では識別子の候補を出す")
	}
}

// ワークスペース: 他の標準モジュールの Public・モジュール名・フォーム名・他モジュールの Public Enum。
func TestWorkspaceIdentifierCompletion(t *testing.T) {
	_, types, dir := newTestWorkspace(t, wsFixture)
	for _, tc := range []struct {
		file string
		identCase
	}{
		{"clsA.cls_utf", identCase{name: "他モジュールの Public", src: "Sub T()\n    x = |\nEnd Sub\n",
			want:    []string{"Pub", "NewA", "g", "C", "Module1", "frmA", "EKind", "kA", "kB", "Me"},
			notWant: []string{"Priv", "P", "SetData", "clsA"},
			kinds: map[string]CompletionItemKind{"Module1": CompletionItemKindModule, "frmA": CompletionItemKindModule,
				"Pub": CompletionItemKindFunction, "g": CompletionItemKindVariable, "C": CompletionItemKindConstant, "kA": CompletionItemKindEnumMember},
			detail: map[string]string{"NewA": "Public Function NewA() As clsA", "kA": "EKind.kA"},
			order:  []string{"Name", "Pub", "MsgBox"}}},
		{"Module1.bas_utf", identCase{name: "自モジュールの Public は自表の 1 件", src: "Sub T()\n    x = |\nEnd Sub\n",
			want: []string{"Pub", "Priv", "P"}, exactly: 1, notWant: []string{"Me"}}},
		{"Other.bas_utf", identCase{name: "New の後にクラス・フォーム", src: "Sub T()\n    Set o = New |\nEnd Sub\n",
			want: []string{"clsA", "frmA", "Collection"}, notWant: []string{"Module1", "TInfo", "EKind"}}},
		{"Other.bas_utf", identCase{name: "As の後にワークスペースの型", src: "Sub T()\n    Dim o As |\nEnd Sub\n",
			want: []string{"clsA", "frmA", "TInfo", "EKind"}, notWant: []string{"Module1"}, order: []string{"clsA", "Long"}}},
	} {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			doc, line, char := wsDocAt(t, types, dir, tc.file, withFixture(tc.file, tc.src))
			checkItems(t, BuildCompletions(doc, line, char), tc.identCase)
		})
	}
}
