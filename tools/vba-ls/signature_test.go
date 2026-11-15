package main

import (
	"strings"
	"testing"
	"unicode/utf16"
)

func signatureAt(t *testing.T, types *CompositeResolver, srcWithCaret string) *SignatureHelp {
	t.Helper()
	if types == nil {
		types = newTestResolver(t)
	}
	src, line, char := caretPos(t, srcWithCaret)
	return BuildSignatureHelp(NewDocument("file:///test.bas", src, ParseSymbols(src), types), line, char)
}

func TestSignatureHelpTable(t *testing.T) {
	const sub = "Sub S(a, Optional b As Long = 1)\nEnd Sub\nFunction F(x As String, ParamArray rest()) As Long\nEnd Function\nSub Foo(a, b)\nEnd Sub\n"
	for _, tc := range []struct {
		name  string
		src   string // Sub の中に置く（"|" がカーソル）
		label string // ラベルの接頭辞（"" なら nil を期待）
		ap    uint32
	}{
		{"組み込み・括弧", "    x = InStr(1, s, |", "InStr(", 2},
		{"括弧なし", "    MsgBox \"a\", |", "MsgBox(", 1},
		{"括弧なしの先頭", "    MsgBox |", "MsgBox(", 0},
		{"入れ子", "    x = Mid(Left(s, |", "Left(", 1},
		{"入れ子の後", "    x = Mid(Left(s, 3), |", "Mid(", 1},
		{"$ 形", "    x = Left$(s, |", "Left$(String, Length As Long) As String", 1},
		{"文字列中のカンマ", "    x = Replace(\"a,b\", |", "Replace(", 1},
		{"文字列の中", "    MsgBox \"a, b|", "MsgBox(", 0},
		{"グループ化の括弧", "    x = Mid(s, (1 + |", "Mid(", 1},
		{"名前付き引数", "    MsgBox \"a\", Title:=|", "MsgBox(", 2},
		{"名前付き引数（大文字小文字）", "    MsgBox \"a\", title:=\"x\", buttons:=|", "MsgBox(", 1},
		{"名前付き引数（不明な名前）", "    MsgBox \"a\", Zzz:=|", "MsgBox(", 1},
		{"行継続", "    x = InStr(1, _\n        s, |", "InStr(", 2},
		{"ParamArray", "    x = Array(1, 2, 3, |", "Array(ParamArray ArgList()) As Variant", 0},
		{"自作 Sub（括弧なし）", "    S 1, |", "S(a, [b As Long = 1])", 1},
		{"自作 Function", "    Call F(|", "F(x As String, ParamArray rest() As Variant) As Long", 0},
		{"自作 Function の ParamArray", "    y = F(\"a\", 1, 2, |", "F(", 1},
		{"TLB のメソッド", "    Dim b As Bodies\n    Set x = b.Item(|", "Item(iIndex As Variant) As Body", 0},
		{"Err.Raise", "    Err.Raise 1000, |", "Raise(Number As Long, [Source], [Description], [HelpFile], [HelpContext])", 1},
		{"Debug.Print", "    Debug.Print a, |", "Print(ParamArray OutputList())", 0},
		{"With の中の .Add", "    Dim c As New Collection\n    With c\n        .Add 1, |\n    End With", "Add(", 1},
		{"Collection.Add 括弧", "    Dim c As New Collection\n    c.Add(1, |", "Add(", 1},
		{"配列の添字", "    Dim arr(3)\n    x = arr(|", "", 0},
		{"配列の添字の中の呼び出し", "    Dim arr(3)\n    x = arr(Len(|", "Len(", 0},
		{"代入", "    x = |", "", 0},
		{"If の条件", "    If x |", "", 0},
		{"コメント中", "    ' MsgBox \"a\", |", "", 0},
		{"文の区切り", "    x = 1: MsgBox \"a\", |", "MsgBox(", 1},
		{"名前付き引数とコロン", "    x = 1: Foo a:=1, |", "Foo(a, b)", 1},
		{"Then の後", "    If x Then MsgBox \"a\", |", "MsgBox(", 1},
		{"未知の関数", "    x = Zzz(1, |", "", 0},
		{"Set の右辺", "    Set o = CreateObject(|", "CreateObject(", 0},
	} {
		t.Run(tc.name, func(t *testing.T) {
			lfAndCRLF(t, sub+"Sub T()\n"+tc.src+"\nEnd Sub\n", func(t *testing.T, src string) {
				sh := signatureAt(t, nil, src)
				if tc.label == "" {
					if sh != nil {
						t.Fatalf("nil のはずが %+v", sh)
					}
					return
				}
				if sh == nil || len(sh.Signatures) != 1 {
					t.Fatalf("シグネチャ = %+v", sh)
				}
				if got := sh.Signatures[0].Label; !strings.HasPrefix(got, tc.label) {
					t.Errorf("ラベル = %q, want 接頭辞 %q", got, tc.label)
				}
				if sh.ActiveParameter != tc.ap {
					t.Errorf("activeParameter = %d, want %d", sh.ActiveParameter, tc.ap)
				}
			})
		})
	}
}

// 引数のラベルの範囲は UTF-16 で、formatParam の表示と一致する。
func TestSignatureParameterLabels(t *testing.T) {
	sh := signatureAt(t, nil, "Sub T()\n    MsgBox \"a\", |\nEnd Sub\n")
	if sh == nil {
		t.Fatal("シグネチャがない")
	}
	sig := sh.Signatures[0]
	u := utf16.Encode([]rune(sig.Label))
	part := func(i int) string {
		l := sig.Parameters[i].Label
		return string(utf16.Decode(u[l[0]:l[1]]))
	}
	if len(sig.Parameters) != 5 || part(0) != "Prompt" || part(1) != "[Buttons As VbMsgBoxStyle = vbOKOnly]" || part(4) != "[Context]" {
		t.Errorf("引数のラベル = %+v（%q）", sig.Parameters, sig.Label)
	}
	if sig.Documentation == nil || sig.Documentation.Value == "" {
		t.Error("説明がない")
	}
	// 日本語の引数名でも UTF-16 の範囲になる
	src := "Sub 処理(ByVal 部品 As String, 個数 As Long)\nEnd Sub\nSub T()\n    処理 \"a\", |\nEnd Sub\n"
	sh = signatureAt(t, nil, src)
	if sh == nil || sh.ActiveParameter != 1 {
		t.Fatalf("シグネチャ = %+v", sh)
	}
	sig = sh.Signatures[0]
	u = utf16.Encode([]rune(sig.Label))
	if got := part(1); got != "個数 As Long" {
		t.Errorf("2 番目の引数のラベル = %q（%q）", got, sig.Label)
	}
}

// ワークスペースのクラスのメソッドと、他の標準モジュールの Public。
func TestSignatureHelpWorkspace(t *testing.T) {
	files := map[string]string{"clsCalc.cls_utf": "Public Function Calc(ByVal a As Long, ByVal b As Long) As Long\nEnd Function\n"}
	for k, v := range wsFixture {
		files[k] = v
	}
	_, types, dir := newTestWorkspace(t, files)
	for _, tc := range []struct {
		src, label string
		ap         uint32
	}{
		{"Sub T()\n    Dim o As New clsCalc\n    x = o.Calc(1, |\nEnd Sub\n", "Calc(ByVal a As Long, ByVal b As Long) As Long", 1},
		{"Sub T()\n    Set x = NewA(|\nEnd Sub\n", "NewA() As clsA", 0},
		{"Sub T()\n    Module1.Pub |\nEnd Sub\n", "Pub()", 0},
	} {
		doc, line, char := wsDocAt(t, types, dir, "Other.bas_utf", tc.src)
		sh := BuildSignatureHelp(doc, line, char)
		if sh == nil || sh.Signatures[0].Label != tc.label || sh.ActiveParameter != tc.ap {
			t.Errorf("%q: %+v", tc.src, sh)
		}
	}
}

// handler: capability と JSON-RPC の応答（見つからなければ null）。
func TestSignatureHelpHandler(t *testing.T) {
	h := newLSPHandler(newTestResolver(t))
	caps := h.initializeResult().Capabilities.SignatureHelpProvider
	if caps == nil || strings.Join(caps.TriggerCharacters, "|") != "(|,| " || strings.Join(caps.RetriggerCharacters, "|") != ")" {
		t.Fatalf("signatureHelpProvider = %+v", caps)
	}
	const uri = "file:///a.bas"
	h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: "Sub M()\n    MsgBox \"a\", \n    x = 1\nEnd Sub\n"}})
	var p SignatureHelpParams
	p.TextDocument.URI = uri
	p.Position = Position{Line: 1, Character: uint32(len("    MsgBox \"a\", "))}
	if sh := h.signatureHelp(p); sh == nil || sh.ActiveParameter != 1 {
		t.Errorf("signatureHelp = %+v", sh)
	}
	p.Position = Position{Line: 2, Character: uint32(len("    x = 1"))}
	if sh := h.signatureHelp(p); sh != nil {
		t.Errorf("呼び出しの外で %+v", sh)
	}
}
