package main

import (
	"strings"
	"testing"
)

// analyzeCompletion の文脈判定（識別子・型・キーワード・宣言位置）。
func TestAnalyzeCompletionContext(t *testing.T) {
	type want struct {
		kind      completionKind
		typeLib   string
		creatable bool
		stmtStart bool
		afterThen bool
		afterCall bool
		keywords  string // CtxKeyword のとき "," 区切り
	}
	ident := want{kind: CtxIdentifier}
	stmt := want{kind: CtxIdentifier, stmtStart: true}
	typ := want{kind: CtxType}
	none := want{kind: CtxNone}
	for _, tc := range []struct {
		name string
		src  string // "|" がカーソル。Sub の中に置く行だけを書く
		want want
	}{
		{"文頭", "    |", stmt},
		{"文頭の入力途中", "    Ms|", stmt},
		{"代入の右辺", "    x = |", ident},
		{"引数の中", "    Call Foo(1, |", ident},
		{"Call の後", "    Call |", want{kind: CtxIdentifier, afterCall: true}},
		{"Call の後の入力途中", "    Call Fo|", want{kind: CtxIdentifier, afterCall: true}},
		{"コロンの後は文頭", "    x = 1: |", stmt},
		{"ラベルの後は文頭", "L1: |", stmt},
		{"名前付き引数は文頭ではない", "    Foo a:=|", ident},
		{"Then の後", "    If x Then |", want{kind: CtxIdentifier, stmtStart: true, afterThen: true}},
		{"Else の後", "    If x Then y = 1 Else |", want{kind: CtxIdentifier, stmtStart: true, afterThen: true}},
		{"文字列の中", "    x = \"a|", none},
		{"コメントの中", "    ' a|", none},
		{"数値リテラル", "    x = 1|", none},
		{"As の後", "    Dim x As |", typ},
		{"As の入力途中", "    Dim x As Lo|", typ},
		{"修飾付きの型", "    Dim x As Scripting.|", want{kind: CtxType, typeLib: "Scripting"}},
		{"修飾付きの型の入力途中", "    Dim x As Scripting.Di|", want{kind: CtxType, typeLib: "Scripting"}},
		{"New の後", "    Set c = New |", want{kind: CtxType, creatable: true}},
		{"As New の後", "    Dim c As New |", want{kind: CtxType, creatable: true}},
		{"New 修飾付き", "    Set c = New Scripting.|", want{kind: CtxType, typeLib: "Scripting", creatable: true}},
		{"引数の型", "Sub B(ByVal p As |", typ},
		{"戻り値の型", "Function F() As |", typ},
		{"Implements", "Implements |", typ},
		{"TypeOf Is", "    If TypeOf o Is |", typ},
		{"Open … As は対象外", "    Open f For Input As |", none},
		{"GoTo の後", "    On Error GoTo |", none},
		{"Resume の後", "    On Error Resume |", want{kind: CtxKeyword, keywords: "Next"}},
		{"Exit の後", "    Exit |", want{kind: CtxKeyword, keywords: "Do,For,Function,Property,Sub"}},
		{"End の後", "    End |", want{kind: CtxKeyword, keywords: "Enum,Function,If,Property,Select,Sub,Type,With"}},
		{"On の後", "    On |", want{kind: CtxKeyword, keywords: "Error"}},
		{"On Error の後", "    On Error |", want{kind: CtxKeyword, keywords: "GoTo,Resume"}},
		{"Option の後", "Option |", want{kind: CtxKeyword, keywords: "Base,Compare,Explicit,Private"}},
		{"Option Compare の後", "Option Compare |", want{kind: CtxKeyword, keywords: "Binary,Database,Text"}},
		{"Dim の名前", "    Dim |", none},
		{"Dim の名前の入力途中", "    Dim ab|", none},
		{"Dim の 2 つ目の名前", "    Dim a As Long, |", none},
		{"Dim の名前の後", "    Dim a |", want{kind: CtxKeyword, keywords: "As"}},
		{"Dim の添字", "    Dim a(|", ident},
		{"Const の値", "Const X = |", ident},
		{"Private WithEvents の名前", "Private WithEvents |", none},
		{"Public の後", "Public |", want{kind: CtxKeyword, keywords: "Const,Declare,Dim,Enum,Event,Function,Property,Sub,Type,WithEvents"}},
		{"Property の後", "Public Property |", want{kind: CtxKeyword, keywords: "Get,Let,Set"}},
		{"Sub の名前", "Sub |", none},
		{"Property Get の名前", "Property Get |", none},
		{"引数名", "Sub B(|", none},
		{"ByVal の後の引数名", "Sub B(ByVal |", none},
		{"Optional の後の引数名", "Sub B(a, Optional |", none},
		{"引数の既定値は式", "Sub B(Optional n As Long = |", ident},
		{"Type の名前", "Private Type |", none},
		{"ReDim は既存の名前", "    ReDim |", ident},
		{"For Each の In の後", "    For Each v In |", ident},
	} {
		t.Run(tc.name, func(t *testing.T) {
			src := tc.src
			if !strings.HasPrefix(src, "Sub ") && !strings.HasPrefix(src, "Function ") && !strings.HasPrefix(src, "Public ") &&
				!strings.HasPrefix(src, "Private ") && !strings.HasPrefix(src, "Option ") && !strings.HasPrefix(src, "Implements ") &&
				!strings.HasPrefix(src, "Const ") && !strings.HasPrefix(src, "Property ") {
				src = "Sub A()\n" + src + "\nEnd Sub\n"
			}
			lfAndCRLF(t, src, func(t *testing.T, src string) {
				clean, line, char := caretPos(t, src)
				doc := NewDocument("file:///test.bas", clean, ParseSymbols(clean), newTestResolver(t))
				ctx := analyzeCompletion(doc, line, char)
				got := want{kind: ctx.Kind, typeLib: ctx.TypeLib, creatable: ctx.Creatable, stmtStart: ctx.StmtStart,
					afterThen: ctx.AfterThen, afterCall: ctx.AfterCall, keywords: strings.Join(ctx.Keywords, ",")}
				if got != tc.want {
					t.Errorf("文脈 = %+v, want %+v", got, tc.want)
				}
			})
		})
	}
}

// Type / Enum の本体の文頭では候補を出さない（メンバー名を書く位置）。型の位置は型の文脈。
func TestAnalyzeCompletionInTypeBody(t *testing.T) {
	for _, tc := range []struct {
		src  string
		want completionKind
	}{
		{"Type T\n    |\nEnd Type\n", CtxNone},
		{"Enum E\n    |\nEnd Enum\n", CtxNone},
		{"Type T\n    a As |\nEnd Type\n", CtxType},
		{"Enum E\n    a = |\nEnd Enum\n", CtxIdentifier},
	} {
		lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
			clean, line, char := caretPos(t, src)
			doc := NewDocument("file:///test.bas", clean, ParseSymbols(clean), newTestResolver(t))
			if got := analyzeCompletion(doc, line, char).Kind; got != tc.want {
				t.Errorf("%q: Kind = %v, want %v", tc.src, got, tc.want)
			}
		})
	}
}
