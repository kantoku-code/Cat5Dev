package main

import (
	"strings"
	"testing"
)

// --- 補完（表駆動。各ケースを LF と CRLF で実行） ---

type completionCase struct {
	name    string
	src     string   // "|" がカーソル位置
	want    []string // 含まれるべき候補（nil なら候補なしを期待）
	notWant []string // 含まれてはいけない候補
}

func runCompletionCases(t *testing.T, types *CompositeResolver, cases []completionCase) {
	t.Helper()
	for _, tc := range cases {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
				items := completeAt(t, types, src)
				if tc.want == nil {
					if items != nil {
						t.Fatalf("候補なしのはずが %v", labelsOf(items))
					}
					return
				}
				labels := labelsOf(items)
				for _, w := range tc.want {
					if !labels[w] {
						t.Errorf("候補に %s がない: %v", w, labels)
					}
				}
				for _, n := range tc.notWant {
					if labels[n] {
						t.Errorf("候補に %s が含まれている: %v", n, labels)
					}
				}
			})
		})
	}
}

var partDocMembers = []string{"Part", "Name", "Save"}

func TestCompletionTable(t *testing.T) {
	runCompletionCases(t, nil, []completionCase{
		{"Private Sub の中", "Private Sub A()\n    Dim d As PartDocument\n    d.|\nEnd Sub\n", partDocMembers, nil},
		{"Friend Function の中", "Friend Function F()\n    Dim d As PartDocument\n    d.|\nEnd Function\n", partDocMembers, nil},
		{"Static Sub の中", "Static Sub S()\n    Dim d As PartDocument\n    d.|\nEnd Sub\n", partDocMembers, nil},
		{"Property Get の中", "Property Get P() As Long\n    Dim d As PartDocument\n    d.|\nEnd Property\n", partDocMembers, nil},
		{"Property Let の引数", "Property Let P(ByVal v As PartDocument)\n    v.|\nEnd Property\n", partDocMembers, nil},
		{"引数の型", "Sub A(ByVal doc As PartDocument, Optional n As Long = 1)\n    doc.|\nEnd Sub\n", partDocMembers, nil},
		{"Function Foo$ の中", "Function Foo$()\n    Dim d As PartDocument\n    d.|\nEnd Function\n", partDocMembers, nil},
		{"As New Scripting.Dictionary", "Sub A()\n    Dim d As New Scripting.Dictionary\n    d.|\nEnd Sub\n", []string{"Add", "Exists", "Keys"}, nil},
		{"Function の戻り値 GetDoc().", "Function GetDoc() As PartDocument\nEnd Function\nSub A()\n    GetDoc().|\nEnd Sub\n", partDocMembers, nil},
		{"Function の戻り値 GetDoc.", "Function GetDoc() As PartDocument\nEnd Function\nSub A()\n    GetDoc.|\nEnd Sub\n", partDocMembers, nil},
		{"With", "Sub A()\n    Dim d As PartDocument\n    With d\n        .|\n    End With\nEnd Sub\n", partDocMembers, nil},
		{"ネストした With", "Sub A()\n    Dim d As PartDocument\n    With d\n        With .Part\n            .|\n        End With\n    End With\nEnd Sub\n", []string{"Bodies", "Name"}, []string{"Save"}},
		{"With 行自身の . は外側", "Sub A()\n    Dim d As PartDocument\n    With d\n        With .|Part\n        End With\n    End With\nEnd Sub\n", partDocMembers, nil},
		{"With New Scripting.Dictionary", "Sub A()\n    With New Scripting.Dictionary\n        .|\n    End With\nEnd Sub\n", []string{"Add", "Exists"}, nil},
		{"With CreateObject", "Sub A()\n    With CreateObject(\"Scripting.Dictionary\")\n        .|\n    End With\nEnd Sub\n", []string{"Add", "Exists"}, nil},
		{"Item(1) の要素", "Sub A()\n    Dim d As PartDocument\n    x = d.Part.Bodies.Item(1).|\nEnd Sub\n", []string{"Name"}, []string{"Count"}},
		{"既定メンバー Bodies(1)", "Sub A()\n    Dim d As PartDocument\n    x = d.Part.Bodies(1).|\nEnd Sub\n", []string{"Name"}, []string{"Count"}},
		{"String の後は候補なし", "Sub A()\n    Dim b As Bodies\n    x = b.Item(1).Name.|\nEnd Sub\n", nil, nil},
		{"文字列内の括弧", "Sub A()\n    Dim b As Bodies\n    x = b.Item(CLng(\"(1\")).|\nEnd Sub\n", []string{"Name"}, nil},
		{"閉じていない呼び出しの中", "Sub A()\n    Dim d As PartDocument\n    Call Foo(d.|\nEnd Sub\n", partDocMembers, nil},
		{"行継続（. で始まる次行）", "Sub A()\n    Dim d As PartDocument\n    x = d _\n        .Part.|\nEnd Sub\n", []string{"Bodies"}, nil},
		{"行継続（引数の途中）", "Sub A()\n    Dim d As PartDocument\n    Call F(1, _\n        d.|\nEnd Sub\n", partDocMembers, nil},
		{"文字列内", "Sub A()\n    Dim d As PartDocument\n    x = \"d.|\"\nEnd Sub\n", nil, nil},
		{"' コメント内", "Sub A()\n    Dim d As PartDocument\n    ' d.|\nEnd Sub\n", nil, nil},
		{"Rem コメント内", "Sub A()\n    Dim d As PartDocument\n    Rem d.|\nEnd Sub\n", nil, nil},
		{":Rem コメント内", "Sub A()\n    Dim d As PartDocument\n    x = 1: Rem d.|\nEnd Sub\n", nil, nil},
		{"継続したコメント内", "Sub A()\n    Dim d As PartDocument\n    ' a _\n    d.|\nEnd Sub\n", nil, nil},
		{"文字列連結の後", "Sub A()\n    Dim d As PartDocument\n    x = \"x\" & d.|\nEnd Sub\n", partDocMembers, nil},
		{"入力途中", "Sub A()\n    Dim d As PartDocument\n    d.Pa|\nEnd Sub\n", partDocMembers, nil},
		{"数値リテラル 1.", "Sub A()\n    x = 1.|\nEnd Sub\n", nil, nil},
		{"数値リテラル .5", "Sub A()\n    Dim d As PartDocument\n    With d\n        x = .5|\n    End With\nEnd Sub\n", nil, nil},
		{"配列そのもの", "Sub A()\n    Dim arr(3) As PartDocument\n    arr.|\nEnd Sub\n", nil, nil},
		{"配列の要素", "Sub A()\n    Dim arr(3) As PartDocument\n    arr(0).|\nEnd Sub\n", partDocMembers, nil},
		{"ローカルがモジュール変数を隠す", "Dim d As Part\nSub A()\n    Dim d As PartDocument\n    d.|\nEnd Sub\n", []string{"Save"}, []string{"Bodies"}},
		{"モジュール変数", "Dim d As Part\nSub A()\n    d.|\nEnd Sub\n", []string{"Bodies"}, nil},
		{"日本語の識別子", "Dim 部品 As PartDocument\nSub A()\n    部品.Part.|\nEnd Sub\n", []string{"Bodies"}, nil},
		{"Excel.Application は補完なし", "Sub A()\n    Set xl = CreateObject(\"Excel.Application\")\n    xl.|\nEnd Sub\n", nil, nil},
		{"VBA の Collection", "Sub A()\n    Dim c As Collection\n    c.|\nEnd Sub\n", []string{"Add", "Count", "Item", "Remove"}, []string{"GetItem"}},
		{"INFITF.Collection", "Sub A()\n    Dim c As INFITF.Collection\n    c.|\nEnd Sub\n", []string{"GetItem"}, []string{"Add"}},
		{"UDT のメンバー", "Type TInfo\n    doc As PartDocument\n    n As Long\nEnd Type\nSub A()\n    Dim v As TInfo\n    v.|\nEnd Sub\n", []string{"doc", "n"}, nil},
		{"UDT のメンバーの先", "Type TInfo\n    doc As PartDocument\nEnd Type\nSub A()\n    Dim v As TInfo\n    v.doc.|\nEnd Sub\n", partDocMembers, nil},
		{"Enum のメンバー", "Enum Color\n    Red = 1\n    Green\nEnd Enum\nSub A()\n    x = Color.|\nEnd Sub\n", []string{"Red", "Green"}, nil},
		{"Me. は 2-0 では候補なし", "Sub A()\n    Me.|\nEnd Sub\n", nil, nil},
		{"Set New 由来", "Sub A()\n    Set d = New PartDocument\n    d.|\nEnd Sub\n", partDocMembers, nil},
	})
}

// DB なし（組み込み型のみ）。
func TestCompletionTableWithoutDB(t *testing.T) {
	runCompletionCases(t, NewDefaultResolver(nil, nil), []completionCase{
		{"CreateObject の Dictionary", "Sub A()\n    Set d = CreateObject(\"Scripting.Dictionary\")\n    d.|\nEnd Sub\n", []string{"Add", "Exists", "Keys"}, nil},
		{"Dim As New Collection", "Sub A()\n    Dim c As New Collection\n    c.|\nEnd Sub\n", []string{"Add", "Count"}, nil},
		{"FSO の OpenTextFile は Scripting の TextStream", "Sub A()\n    Dim fso As New Scripting.FileSystemObject\n    fso.OpenTextFile(\"a\").|\nEnd Sub\n", []string{"ReadAll", "WriteLine"}, nil},
		{"GetFolder.Files(1)", "Sub A()\n    Dim fso As New Scripting.FileSystemObject\n    fso.GetFolder(\"c:\\\").Files(\"a\").|\nEnd Sub\n", []string{"OpenAsTextStream", "ParentFolder"}, nil},
		{"Dictionary の Item は Variant", "Sub A()\n    Dim d As New Scripting.Dictionary\n    d(\"k\").|\nEnd Sub\n", nil, nil},
		{"CATIA 型は DB なしでは候補なし", "Sub A()\n    Dim d As PartDocument\n    d.|\nEnd Sub\n", nil, nil},
	})
}

// --- スキャナ ---

func chainText(ch MemberChain) string {
	var parts []string
	for _, s := range ch.Segs {
		p := s.Name + strings.Repeat("()", s.Calls)
		if s.Bang {
			p = "!" + p
		}
		parts = append(parts, p)
	}
	s := strings.Join(parts, ".")
	if ch.LeadingDot {
		s = "." + s
	}
	return s
}

func TestScanChainBackward(t *testing.T) {
	for _, tc := range []struct {
		input string
		want  string // "." の左側の連鎖。"" は読めない
	}{
		{"partDocument1.", "partDocument1"},
		{"partDocument1.Part.", "partDocument1.Part"},
		{"a.b.c.", "a.b.c"},
		{"部品.", "部品"},
		{"x = 部品.Part.", "部品.Part"},
		{"Debug.Print \"あ\": doc.", "doc"},
		{"x = a.Item(1).Name.", "a.Item().Name"},
		{"x = a.Item(f(1, (2))).", "a.Item()"},
		{"x = Left$(s, 1).", "Left()"},
		{"With d: x = .Part.", ".Part"},
		{"x = a!key.", "a.!key"},
		{"x = a. b.", "a.b"},
		{"x = 1.", ""},
		{"x = (a).", ""},
	} {
		masked, _, _ := maskCode(tc.input)
		got := ""
		if ch, ok := scanChainBackward(masked, len(masked)-1); ok {
			got = chainText(ch)
		}
		if got != tc.want {
			t.Errorf("%q: got %q, want %q", tc.input, got, tc.want)
		}
	}
}

func TestMaskCode(t *testing.T) {
	for _, tc := range []struct {
		in, want string
		str, cmt bool
	}{
		{`x = "a.b" & c`, `x = "   " & c`, false, false},
		{`x = "a""b`, `x = " "" `, true, false},
		{`x = 1 ' c.d`, `x = 1      `, false, true},
		{`Rem c.d`, `       `, false, true},
		{`x = 1: Rem c`, `x = 1:      `, false, true},
		{`Remark = 1`, `Remark = 1`, false, false},
		{`x = "あ" & d`, `x = "   " & d`, false, false},
	} {
		got, s, c := maskCode(tc.in)
		if got != tc.want || s != tc.str || c != tc.cmt {
			t.Errorf("maskCode(%q) = (%q, %v, %v), want (%q, %v, %v)", tc.in, got, s, c, tc.want, tc.str, tc.cmt)
		}
	}
}

func TestLogicalLine(t *testing.T) {
	lines := splitLines("x = d _\r\n    .Part _\n    + 1\ny = 2")
	ll := logicalLineAt(lines, 1)
	if ll.Text != "x = d.Part + 1" || ll.First != 0 || ll.Last != 2 {
		t.Fatalf("論理行 = %+v", ll)
	}
	off := ll.Offset(1, len("    .Pa"))
	if ll.Text[:off] != "x = d.Pa" {
		t.Errorf("Offset = %d (%q)", off, ll.Text[:off])
	}
	if line, col := ll.Physical(strings.Index(ll.Text, "Part")); line != 1 || col != 5 {
		t.Errorf("Physical = (%d,%d), want (1,5)", line, col)
	}
	if got := logicalLineAt(lines, 3); got.Text != "y = 2" {
		t.Errorf("独立した行 = %+v", got)
	}
	if !hasContinuation("x = a _") || hasContinuation("x = a_") || hasContinuation(`x = "_"`) || !hasContinuation("x = a\u3000_  ") {
		t.Error("hasContinuation の判定が誤っている")
	}
}

// --- ホバー ---

func hoverAt(t *testing.T, srcWithCaret string) string {
	t.Helper()
	src, line, char := caretPos(t, srcWithCaret)
	return ResolveHover(newTestDoc(t, src), line, char)
}

func TestHoverTable(t *testing.T) {
	const procSrc = "' 部品を処理する\nPublic Function Run(ByVal doc As PartDocument, Optional n As Long = 1) As Boolean\n    Dim x\n    Const LIMIT As Long = 5\n"
	for _, tc := range []struct {
		name, src, want string
	}{
		{"変数", "Sub A()\n    Dim d As PartDocument\n    x = d|\nEnd Sub\n", "```vba\nDim d As PartDocument\n```"},
		{"型のない変数", procSrc + "    y = x|\nEnd Function\n", "```vba\nDim x\n```"},
		{"引数", procSrc + "    y = do|c\nEnd Function\n", "```vba\n(引数) ByVal doc As PartDocument\n```"},
		{"Optional 引数", procSrc + "    y = n|\nEnd Function\n", "```vba\n(引数) Optional n As Long = 1\n```"},
		{"ローカル定数", procSrc + "    y = LIM|IT\nEnd Function\n", "```vba\nConst LIMIT As Long = 5\n```"},
		{"プロシージャ（ドキュメント付き）", procSrc + "    y = Ru|n(Nothing)\nEnd Function\n", "```vba\nPublic Function Run(ByVal doc As PartDocument, Optional n As Long = 1) As Boolean\n```\n\n部品を処理する"},
		{"チェーン", "Sub A()\n    Dim d As PartDocument\n    x = d.Part.Na|me\nEnd Sub\n", "```vba\nName As String\n```"},
		{"With の中", "Sub A()\n    Dim d As PartDocument\n    With d.Part\n        x = .Na|me\n    End With\nEnd Sub\n", "```vba\nName As String\n```"},
		{"メソッド", "Sub A()\n    Dim b As Bodies\n    Set x = b.It|em(1)\nEnd Sub\n", "```vba\nItem(iIndex As Variant) As Body\n```"},
		{"モジュール定数", "Public Const MAX_N As Long = 10 ' 上限\nSub A()\n    x = MAX|_N\nEnd Sub\n", "```vba\nPublic Const MAX_N As Long = 10\n```\n\n上限"},
		{"文字列内", "Sub A()\n    Dim d As PartDocument\n    x = \"d|\"\nEnd Sub\n", ""},
		{"コメント内", "Sub A()\n    Dim d As PartDocument\n    ' d|\nEnd Sub\n", ""},
		{"未知の識別子", "Sub A()\n    x = zz|z\nEnd Sub\n", ""},
	} {
		t.Run(tc.name, func(t *testing.T) {
			lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
				if got := hoverAt(t, src); got != tc.want {
					t.Errorf("ホバー = %q, want %q", got, tc.want)
				}
			})
		})
	}
}

// --- 診断 ---

func diagsOf(t *testing.T, src string) []Diagnostic {
	t.Helper()
	return BuildDiagnostics(newTestDoc(t, src))
}

func TestDiagnosticsTable(t *testing.T) {
	for _, tc := range []struct {
		name string
		src  string
		want []Range // 期待する範囲（順不同ではなく出現順）
	}{
		// TLB のメンバーの戻り値は基底型のことが多い（Part.UserSurfaces が Collection など）ため、経由した型では診断しない
		{"TLB のチェーンの先は対象外", "Sub A()\n    Dim d As PartDocument\n    d.Part.Bogus\n    d.Part.Bodies(1).Bogus\nEnd Sub\n", nil},
		{"宣言型への直接アクセス", "Sub A()\n    Dim d As PartDocument\n    d.Bogus.Name\nEnd Sub\n",
			[]Range{{Position{2, 6}, Position{2, 11}}}},
		{"組み込み型のチェーン", "Sub A()\n    Dim fso As New Scripting.FileSystemObject\n    fso.GetFolder(\"a\").Files(1).Bogus\nEnd Sub\n",
			[]Range{{Position{2, uint32(u16(`    fso.GetFolder("a").Files(1).`))}, Position{2, uint32(u16(`    fso.GetFolder("a").Files(1).Bogus`))}}}},
		{"Excel は対象外", "Sub A()\n    Set xl = CreateObject(\"Excel.Application\")\n    xl.Workbooks.Add\nEnd Sub\n", nil},
		{"Object は対象外", "Sub A()\n    Dim o As Object\n    o.Foo\nEnd Sub\n", nil},
		{"AnyObject は対象外", "Sub A()\n    Dim a As AnyObject\n    a.Foo\nEnd Sub\n", nil},
		{"With の中", "Sub A()\n    Dim d As PartDocument\n    With d\n        .Bogus\n        .Part.Name = 1\n    End With\nEnd Sub\n",
			[]Range{{Position{3, 9}, Position{3, 14}}}},
		{"行継続の 2 行目", "Sub A()\n    Dim d As PartDocument\n    x = d _\n      .Bogus\nEnd Sub\n",
			[]Range{{Position{3, 7}, Position{3, 12}}}},
		{"文字列とコメントは対象外", "Sub A()\n    Dim d As PartDocument\n    x = \"d.Bogus\" ' d.Bogus\nEnd Sub\n", nil},
		{"配列は対象外", "Sub A()\n    Dim a(3) As PartDocument\n    x = a.Bogus\nEnd Sub\n", nil},
		{"配列の要素", "Sub A()\n    Dim a(3) As PartDocument\n    x = a(0).Bogus\nEnd Sub\n",
			[]Range{{Position{2, 13}, Position{2, 18}}}},
		{"Collection の Item", "Sub A()\n    Dim c As New Collection\n    c.Add 1\n    c.Bogus\nEnd Sub\n",
			[]Range{{Position{3, 6}, Position{3, 11}}}},
		{"数値リテラル", "Sub A()\n    x = 1.5 + .5\nEnd Sub\n", nil},
	} {
		t.Run(tc.name, func(t *testing.T) {
			lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
				diags := diagsOf(t, src)
				if len(diags) != len(tc.want) {
					t.Fatalf("診断 = %+v, want %d 件", diags, len(tc.want))
				}
				for i, w := range tc.want {
					if diags[i].Range != w {
						t.Errorf("診断[%d] の範囲 = %+v, want %+v（%s）", i, diags[i].Range, w, diags[i].Message)
					}
				}
			})
		})
	}
	if d := diagsOf(t, "Sub A()\n    Dim fso As New Scripting.FileSystemObject\n    fso.GetFile(\"a\").Bogus\n    Dim p As PartDocument\n    p.Zzz\nEnd Sub\n"); len(d) != 2 ||
		d[0].Message != "'Scripting.File' に 'Bogus' は存在しません" || d[1].Message != "'PartDocument' に 'Zzz' は存在しません" {
		t.Errorf("メッセージ = %+v", d)
	}
}

// --- セマンティックトークン ---

func TestSemanticTokensScopeAndMembers(t *testing.T) {
	src := "Dim d As Part\n" + // 0
		"Sub A()\n" + // 1
		"    x = y.d\n" + // 2: メンバー名の d は対象外
		"    d.Update\n" + // 3: モジュール変数 d
		"End Sub\n" + // 4
		"Sub B()\n" + // 5
		"    Dim d As String\n" + // 6
		"    d = \"a\"\n" + // 7: ローカルの String が隠す
		"End Sub\n"
	lfAndCRLF(t, src, func(t *testing.T, src string) {
		toks := decodeSemanticTokens(t, BuildSemanticTokens(newTestDoc(t, src)))
		assertTokens(t, toks, []absToken{
			{0, 4, 1, tokenTypeVariable, tokenModDeclaration},
			{0, 9, 4, tokenTypeType, 0},
			{3, 4, 1, tokenTypeVariable, 0},
		})
	})
}
