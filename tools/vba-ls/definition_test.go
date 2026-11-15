package main

import (
	"path/filepath"
	"strings"
	"testing"
)

// declPos は text の中で、行が linePrefix で始まる最初の行と、その行の name の UTF-16 位置を返す。
// linePrefix が name で終わればその位置、そうでなければ linePrefix の後で最初に現れる位置。
func declPos(t *testing.T, text, linePrefix, name string) (int, int) {
	t.Helper()
	for i, l := range splitLines(text) {
		if !strings.HasPrefix(l, linePrefix) {
			continue
		}
		if strings.HasSuffix(linePrefix, name) {
			return i, u16(linePrefix[:len(linePrefix)-len(name)])
		}
		j := strings.Index(l[len(linePrefix):], name)
		if j < 0 {
			t.Fatalf("%q に %q がない", l, name)
		}
		return i, u16(l[:len(linePrefix)+j])
	}
	t.Fatalf("%q で始まる行がない", linePrefix)
	return 0, 0
}

func loc(uri string, line, col int, name string) Location {
	return Location{URI: uri, Range: Range{
		Start: Position{Line: uint32(line), Character: uint32(col)},
		End:   Position{Line: uint32(line), Character: uint32(col + u16(name))},
	}}
}

func TestDefinition(t *testing.T) {
	files := map[string]string{
		"cls吉.cls_utf": "' 説明\nPublic Function 𠮷野() As String\nEnd Function\nPublic Sub B()\n    x = \"あ\": Dim 値 As Long\n    値 = 1\nEnd Sub\n",
	}
	for k, v := range wsFixture {
		files[k] = v
	}
	_, types, dir := newTestWorkspace(t, files)
	uriOf := func(f string) string { return pathToURI(filepath.Join(dir, f)) }
	clsA, mod1 := wsFixture["clsA.cls_utf"], wsFixture["Module1.bas_utf"]

	cl, cc := declPos(t, clsA, "Public Function ", "Clone")
	gl, gc := declPos(t, clsA, "Public Property Get Value", "Value")
	ll, lc := declPos(t, clsA, "Public Property Let Value", "Value")
	pl, pc := declPos(t, mod1, "Public Sub ", "Pub")
	el, ec := declPos(t, mod1, "Public Enum ", "EKind")
	yl, yc := declPos(t, files["cls吉.cls_utf"], "Public Function ", "𠮷野")

	for _, tc := range []struct {
		name, file, src string
		want            []Location
	}{
		{"他クラスのメソッド", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Cl|one\nEnd Sub\n", []Location{loc(uriOf("clsA.cls_utf"), cl, cc, "Clone")}},
		{"Property は Get と Let", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Val|ue = 1\nEnd Sub\n",
			[]Location{loc(uriOf("clsA.cls_utf"), gl, gc, "Value"), loc(uriOf("clsA.cls_utf"), ll, lc, "Value")}},
		{"As clsA はモジュール先頭", "Module1.bas_utf", "Sub T()\n    Dim o As cls|A\nEnd Sub\n", []Location{{URI: uriOf("clsA.cls_utf")}}},
		{"New clsA", "frmA.frm_utf", "Sub T()\n    Set o = New cl|sA\nEnd Sub\n", []Location{{URI: uriOf("clsA.cls_utf")}}},
		{"モジュール名", "clsA.cls_utf", "Sub T()\n    Mod|ule1.Pub\nEnd Sub\n", []Location{{URI: uriOf("Module1.bas_utf")}}},
		{"Module1.Pub", "clsA.cls_utf", "Sub T()\n    Module1.P|ub\nEnd Sub\n", []Location{loc(uriOf("Module1.bas_utf"), pl, pc, "Pub")}},
		{"非修飾の Pub", "clsA.cls_utf", "Sub T()\n    Pu|b\nEnd Sub\n", []Location{loc(uriOf("Module1.bas_utf"), pl, pc, "Pub")}},
		{"Public Enum", "clsA.cls_utf", "Sub T()\n    Dim k As EKi|nd\nEnd Sub\n", []Location{loc(uriOf("Module1.bas_utf"), el, ec, "EKind")}},
		{"TLB のメンバーは nil", "Module1.bas_utf", "Sub T()\n    Dim d As PartDocument\n    d.Pa|rt\nEnd Sub\n", nil},
		{"サロゲートペアの名前", "Module1.bas_utf", "Sub T()\n    Dim y As New cls吉\n    y.𠮷|野\nEnd Sub\n", []Location{loc(uriOf("cls吉.cls_utf"), yl, yc, "𠮷野")}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			src, line, char := caretPos(t, withFixture(tc.file, tc.src))
			doc := NewDocument(uriOf(tc.file), src, ParseSymbols(src), types)
			got := ResolveDefinition(doc, line, char)
			if len(got) != len(tc.want) {
				t.Fatalf("definition = %+v\nwant %+v", got, tc.want)
			}
			for i := range got {
				if got[i] != tc.want[i] {
					t.Errorf("[%d] = %+v\nwant %+v", i, got[i], tc.want[i])
				}
			}
		})
	}

	// 自モジュールのシンボル（日本語の文字列がある行の宣言）は現在のバッファの位置
	src := "Public Sub B()\n    x = \"あ\": Dim 値 As Long\n    値| = 1\nEnd Sub\n"
	clean, line, char := caretPos(t, src)
	doc := NewDocument(uriOf("cls吉.cls_utf"), clean, ParseSymbols(clean), types)
	got := ResolveDefinition(doc, line, char)
	want := loc(uriOf("cls吉.cls_utf"), 1, u16("    x = \"あ\": Dim "), "値")
	if len(got) != 1 || got[0] != want {
		t.Errorf("ローカル変数 = %+v, want %+v", got, want)
	}
}
