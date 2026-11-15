package main

import (
	"strings"
	"testing"
)

func TestTypeHintExtraction(t *testing.T) {
	for _, tc := range []struct {
		decl, name, want string
	}{
		{"Dim v As Object '@As PartDocument", "v", "PartDocument"},
		{"Dim v As Object 'As Viewpoint3D", "v", "Viewpoint3D"},
		{"Dim v As Object ' As Part ", "v", "Part"},
		{"Dim v As Object 'as  MECMOD.Part", "v", "MECMOD.Part"},
		{"Dim v As Variant '@as Part", "v", "Part"},
		{"Dim v '@As Part", "v", "Part"},
		{"Private v As VBA.Object 'As Part", "v", "Part"},
		{"Dim v As Object 'As Part の説明", "v", ""}, // コメント全体がヒントの形でない
		{"Dim v As Object 'As of 2020", "v", ""},  // 型が 2 語
		{"Dim v As Long 'As Part", "v", ""},       // Object / Variant / 型なし以外
		{"Dim v As Scripting.Dictionary 'As Part", "v", ""},
		{"Dim a As Object, v As Object 'As Part", "v", ""}, // 複数の宣言子
		{"Dim v As Object ' 'As Part", "v", ""},
		{"Dim v As Object: x = \"'As Part\"", "v", ""},
	} {
		src := "Sub M()\n    " + tc.decl + "\nEnd Sub\n"
		sym := ParseSymbols(src).FindByName(tc.name, "Sub:M")
		if sym == nil {
			t.Fatalf("%q: シンボルが無い", tc.decl)
		}
		if got := sym.HintRef().String(); got != tc.want {
			t.Errorf("%q: ヒント = %q, want %q", tc.decl, got, tc.want)
		}
	}
	// モジュールレベルの Public・Static も対象
	table := ParseSymbols("Public m_v As Object 'As Part\nSub M()\n    Static s As Object '@As Body\nEnd Sub\n")
	if s := table.FindByName("m_v", "module"); s == nil || s.HintName != "Part" || s.Doc != "" {
		t.Errorf("Public の型ヒント = %+v", s)
	}
	if s := table.FindByName("s", "Sub:M"); s == nil || s.HintName != "Body" {
		t.Errorf("Static の型ヒント = %+v", s)
	}
	if m, typ, ok := plainMember(table.FindByName("m_v", "module"), false); !ok || typ.Name != "Part" || m.ReturnType != "Object" {
		t.Errorf("Public 変数のメンバーの型 = %+v / %+v", m, typ)
	}
}

func TestTypeHintInEditor(t *testing.T) {
	r := v2resolver(t)
	const head = "Sub M()\n    ' 部品文書\n    Dim v As Object '@As PartDocument\n    Dim u As Object 'As Bogus\n"
	if got := labelsOf(completeAt(t, r, head+"    v.|\nEnd Sub\n")); !got["Part"] || !got["Close"] {
		t.Errorf("型ヒントの補完 = %v", got)
	}
	if got := labelsOf(completeAt(t, r, head+"    v.Part.|\nEnd Sub\n")); !got["Bodies"] {
		t.Errorf("型ヒントの連鎖の補完 = %v", got)
	}
	if h := hoverWith(t, r, head+"    x = |v.Part\nEnd Sub\n"); h != "```vba\nDim v As Object  ' 型ヒント: PartDocument\n```\n\n部品文書" {
		t.Errorf("型ヒントのホバー = %q", h)
	}
	// 解決できないヒントは無視する（Object のまま）
	if got := completeAt(t, r, head+"    u.|\nEnd Sub\n"); len(got) != 0 {
		t.Errorf("未解決のヒントで補完が出た: %v", labelsOf(got))
	}
	got := diagMessages(t, r, head+"    v.Bogus\n    u.Bogus\n    x = v.Part.Name\nEnd Sub\n")
	if len(got) != 1 || !strings.Contains(got[0], "'PartDocument' に 'Bogus'") || !strings.HasPrefix(got[0], "W:") {
		t.Errorf("型ヒントの診断 = %v", got)
	}
}
