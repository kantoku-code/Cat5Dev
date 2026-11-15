package main

import (
	"strings"
	"testing"
)

func TestEmbeddedBuiltinsParse(t *testing.T) {
	s, err := parseBuiltinSet(builtinJSON)
	if err != nil {
		t.Fatalf("vba_builtins.json の解析に失敗: %v", err)
	}
	for _, lib := range []string{"VBA", "Scripting", "MSForms"} {
		if s.Library(lib) == nil {
			t.Fatalf("ライブラリ %s がない", lib)
		}
	}
	if s.Library("Unknown") != nil {
		t.Error("未知のライブラリは nil")
	}
	_ = embeddedBuiltins() // panic しないこと

	for _, lib := range s.libs {
		for _, idx := range lib.types {
			for _, m := range idx.members {
				if idx.info.Kind == TypeKindEnum {
					if m.Kind != MemberKindEnumMember || m.Value == "" || m.Doc == "" {
						t.Errorf("%s.%s: Kind = %v, Value = %q, Doc = %q", idx.info.Ref.Name, m.Name, m.Kind, m.Value, m.Doc)
					}
				} else if m.Kind != MemberKindMethod && m.Kind != MemberKindProperty {
					t.Errorf("%s.%s: Kind = %v", idx.info.Ref.Name, m.Name, m.Kind)
				}
			}
		}
	}

	vba := s.libs["vba"]
	if n := len(vba.Members(TypeRef{Name: "Collection"})); n != 4 {
		t.Errorf("Collection のメンバー数 = %d, want 4", n)
	}
	sc := s.libs["scripting"]
	if n := len(sc.Members(TypeRef{Name: "Dictionary"})); n != 10 {
		t.Errorf("Dictionary のメンバー数 = %d, want 10", n)
	}
	for _, name := range []string{"FileSystemObject", "TextStream", "File", "Folder", "Files", "Folders", "Drive", "Drives"} {
		if _, ok := sc.LookupType(name); !ok {
			t.Errorf("Scripting.%s がない", name)
		}
	}
	if ti, ok := vba.LookupType("string"); !ok || ti.Kind != TypeKindPrimitive || ti.Ref.Name != "String" {
		t.Errorf("VBA.String = %+v, %v", ti, ok)
	}
}

func TestBuiltinProviderFromJSON(t *testing.T) {
	const js = `{"version":1,"libraries":[{"name":"Lib","aliases":["L2"],"primitives":["Long"],
"types":[{"name":"Foo","kind":"class","complete":true,"members":[
 {"name":"Bar","kind":"property","type":"Foo"},
 {"name":"Baz","kind":"method","type":"Long"},
 {"name":"Ext","kind":"method","type":"Other"}]}]}]}`
	p, err := NewBuiltinProviderFromJSON([]byte(js), "Lib")
	if err != nil {
		t.Fatal(err)
	}
	if !p.OwnsLibrary("l2") || !p.OwnsLibrary("LIB") || p.OwnsLibrary("x") {
		t.Error("OwnsLibrary が別名・大文字小文字を扱えていない")
	}
	if got, _ := p.MemberType(TypeRef{Name: "Foo"}, "bar"); got != (TypeRef{Lib: "Lib", Name: "Foo"}) {
		t.Errorf("同じライブラリの型は Lib を付ける: %+v", got)
	}
	if got, _ := p.MemberType(TypeRef{Name: "Foo"}, "Baz"); got != (TypeRef{Lib: "VBA", Name: "Long"}) {
		t.Errorf("primitive は VBA を付ける: %+v", got)
	}
	if got, _ := p.MemberType(TypeRef{Name: "Foo"}, "Ext"); got != (TypeRef{Name: "Other"}) {
		t.Errorf("外部の型は Lib なし: %+v", got)
	}

	for _, bad := range []string{
		`{"libraries":[{"name":"L","types":[{"name":"T","kind":"bogus","members":[]}]}]}`,
		`{"libraries":[{"name":"L","types":[{"name":"T","kind":"class","members":[{"name":"A","kind":"field"}]}]}]}`,
		`{"libraries":[{"name":"L","types":[{"name":"T","kind":"class","members":[{"name":"A","kind":"method"},{"name":"a","kind":"method"}]}]}]}`,
		`{`,
	} {
		if _, err := NewBuiltinProviderFromJSON([]byte(bad), "L"); err == nil {
			t.Errorf("不正な JSON がエラーにならない: %s", strings.TrimSpace(bad))
		}
	}
}

// 組み込み関数・定数・列挙・キーワードのデータ。
func TestBuiltinFunctionsAndConstants(t *testing.T) {
	b := embeddedBuiltins()
	vba := b.Library("VBA").(*BuiltinProvider)
	if n := len(vba.Functions()); n != 123 {
		t.Errorf("関数の数 = %d, want 123", n)
	}
	enumMembers := 0
	for _, ti := range vba.Types() {
		if ti.Kind == TypeKindEnum {
			enumMembers += len(vba.Members(ti.Ref))
		}
	}
	if n := len(vba.Constants()) + enumMembers; n < 100 {
		t.Errorf("定数（独立 %d + 列挙メンバー %d）が 100 件未満", len(vba.Constants()), enumMembers)
	}
	if n := len(b.Keywords()); n < 60 {
		t.Errorf("キーワードの数 = %d, want 60 以上", n)
	}
	for _, f := range vba.Functions() {
		if f.Doc == "" || !hasJapanese(f.Doc) {
			t.Errorf("関数 %s の説明が日本語でない: %q", f.Name, f.Doc)
		}
	}
	for _, c := range vba.Constants() {
		if !hasJapanese(c.Doc) || c.Value == "" {
			t.Errorf("定数 %s: Doc = %q, Value = %q", c.Name, c.Doc, c.Value)
		}
	}
	for _, k := range b.Keywords() {
		if !hasJapanese(k.Doc) {
			t.Errorf("キーワード %s の説明が日本語でない: %q", k.Name, k.Doc)
		}
	}

	for _, tc := range []struct{ name, sig string }{
		{"MsgBox", "MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context]) As VbMsgBoxResult"},
		{"instr", "InStr([Start], String1, String2, [Compare As VbCompareMethod = vbBinaryCompare]) As Variant"},
		{"Mid", "Mid(String, Start As Long, [Length]) As Variant"},
		{"Mid$", "Mid$(String, Start As Long, [Length]) As String"},
		{"mid$", "Mid$(String, Start As Long, [Length]) As String"},
		{"UBound", "UBound(ArrayName, [Dimension As Long = 1]) As Long"},
		{"CStr", "CStr(Expression) As String"},
		{"Format", "Format(Expression, [Format], [FirstDayOfWeek As VbDayOfWeek = vbSunday], [FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1]) As Variant"},
		{"Array", "Array(ParamArray ArgList()) As Variant"},
		{"Now", "Now As Variant"},
		{"Kill", "Kill(PathName)"},
	} {
		m, ok := vba.Function(tc.name)
		if !ok {
			t.Errorf("関数 %s が引けない", tc.name)
			continue
		}
		if got := m.Signature(); got != tc.sig {
			t.Errorf("%s: Signature = %q, want %q", tc.name, got, tc.sig)
		}
	}
	for _, name := range []string{"Len$", "MsgBox$", "Bogus"} {
		if _, ok := vba.Function(name); ok {
			t.Errorf("%s は引けないはず", name)
		}
	}
	if c := vba.Constants(); len(c) == 0 {
		t.Fatal("定数がない")
	}
	found := map[string]string{}
	for _, c := range vba.Constants() {
		found[c.Name] = c.Signature()
	}
	if found["vbCrLf"] != "vbCrLf As String = Chr(13) & Chr(10)" {
		t.Errorf("vbCrLf = %q", found["vbCrLf"])
	}
	for _, tc := range []struct{ enum, member, sig string }{
		{"VbMsgBoxResult", "vbYes", "vbYes = 6"},
		{"VbCompareMethod", "vbTextCompare", "vbTextCompare = 1"},
		{"VbStrConv", "vbNarrow", "vbNarrow = 8"},
	} {
		m, ok := vba.Member(TypeRef{Name: tc.enum}, tc.member)
		if !ok || m.Signature() != tc.sig || m.Kind != MemberKindEnumMember {
			t.Errorf("%s.%s = %+v, %v", tc.enum, tc.member, m, ok)
		}
	}
	for _, g := range []struct{ name, typ string }{{"Err", "VBA.ErrObject"}, {"Debug", "VBA.Debug"}} {
		ref, ok := vba.Global(g.name)
		if !ok || ref.String() != g.typ {
			t.Errorf("Global(%s) = %v, %v", g.name, ref, ok)
		}
	}
	if len(vba.Globals()) != 2 {
		t.Errorf("Globals = %+v", vba.Globals())
	}
	for _, m := range []string{"Number", "Description", "Source", "Clear", "Raise"} {
		if _, ok := vba.Member(TypeRef{Name: "ErrObject"}, m); !ok {
			t.Errorf("ErrObject.%s がない", m)
		}
	}
	if _, ok := vba.Member(TypeRef{Name: "Debug"}, "Print"); !ok {
		t.Error("Debug.Print がない")
	}
	creatable := map[string]bool{}
	for _, lib := range []string{"VBA", "Scripting"} {
		for _, ti := range b.Library(lib).(*BuiltinProvider).Types() {
			if ti.Creatable {
				creatable[ti.Ref.Name] = true
			}
		}
	}
	if len(creatable) != 3 || !creatable["Collection"] || !creatable["Dictionary"] || !creatable["FileSystemObject"] {
		t.Errorf("creatable = %v", creatable)
	}
	if n := len(vba.Primitives()); n != 14 {
		t.Errorf("primitive の数 = %d", n)
	}
}

func hasJapanese(s string) bool {
	for _, r := range s {
		if r >= 0x3040 && r <= 0x9FFF {
			return true
		}
	}
	return false
}

func TestBuiltinJSONValidation(t *testing.T) {
	for _, bad := range []string{
		`{"libraries":[{"name":"L","types":[{"name":"E","kind":"enum","members":[{"name":"A","kind":"method"}]}]}]}`,
		`{"libraries":[{"name":"L","types":[{"name":"T","kind":"class","members":[{"name":"A","kind":"enumMember"}]}]}]}`,
		`{"libraries":[{"name":"L","types":[],"functions":[{"name":"F","kind":"proc","module":"M","doc":"説明"}]}]}`,
		`{"libraries":[{"name":"L","types":[],"functions":[{"name":"F","kind":"function","module":"M","doc":"説明"}]}]}`,
		`{"libraries":[{"name":"L","types":[],"functions":[{"name":"F","kind":"function","type":"Long","module":"M","doc":""}]}]}`,
		`{"libraries":[{"name":"L","types":[],"functions":[{"name":"F","kind":"sub","module":"M","doc":"a"},{"name":"f","kind":"sub","module":"M","doc":"b"}]}]}`,
		`{"libraries":[{"name":"L","types":[],"constants":[{"name":"C","type":"Long","value":"1","doc":""}]}]}`,
		`{"libraries":[{"name":"L","types":[{"name":"E","kind":"enum","members":[{"name":"C","kind":"enumMember","value":"1"}]}],"constants":[{"name":"c","type":"Long","value":"1","doc":"d"}]}]}`,
		`{"libraries":[{"name":"L","types":[]}],"keywords":[{"name":"Dim","usage":"bogus","doc":"d"}]}`,
		`{"libraries":[{"name":"L","types":[]}],"keywords":[{"name":"Dim","usage":"statement","doc":""}]}`,
	} {
		if _, err := NewBuiltinProviderFromJSON([]byte(bad), "L"); err == nil {
			t.Errorf("不正な JSON がエラーにならない: %s", bad)
		}
	}
}
