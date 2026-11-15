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
	for _, lib := range []string{"VBA", "Scripting"} {
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
				if m.Kind != MemberKindMethod && m.Kind != MemberKindProperty {
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
