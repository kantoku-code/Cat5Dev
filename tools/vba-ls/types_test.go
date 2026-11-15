package main

import (
	"testing"
)

func TestParseTypeName(t *testing.T) {
	for _, tc := range []struct {
		raw  string
		want TypeRef
	}{
		{"PartDocument", TypeRef{Name: "PartDocument"}},
		{"Scripting.Dictionary", TypeRef{Lib: "Scripting", Name: "Dictionary"}},
		{"Scripting . Dictionary", TypeRef{Lib: "Scripting", Name: "Dictionary"}},
		{"Scripting _\r\n    .Dictionary", TypeRef{Lib: "Scripting", Name: "Dictionary"}},
		{"[Long]", TypeRef{Name: "Long"}},
		{"String * 10", TypeRef{Name: "String"}},
		{"String()", TypeRef{Name: "String", Array: true}},
		{"INFITF.Application", TypeRef{Lib: "INFITF", Name: "Application"}},
		{"", TypeRef{}},
	} {
		if got := ParseTypeName(tc.raw); got != tc.want {
			t.Errorf("ParseTypeName(%q) = %+v, want %+v", tc.raw, got, tc.want)
		}
	}
}

func TestTypeRefFromProgID(t *testing.T) {
	for _, tc := range []struct {
		progID string
		want   TypeRef
	}{
		{"Scripting.Dictionary", TypeRef{Lib: "Scripting", Name: "Dictionary"}},
		{"Excel.Application.16", TypeRef{Lib: "Excel", Name: "Application"}},
		{"CATIA.PartDocument", TypeRef{Lib: "CATIA", Name: "PartDocument"}},
		{"Dictionary", TypeRef{Name: "Dictionary"}},
	} {
		if got := TypeRefFromProgID(tc.progID); got != tc.want {
			t.Errorf("TypeRefFromProgID(%q) = %+v, want %+v", tc.progID, got, tc.want)
		}
	}
}

func TestTypeRefString(t *testing.T) {
	if s := (TypeRef{Lib: "Scripting", Name: "Dictionary"}).String(); s != "Scripting.Dictionary" {
		t.Errorf("String = %q", s)
	}
	if s := (TypeRef{Name: "String", Array: true}).String(); s != "String()" {
		t.Errorf("String = %q", s)
	}
}

func TestResolveOrder(t *testing.T) {
	tdb := newTestTLBDatabase(t)
	if _, err := tdb.db.Exec(`INSERT INTO interfaces(name) VALUES ('TextStream'), ('Application')`); err != nil {
		t.Fatal(err)
	}
	r := NewDefaultResolver(tdb, nil)
	for _, tc := range []struct {
		ref     TypeRef
		wantLib string
		ok      bool
	}{
		{TypeRef{Name: "Collection"}, "VBA", true},                  // 無修飾は VBA が CATIA より優先
		{TypeRef{Lib: "INFITF", Name: "Collection"}, "CATIA", true}, // 修飾ありは担当プロバイダだけ
		{TypeRef{Lib: "vba", Name: "collection"}, "VBA", true},      // 大文字小文字無視
		{TypeRef{Name: "TextStream"}, "CATIA", true},                // CATIA が Scripting より優先
		{TypeRef{Lib: "Scripting", Name: "TextStream"}, "Scripting", true},
		{TypeRef{Name: "Dictionary"}, "Scripting", true},
		{TypeRef{Lib: "Excel", Name: "Application"}, "", false},      // 未知のライブラリは未解決
		{TypeRef{Lib: "Scripting", Name: "PartDocument"}, "", false}, // 担当外
		{TypeRef{Name: "Unknown"}, "", false},
		{TypeRef{}, "", false},
	} {
		ti, _, ok := r.Resolve(tc.ref)
		if ok != tc.ok || (ok && ti.Ref.Lib != tc.wantLib) {
			t.Errorf("Resolve(%+v) = (%+v, %v), want lib %q ok=%v", tc.ref, ti.Ref, ok, tc.wantLib, tc.ok)
		}
	}
}

func TestTypeExistsExcludesPrimitives(t *testing.T) {
	r := newTestResolver(t)
	for _, name := range []string{"String", "Long", "Variant", "Object", "Boolean"} {
		if r.TypeExists(TypeRef{Name: name}) {
			t.Errorf("TypeExists(%s) = true, want false（primitive）", name)
		}
	}
	for _, name := range []string{"PartDocument", "Collection", "Dictionary"} {
		if !r.TypeExists(TypeRef{Name: name}) {
			t.Errorf("TypeExists(%s) = false, want true", name)
		}
	}
	if !r.TypeExists(TypeRef{Name: "PartDocument", Array: true}) {
		t.Error("配列の TypeExists は要素型で判定する")
	}
	if r.Members(TypeRef{Name: "PartDocument", Array: true}) != nil {
		t.Error("配列の Members は nil")
	}
}

func TestMemberTypeQualification(t *testing.T) {
	r := newTestResolver(t)
	for _, tc := range []struct {
		owner  TypeRef
		member string
		want   TypeRef
		ok     bool
	}{
		{TypeRef{Name: "FileSystemObject"}, "OpenTextFile", TypeRef{Lib: "Scripting", Name: "TextStream"}, true},
		{TypeRef{Name: "Dictionary"}, "count", TypeRef{Lib: "VBA", Name: "Long"}, true},
		{TypeRef{Name: "Dictionary"}, "RemoveAll", TypeRef{}, false},
		{TypeRef{Name: "Part"}, "Bodies", TypeRef{Lib: "CATIA", Name: "Bodies"}, true},
		{TypeRef{Name: "PartDocument"}, "Name", TypeRef{Name: "String"}, true},
		{TypeRef{Name: "PartDocument"}, "Save", TypeRef{}, false},
		{TypeRef{Name: "Bodies"}, "Item", TypeRef{Lib: "CATIA", Name: "Body"}, true},
	} {
		got, ok := r.MemberType(tc.owner, tc.member)
		if ok != tc.ok || got != tc.want {
			t.Errorf("MemberType(%v, %s) = (%+v, %v), want (%+v, %v)", tc.owner, tc.member, got, ok, tc.want, tc.ok)
		}
	}
}

func TestDefaultMember(t *testing.T) {
	r := newTestResolver(t)
	for _, tc := range []struct {
		ref  TypeRef
		want string
	}{
		{TypeRef{Name: "Collection"}, "Item"},
		{TypeRef{Name: "Dictionary"}, "Item"},
		{TypeRef{Name: "Bodies"}, "Item"}, // Item メンバーから推定
		{TypeRef{Name: "PartDocument"}, ""},
	} {
		if got := r.DefaultMember(tc.ref); got != tc.want {
			t.Errorf("DefaultMember(%v) = %q, want %q", tc.ref, got, tc.want)
		}
	}
}

func TestCompositeIgnoresNilProviders(t *testing.T) {
	var typedNil *TLBProvider
	r := NewCompositeResolver(nil, typedNil, newTLBProvider(nil), NullProvider{})
	if len(r.providers) != 1 {
		t.Fatalf("providers = %d, want 1（NullProvider のみ）", len(r.providers))
	}
	if _, _, ok := r.Resolve(TypeRef{Name: "Collection"}); ok {
		t.Error("NullProvider だけでは何も解決しない")
	}
	if r.WithLocal(nil) != r {
		t.Error("WithLocal(nil) は同じ resolver を返す")
	}
}

func TestMemberSignature(t *testing.T) {
	r := NewDefaultResolver(nil, nil)
	m, ok := r.Member(TypeRef{Name: "Collection"}, "Add")
	if !ok {
		t.Fatal("Collection.Add がない")
	}
	if got, want := m.Signature(), "Add(Item As Variant, [Key As Variant], [Before As Variant], [After As Variant])"; got != want {
		t.Errorf("Signature = %q, want %q", got, want)
	}
	m, _ = r.Member(TypeRef{Name: "Dictionary"}, "Count")
	if got := m.Signature(); got != "Count As Long" {
		t.Errorf("Signature = %q", got)
	}
	m, _ = r.Member(TypeRef{Name: "Dictionary"}, "Item")
	if got := m.Signature(); got != "Item(Key As Variant) As Variant" {
		t.Errorf("Signature = %q", got)
	}
}
