package main

import (
	"strings"
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

// fakeProvider はテスト用の最小プロバイダ。types のキーは小文字の型名。
type fakeProvider struct {
	lib       string
	qualified bool
	types     map[string]TypeInfo
	members   map[string][]MemberInfo // 小文字の型名 → メンバー
	mtypes    map[string]TypeRef      // "型.メンバー"（小文字）→ 型
}

func (f *fakeProvider) QualifiedOnly() bool { return f.qualified }
func (f *fakeProvider) OwnsLibrary(lib string) bool {
	return f.lib != "" && strings.EqualFold(lib, f.lib)
}
func (f *fakeProvider) LookupType(name string) (TypeInfo, bool) {
	ti, ok := f.types[strings.ToLower(name)]
	return ti, ok
}
func (f *fakeProvider) Members(t TypeRef) []MemberInfo { return f.members[strings.ToLower(t.Name)] }
func (f *fakeProvider) Member(t TypeRef, name string) (MemberInfo, bool) {
	for _, m := range f.Members(t) {
		if strings.EqualFold(m.Name, name) {
			return m, true
		}
	}
	return MemberInfo{}, false
}
func (f *fakeProvider) MemberType(t TypeRef, name string) (TypeRef, bool) {
	r, ok := f.mtypes[strings.ToLower(t.Name+"."+name)]
	return r, ok
}

func TestQualifiedOnlyProvider(t *testing.T) {
	mod := &fakeProvider{lib: moduleLib, qualified: true, types: map[string]TypeInfo{
		"part": {Ref: TypeRef{Lib: moduleLib, Name: "Part"}, Kind: TypeKindModule},
	}}
	r := NewDefaultResolver(newTestTLBDatabase(t), mod)
	ti, _, ok := r.Resolve(TypeRef{Name: "Part"})
	if !ok || ti.Ref.Lib == moduleLib {
		t.Fatalf("無修飾の Part は TLB に解決されるはず: %+v %v", ti, ok)
	}
	ti, _, ok = r.Resolve(TypeRef{Lib: moduleLib, Name: "Part"})
	if !ok || ti.Kind != TypeKindModule {
		t.Fatalf("修飾付きは標準モジュールに解決されるはず: %+v %v", ti, ok)
	}
	if r.TypeExists(TypeRef{Lib: moduleLib, Name: "Part"}) {
		t.Error("標準モジュールは型として存在しない扱い")
	}
	if got := (TypeRef{Lib: moduleLib, Name: "Main"}).String(); got != "Main" {
		t.Errorf("String() = %q, want Main", got)
	}
}

func TestBaseTypeMembers(t *testing.T) {
	base := &fakeProvider{lib: "Lib", types: map[string]TypeInfo{
		"b": {Ref: TypeRef{Lib: "Lib", Name: "B"}, Kind: TypeKindClass},
	}, members: map[string][]MemberInfo{
		"b": {{Name: "Hide", Kind: MemberKindMethod}, {Name: "Name", ReturnType: "String"}},
	}, mtypes: map[string]TypeRef{"b.name": {Lib: "VBA", Name: "String"}}}
	own := &fakeProvider{types: map[string]TypeInfo{
		"f": {Ref: TypeRef{Name: "F"}, Kind: TypeKindClass, Base: TypeRef{Lib: "Lib", Name: "B"}},
		// 自分自身を Base にしても無限に辿らない
		"loop": {Ref: TypeRef{Name: "Loop"}, Kind: TypeKindClass, Base: TypeRef{Name: "Loop"}},
	}, members: map[string][]MemberInfo{
		"f": {{Name: "Name", ReturnType: "Long"}, {Name: "Zed"}},
	}, mtypes: map[string]TypeRef{"f.name": {Lib: "VBA", Name: "Long"}}}
	r := NewCompositeResolver(own, base)
	var names []string
	for _, m := range r.Members(TypeRef{Name: "F"}) {
		names = append(names, m.Name+":"+m.ReturnType)
	}
	if got, want := strings.Join(names, ","), "Hide:,Name:Long,Zed:"; got != want {
		t.Errorf("Members = %s, want %s", got, want)
	}
	if _, ok := r.Member(TypeRef{Name: "F"}, "hide"); !ok {
		t.Error("Base のメンバーが引けない")
	}
	if mt, ok := r.MemberType(TypeRef{Name: "F"}, "Name"); !ok || mt.Name != "Long" {
		t.Errorf("同名は自分側を優先: %+v %v", mt, ok)
	}
	if r.Members(TypeRef{Name: "Loop"}) != nil {
		t.Error("循環した Base は空")
	}
}

func TestWithLocalVariadic(t *testing.T) {
	r := NewDefaultResolver(nil)
	if r.WithLocal() != r || r.WithLocal(nil, nil) != r {
		t.Error("足すものがなければ同じ resolver")
	}
	a := &fakeProvider{types: map[string]TypeInfo{"x": {Ref: TypeRef{Name: "X"}}}}
	b := &fakeProvider{types: map[string]TypeInfo{"x": {Ref: TypeRef{Name: "X"}, Doc: "b"}, "y": {Ref: TypeRef{Name: "Y"}}}}
	w := r.WithLocal(a, nil, b)
	if ti, p, ok := w.Resolve(TypeRef{Name: "X"}); !ok || p != a || ti.Doc != "" {
		t.Error("先に渡したものが優先")
	}
	if _, _, ok := w.Resolve(TypeRef{Name: "Y"}); !ok {
		t.Error("2 つ目も足される")
	}
}

// 型・関数・定数・グローバルの列挙（識別子補完・型補完用）。
func TestCompositeListers(t *testing.T) {
	r := newTestResolver(t)
	names := func(ts []TypeInfo) map[string]bool {
		m := map[string]bool{}
		for _, ti := range ts {
			m[ti.Ref.Name] = true
		}
		return m
	}
	all := names(r.Types())
	for _, w := range []string{"PartDocument", "Document", "Collection", "Dictionary", "ErrObject", "VbMsgBoxResult", "UserForm"} {
		if !all[w] {
			t.Errorf("Types に %s がない", w)
		}
	}
	if all["Long"] {
		t.Error("Types に primitive が含まれている")
	}
	tlb := r.TypesIn("CATIA")
	if len(tlb) != 7 || tlb[0].Ref.Name != "AnyObject" || tlb[0].Kind != TypeKindInterface {
		t.Errorf("TypesIn(CATIA) = %+v", tlb)
	}
	sc := names(r.TypesIn("scripting"))
	if !sc["Dictionary"] || !sc["FileSystemObject"] || sc["Collection"] || sc["PartDocument"] {
		t.Errorf("TypesIn(Scripting) = %v", sc)
	}
	if p := names(r.Primitives()); len(p) != 14 || !p["Long"] || !p["String"] {
		t.Errorf("Primitives = %v", p)
	}
	if m, ok := r.Function("Left$"); !ok || m.Name != "Left$" || m.ReturnType != "String" {
		t.Errorf("Function(Left$) = %+v, %v", m, ok)
	}
	if n := len(r.Functions()); n != 123 {
		t.Errorf("Functions = %d 件", n)
	}
	if n := len(r.Constants()); n != 11 {
		t.Errorf("Constants = %d 件", n)
	}
	if g := r.Globals(); len(g) != 2 || g[0].Name != "Debug" || g[1].Name != "Err" {
		t.Errorf("Globals = %+v", g)
	}
	var nilR *CompositeResolver
	if nilR.Types() != nil || nilR.TypesIn("VBA") != nil || nilR.Functions() != nil || nilR.ProjectMembers() != nil || nilR.ProjectModules() != nil {
		t.Error("nil の resolver は nil を返す")
	}
	if _, ok := nilR.Function("MsgBox"); ok {
		t.Error("nil の resolver は関数を引けない")
	}
	// DB なしでも組み込みは列挙できる
	if n := len(NewDefaultResolver(nil).TypesIn("CATIA")); n != 0 {
		t.Errorf("DB なしの TypesIn(CATIA) = %d 件", n)
	}
}
