package main

import (
	"database/sql"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// tlbV2Schema は tools/tlb-extractor/schema.sql（抽出器と同じスキーマ定義）を読む。
func tlbV2Schema(t *testing.T) string {
	t.Helper()
	b, err := os.ReadFile(filepath.Join("..", "tlb-extractor", "schema.sql"))
	if err != nil {
		t.Fatalf("schema.sql を読めない: %v", err)
	}
	return string(b)
}

// tlbV2TestData はスキーマ v2 のテスト用データ。
// CATBaseUnknown ← CATBaseDispatch ← AnyObject ← {Application, Document ← {PartDocument, ProductDocument}, Collection ← {Documents, Bodies}, Part, Body, InstanceFactory, …}
const tlbV2TestData = `
INSERT INTO meta(key, value) VALUES ('schema_version', '2'), ('doc_source', 'chm-summary');
INSERT INTO libraries(id, name, doc) VALUES (1, 'INFITF', 'CATIA V5 InfInterfaces Object Library'), (2, 'MECMOD', '');
INSERT INTO interfaces(id, name, parent_id, library_id, kind, flags, doc) VALUES
    (1, 'CATBaseUnknown', NULL, 1, 'dispatch', 64, ''),
    (2, 'CATBaseDispatch', 1, 1, 'dispatch', 64, ''),
    (3, 'AnyObject', 2, 1, 'dispatch', 64, 'Represents the base object.'),
    (4, 'Application', 3, 1, 'dispatch', 64, 'Represents the current CNext application.'),
    (5, 'Document', 3, 1, 'dispatch', 64, 'Represents the document.'),
    (6, 'PartDocument', 5, 2, 'dispatch', 64, ''),
    (7, 'ProductDocument', 5, 1, 'dispatch', 64, ''),
    (8, 'Collection', 3, 1, 'dispatch', 64, ''),
    (9, 'Documents', 8, 1, 'dispatch', 64, ''),
    (10, 'Part', 3, 2, 'dispatch', 64, ''),
    (11, 'Bodies', 8, 2, 'dispatch', 64, ''),
    (12, 'Body', 3, 2, 'dispatch', 64, ''),
    (13, 'InstanceFactory', 3, 2, 'dispatch', 64, ''),
    (14, 'HiddenThing', 3, 1, 'dispatch', 80, ''),
    (15, 'TreeVizManipSettingAtt', 3, 1, 'dispatch', 64, ''),
    (16, 'Product', 3, 1, 'dispatch', 64, '');
INSERT INTO methods(id, interface_id, name, return_type, flags, vararg, doc) VALUES
    (1, 4, 'Quit', 'void', 0, 0, 'Exits the application.'),
    (2, 4, 'FileSelectionBox', 'String', 0, 0, 'Displays a modal file selection box.'),
    (3, 4, 'HiddenMethod', 'void', 64, 0, ''),
    (4, 4, 'Format', 'String', 0, 1, ''),
    (5, 9, 'Add', 'Document', 0, 0, 'Creates a Document.'),
    (6, 9, 'Item', 'Document', 0, 0, ''),
    (7, 11, 'Item', 'Body', 0, 0, ''),
    (8, 5, 'Close', 'void', 0, 0, ''),
    (9, 15, 'GetTypeInfo', 'String', 0, 0, ''),
    (10, 4, 'StartCommand', 'void', 0, 0, '');
INSERT INTO parameters(method_id, name, type, position, flags, default_value, doc) VALUES
    (2, 'iTitle', 'String', 0, 1, NULL, 'The title.'),
    (2, 'iExtension', 'String', 1, 1, NULL, ''),
    (2, 'iMode', 'CatFileSelectionMode', 2, 1, NULL, 'The mode.'),
    (4, 'iFormat', 'String', 0, 49, '"x"', ''),
    (4, 'iArgs', 'Variant', 1, 1, NULL, ''),
    (5, 'docType', 'String', 0, 1, NULL, 'The type of the document.'),
    (6, 'iIndex', 'Variant', 0, 1, NULL, ''),
    (7, 'iIndex', 'Variant', 0, 1, NULL, ''),
    (10, 'iCommandId', 'String', 0, 1, NULL, '');
INSERT INTO properties(interface_id, name, type, access, flags, doc) VALUES
    (3, 'Name', 'String', 3, 0, 'Returns or sets the name.'),
    (3, 'Parent', 'AnyObject', 1, 0, ''),
    (3, 'Application', 'Application', 1, 0, ''),
    (4, 'Documents', 'Documents', 1, 0, 'Returns the collection of documents.'),
    (4, 'ActiveDocument', 'Document', 1, 0, 'Returns the active document.'),
    (4, 'Caption', 'String', 3, 0, ''),
    (4, 'HiddenProp', 'Long', 1, 64, ''),
    (4, 'PaperSize', 'CatPaperSize', 3, 0, ''),
    (5, 'FullName', 'String', 1, 0, ''),
    (6, 'Part', 'Part', 1, 0, 'Returns the root part.'),
    (7, 'Product', 'Product', 1, 0, ''),
    (8, 'Count', 'Long', 1, 0, ''),
    (10, 'Bodies', 'Bodies', 1, 0, ''),
    (13, 'InstantiationMode', 'String', 2, 0, 'Sets the mode of instantiation.');
INSERT INTO enums(id, name, library_id, doc) VALUES (1, 'CatPaperSize', 1, 'Paper size.'), (2, 'CatFileSelectionMode', 1, '');
INSERT INTO enum_members(enum_id, name, value, doc) VALUES
    (1, 'catPaperLetter', 0, 'The US Letter format'),
    (1, 'catPaperA4', 6, 'The A4 ISO format'),
    (2, 'CatFileSelectionModeOpen', 0, ''),
    (2, 'CatFileSelectionModeSave', 1, '');
`

// newTestTLBDatabaseV2 はスキーマ v2 のインメモリ DB を作る（DDL は schema.sql）。
func newTestTLBDatabaseV2(t *testing.T) *TLBDatabase {
	t.Helper()
	db, err := sql.Open("sqlite", ":memory:")
	if err != nil {
		t.Fatal(err)
	}
	db.SetMaxOpenConns(1) // :memory: は接続ごとに別 DB
	t.Cleanup(func() { _ = db.Close() })
	if _, err := db.Exec(tlbV2Schema(t)); err != nil {
		t.Fatalf("schema.sql: %v", err)
	}
	if _, err := db.Exec(tlbV2TestData); err != nil {
		t.Fatal(err)
	}
	return &TLBDatabase{db: db}
}

func newTestDocV2(t *testing.T, src string) *Document {
	t.Helper()
	return NewDocument("file:///test.bas", src, ParseSymbols(src), NewDefaultResolver(newTestTLBDatabaseV2(t), nil))
}

func memberNames(ms []MemberInfo) map[string]bool {
	out := map[string]bool{}
	for _, m := range ms {
		out[m.Name] = true
	}
	return out
}

func TestTLBSchemaDetection(t *testing.T) {
	v1 := newTestTLBDatabase(t).Schema()
	if v1.Version != 1 || v1.Hierarchy() || v1.HasTable("enums") || !v1.Has("interfaces", "parent_id") || v1.Has("methods", "doc") {
		t.Errorf("v1: %+v", v1)
	}
	v2 := newTestTLBDatabaseV2(t).Schema()
	if v2.Version != 2 || !v2.Hierarchy() || !v2.HasTable("ENUMS") || !v2.Has("Parameters", "Default_Value") {
		t.Errorf("v2: %+v", v2)
	}
}

// v2 のスキーマでも読み取り専用で開き、ファイルを一切変更しない。
func TestLoadTLBDatabaseReadOnlyV2(t *testing.T) {
	path := filepath.Join(t.TempDir(), "日本語 空白", "catia_api.db")
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		t.Fatal(err)
	}
	w, err := sql.Open("sqlite", path)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Exec(tlbV2Schema(t) + tlbV2TestData); err != nil {
		t.Fatal(err)
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	before := stampOf(t, path)
	db, err := LoadTLBDatabase(path)
	if err != nil {
		t.Fatal(err)
	}
	if v := db.Schema().Version; v != 2 {
		t.Errorf("schema version = %d", v)
	}
	r := NewDefaultResolver(db, nil)
	if ms := r.Members(TypeRef{Name: "PartDocument"}); !memberNames(ms)["Close"] || !memberNames(ms)["Name"] {
		t.Errorf("PartDocument の継承込みのメンバー = %v", memberNames(ms))
	}
	db.Close()
	if after := stampOf(t, path); after != before {
		t.Errorf("DB が変更された: before=%+v after=%+v", before, after)
	}
	assertNoSideFiles(t, path)
}

func TestTLBV2Members(t *testing.T) {
	p := newTLBProvider(newTestTLBDatabaseV2(t)).(*TLBProvider)
	app := TypeRef{Name: "Application"}
	got := memberNames(p.Members(app))
	for _, want := range []string{"Quit", "Documents", "Name", "Parent", "Caption"} {
		if !got[want] {
			t.Errorf("Application.%s が無い: %v", want, got)
		}
	}
	for _, hidden := range []string{"HiddenMethod", "HiddenProp"} {
		if got[hidden] {
			t.Errorf("hidden の %s が補完用の一覧にある", hidden)
		}
		if _, ok := p.Member(app, hidden); !ok {
			t.Errorf("hidden の %s が Member で引けない", hidden)
		}
	}
	if m, _ := p.Member(app, "Documents"); m.Doc != "Returns the collection of documents." || !m.ReadOnly || m.Owner.Name != "Application" {
		t.Errorf("Documents = %+v", m)
	}
	if m, _ := p.Member(app, "Caption"); m.ReadOnly {
		t.Errorf("Caption は読み書きできる: %+v", m)
	}
	if m, _ := p.Member(app, "name"); m.Owner.Name != "AnyObject" || m.Doc == "" {
		t.Errorf("継承した Name = %+v", m)
	}
	m, _ := p.Member(TypeRef{Name: "InstanceFactory"}, "InstantiationMode")
	if !strings.HasPrefix(m.Doc, "（書き込み専用）") || m.ReadOnly || m.ReturnType != "String" {
		t.Errorf("書き込み専用 = %+v", m)
	}
	fs, _ := p.Member(app, "FileSelectionBox")
	if len(fs.Params) != 3 || fs.Params[0].Doc != "The title." || fs.Params[0].Optional {
		t.Errorf("FileSelectionBox = %+v", fs)
	}
	f, _ := p.Member(app, "Format")
	if len(f.Params) != 2 || !f.Params[0].Optional || f.Params[0].Default != `"x"` || !f.Params[1].ParamArray || f.Params[0].ParamArray {
		t.Errorf("Format の引数 = %+v", f.Params)
	}
	if sig := f.Signature(); sig != `Format([iFormat As String = "x"], ParamArray iArgs() As Variant) As String` {
		t.Errorf("Format のシグネチャ = %s", sig)
	}
	// v2 は名前で隠さない（正当な同名メンバー）
	if _, ok := p.Member(TypeRef{Name: "TreeVizManipSettingAtt"}, "GetTypeInfo"); !ok {
		t.Error("TreeVizManipSettingAtt.GetTypeInfo が引けない")
	}
}

func TestTLBV2TypesAndLibraries(t *testing.T) {
	p := newTLBProvider(newTestTLBDatabaseV2(t)).(*TLBProvider)
	names := map[string]TypeInfo{}
	for _, ti := range p.Types() {
		names[ti.Ref.Name] = ti
	}
	if _, ok := names["HiddenThing"]; ok {
		t.Error("hidden のインターフェースが型一覧にある")
	}
	if ti := names["PartDocument"]; ti.Container != "MECMOD" || ti.Kind != TypeKindInterface {
		t.Errorf("PartDocument = %+v", ti)
	}
	if ti, ok := p.LookupType("Application"); !ok || ti.Doc != "Represents the current CNext application." || ti.Container != "INFITF" {
		t.Errorf("Application = %+v", ti)
	}
	if _, ok := p.LookupType("HiddenThing"); !ok {
		t.Error("hidden のインターフェースも型としては引ける")
	}
	for lib, want := range map[string]bool{"CATIA": true, "infitf": true, "MECMOD": true, "SPATypeLib": false, "Excel": false} {
		if got := p.OwnsLibrary(lib); got != want {
			t.Errorf("v2 OwnsLibrary(%q) = %v", lib, got)
		}
	}
	v1 := newTLBProvider(newTestTLBDatabase(t)).(*TLBProvider)
	for lib, want := range map[string]bool{"CATIA": true, "SPATypeLib": true, "Excel": false} {
		if got := v1.OwnsLibrary(lib); got != want {
			t.Errorf("v1 OwnsLibrary(%q) = %v", lib, got)
		}
	}
	r := NewDefaultResolver(newTestTLBDatabaseV2(t), nil)
	if ti, _, ok := r.Resolve(TypeRef{Lib: "MECMOD", Name: "PartDocument"}); !ok || ti.Ref.Name != "PartDocument" {
		t.Errorf("MECMOD.PartDocument が解決できない: %+v", ti)
	}
}

// v1 の DB でも CATBaseUnknown の restricted な関数は候補に出さない。
func TestTLBV1HidesRestrictedNames(t *testing.T) {
	db, err := sql.Open("sqlite", ":memory:")
	if err != nil {
		t.Fatal(err)
	}
	db.SetMaxOpenConns(1)
	defer db.Close()
	if _, err := db.Exec(`
CREATE TABLE interfaces (id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE, parent_id INTEGER);
CREATE TABLE methods (id INTEGER PRIMARY KEY, interface_id INTEGER NOT NULL, name TEXT NOT NULL, return_type TEXT NOT NULL DEFAULT '');
CREATE TABLE properties (id INTEGER PRIMARY KEY, interface_id INTEGER NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL DEFAULT '');
CREATE TABLE parameters (id INTEGER PRIMARY KEY, method_id INTEGER NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL DEFAULT '', position INTEGER NOT NULL DEFAULT 0);
INSERT INTO interfaces(id, name) VALUES (1, 'PartDocument');
INSERT INTO methods(interface_id, name, return_type) VALUES (1, 'QueryInterface', 'void'), (1, 'AddRef', 'Long'),
    (1, 'Destructor', 'void'), (1, 'IsAKindOf', 'Boolean'), (1, 'GetImpl', 'void'), (1, '_NewEnum', 'Object'), (1, 'Save', 'void');
INSERT INTO properties(interface_id, name, type) VALUES (1, 'Part', 'Part');
`); err != nil {
		t.Fatal(err)
	}
	p := newTLBProvider(&TLBDatabase{db: db}).(*TLBProvider)
	got := memberNames(p.Members(TypeRef{Name: "PartDocument"}))
	if len(got) != 2 || !got["Save"] || !got["Part"] {
		t.Errorf("v1 の PartDocument の候補 = %v", got)
	}
	if m, _ := p.Member(TypeRef{Name: "PartDocument"}, "Part"); m.ReadOnly {
		t.Errorf("v1 は access が無いので ReadOnly にしない: %+v", m)
	}
}

// v2resolver はテスト用 v2 DB を含む既定の resolver を作る。
func v2resolver(t *testing.T) *CompositeResolver {
	t.Helper()
	return NewDefaultResolver(newTestTLBDatabaseV2(t), nil)
}

// hoverWith は types を使って "|" の位置のホバーを返す。
func hoverWith(t *testing.T, types *CompositeResolver, srcWithCaret string) string {
	t.Helper()
	src, line, char := caretPos(t, srcWithCaret)
	return ResolveHover(NewDocument("file:///test.bas", src, ParseSymbols(src), types), line, char)
}

// diagMessages は src の診断を「重大度:メッセージ」の一覧で返す。
func diagMessages(t *testing.T, types *CompositeResolver, src string) []string {
	t.Helper()
	var out []string
	for _, d := range BuildDiagnostics(NewDocument("file:///test.bas", src, ParseSymbols(src), types)) {
		out = append(out, strings.Repeat("E", int(d.Severity)%2)+strings.Repeat("W", int(d.Severity)/2)+":"+d.Message)
	}
	return out
}

func TestTLBV2Enums(t *testing.T) {
	p := newTLBProvider(newTestTLBDatabaseV2(t)).(*TLBProvider)
	ti, ok := p.LookupType("catpapersize")
	if !ok || ti.Kind != TypeKindEnum || !ti.Complete || ti.Container != "INFITF" || ti.Doc != "Paper size." || ti.Ref != (TypeRef{Lib: "CATIA", Name: "CatPaperSize"}) {
		t.Fatalf("CatPaperSize = %+v", ti)
	}
	ms := p.Members(ti.Ref)
	if len(ms) != 2 || ms[0].Name != "catPaperA4" || ms[1].Name != "catPaperLetter" {
		t.Errorf("Members = %+v", ms)
	}
	m, ok := p.Member(ti.Ref, "CATPAPERA4")
	if !ok || m.Value != "6" || m.Kind != MemberKindEnumMember || m.Owner.Name != "CatPaperSize" || m.Doc != "The A4 ISO format" {
		t.Errorf("catPaperA4 = %+v", m)
	}
	if got, _ := p.MemberType(TypeRef{Name: "Application"}, "PaperSize"); got != ti.Ref {
		t.Errorf("Enum 型のプロパティの型 = %+v", got)
	}
	if c, ok := p.Constant("catpapera4"); !ok || c.Owner.Name != "CatPaperSize" {
		t.Errorf("Constant = %+v, %v", c, ok)
	}
	kinds := map[string]TypeKind{}
	for _, ti := range p.Types() {
		kinds[ti.Ref.Name] = ti.Kind
	}
	if kinds["CatPaperSize"] != TypeKindEnum || kinds["Application"] != TypeKindInterface {
		t.Errorf("Types = %v", kinds)
	}
	v1 := newTLBProvider(newTestTLBDatabase(t)).(*TLBProvider)
	if _, ok := v1.Constant("catPaperA4"); ok {
		t.Error("v1 に Enum は無い")
	}
}

func TestTLBV2EnumInEditor(t *testing.T) {
	r := v2resolver(t)
	items := completeAt(t, r, "Sub M()\n    x = CatPaperSize.|\nEnd Sub\n")
	if got := labelsOf(items); len(got) != 2 || !got["catPaperA4"] {
		t.Errorf("CatPaperSize. の補完 = %v", got)
	}
	types := labelsOf(completeAt(t, r, "Sub M()\n    Dim s As |\nEnd Sub\n"))
	if !types["CatPaperSize"] || types["HiddenThing"] {
		t.Errorf("型補完に CatPaperSize が無いか hidden がある")
	}
	if h := hoverWith(t, r, "Sub M()\n    x = catPape|rA4\nEnd Sub\n"); !strings.Contains(h, "CatPaperSize.catPaperA4 = 6") || !strings.Contains(h, "The A4 ISO format") {
		t.Errorf("Enum 定数のホバー = %q", h)
	}
	got := diagMessages(t, r, "Sub M()\n    x = CatPaperSize.catPaperA4 + CatPaperSize.catBogus\nEnd Sub\n")
	if len(got) != 1 || !strings.HasPrefix(got[0], "W:") || !strings.Contains(got[0], "catBogus") {
		t.Errorf("Enum の診断 = %v", got)
	}
}

func TestCATIAGlobal(t *testing.T) {
	p := newTLBProvider(newTestTLBDatabaseV2(t)).(*TLBProvider)
	if g, ok := p.Global("catia"); !ok || g != (TypeRef{Lib: "CATIA", Name: "Application"}) {
		t.Errorf("Global(CATIA) = %+v, %v", g, ok)
	}
	if gs := p.Globals(); len(gs) != 1 || gs[0].Name != "CATIA" {
		t.Errorf("Globals = %+v", gs)
	}
	if _, ok := newTLBProvider(newTestTLBDatabase(t)).(*TLBProvider).Global("CATIA"); ok {
		t.Error("Application の無い DB で CATIA が解決された")
	}

	r := v2resolver(t)
	got := labelsOf(completeAt(t, r, "Sub M()\n    CATIA.|\nEnd Sub\n"))
	for _, want := range []string{"Documents", "ActiveDocument", "Quit", "Name"} {
		if !got[want] {
			t.Errorf("CATIA. の補完に %s が無い: %v", want, got)
		}
	}
	if got := labelsOf(completeAt(t, r, "Sub M()\n    CATIA.ActiveDocument.|\nEnd Sub\n")); !got["FullName"] || !got["Close"] {
		t.Errorf("CATIA.ActiveDocument. の補完 = %v", got)
	}
	// 同名のローカル変数が優先される
	if got := labelsOf(completeAt(t, r, "Sub M()\n    Dim CATIA As Part\n    CATIA.|\nEnd Sub\n")); !got["Bodies"] || got["Documents"] {
		t.Errorf("ローカル変数 CATIA の補完 = %v", got)
	}
	var catia *CompletionItem
	items := completeAt(t, r, "Sub M()\n    |\nEnd Sub\n")
	for i := range items {
		if items[i].Label == "CATIA" {
			catia = &items[i]
		}
	}
	if catia == nil || catia.Kind != CompletionItemKindVariable || catia.Detail != "CATIA As Application" {
		t.Errorf("識別子補完の CATIA = %+v", catia)
	}
	if got := labelsOf(completeAt(t, nil, "Sub M()\n    |\nEnd Sub\n")); got["CATIA"] {
		t.Error("Application の無い DB で CATIA が識別子補完に出た")
	}
	if h := hoverWith(t, r, "Sub M()\n    CAT|IA.Quit\nEnd Sub\n"); !strings.Contains(h, "CATIA As Application") {
		t.Errorf("CATIA のホバー = %q", h)
	}
	src, line, char := caretPos(t, "Sub M()\n    s = CATIA.FileSelectionBox(\"a\", |\nEnd Sub\n")
	sh := BuildSignatureHelp(NewDocument("file:///test.bas", src, ParseSymbols(src), r), line, char)
	if sh == nil || !strings.HasPrefix(sh.Signatures[0].Label, "FileSelectionBox(iTitle As String") || sh.ActiveParameter != 1 {
		t.Errorf("CATIA.FileSelectionBox の signatureHelp = %+v", sh)
	}
	if got := diagMessages(t, r, "Sub M()\n    CATIA.Bogus\nEnd Sub\n"); len(got) != 1 || !strings.HasPrefix(got[0], "W:") {
		t.Errorf("CATIA.Bogus の診断 = %v", got)
	}
}

func TestHasMemberInDescendants(t *testing.T) {
	p := newTLBProvider(newTestTLBDatabaseV2(t)).(*TLBProvider)
	for _, tc := range []struct {
		typ, name string
		want      bool
	}{
		{"Document", "Part", true},    // PartDocument にある
		{"Document", "product", true}, // ProductDocument にある
		{"Document", "Name", true},    // AnyObject（祖先）にある
		{"Document", "Bodies", false}, // Part にあるが Document の系統ではない
		{"AnyObject", "Bodies", true}, // Part は AnyObject の派生型
		{"Collection", "Item", true},  // Documents / Bodies にある
		{"PartDocument", "Product", false},
		{"Application", "HiddenMethod", true},
		{"CatPaperSize", "catBogus", false},
		{"NoSuchType", "Name", false},
	} {
		if got := p.HasMemberInDescendants(TypeRef{Name: tc.typ}, tc.name); got != tc.want {
			t.Errorf("HasMemberInDescendants(%s, %s) = %v", tc.typ, tc.name, got)
		}
	}
	if !p.HasHierarchy() || newTLBProvider(newTestTLBDatabase(t)).(*TLBProvider).HasHierarchy() {
		t.Error("HasHierarchy が v2 で true、v1 で false でない")
	}
}

func TestDiagnosticsWithHierarchy(t *testing.T) {
	r := v2resolver(t)
	for _, tc := range []struct {
		body string
		want []string
	}{
		{"Dim d As Document\n    x = d.Part.Bodies.Count", nil},
		{"Dim d As Document\n    x = d.Bodies", []string{"W:'Document' に 'Bodies' は存在しません（派生型を含む）"}},
		{"x = CATIA.ActiveDocument.Part.Name", nil},
		{"x = CATIA.Documents.Item(1).Part.Bodies.Item(1).Name", nil},
		{"x = CATIA.ActiveDocument.Bogus", []string{"W:'CATIA.Document' に 'Bogus' は存在しません（派生型を含む）"}},
		{"x = CATIA.ActiveDocument.Bodies", []string{"W:'CATIA.Document' に 'Bodies' は存在しません（派生型を含む）"}},
		{"CATIA.HiddenMethod", nil},
		{"Dim f As InstanceFactory\n    f.InstantiationMode = \"After\"", nil},
		{"Dim s As String\n    x = s.Bogus", nil},
		{"Dim c As Collection\n    c.Bogus", []string{"E:'Collection' に 'Bogus' は存在しません"}},
	} {
		got := diagMessages(t, r, "Sub M()\n    "+tc.body+"\nEnd Sub\n")
		if strings.Join(got, "|") != strings.Join(tc.want, "|") {
			t.Errorf("%q: 診断 = %v, want %v", tc.body, got, tc.want)
		}
	}
}
