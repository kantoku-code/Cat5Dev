package main

import (
	"crypto/sha256"
	"database/sql"
	"encoding/hex"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

type fileStamp struct {
	size  int64
	mtime time.Time
	hash  string
}

func stampOf(t *testing.T, path string) fileStamp {
	t.Helper()
	fi, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	f, err := os.Open(path)
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	h := sha256.New()
	if _, err := io.Copy(h, f); err != nil {
		t.Fatal(err)
	}
	return fileStamp{size: fi.Size(), mtime: fi.ModTime(), hash: hex.EncodeToString(h.Sum(nil))}
}

func assertNoSideFiles(t *testing.T, path string) {
	t.Helper()
	for _, suffix := range []string{"-journal", "-wal", "-shm"} {
		if _, err := os.Stat(path + suffix); err == nil {
			t.Errorf("%s が作られている", path+suffix)
		}
	}
}

// LoadTLBDatabase は読み取り専用で開き、ファイルを一切変更しない。
func TestLoadTLBDatabaseReadOnly(t *testing.T) {
	dir := filepath.Join(t.TempDir(), "日本語 空白 #x %y")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "catia_api.db")

	// 通常モードで DB を作って閉じる
	w, err := sql.Open("sqlite", path)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := w.Exec(`
CREATE TABLE interfaces (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL UNIQUE, parent_id INTEGER);
CREATE TABLE methods (id INTEGER PRIMARY KEY AUTOINCREMENT, interface_id INTEGER NOT NULL, name TEXT NOT NULL, return_type TEXT NOT NULL DEFAULT '');
CREATE TABLE properties (id INTEGER PRIMARY KEY AUTOINCREMENT, interface_id INTEGER NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL DEFAULT '');
CREATE TABLE parameters (id INTEGER PRIMARY KEY AUTOINCREMENT, method_id INTEGER NOT NULL, name TEXT NOT NULL, type TEXT NOT NULL DEFAULT '', position INTEGER NOT NULL DEFAULT 0);
INSERT INTO interfaces(id, name, parent_id) VALUES (1, 'Base', NULL), (2, 'Derived', 1), (3, 'Loop', 3);
INSERT INTO properties(interface_id, name, type) VALUES (1, 'Name', 'String'), (2, 'Child', 'Derived'), (3, 'X', 'Long');
INSERT INTO methods(interface_id, name, return_type) VALUES (1, 'Run', 'Long'), (2, 'Run', 'Base');
`); err != nil {
		t.Fatal(err)
	}
	if err := w.Close(); err != nil {
		t.Fatal(err)
	}
	before := stampOf(t, path)

	db, err := LoadTLBDatabase(path)
	if err != nil {
		t.Fatalf("LoadTLBDatabase(%q): %v", path, err)
	}
	r := NewDefaultResolver(db, nil)
	ms := r.Members(TypeRef{Name: "Derived"})
	if len(ms) != 3 {
		t.Fatalf("継承込みのメンバー = %+v, want Child/Name/Run", ms)
	}
	if got, _ := r.MemberType(TypeRef{Name: "Derived"}, "Run"); got != (TypeRef{Lib: "CATIA", Name: "Base"}) {
		t.Errorf("派生側のメンバーが優先されていない: %+v", got)
	}
	if got := r.Members(TypeRef{Name: "Loop"}); len(got) != 1 {
		t.Errorf("循環した parent でメンバー = %+v", got)
	}
	if _, err := db.db.Exec(`INSERT INTO interfaces(name) VALUES ('X')`); err == nil || !strings.Contains(strings.ToLower(err.Error()), "readonly") {
		t.Errorf("書き込みが readonly エラーにならない: %v", err)
	}
	db.Close()

	if after := stampOf(t, path); after != before {
		t.Errorf("DB が変更された: before=%+v after=%+v", before, after)
	}
	assertNoSideFiles(t, path)
}

func TestLoadTLBDatabaseMissingFile(t *testing.T) {
	path := filepath.Join(t.TempDir(), "存在しない", "none.db")
	if db, err := LoadTLBDatabase(path); err == nil {
		db.Close()
		t.Fatal("存在しないパスがエラーにならない")
	}
	if _, err := os.Stat(path); err == nil {
		t.Fatal("存在しないパスにファイルが作られた")
	}
	path2 := filepath.Join(t.TempDir(), "none.db")
	if db, err := LoadTLBDatabase(path2); err == nil {
		db.Close()
		t.Fatal("存在しないファイルがエラーにならない")
	}
	if _, err := os.Stat(path2); err == nil {
		t.Fatal("存在しないファイルが作られた")
	}
}

// 同梱の bin/catia_api.db は混入前の原本と同じ内容で、LSP として使っても変更されない。
func TestBundledDatabaseUnchanged(t *testing.T) {
	path, err := filepath.Abs(filepath.Join("..", "..", "bin", "catia_api.db"))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := os.Stat(path); err != nil {
		t.Skipf("bin/catia_api.db がない: %v", err)
	}
	before := stampOf(t, path)
	if orig := filepath.Join("..", "tlb-extractor", "catia_api.db"); fileExists(orig) {
		if o := stampOf(t, orig); o.hash != before.hash {
			t.Errorf("bin/catia_api.db が tools/tlb-extractor/catia_api.db（原本）と異なる")
		}
	}

	// main と同じ経路（環境変数 → loadDB → resolver → handler）で使う
	t.Setenv("CAT5DEV_TLB_DB", path)
	db := loadDB()
	if db == nil {
		t.Fatal("loadDB が nil")
	}
	h := newLSPHandler(NewDefaultResolver(db, nil))
	const uri = "file:///m.bas"
	src := "Sub M()\n    Dim d As PartDocument\n    d.Part.Bodies.Item(1).Name\n    d.Bogus\n    d.\nEnd Sub\n"
	h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: src}})
	var cp CompletionParams
	cp.TextDocument.URI = uri
	cp.Position = Position{Line: 4, Character: 6}
	if items := h.completion(cp); len(items) == 0 {
		t.Error("PartDocument の補完が空")
	}
	_ = h.diagnostics(uri)
	_ = h.semanticTokens(SemanticTokensParams{TextDocument: TextDocumentIdentifier{URI: uri}})

	// 混入していた Dictionary / FileSystemObject は TLB 側にない（Scripting 組み込みで解決される）
	tlb := newTLBProvider(db)
	for _, name := range []string{"Dictionary", "FileSystemObject"} {
		if _, ok := tlb.LookupType(name); ok {
			t.Errorf("bin/catia_api.db に %s が混入している", name)
		}
	}
	db.Close()

	if after := stampOf(t, path); after != before {
		t.Errorf("bin/catia_api.db が変更された: before=%+v after=%+v", before, after)
	}
	assertNoSideFiles(t, path)
}

func fileExists(p string) bool {
	_, err := os.Stat(p)
	return err == nil
}
