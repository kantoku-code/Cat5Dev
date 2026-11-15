package main

import (
	"database/sql"
	"strings"
	"testing"
	"unicode/utf16"
)

// newTestTLBDatabase はインメモリ SQLite で TLBDatabase を作る。
// スキーマは tools/tlb-extractor/extract.py の _create_schema と同じ 4 テーブル。
func newTestTLBDatabase(t *testing.T) *TLBDatabase {
	t.Helper()
	db, err := sql.Open("sqlite", ":memory:")
	if err != nil {
		t.Fatal(err)
	}
	// :memory: は接続ごとに別 DB になるため 1 接続に固定する
	db.SetMaxOpenConns(1)
	t.Cleanup(func() { _ = db.Close() })
	_, err = db.Exec(`
CREATE TABLE interfaces (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    name      TEXT    NOT NULL UNIQUE,
    parent_id INTEGER REFERENCES interfaces(id)
);
CREATE TABLE methods (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    interface_id INTEGER NOT NULL REFERENCES interfaces(id),
    name         TEXT    NOT NULL,
    return_type  TEXT    NOT NULL DEFAULT ''
);
CREATE TABLE properties (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    interface_id INTEGER NOT NULL REFERENCES interfaces(id),
    name         TEXT    NOT NULL,
    type         TEXT    NOT NULL DEFAULT ''
);
CREATE TABLE parameters (
    id        INTEGER PRIMARY KEY AUTOINCREMENT,
    method_id INTEGER NOT NULL REFERENCES methods(id),
    name      TEXT    NOT NULL,
    type      TEXT    NOT NULL DEFAULT '',
    position  INTEGER NOT NULL DEFAULT 0
);
INSERT INTO interfaces(id, name) VALUES (1, 'PartDocument'), (2, 'Part');
INSERT INTO properties(interface_id, name, type) VALUES (1, 'Part', 'Part'), (1, 'Name', 'String');
INSERT INTO methods(interface_id, name, return_type) VALUES (1, 'Save', '');
`)
	if err != nil {
		t.Fatal(err)
	}
	return &TLBDatabase{db: db}
}

// u16 はテスト期待値用の UTF-16 長（実装の utf16Len とは独立に計算する）。
func u16(s string) int { return len(utf16.Encode([]rune(s))) }

func toCRLF(s string) string { return strings.ReplaceAll(s, "\n", "\r\n") }

// --- 補完 ---

func TestCompletionAfterJapaneseOnSameLine(t *testing.T) {
	db := newTestTLBDatabase(t)
	const target = `    Debug.Print "あ": doc.`
	base := "Sub Main()\n    Dim doc As PartDocument\n" + target + "\nEnd Sub\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			table := ParseSymbols(tc.src)
			prefix := lineAt(tc.src, 2, u16(target))
			if prefix != target {
				t.Fatalf("lineAt = %q, want %q", prefix, target)
			}
			items := BuildCompletions(prefix, table, db, scopeAtLine(tc.src, 2))
			if len(items) == 0 {
				t.Fatalf("補完候補が返らない (prefix=%q)", prefix)
			}
		})
	}
}

func TestCompletionAfterSurrogatePair(t *testing.T) {
	db := newTestTLBDatabase(t)
	const target = `    Debug.Print "😀": doc.`
	src := toCRLF("Sub Main()\n    Dim doc As PartDocument\n" + target + "\nEnd Sub\n")
	table := ParseSymbols(src)
	items := BuildCompletions(lineAt(src, 2, u16(target)), table, db, scopeAtLine(src, 2))
	if len(items) == 0 {
		t.Fatal("補完候補が返らない")
	}
}

// --- ホバー ---

func TestHoverAfterJapaneseOnSameLine(t *testing.T) {
	db := newTestTLBDatabase(t)
	base := "Dim doc As PartDocument\nx = \"ああ\": doc.Name\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			table := ParseSymbols(tc.src)
			// カーソルは "Name" の "Na|me"
			word, prev := wordAndPrevWord(tc.src, 1, u16(`x = "ああ": doc.Na`))
			if word != "Name" || prev != "doc" {
				t.Fatalf("wordAndPrevWord = (%q, %q), want (\"Name\", \"doc\")", word, prev)
			}
			if md := ResolveHoverMarkdown(word, prev, table, db, "module"); md == "" {
				t.Fatal("ホバーの Markdown が空")
			}
		})
	}
}

func TestHoverOnVariableAfterSurrogatePair(t *testing.T) {
	src := "Dim doc As PartDocument\nx = \"😀\": doc.Name\n"
	// カーソルは "doc" の "d|oc"
	word, prev := wordAndPrevWord(src, 1, u16(`x = "😀": d`))
	if word != "doc" || prev != "" {
		t.Fatalf("wordAndPrevWord = (%q, %q), want (\"doc\", \"\")", word, prev)
	}
}

// --- 診断 ---

func findDiag(diags []Diagnostic, line uint32) *Diagnostic {
	for i := range diags {
		if diags[i].Range.Start.Line == line {
			return &diags[i]
		}
	}
	return nil
}

func TestDiagnosticsRangeIsUTF16(t *testing.T) {
	db := newTestTLBDatabase(t)
	src := "Sub M()\n" +
		"    Dim doc As PartDocument\n" +
		"x = \"ああ\": doc.Bogus\n" + // line 2
		"x = \"😀\": doc.Zzz\n" + // line 3
		"　　doc.Bad ' あ\n" + // line 4（全角空白でインデント）
		"End Sub\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", src},
		{"CRLF", toCRLF(src)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			diags := BuildDiagnostics(tc.src, ParseSymbols(tc.src), db)
			if len(diags) != 3 {
				t.Fatalf("診断件数 = %d, want 3: %+v", len(diags), diags)
			}
			for _, want := range []struct {
				line       uint32
				start, end uint32
			}{
				{2, 14, 19},
				{3, uint32(u16(`x = "😀": doc.`)), uint32(u16(`x = "😀": doc.Zzz`))},
				{4, 6, 9},
			} {
				d := findDiag(diags, want.line)
				if d == nil {
					t.Fatalf("行 %d の診断がない: %+v", want.line, diags)
				}
				if d.Range.Start.Character != want.start || d.Range.End.Character != want.end || d.Range.End.Line != want.line {
					t.Errorf("行 %d の範囲 = %+v, want (%d,%d)", want.line, d.Range, want.start, want.end)
				}
			}
			if d := findDiag(diags, 3); d != nil && d.Range.Start.Character != 14 {
				t.Errorf("サロゲートペア行の開始 = %d, want 14", d.Range.Start.Character)
			}
		})
	}
}

// --- セマンティックトークン ---

type absToken struct {
	line, col, length int
	tokenType, mods   uint32
}

// decodeSemanticTokens は LSP の delta 形式を絶対位置に戻す。
func decodeSemanticTokens(t *testing.T, data []uint32) []absToken {
	t.Helper()
	if len(data)%5 != 0 {
		t.Fatalf("データ長が 5 の倍数でない: %d", len(data))
	}
	var out []absToken
	line, col := 0, 0
	for i := 0; i < len(data); i += 5 {
		dl, dc := int(data[i]), int(data[i+1])
		if dl == 0 {
			col += dc
		} else {
			line += dl
			col = dc
		}
		out = append(out, absToken{line, col, int(data[i+2]), data[i+3], data[i+4]})
	}
	return out
}

func tokensOnLine(toks []absToken, line int) []absToken {
	var out []absToken
	for _, tk := range toks {
		if tk.line == line {
			out = append(out, tk)
		}
	}
	return out
}

func TestSemanticTokenReferenceColumnIsUTF16(t *testing.T) {
	db := newTestTLBDatabase(t)
	src := "Dim doc As PartDocument\nx = \"ああ\": doc.Save\n"
	toks := tokensOnLine(decodeSemanticTokens(t, BuildSemanticTokens(src, ParseSymbols(src), db)), 1)
	want := []absToken{{1, 10, 3, tokenTypeVariable, 0}}
	if len(toks) != len(want) || toks[0] != want[0] {
		t.Fatalf("行 1 のトークン = %+v, want %+v", toks, want)
	}
}

func TestSemanticTokenNoDuplicateDeclAfterJapanese(t *testing.T) {
	db := newTestTLBDatabase(t)
	const line1 = `    x = "あ": Dim doc As PartDocument`
	src := "Sub M()\n" + line1 + "\nEnd Sub\n"
	toks := tokensOnLine(decodeSemanticTokens(t, BuildSemanticTokens(src, ParseSymbols(src), db)), 1)
	docCol := u16(`    x = "あ": Dim `)
	want := []absToken{
		{1, docCol, 3, tokenTypeVariable, tokenModDeclaration},
		{1, docCol + u16("doc As "), u16("PartDocument"), tokenTypeType, 0},
	}
	if len(toks) != len(want) {
		t.Fatalf("行 1 のトークン = %+v, want %+v（重複トークンが出ている可能性）", toks, want)
	}
	for i := range want {
		if toks[i] != want[i] {
			t.Errorf("トークン[%d] = %+v, want %+v", i, toks[i], want[i])
		}
	}
}

func TestSemanticTokensWithSurrogatePair(t *testing.T) {
	db := newTestTLBDatabase(t)
	const declLine = `    x = "😀": Dim doc As PartDocument`
	const refLine = `    y = "😀": doc.Save`
	base := "Sub M()\n" + declLine + "\n" + refLine + "\nEnd Sub\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			toks := decodeSemanticTokens(t, BuildSemanticTokens(tc.src, ParseSymbols(tc.src), db))
			docCol := u16(`    x = "😀": Dim `)
			want := []absToken{
				{1, docCol, 3, tokenTypeVariable, tokenModDeclaration},
				{1, docCol + u16("doc As "), u16("PartDocument"), tokenTypeType, 0},
				{2, u16(`    y = "😀": `), 3, tokenTypeVariable, 0},
			}
			if len(toks) != len(want) {
				t.Fatalf("トークン = %+v, want %+v", toks, want)
			}
			for i := range want {
				if toks[i] != want[i] {
					t.Errorf("トークン[%d] = %+v, want %+v", i, toks[i], want[i])
				}
			}
		})
	}
}
