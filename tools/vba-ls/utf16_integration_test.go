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
INSERT INTO interfaces(id, name) VALUES (1, 'PartDocument'), (2, 'Part'), (3, 'Bodies'), (4, 'Body'), (5, 'Collection'), (6, 'AnyObject');
INSERT INTO properties(interface_id, name, type) VALUES (1, 'Part', 'Part'), (1, 'Name', 'String'),
    (2, 'Bodies', 'Bodies'), (2, 'Name', 'String'), (3, 'Count', 'Long'), (4, 'Name', 'String'), (6, 'Name', 'String');
INSERT INTO methods(id, interface_id, name, return_type) VALUES (1, 1, 'Save', ''), (2, 3, 'Item', 'Body'), (3, 5, 'GetItem', 'CATBaseDispatch'),
    (4, 3, 'QueryInterface', 'void');
INSERT INTO parameters(method_id, name, type, position) VALUES (2, 'iIndex', 'Variant', 0), (3, 'IDName', 'String', 0);
`)
	if err != nil {
		t.Fatal(err)
	}
	return &TLBDatabase{db: db}
}

// newTestResolver はテスト用 TLB を含む既定の resolver を作る。
func newTestResolver(t *testing.T) *CompositeResolver {
	t.Helper()
	return NewDefaultResolver(newTestTLBDatabase(t), nil)
}

// newTestDoc はテスト用 TLB を含む resolver で src を解析した Document を作る。
func newTestDoc(t *testing.T, src string) *Document {
	t.Helper()
	return NewDocument("file:///test.bas", src, ParseSymbols(src), newTestResolver(t))
}

// caretPos は src 中の "|" を取り除き、その位置（行と UTF-16 位置）を返す。
func caretPos(t *testing.T, src string) (clean string, line, char int) {
	t.Helper()
	i := strings.Index(src, "|")
	if i < 0 {
		t.Fatalf("カーソル位置 | がない: %q", src)
	}
	clean = src[:i] + src[i+1:]
	before := src[:i]
	line = strings.Count(before, "\n")
	lineStart := strings.LastIndex(before, "\n") + 1
	return clean, line, u16(before[lineStart:])
}

// completeAt は "|" の位置で補完を実行する。types が nil ならテスト用 TLB を含む resolver を使う。
func completeAt(t *testing.T, types *CompositeResolver, srcWithCaret string) []CompletionItem {
	t.Helper()
	if types == nil {
		types = newTestResolver(t)
	}
	src, line, char := caretPos(t, srcWithCaret)
	return BuildCompletions(NewDocument("file:///test.bas", src, ParseSymbols(src), types), line, char)
}

// labelsOf は補完候補のラベル集合を返す。
func labelsOf(items []CompletionItem) map[string]bool {
	m := map[string]bool{}
	for _, it := range items {
		m[it.Label] = true
	}
	return m
}

// u16 はテスト期待値用の UTF-16 長（実装の utf16Len とは独立に計算する）。
func u16(s string) int { return len(utf16.Encode([]rune(s))) }

func toCRLF(s string) string { return strings.ReplaceAll(s, "\n", "\r\n") }

// --- 補完 ---

func TestCompletionAfterJapaneseOnSameLine(t *testing.T) {
	db := newTestResolver(t)
	const target = `    Debug.Print "あ": doc.`
	base := "Sub Main()\n    Dim doc As PartDocument\n" + target + "\nEnd Sub\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc := NewDocument("", tc.src, ParseSymbols(tc.src), db)
			if prefix := doc.lineAt(2, u16(target)); prefix != target {
				t.Fatalf("lineAt = %q, want %q", prefix, target)
			}
			if items := BuildCompletions(doc, 2, u16(target)); len(items) == 0 {
				t.Fatal("補完候補が返らない")
			}
		})
	}
}

func TestCompletionAfterSurrogatePair(t *testing.T) {
	db := newTestResolver(t)
	const target = `    Debug.Print "😀": doc.`
	src := toCRLF("Sub Main()\n    Dim doc As PartDocument\n" + target + "\nEnd Sub\n")
	items := BuildCompletions(NewDocument("", src, ParseSymbols(src), db), 2, u16(target))
	if len(items) == 0 {
		t.Fatal("補完候補が返らない")
	}
}

// --- ホバー ---

func TestHoverAfterJapaneseOnSameLine(t *testing.T) {
	db := newTestResolver(t)
	base := "Dim doc As PartDocument\nx = \"ああ\": doc.Name\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			// カーソルは "Name" の "Na|me"
			word, prev := wordAndPrevWord(tc.src, 1, u16(`x = "ああ": doc.Na`))
			if word != "Name" || prev != "doc" {
				t.Fatalf("wordAndPrevWord = (%q, %q), want (\"Name\", \"doc\")", word, prev)
			}
			doc := NewDocument("", tc.src, ParseSymbols(tc.src), db)
			if md := ResolveHover(doc, 1, u16(`x = "ああ": doc.Na`)); md == "" {
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
	db := newTestResolver(t)
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
			diags := BuildDiagnostics(NewDocument("", tc.src, ParseSymbols(tc.src), db))
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
	db := newTestResolver(t)
	src := "Dim doc As PartDocument\nx = \"ああ\": doc.Save\n"
	toks := tokensOnLine(decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, ParseSymbols(src), db))), 1)
	want := []absToken{{1, 10, 3, tokenTypeVariable, 0}}
	if len(toks) != len(want) || toks[0] != want[0] {
		t.Fatalf("行 1 のトークン = %+v, want %+v", toks, want)
	}
}

func TestSemanticTokenNoDuplicateDeclAfterJapanese(t *testing.T) {
	db := newTestResolver(t)
	const line1 = `    x = "あ": Dim doc As PartDocument`
	src := "Sub M()\n" + line1 + "\nEnd Sub\n"
	toks := tokensOnLine(decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, ParseSymbols(src), db))), 1)
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
	db := newTestResolver(t)
	const declLine = `    x = "😀": Dim doc As PartDocument`
	const refLine = `    y = "😀": doc.Save`
	base := "Sub M()\n" + declLine + "\n" + refLine + "\nEnd Sub\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			toks := decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", tc.src, ParseSymbols(tc.src), db)))
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

func TestSemanticTokensIgnoreStringLiterals(t *testing.T) {
	db := newTestResolver(t)
	const line2 = `    MsgBox "doc あdoc" & doc.Name ' doc`
	base := "Sub M()\n    Dim doc As PartDocument\n" + line2 + "\nEnd Sub\n"
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) {
			toks := tokensOnLine(decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", tc.src, ParseSymbols(tc.src), db))), 2)
			want := []absToken{{2, u16(`    MsgBox "doc あdoc" & `), 3, tokenTypeVariable, 0}}
			if len(toks) != len(want) || toks[0] != want[0] {
				t.Fatalf("行 2 のトークン = %+v, want %+v（文字列・コメント内の語が色付けされている可能性）", toks, want)
			}
		})
	}
}
