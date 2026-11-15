package main

import (
	"testing"
)

// --- 項目15: 日本語識別子 ---

func TestJapaneseIdentifiers(t *testing.T) {
	db := newTestTLBDatabase(t)
	base := "Sub 処理()\n    Dim 部品 As PartDocument\n    部品.Bogus\n    x = 部品.Name\nEnd Sub\n"
	lfAndCRLF(t, base, func(t *testing.T, src string) {
		table := ParseSymbols(src)

		s := table.FindByName("部品", "Sub:処理")
		if s == nil {
			t.Fatalf("部品 が Sub:処理 スコープで見つからない: %+v", table.Symbols)
		}
		typeCol := u16("    Dim 部品 As ")
		if s.Line != 1 || s.Col != 8 || s.TypeName != "PartDocument" || s.TypeNameLine != 1 || s.TypeNameCol != typeCol {
			t.Fatalf("部品 = %+v, want (1,8) 型 PartDocument (1,%d)", *s, typeCol)
		}

		if got := scopeAtLine(src, 2); got != "Sub:処理" {
			t.Errorf("scopeAtLine(2) = %q, want %q", got, "Sub:処理")
		}

		diags := BuildDiagnostics(src, table, db)
		if len(diags) != 1 {
			t.Fatalf("診断件数 = %d, want 1: %+v", len(diags), diags)
		}
		if r := diags[0].Range; r.Start.Line != 2 || r.Start.Character != 7 || r.End.Character != 12 {
			t.Errorf("診断範囲 = %+v, want 行2 (7,12)", r)
		}

		if items := BuildCompletions(lineAt(src, 3, u16("    x = 部品.")), table, db, scopeAtLine(src, 3)); len(items) == 0 {
			t.Error("部品. の補完候補が返らない")
		}

		if word, prev := wordAndPrevWord(src, 3, u16("    x = 部品.Na")); word != "Name" || prev != "部品" {
			t.Errorf("wordAndPrevWord = (%q, %q), want (\"Name\", \"部品\")", word, prev)
		}
		if word, prev := wordAndPrevWord(src, 3, u16("    x = 部")); word != "部品" || prev != "" {
			t.Errorf("wordAndPrevWord = (%q, %q), want (\"部品\", \"\")", word, prev)
		}

		toks := decodeSemanticTokens(t, BuildSemanticTokens(src, table, db))
		assertTokens(t, toks, []absToken{
			{1, 8, 2, tokenTypeVariable, tokenModDeclaration},
			{1, typeCol, 12, tokenTypeType, 0},
			{2, 4, 2, tokenTypeVariable, 0},
			{3, 8, 2, tokenTypeVariable, 0},
		})
	})
}

func TestSurrogatePairIdentifier(t *testing.T) {
	db := newTestTLBDatabase(t)
	lfAndCRLF(t, "Dim 𠮷野家 As PartDocument\n𠮷野家.Save\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		s := findSym(t, table, "𠮷野家", "module")
		typeCol := u16("Dim 𠮷野家 As ")
		if s.Col != 4 || s.TypeNameCol != typeCol {
			t.Fatalf("𠮷野家 = %+v, want Col 4, TypeNameCol %d", s, typeCol)
		}
		toks := decodeSemanticTokens(t, BuildSemanticTokens(src, table, db))
		assertTokens(t, toks, []absToken{
			{0, 4, 4, tokenTypeVariable, tokenModDeclaration},
			{0, typeCol, 12, tokenTypeType, 0},
			{1, 0, 4, tokenTypeVariable, 0},
		})
	})
}

func TestScopeAtLineUnicode(t *testing.T) {
	db := newTestTLBDatabase(t)
	base := "Private Sub 処理()\n" + // 0
		"    Dim 部品 As PartDocument\n" + // 1
		"    部品.Bogus\n" + // 2
		"End Sub\n" + // 3
		"　Function 計算１()\n" + // 4（全角空白インデント・全角数字）
		"    y = 2\n" + // 5
		"　End Function\n" + // 6
		"z = 3\n" // 7
	lfAndCRLF(t, base, func(t *testing.T, src string) {
		for _, want := range []struct {
			line  int
			scope string
		}{
			{2, "Sub:処理"},
			{4, "module"},
			{5, "Function:計算１"},
			{7, "module"},
		} {
			if got := scopeAtLine(src, want.line); got != want.scope {
				t.Errorf("scopeAtLine(%d) = %q, want %q", want.line, got, want.scope)
			}
		}
		// Private Sub 内の宣言もパーサと同じスコープ名で解決され、診断が出ること
		table := ParseSymbols(src)
		findSym(t, table, "部品", "Sub:処理")
		diags := BuildDiagnostics(src, table, db)
		if len(diags) != 1 || diags[0].Range.Start.Line != 2 {
			t.Fatalf("診断 = %+v, want 行2 に 1 件", diags)
		}
	})
}
