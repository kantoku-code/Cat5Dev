package main

import (
	"testing"
)

// findSym はテーブルから名前とスコープが一致するシンボルを返す（大文字小文字は区別する）。
func findSym(t *testing.T, table *SymbolTable, name, scope string) Symbol {
	t.Helper()
	for _, s := range table.Symbols {
		if s.Name == name && s.Scope == scope {
			return s
		}
	}
	t.Fatalf("シンボル %q (scope=%q) が見つからない: %+v", name, scope, table.Symbols)
	return Symbol{}
}

// varsOf は Kind が変数のシンボルだけを返す（Sub 自体などのシンボルを除く）。
func varsOf(table *SymbolTable) []Symbol {
	var out []Symbol
	for _, s := range table.Symbols {
		if s.Kind == SymVariable {
			out = append(out, s)
		}
	}
	return out
}

// lfAndCRLF は LF 版と CRLF 版のソースでサブテストを実行する。
func lfAndCRLF(t *testing.T, base string, fn func(t *testing.T, src string)) {
	t.Helper()
	for _, tc := range []struct{ name, src string }{
		{"LF", base},
		{"CRLF", toCRLF(base)},
	} {
		t.Run(tc.name, func(t *testing.T) { fn(t, tc.src) })
	}
}

func assertTokens(t *testing.T, got, want []absToken) {
	t.Helper()
	if len(got) != len(want) {
		t.Fatalf("トークン = %+v, want %+v", got, want)
	}
	for i := range want {
		if got[i] != want[i] {
			t.Errorf("トークン[%d] = %+v, want %+v", i, got[i], want[i])
		}
	}
}

// --- 項目13: Set 由来シンボルの位置 ---

func TestSetNewSymbolPositions(t *testing.T) {
	db := newTestResolver(t)
	lfAndCRLF(t, "Sub M()\n    Set doc = New PartDocument\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		if vars := varsOf(table); len(vars) != 1 {
			t.Fatalf("変数の数 = %d, want 1: %+v", len(vars), vars)
		}
		s := findSym(t, table, "doc", "Sub:M")
		typeCol := u16("    Set doc = New ")
		if s.Line != 1 || s.Col != 8 || s.TypeName != "PartDocument" || !s.HasTypeNamePos || s.TypeNameLine != 1 || s.TypeNameCol != typeCol {
			t.Fatalf("doc = %+v, want Line1/Col8, TypeName PartDocument, 型位置 (1,%d)", s, typeCol)
		}
		toks := decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, table, db)))
		assertTokens(t, toks, []absToken{
			{1, 8, 3, tokenTypeVariable, tokenModDeclaration},
			{1, typeCol, 12, tokenTypeType, 0},
		})
	})
}

func TestSetNewAfterJapanese(t *testing.T) {
	db := newTestResolver(t)
	const line1 = `    x = "あ": Set doc = New PartDocument`
	lfAndCRLF(t, "Sub M()\n"+line1+"\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		s := findSym(t, table, "doc", "Sub:M")
		docCol := u16(`    x = "あ": Set `)
		typeCol := docCol + u16("doc = New ")
		if s.Line != 1 || s.Col != docCol || s.TypeNameLine != 1 || s.TypeNameCol != typeCol || !s.HasTypeNamePos {
			t.Fatalf("doc = %+v, want 変数 (1,%d) 型 (1,%d)", s, docCol, typeCol)
		}
		toks := decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, table, db)))
		assertTokens(t, toks, []absToken{
			{1, docCol, 3, tokenTypeVariable, tokenModDeclaration},
			{1, typeCol, 12, tokenTypeType, 0},
		})
	})
}

func TestSetCreateObjectHasNoTypeToken(t *testing.T) {
	db := newTestResolver(t)
	lfAndCRLF(t, "Sub M()\n    Set doc = CreateObject(\"CATIA.PartDocument\")\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		s := findSym(t, table, "doc", "Sub:M")
		if s.TypeName != "PartDocument" || s.HasTypeNamePos || s.Line != 1 || s.Col != 8 {
			t.Fatalf("doc = %+v, want TypeName PartDocument, HasTypeNamePos=false, (1,8)", s)
		}
		toks := decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, table, db)))
		assertTokens(t, toks, []absToken{
			{1, 8, 3, tokenTypeVariable, tokenModDeclaration},
		})
	})
}

func TestSetNewFillsUntypedDim(t *testing.T) {
	db := newTestResolver(t)
	lfAndCRLF(t, "Sub M()\n    Dim doc\n    Set doc = New PartDocument\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		vars := varsOf(table)
		if len(vars) != 1 {
			t.Fatalf("変数の数 = %d, want 1: %+v", len(vars), vars)
		}
		s := vars[0]
		typeCol := u16("    Set doc = New ")
		if s.Name != "doc" || s.Line != 1 || s.Col != 8 || s.TypeName != "PartDocument" || !s.HasTypeNamePos || s.TypeNameLine != 2 || s.TypeNameCol != typeCol {
			t.Fatalf("doc = %+v, want 宣言 (1,8)、型 PartDocument (2,%d)", s, typeCol)
		}
		toks := decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, table, db)))
		assertTokens(t, toks, []absToken{
			{1, 8, 3, tokenTypeVariable, tokenModDeclaration},
			{2, 8, 3, tokenTypeVariable, 0},
			{2, typeCol, 12, tokenTypeType, 0},
		})
	})
}

func TestSetDoesNotOverrideDimType(t *testing.T) {
	lfAndCRLF(t, "Sub M()\n    Dim doc As PartDocument\n    Set doc = New Part\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		vars := varsOf(table)
		if len(vars) != 1 {
			t.Fatalf("変数の数 = %d, want 1: %+v", len(vars), vars)
		}
		s := vars[0]
		if s.TypeName != "PartDocument" || s.TypeNameLine != 1 || s.TypeNameCol != u16("    Dim doc As ") || !s.HasTypeNamePos {
			t.Fatalf("doc = %+v, want Dim 時の型と位置のまま", s)
		}
	})
}

func TestSetMemberLHSIgnored(t *testing.T) {
	lfAndCRLF(t, "Sub M()\n    Set obj.Prop = New PartDocument\n    Set arr(1) = New PartDocument\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		if vars := varsOf(table); len(vars) != 0 {
			t.Fatalf("変数 = %+v, want なし", vars)
		}
	})
}

func TestSetNewQualifiedType(t *testing.T) {
	lfAndCRLF(t, "Sub M()\n    Set d = New Scripting.Dictionary\nEnd Sub\n", func(t *testing.T, src string) {
		s := findSym(t, ParseSymbols(src), "d", "Sub:M")
		if s.TypeName != "Dictionary" || s.TypeLib != "Scripting" || !s.HasTypeNamePos || s.TypeNameLine != 1 || s.TypeNameCol != u16("    Set d = New Scripting.") {
			t.Fatalf("d = %+v, want TypeName Dictionary, 型位置 (1,%d)", s, u16("    Set d = New Scripting."))
		}
	})
}

func TestDimWithoutAsHasNoTypePos(t *testing.T) {
	lfAndCRLF(t, "Sub M()\n    Dim s\nEnd Sub\n", func(t *testing.T, src string) {
		s := findSym(t, ParseSymbols(src), "s", "Sub:M")
		if s.TypeName != "" || s.HasTypeNamePos {
			t.Fatalf("s = %+v, want 型なし・HasTypeNamePos=false", s)
		}
	})
}

func TestDimAsNew(t *testing.T) {
	db := newTestResolver(t)
	lfAndCRLF(t, "Sub M()\n    Dim o As New PartDocument\nEnd Sub\n", func(t *testing.T, src string) {
		table := ParseSymbols(src)
		s := findSym(t, table, "o", "Sub:M")
		typeCol := u16("    Dim o As New ")
		if s.TypeName != "PartDocument" || !s.HasTypeNamePos || s.Line != 1 || s.Col != 8 || s.TypeNameLine != 1 || s.TypeNameCol != typeCol {
			t.Fatalf("o = %+v, want TypeName PartDocument, 型位置 (1,%d)", s, typeCol)
		}
		toks := decodeSemanticTokens(t, BuildSemanticTokens(NewDocument("", src, table, db)))
		assertTokens(t, toks, []absToken{
			{1, 8, 1, tokenTypeVariable, tokenModDeclaration},
			{1, typeCol, 12, tokenTypeType, 0},
		})
	})
}

// --- 項目16: モジュールレベルの ":" 区切り ---

func TestModuleLevelColonSeparatedDeclarations(t *testing.T) {
	const (
		line1 = "Dim s As String: Dim doc As PartDocument"
		line2 = "Private p As Part : Public d2 As PartDocument"
		line3 = "Const A = 1: Dim e As Part"
		line5 = "    Dim a As Part: Dim b As PartDocument"
	)
	base := "Option Explicit\n" + line1 + "\n" + line2 + "\n" + line3 + "\nSub M()\n" + line5 + "\nEnd Sub\n"
	lfAndCRLF(t, base, func(t *testing.T, src string) {
		table := ParseSymbols(src)
		for _, want := range []struct {
			name, scope, typeName string
			line, col             int
		}{
			{"s", "module", "String", 1, 4},
			{"doc", "module", "PartDocument", 1, u16("Dim s As String: Dim ")},
			{"p", "module", "Part", 2, 8},
			{"d2", "module", "PartDocument", 2, u16("Private p As Part : Public ")},
			{"e", "module", "Part", 3, u16("Const A = 1: Dim ")},
			{"a", "Sub:M", "Part", 5, 8},
			{"b", "Sub:M", "PartDocument", 5, u16("    Dim a As Part: Dim ")},
		} {
			s := findSym(t, table, want.name, want.scope)
			if s.TypeName != want.typeName || s.Line != want.line || s.Col != want.col {
				t.Errorf("%s = %+v, want type %s (%d,%d)", want.name, s, want.typeName, want.line, want.col)
			}
		}
	})
}
