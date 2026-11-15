package main

import "strings"

// Document は 1 つの版のテキストと、その解析結果・型解決をまとめたもの。
// 補完・ホバー・診断・セマンティックトークンはすべてこれを受け取る。
type Document struct {
	URI   string
	Text  string
	Lines []string // "\n" で分割し、末尾の "\r" を除いた行
	Table *SymbolTable
	Types *CompositeResolver
	Self  TypeRef // 2-3: クラスモジュール自身の型（Me.）
}

// NewDocument は Document を作る。table が nil なら空の表、global が nil なら何も解決しない resolver を使う。
func NewDocument(uri, text string, table *SymbolTable, global *CompositeResolver) *Document {
	return newDocumentWithLines(uri, text, splitLines(text), table, global)
}

func newDocumentWithLines(uri, text string, lines []string, table *SymbolTable, global *CompositeResolver) *Document {
	if table == nil {
		table = &SymbolTable{}
	}
	if global == nil {
		global = NewCompositeResolver()
	}
	return &Document{URI: uri, Text: text, Lines: lines, Table: table, Types: global}
}

// splitLines は src を "\n" で分割し、各行末の "\r" を除く。
func splitLines(src string) []string {
	lines := strings.Split(src, "\n")
	for i, l := range lines {
		lines[i] = strings.TrimSuffix(l, "\r")
	}
	return lines
}

// lineAt は line 行目の、UTF-16 位置 char より前のテキストを返す。
func (d *Document) lineAt(line, char int) string {
	if line < 0 || line >= len(d.Lines) {
		return ""
	}
	l := d.Lines[line]
	return l[:utf16ToByteOffset(l, char)]
}
