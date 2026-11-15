package main

import "strings"

// Document は 1 つの版のテキストと、その解析結果・型解決をまとめたもの。
// 補完・ホバー・診断・セマンティックトークン・定義ジャンプはすべてこれを受け取る。
type Document struct {
	URI    string
	Text   string
	Lines  []string // "\n" で分割し、末尾の "\r" を除いた行
	Table  *SymbolTable
	Types  *CompositeResolver
	Self   TypeRef      // Me の型（クラス・フォームはモジュール名。標準モジュール・不明な拡張子はゼロ値）
	Module *ModuleEntry // このファイルのモジュール（*_utf 以外なら nil）
}

// NewDocument は Document を作る。table が nil なら空の表、global が nil なら何も解決しない resolver を使う。
// URI の拡張子が *_utf ならモジュールとして扱い、Me. と自モジュールの Private メンバーを解決できるようにする。
func NewDocument(uri, text string, table *SymbolTable, global *CompositeResolver) *Document {
	return newDocumentWithLines(uri, text, splitLines(text), table, global)
}

func newDocumentWithLines(uri, text string, lines []string, table *SymbolTable, global *CompositeResolver) *Document {
	if table == nil {
		table = &SymbolTable{}
	}
	return newDocumentFull(uri, text, lines, table, newModuleEntryForURI(uri, text, lines, table, 0), global)
}

// newDocumentFull はサーバー用。entry は再解析のときに作ってキャッシュしたもの（nil 可）。
func newDocumentFull(uri, text string, lines []string, table *SymbolTable, entry *ModuleEntry, global *CompositeResolver) *Document {
	if table == nil {
		table = &SymbolTable{}
	}
	if global == nil {
		global = NewCompositeResolver()
	}
	d := &Document{URI: uri, Text: text, Lines: lines, Table: table, Module: entry}
	if entry != nil && entry.Kind != ModuleStd {
		d.Self = entry.selfRef()
	}
	// 解決順の先頭に、現在のモジュールの UDT / Enum と、自モジュール（Private を含む）を足す
	d.Types = global.WithLocal(newModuleTypeProvider(table), newSelfProvider(entry))
	return d
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
