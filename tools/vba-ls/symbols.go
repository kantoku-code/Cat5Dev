package main

import "strings"

// Symbol は VBA コード内の変数宣言情報を表す。
type Symbol struct {
	Name         string // 変数名
	TypeName     string // 型名（"PartDocument" 等）
	Scope        string // "module" または "Sub:FooSub" / "Function:Bar" 等
	Line         int    // 0-based 行番号（変数名の位置）
	Col          int    // 0-based 列番号（変数名の位置、LSP と同じ UTF-16 単位）
	TypeNameLine int    // 0-based 行番号（型名の位置。HasTypeNamePos が true のときのみ有効）
	TypeNameCol  int    // 0-based 列番号（型名の位置、LSP と同じ UTF-16 単位。HasTypeNamePos が true のときのみ有効）
	// HasTypeNamePos はソース上に型名トークンがある（Dim ... As / New）ときのみ true。
	// CreateObject("ProgID") 由来の型など、型名がソースに現れない場合は false。
	HasTypeNamePos bool
}

// SymbolTable はファイル単位のシンボルテーブル。
type SymbolTable struct {
	Symbols []Symbol
}

// FindByName はスコープ優先でシンボルを検索する。
// curScope が一致するシンボルを優先し、次にモジュールスコープを返す。
func (st *SymbolTable) FindByName(name, curScope string) *Symbol {
	var moduleLevel *Symbol
	lname := strings.ToLower(name)
	for i := range st.Symbols {
		s := &st.Symbols[i]
		if strings.ToLower(s.Name) == lname {
			if s.Scope == curScope {
				return s
			}
			if s.Scope == "module" {
				moduleLevel = s
			}
		}
	}
	return moduleLevel
}

// TypeRef はシンボルの型参照を返す。
func (s *Symbol) TypeRef() TypeRef { return ParseTypeName(s.TypeName) }

// Visibility は宣言の公開範囲。
type Visibility int

const (
	VisPrivate Visibility = iota
	VisPublic
	VisFriend
	VisGlobal
	VisLocal
)
