package main

import (
	"sort"
	"strings"
	"time"
)

// SymbolKind はシンボルの種類。
type SymbolKind int

const (
	SymVariable SymbolKind = iota
	SymParam
	SymConst
	SymProcedure // Sub / Function / Declare
	SymProperty  // Property Get / Let / Set
	SymUDT
	SymUDTMember
	SymEnum
	SymEnumMember
	SymEvent
)

// Visibility は宣言の公開範囲。
// 既定値: モジュールレベルの Dim / Const → Private。修飾子のない Sub / Function / Property / Type / Enum / Event / Declare → Public。
// プロシージャ内 → Local。
type Visibility int

const (
	VisPrivate Visibility = iota
	VisPublic
	VisFriend
	VisGlobal
	VisLocal
)

// ProcKind はプロシージャの種類。
type ProcKind int

const (
	ProcSub ProcKind = iota
	ProcFunction
	ProcPropertyGet
	ProcPropertyLet
	ProcPropertySet
	ProcDeclareSub
	ProcDeclareFunction
	ProcEvent
)

// ParamInfo はプロシージャの引数 1 つ。
type ParamInfo struct {
	Name                        string
	Type                        TypeRef // 型なしならゼロ値（Variant）
	ByVal, Optional, ParamArray bool
	Default                     string
	Line, Col                   int // 名前の位置（UTF-16）
}

// ProcInfo はプロシージャ（Sub / Function / Property / Declare / Event）の宣言。
type ProcInfo struct {
	Name       string // 型接尾辞なし
	Kind       ProcKind
	Visibility Visibility
	Static     bool
	Params     []ParamInfo
	ReturnType TypeRef // Function / Property Get / Declare Function。As 句 > 接尾辞 > Variant
	Doc        string
	HasBody    bool // Declare / Event は false（スコープ範囲として扱わない）
	StartLine  int  // 宣言の先頭行（修飾子を含む、0-based）
	EndLine    int  // End 行（0-based、両端を含む）
	NameLine   int
	NameCol    int
}

// ScopeKey はスコープ名（"Sub:Foo" / "Function:Foo" / "PropertyGet:Foo" / "PropertyLet:Foo" / "PropertySet:Foo"）を返す。
func (p *ProcInfo) ScopeKey() string {
	if p == nil {
		return "module"
	}
	switch p.Kind {
	case ProcSub:
		return "Sub:" + p.Name
	case ProcFunction:
		return "Function:" + p.Name
	case ProcPropertyGet:
		return "PropertyGet:" + p.Name
	case ProcPropertyLet:
		return "PropertyLet:" + p.Name
	case ProcPropertySet:
		return "PropertySet:" + p.Name
	case ProcEvent:
		return "Event:" + p.Name
	default:
		return "Declare:" + p.Name
	}
}

// HasReturn は戻り値を持つ（Function / Property Get / Declare Function）かを返す。
func (p *ProcInfo) HasReturn() bool {
	return p.Kind == ProcFunction || p.Kind == ProcPropertyGet || p.Kind == ProcDeclareFunction
}

// Symbol は VBA コード内の宣言 1 つ。
type Symbol struct {
	Name       string
	Kind       SymbolKind
	TypeName   string // 修飾なしの型名（"Dictionary"）。不明なら ""
	TypeLib    string // 修飾子（"Scripting"）。なければ ""
	IsArray    bool
	AsNew      bool
	WithEvents bool
	Static     bool
	Visibility Visibility
	Scope      string    // "module" または ProcInfo.ScopeKey()（UDT のメンバーは "UDT:<型名>"）
	Owner      *ProcInfo // 所属プロシージャ（nil = モジュール）
	Decl       *ProcInfo // Kind が Procedure / Property / Event のときの宣言
	Parent     string    // UDTMember / EnumMember の親の名前
	Value      string    // Const / EnumMember の値の式
	Doc        string
	Line       int // 0-based 行番号（名前の位置）
	Col        int // 0-based 列番号（名前の位置、LSP と同じ UTF-16 単位）
	// TypeNameLine / TypeNameCol は型名（修飾付きなら最後の名前）の位置。HasTypeNamePos が true のときのみ有効。
	// CreateObject("ProgID") 由来や型接尾辞の型など、型名がソースに現れない場合は HasTypeNamePos が false。
	TypeNameLine   int
	TypeNameCol    int
	HasTypeNamePos bool
}

// TypeRef はシンボルの型参照を返す。
func (s *Symbol) TypeRef() TypeRef {
	return TypeRef{Lib: s.TypeLib, Name: s.TypeName, Array: s.IsArray}
}

// QualifiedTypeName は表示用の型名（書かれていた修飾子付き、配列なら "()" 付き）を返す。
func (s *Symbol) QualifiedTypeName() string { return s.TypeRef().String() }

// UDTInfo はユーザー定義型（Type ... End Type）。
type UDTInfo struct {
	Name       string
	Visibility Visibility
	Members    []*Symbol
	StartLine  int
	EndLine    int
	Doc        string
}

// EnumInfo は列挙型（Enum ... End Enum）。
type EnumInfo struct {
	Name       string
	Visibility Visibility
	Members    []*Symbol
	StartLine  int
	EndLine    int
	Doc        string
}

// WithBlock は With ... End With の範囲と対象式。
type WithBlock struct {
	StartLine int
	EndLine   int
	Expr      string // 行継続を空白に正規化済み
	ExprLine  int
	ExprCol   int
}

// ParseStats はパースの統計（性能計測・診断用）。
type ParseStats struct {
	UsedLL       bool          // SLL で構文エラーが出て LL でやり直した
	SyntaxErrors int           // LL 段で検出した構文エラーの数（SLL で成功したときは 0）
	SLLErrorLine int           // SLL 段が失敗したトークンの行（1-based、計測ログ用）
	Duration     time.Duration // 字句解析・構文解析・シンボル収集の合計時間
	Repaired     int           // 解析の前に末尾を補修した論理行の数（入力途中の行。正しいコードでは 0）
}

// SymbolTable はファイル単位のシンボルテーブル。
type SymbolTable struct {
	Symbols    []Symbol
	Procs      []*ProcInfo
	UDTs       []*UDTInfo
	Enums      []*EnumInfo
	Withs      []WithBlock
	Implements []string
	Stats      ParseStats

	scopes []*ProcInfo                 // HasBody のものを StartLine 順に並べたもの
	index  map[string]map[string][]int // scopeKey → 小文字名 → Symbols の添字
}

// finalize はスコープの並べ替えと索引の構築を行う（Symbols を変更したら呼び直すこと）。
func (st *SymbolTable) finalize() {
	if st.scopes == nil {
		for _, p := range st.Procs {
			if p.HasBody {
				st.scopes = append(st.scopes, p)
			}
		}
	}
	sort.SliceStable(st.scopes, func(i, j int) bool { return st.scopes[i].StartLine < st.scopes[j].StartLine })
	st.index = map[string]map[string][]int{}
	for i := range st.Symbols {
		s := &st.Symbols[i]
		m := st.index[s.Scope]
		if m == nil {
			m = map[string][]int{}
			st.index[s.Scope] = m
		}
		key := strings.ToLower(s.Name)
		m[key] = append(m[key], i)
	}
}

func (st *SymbolTable) ensureIndex() {
	if st.index == nil {
		st.finalize()
	}
}

// ProcAt は line（0-based）を含む本体付きプロシージャを返す。なければ nil。
func (st *SymbolTable) ProcAt(line int) *ProcInfo {
	if st == nil {
		return nil
	}
	st.ensureIndex()
	i := sort.Search(len(st.scopes), func(i int) bool { return st.scopes[i].StartLine > line }) - 1
	if i >= 0 && st.scopes[i].EndLine >= line {
		return st.scopes[i]
	}
	return nil
}

// ScopeAt は line を含むスコープ名（プロシージャ外なら "module"）を返す。
func (st *SymbolTable) ScopeAt(line int) string {
	return st.ProcAt(line).ScopeKey()
}

// lookupIn は scope 内で name のシンボルを返す。"module" では Variable / Const > 戻り値のあるプロシージャ > その他 の順に優先する。
func (st *SymbolTable) lookupIn(name, scope string) *Symbol {
	st.ensureIndex()
	idx := st.index[scope][strings.ToLower(name)]
	if len(idx) == 0 {
		return nil
	}
	best, bestRank := -1, 99
	for _, i := range idx {
		s := &st.Symbols[i]
		rank := 2
		switch {
		case s.Kind == SymVariable || s.Kind == SymParam || s.Kind == SymConst:
			rank = 0
		case s.Decl != nil && s.Decl.HasReturn():
			rank = 1
		}
		if rank < bestRank {
			best, bestRank = i, rank
		}
	}
	return &st.Symbols[best]
}

// Lookup は line の位置から見える name のシンボルを返す（プロシージャのスコープ → モジュール）。
func (st *SymbolTable) Lookup(name string, line int) *Symbol {
	if st == nil {
		return nil
	}
	return st.FindByName(name, st.ScopeAt(line))
}

// FindByName はスコープ優先でシンボルを検索する。curScope のシンボルを優先し、次にモジュールスコープを返す。
func (st *SymbolTable) FindByName(name, curScope string) *Symbol {
	if st == nil {
		return nil
	}
	if curScope != "module" {
		if s := st.lookupIn(name, curScope); s != nil {
			return s
		}
	}
	return st.lookupIn(name, "module")
}

// WithAt は line を含む最も内側の With ブロックを返す。With 行自身は含めない
// （With 行の先頭ドットは外側の With に解決させるため）。
func (st *SymbolTable) WithAt(line int) *WithBlock {
	if st == nil {
		return nil
	}
	var best *WithBlock
	for i := range st.Withs {
		w := &st.Withs[i]
		if w.StartLine < line && line <= w.EndLine && (best == nil || w.StartLine >= best.StartLine) {
			best = w
		}
	}
	return best
}

// UDT は名前（大文字小文字無視）でユーザー定義型を返す。
func (st *SymbolTable) UDT(name string) *UDTInfo {
	for _, u := range st.UDTs {
		if strings.EqualFold(u.Name, name) {
			return u
		}
	}
	return nil
}

// Enum は名前（大文字小文字無視）で列挙型を返す。
func (st *SymbolTable) Enum(name string) *EnumInfo {
	for _, e := range st.Enums {
		if strings.EqualFold(e.Name, name) {
			return e
		}
	}
	return nil
}
