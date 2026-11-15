package main

import (
	"reflect"
	"regexp"
	"sort"
	"strings"
)

// TypeRef は型参照。Lib は修飾子（"Scripting" / "VBA" / "CATIA" / "INFITF" など）で、無修飾なら ""。
type TypeRef struct {
	Lib   string
	Name  string // 修飾なし・型接尾辞なし・角括弧なし
	Array bool
}

// IsZero は型が不明（名前なし）かを返す。
func (t TypeRef) IsZero() bool { return t.Name == "" }

// Elem は配列の要素型を返す（配列でなければそのまま）。
func (t TypeRef) Elem() TypeRef {
	t.Array = false
	return t
}

// moduleLib は標準モジュールを型として引くときの修飾子。ソースには現れない（表示でも省く）。
const moduleLib = "#module"

// formBaseType はフォームの基本メンバーの型（vba_builtins.json の MSForms.UserForm）。
var formBaseType = TypeRef{Lib: "MSForms", Name: "UserForm"}

// String は表示用の文字列（"Scripting.Dictionary" / "String()"）を返す。
func (t TypeRef) String() string {
	s := t.Name
	if t.Lib != "" && t.Lib != moduleLib && s != "" {
		s = t.Lib + "." + s
	}
	if t.Array {
		s += "()"
	}
	return s
}

// lineContinuationRe は行継続（空白 + "_" + 改行 + 次行の字下げ）にマッチする。
var lineContinuationRe = regexp.MustCompile(`[ \t\x{3000}]+_[ \t\x{3000}]*\r?\n[ \t\x{3000}]*`)

// ParseTypeName は As 句などに書かれた型名の生テキストを正規化する。
// 行継続と空白を除き、"[Long]"→Long、"String * 10"→String、末尾の "()"→Array とし、
// 最後の "." で Lib と Name に分ける。
func ParseTypeName(raw string) TypeRef {
	s := lineContinuationRe.ReplaceAllString(raw, "")
	s = strings.Join(strings.FieldsFunc(s, isSpaceRune), "")
	if i := strings.IndexByte(s, '*'); i >= 0 {
		s = s[:i] // String * 10
	}
	var ref TypeRef
	if strings.HasSuffix(s, "()") {
		ref.Array = true
		s = strings.TrimSuffix(s, "()")
	}
	s = strings.NewReplacer("[", "", "]", "").Replace(s)
	if i := strings.LastIndexByte(s, '.'); i >= 0 {
		ref.Lib, ref.Name = s[:i], s[i+1:]
	} else {
		ref.Name = s
	}
	ref.Name = strings.TrimRight(ref.Name, "%&!#@$^")
	return ref
}

func isSpaceRune(r rune) bool {
	return r == ' ' || r == '\t' || r == '\r' || r == '\n' || r == '　'
}

// TypeRefFromProgID は CreateObject の ProgID から型参照を作る。
// "Scripting.Dictionary"→{Scripting, Dictionary}、"Excel.Application.16"→{Excel, Application}（末尾の数字成分を除く）。
func TypeRefFromProgID(progID string) TypeRef {
	parts := strings.Split(strings.TrimSpace(progID), ".")
	for len(parts) > 1 && isAllDigits(parts[len(parts)-1]) {
		parts = parts[:len(parts)-1]
	}
	if len(parts) == 1 {
		return TypeRef{Name: parts[0]}
	}
	return TypeRef{Lib: parts[0], Name: parts[len(parts)-1]}
}

func isAllDigits(s string) bool {
	if s == "" {
		return false
	}
	for _, r := range s {
		if r < '0' || r > '9' {
			return false
		}
	}
	return true
}

// suffixTypes は型接尾辞と型名の対応。
var suffixTypes = map[string]string{"%": "Integer", "&": "Long", "!": "Single", "#": "Double", "@": "Currency", "$": "String", "^": "LongLong"}

// TypeKind は型の種類。
type TypeKind int

const (
	TypeKindClass TypeKind = iota
	TypeKindInterface
	TypeKindUDT
	TypeKindEnum
	TypeKindModule // 標準モジュール（moduleLib 修飾でだけ引ける）
	TypeKindPrimitive
)

// TypeInfo は解決済みの型情報。
type TypeInfo struct {
	Ref           TypeRef // Lib は正規化済み（"VBA" / "Scripting" / "CATIA" など）
	Kind          TypeKind
	Complete      bool   // メンバー一覧が網羅的。true のときだけ「存在しない」診断を出す
	DefaultMember string // "Item" など。空なら resolver が "Item" の有無で推定する
	Doc           string
	Base          TypeRef     // 自プロバイダで見つからないメンバーの委譲先（フォーム → MSForms.UserForm）
	Origin        *DeclOrigin // ワークスペースの型なら宣言の場所（定義ジャンプ用）
}

// DeclOrigin は宣言の場所（定義ジャンプ・ホバー用）。Syms が空ならモジュールの先頭（0:0）。
type DeclOrigin struct {
	URI    string
	Module string
	Syms   []*Symbol // Property は Get / Let / Set を行順に
}

// MemberKind はメンバーの種類。
type MemberKind int

const (
	MemberKindProperty MemberKind = iota
	MemberKindMethod
	MemberKindField      // UDT のメンバー
	MemberKindEnumMember // Enum の値
	MemberKindEvent
)

// TLBParam はメソッドの引数情報。
type TLBParam struct {
	Name       string
	Type       string
	Optional   bool
	ByVal      bool
	ParamArray bool
	Array      bool
	Default    string
}

// MemberInfo は型のメンバー 1 件の情報。
type MemberInfo struct {
	Name       string
	ReturnType string // 生の型名（表示用）
	Kind       MemberKind
	Params     []TLBParam
	ReadOnly   bool
	Doc        string     // 2-1: helpstring
	Visibility Visibility  // Private メンバーは自モジュール（Me.）からだけ見える
	Owner      TypeRef     // 定義元（継承元の表示用）
	Origin     *DeclOrigin // ワークスペースのメンバーなら宣言（TLB / 組み込みは nil）
}

// Signature はホバー・補完の詳細表示用のシグネチャ文字列を返す。
func (m MemberInfo) Signature() string {
	ret := ""
	if m.ReturnType != "" && m.ReturnType != "void" {
		ret = " As " + m.ReturnType
	}
	switch m.Kind {
	case MemberKindProperty, MemberKindField:
		if len(m.Params) == 0 {
			return m.Name + ret
		}
	case MemberKindEnumMember:
		return m.Name
	}
	return m.Name + "(" + formatParams(m.Params) + ")" + ret
}

func formatParams(ps []TLBParam) string {
	parts := make([]string, len(ps))
	for i, p := range ps {
		s := p.Name
		if p.ParamArray {
			s = "ParamArray " + s
		} else if p.ByVal {
			s = "ByVal " + s
		}
		if p.Array {
			s += "()"
		}
		if p.Type != "" {
			s += " As " + p.Type
		}
		if p.Optional {
			if p.Default != "" {
				s += " = " + p.Default
			}
			s = "[" + s + "]"
		}
		parts[i] = s
	}
	return strings.Join(parts, ", ")
}

// TypeProvider は型解決の供給元 1 つ。メソッドは並行呼び出しに対して安全であること。
type TypeProvider interface {
	// OwnsLibrary は修飾子 lib（大文字小文字無視）をこのプロバイダが担当するかを返す。
	OwnsLibrary(lib string) bool
	// LookupType は無修飾の型名を引き、Lib を埋めた TypeInfo を返す。
	LookupType(name string) (TypeInfo, bool)
	// Members は継承込みのメンバーを返す。重複は派生側優先で除き、名前順に並べる。
	Members(t TypeRef) []MemberInfo
	// Member は名前（大文字小文字無視）でメンバーを引く。
	Member(t TypeRef, name string) (MemberInfo, bool)
	// MemberType はメンバーの型を返す。型が自プロバイダ内にあれば Lib を付け、そうでなければ Lib="" で返す。
	// void や型なしは false。
	MemberType(t TypeRef, name string) (TypeRef, bool)
}

// GlobalProvider はライブラリの暗黙のグローバル（2-1: CATIA、2-2: Err/Debug）を提供する。
// 標準モジュールの Public は ProjectProvider が扱う。
type GlobalProvider interface {
	Global(name string) (TypeRef, bool)
}

// TypeLister は型の一覧を提供する（2-2: As / New の後の型補完用）。
type TypeLister interface {
	Types() []TypeInfo
}

// qualifiedOnlyProvider は修飾付きの参照（OwnsLibrary が true の Lib）でしか型を引かせないプロバイダ
// （標準モジュール。"Dim p As Part" が同名の標準モジュールに化けないようにする）。
type qualifiedOnlyProvider interface{ QualifiedOnly() bool }

// ProjectProvider はプロジェクト（ワークスペース）の名前解決を提供する。
type ProjectProvider interface {
	// ModuleRef は標準モジュール名を TypeRef{Lib: moduleLib} に、フォーム名を既定インスタンスの型に解決する。
	ModuleRef(name string) (TypeRef, bool)
	// ProjectMember は修飾なしで参照できる標準モジュールの Public メンバーと、その所属モジュールの型を返す。
	ProjectMember(name string) (owner TypeRef, m MemberInfo, ok bool)
}

// NullProvider は何も解決しないプロバイダ。
type NullProvider struct{}

func (NullProvider) OwnsLibrary(string) bool                    { return false }
func (NullProvider) LookupType(string) (TypeInfo, bool)         { return TypeInfo{}, false }
func (NullProvider) Members(TypeRef) []MemberInfo               { return nil }
func (NullProvider) Member(TypeRef, string) (MemberInfo, bool)  { return MemberInfo{}, false }
func (NullProvider) MemberType(TypeRef, string) (TypeRef, bool) { return TypeRef{}, false }

// CompositeResolver は複数のプロバイダを優先順に問い合わせて型を解決する。
type CompositeResolver struct {
	providers []TypeProvider
}

// isNilProvider は interface の nil と、具体型の nil ポインタ（typed nil）の両方を検出する。
func isNilProvider(p TypeProvider) bool {
	if p == nil {
		return true
	}
	v := reflect.ValueOf(p)
	return v.Kind() == reflect.Ptr && v.IsNil()
}

// NewCompositeResolver は nil を除いたプロバイダで resolver を作る。
func NewCompositeResolver(ps ...TypeProvider) *CompositeResolver {
	c := &CompositeResolver{}
	for _, p := range ps {
		if !isNilProvider(p) {
			c.providers = append(c.providers, p)
		}
	}
	return c
}

// NewDefaultResolver は既定の解決順（VBA の参照設定の優先順）でプロバイダを並べる。
// Workspace（クラス・フォーム・標準モジュール）→ VBA 組み込み → TLB(CATIA) → Scripting 組み込み → MSForms。
// VBA でもプロジェクト内の名前は参照設定より優先される。nil の workspace は除かれる。
func NewDefaultResolver(db *TLBDatabase, workspace ...TypeProvider) *CompositeResolver {
	b := embeddedBuiltins()
	ps := append([]TypeProvider{}, workspace...)
	ps = append(ps, b.Library("VBA"), newTLBProvider(db), b.Library("Scripting"), b.Library("MSForms"))
	return NewCompositeResolver(ps...)
}

// WithLocal は local（nil は除く）を先頭に足した浅いコピーを返す。足すものがなければ c 自身。
func (c *CompositeResolver) WithLocal(local ...TypeProvider) *CompositeResolver {
	var add []TypeProvider
	for _, p := range local {
		if !isNilProvider(p) {
			add = append(add, p)
		}
	}
	if len(add) == 0 {
		return c
	}
	ps := make([]TypeProvider, 0, len(c.providers)+len(add))
	ps = append(ps, add...)
	ps = append(ps, c.providers...)
	return &CompositeResolver{providers: ps}
}

// Resolve は型参照を解決する。修飾ありなら担当プロバイダだけに問い合わせ、担当がなければ未解決とする
// （"Excel.Application" を CATIA の Application に誤解決しないため）。
// 修飾専用のプロバイダ（標準モジュール）は無修飾の参照では問い合わせない。
func (c *CompositeResolver) Resolve(ref TypeRef) (TypeInfo, TypeProvider, bool) {
	if c == nil || ref.Name == "" {
		return TypeInfo{}, nil, false
	}
	for _, p := range c.providers {
		if ref.Lib == "" {
			if q, ok := p.(qualifiedOnlyProvider); ok && q.QualifiedOnly() {
				continue
			}
		} else if !p.OwnsLibrary(ref.Lib) {
			continue
		}
		if ti, ok := p.LookupType(ref.Name); ok {
			ti.Ref.Array = ref.Array
			return ti, p, true
		}
	}
	return TypeInfo{}, nil, false
}

// TypeExists は型が解決でき、かつ primitive・標準モジュールでないかを返す。
func (c *CompositeResolver) TypeExists(ref TypeRef) bool {
	ti, _, ok := c.Resolve(ref.Elem())
	return ok && ti.Kind != TypeKindPrimitive && ti.Kind != TypeKindModule
}

// maxBaseDepth は TypeInfo.Base をたどる深さの上限（循環の防止）。
const maxBaseDepth = 4

// Members は型のメンバー一覧を返す。配列なら nil。Base があれば合成する（同名は自分側を優先、名前順）。
func (c *CompositeResolver) Members(ref TypeRef) []MemberInfo { return c.membersDepth(ref, 0) }

func (c *CompositeResolver) membersDepth(ref TypeRef, d int) []MemberInfo {
	if ref.Array || d > maxBaseDepth {
		return nil
	}
	ti, p, ok := c.Resolve(ref)
	if !ok {
		return nil
	}
	own := p.Members(ti.Ref)
	if ti.Base.IsZero() {
		return own
	}
	return mergeMembers(own, c.membersDepth(ti.Base, d+1))
}

// mergeMembers は own と base を名前（大文字小文字無視）で合わせ、同名は own を優先して名前順に並べる。
func mergeMembers(own, base []MemberInfo) []MemberInfo {
	if len(base) == 0 {
		return own
	}
	seen := make(map[string]bool, len(own))
	out := make([]MemberInfo, 0, len(own)+len(base))
	for _, m := range own {
		seen[strings.ToLower(m.Name)] = true
		out = append(out, m)
	}
	for _, m := range base {
		if !seen[strings.ToLower(m.Name)] {
			out = append(out, m)
		}
	}
	sort.SliceStable(out, func(i, j int) bool { return strings.ToLower(out[i].Name) < strings.ToLower(out[j].Name) })
	return out
}

// Member は型のメンバーを名前で引く（自分になければ Base）。
func (c *CompositeResolver) Member(ref TypeRef, name string) (MemberInfo, bool) {
	return c.memberDepth(ref, name, 0)
}

func (c *CompositeResolver) memberDepth(ref TypeRef, name string, d int) (MemberInfo, bool) {
	if ref.Array || d > maxBaseDepth {
		return MemberInfo{}, false
	}
	ti, p, ok := c.Resolve(ref)
	if !ok {
		return MemberInfo{}, false
	}
	if m, ok := p.Member(ti.Ref, name); ok {
		return m, true
	}
	if ti.Base.IsZero() {
		return MemberInfo{}, false
	}
	return c.memberDepth(ti.Base, name, d+1)
}

// MemberType はメンバーの型を返す（自分になければ Base）。
func (c *CompositeResolver) MemberType(ref TypeRef, name string) (TypeRef, bool) {
	return c.memberTypeDepth(ref, name, 0)
}

func (c *CompositeResolver) memberTypeDepth(ref TypeRef, name string, d int) (TypeRef, bool) {
	if ref.Array || d > maxBaseDepth {
		return TypeRef{}, false
	}
	ti, p, ok := c.Resolve(ref)
	if !ok {
		return TypeRef{}, false
	}
	if _, ok := p.Member(ti.Ref, name); ok || ti.Base.IsZero() {
		return p.MemberType(ti.Ref, name)
	}
	return c.memberTypeDepth(ti.Base, name, d+1)
}

// DefaultMember は既定メンバー名を返す。TypeInfo.DefaultMember が空で、Item メンバーがあれば "Item"。
func (c *CompositeResolver) DefaultMember(ref TypeRef) string {
	ti, p, ok := c.Resolve(ref.Elem())
	if !ok {
		return ""
	}
	if ti.DefaultMember != "" {
		return ti.DefaultMember
	}
	if _, ok := p.Member(ti.Ref, "Item"); ok {
		return "Item"
	}
	return ""
}

// ModuleRef は ProjectProvider を実装するプロバイダへ順に問い合わせる（Main. / frmMenu. の修飾子）。
func (c *CompositeResolver) ModuleRef(name string) (TypeRef, bool) {
	if c == nil {
		return TypeRef{}, false
	}
	for _, p := range c.providers {
		if pp, ok := p.(ProjectProvider); ok {
			if t, ok := pp.ModuleRef(name); ok {
				return t, true
			}
		}
	}
	return TypeRef{}, false
}

// ProjectMember は ProjectProvider を実装するプロバイダへ順に問い合わせる（標準モジュールの Public を修飾なしで）。
func (c *CompositeResolver) ProjectMember(name string) (TypeRef, MemberInfo, bool) {
	if c == nil {
		return TypeRef{}, MemberInfo{}, false
	}
	for _, p := range c.providers {
		if pp, ok := p.(ProjectProvider); ok {
			if owner, m, ok := pp.ProjectMember(name); ok {
				return owner, m, true
			}
		}
	}
	return TypeRef{}, MemberInfo{}, false
}

// Types は TypeLister を実装するプロバイダの型をすべて返す（2-2: As / New の後の型補完用）。
func (c *CompositeResolver) Types() []TypeInfo {
	if c == nil {
		return nil
	}
	var out []TypeInfo
	for _, p := range c.providers {
		if l, ok := p.(TypeLister); ok {
			out = append(out, l.Types()...)
		}
	}
	return out
}

// Global は GlobalProvider を実装するプロバイダへ順に問い合わせる。
func (c *CompositeResolver) Global(name string) (TypeRef, bool) {
	if c == nil {
		return TypeRef{}, false
	}
	for _, p := range c.providers {
		if g, ok := p.(GlobalProvider); ok {
			if t, ok := g.Global(name); ok {
				return t, true
			}
		}
	}
	return TypeRef{}, false
}
