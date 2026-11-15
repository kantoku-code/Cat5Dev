package main

import (
	"reflect"
	"regexp"
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

// String は表示用の文字列（"Scripting.Dictionary" / "String()"）を返す。
func (t TypeRef) String() string {
	s := t.Name
	if t.Lib != "" && s != "" {
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
	TypeKindModule // 2-3 で使用
	TypeKindPrimitive
)

// TypeInfo は解決済みの型情報。
type TypeInfo struct {
	Ref           TypeRef // Lib は正規化済み（"VBA" / "Scripting" / "CATIA" など）
	Kind          TypeKind
	Complete      bool   // メンバー一覧が網羅的。true のときだけ「存在しない」診断を出す
	DefaultMember string // "Item" など。空なら resolver が "Item" の有無で推定する
	Doc           string
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
	Visibility Visibility // 2-3: Private メンバー（Me. のときだけ表示）
	Owner      TypeRef    // 定義元（継承元の表示用）
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

// GlobalProvider は暗黙のグローバル（2-1: CATIA、2-2: Err/Debug、2-3: 標準モジュールの Public）を提供する。
type GlobalProvider interface {
	Global(name string) (TypeRef, bool)
}

// TypeLister は型の一覧を提供する（2-2: As / New の後の型補完用）。
type TypeLister interface {
	Types() []TypeInfo
}

// NullProvider は何も解決しないプロバイダ。WorkspaceProvider（2-3）が実装されるまでの差し込み口。
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
// Workspace(2-3) → VBA 組み込み → TLB(CATIA) → Scripting 組み込み。
func NewDefaultResolver(db *TLBDatabase, workspace TypeProvider) *CompositeResolver {
	b := embeddedBuiltins()
	return NewCompositeResolver(workspace, b.Library("VBA"), newTLBProvider(db), b.Library("Scripting"))
}

// WithLocal は local を先頭に足した浅いコピーを返す。
func (c *CompositeResolver) WithLocal(local TypeProvider) *CompositeResolver {
	if isNilProvider(local) {
		return c
	}
	ps := make([]TypeProvider, 0, len(c.providers)+1)
	ps = append(ps, local)
	ps = append(ps, c.providers...)
	return &CompositeResolver{providers: ps}
}

// Resolve は型参照を解決する。修飾ありなら担当プロバイダだけに問い合わせ、担当がなければ未解決とする
// （"Excel.Application" を CATIA の Application に誤解決しないため）。
func (c *CompositeResolver) Resolve(ref TypeRef) (TypeInfo, TypeProvider, bool) {
	if c == nil || ref.Name == "" {
		return TypeInfo{}, nil, false
	}
	for _, p := range c.providers {
		if ref.Lib != "" && !p.OwnsLibrary(ref.Lib) {
			continue
		}
		if ti, ok := p.LookupType(ref.Name); ok {
			ti.Ref.Array = ref.Array
			return ti, p, true
		}
	}
	return TypeInfo{}, nil, false
}

// TypeExists は型が解決でき、かつ primitive でないかを返す。
func (c *CompositeResolver) TypeExists(ref TypeRef) bool {
	ti, _, ok := c.Resolve(ref.Elem())
	return ok && ti.Kind != TypeKindPrimitive
}

// Members は型のメンバー一覧を返す。配列なら nil。
func (c *CompositeResolver) Members(ref TypeRef) []MemberInfo {
	if ref.Array {
		return nil
	}
	ti, p, ok := c.Resolve(ref)
	if !ok {
		return nil
	}
	return p.Members(ti.Ref)
}

// Member は型のメンバーを名前で引く。
func (c *CompositeResolver) Member(ref TypeRef, name string) (MemberInfo, bool) {
	if ref.Array {
		return MemberInfo{}, false
	}
	ti, p, ok := c.Resolve(ref)
	if !ok {
		return MemberInfo{}, false
	}
	return p.Member(ti.Ref, name)
}

// MemberType はメンバーの型を返す。
func (c *CompositeResolver) MemberType(ref TypeRef, name string) (TypeRef, bool) {
	if ref.Array {
		return TypeRef{}, false
	}
	ti, p, ok := c.Resolve(ref)
	if !ok {
		return TypeRef{}, false
	}
	return p.MemberType(ti.Ref, name)
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
