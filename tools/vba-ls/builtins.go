package main

import (
	_ "embed"
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"sync"
)

// vba_builtins.json は VBA / Scripting / MSForms の組み込み型と、VBA の組み込み関数・定数・列挙・キーワードの定義。
//
//go:embed vba_builtins.json
var builtinJSON []byte

type builtinFile struct {
	Version   int              `json:"version"`
	Libraries []builtinLibrary `json:"libraries"`
	Keywords  []builtinKeyword `json:"keywords,omitempty"`
}

type builtinLibrary struct {
	Name       string          `json:"name"`
	Aliases    []string        `json:"aliases,omitempty"`
	Primitives []string        `json:"primitives,omitempty"`
	Types      []builtinType   `json:"types"`
	Globals    []builtinGlobal `json:"globals,omitempty"`   // Err / Debug（2-1: {"name":"CATIA","type":"CATIA.Application"}）
	Functions  []builtinFunc   `json:"functions,omitempty"` // 組み込み関数・ステートメント
	Constants  []builtinConst  `json:"constants,omitempty"` // 列挙に属さない定数（vbCrLf など）
}

type builtinGlobal struct {
	Name string `json:"name"`
	Type string `json:"type"`
	Doc  string `json:"doc,omitempty"`
}

type builtinType struct {
	Name          string          `json:"name"`
	Kind          string          `json:"kind"` // class | interface | enum
	Complete      bool            `json:"complete"`
	Creatable     bool            `json:"creatable,omitempty"` // New で作れる
	DefaultMember string          `json:"defaultMember,omitempty"`
	Doc           string          `json:"doc,omitempty"`
	Members       []builtinMember `json:"members"`
}

type builtinMember struct {
	Name     string         `json:"name"`
	Kind     string         `json:"kind"` // method | property（class / interface）、enumMember（enum）
	Type     string         `json:"type,omitempty"`
	ReadOnly bool           `json:"readOnly,omitempty"`
	Params   []builtinParam `json:"params,omitempty"`
	Value    string         `json:"value,omitempty"` // enumMember の値（表示用の VBA の式）
	Doc      string         `json:"doc,omitempty"`
}

type builtinParam struct {
	Name       string `json:"name"`
	Type       string `json:"type,omitempty"` // 省略は Variant（表示では As を付けない）
	Optional   bool   `json:"optional,omitempty"`
	Default    string `json:"default,omitempty"`
	ParamArray bool   `json:"paramArray,omitempty"`
}

// builtinFunc は組み込み関数・ステートメント 1 つ。
type builtinFunc struct {
	Name          string         `json:"name"`
	Kind          string         `json:"kind"`   // function | sub | property（Date / Now など括弧なしで値を返すもの）
	Module        string         `json:"module"` // Strings | Conversion | Math | DateTime | Information | Interaction | FileSystem
	Type          string         `json:"type,omitempty"`
	Params        []builtinParam `json:"params,omitempty"`
	StringVariant bool           `json:"stringVariant,omitempty"` // Left$ のように String を返す $ 形もある
	Doc           string         `json:"doc"`
}

// builtinConst は列挙に属さない組み込み定数。
type builtinConst struct {
	Name  string `json:"name"`
	Type  string `json:"type"`
	Value string `json:"value"`
	Doc   string `json:"doc"`
}

// builtinKeyword は VBA のキーワード 1 つ（識別子補完の末尾に出す）。
type builtinKeyword struct {
	Name  string `json:"name"`
	Usage string `json:"usage"` // statement（文頭）| expression（式・文の途中）| both
	Doc   string `json:"doc"`
}

type builtinTypeIdx struct {
	info    TypeInfo
	members []MemberInfo
	byName  map[string]int
}

// GlobalInfo は暗黙のグローバル（Err / Debug）1 つ。
type GlobalInfo struct {
	Name string
	Type TypeRef
	Doc  string
}

// BuiltinProvider は組み込みライブラリ 1 つ（VBA / Scripting / MSForms）の TypeProvider。
// 関数・定数・グローバルの列挙（FunctionProvider / GlobalLister）も提供する。
type BuiltinProvider struct {
	lib     string
	aliases map[string]bool
	types   map[string]*builtinTypeIdx
	prims   map[string]string // このライブラリの primitive（小文字 → 正式名）
	allPrim map[string]string // 全ライブラリの primitive（MemberType の修飾に使う）
	globals map[string]TypeRef

	globalList []GlobalInfo    // 名前順
	funcs      []MemberInfo    // 名前順
	funcIdx    map[string]int  // 小文字名 → funcs の添字
	strVariant map[string]bool // $ 形のある関数（小文字名）
	consts     []MemberInfo    // 名前順
	typeList   []TypeInfo      // 名前順（primitive を含まない）
	primList   []TypeInfo      // 名前順
}

// builtinSet は埋め込み JSON から作った全ライブラリ。
type builtinSet struct {
	libs     map[string]*BuiltinProvider
	keywords []builtinKeyword
}

// Library は名前（大文字小文字無視）でライブラリのプロバイダを返す。なければ interface の nil。
func (s *builtinSet) Library(name string) TypeProvider {
	if p, ok := s.libs[strings.ToLower(name)]; ok {
		return p
	}
	return nil
}

// Keywords は VBA のキーワードを JSON の順で返す。
func (s *builtinSet) Keywords() []builtinKeyword { return s.keywords }

var (
	embeddedOnce sync.Once
	embeddedSet  *builtinSet
)

// embeddedBuiltins は埋め込み JSON を初回に 1 回だけ解析して返す。JSON が壊れていれば panic する（テストで検証する）。
func embeddedBuiltins() *builtinSet {
	embeddedOnce.Do(func() {
		s, err := parseBuiltinSet(builtinJSON)
		if err != nil {
			panic(fmt.Sprintf("vba_builtins.json: %v", err))
		}
		embeddedSet = s
	})
	return embeddedSet
}

var keywordUsages = map[string]bool{"statement": true, "expression": true, "both": true}

func parseBuiltinSet(data []byte) (*builtinSet, error) {
	var f builtinFile
	if err := json.Unmarshal(data, &f); err != nil {
		return nil, err
	}
	allPrim := map[string]string{}
	for _, l := range f.Libraries {
		for _, p := range l.Primitives {
			allPrim[strings.ToLower(p)] = p
		}
	}
	s := &builtinSet{libs: map[string]*BuiltinProvider{}}
	for _, l := range f.Libraries {
		p, err := newBuiltinProvider(l, allPrim)
		if err != nil {
			return nil, err
		}
		s.libs[strings.ToLower(l.Name)] = p
	}
	seen := map[string]bool{}
	for _, k := range f.Keywords {
		key := strings.ToLower(k.Name)
		switch {
		case k.Name == "":
			return nil, fmt.Errorf("keyword: name is empty")
		case !keywordUsages[k.Usage]:
			return nil, fmt.Errorf("keyword %s: unknown usage %q", k.Name, k.Usage)
		case k.Doc == "":
			return nil, fmt.Errorf("keyword %s: doc is empty", k.Name)
		case seen[key]:
			return nil, fmt.Errorf("keyword %s: duplicate", k.Name)
		}
		seen[key] = true
		s.keywords = append(s.keywords, k)
	}
	return s, nil
}

// NewBuiltinProviderFromJSON は JSON から lib という名前のライブラリのプロバイダを作る（テスト用）。
func NewBuiltinProviderFromJSON(data []byte, lib string) (*BuiltinProvider, error) {
	s, err := parseBuiltinSet(data)
	if err != nil {
		return nil, err
	}
	p, ok := s.libs[strings.ToLower(lib)]
	if !ok {
		return nil, fmt.Errorf("library %q not found", lib)
	}
	return p, nil
}

func toParams(bps []builtinParam) []TLBParam {
	var out []TLBParam
	for _, bp := range bps {
		out = append(out, TLBParam{Name: bp.Name, Type: bp.Type, Optional: bp.Optional, Default: bp.Default, ParamArray: bp.ParamArray, Array: bp.ParamArray})
	}
	return out
}

func sortMembers(ms []MemberInfo) {
	sort.SliceStable(ms, func(i, j int) bool { return strings.ToLower(ms[i].Name) < strings.ToLower(ms[j].Name) })
}

func newBuiltinProvider(l builtinLibrary, allPrim map[string]string) (*BuiltinProvider, error) {
	if l.Name == "" {
		return nil, fmt.Errorf("library name is empty")
	}
	p := &BuiltinProvider{
		lib:        l.Name,
		aliases:    map[string]bool{strings.ToLower(l.Name): true},
		types:      map[string]*builtinTypeIdx{},
		prims:      map[string]string{},
		allPrim:    allPrim,
		globals:    map[string]TypeRef{},
		funcIdx:    map[string]int{},
		strVariant: map[string]bool{},
	}
	// 非修飾で参照できる名前（関数・定数・Enum メンバー）の重複検査
	unqualified := map[string]string{}
	claim := func(name, what string) error {
		key := strings.ToLower(name)
		if prev, dup := unqualified[key]; dup {
			return fmt.Errorf("%s %s: duplicate name (already %s)", what, name, prev)
		}
		unqualified[key] = what
		return nil
	}
	for _, a := range l.Aliases {
		p.aliases[strings.ToLower(a)] = true
	}
	for _, name := range l.Primitives {
		p.prims[strings.ToLower(name)] = name
		p.primList = append(p.primList, TypeInfo{Ref: TypeRef{Lib: l.Name, Name: name}, Kind: TypeKindPrimitive})
	}
	for _, g := range l.Globals {
		t := ParseTypeName(g.Type)
		p.globals[strings.ToLower(g.Name)] = t
		p.globalList = append(p.globalList, GlobalInfo{Name: g.Name, Type: t, Doc: g.Doc})
	}
	for _, bt := range l.Types {
		kind := TypeKindClass
		switch bt.Kind {
		case "class":
		case "interface":
			kind = TypeKindInterface
		case "enum":
			kind = TypeKindEnum
		default:
			return nil, fmt.Errorf("type %s: unknown kind %q", bt.Name, bt.Kind)
		}
		idx := &builtinTypeIdx{
			info: TypeInfo{
				Ref:           TypeRef{Lib: l.Name, Name: bt.Name},
				Kind:          kind,
				Complete:      bt.Complete,
				Creatable:     bt.Creatable,
				DefaultMember: bt.DefaultMember,
				Doc:           bt.Doc,
			},
			byName: map[string]int{},
		}
		for _, bm := range bt.Members {
			var mk MemberKind
			switch {
			case kind == TypeKindEnum && bm.Kind == "enumMember":
				mk = MemberKindEnumMember
				if err := claim(bm.Name, "enum member"); err != nil {
					return nil, err
				}
			case kind != TypeKindEnum && bm.Kind == "method":
				mk = MemberKindMethod
			case kind != TypeKindEnum && bm.Kind == "property":
				mk = MemberKindProperty
			default:
				return nil, fmt.Errorf("%s.%s: unknown member kind %q", bt.Name, bm.Name, bm.Kind)
			}
			mi := MemberInfo{
				Name:       bm.Name,
				ReturnType: bm.Type,
				Kind:       mk,
				ReadOnly:   bm.ReadOnly,
				Doc:        bm.Doc,
				Visibility: VisPublic,
				Owner:      idx.info.Ref,
				Params:     toParams(bm.Params),
				Value:      bm.Value,
			}
			if mk == MemberKindEnumMember {
				mi.ReturnType, mi.ReadOnly = "Long", true
			}
			idx.members = append(idx.members, mi)
		}
		sortMembers(idx.members)
		for i, m := range idx.members {
			key := strings.ToLower(m.Name)
			if _, dup := idx.byName[key]; dup {
				return nil, fmt.Errorf("%s: duplicate member %s", bt.Name, m.Name)
			}
			idx.byName[key] = i
		}
		p.types[strings.ToLower(bt.Name)] = idx
		p.typeList = append(p.typeList, idx.info)
	}
	for _, bf := range l.Functions {
		mi := MemberInfo{
			Name: bf.Name, ReturnType: bf.Type, Kind: MemberKindMethod, Params: toParams(bf.Params),
			Doc: bf.Doc, Visibility: VisPublic, Owner: TypeRef{Lib: l.Name, Name: bf.Module},
		}
		switch bf.Kind {
		case "function":
		case "property":
			mi.Kind, mi.ReadOnly = MemberKindProperty, true
		case "sub":
			if bf.Type != "" {
				return nil, fmt.Errorf("function %s: sub has a type", bf.Name)
			}
		default:
			return nil, fmt.Errorf("function %s: unknown kind %q", bf.Name, bf.Kind)
		}
		if bf.Kind != "sub" && bf.Type == "" {
			return nil, fmt.Errorf("function %s: type is empty", bf.Name)
		}
		if bf.Doc == "" {
			return nil, fmt.Errorf("function %s: doc is empty", bf.Name)
		}
		if err := claim(bf.Name, "function"); err != nil {
			return nil, err
		}
		if bf.StringVariant {
			p.strVariant[strings.ToLower(bf.Name)] = true
		}
		p.funcs = append(p.funcs, mi)
	}
	sortMembers(p.funcs)
	for i, f := range p.funcs {
		p.funcIdx[strings.ToLower(f.Name)] = i
	}
	for _, bc := range l.Constants {
		if bc.Doc == "" {
			return nil, fmt.Errorf("constant %s: doc is empty", bc.Name)
		}
		if err := claim(bc.Name, "constant"); err != nil {
			return nil, err
		}
		p.consts = append(p.consts, MemberInfo{
			Name: bc.Name, ReturnType: bc.Type, Kind: MemberKindConst, ReadOnly: true, Value: bc.Value,
			Doc: bc.Doc, Visibility: VisPublic, Owner: TypeRef{Lib: l.Name, Name: "Constants"},
		})
	}
	sortMembers(p.consts)
	sort.Slice(p.typeList, func(i, j int) bool { return p.typeList[i].Ref.Name < p.typeList[j].Ref.Name })
	sort.Slice(p.primList, func(i, j int) bool { return p.primList[i].Ref.Name < p.primList[j].Ref.Name })
	sort.Slice(p.globalList, func(i, j int) bool { return p.globalList[i].Name < p.globalList[j].Name })
	return p, nil
}

func (p *BuiltinProvider) OwnsLibrary(lib string) bool { return p.aliases[strings.ToLower(lib)] }

func (p *BuiltinProvider) LookupType(name string) (TypeInfo, bool) {
	key := strings.ToLower(name)
	if t, ok := p.types[key]; ok {
		return t.info, true
	}
	if canon, ok := p.prims[key]; ok {
		return TypeInfo{Ref: TypeRef{Lib: p.lib, Name: canon}, Kind: TypeKindPrimitive}, true
	}
	return TypeInfo{}, false
}

func (p *BuiltinProvider) Members(t TypeRef) []MemberInfo {
	if idx, ok := p.types[strings.ToLower(t.Name)]; ok {
		return idx.members
	}
	return nil
}

func (p *BuiltinProvider) Member(t TypeRef, name string) (MemberInfo, bool) {
	idx, ok := p.types[strings.ToLower(t.Name)]
	if !ok {
		return MemberInfo{}, false
	}
	i, ok := idx.byName[strings.ToLower(name)]
	if !ok {
		return MemberInfo{}, false
	}
	return idx.members[i], true
}

func (p *BuiltinProvider) MemberType(t TypeRef, name string) (TypeRef, bool) {
	m, ok := p.Member(t, name)
	if !ok {
		return TypeRef{}, false
	}
	return p.qualify(m.ReturnType)
}

// qualify はメンバーの型文字列を TypeRef にする。同じライブラリの型なら Lib を付け、primitive なら VBA を付ける。
func (p *BuiltinProvider) qualify(typ string) (TypeRef, bool) {
	if typ == "" || strings.EqualFold(typ, "void") {
		return TypeRef{}, false
	}
	ref := ParseTypeName(typ)
	if ref.Lib != "" {
		return ref, true
	}
	key := strings.ToLower(ref.Name)
	if t, ok := p.types[key]; ok {
		ref.Lib, ref.Name = p.lib, t.info.Ref.Name
	} else if canon, ok := p.allPrim[key]; ok {
		ref.Lib, ref.Name = "VBA", canon
	}
	return ref, true
}

// Global は暗黙のグローバル（Err / Debug）の型を返す。
func (p *BuiltinProvider) Global(name string) (TypeRef, bool) {
	t, ok := p.globals[strings.ToLower(name)]
	return t, ok
}

// Globals は暗黙のグローバルを名前順に返す。
func (p *BuiltinProvider) Globals() []GlobalInfo { return p.globalList }

// Types は型の一覧（primitive を含まない）を名前順に返す。
func (p *BuiltinProvider) Types() []TypeInfo { return p.typeList }

// Primitives は primitive 型（Long / String など）を名前順に返す（型補完用）。
func (p *BuiltinProvider) Primitives() []TypeInfo { return p.primList }

// Functions は組み込み関数・ステートメントを名前順に返す。
func (p *BuiltinProvider) Functions() []MemberInfo { return p.funcs }

// Function は名前（大文字小文字無視）で組み込み関数を引く。"Left$" のような $ 形は、$ 形のある関数だけ引け、
// 名前に $ を付けて戻り値を String にしたものを返す。
func (p *BuiltinProvider) Function(name string) (MemberInfo, bool) {
	key := strings.ToLower(name)
	str := false
	if strings.HasSuffix(key, "$") {
		key = strings.TrimSuffix(key, "$")
		if !p.strVariant[key] {
			return MemberInfo{}, false
		}
		str = true
	}
	i, ok := p.funcIdx[key]
	if !ok {
		return MemberInfo{}, false
	}
	m := p.funcs[i]
	if str {
		m.Name += "$"
		m.ReturnType = "String"
	}
	return m, true
}

// Constants は列挙に属さない組み込み定数を名前順に返す。
func (p *BuiltinProvider) Constants() []MemberInfo { return p.consts }
