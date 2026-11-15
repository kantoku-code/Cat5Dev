package main

import (
	_ "embed"
	"encoding/json"
	"fmt"
	"sort"
	"strings"
	"sync"
)

// vba_builtins.json は VBA / Scripting の組み込み型の定義。2-2 で同じ JSON に組み込み関数（functions）を追加する。
//
//go:embed vba_builtins.json
var builtinJSON []byte

type builtinFile struct {
	Version   int              `json:"version"`
	Libraries []builtinLibrary `json:"libraries"`
}

type builtinLibrary struct {
	Name       string          `json:"name"`
	Aliases    []string        `json:"aliases,omitempty"`
	Primitives []string        `json:"primitives,omitempty"`
	Types      []builtinType   `json:"types"`
	Globals    []builtinGlobal `json:"globals,omitempty"`   // 2-1: {"name":"CATIA","type":"CATIA.Application"}
	Functions  []builtinMember `json:"functions,omitempty"` // 2-2
}

type builtinGlobal struct {
	Name string `json:"name"`
	Type string `json:"type"`
}

type builtinType struct {
	Name          string          `json:"name"`
	Kind          string          `json:"kind"` // class | interface | enum
	Complete      bool            `json:"complete"`
	DefaultMember string          `json:"defaultMember,omitempty"`
	Doc           string          `json:"doc,omitempty"`
	Members       []builtinMember `json:"members"`
}

type builtinMember struct {
	Name     string         `json:"name"`
	Kind     string         `json:"kind"` // method | property
	Type     string         `json:"type,omitempty"`
	ReadOnly bool           `json:"readOnly,omitempty"`
	Params   []builtinParam `json:"params,omitempty"`
	Doc      string         `json:"doc,omitempty"`
}

type builtinParam struct {
	Name     string `json:"name"`
	Type     string `json:"type"`
	Optional bool   `json:"optional,omitempty"`
	Default  string `json:"default,omitempty"`
}

type builtinTypeIdx struct {
	info    TypeInfo
	members []MemberInfo
	byName  map[string]int
}

// BuiltinProvider は組み込みライブラリ 1 つ（VBA / Scripting）の TypeProvider。
type BuiltinProvider struct {
	lib     string
	aliases map[string]bool
	types   map[string]*builtinTypeIdx
	prims   map[string]string // このライブラリの primitive（小文字 → 正式名）
	allPrim map[string]string // 全ライブラリの primitive（MemberType の修飾に使う）
	globals map[string]TypeRef
}

// builtinSet は埋め込み JSON から作った全ライブラリ。
type builtinSet struct {
	libs map[string]*BuiltinProvider
}

// Library は名前（大文字小文字無視）でライブラリのプロバイダを返す。なければ interface の nil。
func (s *builtinSet) Library(name string) TypeProvider {
	if p, ok := s.libs[strings.ToLower(name)]; ok {
		return p
	}
	return nil
}

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

func newBuiltinProvider(l builtinLibrary, allPrim map[string]string) (*BuiltinProvider, error) {
	if l.Name == "" {
		return nil, fmt.Errorf("library name is empty")
	}
	p := &BuiltinProvider{
		lib:     l.Name,
		aliases: map[string]bool{strings.ToLower(l.Name): true},
		types:   map[string]*builtinTypeIdx{},
		prims:   map[string]string{},
		allPrim: allPrim,
		globals: map[string]TypeRef{},
	}
	for _, a := range l.Aliases {
		p.aliases[strings.ToLower(a)] = true
	}
	for _, name := range l.Primitives {
		p.prims[strings.ToLower(name)] = name
	}
	for _, g := range l.Globals {
		p.globals[strings.ToLower(g.Name)] = ParseTypeName(g.Type)
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
				DefaultMember: bt.DefaultMember,
				Doc:           bt.Doc,
			},
			byName: map[string]int{},
		}
		for _, bm := range bt.Members {
			mk := MemberKindMethod
			switch bm.Kind {
			case "method":
			case "property":
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
			}
			for _, bp := range bm.Params {
				mi.Params = append(mi.Params, TLBParam{Name: bp.Name, Type: bp.Type, Optional: bp.Optional, Default: bp.Default})
			}
			idx.members = append(idx.members, mi)
		}
		sort.SliceStable(idx.members, func(i, j int) bool {
			return strings.ToLower(idx.members[i].Name) < strings.ToLower(idx.members[j].Name)
		})
		for i, m := range idx.members {
			key := strings.ToLower(m.Name)
			if _, dup := idx.byName[key]; dup {
				return nil, fmt.Errorf("%s: duplicate member %s", bt.Name, m.Name)
			}
			idx.byName[key] = i
		}
		p.types[strings.ToLower(bt.Name)] = idx
	}
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

// Global は暗黙のグローバル（2-0 では空）を返す。
func (p *BuiltinProvider) Global(name string) (TypeRef, bool) {
	t, ok := p.globals[strings.ToLower(name)]
	return t, ok
}

// Types は型の一覧を返す（2-2 の型補完用）。
func (p *BuiltinProvider) Types() []TypeInfo {
	out := make([]TypeInfo, 0, len(p.types))
	for _, t := range p.types {
		out = append(out, t.info)
	}
	sort.Slice(out, func(i, j int) bool { return out[i].Ref.Name < out[j].Ref.Name })
	return out
}
