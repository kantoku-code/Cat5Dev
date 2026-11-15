package main

import (
	"hash/fnv"
	"sort"
	"strconv"
	"strings"
	"time"
)

// ModuleEntry は 1 モジュールの解析結果と、他モジュールから見える面（公開メンバー）。
// 作成後は不変（更新は新しい entry に置き換える）なので、ロックなしで並行に読める。
type ModuleEntry struct {
	Name       string
	Kind       ModuleKind
	Path       string // 絶対パス（file: 以外のバッファなら ""）
	Key        string // pathKey(Path)（file: 以外なら URI）
	URI        string // Location に使う URI（開いていればクライアントの URI）
	Table      *SymbolTable
	Doc        string // ファイル先頭のコメント
	FromBuffer bool
	BufGen     int      // バッファ由来なら docState.gen
	TextHash   uint64   // 本文のハッシュ（didClose 時にディスクと同じかを判定する）
	Stat       fileStat // ディスク由来の size / modTime（Reconcile 用）

	public  *memberIndex // Public / Friend / Global のみ（他モジュールから見える面）
	all     *memberIndex // Private を含む（自モジュール・Me. 用）
	Events  []*Symbol    // Event 宣言（o. の候補には出さない）
	surface uint64       // public の名前・種別・シグネチャのハッシュ（公開面が変わったかの判定）
}

type fileStat struct {
	Size    int64
	ModTime time.Time
}

// memberIndex はメンバーの一覧と名前引きの表。
type memberIndex struct {
	list   []MemberInfo // 名前順
	types  []TypeRef    // list と同じ添字。メンバーの宣言型（Sub や型なしはゼロ値）
	byName map[string]int
}

func (mi *memberIndex) lookup(name string) (int, bool) {
	if mi == nil {
		return 0, false
	}
	i, ok := mi.byName[strings.ToLower(name)]
	return i, ok
}

// newModuleEntry は解析済みの表からモジュールの entry を作る。
func newModuleEntry(name string, kind ModuleKind, path, uri, text string, lines []string, table *SymbolTable) *ModuleEntry {
	if table == nil {
		table = &SymbolTable{}
	}
	e := &ModuleEntry{Name: name, Kind: kind, Path: path, URI: uri, Table: table, TextHash: textHash(text)}
	if path != "" {
		e.Key = pathKey(path)
	} else {
		e.Key = uri
	}
	e.Doc = moduleDoc(lines)
	e.all = buildMemberIndex(e, true)
	e.public = buildMemberIndex(e, false)
	for i := range table.Symbols {
		s := &table.Symbols[i]
		if s.Kind == SymEvent && s.Owner == nil {
			e.Events = append(e.Events, s)
		}
	}
	e.surface = surfaceHash(e.public)
	return e
}

// newModuleEntryForURI は開いている文書の entry を作る。*_utf 以外の拡張子なら nil。
func newModuleEntryForURI(uri, text string, lines []string, table *SymbolTable, gen int) *ModuleEntry {
	name, kind, ok := moduleIdentityOfURI(uri)
	if !ok {
		return nil
	}
	path, _ := uriToPath(uri)
	e := newModuleEntry(name, kind, path, uri, text, lines, table)
	e.FromBuffer, e.BufGen = true, gen
	return e
}

func textHash(s string) uint64 {
	h := fnv.New64a()
	_, _ = h.Write([]byte(s))
	return h.Sum64()
}

// selfRef はこのモジュールを指す型参照（標準モジュールは moduleLib 修飾）。
func (e *ModuleEntry) selfRef() TypeRef {
	if e.Kind == ModuleStd {
		return TypeRef{Lib: moduleLib, Name: e.Name}
	}
	return TypeRef{Name: e.Name}
}

// typeInfo はクラス / フォーム / 標準モジュールの TypeInfo を返す。
// 構文エラーのあるモジュールは宣言が欠けているかもしれないので Complete にしない（誤った「存在しない」を出さない）。
// フォームはコントロールの定義がファイルに無いため常に Complete=false。
func (e *ModuleEntry) typeInfo() TypeInfo {
	ti := TypeInfo{
		Ref:      e.selfRef(),
		Kind:     TypeKindClass,
		Complete: e.Table.Stats.SyntaxErrors == 0,
		Doc:      e.Doc,
		Origin:   &DeclOrigin{URI: e.URI, Module: e.Name},
	}
	switch e.Kind {
	case ModuleStd:
		ti.Kind = TypeKindModule
	case ModuleForm:
		ti.Complete = false
		ti.Base = formBaseType
	}
	return ti
}

func isExported(v Visibility) bool { return v == VisPublic || v == VisFriend || v == VisGlobal }

// symVisibility はシンボルの公開範囲（プロシージャは宣言の公開範囲）を返す。
func symVisibility(s *Symbol) Visibility {
	if s.Decl != nil {
		return s.Decl.Visibility
	}
	return s.Visibility
}

// toTLBParams はプロシージャの引数を表示用の引数情報に変換する（型には配列記号を含めない）。
func toTLBParams(ps []ParamInfo) []TLBParam {
	if len(ps) == 0 {
		return nil
	}
	out := make([]TLBParam, len(ps))
	for i, p := range ps {
		out[i] = TLBParam{
			Name: p.Name, Type: p.Type.Elem().String(), Optional: p.Optional, ByVal: p.ByVal,
			ParamArray: p.ParamArray, Array: p.Type.Array, Default: p.Default,
		}
	}
	return out
}

// buildMemberIndex はモジュールのメンバー表を作る。
// 対象はモジュールレベルの変数・定数・プロシージャ・Property（Get/Let/Set は 1 件に統合）。
// Event・UDT・Enum・Enum メンバーは含めない。同名が複数あれば（VBA ではコンパイルエラー）行順で最初のもの。
func buildMemberIndex(e *ModuleEntry, includePrivate bool) *memberIndex {
	st := e.Table
	owner := e.selfRef()
	type group struct {
		syms []*Symbol
	}
	order := []string{}
	groups := map[string]*group{}
	for i := range st.Symbols {
		s := &st.Symbols[i]
		if s.Scope != "module" || s.Owner != nil {
			continue
		}
		switch s.Kind {
		case SymVariable, SymConst, SymProcedure, SymProperty:
		default:
			continue
		}
		key := strings.ToLower(s.Name)
		g := groups[key]
		if g == nil {
			g = &group{}
			groups[key] = g
			order = append(order, key)
		}
		g.syms = append(g.syms, s)
	}
	mi := &memberIndex{byName: map[string]int{}}
	type built struct {
		m MemberInfo
		t TypeRef
	}
	var out []built
	for _, key := range order {
		syms := groups[key].syms
		sort.SliceStable(syms, func(i, j int) bool {
			if syms[i].Line != syms[j].Line {
				return syms[i].Line < syms[j].Line
			}
			return syms[i].Col < syms[j].Col
		})
		var m MemberInfo
		var t TypeRef
		var ok bool
		if syms[0].Kind == SymProperty {
			m, t, ok = propertyMember(syms, includePrivate)
		} else {
			m, t, ok = plainMember(syms[0], includePrivate)
		}
		if !ok {
			continue
		}
		m.Owner = owner
		if m.Origin != nil {
			m.Origin.URI, m.Origin.Module = e.URI, e.Name
		}
		out = append(out, built{m, t})
	}
	sort.SliceStable(out, func(i, j int) bool { return strings.ToLower(out[i].m.Name) < strings.ToLower(out[j].m.Name) })
	for i, b := range out {
		mi.list = append(mi.list, b.m)
		mi.types = append(mi.types, b.t)
		mi.byName[strings.ToLower(b.m.Name)] = i
	}
	return mi
}

// plainMember は変数・定数・Sub / Function / Declare をメンバーにする。
func plainMember(s *Symbol, includePrivate bool) (MemberInfo, TypeRef, bool) {
	vis := symVisibility(s)
	if !includePrivate && !isExported(vis) {
		return MemberInfo{}, TypeRef{}, false
	}
	m := MemberInfo{Name: s.Name, Visibility: vis, Doc: s.Doc, Origin: &DeclOrigin{Syms: []*Symbol{s}}}
	var t TypeRef
	switch s.Kind {
	case SymVariable, SymConst:
		m.Kind = MemberKindProperty
		m.ReadOnly = s.Kind == SymConst
		t = s.TypeRef()
		m.ReturnType = t.String()
	case SymProcedure:
		m.Kind = MemberKindMethod
		if d := s.Decl; d != nil {
			m.Params = toTLBParams(d.Params)
			if m.Doc == "" {
				m.Doc = d.Doc
			}
			if d.HasReturn() {
				t = d.ReturnType
				m.ReturnType = t.String()
			}
		}
	}
	return m, t, true
}

// propertyMember は同名の Property Get / Let / Set を 1 件のメンバーにまとめる。
// 型は Get の戻り値（なければ Let/Set の最後の引数）、引数は Get の引数（なければ Let/Set の最後を除いたもの）。
// Let も Set もなければ ReadOnly。公開範囲はアクセサのうち最も公開度の高いもの。
func propertyMember(syms []*Symbol, includePrivate bool) (MemberInfo, TypeRef, bool) {
	var get, letset *ProcInfo
	vis := VisPrivate
	exported := false
	doc := ""
	var origin []*Symbol
	for _, s := range syms {
		d := s.Decl
		if d == nil {
			continue
		}
		origin = append(origin, s)
		if isExported(d.Visibility) {
			exported = true
			if vis == VisPrivate || d.Visibility == VisPublic {
				vis = d.Visibility
			}
		}
		switch d.Kind {
		case ProcPropertyGet:
			if get == nil {
				get = d
			}
		default:
			if letset == nil {
				letset = d
			}
		}
	}
	if len(origin) == 0 || (!includePrivate && !exported) {
		return MemberInfo{}, TypeRef{}, false
	}
	// Doc は Get → Let → Set の順で最初の空でないもの
	for _, k := range []ProcKind{ProcPropertyGet, ProcPropertyLet, ProcPropertySet} {
		for _, s := range origin {
			if doc == "" && s.Decl.Kind == k && s.Decl.Doc != "" {
				doc = s.Decl.Doc
			}
		}
	}
	m := MemberInfo{Name: origin[0].Name, Kind: MemberKindProperty, Visibility: vis, Doc: doc, ReadOnly: letset == nil,
		Origin: &DeclOrigin{Syms: origin}}
	var t TypeRef
	switch {
	case get != nil:
		t = get.ReturnType
		m.Params = toTLBParams(get.Params)
	case letset != nil && len(letset.Params) > 0:
		t = letset.Params[len(letset.Params)-1].Type
		m.Params = toTLBParams(letset.Params[:len(letset.Params)-1])
	}
	m.ReturnType = t.String()
	return m, t, true
}

// surfaceHash は公開面（名前・種別・シグネチャ・型）のハッシュ。本体だけの変更では変わらない。
func surfaceHash(mi *memberIndex) uint64 {
	h := fnv.New64a()
	for i, m := range mi.list {
		_, _ = h.Write([]byte(m.Name + "|" + strconv.Itoa(int(m.Kind)) + "|" + m.Signature() + "|" + mi.types[i].String() + "\n"))
	}
	return h.Sum64()
}

// moduleDoc はファイル先頭のコメント（空行を飛ばし、連続するコメント行）を最大 8 行返す。装飾は除く。
func moduleDoc(lines []string) string {
	const maxLines = 8
	var parts []string
	i := 0
	for i < len(lines) && strings.TrimSpace(lines[i]) == "" {
		i++
	}
	for ; i < len(lines) && len(parts) < maxLines; i++ {
		text, ok := commentText(lines[i])
		if !ok {
			break
		}
		if t := strings.Trim(text, docTrim); t != "" {
			parts = append(parts, t)
		}
	}
	return strings.Join(parts, "\n")
}

// --- selfProvider ---

// selfProvider は現在のモジュール自身の型を Private を含めて提供する（Me. / 標準モジュール内の Module1.Private）。
// Document ごとに解決順の先頭付近へ置く。標準モジュールは moduleLib 修飾でしか引かせない。
type selfProvider struct{ e *ModuleEntry }

func newSelfProvider(e *ModuleEntry) TypeProvider {
	if e == nil {
		return nil
	}
	return &selfProvider{e: e}
}

func (p *selfProvider) QualifiedOnly() bool { return p.e.Kind == ModuleStd }

func (p *selfProvider) OwnsLibrary(lib string) bool {
	return p.e.Kind == ModuleStd && strings.EqualFold(lib, moduleLib)
}

func (p *selfProvider) LookupType(name string) (TypeInfo, bool) {
	if !strings.EqualFold(name, p.e.Name) {
		return TypeInfo{}, false
	}
	return p.e.typeInfo(), true
}

func (p *selfProvider) Members(TypeRef) []MemberInfo { return p.e.all.list }

func (p *selfProvider) Member(_ TypeRef, name string) (MemberInfo, bool) {
	if i, ok := p.e.all.lookup(name); ok {
		return p.e.all.list[i], true
	}
	return MemberInfo{}, false
}

func (p *selfProvider) MemberType(_ TypeRef, name string) (TypeRef, bool) {
	if i, ok := p.e.all.lookup(name); ok {
		t := p.e.all.types[i]
		return t, !t.IsZero()
	}
	return TypeRef{}, false
}

// ModuleRef は自モジュールの名前を修飾子として解決する（ワークスペース外のファイルでも Module1.Priv が引けるように）。
func (p *selfProvider) ModuleRef(name string) (TypeRef, bool) {
	if !strings.EqualFold(name, p.e.Name) || p.e.Kind == ModuleClass {
		return TypeRef{}, false
	}
	return p.e.selfRef(), true
}

// ProjectMember は扱わない（自モジュールの非修飾の名前はシンボル表で解決する）。
func (p *selfProvider) ProjectMember(string) (TypeRef, MemberInfo, bool) {
	return TypeRef{}, MemberInfo{}, false
}
