package main

import (
	"strings"
	"sync"
)

// candGroup は補完候補の並び順のグループ（小さいほど先）。sortText の先頭の数字になる。
// VS Code はファジー一致のスコアが同じ候補の間で sortText を使う。
type candGroup int

const (
	grpLocal   candGroup = iota // 現在のプロシージャのローカル・引数・ローカル定数
	grpModule                   // 自モジュールのモジュールレベル（Private を含む）、Me
	grpProject                  // 他の標準モジュールの Public・モジュール名・他モジュールの Public Enum とメンバー
	grpBuiltin                  // 組み込み関数・定数・Enum メンバー・Err / Debug（型の文脈では組み込み・TLB の型と primitive）
	grpKeyword                  // キーワード
)

// itemSet は名前（大文字小文字無視）で重複を除きながら候補を集める。先に入った候補が勝つ（スコープの隠蔽と同じ）。
type itemSet struct {
	items []CompletionItem
	seen  map[string]bool
}

func newItemSet(capacity int) *itemSet {
	return &itemSet{items: make([]CompletionItem, 0, capacity), seen: make(map[string]bool, capacity)}
}

func (s *itemSet) add(g candGroup, it CompletionItem) {
	key := strings.ToLower(it.Label)
	if key == "" || s.seen[key] {
		return
	}
	s.seen[key] = true
	it.SortText = string(rune('0'+g)) + key
	s.items = append(s.items, it)
}

func (s *itemSet) result() []CompletionItem {
	if len(s.items) == 0 {
		return nil
	}
	return s.items
}

func plainDoc(doc string) *MarkupContent {
	if doc == "" {
		return nil
	}
	return &MarkupContent{Kind: "plaintext", Value: doc}
}

func describe(desc string) *CompletionItemLabelDetails {
	if desc == "" {
		return nil
	}
	return &CompletionItemLabelDetails{Description: desc}
}

// --- 識別子補完 ---

// buildIdentifierItems は "." なしの識別子の候補を返す。サーバー側では絞り込まない（クライアントのファジーフィルタに任せる）。
// 並び: ローカル・引数 → 自モジュール → プロジェクト → 組み込み → キーワード。
func buildIdentifierItems(doc *Document, line int, ctx completionContext) []CompletionItem {
	set := newItemSet(512)
	st := doc.Table
	// ローカル・引数・ローカル定数
	if scope := st.ScopeAt(line); scope != "module" {
		for i := range st.Symbols {
			s := &st.Symbols[i]
			if s.Scope != scope {
				continue
			}
			switch s.Kind {
			case SymVariable, SymParam:
			case SymConst:
				if ctx.AfterCall {
					continue
				}
			default:
				continue
			}
			set.add(grpLocal, symbolItem(s))
		}
	}
	// 自モジュールのモジュールレベル
	for i := range st.Symbols {
		s := &st.Symbols[i]
		if s.Scope != "module" || s.Owner != nil || !moduleSymbolWanted(s, ctx) {
			continue
		}
		set.add(grpModule, symbolItem(s))
	}
	if doc.Module != nil && doc.Module.Kind != ModuleStd && !ctx.AfterCall {
		set.add(grpModule, keywordItem("Me"))
	}
	// プロジェクト: 他の標準モジュールの Public、モジュール名・フォーム名、他モジュールの Public Enum とメンバー
	self := ""
	if doc.Module != nil {
		self = doc.Module.Name
	}
	for _, m := range doc.Types.ProjectMembers() {
		if strings.EqualFold(m.Owner.Name, self) {
			continue // 自モジュールは自表の方を使う
		}
		if ctx.AfterCall && isConstMember(m) {
			continue
		}
		set.add(grpProject, projectMemberItem(m))
	}
	if !ctx.AfterCall {
		for _, ti := range doc.Types.ProjectModules() {
			set.add(grpProject, moduleItem(ti))
		}
		for _, ti := range doc.Types.ProjectTypes() {
			if ti.Kind != TypeKindEnum {
				continue
			}
			set.add(grpProject, typeItem(ti))
			for _, m := range doc.Types.Members(ti.Ref) {
				set.add(grpProject, enumMemberItem(m, ti.Ref.Name))
			}
		}
	} else {
		for _, ti := range doc.Types.ProjectModules() {
			set.add(grpProject, moduleItem(ti)) // Call Main.Foo の修飾子
		}
	}
	// 組み込み
	b := builtinItems()
	for _, it := range b.funcs {
		set.add(grpBuiltin, it)
	}
	if !ctx.AfterCall {
		for _, it := range b.consts {
			set.add(grpBuiltin, it)
		}
	}
	for _, it := range b.globals {
		set.add(grpBuiltin, it)
	}
	// キーワード
	if !ctx.AfterCall {
		for _, it := range keywordsFor(ctx) {
			set.add(grpKeyword, it)
		}
	}
	return set.result()
}

// moduleSymbolWanted は自モジュールのモジュールレベルのシンボルを識別子の候補にするかを返す。
func moduleSymbolWanted(s *Symbol, ctx completionContext) bool {
	switch s.Kind {
	case SymVariable, SymProcedure, SymProperty:
		return true
	case SymConst, SymEnum:
		return !ctx.AfterCall
	case SymEnumMember:
		return !ctx.AfterCall && s.Parent != ""
	}
	return false // Event・UDT（型の文脈だけ）・UDT のメンバー
}

func isConstMember(m MemberInfo) bool {
	return m.Origin != nil && len(m.Origin.Syms) > 0 && m.Origin.Syms[0].Kind == SymConst
}

// symbolItem は自モジュールのシンボルを候補にする（Detail は宣言、Documentation はドキュメントコメント）。
func symbolItem(s *Symbol) CompletionItem {
	kind := CompletionItemKindVariable
	switch s.Kind {
	case SymConst:
		kind = CompletionItemKindConstant
	case SymProcedure:
		kind = CompletionItemKindFunction
	case SymProperty:
		kind = CompletionItemKindProperty
	case SymEnum:
		kind = CompletionItemKindEnum
	case SymEnumMember:
		kind = CompletionItemKindEnumMember
	case SymUDT:
		kind = CompletionItemKindStruct
	}
	doc := s.Doc
	if doc == "" && s.Decl != nil {
		doc = s.Decl.Doc
	}
	return CompletionItem{Label: s.Name, Kind: kind, Detail: declarationText(s), Documentation: plainDoc(doc)}
}

// projectMemberItem は他の標準モジュールの Public メンバーを候補にする（所属モジュールを右側に表示）。
func projectMemberItem(m MemberInfo) CompletionItem {
	it := CompletionItem{Label: m.Name, Kind: CompletionItemKindFunction, Detail: m.Signature(),
		Documentation: plainDoc(m.Doc), LabelDetails: describe(m.Owner.Name)}
	if m.Origin != nil && len(m.Origin.Syms) > 0 {
		s := m.Origin.Syms[0]
		it.Detail = declarationText(s)
		it.Kind = symbolItem(s).Kind
	} else if m.Kind == MemberKindProperty {
		it.Kind = CompletionItemKindProperty
	}
	return it
}

// moduleItem は標準モジュール名・フォーム名（修飾子として使える名前）を候補にする。
func moduleItem(ti TypeInfo) CompletionItem {
	kw := "Module"
	if ti.Base == formBaseType {
		kw = "UserForm"
	}
	return CompletionItem{Label: ti.Ref.Name, Kind: CompletionItemKindModule, Detail: kw + " " + ti.Ref.Name, Documentation: plainDoc(ti.Doc)}
}

// enumMemberItem は Enum のメンバーを候補にする（Detail は "Enum名.メンバー = 値"）。
func enumMemberItem(m MemberInfo, enum string) CompletionItem {
	return CompletionItem{Label: m.Name, Kind: CompletionItemKindEnumMember, Detail: enum + "." + m.Signature(),
		Documentation: plainDoc(m.Doc), LabelDetails: describe(enum)}
}

// builtinMemberItem は組み込みの関数・定数を候補にする。
func builtinMemberItem(m MemberInfo) CompletionItem {
	kind := CompletionItemKindFunction
	if m.Kind == MemberKindConst {
		kind = CompletionItemKindConstant
	}
	return CompletionItem{Label: m.Name, Kind: kind, Detail: m.Signature(), Documentation: plainDoc(m.Doc), LabelDetails: describe(m.Owner.String())}
}

func keywordItem(name string) CompletionItem {
	it := CompletionItem{Label: name, Kind: CompletionItemKindKeyword}
	if k, ok := builtinItems().keywordDoc[strings.ToLower(name)]; ok {
		it.Documentation = plainDoc(k)
	}
	return it
}

// builtinItemCache は組み込みの関数・定数・グローバル・キーワードの候補（初回に 1 回だけ作る）。
type builtinItemCache struct {
	funcs, consts, globals []CompletionItem
	stmtKeywords           []CompletionItem // usage が statement / both
	inlineKeywords         []CompletionItem // stmtKeywords から宣言のキーワードを除いたもの（Then / Else の後）
	exprKeywords           []CompletionItem // usage が expression / both
	keywordDoc             map[string]string
}

var (
	builtinItemsOnce sync.Once
	builtinItemsVal  *builtinItemCache
)

// declKeywords は単一行 If の Then / Else の後には書けない宣言・ブロックのキーワード。
var declKeywords = map[string]bool{
	"dim": true, "const": true, "static": true, "public": true, "private": true, "friend": true, "global": true,
	"sub": true, "function": true, "property": true, "type": true, "enum": true, "declare": true, "event": true,
	"implements": true, "option": true, "elseif": true, "case": true, "loop": true, "next": true, "wend": true,
}

func builtinItems() *builtinItemCache {
	builtinItemsOnce.Do(func() {
		c := &builtinItemCache{keywordDoc: map[string]string{}}
		bs := embeddedBuiltins()
		if vba, ok := bs.Library("VBA").(*BuiltinProvider); ok {
			for _, m := range vba.Functions() {
				c.funcs = append(c.funcs, builtinMemberItem(m))
			}
			for _, m := range vba.Constants() {
				c.consts = append(c.consts, builtinMemberItem(m))
			}
			for _, ti := range vba.Types() {
				if ti.Kind != TypeKindEnum {
					continue
				}
				for _, m := range vba.Members(ti.Ref) {
					it := enumMemberItem(m, ti.Ref.Name)
					it.LabelDetails = describe("VBA." + ti.Ref.Name)
					c.consts = append(c.consts, it)
				}
			}
			for _, g := range vba.Globals() {
				c.globals = append(c.globals, CompletionItem{Label: g.Name, Kind: CompletionItemKindVariable,
					Detail: g.Name + " As " + g.Type.Name, Documentation: plainDoc(g.Doc), LabelDetails: describe("VBA")})
			}
		}
		for _, k := range bs.Keywords() {
			c.keywordDoc[strings.ToLower(k.Name)] = k.Doc
		}
		for _, k := range bs.Keywords() {
			if strings.EqualFold(k.Name, "Me") {
				continue // クラス・フォームのときだけ自モジュールのグループで出す
			}
			it := CompletionItem{Label: k.Name, Kind: CompletionItemKindKeyword, Documentation: plainDoc(k.Doc)}
			if k.Usage != "expression" {
				c.stmtKeywords = append(c.stmtKeywords, it)
				if !declKeywords[strings.ToLower(k.Name)] {
					c.inlineKeywords = append(c.inlineKeywords, it)
				}
			}
			if k.Usage != "statement" {
				c.exprKeywords = append(c.exprKeywords, it)
			}
		}
		builtinItemsVal = c
	})
	return builtinItemsVal
}

// keywordsFor は識別子補完の末尾に付けるキーワード（文頭なら文のキーワード、それ以外は式のキーワード）を返す。
func keywordsFor(ctx completionContext) []CompletionItem {
	b := builtinItems()
	switch {
	case ctx.StmtStart && ctx.AfterThen:
		return b.inlineKeywords
	case ctx.StmtStart:
		return b.stmtKeywords
	}
	return b.exprKeywords
}

// keywordItems は CtxKeyword の候補（決まったキーワード）を返す。
func keywordItems(names []string) []CompletionItem {
	set := newItemSet(len(names))
	for _, n := range names {
		set.add(grpKeyword, keywordItem(n))
	}
	return set.result()
}

// --- 型補完 ---

// buildTypeItems は As / New / Implements / TypeOf…Is の後の型の候補を返す。
// 並び: 自モジュールの UDT / Enum → ワークスペースの型 → 組み込み・TLB の型と primitive。
// New の後は New で作れる型（ワークスペースのクラス・フォーム、Collection / Dictionary / FileSystemObject）だけ。
func buildTypeItems(doc *Document, ctx completionContext) []CompletionItem {
	set := newItemSet(1500)
	if ctx.TypeLib != "" {
		for _, ti := range doc.Types.TypesIn(ctx.TypeLib) {
			if typeWanted(ti, ctx) {
				set.add(grpBuiltin, typeItem(ti))
			}
		}
		if !ctx.Creatable {
			for _, ti := range doc.Types.Primitives() {
				if strings.EqualFold(ti.Ref.Lib, ctx.TypeLib) {
					set.add(grpBuiltin, primitiveItem(ti))
				}
			}
		}
		return set.result()
	}
	if !ctx.Creatable {
		for _, u := range doc.Table.UDTs {
			set.add(grpModule, CompletionItem{Label: u.Name, Kind: CompletionItemKindStruct, Detail: "Type " + u.Name, Documentation: plainDoc(u.Doc)})
		}
		for _, e := range doc.Table.Enums {
			set.add(grpModule, CompletionItem{Label: e.Name, Kind: CompletionItemKindEnum, Detail: "Enum " + e.Name, Documentation: plainDoc(e.Doc)})
		}
	}
	for _, ti := range doc.Types.Types() {
		if !typeWanted(ti, ctx) {
			continue
		}
		g := grpBuiltin
		if ti.Origin != nil {
			g = grpProject
		}
		set.add(g, typeItem(ti))
	}
	if !ctx.Creatable {
		for _, ti := range doc.Types.Primitives() {
			set.add(grpBuiltin, primitiveItem(ti))
		}
	}
	return set.result()
}

func typeWanted(ti TypeInfo, ctx completionContext) bool {
	if ti.Kind == TypeKindModule || ti.Kind == TypeKindPrimitive || ti.Ref == formBaseType {
		return false
	}
	return !ctx.Creatable || ti.Creatable
}

// typeItem は型を候補にする。TLB の型は件数が多いので Detail・Documentation を付けない（所属だけ右側に出す）。
func typeItem(ti TypeInfo) CompletionItem {
	it := CompletionItem{Label: ti.Ref.Name, Kind: CompletionItemKindClass}
	switch {
	case ti.Kind == TypeKindUDT:
		it.Kind, it.Detail = CompletionItemKindStruct, "Type "+ti.Ref.Name
	case ti.Kind == TypeKindEnum:
		it.Kind, it.Detail = CompletionItemKindEnum, "Enum "+ti.Ref.Name
	case ti.Kind == TypeKindInterface:
		it.Kind = CompletionItemKindInterface
	case ti.Base == formBaseType:
		it.Detail = "UserForm " + ti.Ref.Name
	default:
		it.Detail = "Class " + ti.Ref.String()
	}
	if ti.Origin != nil {
		it.LabelDetails = describe(ti.Origin.Module)
	} else {
		it.LabelDetails = describe(ti.Ref.Lib)
	}
	if ti.Kind != TypeKindInterface {
		it.Documentation = plainDoc(ti.Doc)
	}
	return it
}

func primitiveItem(ti TypeInfo) CompletionItem {
	return CompletionItem{Label: ti.Ref.Name, Kind: CompletionItemKindKeyword, LabelDetails: describe(ti.Ref.Lib)}
}
