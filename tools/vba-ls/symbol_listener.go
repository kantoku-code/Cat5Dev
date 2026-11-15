package main

import (
	"regexp"
	"strings"

	"github.com/antlr4-go/antlr/v4"
	vbaantlr "github.com/nori/cat5dev/vba-ls/antlr"
)

// reProgIDArg は CreateObject の第1引数（文字列リテラル）から ProgID を取り出す。
var reProgIDArg = regexp.MustCompile(`^"([^"]+)"`)

// symbolListener は ANTLR4 パースツリーを走査してシンボルを収集するリスナー。
type symbolListener struct {
	*vbaantlr.BasevbaListener
	table *SymbolTable
	lines []string // splitLines の結果（列の UTF-16 変換・ドキュメント抽出用）

	curProc    *ProcInfo // 本体付きプロシージャ（スコープ）
	paramOwner *ProcInfo // 引数を集めるプロシージャ（Declare / Event も含む）

	varVis    Visibility // 処理中の Dim / Public / Private / Static 文の公開範囲
	varStatic bool

	grammarErrors int // 文法では表していない制約の違反数（構文エラーとして数える）
}

func newSymbolListener(lines []string) *symbolListener {
	return &symbolListener{table: &SymbolTable{}, lines: lines, varVis: VisPrivate}
}

func (l *symbolListener) scope() string { return l.curProc.ScopeKey() }

// tokenPos はトークンの 0-based 行と UTF-16 列を返す。
// ANTLR の GetColumn は rune 単位のため、該当行のテキストを使って UTF-16 単位に変換する。
func (l *symbolListener) tokenPos(tok antlr.Token) (line, col int) {
	line = tok.GetLine() - 1
	col = tok.GetColumn()
	if line >= 0 && line < len(l.lines) {
		col = runeToUTF16Offset(l.lines[line], col)
	}
	return line, col
}

func (l *symbolListener) addSymbol(s Symbol) *Symbol {
	l.table.Symbols = append(l.table.Symbols, s)
	return &l.table.Symbols[len(l.table.Symbols)-1]
}

// setTypePos は型名トークンの位置を記録する。
func (l *symbolListener) setTypePos(s *Symbol, tok antlr.Token) {
	if tok == nil {
		s.TypeNameLine, s.TypeNameCol, s.HasTypeNamePos = 0, 0, false
		return
	}
	s.TypeNameLine, s.TypeNameCol = l.tokenPos(tok)
	s.HasTypeNamePos = true
}

// finish は構文エラー時のスコープ補正と索引の構築を行う。
func (l *symbolListener) finish() {
	st := l.table
	if st.Stats.SyntaxErrors > 0 {
		l.applyRegexScopes()
	}
	st.finalize()
}

// applyRegexScopes は構文エラーで木から求めたスコープが崩れている可能性があるため、
// 正規表現で求めたプロシージャ範囲を使い、変数・引数・Const の Scope と Owner を位置から割り当て直す
// （カーソル側の ScopeAt と矛盾させないため）。
func (l *symbolListener) applyRegexScopes() {
	st := l.table
	byKey := map[string]*ProcInfo{}
	for _, p := range st.Procs {
		if p.HasBody {
			byKey[strings.ToLower(p.ScopeKey())] = p
		}
	}
	var scopes []*ProcInfo
	for _, r := range regexScopes(l.lines) {
		if p, ok := byKey[strings.ToLower(r.ScopeKey())]; ok {
			p.StartLine, p.EndLine = r.StartLine, r.EndLine
			r = p
		}
		scopes = append(scopes, r)
	}
	st.scopes = scopes
	st.index = nil
	for i := range st.Symbols {
		s := &st.Symbols[i]
		if s.Kind != SymVariable && s.Kind != SymParam && s.Kind != SymConst {
			continue
		}
		p := st.ProcAt(s.Line)
		if p == nil && s.Owner == nil {
			continue
		}
		s.Owner = p
		s.Scope = p.ScopeKey()
		if p == nil && s.Visibility == VisLocal {
			s.Visibility = VisPrivate
		}
	}
	st.index = nil
}

// --- 型ヘルパー ---

// typeRefOfDefined は "Name" / "Lib.Name" の型参照と、最後の名前のトークンを返す。
func typeRefOfDefined(d vbaantlr.IDefinedTypeExpressionContext) (TypeRef, antlr.Token) {
	if d == nil {
		return TypeRef{}, nil
	}
	var parts []string
	var tok antlr.Token
	if sn := d.SimpleNameExpression(); sn != nil {
		parts = append(parts, sn.GetText())
		tok = sn.GetStart()
	}
	for _, un := range d.AllUnrestrictedName() {
		parts = append(parts, un.GetText())
		tok = un.GetStart()
	}
	return ParseTypeName(strings.Join(parts, ".")), tok
}

// typeRefOfTypeExpr は組み込み型（角括弧を除く）または定義型の型参照を返す。
func typeRefOfTypeExpr(te vbaantlr.ITypeExpressionContext) (TypeRef, antlr.Token) {
	if te == nil {
		return TypeRef{}, nil
	}
	if bt := te.BuiltinType(); bt != nil {
		return ParseTypeName(bt.GetText()), bt.GetStart()
	}
	return typeRefOfDefined(te.DefinedTypeExpression())
}

// typeRefOfAsClause は As 句の型参照・型名トークン・As New かどうかを返す。固定長文字列は String とする。
func typeRefOfAsClause(ac vbaantlr.IAsClauseContext) (TypeRef, antlr.Token, bool) {
	if ac == nil {
		return TypeRef{}, nil, false
	}
	if ao := ac.AsAutoObject(); ao != nil {
		if ct := ao.ClassTypeName(); ct != nil {
			ref, tok := typeRefOfDefined(ct.DefinedTypeExpression())
			return ref, tok, true
		}
		return TypeRef{}, nil, true
	}
	if at := ac.AsType(); at != nil {
		if ts := at.TypeSpec(); ts != nil {
			if fl := ts.FixedLengthStringSpec(); fl != nil {
				return TypeRef{Name: "String"}, fl.GetStart(), false
			}
			ref, tok := typeRefOfTypeExpr(ts.TypeExpression())
			return ref, tok, false
		}
	}
	return TypeRef{}, nil, false
}

// --- プロシージャ ---

func procVisibility(sc vbaantlr.IProcedureScopeContext) Visibility {
	switch {
	case sc == nil:
		return VisPublic
	case sc.PRIVATE() != nil:
		return VisPrivate
	case sc.FRIEND() != nil:
		return VisFriend
	case sc.GLOBAL() != nil:
		return VisGlobal
	}
	return VisPublic
}

// funcNameParts は関数名の先頭トークン・名前（接尾辞なし）・型接尾辞を返す。
func funcNameParts(fn vbaantlr.IFunctionNameContext) (antlr.Token, string, string) {
	if fn == nil {
		return nil, "", ""
	}
	if tn := fn.TypedName(); tn != nil && tn.AmbiguousIdentifier() != nil {
		suffix := ""
		if ts := tn.TypeSuffix(); ts != nil {
			suffix = ts.GetText()
		}
		return tn.AmbiguousIdentifier().GetStart(), tn.AmbiguousIdentifier().GetText(), suffix
	}
	return fn.GetStart(), fn.GetText(), ""
}

func subNameParts(sn vbaantlr.ISubroutineNameContext) (antlr.Token, string) {
	if sn == nil {
		return nil, ""
	}
	return sn.GetStart(), sn.GetText()
}

type procSpec struct {
	kind       ProcKind
	vis        Visibility
	static     bool
	nameTok    antlr.Token
	name       string
	suffix     string
	fnType     vbaantlr.IFunctionTypeContext
	start      antlr.Token // 宣言の先頭（修飾子を含む）
	end        antlr.TerminalNode
	stop       antlr.Token
	hasBody    bool
	symbolKind SymbolKind
}

// beginProc は ProcInfo とモジュールレベルのシンボルを作る。
func (l *symbolListener) beginProc(sp procSpec) *ProcInfo {
	if sp.nameTok == nil || sp.name == "" {
		return nil
	}
	p := &ProcInfo{Name: sp.name, Kind: sp.kind, Visibility: sp.vis, Static: sp.static, HasBody: sp.hasBody}
	p.NameLine, p.NameCol = l.tokenPos(sp.nameTok)
	p.StartLine = p.NameLine
	if sp.start != nil {
		p.StartLine = sp.start.GetLine() - 1
	}
	p.EndLine = p.StartLine
	if sp.end != nil {
		p.EndLine = sp.end.GetSymbol().GetLine() - 1
	} else if sp.stop != nil {
		p.EndLine = sp.stop.GetLine() - 1
	}
	var typeTok antlr.Token
	if p.HasReturn() {
		switch {
		case sp.fnType != nil && sp.fnType.TypeExpression() != nil:
			p.ReturnType, typeTok = typeRefOfTypeExpr(sp.fnType.TypeExpression())
			p.ReturnType.Array = sp.fnType.ArrayDesignator() != nil
		case suffixTypes[sp.suffix] != "":
			p.ReturnType = TypeRef{Name: suffixTypes[sp.suffix]}
		default:
			p.ReturnType = TypeRef{Name: "Variant"}
		}
	}
	p.Doc = docAbove(l.lines, p.StartLine)
	l.table.Procs = append(l.table.Procs, p)

	sym := Symbol{
		Name: p.Name, Kind: sp.symbolKind, Visibility: p.Visibility, Static: p.Static, Scope: "module",
		Decl: p, Doc: p.Doc, Line: p.NameLine, Col: p.NameCol,
		TypeName: p.ReturnType.Name, TypeLib: p.ReturnType.Lib, IsArray: p.ReturnType.Array,
	}
	l.setTypePos(&sym, typeTok)
	l.addSymbol(sym)
	return p
}

func (l *symbolListener) enterBody(p *ProcInfo) {
	if p != nil {
		l.curProc, l.paramOwner = p, p
	}
}

func (l *symbolListener) exitBody() { l.curProc, l.paramOwner = nil, nil }

func (l *symbolListener) EnterSubroutineDeclaration(ctx *vbaantlr.SubroutineDeclarationContext) {
	tok, name := subNameParts(ctx.SubroutineName())
	l.enterBody(l.beginProc(procSpec{
		kind: ProcSub, vis: procVisibility(ctx.ProcedureScope()), static: ctx.InitialStatic() != nil || ctx.TrailingStatic() != nil,
		nameTok: tok, name: name, start: ctx.GetStart(), end: ctx.END(), stop: ctx.GetStop(), hasBody: true, symbolKind: SymProcedure,
	}))
}

func (l *symbolListener) ExitSubroutineDeclaration(*vbaantlr.SubroutineDeclarationContext) {
	l.exitBody()
}

func (l *symbolListener) EnterFunctionDeclaration(ctx *vbaantlr.FunctionDeclarationContext) {
	tok, name, suffix := funcNameParts(ctx.FunctionName())
	l.enterBody(l.beginProc(procSpec{
		kind: ProcFunction, vis: procVisibility(ctx.ProcedureScope()), static: ctx.InitialStatic() != nil || ctx.TrailingStatic() != nil,
		nameTok: tok, name: name, suffix: suffix, fnType: ctx.FunctionType(),
		start: ctx.GetStart(), end: ctx.END(), stop: ctx.GetStop(), hasBody: true, symbolKind: SymProcedure,
	}))
}

func (l *symbolListener) ExitFunctionDeclaration(*vbaantlr.FunctionDeclarationContext) { l.exitBody() }

func (l *symbolListener) EnterPropertyGetDeclaration(ctx *vbaantlr.PropertyGetDeclarationContext) {
	tok, name, suffix := funcNameParts(ctx.FunctionName())
	l.enterBody(l.beginProc(procSpec{
		kind: ProcPropertyGet, vis: procVisibility(ctx.ProcedureScope()), static: ctx.InitialStatic() != nil || ctx.TrailingStatic() != nil,
		nameTok: tok, name: name, suffix: suffix, fnType: ctx.FunctionType(),
		start: ctx.GetStart(), end: ctx.END(), stop: ctx.GetStop(), hasBody: true, symbolKind: SymProperty,
	}))
}

func (l *symbolListener) ExitPropertyGetDeclaration(*vbaantlr.PropertyGetDeclarationContext) {
	l.exitBody()
}

func (l *symbolListener) EnterPropertySetDeclaration(ctx *vbaantlr.PropertySetDeclarationContext) {
	kind := ProcPropertyLet
	if ctx.SET() != nil {
		kind = ProcPropertySet
	}
	tok, name := subNameParts(ctx.SubroutineName())
	l.enterBody(l.beginProc(procSpec{
		kind: kind, vis: procVisibility(ctx.ProcedureScope()), static: ctx.InitialStatic() != nil || ctx.TrailingStatic() != nil,
		nameTok: tok, name: name, start: ctx.GetStart(), end: ctx.END(), stop: ctx.GetStop(), hasBody: true, symbolKind: SymProperty,
	}))
}

func (l *symbolListener) ExitPropertySetDeclaration(*vbaantlr.PropertySetDeclarationContext) {
	l.exitBody()
}

// externalStart は Declare 文の先頭トークン（Public / Private を含む）と公開範囲を返す。
func externalStart(ctx antlr.ParserRuleContext) (antlr.Token, Visibility) {
	start, vis := ctx.GetStart(), VisPublic
	if dcl, ok := ctx.GetParent().(antlr.ParserRuleContext); ok {
		start = dcl.GetStart()
		switch outer := dcl.GetParent().(type) {
		case *vbaantlr.PrivateExternalProcedureDeclarationContext:
			start, vis = outer.GetStart(), VisPrivate
		case *vbaantlr.PublicExternalProcedureDeclarationContext:
			start = outer.GetStart()
		}
	}
	return start, vis
}

func (l *symbolListener) EnterExternalSub(ctx *vbaantlr.ExternalSubContext) {
	tok, name := subNameParts(ctx.SubroutineName())
	start, vis := externalStart(ctx)
	l.paramOwner = l.beginProc(procSpec{
		kind: ProcDeclareSub, vis: vis, nameTok: tok, name: name, start: start, stop: ctx.GetStop(), symbolKind: SymProcedure,
	})
}

func (l *symbolListener) ExitExternalSub(*vbaantlr.ExternalSubContext) { l.paramOwner = nil }

func (l *symbolListener) EnterExternalFunction(ctx *vbaantlr.ExternalFunctionContext) {
	tok, name, suffix := funcNameParts(ctx.FunctionName())
	start, vis := externalStart(ctx)
	l.paramOwner = l.beginProc(procSpec{
		kind: ProcDeclareFunction, vis: vis, nameTok: tok, name: name, suffix: suffix, fnType: ctx.FunctionType(),
		start: start, stop: ctx.GetStop(), symbolKind: SymProcedure,
	})
}

func (l *symbolListener) ExitExternalFunction(*vbaantlr.ExternalFunctionContext) { l.paramOwner = nil }

func (l *symbolListener) EnterEventDeclaration(ctx *vbaantlr.EventDeclarationContext) {
	id := ctx.AmbiguousIdentifier()
	if id == nil {
		return
	}
	l.paramOwner = l.beginProc(procSpec{
		kind: ProcEvent, vis: VisPublic, nameTok: id.GetStart(), name: id.GetText(),
		start: ctx.GetStart(), stop: ctx.GetStop(), symbolKind: SymEvent,
	})
}

func (l *symbolListener) ExitEventDeclaration(*vbaantlr.EventDeclarationContext) { l.paramOwner = nil }

// --- 引数 ---

func (l *symbolListener) EnterPositionalParam(ctx *vbaantlr.PositionalParamContext) {
	if l.paramOwner == nil {
		return
	}
	pm := ctx.ParameterMechanism()
	l.addParam(ctx.ParamDcl(), pm != nil && pm.BYVAL() != nil, false, "")
}

func (l *symbolListener) EnterOptionalParam(ctx *vbaantlr.OptionalParamContext) {
	if l.paramOwner == nil {
		return
	}
	byVal := false
	if op := ctx.OptionalPrefix(); op != nil {
		if pm := op.ParameterMechanism(); pm != nil {
			byVal = pm.BYVAL() != nil
		}
	}
	def := ""
	if dv := ctx.DefaultValue(); dv != nil && dv.ConstantExpression() != nil {
		def = normalizeExpr(dv.ConstantExpression().GetText())
	}
	l.addParam(ctx.ParamDcl(), byVal, true, def)
}

func (l *symbolListener) EnterParamArray(ctx *vbaantlr.ParamArrayContext) {
	if l.paramOwner == nil || ctx.AmbiguousIdentifier() == nil {
		return
	}
	id := ctx.AmbiguousIdentifier()
	pi := ParamInfo{Name: id.GetText(), Type: TypeRef{Name: "Variant", Array: true}, ParamArray: true}
	pi.Line, pi.Col = l.tokenPos(id.GetStart())
	l.appendParam(pi, nil)
}

func (l *symbolListener) addParam(pd vbaantlr.IParamDclContext, byVal, optional bool, def string) {
	if pd == nil {
		return
	}
	pi := ParamInfo{ByVal: byVal, Optional: optional, Default: def}
	var nameTok, typeTok antlr.Token
	if u := pd.UntypedNameParamDcl(); u != nil && u.AmbiguousIdentifier() != nil {
		nameTok, pi.Name = u.AmbiguousIdentifier().GetStart(), u.AmbiguousIdentifier().GetText()
		if pt := u.ParameterType(); pt != nil {
			if pt.TypeExpression() != nil {
				pi.Type, typeTok = typeRefOfTypeExpr(pt.TypeExpression())
			} else if pt.ANY() != nil {
				pi.Type = TypeRef{Name: "Any"}
			}
			pi.Type.Array = pt.ArrayDesignator() != nil
		}
	} else if t := pd.TypedNameParamDcl(); t != nil && t.TypedName() != nil && t.TypedName().AmbiguousIdentifier() != nil {
		tn := t.TypedName()
		nameTok, pi.Name = tn.AmbiguousIdentifier().GetStart(), tn.AmbiguousIdentifier().GetText()
		if ts := tn.TypeSuffix(); ts != nil {
			pi.Type = TypeRef{Name: suffixTypes[ts.GetText()]}
		}
		pi.Type.Array = t.ArrayDesignator() != nil
	}
	if nameTok == nil {
		return
	}
	pi.Line, pi.Col = l.tokenPos(nameTok)
	l.appendParam(pi, typeTok)
}

// appendParam は引数を ProcInfo に追加し、本体のあるプロシージャならそのスコープの SymParam も追加する。
func (l *symbolListener) appendParam(pi ParamInfo, typeTok antlr.Token) {
	p := l.paramOwner
	p.Params = append(p.Params, pi)
	if !p.HasBody {
		return
	}
	sym := Symbol{
		Name: pi.Name, Kind: SymParam, TypeName: pi.Type.Name, TypeLib: pi.Type.Lib, IsArray: pi.Type.Array,
		Visibility: VisLocal, Scope: p.ScopeKey(), Owner: p, Value: pi.Default, Line: pi.Line, Col: pi.Col,
	}
	l.setTypePos(&sym, typeTok)
	l.addSymbol(sym)
}

// --- 変数・定数 ---

// declVisibility は Dim / Public / Private / Global / Static / Const の公開範囲を返す。
func (l *symbolListener) declVisibility(mod vbaantlr.IVariableModifierContext) (Visibility, bool) {
	static := mod != nil && mod.STATIC() != nil
	if l.curProc != nil {
		return VisLocal, static
	}
	switch {
	case mod == nil:
		return VisPrivate, false
	case mod.GLOBAL() != nil:
		return VisGlobal, false
	case mod.PUBLIC() != nil:
		return VisPublic, false
	}
	return VisPrivate, static
}

func (l *symbolListener) EnterVariableDeclaration(ctx *vbaantlr.VariableDeclarationContext) {
	l.varVis, l.varStatic = l.declVisibility(ctx.VariableModifier())
}

func (l *symbolListener) ExitVariableDeclaration(*vbaantlr.VariableDeclarationContext) {
	l.varVis, l.varStatic = VisPrivate, false
}

func (l *symbolListener) newVariable(nameTok antlr.Token, name string) Symbol {
	s := Symbol{Name: name, Kind: SymVariable, Visibility: l.varVis, Static: l.varStatic, Scope: l.scope(), Owner: l.curProc}
	s.Line, s.Col = l.tokenPos(nameTok)
	s.Doc = declDoc(l.lines, s.Line)
	return s
}

func (l *symbolListener) EnterVariableDcl(ctx *vbaantlr.VariableDclContext) {
	if t := ctx.TypedVariableDcl(); t != nil {
		// Dim s$ / Dim a%(10)
		tn := t.TypedName()
		if tn == nil || tn.AmbiguousIdentifier() == nil {
			return
		}
		sym := l.newVariable(tn.AmbiguousIdentifier().GetStart(), tn.AmbiguousIdentifier().GetText())
		if ts := tn.TypeSuffix(); ts != nil {
			sym.TypeName = suffixTypes[ts.GetText()]
		}
		sym.IsArray = t.ArrayDim() != nil
		l.addSymbol(sym)
		return
	}
	u := ctx.UntypedVariableDcl()
	if u == nil || u.AmbiguousIdentifier() == nil {
		return
	}
	id := u.AmbiguousIdentifier()
	sym := l.newVariable(id.GetStart(), id.GetText())
	ac := u.AsClause()
	if arr := u.ArrayClause(); arr != nil {
		// Dim a(3) As Part の As 句は arrayClause の下にある
		sym.IsArray = true
		ac = arr.AsClause()
	}
	ref, typeTok, asNew := typeRefOfAsClause(ac)
	sym.TypeName, sym.TypeLib, sym.AsNew = ref.Name, ref.Lib, asNew
	l.setTypePos(&sym, typeTok)
	l.addSymbol(sym)
}

func (l *symbolListener) EnterWitheventsVariableDcl(ctx *vbaantlr.WitheventsVariableDclContext) {
	id := ctx.AmbiguousIdentifier()
	if id == nil {
		return
	}
	sym := l.newVariable(id.GetStart(), id.GetText())
	sym.WithEvents = true
	if ct := ctx.ClassTypeName(); ct != nil {
		ref, tok := typeRefOfDefined(ct.DefinedTypeExpression())
		sym.TypeName, sym.TypeLib = ref.Name, ref.Lib
		l.setTypePos(&sym, tok)
	}
	l.addSymbol(sym)
}

func (l *symbolListener) EnterConstDeclaration(ctx *vbaantlr.ConstDeclarationContext) {
	list := ctx.ConstItemList()
	if list == nil {
		return
	}
	vis, _ := l.declVisibility(ctx.VariableModifier())
	for _, it := range list.AllConstItem() {
		id := it.AmbiguousIdentifier()
		if id == nil {
			continue
		}
		sym := Symbol{Name: id.GetText(), Kind: SymConst, Visibility: vis, Scope: l.scope(), Owner: l.curProc}
		sym.Line, sym.Col = l.tokenPos(id.GetStart())
		if ca := it.ConstAsClause(); ca != nil && ca.BuiltinType() != nil {
			sym.TypeName = ParseTypeName(ca.BuiltinType().GetText()).Name
			l.setTypePos(&sym, ca.BuiltinType().GetStart())
		} else if ts := it.TypeSuffix(); ts != nil {
			sym.TypeName = suffixTypes[ts.GetText()]
		}
		if ce := it.ConstantExpression(); ce != nil {
			sym.Value = normalizeExpr(ce.GetText())
		}
		sym.Doc = declDoc(l.lines, sym.Line)
		l.addSymbol(sym)
	}
}

// --- ユーザー定義型・列挙型 ---

func cleanName(s string) string { return strings.Trim(strings.TrimSpace(s), "[]") }

func (l *symbolListener) EnterUdtDeclaration(ctx *vbaantlr.UdtDeclarationContext) {
	un := ctx.UntypedName()
	if un == nil {
		return
	}
	u := &UDTInfo{Name: cleanName(un.GetText()), Visibility: VisPublic}
	start := ctx.GetStart()
	switch parent := ctx.GetParent().(type) {
	case *vbaantlr.PublicTypeDeclarationContext:
		start = parent.GetStart()
		if parent.GLOBAL() != nil {
			u.Visibility = VisGlobal
		}
	case *vbaantlr.PrivateTypeDeclarationContext:
		start, u.Visibility = parent.GetStart(), VisPrivate
	}
	u.StartLine = start.GetLine() - 1
	u.EndLine = u.StartLine
	if ctx.END() != nil {
		u.EndLine = ctx.END().GetSymbol().GetLine() - 1
	}
	u.Doc = docAbove(l.lines, u.StartLine)
	l.table.UDTs = append(l.table.UDTs, u)

	sym := Symbol{Name: u.Name, Kind: SymUDT, Visibility: u.Visibility, Scope: "module", Doc: u.Doc}
	sym.Line, sym.Col = l.tokenPos(un.GetStart())
	l.addSymbol(sym)

	ml := ctx.UdtMemberList()
	if ml == nil {
		return
	}
	for _, el := range ml.AllUdtElement() {
		m := el.UdtMember()
		if m == nil {
			continue
		}
		var nameTok antlr.Token
		var name string
		var ac vbaantlr.IAsClauseContext
		isArray := false
		if d := m.UntypedNameMemberDcl(); d != nil && d.AmbiguousIdentifier() != nil {
			nameTok, name = d.AmbiguousIdentifier().GetStart(), d.AmbiguousIdentifier().GetText()
			if oac := d.OptionalArrayClause(); oac != nil {
				isArray = oac.ArrayDim() != nil
				ac = oac.AsClause()
			}
		} else if d := m.ReservedNameMemberDcl(); d != nil && d.ReservedMemberName() != nil {
			nameTok, name = d.ReservedMemberName().GetStart(), d.ReservedMemberName().GetText()
			ac = d.AsClause()
		}
		if nameTok == nil {
			continue
		}
		ms := Symbol{Name: name, Kind: SymUDTMember, Visibility: VisPublic, Scope: "UDT:" + u.Name, Parent: u.Name, IsArray: isArray}
		ms.Line, ms.Col = l.tokenPos(nameTok)
		ref, typeTok, asNew := typeRefOfAsClause(ac)
		ms.TypeName, ms.TypeLib, ms.AsNew = ref.Name, ref.Lib, asNew
		l.setTypePos(&ms, typeTok)
		ms.Doc = declDoc(l.lines, ms.Line)
		u.Members = append(u.Members, &ms)
		l.addSymbol(ms)
	}
}

func (l *symbolListener) EnterEnumDeclaration(ctx *vbaantlr.EnumDeclarationContext) {
	un := ctx.UntypedName()
	if un == nil {
		return
	}
	e := &EnumInfo{Name: cleanName(un.GetText()), Visibility: VisPublic}
	start := ctx.GetStart()
	switch parent := ctx.GetParent().(type) {
	case *vbaantlr.GlobalEnumDeclarationContext:
		start, e.Visibility = parent.GetStart(), VisGlobal
	case *vbaantlr.PublicEnumDeclarationContext:
		start = parent.GetStart()
	case *vbaantlr.PrivateEnumDeclarationContext:
		start, e.Visibility = parent.GetStart(), VisPrivate
	}
	e.StartLine = start.GetLine() - 1
	e.EndLine = e.StartLine
	if ctx.END() != nil {
		e.EndLine = ctx.END().GetSymbol().GetLine() - 1
	}
	e.Doc = docAbove(l.lines, e.StartLine)
	l.table.Enums = append(l.table.Enums, e)

	sym := Symbol{Name: e.Name, Kind: SymEnum, Visibility: e.Visibility, Scope: "module", Doc: e.Doc}
	sym.Line, sym.Col = l.tokenPos(un.GetStart())
	l.addSymbol(sym)

	ml := ctx.EnumMemberList()
	if ml == nil {
		return
	}
	for _, el := range ml.AllEnumElement() {
		m := el.EnumMember()
		if m == nil || m.UntypedName() == nil {
			continue
		}
		ms := Symbol{Name: cleanName(m.UntypedName().GetText()), Kind: SymEnumMember, Visibility: e.Visibility, Scope: "module", Parent: e.Name, TypeName: e.Name}
		ms.Line, ms.Col = l.tokenPos(m.UntypedName().GetStart())
		if ce := m.ConstantExpression(); ce != nil {
			ms.Value = normalizeExpr(ce.GetText())
		}
		ms.Doc = declDoc(l.lines, ms.Line)
		e.Members = append(e.Members, &ms)
		l.addSymbol(ms)
	}
}

// --- For の末尾 ---

func (l *symbolListener) EnterForStatement(ctx *vbaantlr.ForStatementContext) {
	l.checkForTail(ctx.StatementBlock(), ctx.ForTail())
}

func (l *symbolListener) EnterForEachStatement(ctx *vbaantlr.ForEachStatementContext) {
	l.checkForTail(ctx.StatementBlock(), ctx.ForTail())
}

// checkForTail は ", 変数" で閉じる For（"Next j, i" の外側）が、元の文法と同じく
// 「本体の最後の文が Next に変数を持つ For / For Each」の場合だけであることを検査する。
func (l *symbolListener) checkForTail(body vbaantlr.IStatementBlockContext, tail vbaantlr.IForTailContext) {
	if tail == nil || tail.NEXT() != nil {
		return
	}
	if !endsWithExplicitFor(body) {
		l.grammarErrors++
	}
}

// endsWithExplicitFor は文のブロックの最後の文が、Next に変数を持つ For / For Each かを返す。
func endsWithExplicitFor(body vbaantlr.IStatementBlockContext) bool {
	if body == nil {
		return false
	}
	stmts := body.AllBlockStatement()
	if len(stmts) == 0 || stmts[len(stmts)-1].Statement() == nil {
		return false
	}
	cs := stmts[len(stmts)-1].Statement().ControlStatement()
	if cs == nil || cs.ControlStatementExceptMultilineIf() == nil {
		return false
	}
	c := cs.ControlStatementExceptMultilineIf()
	var tail vbaantlr.IForTailContext
	if f := c.ForStatement(); f != nil {
		tail = f.ForTail()
	} else if f := c.ForEachStatement(); f != nil {
		tail = f.ForTail()
	}
	return tail != nil && tail.BoundVariableExpression() != nil
}

// --- With / Implements ---

func (l *symbolListener) EnterWithStatement(ctx *vbaantlr.WithStatementContext) {
	expr := ctx.Expression()
	if expr == nil {
		return
	}
	w := WithBlock{StartLine: ctx.GetStart().GetLine() - 1, Expr: normalizeExpr(expr.GetText())}
	w.EndLine = w.StartLine
	if ctx.END() != nil {
		w.EndLine = ctx.END().GetSymbol().GetLine() - 1
	} else if ctx.GetStop() != nil {
		w.EndLine = ctx.GetStop().GetLine() - 1
	}
	w.ExprLine, w.ExprCol = l.tokenPos(expr.GetStart())
	l.table.Withs = append(l.table.Withs, w)
}

func (l *symbolListener) EnterImplementsDirective(ctx *vbaantlr.ImplementsDirectiveContext) {
	if ct := ctx.ClassTypeName(); ct != nil {
		ref, _ := typeRefOfDefined(ct.DefinedTypeExpression())
		if ref.Name != "" {
			l.table.Implements = append(l.table.Implements, ref.String())
		}
	}
}

// --- Set 代入（Set x = New TypeName） ---

// EnterSetStatement は "Set <変数> = New <型>" / "Set <変数> = CreateObject("ProgID")" から
// 変数の型を登録する。左辺が単純な変数名でない（obj.Prop / arr(i) / With 内の .Member）場合は対象外。
func (l *symbolListener) EnterSetStatement(ctx *vbaantlr.SetStatementContext) {
	lex, rhs := ctx.LExpression(), ctx.Expression()
	if lex == nil || rhs == nil {
		return
	}
	nameTok, name := simpleLExprName(lex)
	if nameTok == nil || name == "" {
		return
	}
	if ne := rhs.NewExpress(); ne != nil {
		// New TypeName / New Lib.TypeName → 最後の名前を型名とし、その位置を型名トークンにする
		if inner := ne.Expression(); inner != nil && inner.LExpression() != nil {
			il := inner.LExpression()
			if typeTok, typeName := lastNameOf(il); typeTok != nil && typeName != "" {
				raw := typeName
				if il.PERIOD() != nil && il.LExpression() != nil {
					raw = il.LExpression().GetText() + "." + typeName
				}
				l.upsertSymbol(name, nameTok, ParseTypeName(raw), typeTok)
			}
		}
		return
	}
	if progID := createObjectProgID(rhs); progID != "" {
		// CreateObject("ProgID") → ProgID から型を作る（ソース上に型名トークンはない）
		if ref := TypeRefFromProgID(progID); ref.Name != "" {
			l.upsertSymbol(name, nameTok, ref, nil)
		}
	}
}

// identOfName は name 規則から識別子部分（型接尾辞を除く）を返す。
func identOfName(n vbaantlr.INameContext) vbaantlr.IAmbiguousIdentifierContext {
	if n == nil {
		return nil
	}
	if un := n.UntypedName(); un != nil {
		return un.AmbiguousIdentifier()
	}
	if tn := n.TypedName(); tn != nil {
		return tn.AmbiguousIdentifier()
	}
	return nil
}

// simpleLExprName は lExpression が単純な名前のとき、その識別子の先頭トークンと名前を返す。
func simpleLExprName(lex vbaantlr.ILExpressionContext) (antlr.Token, string) {
	sn := lex.SimpleNameExpression()
	if sn == nil {
		return nil, ""
	}
	id := identOfName(sn.Name())
	if id == nil {
		return nil, ""
	}
	return id.GetStart(), id.GetText()
}

// lastNameOf は "Foo" または "Lib.Foo" の形の lExpression から最後の名前とその先頭トークンを返す。
func lastNameOf(lex vbaantlr.ILExpressionContext) (antlr.Token, string) {
	if tok, name := simpleLExprName(lex); tok != nil {
		return tok, name
	}
	if lex.PERIOD() != nil {
		if un := lex.UnrestrictedName(); un != nil {
			return un.GetStart(), un.GetText()
		}
	}
	return nil, ""
}

// createObjectProgID は式が CreateObject("ProgID", ...) の呼び出しなら ProgID を返す。
func createObjectProgID(expr vbaantlr.IExpressionContext) string {
	lex := expr.LExpression()
	if lex == nil || lex.LPAREN() == nil || lex.ArgumentList() == nil {
		return ""
	}
	callee := lex.LExpression()
	if callee == nil {
		return ""
	}
	calleeName := ""
	if callee.SimpleNameExpression() != nil {
		calleeName = callee.GetText()
	} else if un := callee.UnrestrictedName(); un != nil {
		calleeName = un.GetText() // VBA.CreateObject(...)
	}
	if !strings.EqualFold(calleeName, "CreateObject") {
		return ""
	}
	if m := reProgIDArg.FindStringSubmatch(lex.ArgumentList().GetText()); m != nil {
		return m[1]
	}
	return ""
}

// upsertSymbol は Set 代入で型が判明した変数を登録する。
// nameTok は左辺の変数名トークン、typeTok は型名トークン（ソース上にない場合は nil）。
// 同じスコープに既存の変数・引数があれば、型が未設定の場合のみ型を補う（宣言位置は変えない）。
func (l *symbolListener) upsertSymbol(name string, nameTok antlr.Token, ref TypeRef, typeTok antlr.Token) {
	scope := l.scope()
	for i := range l.table.Symbols {
		s := &l.table.Symbols[i]
		if s.Scope != scope || !strings.EqualFold(s.Name, name) {
			continue
		}
		if (s.Kind == SymVariable || s.Kind == SymParam) && s.TypeName == "" {
			s.TypeName, s.TypeLib = ref.Name, ref.Lib
			l.setTypePos(s, typeTok)
		}
		return
	}
	vis := VisPrivate
	if l.curProc != nil {
		vis = VisLocal
	}
	sym := Symbol{Name: name, Kind: SymVariable, TypeName: ref.Name, TypeLib: ref.Lib, Visibility: vis, Scope: scope, Owner: l.curProc}
	sym.Line, sym.Col = l.tokenPos(nameTok)
	l.setTypePos(&sym, typeTok)
	l.addSymbol(sym)
}

// --- 式・ドキュメント ---

var spacesRe = regexp.MustCompile(`[ \t\x{3000}]+`)

// normalizeExpr は式のテキストの行継続と連続する空白を 1 つの空白にし、"." の前後の空白を除く。
func normalizeExpr(s string) string {
	s = lineContinuationRe.ReplaceAllString(s, " ")
	s = spacesRe.ReplaceAllString(s, " ")
	s = strings.ReplaceAll(s, " .", ".")
	s = strings.ReplaceAll(s, ". ", ".")
	return strings.TrimSpace(s)
}

// docTrim はドキュメントの各行の両端から除く文字（装飾と空白）。
const docTrim = "*=-/_#~ \t　'"

// commentText はコメント行（' または Rem）ならその本文を返す。
func commentText(line string) (string, bool) {
	t := strings.TrimLeft(line, " \t　")
	switch {
	case strings.HasPrefix(t, "'"):
		return strings.TrimLeft(t, "'"), true
	case len(t) >= 3 && strings.EqualFold(t[:3], "rem") && (len(t) == 3 || t[3] == ' ' || t[3] == '\t'):
		return t[3:], true
	}
	return "", false
}

// docAbove は startLine の直上に連続するコメント行を集める（空行・コードで打ち切り）。
// 装飾（*=-/_#~）と空白を両端から除き、空になった行は捨てる。
func docAbove(lines []string, startLine int) string {
	var parts []string
	for i := startLine - 1; i >= 0 && i < len(lines); i-- {
		text, ok := commentText(lines[i])
		if !ok {
			break
		}
		if t := strings.Trim(text, docTrim); t != "" {
			parts = append(parts, t)
		}
	}
	for i, j := 0, len(parts)-1; i < j; i, j = i+1, j-1 {
		parts[i], parts[j] = parts[j], parts[i]
	}
	return strings.Join(parts, "\n")
}

// trailingComment は行末コメント（コードの後ろの '）の本文を返す。
func trailingComment(line string) string {
	inStr := false
	for i := 0; i < len(line); i++ {
		switch c := line[i]; {
		case c == '"':
			inStr = !inStr
		case c == '\'' && !inStr:
			if strings.TrimSpace(line[:i]) == "" {
				return "" // 行全体がコメント
			}
			return strings.Trim(line[i+1:], docTrim)
		}
	}
	return ""
}

// declDoc は宣言のドキュメント（行末コメントがあればそれ、なければ直上のコメント）を返す。
func declDoc(lines []string, line int) string {
	if line < 0 || line >= len(lines) {
		return ""
	}
	if c := trailingComment(lines[line]); c != "" {
		return c
	}
	return docAbove(lines, line)
}
