package main

import (
	"strings"
	"unicode"
	"unicode/utf8"
)

// identBounds は行 l のバイト位置 b を含む（または直前で終わる）識別子の範囲 [s, e) を返す。
func identBounds(l string, b int) (s, e int) {
	if b > len(l) {
		b = len(l)
	}
	s = identStartBack(l, b)
	e = b
	for e < len(l) {
		r, size := utf8.DecodeRuneInString(l[e:])
		if !isIdentPart(r) {
			break
		}
		e += size
	}
	return s, e
}

// ResolveHover はカーソル位置のホバー Markdown を返す。対象がなければ ""。
// 単独の識別子はシンボルの宣言を、"式.名前" はメンバーのシグネチャを表示する。
func ResolveHover(doc *Document, line, utf16Char int) string {
	if line < 0 || line >= len(doc.Lines) {
		return ""
	}
	l := doc.Lines[line]
	s, e := identBounds(l, utf16ToByteOffset(l, utf16Char))
	if s == e {
		return ""
	}
	if r, _ := utf8.DecodeRuneInString(l[s:e]); !unicode.IsLetter(r) {
		return ""
	}
	ll := logicalLineAt(doc.Lines, line)
	offE := ll.Offset(line, e)
	masked, inStr, inCmt := maskCode(ll.Text[:offE])
	if inStr || inCmt {
		return ""
	}
	ch, ok := scanChainBackward(masked, offE)
	if !ok || len(ch.Segs) == 0 {
		return ""
	}
	last := ch.Segs[len(ch.Segs)-1]
	if !ch.LeadingDot && len(ch.Segs) == 1 {
		if sym := doc.Table.Lookup(last.Name, line); sym != nil {
			return symbolHover(sym)
		}
		return projectHover(doc, last)
	}
	owner := MemberChain{Segs: ch.Segs[:len(ch.Segs)-1], LeadingDot: ch.LeadingDot}
	t, ok := newExprResolver(doc).resolveChain(owner, line)
	if !ok {
		return ""
	}
	m, ok := doc.Types.Member(t, last.Name)
	if !ok {
		return ""
	}
	return memberHover(m)
}

// projectHover は自モジュールに無い単独の識別子（標準モジュールの Public・モジュール名・ワークスペースの型）のホバーを返す。
func projectHover(doc *Document, sg ExprSegment) string {
	if _, m, ok := doc.Types.ProjectMember(sg.Name); ok {
		return memberHover(m)
	}
	if ti, ok := projectType(doc, sg.Name); ok {
		return typeHover(ti)
	}
	return ""
}

// projectType はモジュール名（Main / frmMenu）またはワークスペースの型名（clsA / Public UDT・Enum）を型情報にする。
// TLB・組み込みの型（宣言元の無いもの）は対象外。
func projectType(doc *Document, name string) (TypeInfo, bool) {
	if ref, ok := doc.Types.ModuleRef(name); ok {
		if ti, _, ok := doc.Types.Resolve(ref); ok {
			return ti, true
		}
	}
	if ti, _, ok := doc.Types.Resolve(TypeRef{Name: name}); ok && ti.Origin != nil {
		return ti, true
	}
	return TypeInfo{}, false
}

// memberHover はメンバーのホバー。ワークスペースのメンバーなら宣言（Property は Get / Let / Set の各行）を表示する。
func memberHover(m MemberInfo) string {
	if m.Origin != nil && len(m.Origin.Syms) > 0 {
		decls := make([]string, len(m.Origin.Syms))
		for i, s := range m.Origin.Syms {
			decls[i] = declarationText(s)
		}
		return codeBlock(strings.Join(decls, "\n"), m.Doc)
	}
	return codeBlock(m.Signature(), m.Doc)
}

// typeHover はワークスペースの型（クラス・フォーム・標準モジュール・UDT・Enum）のホバー。
func typeHover(ti TypeInfo) string {
	if ti.Origin != nil && len(ti.Origin.Syms) > 0 {
		return codeBlock(declarationText(ti.Origin.Syms[0]), ti.Doc)
	}
	kw := "Class"
	switch {
	case ti.Kind == TypeKindModule:
		kw = "Module"
	case ti.Base == formBaseType:
		kw = "UserForm"
	}
	return codeBlock(kw+" "+ti.Ref.Name, ti.Doc)
}

func codeBlock(code, doc string) string {
	s := "```vba\n" + code + "\n```"
	if doc != "" {
		s += "\n\n" + doc
	}
	return s
}

// symbolHover はシンボルの宣言表示とドキュメントを返す。
func symbolHover(sym *Symbol) string {
	doc := sym.Doc
	if sym.Decl != nil && doc == "" {
		doc = sym.Decl.Doc
	}
	return codeBlock(declarationText(sym), doc)
}

func visibilityKeyword(v Visibility) string {
	switch v {
	case VisPublic:
		return "Public"
	case VisFriend:
		return "Friend"
	case VisGlobal:
		return "Global"
	case VisPrivate:
		return "Private"
	}
	return ""
}

// asClause は " As 型" を返す（型がなければ ""）。
func asClause(t TypeRef, asNew bool) string {
	if t.IsZero() {
		return ""
	}
	s := " As "
	if asNew {
		s += "New "
	}
	return s + t.Elem().String()
}

// declarationText はシンボルの宣言を VBA の構文で表す。
func declarationText(sym *Symbol) string {
	arr := ""
	if sym.IsArray {
		arr = "()"
	}
	elem := TypeRef{Lib: sym.TypeLib, Name: sym.TypeName}
	switch sym.Kind {
	case SymVariable:
		kw := "Dim"
		if sym.Owner == nil {
			kw = visibilityKeyword(sym.Visibility)
		} else if sym.Static {
			kw = "Static"
		}
		if sym.WithEvents {
			kw += " WithEvents"
		}
		return kw + " " + sym.Name + arr + asClause(elem, sym.AsNew)
	case SymParam:
		if sym.Owner != nil {
			for _, p := range sym.Owner.Params {
				if strings.EqualFold(p.Name, sym.Name) {
					return "(引数) " + paramText(p)
				}
			}
		}
		return "(引数) " + sym.Name + arr + asClause(elem, false)
	case SymConst:
		kw := "Const"
		if sym.Owner == nil && sym.Visibility != VisPrivate {
			kw = visibilityKeyword(sym.Visibility) + " Const"
		}
		s := kw + " " + sym.Name + asClause(elem, false)
		if sym.Value != "" {
			s += " = " + sym.Value
		}
		return s
	case SymProcedure, SymProperty, SymEvent:
		if sym.Decl != nil {
			return procSignature(sym.Decl)
		}
	case SymUDT:
		return strings.TrimSpace(visibilityKeyword(sym.Visibility) + " Type " + sym.Name)
	case SymEnum:
		return strings.TrimSpace(visibilityKeyword(sym.Visibility) + " Enum " + sym.Name)
	case SymEnumMember:
		s := sym.Parent + "." + sym.Name
		if sym.Value != "" {
			s += " = " + sym.Value
		}
		return s
	case SymUDTMember:
		return sym.Parent + "." + sym.Name + arr + asClause(elem, sym.AsNew)
	}
	return sym.Name
}

// paramText は引数 1 つを VBA の構文で表す。
func paramText(p ParamInfo) string {
	s := ""
	switch {
	case p.ParamArray:
		s = "ParamArray "
	case p.Optional:
		s = "Optional "
	}
	if p.ByVal {
		s += "ByVal "
	}
	s += p.Name
	if p.Type.Array {
		s += "()"
	}
	s += asClause(p.Type, false)
	if p.Default != "" {
		s += " = " + p.Default
	}
	return s
}

var procKindKeyword = map[ProcKind]string{
	ProcSub: "Sub", ProcFunction: "Function", ProcPropertyGet: "Property Get", ProcPropertyLet: "Property Let",
	ProcPropertySet: "Property Set", ProcDeclareSub: "Declare Sub", ProcDeclareFunction: "Declare Function", ProcEvent: "Event",
}

// procSignature はプロシージャのシグネチャを VBA の構文で表す。
func procSignature(p *ProcInfo) string {
	var sb strings.Builder
	if v := visibilityKeyword(p.Visibility); v != "" {
		sb.WriteString(v + " ")
	}
	if p.Static {
		sb.WriteString("Static ")
	}
	sb.WriteString(procKindKeyword[p.Kind] + " " + p.Name + "(")
	for i, pi := range p.Params {
		if i > 0 {
			sb.WriteString(", ")
		}
		sb.WriteString(paramText(pi))
	}
	sb.WriteString(")")
	if p.HasReturn() && !p.ReturnType.IsZero() {
		sb.WriteString(" As " + p.ReturnType.String())
	}
	return sb.String()
}
