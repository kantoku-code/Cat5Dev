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
		sym := doc.Table.Lookup(last.Name, line)
		if sym == nil {
			return ""
		}
		return symbolHover(sym)
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
	return codeBlock(m.Signature(), m.Doc)
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
