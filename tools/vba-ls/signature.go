package main

import (
	"regexp"
	"strings"
)

// callSite は呼び出し 1 つと、カーソルがいる引数。
type callSite struct {
	Callee   MemberChain // 最後の段が呼び出す名前
	Suffix   string      // 型接尾辞（Left$( の "$"）
	Paren    bool        // 括弧付き（false は文頭の括弧なし呼び出し）
	ArgIndex int         // カーソルより前の深さ 0 のカンマの数
	NamedArg string      // 現在の引数が "名前:=" ならその名前
}

var (
	namedArgRe = regexp.MustCompile(`^[\s\p{Zs}]*(` + identPattern + `)[\s\p{Zs}]*:=`)
	// stmtCallRe は文頭の括弧なし呼び出しの呼び出し先（"MsgBox " / "Err.Raise " / ".Add "）。
	stmtCallRe = regexp.MustCompile(`^(\.?` + identPattern + `(?:[\s\p{Zs}]*\.[\s\p{Zs}]*` + identPattern + `)*)[\s\p{Zs}]+`)
)

// stmtKeywords は文頭の括弧なし呼び出しと見なさない語（文のキーワード）。
var stmtKeywords = map[string]bool{
	"if": true, "elseif": true, "else": true, "while": true, "until": true, "do": true, "loop": true, "for": true,
	"next": true, "select": true, "case": true, "dim": true, "redim": true, "set": true, "let": true, "call": true,
	"return": true, "print": true, "open": true, "close": true, "goto": true, "gosub": true, "resume": true,
	"exit": true, "end": true, "with": true, "erase": true, "raiseevent": true, "static": true, "const": true,
	"private": true, "public": true, "friend": true, "global": true, "sub": true, "function": true, "property": true,
	"declare": true, "type": true, "enum": true, "event": true, "implements": true, "option": true, "on": true,
	"wend": true, "then": true, "stop": true, "get": true, "put": true, "input": true, "line": true, "write": true,
	"seek": true, "lock": true, "unlock": true, "name": true, "width": true, "attribute": true,
}

// findCallSites はカーソル位置を囲む呼び出しの候補を内側から外側へ返す（最後に文頭の括弧なし呼び出し）。
// masked は maskCode 済みの論理行の [0:off]（文字列の中でもよい。コメントの中なら呼ばない）。
func findCallSites(masked string) []callSite {
	s, _ := stmtStartIndex(masked)
	var out []callSite
	depth, commas := 0, 0
	argEnd := len(masked) // 現在の階層の「現在の引数」の終わり
	argStart := -1        // 現在の階層で最も近い深さ 0 のカンマの位置
	for i := len(masked) - 1; i >= s; i-- {
		switch masked[i] {
		case ')':
			depth++
		case '(':
			if depth > 0 {
				depth--
				continue
			}
			if cs, ok := parenCallee(masked, i); ok {
				from := i + 1
				if argStart >= 0 {
					from = argStart + 1
				}
				cs.ArgIndex = commas
				cs.NamedArg = namedArgOf(masked[from:argEnd])
				out = append(out, cs)
			}
			// 外側の階層へ（グループ化の括弧の中のカンマは外側の引数番号に数えない）
			commas, argStart = 0, -1
		case ',':
			if depth == 0 {
				commas++
				if argStart < 0 {
					argStart = i
				}
			}
		}
	}
	if depth == 0 {
		if cs, ok := stmtCallee(masked, s); ok {
			out = append(out, cs)
		}
	}
	return out
}

// parenCallee は masked[open] の "(" の前が呼び出せる名前（識別子・")"・型接尾辞）なら呼び出し先を返す。
func parenCallee(masked string, open int) (callSite, bool) {
	j := skipBlankBack(masked, open) // "MsgBox (x" の空白区切りも許す
	suffix := ""
	if j >= 2 && strings.IndexByte("$%&#@^", masked[j-1]) >= 0 && isIdentPart(runeBefore(masked, j-1)) {
		suffix, j = masked[j-1:j], j-1
	}
	if !endsChainOperand(masked, j) {
		return callSite{}, false
	}
	ch, ok := scanChainBackward(masked, j)
	if !ok || len(ch.Segs) == 0 {
		return callSite{}, false
	}
	return callSite{Callee: ch, Suffix: suffix, Paren: true}, true
}

// stmtCallee は文頭（masked[s:]）の括弧なし呼び出し（"MsgBox "a", |"）を返す。
func stmtCallee(masked string, s int) (callSite, bool) {
	stmt := masked[s:]
	lead := len(stmt) - len(strings.TrimLeft(stmt, " \t　"))
	m := stmtCallRe.FindStringSubmatchIndex(stmt[lead:])
	if m == nil {
		return callSite{}, false
	}
	chainEnd := s + lead + m[3]
	argsFrom := s + lead + m[1]
	words := identWords(masked[s+lead : chainEnd])
	if len(words) == 0 || stmtKeywords[words[0]] {
		return callSite{}, false
	}
	rest := strings.TrimLeft(masked[argsFrom:], " \t　")
	if strings.HasPrefix(rest, "=") {
		return callSite{}, false // 代入
	}
	ch, ok := scanChainBackward(masked, chainEnd)
	if !ok || len(ch.Segs) == 0 {
		return callSite{}, false
	}
	cs := callSite{Callee: ch}
	depth, last := 0, argsFrom-1
	for i := argsFrom; i < len(masked); i++ {
		switch masked[i] {
		case '(':
			depth++
		case ')':
			depth--
		case ',':
			if depth == 0 {
				cs.ArgIndex++
				last = i
			}
		}
	}
	cs.NamedArg = namedArgOf(masked[last+1:])
	return cs, true
}

func namedArgOf(arg string) string {
	if m := namedArgRe.FindStringSubmatch(arg); m != nil {
		return m[1]
	}
	return ""
}

// BuildSignatureHelp はカーソル位置で入力中の呼び出しのシグネチャを返す。呼び出しが見つからなければ nil。
func BuildSignatureHelp(doc *Document, line, utf16Char int) *SignatureHelp {
	if line < 0 || line >= len(doc.Lines) {
		return nil
	}
	ll := logicalLineAt(doc.Lines, line)
	off := ll.Offset(line, utf16ToByteOffset(doc.Lines[line], utf16Char))
	masked, _, inCmt := maskCode(ll.Text[:off])
	if inCmt {
		return nil
	}
	r := newExprResolver(doc)
	for _, cs := range findCallSites(masked) {
		m, ok := r.resolveCallee(cs, line)
		if !ok {
			continue // 配列の添字・未知の名前は外側へ
		}
		return &SignatureHelp{
			Signatures:      []SignatureInformation{signatureInfo(m)},
			ActiveParameter: activeParameter(m.Params, cs),
		}
	}
	return nil
}

// resolveCallee は呼び出し先のメンバー情報（引数・戻り値・説明）を返す。
func (r *exprResolver) resolveCallee(cs callSite, line int) (MemberInfo, bool) {
	ch := cs.Callee
	last := ch.Segs[len(ch.Segs)-1]
	if last.Calls > 0 || last.Bang {
		return MemberInfo{}, false
	}
	if !ch.LeadingDot && len(ch.Segs) == 1 {
		if sym := r.table.Lookup(last.Name, line); sym != nil {
			if (sym.Kind == SymProcedure || sym.Kind == SymProperty) && sym.Decl != nil {
				return procMember(sym), true
			}
			return MemberInfo{}, false // 変数・定数（配列の添字・既定メンバー）
		}
		if _, m, ok := r.types.ProjectMember(last.Name); ok {
			return callableMember(m)
		}
		if m, ok := r.types.Function(last.Name + cs.Suffix); ok {
			return m, true
		}
		return MemberInfo{}, false
	}
	owner := MemberChain{Segs: ch.Segs[:len(ch.Segs)-1], LeadingDot: ch.LeadingDot, Start: ch.Start}
	t, ok := r.resolveChain(owner, line)
	if !ok {
		return MemberInfo{}, false
	}
	m, ok := r.types.Member(t, last.Name)
	if !ok {
		return MemberInfo{}, false
	}
	return callableMember(m)
}

// callableMember は呼び出しとして表示できるメンバーか（引数のないプロパティ・UDT のメンバー・Enum メンバーは除く）。
func callableMember(m MemberInfo) (MemberInfo, bool) {
	switch m.Kind {
	case MemberKindMethod, MemberKindEvent:
		return m, true
	case MemberKindProperty:
		return m, len(m.Params) > 0
	}
	return MemberInfo{}, false
}

// procMember は自モジュールのプロシージャを MemberInfo にする（Property は宣言の引数をそのまま）。
func procMember(sym *Symbol) MemberInfo {
	d := sym.Decl
	m := MemberInfo{Name: d.Name, Kind: MemberKindMethod, Params: toTLBParams(d.Params), Doc: d.Doc}
	if d.HasReturn() && !d.ReturnType.IsZero() {
		m.ReturnType = d.ReturnType.String()
	}
	return m
}

// signatureInfo はメンバーをシグネチャ表示にする（ラベルと、各引数のラベル内の UTF-16 の範囲）。
func signatureInfo(m MemberInfo) SignatureInformation {
	var sb strings.Builder
	sb.WriteString(m.Name + "(")
	params := make([]ParameterInformation, 0, len(m.Params))
	for i, p := range m.Params {
		if i > 0 {
			sb.WriteString(", ")
		}
		start := utf16Len(sb.String())
		sb.WriteString(formatParam(p))
		params = append(params, ParameterInformation{Label: [2]uint32{uint32(start), uint32(utf16Len(sb.String()))}})
	}
	sb.WriteString(")")
	if ret := m.ReturnType; ret != "" && !strings.EqualFold(ret, "void") {
		sb.WriteString(" As " + ret)
	}
	return SignatureInformation{Label: sb.String(), Documentation: plainDoc(m.Doc), Parameters: params}
}

// activeParameter は強調する引数を決める。名前付き引数は名前で、ParamArray は最後の引数に寄せる。
func activeParameter(params []TLBParam, cs callSite) uint32 {
	if cs.NamedArg != "" {
		for i, p := range params {
			if strings.EqualFold(p.Name, cs.NamedArg) {
				return uint32(i)
			}
		}
	}
	if n := len(params); n > 0 && params[n-1].ParamArray && cs.ArgIndex >= n-1 {
		return uint32(n - 1)
	}
	return uint32(cs.ArgIndex)
}
