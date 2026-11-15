package main

import (
	"regexp"
	"strings"
	"unicode"
	"unicode/utf8"
)

// --- 論理行（行継続 " _" で連結した行） ---

// linePart は論理行の一部分と物理行の対応。
type linePart struct {
	logStart  int // 論理行内のバイト位置
	line      int // 物理行
	byteStart int // 物理行内のバイト位置
	length    int
}

// logicalLine は行継続で連結した 1 つの論理行。
type logicalLine struct {
	Text  string
	First int // 先頭の物理行
	Last  int // 末尾の物理行
	parts []linePart
}

func isBlankByte(b byte) bool { return b == ' ' || b == '\t' }

// trimRightBlank は末尾の空白（全角空白を含む）を除く。
func trimRightBlank(s string) string { return strings.TrimRight(s, " \t\r　") }

// hasContinuation は行が行継続（空白 + "_"）で終わるかを返す。
func hasContinuation(line string) bool {
	t := trimRightBlank(line)
	if !strings.HasSuffix(t, "_") {
		return false
	}
	t = t[:len(t)-1]
	if t == "" {
		return true
	}
	r, _ := utf8.DecodeLastRuneInString(t)
	return r == ' ' || r == '\t' || r == '　'
}

// logicalLineAt は line を含む論理行を返す（前の行にも後ろの行にも、継続している限り連結する）。
func logicalLineAt(lines []string, line int) logicalLine {
	first := line
	for first > 0 && hasContinuation(lines[first-1]) {
		first--
	}
	return buildLogicalLine(lines, first)
}

// buildLogicalLines は文書全体を論理行に分ける（診断用）。
func buildLogicalLines(lines []string) []logicalLine {
	var out []logicalLine
	for i := 0; i < len(lines); {
		ll := buildLogicalLine(lines, i)
		out = append(out, ll)
		i = ll.Last + 1
	}
	return out
}

// buildLogicalLine は first から始まる論理行を作る。
// 継続記号とその前の空白を除き、次の行は左の空白を除いて連結する。次の行が "." か "!" で始まる場合は区切りなし、
// それ以外は空白 1 つで連結する（"MsgBox .Name" を誤って a.Name と読まないため）。
func buildLogicalLine(lines []string, first int) logicalLine {
	ll := logicalLine{First: first, Last: first}
	var sb strings.Builder
	for i := first; i < len(lines); i++ {
		text := lines[i]
		start := 0
		if i > first {
			trimmed := strings.TrimLeft(text, " \t　")
			start = len(text) - len(trimmed)
			if trimmed != "" && trimmed[0] != '.' && trimmed[0] != '!' {
				sb.WriteByte(' ')
			}
		}
		cont := i+1 < len(lines) && hasContinuation(text)
		end := len(text)
		if cont {
			t := trimRightBlank(text)
			end = len(trimRightBlank(t[:len(t)-1]))
			if end < start {
				end = start
			}
		}
		ll.parts = append(ll.parts, linePart{logStart: sb.Len(), line: i, byteStart: start, length: end - start})
		sb.WriteString(text[start:end])
		ll.Last = i
		if !cont {
			break
		}
	}
	ll.Text = sb.String()
	return ll
}

// Offset は物理位置（行・バイト列）を論理行内のバイト位置に変換する。
func (ll logicalLine) Offset(line, byteCol int) int {
	for _, p := range ll.parts {
		if p.line == line {
			off := byteCol - p.byteStart
			if off < 0 {
				off = 0
			}
			if off > p.length {
				off = p.length
			}
			return p.logStart + off
		}
	}
	return len(ll.Text)
}

// Physical は論理行内のバイト位置を物理位置（行・バイト列）に変換する。
func (ll logicalLine) Physical(off int) (line, byteCol int) {
	for i, p := range ll.parts {
		last := i == len(ll.parts)-1
		if off < p.logStart {
			return p.line, p.byteStart // 連結用の空白は次の部分の先頭に寄せる
		}
		if off < p.logStart+p.length || (last && off <= p.logStart+p.length) {
			return p.line, p.byteStart + off - p.logStart
		}
	}
	if n := len(ll.parts); n > 0 {
		p := ll.parts[n-1]
		return p.line, p.byteStart + p.length
	}
	return ll.First, 0
}

// --- マスク ---

// maskCode は文字列リテラルの中身とコメント（' と、文頭 / ":" の直後の Rem）を空白にした文字列を返す（バイト長は保持）。
// 末尾が文字列の中か、コメントの中かも返す。
func maskCode(s string) (masked string, inString, inComment bool) {
	b := []byte(s)
	stmtStart := true // 文の先頭（行頭または ":" の後）で、まだ空白しか見ていない
	for i := 0; i < len(b); i++ {
		c := b[i]
		if inString {
			if c == '"' {
				inString = false
			} else {
				b[i] = ' '
			}
			continue
		}
		switch {
		case c == '"':
			inString = true
			stmtStart = false
		case c == '\'':
			for j := i; j < len(b); j++ {
				b[j] = ' '
			}
			return string(b), false, true
		case c == ':':
			stmtStart = true
		case isBlankByte(c):
		case stmtStart && (c == 'r' || c == 'R') && isRemAt(s, i):
			for j := i; j < len(b); j++ {
				b[j] = ' '
			}
			return string(b), false, true
		default:
			stmtStart = false
		}
	}
	return string(b), inString, false
}

// isRemAt は s[i:] が "Rem" の後に空白か行末が続く形かを返す。
func isRemAt(s string, i int) bool {
	if len(s)-i < 3 || !strings.EqualFold(s[i:i+3], "rem") {
		return false
	}
	return len(s) == i+3 || isBlankByte(s[i+3])
}

// --- 後ろ向きの式スキャナ ---

// ExprSegment はメンバーアクセスの 1 段。
type ExprSegment struct {
	Name  string
	Calls int  // 直後の (...) の数
	Bang  bool // a!b の b
}

// MemberChain は "a.b(1).c" のようなメンバーアクセスの連鎖。
type MemberChain struct {
	Segs       []ExprSegment
	LeadingDot bool // With ブロック内の ".a.b"
	Start      int  // 連鎖の先頭のバイト位置（LeadingDot なら "." の位置）
}

func runeBefore(s string, i int) rune {
	if i <= 0 {
		return utf8.RuneError
	}
	r, _ := utf8.DecodeLastRuneInString(s[:i])
	return r
}

// identStartBack は s[:end] の末尾にある識別子の開始位置を返す（識別子がなければ end）。
func identStartBack(s string, end int) int {
	i := end
	for i > 0 {
		r, size := utf8.DecodeLastRuneInString(s[:i])
		if !isIdentPart(r) {
			break
		}
		i -= size
	}
	return i
}

// matchOpenParenBack は s[closeIdx] の ")" に対応する "(" の位置を返す（なければ -1）。文字列はマスク済みであること。
func matchOpenParenBack(s string, closeIdx int) int {
	depth := 0
	for i := closeIdx; i >= 0; i-- {
		switch s[i] {
		case ')':
			depth++
		case '(':
			depth--
			if depth == 0 {
				return i
			}
		}
	}
	return -1
}

func skipBlankBack(s string, i int) int {
	for i > 0 {
		r, size := utf8.DecodeLastRuneInString(s[:i])
		if r != ' ' && r != '\t' && r != '　' {
			break
		}
		i -= size
	}
	return i
}

// endsChainOperand は s[:i] の末尾が識別子の文字か ")" か（"." の前が式の一部か）を返す。
func endsChainOperand(s string, i int) bool {
	if i <= 0 {
		return false
	}
	if s[i-1] == ')' {
		return true
	}
	return isIdentPart(runeBefore(s, i))
}

// scanChainBackward は masked[:end] の末尾にある "a.b(1).c" を読む（end は識別子の末尾、または ")" の直後）。
// masked は maskCode 済みであること。"(x).y" や数値リテラル "1." は読めない（ok=false）。
func scanChainBackward(s string, end int) (MemberChain, bool) {
	var segs []ExprSegment
	pos := end
	for {
		calls := 0
		for pos > 0 && s[pos-1] == ')' {
			open := matchOpenParenBack(s, pos-1)
			if open < 0 {
				return MemberChain{}, false
			}
			calls++
			pos = skipBlankBack(s, open)
		}
		if calls > 0 && pos > 0 && strings.IndexByte("$%&#@^", s[pos-1]) >= 0 {
			pos-- // Left$(...) などの型接尾辞
		}
		start := identStartBack(s, pos)
		if start == pos {
			return MemberChain{}, false
		}
		first, _ := utf8.DecodeRuneInString(s[start:pos])
		if !unicode.IsLetter(first) {
			return MemberChain{}, false
		}
		segs = append(segs, ExprSegment{Name: s[start:pos], Calls: calls})
		p := skipBlankBack(s, start) // "a. b" を許す
		if p == 0 || (s[p-1] != '.' && s[p-1] != '!') {
			reverseSegs(segs)
			return MemberChain{Segs: segs, Start: start}, true
		}
		sep := p - 1
		if s[sep] == '!' {
			segs[len(segs)-1].Bang = true
		}
		if !endsChainOperand(s, sep) {
			reverseSegs(segs)
			return MemberChain{Segs: segs, LeadingDot: true, Start: sep}, true
		}
		pos = sep
	}
}

func reverseSegs(s []ExprSegment) {
	for i, j := 0, len(s)-1; i < j; i, j = i+1, j-1 {
		s[i], s[j] = s[j], s[i]
	}
}

// --- 補完文脈 ---

type completionKind int

const (
	CtxNone       completionKind = iota
	CtxMember                    // "式." の後
	CtxIdentifier                // 識別子の途中（"." なし）
	CtxType                      // As / New / Implements / TypeOf…Is の後
	CtxKeyword                   // Exit / End / On Error / Option などの後（決まったキーワードだけ）
)

type completionContext struct {
	Kind      completionKind
	Chain     MemberChain // CtxMember のとき "." の左側（LeadingDot なら Segs は空）
	Partial   string      // 入力途中の識別子
	TypeLib   string      // CtxType: "As Scripting.|" の修飾子（"" なら全ライブラリ）
	Creatable bool        // CtxType: New の後（New で作れる型だけ）
	StmtStart bool        // CtxIdentifier: 文頭（文のキーワードも出す）
	AfterThen bool        // CtxIdentifier: 単一行 If の Then / Else の後の文頭（宣言のキーワードは出さない）
	AfterCall bool        // CtxIdentifier: Call の後（定数・Enum メンバー・キーワードを出さない）
	Keywords  []string    // CtxKeyword: 候補のキーワード
}

// analyzeCompletion はカーソル位置（行・UTF-16 位置）の補完文脈を求める。
func analyzeCompletion(doc *Document, line, utf16Char int) completionContext {
	if line < 0 || line >= len(doc.Lines) {
		return completionContext{}
	}
	ll := logicalLineAt(doc.Lines, line)
	off := ll.Offset(line, utf16ToByteOffset(doc.Lines[line], utf16Char))
	masked, inStr, inCmt := maskCode(ll.Text[:off])
	if inStr || inCmt {
		return completionContext{}
	}
	p := identStartBack(masked, off)
	partial := masked[p:off]
	if partial != "" {
		if r, _ := utf8.DecodeRuneInString(partial); !unicode.IsLetter(r) {
			return completionContext{} // 数値リテラル（"1.5" / ".5"）
		}
	}
	dot := skipBlankBack(masked, p)
	if dot > 0 && masked[dot-1] == '.' {
		dot--
		if endsChainOperand(masked, dot) {
			ch, ok := scanChainBackward(masked, dot)
			if !ok {
				return completionContext{}
			}
			// "As Scripting.|" / "New Scripting.|" は修飾付きの型
			if !ch.LeadingDot && len(ch.Segs) == 1 && ch.Segs[0].Calls == 0 && !ch.Segs[0].Bang {
				switch strings.ToLower(wordBefore(masked, ch.Start)) {
				case "as":
					return completionContext{Kind: CtxType, TypeLib: ch.Segs[0].Name, Partial: partial}
				case "new":
					return completionContext{Kind: CtxType, TypeLib: ch.Segs[0].Name, Creatable: true, Partial: partial}
				}
			}
			return completionContext{Kind: CtxMember, Chain: ch, Partial: partial}
		}
		return completionContext{Kind: CtxMember, Chain: MemberChain{LeadingDot: true, Start: dot}, Partial: partial}
	}
	ctx := statementContext(masked[:p])
	ctx.Partial = partial
	if ctx.Kind == CtxIdentifier && ctx.StmtStart && !ctx.AfterThen && ll.First == line &&
		(inTypeOrEnumBody(doc.Table, line) || inTypeOrEnumBodyText(doc.Lines, line)) {
		return completionContext{} // Type / Enum の本体ではメンバー名を書く
	}
	return ctx
}

// statementContext は "." の後でない位置の文脈を、カーソル（入力途中の識別子の先頭）より前のマスク済みテキストから求める。
func statementContext(before string) completionContext {
	s, afterThen := stmtStartIndex(before)
	stmt := strings.TrimLeft(before[s:], " \t　")
	words := identWords(stmt)
	prev := ""
	if n := len(words); n > 0 && len(before) > 0 && isBlankByte(before[len(before)-1]) {
		t := trimRightBlank(stmt)
		if strings.HasSuffix(strings.ToLower(t), words[n-1]) {
			prev = words[n-1]
		}
	}
	first := ""
	if len(words) > 0 {
		first = words[0]
	}
	kw := func(k ...string) completionContext { return completionContext{Kind: CtxKeyword, Keywords: k} }
	switch {
	case prev == "as" && first == "open":
		return completionContext{} // Open … As #n
	case prev == "as":
		return completionContext{Kind: CtxType}
	case prev == "new":
		return completionContext{Kind: CtxType, Creatable: true}
	case prev == "implements" && len(words) == 1:
		return completionContext{Kind: CtxType}
	case prev == "is" && containsWord(words, "typeof"):
		return completionContext{Kind: CtxType}
	case prev == "goto" || prev == "gosub":
		return completionContext{} // ラベル（将来）
	case prev == "resume":
		return kw("Next")
	case wordsAre(words, "exit"):
		return kw("Do", "For", "Function", "Property", "Sub")
	case wordsAre(words, "end"):
		return kw("Enum", "Function", "If", "Property", "Select", "Sub", "Type", "With")
	case wordsAre(words, "on"):
		return kw("Error")
	case wordsAre(words, "on", "error"):
		return kw("GoTo", "Resume")
	case wordsAre(words, "option"):
		return kw("Base", "Compare", "Explicit", "Private")
	case wordsAre(words, "option", "compare"):
		return kw("Binary", "Database", "Text")
	}
	switch declSlotOf(stmt, words) {
	case declName:
		return completionContext{}
	case declAfterName:
		return kw("As")
	case declModifiers:
		return kw("Const", "Declare", "Dim", "Enum", "Event", "Function", "Property", "Sub", "Type", "WithEvents")
	case declProperty:
		return kw("Get", "Let", "Set")
	}
	switch {
	case prev == "call" && len(words) == 1:
		return completionContext{Kind: CtxIdentifier, AfterCall: true}
	case len(words) == 0 && strings.TrimSpace(stmt) == "":
		return completionContext{Kind: CtxIdentifier, StmtStart: true, AfterThen: afterThen}
	}
	return completionContext{Kind: CtxIdentifier}
}

// stmtStartIndex は before の中で最後の文の開始位置を返す。括弧の深さ 0 にある ":"（":=" は名前付き引数なので除く）と、
// 語としての Then / Else の後のうち遅い方。afterThen は Then / Else の後で始まる文か。
func stmtStartIndex(before string) (start int, afterThen bool) {
	depth := 0
	for i := 0; i < len(before); {
		c := before[i]
		switch {
		case c == '(':
			depth++
		case c == ')':
			if depth > 0 {
				depth--
			}
		case c == ':' && depth == 0 && !(i+1 < len(before) && before[i+1] == '='):
			start, afterThen = i+1, false
		default:
			r, size := utf8.DecodeRuneInString(before[i:])
			if isIdentPart(r) {
				j := i + size
				for j < len(before) {
					r2, s2 := utf8.DecodeRuneInString(before[j:])
					if !isIdentPart(r2) {
						break
					}
					j += s2
				}
				w := before[i:j]
				if depth == 0 && (strings.EqualFold(w, "then") || strings.EqualFold(w, "else")) && !precededByDot(before, i) {
					start, afterThen = j, true
				}
				i = j
				continue
			}
			i += size
			continue
		}
		i++
	}
	return start, afterThen
}

func precededByDot(s string, i int) bool {
	j := skipBlankBack(s, i)
	return j > 0 && s[j-1] == '.'
}

// identWords は s（マスク済み）の識別子の語を小文字で返す。数字で始まる語（数値リテラル）は含めない。
func identWords(s string) []string {
	var out []string
	for i := 0; i < len(s); {
		r, size := utf8.DecodeRuneInString(s[i:])
		if !isIdentPart(r) {
			i += size
			continue
		}
		j := i
		for j < len(s) {
			r2, s2 := utf8.DecodeRuneInString(s[j:])
			if !isIdentPart(r2) {
				break
			}
			j += s2
		}
		if unicode.IsLetter(r) {
			out = append(out, strings.ToLower(s[i:j]))
		}
		i = j
	}
	return out
}

// wordBefore は s[:end] の末尾（空白を飛ばす）にある識別子を返す（なければ ""）。
func wordBefore(s string, end int) string {
	e := skipBlankBack(s, end)
	b := identStartBack(s, e)
	if b > 0 && s[b-1] == '.' {
		return ""
	}
	return s[b:e]
}

func wordsAre(words []string, want ...string) bool {
	if len(words) != len(want) {
		return false
	}
	for i := range words {
		if words[i] != want[i] {
			return false
		}
	}
	return true
}

func containsWord(words []string, w string) bool {
	for _, x := range words {
		if x == w {
			return true
		}
	}
	return false
}

// declSlot は宣言文の中での位置。
type declSlot int

const (
	declNone      declSlot = iota // 宣言文ではない、または式・型の位置
	declName                      // 宣言する名前を書く位置（Dim |、Sub |、引数名）
	declAfterName                 // 宣言リストで名前の後（Dim x |）。次は As
	declModifiers                 // 修飾語だけ（Public |）。次は Sub / Const など
	declProperty                  // Property の後（Get / Let / Set）
)

var declModifierWords = map[string]bool{"public": true, "private": true, "friend": true, "global": true, "static": true}

// declSlotOf は文 stmt（文頭からカーソルまで）が宣言文のどの位置にあるかを返す。
func declSlotOf(stmt string, words []string) declSlot {
	i := 0
	for i < len(words) && declModifierWords[words[i]] {
		i++
	}
	mods := i
	if i == len(words) {
		if mods > 0 {
			return declModifiers
		}
		return declNone
	}
	rest := afterWords(stmt, i) // words[i] 以降のテキスト
	switch words[i] {
	case "sub", "function", "event":
		return headerPosition(afterWords(stmt, i+1))
	case "property":
		if i+1 == len(words) {
			return declProperty
		}
		return headerPosition(afterWords(stmt, i+2))
	case "declare":
		j := i + 1
		if j < len(words) && words[j] == "ptrsafe" {
			j++
		}
		if j == len(words) {
			return declNone
		}
		return headerPosition(afterWords(stmt, j+1))
	case "type", "enum":
		if i+1 == len(words) {
			return declName
		}
		return declNone
	case "dim", "const":
		return declListPosition(afterWords(stmt, i+1))
	case "withevents":
		return declListPosition(rest)
	}
	if mods > 0 {
		return declListPosition(rest) // Public x / Private WithEvents x
	}
	return declNone
}

// afterWords は stmt の先頭から n 語を読み飛ばした後のテキストを返す。
func afterWords(stmt string, n int) string {
	i := 0
	for k := 0; k < n; k++ {
		for i < len(stmt) {
			r, size := utf8.DecodeRuneInString(stmt[i:])
			if isIdentPart(r) {
				break
			}
			i += size
		}
		for i < len(stmt) {
			r, size := utf8.DecodeRuneInString(stmt[i:])
			if !isIdentPart(r) {
				break
			}
			i += size
		}
	}
	return stmt[i:]
}

// declListPosition は宣言リスト（"a As Long, b"）のうち最後の項目を見る。
func declListPosition(list string) declSlot {
	depth, last := 0, -1
	for i := 0; i < len(list); i++ {
		switch list[i] {
		case '(':
			depth++
		case ')':
			if depth > 0 {
				depth--
			}
		case ',':
			if depth == 0 {
				last = i
			}
		}
	}
	if depth > 0 {
		return declNone // Dim a(| は添字の式
	}
	item := list[last+1:]
	words := identWords(item)
	if len(words) > 0 && words[0] == "withevents" {
		words = words[1:]
	}
	if strings.Contains(item, "=") || containsWord(words, "as") {
		return declNone
	}
	switch {
	case len(words) == 0:
		return declName
	case len(words) == 1 && strings.TrimRight(item, " \t　") != item:
		return declAfterName
	}
	return declNone
}

// headerPosition はプロシージャの宣言（名前の後のテキスト）のどこにいるかを返す。
func headerPosition(rest string) declSlot {
	open := strings.IndexByte(rest, '(')
	if open < 0 {
		if len(identWords(rest)) <= 1 {
			return declName // プロシージャ名
		}
		return declNone // Declare の Lib "…" など
	}
	depth, last := 0, open
	for i := open; i < len(rest); i++ {
		switch rest[i] {
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				return declNone // ")" の後（As は別の規則）
			}
		case ',':
			if depth == 1 {
				last = i
			}
		}
	}
	if depth != 1 {
		return declNone
	}
	item := rest[last+1:]
	if strings.Contains(item, "=") {
		return declNone // 既定値は式
	}
	words := identWords(item)
	k := 0
	for k < len(words) && (words[k] == "optional" || words[k] == "byval" || words[k] == "byref" || words[k] == "paramarray") {
		k++
	}
	if containsWord(words[k:], "as") || len(words[k:]) > 1 {
		return declNone
	}
	return declName
}

// inTypeOrEnumBody は line が Type / Enum の本体（宣言行と End 行の間）にあるかを返す。
func inTypeOrEnumBody(st *SymbolTable, line int) bool {
	if st == nil {
		return false
	}
	for _, u := range st.UDTs {
		if u.StartLine < line && line < u.EndLine {
			return true
		}
	}
	for _, e := range st.Enums {
		if e.StartLine < line && line < e.EndLine {
			return true
		}
	}
	return false
}

var (
	typeHeaderRe = regexp.MustCompile(`(?i)^` + wsClass + `*(?:(?:Public|Private)` + wsClass + `+)?(?:Type|Enum)` + wsClass + `+` + identPattern + wsClass + `*$`)
	// typeMemberRe は Type / Enum の本体の行（"a As Long" / "a(3) As T" / "a = 1" / "a"）
	typeMemberRe = regexp.MustCompile(`^` + wsClass + `*(` + identPattern + `)` + `(?:` + wsClass + `*\(.*\))?(?:` + wsClass + `+As` + wsClass + `.*|` + wsClass + `*=.*)?` + wsClass + `*$`)
)

// typeMemberKeywords は Type / Enum のメンバー名にならない語（本体の外の文の先頭）。
var typeMemberKeywords = map[string]bool{"end": true, "dim": true, "const": true, "sub": true, "function": true,
	"property": true, "public": true, "private": true, "friend": true, "global": true, "static": true, "declare": true,
	"option": true, "attribute": true, "implements": true, "event": true, "type": true, "enum": true}

// inTypeOrEnumBodyText は、解析結果に Type / Enum が無いとき（本体が空で構文エラーになる入力途中など）のための
// テキストによる判定。line の上を、メンバーらしい行・空行・コメント行の間だけさかのぼり、Type / Enum の宣言行に着けば true。
func inTypeOrEnumBodyText(lines []string, line int) bool {
	for i := line - 1; i >= 0 && i >= line-200; i-- {
		l := lines[i]
		if masked, _, _ := maskCode(l); strings.TrimSpace(masked) == "" {
			continue // 空行・コメント行
		} else {
			l = masked
		}
		if typeHeaderRe.MatchString(l) {
			return true
		}
		m := typeMemberRe.FindStringSubmatch(l)
		if m == nil || typeMemberKeywords[strings.ToLower(m[1])] {
			return false
		}
	}
	return false
}
