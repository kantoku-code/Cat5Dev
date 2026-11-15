package main

import (
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
	CtxIdentifier                // 2-2: 識別子の途中（"." なし）
	CtxType                      // 2-2: As / New の後
)

type completionContext struct {
	Kind    completionKind
	Chain   MemberChain // CtxMember のとき "." の左側（LeadingDot なら Segs は空）
	Partial string      // 入力途中の識別子
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
	if dot == 0 || masked[dot-1] != '.' {
		return completionContext{Kind: CtxIdentifier, Partial: partial}
	}
	dot--
	if endsChainOperand(masked, dot) {
		ch, ok := scanChainBackward(masked, dot)
		if !ok {
			return completionContext{}
		}
		return completionContext{Kind: CtxMember, Chain: ch, Partial: partial}
	}
	return completionContext{Kind: CtxMember, Chain: MemberChain{LeadingDot: true, Start: dot}, Partial: partial}
}
