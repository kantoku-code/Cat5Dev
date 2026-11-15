package main

import (
	"regexp"
	"strings"
	"unicode"
	"unicode/utf8"
)

// --- 入力途中の行の補修 ---
//
// 補完や signatureHelp を使うのは、行を書いている途中（"d."、"MsgBox "a", "、"x = Foo(1, " など）であることが多い。
// そのまま解析すると構文エラーの回復でカーソルより後ろのプロシージャ・モジュール変数がシンボル表から落ちるため、
// 解析の前に「論理行の末尾」だけを補修する。補修は位置を変えない:
//   - 行内の置換は ASCII の 1 バイトを空白 1 バイトにするだけ（バイト長・UTF-16 の位置を保つ）
//   - 追加は物理行の末尾への追記だけ（閉じていない文字列の `"`、閉じていない括弧の `)`）
// 追記が必要な行に行末コメントがあるときは、コメントを空白にしてから追記する（コメントの後ろに字句は無い）。
// 正しいコードは対象にならない（サンプルで補修 0 行を検査する）。
// ブロックの未完（Do / For / Select の未終端）は補修しない（サーバーの直前の正常な表で補完候補を補う）。

// trailingOpWords は行末にあると式が未完の語（小文字）。
var trailingOpWords = map[string]bool{
	"and": true, "or": true, "xor": true, "not": true, "mod": true, "like": true, "is": true,
	"to": true, "step": true, "new": true, "eqv": true, "imp": true,
}

// printStmtRe は Print 文（Debug.Print / Print #n）。行末の "," は正しい形（次のタブ位置）なので補修しない。
var printStmtRe = regexp.MustCompile(`(?i)(?:^|[\s:.])Print(?:[\s#]|$)`)

// repairSource は src の入力途中の行を補修したテキストと、補修した論理行の数を返す。
// lines は splitLines(src) の結果であること（補修しなかったときは src をそのまま返す）。
func repairSource(src string, lines []string) (string, int) {
	var fixed map[int]string // 物理行 → 補修後の行
	n := 0
	for i := 0; i < len(lines); {
		ll := buildLogicalLine(lines, i)
		i = ll.Last + 1
		if !mayNeedRepair(ll.Text) {
			continue
		}
		edits, app, ok := repairLogicalLine(ll)
		if !ok {
			continue
		}
		if fixed == nil {
			fixed = map[int]string{}
		}
		n++
		for _, e := range edits {
			l, ok := fixed[e.line]
			if !ok {
				l = lines[e.line]
			}
			b := []byte(l)
			for k := e.col; k < e.col+e.n && k < len(b); k++ {
				b[k] = ' '
			}
			fixed[e.line] = string(b)
		}
		if app != "" {
			l, ok := fixed[ll.Last]
			if !ok {
				l = lines[ll.Last]
			}
			fixed[ll.Last] = l + app
		}
	}
	if n == 0 {
		return src, 0
	}
	raw := strings.Split(src, "\n")
	if len(raw) != len(lines) {
		return src, 0 // lines が src から作られていない
	}
	for li, l := range fixed {
		if strings.HasSuffix(raw[li], "\r") {
			l += "\r"
		}
		raw[li] = l
	}
	return strings.Join(raw, "\n"), n
}

// mayNeedRepair は補修の対象になりうる行か（大半の行を早く飛ばすための粗い判定）。
func mayNeedRepair(text string) bool {
	if strings.ContainsAny(text, `"(`) {
		return true
	}
	t := trimRightBlank(text)
	if t == "" {
		return false
	}
	switch t[len(t)-1] {
	case '.', ',', '=', '+', '-', '*', '/', '\\', '<', '>', '&', '^':
		return true
	}
	r, _ := utf8.DecodeLastRuneInString(t)
	return unicode.IsLetter(r) // 末尾の語（And / To / New など）
}

// blankEdit は物理行 line のバイト位置 col から n バイトを空白にする補修。
type blankEdit struct{ line, col, n int }

// repairLogicalLine は論理行 1 つの補修を求める。補修が不要なら ok=false。
func repairLogicalLine(ll logicalLine) (edits []blankEdit, app string, ok bool) {
	masked, inStr, cmt := maskCodeAt(ll.Text)
	codeEnd := len(masked)
	if cmt >= 0 {
		codeEnd = cmt
	}
	if len(ll.parts) == 0 || codeEnd < ll.parts[len(ll.parts)-1].logStart {
		return nil, "", false // コードが最後の物理行に無い（コメントが継続している）
	}
	m := []byte(masked[:codeEnd])
	if inStr {
		app += `"`
		m = append(m, '"')
	}
	blank := func(from, to int) {
		for k := from; k < to; k++ {
			m[k] = ' '
			line, col := ll.Physical(k)
			edits = append(edits, blankEdit{line: line, col: col, n: 1})
		}
	}
	isPrint := printStmtRe.Match(m)
	for {
		k := len(trimRightBlank(string(m)))
		if k == 0 {
			break
		}
		c := m[k-1]
		switch {
		case c == '.':
			if s := identStartBack(string(m), k-1); s < k-1 {
				if r, _ := utf8.DecodeRuneInString(string(m[s:])); !unicode.IsLetter(r) {
					goto done // 1. のような数値
				}
			}
			blank(k-1, k)
		case c == ',' && isPrint:
			goto done
		case c == '=' && k >= 2 && m[k-2] == ':':
			blank(k-2, k) // 名前付き引数の ":="
		case strings.IndexByte(",(=+-*/\\<>", c) >= 0:
			blank(k-1, k)
		case c == '&' || c == '^':
			if k >= 2 && !isBlankByte(m[k-2]) && m[k-2] != ')' && m[k-2] != '"' {
				goto done // n& / x^ は型接尾辞
			}
			blank(k-1, k)
		default:
			s := identStartBack(string(m), k)
			if s == k || !trailingOpWords[strings.ToLower(string(m[s:k]))] || (s > 0 && m[s-1] == '.') {
				goto done
			}
			blank(s, k)
		}
	}
done:
	depth := 0
	for _, c := range m {
		switch c {
		case '(':
			depth++
		case ')':
			depth--
		}
	}
	if depth > 0 {
		app += strings.Repeat(")", depth)
	}
	if len(edits) == 0 && app == "" {
		return nil, "", false
	}
	if app != "" && codeEnd < len(ll.Text) {
		// 行末コメントを空白にしてから追記する
		for k := codeEnd; k < len(ll.Text); k++ {
			line, col := ll.Physical(k)
			edits = append(edits, blankEdit{line: line, col: col, n: 1})
		}
	}
	return edits, app, true
}
