package main

import (
	"strings"
	"unicode"
)

// Options はフォーマットオプション
type Options struct {
	IndentSize          int
	CapitalizeKeywords  bool
	FixIndentation      bool
	LineEndings         string // "CRLF" or "LF"
}

// DefaultOptions はデフォルトオプション
func DefaultOptions() Options {
	return Options{
		IndentSize:         4,
		CapitalizeKeywords: true,
		FixIndentation:     true,
		LineEndings:        "CRLF",
	}
}

// Format は VBA ソースコードをフォーマットする
func Format(input string, opts Options) string {
	lines := splitLines(input)

	if opts.CapitalizeKeywords {
		lines = capitalizeKeywords(lines)
	}
	if opts.FixIndentation {
		lines = fixIndentation(lines, opts.IndentSize)
	}

	sep := "\n"
	if opts.LineEndings == "CRLF" {
		sep = "\r\n"
	}
	return strings.Join(lines, sep) + sep
}

// splitLines は改行コードを正規化して行スライスを返す
func splitLines(input string) []string {
	// CRLF → LF → LF で統一してから split
	input = strings.ReplaceAll(input, "\r\n", "\n")
	input = strings.ReplaceAll(input, "\r", "\n")
	// 末尾の改行を除去してから split
	input = strings.TrimRight(input, "\n")
	if input == "" {
		return []string{}
	}
	return strings.Split(input, "\n")
}

// ---------- Stage 2: キーワード大文字化 ----------

// capitalizeKeywords は各行のキーワードを正規ケースに変換する
func capitalizeKeywords(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = capitalizeLine(line)
	}
	return result
}

// capitalizeLine は1行のキーワードを変換する
func capitalizeLine(line string) string {
	var out strings.Builder
	inString := false
	i := 0
	for i < len(line) {
		ch := line[i]

		// 文字列リテラル内
		if inString {
			out.WriteByte(ch)
			if ch == '"' {
				inString = false
			}
			i++
			continue
		}

		// 文字列開始
		if ch == '"' {
			inString = true
			out.WriteByte(ch)
			i++
			continue
		}

		// コメント: ' 以降はそのまま出力
		if ch == '\'' {
			out.WriteString(line[i:])
			return out.String()
		}

		// 識別子 (キーワード候補)
		if isIdentStart(rune(ch)) {
			j := i + 1
			for j < len(line) && isIdentPart(rune(line[j])) {
				j++
			}
			word := line[i:j]
			lower := strings.ToLower(word)

			// 2トークン複合キーワードの先読み
			// "end if/for/with/select/function/sub/property/type/enum/class"
			// "select case", "option explicit/base/compare"
			// "for each", "property get/let/set", "exit sub/function/for/do/while"
			if canonical, ok := Keywords[lower]; ok {
				// 次のトークンと合わせて複合チェック
				rest := line[j:]
				nextWord, nextEnd := peekNextWord(rest)
				compound := lower + " " + strings.ToLower(nextWord)
				switch compound {
				case "end if", "end for", "end with", "end select",
					"end function", "end sub", "end property",
					"end type", "end enum", "end class":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "select case":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "option explicit", "option base", "option compare":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "for each":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "property get", "property let", "property set":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "exit sub", "exit function", "exit for", "exit do", "exit while":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "do while", "do until":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "loop while", "loop until":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "elseif": // すでに単一トークンの場合は不要だが念のため
				}
				out.WriteString(canonical)
			} else {
				out.WriteString(word)
			}
			i = j
			continue
		}

		out.WriteByte(ch)
		i++
	}
	return out.String()
}

// peekNextWord は文字列の先頭の空白をスキップして最初の単語と、
// その単語が終わった位置 (空白含む) を返す
func peekNextWord(s string) (word string, consumed int) {
	i := 0
	for i < len(s) && (s[i] == ' ' || s[i] == '\t') {
		i++
	}
	start := i
	for i < len(s) && isIdentPart(rune(s[i])) {
		i++
	}
	return s[start:i], i
}

func capitalize(word string) string {
	if word == "" {
		return ""
	}
	lower := strings.ToLower(word)
	if canonical, ok := Keywords[lower]; ok {
		return canonical
	}
	return word
}

func isIdentStart(r rune) bool {
	return unicode.IsLetter(r) || r == '_'
}

func isIdentPart(r rune) bool {
	return unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_'
}

// ---------- Stage 4: インデント修正 ----------

type lineKind int

const (
	kindNormal    lineKind = iota
	kindStarter            // インデント増加
	kindEnder              // インデント減少 (Else/Case は減少後増加)
	kindElseCase           // 減少→増加 (Else, ElseIf, Case)
	kindHeader             // 常にカラム0
	kindBlank              // 空行
)

func fixIndentation(lines []string, indentSize int) []string {
	result := make([]string, 0, len(lines))
	depth := 0

	for _, line := range lines {
		trimmed := strings.TrimSpace(line)

		if trimmed == "" {
			result = append(result, "")
			continue
		}

		kind := classifyLine(trimmed)

		switch kind {
		case kindHeader:
			result = append(result, trimmed)

		case kindEnder:
			depth--
			if depth < 0 {
				depth = 0
			}
			result = append(result, indent(trimmed, depth, indentSize))

		case kindElseCase:
			depth--
			if depth < 0 {
				depth = 0
			}
			result = append(result, indent(trimmed, depth, indentSize))
			depth++

		case kindStarter:
			result = append(result, indent(trimmed, depth, indentSize))
			depth++

		default: // kindNormal
			result = append(result, indent(trimmed, depth, indentSize))
		}
	}

	return result
}

// classifyLine は行を分類する
func classifyLine(trimmed string) lineKind {
	// コメント行
	if strings.HasPrefix(trimmed, "'") {
		return kindNormal
	}

	// 条件コンパイル (#If, #Else, #End If) → ヘッダ扱い (カラム0)
	if strings.HasPrefix(trimmed, "#") {
		return kindHeader
	}

	// 先頭トークンを取得
	first, rest := firstToken(trimmed)
	firstLow := strings.ToLower(first)

	// 継続行 (前の行が _ で終わる場合はここでは判定できないが、
	// 本行が _ で終わっても分類は通常通り行う)

	// Attribute, BEGIN, VERSION → header
	switch firstLow {
	case "attribute", "begin", "version":
		return kindHeader
	}

	// End 系 → ender
	if firstLow == "end" {
		return kindEnder
	}

	// Loop, Next, Wend → ender
	if firstLow == "loop" || firstLow == "next" || firstLow == "wend" {
		return kindEnder
	}

	// Else, ElseIf → elseCase
	if firstLow == "else" || firstLow == "elseif" {
		return kindElseCase
	}

	// Case → elseCase (Select Case ブロック内)
	// ただし "Case" で始まる行のみ (Select Case 自体は starter)
	if firstLow == "case" {
		return kindElseCase
	}

	// Select Case → starter
	if firstLow == "select" {
		return kindStarter
	}

	// For, For Each → starter
	if firstLow == "for" {
		return kindStarter
	}

	// Do, Do While, Do Until → starter
	if firstLow == "do" {
		return kindStarter
	}

	// While → starter
	if firstLow == "while" {
		return kindStarter
	}

	// With → starter
	if firstLow == "with" {
		return kindStarter
	}

	// Sub, Function → starter
	if firstLow == "sub" || firstLow == "function" {
		return kindStarter
	}

	// Public/Private/Friend/Static + Sub/Function/Property → starter
	if firstLow == "public" || firstLow == "private" || firstLow == "friend" ||
		firstLow == "static" || firstLow == "global" {
		nextTok, _ := firstToken(strings.TrimSpace(rest))
		nextLow := strings.ToLower(nextTok)
		if nextLow == "sub" || nextLow == "function" || nextLow == "property" {
			return kindStarter
		}
	}

	// Property Get/Let/Set → starter
	if firstLow == "property" {
		return kindStarter
	}

	// Type, Enum, Class → starter
	if firstLow == "type" || firstLow == "enum" || firstLow == "class" {
		return kindStarter
	}

	// If ... Then → starter か normal かを判定
	// 単行 If: "If <expr> Then <stmt>" → normal
	// ブロック If: "If <expr> Then" (Then が行末) → starter
	if firstLow == "if" {
		if isSingleLineIf(trimmed) {
			return kindNormal
		}
		return kindStarter
	}

	return kindNormal
}

// isSingleLineIf は "If ... Then <stmt>" 形式かどうかを判定する
// Then の後にコメント以外のトークンがあれば単行 If
func isSingleLineIf(line string) bool {
	// Then を探す (文字列・コメント考慮の簡易版)
	lower := strings.ToLower(line)
	idx := strings.LastIndex(lower, "then")
	if idx < 0 {
		return false
	}
	// Then の後の文字列
	after := strings.TrimSpace(line[idx+4:])
	// コメントまたは空なら ブロック If
	if after == "" || strings.HasPrefix(after, "'") {
		return false
	}
	return true
}

// firstToken は行の最初のトークン (識別子) と残りの文字列を返す
func firstToken(line string) (token string, rest string) {
	line = strings.TrimSpace(line)
	i := 0
	for i < len(line) && isIdentPart(rune(line[i])) {
		i++
	}
	if i == 0 {
		return "", line
	}
	return line[:i], line[i:]
}

// indent はインデントを付与した文字列を返す
func indent(line string, depth, size int) string {
	if depth <= 0 {
		return line
	}
	return strings.Repeat(" ", depth*size) + line
}
