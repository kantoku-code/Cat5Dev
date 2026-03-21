package main

import (
	"strings"
)

// Options はフォーマットオプション
type Options struct {
	IndentSize         int
	CapitalizeKeywords bool
	FixIndentation     bool
	LineEndings        string // "CRLF" or "LF"

	// 高優先度
	NormalizeOperatorSpacing bool
	TrimTrailingSpace        bool
	IndentContinuationLines  bool
	MaxBlankLines            int // 0=無効

	// 中優先度
	NormalizeCommaSpacing   bool
	SplitColonStatements    bool
	NormalizeThenPlacement  bool
	NormalizeCommentSpace   bool

	// 低優先度
	ExpandTypeSuffixes bool
	NormalizeOnError   bool
}

// DefaultOptions はデフォルトオプション
func DefaultOptions() Options {
	return Options{
		IndentSize:         4,
		CapitalizeKeywords: true,
		FixIndentation:     true,
		LineEndings:        "CRLF",

		NormalizeOperatorSpacing: false,
		TrimTrailingSpace:        true,
		IndentContinuationLines:  true,
		MaxBlankLines:            2,

		NormalizeCommaSpacing:  false,
		SplitColonStatements:   false,
		NormalizeThenPlacement: false,
		NormalizeCommentSpace:  false,

		ExpandTypeSuffixes: false,
		NormalizeOnError:   false,
	}
}

// Format は VBA ソースコードをフォーマットする
func Format(input string, opts Options) string {
	lines := splitLines(input)

	if opts.NormalizeThenPlacement {
		lines = normalizeThenPlacement(lines)
	}
	if opts.SplitColonStatements {
		lines = splitColonStatements(lines)
	}
	if opts.TrimTrailingSpace {
		lines = trimTrailingSpace(lines)
	}
	if opts.CapitalizeKeywords {
		lines = capitalizeKeywords(lines)
	}
	if opts.ExpandTypeSuffixes {
		lines = expandTypeSuffixes(lines)
	}
	if opts.NormalizeOperatorSpacing {
		lines = normalizeOperatorSpacing(lines)
	}
	if opts.NormalizeCommaSpacing {
		lines = normalizeCommaSpacing(lines)
	}
	if opts.NormalizeCommentSpace {
		lines = normalizeCommentSpace(lines)
	}
	// NormalizeOnError: 将来実装予定（現在はスタブ）
	if opts.FixIndentation {
		lines = fixIndentation(lines, opts.IndentSize, opts.IndentContinuationLines)
	}
	if opts.MaxBlankLines > 0 {
		lines = normalizeBlankLines(lines, opts.MaxBlankLines)
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

// ---------- Stage 4: インデント修正 ----------

type lineKind int

const (
	kindNormal        lineKind = iota
	kindStarter                // インデント増加 (+1)
	kindSelectStarter          // Select Case: +2 (Case が -1 するため)
	kindEnder                  // インデント減少 (-1)
	kindEndSelect              // End Select: -2
	kindElseCase               // 減少→増加 (Else, ElseIf, Case): -1 then +1
	kindHeader                 // 常にカラム0
	kindBlank                  // 空行
)

func fixIndentation(lines []string, indentSize int, indentContinuation bool) []string {
	result := make([]string, 0, len(lines))
	depth := 0
	prevWasContinuation := false

	for _, line := range lines {
		trimmed := strings.TrimSpace(line)

		if trimmed == "" {
			result = append(result, "")
			prevWasContinuation = false
			continue
		}

		kind := classifyLine(trimmed)

		// 継続行のインデント調整
		extraIndent := 0
		if indentContinuation && prevWasContinuation {
			extraIndent = 1
		}

		switch kind {
		case kindHeader:
			result = append(result, trimmed)

		case kindEnder:
			depth--
			if depth < 0 {
				depth = 0
			}
			result = append(result, indent(trimmed, depth+extraIndent, indentSize))

		case kindEndSelect:
			depth -= 2
			if depth < 0 {
				depth = 0
			}
			result = append(result, indent(trimmed, depth+extraIndent, indentSize))

		case kindElseCase:
			depth--
			if depth < 0 {
				depth = 0
			}
			result = append(result, indent(trimmed, depth+extraIndent, indentSize))
			depth++

		case kindStarter:
			result = append(result, indent(trimmed, depth+extraIndent, indentSize))
			depth++

		case kindSelectStarter:
			result = append(result, indent(trimmed, depth+extraIndent, indentSize))
			depth += 2

		default: // kindNormal
			result = append(result, indent(trimmed, depth+extraIndent, indentSize))
		}

		// 継続行フラグ更新（コード末尾が _ かどうか）
		prevWasContinuation = indentContinuation && isContinuationLine(trimmed)
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

	// Attribute, BEGIN, VERSION → header
	switch firstLow {
	case "attribute", "begin", "version":
		return kindHeader
	}

	// End 系
	if firstLow == "end" {
		// End Select は -2
		nextTok, _ := firstToken(strings.TrimSpace(rest))
		if strings.ToLower(nextTok) == "select" {
			return kindEndSelect
		}
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
	if firstLow == "case" {
		return kindElseCase
	}

	// Select Case → selectStarter (+2)
	if firstLow == "select" {
		return kindSelectStarter
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
	if firstLow == "if" {
		if isSingleLineIf(trimmed) {
			return kindNormal
		}
		return kindStarter
	}

	return kindNormal
}

// isSingleLineIf は "If ... Then <stmt>" 形式かどうかを判定する
func isSingleLineIf(line string) bool {
	lower := strings.ToLower(line)
	idx := strings.LastIndex(lower, "then")
	if idx < 0 {
		return false
	}
	after := strings.TrimSpace(line[idx+4:])
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
