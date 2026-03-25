package main

import (
	"fmt"
	"regexp"
	"strings"
	"unicode/utf8"
)

// LintOptions は Lint チェックのオプション
type LintOptions struct {
	OptionExplicit    bool `json:"option_explicit"`
	OnErrorResumeNext bool `json:"on_error_resume_next"`
	Goto              bool `json:"goto"`
	MaxLineLength     int  `json:"max_line_length"`
	UnusedVariables   bool `json:"unused_variables"`
	MaxNestingDepth   int  `json:"max_nesting_depth"`
	MaxFunctionLines  int  `json:"max_function_lines"`
	UnmatchedParens   bool `json:"unmatched_parens"`
	UnmatchedBlocks   bool `json:"unmatched_blocks"`
}

// DefaultLintOptions はデフォルト Lint オプション
func DefaultLintOptions() LintOptions {
	return LintOptions{
		OptionExplicit:    true,
		OnErrorResumeNext: true,
		Goto:              true,
		MaxLineLength:     200,
		UnusedVariables:   true,
		MaxNestingDepth:   5,
		MaxFunctionLines:  100,
		UnmatchedParens:   true,
		UnmatchedBlocks:   true,
	}
}

// Diagnostic は診断結果の1件
type Diagnostic struct {
	Line     int    `json:"line"`
	Col      int    `json:"col"`
	EndLine  int    `json:"end_line"`
	EndCol   int    `json:"end_col"`
	Severity string `json:"severity"`
	Code     string `json:"code"`
	Message  string `json:"message"`
}

// Lint は VBA ソースコードを解析して診断結果を返す
func Lint(code string, opts LintOptions) []Diagnostic {
	lines := splitLines(code)
	var diags []Diagnostic

	diags = append(diags, checkOptionExplicit(lines, opts)...)
	diags = append(diags, checkOnErrorResumeNext(lines, opts)...)
	diags = append(diags, checkGoto(lines, opts)...)
	diags = append(diags, checkLineLength(lines, opts)...)
	diags = append(diags, checkNestingDepth(lines, opts)...)
	diags = append(diags, checkFunctionLines(lines, opts)...)
	diags = append(diags, checkUnmatchedParens(lines, opts)...)
	diags = append(diags, checkUnmatchedBlocks(lines, opts)...)
	diags = append(diags, checkUnusedVariables(lines, opts)...)

	return diags
}

// ---------- ヘルパー ----------

func warn(line, col, endLine, endCol int, code, msg string) Diagnostic {
	return Diagnostic{line, col, endLine, endCol, "warning", code, msg}
}

func errDiag(line, col, endLine, endCol int, code, msg string) Diagnostic {
	return Diagnostic{line, col, endLine, endCol, "error", code, msg}
}

// codeText は行のコード部分のみを小文字で返す
func codeTextLower(line string) string {
	segs := parseSegments(line)
	var b strings.Builder
	for _, s := range segs {
		if s.kind == segCode {
			b.WriteString(s.text)
		}
	}
	return strings.ToLower(b.String())
}

// codeText は行のコード部分のみを返す（大文字小文字そのまま）
func codeText(line string) string {
	segs := parseSegments(line)
	var b strings.Builder
	for _, s := range segs {
		if s.kind == segCode {
			b.WriteString(s.text)
		}
	}
	return b.String()
}

// isBlankOrComment は空行またはコメント行かどうかを返す
func isBlankOrComment(line string) bool {
	t := strings.TrimSpace(line)
	return t == "" || strings.HasPrefix(t, "'")
}

// ---------- VBA001: Option Explicit ----------

func checkOptionExplicit(lines []string, opts LintOptions) []Diagnostic {
	if !opts.OptionExplicit {
		return nil
	}
	for _, line := range lines {
		if strings.Contains(strings.ToLower(codeText(line)), "option explicit") {
			return nil
		}
	}
	return []Diagnostic{warn(0, 0, 0, -1, "VBA001", "Option Explicit が宣言されていません")}
}

// ---------- VBA002: On Error Resume Next ----------

func checkOnErrorResumeNext(lines []string, opts LintOptions) []Diagnostic {
	if !opts.OnErrorResumeNext {
		return nil
	}
	var diags []Diagnostic
	for i, line := range lines {
		if strings.Contains(codeTextLower(line), "on error resume next") {
			diags = append(diags, warn(i, 0, i, -1, "VBA002",
				"On Error Resume Next はエラーを無視します。適切なエラーハンドリングを検討してください"))
		}
	}
	return diags
}

// ---------- VBA003: GoTo ----------

var reThen = regexp.MustCompile(`(?i)\bthen\b`)

func checkGoto(lines []string, opts LintOptions) []Diagnostic {
	if !opts.Goto {
		return nil
	}
	var diags []Diagnostic
	for i, line := range lines {
		low := codeTextLower(line)
		if !strings.Contains(low, "goto") {
			continue
		}
		// On Error GoTo は除外
		if strings.Contains(low, "on error") {
			continue
		}
		diags = append(diags, warn(i, 0, i, -1, "VBA003",
			"GoTo 文はフロー制御を複雑にします。構造化制御フローの使用を検討してください"))
	}
	return diags
}

// ---------- VBA004: 最大行長 ----------

func checkLineLength(lines []string, opts LintOptions) []Diagnostic {
	if opts.MaxLineLength <= 0 {
		return nil
	}
	var diags []Diagnostic
	for i, line := range lines {
		// CRLF を除いた実際の文字数（rune 単位）
		l := strings.TrimRight(line, "\r")
		if utf8.RuneCountInString(l) > opts.MaxLineLength {
			diags = append(diags, warn(i, opts.MaxLineLength, i, -1, "VBA004",
				fmt.Sprintf("行が %d 文字を超えています（%d 文字）",
					opts.MaxLineLength, utf8.RuneCountInString(l))))
		}
	}
	return diags
}

// ---------- VBA006: ネスト深さ ----------

func checkNestingDepth(lines []string, opts LintOptions) []Diagnostic {
	if opts.MaxNestingDepth <= 0 {
		return nil
	}
	var diags []Diagnostic
	depth := 0

	for i, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		kind := classifyLine(trimmed)
		switch kind {
		case kindStarter:
			depth++
			if depth > opts.MaxNestingDepth {
				diags = append(diags, warn(i, 0, i, -1, "VBA006",
					fmt.Sprintf("ネストの深さが %d を超えています（現在 %d）",
						opts.MaxNestingDepth, depth)))
			}
		case kindSelectStarter:
			depth += 2
			if depth > opts.MaxNestingDepth {
				diags = append(diags, warn(i, 0, i, -1, "VBA006",
					fmt.Sprintf("ネストの深さが %d を超えています（現在 %d）",
						opts.MaxNestingDepth, depth)))
			}
		case kindEnder:
			depth--
			if depth < 0 {
				depth = 0
			}
		case kindEndSelect:
			depth -= 2
			if depth < 0 {
				depth = 0
			}
		case kindElseCase:
			// 深さ変化なし（-1 then +1）
		}
	}
	return diags
}

// ---------- VBA007: 関数行数 ----------

// isSubFunctionStart は Sub/Function/Property の開始行かどうかを返す
func isSubFunctionStart(trimmed string) bool {
	first, rest := firstToken(trimmed)
	low := strings.ToLower(first)
	// アクセス修飾子をスキップ
	if low == "public" || low == "private" || low == "friend" || low == "static" {
		next, _ := firstToken(strings.TrimSpace(rest))
		low = strings.ToLower(next)
	}
	return low == "sub" || low == "function" || low == "property"
}

// isSubFunctionEnd は End Sub/Function/Property の行かどうかを返す
func isSubFunctionEnd(trimmed string) bool {
	first, rest := firstToken(trimmed)
	if strings.ToLower(first) != "end" {
		return false
	}
	next, _ := firstToken(strings.TrimSpace(rest))
	low := strings.ToLower(next)
	return low == "sub" || low == "function" || low == "property"
}

func checkFunctionLines(lines []string, opts LintOptions) []Diagnostic {
	if opts.MaxFunctionLines <= 0 {
		return nil
	}
	var diags []Diagnostic
	startLine := -1

	for i, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" {
			continue
		}
		if isSubFunctionStart(trimmed) && startLine < 0 {
			startLine = i
			continue
		}
		if isSubFunctionEnd(trimmed) && startLine >= 0 {
			count := i - startLine + 1
			if count > opts.MaxFunctionLines {
				diags = append(diags, warn(startLine, 0, startLine, -1, "VBA007",
					fmt.Sprintf("Sub/Function が %d 行を超えています（%d 行）",
						opts.MaxFunctionLines, count)))
			}
			startLine = -1
		}
	}
	return diags
}

// ---------- VBA008: 括弧の不一致 ----------

func checkUnmatchedParens(lines []string, opts LintOptions) []Diagnostic {
	if !opts.UnmatchedParens {
		return nil
	}
	var diags []Diagnostic
	// 継続行（_で終わる行）を結合してステートメント単位で検査
	i := 0
	for i < len(lines) {
		stmtStart := i
		var stmtCode strings.Builder
		// 継続行を結合
		for i < len(lines) {
			line := lines[i]
			ct := codeText(line)
			stmtCode.WriteString(ct)
			i++
			// 末尾が _ なら継続
			if !strings.HasSuffix(strings.TrimSpace(ct), "_") {
				break
			}
			// _ を除去
			s := stmtCode.String()
			stmtCode.Reset()
			stmtCode.WriteString(strings.TrimRight(strings.TrimSuffix(strings.TrimSpace(s), "_"), " "))
			stmtCode.WriteString(" ")
		}
		code := stmtCode.String()
		open := strings.Count(code, "(")
		close := strings.Count(code, ")")
		if open != close {
			diags = append(diags, errDiag(stmtStart, 0, i-1, -1, "VBA008",
				fmt.Sprintf("括弧が一致していません（'(' %d 個, ')' %d 個）", open, close)))
		}
	}
	return diags
}

// ---------- VBA009: End 忘れ ----------

type blockFrame struct {
	line    int
	keyword string
}

func checkUnmatchedBlocks(lines []string, opts LintOptions) []Diagnostic {
	if !opts.UnmatchedBlocks {
		return nil
	}
	var diags []Diagnostic
	var stack []blockFrame

	for i, line := range lines {
		trimmed := strings.TrimSpace(line)
		if trimmed == "" || strings.HasPrefix(trimmed, "'") {
			continue
		}

		first, rest := firstToken(trimmed)
		low := strings.ToLower(first)

		// アクセス修飾子をスキップして実際のキーワードを得る
		actualLow := low
		actualRest := rest
		if low == "public" || low == "private" || low == "friend" || low == "static" {
			tok, r := firstToken(strings.TrimSpace(rest))
			actualLow = strings.ToLower(tok)
			actualRest = r
		}

		switch actualLow {
		case "sub", "function":
			stack = append(stack, blockFrame{i, actualLow})
		case "property":
			stack = append(stack, blockFrame{i, "property"})
		case "if":
			if !isSingleLineIf(trimmed) {
				stack = append(stack, blockFrame{i, "if"})
			}
		case "for":
			stack = append(stack, blockFrame{i, "for"})
		case "do":
			stack = append(stack, blockFrame{i, "do"})
		case "while":
			stack = append(stack, blockFrame{i, "while"})
		case "with":
			stack = append(stack, blockFrame{i, "with"})
		case "select":
			stack = append(stack, blockFrame{i, "select"})
		case "type", "enum":
			stack = append(stack, blockFrame{i, actualLow})
		case "end":
			nextTok, _ := firstToken(strings.TrimSpace(actualRest))
			nextLow := strings.ToLower(nextTok)
			expected := endKeywordFor(nextLow)
			if expected == "" {
				break
			}
			if len(stack) == 0 {
				diags = append(diags, errDiag(i, 0, i, -1, "VBA009",
					fmt.Sprintf("対応する開始ブロックのない '%s %s' です", first, nextTok)))
				break
			}
			top := stack[len(stack)-1]
			if top.keyword != expected {
				diags = append(diags, errDiag(i, 0, i, -1, "VBA009",
					fmt.Sprintf("'%s %s' が想定されましたが '%s %s' があります（開始: %d 行目）",
						openerFor(top.keyword), closerFor(top.keyword),
						first, nextTok, top.line+1)))
			}
			stack = stack[:len(stack)-1]
		case "next":
			if len(stack) == 0 || stack[len(stack)-1].keyword != "for" {
				if len(stack) == 0 {
					diags = append(diags, errDiag(i, 0, i, -1, "VBA009",
						"対応する For のない Next です"))
				}
				// スタック不一致でも Pop して続行
			} else {
				stack = stack[:len(stack)-1]
			}
		case "loop":
			if len(stack) == 0 || stack[len(stack)-1].keyword != "do" {
				if len(stack) == 0 {
					diags = append(diags, errDiag(i, 0, i, -1, "VBA009",
						"対応する Do のない Loop です"))
				}
			} else {
				stack = stack[:len(stack)-1]
			}
		case "wend":
			if len(stack) == 0 || stack[len(stack)-1].keyword != "while" {
				if len(stack) == 0 {
					diags = append(diags, errDiag(i, 0, i, -1, "VBA009",
						"対応する While のない Wend です"))
				}
			} else {
				stack = stack[:len(stack)-1]
			}
		}
	}

	// スタックに残ったものはすべて End 忘れ
	for _, frame := range stack {
		diags = append(diags, errDiag(frame.line, 0, frame.line, -1, "VBA009",
			fmt.Sprintf("'%s %s' がありません（開始: %d 行目）",
				openerFor(frame.keyword), closerFor(frame.keyword), frame.line+1)))
	}

	return diags
}

func endKeywordFor(nextLow string) string {
	switch nextLow {
	case "sub":
		return "sub"
	case "function":
		return "function"
	case "property":
		return "property"
	case "if":
		return "if"
	case "with":
		return "with"
	case "select":
		return "select"
	case "type":
		return "type"
	case "enum":
		return "enum"
	}
	return ""
}

func openerFor(kw string) string {
	switch kw {
	case "sub":
		return "Sub"
	case "function":
		return "Function"
	case "property":
		return "Property"
	case "if":
		return "If"
	case "for":
		return "For"
	case "do":
		return "Do"
	case "while":
		return "While"
	case "with":
		return "With"
	case "select":
		return "Select"
	case "type":
		return "Type"
	case "enum":
		return "Enum"
	}
	return kw
}

func closerFor(kw string) string {
	switch kw {
	case "for":
		return "Next"
	case "do":
		return "Loop"
	case "while":
		return "Wend"
	case "sub":
		return "End Sub"
	case "function":
		return "End Function"
	case "property":
		return "End Property"
	case "if":
		return "End If"
	case "with":
		return "End With"
	case "select":
		return "End Select"
	case "type":
		return "End Type"
	case "enum":
		return "End Enum"
	}
	return "End " + kw
}

// ---------- VBA005: 未使用変数 ----------

// reDim はDim宣言を解析する正規表現
// Dim varName [As Type][, varName2 [As Type2]]...
var reDimVar = regexp.MustCompile(`(?i)\bDim\s+(\w+)`)

// isWordBoundary はインデックス位置がワード境界かどうかを判定する
func isWordBoundary(s string, start, end int) bool {
	if start > 0 {
		ch := rune(s[start-1])
		if isIdentPart(ch) {
			return false
		}
	}
	if end < len(s) {
		ch := rune(s[end])
		if isIdentPart(ch) {
			return false
		}
	}
	return true
}

func checkUnusedVariables(lines []string, opts LintOptions) []Diagnostic {
	if !opts.UnusedVariables {
		return nil
	}
	var diags []Diagnostic

	// Sub/Function スコープ単位で解析
	type varInfo struct {
		name string
		line int
	}

	scopeStart := -1
	var scopeLines []string
	var scopeStartIdx int

	flushScope := func() {
		if scopeStart < 0 || len(scopeLines) == 0 {
			return
		}
		// Dim 宣言を収集
		var declared []varInfo
		for j, sl := range scopeLines {
			matches := reDimVar.FindAllStringSubmatchIndex(sl, -1)
			for _, m := range matches {
				varName := sl[m[2]:m[3]]
				declared = append(declared, varInfo{varName, scopeStartIdx + j})
			}
		}

		// 各変数が宣言行以外で使われているか確認
		for _, v := range declared {
			used := false
			nameLow := strings.ToLower(v.name)
			for j, sl := range scopeLines {
				if scopeStartIdx+j == v.line {
					continue // 宣言行自体はスキップ
				}
				code := codeText(sl)
				codeLow := strings.ToLower(code)
				// ワード境界マッチ
				idx := 0
				for {
					pos := strings.Index(codeLow[idx:], nameLow)
					if pos < 0 {
						break
					}
					abs := idx + pos
					if isWordBoundary(codeLow, abs, abs+len(nameLow)) {
						used = true
						break
					}
					idx = abs + 1
				}
				if used {
					break
				}
			}
			if !used {
				diags = append(diags, warn(v.line, 0, v.line, -1, "VBA005",
					fmt.Sprintf("変数 '%s' は宣言されていますが使用されていません", v.name)))
			}
		}
		scopeStart = -1
		scopeLines = nil
	}

	for i, line := range lines {
		trimmed := strings.TrimSpace(line)
		if isSubFunctionStart(trimmed) {
			flushScope()
			scopeStart = i
			scopeStartIdx = i
			scopeLines = []string{line}
		} else if isSubFunctionEnd(trimmed) {
			if scopeStart >= 0 {
				scopeLines = append(scopeLines, line)
				flushScope()
			}
		} else if scopeStart >= 0 {
			scopeLines = append(scopeLines, line)
		}
	}
	flushScope()

	return diags
}
