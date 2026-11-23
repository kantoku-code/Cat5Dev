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

	for _, st := range logicalStatements(lines) {
		i := st.line
		kind := kindNormal
		if !st.singleIf {
			kind = classifyLine(st.text)
		}
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
			if !isContinuationLine(strings.TrimSpace(line)) {
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

// ---------- 論理文（VBA006 / VBA009 共通） ----------

// logicalStmt は構造（ブロックの開始・終了）を判定するための 1 文。
// コメントは除き、文字列リテラルは "" に置き換えてある。
type logicalStmt struct {
	line     int    // 文が始まる物理行（0 始まり）
	text     string // トリム済みのコード
	singleIf bool   // 単一行 If（If c Then stmt）。行の残りの文もこれに含む
}

// structuralCode は行のコメントを除き、文字列リテラルを "" に置き換えたコードを返す。
// 2 つ目の戻り値は、コメントが " _" で終わり次の行もコメントの続きになるかどうか。
func structuralCode(line string) (string, bool) {
	var b strings.Builder
	commentCont := false
	for _, s := range parseSegments(line) {
		switch s.kind {
		case segCode:
			b.WriteString(s.text)
		case segString:
			b.WriteString(`""`)
		case segComment:
			commentCont = endsWithContinuation(s.text)
		}
	}
	return strings.TrimSpace(b.String()), commentCont
}

// endsWithContinuation は text が空白 + "_" で終わるかどうかを返す
func endsWithContinuation(text string) bool {
	t := strings.TrimRight(text, " \t")
	if !strings.HasSuffix(t, "_") {
		return false
	}
	t = t[:len(t)-1]
	return t == "" || strings.HasSuffix(t, " ") || strings.HasSuffix(t, "\t")
}

// logicalStatements は行を構造判定用の文に分ける。
//   - 継続行（末尾 _）は結合する（文の行は先頭の物理行）
//   - コロン区切りは別の文に分ける（:= は分けない）
//   - 単一行 If は行の残りをまとめて 1 文にする（If a Then For ...: Next の Next を浮かせない）
//   - 条件付きコンパイル（#If）は最初の分岐だけを対象にし、#ElseIf / #Else の分岐は読み飛ばす
func logicalStatements(lines []string) []logicalStmt {
	type ccFrame struct {
		parentActive bool
		branch       int
	}
	var cc []ccFrame
	active := func() bool {
		if len(cc) == 0 {
			return true
		}
		f := cc[len(cc)-1]
		return f.parentActive && f.branch == 0
	}

	var out []logicalStmt
	commentCont := false
	for i := 0; i < len(lines); i++ {
		if commentCont {
			// 前の行のコメントが " _" で続いている
			commentCont = endsWithContinuation(lines[i])
			continue
		}
		trimmed := strings.TrimSpace(lines[i])
		if strings.HasPrefix(trimmed, "#") {
			tok, _ := firstToken(trimmed[1:])
			switch strings.ToLower(tok) {
			case "if":
				cc = append(cc, ccFrame{parentActive: active()})
			case "elseif", "else":
				if len(cc) > 0 {
					cc[len(cc)-1].branch++
				}
			case "end":
				if len(cc) > 0 {
					cc = cc[:len(cc)-1]
				}
			}
			continue
		}

		start := i
		code, cont := structuralCode(lines[i])
		commentCont = cont
		for strings.HasSuffix(code, "_") && isContinuationLine(strings.TrimSpace(lines[i])) && i+1 < len(lines) {
			code = strings.TrimSpace(strings.TrimSuffix(code, "_"))
			i++
			next, cont := structuralCode(lines[i])
			commentCont = cont
			code += " " + next
		}
		if code == "" || !active() {
			continue
		}
		out = append(out, splitStatements(start, code)...)
	}
	return out
}

// splitStatements は 1 論理行のコードをコロンで文に分ける
func splitStatements(line int, code string) []logicalStmt {
	var parts []string
	var cur strings.Builder
	for k := 0; k < len(code); k++ {
		ch := code[k]
		if ch == ':' && !(k+1 < len(code) && code[k+1] == '=') {
			if p := strings.TrimSpace(cur.String()); p != "" {
				parts = append(parts, p)
			}
			cur.Reset()
			continue
		}
		cur.WriteByte(ch)
	}
	if p := strings.TrimSpace(cur.String()); p != "" {
		parts = append(parts, p)
	}

	var out []logicalStmt
	for k, p := range parts {
		if first, _ := firstToken(p); strings.EqualFold(first, "if") {
			rest := strings.Join(parts[k:], ": ")
			if isSingleLineIfCode(rest) {
				out = append(out, logicalStmt{line: line, text: rest, singleIf: true})
				break
			}
		}
		out = append(out, logicalStmt{line: line, text: p})
	}
	return out
}

// isSingleLineIfCode は（コメント・文字列除去済みの）If 文が単一行 If かどうかを返す。
// 最後の Then の後ろに文があれば単一行 If。
func isSingleLineIfCode(code string) bool {
	locs := reThen.FindAllStringIndex(code, -1)
	if len(locs) == 0 {
		return false
	}
	return strings.TrimSpace(code[locs[len(locs)-1][1]:]) != ""
}

// ---------- VBA009: End 忘れ ----------

type blockFrame struct {
	line    int
	keyword string
}

func isProcKeyword(kw string) bool {
	return kw == "sub" || kw == "function" || kw == "property"
}

// blockKeyword は文の先頭キーワード（アクセス修飾子を除く）を小文字で返す
func blockKeyword(text string) (string, string) {
	tok, rest := firstToken(text)
	low := strings.ToLower(tok)
	for low == "public" || low == "private" || low == "friend" || low == "static" || low == "global" {
		tok, rest = firstToken(strings.TrimSpace(rest))
		low = strings.ToLower(tok)
	}
	return low, strings.TrimSpace(rest)
}

func orphanCloserMessage(kw string) string {
	switch kw {
	case "for", "do", "while":
		return fmt.Sprintf("対応する %s のない %s です", openerFor(kw), closerFor(kw))
	}
	return fmt.Sprintf("対応する開始ブロックのない '%s' です", closerFor(kw))
}

func checkUnmatchedBlocks(lines []string, opts LintOptions) []Diagnostic {
	if !opts.UnmatchedBlocks {
		return nil
	}
	var diags []Diagnostic
	var stack []blockFrame

	missing := func(f blockFrame) {
		diags = append(diags, errDiag(f.line, 0, f.line, -1, "VBA009",
			fmt.Sprintf("'%s' がありません（開始: %d 行目）", closerFor(f.keyword), f.line+1)))
	}
	// closeBlock は kw の閉じを処理する。
	// 一番上が違っても下に対応する開始があれば、間のブロックを閉じ忘れとして報告してそこまで戻す。
	// プロシージャの境界は越えない。対応する開始がなければ閉じだけを報告し、スタックは変えない。
	closeBlock := func(line int, kw string) {
		j := -1
		for idx := len(stack) - 1; idx >= 0; idx-- {
			if stack[idx].keyword == kw {
				j = idx
				break
			}
			if isProcKeyword(stack[idx].keyword) && !isProcKeyword(kw) {
				break
			}
		}
		if j < 0 {
			diags = append(diags, errDiag(line, 0, line, -1, "VBA009", orphanCloserMessage(kw)))
			return
		}
		for idx := len(stack) - 1; idx > j; idx-- {
			missing(stack[idx])
		}
		stack = stack[:j]
	}

	for _, st := range logicalStatements(lines) {
		if st.singleIf {
			continue
		}
		kw, rest := blockKeyword(st.text)
		switch kw {
		case "sub", "function", "property":
			// プロシージャは入れ子にならない。前のプロシージャの閉じ忘れはここで報告して持ち越さない
			for _, f := range stack {
				missing(f)
			}
			stack = append(stack[:0], blockFrame{st.line, kw})
		case "if", "for", "do", "while", "with", "select", "type", "enum":
			stack = append(stack, blockFrame{st.line, kw})
		case "end":
			nextTok, _ := firstToken(rest)
			if expected := endKeywordFor(strings.ToLower(nextTok)); expected != "" {
				closeBlock(st.line, expected)
			}
		case "next":
			// Next i, j は 2 つ閉じる
			for n := strings.Count(rest, ",") + 1; n > 0; n-- {
				closeBlock(st.line, "for")
			}
		case "loop":
			closeBlock(st.line, "do")
		case "wend":
			closeBlock(st.line, "while")
		}
	}

	// スタックに残ったものはすべて End 忘れ
	for _, f := range stack {
		missing(f)
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
		// Dim 宣言を収集（コメント・文字列リテラルは除外）
		var declared []varInfo
		for j, sl := range scopeLines {
			codeLine := codeText(sl)
			matches := reDimVar.FindAllStringSubmatchIndex(codeLine, -1)
			for _, m := range matches {
				varName := codeLine[m[2]:m[3]]
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
