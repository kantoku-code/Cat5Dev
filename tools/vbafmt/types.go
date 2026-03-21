package main

import (
	"strings"
)

// typeSuffixMap は型接尾辞と対応する型名のマップ
var typeSuffixMap = map[byte]string{
	'%': "Integer",
	'&': "Long",
	'!': "Single",
	'#': "Double",
	'@': "Currency",
	'$': "String",
}

// expandTypeSuffixes は型接尾辞を "As Type" に展開する
// 対象行: Dim/ReDim/Public/Private/Static/Const で始まる行のみ
func expandTypeSuffixes(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = expandTypeSuffixLine(line)
	}
	return result
}

func expandTypeSuffixLine(line string) string {
	trimmed := strings.TrimSpace(line)
	if !isDimStatement(trimmed) {
		return line
	}

	// 先頭インデントを保持
	indent := line[:len(line)-len(strings.TrimLeft(line, " \t"))]

	// コメントとコードを分離
	segs := parseSegments(trimmed)
	var codePart strings.Builder
	var commentPart string
	for _, seg := range segs {
		if seg.kind == segComment {
			commentPart = seg.text
		} else {
			codePart.WriteString(seg.text)
		}
	}

	code := codePart.String()

	// "Dim " などのプレフィックスを取得
	keyword, rest := firstTokenStr(code)
	if rest == "" {
		return line
	}

	// "ReDim Preserve" の場合の処理
	nextTok, nextRest := firstTokenStr(strings.TrimLeft(rest, " \t"))
	if strings.ToLower(keyword) == "redim" && strings.ToLower(nextTok) == "preserve" {
		keyword = keyword + " " + nextTok
		rest = nextRest
	}

	// 変数リストを処理（変数リストはカンマ区切り）
	varList := strings.TrimLeft(rest, " \t")
	expanded := expandVarList(varList)

	result := indent + keyword + " " + expanded
	if commentPart != "" {
		result += " " + commentPart
	}
	return result
}

// isDimStatement は行が変数宣言文かどうかを返す
func isDimStatement(trimmed string) bool {
	tok, _ := firstTokenStr(trimmed)
	switch strings.ToLower(tok) {
	case "dim", "redim", "public", "private", "static", "const":
		return true
	}
	return false
}

// firstTokenStr は行の最初のトークン（識別子）と残りの文字列を返す
func firstTokenStr(line string) (token string, rest string) {
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

// expandVarList は "x%, y&, z As Long" のような変数リストを処理する
func expandVarList(varList string) string {
	// カンマで分割（文字列内のカンマは考慮不要: 変数宣言に文字列リテラルはない）
	vars := splitVarList(varList)
	expanded := make([]string, len(vars))
	for i, v := range vars {
		expanded[i] = expandSingleVar(strings.TrimSpace(v))
	}
	return strings.Join(expanded, ", ")
}

// splitVarList はカンマで変数リストを分割する（括弧内のカンマは無視）
func splitVarList(s string) []string {
	var parts []string
	depth := 0
	start := 0
	for i := 0; i < len(s); i++ {
		switch s[i] {
		case '(':
			depth++
		case ')':
			depth--
		case ',':
			if depth == 0 {
				parts = append(parts, s[start:i])
				start = i + 1
			}
		}
	}
	parts = append(parts, s[start:])
	return parts
}

// expandSingleVar は "x%" のような単一変数宣言を "x As Integer" に展開する
// 既に "As Type" がある場合はそのまま返す
// 配列宣言 "x%(10)" も対応: "x%(10)" → "x(10) As Integer"
func expandSingleVar(v string) string {
	if strings.Contains(strings.ToLower(v), " as ") {
		return v // 既に As Type がある
	}

	// 変数名の末尾の接尾辞を探す
	// 配列の場合: "x%(10)" → 名前="x", 接尾辞='%', 配列部="(10)"
	// 通常の場合: "x%" → 名前="x", 接尾辞='%'

	if len(v) == 0 {
		return v
	}

	// 配列部分を分離
	arrayPart := ""
	base := v
	parenIdx := strings.Index(v, "(")
	if parenIdx >= 0 {
		base = v[:parenIdx]
		arrayPart = v[parenIdx:]
	}

	if len(base) == 0 {
		return v
	}

	lastCh := base[len(base)-1]
	typeName, ok := typeSuffixMap[lastCh]
	if !ok {
		return v // 接尾辞なし
	}

	varName := base[:len(base)-1]
	return varName + arrayPart + " As " + typeName
}
