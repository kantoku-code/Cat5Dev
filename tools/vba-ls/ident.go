package main

import (
	"regexp"
	"unicode"
	"unicode/utf8"
)

// VBA 識別子・空白の共通パターン。日本語などの Unicode 識別子を扱うため \w（ASCII のみ）は使わないこと。
// 文法（antlr/vba.g4 の IDENTIFIER）は先頭 "_" を許さないが、行内走査用の正規表現では許容している。
const (
	identStartClass = `\p{L}_`
	identPartClass  = `\p{L}\p{Nd}_`
	identPattern    = `[` + identStartClass + `][` + identPartClass + `]*`
	// wsClass は空白 1 文字。RE2 の \s は全角空白（U+3000）を含まないため \p{Zs} を加える。
	wsClass = `[\s\p{Zs}]`
)

var (
	// procStartRe は Sub / Function の開始行にマッチする（Public / Private / Friend / Global / Static 修飾子付きも含む）。
	// グループ1: "Sub" または "Function"、グループ2: 手続き名。
	procStartRe = regexp.MustCompile(`(?i)^` + wsClass + `*(?:(?:Public|Private|Friend|Global)` + wsClass + `+)?(?:Static` + wsClass + `+)?(Sub|Function)` + wsClass + `+(` + identPattern + `)`)
	// procEndRe は End Sub / End Function 行にマッチする。
	procEndRe = regexp.MustCompile(`(?i)^` + wsClass + `*End` + wsClass + `+(Sub|Function)`)
)

// isIdentPart は r が識別子の 2 文字目以降に使える文字かを返す。
func isIdentPart(r rune) bool {
	return r == '_' || unicode.IsLetter(r) || unicode.IsDigit(r)
}

// precededByIdentPart は s のバイト位置 off の直前が識別子文字かを返す（正規表現の先頭側 \b の代わり）。
func precededByIdentPart(s string, off int) bool {
	if off <= 0 {
		return false
	}
	r, _ := utf8.DecodeLastRuneInString(s[:off])
	return isIdentPart(r)
}
