package main

import (
	"regexp"
	"strings"
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
	// procStartRe は Sub / Function / Property Get・Let・Set の開始行にマッチする
	// （Public / Private / Friend / Global / Static 修飾子付き、名前の型接尾辞付きも含む）。
	// 構文エラーでパースツリーのスコープが崩れたときのフォールバック専用。
	// グループ1: "Sub" / "Function" / "Property"、グループ2: "Get" / "Let" / "Set"（Property のときのみ）、グループ3: 手続き名。
	procStartRe = regexp.MustCompile(`(?i)^` + wsClass + `*(?:(?:Public|Private|Friend|Global)` + wsClass + `+)?(?:Static` + wsClass + `+)?(Sub|Function|Property` + wsClass + `+(Get|Let|Set))` + wsClass + `+(` + identPattern + `)`)
	// procEndRe は End Sub / End Function / End Property 行にマッチする。
	procEndRe = regexp.MustCompile(`(?i)^` + wsClass + `*End` + wsClass + `+(Sub|Function|Property)\b`)
)

// regexScopes は正規表現でプロシージャの範囲を求める（構文エラー時のフォールバック）。
func regexScopes(lines []string) []*ProcInfo {
	var out []*ProcInfo
	var cur *ProcInfo
	for i, line := range lines {
		if m := procStartRe.FindStringSubmatch(line); m != nil && cur == nil {
			kind := ProcSub
			switch {
			case m[2] != "":
				kind = map[string]ProcKind{"get": ProcPropertyGet, "let": ProcPropertyLet, "set": ProcPropertySet}[strings.ToLower(m[2])]
			case strings.EqualFold(m[1], "Function"):
				kind = ProcFunction
			}
			cur = &ProcInfo{Name: m[3], Kind: kind, HasBody: true, StartLine: i, EndLine: len(lines) - 1}
			out = append(out, cur)
		} else if cur != nil && procEndRe.MatchString(line) {
			cur.EndLine = i
			cur = nil
		}
	}
	return out
}

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
