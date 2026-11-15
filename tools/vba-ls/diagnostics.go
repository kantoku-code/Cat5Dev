package main

import (
	"regexp"
	"strings"
)

// memberAccessPattern は "識別子.識別子" を検出する（Unicode 識別子対応）。
// 末尾側は貪欲マッチで識別子の終わりまで取るため \b 不要。先頭側は precededByIdentPart で判定する。
var memberAccessPattern = regexp.MustCompile(`(` + identPattern + `)\.(` + identPattern + `)`)

// BuildDiagnostics は文書全体を走査して、存在しないメンバーへのアクセスを診断として返す。
// メンバー一覧が網羅的（Complete）な型の場合だけ診断する。
func BuildDiagnostics(doc *Document) []Diagnostic {
	var diags []Diagnostic
	curScope := "module"

	for lineIdx, line := range doc.Lines {
		if m := procStartRe.FindStringSubmatch(line); m != nil {
			kind := strings.Title(strings.ToLower(m[1])) //nolint:staticcheck
			curScope = kind + ":" + m[2]
		} else if procEndRe.MatchString(line) {
			curScope = "module"
		}

		// stripped は line とバイト長が等しいので、正規表現のバイト位置は line でもそのまま有効
		stripped := stripCommentAndStrings(line)

		for _, match := range memberAccessPattern.FindAllStringSubmatchIndex(stripped, -1) {
			if precededByIdentPart(stripped, match[2]) {
				continue // 識別子の途中から始まるマッチ（先頭側 \b の代わり）
			}
			varName := stripped[match[2]:match[3]]
			memberName := stripped[match[4]:match[5]]

			sym := doc.Table.FindByName(varName, curScope)
			if sym == nil || sym.TypeName == "" {
				continue
			}
			ti, _, ok := doc.Types.Resolve(sym.TypeRef())
			if !ok || ti.Kind == TypeKindPrimitive || !ti.Complete || ti.Ref.Array {
				continue
			}
			if _, ok := doc.Types.Member(ti.Ref, memberName); ok {
				continue
			}
			diags = append(diags, Diagnostic{
				Range: Range{
					Start: Position{Line: uint32(lineIdx), Character: uint32(byteToUTF16Offset(line, match[4]))},
					End:   Position{Line: uint32(lineIdx), Character: uint32(byteToUTF16Offset(line, match[5]))},
				},
				Severity: DiagnosticSeverityError,
				Source:   "vba-ls",
				Message:  "'" + sym.TypeName + "' に '" + memberName + "' は存在しません",
			})
		}
	}
	return diags
}

// stripCommentAndStrings はコメント（'）と文字列リテラルの中身を空白で置換する。
// 置換はバイト単位で行い、戻り値は line とバイト長が等しい（位置の対応を保つため）。
// 二重引用符と単一引用符はどちらも ASCII なので、UTF-8 の多バイト文字の途中で誤検出することはない。
// "" エスケープは「閉じて直後に開く」扱いになり、引用符自体は残る。
func stripCommentAndStrings(line string) string {
	inStr := false
	result := []byte(line)
	for i, ch := range result {
		if inStr {
			if ch == '"' {
				inStr = false
			} else {
				result[i] = ' '
			}
		} else {
			if ch == '"' {
				inStr = true
			} else if ch == '\'' {
				for j := i; j < len(result); j++ {
					result[j] = ' '
				}
				break
			}
		}
	}
	return string(result)
}
