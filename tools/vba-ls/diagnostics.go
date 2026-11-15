package main

import (
	"regexp"
	"strings"
)

// memberAccessPattern は "識別子.識別子" を検出する。
var memberAccessPattern = regexp.MustCompile(`\b(\w+)\.(\w+)\b`)

// BuildDiagnostics はソース全体を走査して不正なメンバーアクセスを診断として返す。
func BuildDiagnostics(src string, table *SymbolTable, db *TLBDatabase) []Diagnostic {
	var diags []Diagnostic
	lines := strings.Split(src, "\n")
	curScope := "module"

	subRe := regexp.MustCompile(`(?i)^\s*(Sub|Function)\s+(\w+)`)
	endSubRe := regexp.MustCompile(`(?i)^\s*End\s+(Sub|Function)`)

	for lineIdx, line := range lines {
		line = strings.TrimSuffix(line, "\r")
		if m := subRe.FindStringSubmatch(line); m != nil {
			kind := strings.Title(strings.ToLower(m[1])) //nolint:staticcheck
			curScope = kind + ":" + m[2]
		} else if endSubRe.MatchString(line) {
			curScope = "module"
		}

		// stripped は line とバイト長が等しいので、正規表現のバイト位置は line でもそのまま有効
		stripped := stripCommentAndStrings(line)

		for _, match := range memberAccessPattern.FindAllStringSubmatchIndex(stripped, -1) {
			varName := stripped[match[2]:match[3]]
			memberName := stripped[match[4]:match[5]]

			sym := table.FindByName(varName, curScope)
			if sym == nil || sym.TypeName == "" {
				continue
			}
			if db.TypeExists(sym.TypeName) && !hasMember(db, sym.TypeName, memberName) {
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
	}
	return diags
}

func hasMember(db *TLBDatabase, typeName, memberName string) bool {
	return db.HasMember(typeName, memberName)
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
