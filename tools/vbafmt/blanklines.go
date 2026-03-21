package main

import "strings"

// normalizeBlankLines は連続する空行を max 行に圧縮し、
// プロシージャ間（End Sub/Function/Property の次に Sub/Function/Property が来る場合）
// に空行1行を保証する
func normalizeBlankLines(lines []string, max int) []string {
	if max <= 0 {
		return lines
	}

	// Step 1: 連続空行を max 行に圧縮
	result := make([]string, 0, len(lines))
	blankCount := 0
	for _, line := range lines {
		if strings.TrimSpace(line) == "" {
			blankCount++
			if blankCount <= max {
				result = append(result, line)
			}
		} else {
			blankCount = 0
			result = append(result, line)
		}
	}

	// Step 2: プロシージャ間の空行保証
	// "End Sub/Function/Property" の直後（空行なし）に "Sub/Function/Property/Public/Private..." が来る場合
	final := make([]string, 0, len(result))
	for i := 0; i < len(result); i++ {
		final = append(final, result[i])
		if isProcedureEnder(result[i]) {
			// 次の非空行を探す
			j := i + 1
			for j < len(result) && strings.TrimSpace(result[j]) == "" {
				j++
			}
			// 空行がなく、次がプロシージャ開始なら空行を挿入
			if j == i+1 && j < len(result) && isProcedureStarter(result[j]) {
				final = append(final, "")
			}
		}
	}

	return final
}

// isProcedureEnder は行が "End Sub/Function/Property" かどうかを返す
func isProcedureEnder(line string) bool {
	trimmed := strings.TrimSpace(line)
	lower := strings.ToLower(trimmed)
	return lower == "end sub" || lower == "end function" || strings.HasPrefix(lower, "end property")
}

// isProcedureStarter は行がプロシージャ開始かどうかを返す
func isProcedureStarter(line string) bool {
	trimmed := strings.TrimSpace(line)
	lower := strings.ToLower(trimmed)

	// 修飾子を除去
	prefixes := []string{"public ", "private ", "friend ", "static ", "global "}
	for _, p := range prefixes {
		if strings.HasPrefix(lower, p) {
			lower = strings.TrimPrefix(lower, p)
			break
		}
	}

	return strings.HasPrefix(lower, "sub ") ||
		strings.HasPrefix(lower, "function ") ||
		strings.HasPrefix(lower, "property ")
}
