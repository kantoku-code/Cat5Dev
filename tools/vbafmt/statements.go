package main

import "strings"

// normalizeThenPlacement は "If x > 0 _\n    Then" を "If x > 0 Then" に変換する
// "_" で終わる行の次の実質コード行が "Then" のみなら結合する
func normalizeThenPlacement(lines []string) []string {
	result := make([]string, 0, len(lines))
	i := 0
	for i < len(lines) {
		line := lines[i]
		trimmed := strings.TrimSpace(line)

		// 継続行かどうか確認（コード末尾が _ ）
		if isContinuationLine(trimmed) {
			// 次の実質コード行を探す
			j := i + 1
			for j < len(lines) && strings.TrimSpace(lines[j]) == "" {
				j++
			}
			if j < len(lines) && strings.ToLower(strings.TrimSpace(lines[j])) == "then" {
				// 末尾の " _" を除去して "Then" を結合
				base := strings.TrimRight(trimmed, " \t")
				base = base[:len(base)-1] // "_" を除去
				base = strings.TrimRight(base, " \t")
				result = append(result, base+" Then")
				i = j + 1
				continue
			}
		}

		result = append(result, line)
		i++
	}
	return result
}

// isContinuationLine はトリム済みの行が継続行（末尾が _ ）かどうかを返す
// 文字列内の _ でないかを確認する
func isContinuationLine(trimmed string) bool {
	if !strings.HasSuffix(trimmed, "_") {
		return false
	}
	// セグメント解析でコード末尾が _ かどうか確認
	segs := parseSegments(trimmed)
	if len(segs) == 0 {
		return false
	}
	last := segs[len(segs)-1]
	if last.kind != segCode {
		return false
	}
	return strings.HasSuffix(strings.TrimRight(last.text, " \t"), "_")
}

// splitColonStatements はコロン区切り文を複数行に分割する
// "Dim x: x = 1" → ["    Dim x", "    x = 1"]（元行のインデントを維持）
// ラベル（"Label:" 形式）は分割しない
// 文字列・コメント内のコロンはスキップ
func splitColonStatements(lines []string) []string {
	result := make([]string, 0, len(lines))
	for _, line := range lines {
		result = append(result, splitColonLine(line)...)
	}
	return result
}

func splitColonLine(line string) []string {
	// 先頭インデントを取得
	indent := ""
	for _, ch := range line {
		if ch == ' ' || ch == '\t' {
			indent += string(ch)
		} else {
			break
		}
	}
	trimmed := strings.TrimSpace(line)

	// ラベル判定: "Identifier:" の形式で、コロンの後に実質コードがない
	if isLabel(trimmed) {
		return []string{line}
	}

	// コロンでの分割（文字列・コメント内を除く）
	segs := parseSegments(trimmed)
	var parts []string
	var current strings.Builder

	for _, seg := range segs {
		if seg.kind != segCode {
			current.WriteString(seg.text)
			continue
		}
		// コードセグメント内のコロンで分割
		s := seg.text
		for _, ch := range s {
			if ch == ':' {
				part := strings.TrimSpace(current.String())
				if part != "" {
					parts = append(parts, part)
				}
				current.Reset()
			} else {
				current.WriteRune(ch)
			}
		}
	}
	if tail := strings.TrimSpace(current.String()); tail != "" {
		parts = append(parts, tail)
	}

	if len(parts) <= 1 {
		return []string{line}
	}

	// インデントを付けて返す
	result := make([]string, len(parts))
	for i, p := range parts {
		result[i] = indent + p
	}
	return result
}

// isLabel は "Identifier:" の形式かどうかを返す
// コロンが末尾にあり、その前が識別子（スペースなし）の場合のみラベル
func isLabel(trimmed string) bool {
	if !strings.HasSuffix(trimmed, ":") {
		return false
	}
	base := trimmed[:len(trimmed)-1]
	// スペースを含まない識別子のみラベル
	if strings.ContainsAny(base, " \t=<>+-*/\\^&,(") {
		return false
	}
	// 空でなく、識別子文字のみで構成されているか
	if base == "" {
		return false
	}
	for _, ch := range base {
		if !isIdentPart(ch) {
			return false
		}
	}
	return true
}
