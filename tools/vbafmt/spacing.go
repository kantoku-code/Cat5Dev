package main

import "strings"

// trimTrailingSpace は各行末のスペース・タブを除去する
func trimTrailingSpace(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = strings.TrimRight(line, " \t")
	}
	return result
}

// normalizeCommaSpacing はコンマ後にスペースを1つ付与する（文字列・コメント内を除く）
func normalizeCommaSpacing(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = normalizeCommaLine(line)
	}
	return result
}

func normalizeCommaLine(line string) string {
	segs := parseSegments(line)
	var out strings.Builder
	for _, seg := range segs {
		if seg.kind != segCode {
			out.WriteString(seg.text)
			continue
		}
		// コードセグメント内のコンマ処理
		s := seg.text
		var buf strings.Builder
		for i := 0; i < len(s); i++ {
			buf.WriteByte(s[i])
			if s[i] == ',' {
				// 次の文字がスペース以外なら1スペース追加
				if i+1 < len(s) && s[i+1] != ' ' && s[i+1] != '\t' {
					buf.WriteByte(' ')
				}
			}
		}
		out.WriteString(buf.String())
	}
	return out.String()
}

// normalizeCommentSpace はコメント記号 ' の直後にスペースを付与する
// "'comment" → "' comment"（既にスペースある場合は変更しない）
func normalizeCommentSpace(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = normalizeCommentSpaceLine(line)
	}
	return result
}

func normalizeCommentSpaceLine(line string) string {
	segs := parseSegments(line)
	var out strings.Builder
	for _, seg := range segs {
		if seg.kind != segComment {
			out.WriteString(seg.text)
			continue
		}
		// "'" の直後が非スペースならスペースを挿入
		s := seg.text
		if len(s) >= 2 && s[0] == '\'' && s[1] != ' ' && s[1] != '\'' {
			out.WriteByte('\'')
			out.WriteByte(' ')
			out.WriteString(s[1:])
		} else {
			out.WriteString(s)
		}
	}
	return out.String()
}

// normalizeOperatorSpacing は演算子前後にスペースを付与する（文字列・コメント内を除く）
// 対象: = + - * / \ ^ & < > <= >= <>
// 単項マイナス（直前が演算子/(/行頭）はスキップ
// &H 16進リテラルの & はスキップ
// 複合演算子 <=, >=, <> は2文字で1つの演算子として処理
func normalizeOperatorSpacing(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = normalizeOperatorLine(line)
	}
	return result
}

func normalizeOperatorLine(line string) string {
	segs := parseSegments(line)
	var out strings.Builder
	for _, seg := range segs {
		if seg.kind != segCode {
			out.WriteString(seg.text)
			continue
		}
		out.WriteString(processOperatorSpacing(seg.text))
	}
	return out.String()
}

// isOperatorChar は演算子として処理する文字かどうかを返す
func isOperatorChar(ch byte) bool {
	switch ch {
	case '=', '+', '-', '*', '/', '\\', '^', '&', '<', '>':
		return true
	}
	return false
}

// lastNonSpace は文字列の末尾の非スペース文字を返す（なければ0）
func lastNonSpace(s string) byte {
	for i := len(s) - 1; i >= 0; i-- {
		if s[i] != ' ' && s[i] != '\t' {
			return s[i]
		}
	}
	return 0
}

// isUnaryContext は直前の非スペース文字が単項演算子コンテキストかどうかを返す
func isUnaryContext(ch byte) bool {
	if ch == 0 {
		return true // 行頭
	}
	switch ch {
	case '(', ',', '=', '<', '>', '+', '-', '*', '/', '\\', '^', '&':
		return true
	}
	return false
}

func processOperatorSpacing(code string) string {
	// 1パス: 各文字を処理
	var out strings.Builder
	i := 0
	for i < len(code) {
		ch := code[i]

		if !isOperatorChar(ch) {
			out.WriteByte(ch)
			i++
			continue
		}

		// &H 16進リテラルの & はスキップ
		if ch == '&' && i+1 < len(code) && (code[i+1] == 'H' || code[i+1] == 'h') {
			out.WriteByte(ch)
			i++
			continue
		}

		// 複合演算子チェック: <=, >=, <>
		op := string(ch)
		opLen := 1
		if i+1 < len(code) {
			next := code[i+1]
			if (ch == '<' || ch == '>') && (next == '=' || next == '>') {
				op = string([]byte{ch, next})
				opLen = 2
			} else if ch == '<' && next == '>' {
				op = "<>"
				opLen = 2
			}
		}

		// 単項マイナスの判定（- のみ）
		if ch == '-' && opLen == 1 {
			prev := lastNonSpace(out.String())
			if isUnaryContext(prev) {
				out.WriteByte(ch)
				i++
				continue
			}
		}

		// 左側のスペース確保
		current := out.String()
		if len(current) > 0 && current[len(current)-1] != ' ' && current[len(current)-1] != '\t' {
			out.WriteByte(' ')
		}

		out.WriteString(op)
		i += opLen

		// 右側のスペース確保
		if i < len(code) && code[i] != ' ' && code[i] != '\t' {
			out.WriteByte(' ')
		}
	}
	return out.String()
}
