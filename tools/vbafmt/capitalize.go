package main

import (
	"strings"
	"unicode"
)

// capitalizeKeywords は各行のキーワードを正規ケースに変換する
func capitalizeKeywords(lines []string) []string {
	result := make([]string, len(lines))
	for i, line := range lines {
		result[i] = capitalizeLine(line)
	}
	return result
}

// capitalizeLine は1行のキーワードを変換する
func capitalizeLine(line string) string {
	var out strings.Builder
	inString := false
	i := 0
	for i < len(line) {
		ch := line[i]

		// 文字列リテラル内
		if inString {
			out.WriteByte(ch)
			if ch == '"' {
				// "" エスケープ
				if i+1 < len(line) && line[i+1] == '"' {
					out.WriteByte(line[i+1])
					i += 2
					continue
				}
				inString = false
			}
			i++
			continue
		}

		// 文字列開始
		if ch == '"' {
			inString = true
			out.WriteByte(ch)
			i++
			continue
		}

		// コメント: ' 以降はそのまま出力
		if ch == '\'' {
			out.WriteString(line[i:])
			return out.String()
		}

		// 識別子 (キーワード候補)
		if isIdentStart(rune(ch)) {
			j := i + 1
			for j < len(line) && isIdentPart(rune(line[j])) {
				j++
			}
			word := line[i:j]
			lower := strings.ToLower(word)

			if canonical, ok := Keywords[lower]; ok {
				// 次のトークンと合わせて複合チェック
				rest := line[j:]
				nextWord, nextEnd := peekNextWord(rest)
				compound := lower + " " + strings.ToLower(nextWord)
				switch compound {
				case "end if", "end for", "end with", "end select",
					"end function", "end sub", "end property",
					"end type", "end enum", "end class":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "select case":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "option explicit", "option base", "option compare":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "for each":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "property get", "property let", "property set":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "exit sub", "exit function", "exit for", "exit do", "exit while":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "do while", "do until":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				case "loop while", "loop until":
					out.WriteString(canonical + " " + capitalize(nextWord))
					i = j + nextEnd
					continue
				}
				out.WriteString(canonical)
			} else {
				out.WriteString(word)
			}
			i = j
			continue
		}

		out.WriteByte(ch)
		i++
	}
	return out.String()
}

// peekNextWord は文字列の先頭の空白をスキップして最初の単語と、
// その単語が終わった位置 (空白含む) を返す
func peekNextWord(s string) (word string, consumed int) {
	i := 0
	for i < len(s) && (s[i] == ' ' || s[i] == '\t') {
		i++
	}
	start := i
	for i < len(s) && isIdentPart(rune(s[i])) {
		i++
	}
	return s[start:i], i
}

func capitalize(word string) string {
	if word == "" {
		return ""
	}
	lower := strings.ToLower(word)
	if canonical, ok := Keywords[lower]; ok {
		return canonical
	}
	return word
}

func isIdentStart(r rune) bool {
	return unicode.IsLetter(r) || r == '_'
}

func isIdentPart(r rune) bool {
	return unicode.IsLetter(r) || unicode.IsDigit(r) || r == '_'
}
