package main

import "strings"

// LSP の Position.character は UTF-16 コードユニット単位で数える（positionEncoding 未交渉時の既定）。
// 一方 Go の文字列はバイト単位、ANTLR の GetColumn はコードポイント（rune）単位で数えるため、
// サーバー内外の位置をやり取りする際は必ずこのファイルの関数で変換すること。

// lineText は content の 0-based line 行目を、末尾の \r を除いて返す。
// 行が存在しなければ ok=false を返す。
func lineText(content string, line int) (text string, ok bool) {
	if line < 0 {
		return "", false
	}
	lines := strings.Split(content, "\n")
	if line >= len(lines) {
		return "", false
	}
	return strings.TrimSuffix(lines[line], "\r"), true
}

// utf16RuneLen は rune を UTF-16 で表したときのコードユニット数を返す。
func utf16RuneLen(r rune) int {
	if r >= 0x10000 {
		return 2 // サロゲートペア
	}
	return 1
}

// utf16ToByteOffset は行内の UTF-16 位置をバイト位置へ変換する。
// 行長を超える位置は行末に、サロゲートペアの途中を指す位置はその文字の先頭に切り下げる。
func utf16ToByteOffset(line string, char int) int {
	if char <= 0 {
		return 0
	}
	u := 0
	for i, r := range line {
		n := utf16RuneLen(r)
		if u+n > char {
			return i
		}
		u += n
	}
	return len(line)
}

// byteToUTF16Offset は行内のバイト位置を UTF-16 位置へ変換する。
// 行長を超える位置は行末として扱う。
func byteToUTF16Offset(line string, off int) int {
	if off > len(line) {
		off = len(line)
	}
	u := 0
	for i, r := range line {
		if i >= off {
			break
		}
		u += utf16RuneLen(r)
	}
	return u
}

// runeToUTF16Offset は行内の rune 位置（ANTLR の GetColumn の値）を UTF-16 位置へ変換する。
// 行長を超える分は 1 rune = 1 コードユニットとして加算する。
func runeToUTF16Offset(line string, col int) int {
	u, c := 0, 0
	for _, r := range line {
		if c >= col {
			break
		}
		u += utf16RuneLen(r)
		c++
	}
	if col > c {
		u += col - c
	}
	return u
}

// utf16Len は文字列の UTF-16 コードユニット数を返す。
func utf16Len(s string) int {
	return byteToUTF16Offset(s, len(s))
}
