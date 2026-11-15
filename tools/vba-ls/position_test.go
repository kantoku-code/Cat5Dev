package main

import (
	"strings"
	"testing"
)

func TestLineText(t *testing.T) {
	tests := []struct {
		name    string
		content string
		line    int
		want    string
		wantOK  bool
	}{
		{"LF の 1 行目", "a\nbc\n", 1, "bc", true},
		{"CRLF は末尾の \\r を除去", "a\r\nbc\r\n", 1, "bc", true},
		{"最終行（改行なし）", "a\r\nbc", 1, "bc", true},
		{"末尾改行の後の空行", "a\n", 1, "", true},
		{"範囲外の行", "a\nb", 2, "", false},
		{"負の行", "a", -1, "", false},
		{"行中の \\r は残す", "a\rb\r\n", 0, "a\rb", true},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got, ok := lineText(tt.content, tt.line)
			if got != tt.want || ok != tt.wantOK {
				t.Errorf("lineText(%q, %d) = (%q, %v), want (%q, %v)", tt.content, tt.line, got, ok, tt.want, tt.wantOK)
			}
		})
	}
}

func TestUTF16ToByteOffset(t *testing.T) {
	tests := []struct {
		name string
		line string
		char int
		want int
	}{
		{"ASCII 先頭", "doc.", 0, 0},
		{"ASCII 途中", "doc.", 3, 3},
		{"ASCII 行末", "doc.", 4, 4},
		{"ASCII 行末超過は行末", "doc.", 99, 4},
		{"負の位置は 0", "doc.", -1, 0},
		{"空行", "", 3, 0},
		// "あ" は UTF-16 で 1、UTF-8 で 3 バイト
		{"日本語の前", `"あ": doc.`, 1, 1},
		{"日本語の後", `"あ": doc.`, 2, 4},
		{"日本語を含む行の行末", `"あ": doc.`, 9, 11},
		{"日本語を含む行の行末超過", `"あ": doc.`, 10, 11},
		// "a😀b": a=1byte/1unit, 😀=4byte/2unit, b=1byte/1unit
		{"サロゲート 0", "a😀b", 0, 0},
		{"サロゲート 1", "a😀b", 1, 1},
		{"サロゲート 2（ペアの途中は切り下げ）", "a😀b", 2, 1},
		{"サロゲート 3", "a😀b", 3, 5},
		{"サロゲート 4", "a😀b", 4, 6},
		{"サロゲート 5（行末超過）", "a😀b", 5, 6},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := utf16ToByteOffset(tt.line, tt.char); got != tt.want {
				t.Errorf("utf16ToByteOffset(%q, %d) = %d, want %d", tt.line, tt.char, got, tt.want)
			}
		})
	}
}

func TestByteToUTF16Offset(t *testing.T) {
	tests := []struct {
		name string
		line string
		off  int
		want int
	}{
		{"ASCII", "doc.Name", 4, 4},
		{"日本語の後", `x = "ああ": doc.Name`, 14, 10},
		{"日本語を含む行の行末", `x = "ああ": doc.Name`, 22, 18},
		{"行末超過は行末", `x = "ああ": doc.Name`, 99, 18},
		{"サロゲートの後", "a😀b", 5, 3},
		{"サロゲートの行末", "a😀b", 6, 4},
		{"0", "a😀b", 0, 0},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := byteToUTF16Offset(tt.line, tt.off); got != tt.want {
				t.Errorf("byteToUTF16Offset(%q, %d) = %d, want %d", tt.line, tt.off, got, tt.want)
			}
		})
	}
}

func TestRuneToUTF16Offset(t *testing.T) {
	tests := []struct {
		name string
		line string
		col  int
		want int
	}{
		{"ASCII", "Dim doc As PartDocument", 4, 4},
		{"日本語は 1 rune = 1 unit", `x = "あ": Dim doc`, 13, 13},
		{"サロゲートは 1 rune = 2 unit", `x = "😀": Dim doc`, 13, 14},
		{"行末", "a😀b", 3, 4},
		{"行末超過分はそのまま加算", "a😀b", 5, 6},
		{"0", "a😀b", 0, 0},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			if got := runeToUTF16Offset(tt.line, tt.col); got != tt.want {
				t.Errorf("runeToUTF16Offset(%q, %d) = %d, want %d", tt.line, tt.col, got, tt.want)
			}
		})
	}
}

func TestUTF16Len(t *testing.T) {
	tests := []struct {
		s    string
		want int
	}{
		{"", 0},
		{"doc", 3},
		{"部品", 2},
		{"a😀b", 4},
	}
	for _, tt := range tests {
		if got := utf16Len(tt.s); got != tt.want {
			t.Errorf("utf16Len(%q) = %d, want %d", tt.s, got, tt.want)
		}
	}
}

func TestStripCommentAndStringsKeepsByteLength(t *testing.T) {
	sp := func(n int) string { return strings.Repeat(" ", n) }
	tests := []struct {
		name string
		line string
		want string
	}{
		{"文字列内の日本語はバイト数分の空白", `x = "ああ": doc.Bogus`, `x = "` + sp(6) + `": doc.Bogus`},
		{"文字列内のサロゲートペア", `x = "😀": doc.Zzz`, `x = "` + sp(4) + `": doc.Zzz`},
		// "" エスケープは「閉じて直後に開く」扱いで引用符が残る（既存挙動の維持）
		{"\"\" エスケープとコメント", `x = "a""b": doc.Name ' コメント`, `x = " "" ": doc.Name ` + sp(len(`' コメント`))},
		{"文字列外の全角空白は残す", "　doc.Bad ' あ", "　doc.Bad " + sp(len("' あ"))},
		{"文字列内の ' はコメントにしない", `x = "it's": doc.Name`, `x = "` + sp(4) + `": doc.Name`},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			got := stripCommentAndStrings(tt.line)
			if len(got) != len(tt.line) {
				t.Errorf("バイト長 = %d, want %d", len(got), len(tt.line))
			}
			if got != tt.want {
				t.Errorf("stripCommentAndStrings(%q) = %q, want %q", tt.line, got, tt.want)
			}
		})
	}
}
