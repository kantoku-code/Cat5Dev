package main

import (
	"flag"
	"fmt"
	"io"
	"os"
)

func main() {
	indentSize := flag.Int("indent-size", 4, "インデントサイズ")
	capitalizeKeywords := flag.Bool("capitalize-keywords", true, "キーワードの大文字化")
	fixIndentation := flag.Bool("fix-indentation", true, "インデント修正")
	lineEndings := flag.String("line-endings", "CRLF", "改行コード (CRLF または LF)")
	check := flag.Bool("check", false, "フォーマット差分があれば exit 1 (CI用)")
	flag.Parse()

	opts := Options{
		IndentSize:         *indentSize,
		CapitalizeKeywords: *capitalizeKeywords,
		FixIndentation:     *fixIndentation,
		LineEndings:        *lineEndings,
	}

	var input []byte
	var err error

	if flag.NArg() > 0 {
		// ファイルから読込
		input, err = os.ReadFile(flag.Arg(0))
		if err != nil {
			fmt.Fprintf(os.Stderr, "vbafmt: ファイル読込エラー: %v\n", err)
			os.Exit(1)
		}
	} else {
		// stdin から読込
		input, err = io.ReadAll(os.Stdin)
		if err != nil {
			fmt.Fprintf(os.Stderr, "vbafmt: stdin 読込エラー: %v\n", err)
			os.Exit(1)
		}
	}

	formatted := Format(string(input), opts)

	if *check {
		if formatted != string(input) {
			fmt.Fprintln(os.Stderr, "vbafmt: フォーマットが必要なファイルがあります")
			os.Exit(1)
		}
		os.Exit(0)
	}

	fmt.Fprint(os.Stdout, formatted)
}
