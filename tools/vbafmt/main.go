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
	// 高優先度
	normalizeOperatorSpacing := flag.Bool("normalize-operator-spacing", false, "演算子前後のスペース統一")
	trimTrailingSpace := flag.Bool("trim-trailing-space", true, "行末スペースの除去")
	indentContinuationLines := flag.Bool("indent-continuation-lines", true, "継続行のインデント (+1)")
	maxBlankLines := flag.Int("max-blank-lines", 2, "最大連続空行数 (0=無効)")
	// 中優先度
	normalizeCommaSpacing := flag.Bool("normalize-comma-spacing", false, "コンマ後スペース統一")
	splitColonStatements := flag.Bool("split-colon-statements", false, "コロン区切り文の分割")
	normalizeThenPlacement := flag.Bool("normalize-then-placement", false, "Then の同行強制")
	normalizeCommentSpace := flag.Bool("normalize-comment-space", false, "コメント記号後スペース")
	// 低優先度
	expandTypeSuffixes := flag.Bool("expand-type-suffixes", false, "型接尾辞の展開 (Dim x% → Dim x As Integer)")
	normalizeOnError := flag.Bool("normalize-on-error", false, "On Error スタイル統一")
	flag.Parse()

	opts := Options{
		IndentSize:               *indentSize,
		CapitalizeKeywords:       *capitalizeKeywords,
		FixIndentation:           *fixIndentation,
		LineEndings:              *lineEndings,
		NormalizeOperatorSpacing: *normalizeOperatorSpacing,
		TrimTrailingSpace:        *trimTrailingSpace,
		IndentContinuationLines:  *indentContinuationLines,
		MaxBlankLines:            *maxBlankLines,
		NormalizeCommaSpacing:    *normalizeCommaSpacing,
		SplitColonStatements:     *splitColonStatements,
		NormalizeThenPlacement:   *normalizeThenPlacement,
		NormalizeCommentSpace:    *normalizeCommentSpace,
		ExpandTypeSuffixes:       *expandTypeSuffixes,
		NormalizeOnError:         *normalizeOnError,
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
