package main

import (
	"time"

	"github.com/antlr4-go/antlr/v4"
	vbaantlr "github.com/nori/cat5dev/vba-ls/antlr"
)

// ParseSymbols は VBA ソースを ANTLR4 でパースしてシンボルテーブルを返す。
func ParseSymbols(src string) *SymbolTable { return parseModule(src, splitLines(src), parseAuto) }

// ParseSymbolsLines は分割済みの行（splitLines の結果）を使って ParseSymbols と同じ処理をする。
func ParseSymbolsLines(src string, lines []string) *SymbolTable {
	return parseModule(src, lines, parseAuto)
}

func parseModule(src string, lines []string, mode parseMode) *SymbolTable {
	start := time.Now()
	tree, stats := parseTree(src, mode)
	l := newSymbolListener(lines)
	antlr.ParseTreeWalkerDefault.Walk(l, tree)
	stats.Duration = time.Since(start)
	stats.SyntaxErrors += l.grammarErrors // 文法では表していない制約の違反（For の ", i" など）
	l.table.Stats = stats
	l.finish()
	return l.table
}

// --- 二段階パース（SLL → LL） ---
//
// まず高速な SLL 予測でパースし、構文エラーが出たら最初から LL 予測でやり直す。
// 述語（precedence 以外）を持たない文法では、SLL が構文エラーなしで成功すれば LL と同じ木になる。
// 標準の BailErrorStrategy は使えない（Go ランタイムではエラーが SetError(nil) で消えて続行し、
// さらに ReportError が "unknown recognition error type" を stdout に出して LSP の通信路を壊す）ため、
// panic で抜ける自作の ErrorStrategy と recover を使う。

// sllBailout は SLL 段の構文エラーで投げる panic 値。line は失敗したトークンの行（1-based、計測ログ用）。
type sllBailout struct{ line int }

// bailStrategy は最初の構文エラーで sllBailout を panic する ErrorStrategy。
// DefaultErrorStrategy を埋め込み、非公開メソッド reset(Parser) も満たす。何も出力しない。
type bailStrategy struct{ *antlr.DefaultErrorStrategy }

func bailAt(p antlr.Parser) sllBailout {
	if tok := p.GetCurrentToken(); tok != nil {
		return sllBailout{line: tok.GetLine()}
	}
	return sllBailout{}
}

func (b *bailStrategy) ReportError(p antlr.Parser, _ antlr.RecognitionException) { panic(bailAt(p)) }
func (b *bailStrategy) Recover(p antlr.Parser, _ antlr.RecognitionException)     { panic(bailAt(p)) }
func (b *bailStrategy) RecoverInline(p antlr.Parser) antlr.Token                 { panic(bailAt(p)) }
func (b *bailStrategy) Sync(antlr.Parser)                                        {}

// errorCounter は構文エラーの数を数える ErrorListener（何も出力しない）。
type errorCounter struct {
	*antlr.DefaultErrorListener
	n int
}

func (e *errorCounter) SyntaxError(antlr.Recognizer, interface{}, int, int, string, antlr.RecognitionException) {
	e.n++
}

// trySLL は parse を実行し、sllBailout の panic を捕まえたら ok=false を返す。それ以外の panic は再送出する。
// bailLine は失敗したトークンの行（1-based）。
func trySLL(parse func() vbaantlr.IStartRuleContext) (tree vbaantlr.IStartRuleContext, ok bool, bailLine int) {
	defer func() {
		if r := recover(); r != nil {
			b, bail := r.(sllBailout)
			if !bail {
				panic(r)
			}
			tree, ok, bailLine = nil, false, b.line
		}
	}()
	return parse(), true, 0
}

type parseMode int

const (
	parseAuto   parseMode = iota // SLL → 失敗したら LL
	parseLLOnly                  // LL のみ（同等性テスト・計測用。従来の動作）
)

func parseTree(src string, mode parseMode) (vbaantlr.IStartRuleContext, ParseStats) {
	sllLine := 0
	lexer := vbaantlr.NewvbaLexer(antlr.NewInputStream(src))
	lexer.RemoveErrorListeners()
	tokens := antlr.NewCommonTokenStream(lexer, antlr.TokenDefaultChannel)
	tokens.Fill() // 字句解析は 1 回だけ行い、LL 段でも再利用する
	if mode == parseAuto {
		p := vbaantlr.NewvbaParser(tokens)
		p.RemoveErrorListeners()
		p.GetInterpreter().SetPredictionMode(antlr.PredictionModeSLL)
		p.SetErrorHandler(&bailStrategy{antlr.NewDefaultErrorStrategy()})
		t, ok, line := trySLL(p.StartRule)
		if ok {
			return t, ParseStats{}
		}
		sllLine = line
	}
	tokens.Seek(0)
	// SynErr などの状態を引き継がないよう、パーサは新しく作る（DFA は静的に共有される）
	p := vbaantlr.NewvbaParser(tokens)
	p.RemoveErrorListeners()
	cnt := &errorCounter{DefaultErrorListener: antlr.NewDefaultErrorListener()}
	p.AddErrorListener(cnt)
	p.GetInterpreter().SetPredictionMode(antlr.PredictionModeLL)
	t := p.StartRule()
	return t, ParseStats{UsedLL: true, SyntaxErrors: cnt.n, SLLErrorLine: sllLine}
}
