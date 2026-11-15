package main

import (
	"regexp"
	"strings"
	"time"

	"github.com/antlr4-go/antlr/v4"
	vbaantlr "github.com/nori/cat5dev/vba-ls/antlr"
)

// reProgIDArg は CreateObject の第1引数（文字列リテラル）から ProgID を取り出す。
var reProgIDArg = regexp.MustCompile(`^"([^"]+)"`)

// ParseSymbols は VBA ソースを ANTLR4 でパースしてシンボルテーブルを返す。
func ParseSymbols(src string) *SymbolTable { return parseModule(src, splitLines(src), parseAuto) }

// ParseSymbolsLines は分割済みの行（splitLines の結果）を使って ParseSymbols と同じ処理をする。
func ParseSymbolsLines(src string, lines []string) *SymbolTable {
	return parseModule(src, lines, parseAuto)
}

func parseModule(src string, lines []string, mode parseMode) *SymbolTable {
	start := time.Now()
	tree, stats := parseTree(src, mode)
	listener := &symbolListener{
		table: &SymbolTable{},
		scope: "module",
		lines: lines,
	}
	antlr.ParseTreeWalkerDefault.Walk(listener, tree)
	stats.Duration = time.Since(start)
	listener.table.Stats = stats
	return listener.table
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

// symbolListener は ANTLR4 パースツリーを走査してシンボルを収集するリスナー。
type symbolListener struct {
	*vbaantlr.BasevbaListener
	table *SymbolTable
	scope string   // 現在のスコープ名
	lines []string // ソースを "\n" で分割した行（列の UTF-16 変換用）
}

// tokenPos はトークンの 0-based 行と UTF-16 列を返す。
// ANTLR の GetColumn は rune 単位のため、該当行のテキストを使って UTF-16 単位に変換する。
func (l *symbolListener) tokenPos(tok antlr.Token) (line, col int) {
	line = tok.GetLine() - 1
	col = tok.GetColumn()
	if line >= 0 && line < len(l.lines) {
		col = runeToUTF16Offset(l.lines[line], col)
	}
	return line, col
}

// --- スコープ管理 ---

func (l *symbolListener) EnterSubroutineDeclaration(ctx *vbaantlr.SubroutineDeclarationContext) {
	// subroutineName を取得
	if sn := ctx.SubroutineName(); sn != nil {
		l.scope = "Sub:" + sn.GetText()
	}
}

func (l *symbolListener) ExitSubroutineDeclaration(_ *vbaantlr.SubroutineDeclarationContext) {
	l.scope = "module"
}

func (l *symbolListener) EnterFunctionDeclaration(ctx *vbaantlr.FunctionDeclarationContext) {
	if fn := ctx.FunctionName(); fn != nil {
		l.scope = "Function:" + fn.GetText()
	}
}

func (l *symbolListener) ExitFunctionDeclaration(_ *vbaantlr.FunctionDeclarationContext) {
	l.scope = "module"
}

// --- 変数宣言（Dim / Public / Private） ---

func (l *symbolListener) EnterVariableDcl(ctx *vbaantlr.VariableDclContext) {
	// TypedVariableDcl: Dim x%  (型接尾辞付き) — 初期実装では無視
	// UntypedVariableDcl: Dim x As TypeName
	if u := ctx.UntypedVariableDcl(); u != nil {
		ident := u.AmbiguousIdentifier()
		name := ident.GetText()
		typeName := ""
		typeNameLine := 0
		typeNameCol := 0
		hasPos := false
		if ac := u.AsClause(); ac != nil {
			if at := ac.AsType(); at != nil {
				// Dim x As TypeName
				if ts := at.TypeSpec(); ts != nil {
					if te := ts.TypeExpression(); te != nil {
						typeName = strings.TrimSpace(te.GetText())
						typeNameLine, typeNameCol = l.tokenPos(te.GetStart())
						hasPos = true
					}
				}
			} else if ao := ac.AsAutoObject(); ao != nil {
				// Dim x As New TypeName
				if ct := ao.ClassTypeName(); ct != nil {
					typeName = strings.TrimSpace(ct.GetText())
					typeNameLine, typeNameCol = l.tokenPos(ct.GetStart())
					hasPos = true
				}
			}
		}
		if name != "" {
			line, col := l.tokenPos(ident.GetStart())
			l.table.Symbols = append(l.table.Symbols, Symbol{
				Name:           name,
				TypeName:       typeName,
				Scope:          l.scope,
				Line:           line,
				Col:            col,
				TypeNameLine:   typeNameLine,
				TypeNameCol:    typeNameCol,
				HasTypeNamePos: hasPos,
			})
		}
	}
}

// --- Set 代入（Set x = New TypeName） ---

// EnterSetStatement は "Set <変数> = New <型>" / "Set <変数> = CreateObject("ProgID")" から
// 変数の型を登録する。左辺が単純な変数名でない（obj.Prop / arr(i) / With 内の .Member）場合は対象外。
func (l *symbolListener) EnterSetStatement(ctx *vbaantlr.SetStatementContext) {
	lex, rhs := ctx.LExpression(), ctx.Expression()
	if lex == nil || rhs == nil {
		return
	}
	nameTok, name := simpleLExprName(lex)
	if nameTok == nil || name == "" {
		return
	}
	if ne := rhs.NewExpress(); ne != nil {
		// New TypeName / New Lib.TypeName → 最後の名前を型名とし、その位置を型名トークンにする
		if inner := ne.Expression(); inner != nil && inner.LExpression() != nil {
			if typeTok, typeName := lastNameOf(inner.LExpression()); typeTok != nil && typeName != "" {
				l.upsertSymbol(name, nameTok, typeName, typeTok)
			}
		}
		return
	}
	if progID := createObjectProgID(rhs); progID != "" {
		// CreateObject("ProgID") → ProgID の最後のコンポーネントを型名とする（ソース上に型名トークンはない）
		parts := strings.Split(progID, ".")
		if typeName := parts[len(parts)-1]; typeName != "" {
			l.upsertSymbol(name, nameTok, typeName, nil)
		}
	}
}

// identOfName は name 規則から識別子部分（型接尾辞を除く）を返す。
func identOfName(n vbaantlr.INameContext) vbaantlr.IAmbiguousIdentifierContext {
	if n == nil {
		return nil
	}
	if un := n.UntypedName(); un != nil {
		return un.AmbiguousIdentifier()
	}
	if tn := n.TypedName(); tn != nil {
		return tn.AmbiguousIdentifier()
	}
	return nil
}

// simpleLExprName は lExpression が単純な名前のとき、その識別子の先頭トークンと名前を返す。
func simpleLExprName(lex vbaantlr.ILExpressionContext) (antlr.Token, string) {
	sn := lex.SimpleNameExpression()
	if sn == nil {
		return nil, ""
	}
	id := identOfName(sn.Name())
	if id == nil {
		return nil, ""
	}
	return id.GetStart(), id.GetText()
}

// lastNameOf は "Foo" または "Lib.Foo" の形の lExpression から最後の名前とその先頭トークンを返す。
func lastNameOf(lex vbaantlr.ILExpressionContext) (antlr.Token, string) {
	if tok, name := simpleLExprName(lex); tok != nil {
		return tok, name
	}
	if lex.PERIOD() != nil {
		if un := lex.UnrestrictedName(); un != nil {
			return un.GetStart(), un.GetText()
		}
	}
	return nil, ""
}

// createObjectProgID は式が CreateObject("ProgID", ...) の呼び出しなら ProgID を返す。
func createObjectProgID(expr vbaantlr.IExpressionContext) string {
	lex := expr.LExpression()
	if lex == nil || lex.LPAREN() == nil || lex.ArgumentList() == nil {
		return ""
	}
	callee := lex.LExpression()
	if callee == nil {
		return ""
	}
	calleeName := ""
	if callee.SimpleNameExpression() != nil {
		calleeName = callee.GetText()
	} else if un := callee.UnrestrictedName(); un != nil {
		calleeName = un.GetText() // VBA.CreateObject(...)
	}
	if !strings.EqualFold(calleeName, "CreateObject") {
		return ""
	}
	if m := reProgIDArg.FindStringSubmatch(lex.ArgumentList().GetText()); m != nil {
		return m[1]
	}
	return ""
}

// upsertSymbol は Set 代入で型が判明した変数を登録する。
// nameTok は左辺の変数名トークン、typeTok は型名トークン（ソース上にない場合は nil）。
// 同じスコープに既存のシンボルがあれば、型が未設定の場合のみ型を補う（宣言位置は変えない）。
func (l *symbolListener) upsertSymbol(name string, nameTok antlr.Token, typeName string, typeTok antlr.Token) {
	for i := range l.table.Symbols {
		s := &l.table.Symbols[i]
		if strings.EqualFold(s.Name, name) && s.Scope == l.scope {
			if s.TypeName == "" {
				s.TypeName = typeName
				s.TypeNameLine, s.TypeNameCol, s.HasTypeNamePos = 0, 0, false
				if typeTok != nil {
					s.TypeNameLine, s.TypeNameCol = l.tokenPos(typeTok)
					s.HasTypeNamePos = true
				}
			}
			return
		}
	}
	sym := Symbol{Name: name, TypeName: typeName, Scope: l.scope}
	sym.Line, sym.Col = l.tokenPos(nameTok)
	if typeTok != nil {
		sym.TypeNameLine, sym.TypeNameCol = l.tokenPos(typeTok)
		sym.HasTypeNamePos = true
	}
	l.table.Symbols = append(l.table.Symbols, sym)
}
