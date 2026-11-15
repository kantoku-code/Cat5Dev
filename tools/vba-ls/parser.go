package main

import (
	"regexp"
	"strings"

	"github.com/antlr4-go/antlr/v4"
	vbaantlr "github.com/nori/cat5dev/vba-ls/antlr"
)

var reCreateObject = regexp.MustCompile(`(?i)CreateObject\s*\(\s*"([^"]+)"\s*\)`)

// ParseSymbols は VBA ソースを ANTLR4 でパースしてシンボルテーブルを返す。
func ParseSymbols(src string) *SymbolTable {
	is := antlr.NewInputStream(src)
	lexer := vbaantlr.NewvbaLexer(is)
	lexer.RemoveErrorListeners()
	stream := antlr.NewCommonTokenStream(lexer, antlr.TokenDefaultChannel)
	parser := vbaantlr.NewvbaParser(stream)
	parser.RemoveErrorListeners()
	tree := parser.StartRule()

	listener := &symbolListener{
		table: &SymbolTable{},
		scope: "module",
		lines: strings.Split(src, "\n"),
	}
	antlr.ParseTreeWalkerDefault.Walk(listener, tree)
	return listener.table
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
		if ac := u.AsClause(); ac != nil {
			if at := ac.AsType(); at != nil {
				if ts := at.TypeSpec(); ts != nil {
					if te := ts.TypeExpression(); te != nil {
						typeName = strings.TrimSpace(te.GetText())
						typeNameLine, typeNameCol = l.tokenPos(te.GetStart())
					}
				}
			}
		}
		if name != "" {
			line, col := l.tokenPos(ident.GetStart())
			l.table.Symbols = append(l.table.Symbols, Symbol{
				Name:         name,
				TypeName:     typeName,
				Scope:        l.scope,
				Line:         line,
				Col:          col,
				TypeNameLine: typeNameLine,
				TypeNameCol:  typeNameCol,
			})
		}
	}
}

// --- Set 代入（Set x = New TypeName） ---

func (l *symbolListener) EnterSetStatement(ctx *vbaantlr.SetStatementContext) {
	// "Set <lhs> = <rhs>"
	// SetStatement: SET ws implicitCallStmt_S EQ ws? valueStmt
	// lhs からシンボル名を取得するのは複雑なため、テキストを解析
	text := ctx.GetText()
	// 簡易パース: "Set<name>=New<type>" または "Set<name>=<type>"
	lower := strings.ToLower(text)
	if !strings.HasPrefix(lower, "set") {
		return
	}
	inner := text[3:]
	eqIdx := strings.Index(inner, "=")
	if eqIdx < 0 {
		return
	}
	lhs := strings.TrimSpace(inner[:eqIdx])
	rhs := strings.TrimSpace(inner[eqIdx+1:])

	rhsLow := strings.ToLower(rhs)
	if strings.HasPrefix(rhsLow, "new") {
		// New TypeName
		typeName := strings.TrimSpace(rhs[3:])
		if idx := strings.IndexAny(typeName, ".("); idx >= 0 {
			typeName = typeName[:idx]
		}
		if lhs != "" && typeName != "" {
			l.upsertSymbol(lhs, typeName, ctx.GetStart())
		}
	} else if m := reCreateObject.FindStringSubmatch(rhs); m != nil {
		// CreateObject("ProgID") → ProgID の最後のコンポーネントを型名とする
		progID := m[1]
		parts := strings.Split(progID, ".")
		typeName := parts[len(parts)-1]
		if lhs != "" && typeName != "" {
			l.upsertSymbol(lhs, typeName, ctx.GetStart())
		}
	}
}

// upsertSymbol は Set 代入で型が判明した変数を登録する。start は Set 文の先頭トークン。
func (l *symbolListener) upsertSymbol(name, typeName string, start antlr.Token) {
	for i := range l.table.Symbols {
		if strings.EqualFold(l.table.Symbols[i].Name, name) &&
			l.table.Symbols[i].Scope == l.scope {
			if l.table.Symbols[i].TypeName == "" {
				l.table.Symbols[i].TypeName = typeName
			}
			return
		}
	}
	line, col := l.tokenPos(start)
	l.table.Symbols = append(l.table.Symbols, Symbol{
		Name:     name,
		TypeName: typeName,
		Scope:    l.scope,
		Line:     line,
		Col:      col,
	})
}
