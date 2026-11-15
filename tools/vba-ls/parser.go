package main

import (
	"regexp"
	"strings"

	"github.com/antlr4-go/antlr/v4"
	vbaantlr "github.com/nori/cat5dev/vba-ls/antlr"
)

// reProgIDArg は CreateObject の第1引数（文字列リテラル）から ProgID を取り出す。
var reProgIDArg = regexp.MustCompile(`^"([^"]+)"`)

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
