package main

import (
	"reflect"
	"testing"
)

// procByKey は ScopeKey でプロシージャを返す。
func procByKey(t *testing.T, table *SymbolTable, key string) *ProcInfo {
	t.Helper()
	for _, p := range table.Procs {
		if p.ScopeKey() == key {
			return p
		}
	}
	t.Fatalf("プロシージャ %q がない: %+v", key, table.Procs)
	return nil
}

const symbolsSrc = `Option Explicit
'****************************
' 部品ドキュメント
'****************************
Private m_doc As PartDocument
Public g_count As Long ' 件数
Dim m_parts(3) As Part
Dim m_names() As String
Dim s$, n%
Private WithEvents m_app As Application
Public Const MAX_N As Long = 10
Const NAME_PREFIX$ = "P_"
Private Const FLAG = True

Type PointInfo
    X As Double
    Coords(2) As Double
    Name As String
    Doc As PartDocument
End Type

Private Enum Color
    Red = 1
    Green
End Enum

Public Event Changed(ByVal oldValue As Long, newValue As Long)
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)

' 処理を実行する
' 2 行目
Public Function Run(ByVal doc As PartDocument, Optional n As Long = 1, _
        Optional ByRef names As Variant, ParamArray rest()) As Boolean
    Static calls As Long
    Dim p As Part
End Function

Friend Function Name$()
    Dim local1 As String
End Function

Private Sub NoArgs()
End Sub

Function Untyped()
End Function

Function Arr() As String()
End Function

Property Get Value() As Long
    Dim v As Long
End Property

Property Let Value(ByVal newValue As Long)
End Property

Public Property Set Target(ByVal obj As PartDocument)
End Property

Static Sub Counter(a%, b() As Long)
    With m_doc.Part
        .Update
        With .Bodies
            x = .Count
        End With
    End With
End Sub
`

func TestSymbolModel(t *testing.T) {
	lfAndCRLF(t, symbolsSrc, func(t *testing.T, src string) {
		table := ParseSymbols(src)
		if table.Stats.SyntaxErrors != 0 {
			t.Fatalf("構文エラー: %+v", table.Stats)
		}

		t.Run("モジュール変数", func(t *testing.T) {
			for _, w := range []struct {
				name, typeName string
				vis            Visibility
				isArray        bool
			}{
				{"m_doc", "PartDocument", VisPrivate, false},
				{"g_count", "Long", VisPublic, false},
				{"m_parts", "Part", VisPrivate, true}, // As 句は arrayClause の下
				{"m_names", "String", VisPrivate, true},
				{"s", "String", VisPrivate, false}, // 型接尾辞
				{"n", "Integer", VisPrivate, false},
				{"m_app", "Application", VisPrivate, false},
			} {
				s := findSym(t, table, w.name, "module")
				if s.Kind != SymVariable || s.TypeName != w.typeName || s.Visibility != w.vis || s.IsArray != w.isArray {
					t.Errorf("%s = %+v, want 型 %s 公開範囲 %v 配列 %v", w.name, s, w.typeName, w.vis, w.isArray)
				}
			}
			if s := findSym(t, table, "m_app", "module"); !s.WithEvents {
				t.Error("WithEvents が記録されていない")
			}
			if s := findSym(t, table, "m_doc", "module"); s.Doc != "部品ドキュメント" {
				t.Errorf("バナーのドキュメント = %q", s.Doc)
			}
			if s := findSym(t, table, "g_count", "module"); s.Doc != "件数" {
				t.Errorf("行末コメント = %q", s.Doc)
			}
		})

		t.Run("定数", func(t *testing.T) {
			for _, w := range []struct {
				name, typeName, value string
				vis                   Visibility
			}{
				{"MAX_N", "Long", "10", VisPublic},
				{"NAME_PREFIX", "String", `"P_"`, VisPrivate},
				{"FLAG", "", "True", VisPrivate},
			} {
				s := findSym(t, table, w.name, "module")
				if s.Kind != SymConst || s.TypeName != w.typeName || s.Value != w.value || s.Visibility != w.vis {
					t.Errorf("%s = %+v, want 型 %q 値 %q", w.name, s, w.typeName, w.value)
				}
			}
		})

		t.Run("UDTとEnum", func(t *testing.T) {
			u := table.UDT("pointinfo")
			if u == nil || len(u.Members) != 4 || u.Visibility != VisPublic {
				t.Fatalf("UDT = %+v", u)
			}
			want := []struct {
				name, typ string
				arr       bool
			}{{"X", "Double", false}, {"Coords", "Double", true}, {"Name", "String", false}, {"Doc", "PartDocument", false}}
			for i, w := range want {
				m := u.Members[i]
				if m.Name != w.name || m.TypeName != w.typ || m.IsArray != w.arr || m.Kind != SymUDTMember || m.Parent != "PointInfo" {
					t.Errorf("UDT メンバー[%d] = %+v, want %+v", i, m, w)
				}
			}
			if s := table.FindByName("X", "module"); s != nil {
				t.Errorf("UDT のメンバーがモジュールスコープで見える: %+v", s)
			}
			e := table.Enum("Color")
			if e == nil || e.Visibility != VisPrivate || len(e.Members) != 2 || e.Members[0].Value != "1" || e.Members[1].Name != "Green" {
				t.Fatalf("Enum = %+v", e)
			}
			if s := table.FindByName("Green", "module"); s == nil || s.Kind != SymEnumMember || s.Parent != "Color" {
				t.Errorf("Enum のメンバー = %+v", s)
			}
		})

		t.Run("Event と Declare", func(t *testing.T) {
			ev := procByKey(t, table, "Event:Changed")
			if ev.HasBody || len(ev.Params) != 2 || !ev.Params[0].ByVal || ev.Params[1].Type.Name != "Long" {
				t.Errorf("Event = %+v", ev)
			}
			gt := procByKey(t, table, "Declare:GetTickCount")
			if gt.Kind != ProcDeclareFunction || gt.Visibility != VisPrivate || gt.ReturnType.Name != "Long" || gt.HasBody {
				t.Errorf("Declare Function = %+v", gt)
			}
			sl := procByKey(t, table, "Declare:Sleep")
			if sl.Kind != ProcDeclareSub || sl.Visibility != VisPublic || len(sl.Params) != 1 || sl.Params[0].Name != "ms" {
				t.Errorf("Declare Sub = %+v", sl)
			}
			if s := table.FindByName("ms", "Declare:Sleep"); s != nil {
				t.Errorf("Declare の引数はシンボルにしない: %+v", s)
			}
		})

		t.Run("Function の引数・戻り値・範囲", func(t *testing.T) {
			p := procByKey(t, table, "Function:Run")
			if p.Visibility != VisPublic || p.ReturnType != (TypeRef{Name: "Boolean"}) || p.Doc != "処理を実行する\n2 行目" {
				t.Errorf("Run = %+v", p)
			}
			want := []ParamInfo{
				{Name: "doc", Type: TypeRef{Name: "PartDocument"}, ByVal: true},
				{Name: "n", Type: TypeRef{Name: "Long"}, Optional: true, Default: "1"},
				{Name: "names", Type: TypeRef{Name: "Variant"}, Optional: true},
				{Name: "rest", Type: TypeRef{Name: "Variant", Array: true}, ParamArray: true},
			}
			if len(p.Params) != len(want) {
				t.Fatalf("引数 = %+v", p.Params)
			}
			for i, w := range want {
				got := p.Params[i]
				got.Line, got.Col = 0, 0
				if !reflect.DeepEqual(got, w) {
					t.Errorf("引数[%d] = %+v, want %+v", i, got, w)
				}
			}
			startLine := lineOf(t, src, "Public Function Run(")
			if p.StartLine != startLine || p.EndLine != startLine+4 {
				t.Errorf("範囲 = %d-%d, want %d-%d（継続行付きシグネチャ）", p.StartLine, p.EndLine, startLine, startLine+4)
			}
			doc := findSym(t, table, "doc", "Function:Run")
			if doc.Kind != SymParam || doc.Visibility != VisLocal || doc.Owner != p || !doc.HasTypeNamePos {
				t.Errorf("引数のシンボル = %+v", doc)
			}
			calls := findSym(t, table, "calls", "Function:Run")
			if !calls.Static || calls.Visibility != VisLocal {
				t.Errorf("Static ローカル = %+v", calls)
			}
			if s := table.Lookup("rest", startLine+1); s == nil || s.Kind != SymParam {
				t.Errorf("継続行の ParamArray が見つからない: %+v", s)
			}
		})

		t.Run("戻り値の型", func(t *testing.T) {
			for _, w := range []struct {
				key  string
				want TypeRef
				vis  Visibility
			}{
				{"Function:Name", TypeRef{Name: "String"}, VisFriend}, // Function Foo$() のスコープは接尾辞なし
				{"Function:Untyped", TypeRef{Name: "Variant"}, VisPublic},
				{"Function:Arr", TypeRef{Name: "String", Array: true}, VisPublic},
				{"PropertyGet:Value", TypeRef{Name: "Long"}, VisPublic},
			} {
				p := procByKey(t, table, w.key)
				if p.ReturnType != w.want || p.Visibility != w.vis {
					t.Errorf("%s: 戻り値 %+v 公開範囲 %v, want %+v %v", w.key, p.ReturnType, p.Visibility, w.want, w.vis)
				}
			}
			findSym(t, table, "local1", "Function:Name")
			findSym(t, table, "v", "PropertyGet:Value")
			findSym(t, table, "newValue", "PropertyLet:Value")
			findSym(t, table, "obj", "PropertySet:Target")
			if p := procByKey(t, table, "Sub:NoArgs"); p.Visibility != VisPrivate || len(p.Params) != 0 {
				t.Errorf("NoArgs = %+v", p)
			}
			c := procByKey(t, table, "Sub:Counter")
			if !c.Static || len(c.Params) != 2 || c.Params[0].Type.Name != "Integer" || !c.Params[1].Type.Array || c.Params[1].Type.Name != "Long" {
				t.Errorf("Counter = %+v", c)
			}
		})

		t.Run("Lookup の優先順", func(t *testing.T) {
			s := table.Lookup("Value", 0)
			if s == nil || s.Decl == nil || s.Decl.Kind != ProcPropertyGet {
				t.Errorf("モジュールの Value は Property Get を優先: %+v", s)
			}
		})

		t.Run("With", func(t *testing.T) {
			if len(table.Withs) != 2 {
				t.Fatalf("With = %+v", table.Withs)
			}
			outer, inner := table.Withs[0], table.Withs[1]
			if outer.Expr != "m_doc.Part" || inner.Expr != ".Bodies" {
				t.Errorf("With の式 = %q / %q", outer.Expr, inner.Expr)
			}
			if w := table.WithAt(inner.StartLine); w == nil || w.StartLine != outer.StartLine {
				t.Errorf("With 行自身は外側の With: %+v", w)
			}
			if w := table.WithAt(inner.StartLine + 1); w == nil || w.StartLine != inner.StartLine {
				t.Errorf("内側の With の中: %+v", w)
			}
			if w := table.WithAt(outer.StartLine); w != nil {
				t.Errorf("外側の With 行は With の外: %+v", w)
			}
		})
	})
}

// lineOf は src 中で prefix を含む最初の行番号（0-based）を返す。
func lineOf(t *testing.T, src, prefix string) int {
	t.Helper()
	for i, l := range splitLines(src) {
		if len(l) >= len(prefix) && containsStr(l, prefix) {
			return i
		}
	}
	t.Fatalf("%q がない", prefix)
	return -1
}

func containsStr(s, sub string) bool { return len(sub) == 0 || indexOf(s, sub) >= 0 }

func indexOf(s, sub string) int {
	for i := 0; i+len(sub) <= len(s); i++ {
		if s[i:i+len(sub)] == sub {
			return i
		}
	}
	return -1
}

func TestDocComments(t *testing.T) {
	lines := splitLines("'=====\n' 見出し\n'-----\n\n' 本文1\nRem 本文2\nSub A()\n")
	if got := docAbove(lines, 6); got != "本文1\n本文2" {
		t.Errorf("docAbove = %q（空行で打ち切り）", got)
	}
	if got := docAbove(lines, 3); got != "見出し" {
		t.Errorf("装飾行は捨てる: %q", got)
	}
	if got := trailingComment(`x = "a'b" ' 説明 `); got != "説明" {
		t.Errorf("trailingComment = %q", got)
	}
	if got := trailingComment(`' 行全体`); got != "" {
		t.Errorf("行全体のコメントは行末コメントではない: %q", got)
	}
}

// Property の中は Property のスコープになる（修正前は module 扱い）。
func TestPropertyScope(t *testing.T) {
	src := "Private m_doc As Part\nProperty Get Doc() As PartDocument\n    Dim d As PartDocument\n    d.\nEnd Property\n"
	table := ParseSymbols(src)
	if got := table.ScopeAt(3); got != "PropertyGet:Doc" {
		t.Errorf("ScopeAt = %q", got)
	}
	findSym(t, table, "d", "PropertyGet:Doc")
}

// 構文エラーがあっても、正規表現の範囲でスコープを補正する。
func TestRegexScopeFallback(t *testing.T) {
	src := "Sub A()\n    Dim x As Part\n    If Then (\nEnd Sub\nProperty Get B() As Long\n    Dim y As Long\nEnd Property\n"
	table := ParseSymbols(src)
	if table.Stats.SyntaxErrors == 0 {
		t.Skip("構文エラーにならなかった")
	}
	for _, w := range []struct {
		line  int
		scope string
	}{{1, "Sub:A"}, {3, "Sub:A"}, {5, "PropertyGet:B"}} {
		if got := table.ScopeAt(w.line); got != w.scope {
			t.Errorf("ScopeAt(%d) = %q, want %q", w.line, got, w.scope)
		}
	}
	if s := table.Lookup("x", 2); s == nil || s.Scope != "Sub:A" {
		t.Errorf("x = %+v", s)
	}
}

// 型名は修飾子と分けて正規化し、型名トークンの位置は最後の名前にする。
func TestTypeNameNormalization(t *testing.T) {
	src := "Sub M()\n" +
		"    Dim a As Scripting.Dictionary\n" +
		"    Dim b As New Scripting.Dictionary\n" +
		"    Set c = CreateObject(\"Scripting.Dictionary\")\n" +
		"    Set x = CreateObject(\"Excel.Application\")\n" +
		"    Dim i As INFITF.Application\n" +
		"    Dim l As [Long]\n" +
		"    Dim f As String * 10\n" +
		"End Sub\n"
	lfAndCRLF(t, src, func(t *testing.T, src string) {
		table := ParseSymbols(src)
		for _, w := range []struct {
			name, typeName, lib string
			asNew, hasPos       bool
			col                 int
		}{
			{"a", "Dictionary", "Scripting", false, true, u16("    Dim a As Scripting.")},
			{"b", "Dictionary", "Scripting", true, true, u16("    Dim b As New Scripting.")},
			{"c", "Dictionary", "Scripting", false, false, 0},
			{"x", "Application", "Excel", false, false, 0},
			{"i", "Application", "INFITF", false, true, u16("    Dim i As INFITF.")},
			{"l", "Long", "", false, true, u16("    Dim l As ")},
			{"f", "String", "", false, true, u16("    Dim f As ")},
		} {
			s := findSym(t, table, w.name, "Sub:M")
			if s.TypeName != w.typeName || s.TypeLib != w.lib || s.AsNew != w.asNew || s.HasTypeNamePos != w.hasPos || (w.hasPos && s.TypeNameCol != w.col) {
				t.Errorf("%s = %+v, want %s.%s asNew=%v pos=%v col=%d", w.name, s, w.lib, w.typeName, w.asNew, w.hasPos, w.col)
			}
		}
	})
}
