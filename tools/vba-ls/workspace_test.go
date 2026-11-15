package main

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// wsFixture は補完・ホバー・診断・定義ジャンプのテストで使う基本のワークスペース。
var wsFixture = map[string]string{
	"clsA.cls_utf": `'/_/_/_/_/_/_/_/_/
' テスト用クラス A
'/_/_/_/_/_/_/_/_/
Option Explicit
Private m_n As Long
Public Name As String

' 部品を返す
Public Property Get Doc() As PartDocument
End Property

Public Property Get Value() As Long
End Property

Public Property Let Value(ByVal v As Long)
    m_n = v
End Property

Friend Sub FriendSub()
End Sub

Private Sub PrivSub()
End Sub

Public Function Clone() As clsA
End Function

Private Sub Class_Initialize()
End Sub

Public Event Changed()
`,
	"Module1.bas_utf": `Option Explicit
Public g As Long
Public Const C As Long = 1
Private Const P = 2

Public Sub Pub()
End Sub

Private Sub Priv()
End Sub

' A を作る
Public Function NewA() As clsA
End Function

Public Enum EKind
    kA
    kB
End Enum

Public Type TInfo
    doc As PartDocument
End Type
`,
	"frmA.frm_utf": `Option Explicit
Private m_x As Long

Public Sub SetData()
End Sub

Private Sub UserForm_Initialize()
End Sub
`,
}

// newTestWorkspace は files（ファイル名 → 本文）を t.TempDir()/modules に書き、
// テスト用 TLB 付きの resolver と、同期的に IndexAll した Workspace を返す。
func newTestWorkspace(t *testing.T, files map[string]string) (*Workspace, *CompositeResolver, string) {
	t.Helper()
	root := t.TempDir()
	dir := filepath.Join(root, "modules")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	for name, body := range files {
		if err := os.WriteFile(filepath.Join(dir, name), []byte(body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	ws := NewWorkspace()
	ws.SetRoots([]string{root})
	ws.IndexAll(context.Background())
	types := NewDefaultResolver(newTestTLBDatabase(t), ws.TypesProvider(), ws.ModulesProvider())
	return ws, types, dir
}

// wsDocAt は dir/name を srcWithCaret（"|" がカーソル）で開いたものとして Document を作る。
func wsDocAt(t *testing.T, types *CompositeResolver, dir, name, srcWithCaret string) (*Document, int, int) {
	t.Helper()
	src, line, char := caretPos(t, srcWithCaret)
	uri := pathToURI(filepath.Join(dir, name))
	return NewDocument(uri, src, ParseSymbols(src), types), line, char
}

type wsCompletionCase struct {
	name    string
	file    string // 開いているファイル（ワークスペース内）
	src     string
	want    []string
	notWant []string
}

func runWSCompletionCases(t *testing.T, types *CompositeResolver, dir string, cases []wsCompletionCase) {
	t.Helper()
	for _, tc := range cases {
		tc := tc
		t.Run(tc.name, func(t *testing.T) {
			lfAndCRLF(t, withFixture(tc.file, tc.src), func(t *testing.T, src string) {
				doc, line, char := wsDocAt(t, types, dir, tc.file, src)
				items := BuildCompletions(doc, line, char)
				if tc.want == nil {
					if items != nil {
						t.Fatalf("候補なしのはずが %v", labelsOf(items))
					}
					return
				}
				labels := labelsOf(items)
				for _, w := range tc.want {
					if !labels[w] {
						t.Errorf("候補に %s がない: %v", w, labels)
					}
				}
				for _, n := range tc.notWant {
					if labels[n] {
						t.Errorf("候補に %s が含まれている: %v", n, labels)
					}
				}
			})
		})
	}
}

func TestWorkspaceIndex(t *testing.T) {
	files := map[string]string{"temp.txt": "x"}
	for k, v := range wsFixture {
		files[k] = v
	}
	ws, _, dir := newTestWorkspace(t, files)
	// modules の外（root 直下・サブフォルダ）は索引しない
	if err := os.WriteFile(filepath.Join(filepath.Dir(dir), "Outside.bas_utf"), []byte("Public Sub X()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(filepath.Join(dir, "sub"), 0o755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "sub", "Nested.bas_utf"), []byte("Public Sub Y()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	ws2 := NewWorkspace()
	ws2.SetRoots([]string{filepath.Dir(dir)})
	st := ws2.IndexAll(context.Background())
	if st.Parsed != 3 || st.ByKind != [3]int{1, 1, 1} {
		t.Errorf("IndexStats = %+v", st)
	}
	select {
	case <-ws2.Ready():
	default:
		t.Error("IndexAll 後に Ready が close されていない")
	}
	var names []string
	for _, e := range ws2.Modules() {
		names = append(names, e.Name+":"+e.Kind.String())
	}
	if got := strings.Join(names, ","); got != "clsA:class,frmA:form,Module1:std" {
		t.Errorf("Modules = %s", got)
	}
	if !ws.Contains(filepath.Join(dir, "clsB.cls_utf")) || ws.Contains(filepath.Join(dir, "sub", "x.cls_utf")) || ws.Contains(filepath.Join(dir, "a.txt")) {
		t.Error("Contains の判定が違う")
	}
	// modules が無ければフォルダ直下
	if got := indexDirs([]string{dir}); len(got) != 1 || pathKey(got[0]) != pathKey(dir) {
		t.Errorf("indexDirs = %v", got)
	}
	var tl []string
	for _, ti := range ws.TypesProvider().(TypeLister).Types() {
		tl = append(tl, ti.Ref.Name)
	}
	if got := strings.Join(tl, ","); got != "clsA,EKind,frmA,TInfo" {
		t.Errorf("Types = %s", got)
	}
	var pm []string
	for _, m := range ws.ProjectMembers() {
		pm = append(pm, m.Name)
	}
	if got := strings.Join(pm, ","); got != "C,g,NewA,Pub" {
		t.Errorf("ProjectMembers = %s", got)
	}
}

func TestWorkspaceCompletion(t *testing.T) {
	files := map[string]string{
		"Part.cls_utf":         "Public Sub Mine()\nEnd Sub\n",
		"cls部品.cls_utf":        "Public Function 名前() As String\nEnd Function\n",
		"clsWriteOnly.cls_utf": "Public Property Let W(ByVal v As Long)\nEnd Property\n",
	}
	for k, v := range wsFixture {
		files[k] = v
	}
	_, types, dir := newTestWorkspace(t, files)
	runWSCompletionCases(t, types, dir, []wsCompletionCase{
		{"公開メンバーのみ", "Module1.bas_utf", "Sub T()\n    Dim o As New clsA\n    o.|\nEnd Sub\n",
			[]string{"Name", "Doc", "Value", "FriendSub", "Clone"}, []string{"m_n", "PrivSub", "Class_Initialize", "Changed"}},
		{"CATIA 型に続く", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Doc.|\nEnd Sub\n", partDocMembers, nil},
		{"Set New", "frmA.frm_utf", "Sub T()\n    Set x = New clsA\n    x.|\nEnd Sub\n", []string{"Name"}, []string{"m_n"}},
		{"チェーン", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Clone.Clone.Doc.|\nEnd Sub\n", []string{"Part"}, nil},
		{"大文字小文字無視", "Module1.bas_utf", "Sub T()\n    Dim o As CLSA\n    o.|\nEnd Sub\n", []string{"Name"}, nil},
		{"WS のクラスが TLB より優先", "Module1.bas_utf", "Sub T()\n    Dim p As Part\n    p.|\nEnd Sub\n", []string{"Mine"}, []string{"Bodies"}},
		{"ローカル UDT が WS より優先", "Other.bas_utf", "Private Type clsA\n    z As Long\nEnd Type\nSub T()\n    Dim v As clsA\n    v.|\nEnd Sub\n", []string{"z"}, []string{"Name"}},
		{"日本語", "Module1.bas_utf", "Sub T()\n    Dim p As New cls部品\n    p.|\nEnd Sub\n", []string{"名前"}, nil},
		{"Public UDT（他モジュール）", "frmA.frm_utf", "Sub T()\n    Dim v As TInfo\n    v.doc.|\nEnd Sub\n", partDocMembers, nil},
		// 自モジュールの Enum 型の変数と同じ扱い（Enum のメンバーを返す）
		{"Public Enum（他モジュール）の型", "frmA.frm_utf", "Sub T()\n    Dim k As EKind\n    k.|\nEnd Sub\n", []string{"kA", "kB"}, nil},
		{"Let だけの Property", "Module1.bas_utf", "Sub T()\n    Dim o As clsWriteOnly\n    o.|\nEnd Sub\n", []string{"W"}, nil},
	})
	doc, line, char := wsDocAt(t, types, dir, "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.|\nEnd Sub\n")
	n := 0
	for _, it := range BuildCompletions(doc, line, char) {
		if it.Label == "Value" {
			n++
			if it.Detail != "Value As Long" {
				t.Errorf("Value の Detail = %q", it.Detail)
			}
		}
	}
	if n != 1 {
		t.Errorf("Property Get/Let は 1 件に統合（%d 件）", n)
	}
}

// 索引の前（空のワークスペース）でも panic せず、候補なしになる。
func TestWorkspaceCompletionBeforeIndex(t *testing.T) {
	ws := NewWorkspace()
	types := NewDefaultResolver(newTestTLBDatabase(t), ws.TypesProvider(), ws.ModulesProvider())
	if items := completeAt(t, types, "Sub T()\n    Dim o As clsA\n    o.|\nEnd Sub\n"); items != nil {
		t.Errorf("索引前は候補なし: %v", labelsOf(items))
	}
	var nilWS *Workspace
	nilWS.OpenBuffer("file:///c%3A/a.cls_utf")
	nilWS.PutBuffer(nil)
	nilWS.CloseBuffer("file:///c%3A/a.cls_utf")
	nilWS.FilesChanged(nil)
}

func TestWorkspaceMeAndModuleAccess(t *testing.T) {
	_, types, dir := newTestWorkspace(t, wsFixture)
	runWSCompletionCases(t, types, dir, []wsCompletionCase{
		{"クラスの Me.", "clsA.cls_utf", "Sub T()\n    Me.|\nEnd Sub\n", []string{"m_n", "PrivSub", "Class_Initialize", "Name"}, []string{"Changed"}},
		{"フォームの Me.", "frmA.frm_utf", "Sub T()\n    Me.|\nEnd Sub\n", []string{"m_x", "SetData", "Hide", "Show", "Caption"}, nil},
		{"With Me", "clsA.cls_utf", "Sub T()\n    With Me\n        .|\n    End With\nEnd Sub\n", []string{"m_n"}, nil},
		{"標準モジュールの Me. は候補なし", "Module1.bas_utf", "Sub T()\n    Me.|\nEnd Sub\n", nil, nil},
		{"他モジュールから Module1.", "clsA.cls_utf", "Sub T()\n    Module1.|\nEnd Sub\n", []string{"Pub", "NewA", "g", "C"}, []string{"Priv", "P"}},
		{"自モジュールの Module1.", "Module1.bas_utf", "Sub T()\n    Module1.|\nEnd Sub\n", []string{"Priv", "Pub"}, nil},
		{"Module1.NewA().", "frmA.frm_utf", "Sub T()\n    Module1.NewA().|\nEnd Sub\n", []string{"Name", "Doc"}, []string{"m_n"}},
		{"非修飾の NewA().", "frmA.frm_utf", "Sub T()\n    NewA().|\nEnd Sub\n", []string{"Name", "Doc"}, []string{"m_n"}},
		{"非修飾の NewA.", "frmA.frm_utf", "Sub T()\n    NewA.|\nEnd Sub\n", []string{"Name"}, nil},
		{"フォームの既定インスタンス", "clsA.cls_utf", "Sub T()\n    frmA.|\nEnd Sub\n", []string{"SetData", "Hide"}, []string{"m_x"}},
		{"他モジュールの Enum", "clsA.cls_utf", "Sub T()\n    x = EKind.|\nEnd Sub\n", []string{"kA", "kB"}, nil},
		{"クラスに既定インスタンスは無い", "Module1.bas_utf", "Sub T()\n    clsA.|\nEnd Sub\n", nil, nil},
		{"ローカル変数がモジュール名より優先", "clsA.cls_utf", "Sub T()\n    Dim Module1 As clsA\n    Module1.|\nEnd Sub\n", []string{"Name"}, []string{"Pub"}},
	})
}

func TestWorkspaceDiagnostics(t *testing.T) {
	// 末尾の補修（repair.go）で直らない構文エラー（ブロックの未終端）
	files := map[string]string{"clsBad.cls_utf": "Public Sub A()\n    Do\nEnd Sub\nPublic Sub B()\nEnd Sub\n"}
	for k, v := range wsFixture {
		files[k] = v
	}
	_, types, dir := newTestWorkspace(t, files)
	for _, tc := range []struct {
		name, file, src string
		want            []string // メッセージ
	}{
		{"クラスの存在しないメンバー", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Bogus\nEnd Sub\n", []string{"'clsA' に 'Bogus' は存在しません"}},
		{"他モジュールから Private", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.PrivSub\nEnd Sub\n", []string{"'clsA' に 'PrivSub' は存在しません"}},
		{"標準モジュール（#module を出さない）", "clsA.cls_utf", "Sub T()\n    Module1.Bogus\nEnd Sub\n", []string{"'Module1' に 'Bogus' は存在しません"}},
		{"標準モジュールの Private（他から）", "clsA.cls_utf", "Sub T()\n    Module1.Priv\nEnd Sub\n", []string{"'Module1' に 'Priv' は存在しません"}},
		{"正しいアクセスは 0 件", "frmA.frm_utf", "Sub T()\n    Dim o As New clsA\n    o.Name = Module1.NewA().Name\n    Module1.Pub\n    o.Value = Module1.C\n    o.FriendSub\nEnd Sub\n", nil},
		{"フォームの Me.コントロールは 0 件", "frmA.frm_utf", "Sub T()\n    Me.lstX.Clear\n    Me.Hide\n    frmA.lstY.AddItem 1\nEnd Sub\n", nil},
		{"クラスの Me.Bogus は 1 件", "clsA.cls_utf", "Sub T()\n    Me.Bogus\n    Me.PrivSub\nEnd Sub\n", []string{"'clsA' に 'Bogus' は存在しません"}},
		{"構文エラーのあるクラスは 0 件", "Module1.bas_utf", "Sub T()\n    Dim b As clsBad\n    b.Bogus\nEnd Sub\n", nil},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc, _, _ := wsDocAt(t, types, dir, tc.file, withFixture(tc.file, tc.src)+"|")
			var got []string
			for _, d := range BuildDiagnostics(doc) {
				got = append(got, d.Message)
				if d.Severity != DiagnosticSeverityError {
					t.Errorf("ワークスペースの型の診断は Error のはず: %+v", d)
				}
			}
			if strings.Join(got, "\n") != strings.Join(tc.want, "\n") {
				t.Errorf("診断 = %q, want %q", got, tc.want)
			}
		})
	}
}

// WS のクラス型はセマンティックトークンで type になる。
func TestWorkspaceSemanticType(t *testing.T) {
	_, types, dir := newTestWorkspace(t, wsFixture)
	src := "Dim o As clsA\n"
	doc, _, _ := wsDocAt(t, types, dir, "Module1.bas_utf", src+"|")
	toks := decodeSemanticTokens(t, BuildSemanticTokens(doc))
	found := false
	for _, tk := range toks {
		if tk.line == 0 && tk.col == u16("Dim o As ") && tk.tokenType == 0 {
			found = true
		}
	}
	if !found {
		t.Errorf("clsA が type トークンになっていない: %+v", toks)
	}
}

// withFixture は開くファイルが基本フィクスチャにあれば、その本文の後ろに src を足す
// （エディタのバッファはファイル全体なので、Me. や自モジュールのメンバーはこの本文から求まる）。
func withFixture(file, src string) string {
	if body, ok := wsFixture[file]; ok {
		return body + "\n" + src
	}
	return src
}

// 既知の差: クラス内では同じクラスの別インスタンス経由でも Private が候補に出る（VBA ではコンパイルエラー）。
// 誤った診断は出ない方向の緩さなので許容する。
func TestWorkspaceSelfClassPrivateIsLoose(t *testing.T) {
	_, types, dir := newTestWorkspace(t, wsFixture)
	doc, line, char := wsDocAt(t, types, dir, "clsA.cls_utf", withFixture("clsA.cls_utf", "Sub T()\n    Dim o As clsA\n    o.|\nEnd Sub\n"))
	if !labelsOf(BuildCompletions(doc, line, char))["m_n"] {
		t.Error("クラス内では自クラス型の Private も見える（現仕様）")
	}
}

func TestWorkspaceHover(t *testing.T) {
	_, types, dir := newTestWorkspace(t, wsFixture)
	for _, tc := range []struct {
		name, file, src, want string
	}{
		{"Property Get とドキュメント", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Do|c\nEnd Sub\n",
			"```vba\nPublic Property Get Doc() As PartDocument\n```\n\n部品を返す"},
		{"Property Get / Let", "Module1.bas_utf", "Sub T()\n    Dim o As clsA\n    o.Val|ue = 1\nEnd Sub\n",
			"```vba\nPublic Property Get Value() As Long\nPublic Property Let Value(ByVal v As Long)\n```"},
		{"Module1.Pub", "clsA.cls_utf", "Sub T()\n    Module1.Pu|b\nEnd Sub\n", "```vba\nPublic Sub Pub()\n```"},
		{"非修飾の NewA", "frmA.frm_utf", "Sub T()\n    Set o = New|A()\nEnd Sub\n", "```vba\nPublic Function NewA() As clsA\n```\n\nA を作る"},
		{"変数", "frmA.frm_utf", "Sub T()\n    x = Module1.|g\nEnd Sub\n", "```vba\nPublic g As Long\n```"},
		{"As clsA", "Module1.bas_utf", "Sub T()\n    Dim o As cls|A\nEnd Sub\n", "```vba\nClass clsA\n```\n\nテスト用クラス A"},
		{"モジュール名", "clsA.cls_utf", "Sub T()\n    Mod|ule1.Pub\nEnd Sub\n", "```vba\nModule Module1\n```"},
		{"フォーム", "clsA.cls_utf", "Sub T()\n    frm|A.Show\nEnd Sub\n", "```vba\nUserForm frmA\n```"},
		{"他モジュールの Public Enum", "clsA.cls_utf", "Sub T()\n    Dim k As EKi|nd\nEnd Sub\n", "```vba\nPublic Enum EKind\n```"},
		{"TLB の型名は対象外", "clsA.cls_utf", "Sub T()\n    Dim d As PartDocu|ment\nEnd Sub\n", ""},
	} {
		t.Run(tc.name, func(t *testing.T) {
			src, line, char := caretPos(t, withFixture(tc.file, tc.src))
			doc := NewDocument(pathToURI(filepath.Join(dir, tc.file)), src, ParseSymbols(src), types)
			if got := ResolveHover(doc, line, char); got != tc.want {
				t.Errorf("hover = %q\nwant %q", got, tc.want)
			}
		})
	}
}

// プロジェクトの非修飾の名前（標準モジュールの Public・モジュール名・フォーム名）の列挙と、New で作れる型。
func TestWorkspaceProjectListers(t *testing.T) {
	_, types, _ := newTestWorkspace(t, wsFixture)
	got := map[string]bool{}
	for _, m := range types.ProjectMembers() {
		got[m.Name] = true
		if m.Owner.Name != "Module1" {
			t.Errorf("%s の所属 = %v", m.Name, m.Owner)
		}
	}
	for _, w := range []string{"g", "C", "Pub", "NewA"} {
		if !got[w] {
			t.Errorf("ProjectMembers に %s がない: %v", w, got)
		}
	}
	if got["Priv"] || got["P"] || got["SetData"] {
		t.Errorf("Private・フォームのメンバーが含まれている: %v", got)
	}
	var mods []string
	for _, ti := range types.ProjectModules() {
		mods = append(mods, ti.Ref.Name)
	}
	if strings.Join(mods, ",") != "frmA,Module1" {
		t.Errorf("ProjectModules = %v", mods)
	}
	creatable := map[string]bool{}
	for _, ti := range types.Types() {
		if ti.Creatable {
			creatable[ti.Ref.Name] = true
		}
	}
	if !creatable["clsA"] || !creatable["frmA"] || creatable["TInfo"] || creatable["EKind"] || creatable["PartDocument"] {
		t.Errorf("Creatable = %v", creatable)
	}
}
