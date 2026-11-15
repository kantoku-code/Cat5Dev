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
			lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
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
		{"ローカル UDT が WS より優先", "Module1.bas_utf", "Private Type clsA\n    z As Long\nEnd Type\nSub T()\n    Dim v As clsA\n    v.|\nEnd Sub\n", []string{"z"}, []string{"Name"}},
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
