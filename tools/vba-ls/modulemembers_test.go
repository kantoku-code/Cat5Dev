package main

import (
	"strings"
	"testing"
)

func entryOf(t *testing.T, name string, kind ModuleKind, src string) *ModuleEntry {
	t.Helper()
	lines := splitLines(src)
	return newModuleEntry(name, kind, `c:\ws\modules\`+name+".x", "file:///c%3A/ws/modules/"+name+".x", src, lines, ParseSymbolsLines(src, lines))
}

func namesOf(mi *memberIndex) string {
	var s []string
	for _, m := range mi.list {
		s = append(s, m.Name)
	}
	return strings.Join(s, ",")
}

func TestModuleMembersVisibility(t *testing.T) {
	src := `' モジュールの説明
'======
Dim dimVar As Long
Private privVar As String
Public pubVar As PartDocument
Global globVar As Long
Const C1 = 1
Public Const C2 As Long = 2
Sub NoMod()
End Sub
Private Sub PrivSub()
End Sub
Friend Function FriendFn(ByVal a As Long, Optional b() As String) As clsA
End Function
Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Public Event Changed(ByVal v As Long)
Public Enum EKind
    kA
End Enum
Public Type TInfo
    n As Long
End Type
`
	lfAndCRLF(t, src, func(t *testing.T, src string) {
		e := entryOf(t, "Mod", ModuleStd, src)
		if got, want := namesOf(e.public), "C2,FriendFn,GetTickCount,globVar,NoMod,pubVar"; got != want {
			t.Errorf("public = %s, want %s", got, want)
		}
		if got, want := namesOf(e.all), "C1,C2,dimVar,FriendFn,GetTickCount,globVar,NoMod,PrivSub,privVar,pubVar,Sleep"; got != want {
			t.Errorf("all = %s, want %s", got, want)
		}
		if len(e.Events) != 1 || e.Events[0].Name != "Changed" {
			t.Errorf("Events = %v", e.Events)
		}
		if e.Doc != "モジュールの説明" {
			t.Errorf("Doc = %q", e.Doc)
		}
		i, _ := e.public.lookup("friendfn")
		m := e.public.list[i]
		if m.Kind != MemberKindMethod || m.Signature() != "FriendFn(ByVal a As Long, [b() As String]) As clsA" || e.public.types[i].Name != "clsA" {
			t.Errorf("FriendFn = %+v %q %+v", m, m.Signature(), e.public.types[i])
		}
		if m.Owner.Lib != moduleLib || m.Origin == nil || len(m.Origin.Syms) != 1 || m.Origin.Module != "Mod" {
			t.Errorf("Owner / Origin = %+v %+v", m.Owner, m.Origin)
		}
		i, _ = e.public.lookup("C2")
		if !e.public.list[i].ReadOnly || e.public.list[i].ReturnType != "Long" {
			t.Errorf("C2 = %+v", e.public.list[i])
		}
	})
}

func TestModuleMembersPropertyMerge(t *testing.T) {
	src := `Public Property Get Value() As Long
End Property
' 値を設定する
Public Property Let Value(ByVal v As Long)
End Property
Public Property Let OnlyLet(ByVal v As String)
End Property
Public Property Get Mixed() As Long
End Property
Private Property Let Mixed(ByVal v As Long)
End Property
Private Property Get Hidden() As Long
End Property
Public Property Get Item(ByVal i As Long) As String
End Property
Public Property Set Doc(ByVal d As PartDocument)
End Property
`
	e := entryOf(t, "clsA", ModuleClass, src)
	if got, want := namesOf(e.public), "Doc,Item,Mixed,OnlyLet,Value"; got != want {
		t.Errorf("public = %s, want %s", got, want)
	}
	get := func(name string) (MemberInfo, TypeRef) {
		i, ok := e.all.lookup(name)
		if !ok {
			t.Fatalf("%s がない", name)
		}
		return e.all.list[i], e.all.types[i]
	}
	m, ty := get("Value")
	if m.Kind != MemberKindProperty || m.ReadOnly || m.Signature() != "Value As Long" || len(m.Origin.Syms) != 2 || m.Doc != "値を設定する" || ty.Name != "Long" {
		t.Errorf("Value = %+v %q", m, m.Signature())
	}
	m, ty = get("OnlyLet")
	if m.ReadOnly || m.Signature() != "OnlyLet As String" || ty.Name != "String" {
		t.Errorf("OnlyLet = %+v %q", m, m.Signature())
	}
	m, _ = get("Hidden")
	if !m.ReadOnly || m.Visibility != VisPrivate {
		t.Errorf("Hidden = %+v", m)
	}
	m, _ = get("Item")
	if m.Signature() != "Item(ByVal i As Long) As String" {
		t.Errorf("Item = %q", m.Signature())
	}
	m, ty = get("Doc")
	if ty.Name != "PartDocument" || m.Signature() != "Doc As PartDocument" {
		t.Errorf("Doc = %q %+v", m.Signature(), ty)
	}
	if ti := e.typeInfo(); ti.Kind != TypeKindClass || !ti.Complete || ti.Ref.Lib != "" {
		t.Errorf("class typeInfo = %+v", ti)
	}
}

func TestModuleSurfaceHash(t *testing.T) {
	a := entryOf(t, "m", ModuleStd, "Public Sub A()\n    x = 1\nEnd Sub\nPrivate Sub B()\nEnd Sub\n")
	b := entryOf(t, "m", ModuleStd, "Public Sub A()\n    x = 2\n    y = 3\nEnd Sub\nPrivate Sub B(ByVal n As Long)\nEnd Sub\n")
	c := entryOf(t, "m", ModuleStd, "Public Sub A(ByVal n As Long)\nEnd Sub\n")
	if a.surface != b.surface {
		t.Error("本体と Private の変更では公開面は変わらない")
	}
	if a.surface == c.surface {
		t.Error("Public のシグネチャ変更で公開面が変わる")
	}
	f := entryOf(t, "frmA", ModuleForm, "Public Sub SetData()\nEnd Sub\n")
	if ti := f.typeInfo(); ti.Complete || ti.Base != formBaseType {
		t.Errorf("form typeInfo = %+v", ti)
	}
	// 末尾の補修（repair.go）で直らない構文エラー（ブロックの未終端）
	bad := entryOf(t, "clsBad", ModuleClass, "Public Sub A()\n    Do\nEnd Sub\n")
	if bad.Table.Stats.SyntaxErrors == 0 || bad.typeInfo().Complete {
		t.Errorf("構文エラーのあるクラスは Complete にしない: %+v", bad.Table.Stats)
	}
}
