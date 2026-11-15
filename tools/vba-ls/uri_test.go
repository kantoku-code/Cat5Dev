package main

import "testing"

func TestPathToURI(t *testing.T) {
	for _, tc := range []struct{ path, want string }{
		{`c:\Work\TK612_プラグ\modules\clsA.cls_utf`, "file:///c%3A/Work/TK612_%E3%83%97%E3%83%A9%E3%82%B0/modules/clsA.cls_utf"},
		{`C:\a b\#x %y\m.bas_utf`, "file:///c%3A/a%20b/%23x%20%25y/m.bas_utf"},
		{`\\Server\share\m.bas_utf`, "file://server/share/m.bas_utf"},
	} {
		if got := pathToURI(tc.path); got != tc.want {
			t.Errorf("pathToURI(%q) = %q, want %q", tc.path, got, tc.want)
		}
	}
}

func TestURIToPathRoundTrip(t *testing.T) {
	for _, p := range []string{
		`c:\Work\TK612_プラグ\modules\clsA.cls_utf`,
		`C:\a b\#x %y\m.bas_utf`,
		`D:\𠮷\フォーム.frm_utf`,
	} {
		got, ok := uriToPath(pathToURI(p))
		if !ok || pathKey(got) != pathKey(p) {
			t.Errorf("往復 %q → %q (%v)", p, got, ok)
		}
	}
	// VSCode 以外の書き方（大文字ドライブ・%3A なし・日本語そのまま）も同じキーになる
	a, ok1 := uriKey("file:///c%3A/Work/TK612_%E3%83%97%E3%83%A9%E3%82%B0/modules/clsA.cls_utf")
	b, ok2 := uriKey("file:///C:/Work/TK612_プラグ/modules/clsA.cls_utf")
	if !ok1 || !ok2 || a != b {
		t.Errorf("同じファイルのキーが一致しない: %q %q", a, b)
	}
	if p, ok := uriToPath("file://server/share/m.bas_utf"); !ok || pathKey(p) != pathKey(`\\server\share\m.bas_utf`) {
		t.Errorf("UNC: %q %v", p, ok)
	}
	for _, u := range []string{"git:/c%3A/a.cls_utf?x", "untitled:Untitled-1", "", "file://"} {
		if _, ok := uriToPath(u); ok {
			t.Errorf("uriToPath(%q) は ok=false のはず", u)
		}
	}
}

func TestModuleIdentity(t *testing.T) {
	for _, tc := range []struct {
		path string
		name string
		kind ModuleKind
		ok   bool
	}{
		{`c:\m\A.CLS_UTF`, "A", ModuleClass, true},
		{`c:\m\Main.bas_utf`, "Main", ModuleStd, true},
		{`/m/frm部品.frm_utf`, "frm部品", ModuleForm, true},
		{`c:\m\a.bas`, "", 0, false},
		{`c:\m\.bas_utf`, "", 0, false},
	} {
		name, kind, ok := moduleIdentity(tc.path)
		if name != tc.name || kind != tc.kind || ok != tc.ok {
			t.Errorf("moduleIdentity(%q) = %q %v %v", tc.path, name, kind, ok)
		}
	}
	if n, k, ok := moduleIdentityOfURI("git:/c%3A/m/clsA.cls_utf?ref"); !ok || n != "clsA" || k != ModuleClass {
		t.Errorf("git URI: %q %v %v", n, k, ok)
	}
	if _, _, ok := moduleIdentityOfURI("file:///test.bas"); ok {
		t.Error("file:///test.bas はモジュールではない")
	}
}
