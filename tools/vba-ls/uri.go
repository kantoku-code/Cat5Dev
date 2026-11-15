package main

import (
	"net/url"
	"path/filepath"
	"strings"
)

// ModuleKind はモジュールの種別（拡張子で決まる）。
type ModuleKind int

const (
	ModuleStd   ModuleKind = iota // .bas_utf
	ModuleClass                   // .cls_utf
	ModuleForm                    // .frm_utf
)

func (k ModuleKind) String() string {
	switch k {
	case ModuleClass:
		return "class"
	case ModuleForm:
		return "form"
	}
	return "std"
}

var moduleExts = map[string]ModuleKind{".bas_utf": ModuleStd, ".cls_utf": ModuleClass, ".frm_utf": ModuleForm}

// moduleIdentity はファイル名からモジュール名（拡張子を除いたファイル名）と種別を返す。拡張子は大文字小文字無視。
// .cls_utf / .frm_utf にはヘッダ（Attribute VB_Name）が無いため、モジュール名＝ファイル名とする。
func moduleIdentity(path string) (name string, kind ModuleKind, ok bool) {
	base := path
	if i := strings.LastIndexAny(base, `/\`); i >= 0 {
		base = base[i+1:]
	}
	dot := strings.LastIndexByte(base, '.')
	if dot <= 0 {
		return "", 0, false
	}
	kind, ok = moduleExts[strings.ToLower(base[dot:])]
	if !ok {
		return "", 0, false
	}
	return base[:dot], kind, true
}

// moduleIdentityOfURI は URI（file: 以外も可）の末尾のファイル名からモジュール名と種別を返す。
func moduleIdentityOfURI(uri string) (string, ModuleKind, bool) {
	if p, ok := uriToPath(uri); ok {
		return moduleIdentity(p)
	}
	u, err := url.Parse(uri)
	if err != nil || u.Path == "" {
		return "", 0, false
	}
	return moduleIdentity(u.Path)
}

// uriToPath は file: URI を OS のパスにする。file: 以外（untitled: / git: など）は ok=false。
// url.Parse は Path の %XX を復号する（%3A → ":"、日本語も復号）。"/c:/x" → "c:\x"。Host があれば UNC。
func uriToPath(uri string) (string, bool) {
	u, err := url.Parse(uri)
	if err != nil || !strings.EqualFold(u.Scheme, "file") {
		return "", false
	}
	p := u.Path
	if p == "" {
		return "", false
	}
	if u.Host != "" && !strings.EqualFold(u.Host, "localhost") {
		return filepath.FromSlash("//" + u.Host + p), true
	}
	if len(p) >= 3 && p[0] == '/' && isASCIILetter(p[1]) && p[2] == ':' {
		p = p[1:]
	}
	return filepath.FromSlash(p), true
}

func isASCIILetter(b byte) bool { return (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') }

// pathToURI は vscode-uri と同じ形の file: URI を作る: ドライブ文字は小文字、":" は %3A、
// [A-Za-z0-9-._~/] 以外は UTF-8 のバイトごとに %XX（大文字）。
// 例: c:\Work\TK612_プラ\a.cls_utf → file:///c%3A/Work/TK612_%E3%83%97%E3%83%A9/a.cls_utf
func pathToURI(path string) string {
	p := strings.ReplaceAll(path, `\`, "/")
	authority := ""
	if strings.HasPrefix(p, "//") {
		// UNC: //server/share/x
		rest := p[2:]
		i := strings.IndexByte(rest, '/')
		if i < 0 {
			authority, p = rest, "/"
		} else {
			authority, p = rest[:i], rest[i:]
		}
	} else {
		if len(p) >= 2 && isASCIILetter(p[0]) && p[1] == ':' {
			p = strings.ToLower(p[:1]) + p[1:]
		}
		if !strings.HasPrefix(p, "/") {
			p = "/" + p
		}
	}
	var sb strings.Builder
	sb.WriteString("file://")
	sb.WriteString(strings.ToLower(authority))
	const hex = "0123456789ABCDEF"
	for i := 0; i < len(p); i++ {
		c := p[i]
		if isASCIILetter(c) || (c >= '0' && c <= '9') || c == '-' || c == '.' || c == '_' || c == '~' || c == '/' {
			sb.WriteByte(c)
			continue
		}
		sb.WriteByte('%')
		sb.WriteByte(hex[c>>4])
		sb.WriteByte(hex[c&15])
	}
	return sb.String()
}

// pathKey はパスの比較用キー。Windows ではパスの大文字小文字を区別しないため小文字化する。
func pathKey(path string) string { return strings.ToLower(filepath.Clean(path)) }

// uriKey は file: URI の比較用キー（file: 以外は ok=false）。
func uriKey(uri string) (string, bool) {
	p, ok := uriToPath(uri)
	if !ok {
		return "", false
	}
	return pathKey(p), true
}
