package main

import (
	"regexp"
	"strings"
)

// dotPattern はカーソル前の "識別子." パターンを検出する。
var dotPattern = regexp.MustCompile(`([` + identPartClass + `.]+)\.$`)

// BuildCompletions はカーソル位置（行と UTF-16 位置）の補完候補を返す。
func BuildCompletions(doc *Document, line, utf16Char int) []CompletionItem {
	m := dotPattern.FindStringSubmatch(doc.lineAt(line, utf16Char))
	if m == nil {
		return nil
	}
	t, ok := resolveChainType(m[1], doc, doc.Table.ScopeAt(line))
	if !ok {
		return nil
	}
	return memberItems(doc.Types.Members(t))
}

// memberItems はメンバー一覧を補完候補に変換する。
func memberItems(members []MemberInfo) []CompletionItem {
	if len(members) == 0 {
		return nil
	}
	items := make([]CompletionItem, 0, len(members))
	for _, mem := range members {
		kind := CompletionItemKindProperty
		if mem.Kind == MemberKindMethod {
			kind = CompletionItemKindMethod
		}
		items = append(items, CompletionItem{
			Label:  mem.Name,
			Kind:   kind,
			Detail: mem.Signature(),
		})
	}
	return items
}

// resolveChainType は "oDoc" や "oDoc.Part" のような式の最終型を返す。
func resolveChainType(expr string, doc *Document, scope string) (TypeRef, bool) {
	parts := strings.Split(expr, ".")
	sym := doc.Table.FindByName(parts[0], scope)
	if sym == nil || sym.TypeName == "" {
		return TypeRef{}, false
	}
	current := sym.TypeRef()
	for _, part := range parts[1:] {
		next, ok := doc.Types.MemberType(current, part)
		if !ok {
			return TypeRef{}, false
		}
		current = next
	}
	return current, true
}

// ResolveHover はカーソル位置のホバー Markdown を返す。対象がなければ ""。
func ResolveHover(doc *Document, line, utf16Char int) string {
	if line < 0 || line >= len(doc.Lines) {
		return ""
	}
	word, prevWord := wordAndPrevWordInLine(doc.Lines[line], utf16Char)
	if word == "" {
		return ""
	}
	scope := doc.Table.ScopeAt(line)
	if prevWord == "" {
		sym := doc.Table.FindByName(word, scope)
		if sym == nil {
			return ""
		}
		return "**" + sym.Name + "** As " + sym.TypeName
	}
	sym := doc.Table.FindByName(prevWord, scope)
	if sym == nil || sym.TypeName == "" {
		return ""
	}
	if m, ok := doc.Types.Member(sym.TypeRef(), word); ok {
		return "```vba\n" + m.Signature() + "\n```"
	}
	return ""
}
