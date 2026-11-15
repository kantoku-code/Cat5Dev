package main

import (
	"regexp"
	"strings"
)

// dotPattern はカーソル前の "識別子." パターンを検出する。
var dotPattern = regexp.MustCompile(`([` + identPartClass + `.]+)\.$`)

// BuildCompletions はカーソル前テキストを解析して補完候補を返す。
func BuildCompletions(linePrefix string, table *SymbolTable, db *TLBDatabase, scope string) []CompletionItem {
	m := dotPattern.FindStringSubmatch(linePrefix)
	if m == nil {
		return nil
	}
	varName := m[1]

	typeName := resolveChainType(varName, table, db, scope)
	if typeName == "" {
		return nil
	}

	members := db.Members(typeName)
	items := make([]CompletionItem, 0, len(members))
	for _, mem := range members {
		sig := mem.Signature()
		kind := CompletionItemKindProperty
		if mem.Kind == MemberKindMethod {
			kind = CompletionItemKindMethod
		}
		items = append(items, CompletionItem{
			Label:  mem.Name,
			Kind:   kind,
			Detail: sig,
		})
	}
	return items
}

// resolveChainType は "oDoc" や "CATIA.ActiveDocument" のような式の最終型名を返す。
func resolveChainType(expr string, table *SymbolTable, db *TLBDatabase, scope string) string {
	parts := strings.Split(expr, ".")
	// 最初のトークンをシンボルテーブルで解決
	sym := table.FindByName(parts[0], scope)
	if sym == nil || sym.TypeName == "" {
		return ""
	}
	current := sym.TypeName
	// 残りのトークンを DB でチェーン解決
	for _, part := range parts[1:] {
		current = db.GetReturnType(current, part)
		if current == "" {
			return ""
		}
	}
	return current
}

// ResolveHoverMarkdown はホバー対象のシンボル情報を Markdown 文字列で返す。
func ResolveHoverMarkdown(wordAtCursor, prevWord string, table *SymbolTable, db *TLBDatabase, scope string) string {
	if prevWord == "" {
		sym := table.FindByName(wordAtCursor, scope)
		if sym == nil {
			return ""
		}
		return "**" + sym.Name + "** As " + sym.TypeName
	}

	sym := table.FindByName(prevWord, scope)
	if sym == nil || sym.TypeName == "" {
		return ""
	}
	members := db.Members(sym.TypeName)
	lword := strings.ToLower(wordAtCursor)
	for _, m := range members {
		if strings.ToLower(m.Name) == lword {
			return "```vba\n" + m.Signature() + "\n```"
		}
	}
	return ""
}
