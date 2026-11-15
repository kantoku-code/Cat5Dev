package main

// BuildCompletions はカーソル位置（行と UTF-16 位置）の補完候補を返す。
// 2-0 では "式." の後のメンバー補完だけを扱う（識別子補完・型補完は 2-2）。
func BuildCompletions(doc *Document, line, utf16Char int) []CompletionItem {
	ctx := analyzeCompletion(doc, line, utf16Char)
	if ctx.Kind != CtxMember {
		return nil
	}
	t, ok := newExprResolver(doc).resolveChain(ctx.Chain, line)
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
		items = append(items, CompletionItem{
			Label:  mem.Name,
			Kind:   completionKindOf(mem.Kind),
			Detail: mem.Signature(),
		})
	}
	return items
}

func completionKindOf(k MemberKind) CompletionItemKind {
	switch k {
	case MemberKindMethod:
		return CompletionItemKindMethod
	case MemberKindField:
		return CompletionItemKindField
	case MemberKindEnumMember:
		return CompletionItemKindEnumMember
	case MemberKindEvent:
		return CompletionItemKindEvent
	}
	return CompletionItemKindProperty
}
