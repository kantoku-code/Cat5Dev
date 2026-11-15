package main

// BuildCompletions はカーソル位置（行と UTF-16 位置）の補完候補を返す。
func BuildCompletions(doc *Document, line, utf16Char int) []CompletionItem {
	return BuildCompletionsTrigger(doc, line, utf16Char, "")
}

// BuildCompletionsTrigger は起動文字 trigger（"." / " "。手動・入力による起動なら ""）を考慮して補完候補を返す。
// 空白で起動したときは、型（As / New の後）とキーワード（Exit / End の後など）の文脈だけに応答する
// （空白を打つたびに識別子の一覧が出ないように）。
func BuildCompletionsTrigger(doc *Document, line, utf16Char int, trigger string) []CompletionItem {
	ctx := analyzeCompletion(doc, line, utf16Char)
	if trigger == " " && ctx.Kind != CtxType && ctx.Kind != CtxKeyword {
		return nil
	}
	switch ctx.Kind {
	case CtxMember:
		t, ok := newExprResolver(doc).resolveChain(ctx.Chain, line)
		if !ok {
			return nil
		}
		return memberItems(doc.Types.Members(t))
	case CtxIdentifier:
		return buildIdentifierItems(doc, line, ctx)
	case CtxType:
		return buildTypeItems(doc, ctx)
	case CtxKeyword:
		return keywordItems(ctx.Keywords)
	}
	return nil
}

// memberItems はメンバー一覧を補完候補に変換する。
func memberItems(members []MemberInfo) []CompletionItem {
	if len(members) == 0 {
		return nil
	}
	items := make([]CompletionItem, 0, len(members))
	for _, mem := range members {
		items = append(items, CompletionItem{
			Label:         mem.Name,
			Kind:          completionKindOf(mem.Kind),
			Detail:        mem.Signature(),
			Documentation: plainDoc(mem.Doc),
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
	case MemberKindConst:
		return CompletionItemKindConstant
	}
	return CompletionItemKindProperty
}
