package main

import (
	"testing"
)

// stubDB は TLBDatabase のスタブ実装。
type stubDB struct {
	// typeName → members
	data map[string][]MemberInfo
}

func (s *stubDB) Members(typeName string) []MemberInfo {
	return s.data[typeName]
}

func (s *stubDB) GetReturnType(typeName, memberName string) string {
	for _, m := range s.data[typeName] {
		if m.Name == memberName {
			return m.ReturnType
		}
	}
	return ""
}

// stubTLB は *TLBDatabase の代わりに使える薄いラッパー。
// TLBDatabase はインタフェースではないため、テスト用にインタフェースを定義。
type memberResolver interface {
	Members(typeName string) []MemberInfo
	GetReturnType(typeName, memberName string) string
}

// BuildCompletionsFromResolver はテスト用の補完ビルダー。
func BuildCompletionsFromResolver(linePrefix string, table *SymbolTable, db memberResolver, scope string) []CompletionItem {
	m := dotPattern.FindStringSubmatch(linePrefix)
	if m == nil {
		return nil
	}
	varName := m[1]
	typeName := resolveChainTypeWithDB(varName, table, db, scope)
	if typeName == "" {
		return nil
	}
	members := db.Members(typeName)
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

func resolveChainTypeWithDB(expr string, table *SymbolTable, db memberResolver, scope string) string {
	parts := splitDot(expr)
	sym := table.FindByName(parts[0], scope)
	if sym == nil || sym.TypeName == "" {
		return ""
	}
	current := sym.TypeName
	for _, part := range parts[1:] {
		current = db.GetReturnType(current, part)
		if current == "" {
			return ""
		}
	}
	return current
}

func splitDot(s string) []string {
	result := []string{}
	cur := ""
	for _, c := range s {
		if c == '.' {
			result = append(result, cur)
			cur = ""
		} else {
			cur += string(c)
		}
	}
	result = append(result, cur)
	return result
}

func TestChainCompletion(t *testing.T) {
	db := &stubDB{
		data: map[string][]MemberInfo{
			"PartDocument": {
				{Name: "Part", ReturnType: "Part", Kind: MemberKindProperty},
				{Name: "Path", ReturnType: "String", Kind: MemberKindProperty},
			},
			"Part": {
				{Name: "Bodies", ReturnType: "Bodies", Kind: MemberKindProperty},
				{Name: "Name", ReturnType: "String", Kind: MemberKindProperty},
			},
		},
	}
	table := &SymbolTable{
		Symbols: []Symbol{
			{Name: "partDocument1", TypeName: "PartDocument", Scope: "module"},
		},
	}

	t.Run("直接補完", func(t *testing.T) {
		items := BuildCompletionsFromResolver("partDocument1.", table, db, "module")
		if len(items) != 2 {
			t.Fatalf("want 2 items, got %d", len(items))
		}
		if items[0].Label != "Part" {
			t.Errorf("want Part, got %s", items[0].Label)
		}
	})

	t.Run("チェーン補完", func(t *testing.T) {
		items := BuildCompletionsFromResolver("partDocument1.Part.", table, db, "module")
		if len(items) != 2 {
			t.Fatalf("want 2 items, got %d", len(items))
		}
		labels := map[string]bool{}
		for _, item := range items {
			labels[item.Label] = true
		}
		if !labels["Bodies"] || !labels["Name"] {
			t.Errorf("want Bodies and Name, got %v", labels)
		}
	})

	t.Run("存在しないメンバーでチェーン", func(t *testing.T) {
		items := BuildCompletionsFromResolver("partDocument1.Unknown.", table, db, "module")
		if items != nil {
			t.Errorf("want nil, got %v", items)
		}
	})
}

func TestDotPattern(t *testing.T) {
	tests := []struct {
		input string
		want  string
	}{
		{"partDocument1.", "partDocument1"},
		{"partDocument1.Part.", "partDocument1.Part"},
		{"a.b.c.", "a.b.c"},
		{"notrailing", ""},
		{"部品.", "部品"},
		{"x = 部品.Part.", "部品.Part"},
		{"Debug.Print \"あ\": doc.", "doc"},
	}
	for _, tt := range tests {
		m := dotPattern.FindStringSubmatch(tt.input)
		if tt.want == "" {
			if m != nil {
				t.Errorf("input %q: want no match, got %v", tt.input, m)
			}
			continue
		}
		if m == nil {
			t.Errorf("input %q: want match %q, got nil", tt.input, tt.want)
			continue
		}
		if m[1] != tt.want {
			t.Errorf("input %q: want %q, got %q", tt.input, tt.want, m[1])
		}
	}
}
