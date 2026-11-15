package main

import (
	"sort"
	"strings"
)

// moduleTypeProvider は現在のモジュールで定義されたユーザー定義型（Type）と列挙型（Enum）を提供する。
// 解決順の先頭に置く（VBA でもモジュール内の定義が参照設定より優先される）。
type moduleTypeProvider struct {
	table *SymbolTable
}

// newModuleTypeProvider は UDT も Enum もなければ interface の nil を返す。
func newModuleTypeProvider(table *SymbolTable) TypeProvider {
	if table == nil || (len(table.UDTs) == 0 && len(table.Enums) == 0) {
		return nil
	}
	return &moduleTypeProvider{table: table}
}

func (p *moduleTypeProvider) OwnsLibrary(string) bool { return false }

func (p *moduleTypeProvider) LookupType(name string) (TypeInfo, bool) {
	if u := p.table.UDT(name); u != nil {
		return TypeInfo{Ref: TypeRef{Name: u.Name}, Kind: TypeKindUDT, Complete: true, Doc: u.Doc}, true
	}
	if e := p.table.Enum(name); e != nil {
		return TypeInfo{Ref: TypeRef{Name: e.Name}, Kind: TypeKindEnum, Complete: true, Doc: e.Doc}, true
	}
	return TypeInfo{}, false
}

func (p *moduleTypeProvider) Members(t TypeRef) []MemberInfo {
	var out []MemberInfo
	if u := p.table.UDT(t.Name); u != nil {
		for _, m := range u.Members {
			out = append(out, MemberInfo{
				Name: m.Name, ReturnType: m.QualifiedTypeName(), Kind: MemberKindField,
				Doc: m.Doc, Visibility: VisPublic, Owner: TypeRef{Name: u.Name},
			})
		}
	} else if e := p.table.Enum(t.Name); e != nil {
		for _, m := range e.Members {
			doc := m.Doc
			if m.Value != "" {
				doc = strings.TrimSpace("= " + m.Value + "\n" + doc)
			}
			out = append(out, MemberInfo{
				Name: m.Name, ReturnType: "Long", Kind: MemberKindEnumMember,
				Doc: doc, Visibility: VisPublic, Owner: TypeRef{Name: e.Name},
			})
		}
	}
	sort.SliceStable(out, func(i, j int) bool { return strings.ToLower(out[i].Name) < strings.ToLower(out[j].Name) })
	return out
}

func (p *moduleTypeProvider) Member(t TypeRef, name string) (MemberInfo, bool) {
	for _, m := range p.Members(t) {
		if strings.EqualFold(m.Name, name) {
			return m, true
		}
	}
	return MemberInfo{}, false
}

func (p *moduleTypeProvider) MemberType(t TypeRef, name string) (TypeRef, bool) {
	if u := p.table.UDT(t.Name); u != nil {
		for _, m := range u.Members {
			if strings.EqualFold(m.Name, name) {
				return m.TypeRef(), m.TypeName != ""
			}
		}
		return TypeRef{}, false
	}
	if e := p.table.Enum(t.Name); e != nil {
		for _, m := range e.Members {
			if strings.EqualFold(m.Name, name) {
				return TypeRef{Lib: "VBA", Name: "Long"}, true
			}
		}
	}
	return TypeRef{}, false
}
