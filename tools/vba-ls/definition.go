package main

import (
	"unicode"
	"unicode/utf8"
)

// ResolveDefinition はカーソル位置の識別子の宣言位置を返す。見つからなければ nil（応答は null）。
// 自モジュールのシンボルは現在のバッファの位置、他モジュールのメンバー・型はそのモジュールの位置を返す。
// TLB・組み込みの型とメンバーは宣言元が無いので nil。
func ResolveDefinition(doc *Document, line, utf16Char int) []Location {
	if line < 0 || line >= len(doc.Lines) {
		return nil
	}
	l := doc.Lines[line]
	s, e := identBounds(l, utf16ToByteOffset(l, utf16Char))
	if s == e {
		return nil
	}
	if r, _ := utf8.DecodeRuneInString(l[s:e]); !unicode.IsLetter(r) {
		return nil
	}
	ll := logicalLineAt(doc.Lines, line)
	offE := ll.Offset(line, e)
	masked, inStr, inCmt := maskCode(ll.Text[:offE])
	if inStr || inCmt {
		return nil
	}
	ch, ok := scanChainBackward(masked, offE)
	if !ok || len(ch.Segs) == 0 {
		return nil
	}
	last := ch.Segs[len(ch.Segs)-1]
	if !ch.LeadingDot && len(ch.Segs) == 1 {
		if sym := doc.Table.Lookup(last.Name, line); sym != nil {
			return []Location{declLocation(doc.URI, sym)}
		}
		if _, m, ok := doc.Types.ProjectMember(last.Name); ok {
			return originLocations(m.Origin)
		}
		if ti, ok := projectType(doc, last.Name); ok {
			return originLocations(ti.Origin)
		}
		return nil
	}
	owner := MemberChain{Segs: ch.Segs[:len(ch.Segs)-1], LeadingDot: ch.LeadingDot}
	t, ok := newExprResolver(doc).resolveChain(owner, line)
	if !ok {
		return nil
	}
	m, ok := doc.Types.Member(t, last.Name)
	if !ok {
		return nil
	}
	return originLocations(m.Origin)
}

// declLocation は宣言シンボルの名前の範囲（UTF-16）を返す。Symbol.Col は解析時に UTF-16 へ変換済み。
func declLocation(uri string, s *Symbol) Location {
	return Location{URI: uri, Range: Range{
		Start: Position{Line: uint32(s.Line), Character: uint32(s.Col)},
		End:   Position{Line: uint32(s.Line), Character: uint32(s.Col + utf16Len(s.Name))},
	}}
}

// originLocations は宣言元の位置を返す。シンボルが無ければ（クラス・フォーム・標準モジュール）モジュールの先頭。
func originLocations(o *DeclOrigin) []Location {
	if o == nil || o.URI == "" {
		return nil
	}
	if len(o.Syms) == 0 {
		return []Location{{URI: o.URI}}
	}
	out := make([]Location, len(o.Syms))
	for i, s := range o.Syms {
		out[i] = declLocation(o.URI, s)
	}
	return out
}
