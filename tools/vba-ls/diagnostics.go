package main

import (
	"unicode"
	"unicode/utf8"
)

// BuildDiagnostics は文書全体を走査して、存在しないメンバーへのアクセスを診断として返す。
// 論理行（行継続を連結）ごとに文字列・コメントをマスクし、"." の直後の識別子をすべて調べる。
// 受け手の型が解決でき、メンバー一覧が網羅的（Complete）な型の場合だけ診断する。
// 組み込み型・ワークスペースの型は Error、TLB（CATIA）型は Warning にする（TLB に載らない非公開の API があるため）。
// TLB 型は遅延バインディングなので、継承を持つ DB（v2）では宣言型とその派生型のどれにも無い名前だけを、
// 旧 DB（v1）では TLB 全体に無い名前だけを診断する。
// v1 では TLB 型のメンバーの戻り値を経由して求めた型も診断しない（typelib が基底型を宣言していることが多いため。
// 例: Part.UserSurfaces は Collection と宣言されているが実体は UserSurfaces）。v2 は派生型込みで判定するので対象にする。
func BuildDiagnostics(doc *Document) []Diagnostic {
	var diags []Diagnostic
	r := newExprResolver(doc)
	for _, ll := range buildLogicalLines(doc.Lines) {
		masked, _, _ := maskCode(ll.Text)
		for i := 0; i < len(masked); i++ {
			if masked[i] != '.' {
				continue
			}
			// "." の直後（空白を許す）の識別子
			a := i + 1
			for a < len(masked) && (masked[a] == ' ' || masked[a] == '\t') {
				a++
			}
			first, _ := utf8.DecodeRuneInString(masked[a:])
			if a >= len(masked) || !unicode.IsLetter(first) {
				continue
			}
			b := a
			for b < len(masked) {
				rr, size := utf8.DecodeRuneInString(masked[b:])
				if !isIdentPart(rr) {
					break
				}
				b += size
			}
			name := masked[a:b]

			line, _ := ll.Physical(i)
			var ch MemberChain
			if endsChainOperand(masked, i) {
				var ok bool
				if ch, ok = scanChainBackward(masked, i); !ok {
					continue
				}
			} else {
				ch = MemberChain{LeadingDot: true, Start: i}
			}
			r.viaTLBMember = false
			t, ok := r.resolveChain(ch, line)
			if !ok || t.Array {
				continue
			}
			ti, prov, ok := doc.Types.Resolve(t)
			if !ok || ti.Kind == TypeKindPrimitive || !ti.Complete {
				continue
			}
			hier, _ := prov.(MemberHierarchyIndex)
			hierarchy := hier != nil && hier.HasHierarchy()
			if r.viaTLBMember && !hierarchy {
				continue
			}
			if _, ok := doc.Types.Member(t, name); ok {
				continue
			}
			sev := DiagnosticSeverityError
			msg := "'" + t.String() + "' に '" + name + "' は存在しません"
			if _, isTLB := prov.(*TLBProvider); isTLB {
				// CATIA は遅延バインディング。宣言型が基底型（Document 等）でも派生型のメンバーが実行時に動くため、
				// v2 は宣言型と派生型のどれにも無い名前、v1 は TLB のどこにも無い名前だけを警告にする。
				// TLB に載らない API（CATIA.ActiveEditor など）があるので Error にはしない。
				if hierarchy {
					if hier.HasMemberInDescendants(t, name) {
						continue
					}
					msg += "（派生型を含む）"
				} else if idx, ok := prov.(MemberNameIndex); ok && idx.HasMemberNamed(name) {
					continue
				}
				sev = DiagnosticSeverityWarning
			}
			sl, sb := ll.Physical(a)
			el, eb := ll.Physical(b)
			if el != sl {
				continue // 識別子は行をまたがない
			}
			diags = append(diags, Diagnostic{
				Range: Range{
					Start: Position{Line: uint32(sl), Character: uint32(byteToUTF16Offset(doc.Lines[sl], sb))},
					End:   Position{Line: uint32(el), Character: uint32(byteToUTF16Offset(doc.Lines[el], eb))},
				},
				Severity: sev,
				Source:   "vba-ls",
				Message:  msg,
			})
		}
	}
	return diags
}

// stripCommentAndStrings はコメント（'）と文字列リテラルの中身を空白で置換する。
// 置換はバイト単位で行い、戻り値は line とバイト長が等しい（位置の対応を保つため）。
// 二重引用符と単一引用符はどちらも ASCII なので、UTF-8 の多バイト文字の途中で誤検出することはない。
// "" エスケープは「閉じて直後に開く」扱いになり、引用符自体は残る。
func stripCommentAndStrings(line string) string {
	inStr := false
	result := []byte(line)
	for i, ch := range result {
		if inStr {
			if ch == '"' {
				inStr = false
			} else {
				result[i] = ' '
			}
		} else {
			if ch == '"' {
				inStr = true
			} else if ch == '\'' {
				for j := i; j < len(result); j++ {
					result[j] = ' '
				}
				break
			}
		}
	}
	return string(result)
}
