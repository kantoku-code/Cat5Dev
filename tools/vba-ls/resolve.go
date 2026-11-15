package main

import (
	"regexp"
	"strings"
)

// maxWithDepth は With の対象式を解決するときの入れ子の上限（循環の防止）。
const maxWithDepth = 8

// exprResolver はシンボル表と型解決を使って、メンバーアクセスの連鎖の型を求める。
type exprResolver struct {
	table *SymbolTable
	types *CompositeResolver
	self  TypeRef // Me の型（標準モジュールではゼロ値＝Me. は解決しない）
	depth int
	// viaTLBMember は解決の途中で TLB（CATIA）型のメンバーの戻り値・既定メンバーを経由したかを記録する。
	// typelib は戻り値に基底型（Collection / Factory / Viewer など）を宣言していることが多く、
	// 実際のオブジェクトはその派生型なので、診断ではこの型を「確定」とみなさない。
	viaTLBMember bool
}

func newExprResolver(doc *Document) *exprResolver {
	return &exprResolver{table: doc.Table, types: doc.Types, self: doc.Self}
}

// resolveChain は連鎖全体の型を返す。LeadingDot なら With の対象式を起点にする。
func (r *exprResolver) resolveChain(ch MemberChain, line int) (TypeRef, bool) {
	segs := ch.Segs
	var cur TypeRef
	var ok bool
	if ch.LeadingDot {
		cur, ok = r.withTarget(line)
	} else {
		if len(segs) == 0 {
			return TypeRef{}, false
		}
		cur, ok = r.resolveHead(segs[0], line)
		segs = segs[1:]
	}
	for _, sg := range segs {
		if !ok {
			return TypeRef{}, false
		}
		cur, ok = r.step(cur, sg)
	}
	return cur, ok
}

// step は cur 型の値に対する 1 段のメンバーアクセスの型を返す。
func (r *exprResolver) step(cur TypeRef, sg ExprSegment) (TypeRef, bool) {
	if sg.Bang {
		// a!b は既定メンバーに文字列キー "b" を渡す
		t, ok := r.applyIndex(cur, 1)
		if !ok {
			return TypeRef{}, false
		}
		return r.applyIndex(t, sg.Calls)
	}
	m, ok := r.types.Member(cur, sg.Name)
	if !ok {
		return TypeRef{}, false
	}
	r.noteTLB(cur)
	t, ok := r.types.MemberType(cur, sg.Name)
	if !ok {
		return TypeRef{}, false
	}
	calls := sg.Calls
	if calls > 0 && (m.Kind == MemberKindMethod || len(m.Params) > 0) {
		calls-- // 呼び出しそのものの分
	}
	return r.applyIndex(t, calls)
}

// resolveHead は連鎖の先頭の型を返す。解決順:
// Me → 自モジュールのシンボル（ローカル・引数・モジュール変数・プロシージャ）→ モジュール名・フォームの既定インスタンス（Main. / frmMenu.）
// → 標準モジュールの Public（修飾なし）→ ライブラリのグローバル（VBA でもプロジェクトの名前が参照設定より優先）→ Public Enum（EKind.kA）。
func (r *exprResolver) resolveHead(sg ExprSegment, line int) (TypeRef, bool) {
	if strings.EqualFold(sg.Name, "Me") || r.table.Lookup(sg.Name, line) != nil {
		return r.resolveLocalHead(sg, line)
	}
	// 修飾子としてのモジュール名。同名のプロシージャがあっても、括弧なしならモジュールを優先する
	if sg.Calls == 0 {
		if t, ok := r.types.ModuleRef(sg.Name); ok {
			return t, true
		}
	}
	// 標準モジュールの Public メンバー（修飾なし）。呼び出し分の調整は step に任せる
	if owner, _, ok := r.types.ProjectMember(sg.Name); ok {
		return r.step(owner, sg)
	}
	if t, ok := r.types.Global(sg.Name); ok {
		return r.applyIndex(t, sg.Calls)
	}
	// 他モジュールの Public Enum
	if sg.Calls == 0 {
		if ti, _, ok := r.types.Resolve(TypeRef{Name: sg.Name}); ok && ti.Kind == TypeKindEnum {
			return ti.Ref, true
		}
	}
	return TypeRef{}, false
}

// resolveLocalHead は Me と自モジュールのシンボル表で連鎖の先頭を解決する。
func (r *exprResolver) resolveLocalHead(sg ExprSegment, line int) (TypeRef, bool) {
	if strings.EqualFold(sg.Name, "Me") {
		if r.self.IsZero() {
			return TypeRef{}, false
		}
		return r.applyIndex(r.self, sg.Calls)
	}
	if sym := r.table.Lookup(sg.Name, line); sym != nil {
		switch sym.Kind {
		case SymVariable, SymParam, SymConst:
			t := sym.EffectiveTypeRef(r.types) // 'As 型 の型ヒントがあればそれ
			if t.Name == "" {
				return TypeRef{}, false
			}
			return r.applyIndex(t, sg.Calls)
		case SymProcedure, SymProperty:
			if sym.Decl == nil || !sym.Decl.HasReturn() {
				return TypeRef{}, false
			}
			calls := sg.Calls
			if calls > 0 {
				calls-- // 呼び出しそのものの分
			}
			return r.applyIndex(sym.Decl.ReturnType, calls)
		case SymEnum:
			return TypeRef{Name: sym.Name}, sg.Calls == 0
		}
		return TypeRef{}, false
	}
	return TypeRef{}, false
}

// applyIndex は calls 回の添字・既定メンバー呼び出しを適用する。配列なら要素型に、そうでなければ既定メンバーの型にする。
func (r *exprResolver) applyIndex(t TypeRef, calls int) (TypeRef, bool) {
	for i := 0; i < calls; i++ {
		if t.Array {
			t = t.Elem()
			continue
		}
		dm := r.types.DefaultMember(t)
		if dm == "" {
			return TypeRef{}, false
		}
		r.noteTLB(t)
		next, ok := r.types.MemberType(t, dm)
		if !ok {
			return TypeRef{}, false
		}
		t = next
	}
	return t, !t.IsZero()
}

// noteTLB は t が TLB の型なら viaTLBMember を立てる。
func (r *exprResolver) noteTLB(t TypeRef) {
	if _, p, ok := r.types.Resolve(t); ok {
		if _, isTLB := p.(*TLBProvider); isTLB {
			r.viaTLBMember = true
		}
	}
}

// withTarget は line を含む With ブロックの対象式の型を返す。
func (r *exprResolver) withTarget(line int) (TypeRef, bool) {
	w := r.table.WithAt(line)
	if w == nil || r.depth >= maxWithDepth {
		return TypeRef{}, false
	}
	r.depth++
	defer func() { r.depth-- }()
	return r.resolveExprText(w.Expr, w.StartLine)
}

var createObjectRe = regexp.MustCompile(`(?i)^(?:VBA\.)?CreateObject\s*\(\s*"([^"]+)"`)

// resolveExprText は式のテキスト（With の対象式など）の型を返す。
// "New X" と CreateObject("ProgID") を扱い、それ以外は連鎖として全体を読めたときだけ成功する。
func (r *exprResolver) resolveExprText(text string, line int) (TypeRef, bool) {
	t := strings.TrimSpace(text)
	if len(t) > 4 && strings.EqualFold(t[:4], "New ") {
		ref := ParseTypeName(t[4:])
		return ref, ref.Name != ""
	}
	if m := createObjectRe.FindStringSubmatch(t); m != nil {
		ref := TypeRefFromProgID(m[1])
		return ref, ref.Name != ""
	}
	masked, inStr, inCmt := maskCode(t)
	if inStr || inCmt {
		return TypeRef{}, false
	}
	ch, ok := scanChainBackward(masked, len(masked))
	if !ok || ch.Start != 0 {
		return TypeRef{}, false
	}
	return r.resolveChain(ch, line)
}
