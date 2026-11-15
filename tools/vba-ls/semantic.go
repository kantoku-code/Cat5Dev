package main

import (
	"regexp"
	"sort"
	"strings"
)

// セマンティックトークンタイプのインデックス。
// ServerCapabilities の legend.tokenTypes の順序と一致させること。
const (
	tokenTypeType     = uint32(0) // "type"     — CATIA 型名 (PartDocument 等)
	tokenTypeVariable = uint32(1) // "variable" — CATIA 型を持つ変数
)

// セマンティックトークン修飾子のビットマスク。
const (
	tokenModDeclaration = uint32(1 << 0) // "declaration"
)

// rawToken は位置情報付きのトークン1件（ソート用）。
type rawToken struct {
	line      int
	col       int
	length    int
	tokenType uint32
	modifiers uint32
}

var identRe = regexp.MustCompile(`[A-Za-z_]\w*`)

// BuildSemanticTokens は LSP の flat uint32 配列（5-tuple の deltaLine/deltaChar/length/type/modifier）を返す。
// SymbolTable の宣言情報と本文スキャンの両方から CATIA 型関連トークンを収集する。
func BuildSemanticTokens(src string, table *SymbolTable, db *TLBDatabase) []uint32 {
	if table == nil || db == nil {
		return []uint32{}
	}

	// CATIA 型を持つシンボルのセットを構築（大文字小文字無視）
	type catiaVar struct {
		typeName string
	}
	catiaVars := make(map[string]catiaVar) // lower(name) → info
	for _, sym := range table.Symbols {
		if sym.TypeName == "" {
			continue
		}
		if !db.TypeExists(sym.TypeName) {
			continue
		}
		catiaVars[strings.ToLower(sym.Name)] = catiaVar{typeName: sym.TypeName}
	}

	var tokens []rawToken

	// --- 1. 宣言箇所のトークン ---
	for _, sym := range table.Symbols {
		if sym.TypeName == "" || !db.TypeExists(sym.TypeName) {
			continue
		}
		// 変数名トークン（variable + declaration）
		tokens = append(tokens, rawToken{
			line:      sym.Line,
			col:       sym.Col,
			length:    utf16Len(sym.Name),
			tokenType: tokenTypeVariable,
			modifiers: tokenModDeclaration,
		})
		// 型名トークン（type）
		if sym.TypeName != "" {
			tokens = append(tokens, rawToken{
				line:      sym.TypeNameLine,
				col:       sym.TypeNameCol,
				length:    utf16Len(sym.TypeName),
				tokenType: tokenTypeType,
				modifiers: 0,
			})
		}
	}

	// --- 2. 使用箇所のトークン（宣言以外の識別子出現を走査）---
	// 正規表現・コメント判定はバイト位置で行い、トークンの列は UTF-16 位置に変換して扱う。
	lines := strings.Split(src, "\n")
	for lineIdx, line := range lines {
		line = strings.TrimSuffix(line, "\r")
		// コメント・文字列リテラルを除外するため、コメント開始位置（バイト）を特定
		commentStart := commentStartCol(line)

		matches := identRe.FindAllStringIndex(line, -1)
		for _, m := range matches {
			if commentStart >= 0 && m[0] >= commentStart {
				break // コメント部分はスキップ
			}
			word := line[m[0]:m[1]]
			col := byteToUTF16Offset(line, m[0])
			lower := strings.ToLower(word)
			if _, ok := catiaVars[lower]; !ok {
				continue
			}
			// 宣言トークンと重複しないか確認
			if isDeclToken(table, lower, lineIdx, col) {
				continue
			}
			tokens = append(tokens, rawToken{
				line:      lineIdx,
				col:       col,
				length:    utf16Len(word),
				tokenType: tokenTypeVariable,
				modifiers: 0,
			})
		}
	}

	// --- 3. 位置順ソート ---
	sort.Slice(tokens, func(i, j int) bool {
		if tokens[i].line != tokens[j].line {
			return tokens[i].line < tokens[j].line
		}
		return tokens[i].col < tokens[j].col
	})

	// 重複除去（同一位置）
	tokens = dedupTokens(tokens)

	// --- 4. LSP 5-tuple エンコード（delta 形式） ---
	data := make([]uint32, 0, len(tokens)*5)
	prevLine := 0
	prevCol := 0
	for _, t := range tokens {
		deltaLine := t.line - prevLine
		deltaCol := t.col
		if deltaLine == 0 {
			deltaCol = t.col - prevCol
		}
		data = append(data, uint32(deltaLine), uint32(deltaCol), uint32(t.length), t.tokenType, t.modifiers)
		prevLine = t.line
		prevCol = t.col
	}
	return data
}

// commentStartCol は行内のコメント開始位置（バイト）を返す。コメントがなければ -1。
// 文字列リテラル内の ' は無視する。
func commentStartCol(line string) int {
	inStr := false
	for i, ch := range line {
		if ch == '"' {
			inStr = !inStr
		} else if ch == '\'' && !inStr {
			return i
		}
	}
	return -1
}

// isDeclToken はシンボルテーブル内の宣言位置と一致するか確認する。
// col は UTF-16 位置（Symbol.Col と同じ単位）。
func isDeclToken(table *SymbolTable, lowerName string, line, col int) bool {
	for _, sym := range table.Symbols {
		if strings.ToLower(sym.Name) == lowerName && sym.Line == line && sym.Col == col {
			return true
		}
	}
	return false
}

// dedupTokens は同一位置のトークンを除去する（先勝ち）。
func dedupTokens(tokens []rawToken) []rawToken {
	if len(tokens) == 0 {
		return tokens
	}
	out := []rawToken{tokens[0]}
	for i := 1; i < len(tokens); i++ {
		prev := out[len(out)-1]
		cur := tokens[i]
		if cur.line == prev.line && cur.col == prev.col {
			continue
		}
		out = append(out, cur)
	}
	return out
}
