package main

// segKind はセグメントの種類
type segKind int

const (
	segCode    segKind = iota // 通常のコード
	segString                 // 文字列リテラル内 ("..." )
	segComment                // コメント (' 以降)
)

// segment は行の一部とその種類
type segment struct {
	text string
	kind segKind
}

// parseSegments は1行をコード・文字列・コメントの3種に分類したセグメントのスライスを返す。
// VBA の文字列は " で囲まれ、"" は文字列内の " エスケープ。
// ' がコード部に現れた場合、それ以降がコメントになる。
func parseSegments(line string) []segment {
	var segs []segment
	inString := false
	start := 0
	i := 0

	flush := func(end int, kind segKind) {
		if end > start {
			segs = append(segs, segment{line[start:end], kind})
		}
		start = end
	}

	for i < len(line) {
		ch := line[i]

		if inString {
			if ch == '"' {
				// "" はエスケープ
				if i+1 < len(line) && line[i+1] == '"' {
					i += 2
					continue
				}
				// 文字列の終端
				i++
				flush(i, segString)
				inString = false
				continue
			}
			i++
			continue
		}

		// 文字列開始
		if ch == '"' {
			flush(i, segCode)
			inString = true
			i++
			continue
		}

		// コメント開始
		if ch == '\'' {
			flush(i, segCode)
			segs = append(segs, segment{line[i:], segComment})
			return segs
		}

		i++
	}

	// 行末処理
	if inString {
		flush(len(line), segString)
	} else {
		flush(len(line), segCode)
	}

	return segs
}

// codeOnly は行のコード部分のみを返す（文字列・コメント内を除いた部分）
// 文字列・コメント内は元のテキストをそのまま保持して結合して返す
func rebuildFromSegments(segs []segment) string {
	var b []byte
	for _, s := range segs {
		b = append(b, s.text...)
	}
	return string(b)
}
