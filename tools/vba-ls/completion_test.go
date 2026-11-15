package main

import (
	"testing"
)

func TestChainCompletion(t *testing.T) {
	const decl = "Dim partDocument1 As PartDocument\n"

	t.Run("直接補完", func(t *testing.T) {
		labels := labelsOf(completeAt(t, nil, decl+"partDocument1.|"))
		if len(labels) != 3 || !labels["Part"] || !labels["Name"] || !labels["Save"] {
			t.Fatalf("候補 = %v, want Part/Name/Save", labels)
		}
	})

	t.Run("チェーン補完", func(t *testing.T) {
		labels := labelsOf(completeAt(t, nil, decl+"partDocument1.Part.|"))
		if len(labels) != 2 || !labels["Bodies"] || !labels["Name"] {
			t.Errorf("候補 = %v, want Bodies と Name", labels)
		}
	})

	t.Run("存在しないメンバーでチェーン", func(t *testing.T) {
		if items := completeAt(t, nil, decl+"partDocument1.Unknown.|"); items != nil {
			t.Errorf("want nil, got %v", items)
		}
	})

	t.Run("IUnknownのメソッドは候補に出さない", func(t *testing.T) {
		labels := labelsOf(completeAt(t, nil, "Dim b As Bodies\nb.|"))
		if labels["QueryInterface"] || !labels["Item"] || !labels["Count"] {
			t.Errorf("候補 = %v, want Item/Count のみ", labels)
		}
	})
}

// DB がなくても VBA / Scripting の組み込み型は補完できる。
func TestCompletionWithoutDB(t *testing.T) {
	noDB := NewDefaultResolver(nil, nil)
	for _, tc := range []struct {
		name string
		src  string
		want []string
	}{
		{"CreateObjectのDictionary", "Sub M()\n    Set d = CreateObject(\"Scripting.Dictionary\")\n    d.|\nEnd Sub\n", []string{"Add", "Exists", "Keys", "Count"}},
		{"Dim As New Collection", "Sub M()\n    Dim c As New Collection\n    c.|\nEnd Sub\n", []string{"Add", "Count", "Item", "Remove"}},
		{"As Scripting.FileSystemObject", "Dim fso As Scripting.FileSystemObject\nfso.|\n", []string{"OpenTextFile", "GetFile", "FileExists"}},
	} {
		lfAndCRLF(t, tc.src, func(t *testing.T, src string) {
			labels := labelsOf(completeAt(t, noDB, src))
			for _, w := range tc.want {
				if !labels[w] {
					t.Errorf("%s: 候補に %s がない: %v", tc.name, w, labels)
				}
			}
		})
	}

	t.Run("handler", func(t *testing.T) {
		h := newLSPHandler(noDB)
		const uri = "file:///a.bas"
		src := "Dim d As New Scripting.Dictionary\nd."
		h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: src}})
		var p CompletionParams
		p.TextDocument.URI = uri
		p.Position = Position{Line: 1, Character: 2}
		if items := h.completion(p); items == nil || !labelsOf(items)["Exists"] {
			t.Fatalf("handler の補完 = %v, want Exists を含む", items)
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
