package main

import (
	"context"
	"encoding/json"
	"regexp"
	"sync"

	"github.com/sourcegraph/jsonrpc2"
)

const serverName = "vba-ls"
const serverVersion = "0.5.2"

// docState はオープン中の文書 1 件の状態（テキスト・行・最後に解析したシンボル表）。
type docState struct {
	text  string
	lines []string
	table *SymbolTable
}

// documentStore はオープン中のドキュメントを保持する。
type documentStore struct {
	mu   sync.RWMutex
	docs map[string]*docState
}

func newDocumentStore() *documentStore {
	return &documentStore{docs: make(map[string]*docState)}
}

func (s *documentStore) set(uri string, st *docState) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.docs[uri] = st
}

func (s *documentStore) get(uri string) (*docState, bool) {
	s.mu.RLock()
	defer s.mu.RUnlock()
	v, ok := s.docs[uri]
	return v, ok
}

func (s *documentStore) delete(uri string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	delete(s.docs, uri)
}

// lspHandler は jsonrpc2.Handler を実装する LSP サーバー。
type lspHandler struct {
	store     *documentStore
	types     *CompositeResolver // 全文書で共有する型解決（Workspace → VBA → TLB → Scripting）
	workspace TypeProvider       // 2-3 の WorkspaceProvider の差し込み口
}

func newLSPHandler(types *CompositeResolver) *lspHandler {
	if types == nil {
		types = NewDefaultResolver(nil, nil)
	}
	return &lspHandler{
		store: newDocumentStore(),
		types: types,
	}
}

func (h *lspHandler) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	switch req.Method {
	case "initialize":
		_ = conn.Reply(ctx, req.ID, h.initializeResult())
	case "initialized":
		// no-op
	case "shutdown":
		_ = conn.Reply(ctx, req.ID, nil)
	case "exit":
		// サーバー終了
	case "textDocument/didOpen":
		var p DidOpenTextDocumentParams
		if !unmarshalParams(req, &p) {
			return
		}
		h.didOpen(p)
		h.publishDiagnostics(ctx, conn, p.TextDocument.URI)
	case "textDocument/didChange":
		var p DidChangeTextDocumentParams
		if !unmarshalParams(req, &p) || len(p.ContentChanges) == 0 {
			return
		}
		h.didChange(p)
		h.publishDiagnostics(ctx, conn, p.TextDocument.URI)
	case "textDocument/didClose":
		var p DidCloseTextDocumentParams
		if !unmarshalParams(req, &p) {
			return
		}
		h.store.delete(p.TextDocument.URI)
	case "textDocument/completion":
		var p CompletionParams
		if !unmarshalParams(req, &p) {
			_ = conn.Reply(ctx, req.ID, nil)
			return
		}
		items := h.completion(p)
		if items == nil {
			_ = conn.Reply(ctx, req.ID, nil)
			return
		}
		_ = conn.Reply(ctx, req.ID, items)
	case "textDocument/hover":
		var p HoverParams
		if !unmarshalParams(req, &p) {
			_ = conn.Reply(ctx, req.ID, nil)
			return
		}
		if hv := h.hover(p); hv != nil {
			_ = conn.Reply(ctx, req.ID, hv)
			return
		}
		_ = conn.Reply(ctx, req.ID, nil)
	case "textDocument/semanticTokens/full":
		var p SemanticTokensParams
		if !unmarshalParams(req, &p) {
			_ = conn.Reply(ctx, req.ID, SemanticTokens{Data: []uint32{}})
			return
		}
		_ = conn.Reply(ctx, req.ID, SemanticTokens{Data: h.semanticTokens(p)})
	}
}

func unmarshalParams(req *jsonrpc2.Request, v interface{}) bool {
	if req.Params == nil {
		return false
	}
	return json.Unmarshal(*req.Params, v) == nil
}

func (h *lspHandler) initializeResult() InitializeResult {
	return InitializeResult{
		Capabilities: ServerCapabilities{
			TextDocumentSync: TextDocumentSyncKindFull,
			CompletionProvider: &CompletionOptions{
				TriggerCharacters: []string{"."},
			},
			HoverProvider: true,
			SemanticTokensProvider: &SemanticTokensOptions{
				Legend: SemanticTokensLegend{
					TokenTypes:     []string{"type", "variable"},
					TokenModifiers: []string{"declaration"},
				},
				Full: true,
			},
		},
		ServerInfo: &ServerInfo{Name: serverName, Version: serverVersion},
	}
}

func (h *lspHandler) didOpen(p DidOpenTextDocumentParams) {
	h.update(p.TextDocument.URI, p.TextDocument.Text)
}

func (h *lspHandler) didChange(p DidChangeTextDocumentParams) {
	h.update(p.TextDocument.URI, p.ContentChanges[len(p.ContentChanges)-1].Text)
}

// update は文書を解析して保存する。
func (h *lspHandler) update(uri, text string) {
	lines := splitLines(text)
	h.store.set(uri, &docState{text: text, lines: lines, table: ParseSymbolsLines(text, lines)})
}

// document は uri の最新の Document を返す。開かれていなければ nil。
func (h *lspHandler) document(uri string) *Document {
	st, ok := h.store.get(uri)
	if !ok {
		return nil
	}
	return newDocumentWithLines(uri, st.text, st.lines, st.table, h.types)
}

func (h *lspHandler) completion(p CompletionParams) []CompletionItem {
	doc := h.document(p.TextDocument.URI)
	if doc == nil {
		return nil
	}
	return BuildCompletions(doc, int(p.Position.Line), int(p.Position.Character))
}

func (h *lspHandler) hover(p HoverParams) *Hover {
	doc := h.document(p.TextDocument.URI)
	if doc == nil {
		return nil
	}
	md := ResolveHover(doc, int(p.Position.Line), int(p.Position.Character))
	if md == "" {
		return nil
	}
	return &Hover{Contents: MarkupContent{Kind: "markdown", Value: md}}
}

func (h *lspHandler) semanticTokens(p SemanticTokensParams) []uint32 {
	doc := h.document(p.TextDocument.URI)
	if doc == nil {
		return []uint32{}
	}
	return BuildSemanticTokens(doc)
}

func (h *lspHandler) diagnostics(uri string) []Diagnostic {
	doc := h.document(uri)
	if doc == nil {
		return nil
	}
	diags := BuildDiagnostics(doc)
	if diags == nil {
		diags = []Diagnostic{}
	}
	return diags
}

func (h *lspHandler) publishDiagnostics(ctx context.Context, conn *jsonrpc2.Conn, uri string) {
	diags := h.diagnostics(uri)
	if diags == nil {
		return
	}
	_ = conn.Notify(ctx, "textDocument/publishDiagnostics", PublishDiagnosticsParams{
		URI:         uri,
		Diagnostics: diags,
	})
}

// --- ユーティリティ ---

var wordRe = regexp.MustCompile(`[` + identPartClass + `]+`)

// wordAndPrevWord はカーソル位置の語と、その直前に "." で連結された語を返す。
// char は LSP の UTF-16 位置。
func wordAndPrevWord(content string, line, char int) (word, prevWord string) {
	l, ok := lineText(content, line)
	if !ok {
		return
	}
	return wordAndPrevWordInLine(l, char)
}

func wordAndPrevWordInLine(l string, char int) (word, prevWord string) {
	char = utf16ToByteOffset(l, char)
	matches := wordRe.FindAllStringIndex(l, -1)
	for _, m := range matches {
		if m[0] <= char && char <= m[1] {
			word = l[m[0]:m[1]]
			if m[0] > 0 && l[m[0]-1] == '.' {
				for _, pm := range matches {
					if pm[1] == m[0]-1 {
						prevWord = l[pm[0]:pm[1]]
						break
					}
				}
			}
			break
		}
	}
	return
}
