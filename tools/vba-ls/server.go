package main

import (
	"context"
	"encoding/json"
	"regexp"
	"strings"
	"sync"

	"github.com/sourcegraph/jsonrpc2"
)

const serverName = "vba-ls"
const serverVersion = "0.5.2"

// documentStore はオープン中のドキュメントを保持する。
type documentStore struct {
	mu   sync.RWMutex
	docs map[string]string
}

func newDocumentStore() *documentStore {
	return &documentStore{docs: make(map[string]string)}
}
func (s *documentStore) set(uri, content string) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.docs[uri] = content
}
func (s *documentStore) get(uri string) (string, bool) {
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
	store  *documentStore
	db     *TLBDatabase
	tables sync.Map // uri → *SymbolTable
}

func newLSPHandler(db *TLBDatabase) *lspHandler {
	return &lspHandler{
		store: newDocumentStore(),
		db:    db,
	}
}

func (h *lspHandler) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	switch req.Method {
	case "initialize":
		h.handleInitialize(ctx, conn, req)
	case "initialized":
		// no-op
	case "shutdown":
		_ = conn.Reply(ctx, req.ID, nil)
	case "exit":
		// サーバー終了
	case "textDocument/didOpen":
		h.handleDidOpen(ctx, conn, req)
	case "textDocument/didChange":
		h.handleDidChange(ctx, conn, req)
	case "textDocument/didClose":
		h.handleDidClose(ctx, conn, req)
	case "textDocument/completion":
		h.handleCompletion(ctx, conn, req)
	case "textDocument/hover":
		h.handleHover(ctx, conn, req)
	case "textDocument/semanticTokens/full":
		h.handleSemanticTokens(ctx, conn, req)
	}
}

func (h *lspHandler) handleInitialize(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	result := InitializeResult{
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
	_ = conn.Reply(ctx, req.ID, result)
}

func (h *lspHandler) handleDidOpen(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	var p DidOpenTextDocumentParams
	if err := json.Unmarshal(*req.Params, &p); err != nil {
		return
	}
	uri := p.TextDocument.URI
	content := p.TextDocument.Text
	h.store.set(uri, content)
	h.reparse(uri, content)
	h.publishDiagnostics(ctx, conn, uri, content)
}

func (h *lspHandler) handleDidChange(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	var p DidChangeTextDocumentParams
	if err := json.Unmarshal(*req.Params, &p); err != nil {
		return
	}
	if len(p.ContentChanges) == 0 {
		return
	}
	uri := p.TextDocument.URI
	content := p.ContentChanges[len(p.ContentChanges)-1].Text
	h.store.set(uri, content)
	h.reparse(uri, content)
	h.publishDiagnostics(ctx, conn, uri, content)
}

func (h *lspHandler) handleDidClose(_ context.Context, _ *jsonrpc2.Conn, req *jsonrpc2.Request) {
	var p DidCloseTextDocumentParams
	if err := json.Unmarshal(*req.Params, &p); err != nil {
		return
	}
	h.store.delete(p.TextDocument.URI)
	h.tables.Delete(p.TextDocument.URI)
}

func (h *lspHandler) handleCompletion(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	var p CompletionParams
	if err := json.Unmarshal(*req.Params, &p); err != nil {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	uri := p.TextDocument.URI
	content, ok := h.store.get(uri)
	if !ok {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	table := h.getTable(uri)
	if table == nil {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	if h.db == nil {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	linePrefix := lineAt(content, int(p.Position.Line), int(p.Position.Character))
	scope := scopeAtLine(content, int(p.Position.Line))
	items := BuildCompletions(linePrefix, table, h.db, scope)
	_ = conn.Reply(ctx, req.ID, items)
}

func (h *lspHandler) handleHover(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	var p HoverParams
	if err := json.Unmarshal(*req.Params, &p); err != nil {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	uri := p.TextDocument.URI
	content, ok := h.store.get(uri)
	if !ok {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	table := h.getTable(uri)
	if table == nil {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	if h.db == nil {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	word, prevWord := wordAndPrevWord(content, int(p.Position.Line), int(p.Position.Character))
	scope := scopeAtLine(content, int(p.Position.Line))
	md := ResolveHoverMarkdown(word, prevWord, table, h.db, scope)
	if md == "" {
		_ = conn.Reply(ctx, req.ID, nil)
		return
	}
	_ = conn.Reply(ctx, req.ID, Hover{
		Contents: MarkupContent{Kind: "markdown", Value: md},
	})
}

func (h *lspHandler) handleSemanticTokens(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	var p SemanticTokensParams
	if err := json.Unmarshal(*req.Params, &p); err != nil {
		_ = conn.Reply(ctx, req.ID, SemanticTokens{Data: []uint32{}})
		return
	}
	uri := p.TextDocument.URI
	content, ok := h.store.get(uri)
	if !ok {
		_ = conn.Reply(ctx, req.ID, SemanticTokens{Data: []uint32{}})
		return
	}
	table := h.getTable(uri)
	data := BuildSemanticTokens(content, table, h.db)
	_ = conn.Reply(ctx, req.ID, SemanticTokens{Data: data})
}

func (h *lspHandler) reparse(uri, content string) {
	table := ParseSymbols(content)
	h.tables.Store(uri, table)
}

func (h *lspHandler) getTable(uri string) *SymbolTable {
	v, ok := h.tables.Load(uri)
	if !ok {
		return nil
	}
	return v.(*SymbolTable)
}

func (h *lspHandler) publishDiagnostics(ctx context.Context, conn *jsonrpc2.Conn, uri, content string) {
	table := h.getTable(uri)
	if table == nil || h.db == nil {
		return
	}
	diags := BuildDiagnostics(content, table, h.db)
	_ = conn.Notify(ctx, "textDocument/publishDiagnostics", PublishDiagnosticsParams{
		URI:         uri,
		Diagnostics: diags,
	})
}

// --- ユーティリティ ---

// lineAt は line 行目のカーソル位置より前のテキストを返す。
// char は LSP の UTF-16 位置。
func lineAt(content string, line, char int) string {
	l, ok := lineText(content, line)
	if !ok {
		return ""
	}
	return l[:utf16ToByteOffset(l, char)]
}

var wordRe = regexp.MustCompile(`[` + identPartClass + `]+`)

// wordAndPrevWord はカーソル位置の語と、その直前に "." で連結された語を返す。
// char は LSP の UTF-16 位置。
func wordAndPrevWord(content string, line, char int) (word, prevWord string) {
	l, ok := lineText(content, line)
	if !ok {
		return
	}
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

func scopeAtLine(content string, targetLine int) string {
	lines := strings.Split(content, "\n")
	scope := "module"
	for i, line := range lines {
		if i >= targetLine {
			break
		}
		if m := procStartRe.FindStringSubmatch(line); m != nil {
			kind := strings.Title(strings.ToLower(m[1])) //nolint:staticcheck
			scope = kind + ":" + m[2]
		} else if procEndRe.MatchString(line) {
			scope = "module"
		}
	}
	return scope
}
