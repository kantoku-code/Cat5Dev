package main

import (
	"context"
	"encoding/json"
	"net"
	"sync"
	"testing"
	"time"

	"github.com/sourcegraph/jsonrpc2"
)

func completionParams(uri string, line, char int) CompletionParams {
	var p CompletionParams
	p.TextDocument.URI = uri
	p.Position = Position{Line: uint32(line), Character: uint32(char)}
	return p
}

func changeParams(uri, text string) DidChangeTextDocumentParams {
	var p DidChangeTextDocumentParams
	p.TextDocument.URI = uri
	p.ContentChanges = []TextDocumentContentChangeEvent{{Text: text}}
	return p
}

// 入力が続く間は最新の版だけを解析し、途中の版の解析は破棄する。
func TestReparseDebounceParsesLatestOnly(t *testing.T) {
	h := newLSPHandler(newTestResolver(t))
	h.debounce = 50 * time.Millisecond
	var mu sync.Mutex
	parsed := map[int]bool{}
	h.onParsed = func(_ string, gen int, latest bool) {
		mu.Lock()
		parsed[gen] = latest
		mu.Unlock()
	}
	const uri = "file:///a.bas"
	h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: "Dim d As PartDocument\n"}})
	if !h.waitParsed(uri, time.Minute) {
		t.Fatal("didOpen の解析が終わらない")
	}
	for i := 0; i < 5; i++ {
		h.didChange(changeParams(uri, "Dim d As PartDocument\n"+string(rune('a'+i))+" = 1\n"))
	}
	if !h.waitParsed(uri, time.Minute) {
		t.Fatal("didChange の解析が終わらない")
	}
	mu.Lock()
	defer mu.Unlock()
	if latest, ok := parsed[6]; !ok || !latest {
		t.Errorf("最新の版（6）が解析されていない: %v", parsed)
	}
	for gen := 2; gen <= 5; gen++ {
		if _, ok := parsed[gen]; ok {
			t.Errorf("途中の版 %d が解析された: %v", gen, parsed)
		}
	}
}

// 解析が終わる前でも、補完は最新のテキストと直前のシンボル表で動く。
func TestCompletionUsesLatestTextWithPreviousTable(t *testing.T) {
	h := newLSPHandler(newTestResolver(t))
	h.debounce = time.Hour // 再解析させない
	const uri = "file:///a.bas"
	h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: "Sub M()\n    Dim d As PartDocument\n\nEnd Sub\n"}})
	h.waitParsed(uri, time.Minute)
	h.didChange(changeParams(uri, "Sub M()\n    Dim d As PartDocument\n    d.Part.\nEnd Sub\n"))
	if h.waitParsed(uri, 10*time.Millisecond) {
		t.Fatal("再解析されないはずが追いついている")
	}
	items := h.completion(completionParams(uri, 2, len("    d.Part.")))
	if !labelsOf(items)["Bodies"] {
		t.Fatalf("補完 = %v, want Bodies を含む", labelsOf(items))
	}
}

// didClose で予約中の再解析を取り消し、待っている側も解放する。
func TestDidCloseCancelsReparse(t *testing.T) {
	h := newLSPHandler(newTestResolver(t))
	h.debounce = 30 * time.Millisecond
	called := make(chan int, 10)
	h.onParsed = func(_ string, gen int, _ bool) { called <- gen }
	const uri = "file:///a.bas"
	h.didOpen(DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: "Dim d As Part\n"}})
	h.waitParsed(uri, time.Minute)
	<-called
	h.didChange(changeParams(uri, "Dim d As PartDocument\n"))
	h.didClose(uri)
	select {
	case gen := <-called:
		t.Fatalf("close 後に解析された（世代 %d）", gen)
	case <-time.After(150 * time.Millisecond):
	}
	if h.waitParsed(uri, 10*time.Millisecond) || h.document(uri) != nil {
		t.Fatal("close 後も文書が残っている")
	}
}

// JSON-RPC 経由: 解析中も要求に応答し、セマンティックトークンは解析後の結果を返し、診断を通知する。
func TestServerOverJSONRPC(t *testing.T) {
	ctx := context.Background()
	h := newLSPHandler(newTestResolver(t))
	h.debounce = 20 * time.Millisecond
	sc, cc := net.Pipe()
	server := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(sc, jsonrpc2.VSCodeObjectCodec{}), h)
	defer server.Close()

	diagCh := make(chan PublishDiagnosticsParams, 10)
	client := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(cc, jsonrpc2.VSCodeObjectCodec{}),
		jsonrpc2.HandlerWithError(func(_ context.Context, _ *jsonrpc2.Conn, req *jsonrpc2.Request) (interface{}, error) {
			if req.Method == "textDocument/publishDiagnostics" && req.Params != nil {
				var p PublishDiagnosticsParams
				if json.Unmarshal(*req.Params, &p) == nil {
					diagCh <- p
				}
			}
			return nil, nil
		}))
	defer client.Close()

	var init InitializeResult
	if err := client.Call(ctx, "initialize", map[string]interface{}{}, &init); err != nil || init.ServerInfo == nil || init.ServerInfo.Version != serverVersion {
		t.Fatalf("initialize: %v %+v", err, init)
	}
	const uri = "file:///m.bas"
	src := "Sub M()\n    Dim d As PartDocument\n    d.Bogus\n    d.\nEnd Sub\n"
	if err := client.Notify(ctx, "textDocument/didOpen", DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: uri, Text: src}}); err != nil {
		t.Fatal(err)
	}
	var toks SemanticTokens
	if err := client.Call(ctx, "textDocument/semanticTokens/full", SemanticTokensParams{TextDocument: TextDocumentIdentifier{URI: uri}}, &toks); err != nil {
		t.Fatal(err)
	}
	if len(toks.Data) == 0 {
		t.Error("セマンティックトークンが空（解析を待っていない）")
	}
	var items []CompletionItem
	if err := client.Call(ctx, "textDocument/completion", completionParams(uri, 3, len("    d.")), &items); err != nil {
		t.Fatal(err)
	}
	if !labelsOf(items)["Part"] {
		t.Errorf("補完 = %v", labelsOf(items))
	}
	select {
	case p := <-diagCh:
		if p.URI != uri || len(p.Diagnostics) != 1 || p.Diagnostics[0].Range.Start.Line != 2 {
			t.Errorf("診断 = %+v", p)
		}
	case <-time.After(30 * time.Second):
		t.Fatal("診断が通知されない")
	}

	// 変更後も最新の版の診断が通知される
	if err := client.Notify(ctx, "textDocument/didChange", changeParams(uri, "Sub M()\n    Dim d As PartDocument\n    d.Name = 1\nEnd Sub\n")); err != nil {
		t.Fatal(err)
	}
	select {
	case p := <-diagCh:
		if len(p.Diagnostics) != 0 {
			t.Errorf("変更後の診断 = %+v", p)
		}
	case <-time.After(30 * time.Second):
		t.Fatal("変更後の診断が通知されない")
	}
}
