package main

import (
	"context"
	"encoding/json"
	"fmt"
	"os"
	"sort"
	"sync"
	"time"

	"github.com/sourcegraph/jsonrpc2"
)

const serverName = "vba-ls"
const serverVersion = "0.7.1"

// 再解析の待ち時間。jsonrpc2 はハンドラを読み取りループ内で同期実行するため、
// 解析はバックグラウンドで行い、入力が続く間は最新の版だけを解析する。
const (
	defaultReparseDelay = 300 * time.Millisecond
	semanticTokensWait  = 3 * time.Second        // セマンティックトークンは最新の版の解析を待ってから返す（上限）
	definitionWait      = 2 * time.Second        // 定義ジャンプも同様
	diagnosticsRefresh  = 300 * time.Millisecond // 他モジュールの公開面が変わったときの診断の再通知のデバウンス
)

// parseMu は解析を 1 つずつ行うためのロック（開いている文書の再解析とワークスペース索引で共有）。
// ANTLR ランタイムの共有 DFA は内部でロックされているが、この環境では -race で検証できないため直列に保つ。
// 索引はファイルごとにロックを取り直すので、編集中の再解析が待たされるのは最大でファイル 1 つ分。
var parseMu sync.Mutex

// docState はオープン中の文書 1 件の状態。
// テキストは受け取った時点で最新にし、シンボル表は解析が終わった版のもの（古い版のことがある）を保持する。
type docState struct {
	gen      int // テキストの世代（didOpen / didChange ごとに増える）
	text     string
	lines    []string
	table    *SymbolTable  // 最後に解析できたシンボル表（未解析なら nil）
	tableGen int           // table の元になった世代
	timer    *time.Timer   // 予約中の再解析
	ready    chan struct{} // 最新の世代の解析待ち。解析が追いついたら close して nil にする
	entry    *ModuleEntry  // table から作った自モジュールの entry（Me. 用。*_utf 以外は nil）
	// goodTable は構文エラーの無かった最後の表。入力途中のブロックの未完（Do / For の未終端など）で
	// 後ろの宣言が表から落ちている間、識別子補完の候補を補うのに使う（解決・診断には使わない）。
	goodTable *SymbolTable
}

// lspHandler は jsonrpc2.Handler を実装する LSP サーバー。
type lspHandler struct {
	mu   sync.Mutex
	docs map[string]*docState

	types        *CompositeResolver // 全文書で共有する型解決（Workspace → VBA → TLB → Scripting → MSForms）
	ws           *Workspace         // ワークスペース索引（nil 可）
	indexCancel  context.CancelFunc
	refreshTimer *time.Timer // 他文書の診断の再通知
	debounce     time.Duration

	conn     *jsonrpc2.Conn                         // バックグラウンドからの診断の通知用
	onParsed func(uri string, gen int, latest bool) // テスト用フック（解析が終わるたびに呼ぶ）
}

func newLSPHandler(types *CompositeResolver) *lspHandler {
	if types == nil {
		types = NewDefaultResolver(nil)
	}
	return &lspHandler{
		docs:     map[string]*docState{},
		types:    types,
		debounce: defaultReparseDelay,
	}
}

// attachWorkspace はワークスペース索引をつなぐ（types は ws のプロバイダを含めて作っておくこと）。
// 公開面が変わったら、開いている文書の診断を再通知する。
func (h *lspHandler) attachWorkspace(ws *Workspace) {
	h.ws = ws
	if ws != nil {
		ws.mu.Lock()
		ws.onChange = h.scheduleDiagnosticsRefresh
		ws.mu.Unlock()
	}
}

// rootsOf は initialize の引数からワークスペースフォルダのパスを得る（workspaceFolders → rootUri → rootPath）。
func rootsOf(p InitializeParams) []string {
	var roots []string
	for _, f := range p.WorkspaceFolders {
		if path, ok := uriToPath(f.URI); ok {
			roots = append(roots, path)
		}
	}
	if len(roots) == 0 && p.RootURI != nil {
		if path, ok := uriToPath(*p.RootURI); ok {
			roots = append(roots, path)
		}
	}
	if len(roots) == 0 && p.RootPath != nil && *p.RootPath != "" {
		roots = append(roots, *p.RootPath)
	}
	return roots
}

// indexWorkspace は初回索引をバックグラウンドで行う（initialized で起動）。終わったら開いている文書の診断を再通知する。
func (h *lspHandler) indexWorkspace() {
	if h.ws == nil {
		return
	}
	ctx, cancel := context.WithCancel(context.Background())
	h.mu.Lock()
	h.indexCancel = cancel
	h.mu.Unlock()
	st := h.ws.IndexAll(ctx)
	fmt.Fprintf(os.Stderr, "vba-ls: ワークスペース索引 %d 件（std %d / class %d / form %d、失敗 %d、スキップ %d）%v\n",
		st.Parsed, st.ByKind[ModuleStd], st.ByKind[ModuleClass], st.ByKind[ModuleForm], st.Failed, st.Skipped, st.Duration.Round(time.Millisecond))
	h.refreshAllDiagnostics()
}

// scheduleDiagnosticsRefresh は他モジュールの公開面が変わったとき、開いている文書の診断の再通知を予約する。
func (h *lspHandler) scheduleDiagnosticsRefresh() {
	h.mu.Lock()
	defer h.mu.Unlock()
	if h.refreshTimer != nil {
		h.refreshTimer.Stop()
	}
	h.refreshTimer = time.AfterFunc(diagnosticsRefresh, h.refreshAllDiagnostics)
}

// refreshAllDiagnostics は解析済みの開いている文書すべての診断を通知する（再解析はしない）。
func (h *lspHandler) refreshAllDiagnostics() {
	h.mu.Lock()
	var uris []string
	for uri, st := range h.docs {
		if st.table != nil {
			uris = append(uris, uri)
		}
	}
	h.mu.Unlock()
	sort.Strings(uris)
	for _, uri := range uris {
		h.publishDiagnostics(uri)
	}
}

func (h *lspHandler) Handle(ctx context.Context, conn *jsonrpc2.Conn, req *jsonrpc2.Request) {
	h.mu.Lock()
	h.conn = conn
	h.mu.Unlock()
	switch req.Method {
	case "initialize":
		var p InitializeParams
		if unmarshalParams(req, &p) && h.ws != nil {
			h.ws.SetRoots(rootsOf(p))
		}
		_ = conn.Reply(ctx, req.ID, h.initializeResult())
	case "initialized":
		go h.indexWorkspace()
	case "shutdown":
		h.mu.Lock()
		if h.indexCancel != nil {
			h.indexCancel()
		}
		h.mu.Unlock()
		_ = conn.Reply(ctx, req.ID, nil)
	case "workspace/didChangeWatchedFiles":
		var p DidChangeWatchedFilesParams
		if unmarshalParams(req, &p) {
			h.ws.FilesChanged(p.Changes)
		}
	case "textDocument/definition":
		var p DefinitionParams
		if !unmarshalParams(req, &p) {
			_ = conn.Reply(ctx, req.ID, nil)
			return
		}
		id := req.ID
		go func() {
			h.waitParsed(p.TextDocument.URI, definitionWait)
			if locs := h.definition(p); len(locs) > 0 {
				_ = conn.Reply(context.Background(), id, locs)
				return
			}
			_ = conn.Reply(context.Background(), id, nil)
		}()
	case "exit":
		// サーバー終了
	case "textDocument/didOpen":
		var p DidOpenTextDocumentParams
		if unmarshalParams(req, &p) {
			h.didOpen(p)
		}
	case "textDocument/didChange":
		var p DidChangeTextDocumentParams
		if unmarshalParams(req, &p) && len(p.ContentChanges) > 0 {
			h.didChange(p)
		}
	case "textDocument/didClose":
		var p DidCloseTextDocumentParams
		if unmarshalParams(req, &p) {
			h.didClose(p.TextDocument.URI)
		}
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
	case "completionItem/resolve":
		var it CompletionItem
		if !unmarshalParams(req, &it) {
			_ = conn.Reply(ctx, req.ID, nil)
			return
		}
		_ = conn.Reply(ctx, req.ID, h.resolveCompletion(it))
	case "textDocument/signatureHelp":
		var p SignatureHelpParams
		if !unmarshalParams(req, &p) {
			_ = conn.Reply(ctx, req.ID, nil)
			return
		}
		if sh := h.signatureHelp(p); sh != nil {
			_ = conn.Reply(ctx, req.ID, sh)
			return
		}
		_ = conn.Reply(ctx, req.ID, nil)
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
		// 宣言位置は解析結果に依存するため、最新の版の解析を待ってから返す。
		// 読み取りループを止めないよう、待機と応答は別の goroutine で行う。
		id := req.ID
		go func() {
			h.waitParsed(p.TextDocument.URI, semanticTokensWait)
			_ = conn.Reply(context.Background(), id, SemanticTokens{Data: h.semanticTokens(p)})
		}()
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
				TriggerCharacters: []string{".", " "},
				ResolveProvider:   true,
			},
			SignatureHelpProvider: &SignatureHelpOptions{
				TriggerCharacters:   []string{"(", ",", " "},
				RetriggerCharacters: []string{")"},
			},
			HoverProvider:      true,
			DefinitionProvider: true,
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
	h.ws.OpenBuffer(p.TextDocument.URI)
	h.update(p.TextDocument.URI, p.TextDocument.Text, 0)
}

func (h *lspHandler) didChange(p DidChangeTextDocumentParams) {
	h.update(p.TextDocument.URI, p.ContentChanges[len(p.ContentChanges)-1].Text, h.debounce)
}

func (h *lspHandler) didClose(uri string) {
	h.mu.Lock()
	if st := h.docs[uri]; st != nil {
		if st.timer != nil {
			st.timer.Stop()
		}
		if st.ready != nil {
			close(st.ready)
		}
		delete(h.docs, uri)
	}
	h.mu.Unlock()
	// 索引はディスクの内容に戻す
	h.ws.CloseBuffer(uri)
}

// update はテキストを最新にし、delay 後の再解析を予約する（予約済みの古い再解析は取り消す）。
func (h *lspHandler) update(uri, text string, delay time.Duration) {
	h.mu.Lock()
	defer h.mu.Unlock()
	st := h.docs[uri]
	if st == nil {
		st = &docState{}
		h.docs[uri] = st
	}
	st.gen++
	st.text, st.lines = text, splitLines(text)
	if st.timer != nil {
		st.timer.Stop()
	}
	if st.ready == nil {
		st.ready = make(chan struct{})
	}
	gen := st.gen
	st.timer = time.AfterFunc(delay, func() { h.reparse(uri, gen) })
}

// reparse は gen が最新の世代のときだけ解析する（古い要求は破棄する）。
// 解析中に新しい版が届いた場合も、今ある表より新しければ表は更新する（補完・ホバー用）。
func (h *lspHandler) reparse(uri string, gen int) {
	h.mu.Lock()
	st := h.docs[uri]
	if st == nil || st.gen != gen {
		h.mu.Unlock()
		return
	}
	text, lines := st.text, st.lines
	h.mu.Unlock()

	parseMu.Lock()
	table := ParseSymbolsLines(text, lines)
	parseMu.Unlock()
	entry := newModuleEntryForURI(uri, text, lines, table, gen)

	h.mu.Lock()
	if h.docs[uri] != st || (st.table != nil && gen <= st.tableGen) {
		h.mu.Unlock()
		return
	}
	st.table, st.tableGen, st.entry = table, gen, entry
	if table.Stats.SyntaxErrors == 0 {
		st.goodTable = table
	}
	latest := st.gen == gen
	if latest && st.ready != nil {
		close(st.ready)
		st.ready = nil
	}
	hook := h.onParsed
	h.mu.Unlock()

	// 索引の公開面を未保存の編集で置き換える（索引対象外のファイルなら何もしない）
	h.ws.PutBuffer(entry)
	if latest {
		h.publishDiagnostics(uri)
	}
	if hook != nil {
		hook(uri, gen, latest)
	}
}

// waitParsed は uri の最新の版の解析が終わるまで（最大 timeout）待つ。追いついていれば true。
func (h *lspHandler) waitParsed(uri string, timeout time.Duration) bool {
	deadline := time.Now().Add(timeout)
	for {
		h.mu.Lock()
		st := h.docs[uri]
		if st == nil {
			h.mu.Unlock()
			return false
		}
		ch := st.ready
		h.mu.Unlock()
		if ch == nil {
			return true
		}
		remain := time.Until(deadline)
		if remain <= 0 {
			return false
		}
		select {
		case <-ch:
		case <-time.After(remain):
			return false
		}
	}
}

// document は uri の最新のテキストと、最後に解析できたシンボル表から Document を作る。開かれていなければ nil。
func (h *lspHandler) document(uri string) *Document {
	h.mu.Lock()
	st := h.docs[uri]
	if st == nil {
		h.mu.Unlock()
		return nil
	}
	text, lines, table, entry, good := st.text, st.lines, st.table, st.entry, st.goodTable
	h.mu.Unlock()
	doc := newDocumentFull(uri, text, lines, table, entry, h.types)
	if table != nil && table.Stats.SyntaxErrors > 0 && good != nil && good != table {
		doc.Fallback = good
	}
	return doc
}

func (h *lspHandler) definition(p DefinitionParams) []Location {
	doc := h.document(p.TextDocument.URI)
	if doc == nil {
		return nil
	}
	return ResolveDefinition(doc, int(p.Position.Line), int(p.Position.Character))
}

func (h *lspHandler) completion(p CompletionParams) []CompletionItem {
	doc := h.document(p.TextDocument.URI)
	if doc == nil {
		return nil
	}
	trigger := ""
	if p.Context != nil && p.Context.TriggerKind == CompletionTriggerCharacter {
		trigger = p.Context.TriggerCharacter
	}
	return BuildCompletionsTrigger(doc, int(p.Position.Line), int(p.Position.Character), trigger)
}

// resolveCompletion は一覧では省いた詳細（Detail・Documentation）を埋める（Data のある候補だけ）。
func (h *lspHandler) resolveCompletion(it CompletionItem) CompletionItem {
	return ResolveCompletionItem(h.types, it)
}

func (h *lspHandler) signatureHelp(p SignatureHelpParams) *SignatureHelp {
	doc := h.document(p.TextDocument.URI)
	if doc == nil {
		return nil
	}
	return BuildSignatureHelp(doc, int(p.Position.Line), int(p.Position.Character))
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

// publishDiagnostics は診断を通知する（バックグラウンドの解析完了時に呼ぶ）。
func (h *lspHandler) publishDiagnostics(uri string) {
	h.mu.Lock()
	conn := h.conn
	h.mu.Unlock()
	if conn == nil {
		return
	}
	diags := h.diagnostics(uri)
	if diags == nil {
		return
	}
	_ = conn.Notify(context.Background(), "textDocument/publishDiagnostics", PublishDiagnosticsParams{
		URI:         uri,
		Diagnostics: diags,
	})
}
