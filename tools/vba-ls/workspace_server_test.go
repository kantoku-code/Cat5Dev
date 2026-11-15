package main

import (
	"context"
	"encoding/json"
	"net"
	"os"
	"path/filepath"
	"testing"
	"time"

	"github.com/sourcegraph/jsonrpc2"
)

// wsServer はワークスペース付きのサーバーと、JSON-RPC のクライアントを用意する。
type wsServer struct {
	t      *testing.T
	h      *lspHandler
	ws     *Workspace
	client *jsonrpc2.Conn
	root   string
	dir    string
	diags  chan PublishDiagnosticsParams
}

func newWSServer(t *testing.T, files map[string]string, parse func(string, []string) *SymbolTable) *wsServer {
	t.Helper()
	root := t.TempDir()
	dir := filepath.Join(root, "modules")
	if err := os.MkdirAll(dir, 0o755); err != nil {
		t.Fatal(err)
	}
	for name, body := range files {
		if err := os.WriteFile(filepath.Join(dir, name), []byte(body), 0o644); err != nil {
			t.Fatal(err)
		}
	}
	ws := NewWorkspace()
	if parse != nil {
		ws.parse = parse
	}
	h := newLSPHandler(NewDefaultResolver(newTestTLBDatabase(t), ws.TypesProvider(), ws.ModulesProvider()))
	h.attachWorkspace(ws)
	h.debounce = 20 * time.Millisecond
	ctx := context.Background()
	sc, cc := net.Pipe()
	server := jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(sc, jsonrpc2.VSCodeObjectCodec{}), h)
	s := &wsServer{t: t, h: h, ws: ws, root: root, dir: dir, diags: make(chan PublishDiagnosticsParams, 100)}
	s.client = jsonrpc2.NewConn(ctx, jsonrpc2.NewBufferedStream(cc, jsonrpc2.VSCodeObjectCodec{}),
		jsonrpc2.HandlerWithError(func(_ context.Context, _ *jsonrpc2.Conn, req *jsonrpc2.Request) (interface{}, error) {
			if req.Method == "textDocument/publishDiagnostics" && req.Params != nil {
				var p PublishDiagnosticsParams
				if json.Unmarshal(*req.Params, &p) == nil {
					select {
					case s.diags <- p:
					default:
					}
				}
			}
			return nil, nil
		}))
	t.Cleanup(func() {
		_ = s.client.Close()
		_ = server.Close()
	})
	return s
}

func (s *wsServer) uri(name string) string { return pathToURI(filepath.Join(s.dir, name)) }

func (s *wsServer) initialize() InitializeResult {
	s.t.Helper()
	var init InitializeResult
	params := InitializeParams{WorkspaceFolders: []WorkspaceFolder{{URI: pathToURI(s.root), Name: "ws"}}}
	if err := s.client.Call(context.Background(), "initialize", params, &init); err != nil {
		s.t.Fatal(err)
	}
	if err := s.client.Notify(context.Background(), "initialized", map[string]interface{}{}); err != nil {
		s.t.Fatal(err)
	}
	return init
}

func (s *wsServer) waitReady() {
	s.t.Helper()
	select {
	case <-s.ws.Ready():
	case <-time.After(time.Minute):
		s.t.Fatal("索引が終わらない")
	}
}

func (s *wsServer) notify(method string, params interface{}) {
	s.t.Helper()
	if err := s.client.Notify(context.Background(), method, params); err != nil {
		s.t.Fatal(err)
	}
}

func (s *wsServer) open(name, text string) {
	s.notify("textDocument/didOpen", DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: s.uri(name), Text: text}})
}

func (s *wsServer) complete(name string, line, char int) map[string]bool {
	s.t.Helper()
	var items []CompletionItem
	if err := s.client.Call(context.Background(), "textDocument/completion", completionParams(s.uri(name), line, char), &items); err != nil {
		s.t.Fatal(err)
	}
	return labelsOf(items)
}

// eventually は cond が真になるまで（最大 10 秒）待つ。
func eventually(t *testing.T, what string, cond func() bool) {
	t.Helper()
	deadline := time.Now().Add(10 * time.Second)
	for !cond() {
		if time.Now().After(deadline) {
			t.Fatalf("%s にならない", what)
		}
		time.Sleep(10 * time.Millisecond)
	}
}

const module1User = "Sub T()\n    Dim o As clsA\n    o.\nEnd Sub\n"

func TestWorkspaceServerIndexAndBuffers(t *testing.T) {
	s := newWSServer(t, wsFixture, nil)
	init := s.initialize()
	if !init.Capabilities.DefinitionProvider {
		t.Error("definitionProvider が true でない")
	}
	s.waitReady()
	s.open("Other.bas_utf", module1User)
	if !s.h.waitParsed(s.uri("Other.bas_utf"), time.Minute) {
		t.Fatal("解析が終わらない")
	}
	if !s.complete("Other.bas_utf", 2, len("    o."))["Name"] {
		t.Fatal("索引したクラスのメンバーが補完されない")
	}

	// 未保存の編集が他モジュールの補完に反映される
	clsA := wsFixture["clsA.cls_utf"]
	s.open("clsA.cls_utf", clsA)
	s.notify("textDocument/didChange", changeParams(s.uri("clsA.cls_utf"), clsA+"\nPublic Sub NewOne()\nEnd Sub\n"))
	eventually(t, "NewOne が補完される", func() bool { return s.complete("Other.bas_utf", 2, len("    o."))["NewOne"] })

	// 開いたままディスクが変わってもバッファが勝つ
	if err := os.WriteFile(filepath.Join(s.dir, "clsA.cls_utf"), []byte(clsA+"\nPublic Sub DiskOnly()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	s.notify("workspace/didChangeWatchedFiles", DidChangeWatchedFilesParams{Changes: []FileEvent{{URI: s.uri("clsA.cls_utf"), Type: FileChangeChanged}}})
	time.Sleep(4 * watchDebounce)
	if l := s.complete("Other.bas_utf", 2, len("    o.")); l["DiskOnly"] || !l["NewOne"] {
		t.Errorf("開いているファイルはバッファ優先のはず: %v", l)
	}

	// 閉じるとディスクの内容に戻る
	s.notify("textDocument/didClose", DidCloseTextDocumentParams{TextDocument: TextDocumentIdentifier{URI: s.uri("clsA.cls_utf")}})
	eventually(t, "ディスクの内容に戻る", func() bool {
		l := s.complete("Other.bas_utf", 2, len("    o."))
		return l["DiskOnly"] && !l["NewOne"]
	})

	// definition: 他ファイルの宣言
	var locs []Location
	src := "Sub T()\n    Dim o As clsA\n    o.Clone\nEnd Sub\n"
	s.notify("textDocument/didChange", changeParams(s.uri("Other.bas_utf"), src))
	var dp DefinitionParams
	dp.TextDocument.URI = s.uri("Other.bas_utf")
	dp.Position = Position{Line: 2, Character: uint32(len("    o.Cl"))}
	if err := s.client.Call(context.Background(), "textDocument/definition", dp, &locs); err != nil {
		t.Fatal(err)
	}
	if len(locs) != 1 || locs[0].URI != s.uri("clsA.cls_utf") {
		t.Errorf("definition = %+v", locs)
	}

	// git: のバッファは索引に登録しない
	s.open("x", "")
	gitURI := "git:/" + filepath.ToSlash(filepath.Join(s.dir, "Module1.bas_utf")) + "?ref"
	s.notify("textDocument/didOpen", DidOpenTextDocumentParams{TextDocument: TextDocumentItem{URI: gitURI, Text: "Public Sub GitOnly()\nEnd Sub\n"}})
	s.h.waitParsed(gitURI, time.Minute)
	time.Sleep(50 * time.Millisecond)
	if _, m, ok := s.h.types.ProjectMember("GitOnly"); ok {
		t.Errorf("git: のバッファが索引に入った: %+v", m)
	}
}

func TestWorkspaceServerWatchedFiles(t *testing.T) {
	s := newWSServer(t, wsFixture, nil)
	s.initialize()
	s.waitReady()
	// 構文エラーで後続の宣言が落ちないよう、補完位置の後ろにも名前を書いておく
	s.open("Other.bas_utf", "Sub T()\n    Dim o As clsA\n    o.X\n    Dim b As clsB\n    b.X\n    Dim c As clsC\n    c.X\nEnd Sub\n")
	s.h.waitParsed(s.uri("Other.bas_utf"), time.Minute)

	// 作成
	if err := os.WriteFile(filepath.Join(s.dir, "clsB.cls_utf"), []byte("Public Sub BMethod()\nEnd Sub\n"), 0o644); err != nil {
		t.Fatal(err)
	}
	s.notify("workspace/didChangeWatchedFiles", DidChangeWatchedFilesParams{Changes: []FileEvent{{URI: s.uri("clsB.cls_utf"), Type: FileChangeCreated}}})
	eventually(t, "作成したクラスが補完される", func() bool { return s.complete("Other.bas_utf", 4, len("    b."))["BMethod"] })

	// 削除
	if err := os.Remove(filepath.Join(s.dir, "clsA.cls_utf")); err != nil {
		t.Fatal(err)
	}
	s.notify("workspace/didChangeWatchedFiles", DidChangeWatchedFilesParams{Changes: []FileEvent{{URI: s.uri("clsA.cls_utf"), Type: FileChangeDeleted}}})
	eventually(t, "削除したクラスが消える", func() bool { return len(s.complete("Other.bas_utf", 2, len("    o."))) == 0 })

	// リネーム（Deleted + Created）
	if err := os.Rename(filepath.Join(s.dir, "clsB.cls_utf"), filepath.Join(s.dir, "clsC.cls_utf")); err != nil {
		t.Fatal(err)
	}
	s.notify("workspace/didChangeWatchedFiles", DidChangeWatchedFilesParams{Changes: []FileEvent{
		{URI: s.uri("clsB.cls_utf"), Type: FileChangeDeleted}, {URI: s.uri("clsC.cls_utf"), Type: FileChangeCreated},
	}})
	eventually(t, "リネーム後の名前だけが解決する", func() bool {
		return s.complete("Other.bas_utf", 6, len("    c."))["BMethod"] && len(s.complete("Other.bas_utf", 4, len("    b."))) == 0
	})
}

// 索引の完了前でも要求に即座に応答し、完了後に開いている文書の診断が再通知される。
func TestWorkspaceServerBeforeIndex(t *testing.T) {
	release := make(chan struct{})
	blocked := func(text string, lines []string) *SymbolTable {
		<-release
		return ParseSymbolsLines(text, lines)
	}
	s := newWSServer(t, wsFixture, blocked)
	s.initialize()
	src := "Sub T()\n    Dim o As clsA\n    o.Bogus\n    o.\nEnd Sub\n"
	uri := s.uri("Other.bas_utf")
	s.open("Other.bas_utf", src)
	s.h.waitParsed(uri, time.Minute)
	first := <-s.diags
	if len(first.Diagnostics) != 0 {
		t.Errorf("索引前はクラスが分からないので診断なし: %+v", first.Diagnostics)
	}

	done := make(chan struct{})
	go func() {
		defer close(done)
		ctx := context.Background()
		var items []CompletionItem
		_ = s.client.Call(ctx, "textDocument/completion", completionParams(uri, 3, len("    o.")), &items)
		var hv *Hover
		var hp HoverParams
		hp.TextDocument.URI, hp.Position = uri, Position{Line: 2, Character: 6}
		_ = s.client.Call(ctx, "textDocument/hover", hp, &hv)
		var locs []Location
		var dp DefinitionParams
		dp.TextDocument.URI, dp.Position = uri, Position{Line: 2, Character: 6}
		_ = s.client.Call(ctx, "textDocument/definition", dp, &locs)
		var toks SemanticTokens
		_ = s.client.Call(ctx, "textDocument/semanticTokens/full", SemanticTokensParams{TextDocument: TextDocumentIdentifier{URI: uri}}, &toks)
	}()
	select {
	case <-done:
	case <-time.After(10 * time.Second):
		t.Fatal("索引中に要求が応答しない")
	}
	close(release)
	s.waitReady()
	deadline := time.After(10 * time.Second)
	for {
		select {
		case p := <-s.diags:
			if p.URI == uri && len(p.Diagnostics) == 1 && p.Diagnostics[0].Message == "'clsA' に 'Bogus' は存在しません" {
				return
			}
		case <-deadline:
			t.Fatal("索引の完了後に診断が再通知されない")
		}
	}
}
