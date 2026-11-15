package main

// LSP で必要な型を最小限定義する。
// glsp / go-lsp に依存せず、jsonrpc2 だけで動作させる。

type DocumentURI = string

// Position は LSP の Position 型。
type Position struct {
	Line      uint32 `json:"line"`
	Character uint32 `json:"character"`
}

// Range は LSP の Range 型。
type Range struct {
	Start Position `json:"start"`
	End   Position `json:"end"`
}

// Location は LSP の Location 型。
type Location struct {
	URI   DocumentURI `json:"uri"`
	Range Range       `json:"range"`
}

// --- Initialize ---

type WorkspaceFolder struct {
	URI  string `json:"uri"`
	Name string `json:"name"`
}

type InitializeParams struct {
	RootURI               *string           `json:"rootUri,omitempty"`
	RootPath              *string           `json:"rootPath,omitempty"`
	WorkspaceFolders      []WorkspaceFolder `json:"workspaceFolders,omitempty"`
	InitializationOptions interface{}       `json:"initializationOptions,omitempty"`
}

// --- Workspace ---

type FileChangeType int

const (
	FileChangeCreated FileChangeType = 1
	FileChangeChanged FileChangeType = 2
	FileChangeDeleted FileChangeType = 3
)

type FileEvent struct {
	URI  DocumentURI    `json:"uri"`
	Type FileChangeType `json:"type"`
}

type DidChangeWatchedFilesParams struct {
	Changes []FileEvent `json:"changes"`
}

type TextDocumentSyncKind int

const (
	TextDocumentSyncKindNone        TextDocumentSyncKind = 0
	TextDocumentSyncKindFull        TextDocumentSyncKind = 1
	TextDocumentSyncKindIncremental TextDocumentSyncKind = 2
)

type CompletionOptions struct {
	TriggerCharacters []string `json:"triggerCharacters,omitempty"`
}

type SemanticTokensLegend struct {
	TokenTypes     []string `json:"tokenTypes"`
	TokenModifiers []string `json:"tokenModifiers"`
}

type SemanticTokensOptions struct {
	Legend SemanticTokensLegend `json:"legend"`
	Full   bool                 `json:"full"`
}

type ServerCapabilities struct {
	TextDocumentSync       TextDocumentSyncKind   `json:"textDocumentSync"`
	CompletionProvider     *CompletionOptions     `json:"completionProvider,omitempty"`
	HoverProvider          bool                   `json:"hoverProvider,omitempty"`
	DefinitionProvider     bool                   `json:"definitionProvider,omitempty"`
	SemanticTokensProvider *SemanticTokensOptions `json:"semanticTokensProvider,omitempty"`
}

// --- Semantic Tokens ---

type SemanticTokensParams struct {
	TextDocument TextDocumentIdentifier `json:"textDocument"`
}

type SemanticTokens struct {
	Data []uint32 `json:"data"`
}

type ServerInfo struct {
	Name    string `json:"name"`
	Version string `json:"version"`
}

type InitializeResult struct {
	Capabilities ServerCapabilities `json:"capabilities"`
	ServerInfo   *ServerInfo        `json:"serverInfo,omitempty"`
}

// --- TextDocument Sync ---

type TextDocumentItem struct {
	URI        DocumentURI `json:"uri"`
	LanguageID string      `json:"languageId"`
	Version    int         `json:"version"`
	Text       string      `json:"text"`
}

type VersionedTextDocumentIdentifier struct {
	URI     DocumentURI `json:"uri"`
	Version int         `json:"version"`
}

type TextDocumentIdentifier struct {
	URI DocumentURI `json:"uri"`
}

type TextDocumentContentChangeEvent struct {
	Text string `json:"text"`
}

type DidOpenTextDocumentParams struct {
	TextDocument TextDocumentItem `json:"textDocument"`
}

type DidChangeTextDocumentParams struct {
	TextDocument   VersionedTextDocumentIdentifier  `json:"textDocument"`
	ContentChanges []TextDocumentContentChangeEvent `json:"contentChanges"`
}

type DidCloseTextDocumentParams struct {
	TextDocument TextDocumentIdentifier `json:"textDocument"`
}

// --- Completion ---

type TextDocumentPositionParams struct {
	TextDocument TextDocumentIdentifier `json:"textDocument"`
	Position     Position               `json:"position"`
}

type CompletionParams struct {
	TextDocumentPositionParams
}

type CompletionItemKind int

const (
	CompletionItemKindMethod     CompletionItemKind = 2
	CompletionItemKindField      CompletionItemKind = 5
	CompletionItemKindProperty   CompletionItemKind = 10
	CompletionItemKindEnumMember CompletionItemKind = 20
	CompletionItemKindEvent      CompletionItemKind = 23
)

type CompletionItem struct {
	Label  string             `json:"label"`
	Kind   CompletionItemKind `json:"kind,omitempty"`
	Detail string             `json:"detail,omitempty"`
}

// --- Hover ---

type HoverParams struct {
	TextDocumentPositionParams
}

// --- Definition ---

type DefinitionParams struct {
	TextDocumentPositionParams
}

type MarkupContent struct {
	Kind  string `json:"kind"` // "plaintext" | "markdown"
	Value string `json:"value"`
}

type Hover struct {
	Contents MarkupContent `json:"contents"`
}

// --- Diagnostics ---

type DiagnosticSeverity int

const (
	DiagnosticSeverityError       DiagnosticSeverity = 1
	DiagnosticSeverityWarning     DiagnosticSeverity = 2
	DiagnosticSeverityInformation DiagnosticSeverity = 3
	DiagnosticSeverityHint        DiagnosticSeverity = 4
)

type Diagnostic struct {
	Range    Range              `json:"range"`
	Severity DiagnosticSeverity `json:"severity,omitempty"`
	Source   string             `json:"source,omitempty"`
	Message  string             `json:"message"`
}

type PublishDiagnosticsParams struct {
	URI         DocumentURI  `json:"uri"`
	Diagnostics []Diagnostic `json:"diagnostics"`
}
