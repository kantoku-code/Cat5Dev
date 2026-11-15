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
	ResolveProvider   bool     `json:"resolveProvider,omitempty"`
}

type SemanticTokensLegend struct {
	TokenTypes     []string `json:"tokenTypes"`
	TokenModifiers []string `json:"tokenModifiers"`
}

type SemanticTokensOptions struct {
	Legend SemanticTokensLegend `json:"legend"`
	Full   bool                 `json:"full"`
}

type SignatureHelpOptions struct {
	TriggerCharacters   []string `json:"triggerCharacters,omitempty"`
	RetriggerCharacters []string `json:"retriggerCharacters,omitempty"`
}

type ServerCapabilities struct {
	TextDocumentSync       TextDocumentSyncKind   `json:"textDocumentSync"`
	CompletionProvider     *CompletionOptions     `json:"completionProvider,omitempty"`
	SignatureHelpProvider  *SignatureHelpOptions  `json:"signatureHelpProvider,omitempty"`
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

// CompletionTriggerKind: 1=Invoked（Ctrl+Space・入力）, 2=TriggerCharacter, 3=TriggerForIncompleteCompletions
const CompletionTriggerCharacter = 2

type CompletionContext struct {
	TriggerKind      int    `json:"triggerKind"`
	TriggerCharacter string `json:"triggerCharacter,omitempty"`
}

type CompletionParams struct {
	TextDocumentPositionParams
	Context *CompletionContext `json:"context,omitempty"`
}

type CompletionItemKind int

const (
	CompletionItemKindMethod     CompletionItemKind = 2
	CompletionItemKindFunction   CompletionItemKind = 3
	CompletionItemKindField      CompletionItemKind = 5
	CompletionItemKindVariable   CompletionItemKind = 6
	CompletionItemKindClass      CompletionItemKind = 7
	CompletionItemKindInterface  CompletionItemKind = 8
	CompletionItemKindModule     CompletionItemKind = 9
	CompletionItemKindProperty   CompletionItemKind = 10
	CompletionItemKindEnum       CompletionItemKind = 13
	CompletionItemKindKeyword    CompletionItemKind = 14
	CompletionItemKindEnumMember CompletionItemKind = 20
	CompletionItemKindConstant   CompletionItemKind = 21
	CompletionItemKindStruct     CompletionItemKind = 22
	CompletionItemKindEvent      CompletionItemKind = 23
)

type CompletionItemLabelDetails struct {
	Detail      string `json:"detail,omitempty"`
	Description string `json:"description,omitempty"`
}

type CompletionItem struct {
	Label         string                      `json:"label"`
	LabelDetails  *CompletionItemLabelDetails `json:"labelDetails,omitempty"`
	Kind          CompletionItemKind          `json:"kind,omitempty"`
	Detail        string                      `json:"detail,omitempty"`
	Documentation *MarkupContent              `json:"documentation,omitempty"`
	SortText      string                      `json:"sortText,omitempty"`
	// Data は completionItem/resolve で詳細を引くための印（件数の多い TLB の型・Enum 定数だけに付ける）。
	// 名前は Label から引く（応答を小さくするため印だけにする）。
	Data string `json:"data,omitempty"`
}

// CompletionItem.Data の値。
const (
	completionDataType  = "t"
	completionDataConst = "c"
)

// --- Signature Help ---

type SignatureHelpContext struct {
	TriggerKind      int    `json:"triggerKind"`
	TriggerCharacter string `json:"triggerCharacter,omitempty"`
	IsRetrigger      bool   `json:"isRetrigger"`
}

type SignatureHelpParams struct {
	TextDocumentPositionParams
	Context *SignatureHelpContext `json:"context,omitempty"`
}

type ParameterInformation struct {
	Label         [2]uint32      `json:"label"` // シグネチャのラベル内の UTF-16 の [start, end)
	Documentation *MarkupContent `json:"documentation,omitempty"`
}

type SignatureInformation struct {
	Label         string                 `json:"label"`
	Documentation *MarkupContent         `json:"documentation,omitempty"`
	Parameters    []ParameterInformation `json:"parameters"`
}

type SignatureHelp struct {
	Signatures      []SignatureInformation `json:"signatures"`
	ActiveSignature uint32                 `json:"activeSignature"`
	ActiveParameter uint32                 `json:"activeParameter"`
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
