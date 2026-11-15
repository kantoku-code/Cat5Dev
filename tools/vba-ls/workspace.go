package main

import (
	"context"
	"fmt"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"sync/atomic"
	"time"
)

// 索引の上限（巨大なフォルダを開いたときの保護）。
const (
	maxIndexFiles    = 2000
	maxIndexFileSize = 8 << 20
	watchDebounce    = 200 * time.Millisecond
	installBatch     = 8 // 初回索引で何ファイルごとにスナップショットを作り直すか
)

// Workspace はワークスペースのモジュール索引。
// 読み取りは不変のスナップショットを atomic に差し替える方式で、プロバイダのメソッドはロックなしで並行に呼べる。
// 更新は mu で直列化する。ロック順は lspHandler.mu → Workspace.mu（Workspace は mu を持ったまま外を呼ばない）。
type Workspace struct {
	mu        sync.Mutex
	snap      atomic.Pointer[wsSnapshot]
	dirs      []string          // 索引するディレクトリ（絶対パス。順序＝優先順）
	dirKeys   []string          // dirs の pathKey
	open      map[string]string // Key → 開いているバッファの URI（バッファの内容がディスクより優先）
	ready     chan struct{}     // 初回索引の完了で close
	readyOnce sync.Once
	pending   map[string]string // didChangeWatchedFiles で再読込するキー → パス
	timer     *time.Timer       // watched files の Reconcile のデバウンス
	dupWarned map[string]bool

	// parse は解析関数（既定は parseMu を取って ParseSymbolsLines）。テストで遅延・失敗を差し込む。
	parse func(text string, lines []string) *SymbolTable
	// onChange は公開面が変わったときに呼ぶ（mu の外で呼ぶ）。サーバーが他の文書の診断の再通知に使う。
	onChange func()
	// onInstalled はスナップショットを差し替えるたびに呼ぶ（テスト用、mu の外）。
	onInstalled func()
}

type wsSnapshot struct {
	entries map[string]*ModuleEntry // Key → entry（同名の重複も含む全ファイル）
	byName  map[string]*ModuleEntry // 小文字のモジュール名 → 採用した entry
	types   map[string]wsType       // 小文字の型名 → クラス / フォーム、または Public UDT / Enum
	members map[string]wsProjMember // 小文字名 → 標準モジュールの Public メンバー（修飾なしで参照できるもの）
	order   []*ModuleEntry          // byName の値をモジュール名順に
}

type wsType struct {
	entry *ModuleEntry
	udt   *UDTInfo
	enum  *EnumInfo
	sym   *Symbol // UDT / Enum の宣言
}

type wsProjMember struct {
	entry *ModuleEntry
	idx   int // entry.public.list の添字
}

var emptySnapshot = &wsSnapshot{
	entries: map[string]*ModuleEntry{}, byName: map[string]*ModuleEntry{},
	types: map[string]wsType{}, members: map[string]wsProjMember{},
}

// IndexStats は初回索引の統計。
type IndexStats struct {
	Files, Parsed, Skipped, Failed int
	ByKind                         [3]int
	Duration                       time.Duration
}

func NewWorkspace() *Workspace {
	w := &Workspace{
		open:      map[string]string{},
		ready:     make(chan struct{}),
		pending:   map[string]string{},
		dupWarned: map[string]bool{},
		parse: func(text string, lines []string) *SymbolTable {
			parseMu.Lock()
			defer parseMu.Unlock()
			return ParseSymbolsLines(text, lines)
		},
	}
	w.snap.Store(emptySnapshot)
	return w
}

// stripBOM は先頭の UTF-8 BOM（U+FEFF）を除く。
func stripBOM(s string) string { return strings.TrimPrefix(s, "\xef\xbb\xbf") }

func (w *Workspace) snapshot() *wsSnapshot {
	if w == nil {
		return emptySnapshot
	}
	if s := w.snap.Load(); s != nil {
		return s
	}
	return emptySnapshot
}

// SetRoots はワークスペースフォルダ（絶対パス）から索引するディレクトリを決める。
func (w *Workspace) SetRoots(folders []string) {
	dirs := indexDirs(folders)
	w.mu.Lock()
	defer w.mu.Unlock()
	w.dirs = dirs
	w.dirKeys = make([]string, len(dirs))
	for i, d := range dirs {
		w.dirKeys[i] = pathKey(d)
	}
}

// indexDirs は各フォルダについて、<f>/modules がディレクトリならそこ、なければ <f> を返す（どちらも非再帰で走査する）。
// 拡張機能の Pull / Push と同じ規則（modules 以外の temp/ などに置かれた別モジュールを混ぜない）。
func indexDirs(folders []string) []string {
	var out []string
	seen := map[string]bool{}
	for _, f := range folders {
		if f == "" {
			continue
		}
		d := f
		if st, err := os.Stat(filepath.Join(f, "modules")); err == nil && st.IsDir() {
			d = filepath.Join(f, "modules")
		}
		if k := pathKey(d); !seen[k] {
			seen[k] = true
			out = append(out, filepath.Clean(d))
		}
	}
	return out
}

// dirIndex は path の親ディレクトリが何番目の索引ディレクトリかを返す（含まれなければ -1）。mu の下で呼ぶ。
func (w *Workspace) dirIndex(path string) int {
	dk := pathKey(filepath.Dir(path))
	for i, k := range w.dirKeys {
		if k == dk {
			return i
		}
	}
	return -1
}

// Contains は path が索引ディレクトリ直下のモジュールファイル（*_utf）かを返す。
func (w *Workspace) Contains(path string) bool {
	if w == nil || path == "" {
		return false
	}
	if _, _, ok := moduleIdentity(path); !ok {
		return false
	}
	w.mu.Lock()
	defer w.mu.Unlock()
	return w.dirIndex(path) >= 0
}

// Ready は初回索引が終わると close されるチャネルを返す。
func (w *Workspace) Ready() <-chan struct{} { return w.ready }

type diskFile struct {
	path string
	key  string
	stat fileStat
}

// listFiles は索引ディレクトリ直下のモジュールファイルを (ディレクトリの順, キー) の順に返す。
func (w *Workspace) listFiles() ([]diskFile, int) {
	w.mu.Lock()
	dirs := append([]string(nil), w.dirs...)
	w.mu.Unlock()
	var out []diskFile
	skipped := 0
	for _, d := range dirs {
		ents, err := os.ReadDir(d)
		if err != nil {
			continue
		}
		var files []diskFile
		for _, de := range ents {
			if de.IsDir() {
				continue
			}
			p := filepath.Join(d, de.Name())
			if _, _, ok := moduleIdentity(p); !ok {
				continue
			}
			info, err := de.Info()
			if err != nil || !info.Mode().IsRegular() {
				continue
			}
			files = append(files, diskFile{path: p, key: pathKey(p), stat: fileStat{Size: info.Size(), ModTime: info.ModTime()}})
		}
		sort.Slice(files, func(i, j int) bool { return files[i].key < files[j].key })
		out = append(out, files...)
	}
	if len(out) > maxIndexFiles {
		skipped = len(out) - maxIndexFiles
		out = out[:maxIndexFiles]
	}
	return out, skipped
}

// loadDisk はファイルを読んで解析し、ディスク由来の entry を作る。
func (w *Workspace) loadDisk(f diskFile) (*ModuleEntry, error) {
	if f.stat.Size > maxIndexFileSize {
		return nil, fmt.Errorf("%s: %d バイトは上限を超えている", f.path, f.stat.Size)
	}
	data, err := os.ReadFile(f.path)
	if err != nil {
		return nil, err
	}
	text := stripBOM(string(data))
	name, kind, _ := moduleIdentity(f.path)
	lines := splitLines(text)
	table := w.parse(text, lines)
	e := newModuleEntry(name, kind, f.path, pathToURI(f.path), text, lines, table)
	e.Stat = f.stat
	return e, nil
}

// IndexAll はワークスペースを初回索引する（呼び出し側が goroutine で起動する）。
// 開いているバッファのファイルは飛ばす（バッファ側の PutBuffer で入る）。終わると Ready を close する。
func (w *Workspace) IndexAll(ctx context.Context) IndexStats {
	start := time.Now()
	defer w.readyOnce.Do(func() { close(w.ready) })
	files, skipped := w.listFiles()
	st := IndexStats{Files: len(files) + skipped, Skipped: skipped}
	var batch []*ModuleEntry
	lastInstall := time.Now()
	for _, f := range files {
		if ctx.Err() != nil {
			break
		}
		if w.isOpen(f.key) {
			st.Skipped++
			continue
		}
		e, err := w.loadDisk(f)
		if err != nil {
			st.Failed++
			fmt.Fprintf(os.Stderr, "vba-ls: 索引できないファイル: %v\n", err)
			continue
		}
		st.Parsed++
		st.ByKind[e.Kind]++
		batch = append(batch, e)
		if len(batch) >= installBatch || time.Since(lastInstall) > time.Second {
			w.install(batch, nil)
			batch, lastInstall = nil, time.Now()
		}
	}
	w.install(batch, nil)
	st.Duration = time.Since(start)
	return st
}

func (w *Workspace) isOpen(key string) bool {
	w.mu.Lock()
	defer w.mu.Unlock()
	return w.open[key] != ""
}

// OpenBuffer は didOpen で呼ぶ。索引対象のファイルなら、以降はバッファの内容を優先する。
func (w *Workspace) OpenBuffer(uri string) {
	if w == nil {
		return
	}
	p, ok := uriToPath(uri)
	if !ok || !w.Contains(p) {
		return
	}
	w.mu.Lock()
	w.open[pathKey(p)] = uri
	w.mu.Unlock()
}

// PutBuffer は開いている文書の再解析が終わるたびに呼ぶ。古い世代は捨てる。
func (w *Workspace) PutBuffer(e *ModuleEntry) {
	if w == nil || e == nil || e.Path == "" {
		return
	}
	w.mu.Lock()
	if w.open[e.Key] != e.URI {
		w.mu.Unlock()
		return // 閉じられた・索引対象外
	}
	if cur := w.snapshot().entries[e.Key]; cur != nil && cur.FromBuffer && cur.BufGen >= e.BufGen {
		w.mu.Unlock()
		return
	}
	w.mu.Unlock()
	w.install([]*ModuleEntry{e}, nil)
}

// CloseBuffer は didClose で呼ぶ。ディスクの内容に戻す（バックグラウンドで再読込する）。
func (w *Workspace) CloseBuffer(uri string) {
	if w == nil {
		return
	}
	key, ok := uriKey(uri)
	if !ok {
		return
	}
	w.mu.Lock()
	if w.open[key] != uri {
		w.mu.Unlock()
		return
	}
	delete(w.open, key)
	path := ""
	if cur := w.snapshot().entries[key]; cur != nil {
		path = cur.Path
		if cur.FromBuffer {
			// すぐに開き直されたとき、新しいバッファの世代（1 から数え直す）が古い世代に負けないよう、
			// 再読込が終わるまでの間もディスク由来扱いにしておく
			c := *cur
			c.FromBuffer, c.BufGen = false, 0
			old := w.snapshot()
			entries := make(map[string]*ModuleEntry, len(old.entries))
			for k, e := range old.entries {
				entries[k] = e
			}
			entries[key] = &c
			w.snap.Store(w.buildSnapshot(entries))
		}
	}
	w.mu.Unlock()
	if path == "" {
		path, _ = uriToPath(uri)
	}
	go w.reloadDisk(key, path)
}

// reloadDisk は key のファイルをディスクから読み直す（閉じた文書・ファイル変更）。ファイルが無ければ索引から外す。
func (w *Workspace) reloadDisk(key, path string) {
	info, err := os.Stat(path)
	if err != nil || !info.Mode().IsRegular() {
		w.install(nil, []string{key})
		return
	}
	f := diskFile{path: path, key: key, stat: fileStat{Size: info.Size(), ModTime: info.ModTime()}}
	if cur := w.snapshot().entries[key]; cur != nil && f.stat.Size <= maxIndexFileSize {
		if data, err := os.ReadFile(path); err == nil {
			text := stripBOM(string(data))
			if textHash(text) == cur.TextHash {
				// 内容が同じなら解析結果を流用する（URI だけディスクの形に付け替える）
				e := newModuleEntry(cur.Name, cur.Kind, path, pathToURI(path), text, splitLines(text), cur.Table)
				e.Stat = f.stat
				w.install([]*ModuleEntry{e}, nil)
				return
			}
		}
	}
	e, err := w.loadDisk(f)
	if err != nil {
		fmt.Fprintf(os.Stderr, "vba-ls: 索引できないファイル: %v\n", err)
		return // 読めなければ旧エントリを維持する
	}
	w.install([]*ModuleEntry{e}, nil)
}

// FilesChanged は workspace/didChangeWatchedFiles で呼ぶ。対象のファイルをまとめてデバウンスし、Reconcile する。
func (w *Workspace) FilesChanged(changes []FileEvent) {
	if w == nil {
		return
	}
	w.mu.Lock()
	defer w.mu.Unlock()
	n := 0
	for _, c := range changes {
		p, ok := uriToPath(c.URI)
		if !ok {
			continue
		}
		if _, _, ok := moduleIdentity(p); !ok || w.dirIndex(p) < 0 {
			continue
		}
		w.pending[pathKey(p)] = p
		n++
	}
	if n == 0 {
		return
	}
	if w.timer != nil {
		w.timer.Stop()
	}
	w.timer = time.AfterFunc(watchDebounce, func() {
		w.mu.Lock()
		force := w.pending
		w.pending = map[string]string{}
		w.mu.Unlock()
		w.Reconcile(context.Background(), force)
	})
}

// Reconcile はディスクの一覧と索引を突き合わせる。新規・サイズや更新時刻が違う・force のファイルを読み直し、
// 一覧から消えたファイルを外す。開いているファイルはバッファを優先して触らない。
func (w *Workspace) Reconcile(ctx context.Context, force map[string]string) {
	files, _ := w.listFiles()
	snap := w.snapshot()
	onDisk := map[string]bool{}
	var changed []*ModuleEntry
	for _, f := range files {
		if ctx.Err() != nil {
			return
		}
		onDisk[f.key] = true
		if w.isOpen(f.key) {
			continue
		}
		cur := snap.entries[f.key]
		_, forced := force[f.key]
		if cur != nil && !cur.FromBuffer && !forced && cur.Stat.Size == f.stat.Size && cur.Stat.ModTime.Equal(f.stat.ModTime) {
			continue
		}
		e, err := w.loadDisk(f)
		if err != nil {
			fmt.Fprintf(os.Stderr, "vba-ls: 索引できないファイル: %v\n", err)
			continue
		}
		changed = append(changed, e)
	}
	var removed []string
	for key := range snap.entries {
		if !onDisk[key] && !w.isOpen(key) {
			removed = append(removed, key)
		}
	}
	w.install(changed, removed)
}

// install は entry の追加・置換と削除をまとめてスナップショットに反映する。
// ディスク由来の entry は、そのファイルがバッファで開かれていれば捨てる（バッファ優先）。
func (w *Workspace) install(changed []*ModuleEntry, removed []string) {
	if len(changed) == 0 && len(removed) == 0 {
		return
	}
	w.mu.Lock()
	old := w.snapshot()
	entries := make(map[string]*ModuleEntry, len(old.entries)+len(changed))
	for k, e := range old.entries {
		entries[k] = e
	}
	surfaceChanged := false
	for _, key := range removed {
		if _, ok := entries[key]; ok {
			delete(entries, key)
			surfaceChanged = true
		}
	}
	for _, e := range changed {
		if !e.FromBuffer && w.open[e.Key] != "" {
			continue
		}
		if e.FromBuffer && w.open[e.Key] != e.URI {
			continue
		}
		prev := entries[e.Key]
		if prev == nil || prev.surface != e.surface || prev.Name != e.Name || prev.URI != e.URI ||
			prev.Table.Stats.SyntaxErrors != e.Table.Stats.SyntaxErrors || typesSurface(prev) != typesSurface(e) {
			surfaceChanged = true
		}
		entries[e.Key] = e
	}
	ns := w.buildSnapshot(entries)
	w.snap.Store(ns)
	onChange, onInstalled := w.onChange, w.onInstalled
	w.mu.Unlock()
	if surfaceChanged && onChange != nil {
		onChange()
	}
	if onInstalled != nil {
		onInstalled()
	}
}

// typesSurface は Public UDT / Enum の名前とメンバーを表す文字列（公開面の変化の判定用）。
func typesSurface(e *ModuleEntry) string {
	var sb strings.Builder
	for _, u := range e.Table.UDTs {
		if u.Visibility != VisPrivate {
			sb.WriteString("T:" + u.Name)
			for _, m := range u.Members {
				sb.WriteString("," + m.Name + ":" + m.QualifiedTypeName())
			}
			sb.WriteByte(';')
		}
	}
	for _, en := range e.Table.Enums {
		if en.Visibility != VisPrivate {
			sb.WriteString("E:" + en.Name)
			for _, m := range en.Members {
				sb.WriteString("," + m.Name)
			}
			sb.WriteByte(';')
		}
	}
	return sb.String()
}

// buildSnapshot は entries から名前の表を作る。mu の下で呼ぶ。
// 同名のモジュールは (ディレクトリの順, キー) で先のものを採用し、重複は stderr に 1 回だけ出す。
func (w *Workspace) buildSnapshot(entries map[string]*ModuleEntry) *wsSnapshot {
	all := make([]*ModuleEntry, 0, len(entries))
	for _, e := range entries {
		all = append(all, e)
	}
	sort.Slice(all, func(i, j int) bool {
		di, dj := w.dirIndex(all[i].Path), w.dirIndex(all[j].Path)
		if di != dj {
			return di < dj
		}
		return all[i].Key < all[j].Key
	})
	ns := &wsSnapshot{
		entries: entries,
		byName:  map[string]*ModuleEntry{},
		types:   map[string]wsType{},
		members: map[string]wsProjMember{},
	}
	for _, e := range all {
		k := strings.ToLower(e.Name)
		if prev, dup := ns.byName[k]; dup {
			if !w.dupWarned[e.Key] {
				w.dupWarned[e.Key] = true
				fmt.Fprintf(os.Stderr, "vba-ls: モジュール名 '%s' が重複（採用: %s, 無視: %s）\n", e.Name, prev.Path, e.Path)
			}
			continue
		}
		ns.byName[k] = e
		ns.order = append(ns.order, e)
	}
	sort.SliceStable(ns.order, func(i, j int) bool { return strings.ToLower(ns.order[i].Name) < strings.ToLower(ns.order[j].Name) })
	// 型: クラス / フォームが先、次に Public UDT / Enum（まだ無い名前だけ）
	for _, e := range ns.order {
		if e.Kind != ModuleStd {
			ns.types[strings.ToLower(e.Name)] = wsType{entry: e}
		}
	}
	for _, e := range ns.order {
		for _, u := range e.Table.UDTs {
			k := strings.ToLower(u.Name)
			if _, ok := ns.types[k]; u.Visibility == VisPrivate || ok {
				continue
			}
			ns.types[k] = wsType{entry: e, udt: u, sym: declSymbol(e.Table, SymUDT, u.Name)}
		}
		for _, en := range e.Table.Enums {
			k := strings.ToLower(en.Name)
			if _, ok := ns.types[k]; en.Visibility == VisPrivate || ok {
				continue
			}
			ns.types[k] = wsType{entry: e, enum: en, sym: declSymbol(e.Table, SymEnum, en.Name)}
		}
	}
	// 標準モジュールの Public メンバー（修飾なしで参照できるもの）
	for _, e := range ns.order {
		if e.Kind != ModuleStd {
			continue
		}
		for i, m := range e.public.list {
			k := strings.ToLower(m.Name)
			if _, ok := ns.members[k]; !ok {
				ns.members[k] = wsProjMember{entry: e, idx: i}
			}
		}
	}
	return ns
}

// declSymbol はモジュールレベルの kind / name の宣言シンボルを返す。
func declSymbol(st *SymbolTable, kind SymbolKind, name string) *Symbol {
	for i := range st.Symbols {
		s := &st.Symbols[i]
		if s.Kind == kind && s.Scope == "module" && strings.EqualFold(s.Name, name) {
			return s
		}
	}
	return nil
}

// Modules は採用されたモジュールをモジュール名順に返す（2-2 の識別子補完用）。
func (w *Workspace) Modules() []*ModuleEntry {
	return append([]*ModuleEntry(nil), w.snapshot().order...)
}

// ProjectMembers は修飾なしで参照できる標準モジュールの Public メンバーを返す（2-2 の識別子補完用）。
func (w *Workspace) ProjectMembers() []MemberInfo {
	s := w.snapshot()
	out := make([]MemberInfo, 0, len(s.members))
	for _, pm := range s.members {
		out = append(out, pm.entry.public.list[pm.idx])
	}
	sort.Slice(out, func(i, j int) bool { return strings.ToLower(out[i].Name) < strings.ToLower(out[j].Name) })
	return out
}

// Entry は path のファイルの現在の entry を返す（無ければ nil）。
func (w *Workspace) Entry(path string) *ModuleEntry { return w.snapshot().entries[pathKey(path)] }

// TypesProvider はクラス・フォーム・Public UDT / Enum を提供するプロバイダを返す。
func (w *Workspace) TypesProvider() TypeProvider { return &wsTypeProvider{w: w} }

// ModulesProvider は標準モジュール（moduleLib 修飾専用）を提供するプロバイダを返す。
func (w *Workspace) ModulesProvider() TypeProvider { return &wsModuleProvider{w: w} }

// --- wsTypeProvider ---

// wsTypeProvider はワークスペースのクラス・フォーム（Public メンバーのみ）と Public UDT / Enum を提供する。
// ProjectProvider（Main. / frmMenu. / 非修飾の Public）と TypeLister も実装する。
type wsTypeProvider struct{ w *Workspace }

func (p *wsTypeProvider) OwnsLibrary(string) bool { return false }

func (p *wsTypeProvider) LookupType(name string) (TypeInfo, bool) {
	t, ok := p.w.snapshot().types[strings.ToLower(name)]
	if !ok {
		return TypeInfo{}, false
	}
	return t.info(), true
}

func (t wsType) info() TypeInfo {
	e := t.entry
	switch {
	case t.udt != nil:
		return TypeInfo{Ref: TypeRef{Name: t.udt.Name}, Kind: TypeKindUDT, Complete: true, Doc: t.udt.Doc, Origin: t.origin()}
	case t.enum != nil:
		return TypeInfo{Ref: TypeRef{Name: t.enum.Name}, Kind: TypeKindEnum, Complete: true, Doc: t.enum.Doc, Origin: t.origin()}
	}
	return e.typeInfo()
}

func (t wsType) origin() *DeclOrigin {
	o := &DeclOrigin{URI: t.entry.URI, Module: t.entry.Name}
	if t.sym != nil {
		o.Syms = []*Symbol{t.sym}
	}
	return o
}

func (p *wsTypeProvider) Members(ref TypeRef) []MemberInfo {
	t, ok := p.w.snapshot().types[strings.ToLower(ref.Name)]
	if !ok {
		return nil
	}
	if t.udt != nil || t.enum != nil {
		return (&moduleTypeProvider{table: t.entry.Table}).Members(ref)
	}
	return t.entry.public.list
}

func (p *wsTypeProvider) Member(ref TypeRef, name string) (MemberInfo, bool) {
	t, ok := p.w.snapshot().types[strings.ToLower(ref.Name)]
	if !ok {
		return MemberInfo{}, false
	}
	if t.udt != nil || t.enum != nil {
		return (&moduleTypeProvider{table: t.entry.Table}).Member(ref, name)
	}
	if i, ok := t.entry.public.lookup(name); ok {
		return t.entry.public.list[i], true
	}
	return MemberInfo{}, false
}

func (p *wsTypeProvider) MemberType(ref TypeRef, name string) (TypeRef, bool) {
	t, ok := p.w.snapshot().types[strings.ToLower(ref.Name)]
	if !ok {
		return TypeRef{}, false
	}
	if t.udt != nil || t.enum != nil {
		return (&moduleTypeProvider{table: t.entry.Table}).MemberType(ref, name)
	}
	if i, ok := t.entry.public.lookup(name); ok {
		mt := t.entry.public.types[i]
		return mt, !mt.IsZero()
	}
	return TypeRef{}, false
}

// ModuleRef は標準モジュール名を moduleLib 修飾の型に、フォーム名を既定インスタンスの型に解決する。
// クラスは既定インスタンスを持たない（VB_PredeclaredId はファイルから分からない）ので解決しない。
func (p *wsTypeProvider) ModuleRef(name string) (TypeRef, bool) {
	e, ok := p.w.snapshot().byName[strings.ToLower(name)]
	if !ok || e.Kind == ModuleClass {
		return TypeRef{}, false
	}
	return e.selfRef(), true
}

func (p *wsTypeProvider) ProjectMember(name string) (TypeRef, MemberInfo, bool) {
	pm, ok := p.w.snapshot().members[strings.ToLower(name)]
	if !ok {
		return TypeRef{}, MemberInfo{}, false
	}
	return pm.entry.selfRef(), pm.entry.public.list[pm.idx], true
}

// Types はクラス・フォーム・Public UDT / Enum の型情報を名前順に返す（標準モジュールは除く）。
func (p *wsTypeProvider) Types() []TypeInfo {
	s := p.w.snapshot()
	out := make([]TypeInfo, 0, len(s.types))
	for _, t := range s.types {
		out = append(out, t.info())
	}
	sort.Slice(out, func(i, j int) bool { return strings.ToLower(out[i].Ref.Name) < strings.ToLower(out[j].Ref.Name) })
	return out
}

// --- wsModuleProvider ---

// wsModuleProvider は標準モジュールを moduleLib 修飾でだけ提供する（Public メンバーのみ）。
type wsModuleProvider struct{ w *Workspace }

func (p *wsModuleProvider) QualifiedOnly() bool { return true }

func (p *wsModuleProvider) OwnsLibrary(lib string) bool { return strings.EqualFold(lib, moduleLib) }

func (p *wsModuleProvider) std(name string) *ModuleEntry {
	e, ok := p.w.snapshot().byName[strings.ToLower(name)]
	if !ok || e.Kind != ModuleStd {
		return nil
	}
	return e
}

func (p *wsModuleProvider) LookupType(name string) (TypeInfo, bool) {
	if e := p.std(name); e != nil {
		return e.typeInfo(), true
	}
	return TypeInfo{}, false
}

func (p *wsModuleProvider) Members(ref TypeRef) []MemberInfo {
	if e := p.std(ref.Name); e != nil {
		return e.public.list
	}
	return nil
}

func (p *wsModuleProvider) Member(ref TypeRef, name string) (MemberInfo, bool) {
	if e := p.std(ref.Name); e != nil {
		if i, ok := e.public.lookup(name); ok {
			return e.public.list[i], true
		}
	}
	return MemberInfo{}, false
}

func (p *wsModuleProvider) MemberType(ref TypeRef, name string) (TypeRef, bool) {
	if e := p.std(ref.Name); e != nil {
		if i, ok := e.public.lookup(name); ok {
			mt := e.public.types[i]
			return mt, !mt.IsZero()
		}
	}
	return TypeRef{}, false
}
