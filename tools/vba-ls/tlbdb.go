package main

import (
	"database/sql"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"sync"

	_ "modernc.org/sqlite"
)

// TLBDatabase は catia_api.db への読み取り専用アクセスを提供する。
type TLBDatabase struct {
	db *sql.DB
}

// LoadTLBDatabase は SQLite ファイルを読み取り専用（mode=ro）で開く。
// ファイルがなければエラーを返し、新しく作ることはない。DB への書き込みは一切行わない。
func LoadTLBDatabase(path string) (*TLBDatabase, error) {
	dsn, err := readOnlyDSN(path)
	if err != nil {
		return nil, err
	}
	db, err := sql.Open("sqlite", dsn)
	if err != nil {
		return nil, fmt.Errorf("SQLite open: %w", err)
	}
	db.SetMaxOpenConns(1)
	// sql.Open は接続を遅延するため、ここで実際に開いて存在を確かめる
	if err := db.Ping(); err != nil {
		_ = db.Close()
		return nil, fmt.Errorf("SQLite open(ro): %w", err)
	}
	return &TLBDatabase{db: db}, nil
}

// readOnlyDSN は "file:///C:/.../catia_api.db?mode=ro" 形式の URI を作る。
// modernc sqlite は "file:" で始まる DSN を URI としてそのまま sqlite3_open_v2 に渡すため mode=ro が効く。
// 日本語・空白・#・% などは url.URL が % エスケープする。
func readOnlyDSN(path string) (string, error) {
	abs, err := filepath.Abs(path)
	if err != nil {
		return "", err
	}
	p := filepath.ToSlash(abs)
	if !strings.HasPrefix(p, "/") {
		p = "/" + p // Windows: "/C:/..."（SQLite の win VFS が先頭の "/" を除く）
	}
	u := url.URL{Scheme: "file", Path: p, RawQuery: "mode=ro"}
	return u.String(), nil
}

// Close はデータベース接続を閉じる。
func (t *TLBDatabase) Close() {
	if t != nil && t.db != nil {
		_ = t.db.Close()
	}
}

// tlbLibraries は TLBProvider が担当する修飾子（大文字小文字無視）。推定のリストで、
// 2-1 で DB に libraries テーブルを持たせたらそこから作るように置き換える。
var tlbLibraries = map[string]bool{
	"catia": true, "infitf": true, "mecmod": true, "partitf": true, "hybridshapetypelib": true,
	"productstructuretypelib": true, "draftingitf": true, "knowledgewaretypelib": true,
	"spatypelib": true, "navigatortypelib": true,
}

// tlbOpenTypes は多態的な基底型。補完は出すが「存在しない」診断は出さない（Complete=false）。
var tlbOpenTypes = map[string]bool{
	"anyobject": true, "catbasedispatch": true, "catbaseunknown": true, "object": true,
	"idispatch": true, "iunknown": true, "variant": true,
}

// tlbHiddenMembers は IUnknown / IDispatch のメソッド。VBA からは使わないため候補から除く（根本対処は 2-1）。
var tlbHiddenMembers = map[string]bool{
	"queryinterface": true, "addref": true, "release": true, "gettypeinfocount": true,
	"gettypeinfo": true, "getidsofnames": true, "invoke": true,
}

type tlbIface struct {
	id     int64
	name   string
	parent sql.NullInt64
}

type memberSet struct {
	list   []MemberInfo
	byName map[string]int
}

// TLBProvider は catia_api.db の型を TypeProvider として提供する。
// interfaces は初回に全件ロードし、継承込みのメンバーは interface ごとにメモ化する（DB は読み取り専用なので無効化不要）。
type TLBProvider struct {
	db     *TLBDatabase
	once   sync.Once
	byName map[string]tlbIface // 小文字名 → iface
	byID   map[int64]tlbIface

	mu     sync.Mutex
	merged map[int64]*memberSet

	namesOnce sync.Once
	names     map[string]bool // 全インターフェースのメンバー名（小文字）。読み込みに失敗したら nil

	typesOnce sync.Once
	typeList  []TypeInfo // 全インターフェース（IUnknown / IDispatch を除く）。名前順
}

// newTLBProvider は db が nil なら interface の nil を返す（typed nil を作らない）。
func newTLBProvider(db *TLBDatabase) TypeProvider {
	if db == nil || db.db == nil {
		return nil
	}
	return &TLBProvider{db: db, merged: map[int64]*memberSet{}}
}

func (p *TLBProvider) load() {
	p.once.Do(func() {
		p.byName = map[string]tlbIface{}
		p.byID = map[int64]tlbIface{}
		rows, err := p.db.db.Query(`SELECT id, name, parent_id FROM interfaces ORDER BY id`)
		if err != nil {
			fmt.Fprintf(os.Stderr, "vba-ls: interfaces の読み込みに失敗: %v\n", err)
			return
		}
		defer rows.Close()
		for rows.Next() {
			var it tlbIface
			if err := rows.Scan(&it.id, &it.name, &it.parent); err != nil {
				continue
			}
			key := strings.ToLower(it.name)
			if _, dup := p.byName[key]; !dup {
				p.byName[key] = it
			}
			p.byID[it.id] = it
		}
	})
}

func (p *TLBProvider) OwnsLibrary(lib string) bool { return tlbLibraries[strings.ToLower(lib)] }

func (p *TLBProvider) LookupType(name string) (TypeInfo, bool) {
	p.load()
	key := strings.ToLower(name)
	it, ok := p.byName[key]
	if !ok {
		return TypeInfo{}, false
	}
	return TypeInfo{
		Ref:      TypeRef{Lib: "CATIA", Name: it.name},
		Kind:     TypeKindInterface,
		Complete: !tlbOpenTypes[key],
	}, true
}

func (p *TLBProvider) Members(t TypeRef) []MemberInfo {
	if ms := p.memberSetOf(t); ms != nil {
		return ms.list
	}
	return nil
}

func (p *TLBProvider) Member(t TypeRef, name string) (MemberInfo, bool) {
	ms := p.memberSetOf(t)
	if ms == nil {
		return MemberInfo{}, false
	}
	i, ok := ms.byName[strings.ToLower(name)]
	if !ok {
		return MemberInfo{}, false
	}
	return ms.list[i], true
}

func (p *TLBProvider) MemberType(t TypeRef, name string) (TypeRef, bool) {
	m, ok := p.Member(t, name)
	if !ok {
		return TypeRef{}, false
	}
	ret := strings.TrimSpace(m.ReturnType)
	if ret == "" || strings.EqualFold(ret, "void") {
		return TypeRef{}, false
	}
	if it, ok := p.byName[strings.ToLower(ret)]; ok {
		return TypeRef{Lib: "CATIA", Name: it.name}, true
	}
	return ParseTypeName(ret), true
}

func (p *TLBProvider) memberSetOf(t TypeRef) *memberSet {
	p.load()
	it, ok := p.byName[strings.ToLower(t.Name)]
	if !ok {
		return nil
	}
	p.mu.Lock()
	defer p.mu.Unlock()
	if ms, ok := p.merged[it.id]; ok {
		return ms
	}
	ms := p.buildMerged(it)
	p.merged[it.id] = ms
	return ms
}

// buildMerged は parent を最大 20 段・循環ガード付きでたどり、派生側優先で重複を除いて名前順に並べる。
func (p *TLBProvider) buildMerged(it tlbIface) *memberSet {
	ms := &memberSet{byName: map[string]int{}}
	seen := map[int64]bool{}
	cur, ok := it, true
	for depth := 0; ok && depth <= 20 && !seen[cur.id]; depth++ {
		seen[cur.id] = true
		for _, m := range p.ownMembers(cur) {
			key := strings.ToLower(m.Name)
			if _, dup := ms.byName[key]; dup || tlbHiddenMembers[key] {
				continue
			}
			ms.byName[key] = len(ms.list)
			ms.list = append(ms.list, m)
		}
		if !cur.parent.Valid {
			break
		}
		cur, ok = p.byID[cur.parent.Int64]
	}
	sort.SliceStable(ms.list, func(i, j int) bool {
		return strings.ToLower(ms.list[i].Name) < strings.ToLower(ms.list[j].Name)
	})
	for i, m := range ms.list {
		ms.byName[strings.ToLower(m.Name)] = i
	}
	return ms
}

// ownMembers は interface 自身のメソッド（引数付き）とプロパティを返す（メソッドが先）。
func (p *TLBProvider) ownMembers(it tlbIface) []MemberInfo {
	owner := TypeRef{Lib: "CATIA", Name: it.name}
	var out []MemberInfo
	byMethodID := map[int64]int{}

	rows, err := p.db.db.Query(`SELECT id, name, return_type FROM methods WHERE interface_id = ? ORDER BY id`, it.id)
	if err == nil {
		for rows.Next() {
			var id int64
			var name, ret string
			if err := rows.Scan(&id, &name, &ret); err != nil {
				continue
			}
			byMethodID[id] = len(out)
			out = append(out, MemberInfo{Name: name, ReturnType: ret, Kind: MemberKindMethod, Visibility: VisPublic, Owner: owner})
		}
		rows.Close()
	}

	rows, err = p.db.db.Query(`SELECT pa.method_id, pa.name, pa.type FROM parameters pa JOIN methods m ON m.id = pa.method_id
WHERE m.interface_id = ? ORDER BY pa.method_id, pa.position`, it.id)
	if err == nil {
		for rows.Next() {
			var mid int64
			var name, typ string
			if err := rows.Scan(&mid, &name, &typ); err != nil {
				continue
			}
			if i, ok := byMethodID[mid]; ok {
				out[i].Params = append(out[i].Params, TLBParam{Name: name, Type: typ})
			}
		}
		rows.Close()
	}

	rows, err = p.db.db.Query(`SELECT name, type FROM properties WHERE interface_id = ? ORDER BY id`, it.id)
	if err == nil {
		for rows.Next() {
			var name, typ string
			if err := rows.Scan(&name, &typ); err != nil {
				continue
			}
			out = append(out, MemberInfo{Name: name, ReturnType: typ, Kind: MemberKindProperty, Visibility: VisPublic, Owner: owner})
		}
		rows.Close()
	}
	return out
}

// HasMemberNamed は、どれかのインターフェースに name（大文字小文字無視）というメンバーがあるかを返す。
// CATIA は遅延バインディングなので、宣言型が基底型でも派生型のメンバーが実行時に動く。診断ではこの集合に無い名前
// （typo か型情報の欠落）だけを警告にする。初回にメソッドとプロパティの名前を一度だけ読み込む。
// 読み込みに失敗したときは常に true を返す（安全側＝診断しない）。
func (p *TLBProvider) HasMemberNamed(name string) bool {
	p.namesOnce.Do(func() {
		rows, err := p.db.db.Query(`SELECT name FROM methods UNION SELECT name FROM properties`)
		if err != nil {
			fmt.Fprintf(os.Stderr, "vba-ls: メンバー名の読み込みに失敗: %v\n", err)
			return
		}
		defer rows.Close()
		names := map[string]bool{}
		for rows.Next() {
			var n string
			if err := rows.Scan(&n); err != nil {
				continue
			}
			if key := strings.ToLower(n); !tlbHiddenMembers[key] {
				names[key] = true
			}
		}
		if rows.Err() != nil {
			fmt.Fprintf(os.Stderr, "vba-ls: メンバー名の読み込みに失敗: %v\n", rows.Err())
			return
		}
		p.names = names
	})
	if p.names == nil {
		return true
	}
	return p.names[strings.ToLower(name)]
}

// Types は全インターフェース（IUnknown / IDispatch を除く）を名前順に返す（型補完用）。
func (p *TLBProvider) Types() []TypeInfo {
	p.typesOnce.Do(func() {
		p.load()
		for key := range p.byName {
			if key == "iunknown" || key == "idispatch" {
				continue
			}
			if ti, ok := p.LookupType(key); ok {
				p.typeList = append(p.typeList, ti)
			}
		}
		sort.Slice(p.typeList, func(i, j int) bool { return p.typeList[i].Ref.Name < p.typeList[j].Ref.Name })
	})
	return p.typeList
}
