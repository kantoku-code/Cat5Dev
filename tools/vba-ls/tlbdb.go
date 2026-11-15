package main

import (
	"database/sql"
	"fmt"
	"net/url"
	"os"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"

	_ "modernc.org/sqlite"
)

// TLBDatabase は catia_api.db への読み取り専用アクセスを提供する。
type TLBDatabase struct {
	db         *sql.DB
	schemaOnce sync.Once
	schema     *tlbSchema
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

// tlbSchema は DB のテーブル・列の有無とスキーマの版（読み取りだけで検出する）。
// v1（2-0 までの抽出器）は interfaces / methods / properties / parameters の 4 テーブルで、継承・Enum・説明を持たない。
// v2（tools/tlb-extractor/schema.sql）は列とテーブルを足したもので、meta.schema_version = '2'。
type tlbSchema struct {
	Version int                        // meta.schema_version（無ければ 1）
	tables  map[string]bool            // 小文字のテーブル名
	cols    map[string]map[string]bool // テーブル → 小文字の列名
}

// HasTable はテーブルがあるかを返す。
func (s *tlbSchema) HasTable(t string) bool { return s.tables[strings.ToLower(t)] }

// Has はテーブル t に列 col があるかを返す。
func (s *tlbSchema) Has(t, col string) bool { return s.cols[strings.ToLower(t)][strings.ToLower(col)] }

// Hierarchy は継承（parent_id）と自前メンバーだけの構成か（v2 以上）を返す。
// v1 は parent_id が空で各型に平坦なメンバーを持つため、継承をたどる判定に使えない。
func (s *tlbSchema) Hierarchy() bool { return s.Version >= 2 }

// col は SELECT 句の列の式を返す。列が無ければ fallback（定数の式）にする。col は "別名.列" でもよい。
func (s *tlbSchema) col(table, col, fallback string) string {
	name := col
	if i := strings.IndexByte(col, '.'); i >= 0 {
		name = col[i+1:]
	}
	if s.Has(table, name) {
		return col
	}
	return fallback
}

// Schema はスキーマを返す（初回にだけ sqlite_master と pragma_table_info を読む。DB は変更しない）。
func (t *TLBDatabase) Schema() *tlbSchema {
	t.schemaOnce.Do(func() { t.schema = detectSchema(t.db) })
	return t.schema
}

func detectSchema(db *sql.DB) *tlbSchema {
	s := &tlbSchema{Version: 1, tables: map[string]bool{}, cols: map[string]map[string]bool{}}
	rows, err := db.Query(`SELECT name FROM sqlite_master WHERE type = 'table'`)
	if err != nil {
		fmt.Fprintf(os.Stderr, "vba-ls: TLB データベースのスキーマを読めない: %v\n", err)
		return s
	}
	var names []string
	for rows.Next() {
		var n string
		if rows.Scan(&n) == nil {
			names = append(names, n)
		}
	}
	rows.Close()
	for _, n := range names {
		key := strings.ToLower(n)
		s.tables[key] = true
		s.cols[key] = map[string]bool{}
		r, err := db.Query(`SELECT name FROM pragma_table_info(?)`, n)
		if err != nil {
			continue
		}
		for r.Next() {
			var c string
			if r.Scan(&c) == nil {
				s.cols[key][strings.ToLower(c)] = true
			}
		}
		r.Close()
	}
	if s.tables["meta"] {
		var v string
		if db.QueryRow(`SELECT value FROM meta WHERE key = 'schema_version'`).Scan(&v) == nil {
			if n, err := strconv.Atoi(strings.TrimSpace(v)); err == nil && n > 0 {
				s.Version = n
			}
		}
	}
	return s
}

// Close はデータベース接続を閉じる。
func (t *TLBDatabase) Close() {
	if t != nil && t.db != nil {
		_ = t.db.Close()
	}
}

// tlbLibraries は v1 の DB（libraries テーブルが無い）で TLBProvider が担当する修飾子（大文字小文字無視）。
// v2 は libraries テーブルの名前（INFITF / MECMOD …）と "CATIA" を使う。
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

// tlbHiddenMembers は v1 の DB で候補から除く名前。v1 は各型に IUnknown / IDispatch と CATBaseUnknown の
// restricted な関数が平坦に入っているため名前で除く。v2 は抽出時に restricted を除いてあるので使わない
// （TreeVizManipSettingAtt.GetTypeInfo のような正当な同名メンバーを隠さない）。
var tlbHiddenMembers = map[string]bool{
	"queryinterface": true, "addref": true, "release": true, "gettypeinfocount": true,
	"gettypeinfo": true, "getidsofnames": true, "invoke": true,
	"destructor": true, "getmetaobject": true, "isa": true, "isakindof": true, "getimpl": true, "setimpl": true,
	"isnull": true, "isequal": true, "surchargeqi": true, "changecomponentstate": true, "_newenum": true,
}

// TLB の flags・access のビット。
const (
	tlbTypeFlagHidden = 0x10 // TYPEFLAG_FHIDDEN
	tlbFuncFlagHidden = 0x40 // FUNCFLAG_FHIDDEN
	tlbParamFlagOpt   = 0x10 // PARAMFLAG_FOPT
	tlbAccessGet      = 1
	tlbAccessPut      = 2
	tlbAccessPutRef   = 4
)

type tlbIface struct {
	id     int64
	name   string
	parent sql.NullInt64
	lib    sql.NullInt64 // v2: libraries.id
	flags  int64         // v2: TYPEFLAGS（0x10 = hidden）
	doc    string        // v2
}

// memberSet は継承込みのメンバー。list は補完用（hidden を除く）、all と byName は hidden を含む（診断・ホバー用）。
type memberSet struct {
	list   []MemberInfo
	all    []MemberInfo
	byName map[string]int // 小文字名 → all の添字
}

// tlbEnum は TLB の Enum 1 つ。
type tlbEnum struct {
	info    TypeInfo     // Kind=TypeKindEnum, Complete=true
	members []MemberInfo // 定義順（Kind=MemberKindEnumMember, ReturnType "Long", Value）
	sorted  []MemberInfo // 名前順（補完用）
	byName  map[string]int
}

// ownMember は型自身のメンバー 1 件（hidden の印付き）。
type ownMember struct {
	info   MemberInfo
	hidden bool
}

// TLBProvider は catia_api.db の型を TypeProvider として提供する。
// interfaces は初回に全件ロードし、継承込みのメンバーは interface ごとにメモ化する（DB は読み取り専用なので無効化不要）。
type TLBProvider struct {
	db      *TLBDatabase
	once    sync.Once
	schema  *tlbSchema
	byName  map[string]tlbIface // 小文字名 → iface
	byID    map[int64]tlbIface
	libs    map[string]bool       // OwnsLibrary の対象（小文字）
	libName map[int64]string      // libraries.id → 名前（表示用）
	enums   map[string]*tlbEnum   // v2: 小文字の Enum 名 → Enum
	consts  map[string]MemberInfo // v2: 小文字のメンバー名 → Enum メンバー（同名は最初の Enum が勝つ）

	mu     sync.Mutex
	merged map[int64]*memberSet

	namesOnce sync.Once
	names     map[string]bool // 全インターフェースのメンバー名（小文字）。読み込みに失敗したら nil

	typesOnce sync.Once
	typeList  []TypeInfo // 型補完用の型（IUnknown / IDispatch・hidden を除く）。名前順
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
		p.schema = p.db.Schema()
		p.byName = map[string]tlbIface{}
		p.byID = map[int64]tlbIface{}
		p.libs = map[string]bool{}
		p.libName = map[int64]string{}
		p.enums = map[string]*tlbEnum{}
		p.consts = map[string]MemberInfo{}
		s := p.schema
		p.loadLibraries()
		p.loadEnums()
		q := `SELECT id, name, parent_id, ` + s.col("interfaces", "library_id", "NULL") + `, ` +
			s.col("interfaces", "flags", "0") + `, ` + s.col("interfaces", "doc", "''") + ` FROM interfaces ORDER BY id`
		rows, err := p.db.db.Query(q)
		if err != nil {
			fmt.Fprintf(os.Stderr, "vba-ls: interfaces の読み込みに失敗: %v\n", err)
			return
		}
		defer rows.Close()
		for rows.Next() {
			var it tlbIface
			if err := rows.Scan(&it.id, &it.name, &it.parent, &it.lib, &it.flags, &it.doc); err != nil {
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

// loadLibraries は libraries テーブルから修飾子の集合を作る（無ければ v1 の固定リスト）。
func (p *TLBProvider) loadLibraries() {
	p.libs["catia"] = true
	if !p.schema.HasTable("libraries") {
		for k := range tlbLibraries {
			p.libs[k] = true
		}
		return
	}
	rows, err := p.db.db.Query(`SELECT id, name FROM libraries`)
	if err != nil {
		fmt.Fprintf(os.Stderr, "vba-ls: libraries の読み込みに失敗: %v\n", err)
		return
	}
	defer rows.Close()
	for rows.Next() {
		var id int64
		var n string
		if rows.Scan(&id, &n) == nil && n != "" {
			p.libs[strings.ToLower(n)] = true
			p.libName[id] = n
		}
	}
}

// loadEnums は enums / enum_members を読む（v2 のみ。312 種・1,405 メンバー程度なので一度に読む）。
func (p *TLBProvider) loadEnums() {
	if !p.schema.HasTable("enums") || !p.schema.HasTable("enum_members") {
		return
	}
	s := p.schema
	rows, err := p.db.db.Query(`SELECT e.name, ` + s.col("enums", "e.library_id", "NULL") + `, ` + s.col("enums", "e.doc", "''") +
		`, m.name, m.value, ` + s.col("enum_members", "m.doc", "''") +
		` FROM enums e JOIN enum_members m ON m.enum_id = e.id ORDER BY e.id, m.id`)
	if err != nil {
		fmt.Fprintf(os.Stderr, "vba-ls: enums の読み込みに失敗: %v\n", err)
		return
	}
	defer rows.Close()
	for rows.Next() {
		var en, edoc, mn, mdoc string
		var lib sql.NullInt64
		var val int64
		if err := rows.Scan(&en, &lib, &edoc, &mn, &val, &mdoc); err != nil {
			continue
		}
		key := strings.ToLower(en)
		e := p.enums[key]
		if e == nil {
			e = &tlbEnum{byName: map[string]int{}, info: TypeInfo{Ref: TypeRef{Lib: "CATIA", Name: en}, Kind: TypeKindEnum,
				Complete: true, Doc: edoc, Container: p.libName[lib.Int64]}}
			p.enums[key] = e
		}
		m := MemberInfo{Name: mn, ReturnType: "Long", Kind: MemberKindEnumMember, ReadOnly: true, Doc: mdoc,
			Visibility: VisPublic, Owner: e.info.Ref, Value: strconv.FormatInt(val, 10)}
		mk := strings.ToLower(mn)
		if _, dup := e.byName[mk]; !dup {
			e.byName[mk] = len(e.members)
			e.members = append(e.members, m)
		}
		if _, dup := p.consts[mk]; !dup {
			p.consts[mk] = m
		}
	}
	for _, e := range p.enums {
		e.sorted = append([]MemberInfo(nil), e.members...)
		sort.SliceStable(e.sorted, func(i, j int) bool { return strings.ToLower(e.sorted[i].Name) < strings.ToLower(e.sorted[j].Name) })
	}
}

// HasHierarchy は DB が継承（parent_id と自前メンバー）を持つ v2 以上かを返す。
func (p *TLBProvider) HasHierarchy() bool {
	p.load()
	return p.schema.Hierarchy()
}

func (p *TLBProvider) OwnsLibrary(lib string) bool {
	p.load()
	return p.libs[strings.ToLower(lib)]
}

func (p *TLBProvider) LookupType(name string) (TypeInfo, bool) {
	p.load()
	key := strings.ToLower(name)
	it, ok := p.byName[key]
	if !ok {
		if e := p.enums[key]; e != nil {
			return e.info, true
		}
		return TypeInfo{}, false
	}
	return TypeInfo{
		Ref:       TypeRef{Lib: "CATIA", Name: it.name},
		Kind:      TypeKindInterface,
		Complete:  !tlbOpenTypes[key],
		Doc:       it.doc,
		Container: p.libName[it.lib.Int64],
	}, true
}

func (p *TLBProvider) Members(t TypeRef) []MemberInfo {
	if ms := p.memberSetOf(t); ms != nil {
		return ms.list
	}
	if e := p.enumOf(t); e != nil {
		return e.sorted
	}
	return nil
}

// enumOf は t がインターフェースでなく Enum ならその Enum を返す。
func (p *TLBProvider) enumOf(t TypeRef) *tlbEnum {
	p.load()
	key := strings.ToLower(t.Name)
	if _, isIface := p.byName[key]; isIface {
		return nil
	}
	return p.enums[key]
}

// Member は hidden のメンバーも引ける（補完には出さないが、書かれていれば診断しない・ホバーできる）。
func (p *TLBProvider) Member(t TypeRef, name string) (MemberInfo, bool) {
	ms := p.memberSetOf(t)
	if ms == nil {
		if e := p.enumOf(t); e != nil {
			if i, ok := e.byName[strings.ToLower(name)]; ok {
				return e.members[i], true
			}
		}
		return MemberInfo{}, false
	}
	i, ok := ms.byName[strings.ToLower(name)]
	if !ok {
		return MemberInfo{}, false
	}
	return ms.all[i], true
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
	if e := p.enums[strings.ToLower(ret)]; e != nil {
		return e.info.Ref, true
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
	v1 := !p.schema.Hierarchy()
	var all []ownMember
	byName := map[string]bool{}
	seen := map[int64]bool{}
	cur, ok := it, true
	for depth := 0; ok && depth <= 20 && !seen[cur.id]; depth++ {
		seen[cur.id] = true
		for _, om := range p.ownMembers(cur) {
			key := strings.ToLower(om.info.Name)
			if byName[key] || (v1 && tlbHiddenMembers[key]) {
				continue
			}
			byName[key] = true
			all = append(all, om)
		}
		if !cur.parent.Valid {
			break
		}
		cur, ok = p.byID[cur.parent.Int64]
	}
	sort.SliceStable(all, func(i, j int) bool {
		return strings.ToLower(all[i].info.Name) < strings.ToLower(all[j].info.Name)
	})
	ms := &memberSet{all: make([]MemberInfo, len(all)), byName: make(map[string]int, len(all))}
	for i, om := range all {
		ms.all[i] = om.info
		ms.byName[strings.ToLower(om.info.Name)] = i
		if !om.hidden {
			ms.list = append(ms.list, om.info)
		}
	}
	return ms
}

// ownMembers は interface 自身のメソッド（引数付き）とプロパティを返す（メソッドが先）。
// v2 の列（flags / vararg / access / default_value / doc）が無い DB では既定値で埋める。
func (p *TLBProvider) ownMembers(it tlbIface) []ownMember {
	s := p.schema
	owner := TypeRef{Lib: "CATIA", Name: it.name}
	var out []ownMember
	byMethodID := map[int64]int{}
	var varargIDs []int64

	rows, err := p.db.db.Query(`SELECT id, name, return_type, `+s.col("methods", "flags", "0")+`, `+
		s.col("methods", "vararg", "0")+`, `+s.col("methods", "doc", "''")+
		` FROM methods WHERE interface_id = ? ORDER BY id`, it.id)
	if err == nil {
		for rows.Next() {
			var id, flags, vararg int64
			var name, ret, doc string
			if err := rows.Scan(&id, &name, &ret, &flags, &vararg, &doc); err != nil {
				continue
			}
			byMethodID[id] = len(out)
			if vararg != 0 {
				varargIDs = append(varargIDs, id)
			}
			out = append(out, ownMember{
				info:   MemberInfo{Name: name, ReturnType: ret, Kind: MemberKindMethod, Visibility: VisPublic, Owner: owner, Doc: doc},
				hidden: flags&tlbFuncFlagHidden != 0,
			})
		}
		rows.Close()
	}

	rows, err = p.db.db.Query(`SELECT pa.method_id, pa.name, pa.type, `+s.col("parameters", "pa.flags", "0")+`, `+
		s.col("parameters", "pa.default_value", "NULL")+`, `+s.col("parameters", "pa.doc", "''")+
		` FROM parameters pa JOIN methods m ON m.id = pa.method_id WHERE m.interface_id = ? ORDER BY pa.method_id, pa.position`, it.id)
	if err == nil {
		for rows.Next() {
			var mid, flags int64
			var name, typ, doc string
			var def sql.NullString
			if err := rows.Scan(&mid, &name, &typ, &flags, &def, &doc); err != nil {
				continue
			}
			if i, ok := byMethodID[mid]; ok {
				out[i].info.Params = append(out[i].info.Params, TLBParam{
					Name: name, Type: typ, Doc: doc,
					Optional: flags&tlbParamFlagOpt != 0 || def.Valid,
					Default:  def.String,
				})
			}
		}
		rows.Close()
	}
	for _, id := range varargIDs {
		if ps := out[byMethodID[id]].info.Params; len(ps) > 0 {
			ps[len(ps)-1].ParamArray = true
			ps[len(ps)-1].Array = true
		}
	}

	// v1 には access が無い（読み書きの区別が分からない）ので get|put とみなす
	rows, err = p.db.db.Query(`SELECT name, type, `+s.col("properties", "access", "3")+`, `+
		s.col("properties", "flags", "0")+`, `+s.col("properties", "doc", "''")+
		` FROM properties WHERE interface_id = ? ORDER BY id`, it.id)
	if err == nil {
		for rows.Next() {
			var name, typ, doc string
			var access, flags int64
			if err := rows.Scan(&name, &typ, &access, &flags, &doc); err != nil {
				continue
			}
			if access&tlbAccessGet == 0 {
				doc = strings.TrimSpace("（書き込み専用） " + doc)
			}
			out = append(out, ownMember{
				info: MemberInfo{Name: name, ReturnType: typ, Kind: MemberKindProperty, Visibility: VisPublic, Owner: owner,
					Doc: doc, ReadOnly: access&(tlbAccessPut|tlbAccessPutRef) == 0},
				hidden: flags&tlbFuncFlagHidden != 0,
			})
		}
		rows.Close()
	}
	return out
}

// HasMemberNamed は、どれかのインターフェースに name（大文字小文字無視）というメンバーがあるかを返す。
// CATIA は遅延バインディングなので、宣言型が基底型でも派生型のメンバーが実行時に動く。v1 の DB では、
// 診断はこの集合に無い名前（typo か型情報の欠落）だけを警告にする。初回にメソッドとプロパティの名前を一度だけ読み込む。
// 読み込みに失敗したときは常に true を返す（安全側＝診断しない）。
func (p *TLBProvider) HasMemberNamed(name string) bool {
	p.load()
	p.namesOnce.Do(func() {
		rows, err := p.db.db.Query(`SELECT name FROM methods UNION SELECT name FROM properties`)
		if err != nil {
			fmt.Fprintf(os.Stderr, "vba-ls: メンバー名の読み込みに失敗: %v\n", err)
			return
		}
		defer rows.Close()
		v1 := !p.schema.Hierarchy()
		names := map[string]bool{}
		for rows.Next() {
			var n string
			if err := rows.Scan(&n); err != nil {
				continue
			}
			if key := strings.ToLower(n); !v1 || !tlbHiddenMembers[key] {
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

// Constant は TLB の Enum メンバーを非修飾の名前で引く（VBA と同じく catPaperA4 だけで参照できる。ホバー用）。
func (p *TLBProvider) Constant(name string) (MemberInfo, bool) {
	p.load()
	m, ok := p.consts[strings.ToLower(name)]
	return m, ok
}

// Types は型補完用の型（IUnknown / IDispatch と hidden のインターフェースを除く）と Enum を名前順に返す。
func (p *TLBProvider) Types() []TypeInfo {
	p.typesOnce.Do(func() {
		p.load()
		for key, it := range p.byName {
			if key == "iunknown" || key == "idispatch" || it.flags&tlbTypeFlagHidden != 0 {
				continue
			}
			if ti, ok := p.LookupType(key); ok {
				p.typeList = append(p.typeList, ti)
			}
		}
		for key, e := range p.enums {
			if _, dup := p.byName[key]; !dup {
				p.typeList = append(p.typeList, e.info)
			}
		}
		sort.Slice(p.typeList, func(i, j int) bool { return p.typeList[i].Ref.Name < p.typeList[j].Ref.Name })
	})
	return p.typeList
}
