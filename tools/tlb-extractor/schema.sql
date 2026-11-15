-- catia_api.db のスキーマ v2（extract.py が生成し、tools/vba-ls が読み取り専用で読む）。
-- v1（interfaces / methods / properties / parameters の 4 テーブル）に列とテーブルを足したもので、
-- 既存の列名・意味は変えていない（v1 を読む LSP でもそのまま読める）。
-- Go のテスト（tools/vba-ls）もこのファイルを読んでテスト用 DB を作る。

-- schema_version='2', extractor_version, catia_release（例 'B32'）, tlb_count,
-- tlb_digest（TLB の SHA-256 をファイル名順に連結した SHA-256）, doc_source（'none' | 'chm-summary' | 'chm-full'）。
-- 再現性のため時刻は入れない。
CREATE TABLE meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- TLB 1 ファイルにつき 1 行。name は VBA の修飾子として使える名前（INFITF / MECMOD …）。
CREATE TABLE libraries (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    name    TEXT NOT NULL,
    doc     TEXT NOT NULL DEFAULT '',
    guid    TEXT NOT NULL DEFAULT '',
    version TEXT NOT NULL DEFAULT '',
    file    TEXT NOT NULL DEFAULT ''
);

-- dispinterface（dual）と coclass。メンバーは自前の分だけ持ち、継承は parent_id（単一継承）でたどる。
CREATE TABLE interfaces (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT    NOT NULL UNIQUE,
    parent_id  INTEGER REFERENCES interfaces(id),
    library_id INTEGER REFERENCES libraries(id),
    kind       TEXT    NOT NULL DEFAULT 'dispatch', -- 'dispatch' | 'coclass'
    flags      INTEGER NOT NULL DEFAULT 0,          -- TYPEFLAGS（0x10 = hidden）
    doc        TEXT    NOT NULL DEFAULT ''
);

-- restricted（IUnknown / IDispatch / CATBaseUnknown の内部関数）は入れない。
CREATE TABLE methods (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    interface_id INTEGER NOT NULL REFERENCES interfaces(id),
    name         TEXT    NOT NULL,
    return_type  TEXT    NOT NULL DEFAULT '',
    flags        INTEGER NOT NULL DEFAULT 0, -- FUNCFLAGS（0x40 = hidden）
    vararg       INTEGER NOT NULL DEFAULT 0, -- 1 = 最後の引数が ParamArray
    doc          TEXT    NOT NULL DEFAULT ''
);

-- 同名の get / put / putref は 1 行にまとめる。書き込み専用の型は put の最後の引数の型。
CREATE TABLE properties (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    interface_id INTEGER NOT NULL REFERENCES interfaces(id),
    name         TEXT    NOT NULL,
    type         TEXT    NOT NULL DEFAULT '',
    access       INTEGER NOT NULL DEFAULT 1, -- 1=get, 2=put(Let), 4=putref(Set) のビット和
    flags        INTEGER NOT NULL DEFAULT 0, -- FUNCFLAGS（0x40 = hidden）
    doc          TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE parameters (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    method_id     INTEGER NOT NULL REFERENCES methods(id),
    name          TEXT    NOT NULL,
    type          TEXT    NOT NULL DEFAULT '',
    position      INTEGER NOT NULL DEFAULT 0,
    flags         INTEGER NOT NULL DEFAULT 0, -- PARAMFLAGS（1=in, 2=out, 0x10=optional, 0x20=hasdefault）
    default_value TEXT,                       -- VBA の式（'"abc"' / True / 数値）。NULL = 既定値なし
    doc           TEXT    NOT NULL DEFAULT ''
);

-- 同名の enum は最初の TLB（ファイル名順）だけを入れる。foo_<Framework> のダミーは入れない。
CREATE TABLE enums (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    name       TEXT    NOT NULL UNIQUE,
    library_id INTEGER REFERENCES libraries(id),
    doc        TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE enum_members (
    id      INTEGER PRIMARY KEY AUTOINCREMENT,
    enum_id INTEGER NOT NULL REFERENCES enums(id),
    name    TEXT    NOT NULL,
    value   INTEGER NOT NULL,
    doc     TEXT    NOT NULL DEFAULT ''
);

CREATE INDEX idx_iface_name       ON interfaces(name COLLATE NOCASE);
CREATE INDEX idx_method_iface     ON methods(interface_id);
CREATE INDEX idx_method_name      ON methods(name COLLATE NOCASE);
CREATE INDEX idx_prop_iface       ON properties(interface_id);
CREATE INDEX idx_prop_name        ON properties(name COLLATE NOCASE);
CREATE INDEX idx_param_method     ON parameters(method_id);
CREATE INDEX idx_enum_member_enum ON enum_members(enum_id);
