"""
CATIA V5 TLB → SQLite 変換ツール
===================================
CATIA の型ライブラリ（.tlb）を読み込み、VBA Language Server（tools/vba-ls）用の
SQLite データベース catia_api.db（スキーマ v2、schema.sql）を生成する。

- 継承: dual interface の vtable 側の基底を interfaces.parent_id に入れ、各型には自前のメンバーだけを持たせる
- IUnknown / IDispatch / CATBaseUnknown の内部関数（restricted）は入れない。hidden は flags に残す
- 書き込み専用プロパティも properties に入れる（access ビット）
- Enum（enums / enum_members）、TLB ごとのライブラリ名（libraries）、出どころ（meta）
- 出力は決定的（同じ TLB からは同じバイト列の DB ができる）

使い方（UV 仮想環境）:
    # CATIA の標準インストール先から自動検索
    uv run python extract.py

    # ディレクトリ内の全 TLB を処理
    uv run python extract.py "C:/Program Files/Dassault Systemes/B32/win_b64/code/bin"

    # 出力先を明示する場合
    uv run python extract.py --out catia_api.db "C:/Program Files/Dassault Systemes/B32/win_b64/code/bin"

    # 検出された TLB 一覧を確認するだけ
    uv run python extract.py --list

    # 生成済み DB の不変条件を検査する
    uv run python extract.py --verify catia_api.db
"""

from __future__ import annotations

import argparse
import glob
import hashlib
import os
import re
import sqlite3
import sys
from typing import Optional

EXTRACTOR_VERSION = "2.0.0"
SCHEMA_VERSION = "2"
SCHEMA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "schema.sql")

# COM の定数（comtypes に無いものも含めて数値で持つ）
TKIND_ENUM, TKIND_ALIAS, TKIND_DISPATCH, TKIND_COCLASS = 0, 6, 4, 5
TYPEFLAG_FDUAL = 0x40
IMPLTYPEFLAG_FDEFAULT = 0x1
FUNCFLAG_FRESTRICTED = 0x1
FUNCFLAG_FHIDDEN = 0x40
PARAMFLAG_FOPT = 0x10
PARAMFLAG_FHASDEFAULT = 0x20
INVOKE_FUNC, INVOKE_PROPERTYGET, INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF = 1, 2, 4, 8
ACCESS_GET, ACCESS_PUT, ACCESS_PUTREF = 1, 2, 4
REGKIND_NONE = 2

# VARTYPE
VT_EMPTY, VT_I2, VT_I4, VT_R4, VT_R8, VT_CY, VT_DATE, VT_BSTR = 0, 2, 3, 4, 5, 6, 7, 8
VT_DISPATCH, VT_BOOL, VT_VARIANT, VT_UNKNOWN, VT_UI1 = 9, 11, 12, 13, 17
VT_UI2, VT_UI4, VT_I8, VT_INT, VT_UINT, VT_VOID, VT_HRESULT = 18, 19, 20, 22, 23, 24, 25
VT_PTR, VT_SAFEARRAY, VT_USERDEFINED = 26, 27, 29
_VT_BYREF = 0x4000
_VT_ARRAY = 0x2000

_VT_MAP = {
    VT_VOID: "void", VT_HRESULT: "void", VT_EMPTY: "void",
    VT_BOOL: "Boolean", VT_UI1: "Byte",
    VT_I2: "Integer", VT_UI2: "Integer",
    VT_I4: "Long", VT_INT: "Long", VT_UINT: "Long", VT_UI4: "Long", VT_I8: "LongLong",
    VT_R4: "Single", VT_R8: "Double", VT_CY: "Currency", VT_DATE: "Date",
    VT_BSTR: "String", VT_DISPATCH: "Object", VT_UNKNOWN: "Object", VT_VARIANT: "Variant",
}


# ---------------------------------------------------------------------------
# 型・値の変換（comtypes の構造体を受け取るが、それ自体には依存しない）
# ---------------------------------------------------------------------------
def _resolve_type(tdesc, tinfo, depth: int = 0) -> str:
    """TYPEDESC を VBA の型名にする。ポインタは剥がし、typedef（ALIAS）は実体まで解決する。"""
    vt = tdesc.vt & ~_VT_BYREF & ~_VT_ARRAY
    if depth > 16:
        return "Variant"
    if vt == VT_PTR:
        try:
            return _resolve_type(tdesc._.lptdesc.contents, tinfo, depth + 1)
        except Exception:
            return "Object"
    if vt == VT_SAFEARRAY:
        return "Variant"  # CATSafeArrayVariant 等。VBA からは Variant（配列）として扱う
    if vt == VT_USERDEFINED:
        try:
            ref = tinfo.GetRefTypeInfo(tdesc._.hreftype)
            ra = ref.GetTypeAttr()
            if ra.typekind == TKIND_ALIAS:
                return _resolve_type(ra.tdescAlias, ref, depth + 1)
            name = ref.GetDocumentation(-1)[0]
            return str(name) if name else "Variant"
        except Exception:
            return "Variant"
    return _VT_MAP.get(vt, "Variant")


def vba_literal(v) -> Optional[str]:
    """既定値（VARIANT の Python 値）を VBA の式の文字列にする。表せないものは None。"""
    if v is None:
        return None
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, str):
        return '"' + v.replace('"', '""') + '"'
    if isinstance(v, (int, float)):
        return repr(v) if isinstance(v, float) else str(v)
    return None


def _default_value(elemdesc) -> Optional[str]:
    """引数の既定値（PARAMFLAG_FHASDEFAULT のとき）を VBA の式で返す。"""
    pd = elemdesc._.paramdesc
    if not (pd.wParamFlags & PARAMFLAG_FHASDEFAULT) or not pd.pparamdescex:
        return None
    try:
        return vba_literal(pd.pparamdescex.contents.varDefaultValue.value)
    except Exception:
        return None


def _doc_of(tinfo, memid: int = -1) -> tuple[str, str]:
    """(名前, helpstring) を返す。取得失敗時は ("", "")。"""
    try:
        r = tinfo.GetDocumentation(memid)
        return (str(r[0]) if r[0] else ""), (str(r[1]) if r[1] else "")
    except Exception:
        return "", ""


def _dual_parent(tinfo, tattr) -> tuple[Optional[str], int]:
    """(親の名前, 自前の関数の開始位置) を返す。

    dual の dispinterface の関数は「IUnknown 3 + IDispatch 4 + 基底… + 自前」の平坦な並びで、
    自前は末尾の vtable 側の cFuncs 個。親は vtable 側（GetRefTypeOfImplType(-1)）の impltype 0。
    coclass は既定（FDEFAULT）の impltype を親にする。IDispatch / IUnknown は親にしない。
    """
    parent, first = None, 0
    if tattr.typekind == TKIND_COCLASS:
        for k in range(tattr.cImplTypes):
            if tinfo.GetImplTypeFlags(k) & IMPLTYPEFLAG_FDEFAULT:
                parent = _doc_of(tinfo.GetRefTypeInfo(tinfo.GetRefTypeOfImplType(k)))[0]
                break
    elif tattr.wTypeFlags & TYPEFLAG_FDUAL:
        vt = tinfo.GetRefTypeInfo(tinfo.GetRefTypeOfImplType(-1))
        va = vt.GetTypeAttr()
        if va.cImplTypes:
            parent = _doc_of(vt.GetRefTypeInfo(vt.GetRefTypeOfImplType(0)))[0]
        first = max(0, tattr.cFuncs - va.cFuncs)
    elif tattr.cImplTypes:
        # 純粋な dispinterface（CATIA には無い）: 関数は全部自前として扱う
        parent = _doc_of(tinfo.GetRefTypeInfo(tinfo.GetRefTypeOfImplType(0)))[0]
    if parent in ("IDispatch", "IUnknown", ""):
        parent = None
    return parent, first


# ---------------------------------------------------------------------------
# 説明文の供給元（既定は無し。--chm 指定時は ChmDocs）
# ---------------------------------------------------------------------------
class NoDocs:
    source = "none"

    def type_doc(self, name: str) -> str:
        return ""

    def member_doc(self, type_name: str, member: str) -> str:
        return ""

    def param_doc(self, type_name: str, member: str, param: str) -> str:
        return ""

    def enum_doc(self, name: str) -> str:
        return ""

    def enum_member_doc(self, enum: str, member: str) -> str:
        return ""


# ---------------------------------------------------------------------------
# TLB ロード
# ---------------------------------------------------------------------------
def _load_tlb(path: str):
    """TLB を読み込む。失敗時は None を返す。"""
    import comtypes.typeinfo

    for regkind in (REGKIND_NONE, 0):
        try:
            return comtypes.typeinfo.LoadTypeLibEx(path, regkind)
        except Exception:
            continue
    return None


# ---------------------------------------------------------------------------
# TLB → DB 書き込み
# ---------------------------------------------------------------------------
def _process_tlb(path: str, db: "TLBDatabase", docs) -> Optional[tuple[int, int]]:
    """1 つの TLB を処理して DB に書き込む。(型の数, enum の数) を返す。読み込めなければ None。"""
    tlb = _load_tlb(path)
    if tlb is None:
        return None
    lname, ldoc = _doc_of(tlb)
    la = tlb.GetLibAttr()
    lib_id = db.add_library(lname, ldoc, str(la.guid), f"{la.wMajorVerNum}.{la.wMinorVerNum}", os.path.basename(path))
    n_types = n_enums = 0
    for i in range(tlb.GetTypeInfoCount()):
        t = tlb.GetTypeInfo(i)
        a = t.GetTypeAttr()
        if a.typekind in (TKIND_DISPATCH, TKIND_COCLASS):
            if _process_dispatch(t, a, lib_id, db, docs):
                n_types += 1
        elif a.typekind == TKIND_ENUM:
            if _process_enum(t, a, lib_id, db, docs):
                n_enums += 1
        # ALIAS（typedef）・MODULE（IID 定数）・RECORD は VBA から使わないので入れない
    return n_types, n_enums


def _process_dispatch(t, a, lib_id: int, db: "TLBDatabase", docs) -> bool:
    """dispinterface / coclass 1 つを書き込む。自前の範囲の関数だけを入れる。"""
    name = _doc_of(t)[0]
    if not name:
        return False
    iid = db.add_interface(name, lib_id, "coclass" if a.typekind == TKIND_COCLASS else "dispatch",
                           int(a.wTypeFlags), docs.type_doc(name))
    if iid is None:
        db.warn(f"インターフェース {name} が重複している（2 回目以降は無視）")
        return False
    parent, first = _dual_parent(t, a)
    if parent:
        db.set_parent(iid, parent)

    props: dict[str, dict] = {}
    for j in range(first, a.cFuncs):
        fd = t.GetFuncDesc(j)
        if fd.wFuncFlags & FUNCFLAG_FRESTRICTED:
            continue
        fname = _doc_of(t, fd.memid)[0]
        if not fname:
            continue
        try:
            names = list(t.GetNames(fd.memid, fd.cParams + 1))
        except Exception:
            names = [fname]
        if fd.invkind == INVOKE_FUNC:
            vararg = 1 if fd.cParamsOpt == -1 else 0
            mid = db.add_method(iid, fname, _resolve_type(fd.elemdescFunc.tdesc, t), int(fd.wFuncFlags), vararg,
                                docs.member_doc(name, fname))
            for pos in range(fd.cParams):
                ed = fd.lprgelemdescParam[pos]
                pname = str(names[pos + 1]) if pos + 1 < len(names) and names[pos + 1] else f"p{pos}"
                db.add_param(mid, pname, _resolve_type(ed.tdesc, t), pos, int(ed._.paramdesc.wParamFlags),
                             _default_value(ed), docs.param_doc(name, fname, pname))
            continue
        p = props.setdefault(fname, {"type": None, "access": 0, "flags": 0, "order": j})
        p["flags"] |= int(fd.wFuncFlags)
        if fd.invkind == INVOKE_PROPERTYGET:
            p["access"] |= ACCESS_GET
            p["type"] = _resolve_type(fd.elemdescFunc.tdesc, t)
        else:
            p["access"] |= ACCESS_PUT if fd.invkind == INVOKE_PROPERTYPUT else ACCESS_PUTREF
            if p["type"] is None and fd.cParams:
                p["type"] = _resolve_type(fd.lprgelemdescParam[fd.cParams - 1].tdesc, t)
    for pname, p in sorted(props.items(), key=lambda kv: kv[1]["order"]):
        db.add_property(iid, pname, p["type"] or "Variant", p["access"], p["flags"], docs.member_doc(name, pname))
    return True


def _process_enum(t, a, lib_id: int, db: "TLBDatabase", docs) -> bool:
    """enum 1 つを書き込む。foo_<Framework> のダミーと、2 回目以降の同名 enum は入れない。"""
    name = _doc_of(t)[0]
    if not name or name.startswith("foo_"):
        return False
    members = []
    for v in range(a.cVars):
        vd = t.GetVarDesc(v)
        members.append((_doc_of(t, vd.memid)[0], int(vd._.lpvarValue.contents.value)))
    return db.add_enum(name, lib_id, docs.enum_doc(name),
                       [(mn, mv, docs.enum_member_doc(name, mn)) for mn, mv in members])


# ---------------------------------------------------------------------------
# SQLite データベースラッパー
# ---------------------------------------------------------------------------
def load_schema() -> str:
    with open(SCHEMA_PATH, encoding="utf-8") as f:
        return f.read()


class TLBDatabase:
    def __init__(self, path: str):
        self.path = path
        self.conn = sqlite3.connect(path)
        self.conn.executescript(load_schema())
        self._iface: dict[str, int] = {}
        self._parents: list[tuple[int, str]] = []
        self._enums: dict[str, tuple] = {}
        self.warnings: list[str] = []

    def warn(self, msg: str):
        self.warnings.append(msg)

    def add_library(self, name: str, doc: str, guid: str, version: str, file: str) -> int:
        return self.conn.execute(
            "INSERT INTO libraries(name, doc, guid, version, file) VALUES(?,?,?,?,?)",
            (name, doc, guid, version, file)).lastrowid

    def add_interface(self, name: str, lib_id: int, kind: str, flags: int, doc: str) -> Optional[int]:
        """インターフェースを登録して ID を返す。同名が登録済みなら None。"""
        if name in self._iface:
            return None
        iid = self.conn.execute(
            "INSERT INTO interfaces(name, library_id, kind, flags, doc) VALUES(?,?,?,?,?)",
            (name, lib_id, kind, flags, doc)).lastrowid
        self._iface[name] = iid
        return iid

    def set_parent(self, iid: int, parent: str):
        """親は別の TLB にあることが多いので、名前で覚えておき finish で解決する。"""
        self._parents.append((iid, parent))

    def add_method(self, iface_id: int, name: str, return_type: str, flags: int, vararg: int, doc: str) -> int:
        return self.conn.execute(
            "INSERT INTO methods(interface_id, name, return_type, flags, vararg, doc) VALUES(?,?,?,?,?,?)",
            (iface_id, name, return_type, flags, vararg, doc)).lastrowid

    def add_param(self, method_id: int, name: str, param_type: str, position: int, flags: int,
                  default_value: Optional[str], doc: str):
        self.conn.execute(
            "INSERT INTO parameters(method_id, name, type, position, flags, default_value, doc) VALUES(?,?,?,?,?,?,?)",
            (method_id, name, param_type, position, flags, default_value, doc))

    def add_property(self, iface_id: int, name: str, prop_type: str, access: int, flags: int, doc: str):
        self.conn.execute(
            "INSERT INTO properties(interface_id, name, type, access, flags, doc) VALUES(?,?,?,?,?,?)",
            (iface_id, name, prop_type, access, flags, doc))

    def add_enum(self, name: str, lib_id: int, doc: str, members: list[tuple[str, int, str]]) -> bool:
        """enum を登録する。同名が登録済みなら値の一致を確かめて False（不一致なら警告）。"""
        sig = tuple((n, v) for n, v, _ in members)
        if name in self._enums:
            if self._enums[name] != sig:
                self.warn(f"enum {name} が別の値で重複している（最初のものを採用）")
            return False
        self._enums[name] = sig
        eid = self.conn.execute("INSERT INTO enums(name, library_id, doc) VALUES(?,?,?)", (name, lib_id, doc)).lastrowid
        for n, v, d in members:
            self.conn.execute("INSERT INTO enum_members(enum_id, name, value, doc) VALUES(?,?,?,?)", (eid, n, v, d))
        return True

    def set_meta(self, key: str, value: str):
        self.conn.execute("INSERT OR REPLACE INTO meta(key, value) VALUES(?,?)", (key, value))

    def finish(self):
        """親を解決し、コミットして VACUUM で詰める（出力を決定的にする）。"""
        for iid, parent in self._parents:
            pid = self._iface.get(parent)
            if pid is None:
                self.warn(f"親 {parent} が見つからない（interface id {iid}）")
                continue
            self.conn.execute("UPDATE interfaces SET parent_id=? WHERE id=?", (pid, iid))
        self.conn.commit()
        self.conn.execute("VACUUM")

    def close(self):
        self.conn.close()

    def stats(self) -> dict:
        c = self.conn
        return {t: c.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                for t in ("libraries", "interfaces", "methods", "properties", "parameters", "enums", "enum_members")}


# ---------------------------------------------------------------------------
# 検査
# ---------------------------------------------------------------------------
REQUIRED_META = ("schema_version", "extractor_version", "catia_release", "tlb_count", "tlb_digest", "doc_source")


def verify_db(path: str) -> list[str]:
    """生成した DB の不変条件を調べ、違反の説明の一覧を返す（空なら合格）。DB は読み取り専用で開く。"""
    uri = "file:" + os.path.abspath(path).replace("\\", "/") + "?mode=ro"
    c = sqlite3.connect(uri, uri=True)
    errs: list[str] = []
    try:
        meta = dict(c.execute("SELECT key, value FROM meta"))
        for k in REQUIRED_META:
            if k not in meta:
                errs.append(f"meta.{k} が無い")
        if meta.get("schema_version") != SCHEMA_VERSION:
            errs.append(f"schema_version = {meta.get('schema_version')!r}")
        if meta.get("tlb_count") and int(meta["tlb_count"]) != c.execute("SELECT COUNT(*) FROM libraries").fetchone()[0]:
            errs.append("tlb_count と libraries の件数が違う")
        checks = {
            "親の参照切れ": "SELECT COUNT(*) FROM interfaces WHERE parent_id IS NOT NULL AND parent_id NOT IN (SELECT id FROM interfaces)",
            "ライブラリの参照切れ": "SELECT COUNT(*) FROM interfaces WHERE library_id IS NULL OR library_id NOT IN (SELECT id FROM libraries)",
            "メソッドの所属切れ": "SELECT COUNT(*) FROM methods WHERE interface_id NOT IN (SELECT id FROM interfaces)",
            "プロパティの所属切れ": "SELECT COUNT(*) FROM properties WHERE interface_id NOT IN (SELECT id FROM interfaces)",
            "引数の所属切れ": "SELECT COUNT(*) FROM parameters WHERE method_id NOT IN (SELECT id FROM methods)",
            "enum メンバーの所属切れ": "SELECT COUNT(*) FROM enum_members WHERE enum_id NOT IN (SELECT id FROM enums)",
            "名前が空": "SELECT (SELECT COUNT(*) FROM interfaces WHERE name='') + (SELECT COUNT(*) FROM methods WHERE name='')"
                       " + (SELECT COUNT(*) FROM properties WHERE name='') + (SELECT COUNT(*) FROM enum_members WHERE name='')",
            "restricted のメソッド": f"SELECT COUNT(*) FROM methods WHERE flags & {FUNCFLAG_FRESTRICTED}",
            "access が 0 のプロパティ": "SELECT COUNT(*) FROM properties WHERE access = 0 OR access & ~7",
            "同じ型に同名のメンバー": "SELECT COUNT(*) FROM (SELECT 1 FROM (SELECT interface_id AS i, lower(name) AS n FROM methods"
                             " UNION ALL SELECT interface_id, lower(name) FROM properties) GROUP BY i, n HAVING COUNT(*) > 1)",
            "ダミーの enum（foo_*）": "SELECT COUNT(*) FROM enums WHERE name LIKE 'foo\\_%' ESCAPE '\\'",
        }
        for label, sql in checks.items():
            row = c.execute(sql).fetchone()
            n = row[0] if row else 0
            if n:
                errs.append(f"{label}: {n} 件")
        parent = dict(c.execute("SELECT id, parent_id FROM interfaces"))
        for iid in parent:
            seen, cur = set(), iid
            while cur is not None:
                if cur in seen or len(seen) > 20:
                    errs.append(f"継承が循環しているか深すぎる（interface id {iid}）")
                    break
                seen.add(cur)
                cur = parent.get(cur)
    finally:
        c.close()
    return errs


# ---------------------------------------------------------------------------
# TLB ファイル検索
# ---------------------------------------------------------------------------
CATIA_TLB_DEFAULT_PATTERNS = [
    r"C:\Program Files\Dassault Systemes\B*\win_b64\code\bin\*.tlb",
    r"C:\Program Files\Dassault Systemes\B*\intel_a\code\bin\*.tlb",
    r"C:\Program Files (x86)\Dassault Systemes\B*\win_b64\code\bin\*.tlb",
]


def find_tlb_files(paths: list[str], recursive: bool = False) -> list[str]:
    """引数リストから TLB ファイルのパス一覧を返す。"""
    result = []
    seen: set[str] = set()

    def add(p: str):
        p = os.path.abspath(p)
        if p not in seen:
            seen.add(p)
            result.append(p)

    for raw in paths:
        p = os.path.abspath(os.path.expandvars(os.path.expanduser(raw)))
        if os.path.isdir(p):
            pattern = "**/*.tlb" if recursive else "*.tlb"
            for f in glob.glob(os.path.join(p, pattern), recursive=recursive):
                add(f)
        elif "*" in raw or "?" in raw:
            for f in glob.glob(raw, recursive=recursive):
                add(f)
        else:
            add(p)

    return sorted(result)


def find_catia_tlbs_auto() -> list[str]:
    """CATIA の標準インストールパスから TLB を自動検索する。"""
    result = []
    seen: set[str] = set()
    for pat in CATIA_TLB_DEFAULT_PATTERNS:
        for f in glob.glob(pat, recursive=False):
            a = os.path.abspath(f)
            if a not in seen:
                seen.add(a)
                result.append(a)
    return sorted(result)


def _sorted_by_name(files: list[str]) -> list[str]:
    """ファイル名（大文字小文字無視）→ パスの順に並べる（同名 enum の採用順と digest の順序を決める）。"""
    return sorted(files, key=lambda p: (os.path.basename(p).lower(), p.lower()))


def tlb_digest(files: list[str]) -> str:
    """各 TLB の SHA-256 をファイル名順に「名前:ハッシュ」で連結した SHA-256。"""
    h = hashlib.sha256()
    for f in _sorted_by_name(files):
        with open(f, "rb") as fp:
            fh = hashlib.sha256(fp.read()).hexdigest()
        h.update(f"{os.path.basename(f).lower()}:{fh}\n".encode("utf-8"))
    return h.hexdigest()


def catia_release(files: list[str]) -> str:
    """パスに含まれる CATIA のリリース（B32 など）。複数あればカンマ区切り、無ければ ''。"""
    rels = set()
    for f in files:
        m = re.search(r"[\\/](B\d+)[\\/]", f)
        if m:
            rels.add(m.group(1))
    return ",".join(sorted(rels))


# ---------------------------------------------------------------------------
# 生成
# ---------------------------------------------------------------------------
def build_db(tlb_files: list[str], out: str, docs=None, log=print) -> dict:
    """tlb_files から out に DB を作る（既存の out は置き換える）。統計を返す。"""
    docs = docs or NoDocs()
    files = _sorted_by_name(tlb_files)
    if os.path.exists(out):
        os.remove(out)
    db = TLBDatabase(out)
    failed = []
    try:
        for idx, path in enumerate(files, 1):
            fname = os.path.basename(path)
            r = _process_tlb(path, db, docs)
            if r is None:
                failed.append(path)
                log(f"  [{idx:3d}/{len(files)}] {fname}: 読み込み失敗")
                continue
            n_types, n_enums = r
            log(f"  [{idx:3d}/{len(files)}] {fname}: 型 {n_types} / enum {n_enums}")
        db.set_meta("schema_version", SCHEMA_VERSION)
        db.set_meta("extractor_version", EXTRACTOR_VERSION)
        db.set_meta("catia_release", catia_release(files))
        db.set_meta("tlb_count", str(len(files) - len(failed)))
        db.set_meta("tlb_digest", tlb_digest([f for f in files if f not in failed]))
        db.set_meta("doc_source", docs.source)
        db.finish()
        stats = db.stats()
    finally:
        db.close()
    stats["failed"] = failed
    stats["warnings"] = db.warnings
    return stats


# ---------------------------------------------------------------------------
# メインエントリ
# ---------------------------------------------------------------------------
def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser(
        description="CATIA V5 TLB → SQLite 変換ツール（スキーマ v2）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    ap.add_argument("paths", nargs="*", help="TLB ファイルまたはディレクトリ（省略時は自動検索）")
    ap.add_argument("--out", default="catia_api.db", help="出力 SQLite ファイルパス")
    ap.add_argument("--recursive", "-r", action="store_true", help="ディレクトリを再帰検索")
    ap.add_argument("--list", action="store_true", help="処理対象 TLB を一覧表示して終了")
    ap.add_argument("--verify", metavar="DB", help="生成済み DB の不変条件を検査して終了")
    args = ap.parse_args()

    if args.verify:
        errs = verify_db(args.verify)
        for e in errs:
            print(f"NG: {e}")
        print("OK" if not errs else f"{len(errs)} 件の違反")
        sys.exit(1 if errs else 0)

    if args.paths:
        tlb_files = find_tlb_files(args.paths, recursive=args.recursive)
    else:
        print("CATIA TLB ファイルを自動検索中...", flush=True)
        tlb_files = find_catia_tlbs_auto()

    if not tlb_files:
        print("ERROR: TLB ファイルが見つかりませんでした。", file=sys.stderr)
        print("  パスを引数で指定してください:", file=sys.stderr)
        print('  uv run python extract.py "C:/Program Files/Dassault Systemes/B32/win_b64/code/bin"', file=sys.stderr)
        sys.exit(1)

    print(f"対象 TLB ファイル: {len(tlb_files)} 件", flush=True)
    if args.list:
        for f in tlb_files:
            print(f"  {f}")
        return

    stats = build_db(tlb_files, args.out, log=lambda s: print(s, flush=True))
    for w in stats["warnings"]:
        print(f"警告: {w}")
    errs = verify_db(args.out)
    for e in errs:
        print(f"検査 NG: {e}", file=sys.stderr)

    print("\n完了:")
    print(f"  TLB:              {len(tlb_files) - len(stats['failed'])} 成功 / {len(stats['failed'])} 失敗")
    print(f"  インターフェース: {stats['interfaces']:,}")
    print(f"  メソッド:         {stats['methods']:,}")
    print(f"  プロパティ:       {stats['properties']:,}")
    print(f"  Enum:             {stats['enums']:,}（メンバー {stats['enum_members']:,}）")
    print(f"  DB サイズ:        {os.path.getsize(args.out) / 1024 / 1024:.2f} MB")
    print(f"  出力先:           {os.path.abspath(args.out)}")
    if errs:
        sys.exit(1)


if __name__ == "__main__":
    main()
