"""extract.py のテスト（標準ライブラリの unittest のみ）。

実行（リポジトリに .venv / __pycache__ を作らない）:
    PYTHONDONTWRITEBYTECODE=1 UV_PROJECT_ENVIRONMENT=<リポジトリ外の venv> \
        uv run --project tools/tlb-extractor python -m unittest discover -s tools/tlb-extractor

実 TLB を使うテストは CATIA V5 がインストールされていなければ skip する。
"""

from __future__ import annotations

import glob
import hashlib
import os
import shutil
import sqlite3
import tempfile
import unittest
from types import SimpleNamespace

import extract


def _sha256(path: str) -> str:
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


class PureFunctionTest(unittest.TestCase):
    def test_vba_literal(self):
        self.assertEqual(extract.vba_literal('a"b'), '"a""b"')
        self.assertEqual(extract.vba_literal(""), '""')
        self.assertEqual(extract.vba_literal(True), "True")
        self.assertEqual(extract.vba_literal(False), "False")
        self.assertEqual(extract.vba_literal(0), "0")
        self.assertEqual(extract.vba_literal(-3), "-3")
        self.assertEqual(extract.vba_literal(1.5), "1.5")
        self.assertIsNone(extract.vba_literal(None))
        self.assertIsNone(extract.vba_literal(object()))

    def test_default_value(self):
        def ed(flags, value=None, has_ex=True):
            ex = SimpleNamespace(contents=SimpleNamespace(varDefaultValue=SimpleNamespace(value=value))) if has_ex else None
            return SimpleNamespace(_=SimpleNamespace(paramdesc=SimpleNamespace(wParamFlags=flags, pparamdescex=ex)))

        self.assertEqual(extract._default_value(ed(extract.PARAMFLAG_FHASDEFAULT | 1, "x")), '"x"')
        self.assertEqual(extract._default_value(ed(extract.PARAMFLAG_FHASDEFAULT, 5)), "5")
        self.assertIsNone(extract._default_value(ed(1, "x")))  # FHASDEFAULT なし
        self.assertIsNone(extract._default_value(ed(extract.PARAMFLAG_FHASDEFAULT, has_ex=False)))

    def test_catia_release(self):
        self.assertEqual(extract.catia_release([r"C:\Program Files\Dassault Systemes\B32\win_b64\code\bin\a.tlb"]), "B32")
        self.assertEqual(extract.catia_release(["C:/x/B28/a.tlb", "C:/y/B32/b.tlb", "C:/y/B32/c.tlb"]), "B28,B32")
        self.assertEqual(extract.catia_release(["C:/tmp/a.tlb"]), "")

    def test_tlb_digest_is_order_independent(self):
        d = tempfile.mkdtemp()
        try:
            a, b = os.path.join(d, "A.tlb"), os.path.join(d, "b.tlb")
            for p, data in ((a, b"1"), (b, b"2")):
                with open(p, "wb") as f:
                    f.write(data)
            self.assertEqual(extract.tlb_digest([a, b]), extract.tlb_digest([b, a]))
            with open(b, "wb") as f:
                f.write(b"3")
            self.assertNotEqual(extract.tlb_digest([a, b]), extract.tlb_digest([b, a]) + "x")
        finally:
            shutil.rmtree(d)


_IFACE_PAGE = """<html><body>
<section id='maindesc'>
<b><i>   Represents the
 factory.</i></b><br>
 <b>Role</b>: creates a <script type="text/javascript"> activateLinkAnchor('ShapeInstance','','ShapeInstance') </script>
 instance.<br>Second &amp; line.
</section>
<section id='idxsum'>
 <tr><td  id='tdleft'>
    <a href="#Mode"><b>Mode</b></a>
</td><td  id='tdright'>  Sets the
 mode.
</td></tr>
 <tr><td  id='tdleft'>
    <a href="#Run"><b>Run</b></a>
</td><td  id='tdright'>  Runs it.
</td></tr>
</section>
<a id="Mode"></a><dt class='dtmethod'>
<span class='property'>Property Mode</span></dt>
  <dd class='ddmethod'>
 Sets the mode of
 instantiation.<br><b>Role</b>: call it first.
 <! @sample >
 <dl><dt><b>Example:</b><dd>ignored</dl>
</dd>
<a id="Run"></a><dt class='dtmethod'>
<span class='property'>Sub Run(<script type="text/javascript"> activateLink('CATBSTR','CATBSTR')</script> iName)</span></dt>
  <dd class='ddmethod'>
 Runs it.
<dl  class='desclist'>
  <dt> <b>Parameters:</b></dt>
  <dd>
    <dl class='paramlist' >
    <dt><span class='paramitem'> iName</span></dt>
   <dd>   The name of
 the <script>activateLink('Document','Document')</script>.
 </ dd>
    <dt><span class='paramitem'> iLast</span></dt>
   <dd>  Last one.
 <! @sample >
 </dl>
 </dl></dd>
</body></html>"""

_ENUM_PAGE = """</dt><dd class='dd1'>
<b><i> Paper size.</i></b><br>
 It used by the <script type="text/javascript"> activateLinkAnchor('PageSetup','','PageSetup') </script> object.
<dl  class='desclist'>
  <dt> <b>Values:</b></dt>
  <dd>
    <dl class='valuelist'>
    <dt><span class='paramitem'> catPaperA4</span></dt>
   <dd>   The A4 ISO format (210 x 297 mm)
 </ dd>
    <dt><span class='paramitem'> catPaperA3</span></dt>
   <dd>   The A3 format
 </ dd>
 </dl>"""


class ChmTest(unittest.TestCase):
    def test_html_text(self):
        self.assertEqual(extract.html_text("a\n b<br>c &lt;d&gt;"), "a b\nc <d>")
        self.assertEqual(extract.html_text("see <script type='text/javascript'> activateLink('Doc','Document')</script>."),
                         "see Document.")
        self.assertEqual(extract.html_text("x<p><p><p>y"), "x\n\ny")
        self.assertEqual(extract.html_text("<ul><li>one</li><li>two</li></ul>"), "one\ntwo")
        self.assertEqual(extract.html_text("  \n "), "")

    def _docs(self, full):
        d = extract.ChmDocs.__new__(extract.ChmDocs)
        d.full = full
        return d

    def test_interface_summary(self):
        desc, members, params = self._docs(False).parse_interface(_IFACE_PAGE)
        self.assertEqual(desc, "Represents the factory.")
        self.assertEqual(members, {"Mode": "Sets the mode.", "Run": "Runs it."})
        self.assertEqual(params, {})

    def test_interface_full(self):
        desc, members, params = self._docs(True).parse_interface(_IFACE_PAGE)
        self.assertEqual(desc, "Represents the factory.\nRole: creates a ShapeInstance instance.\nSecond & line.")
        self.assertEqual(members["Mode"], "Sets the mode of instantiation.\nRole: call it first.")
        self.assertEqual(members["Run"], "Runs it.")
        self.assertEqual(params, {("Run", "iName"): "The name of the Document.", ("Run", "iLast"): "Last one."})

    def test_enum(self):
        desc, vals = self._docs(False).parse_enum(_ENUM_PAGE)
        self.assertEqual((desc, vals), ("Paper size.", {}))
        desc, vals = self._docs(True).parse_enum(_ENUM_PAGE)
        self.assertEqual(desc, "Paper size.\nIt used by the PageSetup object.")
        self.assertEqual(vals, {"catPaperA4": "The A4 ISO format (210 x 297 mm)", "catPaperA3": "The A3 format"})

    def test_chm_directory(self):
        d = tempfile.mkdtemp()
        try:
            fw = os.path.join(d, "generated", "interfaces", "InfInterfaces")
            os.makedirs(fw)
            with open(os.path.join(fw, "interface_InstanceFactory_1.htm"), "w", encoding="cp1252") as f:
                f.write(_IFACE_PAGE)
            with open(os.path.join(fw, "enum_CatPaperSize_2.htm"), "w", encoding="cp1252") as f:
                f.write(_ENUM_PAGE)
            docs = extract.load_chm_docs(d, full=True, log=lambda s: None)
            self.assertEqual(docs.source, "chm-full")
            self.assertEqual(docs.member_doc("InstanceFactory", "Run"), "Runs it.")
            self.assertEqual(docs.param_doc("InstanceFactory", "Run", "iName"), "The name of the Document.")
            self.assertEqual(docs.enum_member_doc("CatPaperSize", "catPaperA3"), "The A3 format")
            self.assertEqual(docs.type_doc("Nothing"), "")
            self.assertEqual(extract.ChmDocs(d, full=False).source, "chm-summary")
        finally:
            shutil.rmtree(d)


class DatabaseTest(unittest.TestCase):
    """TLBDatabase と verify_db（合成データ）。"""

    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.path = os.path.join(self.dir, "t.db")

    def tearDown(self):
        shutil.rmtree(self.dir)

    def _build(self, **kw):
        db = extract.TLBDatabase(self.path)
        lib = db.add_library("INFITF", "doc", "{G}", "0.0", "Inf.tlb")
        base = db.add_interface("AnyObject", lib, "dispatch", 0x40, "")
        app = db.add_interface("Application", lib, "dispatch", 0x40, "app")
        self.assertIsNone(db.add_interface("Application", lib, "dispatch", 0, ""))  # 同名の 2 回目
        db.set_parent(app, "AnyObject")
        if kw.get("missing_parent"):
            db.set_parent(base, "Nowhere")
        m = db.add_method(app, "Quit", "void", 0, 0, "")
        db.add_param(m, "iMode", "Long", 0, 1, None, "")
        db.add_property(app, "Name", "String", 1, 0, "")
        self.assertTrue(db.add_enum("CatX", lib, "", [("catA", 0, ""), ("catB", 1, "")]))
        self.assertFalse(db.add_enum("CatX", lib, "", [("catA", 0, ""), ("catB", 1, "")]))
        self.assertFalse(db.add_enum("CatX", lib, "", [("catA", 0, ""), ("catB", 2, "")]))
        for k in extract.REQUIRED_META:
            db.set_meta(k, "2" if k == "schema_version" else "1" if k == "tlb_count" else "x")
        db.finish()
        db.close()
        return db

    def test_build_and_verify(self):
        db = self._build()
        self.assertEqual(len(db.warnings), 1, db.warnings)  # 値の違う同名 enum
        self.assertEqual(extract.verify_db(self.path), [])
        c = sqlite3.connect(self.path)
        try:
            self.assertEqual(c.execute("SELECT p.name FROM interfaces i JOIN interfaces p ON p.id = i.parent_id"
                                       " WHERE i.name = 'Application'").fetchone(), ("AnyObject",))
        finally:
            c.close()

    def test_missing_parent_is_warned(self):
        db = self._build(missing_parent=True)
        self.assertTrue(any("Nowhere" in w for w in db.warnings), db.warnings)

    def test_verify_detects_violations(self):
        self._build()
        c = sqlite3.connect(self.path)
        c.execute("UPDATE interfaces SET parent_id = id WHERE name = 'AnyObject'")  # 循環
        c.execute("INSERT INTO enums(name, library_id) VALUES ('foo_X', 1)")
        c.execute("INSERT INTO properties(interface_id, name, type, access) VALUES (99, 'Bad', 'Long', 0)")
        c.execute("UPDATE meta SET value = '1' WHERE key = 'schema_version'")
        c.commit()
        c.close()
        errs = "\n".join(extract.verify_db(self.path))
        for want in ("循環", "foo_", "所属切れ", "access", "schema_version"):
            self.assertIn(want, errs)


def _catia_tlbs() -> list[str]:
    return extract.find_catia_tlbs_auto()


@unittest.skipUnless(_catia_tlbs(), "CATIA V5 の TLB が見つからない")
class RealTLBTest(unittest.TestCase):
    """インストール済みの CATIA の TLB から生成した DB を検査する。"""

    @classmethod
    def setUpClass(cls):
        cls.dir = tempfile.mkdtemp()
        cls.files = _catia_tlbs()
        cls.path = os.path.join(cls.dir, "a.db")
        cls.stats = extract.build_db(cls.files, cls.path, log=lambda s: None)
        cls.c = sqlite3.connect("file:" + cls.path.replace("\\", "/") + "?mode=ro", uri=True)

    @classmethod
    def tearDownClass(cls):
        cls.c.close()
        shutil.rmtree(cls.dir)

    def q(self, sql, *args):
        return self.c.execute(sql, args).fetchall()

    def parent_of(self, name):
        r = self.q("SELECT p.name FROM interfaces i LEFT JOIN interfaces p ON p.id = i.parent_id WHERE i.name = ?", name)
        return r[0][0] if r else "<none>"

    def test_verify(self):
        self.assertEqual(extract.verify_db(self.path), [])
        self.assertEqual(self.stats["failed"], [])
        self.assertEqual(self.stats["warnings"], [])

    def test_meta_and_libraries(self):
        meta = dict(self.q("SELECT key, value FROM meta"))
        self.assertEqual(meta["schema_version"], "2")
        self.assertEqual(meta["doc_source"], "none")
        self.assertEqual(int(meta["tlb_count"]), len(self.files))
        self.assertEqual(self.q("SELECT COUNT(*) FROM libraries")[0][0], len(self.files))
        names = {n for (n,) in self.q("SELECT name FROM libraries")}
        self.assertTrue({"INFITF", "MECMOD"} <= names, names)

    def test_hierarchy(self):
        self.assertEqual(self.parent_of("Application"), "AnyObject")
        self.assertEqual(self.parent_of("AnyObject"), "CATBaseDispatch")
        self.assertEqual(self.parent_of("CATBaseDispatch"), "CATBaseUnknown")
        self.assertIsNone(self.parent_of("CATBaseUnknown"))
        self.assertEqual(self.parent_of("PartDocument"), "Document")
        parent = dict(self.q("SELECT id, parent_id FROM interfaces"))
        deepest = 0
        for iid in parent:
            d, cur = 0, iid
            while cur is not None:
                d += 1
                cur = parent[cur]
            deepest = max(deepest, d)
        self.assertLessEqual(deepest, 9)  # CATBaseUnknown から最大 8 段

    def test_restricted_members_are_excluded(self):
        restricted = ["QueryInterface", "AddRef", "Release", "GetTypeInfoCount", "GetIDsOfNames", "Invoke",
                      "Destructor", "GetMetaObject", "IsA", "IsAKindOf", "GetImpl", "SetImpl", "IsNull", "IsEqual",
                      "SurchargeQI", "ChangeComponentState", "_NewEnum"]
        marks = ",".join("?" * len(restricted))
        rows = self.q(f"SELECT i.name, m.name FROM methods m JOIN interfaces i ON i.id = m.interface_id"
                      f" WHERE m.name IN ({marks}) UNION ALL"
                      f" SELECT i.name, p.name FROM properties p JOIN interfaces i ON i.id = p.interface_id"
                      f" WHERE p.name IN ({marks})", *restricted, *restricted)
        # 正当な同名メンバー（CATIA のリリース番号）だけが残る
        self.assertEqual(rows, [("SystemConfiguration", "Release")])
        self.assertTrue(self.q("SELECT 1 FROM methods m JOIN interfaces i ON i.id = m.interface_id"
                               " WHERE i.name = 'TreeVizManipSettingAtt' AND m.name = 'GetTypeInfo'"))

    def test_write_only_property(self):
        rows = self.q("SELECT p.type, p.access FROM properties p JOIN interfaces i ON i.id = p.interface_id"
                      " WHERE i.name = 'InstanceFactory' AND p.name = 'InstantiationMode'")
        self.assertEqual(rows, [("String", extract.ACCESS_PUT)])

    def test_enums(self):
        rows = self.q("SELECT m.name, m.value FROM enum_members m JOIN enums e ON e.id = m.enum_id"
                      " WHERE e.name = 'CatPaperSize' ORDER BY m.id")
        self.assertEqual(len(rows), 14)
        self.assertIn(("catPaperA4", 6), rows)
        self.assertEqual(self.q("SELECT COUNT(*) FROM enums WHERE name LIKE 'foo%'")[0][0], 0)

    def test_flattened_members_match_typelib(self):
        """各 dispinterface の非 restricted の関数名の集合 ＝ DB の祖先の自前メンバー名の和集合。"""
        own: dict[int, set] = {}
        for iid, n in self.q("SELECT interface_id, name FROM methods UNION ALL SELECT interface_id, name FROM properties"):
            own.setdefault(iid, set()).add(n)
        ids = {n: (i, p) for i, n, p in self.q("SELECT id, name, parent_id FROM interfaces")}
        byid = {i: p for i, p in ids.values()}

        def flattened(name):
            s, cur = set(), ids[name][0]
            while cur is not None:
                s |= own.get(cur, set())
                cur = byid[cur]
            return s

        checked = 0
        for path in self.files:
            tlb = extract._load_tlb(path)
            for i in range(tlb.GetTypeInfoCount()):
                t = tlb.GetTypeInfo(i)
                a = t.GetTypeAttr()
                if a.typekind != extract.TKIND_DISPATCH:
                    continue
                want = set()
                for j in range(a.cFuncs):
                    fd = t.GetFuncDesc(j)
                    if not fd.wFuncFlags & extract.FUNCFLAG_FRESTRICTED:
                        want.add(extract._doc_of(t, fd.memid)[0])
                name = extract._doc_of(t)[0]
                self.assertEqual(flattened(name), want, name)
                checked += 1
        self.assertGreater(checked, 1000)

    def test_reproducible(self):
        other = os.path.join(self.dir, "b.db")
        extract.build_db(list(reversed(self.files)), other, log=lambda s: None)
        self.assertEqual(_sha256(other), _sha256(self.path))


if __name__ == "__main__":
    unittest.main()
