// eval-probe.exe（開発者用・非出荷）: CATIA.SystemService.Evaluate で VBA コード文字列を CATIA 内実行できるかを試す。
// Build: node scripts/probe-evaluate.js が temp/eval-probe.exe にビルドする（.NET Framework 4.x の csc.exe、C# 5 の構文に留める）。
// Usage: eval-probe.exe [--byref] [--project <name>] [--dry-run] <step> [<step> ...]
//        eval-probe.exe [--byref] [--dry-run] [--label <name>] [--no-retry] [<step> ...] vbacall <library> <module> <function> [args...]
//        eval-probe.exe --print-code <step>   （CATIA に接続せず、そのステップのコードを出力するだけ。テスト用）
//   vbaproj-* ステップは直後の引数を .catvba のパス（ExecuteScript の iLibraryName）として取る:
//        eval-probe.exe vbaproj-ping <libraryPath>
//   --dry-run: CATIA に接続せず、各ステップの呼び出し内容（method/library/program/function）を 1 行 JSON で出す（テスト用）
//   --project: vbaproj-filename-project で iLibraryName に使うプロジェクト名（既定 toQI）
//   steps: check basic sub typed params apc apc2 apc2-getobject apc2-catia apc2-application
//          jscript jscript-params jscript-fso jscript-apc jscript-read
//   (apc2 = apc2-getobject + apc2-catia + apc2-application)
//   ExecuteScript(iLibraryName, 2 = catScriptLibraryTypeVBAProject, "C5D_BridgeTest", iFunctionName, iParameters) の系列
//   （VBProject に標準モジュール C5D_BridgeTest が入っている前提。scripts/probe-vbaproject.js から呼ぶ）:
//     vbaproj-ping      C5D_Ping("cat5dev-param") -> 期待 "vba:Long:cat5dev-param"（ByVal で失敗/不一致なら ByRef で再試行）
//     vbaproj-sub       C5D_Noop（Sub、引数なし・戻り値なし）
//     vbaproj-projects  C5D_Projects -> 期待 "count=N|name1|...;err=N"
//     vbaproj-filename  = vbaproj-filename-basename（ファイル名のみ）+ -empty（空文字）+ -project（プロジェクト名）で C5D_Ping
//   汎用: vbacall <library> <module> <function> [args...]   （scripts/probe-bridge.js から呼ぶ）
//     ExecuteScript(library, 2, module, function, [args...]) を 1 回呼ぶ。args は文字列 0 個以上（ByVal で例外なら ByRef で再試行）。
//     vbacall 以降の引数は全て vbacall のもの（-- で始まっても VBA への引数）なので、オプションは vbacall より前に書く。
//     各トークンは "b64:<UTF-8 の base64>" でも渡せる（非 ASCII・改行・引用符をコマンドラインで化けさせないため）。
//     期待値は持たない（match = null）。判定は呼び出し側が行う。--label <name> で出力 JSON の step 名を変えられる（既定 vbacall）。
//   --no-retry: vbacall の ByRef 再試行をしない（ExecuteScript をちょうど 1 回だけ呼ぶ。アンロード等の副作用がある関数用）
//   オブジェクト戻り値: vbacall-obj <library> <module> <function> [<action> ...]   （scripts/probe-bridge.js --remove-experiment から呼ぶ）
//     引数なしの VBA 関数を ExecuteScript で 1 回だけ呼び（ByRef 再試行なし）、戻り値（VT_DISPATCH なら COM プロキシ）を保持したまま、
//     ExecuteScript が戻った後（VBA の実行が終わった後）に、この exe から IDispatch のレイトバインドで action を順に実行する。
//     action（各 1 トークン、b64: 可）:
//       get:<expr>             式の値を読む（例 get:$.Name / get:$.VBProject.Type / get:$.Projects.Item('C5D_Bridge').Mode）
//       type:<expr>            式のオブジェクトの型名（IDispatch.GetTypeInfo の名前）
//       call:<expr>            最後の区間がメソッド呼び出しの式（例 call:$.Collection.Remove($) / call:$.Close(false) / call:$.Apc.End()）
//       let:<expr>=<arg>       プロパティに値を入れる（例 let:$.Mode=0）
//       sleep:<ms>             待つ（CATIA 側の処理を待つため）
//     <expr> = $ | $<n>（戻り値が配列ならその n 番目、0 始まり）に .Name または .Name(<arg>, ...) を続けたもの。
//     <arg> = <expr> | true | false | 整数 | 'text'（'' で ' 自身）| missing（省略引数）
//     出力: 1 行目は ExecuteScript の結果（resultType = ComObject なら VT_DISPATCH/VT_UNKNOWN で返った。isDispatch / typeInfo / arrayLength も出す）、
//     続いて action ごとに {"step","action":<番号>,"op","kind","ok","result","resultType","typeInfo","hresult","message","elapsedMs"}。
//     ExecuteScript が例外ならアクションは実行しない（"skipped" 行を出す）。action の失敗は後続を止めない。
//   eval-probe.exe --obj-selftest <progid> [<action> ...]   （CATIA に接続せず、CreateObject(progid) を戻り値の代わりにして action を実行する。テスト用）
//   --byref: iParameters を VT_BYREF で渡す（params / jscript-params は指定がなくても ByVal 失敗時に ByRef で再試行する）
//   jscript*: 言語 4（JScript）。ES3 構文のみ・全処理を try/catch で囲み、例外は結果文字列に含めて返す。書き込み系 API は呼ばない。
// Output: ステップごとに 1 行 JSON（UTF-8）
//   {"step","language","ok","match","result","resultType","hresult","message","paramMode","attempts","elapsedMs"}
//   ok = Evaluate が例外なく戻った, match = 戻り値が期待どおり（期待値のないステップは null）
// Exit codes: 0 = 全ステップ ok（vbacall-obj は全 action も ok）, 1 = いずれか失敗, 2 = 引数不正, 10 = CATIA 取得失敗, 11 = SystemService 取得失敗
// 注意: VBA のコンパイル/実行時エラーで CATIA がモーダルダイアログを出して止まる恐れがあるため、
//   コードは全て On Error Resume Next 付き・保守的な構文・MsgBox なしにしている。CATIA は読み取りのみ。
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

internal static class EvalProbe
{
    const int Ok = 0, StepFailed = 1, BadArguments = 2, NoCatia = 10, NoSystemService = 11;
    // CATScriptLanguage (bin/catia_api.db)
    const int VBScriptLanguage = 0, VbaLanguage = 1, BasicScriptLanguage = 2, JavaLanguage = 3, JScriptLanguage = 4;
    // CATScriptLibraryType: 2 = catScriptLibraryTypeVBAProject（iLibraryName = .catvba、iProgramName = モジュール名）
    const int LibraryTypeVbaProject = 2;
    const string EvaluateMethod = "Evaluate", ExecuteScriptMethod = "ExecuteScript";
    const string BridgeModule = "C5D_BridgeTest", PingArgument = "cat5dev-param";
    const string NL = "\r\n";

    static readonly TextWriter Out = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };

    sealed class Step
    {
        public string Name;
        public int Language;
        public string Code;
        public object[] Parameters = new object[0];
        public string Expected;       // null = 期待値なし
        public bool ExpectedPrefix;   // true = Expected で始まれば一致
        public Regex ExpectedPattern; // Expected の代わりに正規表現で判定（null = 使わない）
        public bool RetryByRef;       // ByVal で失敗/不一致なら ByRef で再試行
        public string Method = EvaluateMethod;
        public string Library;        // ExecuteScript: iLibraryName
        public string ProgramName;    // ExecuteScript: iProgramName（モジュール名）
        public string FunctionName;   // ExecuteScript: iFunctionName
        public List<ObjAction> Actions; // vbacall-obj: 戻り値に対する操作（null = 通常のステップ）
        public string[] ActionTexts;    // vbacall-obj: 受け取った action の文字列（そのまま）

        public bool IsCheck { get { return Method == EvaluateMethod && Code == null; } }

        public bool HasExpectation { get { return Expected != null || ExpectedPattern != null; } }
    }

    // ---- JScript（言語 4）。ES3 構文（var / function / 配列 join）。単一引用符で書き、C# の逐語的文字列では " を "" と書く。
    // 共通関数: hex(n, width), errText(e) = "err=0xXXXXXXXX message"（| ; 改行は空白に置換）,
    //   q(s) = JSON 文字列リテラル（" \ 制御文字 U+2028/2029 サロゲートを \uXXXX 等でエスケープ、日本語等はそのまま）,
    //   getApc(log) = MSAPC.Apc.7.1 → 6.2 → MSAPC.Apc の順に new ActiveXObject し、各段階の結果を log に積む。
    const string JsPrelude = @"
function hex(n, width) {
    var s = (n >>> 0).toString(16).toUpperCase();
    while (s.length < width) { s = '0' + s; }
    return s;
}
function errText(e) {
    var n = 0, m = '';
    try { n = e.number; m = e.message || e.description; } catch (x) { }
    if (typeof n != 'number') { n = 0; }
    if (!m) { try { m = String(e); } catch (y) { m = '?'; } }
    return 'err=0x' + hex(n, 8) + ' ' + String(m).replace(/[|;\r\n]+/g, ' ');
}
function q(s) {
    var out = '""', i, c, ch;
    s = String(s);
    for (i = 0; i < s.length; i++) {
        ch = s.charAt(i);
        c = s.charCodeAt(i);
        if (ch == '""') { out += '\\""'; }
        else if (ch == '\\') { out += '\\\\'; }
        else if (c < 32 || c == 0x2028 || c == 0x2029 || (c >= 0xD800 && c <= 0xDFFF)) { out += '\\u' + hex(c, 4); }
        else { out += ch; }
    }
    return out + '""';
}
function str(v) { return v == null ? 'null' : q(v); }
function num(v) { return (typeof v == 'number' && isFinite(v)) ? String(v) : 'null'; }
function getApc(log) {
    var ids = ['MSAPC.Apc.7.1', 'MSAPC.Apc.6.2', 'MSAPC.Apc'], i, apc;
    for (i = 0; i < ids.length; i++) {
        try {
            apc = new ActiveXObject(ids[i]);
            log.push(ids[i] + '=ok');
            return apc;
        } catch (e) { log.push(ids[i] + '=' + errText(e)); }
    }
    return null;
}
";

    const string JsParamsCode = @"
function CATMain(a) {
    try { return 'param:' + String(a); } catch (e) { return 'param-error:' + errText(e); }
}
";

    // 作成のみ。メソッドは呼ばずファイルも書かない。
    const string JsFsoCode = @"
function tryCreate(id) {
    var o = null, t;
    try {
        o = new ActiveXObject(id);
        t = typeof o;
        o = null;
        return id + '=ok(' + t + ')';
    } catch (e) { return id + '=' + errText(e); }
}
function CATMain() {
    var r = [];
    try {
        r.push(tryCreate('ADODB.Stream'));
        r.push(tryCreate('Scripting.FileSystemObject'));
    } catch (e) { r.push('unexpected ' + errText(e)); }
    return r.join('|');
}
";

    // 結果: count=N|name1|name2...;<取得ログ>（失敗時 count=-1|<失敗箇所 err=...>;<取得ログ>）
    const string JsApcCode = @"
function CATMain() {
    var log = [], apc = null, projects = null, n = -1, names = [], i;
    try {
        apc = getApc(log);
        if (apc == null) { return 'count=-1|apc=null;' + log.join(';'); }
        try { projects = apc.VBE.VBProjects; } catch (e1) { return 'count=-1|VBE.VBProjects ' + errText(e1) + ';' + log.join(';'); }
        try { n = projects.Count; } catch (e2) { return 'count=-1|VBProjects.Count ' + errText(e2) + ';' + log.join(';'); }
        for (i = 1; i <= n; i++) {
            try { names.push(String(projects.Item(i).Name)); } catch (e3) { names.push('#' + i + ' ' + errText(e3)); }
        }
        return 'count=' + n + (names.length > 0 ? '|' + names.join('|') : '') + ';' + log.join(';');
    } catch (e) { return 'count=-1|unexpected ' + errText(e) + ';' + log.join(';'); }
}
";

    // 結果は JSON 文字列。先頭 3 プロジェクトのコンポーネント（各 50 件まで）を列挙し、
    // 最初に CountOfLines > 0 のモジュール 1 つで Lines(1, min(3, n)) を読む。
    // Lines は引数付きプロパティなので、呼び出し形式 cm.Lines(1, k) と 1 行ずつ cm.Lines(i, 1) の両方を試し、
    // via（lines-call / lines-perline / none）と attempts に結果を残す。
    const string JsReadCode = @"
function readComponent(comps, j, state) {
    var comp, name = null, type = null, count = null, cm = null, errs = [];
    try { comp = comps.Item(j); } catch (e0) { return '{""index"":' + j + ',""error"":' + q('Item ' + errText(e0)) + '}'; }
    try { name = String(comp.Name); } catch (e1) { errs.push('Name ' + errText(e1)); }
    try { type = Number(comp.Type); } catch (e2) { errs.push('Type ' + errText(e2)); }
    try { cm = comp.CodeModule; count = Number(cm.CountOfLines); } catch (e3) { errs.push('CodeModule.CountOfLines ' + errText(e3)); }
    if (state.module == null && typeof count == 'number' && count > 0) {
        state.module = cm;
        state.project = state.currentProject;
        state.component = name;
        state.countOfLines = count;
    }
    return '{""index"":' + j + ',""name"":' + str(name) + ',""type"":' + num(type) + ',""countOfLines"":' + num(count)
        + (errs.length > 0 ? ',""errors"":' + q(errs.join(';')) : '') + '}';
}
function readProject(projects, p, state) {
    var project, name = null, comps = null, total = null, items = [], j, errs = [];
    try { project = projects.Item(p); } catch (e0) { return '{""index"":' + p + ',""error"":' + q('Item ' + errText(e0)) + '}'; }
    try { name = String(project.Name); } catch (e1) { errs.push('Name ' + errText(e1)); }
    state.currentProject = name;
    try { comps = project.VBComponents; total = Number(comps.Count); } catch (e2) { errs.push('VBComponents.Count ' + errText(e2)); comps = null; }
    if (comps != null && typeof total == 'number') {
        for (j = 1; j <= total && j <= state.maxComponents; j++) { items.push(readComponent(comps, j, state)); }
    }
    return '{""index"":' + p + ',""name"":' + str(name) + ',""componentCount"":' + num(total) + ',""components"":[' + items.join(',') + ']'
        + (typeof total == 'number' && total > state.maxComponents ? ',""truncated"":true' : '')
        + (errs.length > 0 ? ',""errors"":' + q(errs.join(';')) : '') + '}';
}
function readSample(state) {
    var cm = state.module, n = state.countOfLines, k = n < 3 ? n : 3, text = null, via = 'none', attempts = [], parts = [];
    var i, c, ok = true, head, nonAscii = 0, codes = [], replacement = 0, perline = null;
    try { text = String(cm.Lines(1, k)); via = 'lines-call'; attempts.push('lines-call=ok'); } catch (e1) { attempts.push('lines-call ' + errText(e1)); }
    for (i = 1; i <= k; i++) {
        try { parts.push(String(cm.Lines(i, 1))); } catch (e2) { attempts.push('lines-perline(' + i + ') ' + errText(e2)); ok = false; break; }
    }
    if (ok) {
        attempts.push('lines-perline=ok');
        perline = parts.join('\r\n');
        if (text == null) { text = perline; via = 'lines-perline'; }
    }
    var s = '{""project"":' + str(state.project) + ',""component"":' + str(state.component) + ',""countOfLines"":' + num(n)
        + ',""requested"":' + num(k) + ',""via"":' + q(via) + ',""attempts"":' + q(attempts.join(';'));
    if (text == null) { return s + '}'; }
    for (i = 0; i < text.length; i++) {
        c = text.charCodeAt(i);
        if (c > 127) { nonAscii++; if (codes.length < 8) { codes.push('U+' + hex(c, 4)); } }
        if (c == 0xFFFD) { replacement++; }
    }
    head = text.substr(0, 40);
    return s + ',""length"":' + text.length + ',""head"":' + q(head) + ',""nonAscii"":' + nonAscii + ',""nonAsciiCodes"":' + q(codes.join(' '))
        + ',""replacementChars"":' + replacement + (perline != null && via == 'lines-call' ? ',""perlineEqual"":' + (perline == text ? 'true' : 'false') : '') + '}';
}
function CATMain() {
    var log = [], apc = null, projects = null, total = -1, p, items = [], sample = 'null';
    var state = { module: null, project: null, component: null, countOfLines: 0, currentProject: null, maxComponents: 50 };
    try {
        apc = getApc(log);
        if (apc == null) { return '{""ok"":false,""apc"":' + q(log.join(';')) + ',""error"":""apc=null""}'; }
        try { projects = apc.VBE.VBProjects; total = Number(projects.Count); } catch (e1) {
            return '{""ok"":false,""apc"":' + q(log.join(';')) + ',""error"":' + q('VBE.VBProjects.Count ' + errText(e1)) + '}';
        }
        for (p = 1; p <= total && p <= 3; p++) { items.push(readProject(projects, p, state)); }
        if (state.module != null) { sample = readSample(state); }
        return '{""ok"":true,""apc"":' + q(log.join(';')) + ',""projectCount"":' + num(total) + ',""projects"":[' + items.join(',') + '],""sample"":' + sample + '}';
    } catch (e) {
        return '{""ok"":false,""apc"":' + q(log.join(';')) + ',""error"":' + q('unexpected ' + errText(e)) + '}';
    }
}
";

    static readonly string ApcCode = string.Join(NL, new[] {
        "Function CATMain()",
        "    On Error Resume Next",
        "    Dim apc, projects, project, s, e, total, i, projectName",
        "    Set apc = Nothing",
        "    s = \"\"",
        "    Err.Clear",
        "    Set apc = CreateObject(\"MSAPC.Apc.7.1\")",
        "    e = Err.Number",
        "    Err.Clear",
        "    s = s & \"create71=\" & CStr(e)",
        "    If apc Is Nothing Then",
        "        Set apc = CreateObject(\"MSAPC.Apc.6.2\")",
        "        e = Err.Number",
        "        Err.Clear",
        "        s = s & \"|create62=\" & CStr(e)",
        "    End If",
        "    If apc Is Nothing Then",
        "        Set apc = CreateObject(\"MSAPC.Apc\")",
        "        e = Err.Number",
        "        Err.Clear",
        "        s = s & \"|create=\" & CStr(e)",
        "    End If",
        "    If apc Is Nothing Then",
        "        CATMain = s & \"|apc=nothing\"",
        "        Exit Function",
        "    End If",
        "    Set projects = Nothing",
        "    Set projects = apc.VBE.VBProjects",
        "    e = Err.Number",
        "    Err.Clear",
        "    s = s & \"|projectsErr=\" & CStr(e)",
        "    If projects Is Nothing Then",
        "        CATMain = s & \"|projects=nothing\"",
        "        Exit Function",
        "    End If",
        "    total = -1",
        "    total = projects.Count",
        "    e = Err.Number",
        "    Err.Clear",
        "    s = s & \"|countErr=\" & CStr(e) & \"|count=\" & CStr(total)",
        "    For i = 1 To total",
        "        projectName = \"\"",
        "        Set project = Nothing",
        "        Set project = projects.Item(i)",
        "        projectName = project.Name",
        "        e = Err.Number",
        "        Err.Clear",
        "        If e = 0 Then",
        "            s = s & \"|\" & projectName",
        "        Else",
        "            s = s & \"|#\" & CStr(i) & \"!err=\" & CStr(e)",
        "        End If",
        "    Next",
        "    CATMain = s",
        "End Function" });

    /// <summary>obj（式で取得）の .VBE.VBProjects.Count を報告する VBA。式は Variant 経由のレイトバインドに留める。</summary>
    static string VbeFromCode(string label, string acquire)
    {
        return string.Join(NL, new[] {
            "Function CATMain()",
            "    On Error Resume Next",
            "    Dim obj, editor, total, e, s",
            "    Set obj = Nothing",
            "    Set editor = Nothing",
            "    Err.Clear",
            "    " + acquire,
            "    e = Err.Number",
            "    Err.Clear",
            "    s = \"" + label + "Err=\" & CStr(e) & \"|type=\" & TypeName(obj)",
            "    If Not IsObject(obj) Then",
            "        CATMain = s & \"|obj=notobject\"",
            "        Exit Function",
            "    End If",
            "    If obj Is Nothing Then",
            "        CATMain = s & \"|obj=nothing\"",
            "        Exit Function",
            "    End If",
            "    Set editor = obj.VBE",
            "    e = Err.Number",
            "    Err.Clear",
            "    s = s & \"|vbeErr=\" & CStr(e)",
            "    If editor Is Nothing Then",
            "        CATMain = s & \"|vbe=nothing\"",
            "        Exit Function",
            "    End If",
            "    total = -1",
            "    total = editor.VBProjects.Count",
            "    e = Err.Number",
            "    Err.Clear",
            "    CATMain = s & \"|countErr=\" & CStr(e) & \"|count=\" & CStr(total)",
            "End Function" });
    }

    static Step[] Expand(string name)
    {
        switch (name)
        {
            case "check":
                return new[] { new Step { Name = name } };
            case "basic":
                return new[] { new Step { Name = name, Language = VbaLanguage, Expected = "ok:2",
                    Code = "Function CATMain()" + NL + "    On Error Resume Next" + NL + "    CATMain = \"ok:\" & CStr(1 + 1)" + NL + "End Function" } };
            case "sub":
                return new[] { new Step { Name = name, Language = VbaLanguage,
                    Code = "Sub CATMain()" + NL + "    On Error Resume Next" + NL + "End Sub" } };
            case "typed":
                // VBScript では構文エラーになる As 付き宣言。"vba:Long:3" なら VBA として実行されている。
                return new[] { new Step { Name = name, Language = VbaLanguage, Expected = "vba:Long:3",
                    Code = string.Join(NL, new[] {
                        "Function CATMain() As String",
                        "    On Error Resume Next",
                        "    Dim n As Long",
                        "    Dim t As String",
                        "    n = 3",
                        "    t = \"abc\"",
                        "    CATMain = \"vba:\" & TypeName(n) & \":\" & CStr(Len(t))",
                        "End Function" }) } };
            case "params":
                return new[] { new Step { Name = name, Language = VbaLanguage, Parameters = new object[] { "cat5dev-param" },
                    Expected = "param:cat5dev-param", RetryByRef = true,
                    Code = "Function CATMain(a)" + NL + "    On Error Resume Next" + NL + "    CATMain = \"param:\" & CStr(a)" + NL + "End Function" } };
            case "apc":
                return new[] { new Step { Name = name, Language = VbaLanguage, Expected = "create71=", ExpectedPrefix = true, Code = ApcCode } };
            case "apc2":
                return Concat(Expand("apc2-getobject"), Expand("apc2-catia"), Expand("apc2-application"));
            case "apc2-getobject":
                return new[] { new Step { Name = name, Language = VbaLanguage,
                    Code = VbeFromCode("getObject", "Set obj = GetObject(, \"MSAPC.Apc\")") } };
            case "apc2-catia":
                return new[] { new Step { Name = name, Language = VbaLanguage,
                    Code = VbeFromCode("catia", "Set obj = CATIA") } };
            case "apc2-application":
                return new[] { new Step { Name = name, Language = VbaLanguage,
                    Code = VbeFromCode("application", "Set obj = Application") } };
            case "jscript":
                return new[] { new Step { Name = name, Language = JScriptLanguage, Expected = "js:2",
                    Code = "function CATMain(){ return \"js:\" + (1+1); }" } };
            case "jscript-params":
                return new[] { new Step { Name = name, Language = JScriptLanguage, Parameters = new object[] { "cat5dev-param" },
                    Expected = "param:cat5dev-param", RetryByRef = true, Code = JsPrelude + JsParamsCode } };
            case "jscript-fso":
                return new[] { new Step { Name = name, Language = JScriptLanguage, Code = JsPrelude + JsFsoCode,
                    ExpectedPattern = new Regex(@"^ADODB\.Stream=ok\(object\)\|Scripting\.FileSystemObject=ok\(object\)$") } };
            case "jscript-apc":
                return new[] { new Step { Name = name, Language = JScriptLanguage, Code = JsPrelude + JsApcCode,
                    ExpectedPattern = new Regex(@"^count=\d+[|;]") } };
            case "jscript-read":
                return new[] { new Step { Name = name, Language = JScriptLanguage, Code = JsPrelude + JsReadCode,
                    ExpectedPattern = new Regex("\"via\":\"lines-(call|perline)\"") } };
            default:
                return null;
        }
    }

    static bool IsLibraryStep(string name) { return name.StartsWith("vbaproj-", StringComparison.Ordinal); }

    static Step BridgeCall(string name, string library, string function, object[] parameters)
    {
        return new Step { Name = name, Language = -1, Method = ExecuteScriptMethod, Library = library, ProgramName = BridgeModule,
            FunctionName = function, Parameters = parameters, RetryByRef = true };
    }

    static Step PingCall(string name, string library)
    {
        Step step = BridgeCall(name, library, "C5D_Ping", new object[] { PingArgument });
        step.Expected = "vba:Long:" + PingArgument;
        return step;
    }

    static string FileNameOnly(string library)
    {
        try { return Path.GetFileName(library); } catch (ArgumentException) { return library; }
    }

    /// <summary>vbaproj-* ステップ（ExecuteScript, iType = 2）。library = .catvba のパス。不明な名前は null。</summary>
    static Step[] ExpandLibrary(string name, string library, string project)
    {
        switch (name)
        {
            case "vbaproj-ping":
                return new[] { PingCall(name, library) };
            case "vbaproj-sub":
                return new[] { BridgeCall(name, library, "C5D_Noop", new object[0]) };
            case "vbaproj-projects":
                {
                    Step step = BridgeCall(name, library, "C5D_Projects", new object[0]);
                    step.ExpectedPattern = new Regex(@"^count=\d+");
                    return new[] { step };
                }
            case "vbaproj-filename":
                return Concat(ExpandLibrary("vbaproj-filename-basename", library, project), ExpandLibrary("vbaproj-filename-empty", library, project),
                    ExpandLibrary("vbaproj-filename-project", library, project));
            case "vbaproj-filename-basename":
                return new[] { PingCall(name, FileNameOnly(library)) };
            case "vbaproj-filename-empty":
                return new[] { PingCall(name, "") };
            case "vbaproj-filename-project":
                return new[] { PingCall(name, project) };
            default:
                return null;
        }
    }

    const string Base64Prefix = "b64:";
    static readonly Encoding StrictUtf8 = new UTF8Encoding(false, true);

    /// <summary>"b64:..." なら UTF-8 の base64 として復号、それ以外はそのまま。復号できなければ null。</summary>
    static string DecodeToken(string token)
    {
        if (!token.StartsWith(Base64Prefix, StringComparison.Ordinal)) { return token; }
        try { return StrictUtf8.GetString(Convert.FromBase64String(token.Substring(Base64Prefix.Length))); }
        catch (FormatException) { return null; }
        catch (ArgumentException) { return null; }
    }

    /// <summary>vbacall: ExecuteScript(library, 2, module, function, args)。期待値なし、ByVal で例外なら ByRef で再試行。</summary>
    static Step VbaCall(string label, string library, string module, string function, object[] parameters)
    {
        return new Step { Name = label ?? "vbacall", Language = -1, Method = ExecuteScriptMethod, Library = library, ProgramName = module,
            FunctionName = function, Parameters = parameters, RetryByRef = true };
    }

    static Step[] Concat(params Step[][] groups)
    {
        List<Step> all = new List<Step>();
        foreach (Step[] group in groups) { all.AddRange(group); }
        return all.ToArray();
    }

    [STAThread]
    static int Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--print-code")
        {
            // CATIA に接続せず、ステップのコードだけを出力する（テストで JScript を検査するため）
            Step[] printed = args.Length == 2 ? Expand(args[1]) : null;
            if (printed == null || printed.Length != 1 || printed[0].Code == null)
            {
                Console.Error.WriteLine("usage: eval-probe --print-code <step with code>");
                return BadArguments;
            }
            Out.Write(printed[0].Code);
            return Ok;
        }
        if (args.Length > 0 && args[0] == "--obj-selftest")
        {
            // CATIA に接続せず、ローカルに作った COM オブジェクト（例 Scripting.Dictionary）に action を実行する（action 評価器の確認用）
            return ObjSelfTest(args);
        }
        bool byRef = false, dryRun = false, noRetry = false;
        string project = "toQI", label = null;
        Step vbaCall = null;
        // (step name, library or null) in argument order; expanded after --project is known
        List<KeyValuePair<string, string>> requested = new List<KeyValuePair<string, string>>();
        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            if (arg == "--byref") { byRef = true; continue; }
            if (arg == "--dry-run") { dryRun = true; continue; }
            if (arg == "--no-retry") { noRetry = true; continue; }
            if (arg == "--label")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--label requires a name"); return BadArguments; }
                label = args[++i];
                continue;
            }
            if (arg == "vbacall")
            {
                // vbacall <library> <module> <function> [args...]: 残りの引数は全て vbacall のもの
                if (i + 3 >= args.Length)
                {
                    Console.Error.WriteLine("usage: eval-probe [--label <name>] vbacall <library> <module> <function> [args...]");
                    return BadArguments;
                }
                List<string> values = new List<string>();
                for (int j = i + 1; j < args.Length; j++)
                {
                    string value = DecodeToken(args[j]);
                    if (value == null) { Console.Error.WriteLine("invalid b64: token at argument " + (j + 1).ToString(CultureInfo.InvariantCulture)); return BadArguments; }
                    values.Add(value);
                }
                object[] parameters = new object[values.Count - 3];
                for (int j = 3; j < values.Count; j++) { parameters[j - 3] = values[j]; }
                vbaCall = VbaCall(label, values[0], values[1], values[2], parameters);
                break;
            }
            if (arg == "vbacall-obj")
            {
                // vbacall-obj <library> <module> <function> [<action> ...]: 残りの引数は全て vbacall-obj のもの
                if (i + 3 >= args.Length)
                {
                    Console.Error.WriteLine("usage: eval-probe [--label <name>] vbacall-obj <library> <module> <function> [<action> ...]");
                    return BadArguments;
                }
                List<string> values = new List<string>();
                for (int j = i + 1; j < args.Length; j++)
                {
                    string value = DecodeToken(args[j]);
                    if (value == null) { Console.Error.WriteLine("invalid b64: token at argument " + (j + 1).ToString(CultureInfo.InvariantCulture)); return BadArguments; }
                    values.Add(value);
                }
                List<ObjAction> actions = new List<ObjAction>();
                string[] texts = new string[values.Count - 3];
                for (int j = 3; j < values.Count; j++)
                {
                    texts[j - 3] = values[j];
                    string error;
                    ObjAction action = ObjActions.Parse(values[j], out error);
                    if (action == null) { Console.Error.WriteLine("invalid action " + (j - 2).ToString(CultureInfo.InvariantCulture) + " '" + values[j] + "': " + error); return BadArguments; }
                    actions.Add(action);
                }
                vbaCall = VbaCall(label ?? "vbacall-obj", values[0], values[1], values[2], new object[0]);
                vbaCall.RetryByRef = false;
                vbaCall.Actions = actions;
                vbaCall.ActionTexts = texts;
                break;
            }
            if (arg == "--project")
            {
                if (i + 1 >= args.Length) { Console.Error.WriteLine("--project requires a project name"); return BadArguments; }
                project = args[++i];
                continue;
            }
            if (IsLibraryStep(arg))
            {
                if (i + 1 >= args.Length || args[i + 1].StartsWith("--", StringComparison.Ordinal))
                {
                    Console.Error.WriteLine(arg + " requires a library path: eval-probe " + arg + " <libraryPath>");
                    return BadArguments;
                }
                requested.Add(new KeyValuePair<string, string>(arg, args[++i]));
                continue;
            }
            requested.Add(new KeyValuePair<string, string>(arg, null));
        }
        List<Step> steps = new List<Step>();
        foreach (KeyValuePair<string, string> item in requested)
        {
            Step[] expanded = item.Value == null ? Expand(item.Key) : ExpandLibrary(item.Key, item.Value, project);
            if (expanded == null)
            {
                Console.Error.WriteLine("unknown step: " + item.Key);
                Console.Error.WriteLine("usage: eval-probe [--byref] [--project <name>] [--dry-run] <check|basic|sub|typed|params|apc|apc2|apc2-getobject|apc2-catia|apc2-application|jscript|jscript-params|jscript-fso|jscript-apc|jscript-read> ...");
                Console.Error.WriteLine("       eval-probe <vbaproj-ping|vbaproj-sub|vbaproj-projects|vbaproj-filename|vbaproj-filename-basename|vbaproj-filename-empty|vbaproj-filename-project> <libraryPath> ...");
                Console.Error.WriteLine("       eval-probe [--label <name>] vbacall <library> <module> <function> [args...]   (tokens may be b64:<UTF-8 base64>)");
                Console.Error.WriteLine("       eval-probe [--label <name>] vbacall-obj <library> <module> <function> [get:<expr>|type:<expr>|call:<expr>|let:<expr>=<arg>|sleep:<ms> ...]");
                return BadArguments;
            }
            steps.AddRange(expanded);
        }
        if (vbaCall != null)
        {
            if (noRetry) { vbaCall.RetryByRef = false; }
            steps.Add(vbaCall);
        }
        if (steps.Count == 0)
        {
            Console.Error.WriteLine("usage: eval-probe [--byref] [--project <name>] [--dry-run] <step> [<step> ...]");
            return BadArguments;
        }
        if (dryRun)
        {
            // CATIA に接続せず、呼び出し内容だけを出す
            foreach (Step step in steps)
            {
                Out.WriteLine("{\"step\":" + Str(step.Name) + ",\"dryRun\":true,\"language\":" + step.Language.ToString(CultureInfo.InvariantCulture)
                    + StepFields(step) + ",\"parameters\":" + Str(Describe(step.Parameters)) + ",\"expected\":" + Str(step.Expected)
                    + ",\"retryByRef\":" + (step.RetryByRef ? "true" : "false") + ActionFields(step) + "}");
            }
            return Ok;
        }

        object app = null, service = null;
        try
        {
            Stopwatch connect = Stopwatch.StartNew();
            try { app = Marshal.GetActiveObject("CATIA.Application"); }
            catch (Exception e) { WriteFailure(steps[0].Name, e, connect, "GetActiveObject(CATIA.Application)"); return NoCatia; }
            if (app == null) { WriteFailure(steps[0].Name, null, connect, "GetActiveObject returned null"); return NoCatia; }
            try { service = app.GetType().InvokeMember("SystemService", BindingFlags.GetProperty, null, app, null); }
            catch (Exception e) { WriteFailure(steps[0].Name, e, connect, "CATIA.SystemService"); return NoSystemService; }
            if (service == null) { WriteFailure(steps[0].Name, null, connect, "SystemService returned null"); return NoSystemService; }

            bool allOk = true;
            foreach (Step step in steps)
            {
                if (step.IsCheck)
                {
                    // check: 接続と SystemService 取得のみ
                    WriteJson(step.Name, -1, true, null, "connected", "String", 0, null, null, new List<string>(), connect.ElapsedMilliseconds, StepFields(step));
                    continue;
                }
                if (step.Actions != null) { if (!RunObjStep(service, step)) { allOk = false; } continue; }
                if (!RunStep(service, step, byRef)) { allOk = false; }
            }
            return allOk ? Ok : StepFailed;
        }
        finally
        {
            if (service != null && Marshal.IsComObject(service)) { Marshal.ReleaseComObject(service); }
            if (app != null && Marshal.IsComObject(app)) { Marshal.ReleaseComObject(app); }
        }
    }

    /// <summary>Evaluate（または ExecuteScript）を呼び、結果を 1 行 JSON で出す。RetryByRef なら ByVal 失敗/不一致時に ByRef で再試行。</summary>
    static bool RunStep(object service, Step step, bool forceByRef)
    {
        Stopwatch watch = Stopwatch.StartNew();
        List<string> attempts = new List<string>();
        bool[] modes = forceByRef ? new[] { true } : step.RetryByRef ? new[] { false, true } : new[] { false };
        bool ok = false;
        bool? match = null;
        object result = null;
        int hresult = 0;
        string message = null, paramMode = null;
        foreach (bool byRef in modes)
        {
            string mode = byRef ? "byref" : "byval";
            // iParameters は試行ごとに複製する（ByRef 呼び出しで書き換えられても次の試行に影響しないように）
            object[] callParameters = (object[])step.Parameters.Clone();
            object[] parameters = step.Method == ExecuteScriptMethod
                ? new object[] { step.Library, LibraryTypeVbaProject, step.ProgramName, step.FunctionName, callParameters }
                : new object[] { step.Code, step.Language, "CATMain", callParameters };
            try
            {
                if (byRef)
                {
                    // 最後の引数（iParameters）を VT_BYREF で渡す
                    ParameterModifier modifier = new ParameterModifier(parameters.Length);
                    modifier[parameters.Length - 1] = true;
                    result = service.GetType().InvokeMember(step.Method, BindingFlags.InvokeMethod, null, service, parameters,
                        new ParameterModifier[] { modifier }, null, null);
                }
                else
                {
                    result = service.GetType().InvokeMember(step.Method, BindingFlags.InvokeMethod, null, service, parameters);
                }
                ok = true;
                hresult = 0;
                message = null;
                paramMode = mode;
                match = Matches(step, result);
                attempts.Add(mode + ":ok" + (match == false ? " (unexpected result: " + Describe(result) + ")" : ""));
                if (match != false) { break; }
            }
            catch (Exception e)
            {
                e = Unwrap(e);
                ok = false;
                match = step.HasExpectation ? false : (bool?)null;
                result = null;
                hresult = e.HResult;
                message = e.Message;
                paramMode = mode;
                attempts.Add(mode + ":0x" + e.HResult.ToString("X8") + " " + e.Message);
            }
        }
        WriteJson(step.Name, step.Language, ok, match, Describe(result), TypeOf(result), hresult, message, paramMode, attempts, watch.ElapsedMilliseconds, StepFields(step));
        if (result != null && Marshal.IsComObject(result)) { Marshal.ReleaseComObject(result); }
        return ok && match != false;
    }

    /// <summary>
    /// vbacall-obj: ExecuteScript を 1 回だけ呼び（ByRef 再試行なし）、戻り値を保持したまま action を順に実行する。
    /// 1 行目 = ExecuteScript の結果、以降 action ごとに 1 行。action の失敗は後続を止めない。COM 参照は最後にまとめて解放する。
    /// </summary>
    static bool RunObjStep(object service, Step step)
    {
        Stopwatch watch = Stopwatch.StartNew();
        List<string> attempts = new List<string>();
        List<object> held = new List<object>();
        object result = null;
        bool ok = false;
        int hresult = 0;
        string message = null;
        object[] parameters = new object[] { step.Library, LibraryTypeVbaProject, step.ProgramName, step.FunctionName, new object[0] };
        try
        {
            result = service.GetType().InvokeMember(step.Method, BindingFlags.InvokeMethod, null, service, parameters);
            ok = true;
            attempts.Add("byval:ok");
        }
        catch (Exception e)
        {
            e = Unwrap(e);
            hresult = e.HResult;
            message = e.Message;
            attempts.Add("byval:0x" + e.HResult.ToString("X8") + " " + e.Message);
        }
        bool allOk = ok;
        try
        {
            WriteJson(step.Name, step.Language, ok, null, Describe(result), TypeOf(result), hresult, message, "byval", attempts, watch.ElapsedMilliseconds,
                StepFields(step) + ActionFields(step) + ResultFields(result));
            for (int i = 0; i < step.Actions.Count; i++)
            {
                if (!ok) { WriteAction(step.Name, i, step.Actions[i], false, null, null, 0, "not run: ExecuteScript failed", null, 0, true); allOk = false; continue; }
                if (!RunAction(step.Name, i, step.Actions[i], result, held)) { allOk = false; }
            }
        }
        finally
        {
            ReleaseAll(held, result);
        }
        return allOk;
    }

    /// <summary>--obj-selftest <progid> [<action> ...]: CreateObject(progid) を戻り値の代わりにして action を実行する。CATIA には触れない。</summary>
    static int ObjSelfTest(string[] args)
    {
        if (args.Length < 2) { Console.Error.WriteLine("usage: eval-probe --obj-selftest <progid> [<action> ...]"); return BadArguments; }
        List<ObjAction> actions = new List<ObjAction>();
        for (int j = 2; j < args.Length; j++)
        {
            string error, text = DecodeToken(args[j]);
            ObjAction action = text == null ? null : ObjActions.Parse(text, out error);
            if (action == null) { Console.Error.WriteLine("invalid action " + (j - 1).ToString(CultureInfo.InvariantCulture) + ": " + args[j]); return BadArguments; }
            actions.Add(action);
        }
        object root = null;
        List<object> held = new List<object>();
        try
        {
            Type type = Type.GetTypeFromProgID(args[1], false);
            if (type == null) { Console.Error.WriteLine("unknown ProgID " + args[1]); return StepFailed; }
            root = Activator.CreateInstance(type);
            Out.WriteLine("{\"step\":\"obj-selftest\",\"progid\":" + Str(args[1]) + ",\"resultType\":" + Str(TypeOf(root)) + ResultFields(root) + "}");
            bool allOk = true;
            for (int i = 0; i < actions.Count; i++) { if (!RunAction("obj-selftest", i, actions[i], root, held)) { allOk = false; } }
            return allOk ? Ok : StepFailed;
        }
        finally { ReleaseAll(held, root); }
    }

    static bool RunAction(string stepName, int index, ObjAction action, object root, List<object> held)
    {
        Stopwatch watch = Stopwatch.StartNew();
        object value = null;
        bool ok = false;
        int hresult = 0;
        string message = null, exception = null;
        try
        {
            switch (action.Kind)
            {
                case "sleep":
                    System.Threading.Thread.Sleep(action.Ms);
                    break;
                case "type":
                    {
                        object target = ObjActions.Eval(action.Target, root, held);
                        value = target != null && Marshal.IsComObject(target) ? (TypeInfoName(target) ?? "(no type info)") : TypeOf(target);
                        break;
                    }
                case "let":
                    ObjActions.Let(action.Target, action.Value, root, held);
                    break;
                default: // get / call
                    value = ObjActions.Eval(action.Target, root, held);
                    break;
            }
            ok = true;
        }
        catch (Exception e)
        {
            e = Unwrap(e);
            hresult = e.HResult;
            message = e.Message;
            exception = e.GetType().Name;
        }
        WriteAction(stepName, index, action, ok, value, value != null && Marshal.IsComObject(value) ? TypeInfoName(value) : null, hresult, message, exception,
            watch.ElapsedMilliseconds, false);
        return ok;
    }

    static void WriteAction(string stepName, int index, ObjAction action, bool ok, object value, string typeInfo, int hresult, string message,
        string exception, long elapsedMs, bool skipped)
    {
        StringBuilder b = new StringBuilder("{");
        b.Append("\"step\":").Append(Str(stepName));
        b.Append(",\"action\":").Append(index.ToString(CultureInfo.InvariantCulture));
        b.Append(",\"op\":").Append(Str(action.Text));
        b.Append(",\"kind\":").Append(Str(action.Kind));
        b.Append(",\"ok\":").Append(ok ? "true" : "false");
        b.Append(",\"skipped\":").Append(skipped ? "true" : "false");
        b.Append(",\"result\":").Append(Str(Describe(value)));
        b.Append(",\"resultType\":").Append(Str(TypeOf(value)));
        b.Append(",\"typeInfo\":").Append(Str(typeInfo));
        b.Append(",\"hresult\":").Append(Str("0x" + hresult.ToString("X8")));
        b.Append(",\"message\":").Append(Str(message));
        b.Append(",\"exception\":").Append(Str(exception));
        b.Append(",\"elapsedMs\":").Append(elapsedMs.ToString(CultureInfo.InvariantCulture));
        b.Append('}');
        Out.WriteLine(b.ToString());
    }

    /// <summary>vbacall-obj の JSON 追加フィールド: actions（受け取った文字列）と parsed（解釈し直した形。--dry-run での確認用）。</summary>
    static string ActionFields(Step step)
    {
        if (step.Actions == null) { return ""; }
        StringBuilder b = new StringBuilder(",\"actions\":[");
        for (int i = 0; i < step.ActionTexts.Length; i++) { if (i > 0) { b.Append(','); } b.Append(Str(step.ActionTexts[i])); }
        b.Append("],\"parsed\":[");
        for (int i = 0; i < step.Actions.Count; i++) { if (i > 0) { b.Append(','); } b.Append(Str(ObjActions.Format(step.Actions[i]))); }
        return b.Append(']').ToString();
    }

    /// <summary>ExecuteScript の戻り値の詳細: isDispatch（IDispatch を QI できたか）、typeInfo、配列なら要素数と各要素の型。</summary>
    static string ResultFields(object result)
    {
        StringBuilder b = new StringBuilder();
        bool com = result != null && Marshal.IsComObject(result);
        b.Append(",\"isDispatch\":").Append(com ? (IsDispatch(result) ? "true" : "false") : "null");
        b.Append(",\"typeInfo\":").Append(Str(com ? TypeInfoName(result) : null));
        Array array = result as Array;
        if (array != null)
        {
            b.Append(",\"arrayLength\":").Append(array.Length.ToString(CultureInfo.InvariantCulture)).Append(",\"elements\":[");
            int k = 0;
            foreach (object item in array)
            {
                if (k++ > 0) { b.Append(','); }
                bool itemCom = item != null && Marshal.IsComObject(item);
                b.Append(Str(TypeOf(item) + (itemCom ? ":" + (TypeInfoName(item) ?? "?") : "")));
            }
            b.Append(']');
        }
        return b.ToString();
    }

    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IDispatchInfo
    {
        [PreserveSig] int GetTypeInfoCount(out uint count);
        [PreserveSig] int GetTypeInfo(uint index, int lcid, out System.Runtime.InteropServices.ComTypes.ITypeInfo info);
    }

    static bool IsDispatch(object value)
    {
        try { IntPtr p = Marshal.GetIDispatchForObject(value); Marshal.Release(p); return true; }
        catch (Exception) { return false; }
    }

    /// <summary>IDispatch.GetTypeInfo(0) の型名（例 _VBProject / IApcProject）。取れなければ "(...)" の説明、IDispatch でなければ null。</summary>
    static string TypeInfoName(object value)
    {
        try
        {
            IDispatchInfo dispatch = value as IDispatchInfo;
            if (dispatch == null) { return null; }
            uint count;
            int hr = dispatch.GetTypeInfoCount(out count);
            if (hr != 0 || count == 0) { return "(no type info: 0x" + hr.ToString("X8") + ")"; }
            System.Runtime.InteropServices.ComTypes.ITypeInfo info;
            hr = dispatch.GetTypeInfo(0, 0, out info);
            if (hr != 0 || info == null) { return "(GetTypeInfo 0x" + hr.ToString("X8") + ")"; }
            string name, doc, helpFile;
            int helpContext;
            info.GetDocumentation(-1, out name, out doc, out helpContext, out helpFile);
            Marshal.ReleaseComObject(info);
            return name;
        }
        catch (Exception e) { return "(type info error 0x" + e.HResult.ToString("X8") + ")"; }
    }

    static void ReleaseAll(List<object> held, object result)
    {
        List<object> all = new List<object>(held);
        Array array = result as Array;
        if (array != null) { foreach (object item in array) { all.Add(item); } }
        all.Add(result);
        List<object> done = new List<object>();
        foreach (object item in all)
        {
            if (item == null || !Marshal.IsComObject(item)) { continue; }
            bool seen = false;
            foreach (object d in done) { if (ReferenceEquals(d, item)) { seen = true; break; } }
            if (seen) { continue; }
            done.Add(item);
            try { Marshal.ReleaseComObject(item); } catch (Exception) { }
        }
    }

    static bool? Matches(Step step, object result)
    {
        if (!step.HasExpectation) { return null; }
        string text = result as string;
        if (text == null) { return false; }
        if (step.ExpectedPattern != null) { return step.ExpectedPattern.IsMatch(text); }
        return step.ExpectedPrefix ? text.StartsWith(step.Expected, StringComparison.Ordinal) : text == step.Expected;
    }

    static string Describe(object value)
    {
        if (value == null) { return null; }
        if (value is DBNull) { return "(Null)"; }
        if (Marshal.IsComObject(value)) { return "(COM object)"; }
        Array array = value as Array;
        if (array != null && !(value is string))
        {
            List<string> items = new List<string>();
            foreach (object item in array) { items.Add(Describe(item) ?? "(Empty)"); }
            return "[" + string.Join(", ", items.ToArray()) + "]";
        }
        IFormattable formattable = value as IFormattable;
        return formattable != null ? formattable.ToString(null, CultureInfo.InvariantCulture) : value.ToString();
    }

    static string TypeOf(object value)
    {
        if (value == null) { return "Empty"; }
        if (Marshal.IsComObject(value)) { return "ComObject"; }
        return value.GetType().Name;
    }

    static Exception Unwrap(Exception e)
    {
        while (e is TargetInvocationException && e.InnerException != null) { e = e.InnerException; }
        return e;
    }

    static void WriteFailure(string step, Exception e, Stopwatch watch, string context)
    {
        Exception inner = e == null ? null : Unwrap(e);
        string message = context + (inner == null ? "" : ": " + inner.Message);
        WriteJson(step, -1, false, null, null, null, inner == null ? 0 : inner.HResult, message, null, new List<string>(), watch.ElapsedMilliseconds, null);
    }

    /// <summary>JSON に追加するフィールド（先頭のカンマ付き）: method と、ExecuteScript なら library/libraryType/program/function。</summary>
    static string StepFields(Step step)
    {
        StringBuilder b = new StringBuilder();
        b.Append(",\"method\":").Append(Str(step.Method));
        if (step.Method == ExecuteScriptMethod)
        {
            b.Append(",\"library\":").Append(Str(step.Library));
            b.Append(",\"libraryType\":").Append(LibraryTypeVbaProject.ToString(CultureInfo.InvariantCulture));
            b.Append(",\"program\":").Append(Str(step.ProgramName));
            b.Append(",\"function\":").Append(Str(step.FunctionName));
            // 実際に渡した iParameters（文字列はそのまま、他は Describe）。node→exe の受け渡しで化けていないかの確認用
            b.Append(",\"arguments\":[");
            for (int i = 0; i < step.Parameters.Length; i++)
            {
                if (i > 0) { b.Append(','); }
                object value = step.Parameters[i];
                b.Append(Str(value as string ?? Describe(value)));
            }
            b.Append(']');
        }
        return b.ToString();
    }

    static void WriteJson(string step, int language, bool ok, bool? match, string result, string resultType, int hresult,
        string message, string paramMode, List<string> attempts, long elapsedMs, string extraFields)
    {
        StringBuilder b = new StringBuilder("{");
        b.Append("\"step\":").Append(Str(step));
        b.Append(",\"language\":").Append(language.ToString(CultureInfo.InvariantCulture));
        b.Append(",\"ok\":").Append(ok ? "true" : "false");
        b.Append(",\"match\":").Append(match == null ? "null" : match.Value ? "true" : "false");
        b.Append(",\"result\":").Append(Str(result));
        b.Append(",\"resultType\":").Append(Str(resultType));
        b.Append(",\"hresult\":").Append(Str("0x" + hresult.ToString("X8")));
        b.Append(",\"message\":").Append(Str(message));
        b.Append(",\"paramMode\":").Append(Str(paramMode));
        b.Append(",\"attempts\":[");
        for (int i = 0; i < attempts.Count; i++) { if (i > 0) { b.Append(','); } b.Append(Str(attempts[i])); }
        b.Append("],\"elapsedMs\":").Append(elapsedMs.ToString(CultureInfo.InvariantCulture));
        if (extraFields != null) { b.Append(extraFields); }
        b.Append('}');
        Out.WriteLine(b.ToString());
    }

    static string Str(string value)
    {
        if (value == null) { return "null"; }
        StringBuilder b = new StringBuilder("\"");
        foreach (char c in value)
        {
            switch (c)
            {
                case '"': b.Append("\\\""); break;
                case '\\': b.Append("\\\\"); break;
                case '\n': b.Append("\\n"); break;
                case '\r': b.Append("\\r"); break;
                case '\t': b.Append("\\t"); break;
                default:
                    if (c < 0x20) { b.Append("\\u").Append(((int)c).ToString("x4")); } else { b.Append(c); }
                    break;
            }
        }
        return b.Append('"').ToString();
    }
}

// ---- vbacall-obj の action（get / type / call / let / sleep）の構文木・構文解析・評価（IDispatch レイトバインド）
internal sealed class ObjExpr
{
    public int Root = -1;                                   // -1 = $（戻り値そのもの）, n = $n（戻り値の配列の n 番目）
    public List<ObjSegment> Segments = new List<ObjSegment>();
}

internal sealed class ObjSegment
{
    public string Name;
    public List<ObjArg> Args;                               // null = 括弧なし（プロパティ取得）
}

internal sealed class ObjArg
{
    public ObjExpr Ref;                                     // $... の式
    public object Value;                                    // bool / int / string
    public bool Missing;                                    // 省略引数（Type.Missing）
}

internal sealed class ObjAction
{
    public string Kind;                                     // get / type / call / let / sleep
    public string Text;                                     // 受け取った文字列
    public ObjExpr Target;
    public ObjArg Value;                                    // let の右辺
    public int Ms;                                          // sleep
}

internal static class ObjActions
{
    const char Quote = (char)39; // '

    /// <summary>1 つの action を解釈する。不正なら null と error。</summary>
    public static ObjAction Parse(string text, out string error)
    {
        error = null;
        int colon = text.IndexOf(':');
        if (colon <= 0) { error = "expected <get|type|call|let|sleep>:..."; return null; }
        string kind = text.Substring(0, colon), rest = text.Substring(colon + 1);
        ObjAction action = new ObjAction { Kind = kind, Text = text };
        int pos = 0;
        try
        {
            switch (kind)
            {
                case "sleep":
                    {
                        int ms;
                        if (!int.TryParse(rest, NumberStyles.None, CultureInfo.InvariantCulture, out ms) || ms > 600000) { throw new FormatException("sleep needs 0..600000 (ms)"); }
                        action.Ms = ms;
                        return action;
                    }
                case "get":
                case "type":
                case "call":
                    action.Target = ParseExpr(rest, ref pos);
                    ExpectEnd(rest, ref pos);
                    if (kind == "call" && (action.Target.Segments.Count == 0 || action.Target.Segments[action.Target.Segments.Count - 1].Args == null))
                    {
                        throw new FormatException("call needs a method call at the end (e.g. call:$.Close(false))");
                    }
                    return action;
                case "let":
                    action.Target = ParseExpr(rest, ref pos);
                    if (action.Target.Segments.Count == 0 || action.Target.Segments[action.Target.Segments.Count - 1].Args != null)
                    {
                        throw new FormatException("let needs a property name at the end (e.g. let:$.Mode=0)");
                    }
                    SkipSpaces(rest, ref pos);
                    if (pos >= rest.Length || rest[pos] != '=') { throw new FormatException("expected '='"); }
                    pos++;
                    action.Value = ParseArg(rest, ref pos);
                    ExpectEnd(rest, ref pos);
                    return action;
                default:
                    error = "unknown action kind '" + kind + "'";
                    return null;
            }
        }
        catch (FormatException e)
        {
            error = e.Message + " (at column " + (colon + 2 + pos).ToString(CultureInfo.InvariantCulture) + ")";
            return null;
        }
    }

    static void SkipSpaces(string s, ref int pos) { while (pos < s.Length && s[pos] == ' ') { pos++; } }

    static void ExpectEnd(string s, ref int pos)
    {
        SkipSpaces(s, ref pos);
        if (pos < s.Length) { throw new FormatException("unexpected '" + s.Substring(pos) + "'"); }
    }

    static string ParseName(string s, ref int pos)
    {
        int start = pos;
        if (pos < s.Length && (char.IsLetter(s[pos]) || s[pos] == '_')) { pos++; } else { throw new FormatException("expected a name"); }
        while (pos < s.Length && (char.IsLetterOrDigit(s[pos]) || s[pos] == '_')) { pos++; }
        return s.Substring(start, pos - start);
    }

    static ObjExpr ParseExpr(string s, ref int pos)
    {
        SkipSpaces(s, ref pos);
        if (pos >= s.Length || s[pos] != '$') { throw new FormatException("expected '$'"); }
        pos++;
        ObjExpr expr = new ObjExpr();
        int start = pos;
        while (pos < s.Length && char.IsDigit(s[pos])) { pos++; }
        if (pos > start) { expr.Root = int.Parse(s.Substring(start, pos - start), CultureInfo.InvariantCulture); }
        while (pos < s.Length && s[pos] == '.')
        {
            pos++;
            ObjSegment segment = new ObjSegment { Name = ParseName(s, ref pos) };
            if (pos < s.Length && s[pos] == '(') { segment.Args = ParseArgs(s, ref pos); }
            expr.Segments.Add(segment);
        }
        return expr;
    }

    static List<ObjArg> ParseArgs(string s, ref int pos)
    {
        pos++; // (
        List<ObjArg> args = new List<ObjArg>();
        SkipSpaces(s, ref pos);
        if (pos < s.Length && s[pos] == ')') { pos++; return args; }
        while (true)
        {
            args.Add(ParseArg(s, ref pos));
            SkipSpaces(s, ref pos);
            if (pos < s.Length && s[pos] == ',') { pos++; continue; }
            if (pos < s.Length && s[pos] == ')') { pos++; return args; }
            throw new FormatException("expected ',' or ')'");
        }
    }

    static ObjArg ParseArg(string s, ref int pos)
    {
        SkipSpaces(s, ref pos);
        if (pos >= s.Length) { throw new FormatException("expected an argument"); }
        char c = s[pos];
        if (c == '$') { return new ObjArg { Ref = ParseExpr(s, ref pos) }; }
        if (c == Quote)
        {
            StringBuilder b = new StringBuilder();
            pos++;
            while (true)
            {
                if (pos >= s.Length) { throw new FormatException("unterminated string"); }
                if (s[pos] == Quote)
                {
                    if (pos + 1 < s.Length && s[pos + 1] == Quote) { b.Append(Quote); pos += 2; continue; }
                    pos++;
                    return new ObjArg { Value = b.ToString() };
                }
                b.Append(s[pos++]);
            }
        }
        if (c == '-' || char.IsDigit(c))
        {
            int start = pos;
            pos++;
            while (pos < s.Length && char.IsDigit(s[pos])) { pos++; }
            int value;
            if (!int.TryParse(s.Substring(start, pos - start), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value)) { throw new FormatException("bad integer"); }
            return new ObjArg { Value = value };
        }
        string word = ParseName(s, ref pos);
        if (word == "true") { return new ObjArg { Value = true }; }
        if (word == "false") { return new ObjArg { Value = false }; }
        if (word == "missing") { return new ObjArg { Missing = true }; }
        throw new FormatException("unknown argument '" + word + "'");
    }

    /// <summary>解釈した action を正規形の文字列に戻す（--dry-run での確認用）。</summary>
    public static string Format(ObjAction action)
    {
        switch (action.Kind)
        {
            case "sleep": return "sleep:" + action.Ms.ToString(CultureInfo.InvariantCulture);
            case "let": return "let:" + FormatExpr(action.Target) + "=" + FormatArg(action.Value);
            default: return action.Kind + ":" + FormatExpr(action.Target);
        }
    }

    static string FormatExpr(ObjExpr expr)
    {
        StringBuilder b = new StringBuilder("$");
        if (expr.Root >= 0) { b.Append(expr.Root.ToString(CultureInfo.InvariantCulture)); }
        foreach (ObjSegment segment in expr.Segments)
        {
            b.Append('.').Append(segment.Name);
            if (segment.Args == null) { continue; }
            b.Append('(');
            for (int i = 0; i < segment.Args.Count; i++) { if (i > 0) { b.Append(", "); } b.Append(FormatArg(segment.Args[i])); }
            b.Append(')');
        }
        return b.ToString();
    }

    static string FormatArg(ObjArg arg)
    {
        if (arg.Ref != null) { return FormatExpr(arg.Ref); }
        if (arg.Missing) { return "missing"; }
        if (arg.Value is bool) { return (bool)arg.Value ? "true" : "false"; }
        if (arg.Value is int) { return ((int)arg.Value).ToString(CultureInfo.InvariantCulture); }
        string q = Quote.ToString();
        return q + ((string)arg.Value).Replace(q, q + q) + q;
    }

    /// <summary>式を評価する。途中で得た COM オブジェクトは held に積む（呼び出し側が最後に解放する）。</summary>
    public static object Eval(ObjExpr expr, object root, List<object> held)
    {
        object current = RootOf(expr.Root, root);
        foreach (ObjSegment segment in expr.Segments) { current = Invoke(current, segment, root, held); }
        return current;
    }

    /// <summary>let: 最後の区間のプロパティに値を入れる（DISPATCH_PROPERTYPUT）。</summary>
    public static void Let(ObjExpr expr, ObjArg value, object root, List<object> held)
    {
        object owner = RootOf(expr.Root, root);
        for (int i = 0; i < expr.Segments.Count - 1; i++) { owner = Invoke(owner, expr.Segments[i], root, held); }
        string name = expr.Segments[expr.Segments.Count - 1].Name;
        if (owner == null) { throw new InvalidOperationException("." + name + " = ... on Empty/Nothing"); }
        owner.GetType().InvokeMember(name, BindingFlags.SetProperty, null, owner, new object[] { ArgValue(value, root, held) }, CultureInfo.InvariantCulture);
    }

    static object RootOf(int index, object root)
    {
        if (index < 0) { return root; }
        Array array = root as Array;
        if (array == null) { throw new InvalidOperationException("$" + index.ToString(CultureInfo.InvariantCulture) + " needs an array result (got " + (root == null ? "Empty" : root.GetType().Name) + ")"); }
        int k = 0;
        foreach (object item in array) { if (k++ == index) { return item; } }
        throw new IndexOutOfRangeException("$" + index.ToString(CultureInfo.InvariantCulture) + " is out of range (array length " + array.Length.ToString(CultureInfo.InvariantCulture) + ")");
    }

    static object Invoke(object target, ObjSegment segment, object root, List<object> held)
    {
        if (target == null) { throw new InvalidOperationException("." + segment.Name + " on Empty/Nothing"); }
        object[] args;
        BindingFlags flags;
        if (segment.Args == null) { args = new object[0]; flags = BindingFlags.GetProperty; }
        else
        {
            args = new object[segment.Args.Count];
            for (int i = 0; i < args.Length; i++) { args[i] = ArgValue(segment.Args[i], root, held); }
            flags = BindingFlags.InvokeMethod | BindingFlags.GetProperty;
        }
        object result = target.GetType().InvokeMember(segment.Name, flags, null, target, args, CultureInfo.InvariantCulture);
        if (result != null && Marshal.IsComObject(result)) { held.Add(result); }
        return result;
    }

    static object ArgValue(ObjArg arg, object root, List<object> held)
    {
        if (arg.Missing) { return Type.Missing; }
        if (arg.Ref != null) { return Eval(arg.Ref, root, held); }
        return arg.Value;
    }
}
