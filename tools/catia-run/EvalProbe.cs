// eval-probe.exe（開発者用・非出荷）: CATIA.SystemService.Evaluate で VBA コード文字列を CATIA 内実行できるかを試す。
// Build: node scripts/probe-evaluate.js が temp/eval-probe.exe にビルドする（.NET Framework 4.x の csc.exe、C# 5 の構文に留める）。
// Usage: eval-probe.exe [--byref] [--project <name>] [--dry-run] <step> [<step> ...]
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
//   --byref: iParameters を VT_BYREF で渡す（params / jscript-params は指定がなくても ByVal 失敗時に ByRef で再試行する）
//   jscript*: 言語 4（JScript）。ES3 構文のみ・全処理を try/catch で囲み、例外は結果文字列に含めて返す。書き込み系 API は呼ばない。
// Output: ステップごとに 1 行 JSON（UTF-8）
//   {"step","language","ok","match","result","resultType","hresult","message","paramMode","attempts","elapsedMs"}
//   ok = Evaluate が例外なく戻った, match = 戻り値が期待どおり（期待値のないステップは null）
// Exit codes: 0 = 全ステップ ok, 1 = いずれか失敗, 2 = 引数不正, 10 = CATIA 取得失敗, 11 = SystemService 取得失敗
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
        bool byRef = false, dryRun = false;
        string project = "toQI";
        // (step name, library or null) in argument order; expanded after --project is known
        List<KeyValuePair<string, string>> requested = new List<KeyValuePair<string, string>>();
        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            if (arg == "--byref") { byRef = true; continue; }
            if (arg == "--dry-run") { dryRun = true; continue; }
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
                return BadArguments;
            }
            steps.AddRange(expanded);
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
                    + ",\"retryByRef\":" + (step.RetryByRef ? "true" : "false") + "}");
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
