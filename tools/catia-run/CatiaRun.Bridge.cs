// catia-run.exe: ブリッジの手順（docs/dev/bridge-contract.md 5・7・8 章）。読み込み → 自分の特定 → 残骸の Close → runtime →
// VBA エディタの確認 → 要求の処理 → m<i>.txt / result.json → finally で guard 付き Close。
// ExecuteScript は IScriptHost の向こう側にあり、CATIA なしの試験では偽物に差し替えられる。
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

/// <summary>ExecuteScript(library, 2, module, function, args)。失敗は例外（TargetInvocationException は呼ぶ側で外す）。</summary>
internal interface IScriptHost
{
    object Execute(string library, string module, string function, object[] args);
}

internal sealed class CatiaScriptHost : IScriptHost
{
    const int LibraryTypeVbaProject = 2; // catScriptLibraryTypeVBAProject
    readonly object service;

    public CatiaScriptHost(object service) { this.service = service; }

    public object Execute(string library, string module, string function, object[] args)
    {
        // iParameters は必ず要素型 object の配列（string[] を渡すと VT_ARRAY|VT_BSTR になり 0x80020005）。ByRef の再試行はしない。
        object[] callArgs = new object[args.Length];
        for (int i = 0; i < args.Length; i++) { callArgs[i] = args[i]; }
        object[] parameters = { library, LibraryTypeVbaProject, module, function, callArgs };
        return service.GetType().InvokeMember("ExecuteScript", BindingFlags.InvokeMethod, null, service, parameters);
    }
}

/// <summary>埋め込みの runtime 本文（マニフェスト リソース C5D_Bridge_Runtime.bas）と、その RT_VERSION。</summary>
internal sealed class RuntimeSource
{
    public const string ResourceName = "C5D_Bridge_Runtime.bas";
    public string Text;
    public string Version;

    static readonly Regex VersionLine = new Regex("^Private Const RT_VERSION As String = \"(c5d-runtime:[0-9]+)\"\r?$",
        RegexOptions.Multiline | RegexOptions.CultureInvariant);

    /// <summary>読めなければ null（error に理由）。</summary>
    public static RuntimeSource Load(out string error)
    {
        error = null;
        try
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
            {
                if (stream == null) { error = "embedded resource " + ResourceName + " is missing"; return null; }
                MemoryStream copy = new MemoryStream();
                stream.CopyTo(copy);
                string text = new UTF8Encoding(false, true).GetString(copy.ToArray());
                Match match = VersionLine.Match(text);
                if (!match.Success) { error = "RT_VERSION is not found in the embedded runtime"; return null; }
                RuntimeSource source = new RuntimeSource();
                source.Text = text;
                source.Version = match.Groups[1].Value;
                return source;
            }
        }
        catch (Exception e) { error = "cannot read the embedded runtime: " + e.Message; return null; }
    }
}

internal sealed class Failure
{
    public string Phase;
    public string Target;
    public int Code;
    public string Message;
}

internal sealed class ResultItem
{
    public string Name;
    public int Type;
    public string File;
}

/// <summary>result.json の中身（version 1。フィールドの順序は operation.catvbs と同じ）。</summary>
internal sealed class ResultData
{
    public int ExpectedCount;
    public int SucceededCount;
    public int UpdatedCount;
    public int DeletedCount;
    public List<Failure> Failures = new List<Failure>();
    public List<ResultItem> Items = new List<ResultItem>();
    public List<string> Projects = new List<string>();

    public string ToJson(BridgeRequest request)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("{\"version\":1");
        sb.Append(",\"operationId\":").Append(JsonText.Quote(request.OperationId));
        sb.Append(",\"stepId\":").Append(JsonText.Quote(request.StepId));
        sb.Append(",\"kind\":").Append(JsonText.Quote(request.Kind));
        sb.Append(",\"targetProject\":").Append(JsonText.Quote(request.TargetProject));
        sb.Append(",\"expectedCount\":").Append(ExpectedCount.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"succeededCount\":").Append(SucceededCount.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"failures\":[");
        for (int i = 0; i < Failures.Count; i++)
        {
            Failure f = Failures[i];
            if (i > 0) { sb.Append(','); }
            sb.Append("{\"phase\":").Append(JsonText.Quote(f.Phase))
              .Append(",\"target\":").Append(JsonText.Quote(f.Target))
              .Append(",\"code\":").Append(f.Code.ToString(CultureInfo.InvariantCulture))
              .Append(",\"message\":").Append(JsonText.Quote(f.Message)).Append('}');
        }
        sb.Append("],\"items\":[");
        for (int i = 0; i < Items.Count; i++)
        {
            ResultItem item = Items[i];
            if (i > 0) { sb.Append(','); }
            sb.Append("{\"name\":").Append(JsonText.Quote(item.Name))
              .Append(",\"type\":").Append(item.Type.ToString(CultureInfo.InvariantCulture))
              .Append(",\"file\":").Append(JsonText.Quote(item.File)).Append('}');
        }
        sb.Append("],\"projects\":[");
        for (int i = 0; i < Projects.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            sb.Append(JsonText.Quote(Projects[i]));
        }
        sb.Append("],\"updatedCount\":").Append(UpdatedCount.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"deletedCount\":").Append(DeletedCount.ToString(CultureInfo.InvariantCulture));
        sb.Append('}');
        return sb.ToString();
    }
}

/// <summary>runtime の 1 回の呼び出しの結果: 成功（Text）、error:（Failure の材料）、異常（例外・文字列でない・形式外）。</summary>
internal sealed class RuntimeReply
{
    public bool Ok;
    public string Text;          // Ok: 戻り値の文字列
    public bool IsError;         // error:<phase>:<n>:<description>
    public string Phase;
    public int Code;
    public string Message;       // IsError: description / 異常: 例外のメッセージなど
    // Ok でも IsError でもなければ異常（Code = HResult または 13）
}

internal static class BridgeText
{
    public const string ProjectName = "C5D_Bridge";
    public const string PushModule = "C5D_Bridge_Push";
    public const string RuntimeModule = "C5D_Bridge_Runtime";
    public const string BridgeVersion = "c5d-bridge:2";
    public const string BridgeVersionV1 = "c5d-bridge:1";

    static readonly Regex ErrorText = new Regex("^error:([a-z0-9.]+):(-?[0-9]+):(.*)$", RegexOptions.Singleline | RegexOptions.CultureInvariant);
    static readonly Regex Escape = new Regex("%(09|0A|0D|25)", RegexOptions.CultureInvariant);

    /// <summary>レコード形式（1 行目 ok、LF 区切りの行、TAB 区切りのフィールド）。形式外（フィールド数違いを含む）なら null。</summary>
    public static List<string[]> ParseRecords(string text, int fieldCount)
    {
        if (text == null) { return null; }
        string[] lines = text.Split('\n');
        if (lines[0] != "ok") { return null; }
        List<string[]> records = new List<string[]>();
        for (int i = 1; i < lines.Length; i++)
        {
            string[] fields = lines[i].Split('\t');
            if (fields.Length != fieldCount) { return null; }
            for (int j = 0; j < fields.Length; j++) { fields[j] = Unescape(fields[j]); }
            records.Add(fields);
        }
        return records;
    }

    /// <summary>%09 %0A %0D %25 を 1 回の走査で戻す。</summary>
    public static string Unescape(string field)
    {
        return Escape.Replace(field, m => ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString());
    }

    public static bool TryParseError(string text, out string phase, out int code, out string message)
    {
        phase = null; code = 0; message = null;
        if (text == null) { return false; }
        Match match = ErrorText.Match(text);
        if (!match.Success) { return false; }
        if (!int.TryParse(match.Groups[2].Value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out code)) { return false; }
        phase = match.Groups[1].Value;
        message = match.Groups[3].Value;
        return true;
    }

    public static Exception Unwrap(Exception e)
    {
        while (e is TargetInvocationException && e.InnerException != null) { e = e.InnerException; }
        return e;
    }

    /// <summary>"0xHRESULT: メッセージ"（今の catia-run.exe と同じ形）。</summary>
    public static string Describe(Exception e)
    {
        e = Unwrap(e);
        return "0x" + e.HResult.ToString("X8", CultureInfo.InvariantCulture) + ": " + e.Message;
    }

    /// <summary>VBA から戻ったオブジェクト（C5D_Self）か。文字列・Empty・Null・配列・数値などは違う。</summary>
    public static bool IsObject(object value)
    {
        if (value == null || value is string || value is DBNull || value is Array) { return false; }
        if (Marshal.IsComObject(value)) { return true; }
        return !value.GetType().IsValueType;
    }

    public static string TypeName(object value)
    {
        if (value == null) { return "Empty"; }
        if (Marshal.IsComObject(value)) { return "Object"; }
        return value.GetType().Name;
    }

    public static void Release(object value)
    {
        if (value == null || !Marshal.IsComObject(value)) { return; }
        try { Marshal.ReleaseComObject(value); } catch (Exception) { }
    }

    public static object GetProperty(object target, string name)
    {
        return target.GetType().InvokeMember(name, BindingFlags.GetProperty, null, target, new object[0], CultureInfo.InvariantCulture);
    }

    public static string Shorten(string text)
    {
        if (text == null) { return "(null)"; }
        string oneLine = text.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
        return oneLine.Length <= 200 ? oneLine : oneLine.Substring(0, 200) + "...";
    }
}

internal sealed class BridgeSession
{
    public const int Ok = 0, ExecuteFailed = 12, BridgeNotReady = 13, NoVbaEditor = 14;
    const int MaxWrecks = 5;
    const int UnexpectedReturnCode = 13;

    static readonly Regex WreckFileName = new Regex(@"^C5D_Bridge.*\.catvba$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    readonly IScriptHost host;
    readonly string link;
    readonly string masterFolder;
    readonly RuntimeSource runtime;
    readonly BridgeRequest request;
    readonly string resultPath;
    readonly string workFolder;
    readonly TextWriter log;
    readonly TextWriter err;

    bool identified;       // 手順 6 で自分を特定できた（finally の Close をしてよい）
    bool userCallsStarted; // 利用者のプロジェクトへの呼び出しを始めた（以降の異常は 13 ではなく 12）

    public BridgeSession(IScriptHost host, string link, string masterPath, RuntimeSource runtime, BridgeRequest request,
        string resultPath, TextWriter log, TextWriter err)
    {
        this.host = host;
        this.link = link;
        this.masterFolder = PathNames.Long(PathNames.Parent(masterPath));
        this.runtime = runtime;
        this.request = request;
        this.resultPath = resultPath;
        this.workFolder = Path.GetDirectoryName(resultPath);
        this.log = log;
        this.err = err;
    }

    /// <summary>手順 5〜12。終了コード（0 / 12 / 13 / 14）を返す。</summary>
    public int Run()
    {
        int code;
        try { code = RunSteps(); }
        catch (Exception e)
        {
            err.WriteLine("error: unexpected failure: " + BridgeText.Describe(e));
            code = userCallsStarted ? ExecuteFailed : BridgeNotReady;
        }
        finally
        {
            if (identified) { CloseSelf(); }
        }
        return code;
    }

    void Log(string stage, Stopwatch watch)
    {
        log.WriteLine(stage + " " + watch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "ms");
    }

    int RunSteps()
    {
        // 5. 読み込み（自動読み込み）と版の確認
        Stopwatch watch = Stopwatch.StartNew();
        string version = PrepareCall("load", link, BridgeText.PushModule, "C5D_Version");
        if (version == null) { return BridgeNotReady; }
        if (version == BridgeText.BridgeVersionV1) { err.WriteLine("bridge v1 is not supported; restart CATIA"); return BridgeNotReady; }
        if (version != BridgeText.BridgeVersion) { err.WriteLine("load: unexpected C5D_Version: " + BridgeText.Shorten(version)); return BridgeNotReady; }
        Log("load", watch);

        // 6. 自分の特定
        watch = Stopwatch.StartNew();
        string whoText = PrepareCall("who", link, BridgeText.PushModule, "C5D_Who");
        if (whoText == null) { return BridgeNotReady; }
        List<string[]> who = ParseWho(whoText);
        if (who == null) { err.WriteLine("who: unexpected C5D_Who: " + BridgeText.Shorten(whoText)); return BridgeNotReady; }
        string[] self = who[0];
        if (self[0] != BridgeText.ProjectName || !SamePath(self[1], link) || !SamePath(self[2], link))
        {
            err.WriteLine("who: the executing project is not this link (Name=" + self[0] + ", DisplayName=" + self[1] + ", FileName=" + self[2]
                + ", link=" + link + ")");
            return BridgeNotReady;
        }
        identified = true;
        bool exact = self[1] == link && self[2] == link;
        bool ignoreCase = string.Equals(self[1], link, StringComparison.OrdinalIgnoreCase) && string.Equals(self[2], link, StringComparison.OrdinalIgnoreCase);
        Log("who " + (exact ? "exact" : ignoreCase ? "ignore-case" : "long-name"), watch);

        // 7. 単一の確認と残骸の Close（これが終わるまでブリッジにも書かない）
        watch = Stopwatch.StartNew();
        int wrecks;
        if (!CloseWrecks(out wrecks)) { return BridgeNotReady; }
        Log("list " + wrecks.ToString(CultureInfo.InvariantCulture) + " wreck(s) closed", watch);

        // 8. runtime
        watch = Stopwatch.StartNew();
        string runtimeState;
        if (!EnsureRuntime(who, out runtimeState)) { return BridgeNotReady; }
        Log("runtime " + runtimeState, watch);

        // 9. VBA エディタの確認（C5D_RT_ListProjects）
        watch = Stopwatch.StartNew();
        ResultData result = new ResultData();
        RuntimeReply projects = RuntimeCall("C5D_RT_ListProjects", 1);
        if (!projects.Ok)
        {
            AddFailure(result, projects, "list.call", request.TargetProject);
            Log("projects failed", watch);
            return WriteResult(result, projects.IsError ? Ok : ExecuteFailed);
        }
        List<string[]> names = BridgeText.ParseRecords(projects.Text, 1);
        if (names.Count == 0)
        {
            // VBA エディタを開く前は C5D_Bridge しか見えないため、「未起動」と「プロジェクトが無い」は区別できない（どちらも 14）。
            err.WriteLine("C5D_RT_ListProjects returned no project: the VBA editor has not been opened, or no VBA project is loaded");
            Log("projects 0", watch);
            return NoVbaEditor;
        }
        Log("projects " + names.Count.ToString(CultureInfo.InvariantCulture), watch);

        // 10. 要求の処理
        watch = Stopwatch.StartNew();
        bool callFailed;
        if (request.Kind == "list")
        {
            foreach (string[] name in names) { result.Projects.Add(name[0]); }
            result.ExpectedCount = result.Projects.Count;
            result.SucceededCount = result.Projects.Count;
            callFailed = false;
        }
        else if (request.Kind == "push") { callFailed = Push(result); }
        else { callFailed = Export(result); }
        Log(request.Kind + " " + result.SucceededCount.ToString(CultureInfo.InvariantCulture) + "/"
            + result.ExpectedCount.ToString(CultureInfo.InvariantCulture), watch);

        // 11. result.json
        return WriteResult(result, callFailed ? ExecuteFailed : Ok);
    }

    /// <summary>大小無視、だめなら両側を長い名前（8.3 短縮名を展開）にして大小無視（PathNames.Same）。</summary>
    static bool SamePath(string a, string b) { return PathNames.Same(a, b); }

    /// <summary>C5D_Who: 1 行目 = Name / DisplayName / FileName、2 行目から = 部品名 / 種類。形式外なら null。</summary>
    static List<string[]> ParseWho(string text)
    {
        string[] lines = text.Split('\n');
        if (lines.Length < 2 || lines[0] != "ok") { return null; }
        List<string[]> rows = new List<string[]>();
        for (int i = 1; i < lines.Length; i++)
        {
            string[] fields = lines[i].Split('\t');
            int expected = i == 1 ? 3 : 2;
            if (fields.Length != expected) { return null; }
            for (int j = 0; j < fields.Length; j++) { fields[j] = BridgeText.Unescape(fields[j]); }
            int type;
            if (i > 1 && !int.TryParse(fields[1], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out type)) { return null; }
            rows.Add(fields);
        }
        return rows;
    }

    /// <summary>ブリッジの準備中の呼び出し。文字列ならそれを、例外・文字列でなければ stderr に出して null。</summary>
    string PrepareCall(string stage, string library, string module, string function, params object[] args)
    {
        object value;
        try { value = host.Execute(library, module, function, args); }
        catch (Exception e)
        {
            err.WriteLine(stage + ": " + module + "." + function + " failed: " + BridgeText.Describe(e));
            return null;
        }
        string text = value as string;
        if (text == null)
        {
            err.WriteLine(stage + ": " + module + "." + function + " returned " + BridgeText.TypeName(value));
            BridgeText.Release(value);
        }
        return text;
    }

    /// <summary>手順 7。C5D_List で自分以外の C5D_Bridge（残骸）を探し、正本のフォルダのものなら外す。だめなら false（13）。</summary>
    bool CloseWrecks(out int closed)
    {
        closed = 0;
        List<string> wrecks;
        if (!ListBridges(out wrecks)) { return false; }
        if (wrecks.Count == 0) { return true; }
        if (wrecks.Count > MaxWrecks)
        {
            err.WriteLine("list: too many C5D_Bridge projects are loaded (" + wrecks.Count.ToString(CultureInfo.InvariantCulture) + "); restart CATIA");
            return false;
        }
        foreach (string wreck in wrecks)
        {
            if (!IsRemovableWreck(wreck))
            {
                err.WriteLine("list: cannot close a C5D_Bridge that is not ours: \"" + wreck + "\"; restart CATIA");
                return false;
            }
        }
        foreach (string wreck in wrecks)
        {
            Stopwatch watch = Stopwatch.StartNew();
            if (!CloseWreck(wreck)) { return false; }
            closed++;
            Log("wreck closed " + Path.GetFileName(wreck), watch);
        }
        List<string> left;
        if (!ListBridges(out left)) { return false; }
        if (left.Count != 0)
        {
            err.WriteLine("list: C5D_Bridge projects are still loaded after closing: " + string.Join(", ", left.ToArray()));
            return false;
        }
        return true;
    }

    /// <summary>C5D_List を読み、自分がちょうど 1 行あることを確かめて、自分以外の C5D_Bridge の DisplayName を返す。</summary>
    bool ListBridges(out List<string> wrecks)
    {
        wrecks = new List<string>();
        string text = PrepareCall("list", link, BridgeText.PushModule, "C5D_List");
        if (text == null) { return false; }
        List<string[]> rows = BridgeText.ParseRecords(text, 3);
        if (rows == null) { err.WriteLine("list: unexpected C5D_List: " + BridgeText.Shorten(text)); return false; }
        int selfRows = 0;
        foreach (string[] row in rows)
        {
            if (row[0] != BridgeText.ProjectName) { continue; } // Name が空の行は利用者のプロジェクトとみなす
            bool me = row[2] == "1";
            bool samePath = SamePath(row[1], link);
            if (me && samePath) { selfRows++; continue; }
            if (me || samePath)
            {
                err.WriteLine("list: inconsistent row for this link (DisplayName=" + row[1] + ", self=" + row[2] + ")");
                return false;
            }
            wrecks.Add(row[1]);
        }
        if (selfRows != 1)
        {
            err.WriteLine("list: this link is listed " + selfRows.ToString(CultureInfo.InvariantCulture) + " time(s)");
            return false;
        }
        return true;
    }

    /// <summary>DisplayName が完全修飾の絶対パス（. / .. を含まない）で、フォルダが正本のフォルダ（長い名前で比べる）で、ファイル名が C5D_Bridge*.catvba。</summary>
    bool IsRemovableWreck(string displayName)
    {
        if (string.IsNullOrEmpty(displayName)) { return false; }
        try
        {
            PathNames.CheckAbsolute("wreck", displayName);
            return PathNames.Same(PathNames.Parent(displayName), masterFolder)
                && WreckFileName.IsMatch(Path.GetFileName(displayName));
        }
        catch (Exception) { return false; }
    }

    /// <summary>残骸自身のパスで C5D_Self → guard（Name / DisplayName）→ Close(False)。</summary>
    bool CloseWreck(string displayName)
    {
        object self = null;
        try
        {
            self = host.Execute(displayName, BridgeText.PushModule, "C5D_Self", new object[0]);
            if (!BridgeText.IsObject(self))
            {
                err.WriteLine("wreck: C5D_Self of \"" + displayName + "\" returned " + (self is string ? BridgeText.Shorten((string)self) : BridgeText.TypeName(self)));
                return false;
            }
            string name = BridgeText.GetProperty(self, "Name") as string;
            string shown = BridgeText.GetProperty(self, "DisplayName") as string;
            if (name != BridgeText.ProjectName || !SamePath(shown, displayName))
            {
                err.WriteLine("wreck: guard failed for \"" + displayName + "\" (Name=" + name + ", DisplayName=" + shown + ")");
                return false;
            }
            self.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, self, new object[] { false }, CultureInfo.InvariantCulture);
            return true;
        }
        catch (Exception e)
        {
            err.WriteLine("wreck: cannot close \"" + displayName + "\": " + BridgeText.Describe(e));
            return false;
        }
        finally { BridgeText.Release(self); }
    }

    /// <summary>手順 8。部品一覧に runtime があり版が同じなら何もしない。無い・違うなら C5D_Install で入れて版を確かめる。</summary>
    bool EnsureRuntime(List<string[]> who, out string state)
    {
        state = null;
        bool present = false;
        for (int i = 1; i < who.Count; i++)
        {
            if (string.Equals(who[i][0], BridgeText.RuntimeModule, StringComparison.OrdinalIgnoreCase)) { present = true; break; }
        }
        if (present)
        {
            string current = PrepareCall("runtime", link, BridgeText.RuntimeModule, "C5D_RT_Version");
            if (current == runtime.Version) { state = "ok " + runtime.Version; return true; }
            err.WriteLine("runtime: installed version " + BridgeText.Shorten(current) + " differs from " + runtime.Version + "; installing");
        }
        string installed = PrepareCall("runtime", link, BridgeText.PushModule, "C5D_Install", BridgeText.RuntimeModule, runtime.Text);
        if (installed == null) { return false; }
        if (installed != "ok") { err.WriteLine("runtime: C5D_Install failed: " + BridgeText.Shorten(installed)); return false; }
        string after = PrepareCall("runtime", link, BridgeText.RuntimeModule, "C5D_RT_Version");
        if (after == null) { return false; }
        if (after != runtime.Version)
        {
            err.WriteLine("runtime: version after install is " + BridgeText.Shorten(after) + ", expected " + runtime.Version);
            return false;
        }
        state = "installed " + runtime.Version;
        return true;
    }

    /// <summary>
    /// C5D_Bridge_Runtime の関数を 1 回呼ぶ。recordFields &gt; 0 ならレコード形式、-1 なら "ok:" 始まり、0 なら "ok" ちょうどを成功とみなす。
    /// </summary>
    RuntimeReply RuntimeCall(string function, int recordFields, params object[] args)
    {
        RuntimeReply reply = new RuntimeReply();
        object value;
        try { value = host.Execute(link, BridgeText.RuntimeModule, function, args); }
        catch (Exception e)
        {
            Exception inner = BridgeText.Unwrap(e);
            err.WriteLine(function + ": " + BridgeText.Describe(inner));
            reply.Code = inner.HResult;
            reply.Message = inner.Message;
            return reply;
        }
        string text = value as string;
        if (text == null)
        {
            reply.Code = UnexpectedReturnCode;
            reply.Message = "Unexpected return value: " + BridgeText.TypeName(value);
            BridgeText.Release(value);
            err.WriteLine(function + ": " + reply.Message);
            return reply;
        }
        bool ok = recordFields > 0 ? BridgeText.ParseRecords(text, recordFields) != null
            : recordFields < 0 ? text.StartsWith("ok:", StringComparison.Ordinal)
            : text == "ok";
        if (ok) { reply.Ok = true; reply.Text = text; return reply; }
        string phase, message;
        int code;
        if (BridgeText.TryParseError(text, out phase, out code, out message))
        {
            reply.IsError = true;
            reply.Phase = phase;
            reply.Code = code;
            reply.Message = message;
            return reply;
        }
        reply.Code = UnexpectedReturnCode;
        reply.Message = "Unexpected return value: " + BridgeText.Shorten(text);
        err.WriteLine(function + ": " + reply.Message);
        return reply;
    }

    /// <summary>5.2 の変換。error: ならそのまま、異常なら callPhase。target は connect. / project. なら targetProject、それ以外は callTarget。</summary>
    void AddFailure(ResultData result, RuntimeReply reply, string callPhase, string callTarget)
    {
        Failure failure = new Failure();
        failure.Phase = reply.IsError ? reply.Phase : callPhase;
        failure.Target = reply.IsError && IsProjectPhase(reply.Phase) ? request.TargetProject : callTarget;
        failure.Code = reply.Code;
        failure.Message = reply.Message ?? "";
        result.Failures.Add(failure);
    }

    static bool IsProjectPhase(string phase)
    {
        return phase.StartsWith("connect.", StringComparison.Ordinal) || phase.StartsWith("project.", StringComparison.Ordinal);
    }

    /// <summary>pull / check。戻り値は「ExecuteScript の異常で止めた」（終了コード 12）。</summary>
    bool Export(ResultData result)
    {
        userCallsStarted = true;
        RuntimeReply list = RuntimeCall("C5D_RT_ListModules", 2, request.TargetProject);
        if (!list.Ok)
        {
            AddFailure(result, list, "export.call", request.TargetProject);
            return !list.IsError;
        }
        List<string[]> components = BridgeText.ParseRecords(list.Text, 2);
        List<string[]> modules = new List<string[]>();
        foreach (string[] component in components)
        {
            int type;
            if (!int.TryParse(component[1], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out type))
            {
                RuntimeReply bad = new RuntimeReply();
                bad.Code = UnexpectedReturnCode;
                bad.Message = "Unexpected return value: module type " + BridgeText.Shorten(component[1]);
                err.WriteLine("C5D_RT_ListModules: " + bad.Message);
                AddFailure(result, bad, "export.call", request.TargetProject);
                return true;
            }
            if (type >= 1 && type <= 3) { modules.Add(component); }
        }
        result.ExpectedCount = modules.Count;
        for (int i = 0; i < modules.Count; i++)
        {
            string name = modules[i][0];
            int type = int.Parse(modules[i][1], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            RuntimeReply read = RuntimeCall("C5D_RT_ReadModule", -1, request.TargetProject, name);
            if (!read.Ok)
            {
                AddFailure(result, read, "export.call", name);
                if (!read.IsError) { return true; }
                if (IsProjectPhase(read.Phase)) { return false; }
                continue;
            }
            string file = "m" + i.ToString(CultureInfo.InvariantCulture) + ".txt";
            try
            {
                File.WriteAllText(Path.Combine(workFolder, file), read.Text.Substring(3), new UTF8Encoding(false));
            }
            catch (Exception e)
            {
                Exception inner = BridgeText.Unwrap(e);
                Failure failure = new Failure();
                failure.Phase = "export.save";
                failure.Target = name;
                failure.Code = inner.HResult;
                failure.Message = inner.Message;
                result.Failures.Add(failure);
                continue;
            }
            ResultItem item = new ResultItem();
            item.Name = name;
            item.Type = type;
            item.File = file;
            result.Items.Add(item);
            result.SucceededCount++;
        }
        return false;
    }

    /// <summary>push。modules を順に Write、続いて deletions を順に Delete。戻り値は「ExecuteScript の異常で止めた」。</summary>
    bool Push(ResultData result)
    {
        result.ExpectedCount = request.Modules.Count + request.Deletions.Count;
        foreach (ModuleRequest module in request.Modules)
        {
            userCallsStarted = true;
            RuntimeReply reply = RuntimeCall("C5D_RT_WriteModule", 0, request.TargetProject, module.Name,
                module.Type.ToString(CultureInfo.InvariantCulture), module.Text);
            if (reply.Ok) { result.UpdatedCount++; result.SucceededCount++; continue; }
            AddFailure(result, reply, "push.call", module.Name);
            if (!reply.IsError) { return true; }
            if (IsProjectPhase(reply.Phase)) { return false; }
        }
        foreach (string name in request.Deletions)
        {
            userCallsStarted = true;
            RuntimeReply reply = RuntimeCall("C5D_RT_DeleteModule", 0, request.TargetProject, name);
            if (reply.Ok) { result.DeletedCount++; result.SucceededCount++; continue; }
            AddFailure(result, reply, "push.call", name);
            if (!reply.IsError) { return true; }
            if (IsProjectPhase(reply.Phase)) { return false; }
        }
        return false;
    }

    /// <summary>result.json.tmp に書いてから rename。書けなければ 12（完了を確認できない）。</summary>
    int WriteResult(ResultData result, int code)
    {
        Stopwatch watch = Stopwatch.StartNew();
        string tmp = resultPath + ".tmp";
        try
        {
            byte[] bytes = new UTF8Encoding(false).GetBytes(result.ToJson(request));
            using (FileStream stream = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                stream.Write(bytes, 0, bytes.Length);
            }
            File.Move(tmp, resultPath);
        }
        catch (Exception e)
        {
            err.WriteLine("error: cannot write result.json: " + BridgeText.Describe(e));
            return ExecuteFailed;
        }
        Log("result", watch);
        return code;
    }

    /// <summary>手順 12。C5D_Self → Name / DisplayName / VBProject.FileName を確かめて Close(False)。失敗は警告だけ。</summary>
    void CloseSelf()
    {
        Stopwatch watch = Stopwatch.StartNew();
        object self = null, project = null;
        try
        {
            self = host.Execute(link, BridgeText.PushModule, "C5D_Self", new object[0]);
            if (!BridgeText.IsObject(self))
            {
                err.WriteLine("warning: bridge close failed: C5D_Self returned "
                    + (self is string ? BridgeText.Shorten((string)self) : BridgeText.TypeName(self)));
                log.WriteLine("close failed " + watch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "ms");
                return;
            }
            string name = BridgeText.GetProperty(self, "Name") as string;
            string shown = BridgeText.GetProperty(self, "DisplayName") as string;
            project = BridgeText.GetProperty(self, "VBProject");
            string file = project == null ? null : BridgeText.GetProperty(project, "FileName") as string;
            if (name != BridgeText.ProjectName || !SamePath(shown, link) || !SamePath(file, link))
            {
                err.WriteLine("warning: bridge close failed: guard mismatch (Name=" + name + ", DisplayName=" + shown + ", FileName=" + file
                    + ", link=" + link + ")");
                log.WriteLine("close skipped " + watch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "ms");
                return;
            }
            self.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, self, new object[] { false }, CultureInfo.InvariantCulture);
            Log("close ok", watch);
        }
        catch (Exception e)
        {
            err.WriteLine("warning: bridge close failed: " + BridgeText.Describe(e));
            log.WriteLine("close failed " + watch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "ms");
        }
        finally
        {
            BridgeText.Release(project);
            BridgeText.Release(self);
        }
    }
}
