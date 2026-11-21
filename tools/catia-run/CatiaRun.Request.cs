// catia-run.exe: コマンドライン引数と request.json の検証（docs/dev/bridge-contract.md 3・4 章）。
// どちらも CATIA に接続する前に行い、不正なら終了コード 2。
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

/// <summary>検証で見つけた不正。Message は stderr に出す（英語）。</summary>
internal sealed class InvalidInputException : Exception
{
    public InvalidInputException(string message) : base(message) { }
}

internal sealed class RunOptions
{
    public bool ValidateOnly;
    public string Bridge;
    public string Request;
    public string Result;

    static readonly Regex MasterName = new Regex(@"^C5D_Bridge-[0-9a-f]{8}\.catvba$", RegexOptions.CultureInvariant);

    /// <summary>
    /// --bridge / --request / --result（どれも必須・1 回だけ・順不同、値は Path.GetFullPath しても変わらない絶対パス）と
    /// --validate-only（1 回まで）。--request と --result は同じフォルダ、--result は未作成、--bridge は正本の名前で存在すること。
    /// </summary>
    public static RunOptions Parse(string[] args)
    {
        RunOptions options = new RunOptions();
        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i];
            if (arg == "--validate-only")
            {
                if (options.ValidateOnly) { throw new InvalidInputException("--validate-only is given twice"); }
                options.ValidateOnly = true;
                continue;
            }
            if (arg != "--bridge" && arg != "--request" && arg != "--result") { throw new InvalidInputException("unknown argument: " + arg); }
            if (i + 1 >= args.Length) { throw new InvalidInputException(arg + " needs a value"); }
            string value = args[++i];
            if ((arg == "--bridge" && options.Bridge != null) || (arg == "--request" && options.Request != null)
                || (arg == "--result" && options.Result != null))
            {
                throw new InvalidInputException(arg + " is given twice");
            }
            CheckAbsolute(arg, value);
            if (arg == "--bridge") { options.Bridge = value; }
            else if (arg == "--request") { options.Request = value; }
            else { options.Result = value; }
        }
        if (options.Bridge == null) { throw new InvalidInputException("--bridge is required"); }
        if (options.Request == null) { throw new InvalidInputException("--request is required"); }
        if (options.Result == null) { throw new InvalidInputException("--result is required"); }
        if (!string.Equals(Path.GetDirectoryName(options.Request), Path.GetDirectoryName(options.Result), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidInputException("--request and --result must be in the same folder");
        }
        if (string.Equals(options.Request, options.Result, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidInputException("--request and --result must differ");
        }
        if (File.Exists(options.Result) || Directory.Exists(options.Result)) { throw new InvalidInputException("--result already exists: " + options.Result); }
        if (!MasterName.IsMatch(Path.GetFileName(options.Bridge)))
        {
            throw new InvalidInputException("--bridge must be named C5D_Bridge-<sha8>.catvba: " + options.Bridge);
        }
        if (!File.Exists(options.Bridge)) { throw new InvalidInputException("--bridge does not exist: " + options.Bridge); }
        return options;
    }

    static void CheckAbsolute(string option, string value)
    {
        string full;
        try { full = Path.GetFullPath(value); }
        catch (Exception e) { throw new InvalidInputException(option + " is not a valid path: " + value + " (" + e.Message + ")"); }
        if (!string.Equals(full, value, StringComparison.Ordinal) || !Path.IsPathRooted(value))
        {
            throw new InvalidInputException(option + " must be a full absolute path: " + value);
        }
    }
}

internal sealed class ModuleRequest
{
    public string Name;
    public int Type;
    public string Text;
}

/// <summary>request.json（version 2）。encoding など未知のフィールドは無視する。</summary>
internal sealed class BridgeRequest
{
    public const long MaxBytes = 64L * 1024 * 1024;

    public string OperationId;
    public string StepId;
    public string Kind;
    public string TargetProject;
    public List<ModuleRequest> Modules = new List<ModuleRequest>();
    public List<string> Deletions = new List<string>();

    public static BridgeRequest Read(string path)
    {
        byte[] bytes;
        try
        {
            FileInfo info = new FileInfo(path);
            if (!info.Exists) { throw new InvalidInputException("request file does not exist: " + path); }
            if (info.Length > MaxBytes) { throw new InvalidInputException("request file is larger than 64 MiB"); }
            bytes = File.ReadAllBytes(path);
        }
        catch (InvalidInputException) { throw; }
        catch (Exception e) { throw new InvalidInputException("cannot read the request file: " + e.Message); }
        if (bytes.LongLength > MaxBytes) { throw new InvalidInputException("request file is larger than 64 MiB"); }
        int offset = bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF ? 3 : 0;
        string text;
        try { text = new UTF8Encoding(false, true).GetString(bytes, offset, bytes.Length - offset); }
        catch (DecoderFallbackException) { throw new InvalidInputException("request file is not valid UTF-8"); }
        object root;
        try { root = JsonParser.Parse(text); }
        catch (FormatException e) { throw new InvalidInputException("request file is not valid JSON: " + e.Message); }
        return FromJson(root);
    }

    public static BridgeRequest FromJson(object rootValue)
    {
        Dictionary<string, object> root = rootValue as Dictionary<string, object>;
        if (root == null) { throw new InvalidInputException("request must be a JSON object"); }
        object version;
        if (!root.TryGetValue("version", out version) || !(version is double) || (double)version != 2)
        {
            throw new InvalidInputException("version must be the number 2");
        }
        BridgeRequest request = new BridgeRequest();
        request.OperationId = NonEmptyString(root, "operationId");
        request.StepId = NonEmptyString(root, "stepId");
        request.Kind = RequiredString(root, "kind");
        if (request.Kind != "list" && request.Kind != "pull" && request.Kind != "check" && request.Kind != "push")
        {
            throw new InvalidInputException("kind must be list, pull, check or push");
        }
        request.TargetProject = RequiredString(root, "targetProject");
        if (request.Kind == "list" && request.TargetProject.Length != 0) { throw new InvalidInputException("targetProject must be empty for list"); }
        if (request.Kind != "list" && request.TargetProject.Length == 0) { throw new InvalidInputException("targetProject must not be empty"); }

        List<object> modules = OptionalArray(root, "modules");
        List<object> deletions = OptionalArray(root, "deletions");
        if (request.Kind != "push" && (modules.Count != 0 || deletions.Count != 0))
        {
            throw new InvalidInputException("modules and deletions are only for push");
        }
        HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (int i = 0; i < modules.Count; i++)
        {
            string where = "modules[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
            Dictionary<string, object> item = modules[i] as Dictionary<string, object>;
            if (item == null) { throw new InvalidInputException(where + " must be an object"); }
            ModuleRequest module = new ModuleRequest();
            module.Name = ModuleName(item, "name", where + ".name");
            object type;
            if (!item.TryGetValue("type", out type) || !(type is double)
                || ((double)type != 1 && (double)type != 2 && (double)type != 3))
            {
                throw new InvalidInputException(where + ".type must be 1, 2 or 3");
            }
            module.Type = (int)(double)type;
            object moduleText;
            if (!item.TryGetValue("text", out moduleText) || !(moduleText is string)) { throw new InvalidInputException(where + ".text must be a string"); }
            module.Text = (string)moduleText;
            if (!names.Add(module.Name)) { throw new InvalidInputException("duplicate module name: " + module.Name); }
            request.Modules.Add(module);
        }
        for (int i = 0; i < deletions.Count; i++)
        {
            string where = "deletions[" + i.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]";
            string name = deletions[i] as string;
            if (name == null) { throw new InvalidInputException(where + " must be a string"); }
            CheckModuleName(name, where);
            if (!names.Add(name)) { throw new InvalidInputException("duplicate module name: " + name); }
            request.Deletions.Add(name);
        }
        return request;
    }

    static string RequiredString(Dictionary<string, object> obj, string key)
    {
        object value;
        if (!obj.TryGetValue(key, out value) || !(value is string)) { throw new InvalidInputException(key + " must be a string"); }
        return (string)value;
    }

    static string NonEmptyString(Dictionary<string, object> obj, string key)
    {
        string value = RequiredString(obj, key);
        if (value.Length == 0) { throw new InvalidInputException(key + " must not be empty"); }
        return value;
    }

    static List<object> OptionalArray(Dictionary<string, object> obj, string key)
    {
        object value;
        if (!obj.TryGetValue(key, out value)) { return new List<object>(); }
        List<object> list = value as List<object>;
        if (list == null) { throw new InvalidInputException(key + " must be an array"); }
        return list;
    }

    static string ModuleName(Dictionary<string, object> obj, string key, string where)
    {
        object value;
        if (!obj.TryGetValue(key, out value) || !(value is string)) { throw new InvalidInputException(where + " must be a string"); }
        CheckModuleName((string)value, where);
        return (string)value;
    }

    /// <summary>空でない、U+0000〜U+001F と \ / : * ? " &lt; &gt; | を含まない。</summary>
    static void CheckModuleName(string name, string where)
    {
        if (name.Length == 0) { throw new InvalidInputException(where + " must not be empty"); }
        foreach (char c in name)
        {
            if (c < ' ' || "\\/:*?\"<>|".IndexOf(c) >= 0) { throw new InvalidInputException(where + " has a character that is not allowed: " + name); }
        }
    }
}
