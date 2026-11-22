// catia-run.exe: 起動中の CATIA に仲介プロジェクト C5D_Bridge（v2）を読み込み、request.json の要求を本物の VBA で処理して
// result.json を書くヘルパー。仕様: docs/dev/bridge-contract.md（3〜7 章）。
// Build: npm run compile:bridge（.NET Framework 4.x 標準の csc.exe。C# 5 の構文に留める。runtime の本文を埋め込む）
// Usage: catia-run.exe [--validate-only] --bridge <master.catvba> --request <dir>\request.json --result <dir>\result.json
// Exit codes (src/catiaExecution.ts と共有):
//   0 = 処理を最後まで行った, 2 = 引数・request.json が不正, 10 = CATIA 取得失敗, 11 = SystemService 取得失敗,
//   12 = 要求の処理中に ExecuteScript が例外, 13 = ブリッジを準備できない,
//   14 = VBA プロジェクトが見えない（VBA エディタが未起動か、プロジェクトが読み込まれていない）, 15 = 別の catia-run.exe が実行中（CATIA に触る前）
// stdout: 段階ごとに 1 行（"<stage> <ms>ms"。人が読むログ）。stderr: エラーの詳細。どちらも UTF-8（BOM なし）。
// 同じフォルダの CatiaRun.*.cs と一緒にコンパイルする（EvalProbe.cs / TypeLibDump.cs は含めない）。
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

internal static class CatiaRun
{
    const int Ok = 0, BadArguments = 2, NoCatia = 10, NoSystemService = 11, BridgeNotReady = 13, RunnerBusy = 15;
    /// <summary>tests/catiaRun.test.js のハーネスもこの名前で Mutex を持つ。</summary>
    internal const string MutexName = @"Local\Cat5Dev.catia-run";
    const string Usage = "usage: catia-run [--validate-only] --bridge <C5D_Bridge-<sha8>.catvba> --request <dir>\\request.json --result <dir>\\result.json";

    // Node 側は UTF-8 で読むため、コンソールのコードページに依存せず UTF-8 で書く。
    static readonly TextWriter Out = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = true };
    static readonly TextWriter Err = new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false)) { AutoFlush = true };

    [STAThread]
    static int Main(string[] args)
    {
        Stopwatch total = Stopwatch.StartNew();
        int code = Run(args);
        Out.WriteLine("exit " + code.ToString(CultureInfo.InvariantCulture) + " " + total.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "ms");
        return code;
    }

    static int Run(string[] args)
    {
        // 1. 引数と request.json の検証
        Stopwatch watch = Stopwatch.StartNew();
        RunOptions options;
        BridgeRequest request;
        try
        {
            options = RunOptions.Parse(args);
            request = BridgeRequest.Read(options.Request);
        }
        catch (InvalidInputException e)
        {
            Err.WriteLine("invalid arguments or request: " + e.Message);
            Err.WriteLine(Usage);
            return BadArguments;
        }
        Log("validate " + request.Kind, watch);
        if (options.ValidateOnly) { return Ok; }

        string error;
        RuntimeSource runtime = RuntimeSource.Load(out error);
        if (runtime == null) { Err.WriteLine("error: " + error); return BridgeNotReady; }

        // 2. 同時に 2 個動かさない（取れなければ CATIA に触る前に 15）
        Mutex mutex = null;
        bool owned = false;
        try
        {
            try
            {
                mutex = new Mutex(false, MutexName);
                try { owned = mutex.WaitOne(0); }
                catch (AbandonedMutexException) { owned = true; } // 前の catia-run.exe が殺された
            }
            catch (Exception e) { Err.WriteLine("error: cannot open the mutex: " + BridgeText.Describe(e)); return BridgeNotReady; }
            if (!owned) { Err.WriteLine("another catia-run.exe is running"); return RunnerBusy; }
            return RunWithCatia(options, request, runtime);
        }
        finally
        {
            if (owned) { try { mutex.ReleaseMutex(); } catch (Exception) { } }
            if (mutex != null) { mutex.Dispose(); }
        }
    }

    static int RunWithCatia(RunOptions options, BridgeRequest request, RuntimeSource runtime)
    {
        object app = null, service = null;
        try
        {
            // 3. CATIA / SystemService
            Stopwatch watch = Stopwatch.StartNew();
            try { app = Marshal.GetActiveObject("CATIA.Application"); }
            catch (Exception e) { Err.WriteLine("CATIA is not running: " + BridgeText.Describe(e)); return NoCatia; }
            if (app == null) { Err.WriteLine("CATIA is not running"); return NoCatia; }
            try { service = app.GetType().InvokeMember("SystemService", BindingFlags.GetProperty, null, app, null); }
            catch (Exception e) { Err.WriteLine("SystemService: " + BridgeText.Describe(e)); return NoSystemService; }
            if (service == null) { Err.WriteLine("SystemService returned Nothing"); return NoSystemService; }
            Log("connect", watch);

            // 4. リンク
            watch = Stopwatch.StartNew();
            string method, error;
            string link = BridgeLink.Create(options.Bridge, out method, out error);
            if (link == null) { Err.WriteLine("link: " + error); return BridgeNotReady; }
            Log("link " + method + " " + Path.GetFileName(link), watch);

            // 5〜12
            BridgeSession session = new BridgeSession(new CatiaScriptHost(service), link, options.Bridge, runtime, request, options.Result, Out, Err);
            return session.Run();
        }
        finally
        {
            BridgeText.Release(service);
            BridgeText.Release(app);
        }
    }

    static void Log(string stage, Stopwatch watch)
    {
        Out.WriteLine(stage + " " + watch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + "ms");
    }
}

/// <summary>正本へのハードリンク（だめなら読み取り専用のコピー）を作る（2.2）。既存のファイルは決して上書きしない。</summary>
internal static class BridgeLink
{
    const int ErrorFileExists = 80, ErrorAlreadyExists = 183;
    const int MaxSuffix = 99;

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    static extern bool CreateHardLinkW(string lpFileName, string lpExistingFileName, IntPtr lpSecurityAttributes);

    /// <summary>作ったリンクのフルパス。作れない・読み取り専用でなければ null（error に理由）。method = hardlink / copy。</summary>
    public static string Create(string master, out string method, out string error)
    {
        method = null;
        error = null;
        string folder = Path.GetDirectoryName(master);
        string stem = Path.GetFileNameWithoutExtension(master) + "-" + CnextPid().ToString(CultureInfo.InvariantCulture) + "-"
            + DateTime.Now.ToString("yyyyMMddHHmmssfff", CultureInfo.InvariantCulture);
        for (int suffix = 0; suffix <= MaxSuffix; suffix++)
        {
            string link = Path.Combine(folder, stem + (suffix == 0 ? "" : "-" + suffix.ToString(CultureInfo.InvariantCulture)) + ".catvba");
            if (File.Exists(link) || Directory.Exists(link)) { continue; }
            if (CreateHardLinkW(link, master, IntPtr.Zero))
            {
                method = "hardlink";
                return CheckReadOnly(link, out error);
            }
            int hardLinkError = Marshal.GetLastWin32Error();
            if (hardLinkError == ErrorFileExists || hardLinkError == ErrorAlreadyExists) { continue; }
            try { File.Copy(master, link, false); }
            catch (Exception e)
            {
                if (File.Exists(link) && e is IOException && !(e is FileNotFoundException)) { continue; } // 同じ名前が先にできた
                error = "cannot create a hard link (win32 error " + hardLinkError.ToString(CultureInfo.InvariantCulture) + ") nor a copy: "
                    + BridgeText.Describe(e);
                return null;
            }
            method = "copy";
            try { File.SetAttributes(link, File.GetAttributes(link) | FileAttributes.ReadOnly); }
            catch (Exception e) { error = "cannot make the copy read-only: " + BridgeText.Describe(e); return null; }
            return CheckReadOnly(link, out error);
        }
        error = "every link name is taken: " + Path.Combine(folder, stem) + "[-1..-99].catvba";
        return null;
    }

    static string CheckReadOnly(string link, out string error)
    {
        error = null;
        try
        {
            if ((File.GetAttributes(link) & FileAttributes.ReadOnly) != 0) { return link; }
            error = "the link is not read-only (the master must be read-only): " + link;
        }
        catch (Exception e) { error = "cannot read the attributes of the link: " + BridgeText.Describe(e); }
        return null;
    }

    /// <summary>CNEXT.exe がちょうど 1 個ならその pid、それ以外は 0（診断用）。</summary>
    static int CnextPid()
    {
        Process[] processes = null;
        try
        {
            processes = Process.GetProcessesByName("CNEXT");
            return processes.Length == 1 ? processes[0].Id : 0;
        }
        catch (Exception) { return 0; }
        finally
        {
            if (processes != null) { foreach (Process p in processes) { p.Dispose(); } }
        }
    }
}
