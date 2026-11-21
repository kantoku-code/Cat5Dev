// catia-run.exe: 起動中の CATIA で SystemService.ExecuteScript を呼ぶヘルパー（旧 run.vbs の置き換え）。
// Build: npm run compile:bridge（.NET Framework 4.x 標準の csc.exe。C# 5 の構文に留める）
// Usage: catia-run.exe <directory> <scriptName>
// Exit codes (src/catiaExecution.ts と共有。10/11 は「CATIA 未着手」として扱われる):
//   0 = 成功, 2 = 引数不正, 10 = CATIA 取得失敗, 11 = SystemService 取得失敗, 12 = ExecuteScript 例外
using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

internal static class CatiaRun
{
    const int Ok = 0, BadArguments = 2, NoCatia = 10, NoSystemService = 11, ExecuteFailed = 12;
    // catScriptLibraryTypeDirectory: iLibraryName をディレクトリとして iProgramName を探す
    const int LibraryTypeDirectory = 1;

    // Node 側は UTF-8 で読むため、コンソールのコードページに依存せず UTF-8 で書く。
    static readonly TextWriter Err = new StreamWriter(Console.OpenStandardError(), new UTF8Encoding(false)) { AutoFlush = true };

    [STAThread]
    static int Main(string[] args)
    {
        if (args.Length != 2 || args[0].Length == 0 || args[1].Length == 0)
        {
            Err.WriteLine("usage: catia-run <directory> <scriptName>");
            return BadArguments;
        }
        object app = null, service = null;
        try
        {
            try { app = Marshal.GetActiveObject("CATIA.Application"); }
            catch (Exception e) { Report(e); return NoCatia; }
            if (app == null) { return NoCatia; }
            try { service = app.GetType().InvokeMember("SystemService", BindingFlags.GetProperty, null, app, null); }
            catch (Exception e) { Report(e); return NoSystemService; }
            if (service == null) { return NoSystemService; }
            // ExecuteScript(iLibraryName, iType, iProgramName, iFunctionName, iParameters)
            // iParameters (CATSafeArrayVariant) は ByRef の可能性がある。VBScript は変数を VT_BYREF で渡していた。
            // 型不一致 (0x80020005 DISP_E_TYPEMISMATCH 等) で失敗する場合は、次のように ByRef 指定に切り替える:
            //   ParameterModifier byRef = new ParameterModifier(5); byRef[4] = true;
            //   InvokeMember("ExecuteScript", BindingFlags.InvokeMethod, null, service, parameters,
            //       new ParameterModifier[] { byRef }, null, null);
            object[] parameters = { args[0], LibraryTypeDirectory, args[1], "CATMain", new object[0] };
            try { service.GetType().InvokeMember("ExecuteScript", BindingFlags.InvokeMethod, null, service, parameters); }
            catch (Exception e) { Report(e); return ExecuteFailed; }
            return Ok;
        }
        finally
        {
            if (service != null && Marshal.IsComObject(service)) { Marshal.ReleaseComObject(service); }
            if (app != null && Marshal.IsComObject(app)) { Marshal.ReleaseComObject(app); }
        }
    }

    /// <summary>"0xHRESULT: メッセージ" を stderr に出す。レイトバインドの例外は InnerException に本体がある。</summary>
    static void Report(Exception e)
    {
        while (e is TargetInvocationException && e.InnerException != null) { e = e.InnerException; }
        Err.WriteLine("0x" + e.HResult.ToString("X8") + ": " + e.Message);
    }
}
