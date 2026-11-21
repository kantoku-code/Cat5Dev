// typelib-dump.exe（開発者用・非出荷）: 型ライブラリ（.dll / .olb / .tlb の TYPELIB リソース）の全メンバーをテキストで出す。
// Build: node scripts/dump-typelib.js が temp/typelib-dump.exe にビルドして実行する（.NET Framework 4.x の csc.exe、C# 5 の構文に留める）。
// Usage: typelib-dump.exe [--max-index <n>] [--highlight <regex>] <file>[\<resource index>] ...
//   <file> にリソース番号が付いていなければ 1..max-index（既定 16）を順に LoadTypeLibEx（REGKIND_NONE、登録しない）で試す。
//   読み込みのみ。COM オブジェクトは作らない（CoCreateInstance しない）ので、CATIA や APC には一切触れない。
// Output (UTF-8, stdout):
//   == <file>\<n>: <ライブラリ名> {GUID} v<major>.<minor> lcid=<n> "<説明>"
//   <kind> <型名> {GUID} flags=<TYPEFLAGS の名前> [cFuncs/cVars/cImplTypes]
//     implements/inherits <型名> [default|source|restricted]
//     [<DISPID 16 進>] <func|get|let|set> <名前>(<引数>) As <型> funcflags=<FUNCFLAGS の名前> invoke=<func|propget|propput|propputref> oVft=<n>
//     [<DISPID>] var <名前> As <型> = <定数値>
//   最後に "== highlight" 節: 名前が --highlight の正規表現（既定 close|remove|unload|...）に一致したメンバーの一覧。
// Exit codes: 0 = 少なくとも 1 つの型ライブラリを読めた, 1 = 1 つも読めなかった, 2 = 引数不正
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Text.RegularExpressions;
using FUNCDESC = System.Runtime.InteropServices.ComTypes.FUNCDESC;
using TYPEDESC = System.Runtime.InteropServices.ComTypes.TYPEDESC;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using TYPEFLAGS = System.Runtime.InteropServices.ComTypes.TYPEFLAGS;
using FUNCFLAGS = System.Runtime.InteropServices.ComTypes.FUNCFLAGS;
using VARFLAGS = System.Runtime.InteropServices.ComTypes.VARFLAGS;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using VARDESC = System.Runtime.InteropServices.ComTypes.VARDESC;
using ELEMDESC = System.Runtime.InteropServices.ComTypes.ELEMDESC;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using VARKIND = System.Runtime.InteropServices.ComTypes.VARKIND;
using INVOKEKIND = System.Runtime.InteropServices.ComTypes.INVOKEKIND;
using IMPLTYPEFLAGS = System.Runtime.InteropServices.ComTypes.IMPLTYPEFLAGS;
using PARAMFLAG = System.Runtime.InteropServices.ComTypes.PARAMFLAG;
using FUNCKIND = System.Runtime.InteropServices.ComTypes.FUNCKIND;
using ITypeInfo = System.Runtime.InteropServices.ComTypes.ITypeInfo;
using ITypeLib = System.Runtime.InteropServices.ComTypes.ITypeLib;

internal static class TypeLibDump
{
    const int RegKindNone = 2;
    const string DefaultHighlight = "close|remove|unload|release|detach|delete|unregister|destroy|terminate|discard|kill|free|reset|clear|mode|projecttype|type$|"
        + "dirty|locked|protection|filename|fullname|path|saved|readonly|apcproject|^projects?$|vbproject|vbe$|unprotect|shutdown|quit|dispose|close";

    [DllImport("oleaut32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    static extern ITypeLib LoadTypeLibEx(string file, int regKind);

    static readonly TextWriter Out = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false)) { AutoFlush = false };
    static readonly List<string> Highlights = new List<string>();
    static Regex highlight = new Regex(DefaultHighlight, RegexOptions.IgnoreCase);

    [STAThread]
    static int Main(string[] args)
    {
        int maxIndex = 16;
        List<string> files = new List<string>();
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--max-index" && i + 1 < args.Length) { maxIndex = int.Parse(args[++i], CultureInfo.InvariantCulture); continue; }
            if (args[i] == "--highlight" && i + 1 < args.Length) { highlight = new Regex(args[++i], RegexOptions.IgnoreCase); continue; }
            files.Add(args[i]);
        }
        if (files.Count == 0)
        {
            Console.Error.WriteLine("usage: typelib-dump [--max-index <n>] [--highlight <regex>] <file>[\\<index>] ...");
            return 2;
        }
        int loaded = 0;
        foreach (string file in files)
        {
            Match m = Regex.Match(file, @"^(.*\.(dll|olb|tlb|exe|ocx))\\(\d+)$", RegexOptions.IgnoreCase);
            if (m.Success || file.EndsWith(".tlb", StringComparison.OrdinalIgnoreCase)) { if (DumpOne(file)) { loaded++; } continue; }
            if (!File.Exists(file)) { Out.WriteLine("== " + file + ": file not found"); continue; }
            for (int n = 1; n <= maxIndex; n++)
            {
                if (DumpOne(file + "\\" + n.ToString(CultureInfo.InvariantCulture))) { loaded++; }
            }
        }
        Out.WriteLine();
        Out.WriteLine("== highlight /" + highlight + "/ (" + Highlights.Count.ToString(CultureInfo.InvariantCulture) + " members)");
        foreach (string h in Highlights) { Out.WriteLine(h); }
        Out.Flush();
        return loaded > 0 ? 0 : 1;
    }

    static bool DumpOne(string spec)
    {
        ITypeLib lib;
        try { lib = LoadTypeLibEx(spec, RegKindNone); }
        catch (Exception e)
        {
            Out.WriteLine("== " + spec + ": not loaded (0x" + e.HResult.ToString("X8") + " " + e.Message.Trim() + ")");
            return false;
        }
        try
        {
            string libName, libDoc, helpFile; int helpContext;
            lib.GetDocumentation(-1, out libName, out libDoc, out helpContext, out helpFile);
            IntPtr pAttr;
            lib.GetLibAttr(out pAttr);
            TYPELIBATTR la = (TYPELIBATTR)Marshal.PtrToStructure(pAttr, typeof(TYPELIBATTR));
            lib.ReleaseTLibAttr(pAttr);
            Out.WriteLine();
            Out.WriteLine("== " + spec + ": " + libName + " {" + la.guid + "} v" + la.wMajorVerNum + "." + la.wMinorVerNum + " lcid=" + la.lcid
                + " syskind=" + la.syskind + " libflags=" + la.wLibFlags + " \"" + libDoc + "\"");
            int count = lib.GetTypeInfoCount();
            for (int i = 0; i < count; i++)
            {
                ITypeInfo info;
                lib.GetTypeInfo(i, out info);
                DumpType(libName, info);
            }
        }
        catch (Exception e)
        {
            Out.WriteLine("!! " + spec + ": dump failed (0x" + e.HResult.ToString("X8") + " " + e.Message.Trim() + ")");
        }
        return true;
    }

    static string Doc(ITypeInfo info, int memid, out string docString)
    {
        string name, helpFile; int helpContext;
        info.GetDocumentation(memid, out name, out docString, out helpContext, out helpFile);
        return name;
    }

    static void DumpType(string libName, ITypeInfo info)
    {
        IntPtr pAttr;
        info.GetTypeAttr(out pAttr);
        TYPEATTR ta = (TYPEATTR)Marshal.PtrToStructure(pAttr, typeof(TYPEATTR));
        info.ReleaseTypeAttr(pAttr);
        string doc;
        string typeName = Doc(info, -1, out doc);
        Out.WriteLine();
        Out.WriteLine(KindName(ta.typekind) + " " + typeName + " {" + ta.guid + "} flags=" + TypeFlags(ta.wTypeFlags)
            + " [funcs=" + ta.cFuncs + " vars=" + ta.cVars + " impl=" + ta.cImplTypes + "]" + (string.IsNullOrEmpty(doc) ? "" : " \"" + doc + "\""));
        if (ta.typekind == TYPEKIND.TKIND_ALIAS) { Out.WriteLine("  alias of " + TypeText(info, ta.tdescAlias)); }
        for (int i = 0; i < ta.cImplTypes; i++)
        {
            try
            {
                int href; info.GetRefTypeOfImplType(i, out href);
                ITypeInfo impl; info.GetRefTypeInfo(href, out impl);
                string d; string implName = Doc(impl, -1, out d);
                IMPLTYPEFLAGS f = 0;
                if (ta.typekind == TYPEKIND.TKIND_COCLASS) { info.GetImplTypeFlags(i, out f); }
                Out.WriteLine("  " + (ta.typekind == TYPEKIND.TKIND_COCLASS ? "implements " : "inherits ") + implName + (f != 0 ? " [" + f + "]" : ""));
            }
            catch (Exception e) { Out.WriteLine("  impl #" + i + " error 0x" + e.HResult.ToString("X8")); }
        }
        for (int i = 0; i < ta.cFuncs; i++)
        {
            IntPtr pFunc;
            info.GetFuncDesc(i, out pFunc);
            FUNCDESC fd = (FUNCDESC)Marshal.PtrToStructure(pFunc, typeof(FUNCDESC));
            try { DumpFunc(libName, typeName, info, fd); }
            catch (Exception e) { Out.WriteLine("  func #" + i + " error 0x" + e.HResult.ToString("X8") + " " + e.Message); }
            info.ReleaseFuncDesc(pFunc);
        }
        for (int i = 0; i < ta.cVars; i++)
        {
            IntPtr pVar;
            info.GetVarDesc(i, out pVar);
            VARDESC vd = (VARDESC)Marshal.PtrToStructure(pVar, typeof(VARDESC));
            try
            {
                string d; string name = Doc(info, vd.memid, out d);
                string value = "";
                if (vd.varkind == VARKIND.VAR_CONST && vd.desc.lpvarValue != IntPtr.Zero)
                {
                    object v = Marshal.GetObjectForNativeVariant(vd.desc.lpvarValue);
                    value = " = " + Convert.ToString(v, CultureInfo.InvariantCulture);
                }
                string line = "  [" + Dispid(vd.memid) + "] var " + name + " As " + TypeText(info, vd.elemdescVar.tdesc) + value
                    + (vd.wVarFlags != 0 ? " varflags=" + VarFlags((VARFLAGS)vd.wVarFlags) : "") + (string.IsNullOrEmpty(d) ? "" : " \"" + d + "\"");
                Out.WriteLine(line);
                if (highlight.IsMatch(name) && ta.typekind != TYPEKIND.TKIND_ENUM) { Highlights.Add(libName + "." + typeName + line); }
            }
            catch (Exception e) { Out.WriteLine("  var #" + i + " error 0x" + e.HResult.ToString("X8")); }
            info.ReleaseVarDesc(pVar);
        }
    }

    static void DumpFunc(string libName, string typeName, ITypeInfo info, FUNCDESC fd)
    {
        string[] names = new string[fd.cParams + 1];
        int got;
        info.GetNames(fd.memid, names, names.Length, out got);
        string d;
        string name = Doc(info, fd.memid, out d);
        List<string> ps = new List<string>();
        int size = Marshal.SizeOf(typeof(ELEMDESC));
        for (int p = 0; p < fd.cParams; p++)
        {
            ELEMDESC ed = (ELEMDESC)Marshal.PtrToStructure(new IntPtr(fd.lprgelemdescParam.ToInt64() + p * size), typeof(ELEMDESC));
            PARAMFLAG pf = ed.desc.paramdesc.wParamFlags;
            string pname = p + 1 < got && names[p + 1] != null ? names[p + 1] : "p" + p;
            StringBuilder b = new StringBuilder();
            if ((pf & PARAMFLAG.PARAMFLAG_FOPT) != 0) { b.Append("[opt] "); }
            if ((pf & PARAMFLAG.PARAMFLAG_FOUT) != 0) { b.Append("[out] "); }
            if ((pf & PARAMFLAG.PARAMFLAG_FRETVAL) != 0) { b.Append("[retval] "); }
            if ((pf & PARAMFLAG.PARAMFLAG_FLCID) != 0) { b.Append("[lcid] "); }
            b.Append(pname).Append(" As ").Append(TypeText(info, ed.tdesc));
            if ((pf & PARAMFLAG.PARAMFLAG_FHASDEFAULT) != 0 && ed.desc.paramdesc.lpVarValue != IntPtr.Zero)
            {
                try
                {
                    // PARAMDESCEX: ULONG cBytes; VARIANT varDefaultValue (8-byte aligned on x64)
                    IntPtr pv = new IntPtr(ed.desc.paramdesc.lpVarValue.ToInt64() + 8);
                    b.Append(" = ").Append(Convert.ToString(Marshal.GetObjectForNativeVariant(pv), CultureInfo.InvariantCulture));
                }
                catch (Exception) { b.Append(" = ?"); }
            }
            ps.Add(b.ToString());
        }
        string kind = fd.invkind == INVOKEKIND.INVOKE_FUNC ? "func" : fd.invkind == INVOKEKIND.INVOKE_PROPERTYGET ? "get"
            : fd.invkind == INVOKEKIND.INVOKE_PROPERTYPUT ? "let" : "set";
        string line = "  [" + Dispid(fd.memid) + "] " + kind + " " + name + "(" + string.Join(", ", ps.ToArray()) + ") As " + TypeText(info, fd.elemdescFunc.tdesc)
            + " funcflags=" + FuncFlags((FUNCFLAGS)fd.wFuncFlags) + " funckind=" + fd.funckind.ToString().Replace("FUNC_", "") + " oVft=" + fd.oVft
            + (string.IsNullOrEmpty(d) ? "" : " \"" + d + "\"");
        Out.WriteLine(line);
        if (highlight.IsMatch(name)) { Highlights.Add(libName + "." + typeName + line); }
    }

    static string Dispid(int memid) { return "0x" + memid.ToString("X8") + "/" + memid.ToString(CultureInfo.InvariantCulture); }

    static string TypeText(ITypeInfo info, TYPEDESC td)
    {
        VarEnum vt = (VarEnum)td.vt;
        switch (vt)
        {
            case VarEnum.VT_PTR:
                {
                    TYPEDESC inner = (TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(TYPEDESC));
                    return TypeText(info, inner) + "*";
                }
            case VarEnum.VT_SAFEARRAY:
                {
                    TYPEDESC inner = (TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(TYPEDESC));
                    return "SAFEARRAY(" + TypeText(info, inner) + ")";
                }
            case VarEnum.VT_CARRAY:
                {
                    TYPEDESC inner = (TYPEDESC)Marshal.PtrToStructure(td.lpValue, typeof(TYPEDESC));
                    return TypeText(info, inner) + "[]";
                }
            case VarEnum.VT_USERDEFINED:
                {
                    try
                    {
                        ITypeInfo r; info.GetRefTypeInfo(unchecked((int)(td.lpValue.ToInt64() & 0xFFFFFFFFL)), out r);
                        string d; return Doc(r, -1, out d);
                    }
                    catch (Exception e) { return "userdefined(href=0x" + (td.lpValue.ToInt64() & 0xFFFFFFFFL).ToString("X") + " err=0x" + e.HResult.ToString("X8") + ")"; }
                }
            default:
                return vt.ToString().Replace("VT_", "");
        }
    }

    static string KindName(TYPEKIND k) { return k.ToString().Replace("TKIND_", "").ToLowerInvariant(); }

    static string TypeFlags(TYPEFLAGS f)
    {
        if (f == 0) { return "0"; }
        return f.ToString().Replace("TYPEFLAG_F", "").Replace(" ", "");
    }

    static string FuncFlags(FUNCFLAGS f)
    {
        if (f == 0) { return "0"; }
        return "0x" + ((int)f).ToString("X") + "(" + f.ToString().Replace("FUNCFLAG_F", "").Replace(" ", "") + ")";
    }

    static string VarFlags(VARFLAGS f) { return "0x" + ((int)f).ToString("X") + "(" + f.ToString().Replace("VARFLAG_F", "").Replace(" ", "") + ")"; }
}
