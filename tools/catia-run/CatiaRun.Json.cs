// catia-run.exe: 厳密な JSON の読み取り（RFC 8259）と、result.json 用の文字列エスケープ。
// .NET Framework 標準の範囲（外部パッケージなし）。C# 5 の構文に留める。
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

/// <summary>JSON の null。Dictionary の値として null を入れないための目印。</summary>
internal sealed class JsonNull
{
    public static readonly JsonNull Value = new JsonNull();
    JsonNull() { }
}

/// <summary>
/// 値の対応: object → Dictionary&lt;string, object&gt;（キーの重複は不正）、array → List&lt;object&gt;、string → string、
/// number → double、true / false → bool、null → JsonNull.Value。入れ子は 64 段まで。不正なら FormatException。
/// </summary>
internal sealed class JsonParser
{
    const int MaxDepth = 64;
    readonly string text;
    int pos;

    JsonParser(string text) { this.text = text; }

    public static object Parse(string text)
    {
        JsonParser parser = new JsonParser(text);
        parser.SkipSpaces();
        object value = parser.ParseValue(0);
        parser.SkipSpaces();
        if (parser.pos != text.Length) { throw parser.Error("unexpected text after the value"); }
        return value;
    }

    FormatException Error(string message)
    {
        return new FormatException("JSON: " + message + " at offset " + pos.ToString(CultureInfo.InvariantCulture));
    }

    void SkipSpaces()
    {
        while (pos < text.Length)
        {
            char c = text[pos];
            if (c != ' ' && c != '\t' && c != '\n' && c != '\r') { return; }
            pos++;
        }
    }

    object ParseValue(int depth)
    {
        if (depth > MaxDepth) { throw Error("nested too deeply"); }
        if (pos >= text.Length) { throw Error("unexpected end"); }
        char c = text[pos];
        switch (c)
        {
            case '{': return ParseObject(depth);
            case '[': return ParseArray(depth);
            case '"': return ParseString();
            case 't': ExpectWord("true"); return true;
            case 'f': ExpectWord("false"); return false;
            case 'n': ExpectWord("null"); return JsonNull.Value;
        }
        if (c == '-' || (c >= '0' && c <= '9')) { return ParseNumber(); }
        throw Error("unexpected character");
    }

    void ExpectWord(string word)
    {
        if (string.CompareOrdinal(text, pos, word, 0, word.Length) != 0) { throw Error("unexpected word"); }
        pos += word.Length;
    }

    Dictionary<string, object> ParseObject(int depth)
    {
        Dictionary<string, object> result = new Dictionary<string, object>(StringComparer.Ordinal);
        pos++; // {
        SkipSpaces();
        if (pos < text.Length && text[pos] == '}') { pos++; return result; }
        while (true)
        {
            SkipSpaces();
            if (pos >= text.Length || text[pos] != '"') { throw Error("object key expected"); }
            string key = ParseString();
            SkipSpaces();
            if (pos >= text.Length || text[pos] != ':') { throw Error("':' expected"); }
            pos++;
            SkipSpaces();
            object value = ParseValue(depth + 1);
            if (result.ContainsKey(key)) { throw Error("duplicate key \"" + key + "\""); }
            result.Add(key, value);
            SkipSpaces();
            if (pos >= text.Length) { throw Error("unexpected end in object"); }
            if (text[pos] == ',') { pos++; continue; }
            if (text[pos] == '}') { pos++; return result; }
            throw Error("',' or '}' expected");
        }
    }

    List<object> ParseArray(int depth)
    {
        List<object> result = new List<object>();
        pos++; // [
        SkipSpaces();
        if (pos < text.Length && text[pos] == ']') { pos++; return result; }
        while (true)
        {
            SkipSpaces();
            result.Add(ParseValue(depth + 1));
            SkipSpaces();
            if (pos >= text.Length) { throw Error("unexpected end in array"); }
            if (text[pos] == ',') { pos++; continue; }
            if (text[pos] == ']') { pos++; return result; }
            throw Error("',' or ']' expected");
        }
    }

    string ParseString()
    {
        pos++; // "
        StringBuilder sb = null;
        int start = pos;
        while (true)
        {
            if (pos >= text.Length) { throw Error("unterminated string"); }
            char c = text[pos];
            if (c == '"')
            {
                string tail = text.Substring(start, pos - start);
                pos++;
                if (sb == null) { return tail; }
                sb.Append(tail);
                return sb.ToString();
            }
            if (c < ' ') { throw Error("control character in string"); }
            if (c != '\\') { pos++; continue; }
            if (sb == null) { sb = new StringBuilder(); }
            sb.Append(text, start, pos - start);
            pos++; // backslash
            if (pos >= text.Length) { throw Error("unterminated escape"); }
            char e = text[pos++];
            switch (e)
            {
                case '"': sb.Append('"'); break;
                case '\\': sb.Append('\\'); break;
                case '/': sb.Append('/'); break;
                case 'b': sb.Append('\b'); break;
                case 'f': sb.Append('\f'); break;
                case 'n': sb.Append('\n'); break;
                case 'r': sb.Append('\r'); break;
                case 't': sb.Append('\t'); break;
                case 'u':
                    if (pos + 4 > text.Length) { throw Error("short \\u escape"); }
                    int code = 0;
                    for (int i = 0; i < 4; i++)
                    {
                        int digit = HexDigit(text[pos + i]);
                        if (digit < 0) { throw Error("bad \\u escape"); }
                        code = code * 16 + digit;
                    }
                    pos += 4;
                    sb.Append((char)code); // 対になっていないサロゲートもそのまま（JSON の文字列としては正しい）
                    break;
                default: throw Error("bad escape");
            }
            start = pos;
        }
    }

    static int HexDigit(char c)
    {
        if (c >= '0' && c <= '9') { return c - '0'; }
        if (c >= 'a' && c <= 'f') { return c - 'a' + 10; }
        if (c >= 'A' && c <= 'F') { return c - 'A' + 10; }
        return -1;
    }

    double ParseNumber()
    {
        int start = pos;
        if (text[pos] == '-') { pos++; }
        if (pos >= text.Length) { throw Error("bad number"); }
        if (text[pos] == '0') { pos++; }
        else if (text[pos] >= '1' && text[pos] <= '9') { while (pos < text.Length && IsDigit(text[pos])) { pos++; } }
        else { throw Error("bad number"); }
        if (pos < text.Length && text[pos] == '.')
        {
            pos++;
            if (pos >= text.Length || !IsDigit(text[pos])) { throw Error("bad number"); }
            while (pos < text.Length && IsDigit(text[pos])) { pos++; }
        }
        if (pos < text.Length && (text[pos] == 'e' || text[pos] == 'E'))
        {
            pos++;
            if (pos < text.Length && (text[pos] == '+' || text[pos] == '-')) { pos++; }
            if (pos >= text.Length || !IsDigit(text[pos])) { throw Error("bad number"); }
            while (pos < text.Length && IsDigit(text[pos])) { pos++; }
        }
        return double.Parse(text.Substring(start, pos - start), NumberStyles.Float, CultureInfo.InvariantCulture);
    }

    static bool IsDigit(char c) { return c >= '0' && c <= '9'; }
}

internal static class JsonText
{
    /// <summary>
    /// JSON の文字列リテラル。" \ と U+0000〜U+001F をエスケープし、サロゲート（対になっていないものを含む）は \uXXXX にする
    /// （UTF-8 に直すときに U+FFFD へ化けないように）。それ以外の文字はそのまま。
    /// </summary>
    public static string Quote(string value)
    {
        StringBuilder sb = new StringBuilder(value.Length + 2);
        sb.Append('"');
        foreach (char c in value)
        {
            switch (c)
            {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (c < ' ' || char.IsSurrogate(c)) { sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture)); }
                    else { sb.Append(c); }
                    break;
            }
        }
        sb.Append('"');
        return sb.ToString();
    }
}
