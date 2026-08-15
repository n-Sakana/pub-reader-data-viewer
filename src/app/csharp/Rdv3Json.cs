// ============================================================================
// Rdv3Json.cs -- a small JSON reader for the settings file.
//
// The distribution is two files: the .cmd and ReaderDataViewer.json. The .cmd
// carries no third-party assembly and compiles under the in-box csc, so the
// reader is here rather than pulled from a package. It is deliberately small:
// the settings file is written by a person and read once at start-up.
//
// Accepted beyond strict JSON, because a settings file is edited by hand:
//   // line comments and  /* block comments */
//   a trailing comma before } or ]
//
// Numbers are read as double and handed out as int/double/bool/string by the
// caller, which also decides what an out-of-range value means. Errors carry the
// line number, because a settings file that silently does nothing is worse than
// one that says what is wrong with it.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public sealed class Rdv3JsonError : Exception
{
    public Rdv3JsonError(string msg, int line)
        : base(msg + " (line " + line.ToString(CultureInfo.InvariantCulture) + ")")
    {
    }
}

public sealed class Rdv3Json
{
    public const int TObject = 0;
    public const int TArray = 1;
    public const int TString = 2;
    public const int TNumber = 3;
    public const int TBool = 4;
    public const int TNull = 5;

    public int Kind;
    public string Str;
    public double Num;
    public bool Flag;
    public List<Rdv3Json> Items;
    public Dictionary<string, Rdv3Json> Members;

    private Rdv3Json(int kind) { Kind = kind; }

    // ---- the public shape a caller reads ----------------------------------
    public bool IsObject { get { return Kind == TObject; } }
    public bool IsArray { get { return Kind == TArray; } }

    public Rdv3Json Member(string name)
    {
        Rdv3Json v;
        if (Kind == TObject && Members != null && Members.TryGetValue(name, out v)) { return v; }
        return null;
    }

    public int Count { get { return (Kind == TArray && Items != null) ? Items.Count : 0; } }

    public Rdv3Json At(int i)
    {
        return (Kind == TArray && Items != null && i >= 0 && i < Items.Count) ? Items[i] : null;
    }

    // Each getter takes the value to use when the member is absent or of the
    // wrong kind, so a partly broken file still starts the app.
    public string GetString(string name, string fallback)
    {
        Rdv3Json v = Member(name);
        return (v != null && v.Kind == TString) ? v.Str : fallback;
    }

    public int GetInt(string name, int fallback)
    {
        Rdv3Json v = Member(name);
        if (v == null || v.Kind != TNumber) { return fallback; }
        double d = v.Num;
        if (d < int.MinValue || d > int.MaxValue) { return fallback; }
        return (int)Math.Round(d);
    }

    public double GetDouble(string name, double fallback)
    {
        Rdv3Json v = Member(name);
        return (v != null && v.Kind == TNumber) ? v.Num : fallback;
    }

    public bool GetBool(string name, bool fallback)
    {
        Rdv3Json v = Member(name);
        return (v != null && v.Kind == TBool) ? v.Flag : fallback;
    }

    public string[] GetStrings(string name)
    {
        Rdv3Json v = Member(name);
        if (v == null || v.Kind != TArray) { return new string[0]; }
        List<string> outp = new List<string>();
        for (int i = 0; i < v.Count; i++)
        {
            Rdv3Json e = v.At(i);
            if (e != null && e.Kind == TString) { outp.Add(e.Str); }
        }
        return outp.ToArray();
    }

    // ---- reading -----------------------------------------------------------
    private string src;
    private int pos;
    private int line;

    public static Rdv3Json Parse(string text)
    {
        Rdv3Json p = new Rdv3Json(TNull);
        p.src = (text == null) ? "" : text;
        p.pos = 0;
        p.line = 1;
        if (p.src.Length > 0 && p.src[0] == '﻿') { p.pos = 1; }
        Rdv3Json v = p.ReadValue();
        p.SkipWhite();
        if (p.pos < p.src.Length) { throw new Rdv3JsonError("trailing text after the top level value", p.line); }
        return v;
    }

    private void SkipWhite()
    {
        while (pos < src.Length)
        {
            char c = src[pos];
            if (c == '\n') { line++; pos++; }
            else if (c == ' ' || c == '\t' || c == '\r') { pos++; }
            else if (c == '/' && pos + 1 < src.Length && src[pos + 1] == '/')
            {
                while (pos < src.Length && src[pos] != '\n') { pos++; }
            }
            else if (c == '/' && pos + 1 < src.Length && src[pos + 1] == '*')
            {
                pos += 2;
                while (pos + 1 < src.Length && !(src[pos] == '*' && src[pos + 1] == '/'))
                {
                    if (src[pos] == '\n') { line++; }
                    pos++;
                }
                pos = Math.Min(src.Length, pos + 2);
            }
            else { return; }
        }
    }

    private char Peek()
    {
        SkipWhite();
        if (pos >= src.Length) { throw new Rdv3JsonError("the file ends too early", line); }
        return src[pos];
    }

    private Rdv3Json ReadValue()
    {
        char c = Peek();
        if (c == '{') { return ReadObject(); }
        if (c == '[') { return ReadArray(); }
        if (c == '"') { Rdv3Json s = new Rdv3Json(TString); s.Str = ReadString(); return s; }
        if (c == '-' || (c >= '0' && c <= '9')) { return ReadNumber(); }
        if (Word("true")) { Rdv3Json b = new Rdv3Json(TBool); b.Flag = true; return b; }
        if (Word("false")) { Rdv3Json b = new Rdv3Json(TBool); b.Flag = false; return b; }
        if (Word("null")) { return new Rdv3Json(TNull); }
        throw new Rdv3JsonError("expected a value", line);
    }

    private bool Word(string w)
    {
        if (pos + w.Length > src.Length) { return false; }
        if (string.CompareOrdinal(src, pos, w, 0, w.Length) != 0) { return false; }
        pos += w.Length;
        return true;
    }

    private Rdv3Json ReadObject()
    {
        Rdv3Json o = new Rdv3Json(TObject);
        o.Members = new Dictionary<string, Rdv3Json>(StringComparer.OrdinalIgnoreCase);
        pos++;                                   // {
        while (true)
        {
            char c = Peek();
            if (c == '}') { pos++; return o; }
            if (c != '"') { throw new Rdv3JsonError("expected a member name in quotes", line); }
            string name = ReadString();
            if (Peek() != ':') { throw new Rdv3JsonError("expected : after the member name", line); }
            pos++;
            o.Members[name] = ReadValue();
            c = Peek();
            if (c == ',') { pos++; continue; }
            if (c == '}') { pos++; return o; }
            throw new Rdv3JsonError("expected , or } in an object", line);
        }
    }

    private Rdv3Json ReadArray()
    {
        Rdv3Json a = new Rdv3Json(TArray);
        a.Items = new List<Rdv3Json>();
        pos++;                                   // [
        while (true)
        {
            char c = Peek();
            if (c == ']') { pos++; return a; }
            a.Items.Add(ReadValue());
            c = Peek();
            if (c == ',') { pos++; continue; }
            if (c == ']') { pos++; return a; }
            throw new Rdv3JsonError("expected , or ] in an array", line);
        }
    }

    private string ReadString()
    {
        pos++;                                   // opening quote
        StringBuilder sb = new StringBuilder();
        while (true)
        {
            if (pos >= src.Length) { throw new Rdv3JsonError("the text is not closed", line); }
            char c = src[pos++];
            if (c == '"') { return sb.ToString(); }
            if (c == '\n') { throw new Rdv3JsonError("a line break inside text", line); }
            if (c != '\\') { sb.Append(c); continue; }
            if (pos >= src.Length) { throw new Rdv3JsonError("the text is not closed", line); }
            char e = src[pos++];
            if (e == 'n') { sb.Append('\n'); }
            else if (e == 't') { sb.Append('\t'); }
            else if (e == 'r') { sb.Append('\r'); }
            else if (e == 'b') { sb.Append('\b'); }
            else if (e == 'f') { sb.Append('\f'); }
            else if (e == '/' || e == '\\' || e == '"') { sb.Append(e); }
            else if (e == 'u')
            {
                if (pos + 4 > src.Length) { throw new Rdv3JsonError("a short \\u escape", line); }
                int code = int.Parse(src.Substring(pos, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                sb.Append((char)code);
                pos += 4;
            }
            else { throw new Rdv3JsonError("an escape that is not understood", line); }
        }
    }

    private Rdv3Json ReadNumber()
    {
        int start = pos;
        if (pos < src.Length && src[pos] == '-') { pos++; }
        while (pos < src.Length)
        {
            char c = src[pos];
            if ((c >= '0' && c <= '9') || c == '.' || c == 'e' || c == 'E' || c == '+' || c == '-') { pos++; }
            else { break; }
        }
        double d;
        if (!double.TryParse(src.Substring(start, pos - start), NumberStyles.Float,
                CultureInfo.InvariantCulture, out d))
        {
            throw new Rdv3JsonError("a number that is not understood", line);
        }
        Rdv3Json n = new Rdv3Json(TNumber);
        n.Num = d;
        return n;
    }
}
