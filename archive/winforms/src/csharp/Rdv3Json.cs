// ============================================================================
// Rdv3Json.cs -- a small, STRICT JSON reader for settings.json.
//
// The .cmd carries no third-party assembly and compiles under the in-box csc,
// so the reader is here rather than pulled from a package. It is small on
// purpose: the file is written by a person and read once at start-up.
//
// Accepted beyond strict JSON, because the file is edited by hand:
//   // line comments and  /* block comments */
//   a trailing comma before } or ]
//
// Everything else is an error that names the line. The typed accessors below
// are the only way the loaders read a member, and every one of them throws
// Rdv3LoadError on a member of the wrong kind, a value out of range, or a
// required member that is missing; Only() throws on a member nobody asked for
// (a typo would otherwise be a setting that silently does nothing). The app
// does not start on any of these: a file that half-applies is worse than one
// that says what is wrong with it.
//
// Every node remembers its line, its path ("watch.targets[0].window") and the
// span of its text in the source, so the settings dialog can rewrite just the
// members it owns and leave every other byte -- comments included -- as it is.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

// one reason the file cannot be used, with the line it was found on
public sealed class Rdv3LoadError : Exception
{
    public readonly int Line;

    public Rdv3LoadError(string msg, int line)
        : base(msg + ((line > 0) ? " (line " + line.ToString(CultureInfo.InvariantCulture) + ")" : ""))
    {
        Line = line;
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
    public List<string> Order;               // member names as written

    public int Line;                         // where the value starts
    public string Path = "";                 // "watch.targets[0].window"
    public int Start, End;                   // the value's text in the source
    public int KeyStart = -1;                // the member name's opening quote

    private Rdv3Json(int kind) { Kind = kind; }

    // ---- shape --------------------------------------------------------------
    public bool IsObject { get { return Kind == TObject; } }
    public bool IsArray { get { return Kind == TArray; } }
    public int Count { get { return (Kind == TArray && Items != null) ? Items.Count : 0; } }

    public Rdv3Json At(int i)
    {
        return (Kind == TArray && Items != null && i >= 0 && i < Items.Count) ? Items[i] : null;
    }

    public Rdv3Json Member(string name)
    {
        Rdv3Json v;
        if (Kind == TObject && Members != null && Members.TryGetValue(name, out v)) { return v; }
        return null;
    }

    public bool Has(string name) { return Member(name) != null; }

    private string Sub(string name) { return (Path.Length == 0) ? name : Path + "." + name; }

    public Rdv3LoadError Fail(string msg) { return new Rdv3LoadError(((Path.Length == 0) ? "" : Path + ": ") + msg, Line); }

    private Rdv3LoadError FailAt(string name, string msg)
    {
        Rdv3Json v = Member(name);
        return new Rdv3LoadError(Sub(name) + ": " + msg, (v == null) ? Line : v.Line);
    }

    private static string KindName(int k)
    {
        switch (k)
        {
            case TObject: return "an object";
            case TArray: return "an array";
            case TString: return "text";
            case TNumber: return "a number";
            case TBool: return "true/false";
        }
        return "null";
    }

    // ---- strict accessors -------------------------------------------------------
    // every member of this object must be one of the given names
    public void Only(params string[] names)
    {
        if (Kind != TObject) { throw Fail("must be an object"); }
        for (int i = 0; i < Order.Count; i++)
        {
            bool ok = false;
            for (int k = 0; k < names.Length; k++) { if (names[k] == Order[i]) { ok = true; break; } }
            if (!ok) { throw FailAt(Order[i], "is not a member this program knows (" + string.Join(", ", names) + ")"); }
        }
    }

    private Rdv3Json Want(string name, int kind, bool required)
    {
        if (Kind != TObject) { throw Fail("must be an object"); }
        Rdv3Json v = Member(name);
        if (v == null)
        {
            if (required) { throw FailAt(name, "is required"); }
            return null;
        }
        if (v.Kind != kind) { throw FailAt(name, "must be " + KindName(kind) + ", not " + KindName(v.Kind)); }
        return v;
    }

    public Rdv3Json Obj(string name, bool required) { return Want(name, TObject, required); }
    public Rdv3Json Arr(string name, bool required) { return Want(name, TArray, required); }

    // text that must be there and must not be blank
    public string Need(string name)
    {
        Rdv3Json v = Want(name, TString, true);
        if (v.Str.Trim().Length == 0) { throw FailAt(name, "must not be blank"); }
        return v.Str;
    }

    // text that may be absent (then def), but if present must be text
    public string StrOr(string name, string def)
    {
        Rdv3Json v = Want(name, TString, false);
        return (v == null) ? def : v.Str;
    }

    // one of a fixed list of words (case-sensitive, as documented)
    public string Word(string name, string def, params string[] allowed)
    {
        string s = StrOr(name, def);
        for (int i = 0; i < allowed.Length; i++) { if (allowed[i] == s) { return s; } }
        throw FailAt(name, "must be one of " + string.Join(" / ", allowed) + ", not " + s);
    }

    public int Int(string name, int lo, int hi) { return IntOf(Want(name, TNumber, true), name, lo, hi); }

    public int IntOr(string name, int def, int lo, int hi)
    {
        Rdv3Json v = Want(name, TNumber, false);
        return (v == null) ? def : IntOf(v, name, lo, hi);
    }

    private int IntOf(Rdv3Json v, string name, int lo, int hi)
    {
        double d = v.Num;
        if (d != Math.Floor(d)) { throw FailAt(name, "must be a whole number"); }
        if (d < lo || d > hi)
        {
            throw FailAt(name, d.ToString(CultureInfo.InvariantCulture) + " is out of range "
                + lo.ToString(CultureInfo.InvariantCulture) + ".." + hi.ToString(CultureInfo.InvariantCulture));
        }
        return (int)d;
    }

    public double Dbl(string name, double lo, double hi) { return DblOf(Want(name, TNumber, true), name, lo, hi); }

    public double DblOr(string name, double def, double lo, double hi)
    {
        Rdv3Json v = Want(name, TNumber, false);
        return (v == null) ? def : DblOf(v, name, lo, hi);
    }

    private double DblOf(Rdv3Json v, string name, double lo, double hi)
    {
        if (v.Num < lo || v.Num > hi)
        {
            throw FailAt(name, v.Num.ToString(CultureInfo.InvariantCulture) + " is out of range "
                + lo.ToString(CultureInfo.InvariantCulture) + ".." + hi.ToString(CultureInfo.InvariantCulture));
        }
        return v.Num;
    }

    public bool Bool(string name) { return Want(name, TBool, true).Flag; }

    public bool BoolOr(string name, bool def)
    {
        Rdv3Json v = Want(name, TBool, false);
        return (v == null) ? def : v.Flag;
    }

    // an array of text; absent = empty, present = every item must be text
    public string[] Strs(string name, bool required)
    {
        Rdv3Json v = Want(name, TArray, required);
        if (v == null) { return new string[0]; }
        string[] outp = new string[v.Count];
        for (int i = 0; i < v.Count; i++)
        {
            Rdv3Json e = v.At(i);
            if (e == null || e.Kind != TString) { throw e.Fail("must be text"); }
            outp[i] = e.Str;
        }
        return outp;
    }

    // an array whose items must all be objects
    public List<Rdv3Json> Objs(string name, bool required)
    {
        Rdv3Json v = Want(name, TArray, required);
        List<Rdv3Json> list = new List<Rdv3Json>();
        if (v == null) { return list; }
        for (int i = 0; i < v.Count; i++)
        {
            Rdv3Json e = v.At(i);
            if (e.Kind != TObject) { throw e.Fail("must be an object"); }
            list.Add(e);
        }
        return list;
    }

    // a number, or 1..4 numbers, CSS-box style: [all] | [v, h] | [t, h, b] | [t, r, b, l]
    public double[] Box(string name)
    {
        Rdv3Json a = Member(name);
        if (a == null) { return null; }
        if (a.Kind == TNumber) { return new double[] { a.Num, a.Num, a.Num, a.Num }; }
        if (a.Kind != TArray || a.Count == 0 || a.Count > 4) { throw FailAt(name, "must be a number or 1..4 numbers"); }
        double[] v = new double[a.Count];
        for (int i = 0; i < a.Count; i++)
        {
            Rdv3Json n = a.At(i);
            if (n.Kind != TNumber) { throw n.Fail("must be a number"); }
            v[i] = n.Num;
        }
        if (v.Length == 1) { return new double[] { v[0], v[0], v[0], v[0] }; }
        if (v.Length == 2) { return new double[] { v[0], v[1], v[0], v[1] }; }
        if (v.Length == 3) { return new double[] { v[0], v[1], v[2], v[1] }; }
        return v;
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
        if (p.src.Length > 0 && p.src[0] == '\uFEFF') { p.pos = 1; }
        Rdv3Json v = p.ReadValue("");
        p.SkipWhite();
        if (p.pos < p.src.Length) { throw new Rdv3LoadError("trailing text after the top level value", p.line); }
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
        if (pos >= src.Length) { throw new Rdv3LoadError("the file ends too early", line); }
        return src[pos];
    }

    private Rdv3Json ReadValue(string path)
    {
        char c = Peek();
        int at = pos;
        int ln = line;
        Rdv3Json v;
        if (c == '{') { v = ReadObject(path); }
        else if (c == '[') { v = ReadArray(path); }
        else if (c == '"') { v = new Rdv3Json(TString); v.Str = ReadString(); }
        else if (c == '-' || (c >= '0' && c <= '9')) { v = ReadNumber(); }
        else if (Word("true")) { v = new Rdv3Json(TBool); v.Flag = true; }
        else if (Word("false")) { v = new Rdv3Json(TBool); v.Flag = false; }
        else if (Word("null")) { v = new Rdv3Json(TNull); }
        else { throw new Rdv3LoadError("expected a value", line); }
        v.Path = path;
        v.Line = ln;
        v.Start = at;
        v.End = pos;
        return v;
    }

    private bool Word(string w)
    {
        if (pos + w.Length > src.Length) { return false; }
        if (string.CompareOrdinal(src, pos, w, 0, w.Length) != 0) { return false; }
        pos += w.Length;
        return true;
    }

    private Rdv3Json ReadObject(string path)
    {
        Rdv3Json o = new Rdv3Json(TObject);
        o.Members = new Dictionary<string, Rdv3Json>(StringComparer.Ordinal);
        o.Order = new List<string>();
        pos++;                                   // {
        while (true)
        {
            char c = Peek();
            if (c == '}') { pos++; return o; }
            if (c != '"') { throw new Rdv3LoadError("expected a member name in quotes", line); }
            int keyAt = pos;
            string name = ReadString();
            if (Peek() != ':') { throw new Rdv3LoadError("expected : after the member name " + name, line); }
            pos++;
            if (o.Members.ContainsKey(name)) { throw new Rdv3LoadError("the member " + name + " is written twice", line); }
            Rdv3Json v = ReadValue((path.Length == 0) ? name : path + "." + name);
            v.KeyStart = keyAt;
            o.Members[name] = v;
            o.Order.Add(name);
            c = Peek();
            if (c == ',') { pos++; continue; }
            if (c == '}') { pos++; return o; }
            throw new Rdv3LoadError("expected , or } in an object", line);
        }
    }

    private Rdv3Json ReadArray(string path)
    {
        Rdv3Json a = new Rdv3Json(TArray);
        a.Items = new List<Rdv3Json>();
        pos++;                                   // [
        while (true)
        {
            char c = Peek();
            if (c == ']') { pos++; return a; }
            a.Items.Add(ReadValue(path + "[" + a.Items.Count.ToString(CultureInfo.InvariantCulture) + "]"));
            c = Peek();
            if (c == ',') { pos++; continue; }
            if (c == ']') { pos++; return a; }
            throw new Rdv3LoadError("expected , or ] in an array", line);
        }
    }

    private string ReadString()
    {
        pos++;                                   // opening quote
        StringBuilder sb = new StringBuilder();
        while (true)
        {
            if (pos >= src.Length) { throw new Rdv3LoadError("the text is not closed", line); }
            char c = src[pos++];
            if (c == '"') { return sb.ToString(); }
            if (c == '\n') { throw new Rdv3LoadError("a line break inside text", line); }
            if (c != '\\') { sb.Append(c); continue; }
            if (pos >= src.Length) { throw new Rdv3LoadError("the text is not closed", line); }
            char e = src[pos++];
            if (e == 'n') { sb.Append('\n'); }
            else if (e == 't') { sb.Append('\t'); }
            else if (e == 'r') { sb.Append('\r'); }
            else if (e == 'b') { sb.Append('\b'); }
            else if (e == 'f') { sb.Append('\f'); }
            else if (e == '/' || e == '\\' || e == '"') { sb.Append(e); }
            else if (e == 'u')
            {
                int code;
                if (pos + 4 > src.Length || !int.TryParse(src.Substring(pos, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                {
                    throw new Rdv3LoadError("a \\u escape needs four hex digits", line);
                }
                sb.Append((char)code);
                pos += 4;
            }
            else { throw new Rdv3LoadError("an escape that is not understood: \\" + e, line); }
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
            throw new Rdv3LoadError("a number that is not understood", line);
        }
        Rdv3Json n = new Rdv3Json(TNumber);
        n.Num = d;
        return n;
    }
}
