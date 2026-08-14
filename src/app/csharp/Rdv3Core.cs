// ============================================================================
// Rdv3Core.cs -- constants, clock and CSV tables for the practical build.
//
// The practical build ("app") is NOT the benchmark. The benchmark sources
// (src\csharp, src\v2\csharp) are frozen; this directory may copy from them
// but never changes them. What is different here:
//
//   - one index method only: the standard Dictionary<string, List<int>>.
//     There is no hand-built hash in this build and no fallback to one.
//   - the merge runs once, in the background, at startup: its output is the
//     integrated ledger, which is persisted and searched. Searching does NOT
//     re-read the CSVs (the old per-trigger re-read contract belongs to the
//     frozen comparison builds).
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// See build\pack_app.ps1.
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

public static class Rdv3Spec
{
    public const int KeyLen = 8;
    public const int Fields = 10;
    public const long Mod = 1000000007L;

    // merge stages: the eight regions whose sum is the merge time on screen
    public const int StReadA = 0;
    public const int StReadB = 1;
    public const int StReadC = 2;
    public const int StIdxA = 3;
    public const int StIdxB = 4;
    public const int StIdxC = 5;
    public const int StJoinAB = 6;
    public const int StJoinBC = 7;
    public const int StageCount = 8;

    public static readonly string[] StageKey =
    { "readA", "readB", "readC", "idxA", "idxB", "idxC", "joinAB", "joinBC" };

    public static bool IsKey(string s)
    {
        if (s == null || s.Length != KeyLen) { return false; }
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] < '0' || s[i] > '9') { return false; }
        }
        return true;
    }
}

public static class Rdv3Clock
{
    private static readonly double TickMs = 1000.0 / (double)Stopwatch.Frequency;

    public static long Now()
    {
        return Stopwatch.GetTimestamp();
    }

    public static double MsSince(long t0)
    {
        return (double)(Stopwatch.GetTimestamp() - t0) * TickMs;
    }

    public static double MsBetween(long t0, long t1)
    {
        return (double)(t1 - t0) * TickMs;
    }

    public static string Fmt(double ms)
    {
        return ms.ToString("N1", CultureInfo.InvariantCulture);
    }
}

// One CSV: the bytes off the disk plus where each row starts and ends.
// Same reading method as the frozen builds (measured there; nothing new).
public sealed class Rdv3Table
{
    public string Name;
    public string Path;
    public byte[] Buf;
    public int Rows;
    public int[] Start;
    public int[] End;
    public int[] KeyAt;
    public int HeadStart, HeadEnd;

    public static Rdv3Table Read(string path, string name)
    {
        Rdv3Table t = new Rdv3Table();
        t.Name = name;
        t.Path = path;
        t.Buf = File.ReadAllBytes(path);

        byte[] b = t.Buf;
        int n = b.Length;
        int cap = 65536;
        int[] st = new int[cap];
        int[] en = new int[cap];
        int rows = 0;
        int pos = 0;
        if (n >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) { pos = 3; }

        while (pos < n)
        {
            int nl = pos;
            while (nl < n && b[nl] != (byte)'\n') { nl++; }
            int e = nl;
            if (e > pos && b[e - 1] == (byte)'\r') { e--; }
            if (e > pos)
            {
                if (rows == cap)
                {
                    cap = cap << 1;
                    Array.Resize(ref st, cap);
                    Array.Resize(ref en, cap);
                }
                st[rows] = pos;
                en[rows] = e;
                rows++;
            }
            pos = nl + 1;
        }
        if (rows < 2) { throw new InvalidDataException(path); }

        t.HeadStart = st[0];
        t.HeadEnd = en[0];
        t.Rows = rows - 1;
        t.Start = new int[t.Rows];
        t.End = new int[t.Rows];
        t.KeyAt = new int[t.Rows];
        for (int i = 0; i < t.Rows; i++)
        {
            t.Start[i] = st[i + 1];
            t.End[i] = en[i + 1];
            t.KeyAt[i] = st[i + 1];
            int ke = t.KeyAt[i] + Rdv3Spec.KeyLen;
            if (ke > t.End[i] || (ke < t.End[i] && b[ke] != (byte)','))
            {
                throw new InvalidDataException(name + " row " + (i + 1).ToString(CultureInfo.InvariantCulture)
                    + ": key must be " + Rdv3Spec.KeyLen.ToString(CultureInfo.InvariantCulture) + " bytes");
            }
        }
        return t;
    }

    // byte offset of field f in row i, -1 if the row has no such field
    public int FieldAt(int i, int f, out int len)
    {
        int p = Start[i];
        int e = End[i];
        for (int k = 0; k < f; k++)
        {
            while (p < e && Buf[p] != (byte)',') { p++; }
            if (p >= e) { len = 0; return -1; }
            p++;
        }
        int q = p;
        while (q < e && Buf[q] != (byte)',') { q++; }
        len = q - p;
        return p;
    }

    public string Field(int i, int f)
    {
        int len;
        int p = FieldAt(i, f, out len);
        if (p < 0 || len <= 0) { return ""; }
        return Encoding.UTF8.GetString(Buf, p, len);
    }
}

// expected.txt is the generator's independent oracle. The practical build keeps
// the same verification the comparison builds ran (rows + join checksum); it is
// logged, and a mismatch is surfaced as an error. Nothing new is checked.
public sealed class Rdv3Expected
{
    public int Rows;
    public long Checksum = -1;
    public bool Loaded;

    public static Rdv3Expected Read(string dir)
    {
        Rdv3Expected e = new Rdv3Expected();
        string p = Path.Combine(dir, "expected.txt");
        if (!File.Exists(p)) { return e; }
        string[] lines = File.ReadAllLines(p);
        for (int i = 0; i < lines.Length; i++)
        {
            int eq = lines[i].IndexOf('=');
            if (eq <= 0) { continue; }
            string k = lines[i].Substring(0, eq);
            string v = lines[i].Substring(eq + 1);
            if (k == "rows") { int.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out e.Rows); }
            else if (k == "joinchecksum") { long.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out e.Checksum); }
        }
        e.Loaded = (e.Rows > 0);
        return e;
    }
}
