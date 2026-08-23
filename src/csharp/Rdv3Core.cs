// ============================================================================
// Rdv3Core.cs -- clock, CSV tables and the data error of the practical build.
//
// The product build is not a benchmark. Frozen benchmark sources live under
// archive\comparisons; this directory contains only the selected product
// implementation:
//
//   - one index method only: the standard Dictionary<string, List<int>>.
//     There is no hand-built hash in this build and no fallback to one.
//   - the merge runs once, in the background, at startup: its output is the
//     integrated ledger, which is persisted and searched. Searching does NOT
//     re-read the CSVs (the old per-trigger re-read contract belongs to the
//     frozen comparison builds).
//
// The CSVs are RAW data and the tables are described by settings.json
// ("data": which files, which column is the key, how they join -- Rdv3Data).
// A table's column names are its header row; the width of its key is the
// width of the key in its first data row, and every other row must agree.
//
// The reader is STRICT. A row with the wrong number of columns, a field that
// begins with a quote (this reader does not unquote, so a quoted field would
// silently shift every column after it), a key of another width, a key with a
// byte outside ASCII: each is an Rdv3DataError with the file and the row, and
// the app does not run on that data. Nothing is skipped or repaired.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// See build\pack_app.ps1.
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

// the input (a CSV or the saved ledger) cannot be used as it is; the message
// is the operator's, naming the file and the row
public sealed class Rdv3DataError : Exception
{
    public Rdv3DataError(string msg) : base(msg) { }
}

public static class Rdv3Spec
{
    public const long Mod = 1000000007L;
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
public sealed class Rdv3Table
{
    public string Name;                      // the table id of the definition ("A")
    public string Path;
    public byte[] Buf;
    public Encoding Enc;
    public int Rows;
    public int[] Start;
    public int[] End;
    // the header row, one name per column
    public string[] Head;
    // the key column: its index in Head, the byte offset of the key in each
    // row, and its width (taken from the first data row)
    public int KeyCol;
    public int[] KeyAt;
    public int KeyLen;

    private static string Fmt(string text, string file, int row)
    {
        return text.Replace("{file}", file).Replace("{row}", row.ToString(CultureInfo.InvariantCulture));
    }

    // the header row only (for the start-up check of the definition against
    // the data, before the full read in the worker)
    public static string[] ReadHead(string path, Encoding enc)
    {
        string file = System.IO.Path.GetFileName(path);
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (StreamReader r = new StreamReader(fs, enc, true))
        {
            string first = r.ReadLine();
            if (first == null || first.Trim().Length == 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoRows, file, 0)); }
            return SplitHead(first, file);
        }
    }

    private static string[] SplitHead(string line, string file)
    {
        if (line.Length > 0 && line[0] == '\uFEFF') { line = line.Substring(1); }
        string[] h = line.Split(',');
        for (int i = 0; i < h.Length; i++)
        {
            h[i] = h[i].Trim();
            if (h[i].Length == 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataBlankHeader, file, 1)); }
            if (h[i][0] == '"') { throw new Rdv3DataError(Fmt(Rdv3Text.DataQuoted, file, 1)); }
            for (int k = 0; k < i; k++)
            {
                if (string.Equals(h[k], h[i], StringComparison.Ordinal))
                {
                    throw new Rdv3DataError(Fmt(Rdv3Text.DataDupHeader, file, 1).Replace("{name}", h[i]));
                }
            }
        }
        return h;
    }

    // the whole table, validated row by row
    public static Rdv3Table Read(string path, string name, Encoding enc, string keyName)
    {
        Rdv3Table t = new Rdv3Table();
        t.Name = name;
        t.Path = path;
        t.Enc = enc;
        t.Buf = File.ReadAllBytes(path);
        string file = System.IO.Path.GetFileName(path);

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
        if (rows < 2) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoRows, file, 0)); }

        t.Head = SplitHead(enc.GetString(b, st[0], en[0] - st[0]), file);
        int cols = t.Head.Length;
        t.KeyCol = -1;
        for (int i = 0; i < cols; i++) { if (t.Head[i] == keyName) { t.KeyCol = i; } }
        if (t.KeyCol < 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoColumn, file, 1).Replace("{name}", keyName)); }

        t.Rows = rows - 1;
        t.Start = new int[t.Rows];
        t.End = new int[t.Rows];
        t.KeyAt = new int[t.Rows];
        t.KeyLen = -1;
        for (int i = 0; i < t.Rows; i++)
        {
            int rs = st[i + 1];
            int re = en[i + 1];
            t.Start[i] = rs;
            t.End[i] = re;
            int row = i + 2;                   // the row number a person sees (header = 1)
            // one pass over the row: count the columns, refuse a quoted field,
            // and find the key
            int field = 0;
            int p = rs;
            int keyAt = -1;
            int keyEnd = -1;
            while (true)
            {
                if (p < re && b[p] == (byte)'"') { throw new Rdv3DataError(Fmt(Rdv3Text.DataQuoted, file, row)); }
                int q = p;
                while (q < re && b[q] != (byte)',') { q++; }
                if (field == t.KeyCol) { keyAt = p; keyEnd = q; }
                field++;
                if (q >= re) { break; }
                p = q + 1;
            }
            if (field != cols)
            {
                throw new Rdv3DataError(Fmt(Rdv3Text.DataColumnCount, file, row)
                    .Replace("{n}", field.ToString(CultureInfo.InvariantCulture))
                    .Replace("{cols}", cols.ToString(CultureInfo.InvariantCulture)));
            }
            int klen = keyEnd - keyAt;
            if (klen <= 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataEmptyKey, file, row).Replace("{name}", keyName)); }
            for (int k = keyAt; k < keyEnd; k++)
            {
                if (b[k] > 127) { throw new Rdv3DataError(Fmt(Rdv3Text.DataKeyNotAscii, file, row).Replace("{name}", keyName)); }
            }
            if (t.KeyLen < 0) { t.KeyLen = klen; }
            else if (klen != t.KeyLen)
            {
                throw new Rdv3DataError(Fmt(Rdv3Text.DataKeyWidth, file, row)
                    .Replace("{name}", keyName)
                    .Replace("{n}", t.KeyLen.ToString(CultureInfo.InvariantCulture)));
            }
            t.KeyAt[i] = keyAt;
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
        return Enc.GetString(Buf, p, len);
    }

    public int ColumnOf(string name)
    {
        for (int i = 0; i < Head.Length; i++) { if (Head[i] == name) { return i; } }
        return -1;
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
