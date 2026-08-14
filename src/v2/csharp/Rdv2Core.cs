// ============================================================================
// Rdv2Core.cs -- merge-select over a one-to-many key.
//
// What changed from src\csharp\RdvCore.cs (which is untouched and still builds
// the 1,000,000-row 1:1 distributables):
//
//   key 1 is legitimately one-to-many. B is the middle table -- a slip with
//   several lines - so one key 1 can own several B rows, each with its own
//   key 2. A search by key 1 therefore yields a SET of candidates, and the
//   operator picks one. An index that keeps "the last row for this key" would
//   silently throw the others away, so no index here is allowed to do that.
//
//   reading and indexing are timed separately, because the whole question this
//   build exists to answer is what the index costs.
//
// The index itself is NOT in this file. Two files define a class called
// RdvIndex with the same three members, and the packer puts exactly one of
// them in each .cmd:
//
//   Rdv2IdxHash.cs   hand-built bucket-chained hash over the raw bytes
//   Rdv2IdxDict.cs   Dictionary<string, List<int>>
//
// Same call sites, no interface, no virtual dispatch, so the measurement is of
// the two structures and not of an abstraction wrapped around them.
//
// C# 5 only, ASCII only. See build\pack_cmd.ps1.
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

public static class Rdv2Spec
{
    public const int  KeyLen = 8;
    public const int  Fields = 10;
    public const long Mod    = 1000000007L;
    public const int  HashBits = 21;
    public const int  HashMask = (1 << 21) - 1;
    public const int  MaxCand = 256;      // candidates held for one key

    public const int StReadA  = 0;
    public const int StReadB  = 1;
    public const int StReadC  = 2;
    public const int StIdxA   = 3;
    public const int StIdxB   = 4;
    public const int StIdxC   = 5;
    public const int StJoinAB = 6;
    public const int StJoinBC = 7;
    public const int StSearch = 8;
    public const int StShow   = 9;
    public const int StageCount = 10;

    public static readonly string[] StageKey =
    { "readA", "readB", "readC", "idxA", "idxB", "idxC", "joinAB", "joinBC", "search", "present" };

    public static int Capacity(int rows)
    {
        int cap = 1024;
        while (cap < rows * 2 && cap < (1 << HashBits)) { cap = cap << 1; }
        return cap;
    }

    // the same polynomial hash the 1:1 build uses, and the same one
    // modRdv2IdxHash.cls runs inside Excel
    public static int Hash(byte[] b, int p)
    {
        int h = b[p];
        h = ((h * 131) + b[p + 1]) & HashMask;
        h = ((h * 131) + b[p + 2]) & HashMask;
        h = ((h * 131) + b[p + 3]) & HashMask;
        h = ((h * 131) + b[p + 4]) & HashMask;
        h = ((h * 131) + b[p + 5]) & HashMask;
        h = ((h * 131) + b[p + 6]) & HashMask;
        h = ((h * 131) + b[p + 7]) & HashMask;
        h = (h ^ (h >> 10)) & HashMask;
        return h;
    }

    public static bool SameKey(byte[] a, int pa, byte[] b, int pb)
    {
        return a[pa] == b[pb] && a[pa + 1] == b[pb + 1] && a[pa + 2] == b[pb + 2] &&
               a[pa + 3] == b[pb + 3] && a[pa + 4] == b[pb + 4] && a[pa + 5] == b[pb + 5] &&
               a[pa + 6] == b[pb + 6] && a[pa + 7] == b[pb + 7];
    }
}

public static class Rdv2Clock
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

    public static string FmtSec(double ms)
    {
        return (ms / 1000.0).ToString("N3", CultureInfo.InvariantCulture);
    }
}

// One CSV: the bytes off the disk plus where each row starts and ends.
// Reading stops here. Indexing is a separate, separately timed step.
public sealed class Rdv2Table
{
    public string Name;
    public string Path;
    public byte[] Buf;
    public int    Rows;
    public int[]  Start;
    public int[]  End;
    public int[]  KeyAt;
    public int    HeadStart, HeadEnd;

    public static Rdv2Table Read(string path, string name)
    {
        Rdv2Table t = new Rdv2Table();
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
            int ke = t.KeyAt[i] + Rdv2Spec.KeyLen;
            if (ke > t.End[i] || (ke < t.End[i] && b[ke] != (byte)','))
            {
                throw new InvalidDataException(name + " row " + (i + 1).ToString(CultureInfo.InvariantCulture)
                    + ": key must be " + Rdv2Spec.KeyLen.ToString(CultureInfo.InvariantCulture) + " bytes");
            }
        }
        return t;
    }

    public string Header()
    {
        return Encoding.UTF8.GetString(Buf, HeadStart, HeadEnd - HeadStart);
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

    public string[] SplitRow(int i)
    {
        string[] v = new string[Rdv2Spec.Fields];
        int p = Start[i];
        int e = End[i];
        for (int f = 0; f < Rdv2Spec.Fields; f++)
        {
            int q = p;
            while (q < e && Buf[q] != (byte)',') { q++; }
            v[f] = (q > p) ? Encoding.UTF8.GetString(Buf, p, q - p) : "";
            p = q + 1;
            if (p > e) { p = e; }
        }
        return v;
    }

    public string[] SplitHeader()
    {
        string[] v = new string[Rdv2Spec.Fields];
        string[] parts = Header().Split(',');
        for (int i = 0; i < Rdv2Spec.Fields; i++) { v[i] = (i < parts.Length) ? parts[i] : ""; }
        return v;
    }
}

// One row of the candidate list: which B row it is, what it joins to, and the
// columns the operator picks by. The columns are chosen from the 10-field
// layout and are listed in README.md; changing them here means changing them
// there and in the VBA build too.
public sealed class Rdv2Cand
{
    public int    RowB = -1, RowA = -1, RowC = -1;
    public string Key2 = "";
    public string Line = "";      // b_line
    public string Slip = "";      // b_slip
    public string Date = "";      // b_date
    public string Qty = "";       // b_qty
    public string Status = "";    // b_status
    public string Item = "";      // c_item
    public string Maker = "";     // c_maker

    public string[] Columns()
    {
        return new string[] { Key2, Line, Slip, Date, Qty, Status, Item, Maker };
    }
}

public sealed class Rdv2Hit
{
    public string     Key1 = "";
    public Rdv2Cand[] Cands = new Rdv2Cand[0];
    public int        Chosen = -1;          // index into Cands once one is picked
    public string[]   NameA, NameB, NameC;
    public string[]   ValA, ValB, ValC;

    public int Count { get { return Cands.Length; } }
}

public sealed class Rdv2Run
{
    public int      Seq;
    public string   Key = "";
    public DateTime When;
    public double[] Stage = new double[Rdv2Spec.StageCount];
    public double   TotalMs;
    public double   DetectMs;
    public double   PickMs = -1;     // display after the operator chose; -1 = not applicable
    public int      Polls;
    public int      CandCount;
    public long     Checksum;
    public long     Probes;
    public int      MatchedAB, MatchedBC;
    public int      Rows;
    public bool     OracleOk;
    public string   OracleNote = "";
    public string   IndexKind = "";
    public Rdv2Hit  Hit = new Rdv2Hit();
    public string   Error = "";

    public double StageSum()
    {
        double s = 0;
        for (int i = 0; i < Rdv2Spec.StageCount; i++) { s += Stage[i]; }
        return s;
    }
}

public sealed class Rdv2Expected
{
    public int  Rows;
    public long Checksum = -1;
    public bool Loaded;

    public static Rdv2Expected Read(string dir)
    {
        Rdv2Expected e = new Rdv2Expected();
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

public sealed class Rdv2Engine
{
    private readonly string dir;
    private readonly Rdv2Expected expected;

    // kept only so the operator's pick can be materialised after the clock has
    // stopped; every trigger throws them away and reads the files again
    private Rdv2Table ta, tb, tc;

    public Rdv2Engine(string dataDir)
    {
        dir = dataDir;
        expected = Rdv2Expected.Read(dataDir);
    }

    public Rdv2Expected Expected { get { return expected; } }
    public string DataDir { get { return dir; } }

    public void CheckFiles()
    {
        string[] need = { "tableA.csv", "tableB.csv", "tableC.csv" };
        for (int i = 0; i < need.Length; i++)
        {
            string p = Path.Combine(dir, need[i]);
            if (!File.Exists(p)) { throw new FileNotFoundException(p); }
        }
    }

    public Rdv2Run Execute(string key, long t0)
    {
        Rdv2Run r = new Rdv2Run();
        r.Key = key;
        r.When = DateTime.Now;
        r.IndexKind = RdvIndex.Kind;
        try
        {
            long m = Rdv2Clock.Now();
            ta = Rdv2Table.Read(Path.Combine(dir, "tableA.csv"), "A");
            r.Stage[Rdv2Spec.StReadA] = Rdv2Clock.MsSince(m);

            m = Rdv2Clock.Now();
            tb = Rdv2Table.Read(Path.Combine(dir, "tableB.csv"), "B");
            r.Stage[Rdv2Spec.StReadB] = Rdv2Clock.MsSince(m);

            m = Rdv2Clock.Now();
            tc = Rdv2Table.Read(Path.Combine(dir, "tableC.csv"), "C");
            r.Stage[Rdv2Spec.StReadC] = Rdv2Clock.MsSince(m);

            m = Rdv2Clock.Now();
            RdvIndex ia = new RdvIndex(ta);
            r.Stage[Rdv2Spec.StIdxA] = Rdv2Clock.MsSince(m);

            m = Rdv2Clock.Now();
            RdvIndex ib = new RdvIndex(tb);
            r.Stage[Rdv2Spec.StIdxB] = Rdv2Clock.MsSince(m);

            m = Rdv2Clock.Now();
            RdvIndex ic = new RdvIndex(tc);
            r.Stage[Rdv2Spec.StIdxC] = Rdv2Clock.MsSince(m);

            r.Rows = tb.Rows;

            int[] ab = new int[tb.Rows];
            int[] bc = new int[tb.Rows];
            int[] found = new int[Rdv2Spec.MaxCand];

            // A-B over every row of B. A holds one row per key 1, but nothing
            // here assumes it: the index answers with a set and the count is
            // carried forward.
            m = Rdv2Clock.Now();
            int matched = 0;
            for (int i = 0; i < tb.Rows; i++)
            {
                int c = ia.Find(tb.Buf, tb.KeyAt[i], found);
                ab[i] = (c > 0) ? found[0] : -1;
                if (c > 0) { matched++; }
            }
            r.MatchedAB = matched;
            r.Stage[Rdv2Spec.StJoinAB] = Rdv2Clock.MsSince(m);

            // B-C over every row of B, on key 2 = field 1 of B
            m = Rdv2Clock.Now();
            matched = 0;
            long chk = 0;
            for (int i = 0; i < tb.Rows; i++)
            {
                int len;
                int p = tb.FieldAt(i, 1, out len);
                int rc = -1;
                if (p >= 0 && len == Rdv2Spec.KeyLen)
                {
                    int c = ic.Find(tb.Buf, p, found);
                    if (c > 0) { rc = found[0]; matched++; }
                }
                bc[i] = rc;
                chk = (chk * 31 + ab[i] + rc) % Rdv2Spec.Mod;
            }
            r.MatchedBC = matched;
            r.Checksum = chk;
            r.Stage[Rdv2Spec.StJoinBC] = Rdv2Clock.MsSince(m);
            r.Probes = ia.Probes + ib.Probes + ic.Probes;

            // search: key 1 -> the set of B rows
            m = Rdv2Clock.Now();
            byte[] kb = Encoding.ASCII.GetBytes(key);
            int n = 0;
            if (kb.Length == Rdv2Spec.KeyLen) { n = ib.Find(kb, 0, found); }
            int[] rowsB = new int[n];
            for (int i = 0; i < n; i++) { rowsB[i] = found[i]; }
            Array.Sort(rowsB);          // file order, so the list is stable
            r.CandCount = n;
            r.Stage[Rdv2Spec.StSearch] = Rdv2Clock.MsSince(m);

            Rdv2Hit hit = new Rdv2Hit();
            hit.Key1 = key;
            hit.Cands = new Rdv2Cand[n];
            for (int i = 0; i < n; i++)
            {
                Rdv2Cand cd = new Rdv2Cand();
                cd.RowB = rowsB[i];
                cd.RowA = ab[rowsB[i]];
                cd.RowC = bc[rowsB[i]];
                hit.Cands[i] = cd;
            }
            r.Hit = hit;

            if (expected.Loaded)
            {
                r.OracleOk = true;
                if (expected.Rows != r.Rows) { r.OracleOk = false; r.OracleNote = "rows"; }
                else if (expected.Checksum != r.Checksum) { r.OracleOk = false; r.OracleNote = "checksum"; }
            }
            else { r.OracleOk = true; r.OracleNote = "no expected.txt"; }
        }
        catch (Exception ex)
        {
            r.Error = ex.GetType().Name + ": " + ex.Message;
        }
        return r;
    }

    // stage 10, part of the measured total: the columns the operator picks by.
    // Only the candidates are materialised here, never all 100,000 rows.
    public void FillCandidates(Rdv2Hit hit)
    {
        for (int i = 0; i < hit.Cands.Length; i++)
        {
            Rdv2Cand c = hit.Cands[i];
            c.Key2 = tb.Field(c.RowB, 1);
            c.Slip = tb.Field(c.RowB, 2);
            c.Date = tb.Field(c.RowB, 3);
            c.Qty = tb.Field(c.RowB, 4);
            c.Status = tb.Field(c.RowB, 7);
            c.Line = tb.Field(c.RowB, 8);
            if (c.RowC >= 0)
            {
                c.Item = tc.Field(c.RowC, 1);
                c.Maker = tc.Field(c.RowC, 2);
            }
        }
    }

    // the 30 fields of one chosen candidate
    public void FillRecord(Rdv2Hit hit, int which)
    {
        if (which < 0 || which >= hit.Cands.Length) { return; }
        Rdv2Cand c = hit.Cands[which];
        hit.Chosen = which;
        hit.NameB = tb.SplitHeader();
        hit.ValB = tb.SplitRow(c.RowB);
        if (c.RowA >= 0) { hit.NameA = ta.SplitHeader(); hit.ValA = ta.SplitRow(c.RowA); }
        else { hit.NameA = ta.SplitHeader(); hit.ValA = null; }
        if (c.RowC >= 0) { hit.NameC = tc.SplitHeader(); hit.ValC = tc.SplitRow(c.RowC); }
        else { hit.NameC = tc.SplitHeader(); hit.ValC = null; }
    }
}
