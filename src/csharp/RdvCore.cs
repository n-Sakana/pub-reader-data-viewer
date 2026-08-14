// ============================================================================
// RdvCore.cs -- the merge-select engine. Shared verbatim by both .cmd builds.
//
// One merge-select is: three CSVs read from disk, two 1:1 joins over all
// 1,000,000 rows, one lookup, one screen. Nothing is cached between triggers.
// Every trigger repeats the whole thing, which is the point of the exercise.
//
// The algorithm here is deliberately the same one modRdvEngine.bas runs inside
// Excel: raw file bytes, a line-offset scan, an open-addressing hash table with
// the polynomial hash in RdvHash, and the same probe order. Same rule on both
// sides, so a difference in the numbers is a difference between the METHODS and
// not between two people's idea of how to write a join.
//
// Cross-checks that make that claim testable, printed on every run:
//   probes   -- total linear-probe steps while building the three tables
//   checksum -- folded over both joins, in B's file order
// Both must match data\expected.txt and must match between VBA and C#.
//
// C# 5 only: this is compiled by Add-Type through the in-box .NET Framework
// csc, so no interpolated strings, no ?., no expression-bodied members.
// ASCII only: the packer turns every non-ASCII character into \uXXXX, and a
// verbatim string would swallow the escape. Keep UI text out of this file.
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

public static class RdvSpec
{
    public const int  KeyLen = 8;           // "00000001" -- leading zeros are data
    public const int  Fields = 10;
    public const long Mod    = 1000000007L;
    public const int  HashBits = 21;        // the hash yields 21 bits, so the
    public const int  HashMask = (1 << 21) - 1;  // table tops out at 2,097,152

    public const int StageLoadA  = 0;
    public const int StageLoadB  = 1;
    public const int StageLoadC  = 2;
    public const int StageJoinAB = 3;
    public const int StageJoinBC = 4;
    public const int StageSelect = 5;
    public const int StageShow   = 6;
    public const int StageCount  = 7;

    public static readonly string[] StageKey =
        { "loadA", "loadB", "loadC", "joinAB", "joinBC", "select", "display" };

    // table capacity: the smallest power of two that is at least twice the row
    // count, never more than the 21 bits the hash produces. modRdvEngine.bas
    // computes this the same way; the probe counts only match if it does.
    public static int Capacity(int rows)
    {
        int cap = 1024;
        while (cap < rows * 2 && cap < (1 << HashBits)) { cap = cap << 1; }
        return cap;
    }
}

public static class RdvHash
{
    // Polynomial hash over the 8 key bytes, masked to 21 bits at every step so
    // VBA can run the identical arithmetic without a Long overflow: the largest
    // intermediate is 2097151 * 131 + 255, which still fits in a signed 32-bit
    // integer. The final xor-shift folds the high bits back down.
    public static int Of(byte[] b, int p)
    {
        int h = b[p];
        h = ((h * 131) + b[p + 1]) & RdvSpec.HashMask;
        h = ((h * 131) + b[p + 2]) & RdvSpec.HashMask;
        h = ((h * 131) + b[p + 3]) & RdvSpec.HashMask;
        h = ((h * 131) + b[p + 4]) & RdvSpec.HashMask;
        h = ((h * 131) + b[p + 5]) & RdvSpec.HashMask;
        h = ((h * 131) + b[p + 6]) & RdvSpec.HashMask;
        h = ((h * 131) + b[p + 7]) & RdvSpec.HashMask;
        h = (h ^ (h >> 10)) & RdvSpec.HashMask;
        return h;
    }
}

// One CSV, held as the bytes that came off the disk plus an index over them.
// No per-row strings: a row is a pair of offsets into Buf, and a key is eight
// bytes at a known offset. Only the record that gets displayed becomes text.
public sealed class RdvTable
{
    public string Name;
    public string Path;
    public byte[] Buf;
    public int    Rows;
    public int[]  Start;         // byte offset of data row i
    public int[]  End;           // one past its last byte, CR/LF excluded
    public int[]  KeyAt;         // byte offset of the key inside that row
    public int[]  Slot;          // open addressing: 0 empty, else row + 1
    public int    Cap;
    public int    Mask;
    public int    HeadStart, HeadEnd;
    public long   Probes;
    public int    Duplicates;

    public string Header()
    {
        return Encoding.UTF8.GetString(Buf, HeadStart, HeadEnd - HeadStart);
    }

    public string Row(int i)
    {
        return Encoding.UTF8.GetString(Buf, Start[i], End[i] - Start[i]);
    }

    // byte offset of field f in row i, and its length
    public bool FieldAt(int i, int f, out int off, out int len)
    {
        int p = Start[i];
        int e = End[i];
        for (int k = 0; k < f; k++)
        {
            while (p < e && Buf[p] != (byte)',') { p++; }
            if (p >= e) { off = 0; len = 0; return false; }
            p++;
        }
        int q = p;
        while (q < e && Buf[q] != (byte)',') { q++; }
        off = p;
        len = q - p;
        return true;
    }

    public string[] SplitRow(int i)
    {
        string[] v = new string[RdvSpec.Fields];
        int p = Start[i];
        int e = End[i];
        for (int f = 0; f < RdvSpec.Fields; f++)
        {
            int q = p;
            while (q < e && Buf[q] != (byte)',') { q++; }
            v[f] = (q > p) ? Encoding.UTF8.GetString(Buf, p, q - p) : "";
            p = q + 1;
            if (p > e) { p = e; }
        }
        return v;
    }

    // Probe with 8 key bytes that live anywhere: another table's buffer during
    // a join, or a small array holding what the reader typed.
    public int Find(byte[] other, int q0)
    {
        int idx = RdvHash.Of(other, q0) & Mask;
        for (; ; )
        {
            int r = Slot[idx];
            if (r == 0) { return -1; }
            int q = KeyAt[r - 1];
            if (Buf[q] == other[q0] && Buf[q + 1] == other[q0 + 1] && Buf[q + 2] == other[q0 + 2] &&
                Buf[q + 3] == other[q0 + 3] && Buf[q + 4] == other[q0 + 4] && Buf[q + 5] == other[q0 + 5] &&
                Buf[q + 6] == other[q0 + 6] && Buf[q + 7] == other[q0 + 7])
            {
                return r - 1;
            }
            idx = (idx + 1) & Mask;
        }
    }
}

public sealed class RdvHit
{
    public bool     Found;
    public string   Key1 = "";
    public string   Key2 = "";
    public int      RowA = -1, RowB = -1, RowC = -1;
    public string[] NameA, NameB, NameC;
    public string[] ValA, ValB, ValC;
}

public sealed class RdvRun
{
    public int      Seq;
    public string   Key = "";
    public DateTime When;
    public double[] Stage = new double[RdvSpec.StageCount];
    public double   TotalMs;
    public double   DetectMs;      // reference only: poll latency, not merge-select
    public int      Polls;
    public long     Checksum;
    public long     Probes;
    public int      MatchedAB, MatchedBC;
    public int      Rows;
    public bool     OracleOk;
    public string   OracleNote = "";
    public RdvHit   Hit = new RdvHit();
    public string   Error = "";

    public double StageSum()
    {
        double s = 0;
        for (int i = 0; i < RdvSpec.StageCount; i++) { s += Stage[i]; }
        return s;
    }
}

public sealed class RdvExpected
{
    public int  Rows;
    public long Checksum = -1;
    public bool Loaded;

    public static RdvExpected Read(string dir)
    {
        RdvExpected e = new RdvExpected();
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

public sealed class RdvEngine
{
    private readonly string dir;
    private readonly RdvExpected expected;
    private static readonly double TickMs = 1000.0 / (double)Stopwatch.Frequency;

    public RdvEngine(string dataDir)
    {
        dir = dataDir;
        expected = RdvExpected.Read(dataDir);
    }

    public RdvExpected Expected { get { return expected; } }
    public string DataDir { get { return dir; } }

    public static double MsSince(long t0)
    {
        return (double)(Stopwatch.GetTimestamp() - t0) * TickMs;
    }

    public static double MsBetween(long t0, long t1)
    {
        return (double)(t1 - t0) * TickMs;
    }

    public void CheckFiles()
    {
        string[] need = { "tableA.csv", "tableB.csv", "tableC.csv" };
        for (int i = 0; i < need.Length; i++)
        {
            string p = Path.Combine(dir, need[i]);
            if (!File.Exists(p)) { throw new FileNotFoundException(p); }
        }
    }

    // ---- stage 1..3: read one CSV and index it -----------------------------
    public RdvTable Load(string fileName, string name)
    {
        RdvTable t = new RdvTable();
        t.Name = name;
        t.Path = Path.Combine(dir, fileName);
        t.Buf = File.ReadAllBytes(t.Path);

        byte[] b = t.Buf;
        int n = b.Length;
        int cap = 65536;
        int[] st = new int[cap];
        int[] en = new int[cap];
        int rows = 0;
        int pos = 0;

        // skip a UTF-8 BOM: it would otherwise land inside the first key
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
        if (rows < 2) { throw new InvalidDataException(t.Path); }

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
        }
        st = null; en = null;

        t.Cap = RdvSpec.Capacity(t.Rows);
        t.Mask = t.Cap - 1;
        t.Slot = new int[t.Cap];

        long probes = 0;
        for (int i = 0; i < t.Rows; i++)
        {
            int p = t.KeyAt[i];
            // the key is field 0 and must be exactly KeyLen bytes; anything else
            // is a data problem and is reported, never worked around
            int ke = p + RdvSpec.KeyLen;
            if (ke > t.End[i] || (ke < t.End[i] && b[ke] != (byte)','))
            {
                throw new InvalidDataException(name + " row " + (i + 1).ToString(CultureInfo.InvariantCulture)
                    + ": key must be " + RdvSpec.KeyLen.ToString(CultureInfo.InvariantCulture) + " bytes");
            }
            int idx = RdvHash.Of(b, p) & t.Mask;
            while (t.Slot[idx] != 0)
            {
                probes++;
                idx = (idx + 1) & t.Mask;
            }
            t.Slot[idx] = i + 1;
        }
        t.Probes = probes;
        return t;
    }

    // ---- stage 4: A-B on key 1, one to one --------------------------------
    public static int JoinAB(RdvTable a, RdvTable b, int[] ab)
    {
        int matched = 0;
        byte[] bb = b.Buf;
        for (int i = 0; i < b.Rows; i++)
        {
            int r = a.Find(bb, b.KeyAt[i]);
            ab[i] = r;
            if (r >= 0) { matched++; }
        }
        return matched;
    }

    // ---- stage 5: B-C on key 2, one to one, and fold the checksum ---------
    public static int JoinBC(RdvTable b, RdvTable c, int[] ab, int[] bc, out long checksum)
    {
        int matched = 0;
        long chk = 0;
        byte[] bb = b.Buf;
        for (int i = 0; i < b.Rows; i++)
        {
            // key 2 is field 1 of B: step over field 0
            int p = b.Start[i];
            int e = b.End[i];
            while (p < e && bb[p] != (byte)',') { p++; }
            p++;
            int r = (p + RdvSpec.KeyLen <= e) ? c.Find(bb, p) : -1;
            bc[i] = r;
            if (r >= 0) { matched++; }
            chk = (chk * 31 + ab[i] + r) % RdvSpec.Mod;
        }
        checksum = chk;
        return matched;
    }

    // ---- one whole merge-select, stages 1..6 -------------------------------
    // t0 is the instant the key was confirmed. The caller stamps stage 7 after
    // the screen is actually updated and fills in TotalMs.
    public RdvRun Execute(string key, long t0)
    {
        RdvRun r = new RdvRun();
        r.Key = key;
        r.When = DateTime.Now;
        try
        {
            long m = Stopwatch.GetTimestamp();
            RdvTable ta = Load("tableA.csv", "A");
            r.Stage[RdvSpec.StageLoadA] = MsSince(m);

            m = Stopwatch.GetTimestamp();
            RdvTable tb = Load("tableB.csv", "B");
            r.Stage[RdvSpec.StageLoadB] = MsSince(m);

            m = Stopwatch.GetTimestamp();
            RdvTable tc = Load("tableC.csv", "C");
            r.Stage[RdvSpec.StageLoadC] = MsSince(m);

            r.Rows = tb.Rows;
            r.Probes = ta.Probes + tb.Probes + tc.Probes;

            int[] ab = new int[tb.Rows];
            int[] bc = new int[tb.Rows];

            m = Stopwatch.GetTimestamp();
            r.MatchedAB = JoinAB(ta, tb, ab);
            r.Stage[RdvSpec.StageJoinAB] = MsSince(m);

            long chk;
            m = Stopwatch.GetTimestamp();
            r.MatchedBC = JoinBC(tb, tc, ab, bc, out chk);
            r.Stage[RdvSpec.StageJoinBC] = MsSince(m);
            r.Checksum = chk;

            m = Stopwatch.GetTimestamp();
            byte[] kb = Encoding.ASCII.GetBytes(key);
            RdvHit hit = new RdvHit();
            hit.Key1 = key;
            if (kb.Length == RdvSpec.KeyLen)
            {
                int bi = tb.Find(kb, 0);
                if (bi >= 0)
                {
                    hit.Found = true;
                    hit.RowB = bi;
                    hit.RowA = ab[bi];
                    hit.RowC = bc[bi];
                    hit.NameB = SplitHeader(tb);
                    hit.ValB = tb.SplitRow(bi);
                    hit.Key2 = hit.ValB[1];
                    if (hit.RowA >= 0) { hit.NameA = SplitHeader(ta); hit.ValA = ta.SplitRow(hit.RowA); }
                    if (hit.RowC >= 0) { hit.NameC = SplitHeader(tc); hit.ValC = tc.SplitRow(hit.RowC); }
                }
            }
            r.Hit = hit;
            r.Stage[RdvSpec.StageSelect] = MsSince(m);

            r.OracleOk = true;
            if (expected.Loaded)
            {
                if (expected.Rows != r.Rows) { r.OracleOk = false; r.OracleNote = "rows"; }
                else if (expected.Checksum != r.Checksum) { r.OracleOk = false; r.OracleNote = "checksum"; }
                else if (r.MatchedAB != r.Rows || r.MatchedBC != r.Rows) { r.OracleOk = false; r.OracleNote = "matched"; }
            }
            else { r.OracleNote = "no expected.txt"; }
        }
        catch (Exception ex)
        {
            r.Error = ex.GetType().Name + ": " + ex.Message;
        }
        return r;
    }

    public static string[] SplitHeader(RdvTable t)
    {
        string[] v = new string[RdvSpec.Fields];
        string h = t.Header();
        string[] parts = h.Split(',');
        for (int i = 0; i < RdvSpec.Fields; i++) { v[i] = (i < parts.Length) ? parts[i] : ""; }
        return v;
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
