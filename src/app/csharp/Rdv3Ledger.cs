// ============================================================================
// Rdv3Ledger.cs -- the integrated ledger: one row per integrated A+B+C record.
//
// The spine is table B: a slip detail line. One B row joins to at most one A
// row (on key 1) and at most one C row (on key 2), so one B row IS one
// integrated record, and the ledger has exactly one row per B row.
//
// Row identity is key 2. Key 1 cannot identify a row -- it is legitimately
// one-to-many (a slip owns several lines). Key 2 is unique by the data model:
// the generator hands it out as a bijection over the B rows, table C is its
// 1:1 partner, and the measured comparison builds already resolve a picked
// candidate by key 2 alone (docs/results2.md). Every carry-over and every
// "processed" write below therefore keys on key 2.
//
// A ledger row is stored as one tab-joined line of the 28 content columns:
//
//   0 key1   1 key2
//   2..10    a_code a_name a_grade a_date a_amount a_rate a_flag a_dept a_note
//   11..18   b_slip b_date b_qty b_unit b_total b_status b_line b_memo
//   19..27   c_item c_maker c_cat c_price c_stock c_loc c_lot c_exp c_remark
//
// key1/key2 appear once because the join guarantees they match their tables.
// "processed" is app state, not content: it lives beside the line, is excluded
// from content comparison, and is carried across rebuilds by the rule in
// CarryProcessed below.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

public sealed class Rdv3MergeResult
{
    public string[] Lines;
    public double[] StageMs = new double[Rdv3Spec.StageCount];
    public int Rows;
    public int MatchedAB;
    public int MatchedBC;
    public long Checksum;
    public int KeysA, KeysB, KeysC;

    public double MergeMs()
    {
        double s = 0;
        for (int i = 0; i < Rdv3Spec.StageCount; i++) { s += StageMs[i]; }
        return s;
    }
}

public static class Rdv3Ledger
{
    public const int ContentCols = 28;

    // ledger column names: the "processed" flag, then the 28 content columns.
    // The xlsx header row, the VBA LEDGER sheet header row and this array must
    // stay identical; Rdv3Xlsx.Read refuses a file whose header differs.
    public static readonly string[] ContentHead =
    {
        "key1", "key2",
        "a_code", "a_name", "a_grade", "a_date", "a_amount", "a_rate", "a_flag", "a_dept", "a_note",
        "b_slip", "b_date", "b_qty", "b_unit", "b_total", "b_status", "b_line", "b_memo",
        "c_item", "c_maker", "c_cat", "c_price", "c_stock", "c_loc", "c_lot", "c_exp", "c_remark"
    };

    public const string ProcessedTrue = "TRUE";
    public const string ProcessedFalse = "FALSE";

    // ---- the merge: CSV -> ledger lines, with the eight timed stages --------
    public static Rdv3MergeResult BuildFromCsv(string dataDir)
    {
        Rdv3MergeResult r = new Rdv3MergeResult();

        long m = Rdv3Clock.Now();
        Rdv3Table ta = Rdv3Table.Read(Path.Combine(dataDir, "tableA.csv"), "A");
        r.StageMs[Rdv3Spec.StReadA] = Rdv3Clock.MsSince(m);

        m = Rdv3Clock.Now();
        Rdv3Table tb = Rdv3Table.Read(Path.Combine(dataDir, "tableB.csv"), "B");
        r.StageMs[Rdv3Spec.StReadB] = Rdv3Clock.MsSince(m);

        m = Rdv3Clock.Now();
        Rdv3Table tc = Rdv3Table.Read(Path.Combine(dataDir, "tableC.csv"), "C");
        r.StageMs[Rdv3Spec.StReadC] = Rdv3Clock.MsSince(m);

        m = Rdv3Clock.Now();
        Rdv3Index ia = new Rdv3Index(ta);
        r.StageMs[Rdv3Spec.StIdxA] = Rdv3Clock.MsSince(m);

        m = Rdv3Clock.Now();
        Rdv3Index ib = new Rdv3Index(tb);
        r.StageMs[Rdv3Spec.StIdxB] = Rdv3Clock.MsSince(m);

        m = Rdv3Clock.Now();
        Rdv3Index ic = new Rdv3Index(tc);
        r.StageMs[Rdv3Spec.StIdxC] = Rdv3Clock.MsSince(m);

        r.Rows = tb.Rows;
        r.KeysA = ia.Keys; r.KeysB = ib.Keys; r.KeysC = ic.Keys;

        int[] ab = new int[tb.Rows];
        int[] bc = new int[tb.Rows];
        List<int> found;

        // A-B over every row of B on key 1. The index answers with a set; A is
        // unique per key 1 in this data but nothing here assumes it.
        m = Rdv3Clock.Now();
        int matched = 0;
        for (int i = 0; i < tb.Rows; i++)
        {
            int c = ia.FindBytes(tb.Buf, tb.KeyAt[i], out found);
            ab[i] = (c > 0) ? found[0] : -1;
            if (c > 0) { matched++; }
        }
        r.MatchedAB = matched;
        r.StageMs[Rdv3Spec.StJoinAB] = Rdv3Clock.MsSince(m);

        // B-C over every row of B on key 2 = field 1 of B. The same join
        // checksum the comparison builds report, so expected.txt still verifies
        // this merge.
        m = Rdv3Clock.Now();
        matched = 0;
        long chk = 0;
        for (int i = 0; i < tb.Rows; i++)
        {
            int len;
            int p = tb.FieldAt(i, 1, out len);
            int rc = -1;
            if (p >= 0 && len == Rdv3Spec.KeyLength)
            {
                int c = ic.FindBytes(tb.Buf, p, out found);
                if (c > 0) { rc = found[0]; matched++; }
            }
            bc[i] = rc;
            chk = (chk * 31 + ab[i] + rc) % Rdv3Spec.Mod;
        }
        r.MatchedBC = matched;
        r.Checksum = chk;
        r.StageMs[Rdv3Spec.StJoinBC] = Rdv3Clock.MsSince(m);

        // ---- end of the timed merge region. Composing the lines below is the
        // ledger materialisation, logged separately, never part of merge time.
        r.Lines = ComposeLines(ta, tb, tc, ab, bc);
        return r;
    }

    private static string[] ComposeLines(Rdv3Table ta, Rdv3Table tb, Rdv3Table tc, int[] ab, int[] bc)
    {
        string[] lines = new string[tb.Rows];
        StringBuilder sb = new StringBuilder(256);
        for (int i = 0; i < tb.Rows; i++)
        {
            sb.Length = 0;
            sb.Append(tb.Field(i, 0)).Append('\t');            // key1
            sb.Append(tb.Field(i, 1));                          // key2
            int ra = ab[i];
            for (int f = 1; f <= 9; f++)                        // a_code..a_note
            {
                sb.Append('\t');
                if (ra >= 0) { sb.Append(ta.Field(ra, f)); }
            }
            for (int f = 2; f <= 9; f++)                        // b_slip..b_memo
            {
                sb.Append('\t').Append(tb.Field(i, f));
            }
            int rc = bc[i];
            for (int f = 1; f <= 9; f++)                        // c_item..c_remark
            {
                sb.Append('\t');
                if (rc >= 0) { sb.Append(tc.Field(rc, f)); }
            }
            lines[i] = sb.ToString();
        }
        return lines;
    }

    // ---- content comparison: the update decision --------------------------
    // The comparison is over the CONTENT of the new merge result vs the saved
    // ledger -- never over CSV timestamps or sizes. Row order counts: the
    // ledger is the B file order, so a reordered B is a different ledger.
    public static bool SameContent(string[] oldLines, string[] newLines, out int firstDiff)
    {
        firstDiff = -1;
        if (oldLines.Length != newLines.Length)
        {
            firstDiff = Math.Min(oldLines.Length, newLines.Length);
            return false;
        }
        for (int i = 0; i < newLines.Length; i++)
        {
            if (!string.Equals(oldLines[i], newLines[i], StringComparison.Ordinal))
            {
                firstDiff = i;
                return false;
            }
        }
        return true;
    }

    // ---- processed carry-over ----------------------------------------------
    // A new row keeps TRUE only if the old ledger has the same key 2 AND the
    // full content line is identical AND the old row was TRUE. A row whose
    // content changed needs processing again, so it resets to FALSE.
    public sealed class CarryStats
    {
        public int Carried, Reset, New, Dropped;
    }

    public static bool[] CarryProcessed(string[] oldLines, bool[] oldProcessed, string[] newLines, CarryStats stats)
    {
        Dictionary<string, int> byKey2 = new Dictionary<string, int>(oldLines.Length, StringComparer.Ordinal);
        for (int i = 0; i < oldLines.Length; i++)
        {
            string k2 = FieldOf(oldLines[i], 1);
            // key 2 is unique by the data model; if a file ever violates that,
            // the first row keeps the identity and the duplicate is logged by
            // the caller via stats -- nothing is silently merged.
            if (!byKey2.ContainsKey(k2)) { byKey2.Add(k2, i); }
        }

        bool[] np = new bool[newLines.Length];
        int used = 0;
        for (int i = 0; i < newLines.Length; i++)
        {
            string k2 = FieldOf(newLines[i], 1);
            int oi;
            if (byKey2.TryGetValue(k2, out oi))
            {
                used++;
                if (string.Equals(oldLines[oi], newLines[i], StringComparison.Ordinal))
                {
                    np[i] = oldProcessed[oi];
                    if (np[i]) { stats.Carried++; }
                }
                else
                {
                    np[i] = false;
                    if (oldProcessed[oi]) { stats.Reset++; }
                }
            }
            else
            {
                np[i] = false;
                stats.New++;
            }
        }
        stats.Dropped = oldLines.Length - used;
        if (stats.Dropped < 0) { stats.Dropped = 0; }
        return np;
    }

    // ---- field access on a ledger line -------------------------------------
    public static string FieldOf(string line, int col)
    {
        int p = 0;
        for (int k = 0; k < col; k++)
        {
            int t = line.IndexOf('\t', p);
            if (t < 0) { return ""; }
            p = t + 1;
        }
        int q = line.IndexOf('\t', p);
        if (q < 0) { q = line.Length; }
        return line.Substring(p, q - p);
    }

    public static string[] SplitLine(string line)
    {
        string[] v = line.Split('\t');
        if (v.Length == ContentCols) { return v; }
        string[] w = new string[ContentCols];
        for (int i = 0; i < ContentCols; i++) { w[i] = (i < v.Length) ? v[i] : ""; }
        return w;
    }

    // the eight identification columns of the candidate list -- the same eight,
    // in the same order, as the measured comparison builds and README.md:
    // key2, b_line, b_slip, b_date, b_qty, b_status, c_item, c_maker
    public static readonly int[] CandCols = { 1, 17, 11, 12, 13, 16, 19, 20 };

    public static string[] CandidateColumns(string line)
    {
        string[] f = SplitLine(line);
        string[] c = new string[CandCols.Length];
        for (int i = 0; i < CandCols.Length; i++) { c[i] = f[CandCols[i]]; }
        return c;
    }

    // the 30-field record view: the ledger row unfolded back into the three
    // 10-field tables (key1/key2 re-shown in each table they belong to)
    public static void RecordView(string line, out string[] valA, out string[] valB, out string[] valC)
    {
        string[] f = SplitLine(line);
        valA = new string[Rdv3Spec.Fields];
        valB = new string[Rdv3Spec.Fields];
        valC = new string[Rdv3Spec.Fields];
        bool hasA = false, hasC = false;
        for (int i = 2; i <= 10; i++) { if (f[i].Length > 0) { hasA = true; break; } }
        for (int i = 19; i <= 27; i++) { if (f[i].Length > 0) { hasC = true; break; } }
        valA[0] = hasA ? f[0] : "";
        for (int i = 1; i < 10; i++) { valA[i] = f[1 + i]; }
        valB[0] = f[0];
        valB[1] = f[1];
        for (int i = 2; i < 10; i++) { valB[i] = f[9 + i]; }
        valC[0] = hasC ? f[1] : "";
        for (int i = 1; i < 10; i++) { valC[i] = f[18 + i]; }
    }

    public static readonly string[] NameA =
    { "key1", "a_code", "a_name", "a_grade", "a_date", "a_amount", "a_rate", "a_flag", "a_dept", "a_note" };
    public static readonly string[] NameB =
    { "key1", "key2", "b_slip", "b_date", "b_qty", "b_unit", "b_total", "b_status", "b_line", "b_memo" };
    public static readonly string[] NameC =
    { "key2", "c_item", "c_maker", "c_cat", "c_price", "c_stock", "c_loc", "c_lot", "c_exp", "c_remark" };

    public static string[] Key1Column(string[] lines)
    {
        string[] keys = new string[lines.Length];
        for (int i = 0; i < lines.Length; i++) { keys[i] = FieldOf(lines[i], 0); }
        return keys;
    }
}
