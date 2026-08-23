// ============================================================================
// Rdv3Ledger.cs -- the integrated ledger: one row per row of the spine table.
//
// The definition (Rdv3Data) says which table is the spine, which tables join
// to it and on which spine column, and which columns -- "<table>.<column>" --
// make up a ledger row, in which order. A spine row joins to at most one row
// of each joined table (a table's key is unique, the index enforces it), so
// one spine row IS one integrated record.
//
// Row identity is the spine's key (IdentityCol). Every carry-over and every
// work-state write keys on it. A ledger row is stored as one tab-joined line
// of the content columns; the column NAMES are the CSV header names in ledger
// order (Rdv3Data.Head), and the program never spells them out. The WORK
// STATE (todo / done) is app state, not content: it lives beside the line as
// the stored string the screen definition maps to a state, is excluded from
// content comparison, and is carried across rebuilds by the rule in
// CarryStates below.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

public sealed class Rdv3MergeResult
{
    public string[] Lines;
    public int Rows;
    public long Checksum;
    // per table (definition order): read ms, index ms, distinct keys
    public double[] ReadMs;
    public double[] IndexMs;
    public int[] Keys;
    // per join (definition order): join ms, matched spine rows
    public double[] JoinMs;
    public int[] Matched;
    // the content column names, in ledger order
    public string[] Head;

    public double MergeMs()
    {
        double s = 0;
        for (int i = 0; i < ReadMs.Length; i++) { s += ReadMs[i] + IndexMs[i]; }
        for (int i = 0; i < JoinMs.Length; i++) { s += JoinMs[i]; }
        return s;
    }
}

public static class Rdv3Ledger
{
    // ---- the merge: CSV -> ledger lines, with the timed stages ------------------
    public static Rdv3MergeResult BuildFromCsv(Rdv3Data d, string dataDir)
    {
        Rdv3MergeResult r = new Rdv3MergeResult();
        int nt = d.Tables.Count;
        r.ReadMs = new double[nt];
        r.IndexMs = new double[nt];
        r.Keys = new int[nt];
        r.JoinMs = new double[d.Joins.Count];
        r.Matched = new int[d.Joins.Count];

        Rdv3Table[] tables = new Rdv3Table[nt];
        Rdv3Index[] index = new Rdv3Index[nt];
        string[][] heads = new string[nt][];
        for (int t = 0; t < nt; t++)
        {
            long m = Rdv3Clock.Now();
            tables[t] = Rdv3Table.Read(Path.Combine(dataDir, d.Tables[t].File), d.Tables[t].Id, d.Enc, d.Tables[t].Key);
            r.ReadMs[t] = Rdv3Clock.MsSince(m);
            heads[t] = tables[t].Head;
        }
        // the definition's names against the headers actually read
        d.Bind(heads);
        for (int t = 0; t < nt; t++)
        {
            long m = Rdv3Clock.Now();
            index[t] = new Rdv3Index(tables[t]);
            r.IndexMs[t] = Rdv3Clock.MsSince(m);
            r.Keys[t] = index[t].Keys;
        }

        Rdv3Table spine = tables[d.SpineOrd];
        r.Rows = spine.Rows;

        // joined[j][i] = the row of join j's table that spine row i joins to, or -1.
        // The checksum is the one the comparison builds report (spine row by
        // row, the joined row numbers in join order), so expected.txt still
        // verifies the shipped merge.
        int[][] joined = new int[d.Joins.Count][];
        List<int> found;
        for (int j = 0; j < d.Joins.Count; j++)
        {
            long m = Rdv3Clock.Now();
            Rdv3JoinDef jd = d.Joins[j];
            Rdv3Index ix = index[jd.TableOrd];
            int[] rows = new int[spine.Rows];
            int matched = 0;
            for (int i = 0; i < spine.Rows; i++)
            {
                int len;
                int p = spine.FieldAt(i, jd.OnField, out len);
                int rc = -1;
                if (p >= 0 && len > 0)
                {
                    int c = ix.FindBytes(spine.Buf, p, len, out found);
                    if (c > 0) { rc = found[0]; matched++; }
                }
                rows[i] = rc;
            }
            joined[j] = rows;
            r.Matched[j] = matched;
            r.JoinMs[j] = Rdv3Clock.MsSince(m);
        }
        long chk = 0;
        for (int i = 0; i < spine.Rows; i++)
        {
            long sum = 0;
            for (int j = 0; j < d.Joins.Count; j++) { sum += joined[j][i]; }
            chk = (chk * 31 + sum) % Rdv3Spec.Mod;
        }
        r.Checksum = chk;

        // ---- end of the timed merge region. Composing the lines below is the
        // ledger materialisation, logged separately, never part of merge time.
        r.Head = d.Head;
        r.Lines = ComposeLines(d, tables, joined);
        return r;
    }

    private static string[] ComposeLines(Rdv3Data d, Rdv3Table[] tables, int[][] joined)
    {
        Rdv3Table spine = tables[d.SpineOrd];
        int nc = d.Columns.Count;
        // for each ledger column: the table, its field, and which join (or -1 = spine)
        int[] tbl = new int[nc];
        int[] fld = new int[nc];
        int[] via = new int[nc];
        for (int c = 0; c < nc; c++)
        {
            tbl[c] = d.Columns[c].TableOrd;
            fld[c] = d.Columns[c].Field;
            via[c] = -1;
            for (int j = 0; j < d.Joins.Count; j++) { if (d.Joins[j].TableOrd == tbl[c]) { via[c] = j; } }
        }
        string[] lines = new string[spine.Rows];
        StringBuilder sb = new StringBuilder(256);
        for (int i = 0; i < spine.Rows; i++)
        {
            sb.Length = 0;
            for (int c = 0; c < nc; c++)
            {
                if (c > 0) { sb.Append('\t'); }
                int row = (via[c] < 0) ? i : joined[via[c]][i];
                if (row >= 0) { sb.Append(tables[tbl[c]].Field(row, fld[c])); }
            }
            lines[i] = sb.ToString();
        }
        return lines;
    }

    // ---- content comparison: the update decision --------------------------
    // The comparison is over the CONTENT of the new merge result vs the saved
    // ledger -- never over CSV timestamps or sizes. Row order counts: the
    // ledger is the spine file's order, so a reordered spine is a different
    // ledger.
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

    // ---- work-state carry-over ---------------------------------------------
    // A new row keeps its stored state only if the old ledger has the same
    // identity AND the full content line is identical. A row whose content
    // changed needs working again, so it goes back to the initial state; so
    // does a row that is new. The stored string is carried verbatim -- the
    // ledger never decides what a state means, the screen definition does.
    public sealed class CarryStats
    {
        public int Carried, Reset, New, Dropped;
    }

    public static string[] CarryStates(string[] oldLines, string[] oldStates, string[] newLines,
                                       int identityCol, string initialStored, CarryStats stats)
    {
        Dictionary<string, int> byId = new Dictionary<string, int>(oldLines.Length, StringComparer.Ordinal);
        for (int i = 0; i < oldLines.Length; i++)
        {
            string id = FieldOf(oldLines[i], identityCol);
            // the new lines come from a table whose key the index proved unique;
            // the OLD lines are whatever the xlsx held, so a duplicate there is
            // possible (a hand edit). The first row keeps the identity and the
            // rest count as dropped -- nothing is merged.
            if (!byId.ContainsKey(id)) { byId.Add(id, i); }
        }

        string[] ns = new string[newLines.Length];
        int used = 0;
        for (int i = 0; i < newLines.Length; i++)
        {
            string id = FieldOf(newLines[i], identityCol);
            int oi;
            if (byId.TryGetValue(id, out oi))
            {
                used++;
                if (string.Equals(oldLines[oi], newLines[i], StringComparison.Ordinal))
                {
                    ns[i] = oldStates[oi];
                    if (!string.Equals(ns[i], initialStored, StringComparison.Ordinal)) { stats.Carried++; }
                }
                else
                {
                    ns[i] = initialStored;
                    if (!string.Equals(oldStates[oi], initialStored, StringComparison.Ordinal)) { stats.Reset++; }
                }
            }
            else
            {
                ns[i] = initialStored;
                stats.New++;
            }
        }
        stats.Dropped = oldLines.Length - used;
        if (stats.Dropped < 0) { stats.Dropped = 0; }
        return ns;
    }

    public static string[] FreshStates(int rows, string initialStored)
    {
        string[] s = new string[rows];
        for (int i = 0; i < rows; i++) { s[i] = initialStored; }
        return s;
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
        return line.Split('\t');
    }

    public static string[] Column(string[] lines, int col)
    {
        string[] keys = new string[lines.Length];
        for (int i = 0; i < lines.Length; i++) { keys[i] = FieldOf(lines[i], col); }
        return keys;
    }
}
