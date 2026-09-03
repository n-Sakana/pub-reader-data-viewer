// ============================================================================
// Rdv3Ledger.cs -- the integrated ledger and its optimized left-join path.
//
// General table operations run in Rdv3Process. The common left-join pipeline
// is compiled to the byte-indexed path here, without changing its JSON
// meaning. Ledger writes use the same merge/replace step in either path.
//
// Row identity is the configured ledger key (IdentityCol). Every carry-over and every
// work-state write keys on it. A ledger row is stored as one tab-joined line
// of the content columns; the column NAMES are the CSV header names in ledger
// order (Rdv3Data.Head), and the program never spells them out. The WORK
// STATE (todo / done) is app state, not content: it lives beside the line as
// the stored string the screen definition maps to a state, is excluded from
// content comparison. Its on-source-change behavior comes from the application
// column definition in settings.json.
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
    internal Rdv3ProcessJobDef Job;
    internal Rdv3PreparedProcess Prepared;

    public double MergeMs()
    {
        double s = 0;
        for (int i = 0; i < ReadMs.Length; i++) { s += ReadMs[i] + IndexMs[i]; }
        for (int i = 0; i < JoinMs.Length; i++) { s += JoinMs[i]; }
        return s;
    }
}

public sealed class Rdv3UpdateResult
{
    public string[] Lines;
    public string[] States;
    public readonly List<string> ResetLines = new List<string>();
    public int Added;
    public int Updated;
    public int Unchanged;
    public int Kept;
    public int Deleted;
}

public sealed class Rdv3DeleteResult
{
    public string[] Lines;
    public string[] States;
    public int Deleted;
}

public static class Rdv3Ledger
{
    // the modulus of the join fingerprint logged with every merge
    private const long Mod = 1000000007L;

    // ---- the merge: CSV -> ledger lines, with the timed stages ------------------
    public static Rdv3MergeResult BuildFromCsv(Rdv3Data d, string dataDir)
    {
        return BuildFromCsv(d, d.UpdateJob, dataDir);
    }

    public static Rdv3MergeResult BuildFromCsv(Rdv3Data d, Rdv3ProcessJobDef job, string dataDir)
    {
        if (d == null || job == null || job.Kind != "update")
        {
            throw new InvalidOperationException("not an update job");
        }
        Rdv3MergeResult r = new Rdv3MergeResult();
        r.Job = job;
        int nt = d.Tables.Count;
        r.ReadMs = new double[nt];
        r.IndexMs = new double[nt];
        r.Keys = new int[nt];
        r.JoinMs = new double[job.Joins.Count];
        r.Matched = new int[job.Joins.Count];

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

        if (!job.FastJoinPlan)
        {
            // the tables just read (and their keys just proved unique) are the
            // pipeline's table inputs; only a non-table input is read here
            r.Prepared = Rdv3Process.Prepare(d, job, dataDir, tables);
            Rdv3ProcessResult process = Rdv3Process.Execute(r.Prepared, new string[0], new string[0], "", false);
            if (process.Kind != "ledger") { throw new InvalidDataException("automatic update job did not produce ledger"); }
            r.Head = d.Head;
            r.Lines = process.Lines;
            r.Rows = process.Lines.Length;
            r.Checksum = 0;
            return r;
        }

        Rdv3Table spine = tables[job.SpineOrd];
        r.Rows = spine.Rows;

        // joined[j][i] = the row of join j's table that spine row i joins to, or -1.
        // The checksum is a fingerprint of the join (spine row by row, the
        // joined row numbers in join order): it is logged with the merge, and
        // the build compares it with the figure the data generator recorded.
        int[][] joined = new int[job.Joins.Count][];
        List<int> found;
        for (int j = 0; j < job.Joins.Count; j++)
        {
            long m = Rdv3Clock.Now();
            Rdv3JoinDef jd = job.Joins[j];
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
            for (int j = 0; j < job.Joins.Count; j++) { sum += joined[j][i]; }
            chk = (chk * 31 + sum) % Mod;
        }
        r.Checksum = chk;

        // ---- end of the timed merge region. Composing the lines below is the
        // ledger materialisation, logged separately, never part of merge time.
        r.Head = d.Head;
        r.Lines = ComposeLines(d, job, tables, joined);
        return r;
    }

    private static string[] ComposeLines(Rdv3Data d, Rdv3ProcessJobDef job,
                                         Rdv3Table[] tables, int[][] joined)
    {
        Rdv3Table spine = tables[job.SpineOrd];
        int nc = d.Columns.Count;
        // for each ledger column: its field and route (-1 = spine, -2 = absent)
        int[] tbl = new int[nc];
        int[] fld = new int[nc];
        int[] via = new int[nc];
        for (int c = 0; c < nc; c++)
        {
            tbl[c] = d.Columns[c].TableOrd;
            fld[c] = d.Columns[c].Field;
            via[c] = (tbl[c] == job.SpineOrd) ? -1 : -2;
            for (int j = 0; j < job.Joins.Count; j++) { if (job.Joins[j].TableOrd == tbl[c]) { via[c] = j; } }
        }
        string[] lines = new string[spine.Rows];
        StringBuilder sb = new StringBuilder(256);
        for (int i = 0; i < spine.Rows; i++)
        {
            sb.Length = 0;
            for (int c = 0; c < nc; c++)
            {
                if (c > 0) { sb.Append('\t'); }
                int row = (via[c] == -2) ? -1 : ((via[c] < 0) ? i : joined[via[c]][i]);
                if (row >= 0) { sb.Append(tables[tbl[c]].Field(row, fld[c])); }
            }
            lines[i] = sb.ToString();
        }
        return lines;
    }

    // ---- content comparison: the update decision --------------------------
    // The comparison is over the CONTENT of the new merge result vs the saved
    // ledger -- never over CSV timestamps or sizes. Row order counts because
    // the ordered operations define the output order too.
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

    public static bool SameLedger(string[] oldLines, string[] oldStates,
                                  string[] newLines, string[] newStates)
    {
        int ignored;
        if (!SameContent(oldLines, newLines, out ignored)) { return false; }
        if (oldStates == null || newStates == null || oldStates.Length != newStates.Length) { return false; }
        for (int i = 0; i < oldStates.Length; i++)
        {
            if (!string.Equals(oldStates[i], newStates[i], StringComparison.Ordinal)) { return false; }
        }
        return true;
    }

    // Apply the one ledger-writing step of an update job. A merge keeps the
    // destination's order and appends new source rows. Replace keeps the
    // source's order and drops destination-only rows. Application-owned state
    // follows its configured on-source-change rule.
    public static Rdv3UpdateResult ApplyUpdate(Rdv3ProcessJobDef job,
                                                string[] targetLines, string[] targetStates,
                                                string[] sourceLines, int identityCol,
                                                string initialStored)
    {
        if (job == null || job.ApplyStep == null) { throw new InvalidOperationException("update job has no ledger-writing step"); }
        string[] oldLines = targetLines ?? new string[0];
        string[] oldStates = targetStates ?? FreshStates(oldLines.Length, initialStored);
        string[] newLines = sourceLines ?? new string[0];
        if (oldStates.Length != oldLines.Length) { throw new InvalidDataException("ledger content and application-column lengths differ"); }

        Dictionary<string, int> oldById = RowMap(oldLines, identityCol, "target");
        Dictionary<string, int> sourceById = RowMap(newLines, identityCol, "source");
        Rdv3ProcessStepDef step = job.ApplyStep;
        if (job.OnSourceChange != "reset" && job.OnSourceChange != "preserve")
        {
            throw new InvalidDataException("application column has no onSourceChange rule");
        }
        string sourceOnly = (step.Operation == "replace") ? "add" : step.SourceOnly;
        string both = (step.Operation == "replace") ? "update" : step.Both;
        string targetOnly = (step.Operation == "replace") ? "delete" : step.TargetOnly;

        List<string> lines = new List<string>(Math.Max(oldLines.Length, newLines.Length));
        List<string> states = new List<string>(Math.Max(oldLines.Length, newLines.Length));
        Rdv3UpdateResult result = new Rdv3UpdateResult();

        if (step.Operation == "replace")
        {
            for (int i = 0; i < newLines.Length; i++)
            {
                string identity = FieldOf(newLines[i], identityCol);
                int oldRow;
                if (oldById.TryGetValue(identity, out oldRow))
                {
                    AddMatched(result, lines, states, oldLines[oldRow], oldStates[oldRow],
                               newLines[i], both, initialStored, job.OnSourceChange);
                }
                else if (sourceOnly == "add")
                {
                    lines.Add(newLines[i]);
                    states.Add(initialStored);
                    result.Added++;
                }
            }
            for (int i = 0; i < oldLines.Length; i++)
            {
                string identity = FieldOf(oldLines[i], identityCol);
                if (!sourceById.ContainsKey(identity)) { result.Deleted++; }
            }
        }
        else
        {
            HashSet<string> consumed = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < oldLines.Length; i++)
            {
                string identity = FieldOf(oldLines[i], identityCol);
                int sourceRow;
                if (sourceById.TryGetValue(identity, out sourceRow))
                {
                    consumed.Add(identity);
                    AddMatched(result, lines, states, oldLines[i], oldStates[i],
                               newLines[sourceRow], both, initialStored, job.OnSourceChange);
                }
                else if (targetOnly == "keep")
                {
                    lines.Add(oldLines[i]);
                    states.Add(oldStates[i]);
                    result.Kept++;
                }
                else { result.Deleted++; }
            }
            if (sourceOnly == "add")
            {
                for (int i = 0; i < newLines.Length; i++)
                {
                    string identity = FieldOf(newLines[i], identityCol);
                    if (consumed.Contains(identity)) { continue; }
                    lines.Add(newLines[i]);
                    states.Add(initialStored);
                    result.Added++;
                }
            }
        }

        result.Lines = lines.ToArray();
        result.States = states.ToArray();
        return result;
    }

    private static void AddMatched(Rdv3UpdateResult result, List<string> lines, List<string> states,
                                   string oldLine, string oldState, string sourceLine,
                                   string both, string initialStored, string onSourceChange)
    {
        if (both == "keep")
        {
            lines.Add(oldLine);
            states.Add(oldState);
            result.Kept++;
            return;
        }
        lines.Add(sourceLine);
        if (string.Equals(oldLine, sourceLine, StringComparison.Ordinal))
        {
            states.Add(oldState);
            result.Unchanged++;
        }
        else
        {
            result.Updated++;
            if (onSourceChange == "preserve") { states.Add(oldState); }
            else
            {
                states.Add(initialStored);
                if (!string.Equals(oldState, initialStored, StringComparison.Ordinal)) { result.ResetLines.Add(sourceLine); }
            }
        }
    }

    // ---- row identity ------------------------------------------------------
    // Every ledger row is told apart by its identity column. One scan finds
    // the rows by identity, or the first blank / duplicated one; the two
    // callers below say it differently.
    private static Dictionary<string, int> ScanIdentities(string[] lines, int identityCol,
                                                          out int blankRow, out int dupFirst, out int dupSecond)
    {
        blankRow = -1;
        dupFirst = -1;
        dupSecond = -1;
        Dictionary<string, int> rows = new Dictionary<string, int>(lines.Length, StringComparer.Ordinal);
        for (int i = 0; i < lines.Length; i++)
        {
            string identity = FieldOf(lines[i], identityCol);
            if (identity.Length == 0) { blankRow = i; return rows; }
            int first;
            if (rows.TryGetValue(identity, out first)) { dupFirst = first; dupSecond = i; return rows; }
            rows.Add(identity, i);
        }
        return rows;
    }

    // the rows by identity, for lines the caller has already validated: a
    // blank or duplicated identity here is a broken invariant, not user data
    public static Dictionary<string, int> RowMap(string[] lines, int identityCol, string where)
    {
        int blank, first, second;
        Dictionary<string, int> rows = ScanIdentities(lines, identityCol, out blank, out first, out second);
        if (blank >= 0) { throw new InvalidDataException("blank row identity in " + where + " at row " + Row(blank)); }
        if (second >= 0) { throw new InvalidDataException("duplicate row identity in " + where + ": " + FieldOf(lines[second], identityCol)); }
        return rows;
    }

    // the saved ledger as read from its file, checked the way a CSV is: a blank
    // or duplicated identity (a hand edit in Excel, say) is a data error that
    // names the file, the column and the rows as the sheet numbers them
    public static void CheckIdentities(string[] lines, int identityCol, string file, string name)
    {
        int blank, first, second;
        ScanIdentities(lines, identityCol, out blank, out first, out second);
        if (blank >= 0)
        {
            throw new Rdv3DataError(Rdv3Text.DataLedgerBlankIdentity
                .Replace("{file}", file).Replace("{row}", Row(blank)).Replace("{name}", name));
        }
        if (second >= 0)
        {
            throw new Rdv3DataError(Rdv3Text.DataLedgerDupIdentity
                .Replace("{file}", file).Replace("{name}", name)
                .Replace("{key}", FieldOf(lines[second], identityCol))
                .Replace("{row1}", Row(first)).Replace("{row2}", Row(second)));
        }
    }

    // the row number a person sees in the sheet (the header is row 1)
    private static string Row(int i)
    {
        return (i + 2).ToString(CultureInfo.InvariantCulture);
    }

    public static Rdv3DeleteResult ApplyDelete(Rdv3Data data, Rdv3ProcessJobDef job,
                                                string dataDir, string[] ledgerLines,
                                                string[] ledgerStates, string initialStored)
    {
        if (data == null || job == null || job.Kind != "delete") { throw new InvalidOperationException("not a delete job"); }
        Rdv3ProcessResult run = Rdv3Process.Run(data, job, dataDir, ledgerLines, ledgerStates, initialStored);
        if (run.Kind != "ledger") { throw new InvalidOperationException("delete job did not produce ledger"); }
        Rdv3DeleteResult result = new Rdv3DeleteResult();
        result.Lines = run.Lines;
        result.States = run.States;
        result.Deleted = run.Deleted;
        return result;
    }

    // ---- work-state carry-over ---------------------------------------------
    // The stored string is carried verbatim or reset according to the
    // application-column rule. The ledger never decides what a state means;
    // the screen definition supplies the stored initial value.
    public sealed class CarryStats
    {
        public int Carried, Reset, New, Dropped;
    }

    public static string[] CarryStates(string[] oldLines, string[] oldStates, string[] newLines,
                                       int identityCol, string initialStored, string onSourceChange,
                                       CarryStats stats)
    {
        if (onSourceChange != "reset" && onSourceChange != "preserve")
        {
            throw new InvalidDataException("application column has no onSourceChange rule");
        }
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
                if (string.Equals(oldLines[oi], newLines[i], StringComparison.Ordinal)
                    || onSourceChange == "preserve")
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
