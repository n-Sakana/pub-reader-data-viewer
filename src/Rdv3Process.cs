// ============================================================================
// Rdv3Process.cs -- typed table-operation pipelines from settings.json.
//
// Every operation consumes values made by inputs or earlier steps. The three
// value kinds are table, rows (a selection tied to one table value), and
// ledger. Validation and execution use the same rules. No job-kind-specific
// sequence is compiled here.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

public sealed class Rdv3ProcessValueResult
{
    public string Kind = "";
    public string[] Columns = new string[0];
    public string[] Lines = new string[0];
    public int Count;
}

public sealed class Rdv3ProcessResult
{
    public readonly List<Rdv3JoinResult> Joins = new List<Rdv3JoinResult>();
    public readonly List<string> Warnings = new List<string>();
    public string Kind = "";
    public string[] Columns = new string[0];
    public string[] Lines = new string[0];
    public string[] States = new string[0];
    public int Deleted;
    public int SkippedInvalid;
    public Rdv3UpdateResult Update;
    public readonly Dictionary<string, Rdv3ProcessValueResult> Values
        = new Dictionary<string, Rdv3ProcessValueResult>(StringComparer.Ordinal);

    public Rdv3ProcessValueResult ValueOf(string name)
    {
        Rdv3ProcessValueResult value;
        return Values.TryGetValue(name, out value) ? value : null;
    }
}

public sealed class Rdv3JoinResult
{
    public string Output;
    public int LeftRows, RightRows, OutputRows, UnmatchedLeft, UnmatchedRight;
}

internal sealed class Rdv3InputResult
{
    public string Id, File;
    public int Rows, SkippedEmpty, SkippedDuplicate, SkippedShort, SkippedBlank, SkippedColumns, SkippedInvalid;

    public Rdv3InputResult(string id, Rdv3Table table)
    {
        Id = id; File = table.Path; Rows = table.Rows;
        SkippedEmpty = table.SkippedEmptyRows; SkippedDuplicate = table.SkippedDuplicateRows;
        SkippedShort = table.InputCounts.ShortRows; SkippedBlank = table.InputCounts.BlankRows;
        SkippedColumns = table.InputCounts.HeaderColumns;
        SkippedInvalid = table.InputCounts.InvalidRows;
    }
}

internal sealed class Rdv3Relation
{
    public string Kind = "table";
    public string[] Columns = new string[0];
    public List<string[]> Rows = new List<string[]>();
    public List<string> Origins = new List<string>();
    public List<string> States;

    public string Origin(int row)
    { return Origins.Count > row ? Origins[row] : Rdv3Text.Format(Rdv3Text.SourceRow, "", row + 1); }

    public int ColumnOf(string name)
    {
        for (int i = 0; i < Columns.Length; i++)
        {
            if (Columns[i] == name) { return i; }
        }
        return -1;
    }

    public int NeedColumn(string name)
    {
        int column = ColumnOf(name);
        if (column < 0) { throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessMissingColumn, name)); }
        return column;
    }

    public int[] NeedColumns(string[] names)
    {
        int[] cols = new int[names.Length];
        for (int i = 0; i < cols.Length; i++) { cols[i] = NeedColumn(names[i]); }
        return cols;
    }

    public string[] ToLines()
    {
        string[] lines = new string[Rows.Count];
        for (int i = 0; i < Rows.Count; i++) { lines[i] = string.Join("\t", Rows[i]); }
        return lines;
    }
}

internal sealed class Rdv3RowSelection
{
    public Rdv3Relation Table;
    public HashSet<int> Rows = new HashSet<int>();
}

internal sealed class Rdv3PreparedProcess
{
    public readonly List<Rdv3InputResult> InputResults = new List<Rdv3InputResult>();
    public readonly List<string> Warnings = new List<string>();
    public Rdv3Data Data;
    public Rdv3ProcessJobDef Job;
    public Dictionary<string, object> Inputs = new Dictionary<string, object>(StringComparer.Ordinal);
}

public static class Rdv3Process
{
    public static Rdv3ProcessResult Run(Rdv3Data data, Rdv3ProcessJobDef job, string dataDir,
                                        string[] ledgerLines, string[] ledgerStates,
                                        string initialStored)
    {
        Rdv3PreparedProcess prepared = Prepare(data, job, dataDir);
        return Execute(prepared, ledgerLines, ledgerStates, initialStored, true);
    }

    internal static Rdv3PreparedProcess Prepare(Rdv3Data data, Rdv3ProcessJobDef job, string dataDir)
    {
        return Prepare(data, job, dataDir, null);
    }

    // tables: the definition's tables the caller has already read and indexed
    // (definition order), so they are not read twice; null reads every input
    internal static Rdv3PreparedProcess Prepare(Rdv3Data data, Rdv3ProcessJobDef job, string dataDir, Rdv3Table[] tables)
    {
        if (data == null || job == null) { throw new ArgumentNullException("data"); }
        Rdv3PreparedProcess prepared = new Rdv3PreparedProcess();
        prepared.Data = data;
        prepared.Job = job;
        for (int i = 0; i < job.Inputs.Count; i++)
        {
            Rdv3ProcessInputDef input = job.Inputs[i];
            Rdv3Table table = (input.IsTable && tables != null) ? tables[input.TableOrd] : null;
            if (table == null)
            {
                string path = Path.IsPathRooted(input.File) ? input.File : Path.Combine(dataDir, input.File);
                table = Rdv3Table.Read(path, input.Id, input.Enc, input.Columns ?? new string[] { input.Column }, input.KeyValidation, input.EncodingSetting, data.SourceReferences(input), input.HeaderRow, input.Delimiter, input.Sheet);
                if (input.IsTable) { data.ValidateInput(table, input.TableOrd); }
                new Rdv3Index(table);                    // enforce the configured duplicate rule
                table.AddWarnings(prepared.Warnings);
            }
            prepared.InputResults.Add(new Rdv3InputResult(input.Id, table));
            prepared.Inputs.Add(input.Id, input.IsTable
                ? RelationOfTable(input, table) : RelationOfValues(input, table));
        }
        return prepared;
    }

    private static Rdv3PreparedProcess PrepareFromHeads(Rdv3Data data, Rdv3ProcessJobDef job,
                                                         string[][] heads)
    {
        Rdv3PreparedProcess prepared = new Rdv3PreparedProcess();
        prepared.Data = data;
        prepared.Job = job;
        for (int i = 0; i < job.Inputs.Count; i++)
        {
            Rdv3ProcessInputDef input = job.Inputs[i];
            Rdv3Relation relation = new Rdv3Relation();
            if (input.IsTable)
            {
                string[] head = heads[input.TableOrd];
                relation.Columns = new string[head.Length];
                for (int c = 0; c < head.Length; c++) { relation.Columns[c] = input.Table + "." + head[c]; }
            }
            else { relation.Columns = new string[] { input.Key }; }
            prepared.Inputs.Add(input.Id, relation);
        }
        return prepared;
    }

    // Walks every job over the input headers alone. Returns the column
    // references the update job's tables and row sets carry (null when that
    // walk failed), so that the caller can accept a ledger column a step makes.
    public static HashSet<string> ValidateColumns(Rdv3Data data, string[][] heads, Rdv3Validation validation = null)
    {
        HashSet<string> produced = null;
        for (int i = 0; i < data.Jobs.Count; i++)
        {
            Rdv3ProcessJobDef job = data.Jobs[i];
            Action check = delegate {
            Rdv3PreparedProcess prepared = PrepareFromHeads(data, job, heads);
            Rdv3ProcessResult result = Execute(prepared, new string[0], new string[0], "", true);
            if (job != data.UpdateJob) { return; }
            produced = new HashSet<string>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, Rdv3ProcessValueResult> value in result.Values)
            { if (value.Value.Kind != "ledger") { produced.UnionWith(value.Value.Columns); } }
            CheckLedgerSource(data, job, result);
            };
            if (validation == null) { check(); }
            else { validation.Check("data.jobs[" + i.ToString(CultureInfo.InvariantCulture) + "] column dependencies", check); }
        }
        if (validation != null) { validation.Finish("job columns", "input types and job preparation"); }
        return produced;
    }

    // Every saved column must reach the ledger-writing step: a column of a
    // table that was never joined, or dropped by a select, would otherwise
    // be written as a silent blank in every row.
    private static void CheckLedgerSource(Rdv3Data data, Rdv3ProcessJobDef job, Rdv3ProcessResult result)
    {
        if (job.ApplyStep == null) { return; }
        Rdv3ProcessValueResult source = result.ValueOf(job.ApplyStep.Target1);
        if (source == null || source.Kind == "ledger") { return; }
        HashSet<string> columns = new HashSet<string>(source.Columns, StringComparer.Ordinal);
        for (int c = 0; c < data.Columns.Count; c++)
        {
            if (Array.IndexOf(data.IdentityCols, c) >= 0 || columns.Contains(data.Columns[c].Ref)) { continue; }
            throw new Rdv3DataError(Rdv3Text.LedgerColumnNotProduced
                .Replace("{name}", data.Columns[c].Ref)
                .Replace("{step}", job.ApplyStep.Operation + " " + job.ApplyStep.Target1));
        }
    }

    internal static Rdv3ProcessResult Execute(Rdv3PreparedProcess prepared,
                                              string[] ledgerLines, string[] ledgerStates,
                                              string initialStored, bool capture)
    {
        if (prepared == null) { throw new ArgumentNullException("prepared"); }
        Rdv3Data data = prepared.Data;
        Rdv3ProcessJobDef job = prepared.Job;
        Dictionary<string, object> values = new Dictionary<string, object>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, object> pair in prepared.Inputs) { values.Add(pair.Key, pair.Value); }
        values.Add("ledger", RelationOfLedger(data, ledgerLines, ledgerStates, initialStored));

        object last = null;
        Rdv3ProcessResult result = new Rdv3ProcessResult();
        result.Warnings.AddRange(prepared.Warnings);
        int directUpdated = 0;
        List<string> directReset = new List<string>();
        for (int i = 0; i < job.Steps.Count; i++)
        {
            Rdv3ProcessStepDef step = job.Steps[i];
            object left = values[step.Target1];
            object right = (step.Target2.Length == 0) ? null : values[step.Target2];
            object output;
            if (step.Operation == "join")
            {
                Rdv3JoinResult joined;
                output = Join((Rdv3Relation)left, (Rdv3Relation)right, step, out joined);
                result.Joins.Add(joined);
            }
            else if (step.Operation == "append")
            {
                output = Append((Rdv3Relation)left, (Rdv3Relation)right);
            }
            else if (step.Operation == "extract")
            {
                output = Extract(left, right, step, result);
            }
            else if (step.Operation == "delete")
            {
                int deleted;
                output = Delete((Rdv3Relation)left, (Rdv3RowSelection)right, out deleted);
                result.Deleted += deleted;
            }
            else if (step.Operation == "update")
            {
                int changed;
                output = Update(data, (Rdv3Relation)left, (Rdv3RowSelection)right, step,
                                job.OnSourceChange, initialStored, directReset, out changed, result);
                directUpdated += changed;
            }
            else if (step.Operation == "select")
            {
                output = Select((Rdv3Relation)left, step);
            }
            else if (step.Operation == "calculate")
            {
                output = Calculate((Rdv3Relation)left, step, result);
            }
            else if (step.Operation == "aggregate")
            {
                output = Aggregate((Rdv3Relation)left, step, result);
            }
            else if (step.Operation == "sort")
            {
                output = Sort((Rdv3Relation)left, step, result);
            }
            else if (step.Operation == "distinct")
            {
                output = Distinct((Rdv3Relation)left, step);
            }
            else if (step.Operation == "merge" || step.Operation == "replace")
            {
                Rdv3UpdateResult update;
                output = WriteLedger(data, job, (Rdv3Relation)left, (Rdv3Relation)right,
                                     step, initialStored, out update, result);
                result.Update = update;
                result.Deleted += update.Deleted;
            }
            else { throw new InvalidOperationException(Rdv3Text.Format(Rdv3Text.ProcessUnknownOperation, step.Operation)); }
            Rdv3Relation outputTable = output as Rdv3Relation;
            if (outputTable != null && outputTable.Kind != "ledger")
            { output = ValidResultTypes(data, outputTable, step, result); }
            values[step.Output] = output;
            last = output;
        }

        Rdv3Relation finalLedger = last as Rdv3Relation;
        if (finalLedger != null && finalLedger.Kind == "ledger")
        {
            ValidateLedgerIdentity(data, job, finalLedger);
        }
        FillResult(result, last);
        if (result.Update == null && last is Rdv3Relation) { result.Update = new Rdv3UpdateResult(); }
        if (result.Update != null)
        {
            result.Update.Lines = result.Lines;
            result.Update.States = result.States;
            result.Update.Deleted = result.Deleted;
            result.Update.Updated += directUpdated;
            result.Update.ResetLines.AddRange(directReset);
        }
        if (capture) { Capture(values, result); }
        return result;
    }

    private static Rdv3Relation RelationOfTable(Rdv3ProcessInputDef input, Rdv3Table table)
    {
        Rdv3Relation relation = new Rdv3Relation();
        relation.Columns = new string[table.Head.Length];
        for (int c = 0; c < table.Head.Length; c++) { relation.Columns[c] = input.Table + "." + table.Head[c]; }
        for (int r = 0; r < table.Rows; r++)
        {
            string[] row = new string[table.Head.Length];
            for (int c = 0; c < row.Length; c++) { row[c] = table.Field(r, c); }
            relation.Rows.Add(row);
            relation.Origins.Add(Rdv3Text.Format(Rdv3Text.SourceRow, Path.GetFileName(table.Path), table.SourceRow(r)));
        }
        return relation;
    }

    private static Rdv3Relation RelationOfValues(Rdv3ProcessInputDef input, Rdv3Table table)
    {
        Rdv3Relation relation = new Rdv3Relation();
        relation.Columns = new string[] { input.Key };
        for (int r = 0; r < table.Rows; r++)
        {
            relation.Rows.Add(new string[] { table.Field(r, table.KeyCol) });
            relation.Origins.Add(Rdv3Text.Format(Rdv3Text.SourceRow, Path.GetFileName(table.Path), table.SourceRow(r)));
        }
        return relation;
    }

    private static Rdv3Relation RelationOfLedger(Rdv3Data data, string[] lines, string[] states,
                                                  string initialStored)
    {
        string[] content = lines ?? new string[0];
        string[] app = states;
        if (app == null)
        {
            app = new string[content.Length];
            for (int i = 0; i < app.Length; i++) { app[i] = initialStored; }
        }
        if (content.Length != app.Length)
        {
            throw new InvalidDataException(Rdv3Text.LedgerLengths);
        }
        Rdv3Relation relation = new Rdv3Relation();
        relation.Kind = "ledger";
        relation.Columns = data.ColumnRefs;
        relation.States = new List<string>(app);
        for (int i = 0; i < content.Length; i++)
        {
            string[] row = Rdv3Ledger.SplitLine(content[i]);
            if (row.Length != relation.Columns.Length)
            {
                throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.LedgerRowColumns, i + 2, relation.Columns.Length, row.Length));
            }
            relation.Rows.Add(row);
            relation.Origins.Add(Rdv3Text.Format(Rdv3Text.SourceRow, "ledger", i + 2));
        }
        return relation;
    }

    private static void ValidateLedgerIdentity(Rdv3Data data, Rdv3ProcessJobDef job, Rdv3Relation ledger)
    {
        ValidateIdentity(data, job, ledger, ledger.NeedColumns(data.IdentityRefs), data.IdentityRefs);
    }

    private static void ValidateIdentity(Rdv3Data data, Rdv3ProcessJobDef job, Rdv3Relation relation,
                                         int[] identity, string[] references)
    {
        string jobName = (job.Name.Length == 0) ? job.Id : job.Name;
        List<string> labels = new List<string>();
        foreach (string reference in references)
        { string label = data.LabelOf(reference); labels.Add(label.Length == 0 ? reference : label); }
        string column = string.Join(" / ", labels.ToArray());
        HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < relation.Rows.Count; i++)
        {
            string value = Rdv3Key.FromCells(relation.Rows[i], identity);
            if (value.Length == 0)
            {
                throw new InvalidDataException(Rdv3Text.ProcessBlankIdentity
                    .Replace("{job}", jobName).Replace("{column}", column));
            }
            if (!seen.Add(value))
            {
                throw new InvalidDataException(Rdv3Text.ProcessDuplicateIdentity
                    .Replace("{job}", jobName).Replace("{column}", column).Replace("{value}", value));
            }
        }
    }

    private static Rdv3Relation Join(Rdv3Relation left, Rdv3Relation right, Rdv3ProcessStepDef step, out Rdv3JoinResult stats)
    {
        stats = new Rdv3JoinResult { Output = step.Output, LeftRows = left.Rows.Count, RightRows = right.Rows.Count };
        int[] lc = left.NeedColumns(step.KeySide(0));
        int[] rc = right.NeedColumns(step.KeySide(1));
        string[] columns = new string[left.Columns.Length + right.Columns.Length];
        Array.Copy(left.Columns, 0, columns, 0, left.Columns.Length);
        for (int c = 0; c < right.Columns.Length; c++)
        {
            if (left.ColumnOf(right.Columns[c]) >= 0)
            {
                throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessDuplicateColumn, right.Columns[c]));
            }
            columns[left.Columns.Length + c] = right.Columns[c];
        }

        Dictionary<string, List<int>> index = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        for (int i = 0; i < right.Rows.Count; i++)
        {
            string key = Rdv3Key.FromCells(right.Rows[i], rc);
            if (key.Length == 0) { continue; }
            List<int> found;
            if (!index.TryGetValue(key, out found))
            {
                found = new List<int>();
                index.Add(key, found);
            }
            found.Add(i);
        }
        bool[] used = new bool[right.Rows.Count];
        Rdv3Relation output = new Rdv3Relation();
        output.Columns = columns;
        string[] blankRight = new string[right.Columns.Length];
        string[] blankLeft = new string[left.Columns.Length];
        for (int i = 0; i < blankRight.Length; i++) { blankRight[i] = ""; }
        for (int i = 0; i < blankLeft.Length; i++) { blankLeft[i] = ""; }
        for (int i = 0; i < left.Rows.Count; i++)
        {
            string key = Rdv3Key.FromCells(left.Rows[i], lc);
            List<int> found;
            if (key.Length > 0 && index.TryGetValue(key, out found))
            {
                for (int f = 0; f < found.Count; f++)
                {
                    used[found[f]] = true;
                    output.Rows.Add(Combine(left.Rows[i], right.Rows[found[f]]));
                    output.Origins.Add(left.Origin(i) + " / " + right.Origin(found[f]));
                }
            }
            else
            {
                stats.UnmatchedLeft++;
                if (step.Condition == "left" || step.Condition == "full")
                { output.Rows.Add(Combine(left.Rows[i], blankRight)); output.Origins.Add(left.Origin(i)); }
            }
        }
        for (int i = 0; i < used.Length; i++) { if (!used[i]) { stats.UnmatchedRight++; } }
        if (step.Condition == "full")
        {
            for (int i = 0; i < right.Rows.Count; i++)
            {
                if (!used[i]) { output.Rows.Add(Combine(blankLeft, right.Rows[i])); output.Origins.Add(right.Origin(i)); }
            }
        }
        stats.OutputRows = output.Rows.Count;
        return output;
    }

    private static string[] Combine(string[] left, string[] right)
    {
        string[] row = new string[left.Length + right.Length];
        Array.Copy(left, 0, row, 0, left.Length);
        Array.Copy(right, 0, row, left.Length, right.Length);
        return row;
    }

    private static Rdv3Relation Append(Rdv3Relation left, Rdv3Relation right)
    {
        if (left.Columns.Length != right.Columns.Length)
        {
            throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessAppendCount, left.Columns.Length, right.Columns.Length));
        }
        for (int c = 0; c < left.Columns.Length; c++)
        {
            if (ColumnName(left.Columns[c]) != ColumnName(right.Columns[c]))
            {
                throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessAppendColumn, c + 1, left.Columns[c], right.Columns[c]));
            }
        }
        Rdv3Relation output = new Rdv3Relation();
        output.Columns = (string[])left.Columns.Clone();
        for (int i = 0; i < left.Rows.Count; i++) { output.Rows.Add((string[])left.Rows[i].Clone()); output.Origins.Add(left.Origin(i)); }
        for (int i = 0; i < right.Rows.Count; i++) { output.Rows.Add((string[])right.Rows[i].Clone()); output.Origins.Add(right.Origin(i)); }
        return output;
    }

    private static object Extract(object leftValue, object rightValue, Rdv3ProcessStepDef step, Rdv3ProcessResult result)
    {
        Rdv3Relation left = leftValue as Rdv3Relation;
        Rdv3Relation right = rightValue as Rdv3Relation;
        if (left != null && step.Where != null)
        {
            int column = left.NeedColumn(step.Where.Column);
            Rdv3RowSelection selected = new Rdv3RowSelection();
            selected.Table = left;
            for (int i = 0; i < left.Rows.Count; i++)
            {
                try { if (Matches(left.Rows[i][column], step.Where)) { selected.Rows.Add(i); } }
                catch (Rdv3RecordError error) { Exclude(result, left, i, step, error.Message); }
            }
            return selected;
        }
        if (left != null && right != null)
        {
            int[] lc = left.NeedColumns(step.KeySide(0));
            int[] rc = right.NeedColumns(step.KeySide(1));
            HashSet<string> wanted = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < right.Rows.Count; i++) { wanted.Add(Rdv3Key.FromCells(right.Rows[i], rc)); }
            Rdv3RowSelection selected = new Rdv3RowSelection();
            selected.Table = left;
            for (int i = 0; i < left.Rows.Count; i++)
            {
                string key = Rdv3Key.FromCells(left.Rows[i], lc);
                bool found = (lc.Length == 1 || key.Length > 0) && wanted.Contains(key);
                if ((step.Condition == "match" && found) || (step.Condition == "exclude" && !found))
                {
                    selected.Rows.Add(i);
                }
            }
            return selected;
        }

        Rdv3RowSelection a = (Rdv3RowSelection)leftValue;
        Rdv3RowSelection b = (Rdv3RowSelection)rightValue;
        if (!object.ReferenceEquals(a.Table, b.Table))
        {
            throw new InvalidDataException(Rdv3Text.ProcessSelectionSource);
        }
        Rdv3RowSelection combined = new Rdv3RowSelection();
        combined.Table = a.Table;
        combined.Rows = new HashSet<int>(a.Rows);
        if (step.Condition == "either") { combined.Rows.UnionWith(b.Rows); }
        else if (step.Condition == "both") { combined.Rows.IntersectWith(b.Rows); }
        else { combined.Rows.ExceptWith(b.Rows); }
        return combined;
    }

    private static bool Matches(string value, Rdv3ProcessWhereDef where)
    {
        if (where.Operator == "equals") { return value == where.Value; }
        if (where.Operator == "notEquals") { return value != where.Value; }
        if (where.Operator == "contains") { return value.IndexOf(where.Value, StringComparison.Ordinal) >= 0; }
        if (where.Operator == "startsWith") { return value.StartsWith(where.Value, StringComparison.Ordinal); }
        if (where.Operator == "endsWith") { return value.EndsWith(where.Value, StringComparison.Ordinal); }
        if (where.Operator == "empty") { return value.Length == 0; }
        if (where.Operator == "notEmpty") { return value.Length > 0; }
        if (value.Length == 0) { return false; }
        decimal left;
        decimal right;
        if (!Rdv3Input.TryNumber(value, out left)
            || !Rdv3Input.TryNumber(where.Value, out right))
        {
            throw new Rdv3RecordError(Rdv3Text.Format(Rdv3Text.RecordNumber, Rdv3Input.Display(value)));
        }
        if (where.Operator == "greater") { return left > right; }
        if (where.Operator == "atLeast") { return left >= right; }
        if (where.Operator == "less") { return left < right; }
        return left <= right;
    }

    private static Rdv3Relation Delete(Rdv3Relation source, Rdv3RowSelection selected, out int deleted)
    {
        if (!object.ReferenceEquals(source, selected.Table))
        {
            throw new InvalidDataException(Rdv3Text.ProcessSelectionSource);
        }
        Rdv3Relation output = NewLike(source);
        for (int i = 0; i < source.Rows.Count; i++)
        {
            if (selected.Rows.Contains(i)) { continue; }
            output.Rows.Add((string[])source.Rows[i].Clone());
            output.Origins.Add(source.Origin(i));
            if (output.States != null) { output.States.Add(source.States[i]); }
        }
        deleted = selected.Rows.Count;
        return output;
    }

    private static Rdv3Relation Update(Rdv3Data data, Rdv3Relation source, Rdv3RowSelection selected,
                                       Rdv3ProcessStepDef step, string onSourceChange,
                                       string initialStored, List<string> resetLines,
                                       out int changed, Rdv3ProcessResult result)
    {
        if (!object.ReferenceEquals(source, selected.Table))
        {
            throw new InvalidDataException(Rdv3Text.ProcessSelectionSource);
        }
        int[] columns = new int[step.Set.Count];
        Rdv3Expression[] expressions = new Rdv3Expression[step.Set.Count];
        for (int i = 0; i < step.Set.Count; i++)
        {
            columns[i] = source.NeedColumn(step.Set[i].Column);
            for (int k = 0; k < i; k++)
            {
                if (columns[k] == columns[i])
                {
                    throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessDuplicateColumn, step.Set[i].Column));
                }
            }
            expressions[i] = Rdv3Expression.Compile(step.Set[i].Expression, source.Columns);
        }
        Rdv3Relation output = NewLike(source);
        changed = 0;
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = (string[])source.Rows[r].Clone();
            bool failed = false;
            if (selected.Rows.Contains(r))
            {
                for (int i = 0; i < columns.Length; i++)
                {
                    try
                    {
                        row[columns[i]] = Evaluate(expressions[i], source.Rows[r]);
                        Rdv3ColumnTypeDef type = data.TypeOf(source.Columns[columns[i]]);
                        if (type != null && type.Type != "text" && row[columns[i]].Length > 0)
                        { CheckTypedResult(type, row[columns[i]], source.Origin(r)); }
                    }
                    catch (Rdv3RecordError error) { Exclude(result, source, r, step, error.Message); failed = true; break; }
                }
            }
            // A failed update never deletes an existing ledger record or
            // leaves half of a multi-column assignment in that record.
            if (failed) { if (source.Kind != "ledger") { continue; } row = (string[])source.Rows[r].Clone(); }
            bool rowChanged = !SameRow(source.Rows[r], row);
            if (rowChanged) { changed++; }
            output.Rows.Add(row);
            output.Origins.Add(source.Origin(r));
            if (output.States != null)
            {
                string state = source.States[r];
                if (rowChanged && onSourceChange == "reset")
                {
                    state = initialStored;
                    if (!string.Equals(source.States[r], initialStored, StringComparison.Ordinal))
                    {
                        resetLines.Add(string.Join("\t", row));
                    }
                }
                output.States.Add(state);
            }
        }
        return output;
    }

    private static bool SameRow(string[] left, string[] right)
    {
        if (left.Length != right.Length) { return false; }
        for (int i = 0; i < left.Length; i++)
        {
            if (!string.Equals(left[i], right[i], StringComparison.Ordinal)) { return false; }
        }
        return true;
    }

    private static Rdv3Relation Select(Rdv3Relation source, Rdv3ProcessStepDef step)
    {
        Rdv3Relation output = new Rdv3Relation();
        output.Columns = new string[step.Columns.Count];
        int[] fields = new int[step.Columns.Count];
        for (int i = 0; i < step.Columns.Count; i++)
        {
            fields[i] = source.NeedColumn(step.Columns[i].Column);
            output.Columns[i] = step.Columns[i].OutputRef(step.Output);
            for (int k = 0; k < i; k++)
            {
                if (output.Columns[k] == output.Columns[i])
                {
                    throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessDuplicateColumn, output.Columns[i]));
                }
            }
        }
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = new string[fields.Length];
            for (int i = 0; i < fields.Length; i++) { row[i] = source.Rows[r][fields[i]]; }
            output.Rows.Add(row);
            output.Origins.Add(source.Origin(r));
        }
        return output;
    }

    private static Rdv3Relation Calculate(Rdv3Relation source, Rdv3ProcessStepDef step, Rdv3ProcessResult result)
    {
        string added = step.Output + "." + step.Column;
        if (source.ColumnOf(added) >= 0) { throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessDuplicateColumn, added)); }
        Rdv3Expression expression = Rdv3Expression.Compile(step.Expression, source.Columns);
        Rdv3Relation output = new Rdv3Relation();
        output.Columns = new string[source.Columns.Length + 1];
        Array.Copy(source.Columns, output.Columns, source.Columns.Length);
        output.Columns[source.Columns.Length] = added;
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = new string[output.Columns.Length];
            Array.Copy(source.Rows[r], row, source.Columns.Length);
            try { row[source.Columns.Length] = Evaluate(expression, source.Rows[r]); }
            catch (Rdv3RecordError error) { Exclude(result, source, r, step, error.Message); continue; }
            output.Rows.Add(row);
            output.Origins.Add(source.Origin(r));
        }
        return output;
    }

    private sealed class GroupValue
    {
        public string[] Keys;
        public decimal[] Sums;
        public int Count;
    }

    private static Rdv3Relation Aggregate(Rdv3Relation source, Rdv3ProcessStepDef step, Rdv3ProcessResult result)
    {
        int[] groups = new int[step.GroupBy.Count];
        for (int i = 0; i < groups.Length; i++) { groups[i] = source.NeedColumn(step.GroupBy[i]); }
        int[] fields = new int[step.Aggregates.Count];
        for (int i = 0; i < fields.Length; i++)
        {
            fields[i] = (step.Aggregates[i].Function == "sum")
                ? source.NeedColumn(step.Aggregates[i].Column) : -1;
        }

        Dictionary<string, GroupValue> byKey = new Dictionary<string, GroupValue>(StringComparer.Ordinal);
        List<GroupValue> ordered = new List<GroupValue>();
        if (groups.Length == 0)
        {
            GroupValue empty = NewGroup(new string[0], step.Aggregates.Count);
            byKey.Add("", empty);
            ordered.Add(empty);
        }
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] keys = new string[groups.Length];
            for (int g = 0; g < groups.Length; g++) { keys[g] = source.Rows[r][groups[g]]; }
            string key = Composite(keys);
            GroupValue value;
            bool exists = byKey.TryGetValue(key, out value);
            if (!exists)
            {
                value = NewGroup(keys, step.Aggregates.Count);
            }
            decimal[] sums = (decimal[])value.Sums.Clone();
            bool invalid = false;
            for (int a = 0; a < step.Aggregates.Count; a++)
            {
                if (fields[a] < 0) { continue; }
                decimal number;
                if (!Rdv3Input.TryNumber(source.Rows[r][fields[a]], out number))
                {
                    Exclude(result, source, r, step, Rdv3Text.Format(Rdv3Text.RecordNumber, Rdv3Input.Display(source.Rows[r][fields[a]])));
                    invalid = true; break;
                }
                try { sums[a] += number; }
                catch (OverflowException) { Exclude(result, source, r, step, Rdv3Text.RecordOverflow); invalid = true; break; }
            }
            // Commit every aggregate together: one invalid cell must not
            // increment count or any earlier sum for this same record.
            if (invalid) { continue; }
            value.Sums = sums;
            value.Count++;
            if (!exists) { byKey.Add(key, value); ordered.Add(value); }
        }

        Rdv3Relation output = new Rdv3Relation();
        output.Columns = new string[groups.Length + step.Aggregates.Count];
        for (int g = 0; g < groups.Length; g++) { output.Columns[g] = step.GroupBy[g]; }
        for (int a = 0; a < step.Aggregates.Count; a++)
        {
            output.Columns[groups.Length + a] = step.Output + "." + step.Aggregates[a].As;
        }
        for (int i = 0; i < output.Columns.Length; i++)
        {
            for (int k = 0; k < i; k++)
            {
                if (output.Columns[k] == output.Columns[i])
                {
                    throw new InvalidDataException(Rdv3Text.Format(Rdv3Text.ProcessDuplicateColumn, output.Columns[i]));
                }
            }
        }
        for (int i = 0; i < ordered.Count; i++)
        {
            GroupValue value = ordered[i];
            string[] row = new string[output.Columns.Length];
            Array.Copy(value.Keys, row, value.Keys.Length);
            for (int a = 0; a < step.Aggregates.Count; a++)
            {
                row[groups.Length + a] = (step.Aggregates[a].Function == "count")
                    ? value.Count.ToString(CultureInfo.InvariantCulture)
                    : value.Sums[a].ToString("G29", CultureInfo.InvariantCulture);
            }
            output.Rows.Add(row);
            output.Origins.Add("groupBy " + Rdv3Input.Display(string.Join(" / ", value.Keys)));
        }
        return output;
    }

    private static GroupValue NewGroup(string[] keys, int aggregates)
    {
        GroupValue value = new GroupValue();
        value.Keys = keys;
        value.Sums = new decimal[aggregates];
        return value;
    }

    private sealed class SortValue
    {
        public string[] Row;
        public int Ord;
    }

    private static Rdv3Relation Sort(Rdv3Relation source, Rdv3ProcessStepDef step, Rdv3ProcessResult result)
    {
        int[] fields = new int[step.Orders.Count];
        for (int i = 0; i < fields.Length; i++) { fields[i] = source.NeedColumn(step.Orders[i].Column); }
        List<SortValue> rows = new List<SortValue>();
        for (int i = 0; i < source.Rows.Count; i++)
        {
            bool invalid = false;
            for (int c = 0; c < fields.Length; c++)
            {
                string cell = source.Rows[i][fields[c]];
                decimal number;
                if (step.Orders[c].Type == "number" && cell.Length > 0 && !Rdv3Input.TryNumber(cell, out number))
                {
                    string reason = Rdv3Text.Format(Rdv3Text.RecordNumber, Rdv3Input.Display(cell));
                    // A sort can directly replace ledger. Never turn a bad
                    // persisted value into an implicit deletion of that row.
                    if (source.Kind == "ledger")
                    { throw new InvalidDataException(source.Origin(i) + ": " + step.Orders[c].Column + ": " + reason); }
                    Exclude(result, source, i, step, reason); invalid = true; break;
                }
            }
            if (invalid) { continue; }
            SortValue value = new SortValue();
            value.Row = source.Rows[i];
            value.Ord = i;
            rows.Add(value);
        }
        rows.Sort(delegate(SortValue a, SortValue b)
        {
            for (int i = 0; i < fields.Length; i++)
            {
                int cmp = Compare(a.Row[fields[i]], b.Row[fields[i]], step.Orders[i].Type,
                                  step.Orders[i].Direction);
                if (cmp != 0) { return cmp; }
            }
            return a.Ord.CompareTo(b.Ord);
        });
        Rdv3Relation output = NewLike(source);
        for (int i = 0; i < rows.Count; i++)
        {
            output.Rows.Add((string[])rows[i].Row.Clone());
            output.Origins.Add(source.Origin(rows[i].Ord));
            if (output.States != null) { output.States.Add(source.States[rows[i].Ord]); }
        }
        return output;
    }

    private static int Compare(string left, string right, string type, string direction)
    {
        if (type == "text")
        {
            int text = string.Compare(left, right, StringComparison.Ordinal);
            return (direction == "descending") ? -text : text;
        }
        bool leftBlank = left.Length == 0;
        bool rightBlank = right.Length == 0;
        if (leftBlank || rightBlank)
        {
            if (leftBlank && rightBlank) { return 0; }
            return leftBlank ? 1 : -1;
        }
        decimal a;
        decimal b;
        if (!Rdv3Input.TryNumber(left, out a)
            || !Rdv3Input.TryNumber(right, out b))
        {
            throw new Rdv3RecordError(Rdv3Text.Format(Rdv3Text.RecordNumber, Rdv3Input.Display(left + " / " + right)));
        }
        int number = a.CompareTo(b);
        return (direction == "descending") ? -number : number;
    }

    private static Rdv3Relation Distinct(Rdv3Relation source, Rdv3ProcessStepDef step)
    {
        int[] fields = new int[step.Columns.Count];
        for (int i = 0; i < fields.Length; i++) { fields[i] = source.NeedColumn(step.Columns[i].Column); }
        HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
        Rdv3Relation output = NewLike(source);
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] key = new string[fields.Length];
            for (int i = 0; i < fields.Length; i++) { key[i] = source.Rows[r][fields[i]]; }
            if (seen.Add(Composite(key)))
            {
                output.Rows.Add((string[])source.Rows[r].Clone());
                output.Origins.Add(source.Origin(r));
                if (output.States != null) { output.States.Add(source.States[r]); }
            }
        }
        return output;
    }

    private static Rdv3Relation WriteLedger(Rdv3Data data, Rdv3ProcessJobDef job,
                                            Rdv3Relation source, Rdv3Relation target,
                                            Rdv3ProcessStepDef step, string initialStored,
                                            out Rdv3UpdateResult update, Rdv3ProcessResult result)
    {
        int[] targetKey = target.NeedColumns(step.KeySide(1));
        if (!SameColumns(targetKey, data.IdentityCols))
        {
            throw new InvalidDataException(Rdv3Text.ProcessLedgerKey);
        }
        int[] sourceKey = source.NeedColumns(step.KeySide(0));
        source = ValidSourceIdentity(data, job, source, sourceKey, step, result);
        // The column map does not change within the step: resolved once here,
        // not by a name search for every cell of every row.
        int[] from = new int[data.Columns.Count];
        Rdv3ColumnTypeDef[] typed = new Rdv3ColumnTypeDef[data.Columns.Count];
        for (int c = 0; c < data.Columns.Count; c++)
        {
            int identityPart = Array.IndexOf(data.IdentityCols, c);
            from[c] = identityPart >= 0 ? sourceKey[identityPart] : source.ColumnOf(data.Columns[c].Ref);
            // ValidateColumns proves this before any row is read; a blank
            // written here would be a silent loss, never an acceptable value.
            if (from[c] < 0)
            {
                throw new InvalidDataException(Rdv3Text.LedgerColumnNotProduced
                    .Replace("{name}", data.Columns[c].Ref).Replace("{step}", step.Operation + " " + step.Target1));
            }
            // Input columns met their declared type when their file was read. A
            // column the job makes has no file, so its values are checked here,
            // before anything reaches the ledger.
            Rdv3ColumnTypeDef type = data.TypeOf(data.Columns[c].Ref);
            typed[c] = (type != null && type.TableOrd < 0) ? type : null;
        }
        List<string> sourceLines = new List<string>();
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = new string[data.Columns.Count];
            bool invalid = false;
            for (int c = 0; c < data.Columns.Count; c++)
            {
                string value = source.Rows[r][from[c]] ?? "";
                if (typed[c] != null && value.Length > 0)
                {
                    try { CheckTypedResult(typed[c], value, Rdv3Key.FromCells(source.Rows[r], sourceKey)); }
                    catch (Rdv3RecordError error) { Exclude(result, source, r, step, error.Message); invalid = true; break; }
                }
                row[c] = value;
            }
            if (!invalid) { sourceLines.Add(string.Join("\t", row)); }
        }
        Rdv3ProcessJobDef apply = new Rdv3ProcessJobDef();
        apply.ApplyStep = step;
        apply.OnSourceChange = job.OnSourceChange;
        string[] targetStates = (target.States == null) ? new string[target.Rows.Count] : target.States.ToArray();
        update = Rdv3Ledger.ApplyUpdate(apply, target.ToLines(), targetStates,
                                       sourceLines.ToArray(), data.IdentityCols, initialStored);
        Rdv3Relation output = RelationOfLedger(data, update.Lines, update.States, initialStored);
        return output;
    }

    private static void CheckTypedResult(Rdv3ColumnTypeDef type, string value, string identity)
    {
        DateTime date;
        decimal number;
        bool valid = (type.Type == "date") ? type.TryDate(value, out date) : type.TryNumber(value, out number);
        if (valid) { return; }
        string displayType = (type.Type == "date")
            ? Rdv3Text.TypeDateFormat.Replace("{format}", type.Format) : Rdv3Text.TypeNumber;
        throw new Rdv3RecordError(Rdv3Text.DataTypedResult.Replace("{name}", type.Ref).Replace("{identity}", identity)
            .Replace("{value}", value).Replace("{type}", displayType) + Rdv3Text.InputFixType.Replace("{ref}", type.Ref));
    }

    private static Rdv3Relation NewLike(Rdv3Relation source)
    {
        Rdv3Relation output = new Rdv3Relation();
        output.Kind = source.Kind;
        output.Columns = (string[])source.Columns.Clone();
        if (source.States != null) { output.States = new List<string>(); }
        return output;
    }

    private static string Evaluate(Rdv3Expression expression, string[] row)
    {
        try { return expression.Evaluate(row); }
        catch (OverflowException) { throw new Rdv3RecordError(Rdv3Text.RecordOverflow); }
        catch (RegexMatchTimeoutException) { throw new Rdv3RecordError(Rdv3Text.RecordRegexTimeout); }
    }

    private static void Exclude(Rdv3ProcessResult result, Rdv3Relation source, int row,
                                Rdv3ProcessStepDef step, string reason)
    {
        List<string> cells = new List<string>();
        for (int c = 0; c < source.Columns.Length; c++)
        { cells.Add(source.Columns[c] + "=" + Rdv3Input.Display(source.Rows[row][c])); }
        string where = Rdv3Text.Format(Rdv3Text.RecordStep, step.Operation, step.Target1, step.Output, source.Origin(row));
        result.Warnings.Add(Rdv3Text.Format(Rdv3Text.RecordExcluded, where,
            reason + " " + Rdv3Text.Format(Rdv3Text.RecordValues, string.Join(" / ", cells.ToArray()))));
        result.SkippedInvalid++;
    }

    private static Rdv3Relation ValidResultTypes(Rdv3Data data, Rdv3Relation source,
                                                 Rdv3ProcessStepDef step, Rdv3ProcessResult result)
    {
        List<int> columns = new List<int>();
        List<Rdv3ColumnTypeDef> types = new List<Rdv3ColumnTypeDef>();
        for (int c = 0; c < source.Columns.Length; c++)
        {
            Rdv3ColumnTypeDef type = data.TypeOf(source.Columns[c]);
            if (type != null && type.Type != "text") { columns.Add(c); types.Add(type); }
        }
        if (columns.Count == 0) { return source; }
        Rdv3Relation output = NewLike(source);
        for (int r = 0; r < source.Rows.Count; r++)
        {
            List<string> errors = new List<string>();
            for (int c = 0; c < columns.Count; c++)
            {
                string value = source.Rows[r][columns[c]];
                if (value.Length == 0) { continue; }
                try { CheckTypedResult(types[c], value, source.Origin(r)); }
                catch (Rdv3RecordError error) { errors.Add(error.Message); }
            }
            if (errors.Count > 0) { Exclude(result, source, r, step, string.Join(" / ", errors.ToArray())); continue; }
            output.Rows.Add(source.Rows[r]); output.Origins.Add(source.Origin(r));
            if (output.States != null) { output.States.Add(source.States[r]); }
        }
        return output;
    }

    private static Rdv3Relation ValidSourceIdentity(Rdv3Data data, Rdv3ProcessJobDef job,
        Rdv3Relation source, int[] identity, Rdv3ProcessStepDef step, Rdv3ProcessResult result)
    {
        Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        HashSet<int> excluded = new HashSet<int>();
        string column = string.Join(" / ", step.KeySide(0));
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string key = Rdv3Key.FromCells(source.Rows[r], identity);
            if (key.Length == 0)
            {
                Exclude(result, source, r, step, Rdv3Text.ProcessBlankIdentity.Replace("{job}", job.Name).Replace("{column}", column));
                excluded.Add(r); continue;
            }
            List<int> group;
            if (!groups.TryGetValue(key, out group)) { group = new List<int>(); groups.Add(key, group); }
            group.Add(r);
        }
        foreach (KeyValuePair<string, List<int>> pair in groups)
        {
            List<int> group = pair.Value;
            if (group.Count < 2) { continue; }
            bool conflict = false;
            for (int i = 1; i < group.Count; i++)
            { if (!SameRow(source.Rows[group[0]], source.Rows[group[i]])) { conflict = true; break; } }
            for (int i = conflict ? 0 : 1; i < group.Count; i++)
            {
                Exclude(result, source, group[i], step,
                    Rdv3Text.ProcessDuplicateIdentity.Replace("{job}", job.Name).Replace("{column}", column).Replace("{value}", Rdv3Input.Display(pair.Key))
                    + " " + (conflict ? Rdv3Text.RecordConflict : Rdv3Text.RecordIdentical));
                excluded.Add(group[i]);
            }
        }
        if (excluded.Count == 0) { return source; }
        Rdv3Relation output = NewLike(source);
        for (int r = 0; r < source.Rows.Count; r++)
        {
            if (excluded.Contains(r)) { continue; }
            output.Rows.Add(source.Rows[r]); output.Origins.Add(source.Origin(r));
            if (output.States != null) { output.States.Add(source.States[r]); }
        }
        return output;
    }

    private static bool SameColumns(int[] a, int[] b)
    {
        if (a.Length != b.Length) { return false; }
        for (int i = 0; i < a.Length; i++) { if (a[i] != b[i]) { return false; } }
        return true;
    }

    private static string ColumnName(string reference)
    {
        int dot = reference.IndexOf('.');
        return (dot < 0) ? reference : reference.Substring(dot + 1);
    }

    private static string Composite(string[] values)
    {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < values.Length; i++)
        {
            string value = values[i] ?? "";
            sb.Append(value.Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(value);
        }
        return sb.ToString();
    }

    private static void FillResult(Rdv3ProcessResult result, object value)
    {
        Rdv3Relation table = value as Rdv3Relation;
        if (table != null)
        {
            result.Kind = table.Kind;
            result.Columns = (string[])table.Columns.Clone();
            result.Lines = table.ToLines();
            result.States = (table.States == null) ? new string[0] : table.States.ToArray();
            return;
        }
        Rdv3RowSelection rows = value as Rdv3RowSelection;
        if (rows != null)
        {
            FillResult(result, SelectedTable(rows));
            result.Kind = "rows";
        }
    }

    private static Rdv3Relation SelectedTable(Rdv3RowSelection selection)
    {
        Rdv3Relation table = NewLike(selection.Table);
        List<int> indices = new List<int>(selection.Rows);
        indices.Sort();
        foreach (int row in indices)
        {
            table.Rows.Add(selection.Table.Rows[row]);
            table.Origins.Add(selection.Table.Origin(row));
            if (table.States != null) { table.States.Add(selection.Table.States[row]); }
        }
        return table;
    }

    private static void Capture(Dictionary<string, object> values, Rdv3ProcessResult result)
    {
        foreach (KeyValuePair<string, object> pair in values)
        {
            Rdv3ProcessValueResult snapshot = new Rdv3ProcessValueResult();
            Rdv3Relation table = pair.Value as Rdv3Relation;
            if (table != null)
            {
                snapshot.Kind = table.Kind;
                snapshot.Columns = (string[])table.Columns.Clone();
                snapshot.Lines = table.ToLines();
                snapshot.Count = table.Rows.Count;
            }
            else
            {
                Rdv3RowSelection rows = (Rdv3RowSelection)pair.Value;
                snapshot.Kind = "rows";
                snapshot.Count = rows.Rows.Count;
                snapshot.Columns = (string[])rows.Table.Columns.Clone();
                snapshot.Lines = SelectedTable(rows).ToLines();
            }
            result.Values[pair.Key] = snapshot;
        }
    }

    public static string DisplayExpression(Rdv3Data data, string expression)
    {
        StringBuilder sb = new StringBuilder();
        int p = 0;
        while (p < expression.Length)
        {
            char ch = expression[p];
            if (ch == '\'')
            {
                int start = p++;
                while (p < expression.Length)
                {
                    if (expression[p++] == '\'' && (p >= expression.Length || expression[p] != '\'')) { break; }
                    if (p < expression.Length && expression[p - 1] == '\'' && expression[p] == '\'') { p++; }
                }
                sb.Append(expression.Substring(start, p - start));
            }
            else if (char.IsWhiteSpace(ch) || ch == '+' || ch == '-' || ch == '*'
                     || ch == '/' || ch == '(' || ch == ')' || ch == ',')
            {
                sb.Append(ch);
                p++;
            }
            else
            {
                int start = p;
                while (p < expression.Length && !char.IsWhiteSpace(expression[p])
                       && "+-*/(),".IndexOf(expression[p]) < 0) { p++; }
                string token = expression.Substring(start, p - start);
                if (Rdv3Expression.IsFunctionName(token)) { sb.Append(token); }
                else
                {
                    string label = data.LabelOf(token);
                    sb.Append(label.Length == 0 ? token : label);
                }
            }
        }
        return sb.ToString();
    }
}
