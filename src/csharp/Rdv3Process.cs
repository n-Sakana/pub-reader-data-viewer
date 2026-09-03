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

public sealed class Rdv3ProcessValueResult
{
    public string Kind = "";
    public string[] Columns = new string[0];
    public string[] Lines = new string[0];
    public int Count;
}

public sealed class Rdv3ProcessResult
{
    public string Kind = "";
    public string[] Columns = new string[0];
    public string[] Lines = new string[0];
    public string[] States = new string[0];
    public int Deleted;
    public Rdv3UpdateResult Update;
    public readonly Dictionary<string, Rdv3ProcessValueResult> Values
        = new Dictionary<string, Rdv3ProcessValueResult>(StringComparer.Ordinal);

    public Rdv3ProcessValueResult ValueOf(string name)
    {
        Rdv3ProcessValueResult value;
        return Values.TryGetValue(name, out value) ? value : null;
    }
}

internal sealed class Rdv3Relation
{
    public string Kind = "table";
    public string[] Columns = new string[0];
    public List<string[]> Rows = new List<string[]>();
    public List<string> States;

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
        if (column < 0) { throw new InvalidDataException("table has no column " + name); }
        return column;
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
                table = Rdv3Table.Read(path, input.Id, data.Enc, input.Column);
                new Rdv3Index(table);                    // the key must be unique
            }
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

    public static void ValidateColumns(Rdv3Data data, string[][] heads)
    {
        for (int i = 0; i < data.Jobs.Count; i++)
        {
            Rdv3PreparedProcess prepared = PrepareFromHeads(data, data.Jobs[i], heads);
            Execute(prepared, new string[0], new string[0], "", false);
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
                output = Join((Rdv3Relation)left, (Rdv3Relation)right, step);
            }
            else if (step.Operation == "append")
            {
                output = Append((Rdv3Relation)left, (Rdv3Relation)right);
            }
            else if (step.Operation == "extract")
            {
                output = Extract(left, right, step);
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
                output = Update((Rdv3Relation)left, (Rdv3RowSelection)right, step,
                                job.OnSourceChange, initialStored, directReset, out changed);
                directUpdated += changed;
            }
            else if (step.Operation == "select")
            {
                output = Select((Rdv3Relation)left, step);
            }
            else if (step.Operation == "calculate")
            {
                output = Calculate((Rdv3Relation)left, step);
            }
            else if (step.Operation == "aggregate")
            {
                output = Aggregate((Rdv3Relation)left, step);
            }
            else if (step.Operation == "sort")
            {
                output = Sort((Rdv3Relation)left, step);
            }
            else if (step.Operation == "distinct")
            {
                output = Distinct((Rdv3Relation)left, step);
            }
            else if (step.Operation == "merge" || step.Operation == "replace")
            {
                Rdv3UpdateResult update;
                output = WriteLedger(data, job, (Rdv3Relation)left, (Rdv3Relation)right,
                                     step, initialStored, out update);
                result.Update = update;
                result.Deleted += update.Deleted;
            }
            else { throw new InvalidOperationException("unknown operation " + step.Operation); }
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
            throw new InvalidDataException("ledger content and application-column lengths differ");
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
                throw new InvalidDataException("ledger row has a different column count");
            }
            relation.Rows.Add(row);
        }
        return relation;
    }

    private static void ValidateLedgerIdentity(Rdv3Data data, Rdv3ProcessJobDef job, Rdv3Relation ledger)
    {
        string reference = data.Columns[data.IdentityCol].Ref;
        int identity = ledger.NeedColumn(reference);
        ValidateIdentity(data, job, ledger, identity, reference);
    }

    private static void ValidateIdentity(Rdv3Data data, Rdv3ProcessJobDef job, Rdv3Relation relation,
                                         int identity, string reference)
    {
        string jobName = (job.Name.Length == 0) ? job.Id : job.Name;
        string column = data.LabelOf(reference);
        if (column.Length == 0) { column = reference; }
        HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < relation.Rows.Count; i++)
        {
            string value = relation.Rows[i][identity];
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

    private static Rdv3Relation Join(Rdv3Relation left, Rdv3Relation right, Rdv3ProcessStepDef step)
    {
        int lc = left.NeedColumn(step.Keys[0]);
        int rc = right.NeedColumn(step.Keys[1]);
        string[] columns = new string[left.Columns.Length + right.Columns.Length];
        Array.Copy(left.Columns, 0, columns, 0, left.Columns.Length);
        for (int c = 0; c < right.Columns.Length; c++)
        {
            if (left.ColumnOf(right.Columns[c]) >= 0)
            {
                throw new InvalidDataException("join would duplicate column " + right.Columns[c]);
            }
            columns[left.Columns.Length + c] = right.Columns[c];
        }

        Dictionary<string, List<int>> index = new Dictionary<string, List<int>>(StringComparer.Ordinal);
        for (int i = 0; i < right.Rows.Count; i++)
        {
            string key = right.Rows[i][rc];
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
            string key = left.Rows[i][lc];
            List<int> found;
            if (key.Length > 0 && index.TryGetValue(key, out found))
            {
                for (int f = 0; f < found.Count; f++)
                {
                    used[found[f]] = true;
                    output.Rows.Add(Combine(left.Rows[i], right.Rows[found[f]]));
                }
            }
            else if (step.Condition == "left" || step.Condition == "full")
            {
                output.Rows.Add(Combine(left.Rows[i], blankRight));
            }
        }
        if (step.Condition == "full")
        {
            for (int i = 0; i < right.Rows.Count; i++)
            {
                if (!used[i]) { output.Rows.Add(Combine(blankLeft, right.Rows[i])); }
            }
        }
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
            throw new InvalidDataException("append requires the same number of columns");
        }
        for (int c = 0; c < left.Columns.Length; c++)
        {
            if (ColumnName(left.Columns[c]) != ColumnName(right.Columns[c]))
            {
                throw new InvalidDataException("append columns differ at " + (c + 1).ToString(CultureInfo.InvariantCulture));
            }
        }
        Rdv3Relation output = new Rdv3Relation();
        output.Columns = (string[])left.Columns.Clone();
        for (int i = 0; i < left.Rows.Count; i++) { output.Rows.Add((string[])left.Rows[i].Clone()); }
        for (int i = 0; i < right.Rows.Count; i++) { output.Rows.Add((string[])right.Rows[i].Clone()); }
        return output;
    }

    private static object Extract(object leftValue, object rightValue, Rdv3ProcessStepDef step)
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
                if (Matches(left.Rows[i][column], step.Where)) { selected.Rows.Add(i); }
            }
            return selected;
        }
        if (left != null && right != null)
        {
            int lc = left.NeedColumn(step.Keys[0]);
            int rc = right.NeedColumn(step.Keys[1]);
            HashSet<string> wanted = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < right.Rows.Count; i++) { wanted.Add(right.Rows[i][rc]); }
            Rdv3RowSelection selected = new Rdv3RowSelection();
            selected.Table = left;
            for (int i = 0; i < left.Rows.Count; i++)
            {
                bool found = wanted.Contains(left.Rows[i][lc]);
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
            throw new InvalidDataException("row sets come from different table values");
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
        if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out left)
            || !decimal.TryParse(where.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out right))
        {
            throw new InvalidDataException("numeric row condition received non-numeric text");
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
            throw new InvalidDataException("delete rows come from a different table value");
        }
        Rdv3Relation output = NewLike(source);
        for (int i = 0; i < source.Rows.Count; i++)
        {
            if (selected.Rows.Contains(i)) { continue; }
            output.Rows.Add((string[])source.Rows[i].Clone());
            if (output.States != null) { output.States.Add(source.States[i]); }
        }
        deleted = selected.Rows.Count;
        return output;
    }

    private static Rdv3Relation Update(Rdv3Relation source, Rdv3RowSelection selected,
                                       Rdv3ProcessStepDef step, string onSourceChange,
                                       string initialStored, List<string> resetLines,
                                       out int changed)
    {
        if (!object.ReferenceEquals(source, selected.Table))
        {
            throw new InvalidDataException("update rows come from a different table value");
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
                    throw new InvalidDataException("update names column twice: " + step.Set[i].Column);
                }
            }
            expressions[i] = Rdv3Expression.Compile(step.Set[i].Expression, source.Columns);
        }
        Rdv3Relation output = NewLike(source);
        changed = 0;
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = (string[])source.Rows[r].Clone();
            if (selected.Rows.Contains(r))
            {
                for (int i = 0; i < columns.Length; i++)
                {
                    row[columns[i]] = expressions[i].Evaluate(source.Rows[r]);
                }
            }
            bool rowChanged = !SameRow(source.Rows[r], row);
            if (rowChanged) { changed++; }
            output.Rows.Add(row);
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
                    throw new InvalidDataException("select names column twice: " + output.Columns[i]);
                }
            }
        }
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = new string[fields.Length];
            for (int i = 0; i < fields.Length; i++) { row[i] = source.Rows[r][fields[i]]; }
            output.Rows.Add(row);
        }
        return output;
    }

    private static Rdv3Relation Calculate(Rdv3Relation source, Rdv3ProcessStepDef step)
    {
        string added = step.Output + "." + step.Column;
        if (source.ColumnOf(added) >= 0) { throw new InvalidDataException("calculate would duplicate column " + added); }
        Rdv3Expression expression = Rdv3Expression.Compile(step.Expression, source.Columns);
        Rdv3Relation output = new Rdv3Relation();
        output.Columns = new string[source.Columns.Length + 1];
        Array.Copy(source.Columns, output.Columns, source.Columns.Length);
        output.Columns[source.Columns.Length] = added;
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = new string[output.Columns.Length];
            Array.Copy(source.Rows[r], row, source.Columns.Length);
            row[source.Columns.Length] = expression.Evaluate(source.Rows[r]);
            output.Rows.Add(row);
        }
        return output;
    }

    private sealed class GroupValue
    {
        public string[] Keys;
        public decimal[] Sums;
        public int Count;
    }

    private static Rdv3Relation Aggregate(Rdv3Relation source, Rdv3ProcessStepDef step)
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
        if (source.Rows.Count == 0 && groups.Length == 0)
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
            if (!byKey.TryGetValue(key, out value))
            {
                value = NewGroup(keys, step.Aggregates.Count);
                byKey.Add(key, value);
                ordered.Add(value);
            }
            value.Count++;
            for (int a = 0; a < step.Aggregates.Count; a++)
            {
                if (fields[a] < 0) { continue; }
                decimal number;
                if (!decimal.TryParse(source.Rows[r][fields[a]], NumberStyles.Number,
                                      CultureInfo.InvariantCulture, out number))
                {
                    throw new InvalidDataException("sum received non-numeric text in " + step.Aggregates[a].Column);
                }
                value.Sums[a] += number;
            }
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
                    throw new InvalidDataException("aggregate names column twice: " + output.Columns[i]);
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

    private static Rdv3Relation Sort(Rdv3Relation source, Rdv3ProcessStepDef step)
    {
        int[] fields = new int[step.Orders.Count];
        for (int i = 0; i < fields.Length; i++) { fields[i] = source.NeedColumn(step.Orders[i].Column); }
        List<SortValue> rows = new List<SortValue>();
        for (int i = 0; i < source.Rows.Count; i++)
        {
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
        if (!decimal.TryParse(left, NumberStyles.Number, CultureInfo.InvariantCulture, out a)
            || !decimal.TryParse(right, NumberStyles.Number, CultureInfo.InvariantCulture, out b))
        {
            throw new InvalidDataException("numeric sort received non-numeric text");
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
                if (output.States != null) { output.States.Add(source.States[r]); }
            }
        }
        return output;
    }

    private static Rdv3Relation WriteLedger(Rdv3Data data, Rdv3ProcessJobDef job,
                                            Rdv3Relation source, Rdv3Relation target,
                                            Rdv3ProcessStepDef step, string initialStored,
                                            out Rdv3UpdateResult update)
    {
        int targetKey = target.NeedColumn(step.Keys[1]);
        if (targetKey != data.IdentityCol)
        {
            throw new InvalidDataException("ledger write target key is not data.ledger.identity");
        }
        int sourceKey = source.NeedColumn(step.Keys[0]);
        ValidateIdentity(data, job, source, sourceKey, step.Keys[0]);
        string[] sourceLines = new string[source.Rows.Count];
        for (int r = 0; r < source.Rows.Count; r++)
        {
            string[] row = new string[data.Columns.Count];
            for (int c = 0; c < data.Columns.Count; c++)
            {
                int from = (c == data.IdentityCol) ? sourceKey : source.ColumnOf(data.Columns[c].Ref);
                if (from >= 0) { row[c] = source.Rows[r][from]; }
                else { row[c] = ""; }
            }
            sourceLines[r] = string.Join("\t", row);
        }
        Rdv3ProcessJobDef apply = new Rdv3ProcessJobDef();
        apply.ApplyStep = step;
        apply.OnSourceChange = job.OnSourceChange;
        string[] targetStates = (target.States == null) ? new string[target.Rows.Count] : target.States.ToArray();
        update = Rdv3Ledger.ApplyUpdate(apply, target.ToLines(), targetStates,
                                       sourceLines, data.IdentityCol, initialStored);
        Rdv3Relation output = RelationOfLedger(data, update.Lines, update.States, initialStored);
        return output;
    }

    private static Rdv3Relation NewLike(Rdv3Relation source)
    {
        Rdv3Relation output = new Rdv3Relation();
        output.Kind = source.Kind;
        output.Columns = (string[])source.Columns.Clone();
        if (source.States != null) { output.States = new List<string>(); }
        return output;
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
            result.Kind = "rows";
            result.Lines = new string[0];
        }
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
                     || ch == '/' || ch == '(' || ch == ')')
            {
                sb.Append(ch);
                p++;
            }
            else
            {
                int start = p;
                while (p < expression.Length && !char.IsWhiteSpace(expression[p])
                       && "+-*/()".IndexOf(expression[p]) < 0) { p++; }
                string token = expression.Substring(start, p - start);
                string label = data.LabelOf(token);
                sb.Append(label.Length == 0 ? token : label);
            }
        }
        return sb.ToString();
    }
}

internal abstract class Rdv3Expression
{
    public abstract string Evaluate(string[] row);

    public static Rdv3Expression Compile(string text, string[] columns)
    {
        Parser parser = new Parser(text, columns);
        Rdv3Expression expression = parser.ParseExpression();
        parser.Finish();
        return expression;
    }

    private sealed class Literal : Rdv3Expression
    {
        private readonly string value;
        public Literal(string v) { value = v; }
        public override string Evaluate(string[] row) { return value; }
    }

    private sealed class Field : Rdv3Expression
    {
        private readonly int column;
        public Field(int c) { column = c; }
        public override string Evaluate(string[] row) { return row[column]; }
    }

    private sealed class Unary : Rdv3Expression
    {
        private readonly Rdv3Expression child;
        public Unary(Rdv3Expression value) { child = value; }
        public override string Evaluate(string[] row)
        {
            decimal value = Number(child.Evaluate(row));
            return (-value).ToString("G29", CultureInfo.InvariantCulture);
        }
    }

    private sealed class Binary : Rdv3Expression
    {
        private readonly char operation;
        private readonly Rdv3Expression left;
        private readonly Rdv3Expression right;
        public Binary(char op, Rdv3Expression a, Rdv3Expression b)
        {
            operation = op;
            left = a;
            right = b;
        }
        public override string Evaluate(string[] row)
        {
            string a = left.Evaluate(row);
            string b = right.Evaluate(row);
            decimal an;
            decimal bn;
            bool aNumber = decimal.TryParse(a, NumberStyles.Number, CultureInfo.InvariantCulture, out an);
            bool bNumber = decimal.TryParse(b, NumberStyles.Number, CultureInfo.InvariantCulture, out bn);
            if (operation == '+' && (!aNumber || !bNumber)) { return a + b; }
            if (!aNumber || !bNumber) { throw new InvalidDataException("arithmetic expression received non-numeric text"); }
            decimal value;
            if (operation == '+') { value = an + bn; }
            else if (operation == '-') { value = an - bn; }
            else if (operation == '*') { value = an * bn; }
            else
            {
                if (bn == 0) { throw new InvalidDataException("division by zero in expression"); }
                value = an / bn;
            }
            return value.ToString("G29", CultureInfo.InvariantCulture);
        }
    }

    private static decimal Number(string text)
    {
        decimal value;
        if (!decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out value))
        {
            throw new InvalidDataException("arithmetic expression received non-numeric text");
        }
        return value;
    }

    private sealed class Parser
    {
        private readonly string text;
        private readonly string[] columns;
        private int position;

        public Parser(string source, string[] names)
        {
            text = source ?? "";
            columns = names;
        }

        public Rdv3Expression ParseExpression()
        {
            Rdv3Expression value = ParseTerm();
            while (true)
            {
                Skip();
                if (!Take('+') && !Take('-')) { return value; }
                char operation = text[position - 1];
                value = new Binary(operation, value, ParseTerm());
            }
        }

        private Rdv3Expression ParseTerm()
        {
            Rdv3Expression value = ParseFactor();
            while (true)
            {
                Skip();
                if (!Take('*') && !Take('/')) { return value; }
                char operation = text[position - 1];
                value = new Binary(operation, value, ParseFactor());
            }
        }

        private Rdv3Expression ParseFactor()
        {
            Skip();
            if (Take('-')) { return new Unary(ParseFactor()); }
            if (Take('('))
            {
                Rdv3Expression value = ParseExpression();
                Skip();
                if (!Take(')')) { throw Error("missing )"); }
                return value;
            }
            if (position < text.Length && text[position] == '\'') { return new Literal(ParseString()); }
            int start = position;
            while (position < text.Length && !char.IsWhiteSpace(text[position])
                   && "+-*/()".IndexOf(text[position]) < 0) { position++; }
            if (start == position) { throw Error("expected a number, quoted text, column, or ("); }
            string token = text.Substring(start, position - start);
            decimal number;
            if (decimal.TryParse(token, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
            {
                return new Literal(number.ToString("G29", CultureInfo.InvariantCulture));
            }
            for (int i = 0; i < columns.Length; i++)
            {
                if (columns[i] == token) { return new Field(i); }
            }
            throw Error("unknown column " + token);
        }

        private string ParseString()
        {
            position++;
            StringBuilder sb = new StringBuilder();
            while (position < text.Length)
            {
                char ch = text[position++];
                if (ch != '\'') { sb.Append(ch); continue; }
                if (position < text.Length && text[position] == '\'')
                {
                    sb.Append('\'');
                    position++;
                    continue;
                }
                return sb.ToString();
            }
            throw Error("quoted text has no closing quote");
        }

        public void Finish()
        {
            Skip();
            if (position != text.Length) { throw Error("unexpected text"); }
        }

        private bool Take(char wanted)
        {
            if (position < text.Length && text[position] == wanted) { position++; return true; }
            return false;
        }

        private void Skip()
        {
            while (position < text.Length && char.IsWhiteSpace(text[position])) { position++; }
        }

        private InvalidDataException Error(string message)
        {
            return new InvalidDataException("expression at " + position.ToString(CultureInfo.InvariantCulture)
                + ": " + message);
        }
    }
}
