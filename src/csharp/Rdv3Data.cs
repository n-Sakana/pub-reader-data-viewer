// ============================================================================
// Rdv3Data.cs -- the data and ordered process-job definition in settings.json.
//
// Jobs are validated by the kinds of values their ordered steps consume and
// produce. A left-join chain is also recognized as an optional fast plan; the
// JSON pipeline remains the single source of truth.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public sealed class Rdv3TableDef
{
    public string Id = "";
    public string Label = "";
    public string File = "";
    public string Key = "";
    public int Ord;
    public string[] Head;
}

public sealed class Rdv3JoinDef
{
    public string Table = "";
    public string On = "";
    public int TableOrd;
    public int OnField = -1;
}

public sealed class Rdv3ColumnRef
{
    public string Ref = "";
    public string Table = "";
    public string Column = "";
    public int TableOrd;
    public int Field = -1;
}

public sealed class Rdv3ApplicationColumnDef
{
    public string Name = "";
    public string OnSourceChange = "";
}

public sealed class Rdv3ProcessColumnDef
{
    public string Column = "";
    public string As = "";

    public string OutputRef(string output)
    {
        return (As.Length == 0) ? Column : output + "." + As;
    }
}

public sealed class Rdv3ProcessWhereDef
{
    public string Column = "";
    public string Operator = "";
    public string Value = "";
}

public sealed class Rdv3ProcessAggregateDef
{
    public string Function = "";
    public string Column = "";
    public string As = "";
}

public sealed class Rdv3ProcessOrderDef
{
    public string Column = "";
    public string Direction = "ascending";
    public string Type = "text";
}

public sealed class Rdv3ProcessSetDef
{
    public string Column = "";
    public string Expression = "";
}

public sealed class Rdv3ProcessInputDef
{
    public string Id = "";
    public string Label = "";
    public string Table = "";
    public string File = "";
    public string Column = "";
    public string Key = "";
    public int TableOrd = -1;
    public bool IsTable { get { return TableOrd >= 0; } }
}

public sealed class Rdv3ProcessStepDef
{
    public string Operation = "";
    public string Target1 = "";
    public string Target2 = "";
    public string Condition = "";
    public string Output = "";
    public string[] Keys = new string[0];
    public List<Rdv3ProcessColumnDef> Columns = new List<Rdv3ProcessColumnDef>();
    public Rdv3ProcessWhereDef Where;
    public string Column = "";
    public string Expression = "";
    public List<string> GroupBy = new List<string>();
    public List<Rdv3ProcessAggregateDef> Aggregates = new List<Rdv3ProcessAggregateDef>();
    public List<Rdv3ProcessOrderDef> Orders = new List<Rdv3ProcessOrderDef>();
    public List<Rdv3ProcessSetDef> Set = new List<Rdv3ProcessSetDef>();
    // Row destinations for a merge. The three names mirror the three
    // possible relationships between source and destination rows.
    public string SourceOnly = "";
    public string Both = "";
    public string TargetOnly = "keep";
    public int Line;
}

public sealed class Rdv3ProcessJobDef
{
    public string Id = "";
    public string Name = "";
    public string Kind = "";
    public int Line;
    public List<Rdv3ProcessInputDef> Inputs = new List<Rdv3ProcessInputDef>();
    public List<Rdv3ProcessStepDef> Steps = new List<Rdv3ProcessStepDef>();

    // Optional fast plan compiled from a left-join pipeline.
    public string Spine = "";
    public int SpineOrd = -1;
    public List<Rdv3JoinDef> Joins = new List<Rdv3JoinDef>();
    public Rdv3ProcessStepDef ApplyStep;
    public bool FastJoinPlan;
    public string FinalKind = "";
    public string OnSourceChange = "";
}

public sealed class Rdv3Data
{
    public string EncodingName = "utf-8";
    public Encoding Enc = new UTF8Encoding(false);
    public List<Rdv3TableDef> Tables = new List<Rdv3TableDef>();
    public Dictionary<string, string> Labels = new Dictionary<string, string>(StringComparer.Ordinal);
    public List<string> LabelOrder = new List<string>();
    public List<Rdv3ProcessJobDef> Jobs = new List<Rdv3ProcessJobDef>();

    // The first update job is the automatic/start-up job.
    public Rdv3ProcessJobDef UpdateJob;
    public string Spine = "";
    public int SpineOrd;
    public List<Rdv3JoinDef> Joins = new List<Rdv3JoinDef>();

    public List<string> SearchRefs = new List<string>();
    public int[] SearchCols = new int[0];
    public string SearchMatch = "exact";
    public List<string> ApplicationColumns = new List<string>();
    public List<Rdv3ApplicationColumnDef> ApplicationColumnDefs = new List<Rdv3ApplicationColumnDef>();
    public bool WorkStateIsApplicationOwned;
    public string WorkStateOnSourceChange = "";
    public int IdentityCol = -1;
    public List<Rdv3ColumnRef> Columns = new List<Rdv3ColumnRef>();

    public Rdv3TableDef TableOf(string id)
    {
        for (int i = 0; i < Tables.Count; i++) { if (Tables[i].Id == id) { return Tables[i]; } }
        return null;
    }

    public Rdv3ProcessJobDef JobOf(string id)
    {
        for (int i = 0; i < Jobs.Count; i++) { if (Jobs[i].Id == id) { return Jobs[i]; } }
        return null;
    }

    public string LabelOf(string reference)
    {
        if (reference == null || reference.Length == 0) { return ""; }
        string label;
        if (Labels.TryGetValue(reference, out label)) { return label; }
        Rdv3TableDef table = TableOf(reference);
        return (table == null) ? "" : table.Label;
    }

    public int IndexOf(string reference)
    {
        for (int i = 0; i < Columns.Count; i++) { if (Columns[i].Ref == reference) { return i; } }
        return -1;
    }

    public string[] ColumnRefs
    {
        get
        {
            string[] r = new string[Columns.Count];
            for (int i = 0; i < r.Length; i++) { r[i] = Columns[i].Ref; }
            return r;
        }
    }

    public string[] Head
    {
        get
        {
            string[] h = new string[Columns.Count];
            for (int i = 0; i < h.Length; i++) { h[i] = Columns[i].Column; }
            return h;
        }
    }

    public static Rdv3Data Read(Rdv3Json o)
    {
        Rdv3Data d = new Rdv3Data();
        o.Only("encoding", "tables", "labels", "jobs", "ledger");

        d.EncodingName = o.StrOr("encoding", "utf-8");
        try
        {
            d.Enc = (string.Equals(d.EncodingName, "utf-8", StringComparison.OrdinalIgnoreCase))
                ? (Encoding)new UTF8Encoding(false) : Encoding.GetEncoding(d.EncodingName);
        }
        catch (Exception ex) { throw o.Member("encoding").Fail("is not an encoding this machine knows (" + ex.Message + ")"); }

        Rdv3Json tables = o.Obj("tables", true);
        if (tables.Order.Count == 0) { throw tables.Fail("names no table"); }
        for (int i = 0; i < tables.Order.Count; i++)
        {
            string id = tables.Order[i];
            Rdv3Json to = tables.Member(id);
            if (id.Trim().Length == 0 || id.IndexOf('.') >= 0 || id.IndexOf(' ') >= 0)
            {
                throw to.Fail("a table id must be a word without . or spaces");
            }
            if (id == "ledger") { throw to.Fail("ledger is a reserved value name"); }
            if (to.Kind != Rdv3Json.TObject) { throw to.Fail("must be an object { label, file, key }"); }
            to.Only("label", "file", "key");
            Rdv3TableDef t = new Rdv3TableDef();
            t.Id = id;
            t.Label = to.Need("label");
            t.File = to.Need("file");
            t.Key = to.Need("key");
            t.Ord = d.Tables.Count;
            d.Tables.Add(t);
            d.Labels.Add(t.Id, t.Label);
        }

        Rdv3Json labels = o.Obj("labels", true);
        for (int i = 0; i < labels.Order.Count; i++)
        {
            string key = labels.Order[i];
            Rdv3Json val = labels.Member(key);
            if (val.Kind != Rdv3Json.TString || val.Str.Trim().Length == 0) { throw val.Fail("must be a non-blank display label"); }
            if (d.Labels.ContainsKey(key)) { throw val.Fail(key + " already has a table label"); }
            d.Labels.Add(key, val.Str);
            d.LabelOrder.Add(key);
        }

        List<Rdv3Json> jobs = o.Objs("jobs", true);
        if (jobs.Count == 0) { throw o.Member("jobs").Fail("holds no job"); }
        for (int i = 0; i < jobs.Count; i++)
        {
            Rdv3ProcessJobDef job = ReadJob(d, jobs[i]);
            if (d.JobOf(job.Id) != null) { throw jobs[i].Member("id").Fail(job.Id + " is used by another job"); }
            d.Jobs.Add(job);
            if (job.Kind == "update" && d.UpdateJob == null) { d.UpdateJob = job; }
        }
        if (d.UpdateJob == null) { throw o.Member("jobs").Fail("must contain at least one update job"); }
        if (d.UpdateJob.FinalKind != "ledger")
        {
            throw jobNodesFor(d, jobs, d.UpdateJob).Member("steps").Fail("the automatic update job must produce ledger");
        }

        Rdv3Json ledger = o.Obj("ledger", true);
        ledger.Only("identity", "search", "columns");
        Rdv3Json columnGroups = ledger.Obj("columns", true);
        columnGroups.Only("source", "application");
        string[] cols = columnGroups.Strs("source", true);
        if (cols.Length == 0) { throw ledger.Member("columns").Fail("names no column"); }
        Rdv3Json colsNode = columnGroups.Member("source");
        for (int i = 0; i < cols.Length; i++)
        {
            Rdv3ColumnRef c = ParseRef(d, cols[i], colsNode.At(i));
            if (d.IndexOf(c.Ref) >= 0) { throw colsNode.At(i).Fail(c.Ref + " is listed twice"); }
            d.Columns.Add(c);
        }

        List<Rdv3Json> appColumns = columnGroups.Objs("application", true);
        if (appColumns.Count == 0) { throw columnGroups.Member("application").Fail("names no application-owned column"); }
        for (int i = 0; i < appColumns.Count; i++)
        {
            Rdv3Json ao = appColumns[i];
            ao.Only("name", "onSourceChange");
            Rdv3ApplicationColumnDef ac = new Rdv3ApplicationColumnDef();
            ac.Name = ao.Need("name");
            ac.OnSourceChange = ao.Word("onSourceChange", "", "reset", "preserve");
            if (d.ApplicationColumns.Contains(ac.Name)) { throw ao.Member("name").Fail(ac.Name + " is listed twice"); }
            if (ac.Name != "workState") { throw ao.Member("name").Fail(ac.Name + " is not an application column this program provides"); }
            d.ApplicationColumns.Add(ac.Name);
            d.ApplicationColumnDefs.Add(ac);
            if (ac.Name == "workState") { d.WorkStateOnSourceChange = ac.OnSourceChange; }
        }
        d.WorkStateIsApplicationOwned = d.ApplicationColumns.Contains("workState");
        if (d.WorkStateOnSourceChange.Length == 0)
        {
            throw columnGroups.Member("application").Fail("must define workState and its onSourceChange rule");
        }
        for (int i = 0; i < d.Jobs.Count; i++) { d.Jobs[i].OnSourceChange = d.WorkStateOnSourceChange; }

        Rdv3ColumnRef identityRef = ParseRef(d, ledger.Need("identity"), ledger.Member("identity"));
        Rdv3TableDef identityTable = d.TableOf(identityRef.Table);
        if (identityTable == null || identityRef.Column != identityTable.Key)
        {
            throw ledger.Member("identity").Fail("must be a table's unique key");
        }
        d.IdentityCol = d.IndexOf(identityRef.Ref);
        if (d.IdentityCol < 0) { throw colsNode.Fail("must include the update key " + identityRef.Ref + " (the row identity)"); }
        d.Spine = identityRef.Table;
        d.SpineOrd = identityRef.TableOrd;
        for (int i = 0; i < d.Jobs.Count; i++)
        {
            if (d.Jobs[i].Kind == "update")
            {
                CompileFastUpdate(d, d.Jobs[i], jobNodesFor(d, jobs, d.Jobs[i]));
            }
        }
        if (d.UpdateJob.FastJoinPlan)
        {
            d.Spine = d.UpdateJob.Spine;
            d.SpineOrd = d.UpdateJob.SpineOrd;
            d.Joins = d.UpdateJob.Joins;
        }
        else { d.Joins = new List<Rdv3JoinDef>(); }
        Rdv3Json search = ledger.Obj("search", true);
        search.Only("columns", "match");
        string[] searchRefs = search.Strs("columns", true);
        if (searchRefs.Length == 0) { throw search.Member("columns").Fail("names no search column"); }
        d.SearchMatch = search.Word("match", "exact", "exact", "contains");
        List<int> searchCols = new List<int>();
        for (int i = 0; i < searchRefs.Length; i++)
        {
            Rdv3ColumnRef c = ParseRef(d, searchRefs[i], search.Member("columns").At(i));
            int col = d.IndexOf(c.Ref);
            if (col < 0) { throw search.Member("columns").At(i).Fail(c.Ref + " must be one of the ledger source columns"); }
            if (d.SearchRefs.Contains(c.Ref)) { throw search.Member("columns").At(i).Fail(c.Ref + " is listed twice"); }
            d.SearchRefs.Add(c.Ref);
            searchCols.Add(col);
        }
        d.SearchCols = searchCols.ToArray();

        ValidateJobRefs(d, jobs);
        return d;
    }

    private static Rdv3Json jobNodesFor(Rdv3Data d, List<Rdv3Json> nodes, Rdv3ProcessJobDef job)
    {
        for (int i = 0; i < d.Jobs.Count; i++) { if (object.ReferenceEquals(d.Jobs[i], job)) { return nodes[i]; } }
        return nodes[0];
    }

    private static Rdv3ProcessJobDef ReadJob(Rdv3Data d, Rdv3Json o)
    {
        o.Only("id", "name", "kind", "inputs", "steps");
        Rdv3ProcessJobDef job = new Rdv3ProcessJobDef();
        job.Id = o.Need("id");
        job.Name = o.Need("name");
        job.Kind = o.Word("kind", "", "update", "delete");
        job.Line = o.Line;

        List<Rdv3Json> inputs = o.Objs("inputs", true);
        Dictionary<string, bool> inputNames = new Dictionary<string, bool>(StringComparer.Ordinal);
        for (int i = 0; i < inputs.Count; i++)
        {
            Rdv3Json io = inputs[i];
            Rdv3ProcessInputDef input = new Rdv3ProcessInputDef();
            string tableId = io.StrOr("table", "");
            if (tableId.Length > 0)
            {
                io.Only("table");
                Rdv3TableDef table = d.TableOf(tableId);
                if (table == null) { throw io.Member("table").Fail(tableId + " is not one of the tables"); }
                input.Id = table.Id;
                input.Label = table.Label;
                input.Table = table.Id;
                input.File = table.File;
                input.Column = table.Key;
                input.Key = table.Id + "." + table.Key;
                input.TableOrd = table.Ord;
            }
            else
            {
                io.Only("id", "label", "file", "column", "key");
                input.Id = io.Need("id");
                input.Label = io.Need("label");
                input.File = io.Need("file");
                input.Column = io.Need("column");
                input.Key = io.Need("key");
            }
            if (input.Id == "ledger") { throw io.Fail("ledger is a reserved value name"); }
            if (inputNames.ContainsKey(input.Id)) { throw io.Fail(input.Id + " is listed twice in inputs"); }
            inputNames.Add(input.Id, true);
            string existingLabel;
            if (d.Labels.TryGetValue(input.Id, out existingLabel))
            {
                if (existingLabel != input.Label) { throw io.Fail(input.Id + " has a different display label elsewhere"); }
            }
            else { d.Labels.Add(input.Id, input.Label); }
            job.Inputs.Add(input);
        }

        List<Rdv3Json> steps = o.Objs("steps", true);
        if (steps.Count == 0) { throw o.Member("steps").Fail("holds no step"); }
        for (int i = 0; i < steps.Count; i++)
        {
            Rdv3Json so = steps[i];
            Rdv3ProcessStepDef step = new Rdv3ProcessStepDef();
            step.Operation = so.Word("operation", "", "join", "extract", "delete", "append", "update",
                                     "merge", "replace", "select", "calculate", "aggregate", "sort", "distinct");
            step.Target1 = so.Need("target1");
            step.Target2 = so.StrOr("target2", "");
            step.Condition = so.StrOr("condition", "");
            step.Output = so.Need("output");
            if (step.Operation == "join")
            {
                so.Only("operation", "target1", "target2", "keys", "condition", "output");
                step.Keys = ReadKeys(so, true);
            }
            else if (step.Operation == "extract")
            {
                so.Only("operation", "target1", "target2", "keys", "condition", "where", "output");
                step.Keys = ReadKeys(so, false);
                step.Where = ReadWhere(so);
            }
            else if (step.Operation == "delete" || step.Operation == "append")
            {
                so.Only("operation", "target1", "target2", "condition", "output");
            }
            else if (step.Operation == "update")
            {
                so.Only("operation", "target1", "target2", "condition", "set", "output");
                step.Set = ReadSet(so);
            }
            else if (step.Operation == "merge")
            {
                so.Only("operation", "target1", "target2", "keys", "condition", "output",
                        "sourceOnly", "both", "targetOnly");
                step.Keys = ReadKeys(so, true);
                step.SourceOnly = so.Word("sourceOnly", "", "add", "ignore");
                step.Both = so.Word("both", "", "update", "keep");
                step.TargetOnly = so.Word("targetOnly", "keep", "keep", "delete");
            }
            else if (step.Operation == "replace")
            {
                so.Only("operation", "target1", "target2", "keys", "condition", "output");
                step.Keys = ReadKeys(so, true);
            }
            else if (step.Operation == "select")
            {
                so.Only("operation", "target1", "target2", "condition", "columns", "output");
                step.Columns = ReadColumns(so);
            }
            else if (step.Operation == "calculate")
            {
                so.Only("operation", "target1", "target2", "condition", "column", "expression", "output");
                step.Column = NeedAlias(so, "column");
                step.Expression = so.Need("expression");
            }
            else if (step.Operation == "aggregate")
            {
                so.Only("operation", "target1", "target2", "condition", "groupBy", "aggregates", "output");
                string[] groups = so.Strs("groupBy", true);
                for (int g = 0; g < groups.Length; g++) { step.GroupBy.Add(NeedRef(groups[g], so.Member("groupBy").At(g))); }
                step.Aggregates = ReadAggregates(so);
            }
            else if (step.Operation == "sort")
            {
                so.Only("operation", "target1", "target2", "condition", "orders", "output");
                step.Orders = ReadOrders(so);
            }
            else if (step.Operation == "distinct")
            {
                so.Only("operation", "target1", "target2", "condition", "columns", "output");
                string[] distinct = so.Strs("columns", true);
                if (distinct.Length == 0) { throw so.Member("columns").Fail("names no column"); }
                for (int c = 0; c < distinct.Length; c++)
                {
                    Rdv3ProcessColumnDef col = new Rdv3ProcessColumnDef();
                    col.Column = NeedRef(distinct[c], so.Member("columns").At(c));
                    step.Columns.Add(col);
                }
            }
            step.Line = so.Line;
            job.Steps.Add(step);
        }

        CompileJob(d, job, o);
        return job;
    }

    private static string[] ReadKeys(Rdv3Json o, bool required)
    {
        string[] keys = o.Strs("keys", required);
        for (int i = 0; i < keys.Length; i++) { keys[i] = NeedRef(keys[i], o.Member("keys").At(i)); }
        return keys;
    }

    private static Rdv3ProcessWhereDef ReadWhere(Rdv3Json o)
    {
        Rdv3Json w = o.Obj("where", false);
        if (w == null) { return null; }
        w.Only("column", "operator", "value");
        Rdv3ProcessWhereDef result = new Rdv3ProcessWhereDef();
        result.Column = NeedRef(w.Need("column"), w.Member("column"));
        result.Operator = w.Word("operator", "", "equals", "notEquals", "contains", "startsWith",
                                 "endsWith", "empty", "notEmpty", "greater", "atLeast", "less", "atMost");
        result.Value = w.StrOr("value", "");
        if (result.Operator != "empty" && result.Operator != "notEmpty" && !w.Has("value"))
        {
            throw w.Fail("value is required for " + result.Operator);
        }
        if ((result.Operator == "empty" || result.Operator == "notEmpty") && w.Has("value"))
        {
            throw w.Member("value").Fail("does not belong to " + result.Operator);
        }
        if (result.Operator == "greater" || result.Operator == "atLeast"
            || result.Operator == "less" || result.Operator == "atMost")
        {
            decimal number;
            if (!decimal.TryParse(result.Value, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
            {
                throw w.Member("value").Fail("must be a number for " + result.Operator);
            }
        }
        return result;
    }

    private static List<Rdv3ProcessColumnDef> ReadColumns(Rdv3Json o)
    {
        List<Rdv3Json> nodes = o.Objs("columns", true);
        if (nodes.Count == 0) { throw o.Member("columns").Fail("names no column"); }
        List<Rdv3ProcessColumnDef> result = new List<Rdv3ProcessColumnDef>();
        for (int i = 0; i < nodes.Count; i++)
        {
            Rdv3Json n = nodes[i];
            n.Only("column", "as");
            Rdv3ProcessColumnDef c = new Rdv3ProcessColumnDef();
            c.Column = NeedRef(n.Need("column"), n.Member("column"));
            c.As = n.Has("as") ? NeedAlias(n, "as") : "";
            result.Add(c);
        }
        return result;
    }

    private static List<Rdv3ProcessAggregateDef> ReadAggregates(Rdv3Json o)
    {
        List<Rdv3Json> nodes = o.Objs("aggregates", true);
        if (nodes.Count == 0) { throw o.Member("aggregates").Fail("names no aggregate"); }
        List<Rdv3ProcessAggregateDef> result = new List<Rdv3ProcessAggregateDef>();
        for (int i = 0; i < nodes.Count; i++)
        {
            Rdv3Json n = nodes[i];
            n.Only("function", "column", "as");
            Rdv3ProcessAggregateDef a = new Rdv3ProcessAggregateDef();
            a.Function = n.Word("function", "", "sum", "count");
            a.Column = n.StrOr("column", "");
            if (a.Column.Length > 0) { a.Column = NeedRef(a.Column, n.Member("column")); }
            if (a.Function == "sum" && a.Column.Length == 0) { throw n.Member("column").Fail("is required for sum"); }
            if (a.Function == "count" && a.Column.Length > 0) { throw n.Member("column").Fail("does not belong to count"); }
            a.As = NeedAlias(n, "as");
            result.Add(a);
        }
        return result;
    }

    private static List<Rdv3ProcessOrderDef> ReadOrders(Rdv3Json o)
    {
        List<Rdv3Json> nodes = o.Objs("orders", true);
        if (nodes.Count == 0) { throw o.Member("orders").Fail("names no order"); }
        List<Rdv3ProcessOrderDef> result = new List<Rdv3ProcessOrderDef>();
        for (int i = 0; i < nodes.Count; i++)
        {
            Rdv3Json n = nodes[i];
            n.Only("column", "direction", "type");
            Rdv3ProcessOrderDef order = new Rdv3ProcessOrderDef();
            order.Column = NeedRef(n.Need("column"), n.Member("column"));
            order.Direction = n.Word("direction", "ascending", "ascending", "descending");
            order.Type = n.Word("type", "text", "text", "number");
            result.Add(order);
        }
        return result;
    }

    private static List<Rdv3ProcessSetDef> ReadSet(Rdv3Json o)
    {
        List<Rdv3Json> nodes = o.Objs("set", true);
        if (nodes.Count == 0) { throw o.Member("set").Fail("names no column update"); }
        List<Rdv3ProcessSetDef> result = new List<Rdv3ProcessSetDef>();
        for (int i = 0; i < nodes.Count; i++)
        {
            Rdv3Json n = nodes[i];
            n.Only("column", "expression");
            Rdv3ProcessSetDef set = new Rdv3ProcessSetDef();
            set.Column = NeedRef(n.Need("column"), n.Member("column"));
            set.Expression = n.Need("expression");
            result.Add(set);
        }
        return result;
    }

    private static string NeedAlias(Rdv3Json o, string name)
    {
        string value = o.Need(name).Trim();
        if (value.IndexOf('.') >= 0) { throw o.Member(name).Fail("must be one column name, without ."); }
        for (int i = 0; i < value.Length; i++)
        {
            if (char.IsWhiteSpace(value[i])) { throw o.Member(name).Fail("must not contain spaces"); }
        }
        return value;
    }

    private static string NeedRef(string value, Rdv3Json at)
    {
        string result = (value == null) ? "" : value.Trim();
        int dot = result.IndexOf('.');
        if (dot <= 0 || dot == result.Length - 1) { throw at.Fail("a column is written <table>.<column>, not " + result); }
        return result;
    }

    private sealed class ProcessType
    {
        public string Kind;
        public int Token;
        public int BaseToken;
    }

    private static void CompileJob(Rdv3Data d, Rdv3ProcessJobDef job, Rdv3Json at)
    {
        Dictionary<string, ProcessType> values = new Dictionary<string, ProcessType>(StringComparer.Ordinal);
        int token = 1;
        for (int i = 0; i < job.Inputs.Count; i++)
        {
            values.Add(job.Inputs[i].Id, NewType("table", token++));
        }
        values.Add("ledger", NewType("ledger", token++));
        ProcessType last = null;
        for (int i = 0; i < job.Steps.Count; i++)
        {
            Rdv3Json sn = at.Member("steps").At(i);
            Rdv3ProcessStepDef step = job.Steps[i];
            ProcessType left;
            ProcessType right = null;
            if (!values.TryGetValue(step.Target1, out left))
            {
                throw sn.Member("target1").Fail(step.Target1 + " has not been defined by an input or earlier output");
            }
            if (step.Target2.Length > 0 && !values.TryGetValue(step.Target2, out right))
            {
                throw sn.Member("target2").Fail(step.Target2 + " has not been defined by an input or earlier output");
            }
            ProcessType output = CompileStepType(step, left, right, token++, sn);
            if (step.Output == "ledger" && output.Kind != "ledger")
            {
                throw sn.Member("output").Fail("ledger is reserved for a ledger value");
            }
            ProcessType old;
            if (values.TryGetValue(step.Output, out old) && step.Output != step.Target1 && step.Output != "ledger")
            {
                throw sn.Member("output").Fail(step.Output + " is already defined");
            }
            values[step.Output] = output;
            last = output;
            if ((step.Operation == "merge" || step.Operation == "replace") && output.Kind == "ledger")
            {
                job.ApplyStep = step;
            }
        }
        job.FinalKind = (last == null) ? "" : last.Kind;
    }

    private static ProcessType CompileStepType(Rdv3ProcessStepDef step, ProcessType left,
                                               ProcessType right, int token, Rdv3Json at)
    {
        bool leftTable = IsTable(left);
        bool rightTable = IsTable(right);
        if (step.Operation == "join")
        {
            NeedTypes(leftTable && rightTable, at, "join requires two tables");
            NeedKeys(step, at);
            NeedCondition(step, at, "match", "left", "full");
            return NewType("table", token);
        }
        if (step.Operation == "append")
        {
            NeedTypes(leftTable && rightTable, at, "append requires two tables");
            NeedEmptyCondition(step, at);
            return NewType("table", token);
        }
        if (step.Operation == "extract")
        {
            if (leftTable && step.Where != null && right == null)
            {
                NeedTypes(step.Keys.Length == 0, at, "a predicate extract does not use keys");
                NeedEmptyCondition(step, at);
                ProcessType rows = NewType("rows", token);
                rows.BaseToken = left.Token;
                return rows;
            }
            if (leftTable && rightTable)
            {
                NeedTypes(step.Where == null, at, "a two-table extract does not use where");
                NeedKeys(step, at);
                NeedCondition(step, at, "match", "exclude");
                ProcessType rows = NewType("rows", token);
                rows.BaseToken = left.Token;
                return rows;
            }
            if (left.Kind == "rows" && right != null && right.Kind == "rows")
            {
                NeedTypes(step.Where == null && step.Keys.Length == 0, at, "a row-set extract uses only its condition");
                NeedTypes(left.BaseToken == right.BaseToken, at, "row sets must come from the same table value");
                NeedCondition(step, at, "either", "both", "exclude");
                ProcessType rows = NewType("rows", token);
                rows.BaseToken = left.BaseToken;
                return rows;
            }
            throw at.Fail("extract requires a table predicate, two tables, or two row sets");
        }
        if (step.Operation == "delete" || step.Operation == "update")
        {
            NeedTypes(leftTable && right != null && right.Kind == "rows" && right.BaseToken == left.Token,
                      at, step.Operation + " requires a table and rows selected from that table value");
            NeedEmptyCondition(step, at);
            return NewType(left.Kind, token);
        }
        if (step.Operation == "sort" || step.Operation == "distinct")
        {
            NeedTypes(leftTable && right == null, at, step.Operation + " requires one table");
            NeedEmptyCondition(step, at);
            return NewType(left.Kind, token);
        }
        if (step.Operation == "select" || step.Operation == "calculate" || step.Operation == "aggregate")
        {
            NeedTypes(leftTable && right == null, at, step.Operation + " requires one table");
            NeedEmptyCondition(step, at);
            return NewType("table", token);
        }
        if (step.Operation == "merge" || step.Operation == "replace")
        {
            NeedTypes(leftTable && right != null && right.Kind == "ledger", at,
                      step.Operation + " requires a source table and ledger");
            NeedTypes(step.Output == "ledger", at, step.Operation + " must output ledger");
            NeedKeys(step, at);
            NeedEmptyCondition(step, at);
            return NewType("ledger", token);
        }
        throw at.Fail("unknown operation " + step.Operation);
    }

    private static ProcessType NewType(string kind, int token)
    {
        ProcessType t = new ProcessType();
        t.Kind = kind;
        t.Token = token;
        return t;
    }

    private static bool IsTable(ProcessType t)
    {
        return t != null && (t.Kind == "table" || t.Kind == "ledger");
    }

    private static void NeedTypes(bool ok, Rdv3Json at, string message)
    {
        if (!ok) { throw at.Fail(message); }
    }

    private static void NeedKeys(Rdv3ProcessStepDef step, Rdv3Json at)
    {
        if (step.Keys.Length != 2) { throw at.Member("keys").Fail("must hold exactly two column names"); }
    }

    private static void NeedCondition(Rdv3ProcessStepDef step, Rdv3Json at, params string[] allowed)
    {
        for (int i = 0; i < allowed.Length; i++) { if (step.Condition == allowed[i]) { return; } }
        throw at.Member("condition").Fail("must be one of " + string.Join(" / ", allowed));
    }

    private static void NeedEmptyCondition(Rdv3ProcessStepDef step, Rdv3Json at)
    {
        if (step.Condition.Length > 0) { throw at.Member("condition").Fail("does not belong to this operation"); }
    }

    private static void CompileFastUpdate(Rdv3Data d, Rdv3ProcessJobDef job, Rdv3Json at)
    {
        job.FastJoinPlan = false;
        job.Spine = "";
        job.SpineOrd = -1;
        job.Joins.Clear();
        if (job.ApplyStep == null || job.Steps.Count < 2
            || !object.ReferenceEquals(job.Steps[job.Steps.Count - 1], job.ApplyStep)) { return; }

        string live = "";
        Rdv3TableDef spine = null;
        List<Rdv3JoinDef> plan = new List<Rdv3JoinDef>();
        for (int i = 0; i < job.Steps.Count - 1; i++)
        {
            Rdv3ProcessStepDef step = job.Steps[i];
            if (step.Operation != "join" || step.Condition != "left" || step.Keys.Length != 2) { return; }
            Rdv3TableDef other;
            Rdv3ColumnRef leftKey;
            Rdv3ColumnRef rightKey;
            try
            {
                leftKey = ParseRef(d, step.Keys[0], at.Member("steps").At(i).Member("keys").At(0));
                rightKey = ParseRef(d, step.Keys[1], at.Member("steps").At(i).Member("keys").At(1));
            }
            catch (Rdv3LoadError) { return; }
            if (i == 0)
            {
                spine = d.TableOf(step.Target1);
                other = d.TableOf(step.Target2);
                if (spine == null || other == null || leftKey.Table != spine.Id) { return; }
            }
            else
            {
                if (step.Target1 != live || leftKey.Table != spine.Id) { return; }
                other = d.TableOf(step.Target2);
                if (other == null) { return; }
            }
            if (rightKey.Table != other.Id || rightKey.Column != other.Key || other.Ord == spine.Ord) { return; }
            for (int j = 0; j < plan.Count; j++) { if (plan[j].TableOrd == other.Ord) { return; } }
            Rdv3JoinDef join = new Rdv3JoinDef();
            join.Table = other.Id;
            join.TableOrd = other.Ord;
            join.On = leftKey.Column;
            plan.Add(join);
            live = step.Output;
        }
        if (plan.Count == 0 || job.ApplyStep.Target1 != live || job.ApplyStep.Target2 != "ledger"
            || job.ApplyStep.Keys.Length != 2) { return; }
        string identity = d.Columns[d.IdentityCol].Ref;
        if (job.ApplyStep.Keys[0] != identity || job.ApplyStep.Keys[1] != identity) { return; }
        job.Spine = spine.Id;
        job.SpineOrd = spine.Ord;
        job.Joins = plan;
        job.FastJoinPlan = true;
    }

    private static void ValidateJobRefs(Rdv3Data d, List<Rdv3Json> jobNodes)
    {
        for (int j = 0; j < d.Jobs.Count; j++)
        {
            Rdv3ProcessJobDef job = d.Jobs[j];
            Rdv3Json node = jobNodes[j];
            for (int i = 0; i < job.Inputs.Count; i++)
            {
                Rdv3ProcessInputDef input = job.Inputs[i];
                if (!input.IsTable)
                {
                    Rdv3ColumnRef key = ParseRef(d, input.Key, node.Member("inputs").At(i).Member("key"));
                    if (d.IndexOf(key.Ref) < 0) { throw node.Member("inputs").At(i).Member("key").Fail(key.Ref + " must be one of the ledger source columns"); }
                }
            }
            for (int i = 0; i < job.Steps.Count; i++)
            {
                Rdv3ProcessStepDef step = job.Steps[i];
                Rdv3Json sn = node.Member("steps").At(i);
                RequireLabel(d, step.Target1, sn.Member("target1"));
                if (step.Target2.Length > 0) { RequireLabel(d, step.Target2, sn.Member("target2")); }
                RequireLabel(d, step.Output, sn.Member("output"));
                for (int k = 0; k < step.Keys.Length; k++) { RequireLabel(d, step.Keys[k], sn.Member("keys").At(k)); }
                if (step.Where != null) { RequireLabel(d, step.Where.Column, sn.Member("where").Member("column")); }
                for (int c = 0; c < step.Columns.Count; c++)
                {
                    RequireLabel(d, step.Columns[c].Column, sn.Member("columns").At(c));
                    if (step.Columns[c].As.Length > 0)
                    {
                        RequireLabel(d, step.Columns[c].OutputRef(step.Output), sn.Member("columns").At(c).Member("as"));
                    }
                }
                if (step.Column.Length > 0)
                {
                    RequireLabel(d, step.Output + "." + step.Column, sn.Member("column"));
                    RequireExpressionLabels(d, step.Expression, sn.Member("expression"));
                }
                for (int g = 0; g < step.GroupBy.Count; g++) { RequireLabel(d, step.GroupBy[g], sn.Member("groupBy").At(g)); }
                for (int a = 0; a < step.Aggregates.Count; a++)
                {
                    if (step.Aggregates[a].Column.Length > 0)
                    {
                        RequireLabel(d, step.Aggregates[a].Column, sn.Member("aggregates").At(a).Member("column"));
                    }
                    RequireLabel(d, step.Output + "." + step.Aggregates[a].As, sn.Member("aggregates").At(a).Member("as"));
                }
                for (int x = 0; x < step.Orders.Count; x++)
                {
                    RequireLabel(d, step.Orders[x].Column, sn.Member("orders").At(x).Member("column"));
                }
                for (int x = 0; x < step.Set.Count; x++)
                {
                    RequireLabel(d, step.Set[x].Column, sn.Member("set").At(x).Member("column"));
                    RequireExpressionLabels(d, step.Set[x].Expression,
                                            sn.Member("set").At(x).Member("expression"));
                }
            }
        }
    }

    private static void RequireLabel(Rdv3Data d, string name, Rdv3Json at)
    {
        if (d.LabelOf(name).Length == 0) { throw at.Fail(name + " has no screen label under data.labels or tables.*.label"); }
    }

    private static void RequireExpressionLabels(Rdv3Data d, string expression, Rdv3Json at)
    {
        int p = 0;
        while (p < expression.Length)
        {
            char ch = expression[p];
            if (char.IsWhiteSpace(ch) || "+-*/()".IndexOf(ch) >= 0) { p++; continue; }
            if (ch == '\'')
            {
                p++;
                while (p < expression.Length)
                {
                    if (expression[p++] != '\'') { continue; }
                    if (p < expression.Length && expression[p] == '\'') { p++; continue; }
                    break;
                }
                continue;
            }
            int start = p;
            while (p < expression.Length && !char.IsWhiteSpace(expression[p])
                   && "+-*/()".IndexOf(expression[p]) < 0) { p++; }
            string token = expression.Substring(start, p - start);
            decimal number;
            if (!decimal.TryParse(token, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
            {
                NeedRef(token, at);
                RequireLabel(d, token, at);
            }
        }
    }

    private static Rdv3ColumnRef ParseRef(Rdv3Data d, string s, Rdv3Json at)
    {
        Rdv3ColumnRef c = new Rdv3ColumnRef();
        c.Ref = (s == null) ? "" : s.Trim();
        int dot = c.Ref.IndexOf('.');
        if (dot <= 0 || dot == c.Ref.Length - 1) { throw at.Fail("a column is written <table>.<column>, not " + c.Ref); }
        c.Table = c.Ref.Substring(0, dot);
        c.Column = c.Ref.Substring(dot + 1);
        Rdv3TableDef t = d.TableOf(c.Table);
        if (t == null) { throw at.Fail(c.Table + " is not one of the tables"); }
        c.TableOrd = t.Ord;
        return c;
    }

    public void Bind(string[][] heads)
    {
        for (int t = 0; t < Tables.Count; t++)
        {
            Tables[t].Head = heads[t];
            if (FieldOf(heads[t], Tables[t].Key) < 0) { throw Missing(Tables[t], Tables[t].Key, "tables." + Tables[t].Id + ".key"); }
        }
        for (int j = 0; j < Jobs.Count; j++)
        {
            Rdv3ProcessJobDef job = Jobs[j];
            if (!job.FastJoinPlan) { continue; }
            string[] spineHead = heads[job.SpineOrd];
            for (int i = 0; i < job.Joins.Count; i++)
            {
                job.Joins[i].OnField = FieldOf(spineHead, job.Joins[i].On);
                if (job.Joins[i].OnField < 0) { throw Missing(Tables[job.SpineOrd], job.Joins[i].On, "jobs.steps.keys"); }
            }
        }
        for (int i = 0; i < Columns.Count; i++)
        {
            Rdv3ColumnRef c = Columns[i];
            c.Field = FieldOf(heads[c.TableOrd], c.Column);
            if (c.Field < 0) { throw Missing(Tables[c.TableOrd], c.Column, "ledger.columns"); }
        }
        Rdv3Process.ValidateColumns(this, heads);
    }

    private static Rdv3DataError Missing(Rdv3TableDef t, string column, string where)
    {
        return new Rdv3DataError(Rdv3Text.DataNoColumn.Replace("{file}", t.File).Replace("{row}", "1")
            .Replace("{name}", column) + " (data." + where + ")");
    }

    private static int FieldOf(string[] head, string name)
    {
        if (head == null) { return -1; }
        for (int i = 0; i < head.Length; i++) { if (head[i] == name) { return i; } }
        return -1;
    }

    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("tables=").Append(Tables.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" jobs=").Append(Jobs.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" update=").Append((UpdateJob == null) ? "-" : UpdateJob.Id);
        sb.Append(" spine=").Append(Spine);
        sb.Append(" joins=");
        for (int i = 0; i < Joins.Count; i++) { if (i > 0) { sb.Append(','); } sb.Append(Joins[i].Table).Append("(on ").Append(Joins[i].On).Append(')'); }
        sb.Append(" columns=").Append(Columns.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" identity=").Append(Columns[IdentityCol].Ref);
        sb.Append(" search=").Append(string.Join(",", SearchRefs.ToArray())).Append("(").Append(SearchMatch).Append(")");
        sb.Append(" appColumns=").Append(string.Join(",", ApplicationColumns.ToArray()));
        sb.Append(" sourceChange=").Append(WorkStateOnSourceChange);
        sb.Append(" encoding=").Append(EncodingName);
        return sb.ToString();
    }
}
