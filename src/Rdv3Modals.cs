// ============================================================================
// Rdv3Modals.cs -- data and responses for the approved v13 HTML dialogs.
// C# 5, ASCII source.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

public static class Rdv3ConfirmForm
{
    public static bool Ask(Rdv3Form owner, string title, string body)
    {
        Rdv3Json result = owner.ShowModal(
            "confirm",
            "{\"title\":" + Rdv3WebJson.Q(title) +
            ",\"body\":" + Rdv3WebJson.Q(body) +
            ",\"ask\":true}");
        return Rdv3Form.Flag(result, "ok", false);
    }

    public static void Tell(Rdv3Form owner, string title, string body)
    {
        owner.ShowModal(
            "confirm",
            "{\"title\":" + Rdv3WebJson.Q(title) +
            ",\"body\":" + Rdv3WebJson.Q(body) +
            ",\"ask\":false}");
    }
}

public static class Rdv3CandidatesForm
{
    public static int Pick(
        Rdv3Form owner,
        Rdv3CandidatesDef def,
        List<Rdv3CandRow> rows,
        int total,
        int selected)
    {
        string content = Rdv3ModalData.Candidates(
            owner,
            def,
            rows,
            total,
            selected,
            true,
            def.Title,
            def.Hint);
        Rdv3Json result = owner.ShowModal("candidates", content);
        if (!Rdv3Form.Flag(result, "ok", false)) { return -1; }
        return Rdv3Form.Number(result, "index", -1);
    }
}

public static class Rdv3LedgerUpdateForm
{
    public static bool Ask(Rdv3Form owner, List<Rdv3CandRow> resetRows)
    {
        string body = Rdv3Text.SharedUpdateBody;
        if (resetRows != null && resetRows.Count > 0)
        {
            body += "\r\n" + Rdv3Text.SharedResetFmt
                .Replace("{state}", owner.Screen.Work.InitialState.Text)
                .Replace("{n}", resetRows.Count.ToString("N0", CultureInfo.InvariantCulture));
        }
        string content = Rdv3ModalData.Candidates(
            owner,
            owner.Screen.Candidates,
            resetRows,
            resetRows == null ? 0 : resetRows.Count,
            -1,
            false,
            Rdv3Text.SharedUpdateTitle,
            body);
        return Rdv3Form.Flag(owner.ShowModal("shared", content), "ok", false);
    }

    public static void TellReset(Rdv3Form owner, List<Rdv3CandRow> resetRows)
    {
        string body = Rdv3Text.NoteUpdated;
        if (resetRows != null && resetRows.Count > 0)
        {
            body += "\r\n" + Rdv3Text.SharedResetFmt
                .Replace("{state}", owner.Screen.Work.InitialState.Text)
                .Replace("{n}", resetRows.Count.ToString("N0", CultureInfo.InvariantCulture));
        }
        owner.ShowModal(
            "sharedTell",
            Rdv3ModalData.Candidates(
                owner,
                owner.Screen.Candidates,
                resetRows,
                resetRows == null ? 0 : resetRows.Count,
                -1,
                false,
                Rdv3Text.SharedUpdateTitle,
                body));
    }
}

public static class Rdv3UnmatchedForm
{
    public static void Tell(Rdv3Form owner, List<Rdv3UnmatchedChange> values)
    {
        List<Rdv3UnmatchedChange> rows = values ?? new List<Rdv3UnmatchedChange>();
        StringBuilder sb = new StringBuilder();
        sb.Append("{\"title\":").Append(Rdv3WebJson.Q(Rdv3Text.UnmatchedTitle));
        sb.Append(",\"body\":").Append(Rdv3WebJson.Q(
            Rdv3Text.UnmatchedBodyFmt.Replace(
                "{n}",
                rows.Count.ToString("N0", CultureInfo.InvariantCulture))));
        sb.Append(",\"rows\":[");
        for (int i = 0; i < rows.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            sb.Append("[").Append(Rdv3WebJson.Q((i + 1).ToString(CultureInfo.InvariantCulture)));
            sb.Append(',').Append(Rdv3WebJson.Q(rows[i].Identity));
            sb.Append(',').Append(Rdv3WebJson.Q(rows[i].Reason == "missing"
                ? Rdv3Text.UnmatchedMissing : Rdv3Text.UnmatchedChanged)).Append(']');
        }
        sb.Append("]}");
        owner.ShowModal("unmatched", sb.ToString());
    }
}

public static class Rdv3ProcessForm
{
    public static bool ShowJob(
        Rdv3Form owner,
        Rdv3Data data,
        string jobId,
        string dataDir,
        string ledgerPath)
    {
        Rdv3ProcessJobDef job = data.JobOf(jobId);
        if (job == null) { return false; }
        bool inputsOk;
        string content = Build(data, job, dataDir, ledgerPath, out inputsOk);
        Rdv3Json result = owner.ShowModal(job.Kind == "delete" ? "delete" : "update", content);
        return inputsOk && Rdv3Form.Flag(result, "ok", false);
    }

    private static string Build(
        Rdv3Data data,
        Rdv3ProcessJobDef job,
        string dataDir,
        string ledgerPath,
        out bool inputsOk)
    {
        inputsOk = true;
        StringBuilder sb = new StringBuilder(8192);
        bool deleting = job.Kind == "delete";
        sb.Append("{\"title\":").Append(Rdv3WebJson.Q(
            deleting ? Rdv3Text.DeleteRecordsTitle : Rdv3Text.UpdateRecordsTitle));
        sb.Append(",\"hint\":").Append(Rdv3WebJson.Q(
            deleting ? Rdv3Text.DeleteRecordsHint : Rdv3Text.UpdateRecordsHint));
        sb.Append(",\"inputTitle\":").Append(Rdv3WebJson.Q(
            Rdv3Text.SecInputs.Replace("{dir}", new DirectoryInfo(dataDir).Name)));
        sb.Append(",\"inputs\":[");
        for (int i = 0; i < job.Inputs.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            Rdv3ProcessInputDef input = job.Inputs[i];
            string path = Path.IsPathRooted(input.File)
                ? input.File : Path.Combine(dataDir, input.File);
            string rows = "";
            string validation;
            bool valid = true;
            if (!File.Exists(path))
            {
                valid = false;
                inputsOk = false;
                validation = Rdv3Text.ValidationMissing;
            }
            else
            {
                try
                {
                    Rdv3Table table = Rdv3Table.Read(
                        path,
                        input.Id,
                        data.Enc,
                        input.Column,
                        input.KeyValidation);
                    new Rdv3Index(table);
                    rows = table.Rows.ToString("N0", CultureInfo.InvariantCulture);
                    List<string> warnings = new List<string>();
                    if (table.InvalidEncodingRow > 0)
                    {
                        warnings.Add(Rdv3Text.ValidationEncodingMismatch.Replace(
                            "{row}",
                            table.InvalidEncodingRow.ToString(CultureInfo.InvariantCulture)));
                    }
                    if (table.ControlCharacterWarning.Length > 0)
                    {
                        warnings.Add(table.ControlCharacterWarning);
                    }
                    validation = warnings.Count == 0
                        ? Rdv3Text.ValidationColumnsMatch
                        : string.Join(" / ", warnings.ToArray());
                }
                catch (Exception exception)
                {
                    valid = false;
                    inputsOk = false;
                    validation = Rdv3Text.ValidationError + " (" + exception.Message + ")";
                }
            }
            sb.Append("{\"id\":").Append(Rdv3WebJson.Q(input.Id));
            sb.Append(",\"file\":").Append(Rdv3WebJson.Q(input.File));
            sb.Append(",\"key\":").Append(Rdv3WebJson.Q(input.Column));
            sb.Append(",\"rows\":").Append(Rdv3WebJson.Q(rows));
            sb.Append(",\"validation\":").Append(Rdv3WebJson.Q(validation));
            sb.Append(",\"valid\":").Append(Rdv3WebJson.B(valid)).Append('}');
        }
        sb.Append("],\"steps\":[");
        for (int i = 0; i < job.Steps.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            Rdv3ProcessStepDef step = job.Steps[i];
            sb.Append('[').Append(Rdv3WebJson.Q((i + 1).ToString(CultureInfo.InvariantCulture)));
            sb.Append(',').Append(Rdv3WebJson.Q(Rdv3Text.OperationLabel(step.Operation)));
            sb.Append(',').Append(Rdv3WebJson.Q(Display(data, step.Target1)));
            sb.Append(',').Append(Rdv3WebJson.Q(Display(data, step.Target2)));
            sb.Append(',').Append(Rdv3WebJson.Q(DisplayKey(data, step)));
            sb.Append(',').Append(Rdv3WebJson.Q(DisplayCondition(data, step)));
            sb.Append(',').Append(Rdv3WebJson.Q(Display(data, step.Output))).Append(']');
        }
        string directory = Path.GetDirectoryName(ledgerPath) ?? "";
        string file = Path.GetFileName(ledgerPath);
        string updated = File.Exists(ledgerPath)
            ? File.GetLastWriteTime(ledgerPath).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)
            : Rdv3Text.LblNeverWritten;
        sb.Append("],\"output\":[");
        sb.Append(Rdv3WebJson.Q(directory)).Append(',');
        sb.Append(Rdv3WebJson.Q(file)).Append(',');
        sb.Append(Rdv3WebJson.Q(updated)).Append(']');
        sb.Append(",\"canRun\":").Append(Rdv3WebJson.B(inputsOk));
        sb.Append(",\"executeText\":").Append(Rdv3WebJson.Q(
            deleting ? Rdv3Text.BtnDelete : Rdv3Text.BtnExecute)).Append('}');
        return sb.ToString();
    }

    private static string Display(Rdv3Data data, string name)
    {
        if (string.IsNullOrEmpty(name)) { return ""; }
        string label = data.LabelOf(name);
        return label.Length == 0 ? name : label;
    }

    private static string DisplayKey(Rdv3Data data, Rdv3ProcessStepDef step)
    {
        List<string> values = new List<string>();
        for (int i = 0; i < step.Keys.Length; i++) { values.Add(Display(data, step.Keys[i])); }
        if (values.Count > 0) { return string.Join(" = ", values.ToArray()); }
        if (step.Operation == "select" || step.Operation == "distinct")
        {
            for (int i = 0; i < step.Columns.Count; i++) { values.Add(Display(data, step.Columns[i].Column)); }
        }
        else if (step.Operation == "calculate" && step.Column.Length > 0)
        {
            values.Add(Display(data, step.Output + "." + step.Column));
        }
        else if (step.Operation == "aggregate")
        {
            for (int i = 0; i < step.GroupBy.Count; i++) { values.Add(Display(data, step.GroupBy[i])); }
        }
        else if (step.Operation == "sort")
        {
            for (int i = 0; i < step.Orders.Count; i++) { values.Add(Display(data, step.Orders[i].Column)); }
        }
        else if (step.Operation == "update")
        {
            for (int i = 0; i < step.Set.Count; i++) { values.Add(Display(data, step.Set[i].Column)); }
        }
        return values.Count == 0 ? "" : string.Join(" / ", values.ToArray());
    }

    private static string DisplayCondition(Rdv3Data data, Rdv3ProcessStepDef step)
    {
        if (step.Operation == "merge")
        {
            return Rdv3Text.MergeDestinations(step.SourceOnly, step.Both, step.TargetOnly);
        }
        if (step.Operation == "join") { return Rdv3Text.JoinConditionLabel(step.Condition); }
        if (step.Where != null)
        {
            string result = Display(data, step.Where.Column) + " " +
                Rdv3Text.PredicateLabel(step.Where.Operator);
            if (step.Where.Operator != "empty" && step.Where.Operator != "notEmpty")
            {
                result += " " + step.Where.Value;
            }
            return result;
        }
        if (step.Operation == "select")
        {
            List<string> mappings = new List<string>();
            for (int i = 0; i < step.Columns.Count; i++)
            {
                if (step.Columns[i].As.Length > 0)
                {
                    mappings.Add(Display(data, step.Columns[i].Column) + " -> " +
                        Display(data, step.Columns[i].OutputRef(step.Output)));
                }
            }
            return string.Join(" / ", mappings.ToArray());
        }
        if (step.Operation == "calculate")
        {
            return Rdv3Process.DisplayExpression(data, step.Expression);
        }
        if (step.Operation == "aggregate")
        {
            List<string> values = new List<string>();
            for (int i = 0; i < step.Aggregates.Count; i++)
            {
                string source = step.Aggregates[i].Column.Length == 0
                    ? "*" : Display(data, step.Aggregates[i].Column);
                values.Add(Rdv3Text.AggregateLabel(step.Aggregates[i].Function) + "(" + source + ") -> " +
                    Display(data, step.Output + "." + step.Aggregates[i].As));
            }
            return string.Join(" / ", values.ToArray());
        }
        if (step.Operation == "sort")
        {
            List<string> values = new List<string>();
            for (int i = 0; i < step.Orders.Count; i++)
            {
                values.Add(Display(data, step.Orders[i].Column) + " " +
                    Rdv3Text.DirectionLabel(step.Orders[i].Direction) + " (" +
                    Rdv3Text.SortTypeLabel(step.Orders[i].Type) + ")");
            }
            return string.Join(" / ", values.ToArray());
        }
        if (step.Operation == "update")
        {
            List<string> values = new List<string>();
            for (int i = 0; i < step.Set.Count; i++)
            {
                values.Add(Display(data, step.Set[i].Column) + " = " +
                    Rdv3Process.DisplayExpression(data, step.Set[i].Expression));
            }
            return string.Join(" / ", values.ToArray());
        }
        return Rdv3Text.ConditionLabel(step.Condition);
    }
}

public sealed class Rdv3ExportFilter
{
    public string Field = "";
    public string Operator = "";
    public string First = "";
    public string Last = "";

    public bool Matches(Rdv3Data data, string[] values, string storedState)
    {
        string value;
        if (Field == "$work") { value = storedState ?? ""; }
        else
        {
            int column = data.IndexOf(Field);
            if (column < 0 || values == null || column >= values.Length) { return false; }
            value = values[column] ?? "";
        }
        Rdv3ColumnTypeDef type = data.TypeOf(Field);
        if (type == null)
        {
            if (Operator == "contains") { return value.IndexOf(First, StringComparison.Ordinal) >= 0; }
            if (Operator == "equals") { return string.Equals(value, First, StringComparison.Ordinal); }
            if (Operator == "startsWith") { return value.StartsWith(First, StringComparison.Ordinal); }
            if (Operator == "notContains") { return value.IndexOf(First, StringComparison.Ordinal) < 0; }
            throw new InvalidOperationException("invalid text export filter operator: " + Operator);
        }
        if (Operator != "range")
        {
            throw new InvalidOperationException("invalid typed export filter operator: " + Operator);
        }
        if (type.Type == "date")
        {
            DateTime actual;
            DateTime first;
            DateTime last;
            if (!type.TryDate(value, out actual) || !type.TryDate(First, out first) ||
                !type.TryDate(Last, out last)) { return false; }
            return actual >= first && actual <= last;
        }
        decimal actualNumber;
        decimal firstNumber;
        decimal lastNumber;
        if (!type.TryNumber(value, out actualNumber) || !type.TryNumber(First, out firstNumber) ||
            !type.TryNumber(Last, out lastNumber)) { return false; }
        return actualNumber >= firstNumber && actualNumber <= lastNumber;
    }
}

public sealed class Rdv3ExportRequest
{
    public string Path = "";
    public List<string> Fields = new List<string>();
    public List<Rdv3ExportFilter> Filters = new List<Rdv3ExportFilter>();

    public bool Matches(Rdv3Data data, string[] values, string storedState)
    {
        for (int i = 0; i < Filters.Count; i++)
        {
            if (!Filters[i].Matches(data, values, storedState)) { return false; }
        }
        return true;
    }
}

public static class Rdv3ExportForm
{
    public static Rdv3ExportRequest Pick(
        Rdv3Form owner,
        Rdv3Data data,
        Rdv3Screen screen,
        string baseDir)
    {
        List<string> refs = new List<string>();
        StringBuilder content = new StringBuilder(8192);
        content.Append("{\"title\":").Append(Rdv3WebJson.Q(Rdv3Text.ExportTitle));
        content.Append(",\"hint\":").Append(Rdv3WebJson.Q(Rdv3Text.ExportHint));
        content.Append(",\"destination\":").Append(Rdv3WebJson.Q(Path.Combine(
            baseDir,
            Rdv3Text.ExportDefaultPath.Replace(
                "{yyyyMMdd-HHmmss}",
                DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)))));
        content.Append(",\"defaults\":").Append(Rdv3WebJson.S(screen.ExportDefaultFields));
        content.Append(",\"fields\":[");
        bool comma = false;
        for (int i = 0; i < data.LabelOrder.Count; i++)
        {
            string reference = data.LabelOrder[i];
            int column = data.IndexOf(reference);
            if (column < 0) { continue; }
            if (comma) { content.Append(','); }
            comma = true;
            refs.Add(reference);
            Rdv3ColumnTypeDef type = data.TypeOf(reference);
            content.Append("{\"ref\":").Append(Rdv3WebJson.Q(reference));
            content.Append(",\"label\":").Append(Rdv3WebJson.Q(
                data.LabelOf(reference) + " (" + reference + ")"));
            content.Append(",\"kind\":").Append(Rdv3WebJson.Q(type == null ? "text" : type.Type));
            content.Append(",\"format\":").Append(Rdv3WebJson.Q(type == null ? "" : type.Format)).Append('}');
        }
        if (comma) { content.Append(','); }
        refs.Add("$work");
        Rdv3StateDef state = screen.Work.InitialTargetState ?? screen.Work.InitialState;
        content.Append("{\"ref\":\"$work\",\"label\":").Append(Rdv3WebJson.Q(
            state == null ? screen.Work.Column : state.Text));
        content.Append(",\"kind\":\"text\",\"format\":\"\"}");
        content.Append("]}");

        Rdv3Json result = owner.ShowModal("export", content.ToString());
        if (!Rdv3Form.Flag(result, "ok", false)) { return null; }
        Rdv3ExportRequest request = new Rdv3ExportRequest();
        request.Path = Rdv3Form.Text(result, "path").Trim();
        Rdv3Json selected = result.Member("fields");
        if (selected != null && selected.Kind == Rdv3Json.TArray)
        {
            for (int i = 0; i < selected.Count; i++)
            {
                Rdv3Json item = selected.At(i);
                if (item != null && item.Kind == Rdv3Json.TString && refs.Contains(item.Str))
                {
                    request.Fields.Add(item.Str);
                }
            }
        }
        Rdv3Json filters = result.Member("filters");
        if (filters != null && filters.Kind == Rdv3Json.TArray)
        {
            for (int i = 0; i < filters.Count; i++)
            {
                Rdv3Json item = filters.At(i);
                if (item == null || item.Kind != Rdv3Json.TObject) { continue; }
                Rdv3ExportFilter filter = new Rdv3ExportFilter();
                filter.Field = Rdv3Form.Text(item, "field");
                filter.Operator = Rdv3Form.Text(item, "operator");
                filter.First = Rdv3Form.Text(item, "first");
                filter.Last = Rdv3Form.Text(item, "last");
                if (refs.Contains(filter.Field)) { request.Filters.Add(filter); }
            }
        }
        if (request.Path.Length == 0 || request.Fields.Count == 0)
        {
            owner.Error(Rdv3Text.ExportNeedField);
            return null;
        }
        return request;
    }
}

internal static class Rdv3ModalData
{
    public static string Candidates(
        Rdv3Form owner,
        Rdv3CandidatesDef def,
        List<Rdv3CandRow> source,
        int total,
        int selected,
        bool selectable,
        string title,
        string hint)
    {
        List<Rdv3CandRow> rows = source ?? new List<Rdv3CandRow>();
        StringBuilder sb = new StringBuilder(16384);
        sb.Append("{\"title\":").Append(Rdv3WebJson.Q(title));
        sb.Append(",\"hint\":").Append(Rdv3WebJson.Q(hint));
        sb.Append(",\"total\":").Append(total.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"selected\":").Append(selected.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"selectable\":").Append(Rdv3WebJson.B(selectable));
        sb.Append(",\"width\":").Append(Rdv3WebJson.N(def.Width));
        sb.Append(",\"maxHeight\":").Append(Rdv3WebJson.N(def.MaxHeight));
        sb.Append(",\"rowHeight\":").Append(Rdv3WebJson.N(def.RowHeight));
        sb.Append(",\"headerHeight\":").Append(Rdv3WebJson.N(def.HeaderHeight));
        sb.Append(",\"columns\":[");
        for (int i = 0; i < def.Columns.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            Rdv3ColumnDef column = def.Columns[i];
            sb.Append("{\"header\":").Append(Rdv3WebJson.Q(column.Header));
            sb.Append(",\"width\":").Append(Rdv3WebJson.N(column.Width));
            sb.Append(",\"align\":").Append(Rdv3WebJson.Q(column.Align));
            sb.Append(",\"bold\":").Append(Rdv3WebJson.B(column.Bold));
            sb.Append(",\"muted\":").Append(Rdv3WebJson.B(column.Muted));
            sb.Append(",\"render\":").Append(Rdv3WebJson.Q(column.Render)).Append('}');
        }
        sb.Append("],\"rows\":[");
        for (int i = 0; i < rows.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            Rdv3View view = new Rdv3View();
            view.Record = Rdv3Ledger.SplitLine(rows[i].Line);
            view.StoredState = rows[i].Stored ?? "";
            view.RowNumber = i + 1;
            sb.Append('[');
            for (int k = 0; k < def.Columns.Count; k++)
            {
                if (k > 0) { sb.Append(','); }
                Rdv3ColumnDef column = def.Columns[k];
                Rdv3Value value = Rdv3Eval.Evaluate(column.Value, view, owner.Fields, owner.Screen.Work);
                string look = "";
                if (column.Render == "tag")
                {
                    if (column.Value != null && column.Value.IsState && column.Value.State == "workStateShort")
                    {
                        look = Rdv3Eval.WorkStateLook(view, owner.Screen.Work);
                    }
                    else
                    {
                        string mapped;
                        if (!column.Looks.TryGetValue(value.Text, out mapped))
                        {
                            column.Looks.TryGetValue("*", out mapped);
                        }
                        look = mapped ?? "neutral";
                    }
                }
                sb.Append("{\"text\":").Append(Rdv3WebJson.Q(value.Text));
                sb.Append(",\"tone\":").Append(value.Tone.ToString(CultureInfo.InvariantCulture));
                sb.Append(",\"look\":").Append(Rdv3WebJson.Q(look)).Append('}');
            }
            sb.Append(']');
        }
        sb.Append("]}");
        return sb.ToString();
    }
}
