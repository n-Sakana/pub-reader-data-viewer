using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

// Evaluation uses the same parser and pipeline as the window. Only the report
// is written; shared locks, pending edits and the configured ledger are untouched.
public static class Rdv3Headless
{
    public static int Run(string appDir, string configPath, string dataDir, bool execute,
                          string outputPath, string baselinePath)
    {
        string stage = "settings file and definitions";
        try
        {
            Rdv3Log.Phase((execute ? "RunUpdate" : "ValidateOnly") + " settings " + configPath);
            Rdv3Config cfg = Rdv3Config.Load(configPath, !execute);
            stage = "paths, input files and job preparation";
            dataDir = Rdv3Files.Full(string.IsNullOrEmpty(dataDir) ? cfg.DataDir : dataDir, appDir);
            if (execute)
            {
                outputPath = Rdv3Files.ReportPath(outputPath, appDir, dataDir,
                    Rdv3Files.Full(cfg.Ledger, appDir), Rdv3Files.Full(cfg.Log, appDir), configPath, cfg.Data);
                if (!string.IsNullOrEmpty(baselinePath) && Rdv3Files.Same(outputPath, baselinePath))
                { throw new IOException("-Output must differ from -BaselineLedger"); }
            }
            string report = Evaluate(cfg, appDir, dataDir, execute, baselinePath);
            if (execute)
            {
                Rdv3Log.Phase("writing report " + outputPath);
                Rdv3Files.WriteNewText(outputPath, report, false);
            }
            Rdv3Log.Feedback("PASS", (execute ? "RunUpdate " : "ValidateOnly ") + configPath);
            // The console stays small; complete rows and intermediate values are
            // in the report, where a caller can compare them without a GUI.
            Rdv3Json parsed = Rdv3Json.Parse(report);
            Console.WriteLine("PASS " + (execute ? "RunUpdate " : "ValidateOnly ") + configPath);
            Rdv3Json summary = parsed.Member("summary");
            foreach (string key in summary.Order)
            { Report(key + "=" + (summary.Member(key).Kind == Rdv3Json.TNull ? "not-run" : summary.Member(key).Num.ToString(CultureInfo.InvariantCulture))); }
            foreach (Rdv3Json join in parsed.Member("joins").Items)
            {
                Report("JOIN " + join.Member("output").Str
                    + " unmatchedLeft=" + join.Member("unmatchedLeft").Num.ToString(CultureInfo.InvariantCulture)
                    + " unmatchedRight=" + join.Member("unmatchedRight").Num.ToString(CultureInfo.InvariantCulture));
            }
            foreach (Rdv3Json warning in parsed.Member("warnings").Items) { Report("WARNING " + warning.Str); }
            if (execute) { Report("OUTPUT " + outputPath); }
            return 0;
        }
        catch (Exception error)
        {
            StringBuilder feedback = new StringBuilder();
            Rdv3ValidationError validation = error as Rdv3ValidationError;
            if (validation == null)
            {
                feedback.AppendLine("FAIL 1 error");
                feedback.AppendLine("  " + configPath + ": " + error.Message);
                feedback.AppendLine("STOP " + stage + "; subsequent checks were not performed.");
            }
            else
            {
                feedback.AppendLine("FAIL " + N(validation.Errors.Length) + " errors");
                foreach (string detail in validation.Errors) { feedback.AppendLine("  " + configPath + ": " + detail); }
                feedback.AppendLine("STOP " + validation.Stage);
                foreach (string remaining in validation.Unchecked) { feedback.AppendLine("NOT CHECKED " + remaining); }
            }
            Rdv3Log.Feedback("VALIDATION", feedback.ToString());
            Rdv3Log.Error(stage, error);
            Console.Error.Write(feedback.ToString());
            return 3;
        }
    }

    public static string Evaluate(Rdv3Config cfg, string appDir, string dataDir, bool execute, string baselinePath)
    {
        Rdv3Config resolved = cfg.Clone();
        resolved.DataDir = dataDir;
        Rdv3Files.ValidateLayout(resolved, appDir, cfg.SourcePath);
        Rdv3Data data = cfg.Data;
        Rdv3Table[] tables = new Rdv3Table[data.Tables.Count];
        string[][] heads = new string[tables.Length][];
        List<string> warnings = new List<string>();
        Rdv3Validation validation = execute ? null : new Rdv3Validation();
        for (int i = 0; i < tables.Length; i++)
        {
            Rdv3TableDef def = data.Tables[i];
            Action read = delegate {
            Rdv3Log.Phase("reading input " + def.Id + " " + Rdv3Files.Full(def.File, dataDir));
            tables[i] = Rdv3Table.Read(Rdv3Files.Full(def.File, dataDir), def.Id, def.Enc,
                def.KeyColumns, def.KeyValidation, def.EncodingSetting, data.SourceReferences(def.Id));
            heads[i] = tables[i].Head;
            new Rdv3Index(tables[i]);
            tables[i].AddWarnings(warnings);
            };
            if (validation == null) { read(); }
            else { validation.Check("input " + def.Id + " (" + def.File + ")", read); }
        }
        if (validation != null) { validation.Finish("input files", "input columns/types and job preparation"); }
        Rdv3Log.Phase("input columns and types");
        data.Bind(heads, validation);
        data.ValidateTypes(tables, validation);
        if (validation != null) { validation.Finish("input types", "job preparation"); }
        Rdv3PreparedProcess update = null;
        List<Rdv3InputResult> inputs = new List<Rdv3InputResult>();
        for (int j = 0; j < data.Jobs.Count; j++)
        {
            Action prepare = delegate {
            Rdv3Log.Phase("preparing job " + data.Jobs[j].Id);
            Rdv3PreparedProcess prepared = Rdv3Process.Prepare(data, data.Jobs[j], dataDir, tables);
            warnings.AddRange(prepared.Warnings);
            if (data.Jobs[j] == data.UpdateJob) { update = prepared; inputs.AddRange(prepared.InputResults); }
            };
            if (validation == null) { prepare(); }
            else { validation.Check("job " + data.Jobs[j].Id, prepare); }
        }
        if (validation != null) { validation.Finish("job preparation", "job execution (use -RunUpdate after fixing the errors)"); }
        string[] before = new string[0], states = new string[0];
        if (execute && !string.IsNullOrEmpty(baselinePath))
        {
            Rdv3Log.Phase("reading baseline " + baselinePath);
            Rdv3LedgerSnapshot snapshot = new Rdv3LedgerStore(baselinePath, data, cfg.Screen.Work, null).Read(data.Head);
            before = snapshot.Lines; states = snapshot.States;
            if (snapshot.Warning.Length > 0) { warnings.Add(snapshot.Warning); }
        }
        Rdv3Log.Phase(execute ? "executing update job " + data.UpdateJob.Id : "validation finished");
        Rdv3ProcessResult result = execute
            ? Rdv3Process.Execute(update, before, states, cfg.Screen.Work.InitialStored, true) : null;
        if (result != null) { warnings.AddRange(result.Warnings); }
        List<string> reset = result == null ? new List<string>()
            : new Rdv3ResetNotice(data.IdentityCols, cfg.Screen.Work.InitialStored).ChangedRows(before, states, result.Lines, result.States);
        StringBuilder sb = new StringBuilder("{\"mode\":");
        sb.Append(Rdv3Json.Quote(execute ? "update" : "validate"));
        sb.Append(",\"job\":").Append(Rdv3Json.Quote(data.UpdateJob.Id));
        int empty = 0, duplicate = 0, shortRows = 0, blankRows = 0, columns = 0;
        foreach (Rdv3InputResult input in inputs)
        {
            empty += input.SkippedEmpty; duplicate += input.SkippedDuplicate;
            shortRows += input.SkippedShort; blankRows += input.SkippedBlank; columns += input.SkippedColumns;
        }
        sb.Append(",\"summary\":{\"rows\":").Append(result == null ? "null" : N(result.Lines.Length));
        sb.Append(",\"skippedEmpty\":").Append(N(empty)).Append(",\"skippedDuplicate\":").Append(N(duplicate));
        sb.Append(",\"skippedShort\":").Append(N(shortRows)).Append(",\"skippedBlank\":").Append(N(blankRows));
        sb.Append(",\"skippedColumns\":").Append(N(columns));
        sb.Append(",\"baselineRows\":").Append(N(before.Length));
        sb.Append(",\"resetRows\":").Append(result == null ? "null" : N(reset.Count)).Append('}');
        sb.Append(",\"inputs\":[");
        for (int i = 0; i < inputs.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            Rdv3InputResult input = inputs[i];
            sb.Append("{\"id\":").Append(Rdv3Json.Quote(input.Id)).Append(",\"file\":").Append(Rdv3Json.Quote(input.File));
            sb.Append(",\"rows\":").Append(N(input.Rows)).Append(",\"skippedEmpty\":").Append(N(input.SkippedEmpty));
            sb.Append(",\"skippedDuplicate\":").Append(N(input.SkippedDuplicate));
            sb.Append(",\"skippedShort\":").Append(N(input.SkippedShort)).Append(",\"skippedBlank\":").Append(N(input.SkippedBlank));
            sb.Append(",\"skippedColumns\":").Append(N(input.SkippedColumns)).Append('}');
        }
        sb.Append("],\"warnings\":").Append(Rdv3WebJson.S(new List<string>(new HashSet<string>(warnings)).ToArray()));
        sb.Append(",\"joins\":[");
        if (result != null)
        {
            for (int i = 0; i < result.Joins.Count; i++)
            {
                if (i > 0) { sb.Append(','); }
                Rdv3JoinResult join = result.Joins[i];
                sb.Append("{\"output\":").Append(Rdv3Json.Quote(join.Output));
                sb.Append(",\"leftRows\":").Append(N(join.LeftRows)).Append(",\"rightRows\":").Append(N(join.RightRows));
                sb.Append(",\"outputRows\":").Append(N(join.OutputRows));
                sb.Append(",\"unmatchedLeft\":").Append(N(join.UnmatchedLeft));
                sb.Append(",\"unmatchedRight\":").Append(N(join.UnmatchedRight)).Append('}');
            }
        }
        sb.Append("]");
        if (result != null)
        {
            sb.Append(",\"columns\":").Append(Rdv3WebJson.S(data.ColumnRefs));
            sb.Append(",\"rows\":").Append(Rows(result.Lines));
            sb.Append(",\"states\":").Append(Rdv3WebJson.S(result.States));
            sb.Append(",\"resetRows\":").Append(Rows(reset.ToArray()));
            sb.Append(",\"values\":{");
            bool comma = false;
            foreach (KeyValuePair<string, Rdv3ProcessValueResult> value in result.Values)
            {
                if (comma) { sb.Append(','); } comma = true;
                sb.Append(Rdv3Json.Quote(value.Key)).Append(":{\"kind\":").Append(Rdv3Json.Quote(value.Value.Kind));
                sb.Append(",\"count\":").Append(N(value.Value.Count));
                sb.Append(",\"columns\":").Append(Rdv3WebJson.S(value.Value.Columns));
                sb.Append(",\"rows\":").Append(Rows(value.Value.Lines)).Append('}');
            }
            sb.Append('}');
        }
        return sb.Append('}').ToString();
    }

    private static void Report(string text)
    {
        Rdv3Log.Feedback("RESULT", text);
        Console.WriteLine(text);
    }

    private static string N(int value) { return value.ToString(CultureInfo.InvariantCulture); }
    private static string Rows(string[] lines)
    {
        StringBuilder sb = new StringBuilder("[");
        for (int i = 0; i < lines.Length; i++)
        { if (i > 0) { sb.Append(','); } sb.Append(Rdv3WebJson.S(lines[i].Split('\t'))); }
        return sb.Append(']').ToString();
    }
}
