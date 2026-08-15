// ============================================================================
// Rdv3Config.cs -- everything the operator may change, in one file next to the
// .cmd. The distribution is two files:
//
//   ReaderDataViewer.cmd     the program
//   ReaderDataViewer.json    these settings
//
// What belongs here: anything that depends on the SITE rather than on the
// program -- which windows to watch, where the data and the ledger are, how
// long a job may take, how the key is shaped. What does NOT belong here: the
// screen's geometry (that is the UI reference, docs\ui-spec.md), the merge
// algorithm, and the save contract.
//
// A missing file is not an error: the built-in defaults are the ones written in
// the shipped ReaderDataViewer.json, and the app logs which it used. A file
// that cannot be parsed is reported on screen and in the log, and the defaults
// are used, because a settings typo must not stop the operator working.
// A single bad VALUE falls back on its own, so the rest of the file still
// applies.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

// ---------------------------------------------------------------------------
// One UIA element matcher. Empty members are not tested, so a matcher with
// nothing set matches anything -- which is why a target always needs at least
// one of automationId / className / name / nameLike / processName.
// ---------------------------------------------------------------------------
public sealed class Rdv3Match
{
    public string AutomationId = "";
    public string ClassName = "";
    public string Name = "";
    public string NameLike = "";          // * and ? wildcards
    public string ProcessName = "";       // without .exe
    public string[] ControlTypes = new string[0];
    public bool RequireValuePattern;
    public int Index;                     // which match to take when several fit
    public bool Descendants = true;       // "scope": "children" looks one level only

    public bool IsEmpty
    {
        get
        {
            return AutomationId.Length == 0 && ClassName.Length == 0 && Name.Length == 0
                && NameLike.Length == 0 && ProcessName.Length == 0 && ControlTypes.Length == 0
                && !RequireValuePattern;
        }
    }

    public static Rdv3Match Read(Rdv3Json o, List<string> notes, string where)
    {
        Rdv3Match m = new Rdv3Match();
        if (o == null || !o.IsObject) { return m; }
        m.AutomationId = o.GetString("automationId", "");
        m.ClassName = o.GetString("className", "");
        m.Name = o.GetString("name", "");
        m.NameLike = o.GetString("nameLike", "");
        m.ProcessName = o.GetString("processName", "");
        m.ControlTypes = o.GetStrings("controlTypes");
        m.RequireValuePattern = o.GetBool("requireValuePattern", false);
        m.Index = Math.Max(0, o.GetInt("index", 0));
        string scope = o.GetString("scope", "").Trim().ToLowerInvariant();
        if (scope == "children") { m.Descendants = false; }
        else if (scope == "descendants" || scope.Length == 0) { m.Descendants = true; }
        else { notes.Add(where + ": scope " + scope + " is not children or descendants"); }
        for (int i = 0; i < m.ControlTypes.Length; i++)
        {
            if (Rdv3Uia.ControlTypeByName(m.ControlTypes[i]) == null)
            {
                notes.Add(where + ": controlType " + m.ControlTypes[i] + " is not a name UI Automation knows");
            }
        }
        return m;
    }

    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        Add(sb, "automationId", AutomationId);
        Add(sb, "class", ClassName);
        Add(sb, "name", Name);
        Add(sb, "nameLike", NameLike);
        Add(sb, "process", ProcessName);
        if (ControlTypes.Length > 0) { Add(sb, "type", string.Join("/", ControlTypes)); }
        if (RequireValuePattern) { Add(sb, "value", "yes"); }
        if (Index > 0) { Add(sb, "index", Index.ToString(CultureInfo.InvariantCulture)); }
        Add(sb, "scope", Descendants ? "descendants" : "children");
        return (sb.Length == 0) ? "(anything)" : sb.ToString();
    }

    public string ToJson()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("{ ");
        J(sb, "automationId", AutomationId);
        J(sb, "className", ClassName);
        J(sb, "name", Name);
        J(sb, "nameLike", NameLike);
        J(sb, "processName", ProcessName);
        if (ControlTypes.Length > 0)
        {
            Comma(sb);
            sb.Append("\"controlTypes\": [");
            for (int i = 0; i < ControlTypes.Length; i++)
            {
                if (i > 0) { sb.Append(", "); }
                sb.Append(Rdv3Config.Q(ControlTypes[i]));
            }
            sb.Append("]");
        }
        if (RequireValuePattern) { Comma(sb); sb.Append("\"requireValuePattern\": true"); }
        if (Index > 0) { Comma(sb); sb.Append("\"index\": ").Append(Rdv3Config.N(Index)); }
        Comma(sb);
        sb.Append("\"scope\": ").Append(Rdv3Config.Q(Descendants ? "descendants" : "children"));
        sb.Append(" }");
        return sb.ToString();
    }

    private static void Comma(StringBuilder sb)
    {
        if (sb.Length > 2) { sb.Append(", "); }
    }

    private static void J(StringBuilder sb, string k, string v)
    {
        if (v == null || v.Length == 0) { return; }
        Comma(sb);
        sb.Append(Rdv3Config.Q(k)).Append(": ").Append(Rdv3Config.Q(v));
    }

    public Rdv3Match Clone()
    {
        Rdv3Match m = new Rdv3Match();
        m.AutomationId = AutomationId; m.ClassName = ClassName; m.Name = Name;
        m.NameLike = NameLike; m.ProcessName = ProcessName;
        m.ControlTypes = (string[])ControlTypes.Clone();
        m.RequireValuePattern = RequireValuePattern; m.Index = Index; m.Descendants = Descendants;
        return m;
    }

    private static void Add(StringBuilder sb, string k, string v)
    {
        if (v == null || v.Length == 0) { return; }
        if (sb.Length > 0) { sb.Append(" "); }
        sb.Append(k).Append("=").Append(v);
    }
}

// ---------------------------------------------------------------------------
// One thing to watch: an application WINDOW and, inside it, the FIELD whose
// text is read. Several targets may be listed and all of them are watched at
// once; whichever produces a confirmed key first drives the search.
// ---------------------------------------------------------------------------
public sealed class Rdv3Target
{
    public bool Enabled = true;
    public string Name = "";
    public Rdv3Match Window = new Rdv3Match();
    // Optional intermediate steps: a real application keeps the field several
    // levels down (tab -> pane -> grid -> cell), and naming those levels is
    // both faster and far less ambiguous than one descendant search.
    public List<Rdv3Match> Steps = new List<Rdv3Match>();
    public Rdv3Match Field = new Rdv3Match();
    public int ReadMode = Rdv3Uia.ReadValue;

    public Rdv3Target Clone()
    {
        Rdv3Target t = new Rdv3Target();
        t.Enabled = Enabled; t.Name = Name; t.ReadMode = ReadMode;
        t.Window = Window.Clone();
        t.Field = Field.Clone();
        t.Steps = new List<Rdv3Match>();
        for (int i = 0; i < Steps.Count; i++) { t.Steps.Add(Steps[i].Clone()); }
        return t;
    }

    public static Rdv3Target Read(Rdv3Json o, int i, List<string> notes)
    {
        Rdv3Target t = new Rdv3Target();
        string where = "watch.targets[" + i.ToString(CultureInfo.InvariantCulture) + "]";
        if (o == null || !o.IsObject) { notes.Add(where + " is not an object; ignored"); t.Enabled = false; return t; }
        t.Enabled = o.GetBool("enabled", true);
        t.Name = o.GetString("name", "");
        t.Window = Rdv3Match.Read(o.Member("window"), notes, where + ".window");
        t.Window.Descendants = false;                 // top level windows are children of the desktop
        Rdv3Json path = o.Member("path");
        if (path != null && path.IsArray)
        {
            for (int k = 0; k < path.Count; k++)
            {
                Rdv3Match step = Rdv3Match.Read(path.At(k), notes,
                    where + ".path[" + k.ToString(CultureInfo.InvariantCulture) + "]");
                if (step.IsEmpty)
                {
                    notes.Add(where + ".path[" + k.ToString(CultureInfo.InvariantCulture)
                        + "] matches nothing in particular; skipped");
                    continue;
                }
                t.Steps.Add(step);
            }
        }
        t.Field = Rdv3Match.Read(o.Member("field"), notes, where + ".field");
        t.ReadMode = Rdv3Uia.ReadModeByName(o.GetString("read", "value"));
        if (t.Window.IsEmpty)
        {
            notes.Add(where + ".window matches nothing in particular; the target is ignored");
            t.Enabled = false;
        }
        if (t.Field.IsEmpty)
        {
            notes.Add(where + ".field matches nothing in particular; the target is ignored");
            t.Enabled = false;
        }
        if (t.Name.Length == 0)
        {
            t.Name = (t.Window.ProcessName.Length > 0) ? t.Window.ProcessName : t.Window.ClassName;
        }
        return t;
    }
}

// ---------------------------------------------------------------------------
public sealed class Rdv3Config
{
    public const int Schema = 1;

    // paths, as written in the file; resolved against the .cmd's folder
    public string DataDir = "data";
    public string Ledger = "ReaderDataViewer-Ledger.xlsx";
    public string Log = "ReaderDataViewer.log";

    public int KeyLength = 8;
    public bool KeyDigitsOnly = true;

    public int PollMs = 40;
    public int StableMs = 120;
    public int RebindMs = 400;
    public bool PreferFocusedWindow = true;
    public List<Rdv3Target> Targets = new List<Rdv3Target>();

    public int CandidateRowsShown = 10;

    public int CheckTimeoutMs = 180000;
    public int SearchTimeoutMs = 30000;
    public int SaveTimeoutMs = 60000;
    public int MarkOverdueMs = 180000;
    public int PumpMs = 1000;

    public string SourcePath = "";
    public bool Loaded;
    public string Error = "";
    public readonly List<string> Notes = new List<string>();

    // ---- the built-in defaults, which are what the shipped file says -------
    public static Rdv3Config Defaults()
    {
        Rdv3Config c = new Rdv3Config();
        Rdv3Target t = new Rdv3Target();
        t.Name = Rdv3Text.LabelNotepad;
        t.Window.ClassName = "Notepad";
        t.Field.ControlTypes = new string[] { "Document", "Edit" };
        t.Field.RequireValuePattern = true;
        c.Targets.Add(t);
        return c;
    }

    public static Rdv3Config Load(string path)
    {
        Rdv3Config c = Defaults();
        c.SourcePath = (path == null) ? "" : path;
        if (path == null || path.Length == 0 || !File.Exists(path)) { return c; }
        string text;
        try { text = File.ReadAllText(path, new UTF8Encoding(false)); }
        catch (Exception ex) { c.Error = ex.Message; return c; }

        Rdv3Json root;
        try { root = Rdv3Json.Parse(text); }
        catch (Rdv3JsonError ex) { c.Error = ex.Message; return c; }
        catch (Exception ex) { c.Error = ex.Message; return c; }
        if (root == null || !root.IsObject) { c.Error = "the top level is not an object"; return c; }

        int schema = root.GetInt("schema", Schema);
        if (schema != Schema)
        {
            c.Notes.Add("schema " + schema.ToString(CultureInfo.InvariantCulture)
                + " is not the one this build knows (" + Schema.ToString(CultureInfo.InvariantCulture)
                + "); the file is read anyway");
        }

        Rdv3Json p = root.Member("paths");
        if (p != null && p.IsObject)
        {
            c.DataDir = p.GetString("dataDir", c.DataDir);
            c.Ledger = p.GetString("ledger", c.Ledger);
            c.Log = p.GetString("log", c.Log);
        }

        Rdv3Json k = root.Member("key");
        if (k != null && k.IsObject)
        {
            c.KeyLength = Clamp(k.GetInt("length", c.KeyLength), 1, 64, "key.length", c);
            c.KeyDigitsOnly = k.GetBool("digitsOnly", c.KeyDigitsOnly);
        }

        Rdv3Json w = root.Member("watch");
        if (w != null && w.IsObject)
        {
            c.PollMs = Clamp(w.GetInt("pollMs", c.PollMs), 5, 5000, "watch.pollMs", c);
            c.StableMs = Clamp(w.GetInt("stableMs", c.StableMs), 0, 60000, "watch.stableMs", c);
            c.RebindMs = Clamp(w.GetInt("rebindMs", c.RebindMs), 50, 60000, "watch.rebindMs", c);
            c.PreferFocusedWindow = w.GetBool("preferFocusedWindow", c.PreferFocusedWindow);
            Rdv3Json ts = w.Member("targets");
            if (ts != null && ts.IsArray)
            {
                List<Rdv3Target> list = new List<Rdv3Target>();
                for (int i = 0; i < ts.Count; i++)
                {
                    Rdv3Target t = Rdv3Target.Read(ts.At(i), i, c.Notes);
                    if (t.Enabled) { list.Add(t); }
                }
                if (list.Count == 0) { c.Notes.Add("watch.targets has nothing usable; the built-in target is kept"); }
                else { c.Targets = list; }
            }
        }

        Rdv3Json s = root.Member("search");
        if (s != null && s.IsObject)
        {
            c.CandidateRowsShown = Clamp(s.GetInt("candidateRowsShown", c.CandidateRowsShown),
                1, 1000, "search.candidateRowsShown", c);
        }

        Rdv3Json j = root.Member("jobs");
        if (j != null && j.IsObject)
        {
            c.CheckTimeoutMs = Clamp(j.GetInt("checkTimeoutMs", c.CheckTimeoutMs), 1000, 3600000, "jobs.checkTimeoutMs", c);
            c.SearchTimeoutMs = Clamp(j.GetInt("searchTimeoutMs", c.SearchTimeoutMs), 1000, 3600000, "jobs.searchTimeoutMs", c);
            c.SaveTimeoutMs = Clamp(j.GetInt("saveTimeoutMs", c.SaveTimeoutMs), 1000, 3600000, "jobs.saveTimeoutMs", c);
            c.MarkOverdueMs = Clamp(j.GetInt("markOverdueMs", c.MarkOverdueMs), 1000, 3600000, "jobs.markOverdueMs", c);
            c.PumpMs = Clamp(j.GetInt("pumpMs", c.PumpMs), 100, 60000, "jobs.pumpMs", c);
        }

        c.Loaded = true;
        return c;
    }

    private static int Clamp(int v, int lo, int hi, string what, Rdv3Config c)
    {
        if (v < lo || v > hi)
        {
            int fixedTo = (v < lo) ? lo : hi;
            c.Notes.Add(what + " " + v.ToString(CultureInfo.InvariantCulture) + " is out of range ["
                + lo.ToString(CultureInfo.InvariantCulture) + ".."
                + hi.ToString(CultureInfo.InvariantCulture) + "]; using "
                + fixedTo.ToString(CultureInfo.InvariantCulture));
            return fixedTo;
        }
        return v;
    }

    // ---- writing it back ---------------------------------------------------
    // The settings dialog writes the whole file. Comments a person added are not
    // kept (the reader accepts them; the writer does not invent them), so the
    // header says who wrote it and when.
    public string Save(string path)
    {
        try
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("// Reader Data Viewer settings\r\n");
            sb.Append("// written by the app on ").Append(DateTime.Now.ToString(
                "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)).Append("\r\n");
            sb.Append("// what each member means: docs/settings.md\r\n");
            sb.Append("{\r\n");
            sb.Append("  \"schema\": ").Append(N(Schema)).Append(",\r\n");
            sb.Append("  \"paths\": {\r\n");
            sb.Append("    \"dataDir\": ").Append(Q(DataDir)).Append(",\r\n");
            sb.Append("    \"ledger\": ").Append(Q(Ledger)).Append(",\r\n");
            sb.Append("    \"log\": ").Append(Q(Log)).Append("\r\n");
            sb.Append("  },\r\n");
            sb.Append("  \"key\": { \"length\": ").Append(N(KeyLength));
            sb.Append(", \"digitsOnly\": ").Append(B(KeyDigitsOnly)).Append(" },\r\n");
            sb.Append("  \"search\": { \"candidateRowsShown\": ").Append(N(CandidateRowsShown)).Append(" },\r\n");
            sb.Append("  \"jobs\": {\r\n");
            sb.Append("    \"checkTimeoutMs\": ").Append(N(CheckTimeoutMs)).Append(",\r\n");
            sb.Append("    \"searchTimeoutMs\": ").Append(N(SearchTimeoutMs)).Append(",\r\n");
            sb.Append("    \"saveTimeoutMs\": ").Append(N(SaveTimeoutMs)).Append(",\r\n");
            sb.Append("    \"markOverdueMs\": ").Append(N(MarkOverdueMs)).Append(",\r\n");
            sb.Append("    \"pumpMs\": ").Append(N(PumpMs)).Append("\r\n");
            sb.Append("  },\r\n");
            sb.Append("  \"watch\": {\r\n");
            sb.Append("    \"pollMs\": ").Append(N(PollMs));
            sb.Append(", \"stableMs\": ").Append(N(StableMs));
            sb.Append(", \"rebindMs\": ").Append(N(RebindMs));
            sb.Append(", \"preferFocusedWindow\": ").Append(B(PreferFocusedWindow)).Append(",\r\n");
            sb.Append("    \"targets\": [\r\n");
            for (int i = 0; i < Targets.Count; i++)
            {
                Rdv3Target t = Targets[i];
                sb.Append("      {\r\n");
                sb.Append("        \"enabled\": ").Append(B(t.Enabled)).Append(",\r\n");
                sb.Append("        \"name\": ").Append(Q(t.Name)).Append(",\r\n");
                sb.Append("        \"window\": ").Append(t.Window.ToJson()).Append(",\r\n");
                sb.Append("        \"path\": [");
                for (int k = 0; k < t.Steps.Count; k++)
                {
                    if (k > 0) { sb.Append(","); }
                    sb.Append("\r\n          ").Append(t.Steps[k].ToJson());
                }
                sb.Append((t.Steps.Count > 0) ? "\r\n        ]" : "]").Append(",\r\n");
                sb.Append("        \"field\": ").Append(t.Field.ToJson()).Append(",\r\n");
                sb.Append("        \"read\": ").Append(Q(Rdv3Uia.ReadModeName(t.ReadMode))).Append("\r\n");
                sb.Append("      }").Append((i + 1 < Targets.Count) ? "," : "").Append("\r\n");
            }
            sb.Append("    ]\r\n");
            sb.Append("  }\r\n");
            sb.Append("}\r\n");
            File.WriteAllText(path, sb.ToString(), new UTF8Encoding(false));
            SourcePath = path;
            Loaded = true;
            Error = "";
            return null;
        }
        catch (Exception ex) { return ex.Message; }
    }

    internal static string Q(string s)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append('"');
        if (s != null)
        {
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"' || c == '\\') { sb.Append('\\').Append(c); }
                else if (c == '\n') { sb.Append("\\n"); }
                else if (c == '\r') { sb.Append("\\r"); }
                else if (c == '\t') { sb.Append("\\t"); }
                else if (c < ' ') { sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture)); }
                else { sb.Append(c); }
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    internal static string N(int v) { return v.ToString(CultureInfo.InvariantCulture); }

    internal static string B(bool v) { return v ? "true" : "false"; }

    // What a running app may adopt without a restart. Paths are deliberately
    // left out: the ledger and the CSVs are read at start-up, and swapping them
    // under a live session would be a different program.
    public void AdoptRuntimeFrom(Rdv3Config o)
    {
        if (o == null) { return; }
        KeyLength = o.KeyLength;
        KeyDigitsOnly = o.KeyDigitsOnly;
        PollMs = o.PollMs;
        StableMs = o.StableMs;
        RebindMs = o.RebindMs;
        PreferFocusedWindow = o.PreferFocusedWindow;
        Targets = o.Targets;
        CandidateRowsShown = o.CandidateRowsShown;
        CheckTimeoutMs = o.CheckTimeoutMs;
        SearchTimeoutMs = o.SearchTimeoutMs;
        SaveTimeoutMs = o.SaveTimeoutMs;
        MarkOverdueMs = o.MarkOverdueMs;
        PumpMs = o.PumpMs;
    }

    // deep enough for the dialog to edit without touching the running settings
    // until the operator says OK
    public Rdv3Config Clone()
    {
        Rdv3Config c = new Rdv3Config();
        c.DataDir = DataDir; c.Ledger = Ledger; c.Log = Log;
        c.KeyLength = KeyLength; c.KeyDigitsOnly = KeyDigitsOnly;
        c.PollMs = PollMs; c.StableMs = StableMs; c.RebindMs = RebindMs;
        c.PreferFocusedWindow = PreferFocusedWindow;
        c.CandidateRowsShown = CandidateRowsShown;
        c.CheckTimeoutMs = CheckTimeoutMs; c.SearchTimeoutMs = SearchTimeoutMs;
        c.SaveTimeoutMs = SaveTimeoutMs; c.MarkOverdueMs = MarkOverdueMs; c.PumpMs = PumpMs;
        c.SourcePath = SourcePath; c.Loaded = Loaded;
        c.Targets = new List<Rdv3Target>();
        for (int i = 0; i < Targets.Count; i++) { c.Targets.Add(Targets[i].Clone()); }
        return c;
    }

    // one line for the log, so a support question starts from what was read
    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(Loaded ? "loaded " : "defaults ");
        sb.Append(SourcePath);
        if (Error.Length > 0) { sb.Append(" ERROR=").Append(Error); }
        sb.Append(" key=").Append(KeyLength.ToString(CultureInfo.InvariantCulture));
        sb.Append(KeyDigitsOnly ? "digits" : "any");
        sb.Append(" poll=").Append(PollMs.ToString(CultureInfo.InvariantCulture));
        sb.Append("/").Append(StableMs.ToString(CultureInfo.InvariantCulture));
        sb.Append(" targets=").Append(Targets.Count.ToString(CultureInfo.InvariantCulture));
        for (int i = 0; i < Targets.Count; i++)
        {
            Rdv3Target t = Targets[i];
            sb.Append(" [").Append(t.Name).Append(": win ").Append(t.Window.Describe());
            for (int k = 0; k < t.Steps.Count; k++)
            {
                sb.Append(" > ").Append(t.Steps[k].Describe());
            }
            sb.Append(" > field ").Append(t.Field.Describe());
            sb.Append(" read=").Append(Rdv3Uia.ReadModeName(t.ReadMode)).Append("]");
        }
        return sb.ToString();
    }
}
