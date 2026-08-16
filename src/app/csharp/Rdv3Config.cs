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

    // Nothing effective in it -- and an empty matcher matches ANYTHING, so this
    // is the question "would this matcher accept any window on the desktop".
    public bool IsEmpty
    {
        get
        {
            return AutomationId.Length == 0 && ClassName.Length == 0 && Name.Length == 0
                && NameLike.Length == 0 && ProcessName.Length == 0 && KnownControlTypes == 0
                && !RequireValuePattern;
        }
    }

    // How many of the listed control types UI Automation actually knows. A name
    // it does not know is dropped from the search condition (Rdv3Uia.ToCondition
    // keeps only the ones it can resolve), so counting it as a constraint here
    // would let a single typo -- "Edti" for "Edit" -- pass for a narrowed matcher
    // and leave one that in fact accepts every window there is.
    public int KnownControlTypes
    {
        get
        {
            int n = 0;
            for (int i = 0; i < ControlTypes.Length; i++)
            {
                if (Rdv3Uia.ControlTypeByName(ControlTypes[i]) != null) { n++; }
            }
            return n;
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

    // Why this target cannot be watched, or "" when it can. A matcher with
    // nothing effective in it matches any window at all, so a target that names
    // neither a window nor a field would bind to whatever happened to be in
    // front and read a number off a screen nobody pointed it at. Asked here,
    // where the file is read, and again at the watcher -- the settings dialog
    // builds targets in memory and never comes through Read().
    public string WhyNotWatchable()
    {
        if (Window.IsEmpty) { return "window matches nothing in particular"; }
        if (Field.IsEmpty) { return "field matches nothing in particular"; }
        return "";
    }

    // Turned on AND able to work: the set that is actually being watched. It is
    // smaller than cfg.Targets, which also holds the ones the operator switched
    // off -- those stay in the file and in the dialog, but the status line is not
    // waiting for them and the watcher does not bind them. One property, so the
    // screen and the watcher cannot come to different answers.
    public bool IsWatchable
    {
        get { return Enabled && WhyNotWatchable().Length == 0; }
    }

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
        if (o == null || !o.IsObject) { notes.Add(where + " is not an object; ignored"); return t; }
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
        // "unusable" is not "turned off": Enabled stays exactly as the operator
        // wrote it, and Load decides what to do with a target that cannot work
        string why = t.WhyNotWatchable();
        if (why.Length > 0) { notes.Add(where + ": " + why + "; the target is ignored"); }
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

        // the built-in default of every member, kept aside: a value that is out
        // of range falls back to its own default, not to the nearest limit
        Rdv3Config d = Defaults();

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
            c.KeyLength = Ranged(k.GetInt("length", c.KeyLength), 1, 64, d.KeyLength, "key.length", c);
            c.KeyDigitsOnly = k.GetBool("digitsOnly", c.KeyDigitsOnly);
        }

        Rdv3Json w = root.Member("watch");
        if (w != null && w.IsObject)
        {
            c.PollMs = Ranged(w.GetInt("pollMs", c.PollMs), 5, 5000, d.PollMs, "watch.pollMs", c);
            c.StableMs = Ranged(w.GetInt("stableMs", c.StableMs), 0, 60000, d.StableMs, "watch.stableMs", c);
            c.RebindMs = Ranged(w.GetInt("rebindMs", c.RebindMs), 50, 60000, d.RebindMs, "watch.rebindMs", c);
            c.PreferFocusedWindow = w.GetBool("preferFocusedWindow", c.PreferFocusedWindow);
            Rdv3Json ts = w.Member("targets");
            if (ts != null && ts.IsArray)
            {
                // The file answered the question, so the file decides -- down to
                // "watch nothing". A target the operator turned OFF is kept: it
                // is still theirs, the dialog still lists it, and the next save
                // still writes it. Only one that cannot work at all is dropped.
                // The built-in target belongs to a file that is missing or does
                // not say, never to one that does.
                List<Rdv3Target> list = new List<Rdv3Target>();
                for (int i = 0; i < ts.Count; i++)
                {
                    Rdv3Target t = Rdv3Target.Read(ts.At(i), i, c.Notes);
                    if (t.WhyNotWatchable().Length == 0) { list.Add(t); }
                }
                if (list.Count == 0) { c.Notes.Add("watch.targets names nothing that can be watched; nothing is watched"); }
                c.Targets = list;
            }
        }

        Rdv3Json s = root.Member("search");
        if (s != null && s.IsObject)
        {
            c.CandidateRowsShown = Ranged(s.GetInt("candidateRowsShown", c.CandidateRowsShown),
                1, 1000, d.CandidateRowsShown, "search.candidateRowsShown", c);
        }

        Rdv3Json j = root.Member("jobs");
        if (j != null && j.IsObject)
        {
            c.CheckTimeoutMs = Ranged(j.GetInt("checkTimeoutMs", c.CheckTimeoutMs), 1000, 3600000, d.CheckTimeoutMs, "jobs.checkTimeoutMs", c);
            c.SearchTimeoutMs = Ranged(j.GetInt("searchTimeoutMs", c.SearchTimeoutMs), 1000, 3600000, d.SearchTimeoutMs, "jobs.searchTimeoutMs", c);
            c.SaveTimeoutMs = Ranged(j.GetInt("saveTimeoutMs", c.SaveTimeoutMs), 1000, 3600000, d.SaveTimeoutMs, "jobs.saveTimeoutMs", c);
            c.MarkOverdueMs = Ranged(j.GetInt("markOverdueMs", c.MarkOverdueMs), 1000, 3600000, d.MarkOverdueMs, "jobs.markOverdueMs", c);
            c.PumpMs = Ranged(j.GetInt("pumpMs", c.PumpMs), 100, 60000, d.PumpMs, "jobs.pumpMs", c);
        }

        c.Loaded = true;
        return c;
    }

    // Out of range means THIS member falls back on its own default, which is what
    // the shipped file and docs\settings.md promise. The nearest limit would be a
    // different answer to a different question: a typed "0" for jobs.checkTimeoutMs
    // is not a request for the shortest legal timeout, and 1000 ms would cut off a
    // start-up check that normally runs longer than that.
    private static int Ranged(int v, int lo, int hi, int def, string what, Rdv3Config c)
    {
        if (v < lo || v > hi)
        {
            c.Notes.Add(what + " " + v.ToString(CultureInfo.InvariantCulture) + " is out of range ["
                + lo.ToString(CultureInfo.InvariantCulture) + ".."
                + hi.ToString(CultureInfo.InvariantCulture) + "]; using the default "
                + def.ToString(CultureInfo.InvariantCulture));
            return def;
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

    // What a running app may adopt without a restart. Paths and key.length are
    // deliberately left out: the ledger and the CSVs are read at start-up, WITH
    // the key length they were read with, and swapping either under a live
    // session would be a different program -- a shorter key would leave the
    // screen asking for numbers the index that is already built cannot answer.
    // key.digitsOnly is NOT in that group: it only decides which characters make
    // a key, nothing built at start-up depends on it, so it applies at once.
    public void AdoptRuntimeFrom(Rdv3Config o)
    {
        if (o == null) { return; }
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

    // What the file now says for the members that only take effect at the next
    // start. The running session goes on using what it started with -- those
    // values live elsewhere by then (the resolved paths in Rdv3App, the key rule
    // in Rdv3Spec) -- but this object is also what the settings dialog opens on
    // and what the next save writes. Without this, the dialog would come back up
    // showing the OLD paths and the next save would quietly undo the change the
    // operator had just made.
    public void AdoptSavedFrom(Rdv3Config o)
    {
        if (o == null) { return; }
        DataDir = o.DataDir;
        Ledger = o.Ledger;
        Log = o.Log;
        KeyLength = o.KeyLength;
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
        // targets= is every target in the file, including the ones that are
        // switched off; watched= is how many of them are actually being looked
        // at. A support question that starts from this line has to be able to
        // tell those apart, and each entry says which it is.
        int watched = 0;
        for (int i = 0; i < Targets.Count; i++) { if (Targets[i].IsWatchable) { watched++; } }
        sb.Append(" targets=").Append(Targets.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" watched=").Append(watched.ToString(CultureInfo.InvariantCulture));
        for (int i = 0; i < Targets.Count; i++)
        {
            Rdv3Target t = Targets[i];
            sb.Append(" [").Append(t.Name);
            if (!t.Enabled) { sb.Append(" OFF"); }
            else if (t.WhyNotWatchable().Length > 0) { sb.Append(" UNUSABLE"); }
            sb.Append(": win ").Append(t.Window.Describe());
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
