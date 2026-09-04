// ============================================================================
// Rdv3Config.cs -- settings.json, the ONE file next to the .cmd. The
// distribution is the program and this file:
//
//   ReaderDataViewer.cmd     the program
//   settings.json            paths, search, watch, jobs   (Rdv3Config, below)
//                            data                         (Rdv3Data.cs)
//                            screen                       (Rdv3Screen.cs)
//
// The file is read once, at start-up, and it is read STRICTLY: it must exist,
// parse, and pass every check, or the app does not start -- with the file,
// the line and the reason on screen and in the log. There is no built-in
// default to fall back on, no member that is quietly ignored, no value that is
// quietly corrected. A file that half-applies would leave the operator
// believing an edit is in effect when it is not.
//
// The settings dialog writes back only the members it edits (paths, search,
// watch). It re-reads the file first -- a file that no longer loads is not
// written over -- and then replaces just the text of those three members, so
// everything else in the file, comments included, stays byte for byte.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

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

    // How many of the listed control types UI Automation actually knows. The
    // reader refuses a name it does not know, so for a matcher that came from
    // the file this is ControlTypes.Length; the picker builds matchers in
    // memory and the count is asked again there.
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

    public static Rdv3Match Read(Rdv3Json o)
    {
        o.Only("automationId", "className", "name", "nameLike", "processName", "controlTypes",
               "requireValuePattern", "index", "scope");
        Rdv3Match m = new Rdv3Match();
        m.AutomationId = o.StrOr("automationId", "");
        m.ClassName = o.StrOr("className", "");
        m.Name = o.StrOr("name", "");
        m.NameLike = o.StrOr("nameLike", "");
        m.ProcessName = o.StrOr("processName", "");
        m.ControlTypes = o.Strs("controlTypes", false);
        m.RequireValuePattern = o.BoolOr("requireValuePattern", false);
        m.Index = o.IntOr("index", 0, 0, 1000);
        m.Descendants = (o.Word("scope", "descendants", "descendants", "children") == "descendants");
        for (int i = 0; i < m.ControlTypes.Length; i++)
        {
            if (Rdv3Uia.ControlTypeByName(m.ControlTypes[i]) == null)
            {
                throw o.Member("controlTypes").At(i).Fail(m.ControlTypes[i] + " is not a control type UI Automation knows");
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
    // front and read a number off a screen nobody pointed it at. The file
    // reader refuses such a target; the settings dialog builds targets in
    // memory and asks here before it lists them as watched.
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

    public static Rdv3Target Read(Rdv3Json o)
    {
        o.Only("enabled", "name", "window", "path", "field", "read");
        Rdv3Target t = new Rdv3Target();
        t.Enabled = o.BoolOr("enabled", true);
        t.Name = o.StrOr("name", "");
        t.Window = Rdv3Match.Read(o.Obj("window", true));
        t.Window.Descendants = false;                 // top level windows are children of the desktop
        List<Rdv3Json> path = o.Objs("path", false);
        for (int k = 0; k < path.Count; k++)
        {
            Rdv3Match step = Rdv3Match.Read(path[k]);
            if (step.IsEmpty) { throw path[k].Fail("matches nothing in particular"); }
            t.Steps.Add(step);
        }
        t.Field = Rdv3Match.Read(o.Obj("field", true));
        t.ReadMode = Rdv3Uia.ReadModeByName(o.Word("read", "value", "value", "text", "name"));
        string why = t.WhyNotWatchable();
        if (why.Length > 0) { throw o.Fail(why); }
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
    public const int Schema = 3;

    // paths, as written in the file; resolved against the .cmd's folder
    public string DataDir = "data";
    public string Ledger = "ReaderDataViewer-Ledger.xlsx";
    public string Log = "ReaderDataViewer.log";

    // what counts as a number: a typed or watched value is confirmed only when
    // the whole of it matches this regular expression
    public const string DefaultKeyPattern = "^[0-9]{8}$";
    public string KeyPattern = DefaultKeyPattern;
    private Regex keyRule;

    public int PollMs = 40;
    public int StableMs = 120;
    public int RebindMs = 400;
    public bool PreferFocusedWindow = true;
    public List<Rdv3Target> Targets = new List<Rdv3Target>();

    public int CandidateRowsShown = 100;

    public int CheckTimeoutMs = 180000;
    public int SearchTimeoutMs = 30000;
    public int SaveTimeoutMs = 60000;
    public int MarkOverdueMs = 180000;
    public int PumpMs = 1000;
    public int LockRetryMs = 250;
    public int LockStaleMs = 600000;
    public int MarkerPollMs = 3000;

    // the other two parts of the same file
    public Rdv3Data Data;
    public Rdv3Screen Screen;

    public string SourcePath = "";

    // where the three members the dialog owns sit in the file's text
    private Rdv3Json pathsNode, searchNode, watchNode;

    // ---- reading -----------------------------------------------------------
    // The whole file, or an Rdv3LoadError that says what is wrong and where.
    public static Rdv3Config Load(string path)
    {
        if (path == null || path.Length == 0 || !File.Exists(path))
        {
            throw new Rdv3LoadError("the file does not exist: " + path, 0);
        }
        string text;
        try { text = File.ReadAllText(path, new UTF8Encoding(false)); }
        catch (Exception ex) { throw new Rdv3LoadError("the file cannot be read: " + ex.Message, 0); }
        Rdv3Config c = Parse(text);
        c.SourcePath = path;
        return c;
    }

    public static Rdv3Config Parse(string text)
    {
        Rdv3Json root = Rdv3Json.Parse(text);
        if (!root.IsObject) { throw new Rdv3LoadError("the top level is not an object", root.Line); }
        root.Only("schema", "paths", "search", "watch", "jobs", "data", "screen");
        Rdv3Config c = new Rdv3Config();

        int schema = root.Int("schema", 1, 1000);
        if (schema != Schema)
        {
            throw root.Member("schema").Fail("this program reads schema " + N(Schema)
                + "; the file says " + N(schema) + " (docs/settings.md)");
        }

        Rdv3Json p = root.Obj("paths", true);
        p.Only("dataDir", "ledger", "log");
        c.DataDir = p.StrOr("dataDir", c.DataDir);
        c.Ledger = p.StrOr("ledger", c.Ledger);
        c.Log = p.StrOr("log", c.Log);
        if (c.DataDir.Trim().Length == 0) { throw p.Member("dataDir").Fail("must not be blank"); }
        if (c.Ledger.Trim().Length == 0) { throw p.Member("ledger").Fail("must not be blank"); }
        if (c.Log.Trim().Length == 0) { throw p.Member("log").Fail("must not be blank"); }
        c.pathsNode = p;

        Rdv3Json s = root.Obj("search", true);
        s.Only("pattern", "candidateRowsShown");
        c.KeyPattern = s.StrOr("pattern", DefaultKeyPattern);
        string why = PatternError(c.KeyPattern);
        if (why != null) { throw s.Member("pattern").Fail("is not a usable regular expression (" + why + ")"); }
        c.CandidateRowsShown = s.IntOr("candidateRowsShown", c.CandidateRowsShown, 1, 1000);
        c.searchNode = s;

        Rdv3Json w = root.Obj("watch", true);
        w.Only("pollMs", "stableMs", "rebindMs", "preferFocusedWindow", "targets");
        c.PollMs = w.IntOr("pollMs", c.PollMs, 5, 5000);
        c.StableMs = w.IntOr("stableMs", c.StableMs, 0, 60000);
        c.RebindMs = w.IntOr("rebindMs", c.RebindMs, 50, 60000);
        c.PreferFocusedWindow = w.BoolOr("preferFocusedWindow", c.PreferFocusedWindow);
        // the file decides what is watched -- down to "nothing" (an empty list).
        // A target the operator turned OFF is kept: it is still theirs, the
        // dialog still lists it, and the next save still writes it.
        List<Rdv3Json> ts = w.Objs("targets", true);
        for (int i = 0; i < ts.Count; i++) { c.Targets.Add(Rdv3Target.Read(ts[i])); }
        c.watchNode = w;

        Rdv3Json j = root.Obj("jobs", true);
        j.Only("checkTimeoutMs", "searchTimeoutMs", "saveTimeoutMs", "markOverdueMs", "pumpMs", "lockRetryMs", "lockStaleMs", "markerPollMs");
        c.CheckTimeoutMs = j.IntOr("checkTimeoutMs", c.CheckTimeoutMs, 1000, 3600000);
        c.SearchTimeoutMs = j.IntOr("searchTimeoutMs", c.SearchTimeoutMs, 1000, 3600000);
        c.SaveTimeoutMs = j.IntOr("saveTimeoutMs", c.SaveTimeoutMs, 1000, 3600000);
        c.MarkOverdueMs = j.IntOr("markOverdueMs", c.MarkOverdueMs, 1000, 3600000);
        c.PumpMs = j.IntOr("pumpMs", c.PumpMs, 100, 60000);
        c.LockRetryMs = j.IntOr("lockRetryMs", c.LockRetryMs, 50, 5000);
        c.LockStaleMs = j.IntOr("lockStaleMs", c.LockStaleMs, 60000, 86400000);
        c.MarkerPollMs = j.IntOr("markerPollMs", c.MarkerPollMs, 500, 60000);

        c.Data = Rdv3Data.Read(root.Obj("data", true));
        c.Screen = Rdv3Screen.Read(root.Obj("screen", true));
        // the screen names ledger columns; they have to be the data's
        c.Screen.Check(c.Data);
        return c;
    }

    // the compiled rule behind KeyPattern, rebuilt when the pattern changes
    public Regex KeyRule
    {
        get
        {
            if (keyRule == null || keyRule.ToString() != KeyPattern)
            {
                keyRule = new Regex(KeyPattern, RegexOptions.CultureInvariant);
            }
            return keyRule;
        }
    }

    public bool IsKey(string s)
    {
        return s != null && s.Length > 0 && KeyRule.IsMatch(s);
    }

    // null when the pattern compiles, otherwise the reason it does not
    public static string PatternError(string pattern)
    {
        if (pattern == null || pattern.Trim().Length == 0) { return "empty"; }
        try { new Regex(pattern, RegexOptions.CultureInvariant); return null; }
        catch (Exception ex) { return ex.Message; }
    }

    // ---- writing it back ---------------------------------------------------
    // Only paths, search and watch are the dialog's to write. The file is read
    // again first: if it no longer loads (someone edited it meanwhile and broke
    // it), nothing is written and the reason comes back. Otherwise the text of
    // those three members is replaced in place and every other byte of the
    // file is left exactly as it was. Returns null on success, else the reason.
    public string Save(string path)
    {
        try
        {
            string text = File.ReadAllText(path, new UTF8Encoding(false));
            Rdv3Config now = Parse(text);
            string nl = (text.IndexOf("\r\n", StringComparison.Ordinal) >= 0) ? "\r\n" : "\n";
            // latest offset first, so the earlier spans stay where they are
            Rdv3Json[] nodes = { now.pathsNode, now.searchNode, now.watchNode };
            string[] bodies = { PathsJson(IndentOf(text, now.pathsNode), nl),
                                SearchJson(),
                                WatchJson(IndentOf(text, now.watchNode), nl) };
            for (int a = 0; a < nodes.Length; a++)
            {
                for (int b = a + 1; b < nodes.Length; b++)
                {
                    if (nodes[b].Start > nodes[a].Start)
                    {
                        Rdv3Json tn = nodes[a]; nodes[a] = nodes[b]; nodes[b] = tn;
                        string tb = bodies[a]; bodies[a] = bodies[b]; bodies[b] = tb;
                    }
                }
            }
            for (int i = 0; i < nodes.Length; i++)
            {
                text = text.Substring(0, nodes[i].Start) + bodies[i] + text.Substring(nodes[i].End);
            }
            // the result has to load too, or it is not written
            Parse(text);
            string tmp = path + ".tmp-" + DateTime.Now.Ticks.ToString(CultureInfo.InvariantCulture);
            File.WriteAllText(tmp, text, new UTF8Encoding(false));
            if (File.Exists(path)) { File.Replace(tmp, path, null); }
            else { File.Move(tmp, path); }
            SourcePath = path;
            return null;
        }
        catch (Exception ex) { return ex.Message; }
    }

    // the whitespace that opens the line the member's name is on
    private static string IndentOf(string text, Rdv3Json node)
    {
        int at = (node.KeyStart >= 0) ? node.KeyStart : node.Start;
        int ls = at;
        while (ls > 0 && text[ls - 1] != '\n') { ls--; }
        int e = ls;
        while (e < at && (text[e] == ' ' || text[e] == '\t')) { e++; }
        return text.Substring(ls, e - ls);
    }

    private string PathsJson(string ind, string nl)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("{").Append(nl);
        sb.Append(ind).Append("  \"dataDir\": ").Append(Q(DataDir)).Append(",").Append(nl);
        sb.Append(ind).Append("  \"ledger\": ").Append(Q(Ledger)).Append(",").Append(nl);
        sb.Append(ind).Append("  \"log\": ").Append(Q(Log)).Append(nl);
        sb.Append(ind).Append("}");
        return sb.ToString();
    }

    private string SearchJson()
    {
        return "{ \"pattern\": " + Q(KeyPattern) + ", \"candidateRowsShown\": " + N(CandidateRowsShown) + " }";
    }

    private string WatchJson(string ind, string nl)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("{").Append(nl);
        sb.Append(ind).Append("  \"pollMs\": ").Append(N(PollMs));
        sb.Append(", \"stableMs\": ").Append(N(StableMs));
        sb.Append(", \"rebindMs\": ").Append(N(RebindMs));
        sb.Append(", \"preferFocusedWindow\": ").Append(B(PreferFocusedWindow)).Append(",").Append(nl);
        sb.Append(ind).Append("  \"targets\": [").Append(nl);
        for (int i = 0; i < Targets.Count; i++)
        {
            Rdv3Target t = Targets[i];
            string ti = ind + "    ";
            sb.Append(ti).Append("{").Append(nl);
            sb.Append(ti).Append("  \"enabled\": ").Append(B(t.Enabled)).Append(",").Append(nl);
            sb.Append(ti).Append("  \"name\": ").Append(Q(t.Name)).Append(",").Append(nl);
            sb.Append(ti).Append("  \"window\": ").Append(t.Window.ToJson()).Append(",").Append(nl);
            sb.Append(ti).Append("  \"path\": [");
            for (int k = 0; k < t.Steps.Count; k++)
            {
                if (k > 0) { sb.Append(","); }
                sb.Append(nl).Append(ti).Append("    ").Append(t.Steps[k].ToJson());
            }
            if (t.Steps.Count > 0) { sb.Append(nl).Append(ti).Append("  "); }
            sb.Append("],").Append(nl);
            sb.Append(ti).Append("  \"field\": ").Append(t.Field.ToJson()).Append(",").Append(nl);
            sb.Append(ti).Append("  \"read\": ").Append(Q(Rdv3Uia.ReadModeName(t.ReadMode))).Append(nl);
            sb.Append(ti).Append("}").Append((i + 1 < Targets.Count) ? "," : "").Append(nl);
        }
        sb.Append(ind).Append("  ]").Append(nl);
        sb.Append(ind).Append("}");
        return sb.ToString();
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
    // left out: the ledger and the CSVs are read at start-up, and swapping
    // them under a live session would be a different program. The key pattern
    // only decides which strings are numbers, nothing built at start-up depends
    // on it, so it applies at once.
    public void AdoptRuntimeFrom(Rdv3Config o)
    {
        if (o == null) { return; }
        KeyPattern = o.KeyPattern;
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
        LockRetryMs = o.LockRetryMs;
        LockStaleMs = o.LockStaleMs;
        MarkerPollMs = o.MarkerPollMs;
    }

    // What the file now says for the members that only take effect at the next
    // start. The running session goes on using what it started with -- the
    // resolved paths live in Rdv3App by then -- but this object is also what
    // the settings dialog opens on and what the next save writes. Without this,
    // the dialog would come back up showing the OLD paths and the next save
    // would quietly undo the change the operator had just made.
    public void AdoptSavedFrom(Rdv3Config o)
    {
        if (o == null) { return; }
        DataDir = o.DataDir;
        Ledger = o.Ledger;
        Log = o.Log;
    }

    // deep enough for the dialog to edit without touching the running settings
    // until the operator says OK
    public Rdv3Config Clone()
    {
        Rdv3Config c = new Rdv3Config();
        c.DataDir = DataDir; c.Ledger = Ledger; c.Log = Log;
        c.KeyPattern = KeyPattern;
        c.PollMs = PollMs; c.StableMs = StableMs; c.RebindMs = RebindMs;
        c.PreferFocusedWindow = PreferFocusedWindow;
        c.CandidateRowsShown = CandidateRowsShown;
        c.CheckTimeoutMs = CheckTimeoutMs; c.SearchTimeoutMs = SearchTimeoutMs;
        c.SaveTimeoutMs = SaveTimeoutMs; c.MarkOverdueMs = MarkOverdueMs; c.PumpMs = PumpMs;
        c.LockRetryMs = LockRetryMs; c.LockStaleMs = LockStaleMs; c.MarkerPollMs = MarkerPollMs;
        c.SourcePath = SourcePath;
        c.Data = Data; c.Screen = Screen;
        c.Targets = new List<Rdv3Target>();
        for (int i = 0; i < Targets.Count; i++) { c.Targets.Add(Targets[i].Clone()); }
        return c;
    }

    // one line for the log, so a support question starts from what was read
    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("loaded ").Append(SourcePath);
        sb.Append(" key=").Append(KeyPattern);
        sb.Append(" poll=").Append(PollMs.ToString(CultureInfo.InvariantCulture));
        sb.Append("/").Append(StableMs.ToString(CultureInfo.InvariantCulture));
        sb.Append(" shared=").Append(MarkerPollMs.ToString(CultureInfo.InvariantCulture));
        sb.Append("/").Append(LockRetryMs.ToString(CultureInfo.InvariantCulture));
        sb.Append("/").Append(LockStaleMs.ToString(CultureInfo.InvariantCulture));
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
