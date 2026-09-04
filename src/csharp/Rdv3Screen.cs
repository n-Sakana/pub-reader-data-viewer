// ============================================================================
// Rdv3Screen.cs -- the "screen" member of settings.json: what the screen shows.
//
// The screen is not written into the program. It is composed from a small set
// of parts that the file arranges, names and binds to data:
//
//   sections      titleBar | keyPanel | columns | fieldList | textBox |
//                 statusBand | sendBar | statusBar  (in display order)
//   values        { field } | { fields, joiner } | { state }   + format, empty
//   judgments     a source value, ordered rules, named results with a look.
//                 No rule matching is "undefined"; a source that cannot be
//                 read is "error". Neither is ever shown as OK.
//   workState     the states a record can be in (todo / done ...), what
//                 each is stored as, the transitions a button may make, and
//                 the ledger column they live in.
//   candidates    the columns of the candidate list.
//
// There is no expression language and no code in the file: every member is a
// literal that one of the parts above understands, and the program does the
// reading, judging, saving and drawing. The reader is strict (Rdv3Json): a
// member the program does not know, a value of the wrong kind, a name that
// refers to nothing -- a state, a judgment, a ledger column -- is an
// Rdv3LoadError, and the app does not start. There is no built-in screen to
// fall back on; the shipped settings.json is the only definition there is.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

// ---------------------------------------------------------------------------
// one value on the screen: where it comes from and how it is shown
// ---------------------------------------------------------------------------
public sealed class Rdv3Bind
{
    public string[] Fields = new string[0];   // "A.a_name" ... (one or several)
    public string Joiner = " \u30fb ";        // between several fields (katakana middle dot)
    public string State = "";                 // or an app value by name
    public Rdv3Format Format;                 // null = as is
    public string Empty = "N/A";              // a record is shown but the value is blank
    public int Line;                          // where it is written, for the column check

    public bool IsField { get { return Fields.Length > 0; } }
    public bool IsState { get { return State.Length > 0; } }

    // the app values a binding may name. Fixed and small on purpose.
    public static readonly string[] StateNames =
    {
        "searchKey", "candidateCount", "workState", "workStateShort", "rowNumber",
        "appState", "watchLabel", "watchDetail", "ledgerFile", "ledgerRows", "ledgerSaved",
        "pendingCount", "mergeMs", "searchMs", "pid", "logName", "clock", "userName", "hostName"
    };

    public static bool IsStateName(string s)
    {
        for (int i = 0; i < StateNames.Length; i++) { if (StateNames[i] == s) { return true; } }
        return false;
    }

    public static Rdv3Bind Read(Rdv3Json o)
    {
        if (o == null) { throw new Rdv3LoadError("a value is required", 0); }
        o.Only("field", "fields", "joiner", "state", "format", "empty");
        Rdv3Bind b = new Rdv3Bind();
        b.Line = o.Line;
        string one = o.StrOr("field", "");
        string[] many = o.Strs("fields", false);
        string st = o.StrOr("state", "");
        int ways = ((one.Length > 0) ? 1 : 0) + ((many.Length > 0) ? 1 : 0) + ((st.Length > 0) ? 1 : 0);
        if (ways != 1) { throw o.Fail("a value names exactly one of field / fields / state"); }
        if (one.Length > 0) { b.Fields = new string[] { one.Trim() }; }
        else if (many.Length > 0)
        {
            b.Fields = new string[many.Length];
            for (int i = 0; i < many.Length; i++) { b.Fields[i] = many[i].Trim(); }
        }
        b.Joiner = o.StrOr("joiner", b.Joiner);
        b.State = st;
        b.Empty = o.StrOr("empty", b.Empty);
        if (b.IsState && !IsStateName(b.State))
        {
            throw o.Member("state").Fail(b.State + " is not a value the program provides (" + string.Join(", ", StateNames) + ")");
        }
        Rdv3Json f = o.Obj("format", false);
        if (f != null) { b.Format = Rdv3Format.Read(f); }
        return b;
    }

    public string Describe()
    {
        if (IsField) { return (Fields.Length == 1) ? Fields[0] : string.Join("+", Fields); }
        return "state:" + State;
    }
}

// how a raw string is turned into the text on screen
public sealed class Rdv3Format
{
    public string Kind = "text";              // text | number | date
    public bool Group;                        // number: thousands separators
    public string From = "yyyyMMdd";          // date: the CSV's shape
    public string To = "yyyy-MM-dd";          // date: the screen's shape

    public static Rdv3Format Read(Rdv3Json o)
    {
        o.Only("kind", "group", "from", "to");
        Rdv3Format f = new Rdv3Format();
        f.Kind = o.Word("kind", "text", "text", "number", "date");
        f.Group = o.BoolOr("group", true);
        f.From = o.StrOr("from", f.From);
        f.To = o.StrOr("to", f.To);
        return f;
    }
}

// ---------------------------------------------------------------------------
// a judgment: source -> rules -> one of the named results
// ---------------------------------------------------------------------------
public sealed class Rdv3Rule
{
    public string[] EqualsAny = new string[0];
    public string Pattern = "";
    public Regex PatternRule;
    public bool Empty;
    public string Result = "";
}

public sealed class Rdv3Result
{
    public string Id = "";
    public string Text = "";
    public string Look = "undefined";        // ok | ng | undefined | error
    public string Icon = "";                 // "" | check

    public static readonly string[] Looks = { "ok", "ng", "undefined", "error" };
}

public sealed class Rdv3Judgment
{
    public const string Undefined = "undefined";
    public const string Error = "error";

    public string Id = "";
    public Rdv3Bind Source;
    public List<Rdv3Rule> Rules = new List<Rdv3Rule>();
    public Dictionary<string, Rdv3Result> Results = new Dictionary<string, Rdv3Result>(StringComparer.Ordinal);

    public Rdv3Result ResultOf(string id)
    {
        Rdv3Result r;
        if (Results.TryGetValue(id, out r)) { return r; }
        // the two built-in outcomes exist even when the file does not spell them out
        r = new Rdv3Result();
        r.Id = id;
        if (id == Error) { r.Text = Rdv3Text.JudgeError; r.Look = "error"; }
        else { r.Text = Rdv3Text.JudgeUndefined; r.Look = "undefined"; }
        return r;
    }

    public static Rdv3Judgment Read(string id, Rdv3Json o)
    {
        if (o.Kind != Rdv3Json.TObject) { throw o.Fail("must be an object"); }
        o.Only("source", "rules", "results");
        Rdv3Judgment j = new Rdv3Judgment();
        j.Id = id;
        j.Source = Rdv3Bind.Read(o.Obj("source", true));
        List<Rdv3Json> rules = o.Objs("rules", true);
        for (int i = 0; i < rules.Count; i++)
        {
            Rdv3Json ro = rules[i];
            ro.Only("equals", "pattern", "empty", "result");
            Rdv3Rule r = new Rdv3Rule();
            r.EqualsAny = ro.Strs("equals", false);
            r.Pattern = ro.StrOr("pattern", "");
            r.Empty = ro.BoolOr("empty", false);
            r.Result = ro.Need("result");
            if (r.Pattern.Length > 0)
            {
                try { r.PatternRule = new Regex(r.Pattern, RegexOptions.CultureInvariant); }
                catch (Exception ex) { throw ro.Member("pattern").Fail("is not a usable regular expression (" + ex.Message + ")"); }
            }
            if (r.EqualsAny.Length == 0 && r.PatternRule == null && !r.Empty) { throw ro.Fail("has no condition (equals / pattern / empty)"); }
            if (r.Result == Undefined || r.Result == Error) { throw ro.Member("result").Fail("may not name the built-in result " + r.Result); }
            j.Rules.Add(r);
        }
        Rdv3Json res = o.Obj("results", true);
        for (int i = 0; i < res.Order.Count; i++)
        {
            string key = res.Order[i];
            Rdv3Json v = res.Member(key);
            if (v.Kind != Rdv3Json.TObject) { throw v.Fail("must be an object"); }
            v.Only("text", "look", "icon");
            Rdv3Result r = new Rdv3Result();
            r.Id = key;
            r.Text = v.StrOr("text", key);
            r.Look = v.Word("look", (key == Error) ? "error" : (key == Undefined) ? "undefined" : "ok", Rdv3Result.Looks);
            r.Icon = v.Word("icon", "", "", "check");
            j.Results[key] = r;
        }
        for (int i = 0; i < j.Rules.Count; i++)
        {
            if (!j.Results.ContainsKey(j.Rules[i].Result))
            {
                throw rules[i].Member("result").Fail(j.Rules[i].Result + " has no entry under results");
            }
        }
        return j;
    }
}

// ---------------------------------------------------------------------------
// the work state of a record
// ---------------------------------------------------------------------------
public sealed class Rdv3StateDef
{
    public string Id = "";
    public string Text = "";
    public string Short = "";
    public string Look = "neutral";          // neutral | accent | outline | faded | error
    public string Stored = "";               // the string written to the ledger column
}

public sealed class Rdv3Transition
{
    public string From = "";
    public string To = "";
    public string Confirm = "";              // "" = no confirmation
    public string Done = "";                 // the completion notice
    public int Line;
}

public sealed class Rdv3WorkState
{
    public static readonly string[] Looks = { "neutral", "accent", "outline", "faded", "error" };

    public string Column = "";               // the ledger column heading
    public List<Rdv3StateDef> States = new List<Rdv3StateDef>();
    public string Initial = "";
    public List<Rdv3Transition> Transitions = new List<Rdv3Transition>();
    public string ButtonText = "{state}";
    public string ButtonTip = "";
    // automatic: a single hit from the watched application advances the
    // state; manual: only the work-state button can advance it.
    public string Trigger = "automatic";

    public Rdv3StateDef ById(string id)
    {
        for (int i = 0; i < States.Count; i++) { if (States[i].Id == id) { return States[i]; } }
        return null;
    }

    public Rdv3StateDef ByStored(string stored)
    {
        for (int i = 0; i < States.Count; i++)
        {
            if (string.Equals(States[i].Stored, stored, StringComparison.Ordinal)) { return States[i]; }
        }
        return null;
    }

    public Rdv3StateDef InitialState { get { return ById(Initial); } }
    public string InitialStored { get { Rdv3StateDef s = InitialState; return (s == null) ? "" : s.Stored; } }
    public Rdv3StateDef InitialTargetState
    {
        get
        {
            Rdv3Transition transition = FromState(Initial);
            return (transition == null) ? null : ById(transition.To);
        }
    }

    public Rdv3Transition FromState(string id)
    {
        for (int i = 0; i < Transitions.Count; i++) { if (Transitions[i].From == id) { return Transitions[i]; } }
        return null;
    }

    public static Rdv3WorkState Read(Rdv3Json o)
    {
        o.Only("store", "states", "initial", "transitions", "button", "trigger");
        Rdv3WorkState w = new Rdv3WorkState();
        w.Trigger = o.Word("trigger", w.Trigger, "automatic", "manual");
        Rdv3Json store = o.Obj("store", true);
        store.Only("column");
        w.Column = store.Need("column");
        List<Rdv3Json> states = o.Objs("states", true);
        if (states.Count == 0) { throw o.Member("states").Fail("names no state"); }
        for (int i = 0; i < states.Count; i++)
        {
            Rdv3Json so = states[i];
            so.Only("id", "text", "short", "look", "stored");
            Rdv3StateDef s = new Rdv3StateDef();
            s.Id = so.Need("id");
            s.Text = so.StrOr("text", s.Id);
            s.Short = so.StrOr("short", s.Text);
            s.Look = so.Word("look", "neutral", Looks);
            s.Stored = so.StrOr("stored", s.Id);
            if (s.Stored.Length == 0) { throw so.Member("stored").Fail("must not be blank (a blank ledger cell is 'no state')"); }
            if (w.ById(s.Id) != null) { throw so.Member("id").Fail(s.Id + " is used twice"); }
            if (w.ByStored(s.Stored) != null) { throw so.Fail("stored value " + s.Stored + " is used twice"); }
            w.States.Add(s);
        }
        w.Initial = o.Need("initial");
        if (w.ById(w.Initial) == null) { throw o.Member("initial").Fail(w.Initial + " is not a state"); }
        List<Rdv3Json> ts = o.Objs("transitions", false);
        for (int i = 0; i < ts.Count; i++)
        {
            Rdv3Json to = ts[i];
            to.Only("from", "to", "confirm", "done");
            Rdv3Transition t = new Rdv3Transition();
            t.From = to.Need("from");
            t.To = to.Need("to");
            t.Confirm = to.StrOr("confirm", "");
            t.Done = to.StrOr("done", "");
            t.Line = to.Line;
            if (w.ById(t.From) == null) { throw to.Member("from").Fail(t.From + " is not a state"); }
            if (w.ById(t.To) == null) { throw to.Member("to").Fail(t.To + " is not a state"); }
            if (w.FromState(t.From) != null) { throw to.Fail("a second transition from " + t.From); }
            w.Transitions.Add(t);
        }
        Rdv3Json btn = o.Obj("button", false);
        if (btn != null)
        {
            btn.Only("text", "tip");
            w.ButtonText = btn.StrOr("text", w.ButtonText);
            w.ButtonTip = btn.StrOr("tip", w.ButtonTip);
        }
        return w;
    }
}

// ---------------------------------------------------------------------------
// the parts a section is made of
// ---------------------------------------------------------------------------
public sealed class Rdv3TagDef
{
    public string Text = "";
    public string Look = "accent";           // accent | neutral | outline
}

public sealed class Rdv3ButtonDef
{
    public string Action = "";               // a named action implemented by the main form
    public string Text = "";
    public string Icon = "";                 // "" | search | refresh | gear
    public string Tip = "";
    public string Job = "";
    public bool Primary;

    public static readonly string[] Actions =
    { "search", "clear", "workState", "tableExport", "updateRecords", "deleteRecords", "sendChanges", "refreshLedger", "settings" };
}

public sealed class Rdv3RowDef
{
    public string Label = "";
    public Rdv3Bind Value;
}

public sealed class Rdv3SegmentDef
{
    public string Prefix = "";               // a fixed word before the value (the "ledger" label)
    public Rdv3Bind Value;
    public bool Bold;
    public bool Dot;                         // the accent dot before the text
}

public sealed class Rdv3ColumnDef
{
    public string Header = "";
    public double Width;                     // 0 = takes the rest
    public string Align = "left";            // left | right
    public Rdv3Bind Value;
    public bool Bold;
    public bool Muted;
    public string Render = "text";           // text | tag
    public Dictionary<string, string> Looks = new Dictionary<string, string>(StringComparer.Ordinal);
}

public sealed class Rdv3Section
{
    public string Type = "";
    public double[] Margin;                  // null = the card's gap and padding
    public int Line;

    // titleBar
    public string Brand = "";
    public List<Rdv3TagDef> Tags = new List<Rdv3TagDef>();
    public List<Rdv3ButtonDef> Buttons = new List<Rdv3ButtonDef>();
    // keyPanel
    public string Label = "";
    public Rdv3Bind Value;
    public string Placeholder = "";
    public string InputLabel = "";
    public double InputWidth = 220;
    public int MaxLength = 64;
    // columns
    public double[] Weights = new double[0];
    public double Gap = 17;
    public double StackBelow = 760;          // narrower than this, the items stack
    public List<Rdv3Section> Items = new List<Rdv3Section>();
    // fieldList
    public string Title = "";
    public double LabelWidth = 104;
    public double RowHeight = 44;
    public List<Rdv3RowDef> Rows = new List<Rdv3RowDef>();
    // textBox: so many lines of text (the card fits around the box)
    public int Lines = 2;
    // statusBand / sendBar / statusBar
    public double Height;
    // statusBand
    public string Judgment = "";
    public List<Rdv3Bind> Sub = new List<Rdv3Bind>();
    public string Joiner = " \u30fb ";
    // statusBar: segments on the left, buttons on the right
    public List<Rdv3SegmentDef> Segments = new List<Rdv3SegmentDef>();

    public static readonly string[] Types =
    { "titleBar", "keyPanel", "columns", "fieldList", "textBox", "statusBand", "sendBar", "statusBar" };

    public static Rdv3Section Read(Rdv3Json o, bool nested)
    {
        Rdv3Section s = new Rdv3Section();
        s.Line = o.Line;
        s.Type = o.Word("type", "", Types);
        if (s.Type.Length == 0) { throw o.Fail("type is required"); }
        if (nested && s.Type != "fieldList" && s.Type != "textBox")
        {
            throw o.Fail("a columns item must be a fieldList or a textBox");
        }
        s.Margin = o.Box("margin");

        if (s.Type == "titleBar")
        {
            o.Only("type", "margin", "brand", "tags", "buttons");
            s.Brand = o.StrOr("brand", Rdv3Text.AppTitle);
            List<Rdv3Json> tags = o.Objs("tags", false);
            for (int i = 0; i < tags.Count; i++)
            {
                tags[i].Only("text", "look");
                Rdv3TagDef d = new Rdv3TagDef();
                d.Text = tags[i].Need("text");
                d.Look = tags[i].Word("look", "accent", "accent", "neutral", "outline");
                s.Tags.Add(d);
            }
            s.Buttons = ReadButtons(o.Objs("buttons", false));
        }
        else if (s.Type == "keyPanel")
        {
            o.Only("type", "margin", "title", "figure", "input", "buttons");
            s.Title = o.StrOr("title", "");
            Rdv3Json fig = o.Obj("figure", true);
            fig.Only("label", "value");
            s.Label = fig.StrOr("label", "");
            s.Value = Rdv3Bind.Read(fig.Obj("value", true));
            Rdv3Json inp = o.Obj("input", false);
            if (inp != null)
            {
                inp.Only("label", "placeholder", "width", "maxLength");
                s.InputLabel = inp.StrOr("label", "");
                s.Placeholder = inp.StrOr("placeholder", "");
                s.InputWidth = inp.DblOr("width", s.InputWidth, 60, 2000);
                s.MaxLength = inp.IntOr("maxLength", s.MaxLength, 1, 256);
            }
            s.Buttons = ReadButtons(o.Objs("buttons", false));
        }
        else if (s.Type == "columns")
        {
            o.Only("type", "margin", "weights", "gap", "stackBelow", "items");
            s.Gap = o.DblOr("gap", s.Gap, 0, 200);
            s.StackBelow = o.DblOr("stackBelow", s.StackBelow, 0, 10000);
            List<Rdv3Json> items = o.Objs("items", true);
            if (items.Count == 0) { throw o.Member("items").Fail("holds no item"); }
            for (int i = 0; i < items.Count; i++) { s.Items.Add(Read(items[i], true)); }
            Rdv3Json w = o.Arr("weights", false);
            List<double> ws = new List<double>();
            if (w != null)
            {
                for (int i = 0; i < w.Count; i++)
                {
                    Rdv3Json v = w.At(i);
                    if (v.Kind != Rdv3Json.TNumber || v.Num <= 0) { throw v.Fail("must be a positive number"); }
                    ws.Add(v.Num);
                }
                if (ws.Count != s.Items.Count) { throw w.Fail("must have one weight per item"); }
            }
            while (ws.Count < s.Items.Count) { ws.Add(1.0); }
            s.Weights = ws.ToArray();
        }
        else if (s.Type == "fieldList")
        {
            o.Only("type", "margin", "title", "labelWidth", "rowHeight", "rows");
            s.Title = o.StrOr("title", "");
            s.LabelWidth = o.DblOr("labelWidth", s.LabelWidth, 20, 1000);
            s.RowHeight = o.DblOr("rowHeight", s.RowHeight, 20, 200);
            List<Rdv3Json> rows = o.Objs("rows", true);
            if (rows.Count == 0) { throw o.Member("rows").Fail("holds no row"); }
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].Only("label", "value");
                Rdv3RowDef d = new Rdv3RowDef();
                d.Label = rows[i].StrOr("label", "");
                d.Value = Rdv3Bind.Read(rows[i].Obj("value", true));
                s.Rows.Add(d);
            }
        }
        else if (s.Type == "textBox")
        {
            o.Only("type", "margin", "title", "lines", "value");
            s.Title = o.StrOr("title", "");
            s.Lines = o.Int("lines", 1, 60);
            s.Value = Rdv3Bind.Read(o.Obj("value", true));
        }
        else if (s.Type == "statusBand")
        {
            o.Only("type", "margin", "label", "height", "judgment", "joiner", "sub");
            s.Label = o.StrOr("label", "");
            s.Height = o.DblOr("height", 84, 30, 500);
            s.Judgment = o.Need("judgment");
            s.Joiner = o.StrOr("joiner", s.Joiner);
            List<Rdv3Json> sub = o.Objs("sub", false);
            for (int i = 0; i < sub.Count; i++) { s.Sub.Add(Rdv3Bind.Read(sub[i])); }
        }
        else if (s.Type == "sendBar")
        {
            o.Only("type", "margin", "height", "value", "buttons");
            s.Height = o.DblOr("height", 31, 24, 200);
            s.Value = Rdv3Bind.Read(o.Obj("value", true));
            s.Buttons = ReadButtons(o.Objs("buttons", true));
            if (s.Buttons.Count != 1 || s.Buttons[0].Action != "sendChanges")
            {
                throw o.Member("buttons").Fail("must hold one sendChanges button");
            }
        }
        else if (s.Type == "statusBar")
        {
            o.Only("type", "margin", "height", "segments", "buttons");
            s.Height = o.DblOr("height", 38, 20, 200);
            List<Rdv3Json> segs = o.Objs("segments", false);
            for (int i = 0; i < segs.Count; i++)
            {
                segs[i].Only("prefix", "value", "bold", "dot");
                Rdv3SegmentDef d = new Rdv3SegmentDef();
                d.Prefix = segs[i].StrOr("prefix", "");
                d.Bold = segs[i].BoolOr("bold", false);
                d.Dot = segs[i].BoolOr("dot", false);
                d.Value = Rdv3Bind.Read(segs[i].Obj("value", true));
                s.Segments.Add(d);
            }
            s.Buttons = ReadButtons(o.Objs("buttons", false));
        }
        return s;
    }

    private static List<Rdv3ButtonDef> ReadButtons(List<Rdv3Json> a)
    {
        List<Rdv3ButtonDef> list = new List<Rdv3ButtonDef>();
        for (int i = 0; i < a.Count; i++)
        {
            Rdv3Json b = a[i];
            b.Only("action", "text", "icon", "tip", "primary", "job");
            Rdv3ButtonDef d = new Rdv3ButtonDef();
            d.Action = b.Word("action", "", Rdv3ButtonDef.Actions);
            if (d.Action.Length == 0) { throw b.Fail("action is required"); }
            d.Text = b.StrOr("text", "");
            d.Icon = b.Word("icon", "", "", "search", "refresh", "gear");
            d.Tip = b.StrOr("tip", "");
            d.Job = b.StrOr("job", "");
            d.Primary = b.BoolOr("primary", false);
            if (d.Text.Length == 0 && d.Action != "workState") { throw b.Fail("text is required"); }
            list.Add(d);
        }
        return list;
    }
}

public sealed class Rdv3CandidatesDef
{
    public string Title = "";
    public string Hint = "";
    public double Width = 980;
    public double MaxHeight = 340;
    public double RowHeight = 46;
    public double HeaderHeight = 38;
    public List<Rdv3ColumnDef> Columns = new List<Rdv3ColumnDef>();

    public static Rdv3CandidatesDef Read(Rdv3Json o)
    {
        o.Only("title", "hint", "width", "maxHeight", "rowHeight", "headerHeight", "columns");
        Rdv3CandidatesDef c = new Rdv3CandidatesDef();
        c.Title = o.StrOr("title", Rdv3Text.PanelCand);
        c.Hint = o.StrOr("hint", "");
        c.Width = o.DblOr("width", c.Width, 300, 4000);
        c.MaxHeight = o.DblOr("maxHeight", c.MaxHeight, 60, 4000);
        c.RowHeight = o.DblOr("rowHeight", c.RowHeight, 20, 200);
        c.HeaderHeight = o.DblOr("headerHeight", c.HeaderHeight, 16, 200);
        List<Rdv3Json> cols = o.Objs("columns", true);
        if (cols.Count == 0) { throw o.Member("columns").Fail("holds no column"); }
        for (int i = 0; i < cols.Count; i++)
        {
            Rdv3Json k = cols[i];
            k.Only("header", "width", "align", "value", "bold", "muted", "render", "looks");
            Rdv3ColumnDef d = new Rdv3ColumnDef();
            d.Header = k.StrOr("header", "");
            d.Width = k.DblOr("width", 0, 0, 2000);
            d.Align = k.Word("align", "left", "left", "right");
            d.Value = Rdv3Bind.Read(k.Obj("value", true));
            d.Bold = k.BoolOr("bold", false);
            d.Muted = k.BoolOr("muted", false);
            d.Render = k.Word("render", "text", "text", "tag");
            Rdv3Json looks = k.Obj("looks", false);
            if (looks != null)
            {
                for (int q = 0; q < looks.Order.Count; q++)
                {
                    Rdv3Json v = looks.Member(looks.Order[q]);
                    if (v.Kind != Rdv3Json.TString) { throw v.Fail("must be a look name"); }
                    bool known = false;
                    for (int x = 0; x < Rdv3WorkState.Looks.Length; x++) { if (Rdv3WorkState.Looks[x] == v.Str) { known = true; } }
                    if (!known) { throw v.Fail(v.Str + " is not a look (" + string.Join(" / ", Rdv3WorkState.Looks) + ")"); }
                    d.Looks[looks.Order[q]] = v.Str;
                }
            }
            c.Columns.Add(d);
        }
        return c;
    }
}

// ---------------------------------------------------------------------------
public sealed class Rdv3Screen
{
    public double CardWidth = 1240;
    // the window's client size at start-up, CSS px (the reference's 1240 is
    // the design width; the operator works in a smaller window)
    public double StartWidth = 840;
    public double StartHeight = 830;
    public double Gap = 17;
    public double[] Padding = { 14.3, 14.3, 14.2, 14.3 };   // top, right, bottom, left
    public string FontFamily = "Yu Gothic UI";
    // point size of the screen font; every height on the screen is measured
    // from it, so raising it here raises the rows and the boxes with it
    public double FontSize = 9;
    // The three emphasized sizes are separate from the body font so a screen
    // definition can keep labels quiet while making the current value and the
    // central judgment easy to find.
    public double KeyValueFontSize = 15;
    public double JudgmentFontSize = 15;
    public double UnsearchedFontSize = 12;
    public List<Rdv3Section> Sections = new List<Rdv3Section>();
    public Rdv3CandidatesDef Candidates;
    public Dictionary<string, Rdv3Judgment> Judgments = new Dictionary<string, Rdv3Judgment>(StringComparer.Ordinal);
    public Rdv3WorkState Work;
    public string[] ExportDefaultFields = new string[0];
    private int exportLine;

    public Rdv3Judgment JudgmentOf(string id)
    {
        Rdv3Judgment j;
        return (id != null && Judgments.TryGetValue(id, out j)) ? j : null;
    }

    // every binding the screen evaluates, so the load can check what it refers to
    public List<Rdv3Bind> AllBindings()
    {
        List<Rdv3Bind> all = new List<Rdv3Bind>();
        for (int i = 0; i < Sections.Count; i++) { Collect(Sections[i], all); }
        if (Candidates != null) { for (int i = 0; i < Candidates.Columns.Count; i++) { if (Candidates.Columns[i].Value != null) { all.Add(Candidates.Columns[i].Value); } } }
        foreach (KeyValuePair<string, Rdv3Judgment> kv in Judgments) { if (kv.Value.Source != null) { all.Add(kv.Value.Source); } }
        return all;
    }

    private static void Collect(Rdv3Section s, List<Rdv3Bind> all)
    {
        if (s.Value != null) { all.Add(s.Value); }
        for (int i = 0; i < s.Rows.Count; i++) { if (s.Rows[i].Value != null) { all.Add(s.Rows[i].Value); } }
        for (int i = 0; i < s.Sub.Count; i++) { all.Add(s.Sub[i]); }
        for (int i = 0; i < s.Segments.Count; i++) { if (s.Segments[i].Value != null) { all.Add(s.Segments[i].Value); } }
        for (int i = 0; i < s.Items.Count; i++) { Collect(s.Items[i], all); }
    }

    // ---- reading the "screen" member -------------------------------------------
    public static Rdv3Screen Read(Rdv3Json root)
    {
        root.Only("card", "judgments", "workState", "export", "sections", "candidates");
        Rdv3Screen s = new Rdv3Screen();
        Rdv3Json card = root.Obj("card", false);
        if (card != null)
        {
            card.Only("width", "startSize", "gap", "padding", "font", "fontSize",
                      "keyValueFontSize", "judgmentFontSize", "unsearchedFontSize");
            s.CardWidth = card.DblOr("width", s.CardWidth, 600, 4000);
            double[] start = card.Box("startSize");
            if (start != null)
            {
                if (card.Member("startSize").Count != 2) { throw card.Member("startSize").Fail("is [width, height]"); }
                s.StartWidth = Math.Max(480, start[0]);
                s.StartHeight = Math.Max(300, start[1]);
            }
            s.Gap = card.DblOr("gap", s.Gap, 0, 200);
            double[] pad = card.Box("padding");
            if (pad != null) { s.Padding = pad; }
            s.FontFamily = card.StrOr("font", s.FontFamily);
            if (s.FontFamily.Trim().Length == 0) { throw card.Member("font").Fail("must name a font"); }
            s.FontSize = card.DblOr("fontSize", s.FontSize, 6, 24);
            s.KeyValueFontSize = card.DblOr("keyValueFontSize", s.KeyValueFontSize, 6, 36);
            s.JudgmentFontSize = card.DblOr("judgmentFontSize", s.JudgmentFontSize, 6, 36);
            s.UnsearchedFontSize = card.DblOr("unsearchedFontSize", s.UnsearchedFontSize, 6, 36);
        }
        Rdv3Json js = root.Obj("judgments", false);
        if (js != null)
        {
            for (int i = 0; i < js.Order.Count; i++)
            {
                s.Judgments[js.Order[i]] = Rdv3Judgment.Read(js.Order[i], js.Member(js.Order[i]));
            }
        }
        s.Work = Rdv3WorkState.Read(root.Obj("workState", true));
        Rdv3Json export = root.Obj("export", true);
        export.Only("defaultFields");
        s.ExportDefaultFields = export.Strs("defaultFields", true);
        s.exportLine = export.Line;
        if (s.ExportDefaultFields.Length == 0) { throw export.Member("defaultFields").Fail("names no field"); }
        s.Candidates = Rdv3CandidatesDef.Read(root.Obj("candidates", true));

        List<Rdv3Json> secs = root.Objs("sections", true);
        if (secs.Count == 0) { throw root.Member("sections").Fail("holds no section"); }
        for (int i = 0; i < secs.Count; i++) { s.Sections.Add(Rdv3Section.Read(secs[i], false)); }
        for (int i = 0; i < s.Sections.Count; i++)
        {
            Rdv3Section sec = s.Sections[i];
            if (sec.Type == "statusBand" && s.JudgmentOf(sec.Judgment) == null)
            {
                throw new Rdv3LoadError(secs[i].Path + ".judgment: " + sec.Judgment + " is not defined under judgments", sec.Line);
            }
        }
        return s;
    }

    // ---- the screen against the data definition ---------------------------------
    // Every "<table>.<column>" the screen names -- in a value, a judgment source,
    // a confirm text -- must be one of data.ledger.columns.
    public void Check(Rdv3Data data)
    {
        List<Rdv3Bind> all = AllBindings();
        for (int i = 0; i < all.Count; i++)
        {
            Rdv3Bind b = all[i];
            if (!b.IsField) { continue; }
            for (int k = 0; k < b.Fields.Length; k++)
            {
                if (data.IndexOf(b.Fields[k]) < 0)
                {
                    throw new Rdv3LoadError("screen: " + b.Fields[k] + " is not one of data.ledger.columns", b.Line);
                }
            }
        }
        for (int i = 0; i < Work.Transitions.Count; i++)
        {
            Rdv3Transition t = Work.Transitions[i];
            CheckTemplate(t.Confirm, data, t.Line);
            CheckTemplate(t.Done, data, t.Line);
        }
        HashSet<string> exportSeen = new HashSet<string>(StringComparer.Ordinal);
        for (int i = 0; i < ExportDefaultFields.Length; i++)
        {
            string reference = ExportDefaultFields[i];
            bool available = reference == "$work";
            for (int k = 0; !available && k < data.LabelOrder.Count; k++)
            {
                if (data.LabelOrder[k] == reference && data.IndexOf(reference) >= 0) { available = true; }
            }
            if (!available)
            {
                throw new Rdv3LoadError("screen.export.defaultFields: " + reference + " is not an export field", exportLine);
            }
            if (!exportSeen.Add(reference))
            {
                throw new Rdv3LoadError("screen.export.defaultFields: " + reference + " is listed twice", exportLine);
            }
        }
        for (int i = 0; i < Sections.Count; i++) { CheckButtons(Sections[i], data); }
    }

    private static void CheckButtons(Rdv3Section section, Rdv3Data data)
    {
        for (int i = 0; i < section.Buttons.Count; i++)
        {
            Rdv3ButtonDef button = section.Buttons[i];
            bool process = button.Action == "updateRecords" || button.Action == "deleteRecords";
            if (!process)
            {
                if (button.Job.Length > 0) { throw new Rdv3LoadError("screen: only updateRecords/deleteRecords buttons may name a job", section.Line); }
                continue;
            }
            Rdv3ProcessJobDef job = data.JobOf(button.Job);
            string kind = (button.Action == "updateRecords") ? "update" : "delete";
            if (job == null) { throw new Rdv3LoadError("screen: button job " + button.Job + " is not one of data.jobs", section.Line); }
            if (job.Kind != kind) { throw new Rdv3LoadError("screen: button job " + button.Job + " is " + job.Kind + ", not " + kind, section.Line); }
            if (job.FinalKind != "ledger") { throw new Rdv3LoadError("screen: button job " + button.Job + " does not produce ledger", section.Line); }
        }
        for (int i = 0; i < section.Items.Count; i++) { CheckButtons(section.Items[i], data); }
    }

    private static void CheckTemplate(string text, Rdv3Data data, int line)
    {
        int i = 0;
        while (text != null && i < text.Length)
        {
            int open = text.IndexOf('{', i);
            if (open < 0) { return; }
            int close = text.IndexOf('}', open + 1);
            if (close < 0) { throw new Rdv3LoadError("screen.workState: a { without } in " + text, line); }
            string name = text.Substring(open + 1, close - open - 1).Trim();
            if (name != "state" && name != "key" && data.IndexOf(name) < 0)
            {
                throw new Rdv3LoadError("screen.workState: {" + name + "} is neither {state}, {key} nor a ledger column", line);
            }
            i = close + 1;
        }
    }

    // one line for the log
    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("sections=").Append(Sections.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" judgments=").Append(Judgments.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" states=").Append((Work == null) ? "0" : Work.States.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" columns=").Append((Candidates == null) ? "0" : Candidates.Columns.Count.ToString(CultureInfo.InvariantCulture));
        sb.Append(" font=").Append(FontFamily);
        sb.Append(" fontSize=").Append(FontSize.ToString(CultureInfo.InvariantCulture));
        sb.Append(" emphasis=").Append(KeyValueFontSize.ToString(CultureInfo.InvariantCulture));
        sb.Append('/').Append(JudgmentFontSize.ToString(CultureInfo.InvariantCulture));
        sb.Append('/').Append(UnsearchedFontSize.ToString(CultureInfo.InvariantCulture));
        sb.Append(" trigger=").Append((Work == null) ? "-" : Work.Trigger);
        return sb.ToString();
    }
}
