// ============================================================================
// Rdv3Values.cs -- evaluating the screen definition against real data.
//
//   Rdv3View     what the screen can see: the search key, the selected record
//                (one ledger line, split), its stored work state, and the app
//                values the status bar names. Owned by the form, written on
//                the UI thread only.
//   Rdv3Fields   "A.a_name" -> ledger column, from the ledger column list of
//                the data definition. The load already proved every name the
//                screen uses is there; a name that still is not resolves to -1
//                and is SHOWN as unresolved, never as a blank.
//   Rdv3Eval     one value binding -> the text and tone to draw; one judgment
//                -> a result; a confirm text with {B.key2} / {state} filled in.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public sealed class Rdv3View
{
    public string SearchKey = "";
    public string[] Record;                 // the selected ledger line, split; null = none
    public string StoredState = "";         // the record's stored work-state string
    public bool Saving;                     // a work-state save is in flight
    public int CandidateCount;
    public int SelectedIndex = -1;
    public int RowNumber;                   // 1-based, while a list row is evaluated

    public string AppState = "";
    public string WatchLabel = "";
    public string WatchDetail = "";
    public string LedgerFile = "";
    public string LedgerRows = "";
    public string LedgerSaved = "";
    public string MergeMs = "";
    public string SearchMs = "";
    public string Pid = "";
    public string LogName = "";
    public string UserName = "";
    public string HostName = "";

    public bool HasRecord { get { return Record != null; } }
}

public sealed class Rdv3Fields
{
    private readonly Dictionary<string, int> byRef = new Dictionary<string, int>(StringComparer.Ordinal);
    private readonly Dictionary<string, int> byName = new Dictionary<string, int>(StringComparer.Ordinal);
    private readonly int count;

    // refs: the ledger columns, "<table>.<column>", in ledger order
    public Rdv3Fields(string[] refs)
    {
        if (refs == null) { return; }
        count = refs.Length;
        for (int i = 0; i < refs.Length; i++)
        {
            string r = refs[i];
            byRef[r] = i;
            int dot = r.IndexOf('.');
            string bare = (dot >= 0) ? r.Substring(dot + 1) : r;
            if (!byName.ContainsKey(bare)) { byName[bare] = i; }
        }
    }

    public static Rdv3Fields Empty { get { return new Rdv3Fields(null); } }

    public int Count { get { return count; } }

    // "A.a_name" (exact), or a bare column name looked up across the tables
    public int IndexOf(string reference)
    {
        if (reference == null) { return -1; }
        string r = reference.Trim();
        int col;
        if (byRef.TryGetValue(r, out col)) { return col; }
        if (byName.TryGetValue(r, out col)) { return col; }
        return -1;
    }

    public bool Has(string reference) { return IndexOf(reference) >= 0; }
}

// the text to draw and how: Normal, Muted (blank / N/A / placeholder) or Error
public sealed class Rdv3Value
{
    public const int Normal = 0;
    public const int Muted = 1;
    public const int Error = 2;

    public string Text = "";
    public int Tone = Normal;

    public Rdv3Value(string text, int tone) { Text = (text == null) ? "" : text; Tone = tone; }
}

// what a judgment decided: Result is null while there is no record to judge
public sealed class Rdv3Verdict
{
    public Rdv3Result Result;
    public string Raw = "";
}

public static class Rdv3Eval
{
    // ---- one binding -> text ----------------------------------------------
    public static Rdv3Value Evaluate(Rdv3Bind b, Rdv3View v, Rdv3Fields f, Rdv3WorkState w)
    {
        if (b == null) { return new Rdv3Value("", Rdv3Value.Muted); }
        if (b.IsState) { return StateValue(b.State, v, w); }
        if (!b.IsField) { return new Rdv3Value("", Rdv3Value.Muted); }
        if (v == null || v.Record == null) { return new Rdv3Value("", Rdv3Value.Muted); }

        StringBuilder sb = new StringBuilder();
        bool unresolved = false;
        for (int i = 0; i < b.Fields.Length; i++)
        {
            int col = (f == null) ? -1 : f.IndexOf(b.Fields[i]);
            if (col < 0) { unresolved = true; continue; }
            string raw = (col < v.Record.Length && v.Record[col] != null) ? v.Record[col] : "";
            if (raw.Length == 0) { continue; }
            if (sb.Length > 0) { sb.Append(b.Joiner); }
            sb.Append(Format(raw, b.Format));
        }
        if (unresolved && sb.Length == 0)
        {
            // the definition names a column the CSVs do not have: say so, in
            // place, rather than show a blank that looks like "no value"
            return new Rdv3Value(Rdv3Text.FieldUnresolved, Rdv3Value.Error);
        }
        if (sb.Length == 0) { return new Rdv3Value(b.Empty, Rdv3Value.Muted); }
        return new Rdv3Value(sb.ToString(), Rdv3Value.Normal);
    }

    private static Rdv3Value StateValue(string name, Rdv3View v, Rdv3WorkState w)
    {
        if (v == null) { return new Rdv3Value("", Rdv3Value.Muted); }
        switch (name)
        {
            case "searchKey": return new Rdv3Value(v.SearchKey, Rdv3Value.Normal);
            case "candidateCount": return new Rdv3Value(v.CandidateCount.ToString(CultureInfo.InvariantCulture), Rdv3Value.Normal);
            case "workState": return WorkStateValue(v, w, false);
            case "workStateShort": return WorkStateValue(v, w, true);
            case "rowNumber": return new Rdv3Value((v.RowNumber > 0) ? v.RowNumber.ToString(CultureInfo.InvariantCulture) : "", Rdv3Value.Normal);
            case "appState": return new Rdv3Value(v.AppState, Rdv3Value.Normal);
            case "watchLabel": return new Rdv3Value(v.WatchLabel, Rdv3Value.Normal);
            case "watchDetail": return new Rdv3Value(v.WatchDetail, Rdv3Value.Normal);
            case "ledgerFile": return new Rdv3Value(v.LedgerFile, Rdv3Value.Normal);
            case "ledgerRows": return new Rdv3Value(v.LedgerRows, Rdv3Value.Normal);
            case "ledgerSaved": return new Rdv3Value(v.LedgerSaved, Rdv3Value.Normal);
            case "mergeMs": return new Rdv3Value(v.MergeMs, Rdv3Value.Normal);
            case "searchMs": return new Rdv3Value(v.SearchMs, Rdv3Value.Normal);
            case "pid": return new Rdv3Value(v.Pid, Rdv3Value.Normal);
            case "logName": return new Rdv3Value(v.LogName, Rdv3Value.Normal);
            case "userName": return new Rdv3Value(v.UserName, Rdv3Value.Normal);
            case "hostName": return new Rdv3Value(v.HostName, Rdv3Value.Normal);
            case "clock": return new Rdv3Value(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), Rdv3Value.Normal);
        }
        return new Rdv3Value("", Rdv3Value.Muted);
    }

    // The record's work state as the screen names it. A stored string the
    // definition does not know is shown AS IS in the error tone: the ledger
    // said something the screen cannot vouch for, and hiding that behind the
    // initial state would be the silent OK this program refuses to give.
    public static Rdv3Value WorkStateValue(Rdv3View v, Rdv3WorkState w, bool shortForm)
    {
        if (v == null || v.Record == null) { return new Rdv3Value("", Rdv3Value.Muted); }
        Rdv3StateDef s = (w == null) ? null : w.ByStored(v.StoredState);
        string text;
        int tone;
        if (s == null)
        {
            text = (v.StoredState.Length == 0) ? Rdv3Text.StateBlank : v.StoredState;
            tone = Rdv3Value.Error;
        }
        else
        {
            text = shortForm ? s.Short : s.Text;
            tone = Rdv3Value.Normal;
        }
        if (v.Saving && !shortForm) { text += Rdv3Text.SavingSuffix; }
        return new Rdv3Value(text, tone);
    }

    // the look of the record's state tag: the state's own, or error
    public static string WorkStateLook(Rdv3View v, Rdv3WorkState w)
    {
        Rdv3StateDef s = (w == null || v == null) ? null : w.ByStored(v.StoredState);
        return (s == null) ? "error" : s.Look;
    }

    // ---- formats -------------------------------------------------------------
    public static string Format(string raw, Rdv3Format fmt)
    {
        if (raw == null) { return ""; }
        if (fmt == null || fmt.Kind == "text") { return raw; }
        if (fmt.Kind == "number")
        {
            decimal d;
            if (decimal.TryParse(raw.Trim(), NumberStyles.Number, CultureInfo.InvariantCulture, out d))
            {
                if (!fmt.Group) { return d.ToString(CultureInfo.InvariantCulture); }
                int scale = Scale(raw.Trim());
                return d.ToString("N" + scale.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            }
            return raw;
        }
        if (fmt.Kind == "date")
        {
            DateTime t;
            if (DateTime.TryParseExact(raw.Trim(), fmt.From, CultureInfo.InvariantCulture, DateTimeStyles.None, out t))
            {
                try { return t.ToString(fmt.To, CultureInfo.InvariantCulture); }
                catch (FormatException) { return raw; }
            }
            return raw;
        }
        return raw;
    }

    private static int Scale(string s)
    {
        int dot = s.IndexOf('.');
        return (dot < 0) ? 0 : (s.Length - dot - 1);
    }

    // ---- judgments -----------------------------------------------------------
    public static Rdv3Verdict Judge(Rdv3Judgment j, Rdv3View v, Rdv3Fields f)
    {
        Rdv3Verdict out1 = new Rdv3Verdict();
        if (v == null || v.Record == null) { return out1; }          // nothing to judge yet
        if (j == null || j.Source == null) { out1.Result = (j == null) ? ErrorResult() : j.ResultOf(Rdv3Judgment.Error); return out1; }
        if (j.Source.IsState)
        {
            out1.Raw = StateValue(j.Source.State, v, null).Text;
        }
        else
        {
            if (j.Source.Fields.Length == 0) { out1.Result = j.ResultOf(Rdv3Judgment.Error); return out1; }
            int col = (f == null) ? -1 : f.IndexOf(j.Source.Fields[0]);
            if (col < 0) { out1.Result = j.ResultOf(Rdv3Judgment.Error); return out1; }
            out1.Raw = (col < v.Record.Length && v.Record[col] != null) ? v.Record[col] : "";
        }
        string raw = out1.Raw;
        for (int i = 0; i < j.Rules.Count; i++)
        {
            Rdv3Rule r = j.Rules[i];
            bool hit = false;
            if (r.Empty && raw.Length == 0) { hit = true; }
            for (int k = 0; !hit && k < r.EqualsAny.Length; k++)
            {
                if (string.Equals(r.EqualsAny[k], raw, StringComparison.Ordinal)) { hit = true; }
            }
            if (!hit && r.PatternRule != null && raw.Length > 0 && r.PatternRule.IsMatch(raw)) { hit = true; }
            if (hit) { out1.Result = j.ResultOf(r.Result); return out1; }
        }
        out1.Result = j.ResultOf(Rdv3Judgment.Undefined);
        return out1;
    }

    private static Rdv3Result ErrorResult()
    {
        Rdv3Result r = new Rdv3Result();
        r.Id = Rdv3Judgment.Error; r.Text = Rdv3Text.JudgeError; r.Look = "error";
        return r;
    }

    // ---- templates: "{B.key2}", "{state}", "{key}" --------------------------
    public static string Template(string text, Rdv3View v, Rdv3Fields f, Rdv3WorkState w)
    {
        if (text == null || text.IndexOf('{') < 0) { return (text == null) ? "" : text; }
        StringBuilder sb = new StringBuilder(text.Length + 16);
        int i = 0;
        while (i < text.Length)
        {
            char c = text[i];
            if (c != '{') { sb.Append(c); i++; continue; }
            int close = text.IndexOf('}', i + 1);
            if (close < 0) { sb.Append(text.Substring(i)); break; }
            string name = text.Substring(i + 1, close - i - 1).Trim();
            string rep;
            if (name == "state") { rep = WorkStateValue(v, w, false).Text; }
            else if (name == "key") { rep = (v == null) ? "" : v.SearchKey; }
            else
            {
                int col = (f == null) ? -1 : f.IndexOf(name);
                if (col >= 0 && v != null && v.Record != null && col < v.Record.Length) { rep = v.Record[col]; }
                else if (col < 0) { rep = "{" + name + "}"; }
                else { rep = ""; }
            }
            sb.Append(rep);
            i = close + 1;
        }
        return sb.ToString();
    }
}
