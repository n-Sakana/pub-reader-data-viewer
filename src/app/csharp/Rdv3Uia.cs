// ============================================================================
// Rdv3Uia.cs -- UI Automation, and nothing but UI Automation.
//
// Every part of watching a foreign application happens through UIA: finding its
// window, walking into the field the operator named, deciding which window has
// the focus, and reading the text. There is no user32 call in this path -- not
// GetForegroundWindow, not FindWindowEx, not a message hook, not a mouse or
// keyboard poll. (The only P/Invoke left in the build is DPI awareness for our
// OWN window, in Rdv3Ui.cs.)
//
// A real line-of-business window is not Notepad: the field is usually several
// levels down, inside a tab, a pane, a grid. So a target is a PATH --
//
//     window   ->  step  ->  step  ->  field
//
// where each step is matched on the properties UIA actually exposes
// (AutomationId first, then ClassName / Name / ControlType), and each step says
// whether to look at direct children or all descendants. The value can be read
// through ValuePattern, TextPattern or the Name property, because plenty of
// business controls expose only one of the three.
//
// Dump() writes what UIA can see, so the operator can copy real AutomationIds
// into the settings file instead of guessing.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Windows.Automation;
using System.Windows.Automation.Text;

public static class Rdv3Uia
{
    public const int ReadValue = 0;        // ValuePattern
    public const int ReadText = 1;         // TextPattern
    public const int ReadName = 2;         // the Name property

    public static int ReadModeByName(string s)
    {
        if (s == null) { return ReadValue; }
        string t = s.Trim().ToLowerInvariant();
        if (t == "text") { return ReadText; }
        if (t == "name") { return ReadName; }
        return ReadValue;
    }

    public static string ReadModeName(int mode)
    {
        if (mode == ReadText) { return "text"; }
        if (mode == ReadName) { return "name"; }
        return "value";
    }

    // ---- control types by the names a person would write --------------------
    private static Dictionary<string, ControlType> types;

    private static Dictionary<string, ControlType> Types()
    {
        if (types != null) { return types; }
        Dictionary<string, ControlType> d = new Dictionary<string, ControlType>(StringComparer.OrdinalIgnoreCase);
        d["Button"] = ControlType.Button;
        d["Calendar"] = ControlType.Calendar;
        d["CheckBox"] = ControlType.CheckBox;
        d["ComboBox"] = ControlType.ComboBox;
        d["Custom"] = ControlType.Custom;
        d["DataGrid"] = ControlType.DataGrid;
        d["DataItem"] = ControlType.DataItem;
        d["Document"] = ControlType.Document;
        d["Edit"] = ControlType.Edit;
        d["Group"] = ControlType.Group;
        d["Header"] = ControlType.Header;
        d["HeaderItem"] = ControlType.HeaderItem;
        d["Hyperlink"] = ControlType.Hyperlink;
        d["Image"] = ControlType.Image;
        d["List"] = ControlType.List;
        d["ListItem"] = ControlType.ListItem;
        d["Menu"] = ControlType.Menu;
        d["MenuBar"] = ControlType.MenuBar;
        d["MenuItem"] = ControlType.MenuItem;
        d["Pane"] = ControlType.Pane;
        d["ProgressBar"] = ControlType.ProgressBar;
        d["RadioButton"] = ControlType.RadioButton;
        d["ScrollBar"] = ControlType.ScrollBar;
        d["Separator"] = ControlType.Separator;
        d["Slider"] = ControlType.Slider;
        d["Spinner"] = ControlType.Spinner;
        d["SplitButton"] = ControlType.SplitButton;
        d["StatusBar"] = ControlType.StatusBar;
        d["Tab"] = ControlType.Tab;
        d["TabItem"] = ControlType.TabItem;
        d["Table"] = ControlType.Table;
        d["Text"] = ControlType.Text;
        d["Thumb"] = ControlType.Thumb;
        d["TitleBar"] = ControlType.TitleBar;
        d["ToolBar"] = ControlType.ToolBar;
        d["ToolTip"] = ControlType.ToolTip;
        d["Tree"] = ControlType.Tree;
        d["TreeItem"] = ControlType.TreeItem;
        d["Window"] = ControlType.Window;
        types = d;
        return types;
    }

    public static ControlType ControlTypeByName(string name)
    {
        ControlType t;
        return Types().TryGetValue((name == null) ? "" : name.Trim(), out t) ? t : null;
    }

    public static string ControlTypeName(ControlType t)
    {
        if (t == null) { return ""; }
        string s = t.ProgrammaticName;                 // "ControlType.Edit"
        int dot = s.LastIndexOf('.');
        return (dot >= 0) ? s.Substring(dot + 1) : s;
    }

    // ---- matching ----------------------------------------------------------
    // The cheap, indexable properties go into the UIA condition so the provider
    // does the work; wildcards and the process name are checked afterwards on
    // the (small) result.
    private static Condition ToCondition(Rdv3Match m)
    {
        List<Condition> parts = new List<Condition>();
        if (m.AutomationId.Length > 0)
        {
            parts.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, m.AutomationId));
        }
        if (m.ClassName.Length > 0)
        {
            parts.Add(new PropertyCondition(AutomationElement.ClassNameProperty, m.ClassName));
        }
        if (m.Name.Length > 0)
        {
            parts.Add(new PropertyCondition(AutomationElement.NameProperty, m.Name));
        }
        if (m.RequireValuePattern)
        {
            parts.Add(new PropertyCondition(AutomationElement.IsValuePatternAvailableProperty, true));
        }
        if (m.ControlTypes.Length > 0)
        {
            List<Condition> ors = new List<Condition>();
            for (int i = 0; i < m.ControlTypes.Length; i++)
            {
                ControlType ct = ControlTypeByName(m.ControlTypes[i]);
                if (ct != null) { ors.Add(new PropertyCondition(AutomationElement.ControlTypeProperty, ct)); }
            }
            if (ors.Count == 1) { parts.Add(ors[0]); }
            else if (ors.Count > 1) { parts.Add(new OrCondition(ors.ToArray())); }
        }
        if (parts.Count == 0) { return Condition.TrueCondition; }
        if (parts.Count == 1) { return parts[0]; }
        return new AndCondition(parts.ToArray());
    }

    public static bool Like(string text, string pattern)
    {
        if (pattern == null || pattern.Length == 0) { return true; }
        if (text == null) { text = ""; }
        return LikeAt(text, 0, pattern, 0);
    }

    private static bool LikeAt(string s, int si, string p, int pi)
    {
        while (pi < p.Length)
        {
            char c = p[pi];
            if (c == '*')
            {
                while (pi + 1 < p.Length && p[pi + 1] == '*') { pi++; }
                if (pi == p.Length - 1) { return true; }
                for (int k = si; k <= s.Length; k++)
                {
                    if (LikeAt(s, k, p, pi + 1)) { return true; }
                }
                return false;
            }
            if (si >= s.Length) { return false; }
            if (c != '?' && char.ToLowerInvariant(c) != char.ToLowerInvariant(s[si])) { return false; }
            si++; pi++;
        }
        return si == s.Length;
    }

    private static bool PostMatch(AutomationElement e, Rdv3Match m)
    {
        try
        {
            if (m.NameLike.Length > 0 && !Like(e.Current.Name, m.NameLike)) { return false; }
            if (m.ProcessName.Length > 0)
            {
                Process pr = Process.GetProcessById(e.Current.ProcessId);
                if (!string.Equals(pr.ProcessName, m.ProcessName, StringComparison.OrdinalIgnoreCase)) { return false; }
            }
            return true;
        }
        catch (Exception) { return false; }
    }

    public static List<AutomationElement> FindAll(AutomationElement from, Rdv3Match m, bool descendants)
    {
        List<AutomationElement> outp = new List<AutomationElement>();
        if (from == null) { return outp; }
        try
        {
            AutomationElementCollection found = from.FindAll(
                descendants ? TreeScope.Descendants : TreeScope.Children, ToCondition(m));
            for (int i = 0; i < found.Count; i++)
            {
                if (PostMatch(found[i], m)) { outp.Add(found[i]); }
            }
        }
        catch (ElementNotAvailableException) { }
        catch (Exception) { }
        return outp;
    }

    public static AutomationElement FindOne(AutomationElement from, Rdv3Match m, bool descendants)
    {
        List<AutomationElement> all = FindAll(from, m, descendants);
        if (all.Count == 0) { return null; }
        return (m.Index < all.Count) ? all[m.Index] : null;
    }

    // ---- the focused top level window, without user32 ----------------------
    public static AutomationElement FocusedWindow()
    {
        try
        {
            AutomationElement e = AutomationElement.FocusedElement;
            AutomationElement root = AutomationElement.RootElement;
            TreeWalker walker = TreeWalker.ControlViewWalker;
            for (int guard = 0; guard < 64 && e != null; guard++)
            {
                AutomationElement parent = walker.GetParent(e);
                if (parent == null) { return e; }
                if (Automation.Compare(parent, root)) { return e; }
                e = parent;
            }
        }
        catch (Exception) { }
        return null;
    }

    public static bool Same(AutomationElement a, AutomationElement b)
    {
        if (a == null || b == null) { return false; }
        try { return Automation.Compare(a, b); }
        catch (Exception) { return false; }
    }

    // ---- reading the value -------------------------------------------------
    public static string ReadValueOf(AutomationElement e, int mode)
    {
        if (e == null) { return null; }
        try
        {
            if (mode == ReadName) { return e.Current.Name; }
            if (mode == ReadText)
            {
                object p;
                if (e.TryGetCurrentPattern(TextPattern.Pattern, out p))
                {
                    return ((TextPattern)p).DocumentRange.GetText(-1);
                }
                return null;
            }
            object v;
            if (e.TryGetCurrentPattern(ValuePattern.Pattern, out v))
            {
                return ((ValuePattern)v).Current.Value;
            }
            return null;
        }
        catch (ElementNotAvailableException) { throw; }
        catch (Exception) { return null; }
    }

    public static bool CanRead(AutomationElement e, int mode)
    {
        return ReadValueOf(e, mode) != null;
    }

    // ---- the tree, so the operator can write the settings file --------------
    // Without this, naming a field inside a real application is guesswork.
    public static string Dump(int maxDepth, string windowLike)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append("UI Automation tree  ").Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss",
            CultureInfo.InvariantCulture)).Append("\r\n");
        sb.Append("columns: depth | ControlType | AutomationId | ClassName | Name | patterns | process\r\n");
        sb.Append("\r\n");
        try
        {
            AutomationElement root = AutomationElement.RootElement;
            AutomationElementCollection wins = root.FindAll(TreeScope.Children, Condition.TrueCondition);
            for (int i = 0; i < wins.Count; i++)
            {
                AutomationElement w = wins[i];
                string name;
                try { name = w.Current.Name; }
                catch (Exception) { continue; }
                if (windowLike != null && windowLike.Length > 0 && !Like(name, windowLike)) { continue; }
                Walk(sb, w, 0, maxDepth);
                sb.Append("\r\n");
            }
        }
        catch (Exception ex) { sb.Append("(walk stopped: ").Append(ex.Message).Append(")\r\n"); }
        return sb.ToString();
    }

    private static void Walk(StringBuilder sb, AutomationElement e, int depth, int maxDepth)
    {
        if (e == null || depth > maxDepth) { return; }
        try
        {
            AutomationElement.AutomationElementInformation c = e.Current;
            sb.Append(new string(' ', depth * 2));
            sb.Append(depth.ToString(CultureInfo.InvariantCulture)).Append(" | ");
            sb.Append(ControlTypeName(c.ControlType)).Append(" | ");
            sb.Append(c.AutomationId).Append(" | ");
            sb.Append(c.ClassName).Append(" | ");
            sb.Append(Trim(c.Name, 60)).Append(" | ");
            sb.Append(Patterns(e)).Append(" | ");
            sb.Append(ProcessNameOf(c.ProcessId));
            sb.Append("\r\n");
        }
        catch (Exception) { return; }
        if (depth == maxDepth) { return; }
        try
        {
            AutomationElementCollection kids = e.FindAll(TreeScope.Children, Condition.TrueCondition);
            for (int i = 0; i < kids.Count && i < 400; i++) { Walk(sb, kids[i], depth + 1, maxDepth); }
        }
        catch (Exception) { }
    }

    private static string Patterns(AutomationElement e)
    {
        StringBuilder sb = new StringBuilder();
        object p;
        try { if (e.TryGetCurrentPattern(ValuePattern.Pattern, out p)) { sb.Append("value"); } }
        catch (Exception) { }
        try
        {
            if (e.TryGetCurrentPattern(TextPattern.Pattern, out p))
            {
                if (sb.Length > 0) { sb.Append("+"); }
                sb.Append("text");
            }
        }
        catch (Exception) { }
        return sb.ToString();
    }

    private static string ProcessNameOf(int pid)
    {
        try { return Process.GetProcessById(pid).ProcessName; }
        catch (Exception) { return ""; }
    }

    private static string Trim(string s, int n)
    {
        if (s == null) { return ""; }
        s = s.Replace("\r", " ").Replace("\n", " ");
        return (s.Length <= n) ? s : (s.Substring(0, n) + "...");
    }
}
