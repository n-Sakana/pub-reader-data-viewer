// ============================================================================
// Rdv3Watch.cs -- reading detection for the practical build. A deliberate copy
// of src\v2\csharp\Rdv2Watch.cs with the same three rules, renamed so the
// frozen comparison sources stay untouched (the same convention Rdv2Watch.cs
// itself followed toward src\csharp\RdvWatch.cs).
//
//   format+length   exactly 8 digits
//   quiet time      the value must survive StableMs of polling
//   no repeat       the same value does not fire twice in a row, unless the
//                   field went empty in between (a deliberate re-read)
//
// The watcher never starts, closes or kills Notepad. It attaches to a window
// that is already there and only ever reads from it.
//
// In the practical build a confirmed number triggers a SEARCH of the current
// ledger -- not a re-merge. The merge belongs to the startup update check.
// ============================================================================

using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;

public sealed class Rdv3Watch
{
    public int PollMs = 40;
    public int StableMs = 120;

    private AutomationElement win;
    private AutomationElement doc;
    private ValuePattern val;
    private int hwnd;
    private string title = "";

    private string pending = "";
    private long pendingSince;
    private int pendingPolls;
    private string lastFired = "";
    private bool sawEmpty = true;
    private volatile bool stop;
    private volatile bool rebindWanted;
    private Thread thread;

    // key, detection latency in ms, poll count, and the confirm timestamp
    public Action<string, double, int, long> OnConfirmed;
    public Action<string, string> OnState;      // state, detail
    public Action<string> OnRaw;                // current field text, every poll

    public int Hwnd { get { return hwnd; } }
    public string Title { get { return title; } }
    public bool Bound { get { return doc != null; } }

    // What the reader last delivered: the final non-empty line of the document.
    public static string Candidate(string s)
    {
        if (s == null) { return ""; }
        int end = s.Length;
        while (end > 0 && (s[end - 1] == '\n' || s[end - 1] == '\r' || s[end - 1] == ' ' || s[end - 1] == '\t'))
        {
            end--;
        }
        int start = end;
        while (start > 0 && s[start - 1] != '\n' && s[start - 1] != '\r') { start--; }
        return s.Substring(start, end - start).Trim();
    }

    // ---- binding -----------------------------------------------------------
    // Never GetActiveObject-style guesswork: walk the desktop's children, take
    // Notepad windows only, and prefer the one in the foreground so the user
    // can choose which window feeds the app just by clicking it.
    public bool Bind()
    {
        try
        {
            AutomationElement root = AutomationElement.RootElement;
            AutomationElementCollection wins = root.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Notepad"));
            if (wins.Count == 0) { doc = null; val = null; win = null; return false; }

            IntPtr fg = Rdv3NativeFg.GetForegroundWindow();
            AutomationElement pick = null;
            for (int i = 0; i < wins.Count; i++)
            {
                if (new IntPtr(wins[i].Current.NativeWindowHandle) == fg) { pick = wins[i]; break; }
            }
            if (pick == null) { pick = wins[wins.Count - 1]; }

            AutomationElement d = FindTextHost(pick);
            if (d == null) { return false; }
            ValuePattern vp = (ValuePattern)d.GetCurrentPattern(ValuePattern.Pattern);

            win = pick;
            doc = d;
            val = vp;
            hwnd = pick.Current.NativeWindowHandle;
            title = pick.Current.Name;
            return true;
        }
        catch (Exception)
        {
            doc = null; val = null; win = null;
            return false;
        }
    }

    // Windows 11 Notepad is a RichEditD2DPT Document; the classic one is an
    // Edit. Both expose ValuePattern, so ask for the pattern, not the class.
    private static AutomationElement FindTextHost(AutomationElement w)
    {
        AutomationElement d = w.FindFirst(TreeScope.Descendants,
            new AndCondition(
                new PropertyCondition(AutomationElement.IsValuePatternAvailableProperty, true),
                new OrCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document),
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit))));
        return d;
    }

    public string ReadField()
    {
        if (val == null) { return null; }
        try { return val.Current.Value; }
        catch (ElementNotAvailableException) { doc = null; val = null; return null; }
        catch (Exception) { return null; }
    }

    // ---- the loop ----------------------------------------------------------
    public void Start()
    {
        stop = false;
        thread = new Thread(Loop);
        thread.IsBackground = true;
        thread.SetApartmentState(ApartmentState.MTA);
        thread.Name = "rdv-watch";
        thread.Start();
    }

    public void Stop()
    {
        stop = true;
        Thread t = thread;
        if (t != null) { t.Join(1500); }
    }

    public void Rebind()
    {
        rebindWanted = true;
    }

    private void Loop()
    {
        bool saidWaiting = false;
        while (!stop)
        {
            if (rebindWanted)
            {
                rebindWanted = false;
                doc = null; val = null; win = null;
                lastFired = ""; pending = ""; sawEmpty = true;
            }
            if (doc == null)
            {
                if (!Bind())
                {
                    if (!saidWaiting) { Fire2("WAITING", ""); saidWaiting = true; }
                    Thread.Sleep(400);
                    continue;
                }
                saidWaiting = false;
                Fire2("WATCHING", title + "  (hwnd " + hwnd.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")");
                pending = ""; pendingPolls = 0;
            }

            string raw = ReadField();
            if (raw == null) { Thread.Sleep(PollMs); continue; }

            string cand = Candidate(raw);
            if (OnRaw != null) { OnRaw(cand); }

            if (cand.Length == 0)
            {
                sawEmpty = true;
                pending = "";
                pendingPolls = 0;
            }
            else if (cand != pending)
            {
                pending = cand;
                pendingSince = Stopwatch.GetTimestamp();
                pendingPolls = 1;
            }
            else
            {
                pendingPolls++;
                double held = Rdv3Clock.MsSince(pendingSince);
                if (held >= StableMs && Rdv3Spec.IsKey(cand) && (cand != lastFired || sawEmpty))
                {
                    lastFired = cand;
                    sawEmpty = false;
                    long confirmAt = Stopwatch.GetTimestamp();
                    if (OnConfirmed != null)
                    {
                        OnConfirmed(cand, Rdv3Clock.MsBetween(pendingSince, confirmAt), pendingPolls, confirmAt);
                    }
                }
            }
            Thread.Sleep(PollMs);
        }
    }

    private void Fire2(string a, string b)
    {
        if (OnState != null) { OnState(a, b); }
    }
}

public static class Rdv3NativeFg
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
