// ============================================================================
// RdvWatch.cs -- reading detection. Polls a Notepad window through UI
// Automation and decides when a number has actually been read.
//
// A card reader types one character at a time, so "the text changed" is not the
// same as "a number arrived". Three rules together, all of them necessary:
//
//   format+length   exactly 8 digits, so a half-typed 4-digit prefix is nothing
//   quiet time      the same value has to survive StableMs of polling, so a
//                   reader that emits 10 digits never fires on its first 8
//   no repeat       the same value does not fire twice in a row...
//                   ...unless the field went empty in between, which is a
//                   deliberate re-read of the same card and must be allowed
//
// The wait for those rules to be satisfied is detection latency. It is reported
// separately and is NOT part of the merge-select measurement: the merge-select
// clock starts at the instant the value is confirmed.
//
// The watcher never starts, closes or kills Notepad. It attaches to a window
// that is already there and only ever reads from it.
// ============================================================================

using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;

public sealed class RdvWatch
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

    public static bool IsKey(string s)
    {
        if (s == null || s.Length != RdvSpec.KeyLen) { return false; }
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] < '0' || s[i] > '9') { return false; }
        }
        return true;
    }

    // What the reader last delivered: the final non-empty line of the document.
    // A reader that ends with Enter leaves the number on the line above, and a
    // stale earlier line must not keep re-firing.
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
    // Notepad windows only, and prefer the one in the foreground so the user can
    // choose which window feeds the app just by clicking it.
    public bool Bind()
    {
        try
        {
            AutomationElement root = AutomationElement.RootElement;
            AutomationElementCollection wins = root.FindAll(TreeScope.Children,
                new PropertyCondition(AutomationElement.ClassNameProperty, "Notepad"));
            if (wins.Count == 0) { doc = null; val = null; win = null; return false; }

            IntPtr fg = NativeFg.GetForegroundWindow();
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

    // Windows 11 Notepad is a RichEditD2DPT Document; the classic one is an Edit.
    // Both expose ValuePattern, so ask for the pattern rather than the class.
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

    public string Read()
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

    // let the operator point the app at a different Notepad window: drop the
    // binding and let the loop pick up whatever is in front now
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

            string raw = Read();
            if (raw == null) { Thread.Sleep(PollMs); continue; }

            string cand = Candidate(raw);
            if (OnRaw != null) { OnRaw(cand); }

            if (cand.Length == 0)
            {
                // the field was cleared: the same number may be read again
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
                double held = RdvEngine.MsSince(pendingSince);
                if (held >= StableMs && IsKey(cand) && (cand != lastFired || sawEmpty))
                {
                    lastFired = cand;
                    sawEmpty = false;
                    long confirmAt = Stopwatch.GetTimestamp();
                    if (OnConfirmed != null)
                    {
                        OnConfirmed(cand, RdvEngine.MsBetween(pendingSince, confirmAt), pendingPolls, confirmAt);
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

public static class NativeFg
{
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();
}
