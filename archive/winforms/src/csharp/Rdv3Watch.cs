// ============================================================================
// Rdv3Watch.cs -- reading detection, driven entirely by UI Automation and
// entirely by the settings file.
//
//   format          the key rule from settings.json (search.pattern)
//   quiet time      the value must survive stableMs of polling
//   no repeat       the same value does not fire twice in a row from the same
//                   source, unless the field went empty in between
//
// SEVERAL TARGETS AT ONCE. Every enabled target in the settings file is bound
// and polled; whichever confirms a key first drives the search. Each source
// keeps its own pending/last-fired state, so two applications showing the same
// number are two separate readings, and one going away does not disturb the
// others. Targets that are not running yet are retried every rebindMs.
//
// The watcher never starts, closes or kills anything. It attaches to windows
// that are already there and only ever reads from them.
//
// A confirmed number triggers a SEARCH of the current ledger -- not a
// re-merge. The merge belongs to the startup update check.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Windows.Automation;

public sealed class Rdv3Watch
{
    // one bound target: the window, the field inside it, and this source's own
    // idea of what it has already reported
    private sealed class Src
    {
        public Rdv3Target T;
        public AutomationElement Win;
        public AutomationElement Field;
        public string Title = "";
        public int Hwnd;
        public string Pending = "";
        public long PendingSince;
        public int PendingPolls;
        public string LastFired = "";
        public bool SawEmpty = true;
    }

    public Rdv3Config Cfg = new Rdv3Config();

    private readonly List<Src> live = new List<Src>();
    private readonly object gate = new object();
    private long lastBindTry;
    private volatile bool stop;
    private volatile bool rebindWanted;
    private Thread thread;
    private string activeName = "";
    private string activeDetail = "";

    // key, detection latency in ms, poll count, the confirm timestamp, and the
    // name of the target it came from: several targets are watched at once, so a
    // number recorded without its source cannot be traced back to the screen it
    // was read off
    public Action<string, double, int, long, string> OnConfirmed;
    public Action<string, string> OnState;      // state, detail
    public Action<string> OnLabel;              // which target the status line names

    public bool Bound
    {
        get { lock (gate) { return live.Count > 0; } }
    }

    // What the reader last delivered: the final non-empty line of the field.
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
    // window -> optional path steps -> field, each matched on the properties UIA
    // exposes. Nothing is guessed and nothing is looked up by window handle.
    private AutomationElement Resolve(Rdv3Target t, AutomationElement win)
    {
        AutomationElement at = win;
        for (int i = 0; i < t.Steps.Count && at != null; i++)
        {
            at = Rdv3Uia.FindOne(at, t.Steps[i], t.Steps[i].Descendants);
        }
        if (at == null) { return null; }
        return Rdv3Uia.FindOne(at, t.Field, t.Field.Descendants);
    }

    private bool AlreadyBound(Rdv3Target t)
    {
        for (int i = 0; i < live.Count; i++) { if (live[i].T == t) { return true; } }
        return false;
    }

    // Bind every enabled target that is not bound yet. Called at most every
    // rebindMs so a target that is not running costs one search per second,
    // not one per poll.
    private bool BindMissing()
    {
        bool changed = false;
        AutomationElement root;
        try { root = AutomationElement.RootElement; }
        catch (Exception) { return false; }
        AutomationElement focused = Cfg.PreferFocusedWindow ? Rdv3Uia.FocusedWindow() : null;
        List<Rdv3Target> targets = Cfg.Targets;

        for (int i = 0; i < targets.Count; i++)
        {
            // IsWatchable, not Enabled: the settings dialog hands over targets it
            // built in memory, which never went through Rdv3Target.Read, so one
            // added and saved with nothing filled in arrives still calling itself
            // enabled -- and an empty matcher accepts any window, which would
            // attach the reader to whatever happened to be in front.
            Rdv3Target t = targets[i];
            if (!t.IsWatchable) { continue; }
            lock (gate) { if (AlreadyBound(t)) { continue; } }

            List<AutomationElement> wins = Rdv3Uia.FindAll(root, t.Window, false);
            if (wins.Count == 0) { continue; }

            AutomationElement pick = null;
            if (focused != null)
            {
                for (int k = 0; k < wins.Count; k++)
                {
                    if (Rdv3Uia.Same(wins[k], focused)) { pick = wins[k]; break; }
                }
            }
            // "which one of the several that match", not "the last one there
            // happens to be". If the window they numbered is not open, the target
            // simply is not bound yet and the status line keeps saying so; sliding
            // to a different window would read a number off somebody else's work.
            // Rdv3Uia.FindOne answers the same way for the path steps and the field.
            if (pick == null)
            {
                if (t.Window.Index >= wins.Count) { continue; }
                pick = wins[t.Window.Index];
            }

            AutomationElement field = Resolve(t, pick);
            if (field == null) { continue; }
            if (!Rdv3Uia.CanRead(field, t.ReadMode)) { continue; }

            Src s = new Src();
            s.T = t;
            s.Win = pick;
            s.Field = field;
            try { s.Title = pick.Current.Name; s.Hwnd = pick.Current.NativeWindowHandle; }
            catch (Exception) { s.Title = ""; }
            lock (gate) { live.Add(s); }
            changed = true;
        }
        return changed;
    }

    private void Drop(Src s)
    {
        lock (gate) { live.Remove(s); }
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
                lock (gate) { live.Clear(); }
                lastBindTry = 0;
                activeDetail = "";
                saidWaiting = false;
            }

            bool haveAll;
            lock (gate) { haveAll = (live.Count >= WatchableCount()); }
            if (!haveAll && (lastBindTry == 0 || Rdv3Clock.MsSince(lastBindTry) >= Cfg.RebindMs))
            {
                lastBindTry = Rdv3Clock.Now();
                if (BindMissing()) { Announce(); saidWaiting = false; }
            }

            Src[] now;
            lock (gate) { now = live.ToArray(); }
            if (now.Length == 0)
            {
                if (!saidWaiting) { Fire2("WAITING", ""); saidWaiting = true; }
                Thread.Sleep(Math.Min(Cfg.RebindMs, 400));
                continue;
            }

            // the focused window's source is looked at first, so that when two
            // applications hold a value the one the operator is using wins
            if (Cfg.PreferFocusedWindow && now.Length > 1)
            {
                AutomationElement f = Rdv3Uia.FocusedWindow();
                if (f != null)
                {
                    for (int i = 0; i < now.Length; i++)
                    {
                        if (Rdv3Uia.Same(now[i].Win, f))
                        {
                            Src tmp = now[0]; now[0] = now[i]; now[i] = tmp;
                            break;
                        }
                    }
                }
            }

            bool lost = false;
            for (int i = 0; i < now.Length; i++)
            {
                Src s = now[i];
                string raw;
                try { raw = Rdv3Uia.ReadValueOf(s.Field, s.T.ReadMode); }
                catch (ElementNotAvailableException) { Drop(s); lost = true; continue; }
                if (raw == null) { Drop(s); lost = true; continue; }

                string cand = Candidate(raw);

                if (cand.Length == 0)
                {
                    s.SawEmpty = true;
                    s.Pending = "";
                    s.PendingPolls = 0;
                    continue;
                }
                if (cand != s.Pending)
                {
                    s.Pending = cand;
                    s.PendingSince = Rdv3Clock.Now();
                    s.PendingPolls = 1;
                    continue;
                }
                s.PendingPolls++;
                double held = Rdv3Clock.MsSince(s.PendingSince);
                if (held >= Cfg.StableMs && Cfg.IsKey(cand) && (cand != s.LastFired || s.SawEmpty))
                {
                    s.LastFired = cand;
                    s.SawEmpty = false;
                    activeName = s.T.Name;
                    activeDetail = Detail(s);
                    Announce();
                    long confirmAt = Rdv3Clock.Now();
                    if (OnConfirmed != null)
                    {
                        OnConfirmed(cand, Rdv3Clock.MsBetween(s.PendingSince, confirmAt), s.PendingPolls,
                            confirmAt, s.T.Name);
                    }
                    break;                     // one reading per tick
                }
            }
            if (lost) { Announce(); }
            Thread.Sleep(Cfg.PollMs);
        }
    }

    // the same question BindMissing asks, so the two cannot disagree and leave
    // the loop searching the whole desktop every rebindMs for a target it will
    // never take
    private int WatchableCount()
    {
        int n = 0;
        List<Rdv3Target> targets = Cfg.Targets;
        for (int i = 0; i < targets.Count; i++) { if (targets[i].IsWatchable) { n++; } }
        return n;
    }

    // the window's title; the handle is logged by the app, never shown
    private static string Detail(Src s)
    {
        return s.Title;
    }

    // What the status line says: the source that last delivered, or the first
    // bound one, plus how many are being watched when it is more than one.
    private void Announce()
    {
        Src[] now;
        lock (gate) { now = live.ToArray(); }
        if (now.Length == 0) { Fire2("WAITING", ""); return; }
        Src pick = now[0];
        for (int i = 0; i < now.Length; i++)
        {
            if (now[i].T.Name == activeName) { pick = now[i]; break; }
        }
        string detail = Detail(pick);
        if (now.Length > 1)
        {
            detail = detail + " [" + now.Length.ToString(CultureInfo.InvariantCulture) + "]";
        }
        if (OnLabel != null) { OnLabel(pick.T.Name); }
        Fire2("WATCHING", detail);
    }

    private void Fire2(string a, string b)
    {
        if (OnState != null) { OnState(a, b); }
    }
}
