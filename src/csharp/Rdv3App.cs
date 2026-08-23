// ============================================================================
// Rdv3App.cs -- entry point, state machine, timing boundaries, execution log.
//
// The two figures the log carries, and exactly where they start and stop:
//
//   merge time   = readA+readB+readC + idxA+idxB+idxC + joinAB+joinBC, measured
//                  inside the background rebuild the update check runs.
//                  Ledger composition, comparison, carry-over, persistence and
//                  the search index are NOT in it; they are logged separately.
//   search time  = search confirmed (button click / watched detection) ->
//                  the single record is rendered, or the candidate list is
//                  built and rendered. A human choosing from the list is never
//                  in it; the post-pick render is logged as "display".
//
// States: BOOT -> CHECKING -> (confirm) -> APPLYING -> READY, or READY with
// the saved ledger when there is no difference or the operator declines, or
// BLOCKED when no ledger exists and none may be built. The refreshLedger button runs
// the same check again from READY. Stale worker results (superseded or timed
// out run IDs) are discarded and logged, never applied.
//
// Nothing runs on broken input. settings.json is read strictly before the
// window opens (Rdv3Program.Run), the definition is checked against the CSV
// headers there too, and a CSV that fails the row checks during the merge
// (Rdv3DataError) stops the app with the reason -- never a fallback.
//
// The work state of a record (todo / done ...) is whatever the screen
// definition says it is. The app only knows: the stored string, the
// transition the button asks for, and that a transition is persisted at once,
// one record at a time, with the exit held until the write is decided.
//
// The execution log is always on: one tab-separated line per event, next to
// the .cmd. The screen shows notices and errors as toasts.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

public sealed class Rdv3Log
{
    private readonly string path;
    private readonly object gate = new object();
    private bool failed;
    public Action<string> OnFail;

    public Rdv3Log(string p)
    {
        path = p;
    }

    public string Path { get { return path; } }

    public void Write(string runId, string section, string detail)
    {
        StringBuilder sb = new StringBuilder(160);
        sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
        sb.Append('\t').Append(runId);
        sb.Append('\t').Append(section);
        sb.Append('\t').Append(detail.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' '));
        sb.Append("\r\n");
        lock (gate)
        {
            try
            {
                File.AppendAllText(path, sb.ToString(), new UTF8Encoding(false));
                failed = false;
            }
            catch (Exception ex)
            {
                // a log that cannot be written must not crash the app, but it
                // must not fail silently either: surface it once
                if (!failed)
                {
                    failed = true;
                    Action<string> h = OnFail;
                    if (h != null) { h(ex.Message); }
                }
            }
        }
    }

    public static string F(double ms)
    {
        return ms.ToString("F2", CultureInfo.InvariantCulture);
    }
}

public sealed class Rdv3App
{
    private const int StBoot = 0;
    private const int StChecking = 1;
    private const int StApplying = 2;
    private const int StReady = 3;
    private const int StBlocked = 4;

    private readonly Rdv3Config cfg;
    private readonly Rdv3Screen screen;
    private readonly Rdv3WorkState work;
    private readonly Rdv3Data dataDef;

    private readonly Rdv3Form form;
    private readonly Rdv3Worker worker = new Rdv3Worker();
    private readonly Rdv3Watch watch = new Rdv3Watch();
    private readonly Rdv3Log log;
    // qualified: the packer hoists every using into one file, where a bare
    // Timer is ambiguous with System.Threading.Timer
    private readonly System.Windows.Forms.Timer watchdog = new System.Windows.Forms.Timer();

    private readonly string dataDir;
    private readonly string ledgerPath;
    private readonly int pid;

    private int state = StBoot;
    private bool watchStarted;

    // the active ledger. Mutated only inside worker jobs (the queue serialises
    // them); the UI receives rendered copies and never touches these arrays.
    private string[] ledLines;
    private string[] ledStates;
    private string[] ledHead;
    private Rdv3Index ledIndex;

    // what the check produced, held between CHECKING and the decision
    private Rdv3MergeResult mergeResult;
    private string[] savedLines;
    private string[] savedStates;

    private int runSeq;
    private int searchSeq;
    private int procSeq;
    private string activeRunId = "";
    private string activeSearchId = "";

    private int shownRow = -1;
    private string shownKey = "";
    private List<int> shownCands;
    private double lastMergeMs = -1;

    // ---- the exit guard around one unfinished state save --------------------
    // A work-state change is persisted immediately, one record at a time
    // (there is no queue and nothing is batched). The one moment its outcome
    // is not yet decided is while Rdv3Xlsx.Write runs on the worker thread,
    // and a close in that window would end the process without the operator
    // ever learning whether the record was saved. So during that window: no
    // second change is accepted, and a close request is refused with the
    // reason on screen. Both flags are touched on the UI thread only.
    private bool savingMark;
    private bool closeAskedWhileSaving;

    // app construction instant, for the boot->operable startup figure
    private readonly long bootT0 = Rdv3Clock.Now();
    private bool startupLogged;

    public Rdv3App(Rdv3Form f, string data, string ledger, string logPath, Rdv3Config settings)
    {
        form = f;
        cfg = settings;
        screen = settings.Screen;
        work = screen.Work;
        dataDef = settings.Data;
        dataDir = data;
        ledgerPath = ledger;
        pid = System.Diagnostics.Process.GetCurrentProcess().Id;
        log = new Rdv3Log(logPath);
        log.OnFail = delegate(string msg) { form.Error(Rdv3Text.ErrLogWrite + msg); };
        // the ledger columns the screen binds to, fixed by the definition
        form.SetFields(new Rdv3Fields(dataDef.ColumnRefs));

        worker.OnError = JobFailed;

        form.OnSearch = ManualSearch;
        form.OnClear = DoClear;
        form.OnWorkState = DoWorkState;
        form.OnRefreshLedger = RefreshLedger;
        form.OnPick = PickCandidate;
        form.OnSettings = OpenSettings;

        watch.Cfg = cfg;
        watch.OnConfirmed = Detected;
        watch.OnState = WatchState;
        watch.OnLabel = delegate(string name) { form.SetWatch(name, watchDetail); };
        form.SetWatch(WatchName(), Rdv3Text.WatchNone);
        form.SetIdentity("PID " + pid.ToString(CultureInfo.InvariantCulture), System.IO.Path.GetFileName(logPath));
        form.SetState(Rdv3Text.StateBoot);
        form.SetLedger(Rdv3Text.NotYet, "", "");

        watchdog.Interval = cfg.PumpMs;
        watchdog.Tick += delegate(object s, EventArgs e) { CheckOverdue(); };

        form.Shown += delegate(object s, EventArgs e)
        {
            log.Write("-", "screen", form.Diag);
            watchdog.Start();
            StartCheck();
        };
        form.FormClosing += delegate(object s, FormClosingEventArgs e)
        {
            // one unfinished save is the only thing that refuses a close
            if (savingMark)
            {
                e.Cancel = true;
                closeAskedWhileSaving = true;
                form.Error(Rdv3Text.ErrCloseWhileSaving);
                log.Write("-", "exit", "close refused: a state save is still in flight");
                return;
            }
            Shutdown();
        };
    }

    public void LogBoot(double compileMs)
    {
        log.Write("-", "settings", cfg.Describe());
        log.Write("-", "data", dataDef.Describe());
        log.Write("-", "screen", screen.Describe());
        log.Write("-", "boot", "pid=" + pid.ToString(CultureInfo.InvariantCulture)
            + " method=csharp-dict data=" + dataDir + " ledger=" + ledgerPath
            + " compile_ms=" + Rdv3Log.F(compileMs));
        log.Write("-", "worker", "kind=thread owner_pid=" + pid.ToString(CultureInfo.InvariantCulture));
    }

    // ---- the update check (at start-up, and from the refreshLedger button) ----
    private void StartCheck()
    {
        state = StChecking;
        runSeq++;
        string rid = "R" + runSeq.ToString(CultureInfo.InvariantCulture);
        activeRunId = rid;
        form.SetState(Rdv3Text.StateChecking);
        form.EnableOps(false);
        log.Write(rid, "decision", "check started");

        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "check";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate { CheckJob(rid); };
        worker.Start();
        worker.Post(job);
    }

    private void RefreshLedger()
    {
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        if (savingMark) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        log.Write("-", "decision", "check requested from the screen");
        StartCheck();
    }

    // worker thread
    private void CheckJob(string rid)
    {
        long t = Rdv3Clock.Now();
        Rdv3MergeResult mr = Rdv3Ledger.BuildFromCsv(dataDef, dataDir);
        double composeMs = Rdv3Clock.MsSince(t) - mr.MergeMs();

        for (int i = 0; i < dataDef.Tables.Count; i++)
        {
            log.Write(rid, "read", "table=" + dataDef.Tables[i].Id + " ms=" + Rdv3Log.F(mr.ReadMs[i]));
        }
        for (int i = 0; i < dataDef.Tables.Count; i++)
        {
            log.Write(rid, "index", "table=" + dataDef.Tables[i].Id + " keys=" + mr.Keys[i].ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(mr.IndexMs[i]));
        }
        for (int i = 0; i < dataDef.Joins.Count; i++)
        {
            log.Write(rid, "join", "pair=" + dataDef.Spine + dataDef.Joins[i].Table + " on=" + dataDef.Joins[i].On
                + " matched=" + mr.Matched[i].ToString(CultureInfo.InvariantCulture) + " ms=" + Rdv3Log.F(mr.JoinMs[i]));
        }
        log.Write(rid, "merge", "rows=" + mr.Rows.ToString(CultureInfo.InvariantCulture)
            + " checksum=" + mr.Checksum.ToString(CultureInfo.InvariantCulture)
            + " ms=" + Rdv3Log.F(mr.MergeMs()) + " compose_ms=" + Rdv3Log.F(composeMs));

        // the same verification the comparison builds ran; logged, and NG is an
        // error on screen. Skipped (and said so) when there is no expected.txt.
        Rdv3Expected exp = Rdv3Expected.Read(dataDir);
        bool oracleOk = true;
        if (exp.Loaded)
        {
            oracleOk = (exp.Rows == mr.Rows && exp.Checksum == mr.Checksum);
            log.Write(rid, "verify", "rows=" + mr.Rows.ToString(CultureInfo.InvariantCulture)
                + " checksum=" + mr.Checksum.ToString(CultureInfo.InvariantCulture)
                + " expected_rows=" + exp.Rows.ToString(CultureInfo.InvariantCulture)
                + " expected_checksum=" + exp.Checksum.ToString(CultureInfo.InvariantCulture)
                + " ok=" + (oracleOk ? "true" : "false"));
        }
        else
        {
            log.Write(rid, "verify", "skipped (no expected.txt)");
        }

        // saved ledger
        string[] oldLines = null;
        string[] oldStates = null;
        string loadError = null;
        bool exists = File.Exists(ledgerPath);
        if (exists)
        {
            try
            {
                t = Rdv3Clock.Now();
                Rdv3Xlsx.Read(ledgerPath, mr.Head, work.Column, out oldLines, out oldStates);
                log.Write(rid, "load", "ledger rows=" + oldLines.Length.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            }
            catch (Exception ex)
            {
                loadError = ex.Message;
                log.Write(rid, "error", "stage=load msg=" + ex.Message);
            }
        }
        else
        {
            log.Write(rid, "load", "ledger missing: " + ledgerPath);
        }

        bool same = false;
        if (oldLines != null)
        {
            t = Rdv3Clock.Now();
            int firstDiff;
            same = Rdv3Ledger.SameContent(oldLines, mr.Lines, out firstDiff);
            log.Write(rid, "compare", "rows_old=" + oldLines.Length.ToString(CultureInfo.InvariantCulture)
                + " rows_new=" + mr.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " same=" + (same ? "true" : "false")
                + (same ? "" : " first_diff_row=" + (firstDiff + 1).ToString(CultureInfo.InvariantCulture))
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
        }

        Rdv3MergeResult mrKeep = mr;
        string[] keepOld = oldLines;
        string[] keepStates = oldStates;
        string keepErr = loadError;
        bool keepExists = exists;
        bool keepSame = same;
        bool keepOracle = oracleOk;
        form.RunOnUi(delegate { EndCheck(rid, mrKeep, keepOld, keepStates, keepExists, keepErr, keepSame, keepOracle); });
    }

    // UI thread
    private void EndCheck(string rid, Rdv3MergeResult mr, string[] oldLines, string[] oldStates,
                          bool ledgerExists, string loadError, bool same, bool oracleOk)
    {
        if (!string.Equals(rid, activeRunId, StringComparison.Ordinal))
        {
            log.Write(rid, "stale", "check result discarded (active=" + activeRunId + ")");
            return;
        }
        mergeResult = mr;
        savedLines = oldLines;
        savedStates = oldStates;
        lastMergeMs = mr.MergeMs();
        form.SetTimes(lastMergeMs, -1);
        if (!oracleOk) { form.Error(Rdv3Text.ErrOracle); }

        if (oldLines == null)
        {
            // missing or unreadable: never silently rebuild -- ask, with the
            // reason on the screen
            string body = ledgerExists
                ? Rdv3Text.ConfirmRebuildBody.Replace("{err}", (loadError == null) ? "?" : loadError)
                : Rdv3Text.ConfirmCreateBody;
            if (ledgerExists) { form.Error(Rdv3Text.ErrLedgerRead + ((loadError == null) ? "" : loadError)); }
            bool yes = form.Ask(Rdv3Text.ConfirmUpdateTitle, body);
            log.Write(rid, "decision", ledgerExists
                ? ("ledger unreadable; rebuild " + (yes ? "approved" : "declined"))
                : ("ledger missing; create " + (yes ? "approved" : "declined")));
            if (yes) { StartApply(rid, false); }
            else if (ledLines != null) { ContinueWithActive(rid, Rdv3Text.NoteRejected); }
            else { EnterBlocked(); }
            return;
        }

        if (same)
        {
            log.Write(rid, "decision", "no difference");
            if (ledLines != null && runSeq > 1) { ContinueWithActive(rid, Rdv3Text.NoteNoDiff); form.Notice(Rdv3Text.NoteNoDiff); }
            else { AdoptLedger(rid, oldLines, oldStates, mr.Head, Rdv3Text.NoteNoDiff); }
            return;
        }

        bool approve = form.Ask(Rdv3Text.ConfirmUpdateTitle, Rdv3Text.ConfirmUpdateBody);
        log.Write(rid, "decision", "difference found; update " + (approve ? "approved" : "rejected"));
        if (approve)
        {
            StartApply(rid, true);
        }
        else if (ledLines != null)
        {
            ContinueWithActive(rid, Rdv3Text.NoteRejected);
        }
        else
        {
            AdoptLedger(rid, oldLines, oldStates, mr.Head, Rdv3Text.NoteRejected);
        }
    }

    // a re-check from READY that changes nothing: the active ledger stands
    private void ContinueWithActive(string rid, string note)
    {
        form.RunOnUi(delegate { EnterReady(rid, note); });
    }

    // adopt = make these lines the active, searchable ledger (worker builds the
    // index so the UI thread never runs over 100,000 rows)
    private void AdoptLedger(string rid, string[] lines, string[] states, string[] head, string note)
    {
        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "adopt";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate
        {
            long t = Rdv3Clock.Now();
            Rdv3Index ix = new Rdv3Index(Rdv3Ledger.Column(lines, dataDef.SearchCol));
            log.Write(rid, "index", "table=LEDGER rows=" + lines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            ledLines = lines;
            ledStates = states;
            ledHead = head;
            ledIndex = ix;
            form.RunOnUi(delegate { EnterReady(rid, note); });
        };
        worker.Post(job);
    }

    private void StartApply(string rid, bool carry)
    {
        state = StApplying;
        form.SetState(Rdv3Text.StateApplying);

        Rdv3MergeResult mr = mergeResult;
        string[] oldLines = savedLines;
        string[] oldStates = savedStates;
        string initial = work.InitialStored;

        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "apply";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate
        {
            long t = Rdv3Clock.Now();
            string[] ns;
            if (carry && oldLines != null)
            {
                Rdv3Ledger.CarryStats st = new Rdv3Ledger.CarryStats();
                ns = Rdv3Ledger.CarryStates(oldLines, oldStates, mr.Lines, dataDef.IdentityCol, initial, st);
                log.Write(rid, "carry", "carried=" + st.Carried.ToString(CultureInfo.InvariantCulture)
                    + " reset=" + st.Reset.ToString(CultureInfo.InvariantCulture)
                    + " new=" + st.New.ToString(CultureInfo.InvariantCulture)
                    + " dropped=" + st.Dropped.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            }
            else
            {
                ns = Rdv3Ledger.FreshStates(mr.Lines.Length, initial);
                log.Write(rid, "carry", "none (fresh ledger)");
            }

            t = Rdv3Clock.Now();
            try
            {
                Rdv3Xlsx.Write(ledgerPath, mr.Head, work.Column, mr.Lines, ns,
                    pid.ToString(CultureInfo.InvariantCulture) + "-" + rid);
            }
            catch (Exception ex)
            {
                log.Write(rid, "error", "stage=persist msg=" + ex.Message);
                form.RunOnUi(delegate { EndApplyFailed(rid, ex.Message); });
                return;
            }
            log.Write(rid, "persist", "target=xlsx rows=" + mr.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));

            t = Rdv3Clock.Now();
            Rdv3Index ix = new Rdv3Index(Rdv3Ledger.Column(mr.Lines, dataDef.SearchCol));
            log.Write(rid, "index", "table=LEDGER rows=" + mr.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));

            ledLines = mr.Lines;
            ledStates = ns;
            ledHead = mr.Head;
            ledIndex = ix;
            form.RunOnUi(delegate { EnterReady(rid, Rdv3Text.NoteUpdated); form.Notice(Rdv3Text.NoteUpdated); });
        };
        worker.Post(job);
    }

    // UI thread. The persist failed: the saved ledger file is untouched, so the
    // saved content stays the active one (or the app is blocked if none).
    private void EndApplyFailed(string rid, string msg)
    {
        if (!string.Equals(rid, activeRunId, StringComparison.Ordinal))
        {
            log.Write(rid, "stale", "apply failure discarded");
            return;
        }
        form.Error(Rdv3Text.ErrPersist + msg);
        if (ledLines != null) { ContinueWithActive(rid, Rdv3Text.NoteRejected); }
        else if (savedLines != null) { AdoptLedger(rid, savedLines, savedStates, mergeResult.Head, Rdv3Text.NoteRejected); }
        else { EnterBlocked(); }
    }

    // the ledger's own last-write time, in the reference's short form
    private string LedgerStamp()
    {
        try
        {
            if (System.IO.File.Exists(ledgerPath))
            {
                return System.IO.File.GetLastWriteTime(ledgerPath)
                    .ToString("MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
        }
        catch (Exception) { }
        return Rdv3Text.NotYet;
    }

    // UI thread
    private void EnterReady(string rid, string note)
    {
        if (!string.Equals(rid, activeRunId, StringComparison.Ordinal))
        {
            log.Write(rid, "stale", "ready transition discarded");
            return;
        }
        state = StReady;
        activeRunId = "";
        mergeResult = null;
        form.EnableOps(true);
        form.EnableWorkState(!savingMark);
        int tone;
        form.SetState(WatchStateText(out tone));
        // the result on show belongs to the ledger that was just replaced, so
        // it is cleared; the search key stays typed for a repeat search
        if (shownCands != null) { ClearShown(); }
        string rowsText = ledLines.Length.ToString("N0", CultureInfo.InvariantCulture);
        form.SetLedger(Rdv3Text.LedgerSegFmt.Replace("{file}", System.IO.Path.GetFileName(ledgerPath)).Replace("{n}", rowsText),
            rowsText, LedgerStamp());
        log.Write(rid, "decision", "ready rows=" + ledLines.Length.ToString(CultureInfo.InvariantCulture)
            + " note=" + note);
        if (!startupLogged)
        {
            // boot->operable, excluding the cmd's csc compile (compile_ms on
            // the boot line) and the console/process start before it
            startupLogged = true;
            log.Write(rid, "startup", "boot_to_ready_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(bootT0)));
        }
        if (!watchStarted)
        {
            watchStarted = true;
            watch.Start();
        }
    }

    private void ClearShown()
    {
        string typed = form.KeyText;
        form.ClearResult();
        form.SetKeyText(typed);
        shownRow = -1;
        shownKey = "";
        shownCands = null;
    }

    // UI thread
    private void EnterBlocked()
    {
        state = StBlocked;
        activeRunId = "";
        form.EnableOps(false);
        form.SetState(Rdv3Text.StateBlocked);
        form.Error(Rdv3Text.ErrNoLedger);
        log.Write("-", "decision", "blocked (no ledger)");
    }

    // ---- search ------------------------------------------------------------
    private void ManualSearch(string key)
    {
        long t0 = Rdv3Clock.Now();
        if (state != StReady)
        {
            form.Error(Rdv3Text.ErrNotReady);
            log.Write("-", "search", "ignored key=" + key + " reason=not-ready");
            return;
        }
        if (Rdv3Config.PatternError(cfg.KeyPattern) != null)
        {
            form.Error(Rdv3Text.ErrBadPattern);
            log.Write("-", "search", "ignored key=" + key + " reason=bad-pattern");
            return;
        }
        if (!cfg.IsKey(key))
        {
            form.Error(Rdv3Text.ErrBadKeyFmt.Replace("{pattern}", cfg.KeyPattern));
            log.Write("-", "search", "ignored key=" + key + " reason=bad-key");
            return;
        }
        Search(key, "manual", "", 0, t0);
    }

    // watch thread. "target" is the name of the watched screen the number came
    // off: with several watched at once, a reading in the log is only traceable
    // back to the work it belongs to if it says which one delivered it.
    private void Detected(string key, double detectMs, int polls, long t0, string target)
    {
        log.Write("-", "detect", "key=" + key + " target=" + target
            + " latency_ms=" + Rdv3Log.F(detectMs)
            + " polls=" + polls.ToString(CultureInfo.InvariantCulture));
        if (state != StReady)
        {
            log.Write("-", "search", "ignored key=" + key + " target=" + target
                + " reason=not-ready (detected)");
            return;
        }
        form.SetKeyText(key);
        Search(key, "detect", target, detectMs, t0);
    }

    // any thread. t0 is the confirm instant: the search clock is already running.
    private void Search(string key, string source, string target, double detectMs, long t0)
    {
        searchSeq++;
        string sid = "S" + searchSeq.ToString(CultureInfo.InvariantCulture);
        activeSearchId = sid;

        Rdv3Job job = new Rdv3Job();
        job.RunId = sid;
        job.Kind = "search";
        job.TimeoutMs = cfg.SearchTimeoutMs;
        job.Work = delegate { SearchJob(sid, key, source, target, t0); };
        worker.Post(job);
    }

    // worker thread
    private void SearchJob(string sid, string key, string source, string target, long t0)
    {
        if (!string.Equals(sid, activeSearchId, StringComparison.Ordinal))
        {
            log.Write(sid, "stale", "search superseded before start key=" + key);
            return;
        }
        string from = "source=" + source + ((target.Length > 0) ? (" target=" + target) : "");
        string[] lines = ledLines;
        Rdv3Index ix = ledIndex;
        string[] states = ledStates;
        List<int> hits = ix.Find(key);
        int n = (hits == null) ? 0 : hits.Count;

        int show = (n > cfg.CandidateRowsShown) ? cfg.CandidateRowsShown : n;
        List<Rdv3CandRow> rows = new List<Rdv3CandRow>(show);
        List<int> candRows = new List<int>(show);
        for (int i = 0; i < show; i++)
        {
            Rdv3CandRow r = new Rdv3CandRow();
            r.Line = lines[hits[i]];
            r.Stored = states[hits[i]];
            rows.Add(r);
            candRows.Add(hits[i]);
        }
        form.RunOnUi(delegate
        {
            form.ShowCandidates(key, rows, n);
            shownKey = key;
            shownCands = candRows;
            shownRow = -1;
            if (n == 1)
            {
                // one hit is selected at once, the way the reference does it
                form.SelectCandidate(0);
                shownRow = candRows[0];
            }
        });
        double elapsed = Rdv3Clock.MsSince(t0);
        log.Write(sid, "search", "key=" + key + " " + from
            + " hits=" + n.ToString(CultureInfo.InvariantCulture) + " ms=" + Rdv3Log.F(elapsed));
        if (n > 1)
        {
            log.Write(sid, "candidate", "count=" + n.ToString(CultureInfo.InvariantCulture)
                + " shown=" + show.ToString(CultureInfo.InvariantCulture));
            // the list opens as a modal; the choice is the operator's and is
            // never part of the search time -- and never the worker's wait
            form.PostOnUi(delegate
            {
                if (!string.Equals(sid, activeSearchId, StringComparison.Ordinal)) { return; }
                int picked = form.PickFromList();
                if (picked >= 0) { PickCandidate(picked); }
            });
        }
        form.SetTimes(-1, elapsed);
    }

    // UI thread. The pick happens after the search clock stopped; its render is
    // the separate "display" figure in the log.
    private void PickCandidate(int i)
    {
        List<int> cands = shownCands;
        if (state != StReady || cands == null || i < 0 || i >= cands.Count) { return; }
        long t0 = Rdv3Clock.Now();
        int row = cands[i];
        string k2 = Rdv3Ledger.FieldOf(ledLines[row], dataDef.IdentityCol);
        form.SelectCandidate(i);
        shownRow = row;
        log.Write("S" + searchSeq.ToString(CultureInfo.InvariantCulture), "display",
            "picked=" + (i + 1).ToString(CultureInfo.InvariantCulture)
            + " key2=" + k2 + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t0)));
    }

    // Which visible candidate that ledger row is RIGHT NOW, or -1. Asked again
    // when the save is decided rather than remembered from when it started: by
    // then the operator may have picked another candidate or searched again,
    // and the row that was saved is the one that has to be told the truth.
    private int CandOf(int ledgerRow)
    {
        List<int> c = shownCands;
        if (c == null) { return -1; }
        for (int i = 0; i < c.Count; i++) { if (c[i] == ledgerRow) { return i; } }
        return -1;
    }

    // ---- the work state ----------------------------------------------------
    // The button asks for the transition that leaves the record's current
    // state. The definition says which one that is, what to confirm, and what
    // to store; the app does the confirming, storing and telling.
    private void DoWorkState()
    {
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        if (savingMark)
        {
            form.Error(Rdv3Text.ErrSaveInFlight);
            log.Write("-", "state", "refused: a state save is still in flight");
            return;
        }
        int row = shownRow;
        Rdv3StateDef init = work.InitialState;
        if (row < 0)
        {
            // the button names the state a fresh record would move to
            Rdv3Transition first = (init == null) ? null : work.FromState(init.Id);
            Rdv3StateDef target = (first == null) ? null : work.ById(first.To);
            form.Error(Rdv3Text.ErrNoRecordShown.Replace("{state}", (target == null) ? "" : target.Text));
            return;
        }
        string stored = ledStates[row];
        Rdv3StateDef cur = work.ByStored(stored);
        if (cur == null)
        {
            form.Error(Rdv3Text.ErrUnknownState.Replace("{stored}", (stored.Length == 0) ? Rdv3Text.StateBlank : stored));
            log.Write("-", "state", "refused: stored value is not a defined state: " + stored);
            return;
        }
        Rdv3Transition tr = work.FromState(cur.Id);
        if (tr == null)
        {
            form.Error(Rdv3Text.ErrNoTransition.Replace("{state}", cur.Text));
            log.Write("-", "state", "refused: no transition from " + cur.Id);
            return;
        }
        Rdv3StateDef to = work.ById(tr.To);
        string k2 = Rdv3Ledger.FieldOf(ledLines[row], dataDef.IdentityCol);
        if (tr.Confirm.Length > 0)
        {
            string body = Rdv3Eval.Template(tr.Confirm, form.View, form.Fields, work);
            string title = Rdv3Text.ConfirmStateTitleFmt.Replace("{state}", to.Text);
            if (!form.Ask(title, body))
            {
                log.Write("-", "state", "declined key2=" + k2 + " to=" + to.Id);
                return;
            }
        }
        procSeq++;
        string pidTag = "P" + procSeq.ToString(CultureInfo.InvariantCulture);
        long t0 = Rdv3Clock.Now();          // confirm instant: the E2E clock

        // the save starts now and its outcome is undecided until the worker
        // reports back: say so on screen, refuse a second change, refuse a close
        savingMark = true;
        closeAskedWhileSaving = false;
        form.EnableWorkState(false);
        form.SetState(Rdv3Text.StateSavingFmt.Replace("{state}", to.Text));
        form.SetStoredState(CandOf(row), to.Stored, true);
        log.Write(pidTag, "state", "save started key2=" + k2 + " from=" + cur.Id + " to=" + to.Id + " (exit held until it is decided)");

        string was = stored;
        string next = to.Stored;
        Rdv3Job job = new Rdv3Job();
        job.RunId = pidTag;
        job.Kind = "state";
        job.TimeoutMs = cfg.SaveTimeoutMs;
        job.Work = delegate
        {
            ledStates[row] = next;
            long t = Rdv3Clock.Now();
            try
            {
                Rdv3Xlsx.Write(ledgerPath, ledHead, work.Column, ledLines, ledStates,
                    pid.ToString(CultureInfo.InvariantCulture) + "-" + pidTag);
            }
            catch (Exception ex)
            {
                // the file could not change, so the memory must not either:
                // screen and ledger never diverge silently
                ledStates[row] = was;
                log.Write(pidTag, "error", "stage=persist key2=" + k2 + " msg=" + ex.Message);
                form.RunOnUi(delegate
                {
                    form.Error(Rdv3Text.ErrPersist + ex.Message);
                    form.SetStoredState(CandOf(row), was, false);
                    EndMarkSave(pidTag, false);
                });
                return;
            }
            double ms = Rdv3Clock.MsSince(t);
            form.RunOnUi(delegate
            {
                form.SetStoredState(CandOf(row), next, false);
                form.SetLedger(Rdv3Text.LedgerSegFmt.Replace("{file}", System.IO.Path.GetFileName(ledgerPath))
                    .Replace("{n}", ledLines.Length.ToString("N0", CultureInfo.InvariantCulture)),
                    ledLines.Length.ToString("N0", CultureInfo.InvariantCulture), LedgerStamp());
                EndMarkSave(pidTag, true);
                form.Notice((tr.Done.Length > 0) ? tr.Done : Rdv3Text.NoteStateSaved.Replace("{state}", to.Text));
            });
            // e2e = confirm click -> file persisted AND the screen updated
            log.Write(pidTag, "state", "key2=" + k2
                + " row=" + (row + 1).ToString(CultureInfo.InvariantCulture)
                + " value=" + next + " persist_ms=" + Rdv3Log.F(ms)
                + " e2e_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t0)));
        };
        worker.Post(job);
    }

    // UI thread. The save is decided -- saved or failed, either is a decision --
    // so operations and the close are released again.
    private void EndMarkSave(string tag, bool ok)
    {
        if (!savingMark) { return; }
        savingMark = false;
        form.EnableWorkState(state == StReady);
        if (state == StReady)
        {
            int tone;
            form.SetState(WatchStateText(out tone));
        }
        log.Write(tag, "exit", "state save decided (" + (ok ? "saved" : "failed") + "); exit released");
        if (closeAskedWhileSaving)
        {
            closeAskedWhileSaving = false;
            if (ok) { form.Notice(Rdv3Text.NoteSaveDoneCanClose); }
            else { form.Error(Rdv3Text.NoteSaveFailedCanClose); }
        }
    }

    private void DoClear()
    {
        form.ClearResult();
        shownRow = -1;
        shownKey = "";
        shownCands = null;
        log.Write("-", "clear", "input and result cleared");
    }

    // ---- settings ------------------------------------------------------------
    // The settings dialog edits a copy; when it comes back, the file is written
    // and the running session adopts what it can without a restart.
    private void OpenSettings()
    {
        Rdv3Config edited = Rdv3SettingsForm.Edit(form, cfg);
        if (edited == null) { return; }
        string err = edited.Save(cfg.SourcePath);
        if (err != null)
        {
            form.Error(Rdv3Text.ErrSettingsSave + err);
            log.Write("-", "settings", "save failed: " + err);
            return;
        }
        cfg.AdoptRuntimeFrom(edited);
        cfg.AdoptSavedFrom(edited);
        form.SetWatch(WatchName(), watchDetail);
        watch.Rebind();
        log.Write("-", "settings", "saved: " + cfg.Describe());
        for (int i = 0; i < cfg.Targets.Count; i++)
        {
            string why = cfg.Targets[i].WhyNotWatchable();
            if (why.Length > 0)
            {
                log.Write("-", "settings", "target [" + cfg.Targets[i].Name + "] is not watched: " + why);
            }
        }
        form.Notice(Rdv3Text.NoteSettingsApplied);
    }

    // ---- watch / timeout / shutdown ---------------------------------------
    private string watchDetail = Rdv3Text.WatchNone;

    // What the screen calls the thing being watched -- counted over the targets
    // that are actually watched, not over every target in the file.
    private string WatchName()
    {
        Rdv3Target only = null;
        int n = 0;
        for (int i = 0; i < cfg.Targets.Count; i++)
        {
            if (!cfg.Targets[i].IsWatchable) { continue; }
            n++;
            if (n == 1) { only = cfg.Targets[i]; }
        }
        if (n == 1) { return only.Name; }
        if (n == 0) { return Rdv3Text.LabelNotepad; }
        return Rdv3Text.LabelWatch;
    }

    private int WatchableCount()
    {
        int n = 0;
        for (int i = 0; i < cfg.Targets.Count; i++) { if (cfg.Targets[i].IsWatchable) { n++; } }
        return n;
    }

    // WHAT THE STATE WORD SAYS when nothing else is going on
    private string WatchStateText(out int tone)
    {
        if (WatchableCount() == 0) { tone = 2; return Rdv3Text.StateNoTarget; }
        if (watch != null && watch.Bound) { tone = 1; return Rdv3Text.StateReady; }
        tone = 2;
        return Rdv3Text.StateWaitingFmt.Replace("{name}", WatchName());
    }

    private void WatchState(string st, string detail)
    {
        if (st == "WATCHING")
        {
            watchDetail = Rdv3Text.WatchConnectedFmt.Replace("{title}", detail);
        }
        else if (st == "WAITING")
        {
            watchDetail = (WatchableCount() == 0) ? Rdv3Text.WatchNoTarget : Rdv3Text.WatchNone;
        }
        form.SetWatch(WatchName(), watchDetail);
        if ((st == "WATCHING" || st == "WAITING") && state == StReady && !savingMark)
        {
            int tone;
            form.SetState(WatchStateText(out tone));
        }
    }

    private void CheckOverdue()
    {
        Rdv3Job j = worker.TakeOverdue();
        if (j == null) { return; }
        log.Write(j.RunId, "timeout", "kind=" + j.Kind + " after_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(j.StartedAt)));
        if (string.Equals(j.Kind, "state", StringComparison.Ordinal))
        {
            // a managed job cannot be aborted, and until the write returns
            // nobody knows whether the record reached the file. Report the
            // delay; keep holding the exit rather than claim a decision.
            form.Error(Rdv3Text.ErrSaveOverdue);
            return;
        }
        if (string.Equals(j.RunId, activeRunId, StringComparison.Ordinal))
        {
            // abandon the run: whatever arrives later is stale by ID
            activeRunId = "";
            form.Error(Rdv3Text.ErrCheckTimeout);
            if (ledLines != null)
            {
                state = StReady;
                form.EnableOps(true);
                int tone;
                form.SetState(WatchStateText(out tone));
            }
            else if (savedLines != null)
            {
                AdoptTimeoutFallbackSaved();
            }
            else
            {
                EnterBlocked();
            }
        }
    }

    // a job threw: the error goes to the screen and the log, the run ID is
    // abandoned, and the app continues on whatever ledger it validly holds.
    // No alternative method is ever tried in its place.
    private void JobFailed(Rdv3Job job, Exception ex)
    {
        log.Write(job.RunId, "error", "kind=" + job.Kind + " msg=" + ex.GetType().Name + ": " + ex.Message);
        if (ex is Rdv3DataError)
        {
            // the CSVs themselves are not usable: say so and stop. Going on with
            // the ledger in memory would be running on data the operator has
            // just been told is broken.
            log.Write(job.RunId, "exit", "stopping: the data cannot be used");
            // queued, not waited for: this is the worker thread, and the close
            // that follows the modal stops the worker -- a synchronous call
            // here would wait for itself
            form.PostOnUi(delegate
            {
                form.Fatal(Rdv3Text.FatalDataTitle, Rdv3Text.FatalData.Replace("{reason}", ex.Message));
            });
            return;
        }
        form.RunOnUi(delegate
        {
            form.Error(Rdv3Text.ErrCheckFailed + ex.Message);
            // a state job that threw anywhere is still a decided save (failed):
            // the guard must never outlive the job that armed it
            if (string.Equals(job.Kind, "state", StringComparison.Ordinal)) { EndMarkSave(job.RunId, false); }
            if (!string.Equals(job.RunId, activeRunId, StringComparison.Ordinal)) { return; }
            activeRunId = "";
            if (ledLines != null)
            {
                state = StReady;
                form.EnableOps(true);
                int tone;
                form.SetState(WatchStateText(out tone));
            }
            else if (savedLines != null && job.Kind != "adopt" && mergeResult != null)
            {
                AdoptTimeoutFallbackSaved();
            }
            else
            {
                EnterBlocked();
            }
        });
    }

    // a check that timed out after the saved ledger was already read: keep the
    // saved ledger active (explicitly reported above; no other method is tried)
    private void AdoptTimeoutFallbackSaved()
    {
        runSeq++;
        string rid = "R" + runSeq.ToString(CultureInfo.InvariantCulture) + "t";
        activeRunId = rid;
        AdoptLedger(rid, savedLines, savedStates, (mergeResult != null) ? mergeResult.Head : ledHead, Rdv3Text.NoteRejected);
    }

    private void Shutdown()
    {
        log.Write("-", "exit", "closing");
        try { watchdog.Stop(); } catch (Exception) { }
        watch.Stop();
        worker.Stop();
        log.Write("-", "exit", "done");
    }
}

public static class Rdv3Program
{
    // configPath: settings.json (the bootstrap's -config, or the one next to
    // the .cmd). baseDir: the .cmd's folder, what relative paths are relative
    // to. The three overrides are the bootstrap's command line ("" = none).
    public static int Run(string configPath, string baseDir, string dataDirArg, string ledgerArg, string logArg, double compileMs)
    {
        if (IntPtr.Size != 8)
        {
            MessageBox.Show(Rdv3Text.ErrNo64, Rdv3Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }

        // ---- the settings file: exists, parses, checks out -- or nothing runs
        Rdv3Config cfg;
        string settingsDir = "";
        try { settingsDir = Path.GetDirectoryName(Path.GetFullPath(configPath)); } catch (Exception) { }
        try
        {
            cfg = Rdv3Config.Load(configPath);
        }
        catch (Rdv3LoadError ex)
        {
            // the log named by the file is unknown when the file is unusable,
            // so the default log next to it takes the line
            Stop(Path.Combine((settingsDir == null) ? "" : settingsDir, "ReaderDataViewer.log"),
                "settings", "not started: " + configPath + ": " + ex.Message,
                Rdv3Text.FatalTitle, Rdv3Text.FatalSettings.Replace("{file}", configPath).Replace("{reason}", ex.Message));
            return 6;
        }

        string dataDir = Resolve((dataDirArg.Length > 0) ? dataDirArg : cfg.DataDir, baseDir);
        string ledgerPath = Resolve((ledgerArg.Length > 0) ? ledgerArg : cfg.Ledger, baseDir);
        string logPath = Resolve((logArg.Length > 0) ? logArg : cfg.Log, baseDir);

        // ---- the data the definition names: the files exist and their headers
        // hold every column the definition uses
        try
        {
            if (!Directory.Exists(dataDir)) { throw new Rdv3DataError(Rdv3Text.ErrDataDir + dataDir); }
            string[][] heads = new string[cfg.Data.Tables.Count][];
            for (int t = 0; t < cfg.Data.Tables.Count; t++)
            {
                string p = Path.Combine(dataDir, cfg.Data.Tables[t].File);
                if (!File.Exists(p)) { throw new Rdv3DataError(Rdv3Text.ErrNoData + p); }
                heads[t] = Rdv3Table.ReadHead(p, cfg.Data.Enc);
            }
            cfg.Data.Bind(heads);
        }
        catch (Exception ex)
        {
            Stop(logPath, "data", "not started: " + ex.Message,
                Rdv3Text.FatalDataTitle, Rdv3Text.FatalData.Replace("{reason}", ex.Message));
            return 3;
        }

        // One writer per ledger FILE -- so the guard has to be the file, not the
        // spelling. Everything downstream uses this same canonical path.
        try { ledgerPath = Path.GetFullPath(ledgerPath); }
        catch (Exception ex)
        {
            Stop(logPath, "ledger", "not started: bad ledger path " + ledgerPath + ": " + ex.Message,
                Rdv3Text.FatalTitle, Rdv3Text.ErrBadLedgerPath + ledgerPath + "\r\n" + ex.Message);
            return 5;
        }
        // a second instance must not fight the first over the xlsx, so it says
        // so and leaves
        string mutexName = "RdvApp-" + ledgerPath.ToLowerInvariant().GetHashCode().ToString("x8", CultureInfo.InvariantCulture);
        bool createdNew;
        using (Mutex mx = new Mutex(true, mutexName, out createdNew))
        {
            if (!createdNew)
            {
                MessageBox.Show(Rdv3Text.ErrAlreadyRunning,
                    Rdv3Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 4;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Rdv3Form form = new Rdv3Form(cfg.Screen);
            Rdv3App app = new Rdv3App(form, dataDir, ledgerPath, logPath, cfg);
            app.LogBoot(compileMs);
            Application.Run(form);
            return 0;
        }
    }

    private static string Resolve(string p, string baseDir)
    {
        if (Path.IsPathRooted(p) || baseDir == null || baseDir.Length == 0) { return p; }
        return Path.Combine(baseDir, p);
    }

    // the one line in the log and the one dialog, then the caller returns
    private static void Stop(string logPath, string section, string detail, string title, string body)
    {
        try { new Rdv3Log(logPath).Write("-", section, detail); } catch (Exception) { }
        MessageBox.Show(body, Rdv3Text.AppTitle + " - " + title, MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
