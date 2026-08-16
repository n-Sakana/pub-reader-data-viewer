// ============================================================================
// Rdv3App.cs -- entry point, state machine, timing boundaries, execution log.
//
// The two figures the screen shows, and exactly where they start and stop:
//
//   merge time   = readA+readB+readC + idxA+idxB+idxC + joinAB+joinBC, measured
//                  inside the background rebuild the startup check runs.
//                  Ledger composition, comparison, carry-over, persistence and
//                  the search index are NOT in it; they are logged separately.
//   search time  = search confirmed (button click / notepad detection) ->
//                  the single record is rendered, or the candidate list is
//                  built and rendered. A human choosing from the list is never
//                  in it; the post-pick render is logged as "display".
//
// States: BOOT -> CHECKING -> (confirm) -> APPLYING -> READY, or READY with
// the saved ledger when there is no difference or the operator declines, or
// BLOCKED when no ledger exists and none may be built. Stale worker results
// (superseded or timed out run IDs) are discarded and logged, never applied.
//
// The execution log is always on: one tab-separated line per event, next to
// the .cmd. The screen shows the two figures and the error line only.
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

    // everything that was a constant here now comes from ReaderDataViewer.json
    private readonly Rdv3Config cfg;

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
    private bool[] ledProcessed;
    private Rdv3Index ledIndex;

    // what the check produced, held between CHECKING and the decision
    private Rdv3MergeResult mergeResult;
    private string[] savedLines;
    private bool[] savedProcessed;

    private int runSeq;
    private int searchSeq;
    private int procSeq;
    private string activeRunId = "";
    private string activeSearchId = "";

    private int shownRow = -1;
    private string shownKey = "";
    private List<int> shownCands;
    private double lastMergeMs = -1;

    // ---- the exit guard around one unfinished "processed" save --------------
    // A "processed" mark is persisted immediately, one record at a time (there
    // is no queue and nothing is batched). The one moment its outcome is not
    // yet decided is while Rdv3Xlsx.Write runs on the worker thread, and a
    // close in that window would end the process without the operator ever
    // learning whether the record was saved. So during that window: no second
    // mark is accepted, and a close request is refused with the reason on
    // screen. Both flags are touched on the UI thread only.
    private bool savingMark;
    private bool closeAskedWhileSaving;

    // app construction instant, for the boot->operable startup figure
    private readonly long bootT0 = Rdv3Clock.Now();
    private bool startupLogged;

    public Rdv3App(Rdv3Form f, string data, string ledger, string logPath, Rdv3Config settings)
    {
        form = f;
        cfg = (settings == null) ? Rdv3Config.Defaults() : settings;
        dataDir = data;
        ledgerPath = ledger;
        pid = System.Diagnostics.Process.GetCurrentProcess().Id;
        log = new Rdv3Log(logPath);
        log.OnFail = delegate(string msg) { form.SetError("log: " + msg); };

        worker.OnError = JobFailed;

        form.OnSearch = ManualSearch;
        form.OnClear = DoClear;
        form.OnProcessed = DoProcessed;
        form.OnRebind = delegate { log.Write("-", "watch", "rebind requested"); watch.Rebind(); };
        form.OnPick = PickCandidate;

        watch.Cfg = cfg;
        watch.OnConfirmed = Detected;
        watch.OnState = WatchState;
        watch.OnLabel = delegate(string name) { form.SetWatchLabel(name); };
        form.OnSettings = OpenSettings;
        form.SetWatchLabel(WatchName());
        form.SetKeyRule(cfg.KeyLength);

        // the session block of the reference screen: who is signed in, on what
        // machine, and whether this session may write -- all real, none of it
        // the sample text from the artifact
        bool canWrite = true;
        try
        {
            string dir = System.IO.Path.GetDirectoryName(ledgerPath);
            if (System.IO.File.Exists(ledgerPath))
            {
                canWrite = (System.IO.File.GetAttributes(ledgerPath) & System.IO.FileAttributes.ReadOnly) == 0;
            }
            else if (dir != null && dir.Length > 0) { canWrite = System.IO.Directory.Exists(dir); }
        }
        catch (Exception) { canWrite = true; }
        form.SetIdentity(Environment.UserName, Environment.MachineName,
            canWrite ? Rdv3Text.RoleNormal : Rdv3Text.RoleReadOnly,
            "PID " + pid.ToString(CultureInfo.InvariantCulture),
            System.IO.Path.GetFileName(logPath));

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
                form.SetError(Rdv3Text.ErrCloseWhileSaving);
                log.Write("-", "exit", "close refused: a processed save is still in flight");
                form.Tell(Rdv3Text.CloseBlockedTitle, Rdv3Text.CloseBlockedBody);
                return;
            }
            Shutdown();
        };
    }

    public void LogBoot(double compileMs)
    {
        log.Write("-", "settings", cfg.Describe());
        for (int i = 0; i < cfg.Notes.Count; i++) { log.Write("-", "settings", cfg.Notes[i]); }
        log.Write("-", "boot", "pid=" + pid.ToString(CultureInfo.InvariantCulture)
            + " method=csharp-dict data=" + dataDir + " ledger=" + ledgerPath
            + " compile_ms=" + Rdv3Log.F(compileMs));
        log.Write("-", "worker", "kind=thread owner_pid=" + pid.ToString(CultureInfo.InvariantCulture));
    }

    // ---- startup check -----------------------------------------------------
    private void StartCheck()
    {
        state = StChecking;
        runSeq++;
        string rid = "R" + runSeq.ToString(CultureInfo.InvariantCulture);
        activeRunId = rid;
        form.SetState(Rdv3Text.StateChecking, 2);
        form.ShowOverlay(Rdv3Text.StateChecking);
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

    // worker thread
    private void CheckJob(string rid)
    {
        long t = Rdv3Clock.Now();
        Rdv3MergeResult mr = Rdv3Ledger.BuildFromCsv(dataDir);
        double composeMs = Rdv3Clock.MsSince(t) - mr.MergeMs();

        for (int i = 0; i < Rdv3Spec.StageCount; i++)
        {
            string sec = (i < 3) ? "read" : (i < 6) ? "index" : "join";
            string what = (i < 3) ? ("table=" + "ABC"[i])
                : (i < 6) ? ("table=" + "ABC"[i - 3] + " keys=" + PickKeys(mr, i - 3))
                : ((i == Rdv3Spec.StJoinAB)
                    ? ("pair=AB matched=" + mr.MatchedAB.ToString(CultureInfo.InvariantCulture))
                    : ("pair=BC matched=" + mr.MatchedBC.ToString(CultureInfo.InvariantCulture)
                       + " checksum=" + mr.Checksum.ToString(CultureInfo.InvariantCulture)));
            log.Write(rid, sec, what + " ms=" + Rdv3Log.F(mr.StageMs[i]));
        }
        log.Write(rid, "merge", "rows=" + mr.Rows.ToString(CultureInfo.InvariantCulture)
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
        bool[] oldProc = null;
        string loadError = null;
        bool exists = File.Exists(ledgerPath);
        if (exists)
        {
            try
            {
                t = Rdv3Clock.Now();
                Rdv3Xlsx.Read(ledgerPath, out oldLines, out oldProc);
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
        bool[] keepProc = oldProc;
        string keepErr = loadError;
        bool keepExists = exists;
        bool keepSame = same;
        bool keepOracle = oracleOk;
        form.RunOnUi(delegate { EndCheck(rid, mrKeep, keepOld, keepProc, keepExists, keepErr, keepSame, keepOracle); });
    }

    private static string PickKeys(Rdv3MergeResult mr, int i)
    {
        int k = (i == 0) ? mr.KeysA : (i == 1) ? mr.KeysB : mr.KeysC;
        return k.ToString(CultureInfo.InvariantCulture);
    }

    // UI thread
    private void EndCheck(string rid, Rdv3MergeResult mr, string[] oldLines, bool[] oldProc,
                          bool ledgerExists, string loadError, bool same, bool oracleOk)
    {
        if (!string.Equals(rid, activeRunId, StringComparison.Ordinal))
        {
            log.Write(rid, "stale", "check result discarded (active=" + activeRunId + ")");
            return;
        }
        mergeResult = mr;
        savedLines = oldLines;
        savedProcessed = oldProc;
        lastMergeMs = mr.MergeMs();
        form.SetMergeMs(lastMergeMs);
        if (!oracleOk) { form.SetError(Rdv3Text.ErrOracle); }

        if (oldLines == null)
        {
            // missing or unreadable: never silently rebuild -- ask, with the
            // reason on the screen
            string body = ledgerExists
                ? Rdv3Text.ConfirmRebuildBody.Replace("{err}", (loadError == null) ? "?" : loadError)
                : Rdv3Text.ConfirmCreateBody;
            if (ledgerExists) { form.SetError(Rdv3Text.ErrLedgerRead + ((loadError == null) ? "" : loadError)); }
            bool yes = form.Ask(Rdv3Text.ConfirmUpdateTitle, body);
            log.Write(rid, "decision", ledgerExists
                ? ("ledger unreadable; rebuild " + (yes ? "approved" : "declined"))
                : ("ledger missing; create " + (yes ? "approved" : "declined")));
            if (yes) { StartApply(rid, false); }
            else { EnterBlocked(); }
            return;
        }

        if (same)
        {
            log.Write(rid, "decision", "no difference");
            AdoptLedger(rid, oldLines, oldProc, Rdv3Text.NoteNoDiff);
            return;
        }

        bool approve = form.Ask(Rdv3Text.ConfirmUpdateTitle, Rdv3Text.ConfirmUpdateBody);
        log.Write(rid, "decision", "difference found; update " + (approve ? "approved" : "rejected"));
        if (approve)
        {
            StartApply(rid, true);
        }
        else
        {
            AdoptLedger(rid, oldLines, oldProc, Rdv3Text.NoteRejected);
        }
    }

    // adopt = make these lines the active, searchable ledger (worker builds the
    // index so the UI thread never runs over 100,000 rows)
    private void AdoptLedger(string rid, string[] lines, bool[] proc, string note)
    {
        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "adopt";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate
        {
            long t = Rdv3Clock.Now();
            Rdv3Index ix = new Rdv3Index(Rdv3Ledger.Key1Column(lines));
            log.Write(rid, "index", "table=LEDGER rows=" + lines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            ledLines = lines;
            ledProcessed = proc;
            ledIndex = ix;
            form.RunOnUi(delegate { EnterReady(rid, note); });
        };
        worker.Post(job);
    }

    private void StartApply(string rid, bool carry)
    {
        state = StApplying;
        form.SetState(Rdv3Text.StateApplying, 2);
        form.ShowOverlay(Rdv3Text.StateApplying);

        Rdv3MergeResult mr = mergeResult;
        string[] oldLines = savedLines;
        bool[] oldProc = savedProcessed;

        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "apply";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate
        {
            long t = Rdv3Clock.Now();
            bool[] np;
            if (carry && oldLines != null)
            {
                Rdv3Ledger.CarryStats st = new Rdv3Ledger.CarryStats();
                np = Rdv3Ledger.CarryProcessed(oldLines, oldProc, mr.Lines, st);
                log.Write(rid, "carry", "carried=" + st.Carried.ToString(CultureInfo.InvariantCulture)
                    + " reset=" + st.Reset.ToString(CultureInfo.InvariantCulture)
                    + " new=" + st.New.ToString(CultureInfo.InvariantCulture)
                    + " dropped=" + st.Dropped.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            }
            else
            {
                np = new bool[mr.Lines.Length];
                log.Write(rid, "carry", "none (fresh ledger)");
            }

            t = Rdv3Clock.Now();
            try
            {
                Rdv3Xlsx.Write(ledgerPath, mr.Lines, np,
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
            Rdv3Index ix = new Rdv3Index(Rdv3Ledger.Key1Column(mr.Lines));
            log.Write(rid, "index", "table=LEDGER rows=" + mr.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));

            ledLines = mr.Lines;
            ledProcessed = np;
            ledIndex = ix;
            form.RunOnUi(delegate { EnterReady(rid, Rdv3Text.NoteUpdated); });
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
        form.SetError(Rdv3Text.ErrPersist + msg);
        if (savedLines != null)
        {
            AdoptLedger(rid, savedLines, savedProcessed, Rdv3Text.NoteRejected);
        }
        else
        {
            EnterBlocked();
        }
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
        form.HideOverlay();
        form.EnableOps(true);
        form.SetState(watch.Bound ? Rdv3Text.StateReady
            : Rdv3Text.StateWaitingFmt.Replace("{name}", WatchName()), watch.Bound ? 1 : 2);
        // status bar: the file and its size; summary block: the same count and
        // the ledger's own last-write stamp (the reference's "最終更新")
        string rowsText = ledLines.Length.ToString("N0", CultureInfo.InvariantCulture);
        form.SetLedgerInfo(System.IO.Path.GetFileName(ledgerPath) + " ・ "
            + Rdv3Text.LedgerRows.Replace("{n}", rowsText));
        form.SetLedgerSummary(rowsText, LedgerStamp());
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

    // UI thread
    private void EnterBlocked()
    {
        state = StBlocked;
        activeRunId = "";
        form.HideOverlay();
        form.EnableOps(false);
        form.SetState(Rdv3Text.StateBlocked, 3);
        form.SetError(Rdv3Text.ErrNoLedger);
        log.Write("-", "decision", "blocked (no ledger)");
    }

    // ---- search ------------------------------------------------------------
    private void ManualSearch(string key)
    {
        long t0 = Rdv3Clock.Now();
        if (state != StReady)
        {
            form.SetError(Rdv3Text.ErrNotReady);
            log.Write("-", "search", "ignored key=" + key + " reason=not-ready");
            return;
        }
        if (!Rdv3Spec.IsKey(key))
        {
            // the rule this session is actually enforcing -- the same one IsKey
            // just applied, and the one the data was read with. cfg holds what
            // the FILE says, which may already be the next start's rule.
            form.SetInputError((Rdv3Spec.KeyDigitsOnly ? Rdv3Text.ErrBadKey : Rdv3Text.ErrBadKeyAny)
                .Replace("{n}", Rdv3Spec.KeyLength.ToString(CultureInfo.InvariantCulture)));
            log.Write("-", "search", "ignored key=" + key + " reason=bad-key");
            return;
        }
        form.SetError("");
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
        // where this search came from, in the one form every line below uses
        string from = "source=" + source + ((target.Length > 0) ? (" target=" + target) : "");
        string[] lines = ledLines;
        Rdv3Index ix = ledIndex;
        bool[] proc = ledProcessed;
        List<int> hits = ix.Find(key);
        int n = (hits == null) ? 0 : hits.Count;

        double elapsed;
        if (n == 1)
        {
            // one hit is still shown IN the list (auto-selected), the way the
            // reference does it -- the record panel fills from that selection
            int row = hits[0];
            List<Rdv3CandRow> one = new List<Rdv3CandRow>(1);
            one.Add(CandRow(lines[row], proc[row]));
            List<int> candRows = new List<int>(1);
            candRows.Add(row);
            form.RunOnUi(delegate
            {
                form.ShowCandidates(key, one, 1, 1);
                form.SelectCandidate(0, lines[row], proc[row], 1);
                shownRow = row;
                shownKey = key;
                shownCands = candRows;
            });
            elapsed = Rdv3Clock.MsSince(t0);
            log.Write(sid, "search", "key=" + key + " " + from + " hits=1 ms=" + Rdv3Log.F(elapsed));
        }
        else if (n == 0)
        {
            form.RunOnUi(delegate
            {
                form.ShowCandidates(key, new List<Rdv3CandRow>(), 0, 0);
                shownRow = -1;
                shownKey = key;
                shownCands = null;
            });
            elapsed = Rdv3Clock.MsSince(t0);
            log.Write(sid, "search", "key=" + key + " " + from + " hits=0 ms=" + Rdv3Log.F(elapsed));
        }
        else
        {
            int show = (n > cfg.CandidateRowsShown) ? cfg.CandidateRowsShown : n;
            List<Rdv3CandRow> rows = new List<Rdv3CandRow>(show);
            List<int> candRows = new List<int>(show);
            for (int i = 0; i < show; i++)
            {
                rows.Add(CandRow(lines[hits[i]], proc[hits[i]]));
                candRows.Add(hits[i]);
            }
            form.RunOnUi(delegate
            {
                form.ShowCandidates(key, rows, n, show);
                shownRow = -1;
                shownKey = key;
                shownCands = candRows;
            });
            elapsed = Rdv3Clock.MsSince(t0);
            log.Write(sid, "search", "key=" + key + " " + from
                + " hits=" + n.ToString(CultureInfo.InvariantCulture) + " ms=" + Rdv3Log.F(elapsed));
            log.Write(sid, "candidate", "count=" + n.ToString(CultureInfo.InvariantCulture)
                + " shown=" + show.ToString(CultureInfo.InvariantCulture));
        }
        form.SetSearchMs(elapsed);
    }

    // one ledger line as the candidate table needs it
    private static Rdv3CandRow CandRow(string line, bool processed)
    {
        Rdv3CandRow r = new Rdv3CandRow();
        r.Cols = Rdv3Ledger.CandidateColumns(line);
        r.Processed = processed;
        r.Line = line;
        return r;
    }

    // UI thread. The pick happens after the search clock stopped; its render is
    // the separate "display" figure in the log.
    private void PickCandidate(int i)
    {
        List<int> cands = shownCands;
        if (state != StReady || cands == null || i < 0 || i >= cands.Count) { return; }
        long t0 = Rdv3Clock.Now();
        int row = cands[i];
        string k2 = Rdv3Ledger.FieldOf(ledLines[row], 1);
        form.SelectCandidate(i, ledLines[row], ledProcessed[row], cands.Count);
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

    // ---- processed ---------------------------------------------------------
    private void DoProcessed()
    {
        if (state != StReady) { form.SetError(Rdv3Text.ErrNotReady); return; }
        // one save at a time: a second mark while the first is unresolved would
        // rewrite the same file from under it
        if (savingMark)
        {
            form.SetError(Rdv3Text.ErrSaveInFlight);
            log.Write("-", "processed", "refused: a processed save is still in flight");
            return;
        }
        int row = shownRow;
        if (row < 0)
        {
            form.SetError(Rdv3Text.ErrNoRecordShown);
            return;
        }
        string k2 = Rdv3Ledger.FieldOf(ledLines[row], 1);
        if (!form.Ask(Rdv3Text.ConfirmProcessedTitle, Rdv3Text.ConfirmProcessedBody.Replace("{key2}", k2)))
        {
            log.Write("-", "processed", "declined key2=" + k2);
            return;
        }
        form.SetError("");
        procSeq++;
        string pidTag = "P" + procSeq.ToString(CultureInfo.InvariantCulture);
        long t0 = Rdv3Clock.Now();          // confirm instant: the E2E clock

        // the save starts now and its outcome is undecided until the worker
        // reports back: say so on screen, refuse a second mark, refuse a close
        savingMark = true;
        closeAskedWhileSaving = false;
        form.EnableProcessed(false);
        form.SetState(Rdv3Text.StateSavingMark, 2);
        form.ShowProcessedState(Rdv3Text.LabelProcessed + ": " + Rdv3Ledger.ProcessedTrue + Rdv3Text.SavingSuffix);
        form.SetCandidateProcessed(CandOf(row), true);
        log.Write(pidTag, "processed", "save started key2=" + k2 + " (exit held until it is decided)");

        Rdv3Job job = new Rdv3Job();
        job.RunId = pidTag;
        job.Kind = "processed";
        job.TimeoutMs = cfg.SaveTimeoutMs;
        job.Work = delegate
        {
            bool was = ledProcessed[row];
            ledProcessed[row] = true;
            long t = Rdv3Clock.Now();
            try
            {
                Rdv3Xlsx.Write(ledgerPath, ledLines, ledProcessed,
                    pid.ToString(CultureInfo.InvariantCulture) + "-" + pidTag);
            }
            catch (Exception ex)
            {
                // the file could not change, so the memory must not either:
                // screen and ledger never diverge silently
                ledProcessed[row] = was;
                log.Write(pidTag, "error", "stage=persist key2=" + k2 + " msg=" + ex.Message);
                form.RunOnUi(delegate
                {
                    form.SetError(Rdv3Text.ErrPersist + ex.Message);
                    // the LIST row belongs to the record that was saved, whether
                    // or not it is still the one selected: put it back to what
                    // the ledger actually holds
                    form.SetCandidateProcessed(CandOf(row), was);
                    if (shownRow == row)
                    {
                        // what the record went back to, which is what it was
                        // before the click -- not FALSE. Marking a row that was
                        // already TRUE and failing to save must not leave the
                        // screen claiming the opposite of the ledger.
                        form.ShowProcessedState(Rdv3Text.LabelProcessed + ": "
                            + (was ? Rdv3Ledger.ProcessedTrue : Rdv3Ledger.ProcessedFalse));
                    }
                    EndMarkSave(pidTag, false);
                });
                return;
            }
            double ms = Rdv3Clock.MsSince(t);
            form.RunOnUi(delegate
            {
                form.SetCandidateProcessed(CandOf(row), true);
                if (shownRow == row)
                {
                    form.ShowProcessedState(Rdv3Text.LabelProcessed + ": " + Rdv3Ledger.ProcessedTrue);
                }
                EndMarkSave(pidTag, true);
            });
            // e2e = confirm click -> file persisted AND the screen updated
            log.Write(pidTag, "processed", "key2=" + k2
                + " row=" + (row + 1).ToString(CultureInfo.InvariantCulture)
                + " value=TRUE persist_ms=" + Rdv3Log.F(ms)
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
        form.EnableProcessed(state == StReady);
        if (state == StReady)
        {
            int tone;
            string t = WatchStateText(out tone);
            form.SetState(t, tone);
        }
        log.Write(tag, "exit", "processed save decided (" + (ok ? "saved" : "failed") + "); exit released");
        if (closeAskedWhileSaving)
        {
            closeAskedWhileSaving = false;
            // a save that WORKED is not an error. The failed one still is: it
            // says the record did not reach the file.
            if (ok) { form.SetNotice(Rdv3Text.NoteSaveDoneCanClose); }
            else { form.SetError(Rdv3Text.NoteSaveFailedCanClose); }
        }
    }

    private void DoClear()
    {
        form.ClearResult();
        form.SetError("");
        shownRow = -1;
        shownKey = "";
        shownCands = null;
        log.Write("-", "clear", "input and result cleared");
    }

    // ---- watch / timeout / shutdown ---------------------------------------
    // The settings dialog edits a copy; when it comes back, the file is written
    // and the running session adopts what it can without a restart.
    private void OpenSettings()
    {
        Rdv3Config edited = Rdv3SettingsForm.Edit(form, cfg);
        if (edited == null) { return; }
        string err = edited.Save(cfg.SourcePath);
        if (err != null)
        {
            form.SetError(Rdv3Text.ErrSettingsSave + err);
            log.Write("-", "settings", "save failed: " + err);
            return;
        }
        // the file is written, so cfg carries all of it now: the members that
        // take effect at once, and the ones that wait for the next start but
        // still have to survive into the next time this dialog opens
        cfg.AdoptRuntimeFrom(edited);
        cfg.AdoptSavedFrom(edited);
        // digitsOnly only decides which characters make a key -- nothing read at
        // start-up depends on it, so it applies now, as it always has. The key
        // LENGTH is the one that has to wait (Rdv3Spec.KeyLength is the width the
        // CSVs were read and the index was built with), which is also why
        // SetKeyRule is not called here.
        Rdv3Spec.KeyDigitsOnly = cfg.KeyDigitsOnly;
        form.SetWatchLabel(WatchName());
        form.SetInputError("");
        watch.Rebind();
        log.Write("-", "settings", "saved: " + cfg.Describe());
        // the key length in force is still the one this session started on; say
        // so rather than let the log imply the file took over
        if (cfg.KeyLength != Rdv3Spec.KeyLength)
        {
            log.Write("-", "settings", "key.length "
                + cfg.KeyLength.ToString(CultureInfo.InvariantCulture)
                + " applies at the next start; this session keeps "
                + Rdv3Spec.KeyLength.ToString(CultureInfo.InvariantCulture));
        }
        for (int i = 0; i < cfg.Targets.Count; i++)
        {
            string why = cfg.Targets[i].WhyNotWatchable();
            if (why.Length > 0)
            {
                log.Write("-", "settings", "target [" + cfg.Targets[i].Name + "] is not watched: " + why);
            }
        }
        form.SetNotice(Rdv3Text.NoteSettingsApplied);
    }

    // What the screen calls the thing being watched -- counted over the targets
    // that are actually watched, not over every target in the file. cfg.Targets
    // also holds the ones the operator switched off (the dialog lists them and
    // the file keeps them); naming the general word because ONE of two targets is
    // off would drop the name the operator recognises from the status line.
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

    // how many targets this configuration actually asks to be watched
    private int WatchableCount()
    {
        int n = 0;
        for (int i = 0; i < cfg.Targets.Count; i++) { if (cfg.Targets[i].IsWatchable) { n++; } }
        return n;
    }

    // WHAT THE STATE LINE SAYS when nothing else is going on. One definition,
    // because there are two callers -- the watcher's own notifications and the
    // end of a processed save -- and the second one used to build its own and
    // got "nothing is configured to be watched" wrong, calling it メモ帳待機.
    private string WatchStateText(out int tone)
    {
        if (WatchableCount() == 0) { tone = 2; return Rdv3Text.StateNoTarget; }
        if (watch != null && watch.Bound) { tone = 1; return Rdv3Text.StateReady; }
        tone = 2;
        return Rdv3Text.StateWaitingFmt.Replace("{name}", WatchName());
    }

    private void WatchState(string st, string detail)
    {
        // while a save is in flight the state line says so; the watch line is
        // still updated, but it does not overwrite that
        if (st == "WATCHING")
        {
            form.SetNotepad(detail);
        }
        else if (st == "WAITING")
        {
            // nothing to watch is a decision the file made, not a window that
            // has not opened yet: say which of the two this is
            form.SetNotepad(WatchableCount() == 0
                ? Rdv3Text.WatchNoTarget
                : Rdv3Text.WatchNoneFmt.Replace("{name}", WatchName()));
        }
        if ((st == "WATCHING" || st == "WAITING") && state == StReady && !savingMark)
        {
            int tone;
            string t = WatchStateText(out tone);
            form.SetState(t, tone);
        }
    }

    private void CheckOverdue()
    {
        Rdv3Job j = worker.TakeOverdue();
        if (j == null) { return; }
        log.Write(j.RunId, "timeout", "kind=" + j.Kind + " after_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(j.StartedAt)));
        if (string.Equals(j.Kind, "processed", StringComparison.Ordinal))
        {
            // a managed job cannot be aborted, and until the write returns
            // nobody knows whether the record reached the file. Report the
            // delay; keep holding the exit rather than claim a decision.
            form.SetError(Rdv3Text.ErrSaveOverdue);
            return;
        }
        if (string.Equals(j.RunId, activeRunId, StringComparison.Ordinal))
        {
            // abandon the run: whatever arrives later is stale by ID
            activeRunId = "";
            form.SetError(Rdv3Text.ErrCheckTimeout);
            if (ledLines != null)
            {
                state = StReady;
                form.HideOverlay();
                form.EnableOps(true);
                form.SetState(Rdv3Text.StateReady, 1);
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
        form.RunOnUi(delegate
        {
            form.SetError(Rdv3Text.ErrCheckFailed + ex.Message);
            // a mark job that threw anywhere is still a decided save (failed):
            // the guard must never outlive the job that armed it
            if (string.Equals(job.Kind, "processed", StringComparison.Ordinal)) { EndMarkSave(job.RunId, false); }
            if (!string.Equals(job.RunId, activeRunId, StringComparison.Ordinal)) { return; }
            activeRunId = "";
            if (ledLines != null)
            {
                state = StReady;
                form.HideOverlay();
                form.EnableOps(true);
                form.SetState(Rdv3Text.StateReady, 1);
            }
            else if (savedLines != null && job.Kind != "adopt")
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
        AdoptLedger(rid, savedLines, savedProcessed, Rdv3Text.NoteRejected);
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
    public static int Run(string dataDir, string ledgerPath, string logPath, string configPath, double compileMs)
    {
        if (IntPtr.Size != 8)
        {
            MessageBox.Show(Rdv3Text.ErrNo64, Rdv3Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }
        string[] need = { "tableA.csv", "tableB.csv", "tableC.csv" };
        for (int i = 0; i < need.Length; i++)
        {
            string p = Path.Combine(dataDir, need[i]);
            if (!File.Exists(p))
            {
                MessageBox.Show(Rdv3Text.ErrNoData + p, Rdv3Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 3;
            }
        }

        // One writer per ledger FILE -- so the guard has to be the file, not the
        // spelling. "x\led.xlsx", "x\.\led.xlsx" and a relative path all name the
        // same xlsx and used to hash three different ways, which let two
        // processes each believe it was the only one and write over each other.
        // Everything downstream uses this same canonical path.
        try { ledgerPath = Path.GetFullPath(ledgerPath); }
        catch (Exception ex)
        {
            MessageBox.Show(Rdv3Text.ErrBadLedgerPath + ledgerPath + "\r\n" + ex.Message,
                Rdv3Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            // the settings come first: the key rule shapes the screen itself
            Rdv3Config cfg = Rdv3Config.Load(configPath);
            Rdv3Spec.KeyLength = cfg.KeyLength;
            Rdv3Spec.KeyDigitsOnly = cfg.KeyDigitsOnly;

            Rdv3Form form = new Rdv3Form();
            Rdv3App app = new Rdv3App(form, dataDir, ledgerPath, logPath, cfg);
            if (cfg.Error.Length > 0) { form.SetError(Rdv3Text.ErrSettingsRead + cfg.Error); }
            app.LogBoot(compileMs);
            Application.Run(form);
            return 0;
        }
    }
}
