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

    private const int CandShowMax = 10;      // candidate rows on screen; more is said, never hidden
    private const double CheckTimeoutMs = 180000.0;

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

    // app construction instant, for the boot->operable startup figure
    private readonly long bootT0 = Rdv3Clock.Now();
    private bool startupLogged;

    public Rdv3App(Rdv3Form f, string data, string ledger, string logPath)
    {
        form = f;
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

        watch.OnConfirmed = Detected;
        watch.OnState = WatchState;

        watchdog.Interval = 1000;
        watchdog.Tick += delegate(object s, EventArgs e) { CheckOverdue(); };

        form.Shown += delegate(object s, EventArgs e)
        {
            watchdog.Start();
            StartCheck();
        };
        form.FormClosing += delegate(object s, FormClosingEventArgs e) { Shutdown(); };
    }

    public void LogBoot(double compileMs)
    {
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
        job.TimeoutMs = CheckTimeoutMs;
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
        job.TimeoutMs = CheckTimeoutMs;
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
        job.TimeoutMs = CheckTimeoutMs;
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
        form.SetState(watch.Bound ? Rdv3Text.StateReady : Rdv3Text.StateWaitingNotepad, watch.Bound ? 1 : 2);
        form.SetLedgerInfo(Rdv3Text.LedgerRows.Replace("{n}",
            ledLines.Length.ToString("N0", CultureInfo.InvariantCulture)) + "   " + note);
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
            form.SetError(Rdv3Text.ErrBadKey);
            log.Write("-", "search", "ignored key=" + key + " reason=bad-key");
            return;
        }
        form.SetError("");
        Search(key, "manual", 0, t0);
    }

    // watch thread
    private void Detected(string key, double detectMs, int polls, long t0)
    {
        log.Write("-", "detect", "key=" + key + " latency_ms=" + Rdv3Log.F(detectMs)
            + " polls=" + polls.ToString(CultureInfo.InvariantCulture));
        if (state != StReady)
        {
            log.Write("-", "search", "ignored key=" + key + " reason=not-ready (detected)");
            return;
        }
        Search(key, "detect", detectMs, t0);
    }

    // any thread. t0 is the confirm instant: the search clock is already running.
    private void Search(string key, string source, double detectMs, long t0)
    {
        searchSeq++;
        string sid = "S" + searchSeq.ToString(CultureInfo.InvariantCulture);
        activeSearchId = sid;

        Rdv3Job job = new Rdv3Job();
        job.RunId = sid;
        job.Kind = "search";
        job.TimeoutMs = 30000;
        job.Work = delegate { SearchJob(sid, key, source, t0); };
        worker.Post(job);
    }

    // worker thread
    private void SearchJob(string sid, string key, string source, long t0)
    {
        if (!string.Equals(sid, activeSearchId, StringComparison.Ordinal))
        {
            log.Write(sid, "stale", "search superseded before start key=" + key);
            return;
        }
        string[] lines = ledLines;
        Rdv3Index ix = ledIndex;
        bool[] proc = ledProcessed;
        List<int> hits = ix.Find(key);
        int n = (hits == null) ? 0 : hits.Count;

        double elapsed;
        if (n == 1)
        {
            int row = hits[0];
            string[] va, vb, vc;
            Rdv3Ledger.RecordView(lines[row], out va, out vb, out vc);
            string k2 = Rdv3Ledger.FieldOf(lines[row], 1);
            bool p = proc[row];
            form.RunOnUi(delegate
            {
                form.ShowRecord(key, Rdv3Text.VerdictOne.Replace("{key2}", k2), 1, va, vb, vc,
                    Rdv3Text.LabelProcessed + ": " + (p ? Rdv3Ledger.ProcessedTrue : Rdv3Ledger.ProcessedFalse));
                shownRow = row;
                shownKey = key;
                shownCands = null;
            });
            elapsed = Rdv3Clock.MsSince(t0);
            log.Write(sid, "search", "key=" + key + " source=" + source + " hits=1 ms=" + Rdv3Log.F(elapsed));
        }
        else if (n == 0)
        {
            form.RunOnUi(delegate
            {
                form.ShowRecord(key, Rdv3Text.VerdictNone, 3, null, null, null, "");
                shownRow = -1;
                shownKey = key;
                shownCands = null;
            });
            elapsed = Rdv3Clock.MsSince(t0);
            log.Write(sid, "search", "key=" + key + " source=" + source + " hits=0 ms=" + Rdv3Log.F(elapsed));
        }
        else
        {
            int show = (n > CandShowMax) ? CandShowMax : n;
            string[][] rows = new string[show][];
            List<int> candRows = new List<int>(show);
            for (int i = 0; i < show; i++)
            {
                rows[i] = Rdv3Ledger.CandidateColumns(lines[hits[i]]);
                candRows.Add(hits[i]);
            }
            string verdict = (show < n)
                ? Rdv3Text.VerdictManyCut.Replace("{n}", n.ToString(CultureInfo.InvariantCulture))
                                         .Replace("{m}", show.ToString(CultureInfo.InvariantCulture))
                : Rdv3Text.VerdictMany.Replace("{n}", n.ToString(CultureInfo.InvariantCulture));
            form.RunOnUi(delegate
            {
                form.ShowCandidates(key, verdict, rows);
                shownRow = -1;
                shownKey = key;
                shownCands = candRows;
            });
            elapsed = Rdv3Clock.MsSince(t0);
            log.Write(sid, "search", "key=" + key + " source=" + source
                + " hits=" + n.ToString(CultureInfo.InvariantCulture) + " ms=" + Rdv3Log.F(elapsed));
            log.Write(sid, "candidate", "count=" + n.ToString(CultureInfo.InvariantCulture)
                + " shown=" + show.ToString(CultureInfo.InvariantCulture));
        }
        form.SetSearchMs(elapsed);
    }

    // UI thread. The pick happens after the search clock stopped; its render is
    // the separate "display" figure in the log.
    private void PickCandidate(int i)
    {
        List<int> cands = shownCands;
        if (state != StReady || cands == null || i < 0 || i >= cands.Count) { return; }
        long t0 = Rdv3Clock.Now();
        int row = cands[i];
        string[] va, vb, vc;
        Rdv3Ledger.RecordView(ledLines[row], out va, out vb, out vc);
        string k2 = Rdv3Ledger.FieldOf(ledLines[row], 1);
        string verdict = Rdv3Text.VerdictPicked.Replace("{key2}", k2)
            .Replace("{n}", cands.Count.ToString(CultureInfo.InvariantCulture))
            .Replace("{i}", (i + 1).ToString(CultureInfo.InvariantCulture));
        form.ShowRecord(shownKey, verdict, 1, va, vb, vc,
            Rdv3Text.LabelProcessed + ": " + (ledProcessed[row] ? Rdv3Ledger.ProcessedTrue : Rdv3Ledger.ProcessedFalse));
        shownRow = row;
        log.Write("S" + searchSeq.ToString(CultureInfo.InvariantCulture), "display",
            "picked=" + (i + 1).ToString(CultureInfo.InvariantCulture)
            + " key2=" + k2 + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t0)));
    }

    // ---- processed ---------------------------------------------------------
    private void DoProcessed()
    {
        if (state != StReady) { form.SetError(Rdv3Text.ErrNotReady); return; }
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

        Rdv3Job job = new Rdv3Job();
        job.RunId = pidTag;
        job.Kind = "processed";
        job.TimeoutMs = 60000;
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
                form.RunOnUi(delegate { form.SetError(Rdv3Text.ErrPersist + ex.Message); });
                return;
            }
            double ms = Rdv3Clock.MsSince(t);
            form.RunOnUi(delegate
            {
                if (shownRow == row)
                {
                    form.ShowProcessedState(Rdv3Text.LabelProcessed + ": " + Rdv3Ledger.ProcessedTrue);
                }
            });
            // e2e = confirm click -> file persisted AND the screen updated
            log.Write(pidTag, "processed", "key2=" + k2
                + " row=" + (row + 1).ToString(CultureInfo.InvariantCulture)
                + " value=TRUE persist_ms=" + Rdv3Log.F(ms)
                + " e2e_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t0)));
        };
        worker.Post(job);
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
    private void WatchState(string st, string detail)
    {
        if (st == "WATCHING")
        {
            form.SetNotepad(detail);
            if (state == StReady) { form.SetState(Rdv3Text.StateReady, 1); }
        }
        else if (st == "WAITING")
        {
            form.SetNotepad(Rdv3Text.NotepadNone);
            if (state == StReady) { form.SetState(Rdv3Text.StateWaitingNotepad, 2); }
        }
    }

    private void CheckOverdue()
    {
        Rdv3Job j = worker.TakeOverdue();
        if (j == null) { return; }
        log.Write(j.RunId, "timeout", "kind=" + j.Kind + " after_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(j.StartedAt)));
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
    public static int Run(string dataDir, string ledgerPath, string logPath, double compileMs)
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

        // one writer per ledger file: a second instance must not fight the
        // first over the xlsx, so it says so and leaves
        string mutexName = "RdvApp-" + ledgerPath.ToLowerInvariant().GetHashCode().ToString("x8", CultureInfo.InvariantCulture);
        bool createdNew;
        using (Mutex mx = new Mutex(true, mutexName, out createdNew))
        {
            if (!createdNew)
            {
                MessageBox.Show(Rdv3Text.AppTitle + " is already running for this ledger.",
                    Rdv3Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 4;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Rdv3Form form = new Rdv3Form();
            Rdv3App app = new Rdv3App(form, dataDir, ledgerPath, logPath);
            app.LogBoot(compileMs);
            Application.Run(form);
            return 0;
        }
    }
}
