// ============================================================================
// Rdv3App.cs -- state machine and timing boundaries.
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
// definition says it is. A transition is saved to the small local pending
// store first. Only an explicit send applies those pending values to a freshly
// read shared ledger while holding the ledger's shared-file lease.
//
// The execution log is always on: one tab-separated line per event, next to
// the .cmd. The screen shows notices in its status bar and errors in warning dialogs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Threading;

public sealed class Rdv3App
{
    private const int StBoot = 0;
    private const int StChecking = 1;
    private const int StApplying = 2;
    private const int StReady = 3;
    private const int StBlocked = 4;
    private const int StReloading = 5;
    private const int StSending = 6;

    private readonly Rdv3Config cfg;
    private readonly Rdv3Screen screen;
    private readonly Rdv3WorkState work;
    private readonly Rdv3Data dataDef;

    private readonly Rdv3Form form;
    private readonly Rdv3Worker worker = new Rdv3Worker();
    private readonly Rdv3Watch watch = new Rdv3Watch();
    private readonly Rdv3Log log;
    private readonly Rdv3PendingStore pending;
    private readonly Rdv3SharedFiles shared;
    private readonly DispatcherTimer watchdog = new DispatcherTimer();

    private readonly string appDir;
    private readonly string dataDir;
    private readonly string ledgerPath;
    private readonly string storageContract;
    private readonly int pid;

    private int state = StBoot;
    private bool watchStarted;

    // the active ledger. Mutated only inside worker jobs (the queue serialises
    // them); the UI receives rendered copies and never touches these arrays.
    private string[] ledLines;
    private string[] ledStates;
    private string[] sharedStates;
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
    private string activeSearchKey = "";

    private int shownRow = -1;
    private string shownKey = "";
    private List<int> shownCands;
    private double lastMergeMs = -1;

    // ---- the exit guard around one unfinished local/shared write ------------
    // The small pending-file write and every shared-ledger write have a clear
    // success/failure outcome. A close is refused only while that outcome is
    // undecided. The wait flag crosses the UI/worker boundary under its gate.
    private readonly Rdv3WriteGuard writes = new Rdv3WriteGuard();
    private readonly Rdv3LedgerStore ledgerStore;
    private readonly Rdv3ResetNotice resetNotice;

    private Rdv3SharedMarker seenMarker;
    private Rdv3SharedMarker deferredMarker;
    private long lastMarkerPoll;
    private bool markerPollQueued;
    private string ledgerObservedStamp;
    private bool sharedReloading;
    private bool markerReadFailed;
    private DateTime reloadRetryAfter = DateTime.MinValue;

    // app construction instant, for the boot->operable startup figure
    private readonly long bootT0 = Rdv3Clock.Now();
    private bool startupLogged;
    private bool headlessSearchStarted;
    private bool headlessCaptureDone;
    private bool headlessModalStarted;

    public Rdv3App(Rdv3Form f, string appBaseDir, string data, string ledger, string logPath, Rdv3Config settings,
                   Rdv3PendingStore pendingChanges, Rdv3SharedFiles sharedFiles, Rdv3SharedMarker initialMarker)
    {
        form = f;
        cfg = settings;
        screen = settings.Screen;
        work = screen.Work;
        dataDef = settings.Data;
        storageContract = Rdv3Files.StorageContract(dataDef, work);
        appDir = appBaseDir;
        dataDir = data;
        ledgerPath = ledger;
        pid = System.Diagnostics.Process.GetCurrentProcess().Id;
        log = new Rdv3Log(logPath);
        pending = pendingChanges;
        shared = sharedFiles;
        ledgerStore = new Rdv3LedgerStore(ledgerPath, dataDef, work, shared);
        resetNotice = new Rdv3ResetNotice(dataDef.IdentityCols, work.InitialStored);
        seenMarker = initialMarker;
        try { ledgerObservedStamp = Rdv3Files.Stamp(ledgerPath); } catch (Exception) { ledgerObservedStamp = null; }
        log.OnFail = delegate(string msg) { form.Error(Rdv3Text.ErrLogWrite + msg); };
        // the ledger columns the screen binds to, fixed by the definition
        form.SetFields(new Rdv3Fields(dataDef.ColumnRefs));

        worker.OnError = JobFailed;

        form.OnSearch = ManualSearch;
        form.OnKeyChanged = delegate(string value)
        {
            if (!string.Equals(Rdv3Input.Cell((value ?? "").Trim()), activeSearchKey, StringComparison.Ordinal)) { ClearShown(); }
        };
        form.OnClear = DoClear;
        form.OnWorkState = DoWorkState;
        form.OnRefreshLedger = RefreshLedger;
        form.OnTableExport = OpenTableExport;
        form.OnUpdateRecords = OpenUpdateJob;
        form.OnDeleteRecords = OpenDeleteJob;
        form.OnSendChanges = SendChanges;
        form.OnSettings = OpenSettings;

        watch.Cfg = cfg;
        watch.OnConfirmed = Detected;
        watch.OnState = WatchState;
        watch.OnLabel = delegate(string name) { form.SetWatch(name, watchDetail); };
        form.SetWatch(WatchName(), Rdv3Text.WatchNone);
        form.SetIdentity("PID " + pid.ToString(CultureInfo.InvariantCulture), System.IO.Path.GetFileName(logPath));
        form.SetState(Rdv3Text.StateBoot);
        form.SetLedger(Rdv3Text.NotYet, "", "");

        watchdog.Interval = TimeSpan.FromMilliseconds(cfg.PumpMs);
        watchdog.Tick += delegate(object s, EventArgs e) { Tick(); };

        form.Shown += delegate(object s, EventArgs e)
        {
            log.Write("-", "screen", form.Diag);
            log.Write("-", "pending", "path=" + pending.Path + " count=" + pending.Count.ToString(CultureInfo.InvariantCulture));
            log.Write("-", "shared", "lock=" + shared.LockPath + " marker=" + shared.MarkerPath);
            watchdog.Start();
            StartCheck();
        };
        form.FormClosing += delegate(object s, ReaderDataViewer.Rdv3FormClosingEventArgs e)
        {
            if (!writes.TryClose())
            {
                e.Cancel = true;
                form.Error(Rdv3Text.ErrCloseWhileWriting);
                log.Write("-", "exit", "close refused: a write is still in flight");
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
            + " data=" + dataDir + " ledger=" + ledgerPath
            + " compile_ms=" + Rdv3Log.F(compileMs));
        log.Write("-", "worker", "kind=thread owner_pid=" + pid.ToString(CultureInfo.InvariantCulture));
    }

    // ---- the update check (at start-up, and from the refreshLedger button) ----
    private void StartCheck()
    {
        StartCheck(dataDef.UpdateJob);
    }

    private void StartCheck(Rdv3ProcessJobDef process)
    {
        ClearShown();
        state = StChecking;
        runSeq++;
        string rid = "R" + runSeq.ToString(CultureInfo.InvariantCulture);
        activeRunId = rid;
        form.SetState(Rdv3Text.StateChecking);
        form.EnableOps(false);
        log.Write(rid, "decision", "check started job=" + process.Id);

        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "check";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate { CheckJob(rid, process); };
        worker.Start();
        worker.Post(job);
    }

    private void RefreshLedger()
    {
        RefreshLedger(dataDef.UpdateJob);
    }

    private void RefreshLedger(Rdv3ProcessJobDef process)
    {
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        if (writes.Pending) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        log.Write("-", "decision", "check requested from the screen job=" + process.Id);
        StartCheck(process);
    }

    // worker thread
    private void CheckJob(string rid, Rdv3ProcessJobDef process)
    {
        long t = Rdv3Clock.Now();
        Rdv3MergeResult mr = Rdv3Ledger.BuildFromCsv(dataDef, process, dataDir);
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
        for (int i = 0; i < process.Joins.Count; i++)
        {
            log.Write(rid, "join", "pair=" + process.Spine + process.Joins[i].Table + " on=" + process.Joins[i].On
                + " matched=" + mr.Matched[i].ToString(CultureInfo.InvariantCulture) + " ms=" + Rdv3Log.F(mr.JoinMs[i]));
        }
        log.Write(rid, "merge", "rows=" + mr.Rows.ToString(CultureInfo.InvariantCulture)
            + " checksum=" + mr.Checksum.ToString(CultureInfo.InvariantCulture)
            + " ms=" + Rdv3Log.F(mr.MergeMs()) + " compose_ms=" + Rdv3Log.F(composeMs));
        for (int i = 0; i < mr.Warnings.Count; i++) { log.Write(rid, "warning", mr.Warnings[i]); }

        // saved ledger
        string[] oldLines = null;
        string[] oldStates = null;
        string loadError = null;
        bool exists = Rdv3Files.Exists(ledgerPath);
        if (exists)
        {
            try
            {
                t = Rdv3Clock.Now();
                ReadLedger(mr.Head, out oldLines, out oldStates);
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
            Rdv3UpdateResult preview = (mr.Prepared == null)
                ? Rdv3Ledger.ApplyUpdate(process, oldLines, oldStates,
                    mr.Lines, dataDef.IdentityCols, work.InitialStored)
                : Rdv3Process.Execute(mr.Prepared, oldLines, oldStates, work.InitialStored, false).Update;
            int firstDiff;
            Rdv3Ledger.SameContent(oldLines, preview.Lines, out firstDiff);
            same = Rdv3Ledger.SameLedger(oldLines, oldStates, preview.Lines, preview.States);
            log.Write(rid, "compare", "rows_old=" + oldLines.Length.ToString(CultureInfo.InvariantCulture)
                + " rows_source=" + mr.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " rows_result=" + preview.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " same=" + (same ? "true" : "false")
                + (same ? "" : " first_diff_row=" + (firstDiff + 1).ToString(CultureInfo.InvariantCulture))
                + " added=" + preview.Added.ToString(CultureInfo.InvariantCulture)
                + " updated=" + preview.Updated.ToString(CultureInfo.InvariantCulture)
                + " kept=" + preview.Kept.ToString(CultureInfo.InvariantCulture)
                + " deleted=" + preview.Deleted.ToString(CultureInfo.InvariantCulture)
                + " reset=" + preview.ResetLines.Count.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
        }

        Rdv3MergeResult mrKeep = mr;
        string[] keepOld = oldLines;
        string[] keepStates = oldStates;
        string keepErr = loadError;
        bool keepExists = exists;
        bool keepSame = same;
        form.RunOnUi(delegate { EndCheck(rid, mrKeep, keepOld, keepStates, keepExists, keepErr, keepSame); });
    }

    // UI thread
    private void EndCheck(string rid, Rdv3MergeResult mr, string[] oldLines, string[] oldStates,
                          bool ledgerExists, string loadError, bool same)
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
        if (mr.Warnings.Count > 0) { form.Error(string.Join(Environment.NewLine, mr.Warnings.ToArray())); }

        if (oldLines == null)
        {
            if (ledgerExists)
            {
                // A read error is not permission to destroy the unreadable file.
                form.Error(Rdv3Text.ErrLedgerRead + (loadError ?? ""));
                if (ledLines != null) { ContinueWithActive(rid, "ledger-read-failed"); }
                else { EnterBlocked(); }
                return;
            }
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
            if (yes) { StartApply(rid); }
            else if (ledLines != null) { ContinueWithActive(rid, Rdv3Text.NoteRejected); }
            else { EnterBlocked(); }
            return;
        }

        if (same)
        {
            log.Write(rid, "decision", "no difference");
            // No source-content change does NOT mean the shared states are old.
            AdoptLedger(rid, oldLines, oldStates, mr.Head, Rdv3Text.NoteNoDiff);
            return;
        }

        bool approve = form.Ask(Rdv3Text.ConfirmUpdateTitle,
            Rdv3Text.UpdateConfirmBody(mr.Job.OnSourceChange, work.InitialState.Text));
        log.Write(rid, "decision", "difference found; update " + (approve ? "approved" : "rejected"));
        if (approve)
        {
            StartApply(rid);
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

    private Rdv3Index BuildSearchIndex(string[] lines)
    {
        return new Rdv3Index(lines, dataDef.SearchCols, dataDef.SearchMatch);
    }

    // the saved ledger, checked the way the CSVs are: the header row must be
    // the definition's (Rdv3Xlsx), and every row needs an identity of its own
    private void ReadLedger(string[] head, out string[] lines, out string[] states)
    {
        Rdv3LedgerSnapshot snapshot = ledgerStore.Read(head);
        lines = snapshot.Lines;
        states = snapshot.States;
        if (snapshot.Warning.Length > 0) { ReportLedgerWarning(snapshot.Warning); }
    }

    private void ReportLedgerWarning(string warning)
    {
        log.Write("-", "warning", warning);
        form.RunOnUi(delegate { form.Error(warning); });
    }

    // a ledger column as the screen names it (data.labels), else its reference
    private string LabelOrRef(string reference)
    {
        string label = dataDef.LabelOf(reference);
        return (label.Length == 0) ? reference : label;
    }

    // make these lines the active, searchable ledger: the shared states with
    // the local pending values on top, and a fresh search index. A null head
    // keeps the current one.
    private void Install(string[] lines, string[] states, string[] head)
    {
        ledLines = lines;
        sharedStates = states;
        ledStates = pending.Overlay(lines, states, dataDef.IdentityCols);
        if (head != null) { ledHead = head; }
        ledIndex = BuildSearchIndex(lines);
    }

    // UI thread. READY again on whatever ledger is active: the operations open,
    // the work-state button unless a save is still undecided, the watch word.
    private void ResumeReady()
    {
        state = StReady;
        form.EnableOps(true);
        form.EnableWorkState(!writes.Pending);
        form.SetState(WatchStateText());
    }

    // adopt = make these lines the active, searchable ledger (worker builds the
    // index so the UI thread never runs over the whole ledger)
    private void AdoptLedger(string rid, string[] lines, string[] states, string[] head, string note)
    {
        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "adopt";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate
        {
            long t = Rdv3Clock.Now();
            Rdv3Index ix = BuildSearchIndex(lines);
            string[] effective = pending.Overlay(lines, states, dataDef.IdentityCols);
            log.Write(rid, "index", "table=LEDGER rows=" + lines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            form.RunOnUi(delegate
            {
                if (writes.Closing || !string.Equals(rid, activeRunId, StringComparison.Ordinal)) { return; }
                ledLines = lines;
                sharedStates = states;
                ledStates = effective;
                ledHead = head;
                ledIndex = ix;
                EnterReady(rid, note);
            });
        };
        worker.Post(job);
    }

    private void StartApply(string rid)
    {
        state = StApplying;
        writes.Begin(true);
        form.SetState(Rdv3Text.StateApplying);
        form.EnableOps(false);

        Rdv3MergeResult source = mergeResult;
        string[] checkedLines = savedLines;
        string[] beforeLines = ledLines;
        string[] beforeStates = ledStates;
        Rdv3Job job = new Rdv3Job();
        job.RunId = rid;
        job.Kind = "apply";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate { ApplyJob(rid, source, checkedLines, beforeLines, beforeStates); };
        worker.Post(job);
    }

    private void ApplyJob(string rid, Rdv3MergeResult source, string[] checkedLines, string[] beforeLines, string[] beforeStates)
    {
        Rdv3ApplyOutcome outcome = ledgerStore.Apply(source, checkedLines,
            pid.ToString(CultureInfo.InvariantCulture) + "-" + rid,
            delegate { return WaitForSharedLock(rid, Rdv3Text.StateApplying); },
            delegate(string stage, string detail) { log.Write(rid, stage, detail); },
            ReportLedgerWarning);
        if (outcome.Error is OperationCanceledException && writes.Closing) { return; }
        if (outcome.Error != null) { log.Write(rid, "error", "stage=persist msg=" + outcome.Error.Message); }
        if (!outcome.CanAdopt)
        {
            form.RunOnUi(delegate { EndApplyFailed(rid, outcome.Error.Message); });
            return;
        }

        long indexAt = Rdv3Clock.Now();
        Install(outcome.Update.Lines, outcome.Update.States, source.Head);
        List<Rdv3CandRow> resets = outcome.Error == null
            ? ResetCandidates(resetNotice.AfterUpdate(outcome.Update, beforeLines, beforeStates, ledStates))
            : new List<Rdv3CandRow>();
        log.Write(rid, "index", "table=LEDGER rows=" + outcome.Update.Lines.Length.ToString(CultureInfo.InvariantCulture)
            + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(indexAt)));
        form.RunOnUi(delegate { EndApply(rid, outcome, resets); });
    }

    private void EndApply(string rid, Rdv3ApplyOutcome outcome, List<Rdv3CandRow> resets)
    {
        RememberMarker(outcome.Marker);
        EndWriteGuard(rid, outcome.Error == null);
        EnterReady(rid, Rdv3Text.NoteUpdated);
        if (outcome.Error != null) { form.Error(Rdv3Text.ErrSharedMarker + outcome.Error.Message); return; }
        form.Notice(Rdv3Text.NoteUpdated);
        if (resets.Count > 0) { form.TellResetRows(resets); }
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
        EndWriteGuard(rid, false);
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
            if (Rdv3Files.Exists(ledgerPath))
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
        activeRunId = "";
        mergeResult = null;
        ResumeReady();
        // the result on show belongs to the ledger that was just replaced, so
        // it is cleared; the search key stays typed for a repeat search
        if (shownCands != null) { ClearShown(); }
        string rowsText = ledLines.Length.ToString("N0", CultureInfo.InvariantCulture);
        form.SetLedger(Rdv3Text.LedgerSegFmt.Replace("{file}", System.IO.Path.GetFileName(ledgerPath)).Replace("{n}", rowsText),
            rowsText, LedgerStamp());
        form.SetPendingCount(pending.Count);
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
        string capturePath = Environment.GetEnvironmentVariable("RDV_HEADLESS_CAPTURE_PATH");
        if (!headlessCaptureDone && capturePath != null && capturePath.Length > 0)
        {
            headlessCaptureDone = true;
            try
            {
                form.CaptureToFile(capturePath);
                log.Write("-", "screen", "headless capture=" + capturePath);
            }
            catch (Exception ex) { log.Write("-", "error", "stage=headless-capture msg=" + ex.Message); }
        }
        string headlessKey = Environment.GetEnvironmentVariable("RDV_HEADLESS_TEST_KEY");
        if (!headlessSearchStarted && headlessKey != null && headlessKey.Length > 0)
        {
            headlessSearchStarted = true;
            form.SetKeyText(headlessKey);
            ManualSearch(headlessKey);
        }
        string headlessModal = Environment.GetEnvironmentVariable("RDV_HEADLESS_OPEN_MODAL");
        if (!headlessModalStarted && headlessModal != null && headlessModal.Length > 0)
        {
            headlessModalStarted = true;
            form.TriggerProbeAction(headlessModal);
        }
    }

    private void ClearShown()
    {
        activeSearchId = "";
        activeSearchKey = "";
        form.ClearResult(true);
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
        key = Rdv3Input.Cell(key);
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
            form.Error(Rdv3Text.ErrBadKeyFmt.Replace("{label}", LabelOrRef(dataDef.SearchRefs[0]))
                .Replace("{pattern}", cfg.KeyPattern));
            log.Write("-", "search", "ignored key=" + key + " reason=bad-key");
            return;
        }
        Search(key, "manual", "", 0, t0);
    }

    // watch thread. "target" is the name of the watched screen the number came
    // off: with several watched at once, a reading in the log is only traceable
    // back to the work it belongs to if it says which one delivered it.
    private bool Detected(string key, double detectMs, int polls, long t0, string target)
    {
        bool accepted = false;
        form.RunOnUi(delegate
        {
            if (writes.Closing || state != StReady || writes.Pending || form.IsModalOpen) { return; }
            form.SetKeyText(key);
            Search(key, "detect", target, detectMs, t0);
            accepted = true;
        });
        if (accepted)
        {
            log.Write("-", "detect", "key=" + key + " target=" + target
                + " latency_ms=" + Rdv3Log.F(detectMs) + " polls=" + polls.ToString(CultureInfo.InvariantCulture));
        }
        return accepted;
    }

    // UI thread. t0 is the confirm instant: the search clock is already running.
    private void Search(string key, string source, string target, double detectMs, long t0)
    {
        if (state != StReady || writes.Pending || form.IsModalOpen) { return; }
        ClearShown();
        searchSeq++;
        string sid = "S" + searchSeq.ToString(CultureInfo.InvariantCulture);
        activeSearchId = sid;
        activeSearchKey = key;

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
            if (writes.Closing || state != StReady || !string.Equals(sid, activeSearchId, StringComparison.Ordinal)
                || !object.ReferenceEquals(lines, ledLines)) { return; }
            form.ShowCandidates(key, rows, n);
            shownKey = key;
            shownCands = candRows;
            shownRow = -1;
            if (n == 0) { form.Notice(Rdv3Text.NoteNotFound); }
            else if (n == 1)
            {
                // one hit is selected at once, the way the reference does it
                form.SelectCandidate(0);
                shownRow = candRows[0];
                if (source == "detect") { AutoCompleteDetected(); }
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
                if (writes.Closing || state != StReady || writes.Pending || form.IsModalOpen
                    || !string.Equals(sid, activeSearchId, StringComparison.Ordinal)
                    || !object.ReferenceEquals(lines, ledLines)) { return; }
                int picked = form.PickFromList();
                if (picked >= 0 && state == StReady && string.Equals(sid, activeSearchId, StringComparison.Ordinal)
                    && object.ReferenceEquals(lines, ledLines))
                {
                    PickCandidate(picked);
                    if (source == "detect") { AutoCompleteDetected(); }
                }
            });
        }
        form.SetTimes(-1, elapsed);
        string searchCapture = Environment.GetEnvironmentVariable("RDV_HEADLESS_SEARCH_CAPTURE_PATH");
        if (searchCapture != null && searchCapture.Length > 0)
        {
            form.RunOnUi(delegate { form.CaptureToFile(searchCapture); });
        }
    }

    // UI thread. The pick happens after the search clock stopped; its render is
    // the separate "display" figure in the log.
    private void PickCandidate(int i)
    {
        List<int> cands = shownCands;
        if (state != StReady || cands == null || i < 0 || i >= cands.Count) { return; }
        long t0 = Rdv3Clock.Now();
        int row = cands[i];
        string k2 = Rdv3Key.FromLine(ledLines[row], dataDef.IdentityCols);
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
        ChangeWorkState(false);
    }

    private void AutoCompleteDetected()
    {
        if (work.Trigger != "automatic") { return; }
        if (shownRow < 0 || shownRow >= ledStates.Length) { return; }
        Rdv3StateDef current = work.ByStored(ledStates[shownRow]);
        if (current == null || current.Id != work.Initial) { return; }
        ChangeWorkState(true);
    }

    private void ChangeWorkState(bool automatic)
    {
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        if (writes.Pending)
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
        string k2 = Rdv3Key.FromLine(ledLines[row], dataDef.IdentityCols);
        if (!automatic && tr.Confirm.Length > 0)
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
        writes.Begin(false);
        form.EnableWorkState(false);
        form.SetState(Rdv3Text.StateSavingFmt.Replace("{state}", to.Text));
        form.SetStoredState(CandOf(row), to.Stored, true);
        log.Write(pidTag, "state", "save started key2=" + k2 + " from=" + cur.Id + " to=" + to.Id + " (exit held until it is decided)");

        string was = stored;
        string next = to.Stored;
        string baseline = sharedStates[row];
        string line = ledLines[row];
        Rdv3Job job = new Rdv3Job();
        job.RunId = pidTag;
        job.Kind = "state";
        job.TimeoutMs = cfg.SaveTimeoutMs;
        job.Work = delegate
        {
            long t = Rdv3Clock.Now();
            try
            {
                pending.Set(k2, next, line, baseline);
                ledStates[row] = next;
            }
            catch (Exception ex)
            {
                ledStates[row] = was;
                log.Write(pidTag, "error", "stage=pending key2=" + k2 + " msg=" + ex.Message);
                form.RunOnUi(delegate
                {
                    form.Error(Rdv3Text.ErrPendingWrite + ex.Message);
                    form.SetStoredState(CandOf(row), was, false);
                    EndMarkSave(pidTag, false);
                });
                return;
            }
            double ms = Rdv3Clock.MsSince(t);
            form.RunOnUi(delegate
            {
                form.SetStoredState(CandOf(row), next, false);
                form.SetPendingCount(pending.Count);
                EndMarkSave(pidTag, true);
                form.Notice((tr.Done.Length > 0) ? tr.Done : Rdv3Text.NoteStateSaved.Replace("{state}", to.Text));
            });
            // e2e = confirm click -> file persisted AND the screen updated
            log.Write(pidTag, "state", "key2=" + k2
                + " row=" + (row + 1).ToString(CultureInfo.InvariantCulture)
                + " value=" + next + " target=local pending=" + pending.Count.ToString(CultureInfo.InvariantCulture)
                + " persist_ms=" + Rdv3Log.F(ms)
                + " e2e_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t0)));
        };
        worker.Post(job);
    }

    // UI thread. The save is decided -- saved or failed, either is a decision --
    // so operations and the close are released again.
    private void EndMarkSave(string tag, bool ok)
    {
        EndWriteGuard(tag, ok);
    }

    private void EndWriteGuard(string tag, bool ok)
    {
        bool requestedClose;
        if (!writes.Complete(out requestedClose)) { return; }
        form.EnableWorkState(state == StReady);
        if (state == StReady) { form.SetState(WatchStateText()); }
        log.Write(tag, "exit", "write decided (" + (ok ? "saved" : "failed") + "); exit released");
        if (requestedClose)
        {
            if (ok) { form.Notice(Rdv3Text.NoteSaveDoneCanClose); }
            else { form.Error(Rdv3Text.NoteSaveFailedCanClose); }
        }
    }

    private void DoClear()
    {
        activeSearchId = "";
        form.ClearResult();
        shownRow = -1;
        shownKey = "";
        shownCands = null;
        log.Write("-", "clear", "input and result cleared");
    }

    private void SendChanges()
    {
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        if (writes.Pending) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        int count = pending.Count;
        string body = Rdv3Text.ConfirmSendBody.Replace("{n}", count.ToString("N0", CultureInfo.InvariantCulture));
        if (!form.Ask(Rdv3Text.SendTitle, body))
        {
            log.Write("-", "send", "declined count=" + count.ToString(CultureInfo.InvariantCulture));
            return;
        }
        if (count == 0)
        {
            form.Notice(Rdv3Text.NoteNoPending);
            log.Write("-", "send", "nothing pending");
            return;
        }

        procSeq++;
        string tag = "T" + procSeq.ToString(CultureInfo.InvariantCulture);
        state = StSending;
        writes.Begin(true);
        form.EnableOps(false);
        form.SetState(Rdv3Text.StateSending);
        log.Write(tag, "send", "started count=" + count.ToString(CultureInfo.InvariantCulture));

        Rdv3Job job = new Rdv3Job();
        job.RunId = tag;
        job.Kind = "send";
        job.TimeoutMs = cfg.MarkOverdueMs;
        job.Work = delegate { SendJob(tag); };
        worker.Post(job);
    }

    private void SendJob(string tag)
    {
        Rdv3LedgerLock ledgerLock = null;
        bool ledgerWritten = false;
        Rdv3SharedMarker marker = null;
        string[] latestLines = null;
        string[] latestStates = null;
        Rdv3PendingApply apply = null;
        try
        {
            ledgerLock = WaitForSharedLock(tag, Rdv3Text.StateSending);
            long t = Rdv3Clock.Now();
            ReadLedger(ledHead, out latestLines, out latestStates);
            log.Write(tag, "load", "under_lock rows=" + latestLines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));

            apply = pending.PrepareSend(latestLines, latestStates, dataDef.IdentityCols, work.InitialStored);
            if (apply.FromInitial + apply.ToInitial > 0)
            {
                t = Rdv3Clock.Now();
                Rdv3Xlsx.Write(ledgerPath, ledHead, work.Column, latestLines, apply.States,
                    pid.ToString(CultureInfo.InvariantCulture) + "-" + tag, storageContract);
                ledgerWritten = true;
                log.Write(tag, "persist", "target=xlsx rows=" + latestLines.Length.ToString(CultureInfo.InvariantCulture)
                    + " changes=" + apply.Resolved.Count.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
                marker = shared.WriteMarker("send", latestLines.Length, apply.FromInitial, apply.ToInitial);
                log.Write(tag, "marker", "version=" + marker.Version.ToString(CultureInfo.InvariantCulture) + " kind=send");
            }
            // Includes idempotent retries after a previously committed send.
            pending.Remove(apply.Resolved);
            ledgerLock.Release();
            ledgerLock = null;

            Install(latestLines, apply.States, null);
            Rdv3PendingApply keepApply = apply;
            Rdv3SharedMarker keepMarker = marker;
            form.RunOnUi(delegate
            {
                if (keepMarker != null) { RememberMarker(keepMarker); }
                EndWriteGuard(tag, true);
                ReadyAfterShared(tag, Rdv3Text.NoteSendDone.Replace("{n}", keepApply.Resolved.Count.ToString("N0", CultureInfo.InvariantCulture)));
                if (keepApply.Resolved.Count > 0)
                {
                    form.Notice(Rdv3Text.NoteSendDone.Replace("{n}", keepApply.Resolved.Count.ToString("N0", CultureInfo.InvariantCulture)));
                }
                if (keepApply.Unmatched.Count > 0) { HandleUnmatched(keepApply.Unmatched); }
            });
        }
        catch (Exception ex)
        {
            if (ex is OperationCanceledException && writes.Closing) { return; }
            log.Write(tag, "error", "stage=send wrote=" + (ledgerWritten ? "true" : "false") + " msg=" + ex.Message);
            if (ledgerWritten && latestLines != null && apply != null) { Install(latestLines, apply.States, null); }
            Rdv3SharedMarker keepMarker = marker;
            Rdv3PendingApply keepApply = apply;
            form.RunOnUi(delegate
            {
                if (keepMarker != null) { RememberMarker(keepMarker); }
                EndWriteGuard(tag, false);
                ReadyAfterShared(tag, "send-failed");
                form.Error(Rdv3Text.ErrSend + ex.Message);
                if (keepApply != null && keepApply.Unmatched.Count > 0) { HandleUnmatched(keepApply.Unmatched); }
            });
        }
        finally
        {
            if (ledgerLock != null) { ledgerLock.Dispose(); }
        }
    }

    private void HandleUnmatched(List<Rdv3UnmatchedChange> rows)
    {
        if (!form.TellUnmatched(rows) || !form.Ask(Rdv3Text.UnmatchedTitle, Rdv3Text.DiscardPendingConfirm)) { return; }
        List<string> ids = new List<string>();
        for (int i = 0; i < rows.Count; i++) { ids.Add(rows[i].Identity); }
        writes.Begin(false);
        form.EnableOps(false);
        Rdv3Job job = new Rdv3Job();
        job.RunId = "P" + (++procSeq).ToString(CultureInfo.InvariantCulture);
        job.Kind = "state";
        job.TimeoutMs = cfg.SaveTimeoutMs;
        job.Work = delegate
        {
            try
            {
                string backup = pending.Path + ".discard-" + Guid.NewGuid().ToString("N") + ".bak";
                File.Copy(pending.Path, backup, false);
                pending.Remove(ids);
                ledStates = pending.Overlay(ledLines, sharedStates, dataDef.IdentityCols);
                form.RunOnUi(delegate
                {
                    EndWriteGuard(job.RunId, true);
                    ReadyAfterShared(job.RunId, "discarded-conflicts");
                    form.Notice(Rdv3Text.PendingBackup + backup);
                });
            }
            catch (Exception ex)
            {
                form.RunOnUi(delegate { EndWriteGuard(job.RunId, false); ResumeReady(); form.Error(ex.Message); });
            }
        };
        worker.Post(job);
    }

    private Rdv3LedgerLock WaitForSharedLock(string tag, string activeText)
    {
        int lastMinute = -1;
        string lastOwner = "";
        while (!writes.Closing)
        {
            Rdv3LockInfo owner;
            Rdv3LedgerLock result = shared.TryAcquire(out owner);
            if (result != null)
            {
                if (!writes.TryEnterSharedWrite())
                {
                    result.Dispose();
                    throw new OperationCanceledException("application is closing");
                }
                log.Write(tag, "lock", "acquired " + shared.LockPath);
                form.PostOnUi(delegate { form.SetState(activeText); });
                return result;
            }

            long staleAgeMs;
            if (shared.TryRemoveStaleLock(cfg.LockStaleMs, out staleAgeMs))
            {
                log.Write(tag, "lock", "removed stale lock age_ms=" + staleAgeMs.ToString(CultureInfo.InvariantCulture)
                    + " path=" + shared.LockPath);
                continue;
            }

            string ownerKey = (owner == null) ? "" : owner.User + "\t" + owner.Host;
            int minute = (owner == null) ? 0 : owner.AgeMinutes;
            if (minute != lastMinute || !string.Equals(ownerKey, lastOwner, StringComparison.Ordinal))
            {
                lastMinute = minute;
                lastOwner = ownerKey;
                string user = (owner == null || owner.User.Length == 0) ? Rdv3Text.NotYet : owner.User;
                string host = (owner == null || owner.Host.Length == 0) ? Rdv3Text.NotYet : owner.Host;
                string message = Rdv3Text.LockWaitingFmt.Replace("{user}", user).Replace("{host}", host)
                    .Replace("{minutes}", minute.ToString(CultureInfo.InvariantCulture));
                log.Write(tag, "lock", "waiting owner=" + user + " host=" + host + " age_min=" + minute.ToString(CultureInfo.InvariantCulture));
                form.PostOnUi(delegate { form.SetState(Rdv3Text.StateLockWaiting); form.Error(message); });
            }
            Thread.Sleep(cfg.LockRetryMs);
        }
        throw new OperationCanceledException("application is closing");
    }

    private void ReadyAfterShared(string tag, string note)
    {
        ResumeReady();
        if (shownCands != null) { ClearShown(); }
        string rowsText = ledLines.Length.ToString("N0", CultureInfo.InvariantCulture);
        form.SetLedger(Rdv3Text.LedgerSegFmt.Replace("{file}", System.IO.Path.GetFileName(ledgerPath)).Replace("{n}", rowsText),
            rowsText, LedgerStamp());
        form.SetPendingCount(pending.Count);
        log.Write(tag, "shared", "ready rows=" + ledLines.Length.ToString(CultureInfo.InvariantCulture)
            + " pending=" + pending.Count.ToString(CultureInfo.InvariantCulture) + " note=" + note);
    }

    private void OpenUpdateJob(string jobId)
    {
        if (writes.Pending || form.IsModalOpen) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        Rdv3ProcessJobDef process = dataDef.JobOf(jobId);
        if (process == null || process.Kind != "update") { return; }
        if (Rdv3ProcessForm.ShowJob(form, dataDef, jobId, dataDir, ledgerPath)) { RefreshLedger(process); }
    }

    private void OpenDeleteJob(string jobId)
    {
        if (writes.Pending || form.IsModalOpen) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        if (Rdv3ProcessForm.ShowJob(form, dataDef, jobId, dataDir, ledgerPath)) { StartDeleteJob(jobId); }
    }

    private void StartDeleteJob(string jobId)
    {
        if (state != StReady || writes.Pending) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        ClearShown();
        Rdv3ProcessJobDef process = dataDef.JobOf(jobId);
        if (process == null || process.Kind != "delete") { return; }
        state = StApplying;
        writes.Begin(true);
        form.EnableOps(false);
        form.SetState(Rdv3Text.StateDeleting);
        string tag = "D" + (++procSeq).ToString(CultureInfo.InvariantCulture);
        log.Write(tag, "delete", "started job=" + jobId);

        Rdv3Job job = new Rdv3Job();
        job.RunId = tag;
        job.Kind = "delete";
        job.TimeoutMs = cfg.MarkOverdueMs;
        job.Work = delegate { DeleteJob(tag, process); };
        worker.Post(job);
    }

    private void DeleteJob(string tag, Rdv3ProcessJobDef process)
    {
        Rdv3LedgerLock ledgerLock = null;
        bool ledgerWritten = false;
        Rdv3SharedMarker marker = null;
        Rdv3DeleteResult result = null;
        try
        {
            ledgerLock = WaitForSharedLock(tag, Rdv3Text.StateDeleting);
            string[] latestLines;
            string[] latestStates;
            long t = Rdv3Clock.Now();
            ReadLedger(ledHead, out latestLines, out latestStates);
            log.Write(tag, "load", "under_lock rows=" + latestLines.Length.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            string[] beforeEffective = pending.Overlay(latestLines, latestStates, dataDef.IdentityCols);

            t = Rdv3Clock.Now();
            result = Rdv3Ledger.ApplyDelete(dataDef, process, dataDir, latestLines, latestStates,
                                            work.InitialStored);
            for (int i = 0; i < result.Warnings.Count; i++) { log.Write(tag, "warning", result.Warnings[i]); }
            string[] afterEffective = pending.Overlay(result.Lines, result.States, dataDef.IdentityCols);
            List<Rdv3CandRow> resetRows = ResetCandidates(resetNotice.ChangedRows(latestLines, beforeEffective,
                result.Lines, afterEffective));
            bool changed = !Rdv3Ledger.SameLedger(latestLines, latestStates,
                                                  result.Lines, result.States);
            log.Write(tag, "delete", "selected=" + result.Deleted.ToString(CultureInfo.InvariantCulture)
                + " result=" + result.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " changed=" + (changed ? "true" : "false")
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            if (changed)
            {
                t = Rdv3Clock.Now();
                Rdv3Xlsx.Write(ledgerPath, ledHead, work.Column, result.Lines, result.States,
                    pid.ToString(CultureInfo.InvariantCulture) + "-" + tag, storageContract);
                ledgerWritten = true;
                log.Write(tag, "persist", "target=xlsx rows=" + result.Lines.Length.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
                marker = shared.WriteMarker("update", result.Lines.Length, 0, 0);
                log.Write(tag, "marker", "version=" + marker.Version.ToString(CultureInfo.InvariantCulture) + " kind=update");
            }
            ledgerLock.Release();
            ledgerLock = null;

            Install(result.Lines, result.States, null);
            Rdv3SharedMarker keepMarker = marker;
            int deleted = result.Deleted;
            form.RunOnUi(delegate
            {
                if (keepMarker != null) { RememberMarker(keepMarker); }
                EndWriteGuard(tag, true);
                string note = Rdv3Text.NoteDeleteDone.Replace("{n}", deleted.ToString("N0", CultureInfo.InvariantCulture));
                ReadyAfterShared(tag, note);
                form.Notice(note);
                if (result.Warnings.Count > 0) { form.Error(string.Join(Environment.NewLine, result.Warnings.ToArray())); }
                if (resetRows.Count > 0) { form.TellResetRows(resetRows); }
            });
        }
        catch (Exception ex)
        {
            if (ex is OperationCanceledException && writes.Closing) { return; }
            log.Write(tag, "error", "stage=delete wrote=" + (ledgerWritten ? "true" : "false") + " msg=" + ex.Message);
            if (ledgerWritten && result != null) { Install(result.Lines, result.States, null); }
            Rdv3SharedMarker keepMarker = marker;
            form.RunOnUi(delegate
            {
                if (keepMarker != null) { RememberMarker(keepMarker); }
                EndWriteGuard(tag, false);
                if (ledgerWritten && result != null) { ReadyAfterShared(tag, "delete-marker-failed"); }
                else { ResumeReady(); }
                form.Error(Rdv3Text.ErrDelete + ex.Message);
            });
        }
        finally
        {
            if (ledgerLock != null) { ledgerLock.Dispose(); }
        }
    }

    private void OpenTableExport()
    {
        if (writes.Pending || form.IsModalOpen) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        if (state != StReady) { form.Error(Rdv3Text.ErrNotReady); return; }
        string exportDir = System.IO.Path.GetDirectoryName(
            System.IO.Path.Combine(appDir, Rdv3Text.ExportDefaultPath));
        try
        {
            if (exportDir != null && exportDir.Length > 0 && !Directory.Exists(exportDir))
            {
                Directory.CreateDirectory(exportDir);
            }
        }
        catch (Exception ex)
        {
            form.Error(Rdv3Text.ErrExport + ex.Message);
            return;
        }
        Rdv3ExportRequest request = Rdv3ExportForm.Pick(form, dataDef, screen, appDir);
        if (request == null) { return; }
        try { request.Path = Rdv3Files.ExportPath(request.Path, appDir, dataDir, ledgerPath, log.Path, cfg.SourcePath, dataDef); }
        catch (Exception ex) { form.Error(Rdv3Text.ErrExport + ex.Message); return; }
        string[] lines = (string[])ledLines.Clone();
        string[] states = (string[])ledStates.Clone();
        writes.Begin(false);
        form.EnableOps(false);
        Rdv3Job job = new Rdv3Job();
        job.RunId = "E" + (++procSeq).ToString(CultureInfo.InvariantCulture);
        job.Kind = "export";
        job.TimeoutMs = cfg.SaveTimeoutMs;
        job.Work = delegate { ExportJob(job.RunId, request, lines, states); };
        worker.Post(job);
    }

    private void ExportJob(string tag, Rdv3ExportRequest request, string[] lines, string[] states)
    {
        bool exportedFile = false;
        try
        {
            List<int> columns = new List<int>();
            for (int i = 0; i < request.Fields.Count; i++)
            {
                columns.Add(request.Fields[i] == "$work" ? -2 : dataDef.IndexOf(request.Fields[i]));
            }
            StringBuilder output = new StringBuilder();
            string[] headings = Rdv3Files.ExportHeaders(request.Fields, dataDef, work.Column);
            for (int i = 0; i < request.Fields.Count; i++)
            {
                if (i > 0) { output.Append(','); }
                output.Append(Rdv3Files.CsvCell(headings[i], request.ExcelSafe));
            }
            output.Append("\r\n");
            int exported = 0;
            for (int row = 0; row < lines.Length; row++)
            {
                string[] values = Rdv3Ledger.SplitLine(lines[row]);
                if (!request.Matches(dataDef, values, states[row])) { continue; }
                for (int i = 0; i < columns.Count; i++)
                {
                    if (i > 0) { output.Append(','); }
                    string value = columns[i] == -2 ? states[row] : values[columns[i]];
                    output.Append(Rdv3Files.CsvCell(value, request.ExcelSafe));
                }
                output.Append("\r\n");
                exported++;
            }
            Rdv3Files.ExportPath(request.Path, appDir, dataDir, ledgerPath, log.Path, cfg.SourcePath, dataDef);
            Rdv3Files.WriteNewText(request.Path, output.ToString());
            exportedFile = true;
            log.Write(tag, "export", "rows=" + exported.ToString(CultureInfo.InvariantCulture) + " source_rows="
                + lines.Length.ToString(CultureInfo.InvariantCulture) + " filters="
                + request.Filters.Count.ToString(CultureInfo.InvariantCulture) + " fields="
                + columns.Count.ToString(CultureInfo.InvariantCulture) + " path=" + request.Path);
            form.RunOnUi(delegate { form.Notice(Rdv3Text.ExportDoneFmt.Replace("{file}", request.Path)); });
        }
        catch (Exception ex)
        {
            log.Write(tag, "error", "stage=export msg=" + ex.Message);
            form.RunOnUi(delegate { form.Error(Rdv3Text.ErrExport + ex.Message); });
        }
        finally
        {
            form.RunOnUi(delegate { EndWriteGuard(tag, exportedFile); ResumeReady(); });
        }
    }

    // ---- settings ------------------------------------------------------------
    // The settings dialog edits a copy; when it comes back, the file is written
    // and the running session adopts what it can without a restart.
    private void OpenSettings()
    {
        if (writes.Pending || form.IsModalOpen) { form.Error(Rdv3Text.ErrSaveInFlight); return; }
        Rdv3Config latest;
        try { latest = Rdv3Config.Load(cfg.SourcePath); }
        catch (Exception ex) { form.Error(Rdv3Text.ErrSettingsSave + ex.Message); return; }
        Rdv3Config edited = Rdv3SettingsForm.Edit(form, latest);
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
        watchdog.Interval = TimeSpan.FromMilliseconds(cfg.PumpMs);
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

    private void Tick()
    {
        CheckOverdue();
        PollSharedMarker();
        QueueDeferredReload();
    }

    private void PollSharedMarker()
    {
        if (writes.Closing || markerPollQueued || (lastMarkerPoll != 0 && Rdv3Clock.MsSince(lastMarkerPoll) < cfg.MarkerPollMs)) { return; }
        lastMarkerPoll = Rdv3Clock.Now();
        markerPollQueued = true;
        Rdv3Job job = new Rdv3Job();
        job.RunId = "marker";
        job.Kind = "marker";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate
        {
            Rdv3SharedMarker marker = null;
            string stamp = null;
            string error = null;
            try { marker = shared.ReadMarker(); } catch (Exception ex) { error = ex.Message; }
            // XLSX and marker are separate commits. Detect a changed workbook
            // even when the writer could not update its notification marker.
            try { stamp = Rdv3Files.Stamp(ledgerPath); } catch (Exception ex) { error = ex.Message; }
            form.PostOnUi(delegate
            {
                markerPollQueued = false;
                if (writes.Closing) { return; }
                if (error != null)
                {
                    if (!markerReadFailed) { form.Error(Rdv3Text.ErrSharedMarker + error); log.Write("-", "marker", error); }
                    markerReadFailed = true;
                }
                else { markerReadFailed = false; }
                bool fileChanged = stamp != null && ledgerObservedStamp != null && stamp != ledgerObservedStamp;
                if (stamp != null) { ledgerObservedStamp = stamp; }
                if (marker != null && !SameMarker(marker, seenMarker)) { ObserveMarker(marker); }
                else if (fileChanged)
                {
                    Rdv3SharedMarker reload = new Rdv3SharedMarker();
                    reload.Kind = "update";
                    reload.WriterId = "workbook-change-without-marker";
                    reload.WrittenUtcTicks = DateTime.UtcNow.Ticks;
                    deferredMarker = reload;
                }
            });
        };
        worker.Post(job);
    }

    private static bool SameMarker(Rdv3SharedMarker a, Rdv3SharedMarker b)
    {
        return a != null && b != null && a.Version == b.Version
            && a.WrittenUtcTicks == b.WrittenUtcTicks && a.WriterId == b.WriterId;
    }

    private void ObserveMarker(Rdv3SharedMarker marker)
    {
        RememberMarker(marker);
        log.Write("-", "marker", "observed version=" + marker.Version.ToString(CultureInfo.InvariantCulture)
            + " kind=" + marker.Kind + " writer=" + marker.WriterId);
        if (string.Equals(marker.WriterId, shared.WriterId, StringComparison.Ordinal)) { return; }

        if (marker.Kind == "send")
        {
            string actor = MarkerActor(marker);
            Rdv3StateDef initial = work.InitialState;
            Rdv3StateDef changed = work.InitialTargetState;
            string body = Rdv3Text.SharedSendBody.Replace("{user}", actor)
                .Replace("{changed}", marker.FromInitial.ToString("N0", CultureInfo.InvariantCulture))
                .Replace("{changedState}", (changed == null) ? "" : changed.Text)
                .Replace("{initial}", marker.ToInitial.ToString("N0", CultureInfo.InvariantCulture))
                .Replace("{initialState}", (initial == null) ? "" : initial.Text);
            form.SharedNotice(body);
            log.Write("-", "notice", "target=status text=" + body);
        }
        deferredMarker = marker;
    }

    private static string MarkerActor(Rdv3SharedMarker marker)
    {
        if (marker.User != null && marker.User.Length > 0) { return marker.User; }
        if (marker.Host != null && marker.Host.Length > 0) { return marker.Host; }
        return Rdv3Text.NotYet;
    }

    private void RememberMarker(Rdv3SharedMarker marker)
    {
        if (marker == null) { return; }
        seenMarker = marker;
        if (marker.FileStamp != null) { ledgerObservedStamp = marker.FileStamp; }
        if (SameMarker(deferredMarker, marker)) { deferredMarker = null; }
    }

    private void QueueDeferredReload()
    {
        if (deferredMarker == null || sharedReloading || writes.Pending || form.IsModalOpen || DateTime.UtcNow < reloadRetryAfter
            || (state != StReady && state != StBlocked)) { return; }
        ClearShown();
        Rdv3SharedMarker marker = deferredMarker;
        deferredMarker = null;
        sharedReloading = true;
        state = StReloading;
        form.EnableOps(false);
        form.SetState(Rdv3Text.StateReloading);
        string tag = "L" + (++procSeq).ToString(CultureInfo.InvariantCulture);
        Rdv3Job job = new Rdv3Job();
        job.RunId = tag;
        job.Kind = "reload";
        job.TimeoutMs = cfg.CheckTimeoutMs;
        job.Work = delegate { ReloadSharedJob(tag, marker); };
        worker.Post(job);
    }

    private void ReloadSharedJob(string tag, Rdv3SharedMarker marker)
    {
        try
        {
            string[] lines;
            string[] states;
            long t = Rdv3Clock.Now();
            string[] expectedHead = (ledHead != null) ? ledHead : ((mergeResult != null) ? mergeResult.Head : null);
            if (expectedHead == null) { throw new InvalidOperationException("ledger header is not available"); }
            ReadLedger(expectedHead, out lines, out states);
            Rdv3Index index = BuildSearchIndex(lines);
            string[] effective = pending.Overlay(lines, states, dataDef.IdentityCols);
            List<Rdv3CandRow> resets = (marker.Kind == "update")
                ? ResetCandidates(resetNotice.ChangedRows(ledLines, ledStates, lines, states)) : new List<Rdv3CandRow>();
            log.Write(tag, "reload", "version=" + marker.Version.ToString(CultureInfo.InvariantCulture)
                + " rows=" + lines.Length.ToString(CultureInfo.InvariantCulture)
                + " reset=" + resets.Count.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            form.RunOnUi(delegate
            {
                sharedReloading = false;
                bool adopt = marker.Kind != "update" || form.AskLedgerSwitch(resets);
                if (adopt)
                {
                    ledLines = lines;
                    sharedStates = states;
                    ledStates = effective;
                    ledHead = expectedHead;
                    ledIndex = index;
                    ReadyAfterShared(tag, "marker-" + marker.Version.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    ResumeReady();
                    form.Error(Rdv3Text.StaleLedger);
                    log.Write(tag, "reload", "switch declined version=" + marker.Version.ToString(CultureInfo.InvariantCulture));
                }
            });
        }
        catch (Exception ex)
        {
            log.Write(tag, "error", "stage=reload msg=" + ex.Message);
            form.RunOnUi(delegate
            {
                sharedReloading = false;
                if (deferredMarker == null) { deferredMarker = marker; }
                reloadRetryAfter = DateTime.UtcNow.AddMilliseconds(cfg.MarkerPollMs);
                if (ledLines != null) { ResumeReady(); } else { EnterBlocked(); }
                form.Error(Rdv3Text.ErrSharedReload + ex.Message);
            });
        }
    }

    private List<Rdv3CandRow> ResetCandidates(List<string> lines)
    {
        List<Rdv3CandRow> rows = new List<Rdv3CandRow>();
        foreach (string line in lines)
        {
            Rdv3CandRow row = new Rdv3CandRow();
            row.Line = line;
            row.Stored = work.InitialStored;
            rows.Add(row);
        }
        return rows;
    }

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
        if (n == 0) { return Rdv3Text.LabelWatch; }
        return Rdv3Text.LabelWatch;
    }

    private int WatchableCount()
    {
        int n = 0;
        for (int i = 0; i < cfg.Targets.Count; i++) { if (cfg.Targets[i].IsWatchable) { n++; } }
        return n;
    }

    // WHAT THE STATE WORD SAYS when nothing else is going on
    private string WatchStateText()
    {
        if (WatchableCount() == 0) { return Rdv3Text.StateNoTarget; }
        if (watch.Bound) { return Rdv3Text.StateReady; }
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
        if ((st == "WATCHING" || st == "WAITING") && state == StReady && !writes.Pending)
        {
            form.SetState(WatchStateText());
        }
    }

    private void CheckOverdue()
    {
        Rdv3Job j = worker.TakeOverdue();
        if (j == null) { return; }
        log.Write(j.RunId, "timeout", "kind=" + j.Kind + " after_ms=" + Rdv3Log.F(Rdv3Clock.MsSince(j.StartedAt)));
        if (string.Equals(j.Kind, "search", StringComparison.Ordinal))
        {
            if (j.RunId == activeSearchId) { ClearShown(); form.Error(Rdv3Text.ErrSearchTimeout); }
            return;
        }
        if (string.Equals(j.Kind, "state", StringComparison.Ordinal)
            || string.Equals(j.Kind, "send", StringComparison.Ordinal)
            || string.Equals(j.Kind, "apply", StringComparison.Ordinal)
            || string.Equals(j.Kind, "delete", StringComparison.Ordinal)
            || string.Equals(j.Kind, "export", StringComparison.Ordinal))
        {
            // a managed job cannot be aborted, and until the write returns
            // nobody knows whether the record reached the file. Report the
            // delay; keep holding the exit rather than claim a decision.
            form.Error(j.Kind == "send" ? Rdv3Text.ErrSendOverdue
                : ((j.Kind == "apply" || j.Kind == "delete") ? Rdv3Text.ErrSharedWriteOverdue : Rdv3Text.ErrSaveOverdue));
            return;
        }
        if (string.Equals(j.Kind, "reload", StringComparison.Ordinal))
        {
            form.Error(Rdv3Text.ErrReloadOverdue);
            return;
        }
        if (string.Equals(j.RunId, activeRunId, StringComparison.Ordinal))
        {
            // abandon the run: whatever arrives later is stale by ID
            activeRunId = "";
            form.Error(Rdv3Text.ErrCheckTimeout);
            if (ledLines != null) { ResumeReady(); }
            else if (savedLines != null) { AdoptTimeoutFallbackSaved(); }
            else { EnterBlocked(); }
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
                if (writes.Pending) { EndWriteGuard(job.RunId, false); }
                form.Fatal(Rdv3Text.FatalDataTitle, Rdv3Text.FatalData.Replace("{reason}", ex.Message));
            });
            return;
        }
        form.RunOnUi(delegate
        {
            form.Error(Rdv3Text.ErrCheckFailed + ex.Message);
            // a state job that threw anywhere is still a decided save (failed):
            // the guard must never outlive the job that armed it
            if (job.Kind == "state" || job.Kind == "apply" || job.Kind == "delete" || job.Kind == "send" || job.Kind == "export")
            { EndWriteGuard(job.RunId, false); }
            if (!string.Equals(job.RunId, activeRunId, StringComparison.Ordinal)) { return; }
            activeRunId = "";
            if (ledLines != null) { ResumeReady(); }
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
