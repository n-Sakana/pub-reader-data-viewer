// Startup validation, path resolution and lifetime of the running app.
using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;

public static class Rdv3Program
{
    private static readonly IntPtr DpiAwarenessContextSystemAware = new IntPtr(-2);

    [DllImport("user32.dll")]
    private static extern bool SetProcessDpiAwarenessContext(IntPtr value);

    // configPath: settings.json (the bootstrap's -config, or the one next to
    // the .cmd). baseDir: the .cmd's folder, what relative paths are relative
    // to. The three overrides are the bootstrap's command line ("" = none).
    public static int Run(string configPath, string baseDir, string dataDirArg, string ledgerArg, string logArg, double compileMs)
    {
        try { SetProcessDpiAwarenessContext(DpiAwarenessContextSystemAware); }
        catch (EntryPointNotFoundException) { }

        if (IntPtr.Size != 8)
        {
            ReaderDataViewer.App.ShowStartupMessage(Rdv3Text.ErrNo64, null);
            return 2;
        }

        // ---- the settings file: exists, parses, checks out -- or nothing runs
        Rdv3Config cfg;
        string settingsDir = "";
        try { settingsDir = Path.GetDirectoryName(Path.GetFullPath(configPath)); } catch (Exception) { }
        try
        {
            Rdv3Log.Phase("window settings " + configPath);
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

        try
        {
            Rdv3Config resolved = cfg.Clone();
            resolved.DataDir = dataDir; resolved.Ledger = ledgerPath; resolved.Log = logPath;
            Rdv3Files.ValidateLayout(resolved, baseDir, configPath);
        }
        catch (Exception ex)
        {
            ReaderDataViewer.App.ShowStartupMessage(Rdv3Text.FatalSettings.Replace("{file}", configPath)
                .Replace("{reason}", ex.Message), null);
            return 6;
        }

        // ---- the data the definition names: the files exist and their headers
        // hold every column the definition uses
        try
        {
            if (!Directory.Exists(dataDir)) { throw new Rdv3DataError(Rdv3Text.ErrDataDir + dataDir); }
            string[][] heads = new string[cfg.Data.Tables.Count][];
            for (int t = 0; t < cfg.Data.Tables.Count; t++)
            {
                string p = Path.Combine(dataDir, cfg.Data.Tables[t].File);
                Rdv3Log.Phase("window input header " + p);
                if (!File.Exists(p)) { throw new Rdv3DataError(Rdv3Text.ErrNoData + p); }
                heads[t] = Rdv3Table.ReadHead(p, cfg.Data.Tables[t].Enc, cfg.Data.Tables[t].EncodingSetting,
                    cfg.Data.SourceReferences(cfg.Data.Tables[t].Id), cfg.Data.Tables[t].HeaderRow, cfg.Data.Tables[t].Delimiter,
                    cfg.Data.Tables[t].Sheet);
            }
            cfg.Data.Bind(heads);
            if (cfg.Data.TypeOrder.Count > 0)
            {
                Rdv3Table[] typedTables = new Rdv3Table[cfg.Data.Tables.Count];
                for (int i = 0; i < cfg.Data.TypeOrder.Count; i++)
                {
                    int tableOrd = cfg.Data.TypeOrder[i].TableOrd;
                    // a type on a column the update job makes has no file to check here
                    if (tableOrd < 0 || typedTables[tableOrd] != null) { continue; }
                    Rdv3TableDef table = cfg.Data.Tables[tableOrd];
                    typedTables[tableOrd] = Rdv3Table.Read(Path.Combine(dataDir, table.File),
                        table.Id, table.Enc, table.KeyColumns, table.KeyValidation, table.EncodingSetting, cfg.Data.SourceReferences(table.Id),
                        table.HeaderRow, table.Delimiter, table.Sheet);
                }
                cfg.Data.ConvertWorkbookDates(typedTables);
                cfg.Data.ValidateTypes(typedTables);
            }
        }
        catch (Exception ex)
        {
            Stop(logPath, "data", "not started: " + ex.Message,
                Rdv3Text.FatalDataTitle, Rdv3Text.FatalData.Replace("{reason}", ex.Message));
            return 3;
        }

        // Every local store and shared companion file derives from this same
        // canonical ledger path.
        try { ledgerPath = Path.GetFullPath(ledgerPath); }
        catch (Exception ex)
        {
            Stop(logPath, "ledger", "not started: bad ledger path " + ledgerPath + ": " + ex.Message,
                Rdv3Text.FatalTitle, Rdv3Text.ErrBadLedgerPath + ledgerPath + "\r\n" + ex.Message);
            return 5;
        }
        // Keep one local session per ledger. Network-wide serialization is done
        // separately by the lock file beside the shared ledger.
        string mutexName = "RdvApp-" + Rdv3Files.SessionKey(ledgerPath);
        bool createdNew;
        using (Mutex mx = new Mutex(true, mutexName, out createdNew))
        {
            if (!createdNew)
            {
                ReaderDataViewer.App.ShowStartupMessage(Rdv3Text.ErrAlreadyRunning, null);
                return 4;
            }

            FileStream localSession;
            try { localSession = Rdv3Files.AcquireLocalSession(Rdv3PendingStore.PathFor(ledgerPath)); }
            catch (Exception ex)
            {
                ReaderDataViewer.App.ShowStartupMessage(Rdv3Text.ErrAlreadyRunning + "\r\n" + ex.Message, null);
                return 4;
            }
            using (localSession)
            {
            Rdv3PendingStore pending;
            try
            {
                string pendingPath = Rdv3PendingStore.PathFor(ledgerPath);
                pending = new Rdv3PendingStore(pendingPath);
                pending.Validate(cfg.Screen.Work);
            }
            catch (Exception ex)
            {
                Stop(logPath, "pending", "not started: " + ex.Message,
                    Rdv3Text.FatalTitle, Rdv3Text.ErrPendingRead + ex.Message);
                return 7;
            }

            Rdv3SharedFiles shared;
            Rdv3SharedMarker initialMarker;
            try
            {
                int currentPid = System.Diagnostics.Process.GetCurrentProcess().Id;
                string instance = Environment.MachineName + "-" + currentPid.ToString(CultureInfo.InvariantCulture)
                    + "-" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
                shared = new Rdv3SharedFiles(ledgerPath, Environment.MachineName, Environment.UserName, instance);
                // Operation lines kept back by an unreachable share get another
                // chance now; whatever still fails waits for the next write.
                try
                {
                    string undelivered = shared.Operations.Flush();
                    if (undelivered != null)
                    { new Rdv3Log(logPath).Write("-", "oplog", "startup flush: " + undelivered); }
                }
                catch (Exception flushError)
                { new Rdv3Log(logPath).Write("-", "oplog", "spool flush failed: " + flushError.Message); }
                initialMarker = null;
                try { initialMarker = shared.ReadMarker(); }
                catch (Exception markerError)
                { new Rdv3Log(logPath).Write("-", "marker", "startup marker unavailable: " + markerError.Message); }
            }
            catch (Exception ex)
            {
                Stop(logPath, "marker", "not started: " + ex.Message,
                    Rdv3Text.FatalTitle, Rdv3Text.ErrSharedMarker + ex.Message);
                return 8;
            }

            Application application = new Application();
            application.DispatcherUnhandledException += delegate(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
            { Rdv3Log.Error("UI dispatcher", e.Exception); };
            Rdv3Log.MonitorUi(delegate(Action action) { application.Dispatcher.BeginInvoke(action); });
            application.ShutdownMode = ShutdownMode.OnMainWindowClose;
            ReaderDataViewer.MainWindow window = new ReaderDataViewer.MainWindow(cfg.Screen);
            Rdv3Form form = new Rdv3Form(window, cfg.Screen);
            Rdv3App app = new Rdv3App(form, baseDir, dataDir, ledgerPath, logPath, cfg, pending, shared, initialMarker);
            app.LogBoot(compileMs);
            Rdv3Log.Phase("window event loop");
            application.Run(window);
            return window.ExitCode;
            }
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
        ReaderDataViewer.App.ShowStartupMessage(
            Rdv3Text.AppTitle + " - " + title + Environment.NewLine + Environment.NewLine + body,
            null);
    }
}
