// Flush each record before continuing: a killed process cannot log its own
// death. An unfinished BEGIN and its last phase must survive without an END.
using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public sealed class Rdv3Log
{
    private readonly string path;
    private readonly object gate = new object();
    private bool failed;
    public Action<string> OnFail;
    public const int MaxBytes = 4 * 1024 * 1024;
    public const int PreviousFiles = 3;
    private static readonly int processId = Process.GetCurrentProcess().Id;
    private static readonly object lifecycle = new object();
    private static bool started;
    private static volatile bool completed;
    private static volatile string phase = "bootstrap";
    private static long phaseAt = DateTime.UtcNow.Ticks;
    private static long uiAt;
    private static int uiPending;
    private static Action<Action> postUi;
    private static Timer heartbeat;
    private static int reportedLogFailure;

    public static string FeedbackPath
    {
        get { return System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ReaderDataViewer", "logs", "feedback.log"); }
    }

    public static string FallbackPath
    {
        get { return System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ReaderDataViewer", "logs", "feedback.log"); }
    }

    public static void Begin(string mode, string context)
    {
        lock (lifecycle)
        {
            if (started) { return; }
            started = true;
            Feedback("BEGIN", "mode=" + mode + " " + context);
            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            { Feedback("FATAL", "terminating=" + e.IsTerminating + " " + e.ExceptionObject); };
            AppDomain.CurrentDomain.ProcessExit += delegate
            { if (!completed) { Feedback("UNEXPECTED EXIT", "No END was reached; last phase=" + phase); } };
            TaskScheduler.UnobservedTaskException += delegate(object sender, UnobservedTaskExceptionEventArgs e)
            { Error("unobserved task", e.Exception); };
            heartbeat = new Timer(Heartbeat, null, 5000, 5000);
        }
    }

    public static void Phase(string value)
    {
        phase = value;
        Interlocked.Exchange(ref phaseAt, DateTime.UtcNow.Ticks);
        Feedback("PHASE", value);
    }

    public static void MonitorUi(Action<Action> dispatch)
    {
        postUi = dispatch;
        Interlocked.Exchange(ref uiAt, DateTime.UtcNow.Ticks);
    }

    private static void Heartbeat(object unused)
    {
        if (completed) { return; }
        long now = DateTime.UtcNow.Ticks;
        string detail = "last phase=" + phase + " phase_ms=" + ((now - Interlocked.Read(ref phaseAt)) / TimeSpan.TicksPerMillisecond);
        Action<Action> dispatch = postUi;
        if (dispatch != null)
        {
            if (Interlocked.CompareExchange(ref uiPending, 1, 0) == 0)
            {
                try { dispatch(delegate { Interlocked.Exchange(ref uiAt, DateTime.UtcNow.Ticks); Interlocked.Exchange(ref uiPending, 0); }); }
                catch (Exception ex) { Error("UI heartbeat dispatch", ex); }
            }
            long delay = (now - Interlocked.Read(ref uiAt)) / TimeSpan.TicksPerMillisecond;
            detail += " ui_ack_ms=" + delay;
            Feedback(delay >= 10000 ? "UI NOT RESPONDING" : "HEARTBEAT", detail);
        }
        else { Feedback("HEARTBEAT", detail + "; still running, completion not confirmed"); }
    }

    public static void End(int code)
    {
        completed = true;
        Timer timer = heartbeat;
        if (timer != null) { timer.Dispose(); }
        Feedback("END", "exit=" + code + " last phase=" + phase);
    }

    public static void Error(string context, Exception error)
    {
        Feedback("ERROR", context + ": " + (error == null ? "unknown error" : error.ToString()));
    }

    public static void Feedback(string section, string detail)
    {
        string line = Line("pid=" + processId, section, detail);
        try { AppendBounded(FeedbackPath, line); return; }
        catch (Exception primary)
        {
            try
            {
                AppendBounded(FallbackPath, Line("pid=" + processId, "LOG FALLBACK", FeedbackPath + ": " + primary.Message) + line);
                return;
            }
            catch (Exception fallback)
            {
                if (Interlocked.Exchange(ref reportedLogFailure, 1) == 0)
                {
                    try { Console.Error.WriteLine("LOG WRITE FAILED " + FeedbackPath + ": " + primary.Message + "; " + FallbackPath + ": " + fallback.Message); }
                    catch (Exception) { }
                }
            }
        }
    }

    private static string Line(string runId, string section, string detail)
    {
        return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture)
            + "\t" + runId + "\t" + section + "\t" + (detail ?? "").Replace('\t', ' ') + "\r\n";
    }

    public static void AppendBounded(string destination, string text)
    {
        destination = System.IO.Path.GetFullPath(destination);
        string mutexName;
        using (SHA256 hash = SHA256.Create())
        { mutexName = "Rdv3Log-" + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(destination.ToUpperInvariant()))).Replace("-", ""); }
        using (Mutex mutex = new Mutex(false, mutexName))
        {
            bool owned = false;
            try
            {
                try { owned = mutex.WaitOne(1000); }
                catch (AbandonedMutexException) { owned = true; }
                if (!owned) { throw new IOException("Log is busy: " + destination); }
                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destination));
                if (Encoding.UTF8.GetByteCount(text) > MaxBytes)
                { text = text.Substring(0, Math.Min(text.Length, MaxBytes / 4 - 100)) + "\r\n[LOG ENTRY TRUNCATED: exceeded 4 MiB]\r\n"; }
                long size = File.Exists(destination) ? new FileInfo(destination).Length : 0;
                if (size + Encoding.UTF8.GetByteCount(text) > MaxBytes)
                {
                    for (int i = PreviousFiles; i >= 1; i--)
                    {
                        string older = destination + "." + i.ToString(CultureInfo.InvariantCulture);
                        string newer = i == 1 ? destination : destination + "." + (i - 1).ToString(CultureInfo.InvariantCulture);
                        if (File.Exists(older)) { File.Delete(older); }
                        if (File.Exists(newer)) { File.Move(newer, older); }
                    }
                }
                using (FileStream file = new FileStream(destination, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(text);
                    file.Write(bytes, 0, bytes.Length);
                    file.Flush(true);
                }
            }
            finally { if (owned) { mutex.ReleaseMutex(); } }
        }
    }

    public Rdv3Log(string p)
    {
        path = p;
    }

    public string Path { get { return path; } }

    public void Write(string runId, string section, string detail)
    {
        Feedback(section, "run=" + runId + " log=" + path + " " + detail);
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
                if (!string.Equals(System.IO.Path.GetFullPath(path), FeedbackPath, StringComparison.OrdinalIgnoreCase))
                { AppendBounded(path, sb.ToString()); }
                failed = false;
            }
            catch (Exception ex)
            {
                // a log that cannot be written must not crash the app, but it
                // must not fail silently either: surface it once
                if (!failed)
                {
                    failed = true;
                    Error("configured log " + path, ex);
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
