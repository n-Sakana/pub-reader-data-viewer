// The per-terminal operation log: one CSV beside the shared ledger for each
// terminal, appended right after this terminal has replaced the ledger file.
// Terminals never share a file, so writers on different PCs cannot corrupt
// each other's records. The ledger write is already committed when a line is
// recorded, so a line that cannot reach the shared folder is never a failure of
// the operation: it waits in a local spool and rides along with the next
// successful append. C# 5.
using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

public sealed class Rdv3OperationLog
{
    public const string Extension = ".csv";
    private const int Attempts = 5;
    private const int RetryMs = 200;
    private readonly string path;
    private readonly string spoolPath;
    private readonly string host;
    private readonly string user;
    private readonly object gate = new object();

    public Rdv3OperationLog(string ledgerPath, string machine, string userName, string spool)
    {
        path = PathFor(ledgerPath, machine);
        spoolPath = spool;
        host = (machine == null) ? "" : machine;
        user = (userName == null) ? "" : userName;
    }

    public string Path { get { return path; } }
    public string SpoolPath { get { return spoolPath; } }

    // <ledger stem><infix><terminal>.csv, for example Ledger-<infix>-PC01.csv next to Ledger.xlsx.
    public static string FileNameFor(string ledgerPath, string machine)
    {
        return System.IO.Path.GetFileNameWithoutExtension(ledgerPath) + Rdv3Text.OpLogInfix + SafeName(machine) + Extension;
    }

    public static string PathFor(string ledgerPath, string machine)
    {
        string full = System.IO.Path.GetFullPath(ledgerPath);
        return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(full), FileNameFor(full, machine));
    }

    // The local spool sits with the pending store, one per ledger.
    public static string SpoolPathFor(string ledgerPath)
    {
        string pending;
        try { pending = Rdv3PendingStore.PathFor(ledgerPath); }
        catch (IOException)
        {
            pending = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ReaderDataViewer"),
                "pending-" + Rdv3Files.SessionKey(ledgerPath) + ".dat");
        }
        string name = System.IO.Path.GetFileNameWithoutExtension(pending);
        if (name.StartsWith("pending-", StringComparison.Ordinal)) { name = name.Substring("pending-".Length); }
        return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(pending), "operations-" + name + ".spool.csv");
    }

    // Every <stem><infix>*.csv in the ledger's folder is some terminal's log and
    // must not be used as an input, a log or an export target.
    public static bool IsOperationLog(string candidate, string ledgerPath)
    {
        string full = System.IO.Path.GetFullPath(candidate);
        string ledger = System.IO.Path.GetFullPath(ledgerPath);
        if (!string.Equals(System.IO.Path.GetDirectoryName(full), System.IO.Path.GetDirectoryName(ledger), StringComparison.OrdinalIgnoreCase))
        { return false; }
        string prefix = System.IO.Path.GetFileNameWithoutExtension(ledger) + Rdv3Text.OpLogInfix;
        string name = System.IO.Path.GetFileName(full);
        return name.Length > prefix.Length + Extension.Length
            && name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            && name.EndsWith(Extension, StringComparison.OrdinalIgnoreCase);
    }

    private static string SafeName(string machine)
    {
        string value = (machine == null) ? "" : machine.Trim();
        if (value.Length == 0) { value = "unknown"; }
        char[] invalid = System.IO.Path.GetInvalidFileNameChars();
        StringBuilder sb = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++) { sb.Append(Array.IndexOf(invalid, value[i]) >= 0 ? '_' : value[i]); }
        return sb.ToString();
    }

    public static string UpdateDetail(Rdv3ProcessJobDef job, Rdv3UpdateResult update, Rdv3WorkState work)
    {
        Rdv3StateDef initial = (work == null) ? null : work.InitialState;
        return Rdv3Text.OpUpdateDetailFmt.Replace("{job}", JobName(job))
            .Replace("{added}", N(update.Added)).Replace("{updated}", N(update.Updated))
            .Replace("{deleted}", N(update.Deleted)).Replace("{reset}", N(update.ResetLines.Count))
            .Replace("{state}", (initial == null) ? "" : initial.Text);
    }

    public static string DeleteDetail(Rdv3ProcessJobDef job, int deleted)
    {
        return Rdv3Text.OpDeleteDetailFmt.Replace("{job}", JobName(job)).Replace("{n}", N(deleted));
    }

    public static string SendDetail(Rdv3StateDef changedState, int changed, Rdv3StateDef initialState, int initial)
    {
        return Rdv3Text.OpSendDetailFmt
            .Replace("{changedState}", (changedState == null) ? "" : changedState.Text).Replace("{changed}", N(changed))
            .Replace("{initialState}", (initialState == null) ? "" : initialState.Text).Replace("{initial}", N(initial));
    }

    private static string JobName(Rdv3ProcessJobDef job)
    {
        if (job == null) { return ""; }
        return (job.Name != null && job.Name.Length > 0) ? job.Name : job.Id;
    }

    private static string N(int value) { return value.ToString("N0", CultureInfo.InvariantCulture); }

    // Appends one line for a ledger replacement this terminal has just made.
    // Returns null when the line reached the shared file, otherwise the reason
    // it was kept in the local spool instead. Never throws.
    public string Record(string operation, int rows, string detail)
    {
        string line = Line(DateTime.Now, operation, rows, detail);
        lock (gate)
        {
            string failure = TryAppend(line);
            if (failure == null) { return null; }
            try { AppendText(spoolPath, line, true); }
            catch (Exception ex) { failure += "; spool: " + ex.Message; }
            return failure;
        }
    }

    // Re-sends spooled lines, oldest first. Null when nothing is left in the spool.
    public string Flush()
    {
        lock (gate)
        {
            if (!HasSpool()) { return null; }
            return TryAppend("");
        }
    }

    public bool HasSpool()
    {
        try { return File.Exists(spoolPath) && new FileInfo(spoolPath).Length > 0; }
        catch (Exception) { return false; }
    }

    private string Line(DateTime at, string operation, int rows, string detail)
    {
        return Cell(at.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)) + "," + Cell(host) + "," + Cell(user)
            + "," + Cell(operation) + "," + rows.ToString(CultureInfo.InvariantCulture) + "," + Cell(detail) + "\r\n";
    }

    private static string Cell(string value)
    {
        string text = (value == null) ? "" : value.Replace("\r", " ").Replace("\n", " ");
        return Rdv3Files.CsvCell(text, false);
    }

    private string TryAppend(string line)
    {
        string spooled = "";
        try { if (HasSpool()) { spooled = File.ReadAllText(spoolPath, Encoding.UTF8); } }
        catch (Exception ex) { return "spool unreadable: " + ex.Message; }
        string text = spooled + line;
        if (text.Length == 0) { return null; }
        Exception last = null;
        for (int attempt = 0; attempt < Attempts; attempt++)
        {
            try
            {
                AppendShared(text);
                if (spooled.Length > 0) { try { File.Delete(spoolPath); } catch (Exception) { } }
                return null;
            }
            catch (IOException ex) { last = ex; }
            catch (UnauthorizedAccessException ex) { last = ex; }
            if (attempt + 1 < Attempts) { Thread.Sleep(RetryMs); }
        }
        return last.Message;
    }

    // One open per append: FileShare.Read denies other writers while the whole
    // text goes in, so a same-named terminal or a second session on this PC can
    // never interleave lines. The folder is not created here: it exists whenever
    // the ledger beside it was just written, and a missing share is a spool case.
    private void AppendShared(string text)
    {
        string mutexName;
        using (SHA256 hash = SHA256.Create())
        { mutexName = "Rdv3OpLog-" + BitConverter.ToString(hash.ComputeHash(Encoding.UTF8.GetBytes(path.ToUpperInvariant()))).Replace("-", ""); }
        using (Mutex mutex = new Mutex(false, mutexName))
        {
            bool owned = false;
            try
            {
                try { owned = mutex.WaitOne(2000); }
                catch (AbandonedMutexException) { owned = true; }
                if (!owned) { throw new IOException("operation log is busy: " + path); }
                using (FileStream file = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read))
                {
                    if (file.Position == 0)
                    {
                        byte[] head = Encoding.UTF8.GetBytes(Rdv3Text.OpLogHeader + "\r\n");
                        file.Write(new byte[] { 0xEF, 0xBB, 0xBF }, 0, 3);
                        file.Write(head, 0, head.Length);
                    }
                    byte[] body = Encoding.UTF8.GetBytes(text);
                    file.Write(body, 0, body.Length);
                    file.Flush(true);
                }
            }
            finally { if (owned) { mutex.ReleaseMutex(); } }
        }
    }

    private static void AppendText(string destination, string text, bool createDirectory)
    {
        if (createDirectory) { Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destination)); }
        using (FileStream file = new FileStream(destination, FileMode.Append, FileAccess.Write, FileShare.Read))
        {
            byte[] body = Encoding.UTF8.GetBytes(text);
            file.Write(body, 0, body.Length);
            file.Flush(true);
        }
    }
}
