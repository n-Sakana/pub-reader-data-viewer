// The per-terminal operation log: one CSV beside the shared ledger for each
// terminal, appended right after this terminal has replaced the ledger file.
// Terminals never share a file, so writers on different PCs cannot corrupt
// each other's records. The ledger write is already committed when a line is
// recorded, so a line that cannot reach the shared folder is never a failure of
// the operation: it waits in a local spool and rides along with the next
// successful append. A spool that has been delivered but cannot be removed is
// remembered by length, so nothing is ever appended twice. C# 5.
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
    private readonly string deliveredPath;
    private readonly string host;
    private readonly string user;
    private readonly object gate = new object();
    // characters at the start of the spool that already reached the shared file
    private int deliveredChars;

    public Rdv3OperationLog(string ledgerPath, string machine, string userName, string spool)
    {
        path = PathFor(ledgerPath, machine);
        spoolPath = spool;
        deliveredPath = spool + ".delivered";
        host = (machine == null) ? "" : machine;
        user = (userName == null) ? "" : userName;
    }

    public string Path { get { return path; } }
    public string SpoolPath { get { return spoolPath; } }

    // <ledger stem><infix><terminal>.csv next to the ledger.
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
            .Replace("{state}", (initial == null || initial.Text == null) ? "" : initial.Text);
    }

    public static string DeleteDetail(Rdv3ProcessJobDef job, int deleted)
    {
        return Rdv3Text.OpDeleteDetailFmt.Replace("{job}", JobName(job)).Replace("{n}", N(deleted));
    }

    // Counts every row whose stored state changed, by the state it was changed
    // to, and names each configured state: the states that are not the initial
    // one first, the initial one last. Nothing is summed into a guessed state.
    public static string SendDetail(Rdv3WorkState work, string[] before, string[] after)
    {
        int[] counts = new int[work.States.Count];
        int rows = Math.Min((before == null) ? 0 : before.Length, (after == null) ? 0 : after.Length);
        for (int i = 0; i < rows; i++)
        {
            if (string.Equals(before[i], after[i], StringComparison.Ordinal)) { continue; }
            for (int s = 0; s < work.States.Count; s++)
            {
                if (string.Equals(work.States[s].Stored, after[i], StringComparison.Ordinal)) { counts[s]++; break; }
            }
        }
        StringBuilder sb = new StringBuilder();
        for (int pass = 0; pass < 2; pass++)
        {
            for (int s = 0; s < work.States.Count; s++)
            {
                bool isInitial = string.Equals(work.States[s].Id, work.Initial, StringComparison.Ordinal);
                if (isInitial != (pass == 1)) { continue; }
                if (sb.Length > 0) { sb.Append(Rdv3Text.OpSendDetailSeparator); }
                sb.Append(Rdv3Text.OpSendDetailItemFmt.Replace("{state}", work.States[s].Text ?? work.States[s].Id).Replace("{n}", N(counts[s])));
            }
        }
        return sb.ToString();
    }

    private static string JobName(Rdv3ProcessJobDef job)
    {
        if (job == null) { return ""; }
        return (job.Name != null && job.Name.Length > 0) ? job.Name : job.Id;
    }

    private static string N(int value) { return value.ToString("N0", CultureInfo.InvariantCulture); }

    // Appends one line for a ledger replacement this terminal has just made.
    // Returns null when the line reached the shared file and the spool is
    // settled; otherwise a message saying what happened instead. Never throws.
    public string Record(string operation, int rows, string detail)
    {
        string line = Line(DateTime.Now, operation, rows, detail);
        lock (gate)
        {
            bool delivered;
            string outcome = TryAppend(line, out delivered);
            if (delivered) { return outcome; }
            try { AppendText(spoolPath, line); }
            catch (Exception ex) { outcome = "lost: " + outcome + "; spool: " + ex.Message; }
            return outcome;
        }
    }

    public static string FailureNotice(string outcome)
    {
        return outcome != null && outcome.StartsWith("lost: ", StringComparison.Ordinal)
            ? "台帳への保存は完了しました。操作ログと控えの保存に失敗しました。この操作の再送信・再実行は不要です。\n" + outcome
            : null;
    }

    // Re-sends spooled lines, oldest first. Null when nothing is left to send
    // and the spool is settled.
    public string Flush()
    {
        lock (gate)
        {
            bool delivered;
            return TryAppend("", out delivered);
        }
    }

    public bool HasSpool()
    {
        try
        {
            string all, pending;
            ReadSpool(out all, out pending);
            return pending.Length > 0;
        }
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

    // all: the whole spool file; pending: the part of it not yet delivered.
    private void ReadSpool(out string all, out string pending)
    {
        all = File.Exists(spoolPath) ? File.ReadAllText(spoolPath, Encoding.UTF8) : "";
        int skip = deliveredChars;
        try
        {
            int recorded;
            if (File.Exists(deliveredPath) && int.TryParse(File.ReadAllText(deliveredPath, Encoding.ASCII).Trim(),
                NumberStyles.Integer, CultureInfo.InvariantCulture, out recorded) && recorded > skip) { skip = recorded; }
        }
        catch (Exception) { }
        // A spool shorter than the delivered length was replaced meanwhile:
        // nothing of it is known to be delivered.
        if (skip > all.Length) { skip = 0; }
        pending = all.Substring(skip);
    }

    // delivered: true when the text (spooled lines and the new line) reached the
    // shared file, even if the spool could not be settled afterwards.
    private string TryAppend(string line, out bool delivered)
    {
        delivered = false;
        string all, pending;
        try { ReadSpool(out all, out pending); }
        catch (Exception ex) { return "spooled: spool unreadable: " + ex.Message; }
        string text = pending + line;
        if (text.Length == 0)
        {
            if (all.Length == 0) { return null; }
            // everything in the spool is delivered; only its removal is outstanding
            string outstanding = ClearSpool(all);
            return outstanding == null ? null : "written; " + outstanding;
        }
        Exception last = null;
        for (int attempt = 0; attempt < Attempts; attempt++)
        {
            try
            {
                AppendShared(text);
                delivered = true;
                string cleared = (all.Length == 0) ? null : ClearSpool(all);
                return cleared == null ? null : "written; " + cleared;
            }
            catch (IOException ex) { last = ex; }
            catch (UnauthorizedAccessException ex) { last = ex; }
            if (attempt + 1 < Attempts) { Thread.Sleep(RetryMs); }
        }
        return "spooled: " + last.Message;
    }

    // Removes a delivered spool: delete, else truncate, else record how much of
    // it is delivered so the next send skips it. Null when the spool is gone.
    private string ClearSpool(string delivered)
    {
        Exception last = null;
        try
        {
            File.Delete(spoolPath);
            if (File.Exists(deliveredPath)) { File.Delete(deliveredPath); }
            deliveredChars = 0;
            return null;
        }
        catch (Exception ex) { last = ex; }
        try
        {
            using (FileStream file = new FileStream(spoolPath, FileMode.Truncate, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete))
            { file.Flush(true); }
            if (File.Exists(deliveredPath)) { try { File.Delete(deliveredPath); } catch (Exception) { } }
            deliveredChars = 0;
            return null;
        }
        catch (Exception ex) { last = ex; }
        deliveredChars = delivered.Length;
        try
        {
            File.WriteAllText(deliveredPath, delivered.Length.ToString(CultureInfo.InvariantCulture), Encoding.ASCII);
            return "spool not cleared (" + last.Message + "); delivered length recorded";
        }
        catch (Exception ex) { return "spool not cleared (" + last.Message + "); delivered length kept in memory only: " + ex.Message; }
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

    private static void AppendText(string destination, string text)
    {
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destination));
        using (FileStream file = new FileStream(destination, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
        {
            byte[] body = Encoding.UTF8.GetBytes(text);
            file.Write(body, 0, body.Length);
            file.Flush(true);
        }
    }
}
