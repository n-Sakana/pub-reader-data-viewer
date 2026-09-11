// ============================================================================
// Rdv3Shared.cs -- local pending changes and shared-ledger coordination.
//
// A state button writes only the small, machine-local pending file.  The xlsx
// is touched by a send or an update, while a CreateNew lock file is present.
// After the xlsx has been atomically replaced, a one-line marker tells the
// other running copies which version exists without making them poll the xlsx.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

public sealed class Rdv3PendingEntry
{
    public string Identity = "";
    public string Stored = "";
    public string Digest = "";
    public string BaselineStored = "";
    public bool BaselineKnown;

    public Rdv3PendingEntry Copy()
    {
        Rdv3PendingEntry e = new Rdv3PendingEntry();
        e.Identity = Identity;
        e.Stored = Stored;
        e.Digest = Digest;
        e.BaselineStored = BaselineStored;
        e.BaselineKnown = BaselineKnown;
        return e;
    }
}

public sealed class Rdv3UnmatchedChange
{
    public string Identity = "";
    public string Reason = "";              // missing | changed | state-conflict | legacy
}

public sealed class Rdv3PendingApply
{
    public string[] States;
    public readonly List<string> Resolved = new List<string>();
    public readonly List<Rdv3UnmatchedChange> Unmatched = new List<Rdv3UnmatchedChange>();
    public int ToInitial;
    public int FromInitial;
}

public sealed class Rdv3PendingStore
{
    private const string Header = "RDV-PENDING-2";
    private const string LegacyHeader = "RDV-PENDING-1";
    private readonly string path;
    private Dictionary<string, Rdv3PendingEntry> entries =
        new Dictionary<string, Rdv3PendingEntry>(StringComparer.Ordinal);

    public Rdv3PendingStore(string file)
    {
        if (file == null || file.Length == 0) { throw new ArgumentException("pending file path is blank"); }
        path = file;
        Load();
    }

    public string Path { get { return path; } }
    public int Count { get { return entries.Count; } }

    public static string PathFor(string ledgerPath)
    {
        string root = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (root == null || root.Length == 0) { throw new IOException("LOCALAPPDATA is not available"); }
        string canonical = System.IO.Path.GetFullPath(ledgerPath).ToLowerInvariant();
        string key = DigestOf(canonical).Replace("/", "_").Replace("+", "-").TrimEnd('=');
        if (key.Length > 20) { key = key.Substring(0, 20); }
        return System.IO.Path.Combine(System.IO.Path.Combine(root, "ReaderDataViewer"), "pending-" + key + ".dat");
    }

    // The companion travels with the ledger; the pending values remain private
    // in this Windows user's local store. A copied ledger gets an independent
    // snapshot while a moved ledger continues using the same store.
    public static string LocationPathFor(string ledgerPath)
    {
        string current = System.IO.Path.GetFullPath(ledgerPath);
        string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrEmpty(local)) { throw new IOException("LOCALAPPDATA is not available"); }
        string directory = System.IO.Path.Combine(local, "ReaderDataViewer");
        string selected;
        string owner = Environment.MachineName + "\n" + System.Security.Principal.WindowsIdentity.GetCurrent().User.Value;
        string key = DigestOf(owner).Replace("/", "_").Replace("+", "-").Substring(0, 20);
        string companion = current + ".local-" + key + ".state";
        if (File.Exists(companion))
        {
            string[] fields = File.ReadAllLines(companion, new UTF8Encoding(false, true));
            if (fields.Length != 3 || fields[0] != "RDV-LOCAL-1"
                || !System.Text.RegularExpressions.Regex.IsMatch(fields[1], "^pending-[a-f0-9]{32}\\.dat$"))
            { throw new InvalidDataException("未送信データの保存先情報が読めません。台帳と未送信データは変更していません: " + companion); }
            string previous = new UTF8Encoding(false, true).GetString(Convert.FromBase64String(fields[2]));
            selected = System.IO.Path.Combine(directory, fields[1]);
            if (!File.Exists(selected))
            { throw new FileNotFoundException("参照先の未送信データが見つかりません。別の保存先へ切り替えず停止しました。", selected); }
            new Rdv3PendingStore(selected); // never replace an unreadable store with an empty one
            if (!Rdv3Files.Same(previous, current) && File.Exists(previous))
            {
                // The original still exists: copying must not share mutable
                // pending values with the original application folder.
                string copy = System.IO.Path.Combine(directory, "pending-" + Guid.NewGuid().ToString("N") + ".dat");
                AtomicWrite(copy, File.ReadAllText(selected, new UTF8Encoding(false, true)));
                selected = copy;
            }
        }
        else
        {
            // A new reference creates a new store. No historical path is searched.
            selected = System.IO.Path.Combine(directory, "pending-" + Guid.NewGuid().ToString("N") + ".dat");
            AtomicWrite(selected, Header + "\n");
        }
        string content = "RDV-LOCAL-1\n" + System.IO.Path.GetFileName(selected) + "\n"
            + Convert.ToBase64String(Encoding.UTF8.GetBytes(current)) + "\n";
        if (!File.Exists(companion) || File.ReadAllText(companion, Encoding.UTF8) != content)
        { AtomicWrite(companion, content); }
        return selected;
    }

    public List<Rdv3PendingEntry> Snapshot()
    {
        List<Rdv3PendingEntry> result = new List<Rdv3PendingEntry>();
        foreach (KeyValuePair<string, Rdv3PendingEntry> pair in entries) { result.Add(pair.Value.Copy()); }
        result.Sort(delegate(Rdv3PendingEntry a, Rdv3PendingEntry b)
        {
            return string.Compare(a.Identity, b.Identity, StringComparison.Ordinal);
        });
        return result;
    }

    public void Validate(Rdv3WorkState work)
    {
        foreach (KeyValuePair<string, Rdv3PendingEntry> pair in entries)
        {
            if (work.ByStored(pair.Value.Stored) == null)
            {
                throw new InvalidDataException("pending state is not defined: " + pair.Value.Stored);
            }
        }
    }

    public void Set(string identity, string stored, string line, string sharedStored)
    {
        if (identity == null || identity.Length == 0) { throw new InvalidDataException("pending identity is blank"); }
        Dictionary<string, Rdv3PendingEntry> next = CopyEntries();
        if (string.Equals(stored, sharedStored, StringComparison.Ordinal))
        {
            next.Remove(identity);
        }
        else
        {
            Rdv3PendingEntry e = new Rdv3PendingEntry();
            e.Identity = identity;
            e.Stored = (stored == null) ? "" : stored;
            e.Digest = DigestOf((line == null) ? "" : line);
            Rdv3PendingEntry before;
            // Editing a pending value must not silently rebase an old intent.
            if (entries.TryGetValue(identity, out before))
            {
                e.Digest = before.Digest;
                e.BaselineStored = before.BaselineStored;
                e.BaselineKnown = before.BaselineKnown;
            }
            else
            {
                e.BaselineStored = sharedStored ?? "";
                e.BaselineKnown = true;
            }
            next[identity] = e;
        }
        Commit(next);
    }

    public void Remove(List<string> identities)
    {
        if (identities == null || identities.Count == 0) { return; }
        Dictionary<string, Rdv3PendingEntry> next = CopyEntries();
        for (int i = 0; i < identities.Count; i++) { next.Remove(identities[i]); }
        Commit(next);
    }

    public string[] Overlay(string[] lines, string[] sharedStates, int identityColumn)
    { return Overlay(lines, sharedStates, new int[] { identityColumn }); }

    public string[] Overlay(string[] lines, string[] sharedStates, int[] identityColumn)
    {
        if (lines == null || sharedStates == null || lines.Length != sharedStates.Length)
        {
            throw new InvalidDataException("ledger lines and states do not match");
        }
        string[] result = (string[])sharedStates.Clone();
        if (entries.Count == 0) { return result; }
        Dictionary<string, int> rows = Rdv3Ledger.RowMap(lines, identityColumn, "ledger");
        foreach (KeyValuePair<string, Rdv3PendingEntry> pair in entries)
        {
            int row;
            if (!rows.TryGetValue(pair.Key, out row)) { continue; }
            if (!string.Equals(DigestOf(lines[row]), pair.Value.Digest, StringComparison.Ordinal)) { continue; }
            if (!string.Equals(sharedStates[row], pair.Value.Stored, StringComparison.Ordinal)
                && (!pair.Value.BaselineKnown || !string.Equals(sharedStates[row], pair.Value.BaselineStored, StringComparison.Ordinal)))
            { continue; }
            result[row] = pair.Value.Stored;
        }
        return result;
    }

    public Rdv3PendingApply PrepareSend(string[] lines, string[] sharedStates, int identityColumn, string initialStored)
    { return PrepareSend(lines, sharedStates, new int[] { identityColumn }, initialStored); }

    public Rdv3PendingApply PrepareSend(string[] lines, string[] sharedStates, int[] identityColumn, string initialStored)
    {
        Rdv3PendingApply result = new Rdv3PendingApply();
        if (lines == null || sharedStates == null || lines.Length != sharedStates.Length)
        {
            throw new InvalidDataException("ledger lines and states do not match");
        }
        result.States = (string[])sharedStates.Clone();
        Dictionary<string, int> rows = Rdv3Ledger.RowMap(lines, identityColumn, "ledger");
        List<Rdv3PendingEntry> pending = Snapshot();
        for (int i = 0; i < pending.Count; i++)
        {
            Rdv3PendingEntry e = pending[i];
            int row;
            if (!rows.TryGetValue(e.Identity, out row))
            {
                result.Unmatched.Add(Unmatched(e.Identity, "missing"));
                continue;
            }
            if (!string.Equals(DigestOf(lines[row]), e.Digest, StringComparison.Ordinal))
            {
                result.Unmatched.Add(Unmatched(e.Identity, "changed"));
                continue;
            }
            // An interrupted send may already have committed the XLSX. Retry is
            // idempotent; never rewrite or count a value that is already present.
            if (string.Equals(sharedStates[row], e.Stored, StringComparison.Ordinal))
            {
                result.Resolved.Add(e.Identity);
                continue;
            }
            if (!e.BaselineKnown)
            {
                result.Unmatched.Add(Unmatched(e.Identity, "legacy"));
                continue;
            }
            if (!string.Equals(sharedStates[row], e.BaselineStored, StringComparison.Ordinal))
            {
                result.Unmatched.Add(Unmatched(e.Identity, "state-conflict"));
                continue;
            }
            result.States[row] = e.Stored;
            result.Resolved.Add(e.Identity);
            if (string.Equals(e.Stored, initialStored, StringComparison.Ordinal)) { result.ToInitial++; }
            else { result.FromInitial++; }
        }
        return result;
    }

    public static string DigestOf(string value)
    {
        byte[] bytes = Encoding.UTF8.GetBytes((value == null) ? "" : value);
        using (SHA256 sha = SHA256.Create()) { return Convert.ToBase64String(sha.ComputeHash(bytes)); }
    }

    private static Rdv3UnmatchedChange Unmatched(string identity, string reason)
    {
        Rdv3UnmatchedChange u = new Rdv3UnmatchedChange();
        u.Identity = identity;
        u.Reason = reason;
        return u;
    }

    private Dictionary<string, Rdv3PendingEntry> CopyEntries()
    {
        Dictionary<string, Rdv3PendingEntry> copy = new Dictionary<string, Rdv3PendingEntry>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, Rdv3PendingEntry> pair in entries) { copy.Add(pair.Key, pair.Value.Copy()); }
        return copy;
    }

    private void Load()
    {
        if (!Rdv3Files.Exists(path)) { return; }
        string[] lines = File.ReadAllLines(path, new UTF8Encoding(false, true));
        if (lines.Length == 0 || (lines[0] != Header && lines[0] != LegacyHeader))
        {
            throw new InvalidDataException("pending file has an unknown format: " + path);
        }
        bool legacy = lines[0] == LegacyHeader;
        Dictionary<string, Rdv3PendingEntry> loaded = new Dictionary<string, Rdv3PendingEntry>(StringComparer.Ordinal);
        for (int i = 1; i < lines.Length; i++)
        {
            if (lines[i].Length == 0) { continue; }
            string[] cells = lines[i].Split('\t');
            if (cells.Length != (legacy ? 3 : 5)) { throw new InvalidDataException("pending file row is invalid: " + (i + 1).ToString(CultureInfo.InvariantCulture)); }
            Rdv3PendingEntry e = new Rdv3PendingEntry();
            try
            {
                e.Identity = Decode(cells[0]);
                e.Stored = Decode(cells[1]);
                e.Digest = cells[2];
                if (Convert.FromBase64String(e.Digest).Length != 32) { throw new FormatException("invalid digest"); }
                if (!legacy)
                {
                    if (cells[3] != "0" && cells[3] != "1") { throw new FormatException("invalid baseline flag"); }
                    e.BaselineKnown = cells[3] == "1";
                    e.BaselineStored = Decode(cells[4]);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidDataException("pending file row cannot be decoded: " + (i + 1).ToString(CultureInfo.InvariantCulture), ex);
            }
            if (e.Identity.Length == 0 || e.Digest.Length == 0 || loaded.ContainsKey(e.Identity))
            {
                throw new InvalidDataException("pending file row is invalid: " + (i + 1).ToString(CultureInfo.InvariantCulture));
            }
            loaded.Add(e.Identity, e);
        }
        entries = loaded;
    }

    private void Commit(Dictionary<string, Rdv3PendingEntry> next)
    {
        StringBuilder text = new StringBuilder();
        text.Append(Header).Append("\r\n");
        List<Rdv3PendingEntry> ordered = new List<Rdv3PendingEntry>(next.Values);
        ordered.Sort(delegate(Rdv3PendingEntry a, Rdv3PendingEntry b)
        {
            return string.Compare(a.Identity, b.Identity, StringComparison.Ordinal);
        });
        for (int i = 0; i < ordered.Count; i++)
        {
            text.Append(Encode(ordered[i].Identity)).Append('\t');
            text.Append(Encode(ordered[i].Stored)).Append('\t');
            text.Append(ordered[i].Digest).Append('\t');
            text.Append(ordered[i].BaselineKnown ? "1" : "0").Append('\t');
            text.Append(Encode(ordered[i].BaselineStored)).Append("\r\n");
        }
        AtomicWrite(path, text.ToString());
        entries = next;
    }

    private static string Encode(string value)
    {
        return Convert.ToBase64String(Encoding.UTF8.GetBytes((value == null) ? "" : value));
    }

    private static string Decode(string value)
    {
        return new UTF8Encoding(false, true).GetString(Convert.FromBase64String(value));
    }

    internal static void AtomicWrite(string target, string text)
    {
        string dir = System.IO.Path.GetDirectoryName(target);
        if (dir != null && dir.Length > 0 && !Directory.Exists(dir)) { Directory.CreateDirectory(dir); }
        string temp = target + ".tmp-" + System.Diagnostics.Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture)
            + "-" + DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture);
        try
        {
            File.WriteAllText(temp, text, new UTF8Encoding(false));
            if (File.Exists(target)) { File.Replace(temp, target, null); }
            else { File.Move(temp, target); }
        }
        finally
        {
            if (File.Exists(temp)) { try { File.Delete(temp); } catch (Exception) { } }
        }
    }
}

public sealed class Rdv3SharedMarker
{
    public string FileStamp; // local only; captured by the writer under the ledger lock
    public long Version;
    public string Host = "";
    public string User = "";
    public string WriterId = "";
    public long WrittenUtcTicks;
    public int Rows;
    public string Kind = "";                 // update | send
    public int FromInitial;
    public int ToInitial;
}

public sealed class Rdv3LockInfo
{
    public string Host = "";
    public string User = "";
    public long TakenUtcTicks;

    public int AgeMinutes
    {
        get
        {
            long age = DateTime.UtcNow.Ticks - TakenUtcTicks;
            if (age <= 0) { return 0; }
            double minutes = TimeSpan.FromTicks(age).TotalMinutes;
            return (int)Math.Floor(minutes);
        }
    }
}

public sealed class Rdv3LedgerLock : IDisposable
{
    private readonly FileStream lease;
    private bool released;

    internal Rdv3LedgerLock(FileStream heldLease)
    {
        lease = heldLease;
    }

    public void Release()
    {
        if (released) { return; }
        released = true;
        lease.Dispose();
    }

    public void Dispose()
    {
        if (released) { return; }
        released = true;
        try { lease.Dispose(); } catch (Exception) { }
    }
}

public sealed class Rdv3SharedFiles
{
    private const string LockHeader = "RDV-LOCK-1";
    private const string MarkerHeader = "RDV-MARKER-1";
    private readonly string lockPath;
    private readonly string markerPath;
    private readonly string host;
    private readonly string user;
    private readonly string writerId;
    private readonly Rdv3OperationLog operations;

    public Rdv3SharedFiles(string ledgerPath, string machine, string userName, string instance)
        : this(ledgerPath, machine, userName, instance, null) { }

    public Rdv3SharedFiles(string ledgerPath, string machine, string userName, string instance, string operationSpool)
    {
        lockPath = ledgerPath + ".lock";
        markerPath = ledgerPath + ".version";
        host = (machine == null) ? "" : machine;
        user = (userName == null) ? "" : userName;
        writerId = (instance == null) ? "" : instance;
        operations = new Rdv3OperationLog(ledgerPath, host, user,
            (operationSpool == null) ? Rdv3OperationLog.SpoolPathFor(ledgerPath) : operationSpool);
    }

    public string LockPath { get { return lockPath; } }
    public string MarkerPath { get { return markerPath; } }
    public string WriterId { get { return writerId; } }
    public Rdv3OperationLog Operations { get { return operations; } }

    // Records that this terminal has just replaced the shared ledger. The
    // ledger is already saved, so nothing here may throw back into the job:
    // a line that cannot be appended is spooled and the reason is returned.
    public string RecordOperation(string operation, int rows, string detail)
    {
        try { return operations.Record(operation, rows, detail); }
        catch (Exception ex) { return "lost: " + ex.Message; }
    }

    public Rdv3LedgerLock TryAcquire(out Rdv3LockInfo owner)
    {
        owner = null;
        FileStream stream = null;
        try
        {
            string dir = System.IO.Path.GetDirectoryName(lockPath);
            if (dir != null && dir.Length > 0 && !Directory.Exists(dir)) { Directory.CreateDirectory(dir); }
            stream = new FileStream(lockPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read,
                4096, FileOptions.DeleteOnClose);
            using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false), 1024, true))
            {
                writer.Write(LockHeader);
                writer.Write('\t'); writer.Write(Encode(host));
                writer.Write('\t'); writer.Write(Encode(user));
                writer.Write('\t'); writer.Write(DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture));
                writer.Write("\r\n");
                writer.Flush();
                stream.Flush();
            }
            return new Rdv3LedgerLock(stream);
        }
        catch (IOException error)
        {
            if (stream != null) { try { stream.Dispose(); } catch (Exception) { } }
            // The winner can close its DeleteOnClose lease between CreateNew
            // rejecting us and this catch. File.Exists cannot identify that
            // race. Retry only a creation collision, never a failed lock write.
            int code = error.HResult & 0xffff;
            if (stream != null || (code != 80 && code != 183 && code != 32)) { throw; }
            owner = ReadLock();
            return null;
        }
        catch
        {
            if (stream != null) { try { stream.Dispose(); } catch (Exception) { } }
            throw;
        }
    }

    public bool TryRemoveStaleLock(int staleMs, out long ageMs)
    {
        ageMs = 0;
        DateTime modified;
        try { modified = File.GetLastWriteTimeUtc(lockPath); }
        catch (Exception) { return false; }
        double elapsed = (DateTime.UtcNow - modified).TotalMilliseconds;
        if (elapsed <= 0 || elapsed < staleMs) { return false; }
        ageMs = (elapsed >= long.MaxValue) ? long.MaxValue : (long)Math.Floor(elapsed);
        try
        {
            File.Delete(lockPath);
            return !File.Exists(lockPath);
        }
        catch (IOException) { return false; }
        catch (UnauthorizedAccessException) { return false; }
    }

    public Rdv3SharedMarker ReadMarker()
    {
        if (!File.Exists(markerPath)) { return null; }
        string line = ReadShared(markerPath).TrimEnd('\r', '\n');
        string[] cells = line.Split('\t');
        if (cells.Length != 10 || !string.Equals(cells[0], MarkerHeader, StringComparison.Ordinal))
        {
            throw new InvalidDataException("shared marker has an unknown format: " + markerPath);
        }
        Rdv3SharedMarker m = new Rdv3SharedMarker();
        try
        {
            m.Version = long.Parse(cells[1], CultureInfo.InvariantCulture);
            m.Host = Decode(cells[2]);
            m.User = Decode(cells[3]);
            m.WriterId = Decode(cells[4]);
            m.WrittenUtcTicks = long.Parse(cells[5], CultureInfo.InvariantCulture);
            m.Rows = int.Parse(cells[6], CultureInfo.InvariantCulture);
            m.Kind = cells[7];
            m.FromInitial = int.Parse(cells[8], CultureInfo.InvariantCulture);
            m.ToInitial = int.Parse(cells[9], CultureInfo.InvariantCulture);
        }
        catch (Exception ex) { throw new InvalidDataException("shared marker cannot be decoded: " + markerPath, ex); }
        if (m.Version <= 0 || m.WrittenUtcTicks <= 0 || m.Rows < 0 || m.FromInitial < 0 || m.ToInitial < 0
            || (m.Kind != "update" && m.Kind != "send"))
        {
            throw new InvalidDataException("shared marker has invalid values: " + markerPath);
        }
        return m;
    }

    public Rdv3SharedMarker WriteMarker(string kind, int rows, int fromInitial, int toInitial)
    {
        Rdv3SharedMarker previous = ReadMarker();
        Rdv3SharedMarker m = new Rdv3SharedMarker();
        m.Version = (previous == null) ? 1 : checked(previous.Version + 1);
        m.Host = host;
        m.User = user;
        m.WriterId = writerId;
        m.WrittenUtcTicks = DateTime.UtcNow.Ticks;
        m.Rows = rows;
        m.Kind = kind;
        m.FromInitial = fromInitial;
        m.ToInitial = toInitial;
        StringBuilder line = new StringBuilder();
        line.Append(MarkerHeader).Append('\t').Append(m.Version.ToString(CultureInfo.InvariantCulture));
        line.Append('\t').Append(Encode(m.Host)).Append('\t').Append(Encode(m.User));
        line.Append('\t').Append(Encode(m.WriterId)).Append('\t').Append(m.WrittenUtcTicks.ToString(CultureInfo.InvariantCulture));
        line.Append('\t').Append(rows.ToString(CultureInfo.InvariantCulture)).Append('\t').Append(kind);
        line.Append('\t').Append(fromInitial.ToString(CultureInfo.InvariantCulture));
        line.Append('\t').Append(toInitial.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
        Rdv3PendingStore.AtomicWrite(markerPath, line.ToString());
        m.FileStamp = Rdv3Files.Stamp(markerPath.Substring(0, markerPath.Length - ".version".Length));
        return m;
    }

    private Rdv3LockInfo ReadLock()
    {
        Rdv3LockInfo info = new Rdv3LockInfo();
        try
        {
            string line = ReadShared(lockPath).TrimEnd('\r', '\n');
            string[] cells = line.Split('\t');
            if (cells.Length == 4 && string.Equals(cells[0], LockHeader, StringComparison.Ordinal))
            {
                info.Host = Decode(cells[1]);
                info.User = Decode(cells[2]);
                info.TakenUtcTicks = long.Parse(cells[3], CultureInfo.InvariantCulture);
                return info;
            }
        }
        catch (Exception) { }
        try { info.TakenUtcTicks = File.GetLastWriteTimeUtc(lockPath).Ticks; }
        catch (Exception) { info.TakenUtcTicks = DateTime.UtcNow.Ticks; }
        return info;
    }

    // A shared file is read while another copy of the app may be replacing it
    // (write-temp-then-replace). The open allows that replace, and the copy
    // that is reading finishes the version it opened.
    private static string ReadShared(string path)
    {
        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (StreamReader reader = new StreamReader(stream, new UTF8Encoding(false)))
        {
            return reader.ReadToEnd();
        }
    }

    private static string Encode(string value)
    {
        return Convert.ToBase64String(Encoding.UTF8.GetBytes((value == null) ? "" : value));
    }

    private static string Decode(string value)
    {
        return Encoding.UTF8.GetString(Convert.FromBase64String(value));
    }
}
