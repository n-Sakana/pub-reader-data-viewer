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

    public Rdv3PendingEntry Copy()
    {
        Rdv3PendingEntry e = new Rdv3PendingEntry();
        e.Identity = Identity;
        e.Stored = Stored;
        e.Digest = Digest;
        return e;
    }
}

public sealed class Rdv3UnmatchedChange
{
    public string Identity = "";
    public string Reason = "";              // missing | changed
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
    private const string Header = "RDV-PENDING-1";
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
            result[row] = pair.Value.Stored;
        }
        return result;
    }

    public Rdv3PendingApply PrepareSend(string[] lines, string[] sharedStates, int identityColumn, string initialStored)
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
        if (!File.Exists(path)) { return; }
        string[] lines = File.ReadAllLines(path, new UTF8Encoding(false));
        if (lines.Length == 0 || !string.Equals(lines[0], Header, StringComparison.Ordinal))
        {
            throw new InvalidDataException("pending file has an unknown format: " + path);
        }
        Dictionary<string, Rdv3PendingEntry> loaded = new Dictionary<string, Rdv3PendingEntry>(StringComparer.Ordinal);
        for (int i = 1; i < lines.Length; i++)
        {
            if (lines[i].Length == 0) { continue; }
            string[] cells = lines[i].Split('\t');
            if (cells.Length != 3) { throw new InvalidDataException("pending file row is invalid: " + (i + 1).ToString(CultureInfo.InvariantCulture)); }
            Rdv3PendingEntry e = new Rdv3PendingEntry();
            try
            {
                e.Identity = Decode(cells[0]);
                e.Stored = Decode(cells[1]);
                e.Digest = cells[2];
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
            text.Append(ordered[i].Digest).Append("\r\n");
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
        return Encoding.UTF8.GetString(Convert.FromBase64String(value));
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
    private readonly string path;
    private bool released;

    internal Rdv3LedgerLock(string file) { path = file; }

    public void Release()
    {
        if (released) { return; }
        File.Delete(path);
        released = true;
    }

    public void Dispose()
    {
        if (released) { return; }
        try { File.Delete(path); } catch (Exception) { }
        released = true;
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

    public Rdv3SharedFiles(string ledgerPath, string machine, string userName, string instance)
    {
        lockPath = ledgerPath + ".lock";
        markerPath = ledgerPath + ".version";
        host = (machine == null) ? "" : machine;
        user = (userName == null) ? "" : userName;
        writerId = (instance == null) ? "" : instance;
    }

    public string LockPath { get { return lockPath; } }
    public string MarkerPath { get { return markerPath; } }
    public string WriterId { get { return writerId; } }

    public Rdv3LedgerLock TryAcquire(out Rdv3LockInfo owner)
    {
        owner = null;
        try
        {
            string dir = System.IO.Path.GetDirectoryName(lockPath);
            if (dir != null && dir.Length > 0 && !Directory.Exists(dir)) { Directory.CreateDirectory(dir); }
            using (FileStream stream = new FileStream(lockPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false)))
            {
                writer.Write(LockHeader);
                writer.Write('\t'); writer.Write(Encode(host));
                writer.Write('\t'); writer.Write(Encode(user));
                writer.Write('\t'); writer.Write(DateTime.UtcNow.Ticks.ToString(CultureInfo.InvariantCulture));
                writer.Write("\r\n");
                writer.Flush();
                stream.Flush();
            }
            return new Rdv3LedgerLock(lockPath);
        }
        catch (IOException)
        {
            if (!File.Exists(lockPath)) { throw; }
            owner = ReadLock();
            return null;
        }
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
