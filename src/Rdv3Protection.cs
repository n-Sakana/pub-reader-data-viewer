using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

// Business definitions omit presentation and locations, but retain input format
// and ordered operations. JSON member order and whitespace have no meaning.
public static class Rdv3BusinessDefinition
{
    public static void CheckCurrent(Rdv3Config loaded)
    {
        Rdv3Config current = Rdv3Config.Load(loaded.SourcePath);
        if (current.Data.Definition != loaded.Data.Definition
            || Rdv3Files.LegacyStorageContract(current.Data, current.Screen.Work)
                != Rdv3Files.LegacyStorageContract(loaded.Data, loaded.Screen.Work))
        { throw new InvalidDataException(Rdv3Text.BusinessDefinitionMismatch); }
    }

    public static string Normalize(Rdv3Json data, bool includeProtection)
    { return Canonical(data, "data", includeProtection); }

    private static string Canonical(Rdv3Json node, string path, bool policy)
    {
        if (node.Kind == Rdv3Json.TObject)
        {
            List<string> keys = new List<string>(node.Members.Keys);
            keys.Sort(StringComparer.Ordinal);
            List<string> parts = new List<string>();
            foreach (string key in keys)
            {
                bool input = path.StartsWith("data.tables.", StringComparison.Ordinal)
                    && path.Split('.').Length == 3 || path == "data.jobs[].inputs[]";
                if ((path == "data" && key == "labels") || (path == "data.ledger" && key == "search")
                    || (path == "data.jobs[]" && key == "name") || (input && key == "label")
                    || (!policy && path == "data.ledger" && key == "protectStates")) { continue; }
                Rdv3Json value = node.Members[key];
                string text;
                if (input && key == "file")
                { text = Rdv3Json.Quote(System.IO.Path.GetExtension(value.Str).ToLowerInvariant()); }
                else if ((input && key == "key") || (path == "data.ledger" && key == "identity"))
                { text = value.IsArray ? Canonical(value, path + "." + key, policy) : "[" + Canonical(value, path + "." + key, policy) + "]"; }
                else { text = Canonical(value, path + "." + key, policy); }
                parts.Add(Rdv3Json.Quote(key) + ":" + text);
            }
            return "{" + string.Join(",", parts.ToArray()) + "}";
        }
        if (node.Kind == Rdv3Json.TArray)
        {
            List<string> parts = new List<string>();
            foreach (Rdv3Json child in node.Items) { parts.Add(Canonical(child, path + "[]", policy)); }
            if (path == "data.ledger.protectStates") { parts.Sort(StringComparer.Ordinal); }
            return "[" + string.Join(",", parts.ToArray()) + "]";
        }
        if (node.Kind == Rdv3Json.TString) { return Rdv3Json.Quote(node.Str); }
        if (node.Kind == Rdv3Json.TNumber) { return node.Num.ToString("R", CultureInfo.InvariantCulture); }
        if (node.Kind == Rdv3Json.TBool) { return node.Flag ? "true" : "false"; }
        return "null";
    }

    public static string Bound(Rdv3Data data)
    {
        List<Rdv3TableDef> tables = new List<Rdv3TableDef>(data.Tables);
        tables.Sort(delegate(Rdv3TableDef a, Rdv3TableDef b) { return StringComparer.Ordinal.Compare(a.Id, b.Id); });
        StringBuilder result = new StringBuilder("{\"configuration\":").Append(data.Definition);
        result.Append(",\"headers\":{");
        for (int i = 0; i < tables.Count; i++)
        {
            if (i > 0) { result.Append(','); }
            result.Append(Rdv3Json.Quote(tables[i].Id)).Append(':').Append(Rdv3WebJson.S(tables[i].Head));
        }
        result.Append("},\"fileInputs\":{");
        List<string> entries = new List<string>();
        foreach (Rdv3ProcessJobDef job in data.Jobs)
        {
            foreach (Rdv3ProcessInputDef input in job.Inputs)
            {
                if (!input.IsTable) { entries.Add(Rdv3Json.Quote(job.Id + "/" + input.Id) + ":" + Rdv3WebJson.S(input.Head)); }
            }
        }
        entries.Sort(StringComparer.Ordinal);
        return result.Append(string.Join(",", entries.ToArray())).Append("}}").ToString();
    }

    public static void BindFileInputs(Rdv3Data data, string directory)
    {
        foreach (Rdv3ProcessJobDef job in data.Jobs)
        {
            foreach (Rdv3ProcessInputDef input in job.Inputs)
            {
                if (!input.IsTable)
                { input.Head = Rdv3Table.ReadHead(Rdv3Files.Full(input.File, directory), input.Enc, input.EncodingSetting,
                    data.SourceReferences(input), input.HeaderRow, input.Delimiter, input.Sheet); }
            }
        }
    }
}

public sealed class Rdv3DeletedRecord
{
    public string Line;
    public string State;
    public string DeletedAt;
    internal string ArchiveName = "";
}

internal sealed class Rdv3ArchiveReference
{
    internal string Name, Digest, Payload;
}

// Archive batches are immutable. The main workbook atomically commits their
// references only after the separate workbooks have been saved successfully.
public sealed class Rdv3LedgerProtection
{
    public string Definition = "";
    public bool Legacy;
    public readonly List<Rdv3DeletedRecord> Deleted = new List<Rdv3DeletedRecord>();
    internal readonly List<Rdv3ArchiveReference> Archives = new List<Rdv3ArchiveReference>();

    public static Rdv3LedgerProtection Create(Rdv3Data data)
    {
        foreach (Rdv3TableDef table in data.Tables)
        { if (table.Head == null) { throw new InvalidOperationException("Bind input headers before creating a ledger"); } }
        return new Rdv3LedgerProtection { Definition = Rdv3BusinessDefinition.Bound(data) };
    }

    public void RequireWritable(Rdv3Data data)
    {
        if (Legacy) { throw new InvalidDataException(Rdv3Text.LegacyLedgerNeedsMigration); }
        if (Definition != Rdv3BusinessDefinition.Bound(data))
        { throw new InvalidDataException(Rdv3Text.BusinessDefinitionMismatch); }
    }

    internal static Rdv3LedgerProtection Read(ZipArchive zip, string expected)
    {
        ZipArchiveEntry entry = Rdv3Xlsx.MetadataPart(zip, "protection");
        if (entry == null) { return new Rdv3LedgerProtection { Legacy = true }; }
        XmlDocument doc = new XmlDocument();
        doc.XmlResolver = null;
        using (XmlReader reader = XmlReader.Create(entry.Open(), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
        { doc.Load(reader); }
        XmlElement root = doc.DocumentElement;
        if (root == null || root.Name != "protection" || (root.GetAttribute("version") != "1" && root.GetAttribute("version") != "2"))
        { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
        XmlNode definition = root.SelectSingleNode("definition");
        XmlNode deleted = root.SelectSingleNode("deleted");
        if (definition == null || deleted == null || definition.InnerText.Length == 0)
        { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
        Rdv3LedgerProtection result = new Rdv3LedgerProtection { Definition = definition.InnerText };
        if (expected != null && result.Definition != expected)
        { throw new InvalidDataException(Rdv3Text.BusinessDefinitionMismatch); }
        foreach (XmlNode row in deleted.ChildNodes)
        {
            XmlNode line = row.SelectSingleNode("line"), state = row.SelectSingleNode("state"), at = row.SelectSingleNode("at");
            if (row.Name != "row" || line == null || state == null || at == null)
            { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
            result.Deleted.Add(new Rdv3DeletedRecord { Line = line.InnerText, State = state.InnerText, DeletedAt = at.InnerText });
        }
        if (root.GetAttribute("version") == "2")
        {
            XmlNode archives = root.SelectSingleNode("archives");
            if (archives == null || result.Deleted.Count != 0) { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
            HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XmlNode file in archives.ChildNodes)
            {
                XmlElement item = file as XmlElement;
                if (item == null || item.Name != "file") { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
                string name = item.GetAttribute("name"), digest = item.GetAttribute("sha256");
                if (name.Length == 0 || Path.GetFileName(name) != name || name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                    || !name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || digest.Length != 64 || !names.Add(name))
                { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
                result.Archives.Add(new Rdv3ArchiveReference { Name = name, Digest = digest });
            }
        }
        return result;
    }

    internal void Write(ZipArchive zip)
    {
        if (Legacy || Definition.Length == 0) { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
        ZipArchiveEntry entry = zip.CreateEntry("rdv-protection.xml", CompressionLevel.Fastest);
        using (XmlWriter writer = XmlWriter.Create(entry.Open(), new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = true, NewLineHandling = NewLineHandling.Entitize }))
        {
            writer.WriteStartElement("protection"); writer.WriteAttributeString("version", "2");
            writer.WriteElementString("definition", Definition); writer.WriteStartElement("deleted");
            writer.WriteEndElement();
            writer.WriteStartElement("archives");
            foreach (Rdv3ArchiveReference file in Archives)
            {
                writer.WriteStartElement("file");
                writer.WriteAttributeString("name", file.Name); writer.WriteAttributeString("sha256", file.Digest);
                writer.WriteEndElement();
            }
            writer.WriteEndElement(); writer.WriteEndElement();
        }
    }

    private static string ArchiveDirectory(string ledger)
    { return Path.Combine(Path.GetDirectoryName(Path.GetFullPath(ledger)), "archived"); }

    private static string Digest(Stream stream)
    {
        using (System.Security.Cryptography.SHA256 hash = System.Security.Cryptography.SHA256.Create())
        { return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant(); }
    }

    private static string Payload(IList<Rdv3DeletedRecord> rows)
    {
        using (MemoryStream memory = new MemoryStream())
        {
            using (BinaryWriter writer = new BinaryWriter(memory, Encoding.UTF8, true))
            { foreach (Rdv3DeletedRecord row in rows) { writer.Write(row.Line); writer.Write(row.State); writer.Write(row.DeletedAt); } }
            memory.Position = 0; return Digest(memory);
        }
    }

    internal void LoadArchives(string ledger, string[] head, string stateHead)
    {
        string[] archiveHead = new string[head.Length + 1];
        Array.Copy(head, archiveHead, head.Length); archiveHead[head.Length] = "削除日時";
        foreach (Rdv3ArchiveReference reference in Archives)
        {
            string file = Path.Combine(ArchiveDirectory(ledger), reference.Name);
            string[] lines, states; string warning; Rdv3LedgerProtection ignored;
            // Keep the immutable batch locked against edits while verifying and reading.
            using (FileStream stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (Digest(stream) != reference.Digest) { throw new InvalidDataException("削除済み台帳が変更されています: " + file); }
                Rdv3Xlsx.ReadProtected(file, archiveHead, stateHead, out lines, out states, out warning, null, null, null, out ignored, false);
            }
            if (warning.Length > 0) { throw new InvalidDataException(warning); }
            List<Rdv3DeletedRecord> batch = new List<Rdv3DeletedRecord>();
            for (int i = 0; i < lines.Length; i++)
            {
                int split = lines[i].LastIndexOf('\t');
                if (split < 0) { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
                batch.Add(new Rdv3DeletedRecord { Line = lines[i].Substring(0, split), State = states[i], DeletedAt = lines[i].Substring(split + 1), ArchiveName = reference.Name });
            }
            reference.Payload = Payload(batch); Deleted.AddRange(batch);
        }
    }

    internal Rdv3LedgerProtection PrepareArchives(string ledger, string[] head, string stateHead, string runId, List<string> created)
    {
        Rdv3LedgerProtection prepared = new Rdv3LedgerProtection { Definition = Definition, Legacy = Legacy };
        Dictionary<string, List<Rdv3DeletedRecord>> groups = new Dictionary<string, List<Rdv3DeletedRecord>>(StringComparer.Ordinal);
        foreach (Rdv3DeletedRecord row in Deleted)
        {
            List<Rdv3DeletedRecord> batch;
            if (!groups.TryGetValue(row.ArchiveName, out batch)) { batch = new List<Rdv3DeletedRecord>(); groups.Add(row.ArchiveName, batch); }
            batch.Add(row);
        }
        foreach (KeyValuePair<string, List<Rdv3DeletedRecord>> group in groups)
        {
            string payload = Payload(group.Value);
            Rdv3ArchiveReference reference = Archives.Find(delegate(Rdv3ArchiveReference item) { return item.Name == group.Key && item.Payload == payload; });
            if (reference != null)
            {
                using (FileStream stream = new FileStream(Path.Combine(ArchiveDirectory(ledger), reference.Name), FileMode.Open, FileAccess.Read, FileShare.Read))
                { if (Digest(stream) != reference.Digest) { throw new InvalidDataException("削除済み台帳が変更されています: " + reference.Name); } }
                prepared.Archives.Add(reference);
            }
            else
            {
                string directory = ArchiveDirectory(ledger); Directory.CreateDirectory(directory);
                string name = Path.GetFileNameWithoutExtension(ledger) + "_削除済み_" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N") + ".xlsx";
                string file = Path.Combine(directory, name);
                string[] archiveHead = new string[head.Length + 1]; Array.Copy(head, archiveHead, head.Length); archiveHead[head.Length] = "削除日時";
                string[] lines = new string[group.Value.Count], states = new string[lines.Length];
                for (int i = 0; i < lines.Length; i++) { lines[i] = group.Value[i].Line + "\t" + group.Value[i].DeletedAt; states[i] = group.Value[i].State; }
                Rdv3Xlsx.Write(file, archiveHead, stateHead, lines, states, runId, null, null, false);
                created.Add(file);
                using (FileStream stream = File.OpenRead(file)) { reference = new Rdv3ArchiveReference { Name = name, Digest = Digest(stream), Payload = payload }; }
                prepared.Archives.Add(reference);
            }
            foreach (Rdv3DeletedRecord row in group.Value)
            { prepared.Deleted.Add(new Rdv3DeletedRecord { Line = row.Line, State = row.State, DeletedAt = row.DeletedAt, ArchiveName = reference.Name }); }
        }
        return prepared;
    }

    internal void AdoptArchives(Rdv3LedgerProtection saved)
    { Archives.Clear(); Archives.AddRange(saved.Archives); Deleted.Clear(); Deleted.AddRange(saved.Deleted); }

    public void Validate(Rdv3Data data, Rdv3WorkState work, string[] liveLines)
    {
        HashSet<string> identities = new HashSet<string>(StringComparer.Ordinal);
        foreach (string line in liveLines) { identities.Add(Rdv3Key.FromLine(line, data.IdentityCols)); }
        foreach (Rdv3DeletedRecord row in Deleted)
        {
            if (row.Line == null || row.Line.Split('\t').Length != data.Columns.Count || work.ByStored(row.State) == null)
            { throw new InvalidDataException(Rdv3Text.ProtectionInvalid); }
            string id = Rdv3Key.FromLine(row.Line, data.IdentityCols);
            if (id.Length == 0 || !identities.Add(id)) { throw new InvalidDataException(Rdv3Text.ArchiveIdentityConflict); }
        }
    }

    public void ArchiveRemoved(Rdv3Data data, string[] before, string[] states, string[] after)
    {
        HashSet<string> remaining = new HashSet<string>(StringComparer.Ordinal);
        foreach (string line in after) { remaining.Add(Rdv3Key.FromLine(line, data.IdentityCols)); }
        string at = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
        for (int i = 0; i < before.Length; i++)
        {
            if (!remaining.Contains(Rdv3Key.FromLine(before[i], data.IdentityCols)))
            { Deleted.Add(new Rdv3DeletedRecord { Line = before[i], State = states[i], DeletedAt = at }); }
        }
    }

    // Run after the configured pipeline for previews and again under the shared
    // lease. Local pending states are intentionally absent from this API.
    public void ProtectUpdate(Rdv3Data data, string initial, string[] before, string[] states,
                              string[] incoming, Rdv3UpdateResult update)
    {
        if (Deleted.Count == 0 && data.ProtectedStates.Length == 0) { return; }
        before = before ?? new string[0]; states = states ?? new string[0];
        Dictionary<string, int> old = Rdv3Ledger.RowMap(before, data.IdentityCols, "protected update");
        HashSet<string> removed = new HashSet<string>(StringComparer.Ordinal);
        foreach (Rdv3DeletedRecord row in Deleted) { removed.Add(Rdv3Key.FromLine(row.Line, data.IdentityCols)); }
        HashSet<string> locked = new HashSet<string>(data.ProtectedStates, StringComparer.Ordinal);
        HashSet<string> source = new HashSet<string>(StringComparer.Ordinal);
        foreach (string line in incoming) { source.Add(Rdv3Key.FromLine(line, data.IdentityCols)); }
        List<string> lines = new List<string>(), outputStates = new List<string>();
        HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
        update.Added = update.Updated = update.Unchanged = update.Kept = update.Deleted = 0;
        update.Protected = update.SkippedDeleted = 0; update.ResetLines.Clear();
        for (int i = 0; i < update.Lines.Length; i++)
        {
            string line = update.Lines[i], state = update.States[i], id = Rdv3Key.FromLine(line, data.IdentityCols);
            if (removed.Contains(id)) { update.SkippedDeleted++; continue; }
            int previous;
            if (old.TryGetValue(id, out previous))
            {
                if (locked.Contains(states[previous]))
                { line = before[previous]; state = states[previous]; update.Protected++; update.Kept++; }
                else if (!source.Contains(id)) { update.Kept++; }
                else if (line == before[previous]) { update.Unchanged++; }
                else { update.Updated++; }
                if (states[previous] != initial && state == initial) { update.ResetLines.Add(line); }
            }
            else { update.Added++; }
            lines.Add(line); outputStates.Add(state); seen.Add(id);
        }
        for (int i = 0; i < before.Length; i++)
        {
            if (seen.Contains(Rdv3Key.FromLine(before[i], data.IdentityCols))) { continue; }
            if (locked.Contains(states[i]))
            { lines.Add(before[i]); outputStates.Add(states[i]); update.Protected++; update.Kept++; }
            else { update.Deleted++; }
        }
        update.Lines = lines.ToArray(); update.States = outputStates.ToArray();
    }

    public void Restore(Rdv3Data data, IList<Rdv3DeletedRecord> requested, ref string[] lines, ref string[] states)
    {
        Dictionary<string, int> active = Rdv3Ledger.RowMap(lines, data.IdentityCols, "restore");
        Dictionary<string, Rdv3DeletedRecord> archived = new Dictionary<string, Rdv3DeletedRecord>(StringComparer.Ordinal);
        foreach (Rdv3DeletedRecord row in Deleted) { archived.Add(Rdv3Key.FromLine(row.Line, data.IdentityCols), row); }
        HashSet<string> restore = new HashSet<string>(StringComparer.Ordinal);
        foreach (Rdv3DeletedRecord requestedRow in requested)
        {
            string id = Rdv3Key.FromLine(requestedRow.Line, data.IdentityCols);
            Rdv3DeletedRecord actual;
            if (active.ContainsKey(id) || !restore.Add(id) || !archived.TryGetValue(id, out actual)
                || actual.Line != requestedRow.Line || actual.State != requestedRow.State || actual.DeletedAt != requestedRow.DeletedAt)
            { throw new InvalidDataException(Rdv3Text.ArchiveIdentityConflict); }
        }
        List<string> result = new List<string>(lines), resultStates = new List<string>(states);
        foreach (Rdv3DeletedRecord row in Deleted)
        { if (restore.Contains(Rdv3Key.FromLine(row.Line, data.IdentityCols))) { result.Add(row.Line); resultStates.Add(row.State); } }
        Deleted.RemoveAll(delegate(Rdv3DeletedRecord row) { return restore.Contains(Rdv3Key.FromLine(row.Line, data.IdentityCols)); });
        lines = result.ToArray(); states = resultStates.ToArray();
    }
}
