using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

public static class ArchiveSeparation
{
    private static void Check(bool value, string label)
    { if (!value) { throw new Exception(label); } Console.WriteLine("PASS " + label); }
    private static string Metadata(string path)
    { using (FileStream stream = File.OpenRead(path)) using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read)) using (StreamReader reader = new StreamReader(zip.GetEntry("rdv-protection.xml").Open())) { return reader.ReadToEnd(); } }
    private static Rdv3LedgerProtection Read(string path, Rdv3Data data, Rdv3WorkState work, out string[] lines, out string[] states)
    { string warning; Rdv3LedgerProtection p; Rdv3Xlsx.ReadProtected(path, data.Head, work.Column, out lines, out states, out warning, null, null, Rdv3BusinessDefinition.Bound(data), out p); Check(warning.Length == 0, "read without data warnings"); p.Validate(data, work, lines); return p; }
    public static void Run(string root, string output)
    {
        Directory.CreateDirectory(output);
        Rdv3Config cfg = Rdv3Config.Load(Path.Combine(root, "configs/sample/settings.json"));
        Rdv3Data data = cfg.Data; Rdv3WorkState work = cfg.Screen.Work;
        string input = Path.Combine(root, "samples/current/data");
        Rdv3MergeResult source = Rdv3Ledger.BuildFromCsv(data, input);
        Rdv3ProcessResult start = Rdv3Process.Run(data, data.UpdateJob, input, new string[0], new string[0], work.InitialStored);
        Check(start.Lines.Length == 3, "initial three records");
        string file = Path.Combine(output, "統合台帳.xlsx");
        Rdv3LedgerProtection protection = Rdv3LedgerProtection.Create(data);
        int a = Array.FindIndex(start.Lines, line => line.Contains("AB10000001CD"));
        start.States[a] = "TRUE";
        Rdv3Xlsx.Write(file, data.Head, work.Column, start.Lines, start.States, "initial", null, protection);
        Check(!Directory.Exists(Path.Combine(output, "archived")), "no empty archive on creation");
        Rdv3ProcessResult deletion = Rdv3Process.Run(data, data.JobOf("delete-processed-records"), input, start.Lines, start.States, work.InitialStored);
        protection.ArchiveRemoved(data, start.Lines, start.States, deletion.Lines);
        string removedLine = protection.Deleted[0].Line, removedAt = protection.Deleted[0].DeletedAt;

        byte[] original = File.ReadAllBytes(file);
        bool denied = false;
        using (FileStream held = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
        { try { Rdv3Xlsx.Write(file, data.Head, work.Column, deletion.Lines, deletion.States, "locked", null, protection); } catch (IOException) { denied = true; } }
        Check(denied && original.SequenceEqual(File.ReadAllBytes(file)), "failed main replacement preserves original ledger");
        Check(Directory.GetFiles(Path.Combine(output, "archived"), "*.xlsx").Length == 0, "failed deletion leaves no committed archive");

        Rdv3Xlsx.Write(file, data.Head, work.Column, deletion.Lines, deletion.States, "delete", null, protection);
        string[] lines, states;
        protection = Read(file, data, work, out lines, out states);
        Check(lines.Length == 2 && protection.Deleted.Count == 1, "two active records and one archived record");
        string meta = Metadata(file);
        Check(!meta.Contains("<row>") && !meta.Contains("AB10000001CD"), "main metadata has references and no deleted record content");
        string archive = Directory.GetFiles(Path.Combine(output, "archived"), "*.xlsx").Single();
        string[] archiveHead = data.Head.Concat(new[] { "削除日時" }).ToArray(); string[] archiveLines, archiveStates; string warning;
        Rdv3LedgerProtection ignored;
        Rdv3Xlsx.ReadProtected(archive, archiveHead, work.Column, out archiveLines, out archiveStates, out warning, null, null, null, out ignored);
        Check(archiveLines.Length == 1 && archiveLines[0].EndsWith("\t" + removedAt) && archiveStates[0] == "TRUE", "separate Excel contains confirmation state and deletion date");
        Check(protection.Deleted[0].Line == removedLine && protection.Deleted[0].DeletedAt == removedAt, "archive preserves full text and deletion time");
        byte[] archivedBytes = File.ReadAllBytes(archive);
        Rdv3Xlsx.Write(file, data.Head, work.Column, lines, states, "send", null, protection);
        Check(Directory.GetFiles(Path.Combine(output, "archived"), "*.xlsx").Length == 1 && archivedBytes.SequenceEqual(File.ReadAllBytes(archive)), "state-only save reuses archive without rewriting");
        Rdv3ProcessResult reimport = Rdv3Process.Run(data, data.UpdateJob, input, lines, states, work.InitialStored);
        protection.ProtectUpdate(data, work.InitialStored, lines, states, source.Lines, reimport.Update);
        Check(reimport.Update.Lines.Length == 2 && reimport.Update.SkippedDeleted == 1, "deleted number cannot return from CSV");

        File.Move(archive, archive + ".held");
        bool missing = false;
        try { Read(file, data, work, out lines, out states); } catch (FileNotFoundException) { missing = true; }
        finally { File.Move(archive + ".held", archive); }
        Check(missing, "missing archive fails instead of losing deletion protection");
        File.WriteAllBytes(archive, new byte[] { 1, 2, 3 });
        bool corrupt = false;
        try { Read(file, data, work, out lines, out states); } catch (InvalidDataException) { corrupt = true; }
        finally { File.WriteAllBytes(archive, archivedBytes); }
        Check(corrupt, "modified archive detected");

        protection = Read(file, data, work, out lines, out states);
        protection.Restore(data, protection.Deleted.ToArray(), ref lines, ref states);
        Rdv3Xlsx.Write(file, data.Head, work.Column, lines, states, "restore", null, protection);
        protection = Read(file, data, work, out lines, out states);
        Check(lines.Length == 3 && protection.Deleted.Count == 0, "existing explicit restore remains functional");

        // An actual v1 workbook with embedded deleted data upgrades on save.
        string legacyDir = Path.Combine(output, "embedded"); Directory.CreateDirectory(legacyDir);
        string legacy = Path.Combine(legacyDir, "統合台帳.xlsx");
        Rdv3Xlsx.Write(legacy, data.Head, work.Column, deletion.Lines, deletion.States, "legacy", null, Rdv3LedgerProtection.Create(data));
        using (FileStream stream = new FileStream(legacy, FileMode.Open, FileAccess.ReadWrite))
        using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Update))
        {
            zip.GetEntry("rdv-protection.xml").Delete();
            using (XmlWriter writer = XmlWriter.Create(zip.CreateEntry("rdv-protection.xml").Open(), new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = true, NewLineHandling = NewLineHandling.Entitize }))
            {
                writer.WriteStartElement("protection"); writer.WriteAttributeString("version", "1"); writer.WriteElementString("definition", Rdv3BusinessDefinition.Bound(data));
                writer.WriteStartElement("deleted"); writer.WriteStartElement("row"); writer.WriteElementString("line", removedLine); writer.WriteElementString("state", "TRUE"); writer.WriteElementString("at", removedAt);
                writer.WriteEndElement(); writer.WriteEndElement(); writer.WriteEndElement();
            }
        }
        protection = Read(legacy, data, work, out lines, out states);
        Rdv3Xlsx.Write(legacy, data.Head, work.Column, lines, states, "migrate", null, protection);
        protection = Read(legacy, data, work, out lines, out states);
        Check(protection.Deleted.Count == 1 && protection.Deleted[0].Line == removedLine && !Metadata(legacy).Contains("<row>"), "embedded records migrate losslessly to separate workbook");

        // Adding a second deletion must not copy the first batch again.
        List<string> remaining = new List<string>(lines); List<string> remainingStates = new List<string>(states);
        remaining.RemoveAt(0); remainingStates.RemoveAt(0);
        protection.ArchiveRemoved(data, lines, states, remaining.ToArray());
        Rdv3Xlsx.Write(legacy, data.Head, work.Column, remaining.ToArray(), remainingStates.ToArray(), "second-delete", null, protection);
        protection = Read(legacy, data, work, out lines, out states);
        string[] batches = Directory.GetFiles(Path.Combine(legacyDir, "archived"), "*.xlsx");
        Check(batches.Length == 2 && protection.Deleted.Count == 2, "successive deletions add a batch without copying prior records");
        foreach (string batch in batches) { Rdv3Xlsx.ReadProtected(batch, archiveHead, work.Column, out archiveLines, out archiveStates, out warning, null, null, null, out ignored); Check(archiveLines.Length == 1, "one removed record per deletion batch"); }
        Console.WriteLine("Archive separation checks complete: " + output);
    }
}
