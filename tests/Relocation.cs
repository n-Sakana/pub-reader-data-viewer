using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

public static class Relocation
{
    private static void Check(bool result, string label)
    { if (!result) { throw new Exception(label); } Console.WriteLine("PASS " + label); }
    private static string StateHeading(string file)
    {
        using (Stream fileStream = File.OpenRead(file))
        using (System.IO.Compression.ZipArchive zip = new System.IO.Compression.ZipArchive(fileStream))
        using (Stream stream = zip.GetEntry("xl/worksheets/sheet1.xml").Open())
        {
            System.Xml.XmlDocument document = new System.Xml.XmlDocument();
            document.Load(stream);
            System.Xml.XmlNamespaceManager ns = new System.Xml.XmlNamespaceManager(document.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            return document.SelectSingleNode("/s:worksheet/s:sheetData/s:row[1]/s:c[1]/s:is/s:t", ns).InnerText;
        }
    }
    private static void Copy(string from, string to)
    {
        Directory.CreateDirectory(to);
        foreach (string file in Directory.GetFiles(from)) { File.Copy(file, Path.Combine(to, Path.GetFileName(file))); }
        foreach (string directory in Directory.GetDirectories(from)) { Copy(directory, Path.Combine(to, Path.GetFileName(directory))); }
    }
    private static void Move(string evidence, string from, string to)
    {
        string boundary = Path.GetFullPath(evidence).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!Path.GetFullPath(from).StartsWith(boundary, StringComparison.OrdinalIgnoreCase)
            || !Path.GetFullPath(to).StartsWith(boundary, StringComparison.OrdinalIgnoreCase)) { throw new Exception("Move escaped evidence"); }
        Directory.Move(from, to);
    }
    private static Rdv3LedgerProtection Read(string file, Rdv3Config cfg, out string[] lines, out string[] states)
    {
        string warning; Rdv3LedgerProtection protection;
        Rdv3Xlsx.ReadProtected(file, cfg.Data.Head, cfg.Screen.Work.Column, out lines, out states, out warning,
            Rdv3Files.StorageContract(cfg.Data, cfg.Screen.Work), Rdv3Files.LegacyStorageContract(cfg.Data, cfg.Screen.Work),
            Rdv3BusinessDefinition.Bound(cfg.Data), out protection);
        Check(warning.Length == 0, "read without warning");
        protection.RequireWritable(cfg.Data); protection.Validate(cfg.Data, cfg.Screen.Work, lines);
        return protection;
    }
    public static void Run(string root, string evidence)
    {
        Directory.CreateDirectory(evidence);
        string before = Path.Combine(evidence, "移動前 [1] & (空白) ! %literal% ' ; + #");
        string after = Path.Combine(evidence, "移動後 [2] & (空白) ! %literal% ' ; + #");
        string copied = Path.Combine(evidence, "コピー [3]");
        Directory.CreateDirectory(before);
        Rdv3Config cfg = Rdv3Config.Load(Path.Combine(root, "configs/sample/settings.json"));
        string input = Path.Combine(root, "samples/current/data");
        Rdv3Ledger.BuildFromCsv(cfg.Data, input);
        Rdv3WorkState work = cfg.Screen.Work;
        Check(work.Column == "確認状態", "canonical configuration uses confirmation heading");
        string contract = Rdv3Files.StorageContract(cfg.Data, work);
        Rdv3ProcessResult start = Rdv3Process.Run(cfg.Data, cfg.Data.UpdateJob, input, new string[0], new string[0], work.InitialStored);
        int a = Array.FindIndex(start.Lines, line => line.Contains("AB10000001CD"));
        start.States[a] = "TRUE";
        Rdv3ProcessResult deleted = Rdv3Process.Run(cfg.Data, cfg.Data.JobOf("delete-processed-records"), input, start.Lines, start.States, work.InitialStored);
        Rdv3LedgerProtection protection = Rdv3LedgerProtection.Create(cfg.Data);
        protection.ArchiveRemoved(cfg.Data, start.Lines, start.States, deleted.Lines);
        string file = Path.Combine(before, "統合台帳.xlsx");
        Rdv3Xlsx.Write(file, cfg.Data.Head, work.Column, deleted.Lines, deleted.States, "new", contract, protection);
        Check(StateHeading(file) == "確認状態", "new ledger uses canonical heading");
        string archiveName = Path.GetFileName(Directory.GetFiles(Path.Combine(before, "archived")).Single());
        byte[] archived = File.ReadAllBytes(Path.Combine(before, "archived", archiveName));
        byte[] ledgerBytes = File.ReadAllBytes(file);
        int b = Array.FindIndex(deleted.Lines, line => line.Contains("AB10000002CD"));
        string identity = Rdv3Key.FromLine(deleted.Lines[b], cfg.Data.IdentityCols);
        string pendingPath = Rdv3PendingStore.LocationPathFor(file);
        Rdv3PendingStore pending = new Rdv3PendingStore(pendingPath);
        pending.Set(identity, "TRUE", deleted.Lines[b], "FALSE");
        Check(Rdv3PendingStore.LocationPathFor(file) == pendingPath && pending.Count == 1, "canonical reference remains stable across reopen");
        string companion = Directory.GetFiles(before, "*.state").Single();
        byte[] referenceBytes = File.ReadAllBytes(companion);
        try
        {
            File.WriteAllText(companion, "invalid reference", Encoding.UTF8);
            bool rejected = false;
            try { Rdv3PendingStore.LocationPathFor(file); } catch (InvalidDataException) { rejected = true; }
            Check(rejected && new Rdv3PendingStore(pendingPath).Count == 1 && File.ReadAllBytes(file).SequenceEqual(ledgerBytes), "invalid reference stops without replacing saved or unsent data");
            string[] fields = Encoding.UTF8.GetString(referenceBytes).TrimEnd('\n').Split('\n');
            fields[1] = "pending-" + Guid.NewGuid().ToString("N") + ".dat";
            File.WriteAllText(companion, string.Join("\n", fields) + "\n", Encoding.UTF8);
            rejected = false;
            try { Rdv3PendingStore.LocationPathFor(file); } catch (FileNotFoundException) { rejected = true; }
            Check(rejected && new Rdv3PendingStore(pendingPath).Count == 1, "missing referenced store stops instead of falling back to an empty store");
        }
        finally { File.WriteAllBytes(companion, referenceBytes); }
        Copy(before, copied);
        string copyPendingPath = Rdv3PendingStore.LocationPathFor(Path.Combine(copied, "統合台帳.xlsx"));
        Rdv3PendingStore copyPending = new Rdv3PendingStore(copyPendingPath);
        Check(copyPendingPath != pendingPath && copyPending.Count == 1, "copied folder receives an independent pending snapshot");
        copyPending.Remove(new List<string> { identity });
        Check(new Rdv3PendingStore(pendingPath).Count == 1, "editing copied pending state does not affect original");
        Move(evidence, before, after);
        file = Path.Combine(after, "統合台帳.xlsx");
        companion = Path.Combine(after, Path.GetFileName(companion));
        FileAttributes attributes = File.GetAttributes(companion);
        try
        {
            File.SetAttributes(companion, attributes | FileAttributes.ReadOnly);
            bool rejected = false;
            try { Rdv3PendingStore.LocationPathFor(file); } catch (UnauthorizedAccessException) { rejected = true; } catch (IOException) { rejected = true; }
            Check(rejected && File.ReadAllBytes(companion).SequenceEqual(referenceBytes)
                && new Rdv3PendingStore(pendingPath).Count == 1, "reference save failure stops without switching storage paths");
        }
        finally { File.SetAttributes(companion, attributes); }
        string relocated = Rdv3PendingStore.LocationPathFor(file);
        Check(relocated == pendingPath && new Rdv3PendingStore(relocated).Count == 1, "moving folder keeps the private unsent state");
        Check(File.ReadAllBytes(file).SequenceEqual(ledgerBytes), "move leaves saved ledger bytes unchanged");
        string[] lines, states;
        protection = Read(file, cfg, out lines, out states);
        Check(lines.Length == 2 && protection.Deleted.Count == 1 && states[b] == "FALSE", "ledger and archive read without promoting unsent state");
        pending = new Rdv3PendingStore(relocated);
        Check(pending.Overlay(lines, states, cfg.Data.IdentityCols)[b] == "TRUE", "moved local state is visible before sending");
        Rdv3ProcessResult unchanged = Rdv3Process.Run(cfg.Data, cfg.Data.UpdateJob, input, lines, states, work.InitialStored);
        protection.ProtectUpdate(cfg.Data, work.InitialStored, lines, states, start.Lines, unchanged.Update);
        Check(unchanged.Update.Lines.SequenceEqual(lines) && unchanged.Update.States.SequenceEqual(states)
            && unchanged.Update.SkippedDeleted == 1, "move-only update has no content difference and blocks deleted reimport");
        Rdv3PendingApply send = pending.PrepareSend(lines, states, cfg.Data.IdentityCols, work.InitialStored);
        Check(send.Unmatched.Count == 0 && send.Resolved.Count == 1, "pending state is sendable after move");
        Rdv3Xlsx.Write(file, cfg.Data.Head, work.Column, lines, send.States, "send", contract, protection);
        pending.Remove(send.Resolved);
        protection = Read(file, cfg, out lines, out states);
        Check(StateHeading(file) == "確認状態" && states[b] == "TRUE" && pending.Count == 0, "send retains canonical heading and confirmation values");
        Check(File.ReadAllBytes(Path.Combine(after, "archived", archiveName)).SequenceEqual(archived), "send leaves archive bytes unchanged");
        Rdv3ProcessResult removal = Rdv3Process.Run(cfg.Data, cfg.Data.JobOf("delete-processed-records"), input, lines, states, work.InitialStored);
        Check(removal.Deleted == 1, "deletion still requires matching saved confirmation");
        protection.ArchiveRemoved(cfg.Data, lines, states, removal.Lines);
        Rdv3Xlsx.Write(file, cfg.Data.Head, work.Column, removal.Lines, removal.States, "delete", contract, protection);
        protection = Read(file, cfg, out lines, out states);
        Check(lines.Length == 1 && protection.Deleted.Count == 2, "successive deletions retain all archive protection");
        Move(evidence, after, before);
        Check(Rdv3PendingStore.LocationPathFor(Path.Combine(before, "統合台帳.xlsx")) == pendingPath
            && new Rdv3PendingStore(pendingPath).Count == 0, "moving back cannot resurrect sent changes");
        Console.WriteLine("Canonical storage and relocation checks complete.");
    }
}
