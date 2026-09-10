using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

// Integration checks write only the new evidence directory. Each milestone
// includes an actual XLSX copy that another process/reviewer can read.
public static class Rdv3ProtectionWorkflow
{
    private static void Check(bool pass, string name)
    { if (!pass) { throw new Exception(name); } Console.WriteLine("PASS " + name); }
    private static string Cell(string line, Rdv3Data data, string field)
    { return line.Split('\t')[Array.IndexOf(data.ColumnRefs, field)]; }
    private static void CopyInputs(string from, string to)
    {
        Directory.CreateDirectory(to);
        foreach (string file in Directory.GetFiles(from))
        { if (Path.GetExtension(file) == ".csv" || Path.GetExtension(file) == ".xlsx") { File.Copy(file, Path.Combine(to, Path.GetFileName(file)), true); } }
    }
    private static Rdv3LedgerLock Acquire(Rdv3SharedFiles shared)
    { Rdv3LockInfo owner; Rdv3LedgerLock lease = shared.TryAcquire(out owner); if (lease == null) { throw new Exception("test lease unavailable"); } return lease; }
    private static void Reject(Action action, string name)
    { bool refused = false; try { action(); } catch (InvalidDataException) { refused = true; } Check(refused, name); }
    public static int Run(string root, string evidence)
    {
        Directory.CreateDirectory(evidence);
        string dir = Path.Combine(evidence, "app"), dataDir = Path.Combine(dir, "data"); Directory.CreateDirectory(dir);
        CopyInputs(Path.Combine(root, "tests", "fixtures", "sample-v4"), dataDir);
        string config = Path.Combine(dir, "settings.json"); File.Copy(Path.Combine(root, "tests", "fixtures", "sample-v4", "settings.json"), config, false);
        Rdv3Config cfg = Rdv3Config.Load(config); Rdv3Data data = cfg.Data; Rdv3WorkState work = cfg.Screen.Work;
        string ledger = Rdv3Files.Full(cfg.Ledger, dir);
        Rdv3SharedFiles shared = new Rdv3SharedFiles(ledger, "test", "tester", "protection");
        Rdv3MergeResult source = Rdv3Ledger.BuildFromCsv(data, dataDir);
        Rdv3LedgerStore store = new Rdv3LedgerStore(ledger, data, work, shared, delegate { Rdv3BusinessDefinition.CheckCurrent(cfg); });
        Action<string, string> trace = delegate(string stage, string detail) { Console.WriteLine(stage + " " + detail); };
        Action<string> warn = delegate(string warning) { Console.WriteLine("WARNING " + warning); };
        Rdv3ApplyOutcome create = store.Apply(source, null, "create", delegate { return Acquire(shared); }, trace, warn);
        Check(create.Error == null && create.Committed, "create source pipeline");
        Rdv3LedgerSnapshot initial = store.Read(data.Head);
        int paid = 0, noApp = 0;
        foreach (string line in initial.Lines) { if (Cell(line, data, "PAYMAP.決済確認済") != "") { paid++; } if (Cell(line, data, "APP.申請番号") == "") { noApp++; } }
        Check(initial.Lines.Length == 100 && paid == 80 && noApp == 20, "initial 100 / paid 80 / unpaid 20 / APP missing 20");
        File.Copy(ledger, Path.Combine(evidence, "01-initial.xlsx"));
        Rdv3ProcessJobDef deletion = data.JobOf("delete-processed-records");
        Rdv3PendingStore pending = new Rdv3PendingStore(Path.Combine(evidence, "pending.dat"));
        for (int i = 0; i < 21; i++) { pending.Set(Rdv3Key.FromLine(initial.Lines[i], data.IdentityCols), "TRUE", initial.Lines[i], initial.States[i]); }
        Rdv3DeleteResult beforeSend = Rdv3Ledger.ApplyDelete(data, deletion, dataDir, initial.Lines, initial.States, work.InitialStored);
        Check(beforeSend.Deleted == 0 && beforeSend.Lines.Length == 100, "local done changes do not permit deletion");
        using (Rdv3LedgerLock lease = Acquire(shared))
        {
            Rdv3LedgerSnapshot latest = store.Read(data.Head);
            Rdv3PendingApply sent = pending.PrepareSend(latest.Lines, latest.States, data.IdentityCols, work.InitialStored);
            Check(sent.Resolved.Count == 21 && sent.Unmatched.Count == 0, "send all 21 saved states");
            store.Write(data.Head, latest.Lines, sent.States, "send", latest.Protection); pending.Remove(sent.Resolved); lease.Release();
        }
        File.Copy(ledger, Path.Combine(evidence, "02-sent.xlsx"));
        using (Rdv3LedgerLock lease = Acquire(shared))
        {
            Rdv3LedgerSnapshot latest = store.Read(data.Head);
            Rdv3DeleteResult result = Rdv3Ledger.ApplyDelete(data, deletion, dataDir, latest.Lines, latest.States, work.InitialStored);
            latest.Protection.ArchiveRemoved(data, latest.Lines, latest.States, result.Lines);
            store.Write(data.Head, result.Lines, result.States, "delete", latest.Protection); lease.Release();
        }
        Rdv3LedgerSnapshot removed = store.Read(data.Head);
        Check(removed.Lines.Length == 80 && removed.Protection.Deleted.Count == 20, "delete 20 into archive, keep 80 live");
        for (int i = 0; i < 20; i++) { Check(removed.Protection.Deleted[i].Line == initial.Lines[i] && removed.Protection.Deleted[i].State == "TRUE", "archive full row " + i); }
        File.Copy(ledger, Path.Combine(evidence, "03-deleted.xlsx"));
        pending.Set(Rdv3Key.FromLine(initial.Lines[21], data.IdentityCols), "TRUE", initial.Lines[21], "FALSE");
        CopyInputs(Path.Combine(root, "samples", "next-period"), dataDir);
        source = Rdv3Ledger.BuildFromCsv(data, dataDir);
        Rdv3ApplyOutcome next = store.Apply(source, removed.Lines, "next", delegate { return Acquire(shared); }, trace, warn);
        Check(next.Error == null && next.Committed, "next-period update committed");
        Rdv3LedgerSnapshot after = store.Read(data.Head);
        Dictionary<string, int> indices = Rdv3Ledger.RowMap(after.Lines, data.IdentityCols, "next readback");
        string a = Rdv3Key.FromLine(initial.Lines[20], data.IdentityCols), b = Rdv3Key.FromLine(initial.Lines[0], data.IdentityCols);
        string u = Rdv3Key.FromLine(initial.Lines[21], data.IdentityCols), blank = Rdv3Key.FromLine(initial.Lines[22], data.IdentityCols);
        string absent = Rdv3Key.FromLine(initial.Lines[23], data.IdentityCols);
        Check(after.Lines.Length == 81 && after.Protection.Deleted.Count == 20 && !indices.ContainsKey(b), "old deleted keys excluded and new C added");
        Check(after.Lines[indices[a]] == initial.Lines[20] && after.States[indices[a]] == "TRUE", "saved A full content/state preserved");
        Check(after.Lines[indices[absent]] == initial.Lines[23], "row absent from current source retained");
        Check(Cell(after.Lines[indices[blank]],data,"PAY.利用者氏名") == "" && Cell(after.Lines[indices[blank]],data,"APP.備考欄") == "", "unprocessed old values replaced with blanks");
        Check(after.Lines[indices[u]] != initial.Lines[21] && after.States[indices[u]] == "FALSE", "pending row content updated using saved state");
        Check(pending.PrepareSend(after.Lines,after.States,data.IdentityCols,work.InitialStored).Unmatched.Count == 1, "pending conflict survives update");
        for (int i=0;i<20;i++) { Check(after.Protection.Deleted[i].Line == initial.Lines[i], "archive survives reimport " + i); }
        File.Copy(ledger,Path.Combine(evidence,"04-next-period.xlsx"));
        CopyInputs(Path.Combine(root,"samples","partial-pay"),dataDir);
        source=Rdv3Ledger.BuildFromCsv(data,dataDir);
        Rdv3ApplyOutcome partial=store.Apply(source,after.Lines,"partial",delegate{return Acquire(shared);},trace,warn);
        after=store.Read(data.Head); indices=Rdv3Ledger.RowMap(after.Lines,data.IdentityCols,"partial");
        Check(partial.Error==null && after.Lines.Length==81 && after.Lines[indices[a]]==initial.Lines[20] && !indices.ContainsKey(b),"partial-file update protects A/B");
        Check(Cell(after.Lines[indices[blank]],data,"PAY.決済金額")=="3333" && Cell(after.Lines[indices[blank]],data,"PAY.利用者氏名")=="部分差替 新名","partial-file correction adopted");
        File.Copy(ledger,Path.Combine(evidence,"05-partial.xlsx"));
        using(Rdv3LedgerLock lease=Acquire(shared)) {
            after=store.Read(data.Head);string[] lines=after.Lines,states=after.States;
            after.Protection.Restore(data,new Rdv3DeletedRecord[]{after.Protection.Deleted[0]},ref lines,ref states);
            store.Write(data.Head,lines,states,"restore-one",after.Protection);lease.Release();
        }
        after=store.Read(data.Head);indices=Rdv3Ledger.RowMap(after.Lines,data.IdentityCols,"restored");
        Check(after.Lines.Length==82&&after.Protection.Deleted.Count==19&&after.Lines[indices[b]]==initial.Lines[0]&&after.States[indices[b]]=="TRUE","B explicitly restored in full");
        File.Copy(ledger,Path.Combine(evidence,"06-restored-one.xlsx"));
        using(Rdv3LedgerLock lease=Acquire(shared)) {
            after=store.Read(data.Head);string[] lines=after.Lines,states=after.States;
            after.Protection.Restore(data,new List<Rdv3DeletedRecord>(after.Protection.Deleted),ref lines,ref states);
            store.Write(data.Head,lines,states,"restore-all",after.Protection);lease.Release();
        }
        File.Copy(ledger,Path.Combine(evidence,"07-restored-all.xlsx"));
        Rdv3Config restarted=Rdv3Config.Load(config);Rdv3Ledger.BuildFromCsv(restarted.Data,dataDir);
        after=new Rdv3LedgerStore(ledger,restarted.Data,restarted.Screen.Work,null).Read(restarted.Data.Head);
        Check(after.Lines.Length==101&&after.Protection.Deleted.Count==0,"restart reads 101 live and empty archive");
        string valid=File.ReadAllText(config,Encoding.UTF8);byte[] beforeBytes=File.ReadAllBytes(ledger);
        string[] edits={valid.Replace("\"condition\": \"match\"","\"condition\": \"left\""),valid.Replace("[0-9]{8}","[0-9]{7}"),valid.Replace("\"type\": \"number\"","\"type\": \"text\""),valid.Replace("\"value\": \"TRUE\"","\"value\": \"FALSE\"")};
        for(int i=0;i<edits.Length;i++) {
            Check(edits[i]!=valid,"mutation exists "+i);File.WriteAllText(config,edits[i],new UTF8Encoding(false));
            bool rejected=false;try{store.Write(data.Head,after.Lines,after.States,"changed-config",after.Protection);}catch(Exception){rejected=true;}
            Check(rejected,"direct JSON definition edit rejected "+i);
            Check(Convert.ToBase64String(beforeBytes)==Convert.ToBase64String(File.ReadAllBytes(ledger)),"JSON refusal left ledger unchanged "+i);
        }
        File.WriteAllText(config,valid,new UTF8Encoding(false));
        store.CheckWritable(after.Protection);
        Check(true,"valid original configuration remains writable");
        return 0;
    }
}
