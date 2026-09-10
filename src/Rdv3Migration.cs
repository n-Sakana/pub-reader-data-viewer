using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

public static class Rdv3Migration
{
    // The old file cannot prove which join definition created it. Requiring
    // its original JSON and explicit operator attestation makes this gap
    // visible; migration never writes back to the legacy pathname.
    public static int Run(string appDir, string originalPath, string configPath,
                          string dataDir, string sourcePath, string outputPath, bool confirmed)
    {
        try
        {
            if (!confirmed) { throw new InvalidOperationException(Rdv3Text.MigrationNeedsConfirmation); }
            Rdv3Config original = Rdv3Config.Load(Path.GetFullPath(originalPath));
            Rdv3Config current = Rdv3Config.Load(Path.GetFullPath(configPath));
            if (string.IsNullOrEmpty(dataDir)) { dataDir = Rdv3Files.Full(current.DataDir, appDir); }
            sourcePath = Path.GetFullPath(sourcePath);
            Migrate(original, current, appDir, Path.GetFullPath(dataDir), sourcePath, outputPath);
            Console.WriteLine(Rdv3Text.MigrationDone + Path.GetFullPath(outputPath));
            return 0;
        }
        catch (Exception error)
        { Rdv3Log.Error("ledger migration", error); Console.Error.WriteLine(error.Message); return 3; }
    }

    internal static void Migrate(Rdv3Config original, Rdv3Config current, string appDir,
                                 string dataDir, string sourcePath, string outputPath)
    {
        if (original.Data.LegacyDefinition != current.Data.LegacyDefinition
            || Rdv3Files.LegacyStorageContract(original.Data, original.Screen.Work)
                != Rdv3Files.LegacyStorageContract(current.Data, current.Screen.Work))
        { throw new InvalidDataException(Rdv3Text.MigrationDefinitionMismatch); }
        string logPath = Rdv3Files.Full(current.Log, appDir);
        outputPath = Rdv3Files.MigrationPath(outputPath, appDir, dataDir, sourcePath, logPath, current.SourcePath, current.Data);
        Rdv3Files.MigrationPath(outputPath, appDir, dataDir, sourcePath, logPath, original.SourcePath, original.Data);
        string[][] heads = new string[current.Data.Tables.Count][];
        for (int i = 0; i < heads.Length; i++)
        {
            Rdv3TableDef table = current.Data.Tables[i];
            heads[i] = Rdv3Table.ReadHead(Rdv3Files.Full(table.File, dataDir), table.Enc, table.EncodingSetting,
                current.Data.SourceReferences(table.Id), table.HeaderRow, table.Delimiter, table.Sheet);
        }
        current.Data.Bind(heads);
        Dictionary<string, string[]> byId = new Dictionary<string, string[]>(StringComparer.Ordinal);
        for (int i = 0; i < heads.Length; i++) { byId.Add(current.Data.Tables[i].Id, heads[i]); }
        string[][] originalHeads = new string[original.Data.Tables.Count][];
        for (int i = 0; i < originalHeads.Length; i++) { originalHeads[i] = byId[original.Data.Tables[i].Id]; }
        original.Data.Bind(originalHeads);
        Rdv3BusinessDefinition.BindFileInputs(current.Data, dataDir);
        // Business definitions were matched by id; the current paths may move.
        foreach (Rdv3ProcessJobDef job in original.Data.Jobs)
        {
            Rdv3ProcessJobDef now = current.Data.JobOf(job.Id);
            for (int i = 0; i < job.Inputs.Count; i++) { job.Inputs[i].Head = now.Inputs[i].Head; }
        }
        Rdv3SharedFiles shared = new Rdv3SharedFiles(sourcePath, Environment.MachineName, Environment.UserName, "migration");
        Rdv3LockInfo owner;
        using (Rdv3LedgerLock lease = shared.TryAcquire(out owner))
        {
            if (lease == null) { throw new IOException(Rdv3Text.MigrationBusy); }
            Rdv3LedgerSnapshot source = new Rdv3LedgerStore(sourcePath, original.Data, original.Screen.Work, shared).Read(original.Data.Head);
            if (!source.Protection.Legacy) { throw new InvalidDataException(Rdv3Text.MigrationAlreadyBound); }
            if (source.Warning.Length > 0) { throw new InvalidDataException(source.Warning); }
            Rdv3LedgerProtection protection = Rdv3LedgerProtection.Create(current.Data);
            protection.Validate(current.Data, current.Screen.Work, source.Lines);
            Rdv3Xlsx.Write(outputPath, current.Data.Head, current.Screen.Work.Column, source.Lines, source.States,
                "migration", Rdv3Files.StorageContract(current.Data, current.Screen.Work), protection, false);
            Console.WriteLine("rows=" + source.Lines.Length.ToString(CultureInfo.InvariantCulture));
            lease.Release();
        }
    }
}
