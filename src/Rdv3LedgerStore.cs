using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

internal sealed class Rdv3LedgerSnapshot
{
    public readonly string[] Lines;
    public readonly string[] States;
    public readonly string Warning;
    public readonly Rdv3LedgerProtection Protection;

    public Rdv3LedgerSnapshot(string[] lines, string[] states, string warning, Rdv3LedgerProtection protection)
    {
        Lines = lines;
        States = states;
        Warning = warning;
        Protection = protection;
    }
}

internal sealed class Rdv3ApplyOutcome
{
    public readonly Rdv3UpdateResult Update;
    public readonly Rdv3SharedMarker Marker;
    public readonly bool Committed;
    public readonly Exception Error;
    public readonly string[] Warnings;

    public Rdv3ApplyOutcome(Rdv3UpdateResult update, Rdv3SharedMarker marker, bool committed, Exception error, string[] warnings = null)
    {
        Update = update;
        Marker = marker;
        Committed = committed;
        Error = error;
        Warnings = warnings ?? new string[0];
    }

    // A notification or lease-release failure cannot undo the committed XLSX.
    public bool CanAdopt { get { return Error == null || Committed; } }
}

internal sealed class Rdv3LedgerStore
{
    private readonly string path;
    private readonly Rdv3Data data;
    private readonly Rdv3WorkState work;
    private readonly Rdv3SharedFiles shared;
    private readonly string contract;
    private readonly Action checkDefinition;

    public Rdv3LedgerStore(string ledgerPath, Rdv3Data definition, Rdv3WorkState state, Rdv3SharedFiles sharedFiles, Action checkCurrentDefinition = null)
    {
        path = ledgerPath;
        data = definition;
        work = state;
        shared = sharedFiles;
        contract = Rdv3Files.StorageContract(data, work);
        checkDefinition = checkCurrentDefinition;
    }

    public Rdv3LedgerSnapshot Read(string[] head)
    {
        string[] lines, states;
        string warning;
        Rdv3LedgerProtection protection;
        Rdv3Xlsx.ReadProtected(path, head, work.Column, out lines, out states, out warning, contract,
            Rdv3Files.LegacyStorageContract(data, work), Rdv3BusinessDefinition.Bound(data), out protection);
        string reference = string.Join(" / ", data.IdentityRefs);
        string label = string.Join(" / ", Array.ConvertAll(data.IdentityRefs, data.LabelOf));
        Rdv3Ledger.CheckIdentities(lines, data.IdentityCols, Path.GetFileName(path), label.Length == 0 ? reference : label);
        for (int i = 0; i < states.Length; i++)
        {
            if (work.ByStored(states[i]) == null)
            { throw new InvalidDataException(Rdv3Text.LedgerStateInvalid + (i + 2).ToString(CultureInfo.InvariantCulture)); }
        }
        protection.Validate(data, work, lines);
        return new Rdv3LedgerSnapshot(lines, states, warning, protection);
    }

    public void CheckWritable(Rdv3LedgerProtection protection)
    {
        if (checkDefinition != null) { checkDefinition(); }
        protection.RequireWritable(data);
    }

    public void Write(string[] head, string[] lines, string[] states, string tag, Rdv3LedgerProtection protection)
    {
        CheckWritable(protection);
        protection.Validate(data, work, lines);
        Rdv3Xlsx.Write(path, head, work.Column, lines, states, tag, contract, protection);
    }

    public Rdv3ApplyOutcome Apply(Rdv3MergeResult source, string[] checkedLines, string tag,
                                  Func<Rdv3LedgerLock> acquire, Action<string, string> trace, Action<string> warn)
    {
        Rdv3LedgerLock lease = null;
        Rdv3UpdateResult update = null;
        Rdv3SharedMarker marker = null;
        List<string> warnings = new List<string>();
        bool committed = false;
        try
        {
            lease = acquire();
            if (lease == null) { throw new InvalidOperationException("shared lease was not acquired"); }
            long t = Rdv3Clock.Now();
            string[] latestLines = null, latestStates = null;
            Rdv3LedgerProtection protection = Rdv3LedgerProtection.Create(data);
            if (Rdv3Files.Exists(path))
            {
                Rdv3LedgerSnapshot latest = Read(source.Head);
                latestLines = latest.Lines;
                latestStates = latest.States;
                protection = latest.Protection;
                protection.RequireWritable(data);
                if (latest.Warning.Length > 0) { warn(latest.Warning); }
                trace("load", "under_lock rows=" + latestLines.Length.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            }

            // Content changes invalidate the confirmed preview. State-only sends
            // remain valid and must be read under this same lease before merging.
            int firstChanged;
            if ((checkedLines == null) != (latestLines == null)
                || (checkedLines != null && !Rdv3Ledger.SameContent(checkedLines, latestLines, out firstChanged)))
            { throw new IOException(Rdv3Text.UpdateChangedDuringCheck); }
            t = Rdv3Clock.Now();
            if (source.Prepared == null)
            { update = Rdv3Ledger.ApplyUpdate(source.Job, latestLines, latestStates, source.Lines, data.IdentityCols, work.InitialStored); }
            else
            {
                Rdv3ProcessResult executed = Rdv3Process.Execute(source.Prepared, latestLines, latestStates, work.InitialStored, false);
                update = executed.Update;
                foreach (string warning in executed.Warnings)
                {
                    trace("warning", warning);
                    if (!source.Warnings.Contains(warning)) { warnings.Add(warning); }
                }
                // Record warnings are shown together after releasing the
                // lease. Holding shared I/O open for one OK per row would
                // block every other operator until all warnings were read.
            }
            string operation = source.Job.ApplyStep == null ? "pipeline" : source.Job.ApplyStep.Operation;
            protection.ProtectUpdate(data, work.InitialStored, latestLines, latestStates, source.Lines, update);
            if (latestLines != null) { protection.ArchiveRemoved(data, latestLines, latestStates, update.Lines); }
            trace("apply", "operation=" + operation
                + " source=" + source.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " result=" + update.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " added=" + update.Added.ToString(CultureInfo.InvariantCulture)
                + " updated=" + update.Updated.ToString(CultureInfo.InvariantCulture)
                + " unchanged=" + update.Unchanged.ToString(CultureInfo.InvariantCulture)
                + " kept=" + update.Kept.ToString(CultureInfo.InvariantCulture)
                + " deleted=" + update.Deleted.ToString(CultureInfo.InvariantCulture)
                + " protected=" + update.Protected.ToString(CultureInfo.InvariantCulture)
                + " skipped_deleted=" + update.SkippedDeleted.ToString(CultureInfo.InvariantCulture)
                + " reset=" + update.ResetLines.Count.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            if (latestLines == null || !Rdv3Ledger.SameLedger(latestLines, latestStates, update.Lines, update.States))
            {
                t = Rdv3Clock.Now();
                Write(source.Head, update.Lines, update.States, tag, protection);
                committed = true;
                trace("persist", "target=xlsx rows=" + update.Lines.Length.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
                // The operation line goes before the notification: a marker
                // failure is reported to the operator, but the file was replaced
                // either way and the record must say so.
                string spooled = shared.RecordOperation(latestLines == null ? Rdv3Text.OpCreate : Rdv3Text.OpUpdate,
                    update.Lines.Length, Rdv3OperationLog.UpdateDetail(source.Job, update, work));
                trace("oplog", spooled == null ? "written " + shared.Operations.Path : spooled);
                string operationWarning = Rdv3OperationLog.FailureNotice(spooled);
                if (operationWarning != null) { warnings.Add(operationWarning); }
                marker = shared.WriteMarker("update", update.Lines.Length, 0, 0);
                trace("marker", "version=" + marker.Version.ToString(CultureInfo.InvariantCulture) + " kind=update");
            }
            else { CheckWritable(protection); trace("persist", "skipped (latest ledger already has this result)"); }
            lease.Release();
            return new Rdv3ApplyOutcome(update, marker, committed, null, warnings.ToArray());
        }
        catch (Exception error)
        {
            return new Rdv3ApplyOutcome(update, marker, committed, error, warnings.ToArray());
        }
        finally
        {
            if (lease != null) { lease.Dispose(); }
        }
    }
}
