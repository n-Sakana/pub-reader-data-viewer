using System;
using System.Globalization;
using System.IO;

internal sealed class Rdv3LedgerSnapshot
{
    public readonly string[] Lines;
    public readonly string[] States;
    public readonly string Warning;

    public Rdv3LedgerSnapshot(string[] lines, string[] states, string warning)
    {
        Lines = lines;
        States = states;
        Warning = warning;
    }
}

internal sealed class Rdv3ApplyOutcome
{
    public readonly Rdv3UpdateResult Update;
    public readonly Rdv3SharedMarker Marker;
    public readonly bool Committed;
    public readonly Exception Error;

    public Rdv3ApplyOutcome(Rdv3UpdateResult update, Rdv3SharedMarker marker, bool committed, Exception error)
    {
        Update = update;
        Marker = marker;
        Committed = committed;
        Error = error;
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

    public Rdv3LedgerStore(string ledgerPath, Rdv3Data definition, Rdv3WorkState state, Rdv3SharedFiles sharedFiles)
    {
        path = ledgerPath;
        data = definition;
        work = state;
        shared = sharedFiles;
        contract = Rdv3Files.StorageContract(data, work);
    }

    public Rdv3LedgerSnapshot Read(string[] head)
    {
        string[] lines, states;
        string warning;
        Rdv3Xlsx.Read(path, head, work.Column, out lines, out states, out warning, contract);
        string reference = data.Columns[data.IdentityCol].Ref;
        string label = data.LabelOf(reference);
        Rdv3Ledger.CheckIdentities(lines, data.IdentityCol, Path.GetFileName(path), label.Length == 0 ? reference : label);
        for (int i = 0; i < states.Length; i++)
        {
            if (work.ByStored(states[i]) == null)
            { throw new InvalidDataException(Rdv3Text.LedgerStateInvalid + (i + 2).ToString(CultureInfo.InvariantCulture)); }
        }
        return new Rdv3LedgerSnapshot(lines, states, warning);
    }

    public Rdv3ApplyOutcome Apply(Rdv3MergeResult source, string[] checkedLines, string tag,
                                  Func<Rdv3LedgerLock> acquire, Action<string, string> trace, Action<string> warn)
    {
        Rdv3LedgerLock lease = null;
        Rdv3UpdateResult update = null;
        Rdv3SharedMarker marker = null;
        bool committed = false;
        try
        {
            lease = acquire();
            if (lease == null) { throw new InvalidOperationException("shared lease was not acquired"); }
            long t = Rdv3Clock.Now();
            string[] latestLines = null, latestStates = null;
            if (Rdv3Files.Exists(path))
            {
                Rdv3LedgerSnapshot latest = Read(source.Head);
                latestLines = latest.Lines;
                latestStates = latest.States;
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
            update = source.Prepared == null
                ? Rdv3Ledger.ApplyUpdate(source.Job, latestLines, latestStates, source.Lines, data.IdentityCol, work.InitialStored)
                : Rdv3Process.Execute(source.Prepared, latestLines, latestStates, work.InitialStored, false).Update;
            string operation = source.Job.ApplyStep == null ? "pipeline" : source.Job.ApplyStep.Operation;
            trace("apply", "operation=" + operation
                + " source=" + source.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " result=" + update.Lines.Length.ToString(CultureInfo.InvariantCulture)
                + " added=" + update.Added.ToString(CultureInfo.InvariantCulture)
                + " updated=" + update.Updated.ToString(CultureInfo.InvariantCulture)
                + " unchanged=" + update.Unchanged.ToString(CultureInfo.InvariantCulture)
                + " kept=" + update.Kept.ToString(CultureInfo.InvariantCulture)
                + " deleted=" + update.Deleted.ToString(CultureInfo.InvariantCulture)
                + " reset=" + update.ResetLines.Count.ToString(CultureInfo.InvariantCulture)
                + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
            if (latestLines == null || !Rdv3Ledger.SameLedger(latestLines, latestStates, update.Lines, update.States))
            {
                t = Rdv3Clock.Now();
                Rdv3Xlsx.Write(path, source.Head, work.Column, update.Lines, update.States, tag, contract);
                committed = true;
                trace("persist", "target=xlsx rows=" + update.Lines.Length.ToString(CultureInfo.InvariantCulture)
                    + " ms=" + Rdv3Log.F(Rdv3Clock.MsSince(t)));
                marker = shared.WriteMarker("update", update.Lines.Length, 0, 0);
                trace("marker", "version=" + marker.Version.ToString(CultureInfo.InvariantCulture) + " kind=update");
            }
            else { trace("persist", "skipped (latest ledger already has this result)"); }
            lease.Release();
            return new Rdv3ApplyOutcome(update, marker, committed, null);
        }
        catch (Exception error)
        {
            return new Rdv3ApplyOutcome(update, marker, committed, error);
        }
        finally
        {
            if (lease != null) { lease.Dispose(); }
        }
    }
}
