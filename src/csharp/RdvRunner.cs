// ============================================================================
// RdvRunner.cs -- what happens between "a number was confirmed" and "the screen
// is up". Shared by both .cmd builds; only the two delegates differ.
//
// The order below is the whole measurement contract:
//
//   t0            the watcher confirmed a number (start of the merge-select)
//   stages 1..6   read three CSVs, join A-B, join B-C, look one row up
//   stage 7       paint it
//   stop clock    total = now - t0
//   afterwards    paint the timings, append to the log
//
// Nothing is reused between runs. The engine allocates fresh tables every time
// and lets the old ones go, on purpose: this measures the cost of doing the
// whole job, not the cost of remembering last time's answer.
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

public sealed class RdvRunner
{
    private readonly RdvEngine eng;
    private readonly string logPath;
    private int seq;
    private int busy;

    public Action<RdvRun> Paint;    // stage 7: must not return until it is on screen
    public Action<RdvRun> Stats;    // after the clock stops
    public Action<string> Busy;     // "BUSY" / "WATCHING"

    public RdvRunner(RdvEngine engine, string log)
    {
        eng = engine;
        logPath = log;
    }

    public RdvEngine Engine { get { return eng; } }

    public void Handle(string key, double detectMs, int polls, long t0)
    {
        // one merge-select at a time; a second trigger while busy is dropped
        if (Interlocked.CompareExchange(ref busy, 1, 0) != 0) { return; }
        try
        {
            if (Busy != null) { Busy("BUSY"); }
            RdvRun r = eng.Execute(key, t0);
            r.Seq = Interlocked.Increment(ref seq);
            r.DetectMs = detectMs;
            r.Polls = polls;

            long m = Stopwatch.GetTimestamp();
            if (Paint != null) { Paint(r); }
            r.Stage[RdvSpec.StageShow] = RdvEngine.MsSince(m);
            r.TotalMs = RdvEngine.MsSince(t0);

            if (Stats != null) { Stats(r); }
            Log(r);
            if (Busy != null) { Busy("WATCHING"); }
        }
        finally
        {
            Interlocked.Exchange(ref busy, 0);
        }
    }

    // opt-in only: the app writes nothing unless a log path was passed on the
    // command line. It never writes to the CSVs or to a workbook.
    private void Log(RdvRun r)
    {
        if (logPath == null || logPath.Length == 0) { return; }
        try
        {
            StringBuilder sb = new StringBuilder();
            if (!File.Exists(logPath))
            {
                sb.Append("seq\ttime\tkey\t");
                for (int i = 0; i < RdvSpec.StageCount; i++) { sb.Append(RdvSpec.StageKey[i]).Append('\t'); }
                sb.Append("other\ttotal\tdetect\tpolls\trows\tprobes\tchecksum\tmatchedAB\tmatchedBC\toracle\terror\r\n");
            }
            sb.Append(r.Seq.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.When.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Key).Append('\t');
            for (int i = 0; i < RdvSpec.StageCount; i++)
            {
                sb.Append(r.Stage[i].ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            }
            sb.Append((r.TotalMs - r.StageSum()).ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.TotalMs.ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.DetectMs.ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Polls.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Rows.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Probes.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Checksum.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.MatchedAB.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.MatchedBC.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.OracleOk ? "ok" : ("NG:" + r.OracleNote)).Append('\t');
            sb.Append(r.Error.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' '));
            sb.Append("\r\n");
            File.AppendAllText(logPath, sb.ToString(), new UTF8Encoding(false));
        }
        catch (Exception) { }
    }
}

public static class RdvBoot
{
    public static string DataNote(RdvEngine eng, double compileMs, string extra)
    {
        StringBuilder sb = new StringBuilder();
        int rows = eng.Expected.Loaded ? eng.Expected.Rows : 0;
        sb.Append(rows.ToString("N0", CultureInfo.InvariantCulture)).Append(' ').Append(RdvText.BootData);
        sb.Append("    ").Append(RdvText.BootCompiled).Append(' ');
        sb.Append((compileMs / 1000.0).ToString("N2", CultureInfo.InvariantCulture)).Append(" s");
        if (extra != null && extra.Length > 0) { sb.Append("    ").Append(extra); }
        return sb.ToString();
    }

    public static bool Preflight(RdvEngine eng, out string error)
    {
        error = "";
        if (IntPtr.Size != 8) { error = RdvText.ErrNo64; return false; }
        try { eng.CheckFiles(); }
        catch (Exception ex) { error = RdvText.ErrNoData + ex.Message; return false; }
        return true;
    }
}

