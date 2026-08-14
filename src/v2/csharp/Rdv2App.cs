// ============================================================================
// Rdv2App.cs -- the order of events, and the entry point.
//
//   t0                  a number was confirmed
//   stages 1..9         read three CSVs, build three indexes, both joins, search
//   stage 10            present: one candidate -> the record; several -> the
//                       candidate list, built and filled
//   stop clock          total = now - t0. This is the figure the four methods
//                       are compared on, and it never contains a human.
//   (dialog)            the operator reads and picks. Not measured at all.
//   pick                the chosen candidate's record, timed separately.
//
// One merge-select at a time, and nothing is carried over between them: every
// trigger reads all three files again and rebuilds all three indexes.
// ============================================================================

using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

public sealed class Rdv2Runner
{
    private readonly Rdv2Engine eng;
    private readonly Rdv2Form ui;
    private readonly string logPath;
    private int seq;
    private int busy;

    public Rdv2Runner(Rdv2Engine engine, Rdv2Form form, string log)
    {
        eng = engine;
        ui = form;
        logPath = log;
    }

    public void Handle(string key, double detectMs, int polls, long t0)
    {
        if (Interlocked.CompareExchange(ref busy, 1, 0) != 0) { return; }
        try
        {
            ui.SetState("BUSY", "");
            Rdv2Run r = eng.Execute(key, t0);
            r.Seq = Interlocked.Increment(ref seq);
            r.DetectMs = detectMs;
            r.Polls = polls;

            long m = Rdv2Clock.Now();
            if (r.Error.Length > 0 || r.CandCount == 0)
            {
                ui.ShowRecord(r);
            }
            else
            {
                eng.FillCandidates(r.Hit);
                if (r.CandCount == 1)
                {
                    eng.FillRecord(r.Hit, 0);
                    ui.ShowRecord(r);
                }
                else
                {
                    ui.PreparePick(r);
                }
            }
            r.Stage[Rdv2Spec.StShow] = Rdv2Clock.MsSince(m);
            r.TotalMs = Rdv2Clock.MsSince(t0);
            ui.ShowStats(r);

            if (r.Error.Length == 0 && r.CandCount > 1)
            {
                ui.SetState("PICK", "");
                int chosen = ui.RunPick();
                if (chosen >= 0)
                {
                    long p0 = Rdv2Clock.Now();
                    eng.FillRecord(r.Hit, chosen);
                    ui.ShowRecord(r);
                    r.PickMs = Rdv2Clock.MsSince(p0);
                    ui.ShowStats(r);
                }
            }
            Log(r);
            ui.SetState("WATCHING", "");
        }
        finally
        {
            Interlocked.Exchange(ref busy, 0);
        }
    }

    // opt-in only: nothing is written unless -log was passed
    private void Log(Rdv2Run r)
    {
        if (logPath == null || logPath.Length == 0) { return; }
        try
        {
            StringBuilder sb = new StringBuilder();
            if (!File.Exists(logPath))
            {
                sb.Append("seq\ttime\tkey\tindex\t");
                for (int i = 0; i < Rdv2Spec.StageCount; i++) { sb.Append(Rdv2Spec.StageKey[i]).Append('\t'); }
                sb.Append("other\ttotal\tpick\tdetect\tpolls\tcands\trows\tprobes\tchecksum\tmatchedAB\tmatchedBC\toracle\terror\r\n");
            }
            sb.Append(r.Seq.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.When.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Key).Append('\t');
            sb.Append(r.IndexKind).Append('\t');
            for (int i = 0; i < Rdv2Spec.StageCount; i++)
            {
                sb.Append(r.Stage[i].ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            }
            sb.Append((r.TotalMs - r.StageSum()).ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.TotalMs.ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.PickMs.ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.DetectMs.ToString("F2", CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.Polls.ToString(CultureInfo.InvariantCulture)).Append('\t');
            sb.Append(r.CandCount.ToString(CultureInfo.InvariantCulture)).Append('\t');
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

public static class Rdv2Program
{
    public static int Run(string dataDir, double compileMs, string logPath, string autoPick)
    {
        if (IntPtr.Size != 8)
        {
            MessageBox.Show(Rdv2Text.ErrNo64, Rdv2Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }
        Rdv2Engine eng = new Rdv2Engine(dataDir);
        try { eng.CheckFiles(); }
        catch (Exception ex)
        {
            MessageBox.Show(Rdv2Text.ErrNoData + ex.Message, Rdv2Text.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 3;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        string method = (RdvIndex.Kind == "dict") ? Rdv2Text.MethodDict : Rdv2Text.MethodHash;
        StringBuilder note = new StringBuilder();
        note.Append(eng.Expected.Loaded ? eng.Expected.Rows.ToString("N0", CultureInfo.InvariantCulture) : "?");
        note.Append(' ').Append(Rdv2Text.BootData);
        note.Append("    ").Append(Rdv2Text.LabelIndex).Append(' ').Append(RdvIndex.Kind);
        note.Append("    ").Append(Rdv2Text.BootCompiled).Append(' ');
        note.Append((compileMs / 1000.0).ToString("N2", CultureInfo.InvariantCulture)).Append(" s");

        Rdv2Form form = new Rdv2Form(method, dataDir, note.ToString());
        Rdv2Runner runner = new Rdv2Runner(eng, form, logPath);

        // -autopick <n> answers the dialog for an unattended measurement run.
        // It changes nothing that is measured: the choice is made after the
        // clock has stopped, exactly where a person would have made it.
        if (autoPick != null && autoPick.Length > 0)
        {
            int which;
            // a negative value means "leave the dialog to a person", the same
            // as not passing it at all
            if (int.TryParse(autoPick, NumberStyles.Integer, CultureInfo.InvariantCulture, out which) && which >= 0)
            {
                Rdv2AutoPick.Install(form, which);
            }
        }

        Rdv2Watch watch = new Rdv2Watch();
        watch.OnConfirmed = runner.Handle;
        watch.OnState = form.SetState;
        watch.OnRaw = form.SetRaw;
        form.OnRebind = watch.Rebind;
        form.OnManual = delegate(string k)
        {
            long t0 = Rdv2Clock.Now();
            ThreadPool.QueueUserWorkItem(delegate(object o) { runner.Handle(k, 0, 0, t0); });
        };
        form.Shown += delegate(object s, EventArgs e) { watch.Start(); };
        form.FormClosing += delegate(object s, FormClosingEventArgs e) { watch.Stop(); };

        Application.Run(form);
        return 0;
    }
}

// Presses the candidate dialog's button for a benchmark run. A timer on the UI
// thread notices the modal form and picks the n-th row.
public static class Rdv2AutoPick
{
    public static void Install(Form owner, int which)
    {
        // qualified: the packer hoists every using into one file, where a bare
        // Timer is ambiguous with System.Threading.Timer
        System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
        t.Interval = 120;
        t.Tick += delegate(object s, EventArgs e)
        {
            FormCollection fc = Application.OpenForms;
            for (int i = 0; i < fc.Count; i++)
            {
                Rdv2PickForm p = fc[i] as Rdv2PickForm;
                if (p != null && p.Visible)
                {
                    p.PickRow(which);
                    return;
                }
            }
        };
        t.Start();
    }
}
