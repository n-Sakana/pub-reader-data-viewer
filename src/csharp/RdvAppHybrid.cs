// ============================================================================
// RdvAppHybrid.cs -- entry point of the hybrid build: this process watches,
// reads, joins and measures; the workbook shows the answer.
//
// Threads, and why there are three:
//   main    STA, runs the little control panel. Nothing heavy happens here.
//   engine  STA, creates Excel and does every merge-select and every COM write.
//           Same apartment as the Excel instance it made, so the calls are
//           direct rather than marshalled across an apartment boundary.
//   watch   MTA, polls Notepad through UI Automation and hands a confirmed
//           number to the engine thread.
// ============================================================================

using System;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

public sealed class RdvPanel : Form
{
    private readonly Label lblState = new Label();
    private readonly Label lblNotepad = new Label();
    private readonly Label lblData = new Label();
    private readonly Label lblLast = new Label();
    private readonly Label lblErr = new Label();
    private readonly Label lblHint = new Label();
    public Action OnRebind;

    public RdvPanel(string dataDir, string note)
    {
        Text = RdvText.AppTitle + " - " + RdvText.MethodHybrid;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(560, 250);
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        BackColor = Color.White;
        try { Font = new Font("Yu Gothic UI", 9.5f); }
        catch (Exception) { }

        Label t = new Label();
        t.Text = RdvText.AppTitle + "   /   " + RdvText.MethodHybrid;
        t.Font = new Font(Font.FontFamily, 12f, FontStyle.Bold);
        t.AutoSize = true;
        t.Location = new Point(14, 12);
        Controls.Add(t);

        lblState.Text = RdvText.StateBoot;
        lblState.Font = new Font(Font.FontFamily, 11f, FontStyle.Bold);
        lblState.AutoSize = true;
        lblState.Location = new Point(16, 44);
        Controls.Add(lblState);

        lblNotepad.Text = RdvText.NotepadNone;
        lblNotepad.AutoSize = true;
        lblNotepad.Location = new Point(16, 74);
        Controls.Add(lblNotepad);

        lblData.Text = RdvText.LabelData + " " + dataDir + "   " + note;
        lblData.AutoSize = false;
        lblData.Size = new Size(520, 34);
        lblData.Location = new Point(16, 96);
        Controls.Add(lblData);

        lblLast.Text = "";
        lblLast.AutoSize = true;
        lblLast.Font = new Font("Consolas", 10f, FontStyle.Bold);
        lblLast.Location = new Point(16, 136);
        Controls.Add(lblLast);

        lblHint.Text = RdvText.HybridHint;
        lblHint.AutoSize = true;
        lblHint.ForeColor = Color.FromArgb(96, 104, 116);
        lblHint.Location = new Point(16, 164);
        Controls.Add(lblHint);

        lblErr.Text = "";
        lblErr.ForeColor = Color.FromArgb(178, 40, 44);
        lblErr.AutoSize = false;
        lblErr.Size = new Size(520, 30);
        lblErr.Location = new Point(16, 188);
        Controls.Add(lblErr);

        Button b = new Button();
        b.Text = RdvText.BtnRebind;
        b.Size = new Size(150, 26);
        b.Location = new Point(16, 214);
        b.FlatStyle = FlatStyle.System;
        b.Click += delegate(object s, EventArgs e) { if (OnRebind != null) { OnRebind(); } };
        Controls.Add(b);
    }

    public void SetState(string state, string detail)
    {
        if (IsDisposed) { return; }
        try
        {
            Invoke(new Action(delegate
            {
                if (state == "WATCHING")
                {
                    lblState.Text = RdvText.StateWatching;
                    lblState.ForeColor = Color.FromArgb(22, 122, 72);
                    if (detail.Length > 0) { lblNotepad.Text = RdvText.LabelNotepad + " " + detail; }
                }
                else if (state == "WAITING")
                {
                    lblState.Text = RdvText.StateWaiting;
                    lblState.ForeColor = Color.FromArgb(176, 104, 8);
                    lblNotepad.Text = RdvText.NotepadNone;
                }
                else if (state == "BUSY")
                {
                    lblState.Text = RdvText.StateBusy;
                    lblState.ForeColor = Color.FromArgb(176, 104, 8);
                }
                else { lblState.Text = state; }
            }));
        }
        catch (Exception) { }
    }

    public void SetLast(RdvRun r)
    {
        if (IsDisposed) { return; }
        try
        {
            Invoke(new Action(delegate
            {
                lblLast.Text = "#" + r.Seq.ToString(CultureInfo.InvariantCulture) + "  " + r.Key
                    + "   " + RdvEngine.FmtSec(r.TotalMs) + " s";
                lblErr.Text = r.Error;
            }));
        }
        catch (Exception) { }
    }
}

public static class RdvProgramHybrid
{
    private static readonly object gate = new object();
    private static string qKey;
    private static double qDetect;
    private static int qPolls;
    private static long qT0;
    private static volatile bool stop;

    public static int Run(string dataDir, string workbookPath, double compileMs, string logPath)
    {
        RdvEngine eng = new RdvEngine(dataDir);
        string err;
        if (!RdvBoot.Preflight(eng, out err))
        {
            MessageBox.Show(err, RdvText.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }
        if (!File.Exists(workbookPath))
        {
            MessageBox.Show(RdvText.ErrNoBook + workbookPath, RdvText.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 3;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        RdvExcelFront front = new RdvExcelFront();
        double excelMs = 0;
        Exception startErr = null;
        long b0 = Stopwatch.GetTimestamp();

        // Excel is started on the engine thread so that every later COM call
        // stays inside the apartment that created it
        RdvPanel panel = null;
        RdvRunner runner = new RdvRunner(eng, logPath);
        ManualResetEvent ready = new ManualResetEvent(false);

        Thread engine = new Thread(delegate()
        {
            try
            {
                front.Start(workbookPath);
                excelMs = RdvEngine.MsSince(b0);
            }
            catch (Exception ex) { startErr = ex; }
            ready.Set();
            if (startErr != null) { return; }

            while (!stop)
            {
                string k = null;
                double d = 0; int p = 0; long t0 = 0;
                lock (gate)
                {
                    if (qKey != null) { k = qKey; d = qDetect; p = qPolls; t0 = qT0; qKey = null; }
                }
                if (k != null)
                {
                    runner.Handle(k, d, p, t0);
                }
                else
                {
                    Application.DoEvents();     // keep the STA pumping for COM
                    Thread.Sleep(5);
                }
            }
            front.Shutdown();
        });
        engine.SetApartmentState(ApartmentState.STA);
        engine.IsBackground = true;
        engine.Name = "rdv-engine";
        engine.Start();
        ready.WaitOne(120000);

        if (startErr != null)
        {
            MessageBox.Show(RdvText.ErrNoExcel + RdvExcelFront.Explain(startErr), RdvText.AppTitle,
                MessageBoxButtons.OK, MessageBoxIcon.Error);
            stop = true;
            return 4;
        }

        string note = RdvBoot.DataNote(eng, compileMs, RdvText.BootExcel + " "
            + RdvEngine.FmtSec(excelMs) + " s");
        panel = new RdvPanel(dataDir, note);

        RdvWatch watch = new RdvWatch();
        watch.OnState = delegate(string s, string dt)
        {
            panel.SetState(s, dt);
            try { front.Header(RdvText.StateWatching, dt, dataDir); }
            catch (Exception) { }
        };
        watch.OnConfirmed = delegate(string key, double detect, int polls, long t0)
        {
            lock (gate) { qKey = key; qDetect = detect; qPolls = polls; qT0 = t0; }
        };
        runner.Paint = front.ShowRecord;
        runner.Stats = delegate(RdvRun r) { front.ShowStats(r); panel.SetLast(r); };
        runner.Busy = delegate(string s) { panel.SetState(s, ""); };
        panel.OnRebind = watch.Rebind;

        try { front.Header(RdvText.StateWatching, "", dataDir); }
        catch (Exception) { }

        // if the operator closes Excel, this process has nothing left to show
        System.Windows.Forms.Timer alive = new System.Windows.Forms.Timer();
        alive.Interval = 1500;
        alive.Tick += delegate(object s, EventArgs e)
        {
            if (!front.Alive())
            {
                alive.Stop();
                panel.SetState(RdvText.ExcelClosed, "");
                panel.Close();
            }
        };

        panel.Shown += delegate(object s, EventArgs e) { watch.Start(); alive.Start(); };
        panel.FormClosing += delegate(object s, FormClosingEventArgs e)
        {
            alive.Stop();
            watch.Stop();
            stop = true;
            engine.Join(20000);
        };

        Application.Run(panel);
        return 0;
    }
}
