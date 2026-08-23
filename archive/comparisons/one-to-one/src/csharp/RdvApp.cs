// ============================================================================
// RdvApp.cs -- entry point of the C#-only build.
//
// Compilation and window creation happen here, before the watcher starts, so
// none of it can leak into a merge-select measurement.
// ============================================================================

using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;

// ---- method 2: C# only -----------------------------------------------------
public static class RdvProgramForms
{
    public static int Run(string dataDir, double compileMs, string logPath)
    {
        RdvEngine eng = new RdvEngine(dataDir);
        string err;
        if (!RdvBoot.Preflight(eng, out err))
        {
            MessageBox.Show(err, RdvText.AppTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        RdvForm form = new RdvForm(RdvText.MethodCs, dataDir, RdvBoot.DataNote(eng, compileMs, ""));
        RdvRunner runner = new RdvRunner(eng, logPath);
        runner.Paint = form.ShowRecord;
        runner.Stats = form.ShowStats;
        runner.Busy = delegate(string s) { form.SetState(s, ""); };

        RdvWatch watch = new RdvWatch();
        watch.OnConfirmed = runner.Handle;
        watch.OnState = form.SetState;
        watch.OnRaw = form.SetRaw;

        form.OnRebind = watch.Rebind;
        form.OnManual = delegate(string k)
        {
            long t0 = Stopwatch.GetTimestamp();
            ThreadPool.QueueUserWorkItem(delegate(object o) { runner.Handle(k, 0, 0, t0); });
        };
        form.Shown += delegate(object s, EventArgs e) { watch.Start(); };
        form.FormClosing += delegate(object s, FormClosingEventArgs e) { watch.Stop(); };

        Application.Run(form);
        return 0;
    }
}
