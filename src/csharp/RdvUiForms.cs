// ============================================================================
// RdvUiForms.cs -- the WinForms screen for the C#-only build.
//
// Display is stage 7 of the merge-select, so ShowRecord() has to be honest
// about it: the record is pushed to the controls on the UI thread and the form
// is repainted before the clock stops. The timings themselves are painted
// afterwards, outside the measured region -- otherwise the screen would be
// timing the drawing of its own stopwatch.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

public sealed class RdvForm : Form
{
    private readonly Label lblMethod = new Label();
    private readonly Label lblState = new Label();
    private readonly Label lblNotepad = new Label();
    private readonly Label lblData = new Label();
    private readonly Label lblKey = new Label();
    private readonly Label lblVerdict = new Label();
    private readonly Label lblRaw = new Label();
    private readonly Label lblError = new Label();
    private readonly ListView lvRec = new ListView();
    private readonly ListView lvTime = new ListView();
    private readonly ListView lvHist = new ListView();
    private readonly TextBox txtManual = new TextBox();
    private readonly Button btnManual = new Button();
    private readonly Button btnRebind = new Button();

    public Action<string> OnManual;
    public Action OnRebind;

    private static readonly Color Ink = Color.FromArgb(28, 32, 38);
    private static readonly Color Sub = Color.FromArgb(96, 104, 116);
    private static readonly Color Line = Color.FromArgb(214, 219, 226);
    private static readonly Color Paper = Color.FromArgb(248, 249, 251);
    private static readonly Color Good = Color.FromArgb(22, 122, 72);
    private static readonly Color Warn = Color.FromArgb(176, 104, 8);
    private static readonly Color Bad = Color.FromArgb(178, 40, 44);

    private static Font UiFont(float size, FontStyle st)
    {
        try { return new Font("Yu Gothic UI", size, st); }
        catch (Exception) { return new Font(FontFamily.GenericSansSerif, size, st); }
    }

    private static Font MonoFont(float size, FontStyle st)
    {
        try { return new Font("Consolas", size, st); }
        catch (Exception) { return new Font(FontFamily.GenericMonospace, size, st); }
    }

    public RdvForm(string method, string dataDir, string dataNote)
    {
        Text = RdvText.AppTitle + " - " + method;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1180, 800);
        MinimumSize = new Size(920, 640);
        BackColor = Color.White;
        Font = UiFont(9.5f, FontStyle.Regular);

        TableLayoutPanel root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 1;
        root.RowCount = 5;
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 250));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        root.Padding = new Padding(14, 12, 14, 10);
        Controls.Add(root);

        // ---- header ---------------------------------------------------------
        Panel head = new Panel();
        head.Dock = DockStyle.Fill;
        lblMethod.Text = RdvText.AppTitle + "   /   " + method;
        lblMethod.Font = UiFont(14f, FontStyle.Bold);
        lblMethod.ForeColor = Ink;
        lblMethod.AutoSize = true;
        lblMethod.Location = new Point(0, 0);
        head.Controls.Add(lblMethod);

        lblState.Text = RdvText.StateBoot;
        lblState.Font = UiFont(11f, FontStyle.Bold);
        lblState.TextAlign = ContentAlignment.MiddleRight;
        lblState.Dock = DockStyle.Right;
        lblState.Width = 320;
        lblState.ForeColor = Sub;
        head.Controls.Add(lblState);

        lblNotepad.Text = RdvText.NotepadNone;
        lblNotepad.ForeColor = Sub;
        lblNotepad.AutoSize = true;
        lblNotepad.Location = new Point(2, 34);
        head.Controls.Add(lblNotepad);

        lblData.Text = RdvText.LabelData + " " + dataDir + "    " + dataNote;
        lblData.ForeColor = Sub;
        lblData.AutoSize = true;
        lblData.Location = new Point(2, 56);
        head.Controls.Add(lblData);

        btnRebind.Text = RdvText.BtnRebind;
        btnRebind.Size = new Size(150, 26);
        btnRebind.Location = new Point(2, 76);
        btnRebind.FlatStyle = FlatStyle.System;
        btnRebind.Click += delegate(object s, EventArgs e) { if (OnRebind != null) { OnRebind(); } };
        head.Controls.Add(btnRebind);

        txtManual.Location = new Point(166, 78);
        txtManual.Width = 110;
        txtManual.Font = MonoFont(10f, FontStyle.Regular);
        txtManual.MaxLength = RdvSpec.KeyLen;
        head.Controls.Add(txtManual);

        btnManual.Text = RdvText.BtnManual;
        btnManual.Size = new Size(120, 26);
        btnManual.Location = new Point(282, 76);
        btnManual.FlatStyle = FlatStyle.System;
        btnManual.Click += delegate(object s, EventArgs e)
        {
            string k = txtManual.Text.Trim();
            if (OnManual != null && RdvWatch.IsKey(k)) { OnManual(k); }
        };
        head.Controls.Add(btnManual);

        lblRaw.Text = "";
        lblRaw.ForeColor = Sub;
        lblRaw.Font = MonoFont(9f, FontStyle.Regular);
        lblRaw.AutoSize = true;
        lblRaw.Location = new Point(412, 82);
        head.Controls.Add(lblRaw);
        root.Controls.Add(head, 0, 0);

        // ---- current key ----------------------------------------------------
        Panel keyp = new Panel();
        keyp.Dock = DockStyle.Fill;
        keyp.BackColor = Paper;
        keyp.Padding = new Padding(12, 6, 12, 6);

        Label cap = new Label();
        cap.Text = RdvText.LabelCurrent;
        cap.ForeColor = Sub;
        cap.AutoSize = true;
        cap.Location = new Point(14, 8);
        keyp.Controls.Add(cap);

        lblKey.Text = "--------";
        lblKey.Font = MonoFont(24f, FontStyle.Bold);
        lblKey.ForeColor = Ink;
        lblKey.AutoSize = true;
        lblKey.Location = new Point(12, 24);
        keyp.Controls.Add(lblKey);

        lblVerdict.Text = "";
        lblVerdict.Font = UiFont(12f, FontStyle.Bold);
        lblVerdict.AutoSize = true;
        lblVerdict.Location = new Point(230, 32);
        keyp.Controls.Add(lblVerdict);
        root.Controls.Add(keyp, 0, 1);

        // ---- joined record --------------------------------------------------
        SetupList(lvRec);
        lvRec.Columns.Add("#", 34, HorizontalAlignment.Right);
        lvRec.Columns.Add(RdvText.ColFieldA, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(RdvText.ColValueA, 190, HorizontalAlignment.Left);
        lvRec.Columns.Add(RdvText.ColFieldB, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(RdvText.ColValueB, 190, HorizontalAlignment.Left);
        lvRec.Columns.Add(RdvText.ColFieldC, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(RdvText.ColValueC, 190, HorizontalAlignment.Left);
        for (int i = 0; i < RdvSpec.Fields; i++)
        {
            ListViewItem it = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
            for (int k = 0; k < 6; k++) { it.SubItems.Add(""); }
            lvRec.Items.Add(it);
        }
        GroupBox gRec = Boxed(RdvText.BoxRecord, lvRec);
        root.Controls.Add(gRec, 0, 2);

        // ---- timings + history ----------------------------------------------
        TableLayoutPanel bottom = new TableLayoutPanel();
        bottom.Dock = DockStyle.Fill;
        bottom.ColumnCount = 2;
        bottom.RowCount = 1;
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 430));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        SetupList(lvTime);
        lvTime.Columns.Add(RdvText.ColStage, 210, HorizontalAlignment.Left);
        lvTime.Columns.Add(RdvText.ColMs, 92, HorizontalAlignment.Right);
        lvTime.Columns.Add(RdvText.ColShare, 76, HorizontalAlignment.Right);
        bottom.Controls.Add(Boxed(RdvText.BoxStages, lvTime), 0, 0);

        SetupList(lvHist);
        lvHist.Columns.Add("#", 40, HorizontalAlignment.Right);
        lvHist.Columns.Add(RdvText.ColTime, 92, HorizontalAlignment.Left);
        lvHist.Columns.Add(RdvText.ColKey, 92, HorizontalAlignment.Left);
        lvHist.Columns.Add(RdvText.ColTotal, 88, HorizontalAlignment.Right);
        lvHist.Columns.Add(RdvText.ColDetect, 88, HorizontalAlignment.Right);
        lvHist.Columns.Add(RdvText.ColCheck, 150, HorizontalAlignment.Left);
        bottom.Controls.Add(Boxed(RdvText.BoxHistory, lvHist), 1, 0);
        root.Controls.Add(bottom, 0, 3);

        // ---- error ------------------------------------------------------------
        lblError.Dock = DockStyle.Fill;
        lblError.ForeColor = Bad;
        lblError.Font = UiFont(9.5f, FontStyle.Bold);
        lblError.TextAlign = ContentAlignment.MiddleLeft;
        lblError.Text = "";
        root.Controls.Add(lblError, 0, 4);
    }

    private static void SetupList(ListView lv)
    {
        lv.Dock = DockStyle.Fill;
        lv.View = View.Details;
        lv.FullRowSelect = true;
        lv.GridLines = true;
        lv.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        lv.MultiSelect = false;
        lv.HideSelection = true;
        lv.BorderStyle = BorderStyle.None;
        lv.Font = MonoFont(9.5f, FontStyle.Regular);
    }

    private static GroupBox Boxed(string title, Control inner)
    {
        GroupBox g = new GroupBox();
        g.Text = title;
        g.Dock = DockStyle.Fill;
        g.Padding = new Padding(8, 6, 8, 8);
        g.ForeColor = Sub;
        inner.ForeColor = Ink;
        g.Controls.Add(inner);
        return g;
    }

    // ---- called from the watcher thread ------------------------------------
    public void SetState(string state, string detail)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate
        {
            if (state == "WATCHING")
            {
                lblState.Text = RdvText.StateWatching;
                lblState.ForeColor = Good;
                lblNotepad.Text = RdvText.LabelNotepad + " " + detail;
            }
            else if (state == "WAITING")
            {
                lblState.Text = RdvText.StateWaiting;
                lblState.ForeColor = Warn;
                lblNotepad.Text = RdvText.NotepadNone;
            }
            else if (state == "BUSY")
            {
                lblState.Text = RdvText.StateBusy;
                lblState.ForeColor = Warn;
            }
            else
            {
                lblState.Text = state;
                lblState.ForeColor = Sub;
            }
        }));
    }

    public void SetRaw(string s)
    {
        if (IsDisposed) { return; }
        BeginInvoke(new Action(delegate
        {
            lblRaw.Text = RdvText.LabelField + " [" + s + "]";
        }));
    }

    public void SetError(string s)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate { lblError.Text = s; }));
    }

    // stage 7. Everything the operator has to read is on screen when this
    // returns: values in the grid, and the pixels actually painted.
    public void ShowRecord(RdvRun run)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate
        {
            lblKey.Text = run.Key;
            RdvHit h = run.Hit;
            if (run.Error.Length > 0)
            {
                lblVerdict.Text = RdvText.VerdictError;
                lblVerdict.ForeColor = Bad;
            }
            else if (h != null && h.Found)
            {
                lblVerdict.Text = RdvText.VerdictHit + "   key2 = " + h.Key2;
                lblVerdict.ForeColor = Good;
            }
            else
            {
                lblVerdict.Text = RdvText.VerdictMiss;
                lblVerdict.ForeColor = Bad;
            }

            lvRec.BeginUpdate();
            for (int i = 0; i < RdvSpec.Fields; i++)
            {
                ListViewItem it = lvRec.Items[i];
                it.SubItems[1].Text = Pick(h == null ? null : h.NameA, i);
                it.SubItems[2].Text = Pick(h == null ? null : h.ValA, i);
                it.SubItems[3].Text = Pick(h == null ? null : h.NameB, i);
                it.SubItems[4].Text = Pick(h == null ? null : h.ValB, i);
                it.SubItems[5].Text = Pick(h == null ? null : h.NameC, i);
                it.SubItems[6].Text = Pick(h == null ? null : h.ValC, i);
            }
            lvRec.EndUpdate();
            Refresh();
        }));
    }

    private static string Pick(string[] a, int i)
    {
        if (a == null || i >= a.Length) { return ""; }
        return a[i];
    }

    // painted after the clock has stopped
    public void ShowStats(RdvRun run)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate
        {
            lvTime.BeginUpdate();
            lvTime.Items.Clear();
            double total = run.TotalMs;
            for (int i = 0; i < RdvSpec.StageCount; i++)
            {
                AddTime(RdvText.StageName[i], run.Stage[i], total, false);
            }
            double other = total - run.StageSum();
            AddTime(RdvText.StageOther, other, total, false);
            AddTime(RdvText.StageTotal, total, total, true);
            lvTime.Items.Add(new ListViewItem(new string[] { "", "", "" }));
            ListViewItem d = new ListViewItem(RdvText.StageDetect);
            d.SubItems.Add(RdvEngine.Fmt(run.DetectMs));
            d.SubItems.Add(run.Polls.ToString(CultureInfo.InvariantCulture) + "p");
            d.ForeColor = Sub;
            lvTime.Items.Add(d);
            ListViewItem rw = new ListViewItem(RdvText.StageRows);
            rw.SubItems.Add(run.Rows.ToString("N0", CultureInfo.InvariantCulture));
            rw.SubItems.Add("");
            rw.ForeColor = Sub;
            lvTime.Items.Add(rw);
            ListViewItem pb = new ListViewItem(RdvText.StageProbes);
            pb.SubItems.Add(run.Probes.ToString("N0", CultureInfo.InvariantCulture));
            pb.SubItems.Add("");
            pb.ForeColor = Sub;
            lvTime.Items.Add(pb);
            lvTime.EndUpdate();

            ListViewItem hi = new ListViewItem(run.Seq.ToString(CultureInfo.InvariantCulture));
            hi.SubItems.Add(run.When.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
            hi.SubItems.Add(run.Key);
            hi.SubItems.Add(RdvEngine.Fmt(run.TotalMs));
            hi.SubItems.Add(RdvEngine.Fmt(run.DetectMs));
            hi.SubItems.Add(Verdict(run));
            if (!run.OracleOk || run.Error.Length > 0) { hi.ForeColor = Bad; }
            lvHist.Items.Insert(0, hi);
            while (lvHist.Items.Count > 40) { lvHist.Items.RemoveAt(lvHist.Items.Count - 1); }

            lblError.Text = run.Error;
        }));
    }

    private static string Verdict(RdvRun r)
    {
        if (r.Error.Length > 0) { return RdvText.VerdictError; }
        if (!r.OracleOk) { return RdvText.OracleBad + " " + r.OracleNote; }
        return RdvText.OracleOk;
    }

    private void AddTime(string name, double ms, double total, bool bold)
    {
        ListViewItem it = new ListViewItem(name);
        it.SubItems.Add(RdvEngine.Fmt(ms));
        it.SubItems.Add(total > 0 ? (100.0 * ms / total).ToString("N1", CultureInfo.InvariantCulture) + "%" : "");
        if (bold) { it.Font = new Font(lvTime.Font, FontStyle.Bold); }
        lvTime.Items.Add(it);
    }
}
