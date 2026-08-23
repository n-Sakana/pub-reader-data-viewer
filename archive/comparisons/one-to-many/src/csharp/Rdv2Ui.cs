// ============================================================================
// Rdv2Ui.cs -- the WinForms screen and the candidate dialog.
//
// Where the clock stops, and why it stops there:
//
//   one candidate    the record is on screen -> stop
//   several          the dialog is built and its rows are filled -> stop.
//                    ShowDialog() comes AFTER the clock has stopped, so the
//                    time a person spends reading the list is not in the number.
//   after the pick   a separate figure, reported as "選択後の表示".
//
// PreparePick() therefore does all the work (create the form, fill the list,
// lay out the columns) and RunPick() only shows it.
// ============================================================================

using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

public sealed class Rdv2PickForm : Form
{
    private readonly ListView lv = new ListView();
    private readonly Button ok = new Button();
    private readonly Button cancel = new Button();
    public int Chosen = -1;

    public Rdv2PickForm(string key, Rdv2Cand[] cands)
    {
        Text = Rdv2Text.PickTitle;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(880, 420);
        MinimizeBox = false;
        MaximizeBox = false;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        BackColor = Color.White;
        try { Font = new Font("Yu Gothic UI", 9.5f); }
        catch (Exception) { }

        Label head = new Label();
        head.Text = Rdv2Text.LabelCurrent + "  " + key + "     "
                  + cands.Length.ToString(CultureInfo.InvariantCulture) + " " + Rdv2Text.PickHint;
        head.Font = new Font(Font.FontFamily, 11f, FontStyle.Bold);
        head.AutoSize = true;
        head.Location = new Point(14, 12);
        Controls.Add(head);

        lv.Location = new Point(14, 42);
        lv.Size = new Size(852, 306);
        lv.View = View.Details;
        lv.FullRowSelect = true;
        lv.GridLines = true;
        lv.MultiSelect = false;
        lv.HideSelection = false;
        lv.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        try { lv.Font = new Font("Consolas", 9.5f); }
        catch (Exception) { }
        lv.Columns.Add(Rdv2Text.PickNo, 40, HorizontalAlignment.Right);
        lv.Columns.Add(Rdv2Text.PickKey2, 92, HorizontalAlignment.Left);
        lv.Columns.Add(Rdv2Text.PickLine, 60, HorizontalAlignment.Right);
        lv.Columns.Add(Rdv2Text.PickSlip, 110, HorizontalAlignment.Left);
        lv.Columns.Add(Rdv2Text.PickDate, 92, HorizontalAlignment.Left);
        lv.Columns.Add(Rdv2Text.PickQty, 60, HorizontalAlignment.Right);
        lv.Columns.Add(Rdv2Text.PickStatus, 70, HorizontalAlignment.Left);
        lv.Columns.Add(Rdv2Text.PickItem, 110, HorizontalAlignment.Left);
        lv.Columns.Add(Rdv2Text.PickMaker, 110, HorizontalAlignment.Left);
        for (int i = 0; i < cands.Length; i++)
        {
            ListViewItem it = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
            string[] c = cands[i].Columns();
            for (int k = 0; k < c.Length; k++) { it.SubItems.Add(c[k]); }
            lv.Items.Add(it);
        }
        if (lv.Items.Count > 0) { lv.Items[0].Selected = true; }
        lv.DoubleClick += delegate(object s, EventArgs e) { Accept(); };
        Controls.Add(lv);

        ok.Text = Rdv2Text.BtnPick;
        ok.Size = new Size(160, 28);
        ok.Location = new Point(560, 360);
        ok.FlatStyle = FlatStyle.System;
        ok.Click += delegate(object s, EventArgs e) { Accept(); };
        Controls.Add(ok);

        cancel.Text = Rdv2Text.BtnCancel;
        cancel.Size = new Size(120, 28);
        cancel.Location = new Point(730, 360);
        cancel.FlatStyle = FlatStyle.System;
        cancel.Click += delegate(object s, EventArgs e) { Chosen = -1; DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;
    }

    private void Accept()
    {
        if (lv.SelectedIndices.Count == 0) { return; }
        Chosen = lv.SelectedIndices[0];
        DialogResult = DialogResult.OK;
        Close();
    }

    // used only by -autopick, so an unattended benchmark can answer the dialog.
    // It takes the same path a click does, after the clock has already stopped.
    public void PickRow(int which)
    {
        if (lv.Items.Count == 0) { return; }
        if (which < 0 || which >= lv.Items.Count) { which = 0; }
        lv.Items[which].Selected = true;
        Accept();
    }
}

public sealed class Rdv2Form : Form
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
    private Rdv2PickForm pick;

    public Action<string> OnManual;
    public Action OnRebind;

    private static readonly Color Ink = Color.FromArgb(28, 32, 38);
    private static readonly Color Sub = Color.FromArgb(96, 104, 116);
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

    public Rdv2Form(string method, string dataDir, string dataNote)
    {
        Text = Rdv2Text.AppTitle + " - " + method;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1180, 820);
        MinimumSize = new Size(940, 660);
        BackColor = Color.White;
        Font = UiFont(9.5f, FontStyle.Regular);

        TableLayoutPanel root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 1;
        root.RowCount = 5;
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 104));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 72));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 268));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 40));
        root.Padding = new Padding(14, 12, 14, 10);
        Controls.Add(root);

        Panel head = new Panel();
        head.Dock = DockStyle.Fill;
        lblMethod.Text = Rdv2Text.AppTitle + "   /   " + method;
        lblMethod.Font = UiFont(14f, FontStyle.Bold);
        lblMethod.ForeColor = Ink;
        lblMethod.AutoSize = true;
        lblMethod.Location = new Point(0, 0);
        head.Controls.Add(lblMethod);

        lblState.Text = Rdv2Text.StateBoot;
        lblState.Font = UiFont(11f, FontStyle.Bold);
        lblState.TextAlign = ContentAlignment.MiddleRight;
        lblState.Dock = DockStyle.Right;
        lblState.Width = 340;
        lblState.ForeColor = Sub;
        head.Controls.Add(lblState);

        lblNotepad.Text = Rdv2Text.NotepadNone;
        lblNotepad.ForeColor = Sub;
        lblNotepad.AutoSize = true;
        lblNotepad.Location = new Point(2, 34);
        head.Controls.Add(lblNotepad);

        lblData.Text = Rdv2Text.LabelData + " " + dataDir + "    " + dataNote;
        lblData.ForeColor = Sub;
        lblData.AutoSize = true;
        lblData.Location = new Point(2, 56);
        head.Controls.Add(lblData);

        btnRebind.Text = Rdv2Text.BtnRebind;
        btnRebind.Size = new Size(150, 26);
        btnRebind.Location = new Point(2, 76);
        btnRebind.FlatStyle = FlatStyle.System;
        btnRebind.Click += delegate(object s, EventArgs e) { if (OnRebind != null) { OnRebind(); } };
        head.Controls.Add(btnRebind);

        txtManual.Location = new Point(166, 78);
        txtManual.Width = 110;
        txtManual.Font = MonoFont(10f, FontStyle.Regular);
        txtManual.MaxLength = Rdv2Spec.KeyLen;
        head.Controls.Add(txtManual);

        btnManual.Text = Rdv2Text.BtnManual;
        btnManual.Size = new Size(120, 26);
        btnManual.Location = new Point(282, 76);
        btnManual.FlatStyle = FlatStyle.System;
        btnManual.Click += delegate(object s, EventArgs e)
        {
            string k = txtManual.Text.Trim();
            if (OnManual != null && Rdv2Watch.IsKey(k)) { OnManual(k); }
        };
        head.Controls.Add(btnManual);

        lblRaw.Text = "";
        lblRaw.ForeColor = Sub;
        lblRaw.Font = MonoFont(9f, FontStyle.Regular);
        lblRaw.AutoSize = true;
        lblRaw.Location = new Point(412, 82);
        head.Controls.Add(lblRaw);
        root.Controls.Add(head, 0, 0);

        Panel keyp = new Panel();
        keyp.Dock = DockStyle.Fill;
        keyp.BackColor = Paper;
        Label cap = new Label();
        cap.Text = Rdv2Text.LabelCurrent;
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

        SetupList(lvRec);
        lvRec.Columns.Add("#", 34, HorizontalAlignment.Right);
        lvRec.Columns.Add(Rdv2Text.ColFieldA, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv2Text.ColValueA, 190, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv2Text.ColFieldB, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv2Text.ColValueB, 190, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv2Text.ColFieldC, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv2Text.ColValueC, 190, HorizontalAlignment.Left);
        for (int i = 0; i < Rdv2Spec.Fields; i++)
        {
            ListViewItem it = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
            for (int k = 0; k < 6; k++) { it.SubItems.Add(""); }
            lvRec.Items.Add(it);
        }
        root.Controls.Add(Boxed(Rdv2Text.BoxRecord, lvRec), 0, 2);

        TableLayoutPanel bottom = new TableLayoutPanel();
        bottom.Dock = DockStyle.Fill;
        bottom.ColumnCount = 2;
        bottom.RowCount = 1;
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 440));
        bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        SetupList(lvTime);
        lvTime.Columns.Add(Rdv2Text.ColStage, 216, HorizontalAlignment.Left);
        lvTime.Columns.Add(Rdv2Text.ColMs, 92, HorizontalAlignment.Right);
        lvTime.Columns.Add(Rdv2Text.ColShare, 76, HorizontalAlignment.Right);
        bottom.Controls.Add(Boxed(Rdv2Text.BoxStages, lvTime), 0, 0);

        SetupList(lvHist);
        lvHist.Columns.Add("#", 40, HorizontalAlignment.Right);
        lvHist.Columns.Add(Rdv2Text.ColTime, 84, HorizontalAlignment.Left);
        lvHist.Columns.Add(Rdv2Text.ColKey, 88, HorizontalAlignment.Left);
        lvHist.Columns.Add(Rdv2Text.ColCand, 60, HorizontalAlignment.Right);
        lvHist.Columns.Add(Rdv2Text.ColTotal, 84, HorizontalAlignment.Right);
        lvHist.Columns.Add(Rdv2Text.ColPick, 78, HorizontalAlignment.Right);
        lvHist.Columns.Add(Rdv2Text.ColDetect, 78, HorizontalAlignment.Right);
        lvHist.Columns.Add(Rdv2Text.ColCheck, 120, HorizontalAlignment.Left);
        bottom.Controls.Add(Boxed(Rdv2Text.BoxHistory, lvHist), 1, 0);
        root.Controls.Add(bottom, 0, 3);

        lblError.Dock = DockStyle.Fill;
        lblError.ForeColor = Bad;
        lblError.Font = UiFont(9.5f, FontStyle.Bold);
        lblError.TextAlign = ContentAlignment.MiddleLeft;
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

    public void SetState(string state, string detail)
    {
        if (IsDisposed) { return; }
        try
        {
            Invoke(new Action(delegate
            {
                if (state == "WATCHING")
                {
                    lblState.Text = Rdv2Text.StateWatching;
                    lblState.ForeColor = Good;
                    if (detail.Length > 0) { lblNotepad.Text = Rdv2Text.LabelNotepad + " " + detail; }
                }
                else if (state == "WAITING")
                {
                    lblState.Text = Rdv2Text.StateWaiting;
                    lblState.ForeColor = Warn;
                    lblNotepad.Text = Rdv2Text.NotepadNone;
                }
                else if (state == "BUSY") { lblState.Text = Rdv2Text.StateBusy; lblState.ForeColor = Warn; }
                else if (state == "PICK") { lblState.Text = Rdv2Text.StatePick; lblState.ForeColor = Warn; }
                else { lblState.Text = state; lblState.ForeColor = Sub; }
            }));
        }
        catch (Exception) { }
    }

    public void SetRaw(string s)
    {
        if (IsDisposed) { return; }
        try { BeginInvoke(new Action(delegate { lblRaw.Text = Rdv2Text.LabelField + " [" + s + "]"; })); }
        catch (Exception) { }
    }

    // the record for one candidate: this is the end of the measured region when
    // exactly one candidate exists, and the "選択後の表示" figure when the
    // operator picked from several
    public void ShowRecord(Rdv2Run run)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate
        {
            lblKey.Text = run.Key;
            Rdv2Hit h = run.Hit;
            if (run.Error.Length > 0) { lblVerdict.Text = Rdv2Text.VerdictErr; lblVerdict.ForeColor = Bad; }
            else if (h.Count == 0) { lblVerdict.Text = Rdv2Text.VerdictNone; lblVerdict.ForeColor = Bad; }
            else if (h.Chosen >= 0)
            {
                string tail = (h.Count > 1)
                    ? ("   " + Rdv2Text.VerdictMany.Replace("{n}", h.Count.ToString(CultureInfo.InvariantCulture)))
                    : "";
                lblVerdict.Text = Rdv2Text.VerdictOne + "   key2 = " + h.Cands[h.Chosen].Key2 + tail;
                lblVerdict.ForeColor = Good;
            }
            lvRec.BeginUpdate();
            for (int i = 0; i < Rdv2Spec.Fields; i++)
            {
                ListViewItem it = lvRec.Items[i];
                it.SubItems[1].Text = Pick(h.NameA, i);
                it.SubItems[2].Text = Pick(h.ValA, i);
                it.SubItems[3].Text = Pick(h.NameB, i);
                it.SubItems[4].Text = Pick(h.ValB, i);
                it.SubItems[5].Text = Pick(h.NameC, i);
                it.SubItems[6].Text = Pick(h.ValC, i);
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

    // build the dialog and fill it. Still inside the measured region: this is
    // the "選択肢作成" the specification asks for.
    public void PreparePick(Rdv2Run run)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate
        {
            lblKey.Text = run.Key;
            lblVerdict.Text = Rdv2Text.VerdictMany.Replace("{n}", run.CandCount.ToString(CultureInfo.InvariantCulture));
            lblVerdict.ForeColor = Warn;
            pick = new Rdv2PickForm(run.Key, run.Hit.Cands);
            pick.CreateControl();
            Refresh();
        }));
    }

    // outside the measured region: however long the operator takes is theirs
    public int RunPick()
    {
        if (IsDisposed || pick == null) { return -1; }
        object r = Invoke(new Func<int>(delegate
        {
            using (Rdv2PickForm f = pick)
            {
                pick = null;
                f.ShowDialog(this);
                return f.Chosen;
            }
        }));
        return (int)r;
    }

    public void ShowStats(Rdv2Run run)
    {
        if (IsDisposed) { return; }
        Invoke(new Action(delegate
        {
            lvTime.BeginUpdate();
            lvTime.Items.Clear();
            double total = run.TotalMs;
            for (int i = 0; i < Rdv2Spec.StageCount; i++) { AddTime(Rdv2Text.StageName[i], run.Stage[i], total, false); }
            AddTime(Rdv2Text.StageOther, total - run.StageSum(), total, false);
            AddTime(Rdv2Text.StageTotal, total, total, true);
            lvTime.Items.Add(new ListViewItem(new string[] { "", "", "" }));
            Info(Rdv2Text.StagePick, run.PickMs >= 0 ? Rdv2Clock.Fmt(run.PickMs) : "-");
            Info(Rdv2Text.StageDetect, Rdv2Clock.Fmt(run.DetectMs));
            Info(Rdv2Text.StageCand, run.CandCount.ToString("N0", CultureInfo.InvariantCulture));
            Info(Rdv2Text.StageRows, run.Rows.ToString("N0", CultureInfo.InvariantCulture));
            Info(Rdv2Text.StageProbes, run.Probes.ToString("N0", CultureInfo.InvariantCulture));
            lvTime.EndUpdate();

            ListViewItem hi = new ListViewItem(run.Seq.ToString(CultureInfo.InvariantCulture));
            hi.SubItems.Add(run.When.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
            hi.SubItems.Add(run.Key);
            hi.SubItems.Add(run.CandCount.ToString(CultureInfo.InvariantCulture));
            hi.SubItems.Add(Rdv2Clock.Fmt(run.TotalMs));
            hi.SubItems.Add(run.PickMs >= 0 ? Rdv2Clock.Fmt(run.PickMs) : "-");
            hi.SubItems.Add(Rdv2Clock.Fmt(run.DetectMs));
            hi.SubItems.Add(run.Error.Length > 0 ? Rdv2Text.VerdictErr
                : (run.OracleOk ? Rdv2Text.OracleOk : Rdv2Text.OracleBad + " " + run.OracleNote));
            if (!run.OracleOk || run.Error.Length > 0) { hi.ForeColor = Bad; }
            // one line per run: a repeat of the same seq replaces its line
            if (lvHist.Items.Count > 0 && lvHist.Items[0].Text == hi.Text) { lvHist.Items.RemoveAt(0); }
            lvHist.Items.Insert(0, hi);
            while (lvHist.Items.Count > 40) { lvHist.Items.RemoveAt(lvHist.Items.Count - 1); }
            lblError.Text = run.Error;
        }));
    }

    private void Info(string name, string value)
    {
        ListViewItem it = new ListViewItem(name);
        it.SubItems.Add(value);
        it.SubItems.Add("");
        it.ForeColor = Sub;
        lvTime.Items.Add(it);
    }

    private void AddTime(string name, double ms, double total, bool bold)
    {
        ListViewItem it = new ListViewItem(name);
        it.SubItems.Add(Rdv2Clock.Fmt(ms));
        it.SubItems.Add(total > 0 ? (100.0 * ms / total).ToString("N1", CultureInfo.InvariantCulture) + "%" : "");
        if (bold) { it.Font = new Font(lvTime.Font, FontStyle.Bold); }
        lvTime.Items.Add(it);
    }
}
