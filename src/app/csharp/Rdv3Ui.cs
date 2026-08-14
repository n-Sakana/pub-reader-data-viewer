// ============================================================================
// Rdv3Ui.cs -- the WinForms screen of the practical build.
//
// Three regions, the same three the VBA sheet build draws:
//
//   status     state, notepad binding, ledger size, the TWO performance
//              figures the spec allows on screen (merge time, search time),
//              and the error line. Nothing else is displayed as a number;
//              every finer figure goes to the text log.
//   result     either the 30-field record of one integrated row, or the
//              candidate list (same eight identification columns as the
//              measured builds). Clicking a candidate row shows its record.
//   operations search box + 検索 / 内容クリア / 処理済み buttons.
//
// While the startup update check runs, an overlay with an animated label and a
// marquee bar sits over the result region; the UI thread is never the one
// doing the checking, so the animation actually moves.
//
// Every public method here marshals itself onto the UI thread: the worker
// calls them directly and never touches a control.
// ============================================================================

using System;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

public sealed class Rdv3Form : Form
{
    private readonly Label lblTitle = new Label();
    private readonly Label lblState = new Label();
    private readonly Label lblNotepad = new Label();
    private readonly Label lblLedger = new Label();
    private readonly Label lblMerge = new Label();
    private readonly Label lblSearch = new Label();
    private readonly Label lblError = new Label();
    private readonly Label lblKey = new Label();
    private readonly Label lblVerdict = new Label();
    private readonly Label lblProcessed = new Label();

    private readonly ListView lvRec = new ListView();
    private readonly ListView lvCand = new ListView();
    private readonly GroupBox boxRec;
    private readonly GroupBox boxCand;

    private readonly TextBox txtKey = new TextBox();
    private readonly Button btnSearch = new Button();
    private readonly Button btnClear = new Button();
    private readonly Button btnProcessed = new Button();
    private readonly Button btnRebind = new Button();
    private readonly ToolTip tips = new ToolTip();

    private readonly Panel overlay = new Panel();
    private readonly Label lblChecking = new Label();
    private readonly ProgressBar barChecking = new ProgressBar();
    // qualified: the packer hoists every using into one file, where a bare
    // Timer is ambiguous with System.Threading.Timer
    private readonly System.Windows.Forms.Timer animTimer = new System.Windows.Forms.Timer();
    private int animStep;
    private string overlayBase = "";

    public Action<string> OnSearch;      // manual search (the key from the box)
    public Action OnClear;
    public Action OnProcessed;
    public Action OnRebind;
    public Action<int> OnPick;           // candidate row clicked

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

    public Rdv3Form()
    {
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.Method;
        StartPosition = FormStartPosition.CenterScreen;
        ClientSize = new Size(1120, 760);
        MinimumSize = new Size(940, 620);
        BackColor = Color.White;
        Font = UiFont(9.5f, FontStyle.Regular);

        TableLayoutPanel root = new TableLayoutPanel();
        root.Dock = DockStyle.Fill;
        root.ColumnCount = 1;
        root.RowCount = 4;
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 128));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 64));
        root.Padding = new Padding(14, 10, 14, 8);
        Controls.Add(root);

        // ---- status region -------------------------------------------------
        Panel head = new Panel();
        head.Dock = DockStyle.Fill;

        lblTitle.Text = Rdv3Text.AppTitle + "   /   " + Rdv3Text.Method;
        lblTitle.Font = UiFont(14f, FontStyle.Bold);
        lblTitle.ForeColor = Ink;
        lblTitle.AutoSize = true;
        lblTitle.Location = new Point(0, 0);
        head.Controls.Add(lblTitle);

        lblState.Text = Rdv3Text.StateBoot;
        lblState.Font = UiFont(12f, FontStyle.Bold);
        lblState.TextAlign = ContentAlignment.MiddleRight;
        lblState.Dock = DockStyle.Right;
        lblState.Width = 380;
        lblState.ForeColor = Sub;
        head.Controls.Add(lblState);

        lblNotepad.Text = Rdv3Text.LabelNotepad + ": " + Rdv3Text.NotepadNone;
        lblNotepad.ForeColor = Sub;
        lblNotepad.AutoSize = true;
        lblNotepad.Location = new Point(2, 32);
        head.Controls.Add(lblNotepad);

        lblLedger.Text = Rdv3Text.LabelLedger + ": " + Rdv3Text.NotYet;
        lblLedger.ForeColor = Sub;
        lblLedger.AutoSize = true;
        lblLedger.Location = new Point(2, 54);
        head.Controls.Add(lblLedger);

        Label capM = new Label();
        capM.Text = Rdv3Text.LabelMergeMs;
        capM.ForeColor = Sub;
        capM.AutoSize = true;
        capM.Location = new Point(2, 82);
        head.Controls.Add(capM);

        lblMerge.Text = Rdv3Text.NotYet;
        lblMerge.Font = MonoFont(15f, FontStyle.Bold);
        lblMerge.ForeColor = Ink;
        lblMerge.AutoSize = true;
        lblMerge.Location = new Point(96, 76);
        head.Controls.Add(lblMerge);

        Label capS = new Label();
        capS.Text = Rdv3Text.LabelSearchMs;
        capS.ForeColor = Sub;
        capS.AutoSize = true;
        capS.Location = new Point(320, 82);
        head.Controls.Add(capS);

        lblSearch.Text = Rdv3Text.NotYet;
        lblSearch.Font = MonoFont(15f, FontStyle.Bold);
        lblSearch.ForeColor = Ink;
        lblSearch.AutoSize = true;
        lblSearch.Location = new Point(414, 76);
        head.Controls.Add(lblSearch);

        lblError.Text = "";
        lblError.ForeColor = Bad;
        lblError.Font = UiFont(9.5f, FontStyle.Bold);
        lblError.AutoSize = true;
        lblError.Location = new Point(2, 106);
        head.Controls.Add(lblError);
        root.Controls.Add(head, 0, 0);

        // ---- current key ---------------------------------------------------
        Panel keyp = new Panel();
        keyp.Dock = DockStyle.Fill;
        keyp.BackColor = Paper;

        Label cap = new Label();
        cap.Text = Rdv3Text.LabelKey;
        cap.ForeColor = Sub;
        cap.AutoSize = true;
        cap.Location = new Point(14, 6);
        keyp.Controls.Add(cap);

        lblKey.Text = "--------";
        lblKey.Font = MonoFont(22f, FontStyle.Bold);
        lblKey.ForeColor = Ink;
        lblKey.AutoSize = true;
        lblKey.Location = new Point(12, 22);
        keyp.Controls.Add(lblKey);

        lblVerdict.Text = "";
        lblVerdict.Font = UiFont(11.5f, FontStyle.Bold);
        lblVerdict.AutoSize = true;
        lblVerdict.Location = new Point(220, 28);
        keyp.Controls.Add(lblVerdict);

        lblProcessed.Text = "";
        lblProcessed.Font = UiFont(10f, FontStyle.Bold);
        lblProcessed.AutoSize = true;
        lblProcessed.TextAlign = ContentAlignment.MiddleRight;
        lblProcessed.Location = new Point(820, 28);
        lblProcessed.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        keyp.Controls.Add(lblProcessed);
        root.Controls.Add(keyp, 0, 1);

        // ---- result region: record / candidates ---------------------------
        Panel mid = new Panel();
        mid.Dock = DockStyle.Fill;

        SetupList(lvRec);
        lvRec.Columns.Add("#", 34, HorizontalAlignment.Right);
        lvRec.Columns.Add(Rdv3Text.ColFieldA, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv3Text.ColValueA, 186, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv3Text.ColFieldB, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv3Text.ColValueB, 186, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv3Text.ColFieldC, 96, HorizontalAlignment.Left);
        lvRec.Columns.Add(Rdv3Text.ColValueC, 186, HorizontalAlignment.Left);
        for (int i = 0; i < Rdv3Spec.Fields; i++)
        {
            ListViewItem it = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
            for (int k = 0; k < 6; k++) { it.SubItems.Add(""); }
            lvRec.Items.Add(it);
        }
        boxRec = Boxed(Rdv3Text.BoxRecord, lvRec);
        boxRec.Dock = DockStyle.Fill;
        mid.Controls.Add(boxRec);

        SetupList(lvCand);
        lvCand.Columns.Add(Rdv3Text.PickNo, 40, HorizontalAlignment.Right);
        lvCand.Columns.Add(Rdv3Text.PickKey2, 92, HorizontalAlignment.Left);
        lvCand.Columns.Add(Rdv3Text.PickLine, 60, HorizontalAlignment.Right);
        lvCand.Columns.Add(Rdv3Text.PickSlip, 110, HorizontalAlignment.Left);
        lvCand.Columns.Add(Rdv3Text.PickDate, 92, HorizontalAlignment.Left);
        lvCand.Columns.Add(Rdv3Text.PickQty, 60, HorizontalAlignment.Right);
        lvCand.Columns.Add(Rdv3Text.PickStatus, 70, HorizontalAlignment.Left);
        lvCand.Columns.Add(Rdv3Text.PickItem, 110, HorizontalAlignment.Left);
        lvCand.Columns.Add(Rdv3Text.PickMaker, 110, HorizontalAlignment.Left);
        lvCand.MouseClick += CandClicked;
        boxCand = Boxed(Rdv3Text.BoxCand, lvCand);
        boxCand.Dock = DockStyle.Fill;
        boxCand.Visible = false;
        mid.Controls.Add(boxCand);

        // ---- checking overlay ---------------------------------------------
        overlay.Dock = DockStyle.Fill;
        overlay.BackColor = Paper;
        overlay.Visible = false;
        lblChecking.Text = Rdv3Text.StateChecking;
        lblChecking.Font = UiFont(16f, FontStyle.Bold);
        lblChecking.ForeColor = Ink;
        lblChecking.AutoSize = true;
        lblChecking.Location = new Point(40, 60);
        overlay.Controls.Add(lblChecking);
        barChecking.Style = ProgressBarStyle.Marquee;
        barChecking.MarqueeAnimationSpeed = 30;
        barChecking.Location = new Point(44, 104);
        barChecking.Size = new Size(360, 18);
        overlay.Controls.Add(barChecking);
        mid.Controls.Add(overlay);
        overlay.BringToFront();
        root.Controls.Add(mid, 0, 2);

        animTimer.Interval = 250;
        animTimer.Tick += delegate(object s, EventArgs e)
        {
            animStep = (animStep + 1) % 4;
            lblChecking.Text = overlayBase + new string('.', animStep);
        };

        // ---- operations region --------------------------------------------
        Panel ops = new Panel();
        ops.Dock = DockStyle.Fill;

        Label capK = new Label();
        capK.Text = Rdv3Text.LabelSearchBox;
        capK.ForeColor = Sub;
        capK.AutoSize = true;
        capK.Location = new Point(2, 4);
        ops.Controls.Add(capK);

        txtKey.Location = new Point(4, 24);
        txtKey.Width = 130;
        txtKey.Font = MonoFont(12f, FontStyle.Regular);
        txtKey.MaxLength = Rdv3Spec.KeyLen;
        txtKey.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                FireSearch();
            }
        };
        ops.Controls.Add(txtKey);

        btnSearch.Text = Rdv3Text.BtnSearch;
        btnSearch.Size = new Size(110, 30);
        btnSearch.Location = new Point(150, 21);
        btnSearch.FlatStyle = FlatStyle.System;
        btnSearch.Click += delegate(object s, EventArgs e) { FireSearch(); };
        tips.SetToolTip(btnSearch, Rdv3Text.TipSearch);
        ops.Controls.Add(btnSearch);

        btnClear.Text = Rdv3Text.BtnClear;
        btnClear.Size = new Size(110, 30);
        btnClear.Location = new Point(270, 21);
        btnClear.FlatStyle = FlatStyle.System;
        btnClear.Click += delegate(object s, EventArgs e) { if (OnClear != null) { OnClear(); } };
        tips.SetToolTip(btnClear, Rdv3Text.TipClear);
        ops.Controls.Add(btnClear);

        btnProcessed.Text = Rdv3Text.BtnProcessed;
        btnProcessed.Size = new Size(110, 30);
        btnProcessed.Location = new Point(390, 21);
        btnProcessed.FlatStyle = FlatStyle.System;
        btnProcessed.Click += delegate(object s, EventArgs e) { if (OnProcessed != null) { OnProcessed(); } };
        tips.SetToolTip(btnProcessed, Rdv3Text.TipProcessed);
        ops.Controls.Add(btnProcessed);

        btnRebind.Text = Rdv3Text.BtnRebind;
        btnRebind.Size = new Size(130, 30);
        btnRebind.Location = new Point(540, 21);
        btnRebind.FlatStyle = FlatStyle.System;
        btnRebind.Click += delegate(object s, EventArgs e) { if (OnRebind != null) { OnRebind(); } };
        ops.Controls.Add(btnRebind);

        root.Controls.Add(ops, 0, 3);
    }

    private void FireSearch()
    {
        if (OnSearch != null) { OnSearch(txtKey.Text.Trim()); }
    }

    private void CandClicked(object sender, MouseEventArgs e)
    {
        ListViewItem it = lvCand.GetItemAt(e.X, e.Y);
        if (it != null && OnPick != null) { OnPick(it.Index); }
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
        g.Padding = new Padding(8, 6, 8, 8);
        g.ForeColor = Sub;
        inner.ForeColor = Ink;
        g.Controls.Add(inner);
        return g;
    }

    private void OnUi(Action a)
    {
        if (IsDisposed) { return; }
        try
        {
            if (InvokeRequired) { Invoke(a); } else { a(); }
        }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { }
    }

    public void RunOnUi(Action a)
    {
        OnUi(a);
    }

    // ---- status ------------------------------------------------------------
    public void SetState(string text, int tone)
    {
        OnUi(delegate
        {
            lblState.Text = text;
            lblState.ForeColor = (tone == 1) ? Good : (tone == 2) ? Warn : (tone == 3) ? Bad : Sub;
        });
    }

    public void SetNotepad(string detail)
    {
        OnUi(delegate { lblNotepad.Text = Rdv3Text.LabelNotepad + ": " + detail; });
    }

    public void SetLedgerInfo(string text)
    {
        OnUi(delegate { lblLedger.Text = Rdv3Text.LabelLedger + ": " + text; });
    }

    public void SetMergeMs(double ms)
    {
        OnUi(delegate { lblMerge.Text = (ms >= 0) ? Rdv3Clock.Fmt(ms) + Rdv3Text.MsUnit : Rdv3Text.NotYet; });
    }

    public void SetSearchMs(double ms)
    {
        OnUi(delegate { lblSearch.Text = (ms >= 0) ? Rdv3Clock.Fmt(ms) + Rdv3Text.MsUnit : Rdv3Text.NotYet; });
    }

    public void SetError(string text)
    {
        OnUi(delegate { lblError.Text = (text == null) ? "" : text; });
    }

    public void ShowOverlay(string baseText)
    {
        OnUi(delegate
        {
            overlayBase = baseText;
            lblChecking.Text = baseText;
            animStep = 0;
            overlay.Visible = true;
            overlay.BringToFront();
            animTimer.Start();
        });
    }

    public void HideOverlay()
    {
        OnUi(delegate
        {
            animTimer.Stop();
            overlay.Visible = false;
        });
    }

    public void EnableOps(bool on)
    {
        OnUi(delegate
        {
            txtKey.Enabled = on;
            btnSearch.Enabled = on;
            btnClear.Enabled = on;
            btnProcessed.Enabled = on;
        });
    }

    // ---- result ------------------------------------------------------------
    public void ShowRecord(string key, string verdict, int tone, string[] valA, string[] valB, string[] valC, string processedText)
    {
        OnUi(delegate
        {
            lblKey.Text = (key.Length > 0) ? key : "--------";
            lblVerdict.Text = verdict;
            lblVerdict.ForeColor = (tone == 1) ? Good : (tone == 2) ? Warn : (tone == 3) ? Bad : Sub;
            lblProcessed.Text = processedText;
            lblProcessed.ForeColor = (processedText.IndexOf(Rdv3Ledger.ProcessedTrue, StringComparison.Ordinal) >= 0) ? Good : Sub;
            lvRec.BeginUpdate();
            for (int i = 0; i < Rdv3Spec.Fields; i++)
            {
                ListViewItem it = lvRec.Items[i];
                it.SubItems[1].Text = Rdv3Ledger.NameA[i];
                it.SubItems[2].Text = Pick(valA, i);
                it.SubItems[3].Text = Rdv3Ledger.NameB[i];
                it.SubItems[4].Text = Pick(valB, i);
                it.SubItems[5].Text = Rdv3Ledger.NameC[i];
                it.SubItems[6].Text = Pick(valC, i);
            }
            lvRec.EndUpdate();
            boxCand.Visible = false;
            boxRec.Visible = true;
            boxRec.BringToFront();
        });
    }

    private static string Pick(string[] a, int i)
    {
        if (a == null || i >= a.Length) { return ""; }
        return a[i];
    }

    public void ShowCandidates(string key, string verdict, string[][] rows)
    {
        OnUi(delegate
        {
            lblKey.Text = key;
            lblVerdict.Text = verdict;
            lblVerdict.ForeColor = Warn;
            lblProcessed.Text = "";
            lvCand.BeginUpdate();
            lvCand.Items.Clear();
            for (int i = 0; i < rows.Length; i++)
            {
                ListViewItem it = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
                for (int k = 0; k < rows[i].Length; k++) { it.SubItems.Add(rows[i][k]); }
                lvCand.Items.Add(it);
            }
            lvCand.EndUpdate();
            boxRec.Visible = false;
            boxCand.Visible = true;
            boxCand.BringToFront();
        });
    }

    public void ClearResult()
    {
        OnUi(delegate
        {
            txtKey.Text = "";
            lblKey.Text = "--------";
            lblVerdict.Text = "";
            lblProcessed.Text = "";
            lvRec.BeginUpdate();
            for (int i = 0; i < Rdv3Spec.Fields; i++)
            {
                ListViewItem it = lvRec.Items[i];
                for (int k = 1; k <= 6; k++) { it.SubItems[k].Text = ""; }
            }
            lvRec.EndUpdate();
            lvCand.Items.Clear();
            boxCand.Visible = false;
            boxRec.Visible = true;
            boxRec.BringToFront();
        });
    }

    public void ShowProcessedState(string processedText)
    {
        OnUi(delegate
        {
            lblProcessed.Text = processedText;
            lblProcessed.ForeColor = (processedText.IndexOf(Rdv3Ledger.ProcessedTrue, StringComparison.Ordinal) >= 0) ? Good : Sub;
        });
    }

    public bool Ask(string title, string body)
    {
        return MessageBox.Show(this, body, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
    }
}
