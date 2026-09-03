// ============================================================================
// Rdv3Ui.cs -- the definition-driven WinForms screen.
//
// The screen is made only from standard WinForms controls.  In particular,
// borders, buttons, focus cues and the old status bar are drawn by Windows;
// this file contains no custom painting.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

public sealed class Rdv3CandRow
{
    public string Line;
    public string Stored;
}

public sealed class Rdv3Form : Form
{
    private sealed class BoundControl
    {
        public Control Control;
        public Rdv3Bind Bind;
        public bool EmptyAsDash;
    }

    public readonly Rdv3Screen Screen;
    public readonly Rdv3View View = new Rdv3View();

    public Action<string> OnSearch;
    public Action OnClear;
    public Action OnWorkState;
    public Action OnRefreshLedger;
    public Action OnTableExport;
    public Action<string> OnUpdateRecords;
    public Action<string> OnDeleteRecords;
    public Action OnSendChanges;
    public Action OnSettings;

    private Rdv3Fields fields = Rdv3Fields.Empty;
    private readonly List<BoundControl> bound = new List<BoundControl>();
    private readonly List<Control> operationControls = new List<Control>();
    private readonly Dictionary<Control, Rdv3ButtonDef> buttonDefs = new Dictionary<Control, Rdv3ButtonDef>();
    private readonly List<Rdv3SegmentDef> segmentDefs = new List<Rdv3SegmentDef>();
    private readonly List<ToolStripStatusLabel> statusPanels = new List<ToolStripStatusLabel>();
    private readonly Panel contentHost = new Panel();
    private readonly TableLayoutPanel content = new TableLayoutPanel();
    private readonly Panel commandBar = new Panel();
    private readonly FlowLayoutPanel commandFlow = new FlowLayoutPanel();
    private readonly Panel statusHost = new Panel();
    private readonly StatusStrip status = new StatusStrip();
    private readonly System.Windows.Forms.Timer clock = new System.Windows.Forms.Timer();
    private readonly ToolTip tips = new ToolTip();

    private TextBox txtKey;
    private CheckBox btnWork;
    private Label judgmentText;
    private Label judgmentSub;
    private Label pendingLabel;
    private Rdv3Section sendDef;
    private Font judgmentResultFont;
    private Font judgmentUnsearchedFont;
    private Rdv3Judgment judgmentDef;
    private List<Rdv3CandRow> cands = new List<Rdv3CandRow>();
    private int candTotal;
    private Rdv3Toast toast;
    private string sharedNotice = "";
    private long sharedNoticeAt;

    public Rdv3Form(Rdv3Screen screen)
    {
        Screen = screen;
        SuspendLayout();
        // new Font() does not throw for a family this machine does not have:
        // it quietly hands back Microsoft Sans Serif, which carries no Japanese
        // glyphs. Compare the name we asked for with the one we got.
        Font named = null;
        try { named = new Font(screen.FontFamily, 9.0f, FontStyle.Regular); }
        catch { named = null; }
        if (named != null && !string.Equals(named.Name, screen.FontFamily, StringComparison.OrdinalIgnoreCase))
        {
            named.Dispose();
            named = null;
        }
        Font = (named != null) ? named : SystemFonts.MessageBoxFont;
        AutoScaleDimensions = new SizeF(96.0f, 96.0f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size((int)Math.Round(screen.StartWidth), (int)Math.Round(screen.StartHeight));

        Text = Rdv3Text.AppTitle;
        FormBorderStyle = FormBorderStyle.Sizable;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = true;
        MinimizeBox = true;
        ShowIcon = true;
        KeyPreview = true;
        BackColor = SystemColors.Control;
        MinimumSize = new Size(680, 520);

        contentHost.Name = "contentHost";
        contentHost.Dock = DockStyle.Fill;
        contentHost.AutoScroll = true;
        contentHost.BackColor = SystemColors.Control;

        content.Name = "content";
        content.Dock = DockStyle.Top;
        content.AutoSize = true;
        content.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        content.ColumnCount = 1;
        content.RowCount = 0;
        content.Padding = Box(screen.Padding, 8);
        content.BackColor = SystemColors.Control;
        contentHost.Controls.Add(content);

        commandBar.Name = "commandBar";
        commandBar.Dock = DockStyle.Fill;
        commandBar.AutoSize = true;
        commandBar.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        commandBar.MinimumSize = new Size(0, 36);
        commandBar.Padding = new Padding(6, 4, 6, 3);
        commandBar.BackColor = SystemColors.Control;
        commandFlow.Name = "commandButtons";
        commandFlow.Dock = DockStyle.Top;
        commandFlow.AutoSize = true;
        commandFlow.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        commandFlow.FlowDirection = FlowDirection.RightToLeft;
        commandFlow.WrapContents = false;
        commandFlow.BackColor = SystemColors.Control;
        commandBar.Controls.Add(commandFlow);

        statusHost.Name = "statusHost";
        statusHost.Dock = DockStyle.Bottom;
        statusHost.Height = 20;
        statusHost.BackColor = SystemColors.Control;
        status.Name = "statusBar";
        status.Dock = DockStyle.Fill;
        status.AutoSize = false;
        status.SizingGrip = true;
        // the system renderer keeps the strip on the same footing as the rest
        // of the screen; the professional one paints a gradient of its own
        status.RenderMode = ToolStripRenderMode.System;
        status.GripStyle = ToolStripGripStyle.Hidden;
        status.Padding = new Padding(1, 0, 14, 0);

        statusHost.Controls.Add(status);
        Controls.Add(contentHost);
        Controls.Add(statusHost);

        BuildSections();
        clock.Interval = 1000;
        clock.Tick += delegate { RefreshValues(); };
        clock.Start();
        FormClosed += delegate { clock.Stop(); if (toast != null) { toast.Dismiss(); } };
        Shown += delegate { if (txtKey != null && txtKey.Enabled) { txtKey.Focus(); } };
        RefreshValues();
        ResumeLayout(false);
        PerformLayout();
    }

    private static Padding Box(double[] a, int fallback)
    {
        if (a == null || a.Length == 0) { return new Padding(fallback); }
        if (a.Length == 1) { return new Padding((int)Math.Round(a[0])); }
        if (a.Length == 2) { return new Padding((int)Math.Round(a[1]), (int)Math.Round(a[0]), (int)Math.Round(a[1]), (int)Math.Round(a[0])); }
        return new Padding((int)Math.Round(a[3]), (int)Math.Round(a[0]), (int)Math.Round(a[1]), (int)Math.Round(a[2]));
    }

    private void BuildSections()
    {
        Rdv3Section statusSection = null;
        Rdv3Section sendSection = null;
        for (int i = 0; i < Screen.Sections.Count; i++)
        {
            Rdv3Section s = Screen.Sections[i];
            if (s.Type == "titleBar")
            {
                if (s.Brand.Length > 0) { Text = s.Brand; }
                AddTitleButtons(s.Buttons);
                continue;
            }
            if (s.Type == "statusBar") { statusSection = s; continue; }
            if (s.Type == "sendBar") { sendSection = s; continue; }
            Control c = BuildSection(s, "section" + i.ToString(CultureInfo.InvariantCulture));
            if (c == null) { continue; }
            c.Margin = SectionMargin(s);
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.Controls.Add(c, 0, content.RowCount++);
        }
        if (statusSection != null) { BuildBottom(statusSection, sendSection); }
        content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
    }

    private Padding SectionMargin(Rdv3Section s)
    {
        if (s.Margin != null) { return Box(s.Margin, 0); }
        return new Padding(0, 0, 0, Math.Max(0, (int)Math.Round(Screen.Gap)));
    }

    private Control BuildSection(Rdv3Section s, string name)
    {
        if (s.Type == "keyPanel") { return BuildKeyPanel(s, name); }
        if (s.Type == "columns") { return BuildColumns(s, name); }
        if (s.Type == "fieldList") { return BuildFieldList(s, name); }
        if (s.Type == "textBox") { return BuildTextBox(s, name); }
        if (s.Type == "statusBand") { return BuildStatusBand(s, name); }
        return null;
    }

    private Control BuildKeyPanel(Rdv3Section s, string name)
    {
        GroupBox box = NewGroup(name, s.Title);
        // the figure below the caption is set in 15 pt bold: the panel has to
        // clear the caption, the label row and that line, or the number is cut
        box.Height = 80;
        box.AutoSize = false;

        TableLayoutPanel split = new TableLayoutPanel();
        split.Dock = DockStyle.Fill;
        split.ColumnCount = 2;
        split.RowCount = 1;
        split.Padding = new Padding(7, 5, 7, 6);
        split.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220f));
        split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        split.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        TableLayoutPanel figure = new TableLayoutPanel();
        figure.Dock = DockStyle.Fill;
        figure.RowCount = 2;
        figure.ColumnCount = 1;
        figure.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        figure.RowStyles.Add(new RowStyle(SizeType.Absolute, 18f));
        figure.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        Label fl = NewLabel(name + ".figureLabel", s.Label);
        fl.Dock = DockStyle.Fill;
        Label fv = NewReadOnlyLabel(name + ".figure");
        fv.Dock = DockStyle.Fill;
        fv.Font = new Font(Font.FontFamily, 15.0f, FontStyle.Bold);
        fv.TextAlign = ContentAlignment.MiddleCenter;
        figure.Controls.Add(fl, 0, 0);
        figure.Controls.Add(fv, 0, 1);
        Bind(fv, s.Value, true);

        FlowLayoutPanel input = new FlowLayoutPanel();
        input.Name = name + ".inputRow";
        input.Dock = DockStyle.Fill;
        input.FlowDirection = FlowDirection.LeftToRight;
        input.WrapContents = false;
        input.Padding = new Padding(6, 13, 0, 0);
        Label il = NewLabel(name + ".inputLabel", s.InputLabel);
        il.AutoSize = true;
        il.Margin = new Padding(0, 5, 5, 0);
        txtKey = new TextBox();
        txtKey.Name = "searchInput";
        txtKey.Width = (int)Math.Round(s.InputWidth);
        txtKey.MaxLength = s.MaxLength;
        txtKey.BackColor = SystemColors.Window;
        txtKey.BorderStyle = BorderStyle.Fixed3D;
        txtKey.Margin = new Padding(0, 2, 5, 0);
        txtKey.KeyDown += delegate(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; RaiseSearch(); }
        };
        input.Controls.Add(il);
        input.Controls.Add(txtKey);
        for (int i = 0; i < s.Buttons.Count; i++)
        {
            Control made = CreateButton(s.Buttons[i]);
            input.Controls.Add(made);
            Button primary = made as Button;
            if (primary != null && s.Buttons[i].Primary && AcceptButton == null) { AcceptButton = primary; }
        }

        split.Controls.Add(figure, 0, 0);
        split.Controls.Add(input, 1, 0);
        box.Controls.Add(split);
        return box;
    }

    private Control BuildColumns(Rdv3Section s, string name)
    {
        TableLayoutPanel row = new TableLayoutPanel();
        row.Name = name;
        row.Dock = DockStyle.Top;
        row.AutoSize = true;
        row.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        row.RowCount = 1;
        row.ColumnCount = s.Items.Count;
        for (int i = 0; i < s.Items.Count; i++)
        {
            float weight = (float)(s.Weights[i] * 100.0 / Sum(s.Weights));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, weight));
            Control child = BuildSection(s.Items[i], name + ".item" + i.ToString(CultureInfo.InvariantCulture));
            child.Dock = DockStyle.Fill;
            child.Margin = new Padding(i == 0 ? 0 : (int)Math.Round(s.Gap), 0, 0, 0);
            row.Controls.Add(child, i, 0);
        }
        return row;
    }

    private static double Sum(double[] a)
    {
        double n = 0;
        for (int i = 0; i < a.Length; i++) { n += a[i]; }
        return (n <= 0) ? 1 : n;
    }

    private static int DesignFontHeight(Font font)
    {
        int em = font.FontFamily.GetEmHeight(font.Style);
        int line = font.FontFamily.GetLineSpacing(font.Style);
        if (em <= 0 || line <= 0) { return 12; }
        return Math.Max(1, (int)Math.Ceiling(font.SizeInPoints * 96.0f * line / (72.0f * em)));
    }

    private Control BuildFieldList(Rdv3Section s, string name)
    {
        GroupBox box = NewGroup(name, s.Title);
        TableLayoutPanel rows = new TableLayoutPanel();
        rows.Name = name + ".rows";
        rows.Dock = DockStyle.Top;
        rows.Padding = new Padding(7, 5, 7, 7);
        rows.ColumnCount = 2;
        rows.RowCount = s.Rows.Count;
        rows.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (float)s.LabelWidth));
        rows.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        int rh = Math.Max(23, (int)Math.Round(s.RowHeight));
        rows.Height = rows.Padding.Vertical + rh * s.Rows.Count;
        for (int i = 0; i < s.Rows.Count; i++)
        {
            rows.RowStyles.Add(new RowStyle(SizeType.Absolute, rh));
            Label l = NewLabel(name + ".label" + i.ToString(CultureInfo.InvariantCulture), s.Rows[i].Label);
            l.Dock = DockStyle.Fill;
            l.TextAlign = ContentAlignment.MiddleLeft;
            Label v = NewReadOnlyLabel(name + ".value" + i.ToString(CultureInfo.InvariantCulture));
            v.Dock = DockStyle.Fill;
            v.TextAlign = ContentAlignment.MiddleRight;
            v.Margin = new Padding(2, 1, 0, 1);
            rows.Controls.Add(l, 0, i);
            rows.Controls.Add(v, 1, i);
            Bind(v, s.Rows[i].Value, true);
        }
        box.Height = 25 + rh * s.Rows.Count;
        box.AutoSize = false;
        box.Controls.Add(rows);
        return box;
    }

    private Control BuildTextBox(Rdv3Section s, string name)
    {
        GroupBox box = NewGroup(name, s.Title);
        TextBox v = NewReadOnlyTextBox(name + ".value");
        v.Multiline = true;
        v.ScrollBars = ScrollBars.Vertical;
        v.Dock = DockStyle.Fill;
        v.Margin = new Padding(7);
        box.Padding = new Padding(8, 4, 8, 7);
        box.Height = 28 + Math.Max(1, s.Lines) * (DesignFontHeight(Font) + 2);
        box.AutoSize = false;
        box.Controls.Add(v);
        Bind(v, s.Value, true);
        return box;
    }

    private Control BuildStatusBand(Rdv3Section s, string name)
    {
        Panel box = new Panel();
        box.Name = name;
        box.Dock = DockStyle.Top;
        box.Height = Math.Max(32, (int)Math.Round(s.Height));
        box.AutoSize = false;
        box.BorderStyle = BorderStyle.Fixed3D;
        box.BackColor = SystemColors.Control;
        Label title = NewLabel(name + ".label", s.Label);
        title.AutoSize = true;
        judgmentText = NewLabel(name + ".judgment", Rdv3Text.Unsearched);
        judgmentText.AutoSize = true;
        judgmentResultFont = new Font(Font.FontFamily, 15.0f, FontStyle.Bold);
        judgmentUnsearchedFont = new Font(Font.FontFamily, 12.0f, FontStyle.Bold);
        judgmentText.Font = judgmentUnsearchedFont;
        judgmentText.ForeColor = Color.FromArgb(96, 96, 96);
        judgmentSub = NewLabel(name + ".sub", "");
        judgmentSub.AutoSize = true;
        box.Controls.Add(title);
        box.Controls.Add(judgmentText);
        box.Controls.Add(judgmentSub);
        box.Layout += delegate
        {
            int gap = judgmentSub.Text.Length == 0 ? 0 : 8;
            int total = judgmentText.Width + gap + judgmentSub.Width;
            title.Location = new Point(6, Math.Max(0, (box.ClientSize.Height - title.Height) / 2));
            judgmentText.Location = new Point(Math.Max(0, (box.ClientSize.Width - total) / 2),
                Math.Max(0, (box.ClientSize.Height - judgmentText.Height) / 2));
            judgmentSub.Location = new Point(judgmentText.Right + gap,
                Math.Max(0, (box.ClientSize.Height - judgmentSub.Height) / 2));
        };
        judgmentText.TextChanged += delegate { box.PerformLayout(); };
        judgmentSub.TextChanged += delegate { box.PerformLayout(); };
        judgmentDef = Screen.JudgmentOf(s.Judgment);
        // the sub line is composed from all of s.Sub at once, in RefreshJudgment
        judgmentSub.Tag = s;
        return box;
    }

    private void BuildBottom(Rdv3Section s, Rdv3Section send)
    {
        for (int i = s.Buttons.Count - 1; i >= 0; i--) { commandFlow.Controls.Add(CreateButton(s.Buttons[i])); }
        commandBar.Margin = new Padding(0);
        content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        content.Controls.Add(commandBar, 0, content.RowCount++);
        if (send != null)
        {
            Control sendBar = BuildSendBar(send);
            sendBar.Margin = new Padding(0);
            content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            content.Controls.Add(sendBar, 0, content.RowCount++);
        }
        statusHost.Height = Math.Max(20, (int)Math.Round(s.Height));
        int n = s.Segments.Count;
        for (int i = 0; i < n; i++)
        {
            ToolStripStatusLabel p = new ToolStripStatusLabel();
            p.Name = "statusBar.segment" + i.ToString(CultureInfo.InvariantCulture);
            p.Spring = (i == 2);
            p.TextAlign = ContentAlignment.MiddleLeft;
            p.BorderSides = ToolStripStatusLabelBorderSides.All;
            p.BorderStyle = Border3DStyle.SunkenOuter;
            p.Margin = new Padding(0, 0, 2, 0);
            status.Items.Add(p);
            statusPanels.Add(p);
            if (i < s.Segments.Count) { segmentDefs.Add(s.Segments[i]); }
        }
    }

    private Control BuildSendBar(Rdv3Section s)
    {
        sendDef = s;
        TableLayoutPanel bar = new TableLayoutPanel();
        bar.Name = "sendBar";
        bar.Dock = DockStyle.Top;
        bar.AutoSize = false;
        bar.Height = Math.Max(28, (int)Math.Round(s.Height));
        bar.ColumnCount = 2;
        bar.RowCount = 2;
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.RowStyles.Add(new RowStyle(SizeType.Absolute, 2f));
        bar.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        Label separator = new Label();
        separator.Name = "sendBar.separator";
        separator.BorderStyle = BorderStyle.Fixed3D;
        separator.Dock = DockStyle.Fill;
        separator.Margin = new Padding(0);
        bar.Controls.Add(separator, 0, 0);
        bar.SetColumnSpan(separator, 2);

        pendingLabel = NewLabel("sendBar.pending", "");
        pendingLabel.Dock = DockStyle.Fill;
        pendingLabel.TextAlign = ContentAlignment.MiddleLeft;
        pendingLabel.Margin = new Padding(0, 5, 0, 0);
        bar.Controls.Add(pendingLabel, 0, 1);

        Control button = CreateButton(s.Buttons[0]);
        button.Name = "button.sendChanges";
        button.Margin = new Padding(5, 4, 0, 0);
        bar.Controls.Add(button, 1, 1);
        return bar;
    }

    private void AddTitleButtons(List<Rdv3ButtonDef> defs)
    {
        for (int i = defs.Count - 1; i >= 0; i--) { commandFlow.Controls.Add(CreateButton(defs[i])); }
    }

    private Control CreateButton(Rdv3ButtonDef d)
    {
        if (d.Action == "workState")
        {
            CheckBox cb = new CheckBox();
            cb.Name = "button.workState";
            cb.Appearance = Appearance.Button;
            cb.AutoCheck = false;
            cb.TextAlign = ContentAlignment.MiddleCenter;
            cb.FlatStyle = FlatStyle.Standard;
            cb.UseVisualStyleBackColor = false;
            cb.AutoSize = true;
            cb.MinimumSize = new Size(0, 25);
            cb.Margin = new Padding(0, 1, 5, 0);
            cb.Click += delegate { if (OnWorkState != null) { OnWorkState(); } };
            if (Screen.Work != null && Screen.Work.ButtonTip.Length > 0) { tips.SetToolTip(cb, Screen.Work.ButtonTip); }
            btnWork = cb;
            operationControls.Add(cb);
            buttonDefs[cb] = d;
            return cb;
        }

        Button b = new Button();
        b.Name = "button." + d.Action;
        b.Text = d.Text;
        b.FlatStyle = FlatStyle.Standard;
        b.UseVisualStyleBackColor = false;
        b.AutoSize = true;
        b.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        b.MinimumSize = new Size(0, 25);
        b.Margin = new Padding(0, 1, 5, 0);
        b.NotifyDefault(d.Primary);
        b.Click += delegate { RunButton(d); };
        if (d.Tip.Length > 0) { tips.SetToolTip(b, d.Tip); }
        operationControls.Add(b);
        buttonDefs[b] = d;
        return b;
    }

    private void RunButton(Rdv3ButtonDef d)
    {
        if (d.Action == "search") { RaiseSearch(); }
        else if (d.Action == "clear") { if (OnClear != null) { OnClear(); } }
        else if (d.Action == "refreshLedger") { if (OnRefreshLedger != null) { OnRefreshLedger(); } }
        else if (d.Action == "tableExport") { if (OnTableExport != null) { OnTableExport(); } }
        else if (d.Action == "updateRecords") { if (OnUpdateRecords != null) { OnUpdateRecords(d.Job); } }
        else if (d.Action == "deleteRecords") { if (OnDeleteRecords != null) { OnDeleteRecords(d.Job); } }
        else if (d.Action == "sendChanges") { if (OnSendChanges != null) { OnSendChanges(); } }
        else if (d.Action == "settings") { if (OnSettings != null) { OnSettings(); } }
    }

    private void RaiseSearch()
    {
        if (OnSearch != null) { OnSearch(KeyText); }
    }

    private static GroupBox NewGroup(string name, string title)
    {
        GroupBox b = new GroupBox();
        b.Name = name;
        b.Text = title;
        b.Dock = DockStyle.Top;
        b.BackColor = SystemColors.Control;
        return b;
    }

    private static Label NewLabel(string name, string text)
    {
        Label l = new Label();
        l.Name = name;
        l.Text = text;
        l.AutoEllipsis = true;
        l.UseCompatibleTextRendering = false;
        return l;
    }

    private static Label NewReadOnlyLabel(string name)
    {
        Label l = new Label();
        l.Name = name;
        l.TabStop = false;
        l.BackColor = SystemColors.Control;
        l.ForeColor = SystemColors.ControlText;
        l.BorderStyle = BorderStyle.Fixed3D;
        l.TextAlign = ContentAlignment.MiddleLeft;
        l.AutoEllipsis = true;
        l.UseCompatibleTextRendering = false;
        return l;
    }

    private static TextBox NewReadOnlyTextBox(string name)
    {
        TextBox t = new TextBox();
        t.Name = name;
        t.ReadOnly = true;
        t.TabStop = false;
        t.BackColor = SystemColors.Control;
        t.ForeColor = SystemColors.ControlText;
        t.BorderStyle = BorderStyle.Fixed3D;
        return t;
    }

    private void Bind(Control c, Rdv3Bind b, bool dash)
    {
        BoundControl x = new BoundControl();
        x.Control = c;
        x.Bind = b;
        x.EmptyAsDash = dash;
        bound.Add(x);
    }

    private void Ui(Action a)
    {
        if (IsDisposed) { return; }
        if (InvokeRequired) { Invoke(a); } else { a(); }
    }

    public void RunOnUi(Action a) { Ui(a); }

    public void PostOnUi(Action a)
    {
        if (IsDisposed) { return; }
        if (IsHandleCreated) { BeginInvoke(a); } else { a(); }
    }

    public void SetFields(Rdv3Fields f) { Ui(delegate { fields = (f == null) ? Rdv3Fields.Empty : f; RefreshValues(); }); }
    public Rdv3Fields Fields { get { return fields; } }
    public void SetState(string text) { Ui(delegate { View.AppState = (text == null) ? "" : text; RefreshValues(); }); }

    public void SetWatch(string label, string detail)
    {
        Ui(delegate
        {
            View.WatchLabel = ((label == null || label.Length == 0) ? Rdv3Text.LabelWatch : label)
                + ((detail == null || detail.Length == 0) ? "" : (" " + detail));
            View.WatchDetail = (detail == null) ? "" : detail;
            RefreshValues();
        });
    }

    public void SetLedger(string file, string rows, string saved)
    {
        Ui(delegate
        {
            View.LedgerFile = (file == null) ? "" : file;
            View.LedgerRows = (rows == null) ? "" : rows;
            View.LedgerSaved = (saved == null) ? "" : saved;
            RefreshValues();
        });
    }

    public void SetPendingCount(int count)
    {
        Ui(delegate { View.PendingCount = Math.Max(0, count); RefreshValues(); });
    }

    public void SetTimes(double mergeMs, double searchMs)
    {
        Ui(delegate
        {
            if (mergeMs >= 0) { View.MergeMs = Rdv3Clock.Fmt(mergeMs) + Rdv3Text.MsUnit; }
            if (searchMs >= 0) { View.SearchMs = Rdv3Clock.Fmt(searchMs) + Rdv3Text.MsUnit; }
            RefreshValues();
        });
    }

    public void SetIdentity(string pid, string logName)
    {
        Ui(delegate { View.Pid = pid; View.LogName = logName; RefreshValues(); });
    }

    public void EnableOps(bool on)
    {
        Ui(delegate
        {
            if (txtKey != null) { txtKey.Enabled = on; }
            for (int i = 0; i < operationControls.Count; i++)
            {
                Rdv3ButtonDef d = buttonDefs[operationControls[i]];
                if (d.Action != "settings") { operationControls[i].Enabled = on; }
            }
        });
    }

    public void EnableWorkState(bool on) { Ui(delegate { if (btnWork != null) { btnWork.Enabled = on; } }); }

    public void ShowCandidates(string key, List<Rdv3CandRow> rows, int totalHits)
    {
        Ui(delegate
        {
            View.SearchKey = (key == null) ? "" : key;
            cands = (rows == null) ? new List<Rdv3CandRow>() : rows;
            candTotal = totalHits;
            View.CandidateCount = totalHits;
            View.SelectedIndex = -1;
            View.Record = null;
            View.StoredState = "";
            View.Saving = false;
            UpdateWorkButton();
            RefreshValues();
        });
    }

    public void SelectCandidate(int index)
    {
        Ui(delegate
        {
            if (index < 0 || index >= cands.Count) { return; }
            View.SelectedIndex = index;
            View.RowNumber = index + 1;
            View.Record = Rdv3Ledger.SplitLine(cands[index].Line);
            View.StoredState = (cands[index].Stored == null) ? "" : cands[index].Stored;
            View.Saving = false;
            UpdateWorkButton();
            RefreshValues();
        });
    }

    public List<Rdv3CandRow> Candidates { get { return cands; } }
    public int CandidateTotal { get { return candTotal; } }

    public void SetStoredState(int index, string stored, bool saving)
    {
        Ui(delegate
        {
            if (index >= 0 && index < cands.Count) { cands[index].Stored = stored; }
            if (index == View.SelectedIndex && View.Record != null) { View.StoredState = stored; }
            View.Saving = saving && index == View.SelectedIndex;
            UpdateWorkButton();
            RefreshValues();
        });
    }

    private void UpdateWorkButton()
    {
        if (btnWork == null || Screen.Work == null) { return; }
        Rdv3StateDef state = View.HasRecord ? Screen.Work.ByStored(View.StoredState) : Screen.Work.InitialState;
        string stateText = (state == null) ? "" : state.Text;
        btnWork.Text = Screen.Work.ButtonText.Replace("{state}", stateText);
        bool down = View.HasRecord && state != null && state.Id != Screen.Work.Initial;
        btnWork.Checked = down;
        btnWork.ForeColor = down ? Color.Navy : SystemColors.ControlText;
        FontStyle style = down ? FontStyle.Bold : FontStyle.Regular;
        if (btnWork.Font.Style != style) { btnWork.Font = new Font(btnWork.Font, style); }
    }

    public void ClearResult()
    {
        Ui(delegate
        {
            if (txtKey != null) { txtKey.Text = ""; }
            View.SearchKey = "";
            cands = new List<Rdv3CandRow>();
            candTotal = 0;
            View.CandidateCount = 0;
            View.SelectedIndex = -1;
            View.RowNumber = 0;
            View.Record = null;
            View.StoredState = "";
            View.Saving = false;
            UpdateWorkButton();
            RefreshValues();
        });
    }

    public string KeyText { get { return (txtKey == null) ? "" : txtKey.Text.Trim(); } }
    public void SetKeyText(string s) { Ui(delegate { if (txtKey != null) { txtKey.Text = (s == null) ? "" : s; } }); }

    private void RefreshValues()
    {
        for (int i = 0; i < bound.Count; i++)
        {
            BoundControl x = bound[i];
            Rdv3Value value = Rdv3Eval.Evaluate(x.Bind, View, fields, Screen.Work);
            string t = value.Text;
            if (x.EmptyAsDash && (t == null || t.Length == 0)) { t = Rdv3Text.Dash; }
            x.Control.Text = t;
            x.Control.ForeColor = (value.Tone == Rdv3Value.Error) ? Color.Maroon
                : (value.Tone == Rdv3Value.Muted) ? SystemColors.GrayText : SystemColors.ControlText;
        }

        RefreshJudgment();
        RefreshStatus();
        RefreshPending();
        UpdateWorkButton();
    }

    private void RefreshJudgment()
    {
        if (judgmentText == null) { return; }
        if (!View.HasRecord)
        {
            judgmentText.Font = judgmentUnsearchedFont;
            judgmentText.Text = Rdv3Text.Unsearched;
            judgmentText.ForeColor = Color.FromArgb(96, 96, 96);
            judgmentSub.Text = "";
            return;
        }
        Rdv3Verdict verdict = Rdv3Eval.Judge(judgmentDef, View, fields);
        if (verdict.Result == null)
        {
            judgmentText.Font = judgmentUnsearchedFont;
            judgmentText.Text = Rdv3Text.Unsearched;
            judgmentText.ForeColor = Color.FromArgb(96, 96, 96);
        }
        else
        {
            judgmentText.Font = judgmentResultFont;
            judgmentText.Text = verdict.Result.Text;
            judgmentText.ForeColor = (verdict.Result.Look == "ok") ? Color.DarkGreen
                : (verdict.Result.Look == "ng" || verdict.Result.Look == "error") ? Color.Maroon : SystemColors.ControlText;
        }
        Rdv3Section s = (judgmentSub == null) ? null : judgmentSub.Tag as Rdv3Section;
        if (s != null)
        {
            List<string> values = new List<string>();
            for (int i = 0; i < s.Sub.Count; i++)
            {
                string text = Rdv3Eval.Evaluate(s.Sub[i], View, fields, Screen.Work).Text;
                if (text.Length > 0) { values.Add(text); }
            }
            judgmentSub.Text = string.Join(s.Joiner, values.ToArray());
        }
    }

    private void RefreshStatus()
    {
        int n = Math.Min(segmentDefs.Count, statusPanels.Count);
        status.AccessibleName = "";
        for (int i = 0; i < n; i++)
        {
            Rdv3SegmentDef d = segmentDefs[i];
            string t = Rdv3Eval.Evaluate(d.Value, View, fields, Screen.Work).Text;
            statusPanels[i].Text = d.Prefix + t;
            statusPanels[i].ForeColor = SystemColors.ControlText;
            statusPanels[i].ToolTipText = "";
        }
        if (sharedNotice.Length > 0)
        {
            if (Rdv3Clock.MsSince(sharedNoticeAt) >= Screen.ToastMs) { sharedNotice = ""; }
            else if (n > 0)
            {
                int panel = (n > 2) ? 2 : n - 1;
                statusPanels[panel].Text = sharedNotice;
                statusPanels[panel].ForeColor = Color.DarkGreen;
                statusPanels[panel].ToolTipText = sharedNotice;
                status.AccessibleName = sharedNotice;
            }
        }
    }

    private void RefreshPending()
    {
        if (pendingLabel == null || sendDef == null || sendDef.Value == null) { return; }
        pendingLabel.Text = Rdv3Eval.Evaluate(sendDef.Value, View, fields, Screen.Work).Text;
        pendingLabel.ForeColor = (View.PendingCount == 0) ? SystemColors.ControlText : Color.Maroon;
        FontStyle style = (View.PendingCount == 0) ? FontStyle.Regular : FontStyle.Bold;
        if (pendingLabel.Font.Style != style) { pendingLabel.Font = new Font(pendingLabel.Font, style); }
    }

    public string GeometryDump()
    {
        List<Control> controls = new List<Control>();
        CollectControls(this, controls);
        controls.Sort(delegate(Control a, Control b) { return string.Compare(a.Name, b.Name, StringComparison.Ordinal); });
        StringBuilder sb = new StringBuilder();
        sb.Append("{\"card\":[").Append(ClientSize.Width).Append(",").Append(ClientSize.Height).Append("],\"el\":{");
        bool comma = false;
        for (int i = 0; i < controls.Count; i++)
        {
            Control c = controls[i];
            if (c.Name == null || c.Name.Length == 0) { continue; }
            Point p = PointToClient(c.Parent.PointToScreen(c.Location));
            if (comma) { sb.Append(","); }
            comma = true;
            sb.Append("\"").Append(Json(c.Name)).Append("\":[").Append(p.X).Append(",").Append(p.Y).Append(",")
              .Append(c.Width).Append(",").Append(c.Height).Append("]");
        }
        sb.Append("},\"clipped\":[]}");
        return sb.ToString();
    }

    public void CaptureToFile(string path)
    {
        PerformLayout();
        using (Bitmap bitmap = new Bitmap(Width, Height))
        {
            DrawToBitmap(bitmap, new Rectangle(0, 0, bitmap.Width, bitmap.Height));
            bitmap.Save(path, System.Drawing.Imaging.ImageFormat.Png);
        }
    }

    private static void CollectControls(Control root, List<Control> all)
    {
        foreach (Control c in root.Controls) { all.Add(c); CollectControls(c, all); }
    }

    private static string Json(string s) { return s.Replace("\\", "\\\\").Replace("\"", "\\\""); }

    public Rectangle CardBounds { get { return RectangleToScreen(new Rectangle(0, 0, ClientSize.Width, ClientSize.Height)); } }

    public void Notice(string text) { ShowToast(text, true); }
    public void Error(string text) { ShowToast(text, false); }

    public void SharedNotice(string text)
    {
        Ui(delegate
        {
            sharedNotice = (text == null) ? "" : text;
            sharedNoticeAt = Rdv3Clock.Now();
            RefreshStatus();
        });
    }

    private void ShowToast(string text, bool completion)
    {
        Ui(delegate
        {
            if (text == null || text.Length == 0) { return; }
            if (toast == null) { toast = new Rdv3Toast(this); }
            toast.Show(text, completion, Screen.ToastMs);
        });
    }

    public bool Ask(string title, string body) { return Rdv3ConfirmForm.Ask(this, title, body); }

    public void Tell(string title, string body) { Rdv3ConfirmForm.Tell(this, title, body); }

    public bool AskLedgerSwitch(List<Rdv3CandRow> resetRows) { return Rdv3LedgerUpdateForm.Ask(this, resetRows); }

    public void TellResetRows(List<Rdv3CandRow> resetRows) { Rdv3LedgerUpdateForm.TellReset(this, resetRows); }

    public void TellUnmatched(List<Rdv3UnmatchedChange> rows) { Rdv3UnmatchedForm.Tell(this, rows); }

    public void Fatal(string title, string body)
    {
        Rdv3ConfirmForm.Tell(this, title, body);
        Close();
    }

    public int PickFromList()
    {
        if (cands.Count == 0) { return -1; }
        return Rdv3CandidatesForm.Pick(this, Screen.Candidates, cands, candTotal, View.SelectedIndex);
    }

    public string Diag
    {
        get
        {
            Rectangle wa = System.Windows.Forms.Screen.FromControl(this).WorkingArea;
            return "client=" + ClientSize.Width + "x" + ClientSize.Height
                + " work=" + wa.Width + "x" + wa.Height
                + " font=" + Font.FontFamily.Name + " controls=" + CountControls(this).ToString(CultureInfo.InvariantCulture);
        }
    }

    private static int CountControls(Control root)
    {
        int n = root.Controls.Count;
        foreach (Control c in root.Controls) { n += CountControls(c); }
        return n;
    }
}
