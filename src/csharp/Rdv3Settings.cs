// ============================================================================
// Rdv3Settings.cs -- settings and UI Automation picker using stock controls.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Automation;
using System.Windows.Forms;

public sealed class Rdv3SettingsForm : Rdv3Dialog
{
    private readonly Rdv3Config cfg;
    private readonly TextBox dataPath = new TextBox();
    private readonly TextBox ledgerPath = new TextBox();
    private readonly TextBox logPath = new TextBox();
    private readonly TextBox pattern = new TextBox();
    private readonly NumericUpDown candidateRows = new NumericUpDown();
    private readonly TextBox target = new TextBox();
    private readonly TextBox read = new TextBox();
    public Rdv3Config Result;

    public static Rdv3Config Edit(Rdv3Form owner, Rdv3Config current)
    {
        using (Rdv3SettingsForm f = new Rdv3SettingsForm(owner, current.Clone()))
        {
            return (f.ShowOver(owner) == DialogResult.OK) ? f.Result : null;
        }
    }

    private Rdv3SettingsForm(Rdv3Form owner, Rdv3Config working)
        : base(Rdv3Text.SettingsTitle, 540, 376, owner)
    {
        cfg = working;
        Label hint = LabelOf("settings.hint", Rdv3Text.SettingsHint);
        hint.Dock = DockStyle.Top;
        hint.AutoSize = true;
        hint.Padding = new Padding(Rdv3Metrics.Gap, Rdv3Metrics.Gap, Rdv3Metrics.Gap, 0);

        TableLayoutPanel body = new TableLayoutPanel();
        body.Name = "settings.body";
        body.Dock = DockStyle.Fill;
        body.Padding = new Padding(Rdv3Metrics.Gap);
        body.ColumnCount = 1;
        // each group is as tall as it needs to be; the spare row at the bottom
        // takes whatever is left over so no group has to stretch
        body.RowCount = 4;
        int pathsRow = GroupHeight(ButtonFieldRow, 3, Padding.Empty) + Rdv3Metrics.Gap;
        int searchRow = GroupHeight(RowHeight, 2, Padding.Empty) + Rdv3Metrics.Gap;
        int watchRow = Rdv3Metrics.Caption(Font) + Rdv3Metrics.Gap * 3
            + ButtonFieldRow + RowHeight;
        body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        body.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        body.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        body.Controls.Add(BuildPaths(), 0, 0);
        body.Controls.Add(BuildSearch(), 0, 1);
        body.Controls.Add(BuildWatch(), 0, 2);

        FlowLayoutPanel foot = new FlowLayoutPanel();
        foot.Name = "settings.buttons";
        foot.Dock = DockStyle.Bottom;
        foot.FlowDirection = FlowDirection.RightToLeft;
        foot.Padding = new Padding(Rdv3Metrics.Gap);
        foot.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button cancel = ButtonOf("settings.cancel", Rdv3Text.BtnCancel, DialogResult.Cancel);
        Button save = ButtonOf("settings.save", Rdv3Text.BtnOk, DialogResult.None);
        save.Click += delegate { SaveAndClose(); };
        foot.Controls.Add(cancel);
        foot.Controls.Add(save);
        // The controls inside each framed settings group finish two shared
        // gaps from the client edge (body-to-frame, then frame-to-content).
        // Put the footer on that same vertical line.
        AlignRightButtons(foot, Rdv3Metrics.Gap * 2);
        AcceptButton = save;
        CancelButton = cancel;

        ClientSize = new Size(ClientSize.Width,
            Rdv3Metrics.Caption(Font) + hint.Padding.Vertical + body.Padding.Vertical
            + pathsRow + searchRow + watchRow + foot.Height);
        Controls.Add(body);
        Controls.Add(foot);
        Controls.Add(hint);
        LoadAll();
        FinishLayout();
    }

    public static Rdv3SettingsForm ForCheck(Rdv3Config cfg)
    {
        return new Rdv3SettingsForm(null, cfg);
    }

    private GroupBox BuildPaths()
    {
        GroupBox box = Group("settings.pathsGroup", Rdv3Text.SecPlaces);
        TableLayoutPanel grid = Grid("settings.paths", 3, 82, true);
        AddPathRow(grid, 0, Rdv3Text.LblDataShort, dataPath, delegate { BrowseData(); });
        AddPathRow(grid, 1, Rdv3Text.LblLedger, ledgerPath, delegate { BrowseFile(ledgerPath, "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*"); });
        AddPathRow(grid, 2, Rdv3Text.LblLog, logPath, delegate { BrowseFile(logPath, "Log (*.log)|*.log|All files (*.*)|*.*"); });
        box.Controls.Add(grid);
        return box;
    }

    private GroupBox BuildSearch()
    {
        GroupBox box = Group("settings.searchGroup", Rdv3Text.SecSearch);
        TableLayoutPanel grid = Grid("settings.search", 2, 116, false);
        PrepareEdit(pattern, "settings.pattern");
        StretchAcross(pattern);
        pattern.Margin = new Padding(0, 1, 0, 1);
        candidateRows.Name = "settings.candidateRows";
        candidateRows.Minimum = 1;
        candidateRows.Maximum = 1000;
        candidateRows.Width = 74;
        // A spinner insists on its own height, so a Fill dock left it sitting
        // at the top of its cell while the label beside it stayed in the
        // middle. Anchoring it on one side only centres it in the row, and its
        // own margin has to be the same as the boxes above it or the anchor
        // has nothing to centre inside.
        // (a spinner is a container control and scales its own margin a second
        // time on a high-dpi screen, which is what pushed it down; with no
        // margin of its own the anchor has the whole row to centre it in)
        candidateRows.Margin = Padding.Empty;
        candidateRows.Anchor = AnchorStyles.Left;
        AddRow(grid, 0, Rdv3Text.LblKeyPatternShort, pattern);
        AddRow(grid, 1, Rdv3Text.LblCandidateRows, candidateRows);
        box.Controls.Add(grid);
        return box;
    }

    private GroupBox BuildWatch()
    {
        GroupBox box = Group("settings.watchGroup", Rdv3Text.SecTargets);
        TableLayoutPanel grid = new TableLayoutPanel();
        grid.Name = "settings.watch";
        grid.Dock = DockStyle.Top;
        grid.AutoSize = true;
        grid.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        grid.Padding = Padding.Empty;
        grid.ColumnCount = 5;
        grid.RowCount = 2;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, ButtonFieldRow));
        grid.RowStyles.Add(new RowStyle(SizeType.Absolute, RowHeight));
        target.Name = "settings.target";
        target.ReadOnly = true;
        target.TabStop = false;
        target.BackColor = SystemColors.Control;
        target.BorderStyle = BorderStyle.Fixed3D;
        StretchAcross(target);
        target.Margin = new Padding(0, 1, 0, 1);
        read.Name = "settings.read";
        read.ReadOnly = true;
        read.TabStop = false;
        read.BackColor = SystemColors.Control;
        read.BorderStyle = BorderStyle.Fixed3D;
        StretchAcross(read);
        read.Margin = new Padding(0, 1, 0, 1);
        Button pick = ButtonOf("settings.pick", Rdv3Text.BtnInspect, DialogResult.None);
        StretchAcross(pick);
        pick.Margin = Padding.Empty;
        pick.Click += delegate { PickTarget(); };
        AddCellLabel(grid, 0, Rdv3Text.LblTarget);
        AddCellLabel(grid, 1, Rdv3Text.LblRead);
        grid.Controls.Add(target, 2, 0);
        grid.Controls.Add(pick, 4, 0);
        grid.Controls.Add(read, 2, 1);
        grid.SetColumnSpan(read, 3);
        box.Controls.Add(grid);
        return box;
    }

    // The caption line and the frame belong to the GroupBox, and only it knows
    // how tall they come out on this screen. Working the total out here left it
    // a pixel or two short of its own rows, and the row that did not fit was
    // drawn over the bottom of the frame.
    private static GroupBox Group(string name, string title)
    {
        GroupBox box = new GroupBox();
        box.Name = name;
        box.Text = title;
        box.Dock = DockStyle.Top;
        box.AutoSize = true;
        box.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        box.Padding = new Padding(Rdv3Metrics.Gap);
        box.Margin = new Padding(0, 0, 0, Rdv3Metrics.Gap);
        return box;
    }

    private TableLayoutPanel Grid(string name, int rows, int labelWidth, bool browse)
    {
        TableLayoutPanel grid = new TableLayoutPanel();
        grid.Name = name;
        grid.Dock = DockStyle.Top;
        grid.AutoSize = true;
        grid.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        grid.Padding = Padding.Empty;
        grid.ColumnCount = browse ? 5 : 3;
        grid.RowCount = rows;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, labelWidth));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        if (browse)
        {
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        }
        float height = browse ? ButtonFieldRow : RowHeight;
        for (int i = 0; i < rows; i++) { grid.RowStyles.Add(new RowStyle(SizeType.Absolute, height)); }
        return grid;
    }

    // one row of the dialog: a labelled box, measured from the font
    private int RowHeight { get { return Rdv3Metrics.FieldRow(Font); } }

    // a row that also carries a button; a button will not go below the size it
    // needs for its own caption, so the row has to be at least that tall or it
    // hangs over the bottom of the group and rubs out the frame
    private int ButtonFieldRow
    {
        get { return Math.Max(RowHeight, Rdv3Metrics.ButtonHeight(Font)); }
    }

    // a group of rows: the caption, the shared GroupBox padding, the grid's
    // padding, and the rows themselves
    private int GroupHeight(int rowHeight, int rows, Padding gridPadding)
    {
        return Rdv3Metrics.Caption(Font) + Rdv3Metrics.Gap * 2
            + gridPadding.Vertical + rows * rowHeight;
    }

    private static void PrepareEdit(TextBox box, string name)
    {
        box.Name = name;
        box.BackColor = SystemColors.Window;
        box.BorderStyle = BorderStyle.Fixed3D;
    }

    private static void AddPathRow(TableLayoutPanel grid, int row, string label, TextBox box, Action browse)
    {
        PrepareEdit(box, "settings.path" + row.ToString(CultureInfo.InvariantCulture));
        StretchAcross(box);
        box.Margin = new Padding(0, 1, 0, 1);
        Button button = ButtonOf("settings.browse" + row.ToString(CultureInfo.InvariantCulture), Rdv3Text.BtnBrowse, DialogResult.None);
        StretchAcross(button);
        button.Margin = Padding.Empty;
        button.Click += delegate { browse(); };
        AddCellLabel(grid, row, label);
        grid.Controls.Add(box, 2, row);
        grid.Controls.Add(button, 4, row);
    }

    private static void AddRow(TableLayoutPanel grid, int row, string label, Control value)
    {
        AddCellLabel(grid, row, label);
        grid.Controls.Add(value, 2, row);
    }

    private static void AddCellLabel(TableLayoutPanel grid, int row, string text)
    {
        Label label = LabelOf(grid.Name + ".label" + row.ToString(CultureInfo.InvariantCulture), text);
        label.Dock = DockStyle.Fill;
        label.Margin = Padding.Empty;
        label.TextAlign = ContentAlignment.MiddleLeft;
        grid.Controls.Add(label, 0, row);
    }

    private void LoadAll()
    {
        dataPath.Text = cfg.DataDir;
        ledgerPath.Text = cfg.Ledger;
        logPath.Text = cfg.Log;
        pattern.Text = cfg.KeyPattern;
        candidateRows.Value = Math.Max(candidateRows.Minimum, Math.Min(candidateRows.Maximum, cfg.CandidateRowsShown));
        RefreshTarget();
    }

    private Rdv3Target CurrentTarget { get { return (cfg.Targets.Count == 0) ? null : cfg.Targets[0]; } }

    private void RefreshTarget()
    {
        Rdv3Target t = CurrentTarget;
        target.Text = Summary(t);
        string mode = (t == null) ? Rdv3Text.NoValue
            : (t.ReadMode == Rdv3Uia.ReadValue) ? Rdv3Text.ReadValuePattern
            : (t.ReadMode == Rdv3Uia.ReadText) ? Rdv3Text.ReadTextPattern : Rdv3Text.ReadNameProperty;
        read.Text = Rdv3Text.ReadSummaryFmt.Replace("{mode}", mode).Replace("{poll}", cfg.PollMs.ToString(CultureInfo.InvariantCulture));
    }

    private static string Summary(Rdv3Target t)
    {
        if (t == null) { return Rdv3Text.NoteNoTargetShort; }
        string name = (t.Name.Length > 0) ? t.Name : Rdv3Text.NoValue;
        string kind = t.Window.ClassName.Length > 0 ? "className"
            : (t.Window.AutomationId.Length > 0 ? "automationId" : "name");
        string value = t.Window.ClassName.Length > 0 ? t.Window.ClassName
            : (t.Window.AutomationId.Length > 0 ? t.Window.AutomationId : t.Window.NameLike);
        if (value.Length == 0) { value = Rdv3Text.NoValue; }
        return Rdv3Text.NoteTargetSummary.Replace("{name}", name).Replace("{kind}", kind).Replace("{value}", value);
    }

    private void PickTarget()
    {
        Rdv3Target picked = Rdv3PickerForm.Pick(this);
        if (picked == null) { return; }
        Rdv3Target old = CurrentTarget;
        if (old != null)
        {
            if (old.Name.Length > 0 && !old.Name.StartsWith(Rdv3Text.SecTarget)) { picked.Name = old.Name; }
            picked.Enabled = old.Enabled;
            cfg.Targets[0] = picked;
        }
        else { cfg.Targets.Add(picked); }
        RefreshTarget();
    }

    private void BrowseData()
    {
        using (FolderBrowserDialog f = new FolderBrowserDialog())
        {
            try { if (Directory.Exists(dataPath.Text)) { f.SelectedPath = dataPath.Text; } } catch { }
            if (f.ShowDialog(this) == DialogResult.OK) { dataPath.Text = f.SelectedPath; }
        }
    }

    private void BrowseFile(TextBox box, string filter)
    {
        using (SaveFileDialog f = new SaveFileDialog())
        {
            f.Filter = filter;
            try
            {
                f.InitialDirectory = Path.GetDirectoryName(box.Text);
                f.FileName = Path.GetFileName(box.Text);
            }
            catch { }
            if (f.ShowDialog(this) == DialogResult.OK) { box.Text = f.FileName; }
        }
    }

    private void SaveAndClose()
    {
        string p = pattern.Text.Trim();
        string why = Rdv3Config.PatternError(p);
        if (why != null)
        {
            MessageBox.Show(this, Rdv3Text.ErrPatternTyped + why, Rdv3Text.SettingsTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            pattern.Focus();
            return;
        }
        if (dataPath.Text.Trim().Length == 0 || ledgerPath.Text.Trim().Length == 0 || logPath.Text.Trim().Length == 0)
        {
            MessageBox.Show(this, Rdv3Text.ErrPathBlank, Rdv3Text.SettingsTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        cfg.DataDir = dataPath.Text.Trim();
        cfg.Ledger = ledgerPath.Text.Trim();
        cfg.Log = logPath.Text.Trim();
        cfg.KeyPattern = p;
        cfg.CandidateRowsShown = Decimal.ToInt32(candidateRows.Value);
        Result = cfg;
        DialogResult = DialogResult.OK;
        Close();
    }
}

public sealed class Rdv3PickerForm : Rdv3Dialog
{
    private readonly System.Windows.Forms.Timer tick = new System.Windows.Forms.Timer();
    private readonly TextBox[] values = new TextBox[6];
    private AutomationElement hover;
    private string ctrlType = "", autoId = "", className = "", elName = "", value = "", proc = "";
    private bool armed;
    private bool canRead;
    private Rdv3Target result;

    public static Rdv3Target Pick(IWin32Window owner)
    {
        using (Rdv3PickerForm f = new Rdv3PickerForm())
        {
            return (f.ShowDialog(owner) == DialogResult.OK) ? f.result : null;
        }
    }

    private Rdv3PickerForm()
        : base(Rdv3Text.PickTitle, 460, 330, null)
    {
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        Rectangle wa = System.Windows.Forms.Screen.PrimaryScreen.WorkingArea;
        Location = new Point(wa.Right - Width - 20, wa.Bottom - Height - 56);

        Label how = LabelOf("picker.how", Rdv3Text.PickHow);
        how.Dock = DockStyle.Top;
        how.Padding = new Padding(Rdv3Metrics.Gap);
        how.Height = Rdv3Metrics.Caption(Font) + how.Padding.Vertical;
        how.Font = new Font(Font, FontStyle.Bold);

        TableLayoutPanel grid = new TableLayoutPanel();
        grid.Name = "picker.values";
        grid.Dock = DockStyle.Fill;
        grid.Padding = new Padding(Rdv3Metrics.Gap);
        grid.ColumnCount = 2;
        grid.RowCount = 7;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 118f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        string[] labels = { Rdv3Text.LblControlTypes, Rdv3Text.LblAutomationId, Rdv3Text.LblClassName,
                            Rdv3Text.LblName, Rdv3Text.LblProcessOf, Rdv3Text.PickReading };
        for (int i = 0; i < 6; i++)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, Rdv3Metrics.FieldRow(Font) + 6));
            Label label = LabelOf("picker.label" + i.ToString(CultureInfo.InvariantCulture), labels[i]);
            label.Dock = DockStyle.Fill;
            label.TextAlign = ContentAlignment.MiddleLeft;
            values[i] = ReadOnlyBox("picker.value" + i.ToString(CultureInfo.InvariantCulture));
            StretchAcross(values[i]);
            values[i].Margin = new Padding(Rdv3Metrics.Gap, 2, 0, 2);
            grid.Controls.Add(label, 0, i);
            grid.Controls.Add(values[i], 1, i);
        }
        // the spare row keeps the leftover height out of the last field's row
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        FlowLayoutPanel foot = new FlowLayoutPanel();
        foot.Name = "picker.buttons";
        foot.Dock = DockStyle.Bottom;
        foot.FlowDirection = FlowDirection.RightToLeft;
        foot.Padding = new Padding(Rdv3Metrics.Gap);
        foot.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button close = ButtonOf("picker.close", Rdv3Text.BtnClose, DialogResult.Cancel);
        Label esc = LabelOf("picker.escape", Rdv3Text.PickEsc);
        esc.AutoSize = true;
        esc.Margin = Padding.Empty;
        foot.Controls.Add(close);
        foot.Controls.Add(esc);
        AlignRightButtons(foot);
        CancelButton = close;

        Controls.Add(grid);
        Controls.Add(foot);
        Controls.Add(how);
        tick.Interval = 90;
        tick.Tick += delegate { Look(); };
        tick.Start();
        FormClosed += delegate { tick.Stop(); };
        RefreshValues();
        FinishLayout();
    }

    public static Rdv3PickerForm ForCheck()
    {
        Rdv3PickerForm f = new Rdv3PickerForm();
        f.tick.Stop();
        return f;
    }

    public void SetSample(string type, string id, string cls, string name, string process, string readValue)
    {
        ctrlType = type;
        autoId = id;
        className = cls;
        elName = name;
        proc = process;
        value = (readValue == null) ? "" : readValue;
        canRead = value.Length > 0;
        RefreshValues();
    }

    private void RefreshValues()
    {
        string[] data = { ctrlType, autoId, className, elName, proc, canRead ? value : Rdv3Text.PickNoRead };
        for (int i = 0; i < values.Length; i++)
        {
            if (values[i] == null) { continue; }
            values[i].Text = (data[i] == null || data[i].Length == 0) ? Rdv3Text.NotYet : data[i];
            values[i].ForeColor = (i == 5 && !canRead) ? SystemColors.GrayText : SystemColors.ControlText;
        }
    }

    private void Look()
    {
        Point p = Cursor.Position;
        if (Bounds.Contains(p)) { return; }
        AutomationElement e = null;
        try { e = AutomationElement.FromPoint(new System.Windows.Point(p.X, p.Y)); }
        catch { e = null; }
        if (e != null && !Rdv3Uia.Same(e, hover))
        {
            hover = e;
            try
            {
                AutomationElement.AutomationElementInformation c = e.Current;
                ctrlType = Rdv3Uia.ControlTypeName(c.ControlType);
                autoId = c.AutomationId;
                className = c.ClassName;
                elName = Cut(c.Name, 46);
                proc = ProcName(c.ProcessId);
            }
            catch { }
            string v = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadValue);
            if (v == null) { v = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadText); }
            canRead = v != null;
            value = (v == null) ? "" : Cut(Rdv3Watch.Candidate(v), 46);
            RefreshValues();
        }
        bool down = (Control.ModifierKeys & (Keys.Control | Keys.Shift)) == (Keys.Control | Keys.Shift);
        if (down && !armed && hover != null) { armed = true; Take(); }
        else if (!down) { armed = false; }
    }

    private static string ProcName(int pid)
    {
        try { return System.Diagnostics.Process.GetProcessById(pid).ProcessName; }
        catch { return ""; }
    }

    private static string Cut(string s, int n)
    {
        if (s == null) { return ""; }
        s = s.Replace("\r", " ").Replace("\n", " ");
        return (s.Length <= n) ? s : s.Substring(0, n) + "...";
    }

    private void Take()
    {
        AutomationElement e = hover;
        if (e == null) { return; }
        try
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement root = AutomationElement.RootElement;
            List<AutomationElement> chain = new List<AutomationElement>();
            AutomationElement at = e;
            for (int guard = 0; guard < 64 && at != null; guard++)
            {
                AutomationElement parent = walker.GetParent(at);
                if (parent == null || Rdv3Uia.Same(parent, root)) { break; }
                chain.Add(parent);
                at = parent;
            }
            chain.Reverse();
            AutomationElement win = chain.Count > 0 ? chain[0] : e;
            Rdv3Target t = new Rdv3Target();
            t.Window = new Rdv3Match();
            t.Window.Descendants = false;
            t.Window.ClassName = win.Current.ClassName;
            t.Window.ProcessName = ProcName(win.Current.ProcessId);
            if (win.Current.AutomationId.Length > 0) { t.Window.AutomationId = win.Current.AutomationId; }
            t.Name = win.Current.Name.Length > 0 ? Cut(win.Current.Name, 24) : t.Window.ProcessName;
            for (int i = 1; i < chain.Count; i++)
            {
                string id = chain[i].Current.AutomationId;
                if (id.Length == 0) { continue; }
                Rdv3Match step = new Rdv3Match();
                step.AutomationId = id;
                step.Descendants = true;
                t.Steps.Add(step);
            }
            t.Field = new Rdv3Match();
            t.Field.Descendants = true;
            AutomationElement.AutomationElementInformation c = e.Current;
            if (c.AutomationId.Length > 0) { t.Field.AutomationId = c.AutomationId; }
            else if (c.ClassName.Length > 0) { t.Field.ClassName = c.ClassName; }
            string ct = Rdv3Uia.ControlTypeName(c.ControlType);
            if (ct.Length > 0) { t.Field.ControlTypes = new string[] { ct }; }
            if (t.Field.IsEmpty && c.Name.Length > 0) { t.Field.Name = c.Name; }
            t.ReadMode = Rdv3Uia.ReadValue;
            if (Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadValue) == null)
            {
                t.ReadMode = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadText) != null ? Rdv3Uia.ReadText : Rdv3Uia.ReadName;
            }
            if (t.ReadMode == Rdv3Uia.ReadValue) { t.Field.RequireValuePattern = true; }
            AutomationElement anchor = win;
            for (int i = 1; i < chain.Count; i++)
            {
                if (chain[i].Current.AutomationId.Length > 0) { anchor = chain[i]; }
            }
            List<AutomationElement> same = Rdv3Uia.FindAll(anchor, t.Field, t.Field.Descendants);
            for (int i = 0; i < same.Count; i++)
            {
                if (Rdv3Uia.Same(same[i], e)) { t.Field.Index = i; break; }
            }
            result = t;
            DialogResult = DialogResult.OK;
            Close();
        }
        catch { }
    }
}
