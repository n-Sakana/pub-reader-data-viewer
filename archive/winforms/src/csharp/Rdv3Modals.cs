// ============================================================================
// Rdv3Modals.cs -- standard WinForms dialogs used by the viewer.
//
// Every visible part is a stock control. List-shaped data always uses the
// same Details ListView, including candidates, job inputs and process steps.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

public class Rdv3Dialog : Form
{
    protected Rdv3Dialog(string title, int width, int height, Rdv3Form owner)
    {
        SuspendLayout();
        Font = (owner != null) ? owner.Font
            : (Rdv3Metrics.ScreenFont != null) ? Rdv3Metrics.ScreenFont : SystemFonts.MessageBoxFont;
        AutoScaleDimensions = new SizeF(96.0f, 96.0f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ClientSize = new Size(width, height);

        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        BackColor = SystemColors.Control;
    }

    protected DialogResult ShowOver(IWin32Window owner)
    {
        return (owner == null) ? ShowDialog() : ShowDialog(owner);
    }

    protected void FinishLayout()
    {
        ResumeLayout(false);
        PerformLayout();
    }

    public string GeometryDump()
    {
        List<Control> all = new List<Control>();
        Collect(this, all);
        all.Sort(delegate(Control a, Control b) { return string.Compare(a.Name, b.Name, StringComparison.Ordinal); });
        StringBuilder sb = new StringBuilder();
        sb.Append("{\"card\":[").Append(ClientSize.Width).Append(",").Append(ClientSize.Height).Append("],\"el\":{");
        bool comma = false;
        for (int i = 0; i < all.Count; i++)
        {
            Control c = all[i];
            if (c.Name == null || c.Name.Length == 0) { continue; }
            Point p = PointToClient(c.Parent.PointToScreen(c.Location));
            if (comma) { sb.Append(','); }
            comma = true;
            sb.Append('"').Append(Escape(c.Name)).Append("\":[")
              .Append(p.X).Append(',').Append(p.Y).Append(',').Append(c.Width).Append(',').Append(c.Height).Append(']');
        }
        sb.Append("},");
        Rdv3Geometry.AppendClipped(sb, this);
        sb.Append('}');
        return sb.ToString();
    }

    private static void Collect(Control root, List<Control> all)
    {
        foreach (Control c in root.Controls) { all.Add(c); Collect(c, all); }
    }

    private static string Escape(string s) { return s.Replace("\\", "\\\\").Replace("\"", "\\\""); }

    protected static Label LabelOf(string name, string text)
    {
        Label l = new Label();
        l.Name = name;
        l.Text = text;
        l.AutoEllipsis = true;
        l.UseCompatibleTextRendering = false;
        return l;
    }

    protected static TextBox ReadOnlyBox(string name)
    {
        TextBox t = new TextBox();
        t.Name = name;
        t.ReadOnly = true;
        t.TabStop = false;
        t.BackColor = SystemColors.Control;
        t.BorderStyle = BorderStyle.Fixed3D;
        return t;
    }

    protected static ListView DetailsList(string name)
    {
        ListView v = new ListView();
        v.Name = name;
        v.View = View.Details;
        v.FullRowSelect = true;
        v.GridLines = true;
        v.HideSelection = false;
        v.MultiSelect = false;
        v.HeaderStyle = ColumnHeaderStyle.Nonclickable;
        v.BackColor = SystemColors.Window;
        v.BorderStyle = BorderStyle.Fixed3D;
        return v;
    }

    protected static void FitColumns(ListView view, int[] weights)
    {
        if (view == null || weights == null || view.Columns.Count != weights.Length) { return; }
        int available = view.ClientSize.Width - SystemInformation.VerticalScrollBarWidth - 6;
        if (available <= 0) { return; }
        int total = 0;
        for (int i = 0; i < weights.Length; i++) { total += Math.Max(1, weights[i]); }
        int used = 0;
        for (int i = 0; i < weights.Length; i++)
        {
            int width = (i == weights.Length - 1)
                ? Math.Max(1, available - used)
                : Math.Max(1, available * Math.Max(1, weights[i]) / total);
            view.Columns[i].Width = width;
            used += width;
        }
    }

    protected static void KeepColumnsFitted(ListView view, int[] weights)
    {
        view.Resize += delegate { FitColumns(view, weights); };
        FitColumns(view, weights);
    }

    // A box that keeps its own height sits at the top of a Fill dock and ends
    // up out of line with the label beside it; anchored on the two sides it
    // stretches across the row and stays in the middle of it.
    protected static void StretchAcross(Control c)
    {
        c.Dock = DockStyle.None;
        c.Anchor = AnchorStyles.Left | AnchorStyles.Right;
        // A table row already centres a control vertically.  Keeping the
        // button's default top/bottom margin as well would take that room away
        // from its preferred height at 96 dpi.
        if (c is ButtonBase) { c.Margin = new Padding(c.Margin.Left, 0, c.Margin.Right, 0); }
    }

    protected static Button ButtonOf(string name, string text, DialogResult result)
    {
        Button b = new Button();
        b.Name = name;
        b.Text = text;
        if (Rdv3Metrics.ScreenFont != null) { b.Font = Rdv3Metrics.ScreenFont; }
        b.DialogResult = result;
        b.FlatStyle = FlatStyle.Standard;
        b.UseVisualStyleBackColor = false;
        b.AutoSize = true;
        b.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        b.MinimumSize = new Size(82, 26);
        return b;
    }

    // FlowLayoutPanel consumes the horizontal sum of its padding from the
    // right when it flows right-to-left. Choose that inset once, remove the
    // outer button margin, and use the shared gap only between neighbouring
    // buttons. The right edge is then decided once per dialog.
    protected static void AlignRightButtons(FlowLayoutPanel bar)
    {
        AlignRightButtons(bar, Rdv3Metrics.Gap);
    }

    protected static void AlignRightButtons(FlowLayoutPanel bar, int rightInset)
    {
        Padding padding = bar.Padding;
        bar.FlowDirection = FlowDirection.RightToLeft;
        bar.WrapContents = false;
        bar.Padding = new Padding(Math.Max(0, rightInset), padding.Top, 0, padding.Bottom);
        for (int i = 0; i < bar.Controls.Count; i++)
        {
            Control button = bar.Controls[i];
            button.Margin = new Padding(0, 0, i == 0 ? 0 : Rdv3Metrics.Gap, 0);
        }
    }
}

public sealed class Rdv3ConfirmForm : Rdv3Dialog
{
    public static bool Ask(Rdv3Form owner, string title, string body)
    {
        using (Rdv3ConfirmForm f = new Rdv3ConfirmForm(owner, title, body, false))
        {
            return f.ShowOver(owner) == DialogResult.Yes;
        }
    }

    public static void Tell(Rdv3Form owner, string title, string body)
    {
        using (Rdv3ConfirmForm f = new Rdv3ConfirmForm(owner, title, body, true)) { f.ShowOver(owner); }
    }

    private Rdv3ConfirmForm(Rdv3Form owner, string title, string text, bool one)
        : base(title, 470, 190, owner)
    {
        Label body = LabelOf("confirm.body", text);
        body.Dock = DockStyle.Fill;
        body.Padding = new Padding(Rdv3Metrics.Gap);
        body.TextAlign = ContentAlignment.MiddleLeft;
        body.AutoEllipsis = false;

        FlowLayoutPanel buttons = new FlowLayoutPanel();
        buttons.Name = "confirm.buttons";
        buttons.Dock = DockStyle.Bottom;
        buttons.FlowDirection = FlowDirection.RightToLeft;
        buttons.WrapContents = false;
        buttons.Padding = new Padding(Rdv3Metrics.Gap);
        buttons.Height = Rdv3Metrics.ButtonBarHeight(Font);

        if (one)
        {
            Button ok = ButtonOf("confirm.ok", Rdv3Text.BtnOk, DialogResult.OK);
            buttons.Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = ok;
        }
        else
        {
            Button no = ButtonOf("confirm.no", Rdv3Text.BtnNo, DialogResult.No);
            Button yes = ButtonOf("confirm.yes", Rdv3Text.BtnYes, DialogResult.Yes);
            buttons.Controls.Add(no);
            buttons.Controls.Add(yes);
            AcceptButton = yes;
            CancelButton = no;
        }
        AlignRightButtons(buttons);
        int textWidth = Math.Max(120, ClientSize.Width - body.Padding.Horizontal);
        int logicalLineHeight = Rdv3Metrics.TextHeight(Font);
        double measureScale = Math.Max(1.0, (double)Font.Height / logicalLineHeight);
        int physicalTextWidth = Math.Max(120, (int)Math.Round(textWidth * measureScale));
        Size measured = TextRenderer.MeasureText(text, Font, new Size(physicalTextWidth, Int32.MaxValue),
            TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
        int measuredLogicalHeight = (int)Math.Ceiling(measured.Height / measureScale);
        int scaledPaddingHeight = (int)Math.Ceiling(body.Padding.Vertical * measureScale);
        int explicitLines = 1;
        for (int i = 0; i < text.Length; i++) { if (text[i] == '\n') { explicitLines++; } }
        int explicitHeight = explicitLines * logicalLineHeight + scaledPaddingHeight + 4;
        int bodyHeight = Math.Max(logicalLineHeight * 2 + 20,
            Math.Max(measuredLogicalHeight + scaledPaddingHeight, explicitHeight));
        ClientSize = new Size(ClientSize.Width, bodyHeight + buttons.Height);
        Controls.Add(body);
        Controls.Add(buttons);
        FinishLayout();
    }

    public static Rdv3ConfirmForm ForCheck(string title, string body)
    {
        return new Rdv3ConfirmForm(null, title, body, false);
    }
}

public sealed class Rdv3CandidatesForm : Rdv3Dialog
{
    private readonly Rdv3CandidatesDef def;
    private readonly Rdv3Fields fields;
    private readonly Rdv3WorkState work;
    private readonly Label hint;
    private readonly ListView table;
    public int Picked = -1;

    public static int Pick(Rdv3Form owner, Rdv3CandidatesDef def, List<Rdv3CandRow> rows, int total, int selected)
    {
        using (Rdv3CandidatesForm f = new Rdv3CandidatesForm(owner, def, owner.Fields, owner.Screen.Work, rows, total, selected))
        {
            return (f.ShowOver(owner) == DialogResult.OK) ? f.Picked : -1;
        }
    }

    private Rdv3CandidatesForm(Rdv3Form owner, Rdv3CandidatesDef d, Rdv3Fields fs, Rdv3WorkState w,
                               List<Rdv3CandRow> values, int total, int selected)
        : base(d.Title, Math.Min(980, Math.Max(620, (int)Math.Round(d.Width))),
               Math.Min(500, Math.Max(260, (int)Math.Round(d.MaxHeight) + 92)), owner)
    {
        def = d;
        fields = fs;
        work = w;
        hint = LabelOf("candidates.hint", d.Hint);
        hint.Dock = DockStyle.Top;
        hint.AutoSize = true;
        hint.Padding = new Padding(Rdv3Metrics.Gap, Rdv3Metrics.Gap, Rdv3Metrics.Gap, 0);

        table = DetailsList("candidates.list");
        table.Dock = DockStyle.Fill;
        int[] columnWeights = new int[d.Columns.Count];
        for (int i = 0; i < d.Columns.Count; i++)
        {
            Rdv3ColumnDef c = d.Columns[i];
            ColumnHeader h = new ColumnHeader();
            h.Text = c.Header;
            h.Width = (c.Width > 0) ? (int)Math.Round(c.Width) : 100;
            columnWeights[i] = h.Width;
            h.TextAlign = (c.Align == "right") ? HorizontalAlignment.Right : HorizontalAlignment.Left;
            table.Columns.Add(h);
        }
        KeepColumnsFitted(table, columnWeights);
        table.DoubleClick += delegate { AcceptSelection(); };
        table.KeyDown += delegate(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; AcceptSelection(); }
        };

        FlowLayoutPanel foot = new FlowLayoutPanel();
        foot.Name = "candidates.buttons";
        foot.Dock = DockStyle.Bottom;
        foot.FlowDirection = FlowDirection.RightToLeft;
        foot.Padding = new Padding(Rdv3Metrics.Gap);
        foot.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button cancel = ButtonOf("candidates.cancel", Rdv3Text.BtnCancel, DialogResult.Cancel);
        Button ok = ButtonOf("candidates.ok", Rdv3Text.BtnOk, DialogResult.None);
        ok.Click += delegate { AcceptSelection(); };
        foot.Controls.Add(cancel);
        foot.Controls.Add(ok);
        AlignRightButtons(foot);
        AcceptButton = ok;
        CancelButton = cancel;

        Panel center = new Panel();
        center.Name = "candidates.center";
        center.Dock = DockStyle.Fill;
        center.Padding = new Padding(Rdv3Metrics.Gap);
        center.Controls.Add(table);
        Controls.Add(center);
        Controls.Add(foot);
        Controls.Add(hint);
        SetRows(values, total, selected);
        FinishLayout();
    }

    private void SetRows(List<Rdv3CandRow> values, int total, int selected)
    {
        List<Rdv3CandRow> rows = (values == null) ? new List<Rdv3CandRow>() : values;
        hint.Text = def.Hint + "  " + Rdv3Text.CandidateHitsFmt.Replace("{n}", total.ToString("N0", CultureInfo.InvariantCulture));
        table.BeginUpdate();
        table.Items.Clear();
        for (int i = 0; i < rows.Count; i++)
        {
            Rdv3View view = new Rdv3View();
            view.Record = Rdv3Ledger.SplitLine(rows[i].Line);
            view.StoredState = (rows[i].Stored == null) ? "" : rows[i].Stored;
            view.RowNumber = i + 1;
            ListViewItem item = null;
            for (int k = 0; k < def.Columns.Count; k++)
            {
                string value = Rdv3Eval.Evaluate(def.Columns[k].Value, view, fields, work).Text;
                if (k == 0) { item = new ListViewItem(value); }
                else { item.SubItems.Add(value); }
            }
            item.Tag = i;
            table.Items.Add(item);
        }
        table.EndUpdate();
        int choose = (selected >= 0 && selected < table.Items.Count) ? selected : ((table.Items.Count > 0) ? 0 : -1);
        if (choose >= 0)
        {
            table.Items[choose].Selected = true;
            table.Items[choose].Focused = true;
            table.Items[choose].EnsureVisible();
        }
    }

    private void AcceptSelection()
    {
        if (table.SelectedIndices.Count == 0) { return; }
        Picked = table.SelectedIndices[0];
        DialogResult = DialogResult.OK;
        Close();
    }

    public static Rdv3CandidatesForm ForCheck(Rdv3CandidatesDef def, Rdv3Fields fields, Rdv3WorkState work,
                                               List<Rdv3CandRow> rows, int total, int selected)
    {
        return new Rdv3CandidatesForm(null, def, fields, work, rows, total, selected);
    }
}

public sealed class Rdv3LedgerUpdateForm : Rdv3Dialog
{
    public static bool Ask(Rdv3Form owner, List<Rdv3CandRow> resetRows)
    {
        using (Rdv3LedgerUpdateForm f = new Rdv3LedgerUpdateForm(owner, resetRows, true))
        {
            return f.ShowOver(owner) == DialogResult.OK;
        }
    }

    public static void TellReset(Rdv3Form owner, List<Rdv3CandRow> resetRows)
    {
        using (Rdv3LedgerUpdateForm f = new Rdv3LedgerUpdateForm(owner, resetRows, false)) { f.ShowOver(owner); }
    }

    public static Rdv3LedgerUpdateForm ForCheck(Rdv3Form owner, List<Rdv3CandRow> resetRows)
    {
        return new Rdv3LedgerUpdateForm(owner, resetRows, false);
    }

    private Rdv3LedgerUpdateForm(Rdv3Form owner, List<Rdv3CandRow> values, bool askSwitch)
        : base(Rdv3Text.SharedUpdateTitle, 744,
               (values != null && values.Count > 0) ? 330 : 190, owner)
    {
        List<Rdv3CandRow> rows = (values == null) ? new List<Rdv3CandRow>() : values;
        Label body = LabelOf("sharedUpdate.body", askSwitch ? Rdv3Text.SharedUpdateBody : Rdv3Text.NoteUpdated);
        body.Dock = DockStyle.Top;
        body.Height = (rows.Count > 0) ? 52 : 112;
        body.Padding = new Padding(Rdv3Metrics.Gap);
        body.TextAlign = ContentAlignment.TopLeft;
        if (rows.Count > 0)
        {
            body.Text += "\r\n" + Rdv3Text.SharedResetFmt
                .Replace("{state}", owner.Screen.Work.InitialState.Text)
                .Replace("{n}", rows.Count.ToString("N0", CultureInfo.InvariantCulture));
        }

        FlowLayoutPanel buttons = new FlowLayoutPanel();
        buttons.Name = "sharedUpdate.buttons";
        buttons.Dock = DockStyle.Bottom;
        buttons.FlowDirection = FlowDirection.RightToLeft;
        buttons.WrapContents = false;
        buttons.Padding = new Padding(Rdv3Metrics.Gap);
        buttons.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button ok = ButtonOf("sharedUpdate.ok", Rdv3Text.BtnOk, DialogResult.OK);
        if (askSwitch)
        {
            Button cancel = ButtonOf("sharedUpdate.cancel", Rdv3Text.BtnCancel, DialogResult.Cancel);
            buttons.Controls.Add(cancel);
            CancelButton = cancel;
        }
        buttons.Controls.Add(ok);
        AlignRightButtons(buttons);
        AcceptButton = ok;
        if (!askSwitch) { CancelButton = ok; }

        if (rows.Count > 0)
        {
            ListView table = DetailsList("sharedUpdate.list");
            table.Dock = DockStyle.Fill;
            Rdv3CandidatesDef def = owner.Screen.Candidates;
            int[] weights = new int[def.Columns.Count];
            for (int i = 0; i < def.Columns.Count; i++)
            {
                Rdv3ColumnDef c = def.Columns[i];
                int width = (c.Width > 0) ? (int)Math.Round(c.Width) : 100;
                table.Columns.Add(c.Header, width,
                    c.Align == "right" ? HorizontalAlignment.Right : HorizontalAlignment.Left);
                weights[i] = width;
            }
            KeepColumnsFitted(table, weights);
            for (int i = 0; i < rows.Count; i++)
            {
                Rdv3View view = new Rdv3View();
                view.Record = Rdv3Ledger.SplitLine(rows[i].Line);
                view.StoredState = rows[i].Stored;
                view.RowNumber = i + 1;
                ListViewItem item = null;
                for (int k = 0; k < def.Columns.Count; k++)
                {
                    string text = Rdv3Eval.Evaluate(def.Columns[k].Value, view, owner.Fields, owner.Screen.Work).Text;
                    if (k == 0) { item = new ListViewItem(text); }
                    else { item.SubItems.Add(text); }
                }
                table.Items.Add(item);
            }
            Panel center = new Panel();
            center.Name = "sharedUpdate.center";
            center.Dock = DockStyle.Fill;
            center.Padding = new Padding(Rdv3Metrics.Gap);
            center.Controls.Add(table);
            Controls.Add(center);
        }
        Controls.Add(buttons);
        Controls.Add(body);
        FinishLayout();
    }
}

public sealed class Rdv3UnmatchedForm : Rdv3Dialog
{
    public static void Tell(Rdv3Form owner, List<Rdv3UnmatchedChange> values)
    {
        using (Rdv3UnmatchedForm f = new Rdv3UnmatchedForm(owner, values)) { f.ShowOver(owner); }
    }

    private Rdv3UnmatchedForm(Rdv3Form owner, List<Rdv3UnmatchedChange> values)
        : base(Rdv3Text.UnmatchedTitle, 560, 280, owner)
    {
        List<Rdv3UnmatchedChange> rows = (values == null) ? new List<Rdv3UnmatchedChange>() : values;
        Label body = LabelOf("unmatched.body", Rdv3Text.UnmatchedBodyFmt.Replace("{n}", rows.Count.ToString("N0", CultureInfo.InvariantCulture)));
        body.Dock = DockStyle.Top;
        body.Height = 48;
        body.Padding = new Padding(Rdv3Metrics.Gap, Rdv3Metrics.Gap, Rdv3Metrics.Gap, 0);

        ListView table = DetailsList("unmatched.list");
        table.Dock = DockStyle.Fill;
        table.Columns.Add(Rdv3Text.ColNumber, 42, HorizontalAlignment.Right);
        table.Columns.Add(Rdv3Text.ColKey, 210);
        table.Columns.Add(Rdv3Text.ColReason, 270);
        KeepColumnsFitted(table, new int[] { 42, 210, 270 });
        for (int i = 0; i < rows.Count; i++)
        {
            ListViewItem item = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
            item.SubItems.Add(rows[i].Identity);
            item.SubItems.Add(rows[i].Reason == "missing" ? Rdv3Text.UnmatchedMissing : Rdv3Text.UnmatchedChanged);
            table.Items.Add(item);
        }

        FlowLayoutPanel buttons = new FlowLayoutPanel();
        buttons.Name = "unmatched.buttons";
        buttons.Dock = DockStyle.Bottom;
        buttons.FlowDirection = FlowDirection.RightToLeft;
        buttons.Padding = new Padding(Rdv3Metrics.Gap);
        buttons.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button ok = ButtonOf("unmatched.ok", Rdv3Text.BtnOk, DialogResult.OK);
        buttons.Controls.Add(ok);
        AlignRightButtons(buttons);
        AcceptButton = ok;
        CancelButton = ok;

        Panel center = new Panel();
        center.Name = "unmatched.center";
        center.Dock = DockStyle.Fill;
        center.Padding = new Padding(Rdv3Metrics.Gap);
        center.Controls.Add(table);
        Controls.Add(center);
        Controls.Add(buttons);
        Controls.Add(body);
        FinishLayout();
    }
}

public sealed class Rdv3ProcessForm : Rdv3Dialog
{
    private readonly Rdv3Data data;
    private readonly Rdv3ProcessJobDef job;
    private readonly string dataDir;
    private readonly string ledgerPath;
    private readonly ListView inputs;
    private readonly ListView steps;
    private readonly Button execute;
    private bool inputsOk = true;

    public static bool ShowJob(Rdv3Form owner, Rdv3Data data, string jobId, string dataDir, string ledgerPath)
    {
        Rdv3ProcessJobDef job = data.JobOf(jobId);
        if (job == null) { return false; }
        using (Rdv3ProcessForm f = new Rdv3ProcessForm(owner, data, job, dataDir, ledgerPath))
        {
            return f.ShowOver(owner) == DialogResult.OK;
        }
    }

    public static Rdv3ProcessForm ForCheck(Rdv3Data data, string jobId, string dataDir, string ledgerPath)
    {
        Rdv3ProcessJobDef job = data.JobOf(jobId);
        if (job == null) { return null; }
        return new Rdv3ProcessForm(null, data, job, dataDir, ledgerPath);
    }

    private Rdv3ProcessForm(Rdv3Form owner, Rdv3Data d, Rdv3ProcessJobDef j, string dir, string ledger)
        : base(j.Kind == "delete" ? Rdv3Text.DeleteRecordsTitle : Rdv3Text.UpdateRecordsTitle, 640, 442, owner)
    {
        data = d;
        job = j;
        dataDir = dir;
        ledgerPath = ledger;

        Label hint = LabelOf("process.hint", j.Kind == "delete" ? Rdv3Text.DeleteRecordsHint : Rdv3Text.UpdateRecordsHint);
        hint.Dock = DockStyle.Top;
        hint.AutoSize = true;
        hint.Padding = new Padding(Rdv3Metrics.Gap, Rdv3Metrics.Gap, Rdv3Metrics.Gap, 0);

        TableLayoutPanel body = new TableLayoutPanel();
        body.Name = "process.body";
        body.Dock = DockStyle.Fill;
        body.Padding = new Padding(Rdv3Metrics.Gap);
        body.ColumnCount = 1;
        body.RowCount = 3;
        // Each group holds a caption, its own padding and a list; the number
        // of rows each list shows is what the height is for, so ask the font
        // how tall a row and a header come out rather than pinning the sum.
        int listChrome = Rdv3Metrics.Caption(Font) + Rdv3Metrics.Gap * 3;
        int inputsRow = listChrome + Rdv3Metrics.ListHeight(Font, InputListRows);
        int stepsRow = listChrome + Rdv3Metrics.ListHeight(Font, StepListRows);
        body.RowStyles.Add(new RowStyle(SizeType.Absolute, inputsRow));
        // the step list takes whatever the dialog has left over
        body.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        body.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        GroupBox inputGroup = new GroupBox();
        inputGroup.Name = "process.inputsGroup";
        inputGroup.Text = Rdv3Text.SecInputs.Replace("{dir}", new DirectoryInfo(dir).Name);
        inputGroup.Dock = DockStyle.Fill;
        inputGroup.Padding = new Padding(Rdv3Metrics.Gap);
        inputGroup.Margin = new Padding(0, 0, 0, Rdv3Metrics.Gap);
        inputs = DetailsList("process.inputs");
        inputs.Dock = DockStyle.Fill;
        inputs.Columns.Add(j.Kind == "delete" ? Rdv3Text.ColDeleteInput : Rdv3Text.ColInput, 40);
        inputs.Columns.Add(Rdv3Text.ColFile, 150);
        inputs.Columns.Add(Rdv3Text.ColKey, 62);
        inputs.Columns.Add(Rdv3Text.ColRows, 74, HorizontalAlignment.Right);
        inputs.Columns.Add(Rdv3Text.ColValidation, 250);
        KeepColumnsFitted(inputs, new int[] { 40, 150, 62, 74, 250 });
        inputGroup.Controls.Add(inputs);

        GroupBox stepGroup = new GroupBox();
        stepGroup.Name = "process.stepsGroup";
        stepGroup.Text = Rdv3Text.SecProcess;
        stepGroup.Dock = DockStyle.Fill;
        stepGroup.Padding = new Padding(Rdv3Metrics.Gap);
        stepGroup.Margin = new Padding(0, 0, 0, Rdv3Metrics.Gap);
        steps = DetailsList("process.steps");
        steps.Dock = DockStyle.Fill;
        steps.Columns.Add(Rdv3Text.ColNumber, 30, HorizontalAlignment.Right);
        steps.Columns.Add(Rdv3Text.ColOperation, 58);
        steps.Columns.Add(Rdv3Text.ColTarget1, 92);
        steps.Columns.Add(Rdv3Text.ColTarget2, 110);
        steps.Columns.Add(Rdv3Text.ColKey, 78);
        steps.Columns.Add(Rdv3Text.ColCondition, 84);
        steps.Columns.Add(Rdv3Text.ColOutput, 120);
        KeepColumnsFitted(steps, new int[] { 26, 46, 72, 94, 56, 210, 82 });
        stepGroup.Controls.Add(steps);

        GroupBox outGroup = new GroupBox();
        outGroup.Name = "process.outputGroup";
        outGroup.Text = Rdv3Text.SecOutput;
        outGroup.Dock = DockStyle.Top;
        // the caption and the frame are the box's own business: let it add them
        // to the rows rather than working the total out here and coming up a
        // pixel short, which drew the last row over the bottom of the frame
        outGroup.AutoSize = true;
        outGroup.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        outGroup.Margin = new Padding(0);
        outGroup.Padding = new Padding(Rdv3Metrics.Gap);
        outGroup.Controls.Add(BuildOutput());

        body.Controls.Add(inputGroup, 0, 0);
        body.Controls.Add(stepGroup, 0, 1);
        body.Controls.Add(outGroup, 0, 2);

        FlowLayoutPanel foot = new FlowLayoutPanel();
        foot.Name = "process.buttons";
        foot.Dock = DockStyle.Bottom;
        foot.FlowDirection = FlowDirection.RightToLeft;
        foot.Padding = new Padding(Rdv3Metrics.Gap);
        foot.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button cancel = ButtonOf("process.cancel", Rdv3Text.BtnCancel, DialogResult.Cancel);
        execute = ButtonOf("process.execute", j.Kind == "delete" ? Rdv3Text.BtnDelete : Rdv3Text.BtnExecute, DialogResult.OK);
        foot.Controls.Add(cancel);
        foot.Controls.Add(execute);
        AlignRightButtons(foot);
        AcceptButton = execute;
        CancelButton = cancel;

        Controls.Add(body);
        Controls.Add(foot);
        Controls.Add(hint);
        FillInputs();
        FillSteps();
        execute.Enabled = inputsOk;
        if (!inputsOk) { ToolTip tip = new ToolTip(); tip.SetToolTip(execute, Rdv3Text.ProcessNotRun); }
        // the dialog is exactly what its parts add up to; leftover height used
        // to fall into the last row of the output grid and drag its label down
        ClientSize = new Size(ClientSize.Width,
            Rdv3Metrics.Caption(Font) + hint.Padding.Vertical + body.Padding.Vertical
            + inputsRow + stepsRow + OutputHeight() + foot.Height);
        FinishLayout();
    }

    // how many rows of each list the dialog makes room for
    private const int InputListRows = 3;
    private const int StepListRows = 5;
    private const int OutputRows = 3;

    private int OutputHeight()
    {
        return Rdv3Metrics.Caption(Font) + Rdv3Metrics.Gap * 2
            + OutputRows * Rdv3Metrics.FieldRow(Font);
    }

    private Control BuildOutput()
    {
        TableLayoutPanel grid = new TableLayoutPanel();
        grid.Name = "process.output";
        grid.Dock = DockStyle.Top;
        grid.ColumnCount = 3;
        grid.RowCount = OutputRows + 1;
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 82f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        string dir = Path.GetDirectoryName(ledgerPath);
        string file = Path.GetFileName(ledgerPath);
        string updated = File.Exists(ledgerPath)
            ? File.GetLastWriteTime(ledgerPath).ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)
            : Rdv3Text.LblNeverWritten;
        string[] labels = { Rdv3Text.LblPath, Rdv3Text.LblFileName, Rdv3Text.LblLastWrite };
        string[] values = { dir, file, updated };
        int rowHeight = Rdv3Metrics.FieldRow(Font);
        for (int i = 0; i < OutputRows; i++)
        {
            grid.RowStyles.Add(new RowStyle(SizeType.Absolute, rowHeight));
            Label label = LabelOf("process.outputLabel" + i.ToString(CultureInfo.InvariantCulture), labels[i]);
            label.Dock = DockStyle.Fill;
            label.Margin = Padding.Empty;
            label.TextAlign = ContentAlignment.MiddleLeft;
            TextBox value = ReadOnlyBox("process.outputValue" + i.ToString(CultureInfo.InvariantCulture));
            value.Dock = DockStyle.Fill;
            value.Margin = new Padding(0, 1, 0, 1);
            value.Text = values[i];
            grid.Controls.Add(label, 0, i);
            grid.Controls.Add(value, 2, i);
        }
        // a table layout hands any room it has left over to its last row, and
        // that stretched the last label away from the box beside it
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        grid.Height = OutputRows * rowHeight;
        return grid;
    }

    private void FillInputs()
    {
        inputs.BeginUpdate();
        for (int i = 0; i < job.Inputs.Count; i++)
        {
            Rdv3ProcessInputDef input = job.Inputs[i];
            string path = Path.IsPathRooted(input.File) ? input.File : Path.Combine(dataDir, input.File);
            string rows = "";
            string validation;
            Color color;
            if (!File.Exists(path))
            {
                inputsOk = false;
                validation = Rdv3Text.ValidationMissing;
                color = Color.Maroon;
            }
            else
            {
                try
                {
                    Rdv3Table table = Rdv3Table.Read(path, input.Id, data.Enc, input.Column,
                        input.KeyValidation);
                    new Rdv3Index(table);
                    rows = table.Rows.ToString("N0", CultureInfo.InvariantCulture);
                    List<string> warnings = new List<string>();
                    if (table.InvalidEncodingRow > 0)
                    {
                        warnings.Add(Rdv3Text.ValidationEncodingMismatch.Replace("{row}", table.InvalidEncodingRow.ToString(CultureInfo.InvariantCulture)));
                    }
                    if (table.ControlCharacterWarning.Length > 0) { warnings.Add(table.ControlCharacterWarning); }
                    if (warnings.Count > 0)
                    {
                        validation = string.Join(" / ", warnings.ToArray());
                        color = Color.Maroon;
                    }
                    else
                    {
                        validation = Rdv3Text.ValidationColumnsMatch;
                        color = Color.DarkGreen;
                    }
                }
                catch
                {
                    inputsOk = false;
                    validation = Rdv3Text.ValidationError;
                    color = Color.Maroon;
                }
            }
            ListViewItem item = new ListViewItem(input.Id);
            item.SubItems.Add(input.File);
            item.SubItems.Add(input.Column);
            item.SubItems.Add(rows);
            item.SubItems.Add(validation);
            item.UseItemStyleForSubItems = false;
            item.SubItems[4].ForeColor = color;
            inputs.Items.Add(item);
        }
        if (inputs.Items.Count > 0) { inputs.Items[0].Selected = true; }
        inputs.EndUpdate();
    }

    private void FillSteps()
    {
        steps.BeginUpdate();
        for (int i = 0; i < job.Steps.Count; i++)
        {
            Rdv3ProcessStepDef step = job.Steps[i];
            ListViewItem item = new ListViewItem((i + 1).ToString(CultureInfo.InvariantCulture));
            item.SubItems.Add(Rdv3Text.OperationLabel(step.Operation));
            item.SubItems.Add(Display(step.Target1));
            item.SubItems.Add(Display(step.Target2));
            item.SubItems.Add(DisplayKey(step));
            item.SubItems.Add(DisplayCondition(step));
            item.SubItems.Add(Display(step.Output));
            steps.Items.Add(item);
        }
        steps.EndUpdate();
    }

    private string Display(string name)
    {
        if (name == null || name.Length == 0) { return ""; }
        string label = data.LabelOf(name);
        return (label.Length == 0) ? name : label;
    }

    private string DisplayKey(Rdv3ProcessStepDef step)
    {
        List<string> values = new List<string>();
        for (int i = 0; i < step.Keys.Length; i++) { values.Add(Display(step.Keys[i])); }
        if (values.Count > 0) { return string.Join(" = ", values.ToArray()); }
        if (step.Operation == "select" || step.Operation == "distinct")
        {
            for (int i = 0; i < step.Columns.Count; i++) { values.Add(Display(step.Columns[i].Column)); }
        }
        else if (step.Operation == "calculate" && step.Column.Length > 0)
        {
            values.Add(Display(step.Output + "." + step.Column));
        }
        else if (step.Operation == "aggregate")
        {
            for (int i = 0; i < step.GroupBy.Count; i++) { values.Add(Display(step.GroupBy[i])); }
        }
        else if (step.Operation == "sort")
        {
            for (int i = 0; i < step.Orders.Count; i++) { values.Add(Display(step.Orders[i].Column)); }
        }
        else if (step.Operation == "update")
        {
            for (int i = 0; i < step.Set.Count; i++) { values.Add(Display(step.Set[i].Column)); }
        }
        return (values.Count == 0) ? "" : string.Join(" / ", values.ToArray());
    }

    private string DisplayCondition(Rdv3ProcessStepDef step)
    {
        if (step.Operation == "merge")
        {
            return Rdv3Text.MergeDestinations(step.SourceOnly, step.Both, step.TargetOnly);
        }
        if (step.Operation == "join") { return Rdv3Text.JoinConditionLabel(step.Condition); }
        if (step.Where != null)
        {
            string result = Display(step.Where.Column) + " " + Rdv3Text.PredicateLabel(step.Where.Operator);
            if (step.Where.Operator != "empty" && step.Where.Operator != "notEmpty")
            {
                result += " " + step.Where.Value;
            }
            return result;
        }
        if (step.Operation == "select")
        {
            List<string> mappings = new List<string>();
            for (int i = 0; i < step.Columns.Count; i++)
            {
                if (step.Columns[i].As.Length > 0)
                {
                    mappings.Add(Display(step.Columns[i].Column) + " -> "
                        + Display(step.Columns[i].OutputRef(step.Output)));
                }
            }
            return (mappings.Count == 0) ? "" : string.Join(" / ", mappings.ToArray());
        }
        if (step.Operation == "calculate")
        {
            return Rdv3Process.DisplayExpression(data, step.Expression);
        }
        if (step.Operation == "aggregate")
        {
            List<string> aggregates = new List<string>();
            for (int i = 0; i < step.Aggregates.Count; i++)
            {
                string source = (step.Aggregates[i].Column.Length == 0)
                    ? "*" : Display(step.Aggregates[i].Column);
                aggregates.Add(Rdv3Text.AggregateLabel(step.Aggregates[i].Function) + "(" + source + ") -> "
                    + Display(step.Output + "." + step.Aggregates[i].As));
            }
            return string.Join(" / ", aggregates.ToArray());
        }
        if (step.Operation == "sort")
        {
            List<string> orders = new List<string>();
            for (int i = 0; i < step.Orders.Count; i++)
            {
                orders.Add(Display(step.Orders[i].Column) + " " + Rdv3Text.DirectionLabel(step.Orders[i].Direction)
                    + " (" + Rdv3Text.SortTypeLabel(step.Orders[i].Type) + ")");
            }
            return string.Join(" / ", orders.ToArray());
        }
        if (step.Operation == "update")
        {
            List<string> updates = new List<string>();
            for (int i = 0; i < step.Set.Count; i++)
            {
                updates.Add(Display(step.Set[i].Column) + " = "
                    + Rdv3Process.DisplayExpression(data, step.Set[i].Expression));
            }
            return string.Join(" / ", updates.ToArray());
        }
        return Rdv3Text.ConditionLabel(step.Condition);
    }
}

public sealed class Rdv3ExportFilter
{
    public string Field = "";
    public string Operator = "";
    public string First = "";
    public string Last = "";

    public bool Matches(Rdv3Data data, string[] values, string storedState)
    {
        string value;
        if (Field == "$work") { value = storedState ?? ""; }
        else
        {
            int column = data.IndexOf(Field);
            if (column < 0 || values == null || column >= values.Length) { return false; }
            value = values[column] ?? "";
        }
        Rdv3ColumnTypeDef type = data.TypeOf(Field);
        if (type == null)
        {
            if (Operator == "contains") { return value.IndexOf(First, StringComparison.Ordinal) >= 0; }
            if (Operator == "equals") { return string.Equals(value, First, StringComparison.Ordinal); }
            if (Operator == "startsWith") { return value.StartsWith(First, StringComparison.Ordinal); }
            if (Operator == "notContains") { return value.IndexOf(First, StringComparison.Ordinal) < 0; }
            throw new InvalidOperationException("invalid text export filter operator: " + Operator);
        }
        if (Operator != "range") { throw new InvalidOperationException("invalid typed export filter operator: " + Operator); }
        if (type.Type == "date")
        {
            DateTime actual, first, last;
            if (!type.TryDate(value, out actual) || !type.TryDate(First, out first) || !type.TryDate(Last, out last))
            {
                return false;
            }
            return actual >= first && actual <= last;
        }
        decimal actualNumber, firstNumber, lastNumber;
        if (!type.TryNumber(value, out actualNumber) || !type.TryNumber(First, out firstNumber)
            || !type.TryNumber(Last, out lastNumber)) { return false; }
        return actualNumber >= firstNumber && actualNumber <= lastNumber;
    }
}

public sealed class Rdv3ExportRequest
{
    public string Path = "";
    public List<string> Fields = new List<string>();
    public List<Rdv3ExportFilter> Filters = new List<Rdv3ExportFilter>();

    public bool Matches(Rdv3Data data, string[] values, string storedState)
    {
        for (int i = 0; i < Filters.Count; i++)
        {
            if (!Filters[i].Matches(data, values, storedState)) { return false; }
        }
        return true;
    }
}

public sealed class Rdv3ExportForm : Rdv3Dialog
{
    private sealed class ExportField
    {
        public string Ref;
        public string Text;
        public Rdv3ColumnTypeDef Type;
        public override string ToString() { return Text; }
    }

    private sealed class FilterOperator
    {
        public string Code;
        public string Text;
        public override string ToString() { return Text; }
    }

    private readonly Rdv3Data data;
    private readonly Rdv3Screen screen;
    private readonly ListBox from = new ListBox();
    private readonly ListBox to = new ListBox();
    private readonly Label selectedTitle;
    private readonly TextBox destination = new TextBox();
    private readonly List<ExportField> all = new List<ExportField>();
    private readonly List<string> defaults = new List<string>();
    private readonly ComboBox filterField = new ComboBox();
    private readonly ComboBox filterOperator = new ComboBox();
    private readonly TextBox filterFirstText = new TextBox();
    private readonly TextBox filterLastText = new TextBox();
    private readonly DateTimePicker filterFirstDate = new DateTimePicker();
    private readonly DateTimePicker filterLastDate = new DateTimePicker();
    private readonly Label filterRangeMark;
    private readonly ListView filterList;
    private readonly List<Rdv3ExportFilter> filters = new List<Rdv3ExportFilter>();
    public Rdv3ExportRequest Result;

    public static Rdv3ExportRequest Pick(Rdv3Form owner, Rdv3Data data, Rdv3Screen screen, string baseDir)
    {
        using (Rdv3ExportForm f = new Rdv3ExportForm(owner, data, screen, baseDir))
        {
            return (f.ShowOver(owner) == DialogResult.OK) ? f.Result : null;
        }
    }

    public static Rdv3ExportForm ForCheck(Rdv3Data data, Rdv3Screen screen, string baseDir)
    {
        return new Rdv3ExportForm(null, data, screen, baseDir);
    }

    private Rdv3ExportForm(Rdv3Form owner, Rdv3Data d, Rdv3Screen s, string baseDir)
        : base(Rdv3Text.ExportTitle, 600, 330, owner)
    {
        data = d;
        screen = s;
        Label hint = LabelOf("export.hint", Rdv3Text.ExportHint);
        hint.Dock = DockStyle.Top;
        hint.AutoSize = true;
        hint.Padding = new Padding(Rdv3Metrics.Gap, Rdv3Metrics.Gap, Rdv3Metrics.Gap, 0);

        TableLayoutPanel picker = new TableLayoutPanel();
        picker.Name = "export.picker";
        picker.Dock = DockStyle.Top;
        picker.Padding = new Padding(Rdv3Metrics.Gap);
        // the heading row carries a button, so it has to be as tall as one:
        // pinned at 27 it cut the bottom off and left the right edge standing
        // on its own beside the heading
        int headRow = Rdv3Metrics.ButtonRow(Font);
        int listRows = 10;
        picker.Height = picker.Padding.Vertical + headRow + listRows * Rdv3Metrics.ListRow(Font);
        picker.ColumnCount = 3;
        picker.RowCount = 2;
        picker.RowStyles.Add(new RowStyle(SizeType.Absolute, headRow));
        picker.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        Label available = LabelOf("export.availableTitle", Rdv3Text.ExportAvailable);
        available.Dock = DockStyle.Fill;
        available.Margin = Padding.Empty;
        available.TextAlign = ContentAlignment.MiddleLeft;

        TableLayoutPanel rightHead = new TableLayoutPanel();
        rightHead.Name = "export.selectedHeader";
        rightHead.Dock = DockStyle.Fill;
        rightHead.Margin = Padding.Empty;
        rightHead.ColumnCount = 2;
        rightHead.RowCount = 1;
        rightHead.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        rightHead.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        rightHead.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        selectedTitle = LabelOf("export.selectedTitle", "");
        selectedTitle.Dock = DockStyle.Fill;
        selectedTitle.Margin = Padding.Empty;
        selectedTitle.TextAlign = ContentAlignment.MiddleLeft;
        Button reset = ButtonOf("export.reset", Rdv3Text.ExportDefault, DialogResult.None);
        reset.Margin = new Padding(Rdv3Metrics.Gap, 0, 0, 0);
        StretchAcross(reset);
        reset.Click += delegate { ResetFields(); };
        rightHead.Controls.Add(selectedTitle, 0, 0);
        rightHead.Controls.Add(reset, 1, 0);

        from.Name = "export.available";
        from.Dock = DockStyle.Fill;
        from.Margin = Padding.Empty;
        from.BackColor = SystemColors.Window;
        from.IntegralHeight = false;
        from.DoubleClick += delegate { MoveItem(from, to); };
        to.Name = "export.selected";
        to.Dock = DockStyle.Fill;
        to.Margin = Padding.Empty;
        to.BackColor = SystemColors.Window;
        to.IntegralHeight = false;
        to.DoubleClick += delegate { MoveItem(to, from); };

        TableLayoutPanel arrows = new TableLayoutPanel();
        arrows.Name = "export.arrows";
        arrows.Dock = DockStyle.Fill;
        arrows.Padding = new Padding(Rdv3Metrics.Gap);
        arrows.ColumnCount = 1;
        arrows.RowCount = 5;
        arrows.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        arrows.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
        arrows.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        arrows.RowStyles.Add(new RowStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        arrows.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        arrows.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
        Button right = ButtonOf("export.right", Rdv3Text.BtnMoveRight, DialogResult.None);
        Button left = ButtonOf("export.left", Rdv3Text.BtnMoveLeft, DialogResult.None);
        right.AutoSize = false;
        right.MinimumSize = Size.Empty;
        left.AutoSize = false;
        left.MinimumSize = Size.Empty;
        right.Size = new Size(34, 25);
        left.Size = new Size(34, 25);
        right.Anchor = AnchorStyles.None;
        left.Anchor = AnchorStyles.None;
        right.Margin = Padding.Empty;
        left.Margin = Padding.Empty;
        right.Click += delegate { MoveItem(from, to); };
        left.Click += delegate { MoveItem(to, from); };
        arrows.Controls.Add(right, 0, 1);
        arrows.Controls.Add(left, 0, 3);
        int arrowColumnWidth = Math.Max(right.Width, left.Width)
            + arrows.Padding.Horizontal + arrows.Margin.Horizontal;

        picker.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        picker.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, arrowColumnWidth));
        picker.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

        picker.Controls.Add(available, 0, 0);
        picker.Controls.Add(rightHead, 2, 0);
        picker.Controls.Add(from, 0, 1);
        picker.Controls.Add(arrows, 1, 1);
        picker.Controls.Add(to, 2, 1);

        GroupBox filterGroup = new GroupBox();
        filterGroup.Name = "export.filterGroup";
        filterGroup.Text = Rdv3Text.ExportFilterGroup;
        filterGroup.Dock = DockStyle.Top;
        filterGroup.Padding = new Padding(Rdv3Metrics.Gap);

        TableLayoutPanel filterGrid = new TableLayoutPanel();
        filterGrid.Name = "export.filterGrid";
        filterGrid.Dock = DockStyle.Fill;
        filterGrid.Padding = Padding.Empty;
        filterGrid.ColumnCount = 6;
        filterGrid.RowCount = 3;
        Button addFilter = ButtonOf("export.filter.add", Rdv3Text.ExportFilterAdd, DialogResult.None);
        StretchAcross(addFilter);
        addFilter.Margin = new Padding(Rdv3Metrics.Gap, 0, 0, 0);
        addFilter.Click += delegate { AddFilter(); };
        Button removeFilter = ButtonOf("export.filter.remove", Rdv3Text.ExportFilterRemove, DialogResult.None);
        removeFilter.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
        removeFilter.Margin = new Padding(Rdv3Metrics.Gap, Rdv3Metrics.Gap, 0, 0);
        removeFilter.Click += delegate { RemoveFilter(); };
        int filterActionWidth = Math.Max(
            addFilter.GetPreferredSize(Size.Empty).Width + addFilter.Margin.Horizontal,
            removeFilter.GetPreferredSize(Size.Empty).Width + removeFilter.Margin.Horizontal);
        filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 155f));
        filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110f));
        filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 24f));
        filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        filterGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, filterActionWidth));
        int filterHead = Rdv3Metrics.TextHeight(Font) + 4;
        int filterInput = Math.Max(Rdv3Metrics.FieldRow(Font), Rdv3Metrics.ButtonHeight(Font)) + 3;
        int filterRows = 3;
        int filterListHeight = Rdv3Metrics.ListHeight(Font, filterRows);
        filterGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, filterHead));
        filterGrid.RowStyles.Add(new RowStyle(SizeType.Absolute, filterInput));
        filterGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        filterGroup.Height = Rdv3Metrics.Caption(Font) + filterGroup.Padding.Vertical
            + filterGrid.Padding.Vertical + filterHead + filterInput + filterListHeight;

        Label fieldHead = LabelOf("export.filter.fieldLabel", Rdv3Text.ExportFilterField);
        Label operatorHead = LabelOf("export.filter.operatorLabel", Rdv3Text.ExportFilterCondition);
        Label valueHead = LabelOf("export.filter.valueLabel", Rdv3Text.ExportFilterValue);
        fieldHead.Dock = DockStyle.Fill;
        operatorHead.Dock = DockStyle.Fill;
        valueHead.Dock = DockStyle.Fill;
        fieldHead.Margin = new Padding(0, 0, Rdv3Metrics.Gap, 0);
        operatorHead.Margin = new Padding(0, 0, Rdv3Metrics.Gap, 0);
        valueHead.Margin = Padding.Empty;
        fieldHead.TextAlign = ContentAlignment.BottomLeft;
        operatorHead.TextAlign = ContentAlignment.BottomLeft;
        valueHead.TextAlign = ContentAlignment.BottomLeft;
        filterGrid.Controls.Add(fieldHead, 0, 0);
        filterGrid.Controls.Add(operatorHead, 1, 0);
        filterGrid.Controls.Add(valueHead, 2, 0);
        filterGrid.SetColumnSpan(valueHead, 3);

        filterField.Name = "export.filter.field";
        filterField.DropDownStyle = ComboBoxStyle.DropDownList;
        filterField.Dock = DockStyle.Fill;
        filterField.Margin = new Padding(0, 1, Rdv3Metrics.Gap, 1);
        filterField.SelectedIndexChanged += delegate { UpdateFilterEditor(); };
        filterOperator.Name = "export.filter.operator";
        filterOperator.DropDownStyle = ComboBoxStyle.DropDownList;
        filterOperator.Dock = DockStyle.Fill;
        filterOperator.Margin = new Padding(0, 1, Rdv3Metrics.Gap, 1);

        Panel firstHost = FilterValueHost("export.filter.firstHost", filterFirstText, filterFirstDate);
        Panel lastHost = FilterValueHost("export.filter.lastHost", filterLastText, filterLastDate);
        filterFirstText.Name = "export.filter.value1";
        filterLastText.Name = "export.filter.value2";
        filterFirstDate.Name = "export.filter.date1";
        filterLastDate.Name = "export.filter.date2";
        filterRangeMark = LabelOf("export.filter.rangeMark", Rdv3Text.ExportFilterRangeMark);
        // Let the glyph keep its native preferred height, then let the table
        // cell centre it. A Fill-docked Label can retain the row's clipped GDI
        // text rectangle after DPI scaling even though its outer bounds fit.
        filterRangeMark.AutoEllipsis = false;
        filterRangeMark.AutoSize = true;
        filterRangeMark.Dock = DockStyle.None;
        filterRangeMark.Anchor = AnchorStyles.None;
        filterRangeMark.Margin = Padding.Empty;
        filterRangeMark.TextAlign = ContentAlignment.MiddleCenter;
        filterList = DetailsList("export.filter.list");
        filterList.Dock = DockStyle.Fill;
        filterList.Margin = new Padding(0, Rdv3Metrics.Gap, 0, 0);
        filterList.Columns.Add(Rdv3Text.ExportFilterField, 180);
        filterList.Columns.Add(Rdv3Text.ExportFilterCondition, 110);
        filterList.Columns.Add(Rdv3Text.ExportFilterValue, 220);
        KeepColumnsFitted(filterList, new int[] { 180, 110, 220 });
        filterGrid.Controls.Add(filterField, 0, 1);
        filterGrid.Controls.Add(filterOperator, 1, 1);
        filterGrid.Controls.Add(firstHost, 2, 1);
        filterGrid.Controls.Add(filterRangeMark, 3, 1);
        filterGrid.Controls.Add(lastHost, 4, 1);
        filterGrid.Controls.Add(addFilter, 5, 1);
        filterGrid.Controls.Add(filterList, 0, 2);
        filterGrid.SetColumnSpan(filterList, 5);
        filterGrid.Controls.Add(removeFilter, 5, 2);
        filterGroup.Controls.Add(filterGrid);

        TableLayoutPanel path = new TableLayoutPanel();
        path.Name = "export.destinationRow";
        path.Dock = DockStyle.Top;
        path.Padding = new Padding(Rdv3Metrics.Gap);
        // the row carries a button as well as a box, and a button will not go
        // below the size its own caption needs
        path.Height = path.Padding.Vertical
            + Math.Max(Rdv3Metrics.FieldRow(Font), Rdv3Metrics.ButtonHeight(Font));
        path.ColumnCount = 5;
        path.RowCount = 1;
        Label pl = LabelOf("export.destinationLabel", Rdv3Text.ExportDestination);
        pl.Dock = DockStyle.Fill;
        pl.Margin = Padding.Empty;
        pl.TextAlign = ContentAlignment.MiddleLeft;
        destination.Name = "export.destination";
        StretchAcross(destination);
        destination.Margin = new Padding(0, 1, 0, 1);
        destination.BackColor = SystemColors.Window;
        destination.BorderStyle = BorderStyle.Fixed3D;
        destination.Text = Path.Combine(baseDir, Rdv3Text.ExportDefaultPath.Replace("{yyyyMMdd-HHmmss}", DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)));
        Button browse = ButtonOf("export.browse", Rdv3Text.BtnBrowse, DialogResult.None);
        StretchAcross(browse);
        browse.Margin = Padding.Empty;
        browse.Click += delegate { Browse(); };
        int browseColumnWidth = browse.GetPreferredSize(Size.Empty).Width + browse.Margin.Horizontal;
        path.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 62f));
        path.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        path.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        path.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Rdv3Metrics.Gap));
        path.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, browseColumnWidth));
        path.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        path.Controls.Add(pl, 0, 0);
        path.Controls.Add(destination, 2, 0);
        path.Controls.Add(browse, 4, 0);

        FlowLayoutPanel foot = new FlowLayoutPanel();
        foot.Name = "export.buttons";
        foot.Dock = DockStyle.Bottom;
        foot.FlowDirection = FlowDirection.RightToLeft;
        foot.Padding = new Padding(Rdv3Metrics.Gap);
        foot.Height = Rdv3Metrics.ButtonBarHeight(Font);
        Button cancel = ButtonOf("export.cancel", Rdv3Text.BtnCancel, DialogResult.Cancel);
        Button ok = ButtonOf("export.ok", Rdv3Text.BtnOk, DialogResult.None);
        ok.Click += delegate { AcceptExport(); };
        foot.Controls.Add(cancel);
        foot.Controls.Add(ok);
        AlignRightButtons(foot);
        AcceptButton = ok;
        CancelButton = cancel;

        ClientSize = new Size(ClientSize.Width,
            Rdv3Metrics.Caption(Font) + hint.Padding.Vertical
            + picker.Height + filterGroup.Height + path.Height + foot.Height);
        Controls.Add(path);
        Controls.Add(filterGroup);
        Controls.Add(picker);
        Controls.Add(foot);
        Controls.Add(hint);
        BuildFields();
        ResetFields();
        BuildFilterFields();
        FinishLayout();
    }

    private static Panel FilterValueHost(string name, TextBox text, DateTimePicker date)
    {
        Panel host = new Panel();
        host.Name = name;
        host.Dock = DockStyle.Fill;
        host.Margin = new Padding(0, 1, 0, 1);
        text.Dock = DockStyle.Fill;
        text.BackColor = SystemColors.Window;
        text.BorderStyle = BorderStyle.Fixed3D;
        date.Dock = DockStyle.Fill;
        date.Format = DateTimePickerFormat.Custom;
        host.Controls.Add(text);
        host.Controls.Add(date);
        return host;
    }

    private void BuildFields()
    {
        for (int i = 0; i < data.LabelOrder.Count; i++)
        {
            string reference = data.LabelOrder[i];
            int column = data.IndexOf(reference);
            if (column < 0) { continue; }
            Rdv3ColumnRef col = data.Columns[column];
            string label = data.LabelOf(reference);
            ExportField f = new ExportField();
            f.Ref = col.Ref;
            f.Text = label + " (" + col.Ref + ")";
            f.Type = data.TypeOf(col.Ref);
            all.Add(f);
        }
        ExportField work = new ExportField();
        work.Ref = "$work";
        Rdv3StateDef workState = screen.Work.InitialTargetState;
        if (workState == null) { workState = screen.Work.InitialState; }
        work.Text = (workState == null) ? screen.Work.Column : workState.Text;
        all.Add(work);
        for (int i = 0; i < screen.ExportDefaultFields.Length; i++) { defaults.Add(screen.ExportDefaultFields[i]); }
    }

    private void BuildFilterFields()
    {
        filterField.Items.Clear();
        for (int i = 0; i < all.Count; i++) { filterField.Items.Add(all[i]); }
        if (filterField.Items.Count > 0) { filterField.SelectedIndex = 0; }
    }

    private void UpdateFilterEditor()
    {
        ExportField field = filterField.SelectedItem as ExportField;
        bool typed = field != null && field.Type != null;
        bool date = typed && field.Type.Type == "date";
        filterOperator.Items.Clear();
        if (typed)
        {
            filterOperator.Items.Add(FilterOp("range", Rdv3Text.ExportFilterRange));
        }
        else
        {
            filterOperator.Items.Add(FilterOp("contains", Rdv3Text.ExportFilterContains));
            filterOperator.Items.Add(FilterOp("equals", Rdv3Text.ExportFilterEquals));
            filterOperator.Items.Add(FilterOp("startsWith", Rdv3Text.ExportFilterStarts));
            filterOperator.Items.Add(FilterOp("notContains", Rdv3Text.ExportFilterNotContains));
        }
        if (filterOperator.Items.Count > 0) { filterOperator.SelectedIndex = 0; }
        filterFirstText.Visible = !date;
        filterLastText.Visible = typed && !date;
        filterFirstDate.Visible = date;
        filterLastDate.Visible = date;
        filterRangeMark.Visible = typed;
        if (date)
        {
            filterFirstDate.CustomFormat = field.Type.Format;
            filterLastDate.CustomFormat = field.Type.Format;
        }
    }

    private static FilterOperator FilterOp(string code, string label)
    {
        FilterOperator op = new FilterOperator();
        op.Code = code;
        op.Text = label;
        return op;
    }

    private void AddFilter()
    {
        ExportField field = filterField.SelectedItem as ExportField;
        FilterOperator op = filterOperator.SelectedItem as FilterOperator;
        if (field == null || op == null) { return; }
        Rdv3ExportFilter filter = new Rdv3ExportFilter();
        filter.Field = field.Ref;
        filter.Operator = op.Code;
        if (field.Type == null)
        {
            filter.First = filterFirstText.Text;
            if (filter.First.Trim().Length == 0)
            {
                MessageBox.Show(this, Rdv3Text.ExportFilterNeedValue, Rdv3Text.ExportTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }
        else if (field.Type.Type == "date")
        {
            if (filterFirstDate.Value.Date > filterLastDate.Value.Date)
            {
                ShowFilterOrderError();
                return;
            }
            filter.First = filterFirstDate.Value.ToString(field.Type.Format, CultureInfo.InvariantCulture);
            filter.Last = filterLastDate.Value.ToString(field.Type.Format, CultureInfo.InvariantCulture);
        }
        else
        {
            decimal first, last;
            if (!field.Type.TryNumber(filterFirstText.Text, out first)
                || !field.Type.TryNumber(filterLastText.Text, out last))
            {
                MessageBox.Show(this, Rdv3Text.ExportFilterNeedNumber, Rdv3Text.ExportTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (first > last)
            {
                ShowFilterOrderError();
                return;
            }
            filter.First = first.ToString(CultureInfo.InvariantCulture);
            filter.Last = last.ToString(CultureInfo.InvariantCulture);
        }
        filters.Add(filter);
        ListViewItem item = new ListViewItem(field.Text);
        item.SubItems.Add(op.Text);
        item.SubItems.Add(filter.Last.Length == 0 ? filter.First
            : filter.First + " " + Rdv3Text.ExportFilterRangeMark + " " + filter.Last);
        item.Tag = filter;
        filterList.Items.Add(item);
        // Keep the item just added as the explicit removal target. This is
        // useful with a keyboard as well as after clicking the Add button.
        item.Selected = true;
        item.Focused = true;
        item.EnsureVisible();
    }

    private void ShowFilterOrderError()
    {
        MessageBox.Show(this, Rdv3Text.ExportFilterOrder, Rdv3Text.ExportTitle,
            MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    private void RemoveFilter()
    {
        ListViewItem item = (filterList.SelectedItems.Count > 0) ? filterList.SelectedItems[0]
            : (filterList.Items.Count == 1) ? filterList.Items[0] : null;
        if (item == null) { return; }
        Rdv3ExportFilter filter = item.Tag as Rdv3ExportFilter;
        if (filter != null) { filters.Remove(filter); }
        filterList.Items.Remove(item);
    }

    private void ResetFields()
    {
        from.Items.Clear();
        to.Items.Clear();
        for (int i = 0; i < defaults.Count; i++)
        {
            ExportField selected = FieldOf(defaults[i]);
            if (selected != null) { to.Items.Add(selected); }
        }
        for (int i = 0; i < all.Count; i++)
        {
            if (!defaults.Contains(all[i].Ref)) { from.Items.Add(all[i]); }
        }
        UpdateSelectedTitle();
    }

    private ExportField FieldOf(string reference)
    {
        for (int i = 0; i < all.Count; i++) { if (all[i].Ref == reference) { return all[i]; } }
        return null;
    }

    private void MoveItem(ListBox source, ListBox target)
    {
        int i = source.SelectedIndex;
        if (i < 0) { return; }
        object item = source.Items[i];
        source.Items.RemoveAt(i);
        target.Items.Add(item);
        if (source.Items.Count > 0) { source.SelectedIndex = Math.Min(i, source.Items.Count - 1); }
        UpdateSelectedTitle();
    }

    private void UpdateSelectedTitle()
    {
        selectedTitle.Text = Rdv3Text.ExportSelectedFmt.Replace("{n}", to.Items.Count.ToString(CultureInfo.InvariantCulture));
    }

    private void Browse()
    {
        using (SaveFileDialog dialog = new SaveFileDialog())
        {
            dialog.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.*";
            dialog.AddExtension = true;
            dialog.DefaultExt = "csv";
            try
            {
                dialog.InitialDirectory = Path.GetDirectoryName(destination.Text);
                dialog.FileName = Path.GetFileName(destination.Text);
            }
            catch { }
            if (dialog.ShowDialog(this) == DialogResult.OK) { destination.Text = dialog.FileName; }
        }
    }

    private void AcceptExport()
    {
        if (to.Items.Count == 0)
        {
            MessageBox.Show(this, Rdv3Text.ExportNeedField, Rdv3Text.ExportTitle, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        Result = new Rdv3ExportRequest();
        Result.Path = destination.Text.Trim();
        for (int i = 0; i < to.Items.Count; i++) { Result.Fields.Add(((ExportField)to.Items[i]).Ref); }
        for (int i = 0; i < filters.Count; i++) { Result.Filters.Add(filters[i]); }
        DialogResult = DialogResult.OK;
        Close();
    }
}
