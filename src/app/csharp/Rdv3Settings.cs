// ============================================================================
// Rdv3Settings.cs -- the settings dialog and the UI Automation inspector.
//
// Both are modal windows dressed in the same skin as the main screen
// (Rdv3Skin: the same background, hairline panels, section headings, muted
// labels and buttons), so the dialog looks like part of the application and
// not like a WinForms property sheet.
//
// The dialog edits a COPY of the settings. Nothing changes until the operator
// presses save; then the file is written and the running app adopts what can
// be adopted without a restart (see Rdv3Config.AdoptRuntimeFrom).
//
// The inspector is the part that makes naming a field in a real business
// application practical: it walks the live UI Automation tree, shows what each
// element exposes (ControlType, AutomationId, ClassName, Name, patterns,
// process) and turns the element the operator picked into a target -- window,
// the path down to it, the field itself, and how its value can be read.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text;
using System.Windows.Automation;
using System.Windows.Forms;

// ---------------------------------------------------------------------------
// shared bits of chrome, so both dialogs look like the main screen
// ---------------------------------------------------------------------------
internal static class Rdv3Dlg
{
    public static int P(double css) { return Rdv3Skin.P(css); }

    public static void Skin(Rdv3Dialog f, int wCss, int hCss)
    {
        f.FormBorderStyle = FormBorderStyle.FixedDialog;
        f.MaximizeBox = false;
        f.MinimizeBox = false;
        f.ShowInTaskbar = false;
        f.StartPosition = FormStartPosition.CenterParent;
        f.BackColor = Rdv3Skin.Bg;
        f.AutoScaleMode = AutoScaleMode.None;
        f.Font = Rdv3Skin.F(13, FontStyle.Regular);
        f.ClientSize = new Size(P(wCss), P(hCss));
    }

    public static TextBox Box(Control parent, int x, int y, int w)
    {
        TextBox t = new TextBox();
        t.BorderStyle = BorderStyle.None;
        t.BackColor = Rdv3Skin.Surface;
        t.ForeColor = Rdv3Skin.Ink;
        t.Font = Rdv3Skin.F(13, FontStyle.Regular);
        int h = t.PreferredHeight;
        t.SetBounds(x + P(6), y + (P(26) - h) / 2, w - P(12), h);
        parent.Controls.Add(t);
        return t;
    }

    public static CheckBox Check(Control parent, int x, int y, int w, string text)
    {
        CheckBox c = new CheckBox();
        c.Text = text;
        c.FlatStyle = FlatStyle.Flat;
        c.BackColor = Rdv3Skin.Bg;
        c.ForeColor = Rdv3Skin.Ink;
        c.Font = Rdv3Skin.F(12, FontStyle.Regular);
        c.FlatAppearance.BorderSize = 0;
        c.SetBounds(x, y, w, P(26));
        parent.Controls.Add(c);
        return c;
    }

    public static ComboBox Combo(Control parent, int x, int y, int w, string[] items)
    {
        ComboBox c = new ComboBox();
        c.DropDownStyle = ComboBoxStyle.DropDownList;
        c.FlatStyle = FlatStyle.Flat;
        c.BackColor = Rdv3Skin.Surface;
        c.ForeColor = Rdv3Skin.Ink;
        c.Font = Rdv3Skin.F(13, FontStyle.Regular);
        c.Items.AddRange(items);
        c.SetBounds(x, y, w, P(26));
        parent.Controls.Add(c);
        return c;
    }

    public static Rdv3Btn Button(Control parent, int x, int y, int w, string text, bool primary, int icon)
    {
        Rdv3Btn b = new Rdv3Btn();
        b.Text = text;
        b.Primary = primary;
        b.Icon = icon;
        b.Font = Rdv3Skin.S(14);
        b.SetBounds(x, y, w, P(31.7));
        parent.Controls.Add(b);
        return b;
    }

    // a field caption above its box, in the main screen's muted tone
    public static void Caption(Graphics g, string s, int x, int y, int w)
    {
        Rdv3Skin.DrawIn(g, s, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(x, y, w, P(15)), false);
    }

    public static void Section(Graphics g, string title, Rectangle r)
    {
        Rdv3Skin.Blueprint(g, r);
        Rdv3Skin.DrawIn(g, title, Rdv3Skin.S(18), Rdv3Skin.Ink,
            new Rectangle(r.X + P(13.6), r.Y + P(8), r.Width - P(27.2), P(22)), false);
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y + P(40), r.Width, Rdv3Skin.Hair());
    }

    public static void BoxFrame(Graphics g, Rectangle r, bool focused)
    {
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Surface)) { g.FillRectangle(b, r); }
        Rdv3Skin.Frame(g, focused ? Rdv3Skin.Accent : Rdv3Skin.Divider, r);
    }
}

// a modal that paints itself the way the main screen does
public abstract class Rdv3Dialog : Form
{
    protected Rdv3Dialog()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
    }
}

// ---------------------------------------------------------------------------
// the settings dialog
// ---------------------------------------------------------------------------
public sealed class Rdv3SettingsForm : Rdv3Dialog
{
    private readonly Rdv3Config cfg;
    private readonly ListBox lstTargets = new ListBox();
    private readonly ListBox lstPath = new ListBox();

    private TextBox tName, tWinId, tWinClass, tWinLike, tWinProc;
    private TextBox tFldId, tFldClass, tFldTypes, tFldIndex;
    private TextBox tKeyLen, tPoll, tStable, tRebind, tCand;
    private TextBox tData, tLedger, tLog;
    private CheckBox cEnabled, cDigits, cFocus, cFieldValue, cFieldChildren;
    private ComboBox cbRead;
    private bool loading;

    public Rdv3Config Result;

    // rectangles the painter and the layout share
    private Rectangle rTargets, rTarget, rBehaviour, rPaths, rFoot;

    public static Rdv3Config Edit(IWin32Window owner, Rdv3Config current)
    {
        Rdv3SettingsForm f = new Rdv3SettingsForm(current.Clone());
        DialogResult dr = f.ShowDialog(owner);
        Rdv3Config outp = (dr == DialogResult.OK) ? f.Result : null;
        f.Dispose();
        return outp;
    }

    private Rdv3SettingsForm(Rdv3Config working)
    {
        cfg = working;
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.SettingsTitle;
        Rdv3Dlg.Skin(this, 980, 780);
        Build();
        LoadAll();
    }

    private void Build()
    {
        int pad = Rdv3Dlg.P(13.6);
        int colW = Rdv3Dlg.P(300);
        int right = ClientSize.Width - pad;
        rTargets = new Rectangle(pad, pad, colW, Rdv3Dlg.P(300));
        rTarget = new Rectangle(pad + colW + pad, pad, right - (pad + colW + pad), Rdv3Dlg.P(470));
        rBehaviour = new Rectangle(pad, rTargets.Bottom + pad, colW, Rdv3Dlg.P(266));
        rPaths = new Rectangle(rTarget.X, rTarget.Bottom + pad, rTarget.Width, Rdv3Dlg.P(226));
        rFoot = new Rectangle(0, ClientSize.Height - Rdv3Dlg.P(48), ClientSize.Width, Rdv3Dlg.P(48));

        // ---- targets --------------------------------------------------------
        lstTargets.SetBounds(rTargets.X + Rdv3Dlg.P(10), rTargets.Y + Rdv3Dlg.P(48),
            rTargets.Width - Rdv3Dlg.P(20), Rdv3Dlg.P(190));
        SkinList(lstTargets);
        lstTargets.SelectedIndexChanged += delegate { LoadTarget(); };
        Controls.Add(lstTargets);
        int by = lstTargets.Bottom + Rdv3Dlg.P(10);
        int bw = (lstTargets.Width - Rdv3Dlg.P(12)) / 3;
        Rdv3Dlg.Button(this, lstTargets.Left, by, bw, Rdv3Text.BtnAdd, false, 0).Click += delegate { AddTarget(); };
        Rdv3Dlg.Button(this, lstTargets.Left + bw + Rdv3Dlg.P(6), by, bw, Rdv3Text.BtnCopy, false, 0)
            .Click += delegate { CopyTarget(); };
        Rdv3Dlg.Button(this, lstTargets.Left + 2 * (bw + Rdv3Dlg.P(6)), by, bw, Rdv3Text.BtnRemove, false, 0)
            .Click += delegate { RemoveTarget(); };

        // ---- the selected target -------------------------------------------
        int x = rTarget.X + Rdv3Dlg.P(13.6);
        int w = rTarget.Width - Rdv3Dlg.P(27.2);
        int y = rTarget.Y + Rdv3Dlg.P(52);
        int half = (w - Rdv3Dlg.P(10)) / 2;
        tName = Rdv3Dlg.Box(this, x, y + Rdv3Dlg.P(15), half);
        cEnabled = Rdv3Dlg.Check(this, x + half + Rdv3Dlg.P(10), y + Rdv3Dlg.P(15), half, Rdv3Text.LblEnabled);
        y += Rdv3Dlg.P(48);
        tWinId = Rdv3Dlg.Box(this, x, y + Rdv3Dlg.P(15), half);
        tWinClass = Rdv3Dlg.Box(this, x + half + Rdv3Dlg.P(10), y + Rdv3Dlg.P(15), half);
        y += Rdv3Dlg.P(44);
        tWinLike = Rdv3Dlg.Box(this, x, y + Rdv3Dlg.P(15), half);
        tWinProc = Rdv3Dlg.Box(this, x + half + Rdv3Dlg.P(10), y + Rdv3Dlg.P(15), half);
        y += Rdv3Dlg.P(44);
        lstPath.SetBounds(x, y + Rdv3Dlg.P(22), w, Rdv3Dlg.P(52));
        SkinList(lstPath);
        Controls.Add(lstPath);
        Rdv3Dlg.Button(this, x + w - Rdv3Dlg.P(90), y - Rdv3Dlg.P(14), Rdv3Dlg.P(80),
            Rdv3Text.BtnRemove, false, 0).Click += delegate { RemoveStep(); };
        y += Rdv3Dlg.P(88);
        tFldId = Rdv3Dlg.Box(this, x, y + Rdv3Dlg.P(15), half);
        tFldClass = Rdv3Dlg.Box(this, x + half + Rdv3Dlg.P(10), y + Rdv3Dlg.P(15), half);
        y += Rdv3Dlg.P(44);
        tFldTypes = Rdv3Dlg.Box(this, x, y + Rdv3Dlg.P(15), half);
        tFldIndex = Rdv3Dlg.Box(this, x + half + Rdv3Dlg.P(10), y + Rdv3Dlg.P(15), Rdv3Dlg.P(60));
        cFieldValue = Rdv3Dlg.Check(this, x + half + Rdv3Dlg.P(80), y + Rdv3Dlg.P(12),
            half - Rdv3Dlg.P(80), "ValuePattern");
        y += Rdv3Dlg.P(44);
        cbRead = Rdv3Dlg.Combo(this, x, y + Rdv3Dlg.P(15), Rdv3Dlg.P(120),
            new string[] { "value", "text", "name" });
        cFieldChildren = Rdv3Dlg.Check(this, x + Rdv3Dlg.P(140), y + Rdv3Dlg.P(12),
            w - Rdv3Dlg.P(140), Rdv3Text.LblScopeChildren);

        // ---- behaviour -------------------------------------------------------
        int bx = rBehaviour.X + Rdv3Dlg.P(13.6);
        int bwid = rBehaviour.Width - Rdv3Dlg.P(27.2);
        int bh = (bwid - Rdv3Dlg.P(10)) / 2;
        int byy = rBehaviour.Y + Rdv3Dlg.P(52);
        tKeyLen = Rdv3Dlg.Box(this, bx, byy + Rdv3Dlg.P(15), bh);
        cDigits = Rdv3Dlg.Check(this, bx + bh + Rdv3Dlg.P(10), byy + Rdv3Dlg.P(12), bh, Rdv3Text.LblDigitsOnly);
        byy += Rdv3Dlg.P(44);
        tPoll = Rdv3Dlg.Box(this, bx, byy + Rdv3Dlg.P(15), bh);
        tStable = Rdv3Dlg.Box(this, bx + bh + Rdv3Dlg.P(10), byy + Rdv3Dlg.P(15), bh);
        byy += Rdv3Dlg.P(44);
        tRebind = Rdv3Dlg.Box(this, bx, byy + Rdv3Dlg.P(15), bh);
        tCand = Rdv3Dlg.Box(this, bx + bh + Rdv3Dlg.P(10), byy + Rdv3Dlg.P(15), bh);
        byy += Rdv3Dlg.P(44);
        cFocus = Rdv3Dlg.Check(this, bx, byy + Rdv3Dlg.P(10), bwid, Rdv3Text.LblPreferFocus);

        // ---- paths -----------------------------------------------------------
        int px = rPaths.X + Rdv3Dlg.P(13.6);
        int pw = rPaths.Width - Rdv3Dlg.P(27.2);
        int py = rPaths.Y + Rdv3Dlg.P(52);
        tData = Rdv3Dlg.Box(this, px, py + Rdv3Dlg.P(15), pw);
        py += Rdv3Dlg.P(48);
        tLedger = Rdv3Dlg.Box(this, px, py + Rdv3Dlg.P(15), pw);
        py += Rdv3Dlg.P(48);
        tLog = Rdv3Dlg.Box(this, px, py + Rdv3Dlg.P(15), pw);

        // ---- footer ----------------------------------------------------------
        int fw = Rdv3Dlg.P(150);
        int fy = rFoot.Y + Rdv3Dlg.P(8);
        Rdv3Dlg.Button(this, pad, fy, Rdv3Dlg.P(170), Rdv3Text.BtnInspect, false, 4)
            .Click += delegate { Inspect(); };
        Rdv3Btn ok = Rdv3Dlg.Button(this, ClientSize.Width - pad - fw * 2 - Rdv3Dlg.P(8), fy, fw,
            Rdv3Text.BtnSave, true, 2);
        ok.Click += delegate { SaveAndClose(); };
        Rdv3Btn no = Rdv3Dlg.Button(this, ClientSize.Width - pad - fw, fy, fw, Rdv3Text.BtnCancel, false, 0);
        no.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        CancelButton = no;
    }

    private static void SkinList(ListBox l)
    {
        l.BorderStyle = BorderStyle.None;
        l.BackColor = Rdv3Skin.Surface;
        l.ForeColor = Rdv3Skin.Ink;
        l.Font = Rdv3Skin.F(13, FontStyle.Regular);
        l.IntegralHeight = false;
        l.DrawMode = DrawMode.OwnerDrawFixed;
        l.ItemHeight = Rdv3Skin.P(24);
        l.DrawItem += delegate(object s, DrawItemEventArgs e)
        {
            ListBox box = (ListBox)s;
            if (e.Index < 0) { return; }
            bool sel = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            using (SolidBrush b = new SolidBrush(sel ? Rdv3Skin.Accent100 : Rdv3Skin.Surface))
            {
                e.Graphics.FillRectangle(b, e.Bounds);
            }
            if (sel) { Rdv3Skin.Line(e.Graphics, Rdv3Skin.Accent, e.Bounds.X, e.Bounds.Y, Rdv3Skin.P(2), e.Bounds.Height); }
            Rdv3Skin.DrawIn(e.Graphics, Convert.ToString(box.Items[e.Index]), box.Font, Rdv3Skin.Ink,
                new Rectangle(e.Bounds.X + Rdv3Skin.P(8), e.Bounds.Y, e.Bounds.Width - Rdv3Skin.P(12), e.Bounds.Height),
                false);
        };
    }

    // ---- painting ----------------------------------------------------------
    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        Rdv3Dlg.Section(g, Rdv3Text.SecTargets, rTargets);
        Rdv3Dlg.Section(g, Rdv3Text.SecTarget, rTarget);
        Rdv3Dlg.Section(g, Rdv3Text.SecBehaviour, rBehaviour);
        Rdv3Dlg.Section(g, Rdv3Text.SecPaths, rPaths);

        Cap(g, tName, Rdv3Text.LblName);
        Cap(g, tWinId, Rdv3Text.LblWindow + " / " + Rdv3Text.LblAutomationId);
        Cap(g, tWinClass, Rdv3Text.LblClassName);
        Cap(g, tWinLike, Rdv3Text.LblNameLike);
        Cap(g, tWinProc, Rdv3Text.LblProcess);
        Rdv3Dlg.Caption(g, Rdv3Text.LblPath, lstPath.Left, lstPath.Top - Rdv3Skin.P(18),
            lstPath.Width - Rdv3Skin.P(100));
        Cap(g, tFldId, Rdv3Text.LblField + " / " + Rdv3Text.LblAutomationId);
        Cap(g, tFldClass, Rdv3Text.LblClassName);
        Cap(g, tFldTypes, Rdv3Text.LblControlTypes);
        Cap(g, tFldIndex, Rdv3Text.LblIndex);
        Cap(g, cbRead, Rdv3Text.LblRead);
        Cap(g, tKeyLen, Rdv3Text.LblKeyLen);
        Cap(g, tPoll, Rdv3Text.LblPollMs);
        Cap(g, tStable, Rdv3Text.LblStableMs);
        Cap(g, tRebind, Rdv3Text.LblRebindMs);
        Cap(g, tCand, Rdv3Text.LblCandRows);
        Cap(g, tData, Rdv3Text.LblDataDir);
        Cap(g, tLedger, Rdv3Text.LblLedger);
        Cap(g, tLog, Rdv3Text.LblLog);

        Frame(g, tName); Frame(g, tWinId); Frame(g, tWinClass); Frame(g, tWinLike); Frame(g, tWinProc);
        Frame(g, tFldId); Frame(g, tFldClass); Frame(g, tFldTypes); Frame(g, tFldIndex);
        Frame(g, tKeyLen); Frame(g, tPoll); Frame(g, tStable); Frame(g, tRebind); Frame(g, tCand);
        Frame(g, tData); Frame(g, tLedger); Frame(g, tLog);
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, Grow(lstTargets));
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, Grow(lstPath));

        Rdv3Skin.DrawIn(g, Rdv3Text.NoteRestart, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(rPaths.X + Rdv3Skin.P(13.6), tLog.Bottom + Rdv3Skin.P(14),
                rPaths.Width - Rdv3Skin.P(27.2), Rdv3Skin.P(18)), false);

        using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent900)) { g.FillRectangle(b, rFoot); }
        string where = Rdv3Text.NoteSavedTo + " " + cfg.SourcePath;
        int wx = Rdv3Skin.P(13.6) + Rdv3Skin.P(170) + Rdv3Skin.P(14);
        Rdv3Skin.DrawIn(g, where, Rdv3Skin.F(12, FontStyle.Regular), Rdv3Skin.Bg,
            new Rectangle(wx, rFoot.Y, ClientSize.Width - wx - Rdv3Skin.P(340), rFoot.Height), false);
    }

    private void Cap(Graphics g, Control c, string s)
    {
        Rdv3Dlg.Caption(g, s, c.Left - Rdv3Skin.P(6), c.Top - Rdv3Skin.P(18), c.Width + Rdv3Skin.P(60));
    }

    private void Frame(Graphics g, TextBox t)
    {
        Rdv3Dlg.BoxFrame(g, Grow(t), t.Focused);
    }

    private static Rectangle Grow(Control c)
    {
        int p = Rdv3Skin.P(6);
        return new Rectangle(c.Left - p, c.Top - p, c.Width + 2 * p, c.Height + 2 * p);
    }

    // ---- data in and out ---------------------------------------------------
    private void LoadAll()
    {
        loading = true;
        lstTargets.Items.Clear();
        for (int i = 0; i < cfg.Targets.Count; i++) { lstTargets.Items.Add(Label(cfg.Targets[i])); }
        tKeyLen.Text = cfg.KeyLength.ToString(CultureInfo.InvariantCulture);
        cDigits.Checked = cfg.KeyDigitsOnly;
        tPoll.Text = cfg.PollMs.ToString(CultureInfo.InvariantCulture);
        tStable.Text = cfg.StableMs.ToString(CultureInfo.InvariantCulture);
        tRebind.Text = cfg.RebindMs.ToString(CultureInfo.InvariantCulture);
        tCand.Text = cfg.CandidateRowsShown.ToString(CultureInfo.InvariantCulture);
        cFocus.Checked = cfg.PreferFocusedWindow;
        tData.Text = cfg.DataDir;
        tLedger.Text = cfg.Ledger;
        tLog.Text = cfg.Log;
        loading = false;
        if (lstTargets.Items.Count > 0) { lstTargets.SelectedIndex = 0; } else { LoadTarget(); }
    }

    private static string Label(Rdv3Target t)
    {
        string s = t.Name;
        if (!t.Enabled) { s = s + "  (" + Rdv3Text.LblEnabled + ": -)"; }
        return s;
    }

    private Rdv3Target Current
    {
        get
        {
            int i = lstTargets.SelectedIndex;
            return (i >= 0 && i < cfg.Targets.Count) ? cfg.Targets[i] : null;
        }
    }

    private void LoadTarget()
    {
        Rdv3Target t = Current;
        loading = true;
        lstPath.Items.Clear();
        if (t == null)
        {
            tName.Text = ""; cEnabled.Checked = false;
            tWinId.Text = ""; tWinClass.Text = ""; tWinLike.Text = ""; tWinProc.Text = "";
            tFldId.Text = ""; tFldClass.Text = ""; tFldTypes.Text = ""; tFldIndex.Text = "0";
            cFieldValue.Checked = false; cFieldChildren.Checked = false;
            cbRead.SelectedIndex = 0;
        }
        else
        {
            tName.Text = t.Name;
            cEnabled.Checked = t.Enabled;
            tWinId.Text = t.Window.AutomationId;
            tWinClass.Text = t.Window.ClassName;
            tWinLike.Text = t.Window.NameLike;
            tWinProc.Text = t.Window.ProcessName;
            for (int i = 0; i < t.Steps.Count; i++) { lstPath.Items.Add(t.Steps[i].Describe()); }
            tFldId.Text = t.Field.AutomationId;
            tFldClass.Text = t.Field.ClassName;
            tFldTypes.Text = string.Join(",", t.Field.ControlTypes);
            tFldIndex.Text = t.Field.Index.ToString(CultureInfo.InvariantCulture);
            cFieldValue.Checked = t.Field.RequireValuePattern;
            cFieldChildren.Checked = !t.Field.Descendants;
            cbRead.SelectedIndex = t.ReadMode;
        }
        loading = false;
        Invalidate();
    }

    private void StoreTarget()
    {
        Rdv3Target t = Current;
        if (t == null || loading) { return; }
        t.Name = tName.Text.Trim();
        t.Enabled = cEnabled.Checked;
        t.Window.AutomationId = tWinId.Text.Trim();
        t.Window.ClassName = tWinClass.Text.Trim();
        t.Window.NameLike = tWinLike.Text.Trim();
        t.Window.ProcessName = tWinProc.Text.Trim();
        t.Field.AutomationId = tFldId.Text.Trim();
        t.Field.ClassName = tFldClass.Text.Trim();
        t.Field.ControlTypes = Split(tFldTypes.Text);
        t.Field.Index = ToInt(tFldIndex.Text, 0);
        t.Field.RequireValuePattern = cFieldValue.Checked;
        t.Field.Descendants = !cFieldChildren.Checked;
        t.ReadMode = Math.Max(0, cbRead.SelectedIndex);
    }

    private static string[] Split(string s)
    {
        if (s == null) { return new string[0]; }
        string[] raw = s.Split(new char[] { ',', ' ', '/' }, StringSplitOptions.RemoveEmptyEntries);
        List<string> outp = new List<string>();
        for (int i = 0; i < raw.Length; i++) { outp.Add(raw[i].Trim()); }
        return outp.ToArray();
    }

    private static int ToInt(string s, int fallback)
    {
        int v;
        return int.TryParse((s == null) ? "" : s.Trim(), NumberStyles.Integer,
            CultureInfo.InvariantCulture, out v) ? v : fallback;
    }

    private void AddTarget()
    {
        StoreTarget();
        Rdv3Target t = new Rdv3Target();
        t.Name = Rdv3Text.SecTarget + " " + (cfg.Targets.Count + 1).ToString(CultureInfo.InvariantCulture);
        t.Field.RequireValuePattern = true;
        cfg.Targets.Add(t);
        lstTargets.Items.Add(Label(t));
        lstTargets.SelectedIndex = cfg.Targets.Count - 1;
    }

    private void CopyTarget()
    {
        StoreTarget();
        Rdv3Target t = Current;
        if (t == null) { return; }
        Rdv3Target c = t.Clone();
        cfg.Targets.Add(c);
        lstTargets.Items.Add(Label(c));
        lstTargets.SelectedIndex = cfg.Targets.Count - 1;
    }

    private void RemoveTarget()
    {
        int i = lstTargets.SelectedIndex;
        if (i < 0) { return; }
        cfg.Targets.RemoveAt(i);
        lstTargets.Items.RemoveAt(i);
        if (lstTargets.Items.Count > 0) { lstTargets.SelectedIndex = Math.Min(i, lstTargets.Items.Count - 1); }
        else { LoadTarget(); }
    }

    private void RemoveStep()
    {
        Rdv3Target t = Current;
        int i = lstPath.SelectedIndex;
        if (t == null || i < 0 || i >= t.Steps.Count) { return; }
        t.Steps.RemoveAt(i);
        lstPath.Items.RemoveAt(i);
    }

    private void Inspect()
    {
        StoreTarget();
        Rdv3Target picked = Rdv3InspectForm.Pick(this);
        if (picked == null) { return; }
        Rdv3Target t = Current;
        if (t == null)
        {
            cfg.Targets.Add(picked);
            lstTargets.Items.Add(Label(picked));
            lstTargets.SelectedIndex = cfg.Targets.Count - 1;
            return;
        }
        // keep the operator's own name if they gave one
        if (t.Name.Length > 0 && !t.Name.StartsWith(Rdv3Text.SecTarget)) { picked.Name = t.Name; }
        cfg.Targets[lstTargets.SelectedIndex] = picked;
        lstTargets.Items[lstTargets.SelectedIndex] = Label(picked);
        LoadTarget();
    }

    private void SaveAndClose()
    {
        StoreTarget();
        cfg.KeyLength = Math.Max(1, Math.Min(64, ToInt(tKeyLen.Text, cfg.KeyLength)));
        cfg.KeyDigitsOnly = cDigits.Checked;
        cfg.PollMs = Math.Max(5, Math.Min(5000, ToInt(tPoll.Text, cfg.PollMs)));
        cfg.StableMs = Math.Max(0, Math.Min(60000, ToInt(tStable.Text, cfg.StableMs)));
        cfg.RebindMs = Math.Max(50, Math.Min(60000, ToInt(tRebind.Text, cfg.RebindMs)));
        cfg.CandidateRowsShown = Math.Max(1, Math.Min(1000, ToInt(tCand.Text, cfg.CandidateRowsShown)));
        cfg.PreferFocusedWindow = cFocus.Checked;
        cfg.DataDir = tData.Text.Trim();
        cfg.Ledger = tLedger.Text.Trim();
        cfg.Log = tLog.Text.Trim();
        Result = cfg;
        DialogResult = DialogResult.OK;
        Close();
    }
}

// ---------------------------------------------------------------------------
// the UI Automation inspector
// ---------------------------------------------------------------------------
public sealed class Rdv3InspectForm : Rdv3Dialog
{
    private readonly TreeView tree = new TreeView();
    private Rdv3Target result;
    private Rectangle rTree, rInfo, rFoot;
    private string info = "";

    public static Rdv3Target Pick(IWin32Window owner)
    {
        Rdv3InspectForm f = new Rdv3InspectForm();
        DialogResult dr = f.ShowDialog(owner);
        Rdv3Target t = (dr == DialogResult.OK) ? f.result : null;
        f.Dispose();
        return t;
    }

    private Rdv3InspectForm()
    {
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.BtnInspect;
        Rdv3Dlg.Skin(this, 900, 620);
        int pad = Rdv3Dlg.P(13.6);
        rTree = new Rectangle(pad, pad, Rdv3Dlg.P(520), ClientSize.Height - pad * 2 - Rdv3Dlg.P(48));
        rInfo = new Rectangle(rTree.Right + pad, pad, ClientSize.Width - rTree.Right - pad * 2, rTree.Height);
        rFoot = new Rectangle(0, ClientSize.Height - Rdv3Dlg.P(48), ClientSize.Width, Rdv3Dlg.P(48));

        tree.SetBounds(rTree.X + Rdv3Dlg.P(8), rTree.Y + Rdv3Dlg.P(44),
            rTree.Width - Rdv3Dlg.P(16), rTree.Height - Rdv3Dlg.P(52));
        tree.BorderStyle = BorderStyle.None;
        tree.BackColor = Rdv3Skin.Surface;
        tree.ForeColor = Rdv3Skin.Ink;
        tree.Font = Rdv3Skin.F(12, FontStyle.Regular);
        tree.HideSelection = false;
        tree.BeforeExpand += delegate(object s, TreeViewCancelEventArgs e) { Expand(e.Node); };
        tree.AfterSelect += delegate { ShowInfo(); };
        Controls.Add(tree);

        int fy = rFoot.Y + Rdv3Dlg.P(8);
        Rdv3Dlg.Button(this, Rdv3Dlg.P(13.6), fy, Rdv3Dlg.P(120), Rdv3Text.BtnRefresh, false, 3)
            .Click += delegate { Reload(); };
        int xClose = ClientSize.Width - Rdv3Dlg.P(13.6) - Rdv3Dlg.P(120);
        int xUse = xClose - Rdv3Dlg.P(10) - Rdv3Dlg.P(260);
        Rdv3Btn use = Rdv3Dlg.Button(this, xUse, fy, Rdv3Dlg.P(260), Rdv3Text.BtnUseElement, true, 2);
        use.Click += delegate { Use(); };
        Rdv3Btn close = Rdv3Dlg.Button(this, xClose, fy, Rdv3Dlg.P(120), Rdv3Text.BtnClose, false, 0);
        close.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        CancelButton = close;
        Reload();
    }

    // Each node holds its element; children are filled in on demand, because a
    // business application's tree is far too large to walk in one go.
    private sealed class Node
    {
        public AutomationElement E;
        public bool Filled;
    }

    private void Reload()
    {
        tree.BeginUpdate();
        tree.Nodes.Clear();
        try
        {
            AutomationElement root = AutomationElement.RootElement;
            AutomationElementCollection wins = root.FindAll(TreeScope.Children, Condition.TrueCondition);
            for (int i = 0; i < wins.Count; i++)
            {
                TreeNode n = Make(wins[i]);
                if (n != null) { tree.Nodes.Add(n); }
            }
        }
        catch (Exception ex) { info = ex.Message; }
        tree.EndUpdate();
        Invalidate();
    }

    private static TreeNode Make(AutomationElement e)
    {
        string text;
        try
        {
            AutomationElement.AutomationElementInformation c = e.Current;
            StringBuilder sb = new StringBuilder();
            sb.Append(Rdv3Uia.ControlTypeName(c.ControlType));
            if (c.AutomationId.Length > 0) { sb.Append("  #").Append(c.AutomationId); }
            if (c.Name.Length > 0) { sb.Append("  ").Append(Cut(c.Name, 40)); }
            if (c.ClassName.Length > 0) { sb.Append("  [").Append(c.ClassName).Append("]"); }
            text = sb.ToString();
        }
        catch (Exception) { return null; }
        TreeNode n = new TreeNode(text);
        Node d = new Node();
        d.E = e;
        n.Tag = d;
        n.Nodes.Add(new TreeNode("..."));          // a placeholder so it can expand
        return n;
    }

    private static string Cut(string s, int n)
    {
        if (s == null) { return ""; }
        s = s.Replace("\r", " ").Replace("\n", " ");
        return (s.Length <= n) ? s : (s.Substring(0, n) + "...");
    }

    private void Expand(TreeNode n)
    {
        Node d = n.Tag as Node;
        if (d == null || d.Filled) { return; }
        d.Filled = true;
        n.Nodes.Clear();
        try
        {
            AutomationElementCollection kids = d.E.FindAll(TreeScope.Children, Condition.TrueCondition);
            for (int i = 0; i < kids.Count && i < 500; i++)
            {
                TreeNode k = Make(kids[i]);
                if (k != null) { n.Nodes.Add(k); }
            }
        }
        catch (Exception) { }
    }

    private AutomationElement Selected
    {
        get
        {
            TreeNode n = tree.SelectedNode;
            Node d = (n == null) ? null : (n.Tag as Node);
            return (d == null) ? null : d.E;
        }
    }

    private void ShowInfo()
    {
        AutomationElement e = Selected;
        if (e == null) { info = ""; Invalidate(); return; }
        StringBuilder sb = new StringBuilder();
        try
        {
            AutomationElement.AutomationElementInformation c = e.Current;
            Line(sb, Rdv3Text.LblControlTypes, Rdv3Uia.ControlTypeName(c.ControlType));
            Line(sb, Rdv3Text.LblAutomationId, c.AutomationId);
            Line(sb, Rdv3Text.LblClassName, c.ClassName);
            Line(sb, Rdv3Text.LblName, Cut(c.Name, 60));
            Line(sb, Rdv3Text.LblPatterns, Patterns(e));
            Line(sb, Rdv3Text.LblProcessOf, ProcName(c.ProcessId));
            string v = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadValue);
            if (v == null) { v = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadText); }
            if (v != null) { Line(sb, "value", Cut(Rdv3Watch.Candidate(v), 60)); }
        }
        catch (Exception ex) { sb.Append(ex.Message); }
        info = sb.ToString();
        Invalidate();
    }

    private static void Line(StringBuilder sb, string k, string v)
    {
        sb.Append(k).Append(":\t").Append((v == null) ? "" : v).Append("\n");
    }

    private static string Patterns(AutomationElement e)
    {
        StringBuilder sb = new StringBuilder();
        object p;
        try { if (e.TryGetCurrentPattern(ValuePattern.Pattern, out p)) { sb.Append("value"); } }
        catch (Exception) { }
        try
        {
            if (e.TryGetCurrentPattern(TextPattern.Pattern, out p))
            {
                if (sb.Length > 0) { sb.Append(" + "); }
                sb.Append("text");
            }
        }
        catch (Exception) { }
        if (sb.Length == 0) { sb.Append("name"); }
        return sb.ToString();
    }

    private static string ProcName(int pid)
    {
        try { return System.Diagnostics.Process.GetProcessById(pid).ProcessName; }
        catch (Exception) { return ""; }
    }

    // ---- turning the picked element into a target --------------------------
    // Walk up to the top level window, keep the levels that carry a stable
    // AutomationId as path steps, and describe the element itself as the field.
    private void Use()
    {
        AutomationElement e = Selected;
        if (e == null) { return; }
        try
        {
            TreeWalker w = TreeWalker.ControlViewWalker;
            AutomationElement root = AutomationElement.RootElement;
            List<AutomationElement> chain = new List<AutomationElement>();
            AutomationElement at = e;
            for (int guard = 0; guard < 64 && at != null; guard++)
            {
                AutomationElement parent = w.GetParent(at);
                if (parent == null || Rdv3Uia.Same(parent, root)) { break; }
                chain.Add(parent);
                at = parent;
            }
            chain.Reverse();                      // window first
            AutomationElement win = (chain.Count > 0) ? chain[0] : e;

            Rdv3Target t = new Rdv3Target();
            t.Window = new Rdv3Match();
            t.Window.Descendants = false;
            t.Window.ClassName = win.Current.ClassName;
            t.Window.ProcessName = ProcName(win.Current.ProcessId);
            if (win.Current.AutomationId.Length > 0) { t.Window.AutomationId = win.Current.AutomationId; }
            t.Name = (win.Current.Name.Length > 0) ? Cut(win.Current.Name, 24) : t.Window.ProcessName;

            for (int i = 1; i < chain.Count; i++)
            {
                string id = chain[i].Current.AutomationId;
                if (id.Length == 0) { continue; }         // only stable levels are worth naming
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
                t.ReadMode = (Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadText) != null)
                    ? Rdv3Uia.ReadText : Rdv3Uia.ReadName;
            }
            if (t.ReadMode == Rdv3Uia.ReadValue) { t.Field.RequireValuePattern = true; }

            result = t;
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            info = ex.Message;
            Invalidate();
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        Rdv3Dlg.Section(g, Rdv3Text.ColElement, rTree);
        Rdv3Dlg.Section(g, Rdv3Text.SecTarget, rInfo);
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, new Rectangle(tree.Left - Rdv3Skin.P(4), tree.Top - Rdv3Skin.P(4),
            tree.Width + Rdv3Skin.P(8), tree.Height + Rdv3Skin.P(8)));

        int x = rInfo.X + Rdv3Skin.P(13.6);
        int y = rInfo.Y + Rdv3Skin.P(52);
        int w = rInfo.Width - Rdv3Skin.P(27.2);
        Font fk = Rdv3Skin.F(11, FontStyle.Regular);
        Font fv = Rdv3Skin.F(12, FontStyle.Regular);
        string[] rows = info.Split('\n');
        for (int i = 0; i < rows.Length; i++)
        {
            if (rows[i].Length == 0) { continue; }
            int tab = rows[i].IndexOf('\t');
            string k = (tab >= 0) ? rows[i].Substring(0, tab) : rows[i];
            string v = (tab >= 0) ? rows[i].Substring(tab + 1) : "";
            Rdv3Skin.DrawIn(g, k, fk, Rdv3Skin.N700, new Rectangle(x, y, w, Rdv3Skin.P(15)), false);
            Rdv3Skin.DrawIn(g, v, fv, Rdv3Skin.Ink, new Rectangle(x, y + Rdv3Skin.P(15), w, Rdv3Skin.P(19)), false);
            y += Rdv3Skin.P(38);
        }
        Rdv3Skin.DrawIn(g, Rdv3Text.NoteInspectHint, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(x, rInfo.Bottom - Rdv3Skin.P(56), w, Rdv3Skin.P(48)), false);

        using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent900)) { g.FillRectangle(b, rFoot); }
    }
}
