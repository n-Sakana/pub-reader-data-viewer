// ============================================================================
// Rdv3Settings.cs -- the settings modal and the element picker, to the v2
// reference.
//
// The modal is one page, 560 px wide: the watched targets as cards, the
// "pick from the screen" button, the key pattern, and the three paths.
// Everything else the settings file holds (poll timings, job limits) is kept
// as written and not edited here -- the reference offers no controls for it.
//
// THE PICKER. Naming a field inside a business application by typing
// identifiers is hopeless. So: hover over the field, read its identity live,
// take it with Ctrl+Shift. UI Automation the whole way --
// AutomationElement.FromPoint for what is under the cursor, Control.ModifierKeys
// for the shortcut. No hooks, no RegisterHotKey, no user32 mouse polling.
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
// a borderless value box inside a painted rounded field
// ---------------------------------------------------------------------------
internal sealed class Rdv3SetBox : TextBox
{
    private readonly float sc;

    public Rdv3SetBox(float scale, bool mono)
    {
        sc = scale;
        BorderStyle = BorderStyle.None;
        BackColor = Rdv3Skin.N200;
        ForeColor = Rdv3Skin.Ink;
        Font = mono ? new Font("Consolas", (float)(16 * sc), FontStyle.Regular, GraphicsUnit.Pixel)
                    : Rdv3Skin.Px(16, sc);
    }

    // the reference's input padding is 0 12px; the frame is the caller's
    public void PlaceIn(Rectangle frame)
    {
        int pad = (int)Math.Round(12.0 * sc);
        int h = PreferredHeight;
        SetBounds(frame.X + pad, frame.Y + (frame.Height - h) / 2, Math.Max(1, frame.Width - 2 * pad), h);
    }
}

// ---------------------------------------------------------------------------
// the target list: each card says what the target is called and what it
// points at; the selected one is accent-100
// ---------------------------------------------------------------------------
internal sealed class Rdv3SetCards : Control
{
    private readonly float sc;
    public List<string[]> Items = new List<string[]>();   // { name, summary }
    public int Index;
    public Action OnPick;
    public string Empty = "";
    private int hot = -1;
    private int top;
    private readonly Font fName, fSum;
    public const double CardH = 62;

    public Rdv3SetCards(float scale)
    {
        sc = scale;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.White;
        fName = Rdv3Skin.S(16, sc);
        fSum = Rdv3Skin.F(14, FontStyle.Regular, sc);
        Cursor = Cursors.Hand;
    }

    private int P(double v) { return (int)Math.Round(v * sc); }
    public int RowH { get { return P(CardH); } }
    public int Rows { get { return Math.Max(1, (Height - 2) / Math.Max(1, RowH)); } }
    private bool NeedBar { get { return Items.Count > Rows; } }

    public void ScrollToSelected()
    {
        if (Index < top) { top = Index; }
        else if (Index >= top + Rows) { top = Index - Rows + 1; }
        int max = Math.Max(0, Items.Count - Rows);
        if (top > max) { top = max; }
        if (top < 0) { top = 0; }
    }

    private int RowAt(int y)
    {
        if (y < 1) { return -1; }
        int i = top + (y - 1) / Math.Max(1, RowH);
        return (i >= 0 && i < Items.Count) ? i : -1;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        int i = RowAt(e.Y);
        if (i != hot) { hot = i; Invalidate(); }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e) { hot = -1; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        int i = RowAt(e.Y);
        if (i >= 0) { Index = i; ScrollToSelected(); Invalidate(); if (OnPick != null) { OnPick(); } }
        base.OnMouseClick(e);
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        int max = Math.Max(0, Items.Count - Rows);
        top = Math.Max(0, Math.Min(max, top + ((e.Delta > 0) ? -1 : 1)));
        Invalidate();
        base.OnMouseWheel(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        Rectangle all = new Rectangle(0, 0, Width, Height);
        Rdv3Skin.FillRound(g, Rdv3Skin.White, all, P(10));
        int w = Width - 2 - (NeedBar ? P(10) : 0);
        if (Items.Count == 0)
        {
            Rdv3Skin.DrawIn(g, Empty, fSum, Rdv3Skin.N600, new Rectangle(P(15), 0, Width - P(30), Height), 0);
        }
        int y = 1;
        for (int i = top; i < Items.Count && y < Height - 1; i++)
        {
            Rectangle card = new Rectangle(1, y, w, RowH);
            Color back = (i == Index) ? Rdv3Skin.Accent100 : ((i == hot) ? Rdv3Skin.Mix(Rdv3Skin.White, 0.04) : Rdv3Skin.White);
            using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, card); }
            Rectangle name = new Rectangle(card.X + P(14), card.Y + P(10), card.Width - P(28), P(24.8));
            Rectangle sum = new Rectangle(card.X + P(14), card.Y + P(35.9), card.Width - P(28), P(21.7));
            Rdv3Skin.DrawIn(g, Items[i][0], fName, Rdv3Skin.Ink, name, 0);
            Rdv3Skin.DrawIn(g, Items[i][1], fSum, Rdv3Skin.N600, sum, 0);
            if (i + 1 < Items.Count) { Rdv3Skin.Line(g, Rdv3Skin.Mix(back, 0.08), card.X, card.Bottom - 1, card.Width, 1); }
            y += RowH;
        }
        if (NeedBar)
        {
            Rectangle track = new Rectangle(Width - 1 - P(10), 1, P(10), Height - 2);
            double frac = Rows / (double)Math.Max(1, Items.Count);
            int th = Math.Max(P(24), (int)(track.Height * frac));
            int max = Math.Max(0, Items.Count - Rows);
            int off = (max == 0) ? 0 : (int)((track.Height - th) * (top / (double)max));
            using (SolidBrush b = new SolidBrush(Rdv3Skin.N400)) { g.FillRectangle(b, track.X + P(3), track.Y + off, P(4), th); }
        }
        Rdv3Skin.FrameRound(g, Rdv3Skin.Mix(Rdv3Skin.White, 0.09), all, P(10));
    }
}

// ---------------------------------------------------------------------------
// the settings modal
// ---------------------------------------------------------------------------
public sealed class Rdv3SettingsForm : Rdv3Dialog
{
    private readonly Rdv3Config cfg;
    private Rdv3SetCards cards;
    private Rdv3SetBox tPattern, tData, tLedger, tLog;
    private Rdv3Btn btnClose, btnPick, btnSave, btnCancel;
    private Font fTitle, fSec, fLabel, fNote;
    public Rdv3Config Result;
    // why the last save was refused, shown in place of the note under the
    // field concerned until the next attempt
    private string problemPattern = "";
    private string problemFiles = "";

    private const double ModalW = 560;
    private const double ModalH = 769.5;
    private const double HeadH = 59;
    private const double FootH = 63;
    private const double LabelH = 20.15;
    private const double SmallH = 21.7;

    public static Rdv3Config Edit(Rdv3Form owner, Rdv3Config current)
    {
        Rdv3SettingsForm f = new Rdv3SettingsForm(owner, current.Clone());
        DialogResult dr = f.ShowOver(owner);
        Rdv3Config outp = (dr == DialogResult.OK) ? f.Result : null;
        f.Dispose();
        return outp;
    }

    private Rdv3SettingsForm(Rdv3Form owner, Rdv3Config working)
    {
        cfg = working;
        Sc = FitScale(ModalW, ModalH, (owner == null) ? Rdv3Skin.Scale : owner.Sc, owner);
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.SettingsTitle;
        MakeFonts();
        SizeTo(ModalW, ModalH);
        Build();
        LoadAll();
        Layout1();
    }

    // the headless check builds the dialog and renders it without showing it
    public static Rdv3SettingsForm ForCheck(Rdv3Config cfg, float scale)
    {
        Rdv3SettingsForm f = new Rdv3SettingsForm(null, cfg);
        f.SetDesignScale(scale);
        return f;
    }

    public void SetDesignScale(float s)
    {
        Sc = Math.Max(0.5f, s);
        MakeFonts();
        SizeTo(ModalW, ModalH);
        Controls.Clear();
        Build();
        LoadAll();
        Layout1();
        Invalidate();
    }

    private void MakeFonts()
    {
        fTitle = Fb(20);
        fSec = Sf(13);
        fLabel = Fn(14);
        fNote = Fn(13);
    }

    private void Build()
    {
        btnClose = MakeBtn("", Rdv3Btn.Round, 13, "close");
        btnClose.IconCss = 14;
        btnClose.Click += delegate { Cancel(); };

        cards = new Rdv3SetCards(Sc);
        cards.Empty = Rdv3Text.NoteNoTargetShort;
        cards.OnPick = delegate { Invalidate(); };
        Controls.Add(cards);

        btnPick = MakeBtn(Rdv3Text.BtnInspect, Rdv3Btn.Soft, 15, "search");
        btnPick.IconCss = 15;
        btnPick.Click += delegate { Pick(); };

        tPattern = Box(true);
        tData = Box(false);
        tLedger = Box(false);
        tLog = Box(false);

        btnSave = MakeBtn(Rdv3Text.BtnSave, Rdv3Btn.Primary, 15, "");
        btnSave.BackColor = Rdv3Skin.N100;
        btnSave.Click += delegate { SaveAndClose(); };
        btnCancel = MakeBtn(Rdv3Text.BtnCancel, Rdv3Btn.Secondary, 15, "");
        btnCancel.BackColor = Rdv3Skin.N100;
        btnCancel.Click += delegate { Cancel(); };
        CancelButton = btnCancel;
    }

    private Rdv3SetBox Box(bool mono)
    {
        Rdv3SetBox b = new Rdv3SetBox(Sc, mono);
        b.GotFocus += delegate { Invalidate(); };
        b.LostFocus += delegate { Invalidate(); };
        Controls.Add(b);
        return b;
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Layout1();
    }

    // ---- layout: the reference's own coordinates ----------------------------
    private void Layout1()
    {
        if (fTitle == null) { return; }
        ClearRects();
        double x = 20, w = ModalW - 40;
        Put("st.head", 0, 0, ModalW, HeadH);
        Put("st.title", x, 16, Math.Max(40, MW(Rdv3Text.SettingsTitle, fTitle)), 31);
        Put("st.close", ModalW - 20 - 30, 16.5, 30, 30);
        Place(btnClose, "st.close");

        double y = HeadH + 4;
        Put("st.sec.targets", x, y, w, LabelH); y += LabelH + 8;
        Put("st.cards", x, y, w, 64);
        Place(cards, "st.cards");
        y += 64 + 10;
        Put("st.pick", x, y, w, 38);
        Place(btnPick, "st.pick");
        y += 38 + 8;
        Put("st.note.target", x, y, w, SmallH); y += SmallH;
        y += 18; Put("st.rule1", x, y, w, 1); y += 1 + 16;
        Put("st.sec.pattern", x, y, w, LabelH); y += LabelH + 4 + 6;
        Put("st.pattern", x, y, w, 40);
        tPattern.PlaceIn(At("st.pattern"));
        y += 40 + 8;
        Put("st.note.pattern", x, y, w, LabelH); y += LabelH;
        y += 18; Put("st.rule2", x, y, w, 1); y += 1 + 16;
        Put("st.sec.files", x, y, w, LabelH); y += LabelH + 10;
        y = LayPath("st.data", tData, x, y, w);
        y += 12;
        y = LayPath("st.ledger", tLedger, x, y, w);
        y += 12;
        y = LayPath("st.log", tLog, x, y, w);
        y += 8;
        Put("st.note.files", x, y, w, LabelH);

        Put("st.foot", 0, ModalH - FootH, ModalW, FootH);
        Put("st.foot.rule", 0, ModalH - FootH, ModalW, 1);
        double by = ModalH - FootH + 12.5;
        double wSave = Math.Max(74, MW(Rdv3Text.BtnSave, btnSave.Font) + 44);
        double wCancel = Math.Max(62, MW(Rdv3Text.BtnCancel, btnCancel.Font) + 32);
        Put("st.save", ModalW - 20 - wSave, by, wSave, 38);
        Put("st.cancel", ModalW - 20 - wSave - 8 - wCancel, by, wCancel, 38);
        Place(btnSave, "st.save");
        Place(btnCancel, "st.cancel");
        DragZone = At("st.head");
    }

    private double LayPath(string k, Rdv3SetBox box, double x, double y, double w)
    {
        Put(k + ".label", x, y, w, SmallH);
        y += SmallH + 6;
        Put(k, x, y, w, 40);
        box.PlaceIn(At(k));
        return y + 40;
    }

    // ---- painting ---------------------------------------------------------
    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        ClearClip();
        using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, ClientRectangle); }
        T(g, "st.title", Rdv3Text.SettingsTitle, fTitle, Rdv3Skin.Ink, 0);

        T(g, "st.sec.targets", Rdv3Text.SecTargets, fSec, Rdv3Skin.N600, 0);
        T(g, "st.note.target", Rdv3Text.NoteReadTarget.Replace("{sum}", Summary(Current)), fLabel, Rdv3Skin.N600, 0);
        Rule9(g, "st.rule1");
        T(g, "st.sec.pattern", Rdv3Text.SecKeyPattern, fSec, Rdv3Skin.N600, 0);
        Field(g, "st.pattern", tPattern.Focused);
        if (problemPattern.Length > 0) { T(g, "st.note.pattern", problemPattern, fNote, Rdv3Skin.Danger, 0); }
        else { T(g, "st.note.pattern", Rdv3Text.NoteKeyPattern, fNote, Rdv3Skin.N600, 0); }
        Rule9(g, "st.rule2");
        T(g, "st.sec.files", Rdv3Text.SecPaths, fSec, Rdv3Skin.N600, 0);
        T(g, "st.data.label", Rdv3Text.LblDataDir, fLabel, Rdv3Skin.N700, 0);
        Field(g, "st.data", tData.Focused);
        T(g, "st.ledger.label", Rdv3Text.LblLedger, fLabel, Rdv3Skin.N700, 0);
        Field(g, "st.ledger", tLedger.Focused);
        T(g, "st.log.label", Rdv3Text.LblLog, fLabel, Rdv3Skin.N700, 0);
        Field(g, "st.log", tLog.Focused);
        if (problemFiles.Length > 0) { T(g, "st.note.files", problemFiles, fNote, Rdv3Skin.Danger, 0); }
        else { T(g, "st.note.files", Rdv3Text.NoteFilesBase, fNote, Rdv3Skin.N600, 0); }

        Rectangle foot = At("st.foot");
        Rectangle footRound = new Rectangle(foot.X, foot.Y - PX(12), foot.Width, foot.Height + PX(12));
        Rdv3Skin.FillRound(g, Rdv3Skin.N100, footRound, PX(12));
        Rdv3Skin.Line(g, Rdv3Skin.White, foot.X, foot.Y - PX(12), foot.Width, PX(12));
        Rule(g, "st.foot.rule");
    }

    private void Rule9(Graphics g, string k)
    {
        Rectangle r = At(k);
        Rdv3Skin.Line(g, Rdv3Skin.Mix(Rdv3Skin.White, 0.09), r.X, r.Y, r.Width, 1);
    }

    // ---- data ---------------------------------------------------------------
    private Rdv3Target Current
    {
        get
        {
            int i = cards.Index;
            return (i >= 0 && i < cfg.Targets.Count) ? cfg.Targets[i] : null;
        }
    }

    private void LoadAll()
    {
        RefreshCards();
        tPattern.Text = cfg.KeyPattern;
        tData.Text = cfg.DataDir;
        tLedger.Text = cfg.Ledger;
        tLog.Text = cfg.Log;
    }

    private void RefreshCards()
    {
        cards.Items.Clear();
        for (int i = 0; i < cfg.Targets.Count; i++)
        {
            Rdv3Target t = cfg.Targets[i];
            string name = (t.Name.Length > 0) ? t.Name : Rdv3Text.SecTarget;
            if (!t.Enabled) { name += Rdv3Text.LblDisabled; }
            cards.Items.Add(new string[] { name, Summary(t) });
        }
        if (cards.Index >= cfg.Targets.Count) { cards.Index = Math.Max(0, cfg.Targets.Count - 1); }
        cards.ScrollToSelected();
        cards.Invalidate();
    }

    private static string Summary(Rdv3Target t)
    {
        if (t == null) { return Rdv3Text.NoValue; }
        string win = t.Window.ProcessName.Length > 0 ? t.Window.ProcessName
            : (t.Window.ClassName.Length > 0 ? t.Window.ClassName : t.Window.NameLike);
        string fld = t.Field.AutomationId.Length > 0 ? ("#" + t.Field.AutomationId)
            : (t.Field.ControlTypes.Length > 0 ? string.Join("/", t.Field.ControlTypes) : t.Field.ClassName);
        if (win.Length == 0) { win = Rdv3Text.NoValue; }
        if (fld.Length == 0) { fld = Rdv3Text.NoValue; }
        return Rdv3Text.NoteTargetSummary.Replace("{win}", win).Replace("{field}", fld);
    }

    private void Pick()
    {
        Rdv3Target picked = Rdv3PickerForm.Pick(this, Sc);
        if (picked == null) { return; }
        Rdv3Target t = Current;
        if (t == null)
        {
            cfg.Targets.Add(picked);
            cards.Index = cfg.Targets.Count - 1;
        }
        else
        {
            if (t.Name.Length > 0 && !t.Name.StartsWith(Rdv3Text.SecTarget)) { picked.Name = t.Name; }
            picked.Enabled = t.Enabled;
            cfg.Targets[cards.Index] = picked;
        }
        RefreshCards();
        Invalidate();
    }

    private void Cancel()
    {
        DialogResult = DialogResult.Cancel;
        Close();
    }

    // Nothing is saved half-right: a pattern that does not compile or a blank
    // path keeps the dialog open and says why, instead of being dropped in
    // silence while the rest is written.
    private void SaveAndClose()
    {
        string pat = tPattern.Text.Trim();
        string why = Rdv3Config.PatternError(pat);
        problemPattern = (why == null) ? "" : (Rdv3Text.ErrPatternTyped + why);
        bool blank = tData.Text.Trim().Length == 0 || tLedger.Text.Trim().Length == 0 || tLog.Text.Trim().Length == 0;
        problemFiles = blank ? Rdv3Text.ErrPathBlank : "";
        if (problemPattern.Length > 0 || problemFiles.Length > 0) { Invalidate(); return; }
        cfg.KeyPattern = pat;
        cfg.DataDir = tData.Text.Trim();
        cfg.Ledger = tLedger.Text.Trim();
        cfg.Log = tLog.Text.Trim();
        Result = cfg;
        DialogResult = DialogResult.OK;
        Close();
    }

}

// ---------------------------------------------------------------------------
// the picker: hover over the field, read what it is, take it with Ctrl+Shift
// ---------------------------------------------------------------------------
public sealed class Rdv3PickerForm : Rdv3Dialog
{
    private readonly System.Windows.Forms.Timer tick = new System.Windows.Forms.Timer();
    private AutomationElement hover;
    private string ctrlType = "", autoId = "", className = "", elName = "", value = "", proc = "";
    private bool armed;
    private bool canRead;
    private Rdv3Target result;
    private Rdv3Btn btnClose;
    private Font fTitle, fHow, fKey, fVal, fValBold, fTag;

    private const double PickW = 460;
    private const double PickH = 323.5;
    private const double RowH = 27.6;

    public static Rdv3Target Pick(IWin32Window owner, float screenScale)
    {
        Rdv3PickerForm f = new Rdv3PickerForm(screenScale);
        DialogResult dr = f.ShowDialog(owner);
        Rdv3Target t = (dr == DialogResult.OK) ? f.result : null;
        f.Dispose();
        return t;
    }

    private Rdv3PickerForm(float screenScale)
    {
        Sc = FitScale(PickW, PickH, screenScale, null);
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.PickTitle;
        TopMost = true;
        MakeFonts();
        SizeTo(PickW, PickH);

        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Location = new Point(wa.Right - Width - PX(20), wa.Bottom - Height - PX(56));

        btnClose = MakeBtn(Rdv3Text.BtnClose, Rdv3Btn.Secondary, 16, "");
        btnClose.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        CancelButton = btnClose;

        tick.Interval = 90;
        tick.Tick += delegate { Look(); };
        tick.Start();
        Layout1();
    }

    // the acceptance harness renders the panel without a live cursor under it
    public static Rdv3PickerForm ForCheck(float scale)
    {
        Rdv3PickerForm f = new Rdv3PickerForm(scale);
        f.tick.Stop();
        f.SetDesignScale(scale);
        return f;
    }

    public void SetDesignScale(float s)
    {
        Sc = Math.Max(0.5f, s);
        MakeFonts();
        SizeTo(PickW, PickH);
        btnClose.Sc = Sc; btnClose.Font = Sf(16);
        Layout1();
        Invalidate();
    }

    private void MakeFonts()
    {
        fTitle = Sf(22);
        fHow = Fb(15);
        fKey = Fn(14);
        fVal = Fn(15);
        fValBold = Fb(15);
        fTag = Fn(13);
    }

    public void SetSample(string type, string id, string cls, string name, string process, string read)
    {
        ctrlType = type; autoId = id; className = cls; elName = name; proc = process;
        value = read; canRead = (read != null && read.Length > 0);
        Layout1();
        Invalidate();
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Layout1();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        tick.Stop();
        base.OnFormClosed(e);
    }

    private void Layout1()
    {
        if (fTitle == null) { return; }
        ClearRects();
        double x = 14.3, w = PickW - 28.6;
        Put("pk.title", x, 14.3, Math.Max(119.1, MW(Rdv3Text.PickTitle, fTitle)), 34.1);
        Size tz = Rdv3Skin.TagSize(Rdv3Text.TagTopMost, fTag, Sc);
        double tw = tz.Width / (double)Sc, th = tz.Height / (double)Sc;
        Put("pk.tag", PickW - 14.3 - tw, 14.3 + (34.1 - th) / 2, tw, th);
        Put("pk.how", x, 56.4, w, 23.3);
        Put("pk.rule", x, 89.7, w, 1);
        double y = 101.7;
        for (int i = 0; i < 6; i++)
        {
            Put("pk.k" + i.ToString(CultureInfo.InvariantCulture), x, y + (RowH - 21.7) / 2, 120, 21.7);
            Put("pk.v" + i.ToString(CultureInfo.InvariantCulture), x + 120, y + (RowH - 23.3) / 2, w - 120, 23.3);
            y += RowH;
        }
        double fy = PickH - 14.3 - 30;
        Put("pk.esc", x, fy + (30 - 21.7) / 2, Math.Max(62.6, MW(Rdv3Text.PickEsc, fKey)), 21.7);
        double cw = Math.Max(72, MW(Rdv3Text.BtnClose, btnClose.Font) + 24);
        Put("pk.close", PickW - 14.3 - cw, fy, cw, 30);
        Place(btnClose, "pk.close");
        DragZone = new Rectangle(0, 0, ClientSize.Width, At("pk.rule").Y);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics gr = e.Graphics;
        gr.SmoothingMode = SmoothingMode.AntiAlias;
        gr.PixelOffsetMode = PixelOffsetMode.Half;
        ClearClip();
        using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { gr.FillRectangle(b, ClientRectangle); }
        T(gr, "pk.title", Rdv3Text.PickTitle, fTitle, Rdv3Skin.Ink, 0);
        Rdv3Skin.Tag(gr, Rdv3Text.TagTopMost, fTag, Rdv3Skin.TagAccent, At("pk.tag"), Sc);
        T(gr, "pk.how", Rdv3Text.PickHow, fHow, Rdv3Skin.Accent700, 0);
        Rule(gr, "pk.rule");
        Row(gr, 0, Rdv3Text.LblControlTypes, ctrlType, false, false);
        Row(gr, 1, Rdv3Text.LblAutomationId, autoId, true, false);
        Row(gr, 2, Rdv3Text.LblClassName, className, false, false);
        Row(gr, 3, Rdv3Text.LblName, elName, false, false);
        Row(gr, 4, Rdv3Text.LblProcessOf, proc, false, false);
        Row(gr, 5, Rdv3Text.PickReading, canRead ? value : Rdv3Text.PickNoRead, false, !canRead);
        T(gr, "pk.esc", Rdv3Text.PickEsc, fKey, Rdv3Skin.N700, 0);
    }

    private void Row(Graphics g, int i, string k, string v, bool strong, bool muted)
    {
        T(g, "pk.k" + i.ToString(CultureInfo.InvariantCulture), k, fKey, Rdv3Skin.N700, 0);
        string shown = (v == null || v.Length == 0) ? Rdv3Text.NotYet : v;
        T(g, "pk.v" + i.ToString(CultureInfo.InvariantCulture), shown, strong ? fValBold : fVal, muted ? Rdv3Skin.N600 : Rdv3Skin.Ink, 0);
    }

    // What is under the cursor, and is the shortcut down? Both come from UIA and
    // WinForms: AutomationElement.FromPoint and Control.ModifierKeys. No hook.
    private void Look()
    {
        Point p = Cursor.Position;
        if (Bounds.Contains(p)) { return; }              // do not describe ourselves
        AutomationElement e = null;
        try { e = AutomationElement.FromPoint(new System.Windows.Point(p.X, p.Y)); }
        catch (Exception) { e = null; }

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
            catch (Exception) { }
            string v = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadValue);
            if (v == null) { v = Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadText); }
            canRead = (v != null);
            value = (v == null) ? "" : Cut(Rdv3Watch.Candidate(v), 46);
            Invalidate();
        }

        bool down = (Control.ModifierKeys & (Keys.Control | Keys.Shift)) == (Keys.Control | Keys.Shift);
        if (down && !armed && hover != null)
        {
            armed = true;
            Take();
        }
        else if (!down)
        {
            armed = false;
        }
    }

    private static string ProcName(int pid)
    {
        try { return System.Diagnostics.Process.GetProcessById(pid).ProcessName; }
        catch (Exception) { return ""; }
    }

    private static string Cut(string s, int n)
    {
        if (s == null) { return ""; }
        s = s.Replace("\r", " ").Replace("\n", " ");
        return (s.Length <= n) ? s : (s.Substring(0, n) + "...");
    }

    // walk up to the window, keep the levels that carry an AutomationId
    private void Take()
    {
        AutomationElement e = hover;
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
            chain.Reverse();
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
                t.ReadMode = (Rdv3Uia.ReadValueOf(e, Rdv3Uia.ReadText) != null)
                    ? Rdv3Uia.ReadText : Rdv3Uia.ReadName;
            }
            if (t.ReadMode == Rdv3Uia.ReadValue) { t.Field.RequireValuePattern = true; }

            // WHICH of the fields that match. Without an automationId the same
            // container commonly holds several identical Edits, and the resolver
            // takes Field.Index from the same anchor with the same matcher --
            // so the ordinal is measured here exactly the way it will be read
            // back. Left at 0, picking the second box saved a target that reads
            // the first.
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
        catch (Exception) { }
    }
}
