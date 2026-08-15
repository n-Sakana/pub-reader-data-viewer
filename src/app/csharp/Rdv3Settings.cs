// ============================================================================
// Rdv3Settings.cs -- the settings dialog and the element picker.
//
// DESIGN NOTES, because the first attempt got these wrong.
//
//  * The dark navy bar is the STATUS BAR: it means "this application's live
//    state", there is one per app window, and it sits at the bottom of it. A
//    modal has no status, so its action row is plain surface with a hairline
//    above it. Reusing the navy here changed what the colour means.
//  * Three type sizes, no others: 18/600 section heading, 13/400 value,
//    11/400 label. Nothing is left at the WinForms default -- that is what made
//    the checkbox labels shout next to everything else.
//  * Controls are chosen for what they say: a segmented control for "one of
//    three", a drawn checkbox for a switch, cards for the target list (name
//    AND what it points at), chips for the path. All drawn from the main
//    screen's skin, so selection means the same thing here as there
//    (accent-100 fill, 2 px accent bar).
//  * One section at a time, chosen in a left rail. Everything laid out flat
//    was the real problem: no hierarchy, so every field competed with the rest.
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
// shared measurements and painting
// ---------------------------------------------------------------------------
internal static class Rdv3Dlg
{
    public static int P(double css) { return Rdv3Skin.P(css); }

    public const double LabelH = 15;        // the 11 px caption
    public const double BoxH = 28;          // the 13 px value box
    public const double FieldH = 49;        // caption + box + the gap below

    public static void Caption(Graphics g, string s, int x, int y, int w)
    {
        Rdv3Skin.DrawIn(g, s, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(x, y, w, P(LabelH)), false);
    }

    public static void Heading(Graphics g, string s, int x, int y, int w)
    {
        Rdv3Skin.DrawIn(g, s, Rdv3Skin.S(18), Rdv3Skin.Ink, new Rectangle(x, y, w, P(24)), false);
    }

    public static void SubHeading(Graphics g, string s, int x, int y, int w)
    {
        Rdv3Skin.DrawIn(g, s, Rdv3Skin.S(13), Rdv3Skin.Accent700, new Rectangle(x, y, w, P(18)), false);
        Rdv3Skin.Line(g, Rdv3Skin.Divider, x, y + P(22), w, Rdv3Skin.Hair());
    }

    public static void FieldFrame(Graphics g, Control c)
    {
        int p = P(6);
        Rectangle r = new Rectangle(c.Left - p, c.Top - p, c.Width + 2 * p, c.Height + 2 * p);
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Surface)) { g.FillRectangle(b, r); }
        Rdv3Skin.Frame(g, c.Focused ? Rdv3Skin.Accent : Rdv3Skin.Divider, r);
    }

    public static void Chip(Graphics g, string s, Rectangle r, bool selected)
    {
        using (SolidBrush b = new SolidBrush(selected ? Rdv3Skin.Accent100 : Rdv3Skin.N100))
        {
            g.FillRectangle(b, r);
        }
        if (selected) { Rdv3Skin.Frame(g, Rdv3Skin.Accent, r); }
        Rdv3Skin.DrawIn(g, s, Rdv3Skin.F(11, FontStyle.Regular),
            selected ? Rdv3Skin.Accent800 : Rdv3Skin.N800,
            new Rectangle(r.X + P(8), r.Y, r.Width - P(16), r.Height), false);
    }
}

// a text field that looks like the search box on the main screen
internal sealed class Rdv3Text1 : TextBox
{
    public Rdv3Text1()
    {
        BorderStyle = BorderStyle.None;
        BackColor = Rdv3Skin.Surface;
        ForeColor = Rdv3Skin.Ink;
        Font = Rdv3Skin.F(13, FontStyle.Regular);
    }

    public Rdv3Text1 Place(Control parent, int x, int y, int w)
    {
        int h = PreferredHeight;
        SetBounds(x + Rdv3Dlg.P(6), y + (Rdv3Dlg.P(Rdv3Dlg.BoxH) - h) / 2, w - Rdv3Dlg.P(12), h);
        parent.Controls.Add(this);
        return this;
    }
}

// a switch, drawn: a 15 px box with an accent tick and a 13 px label
internal sealed class Rdv3Check : Control
{
    private bool over;
    public bool Checked;
    public Action OnToggle;

    public Rdv3Check(string text)
    {
        Text = text;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
        Font = Rdv3Skin.F(13, FontStyle.Regular);
        Cursor = Cursors.Hand;
    }

    protected override void OnMouseEnter(EventArgs e) { over = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { over = false; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        Checked = !Checked;
        Invalidate();
        if (OnToggle != null) { OnToggle(); }
        base.OnMouseClick(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(BackColor)) { g.FillRectangle(b, ClientRectangle); }
        int s = Rdv3Skin.P(15);
        Rectangle box = new Rectangle(0, (Height - s) / 2, s, s);
        using (SolidBrush b = new SolidBrush(Checked ? Rdv3Skin.Accent : Rdv3Skin.Surface))
        {
            g.FillRectangle(b, box);
        }
        Rdv3Skin.Frame(g, Checked ? Rdv3Skin.Accent
            : (over ? Rdv3Skin.Mix(Rdv3Skin.Surface, 0.45) : Rdv3Skin.Divider), box);
        if (Checked)
        {
            using (Pen p = new Pen(Rdv3Skin.Bg, Math.Max(1f, s * 0.15f)))
            {
                p.StartCap = LineCap.Round;
                p.EndCap = LineCap.Round;
                g.DrawLines(p, new Point[] {
                    new Point(box.X + (int)(s * 0.24), box.Y + (int)(s * 0.52)),
                    new Point(box.X + (int)(s * 0.44), box.Y + (int)(s * 0.72)),
                    new Point(box.X + (int)(s * 0.78), box.Y + (int)(s * 0.28)) });
            }
        }
        Rdv3Skin.DrawIn(g, Text, Font, Rdv3Skin.Ink,
            new Rectangle(box.Right + Rdv3Skin.P(8), 0, Width - box.Width - Rdv3Skin.P(8), Height), false);
    }
}

// one of N, as the reference's segmented control
internal sealed class Rdv3Seg : Control
{
    private readonly string[] items;
    private int hot = -1;
    public int Index;
    public Action OnPick;

    public Rdv3Seg(string[] options)
    {
        items = options;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
        Font = Rdv3Skin.F(13, FontStyle.Regular);
        Cursor = Cursors.Hand;
    }

    private int At(int x)
    {
        int w = Width / Math.Max(1, items.Length);
        int i = (w > 0) ? (x / w) : 0;
        return (i >= 0 && i < items.Length) ? i : -1;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        int i = At(e.X);
        if (i != hot) { hot = i; Invalidate(); }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e) { hot = -1; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        int i = At(e.X);
        if (i >= 0) { Index = i; Invalidate(); if (OnPick != null) { OnPick(); } }
        base.OnMouseClick(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        int w = Width / Math.Max(1, items.Length);
        for (int i = 0; i < items.Length; i++)
        {
            Rectangle r = new Rectangle(i * w, 0, (i == items.Length - 1) ? (Width - i * w) : w, Height);
            Color back = (i == Index) ? Rdv3Skin.Accent : ((i == hot) ? Rdv3Skin.Mix(0.07) : Rdv3Skin.Bg);
            using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, r); }
            Rdv3Skin.DrawIn(g, items[i], Font, (i == Index) ? Rdv3Skin.Bg : Rdv3Skin.Ink, r, false);
            if (i > 0) { Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, Rdv3Skin.Hair(), r.Height); }
        }
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }
}

// the target list: each row says what it is AND what it points at
internal sealed class Rdv3Cards : Control
{
    public List<string[]> Items = new List<string[]>();   // { title, summary }
    public int Index;
    private int hot = -1;
    private int top;
    public Action OnPick;

    public Rdv3Cards()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Surface;
    }

    public int RowH { get { return Rdv3Skin.P(46); } }

    private int At(int y)
    {
        int i = top + y / Math.Max(1, RowH);
        return (i >= 0 && i < Items.Count) ? i : -1;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        int i = At(e.Y);
        if (i != hot) { hot = i; Invalidate(); }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e) { hot = -1; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        int i = At(e.Y);
        if (i >= 0) { Index = i; Invalidate(); if (OnPick != null) { OnPick(); } }
        base.OnMouseClick(e);
    }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        int vis = Math.Max(1, Height / RowH);
        int max = Math.Max(0, Items.Count - vis);
        top = Math.Max(0, Math.Min(max, top + ((e.Delta > 0) ? -1 : 1)));
        Invalidate();
        base.OnMouseWheel(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Surface)) { g.FillRectangle(b, ClientRectangle); }
        Font ft = Rdv3Skin.F(13, FontStyle.Bold);
        Font fs = Rdv3Skin.F(11, FontStyle.Regular);
        int y = 0;
        for (int i = top; i < Items.Count && y < Height; i++)
        {
            Rectangle r = new Rectangle(0, y, Width, RowH);
            if (i == Index)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent100)) { g.FillRectangle(b, r); }
                Rdv3Skin.Line(g, Rdv3Skin.Accent, 0, y, Rdv3Skin.P(2), RowH);
            }
            else if (i == hot)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Mix(Rdv3Skin.Surface, 0.05)))
                {
                    g.FillRectangle(b, r);
                }
            }
            int x = Rdv3Skin.P(12);
            Rdv3Skin.DrawIn(g, Items[i][0], ft, Rdv3Skin.Ink,
                new Rectangle(x, y + Rdv3Skin.P(5), Width - x - Rdv3Skin.P(10), Rdv3Skin.P(18)), false);
            Rdv3Skin.DrawIn(g, Items[i][1], fs, Rdv3Skin.N700,
                new Rectangle(x, y + Rdv3Skin.P(24), Width - x - Rdv3Skin.P(10), Rdv3Skin.P(16)), false);
            if (i + 1 < Items.Count)
            {
                Rdv3Skin.Line(g, Rdv3Skin.RowLine, x, y + RowH - Rdv3Skin.Hair(), Width - x, Rdv3Skin.Hair());
            }
            y += RowH;
        }
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }
}

// the left rail: which part of the settings is on screen
internal sealed class Rdv3Rail : Control
{
    public string[] Items = new string[0];
    public int Index;
    private int hot = -1;
    public Action OnPick;

    public Rdv3Rail()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
    }

    private int RowH { get { return Rdv3Skin.P(42); } }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        int i = e.Y / Math.Max(1, RowH);
        if (i >= Items.Length) { i = -1; }
        if (i != hot) { hot = i; Invalidate(); }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e) { hot = -1; Invalidate(); base.OnMouseLeave(e); }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        int i = e.Y / Math.Max(1, RowH);
        if (i >= 0 && i < Items.Length) { Index = i; Invalidate(); if (OnPick != null) { OnPick(); } }
        base.OnMouseClick(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        Font f = Rdv3Skin.S(13);
        for (int i = 0; i < Items.Length; i++)
        {
            Rectangle r = new Rectangle(0, i * RowH, Width, RowH);
            if (i == Index)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent100)) { g.FillRectangle(b, r); }
                Rdv3Skin.Line(g, Rdv3Skin.Accent, 0, r.Y, Rdv3Skin.P(2), r.Height);
            }
            else if (i == hot)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Mix(0.05))) { g.FillRectangle(b, r); }
            }
            Rdv3Skin.DrawIn(g, Items[i], f, (i == Index) ? Rdv3Skin.Accent800 : Rdv3Skin.Ink,
                new Rectangle(Rdv3Skin.P(18), r.Y, r.Width - Rdv3Skin.P(26), r.Height), false);
        }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, Width - Rdv3Skin.Hair(), 0, Rdv3Skin.Hair(), Height);
    }
}

// ---------------------------------------------------------------------------
// a modal dressed like the main screen
// ---------------------------------------------------------------------------
public abstract class Rdv3Dialog : Form
{
    protected Rdv3Dialog()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Rdv3Skin.Bg;
        AutoScaleMode = AutoScaleMode.None;
        Font = Rdv3Skin.F(13, FontStyle.Regular);
    }

    protected Rdv3Btn Btn(int x, int y, int w, string text, bool primary, int icon)
    {
        Rdv3Btn b = new Rdv3Btn();
        b.Text = text;
        b.Primary = primary;
        b.Icon = icon;
        b.Font = Rdv3Skin.S(14);
        b.SetBounds(x, y, w, Rdv3Skin.P(31.7));
        Controls.Add(b);
        return b;
    }
}

// ---------------------------------------------------------------------------
// the settings dialog
// ---------------------------------------------------------------------------
public sealed class Rdv3SettingsForm : Rdv3Dialog
{
    private readonly Rdv3Config cfg;
    private readonly Rdv3Rail rail = new Rdv3Rail();
    private readonly Rdv3Cards cards = new Rdv3Cards();

    private Rdv3Text1 tName, tWinId, tWinClass, tWinLike, tWinProc;
    private Rdv3Text1 tFldId, tFldClass, tFldTypes, tFldIndex;
    private Rdv3Text1 tKeyLen, tPoll, tStable, tRebind, tCand;
    private Rdv3Text1 tData, tLedger, tLog;
    private Rdv3Check cEnabled, cValuePattern, cDigits, cFocus;
    private Rdv3Seg segRead, segScope;
    private Rdv3Btn btnAdd, btnCopy, btnDel, btnPick, btnStepDel, btnSave, btnCancel;
    private readonly List<Control> pageTargets = new List<Control>();
    private readonly List<Control> pageBehaviour = new List<Control>();
    private readonly List<Control> pageFiles = new List<Control>();
    private bool loading;
    private int stepSel = -1;
    private Rectangle rChips;

    public Rdv3Config Result;

    private const int Wide = 1060;
    private const int High = 640;
    private const int RailW = 180;
    private const int FootH = 56;

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
        ClientSize = new Size(Rdv3Dlg.P(Wide), Rdv3Dlg.P(High));
        Build();
        LoadAll();
        ShowPage(0);
    }

    private int CX { get { return Rdv3Dlg.P(RailW) + Rdv3Dlg.P(24); } }
    private int ContentW { get { return ClientSize.Width - CX - Rdv3Dlg.P(24); } }

    private void Build()
    {
        rail.Items = new string[] { Rdv3Text.SecTargets, Rdv3Text.SecBehaviour, Rdv3Text.SecPaths };
        rail.SetBounds(0, 0, Rdv3Dlg.P(RailW), ClientSize.Height - Rdv3Dlg.P(FootH));
        rail.OnPick = delegate { ShowPage(rail.Index); };
        Controls.Add(rail);

        BuildTargets();
        BuildBehaviour();
        BuildFiles();

        int fy = ClientSize.Height - Rdv3Dlg.P(FootH) + Rdv3Dlg.P(12);
        int bw = Rdv3Dlg.P(140);
        btnCancel = Btn(ClientSize.Width - Rdv3Dlg.P(24) - bw, fy, bw, Rdv3Text.BtnCancel, false, 0);
        btnCancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        btnSave = Btn(btnCancel.Left - Rdv3Dlg.P(10) - bw, fy, bw, Rdv3Text.BtnSave, true, 2);
        btnSave.Click += delegate { SaveAndClose(); };
        CancelButton = btnCancel;
    }

    // ---- page 1: what to watch ---------------------------------------------
    private void BuildTargets()
    {
        int x = CX, y = Rdv3Dlg.P(64);
        int listW = Rdv3Dlg.P(300);
        cards.SetBounds(x, y, listW, Rdv3Dlg.P(276));
        cards.OnPick = delegate { LoadTarget(); };
        Add(pageTargets, cards);

        int by = cards.Bottom + Rdv3Dlg.P(12);
        int bw = (listW - Rdv3Dlg.P(12)) / 3;
        btnAdd = Btn(x, by, bw, Rdv3Text.BtnAdd, false, 0);
        btnAdd.Click += delegate { AddTarget(); };
        btnCopy = Btn(x + bw + Rdv3Dlg.P(6), by, bw, Rdv3Text.BtnCopy, false, 0);
        btnCopy.Click += delegate { CopyTarget(); };
        btnDel = Btn(x + 2 * (bw + Rdv3Dlg.P(6)), by, bw, Rdv3Text.BtnRemove, false, 0);
        btnDel.Click += delegate { RemoveTarget(); };
        Add(pageTargets, btnAdd); Add(pageTargets, btnCopy); Add(pageTargets, btnDel);

        int ex = x + listW + Rdv3Dlg.P(28);
        int ew = ClientSize.Width - Rdv3Dlg.P(24) - ex;
        int col = (ew - Rdv3Dlg.P(20)) / 2;
        int c2 = ex + col + Rdv3Dlg.P(20);
        int ey = y;

        tName = Place(pageTargets, new Rdv3Text1(), ex, ey, col);
        cEnabled = Chk(pageTargets, Rdv3Text.LblEnabled, c2, ey + Rdv3Dlg.P(16), col);
        ey += Rdv3Dlg.P(Rdv3Dlg.FieldH) + Rdv3Dlg.P(14);          // + the subheading below

        ey += Rdv3Dlg.P(26);
        tWinId = Place(pageTargets, new Rdv3Text1(), ex, ey, col);
        tWinClass = Place(pageTargets, new Rdv3Text1(), c2, ey, col);
        ey += Rdv3Dlg.P(Rdv3Dlg.FieldH);
        tWinLike = Place(pageTargets, new Rdv3Text1(), ex, ey, col);
        tWinProc = Place(pageTargets, new Rdv3Text1(), c2, ey, col);
        ey += Rdv3Dlg.P(Rdv3Dlg.FieldH) + Rdv3Dlg.P(14);

        ey += Rdv3Dlg.P(26);
        tFldId = Place(pageTargets, new Rdv3Text1(), ex, ey, col);
        tFldClass = Place(pageTargets, new Rdv3Text1(), c2, ey, col);
        ey += Rdv3Dlg.P(Rdv3Dlg.FieldH);
        tFldTypes = Place(pageTargets, new Rdv3Text1(), ex, ey, col);
        tFldIndex = Place(pageTargets, new Rdv3Text1(), c2, ey, Rdv3Dlg.P(70));
        cValuePattern = Chk(pageTargets, "ValuePattern", c2 + Rdv3Dlg.P(84), ey + Rdv3Dlg.P(4), col - Rdv3Dlg.P(84));
        ey += Rdv3Dlg.P(Rdv3Dlg.FieldH);

        segRead = new Rdv3Seg(new string[] { "value", "text", "name" });
        segRead.SetBounds(ex + Rdv3Dlg.P(6), ey + Rdv3Dlg.P(15), Rdv3Dlg.P(210), Rdv3Dlg.P(28));
        segRead.OnPick = delegate { StoreTarget(); };
        Add(pageTargets, segRead);
        segScope = new Rdv3Seg(new string[] { Rdv3Text.LblScopeDesc, Rdv3Text.LblScopeChildren });
        segScope.SetBounds(c2 + Rdv3Dlg.P(6), ey + Rdv3Dlg.P(15), Rdv3Dlg.P(210), Rdv3Dlg.P(28));
        segScope.OnPick = delegate { StoreTarget(); };
        Add(pageTargets, segScope);
        ey += Rdv3Dlg.P(Rdv3Dlg.FieldH) + Rdv3Dlg.P(6);

        rChips = new Rectangle(ex + Rdv3Dlg.P(6), ey + Rdv3Dlg.P(18), ew - Rdv3Dlg.P(12), Rdv3Dlg.P(30));
        btnStepDel = Btn(ex + ew - Rdv3Dlg.P(90), ey - Rdv3Dlg.P(6), Rdv3Dlg.P(90), Rdv3Text.BtnRemove, false, 0);
        btnStepDel.Click += delegate { RemoveStep(); };
        Add(pageTargets, btnStepDel);
        ey += Rdv3Dlg.P(58);

        btnPick = Btn(ex + Rdv3Dlg.P(6), ey, Rdv3Dlg.P(210), Rdv3Text.BtnInspect, false, 4);
        btnPick.Click += delegate { Pick(); };
        Add(pageTargets, btnPick);
    }

    // ---- page 2: how it behaves --------------------------------------------
    private void BuildBehaviour()
    {
        int x = CX, y = Rdv3Dlg.P(64);
        int col = Rdv3Dlg.P(150);
        int c2 = x + col + Rdv3Dlg.P(28);

        y += Rdv3Dlg.P(26);
        tKeyLen = Place(pageBehaviour, new Rdv3Text1(), x, y, Rdv3Dlg.P(90));
        cDigits = Chk(pageBehaviour, Rdv3Text.LblDigitsOnly, x + Rdv3Dlg.P(120), y + Rdv3Dlg.P(4), Rdv3Dlg.P(200));
        y += Rdv3Dlg.P(Rdv3Dlg.FieldH) + Rdv3Dlg.P(14);

        y += Rdv3Dlg.P(26);
        tPoll = Place(pageBehaviour, new Rdv3Text1(), x, y, Rdv3Dlg.P(110));
        tStable = Place(pageBehaviour, new Rdv3Text1(), c2, y, Rdv3Dlg.P(110));
        y += Rdv3Dlg.P(Rdv3Dlg.FieldH);
        tRebind = Place(pageBehaviour, new Rdv3Text1(), x, y, Rdv3Dlg.P(110));
        tCand = Place(pageBehaviour, new Rdv3Text1(), c2, y, Rdv3Dlg.P(110));
        y += Rdv3Dlg.P(Rdv3Dlg.FieldH);
        cFocus = Chk(pageBehaviour, Rdv3Text.LblPreferFocus, x, y + Rdv3Dlg.P(6), Rdv3Dlg.P(360));
    }

    // ---- page 3: where the files are ---------------------------------------
    private void BuildFiles()
    {
        int x = CX, y = Rdv3Dlg.P(64);
        int w = ContentW - Rdv3Dlg.P(40);
        y += Rdv3Dlg.P(26);
        tData = Place(pageFiles, new Rdv3Text1(), x, y, w);
        y += Rdv3Dlg.P(Rdv3Dlg.FieldH);
        tLedger = Place(pageFiles, new Rdv3Text1(), x, y, w);
        y += Rdv3Dlg.P(Rdv3Dlg.FieldH);
        tLog = Place(pageFiles, new Rdv3Text1(), x, y, w);
    }

    private void Add(List<Control> page, Control c)
    {
        page.Add(c);
        if (!Controls.Contains(c)) { Controls.Add(c); }
    }

    private Rdv3Text1 Place(List<Control> page, Rdv3Text1 t, int x, int y, int w)
    {
        t.Place(this, x, y + Rdv3Dlg.P(Rdv3Dlg.LabelH), w);
        t.TextChanged += delegate { StoreTarget(); };
        page.Add(t);
        return t;
    }

    private Rdv3Check Chk(List<Control> page, string text, int x, int y, int w)
    {
        Rdv3Check c = new Rdv3Check(text);
        c.SetBounds(x, y, w, Rdv3Dlg.P(26));
        c.OnToggle = delegate { StoreTarget(); Invalidate(); };
        Add(page, c);
        return c;
    }

    private void ShowPage(int i)
    {
        rail.Index = i;
        rail.Invalidate();
        Vis(pageTargets, i == 0);
        Vis(pageBehaviour, i == 1);
        Vis(pageFiles, i == 2);
        Invalidate();
    }

    private static void Vis(List<Control> page, bool on)
    {
        for (int k = 0; k < page.Count; k++) { page[k].Visible = on; }
    }

    // ---- painting ----------------------------------------------------------
    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }

        int x = CX;
        Rdv3Dlg.Heading(g, rail.Items[rail.Index], x, Rdv3Dlg.P(24), ContentW);

        if (rail.Index == 0) { PaintTargets(g, x); }
        else if (rail.Index == 1) { PaintBehaviour(g, x); }
        else { PaintFiles(g, x); }

        // the action row: surface and a hairline, NOT the status bar's navy
        int fy = ClientSize.Height - Rdv3Dlg.P(FootH);
        using (SolidBrush b = new SolidBrush(Rdv3Skin.N100))
        {
            g.FillRectangle(b, new Rectangle(0, fy, ClientSize.Width, Rdv3Dlg.P(FootH)));
        }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, 0, fy, ClientSize.Width, Rdv3Skin.Hair());
        Rdv3Skin.DrawIn(g, Rdv3Text.NoteSavedTo + "  " + cfg.SourcePath,
            Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(Rdv3Dlg.P(24), fy, btnSave.Left - Rdv3Dlg.P(40), Rdv3Dlg.P(FootH)), false);
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }

    private void PaintTargets(Graphics g, int x)
    {
        Cap(g, tName, Rdv3Text.LblName);
        int ex = tWinId.Left - Rdv3Dlg.P(6);
        int ew = ClientSize.Width - Rdv3Dlg.P(24) - ex;
        Rdv3Dlg.SubHeading(g, Rdv3Text.SecWindow, ex, tWinId.Top - Rdv3Dlg.P(46), ew);
        Cap(g, tWinId, Rdv3Text.LblAutomationId);
        Cap(g, tWinClass, Rdv3Text.LblClassName);
        Cap(g, tWinLike, Rdv3Text.LblNameLike);
        Cap(g, tWinProc, Rdv3Text.LblProcess);
        Rdv3Dlg.SubHeading(g, Rdv3Text.SecField, ex, tFldId.Top - Rdv3Dlg.P(46), ew);
        Cap(g, tFldId, Rdv3Text.LblAutomationId);
        Cap(g, tFldClass, Rdv3Text.LblClassName);
        Cap(g, tFldTypes, Rdv3Text.LblControlTypes);
        Cap(g, tFldIndex, Rdv3Text.LblIndex);
        Rdv3Dlg.Caption(g, Rdv3Text.LblRead, segRead.Left, segRead.Top - Rdv3Dlg.P(17), segRead.Width);
        Rdv3Dlg.Caption(g, Rdv3Text.LblScope, segScope.Left, segScope.Top - Rdv3Dlg.P(17), segScope.Width);
        Rdv3Dlg.Caption(g, Rdv3Text.LblPath, rChips.X, rChips.Y - Rdv3Dlg.P(17), rChips.Width - Rdv3Dlg.P(100));

        Frame(g, tName); Frame(g, tWinId); Frame(g, tWinClass); Frame(g, tWinLike); Frame(g, tWinProc);
        Frame(g, tFldId); Frame(g, tFldClass); Frame(g, tFldTypes); Frame(g, tFldIndex);

        // the path, as chips
        Rdv3Target t = Current;
        int cx = rChips.X;
        if (t == null || t.Steps.Count == 0)
        {
            Rdv3Skin.DrawIn(g, "-", Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N400,
                new Rectangle(cx, rChips.Y, rChips.Width, rChips.Height), false);
        }
        else
        {
            for (int i = 0; i < t.Steps.Count; i++)
            {
                string s = Chip(t.Steps[i]);
                int w = Rdv3Skin.Measure(s, Rdv3Skin.F(11, FontStyle.Regular)).Width + Rdv3Skin.P(20);
                if (cx + w > rChips.Right) { break; }
                Rdv3Dlg.Chip(g, s, new Rectangle(cx, rChips.Y, w, Rdv3Skin.P(24)), i == stepSel);
                cx += w + Rdv3Skin.P(8);
            }
        }
    }

    private static string Chip(Rdv3Match m)
    {
        if (m.AutomationId.Length > 0) { return "#" + m.AutomationId; }
        if (m.ClassName.Length > 0) { return m.ClassName; }
        if (m.Name.Length > 0) { return m.Name; }
        return m.Describe();
    }

    private void PaintBehaviour(Graphics g, int x)
    {
        Rdv3Dlg.SubHeading(g, Rdv3Text.SecKey, x, tKeyLen.Top - Rdv3Dlg.P(46), ContentW - Rdv3Dlg.P(40));
        Cap(g, tKeyLen, Rdv3Text.LblKeyLen);
        Unit(g, tKeyLen, Rdv3Text.LblUnitDigits);
        Rdv3Dlg.SubHeading(g, Rdv3Text.SecWatch, x, tPoll.Top - Rdv3Dlg.P(46), ContentW - Rdv3Dlg.P(40));
        Cap(g, tPoll, Rdv3Text.LblPollMs);
        Cap(g, tStable, Rdv3Text.LblStableMs);
        Cap(g, tRebind, Rdv3Text.LblRebindMs);
        Cap(g, tCand, Rdv3Text.LblCandRows);
        Unit(g, tCand, Rdv3Text.LblUnitRows);
        Frame(g, tKeyLen); Frame(g, tPoll); Frame(g, tStable); Frame(g, tRebind); Frame(g, tCand);
    }

    private void PaintFiles(Graphics g, int x)
    {
        Cap(g, tData, Rdv3Text.LblDataDir);
        Cap(g, tLedger, Rdv3Text.LblLedger);
        Cap(g, tLog, Rdv3Text.LblLog);
        Frame(g, tData); Frame(g, tLedger); Frame(g, tLog);
        Rdv3Skin.DrawIn(g, Rdv3Text.NoteRestart, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(tLog.Left - Rdv3Dlg.P(6), tLog.Bottom + Rdv3Dlg.P(20), ContentW, Rdv3Dlg.P(18)), false);
    }

    private void Cap(Graphics g, Control c, string s)
    {
        Rdv3Dlg.Caption(g, s, c.Left - Rdv3Dlg.P(6), c.Top - Rdv3Dlg.P(21), c.Width + Rdv3Dlg.P(12));
    }

    private void Unit(Graphics g, Control c, string s)
    {
        Rdv3Skin.DrawIn(g, s, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(c.Right + Rdv3Dlg.P(12), c.Top, Rdv3Dlg.P(40), c.Height), false);
    }

    // NOT gated on Control.Visible: a form that has never been shown reports
    // every child as invisible, and the frames would vanish from the acceptance
    // renders. The page paint methods already draw only their own page.
    private void Frame(Graphics g, Control c)
    {
        Rdv3Dlg.FieldFrame(g, c);
    }

    // ---- data --------------------------------------------------------------
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
        loading = true;
        RefreshCards();
        tKeyLen.Text = N(cfg.KeyLength);
        cDigits.Checked = cfg.KeyDigitsOnly;
        tPoll.Text = N(cfg.PollMs);
        tStable.Text = N(cfg.StableMs);
        tRebind.Text = N(cfg.RebindMs);
        tCand.Text = N(cfg.CandidateRowsShown);
        cFocus.Checked = cfg.PreferFocusedWindow;
        tData.Text = cfg.DataDir;
        tLedger.Text = cfg.Ledger;
        tLog.Text = cfg.Log;
        loading = false;
        LoadTarget();
    }

    private static string N(int v) { return v.ToString(CultureInfo.InvariantCulture); }

    private void RefreshCards()
    {
        cards.Items.Clear();
        for (int i = 0; i < cfg.Targets.Count; i++)
        {
            Rdv3Target t = cfg.Targets[i];
            string title = t.Name;
            if (!t.Enabled) { title = title + "   (" + Rdv3Text.LblEnabled + " -)"; }
            cards.Items.Add(new string[] { title, Summary(t) });
        }
        if (cards.Index >= cfg.Targets.Count) { cards.Index = Math.Max(0, cfg.Targets.Count - 1); }
        cards.Invalidate();
    }

    private static string Summary(Rdv3Target t)
    {
        string win = t.Window.ProcessName.Length > 0 ? t.Window.ProcessName
            : (t.Window.ClassName.Length > 0 ? t.Window.ClassName : t.Window.NameLike);
        string fld = t.Field.AutomationId.Length > 0 ? ("#" + t.Field.AutomationId)
            : (t.Field.ControlTypes.Length > 0 ? string.Join("/", t.Field.ControlTypes) : t.Field.ClassName);
        if (win.Length == 0) { win = "-"; }
        if (fld.Length == 0) { fld = "-"; }
        return Rdv3Text.NoteTargetSummary.Replace("{win}", win).Replace("{field}", fld);
    }

    private void LoadTarget()
    {
        Rdv3Target t = Current;
        loading = true;
        stepSel = (t != null && t.Steps.Count > 0) ? 0 : -1;
        if (t == null)
        {
            tName.Text = ""; cEnabled.Checked = false;
            tWinId.Text = ""; tWinClass.Text = ""; tWinLike.Text = ""; tWinProc.Text = "";
            tFldId.Text = ""; tFldClass.Text = ""; tFldTypes.Text = ""; tFldIndex.Text = "0";
            cValuePattern.Checked = false;
            segRead.Index = 0; segScope.Index = 0;
        }
        else
        {
            tName.Text = t.Name;
            cEnabled.Checked = t.Enabled;
            tWinId.Text = t.Window.AutomationId;
            tWinClass.Text = t.Window.ClassName;
            tWinLike.Text = t.Window.NameLike;
            tWinProc.Text = t.Window.ProcessName;
            tFldId.Text = t.Field.AutomationId;
            tFldClass.Text = t.Field.ClassName;
            tFldTypes.Text = string.Join(",", t.Field.ControlTypes);
            tFldIndex.Text = N(t.Field.Index);
            cValuePattern.Checked = t.Field.RequireValuePattern;
            segRead.Index = t.ReadMode;
            segScope.Index = t.Field.Descendants ? 0 : 1;
        }
        loading = false;
        cEnabled.Invalidate(); cValuePattern.Invalidate();
        segRead.Invalidate(); segScope.Invalidate();
        Invalidate();
    }

    private void StoreTarget()
    {
        if (loading) { return; }
        Rdv3Target t = Current;
        if (t == null) { return; }
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
        t.Field.RequireValuePattern = cValuePattern.Checked;
        t.Field.Descendants = (segScope.Index == 0);
        t.ReadMode = segRead.Index;
        RefreshCards();
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
        Rdv3Target t = new Rdv3Target();
        t.Name = Rdv3Text.SecTarget + " " + N(cfg.Targets.Count + 1);
        t.Field.RequireValuePattern = true;
        cfg.Targets.Add(t);
        RefreshCards();
        cards.Index = cfg.Targets.Count - 1;
        LoadTarget();
    }

    private void CopyTarget()
    {
        Rdv3Target t = Current;
        if (t == null) { return; }
        cfg.Targets.Add(t.Clone());
        RefreshCards();
        cards.Index = cfg.Targets.Count - 1;
        LoadTarget();
    }

    private void RemoveTarget()
    {
        int i = cards.Index;
        if (i < 0 || i >= cfg.Targets.Count) { return; }
        cfg.Targets.RemoveAt(i);
        cards.Index = Math.Max(0, Math.Min(i, cfg.Targets.Count - 1));
        RefreshCards();
        LoadTarget();
    }

    private void RemoveStep()
    {
        Rdv3Target t = Current;
        if (t == null || stepSel < 0 || stepSel >= t.Steps.Count) { return; }
        t.Steps.RemoveAt(stepSel);
        stepSel = (t.Steps.Count > 0) ? 0 : -1;
        RefreshCards();
        Invalidate();
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        // the chips are painted, so their hit test lives here
        Rdv3Target t = Current;
        if (rail.Index == 0 && t != null && rChips.Contains(e.Location))
        {
            int cx = rChips.X;
            for (int i = 0; i < t.Steps.Count; i++)
            {
                string s = Chip(t.Steps[i]);
                int w = Rdv3Skin.Measure(s, Rdv3Skin.F(11, FontStyle.Regular)).Width + Rdv3Skin.P(20);
                if (e.X >= cx && e.X < cx + w) { stepSel = i; Invalidate(); break; }
                cx += w + Rdv3Skin.P(8);
            }
        }
        base.OnMouseClick(e);
    }

    private void Pick()
    {
        Rdv3Target picked = Rdv3PickerForm.Pick(this);
        if (picked == null) { return; }
        Rdv3Target t = Current;
        if (t == null)
        {
            cfg.Targets.Add(picked);
            RefreshCards();
            cards.Index = cfg.Targets.Count - 1;
        }
        else
        {
            if (t.Name.Length > 0 && !t.Name.StartsWith(Rdv3Text.SecTarget)) { picked.Name = t.Name; }
            picked.Enabled = t.Enabled;
            cfg.Targets[cards.Index] = picked;
            RefreshCards();
        }
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
// the picker: hover over the field, take it with Ctrl+Shift
// ---------------------------------------------------------------------------
public sealed class Rdv3PickerForm : Rdv3Dialog
{
    private readonly System.Windows.Forms.Timer tick = new System.Windows.Forms.Timer();
    private AutomationElement hover;
    private string ctrlType = "", autoId = "", className = "", elName = "", value = "", proc = "";
    private bool armed;
    private Rdv3Target result;

    public static Rdv3Target Pick(IWin32Window owner)
    {
        Rdv3PickerForm f = new Rdv3PickerForm();
        DialogResult dr = f.ShowDialog(owner);
        Rdv3Target t = (dr == DialogResult.OK) ? f.result : null;
        f.Dispose();
        return t;
    }

    private Rdv3PickerForm()
    {
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.PickTitle;
        ClientSize = new Size(Rdv3Skin.P(440), Rdv3Skin.P(250));
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        KeyPreview = true;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Location = new Point(wa.Right - Width - Rdv3Skin.P(24), wa.Bottom - Height - Rdv3Skin.P(24));
        int by = ClientSize.Height - Rdv3Skin.P(44);
        Rdv3Btn close = Btn(ClientSize.Width - Rdv3Skin.P(16) - Rdv3Skin.P(110), by, Rdv3Skin.P(110),
            Rdv3Text.BtnClose, false, 0);
        close.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        CancelButton = close;
        tick.Interval = 90;
        tick.Tick += delegate { Look(); };
        tick.Start();
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        tick.Stop();
        base.OnFormClosed(e);
    }

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
        base.OnKeyDown(e);
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
            if (v == null) { v = null; }
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

            result = t;
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception) { }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        int pad = Rdv3Skin.P(16);
        int w = ClientSize.Width - pad * 2;

        Rdv3Dlg.Heading(g, Rdv3Text.PickTitle, pad, Rdv3Skin.P(14), w);
        Rdv3Skin.DrawIn(g, Rdv3Text.PickHow, Rdv3Skin.F(12, FontStyle.Regular), Rdv3Skin.Accent700,
            new Rectangle(pad, Rdv3Skin.P(42), w, Rdv3Skin.P(18)), false);
        Rdv3Skin.Line(g, Rdv3Skin.Divider, pad, Rdv3Skin.P(68), w, Rdv3Skin.Hair());

        int y = Rdv3Skin.P(78);
        Row(g, pad, ref y, w, Rdv3Text.LblControlTypes, ctrlType);
        Row(g, pad, ref y, w, Rdv3Text.LblAutomationId, (autoId.Length > 0) ? autoId : "-");
        Row(g, pad, ref y, w, Rdv3Text.LblClassName, className);
        Row(g, pad, ref y, w, Rdv3Text.LblName, elName);
        Row(g, pad, ref y, w, Rdv3Text.PickReading, (value.Length > 0) ? value : Rdv3Text.PickNoRead);

        Rdv3Skin.DrawIn(g, Rdv3Text.PickEsc, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(pad, ClientSize.Height - Rdv3Skin.P(38), Rdv3Skin.P(200), Rdv3Skin.P(20)), false);
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }

    private static void Row(Graphics g, int x, ref int y, int w, string k, string v)
    {
        int kw = Rdv3Skin.P(120);
        Rdv3Skin.DrawIn(g, k, Rdv3Skin.F(11, FontStyle.Regular), Rdv3Skin.N700,
            new Rectangle(x, y, kw, Rdv3Skin.P(18)), false);
        Rdv3Skin.DrawIn(g, (v == null || v.Length == 0) ? "-" : v,
            Rdv3Skin.F(12, (k == Rdv3Text.LblAutomationId) ? FontStyle.Bold : FontStyle.Regular),
            (k == Rdv3Text.LblAutomationId) ? Rdv3Skin.Ink : Rdv3Skin.Ink,
            new Rectangle(x + kw, y, w - kw, Rdv3Skin.P(18)), false);
        y += Rdv3Skin.P(24);
    }
}
