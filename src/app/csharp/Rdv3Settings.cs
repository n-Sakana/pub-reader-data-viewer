// ============================================================================
// Rdv3Settings.cs -- the settings modal and the element picker, rebuilt to the
// reference artifact "Reader Data Viewer_ver4.html".
//
// Same method as the main screen: the modal was opened in a browser, every
// element's rectangle, font, colour and border was read out numerically, the
// table went into Rdv3SetGeom.cs and docs\ui-ref-settings-geom.json, and THIS
// file is laid out from that table. build\test_settings_geometry.ps1 renders
// the dialog without ever showing it and compares the two, element by element.
//
// What the reference decided, and this file therefore keeps:
//
//  * A HEADER / TABS / BODY / FOOTER modal, not a left rail. Three tabs
//    (targets, behaviour, files) with a 2 px accent underline on the current
//    one; the file being edited is named quietly at the right of the tab row.
//  * The navy status bar is NOT reused: the footer is N100 with a hairline.
//  * Section headings are 10 px accent with a rule under them and they carry
//    the question they answer ("window -- which screen"), so the heading is
//    worth reading. Values are 13 px, captions 11 px. Three sizes, no others.
//  * The target list is a framed box of cards: name, an enabled/disabled tag,
//    and what the target actually points at. Selection is accent-100 with a
//    2 px accent bar, exactly as a selected row means on the main screen.
//  * One-of-N is a segmented control, a switch is a drawn checkbox, and the
//    path steps are a row of their own with a quiet "delete" at the end.
//
// The modal is FULL at 1060 x 764 -- both columns run to the last pixel of the
// body -- so it does not stretch. When the work area cannot hold the design at
// the current Windows scaling, the WHOLE dialog scales down (FitScale) and the
// proportions survive; nothing is re-arranged and nothing is dropped.
//
// THE PICKER. Naming a field inside a business application by typing
// identifiers is hopeless. So: hover over the field, read its identity live,
// take it with Ctrl+Shift (or the button). UI Automation the whole way --
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
// a modal dressed like the main screen, laid out from a measured table
// ---------------------------------------------------------------------------
public abstract class Rdv3Dialog : Form
{
    // the dialog's own scale: the screen's, unless the design does not fit
    protected float Sc = 1f;

    private readonly Dictionary<string, Rectangle> rc = new Dictionary<string, Rectangle>();
    private readonly List<string> clip = new List<string>();

    // The dialog carries its OWN chrome, so it has no system frame.
    //
    // The reference is a modal inside a browser page: no title bar exists, so
    // the card draws its own header -- title, subtitle and the dismiss mark.
    // Put that on a FixedDialog and Windows adds a second title with a second
    // close button on top of it, which is what shipped and was wrong. The
    // header IS the title bar: one title, one close, and the card is dragged
    // by the header the way a title bar drags a window.
    private bool dragging;
    private Point dragFrom;

    // the part of the client area that behaves like a title bar (set by the
    // dialog once it knows its own layout)
    protected Rectangle DragZone = Rectangle.Empty;

    protected Rdv3Dialog()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        FormBorderStyle = FormBorderStyle.None;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.CenterParent;
        BackColor = Rdv3Skin.Bg;
        AutoScaleMode = AutoScaleMode.None;
        KeyPreview = true;
    }

    // CS_DROPSHADOW: the reference's "0 12px 32px rgba(43,43,45,.22)", drawn by
    // the window manager. A managed class style, not an API call.
    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams cp = base.CreateParams;
            cp.ClassStyle |= 0x00020000;
            return cp;
        }
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (e.Button == MouseButtons.Left && DragZone.Contains(e.Location))
        {
            dragging = true;
            dragFrom = e.Location;
        }
        base.OnMouseDown(e);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (dragging)
        {
            Location = new Point(Location.X + e.X - dragFrom.X, Location.Y + e.Y - dragFrom.Y);
        }
        base.OnMouseMove(e);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        dragging = false;
        base.OnMouseUp(e);
    }

    // The design is a fixed shape. If the screen's scaling would push it past
    // the work area the WHOLE dialog is scaled down instead: proportions and
    // reading order survive, only the type gets smaller. There is no system
    // frame to allow for -- the client area is the whole window.
    protected static float FitScale(double cssW, double cssH)
    {
        float sc = Rdv3Skin.Scale;
        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        double mw = Math.Max(360.0, wa.Width - 24.0);
        double mh = Math.Max(280.0, wa.Height - 24.0);
        if (cssW * sc > mw) { sc = (float)(mw / cssW); }
        if (cssH * sc > mh) { sc = (float)(mh / cssH); }
        if (sc > Rdv3Skin.Scale) { sc = Rdv3Skin.Scale; }
        return Math.Max(0.7f, sc);
    }

    protected int PX(double css) { return (int)Math.Round(css * Sc); }

    protected Font Fn(double css) { return Rdv3Skin.F(css, FontStyle.Regular, Sc); }
    protected Font Fb(double css) { return Rdv3Skin.F(css, FontStyle.Bold, Sc); }
    protected Font Sf(double css) { return Rdv3Skin.S(css, Sc); }

    protected void ClearRects() { rc.Clear(); }

    protected void Put(string k, double x, double y, double w, double h)
    {
        int x0 = PX(x), y0 = PX(y);
        rc[k] = new Rectangle(x0, y0, PX(x + w) - x0, PX(y + h) - y0);
    }

    protected void PutR(string k, Rectangle r) { rc[k] = r; }

    protected void Put(string k, Rdv3SetGeom.R r) { Put(k, r.X, r.Y, r.W, r.H); }

    protected Rectangle At(string k)
    {
        Rectangle r;
        return rc.TryGetValue(k, out r) ? r : Rectangle.Empty;
    }

    protected bool Has(string k) { return rc.ContainsKey(k); }

    // ---- drawing, always inside the measured rectangle ---------------------
    // A string that does not fit its box is the one failure the eye forgives
    // and the operator does not, so every draw records it and the acceptance
    // dump reports the list.
    protected void ClearClip() { clip.Clear(); }

    private void Note(string k, string s, Font f, Rectangle r)
    {
        if (s == null || s.Length == 0) { return; }
        if (Rdv3Skin.Measure(s, f).Width > r.Width && !clip.Contains(k)) { clip.Add(k); }
    }

    protected void T(Graphics g, string k, string s, Font f, Color c)
    {
        Rectangle r = At(k);
        Note(k, s, f, r);
        Rdv3Skin.DrawIn(g, s, f, c, r, false);
    }

    protected void TR(Graphics g, string k, string s, Font f, Color c)
    {
        Rectangle r = At(k);
        Note(k, s, f, r);
        Rdv3Skin.DrawIn(g, s, f, c, r, true);
    }

    protected void TC(Graphics g, string k, string s, Font f, Color c)
    {
        Rectangle r = At(k);
        Note(k, s, f, r);
        if (s == null || s.Length == 0) { return; }
        Size z = Rdv3Skin.Measure(s, f);
        Rdv3Skin.Draw(g, s, f, c, r.X + (r.Width - z.Width) / 2, r.Y + (r.Height - z.Height) / 2);
    }

    // the two-line hint under the pick button. Japanese has no spaces, so the
    // break is by character width, which is what the browser does here too.
    protected void TWrap(Graphics g, string k, string s, Font f, Color c, double lineCss)
    {
        Rectangle r = At(k);
        if (s == null || s.Length == 0) { return; }
        int lh = PX(lineCss);
        // how many lines the DESIGN gives, not how many fit after two roundings:
        // integer division turned a two line box into one at 125% (41 / 21).
        int lines = Math.Max(1, (int)Math.Round(r.Height / (double)Math.Max(1, lh)));
        int at = 0, drawn = 0, y = r.Y;
        while (at < s.Length && drawn < lines)
        {
            int take = FitCount(s, at, f, r.Width);
            Rdv3Skin.Draw(g, s.Substring(at, take), f, c, r.X, y);
            at += take;
            y += lh;
            drawn++;
        }
        if (at < s.Length && !clip.Contains(k)) { clip.Add(k); }
    }

    private static int FitCount(string s, int from, Font f, int w)
    {
        int lo = 1, hi = s.Length - from;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) / 2;
            if (Rdv3Skin.Measure(s.Substring(from, mid), f).Width <= w) { lo = mid; } else { hi = mid - 1; }
        }
        return Math.Max(1, lo);
    }

    protected void Rule(Graphics g, string k)
    {
        Rectangle r = At(k);
        if (r.Width <= 0) { return; }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, r.Width, Rdv3Skin.Hair());
    }

    // a text field: the reference's surface box with a hairline, accent while
    // it has the keyboard
    protected void Field(Graphics g, string k, bool focused)
    {
        Rectangle r = At(k);
        if (r.Width <= 0) { return; }
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Surface)) { g.FillRectangle(b, r); }
        Rdv3Skin.Frame(g, focused ? Rdv3Skin.Accent : Rdv3Skin.Divider, r);
    }

    // the 10 px accent section heading with its rule
    protected void Section(Graphics g, string k, string text)
    {
        Rectangle r = At(k);
        using (Font f = Fn(10))
        {
            Note(k, text, f, r);
            Rdv3Skin.Draw(g, text, f, Rdv3Skin.Accent, r.X, r.Y);
        }
        Rule(g, k + ".rule");
    }

    protected void Chip(Graphics g, string k, string text, Color back, Color ink)
    {
        Rectangle r = At(k);
        if (r.Width <= 0) { return; }
        using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, r); }
        using (Font f = Fn(10))
        {
            Size z = Rdv3Skin.Measure(text, f);
            Rdv3Skin.Draw(g, text, f, ink, r.X + (r.Width - z.Width) / 2, r.Y + (r.Height - z.Height) / 2);
        }
    }

    // ---- the acceptance dump ----------------------------------------------
    public string GeometryDump()
    {
        List<string> keys = new List<string>(rc.Keys);
        keys.Sort(StringComparer.Ordinal);
        StringBuilder sb = new StringBuilder();
        sb.Append("{\"scale\":").Append(Sc.ToString("0.###", CultureInfo.InvariantCulture));
        sb.Append(",\"client\":[").Append(Css(ClientSize.Width)).Append(",")
          .Append(Css(ClientSize.Height)).Append("],\"el\":{");
        for (int i = 0; i < keys.Count; i++)
        {
            Rectangle r = rc[keys[i]];
            if (i > 0) { sb.Append(","); }
            sb.Append("\"").Append(keys[i]).Append("\":[")
              .Append(Css(r.X)).Append(",").Append(Css(r.Y)).Append(",")
              .Append(Css(r.Width)).Append(",").Append(Css(r.Height)).Append("]");
        }
        sb.Append("},\"clipped\":[");
        for (int i = 0; i < clip.Count; i++)
        {
            if (i > 0) { sb.Append(","); }
            sb.Append("\"").Append(clip[i]).Append("\"");
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private string Css(int devicePx)
    {
        return (devicePx / (double)Sc).ToString("0.#", CultureInfo.InvariantCulture);
    }
}

// ---------------------------------------------------------------------------
// the 13 px value box. The frame is painted by the dialog, so the control is
// the text only and sits inside the measured rectangle.
// ---------------------------------------------------------------------------
internal sealed class Rdv3SetBox : TextBox
{
    public float Sc = 1f;
    public Action OnEdit;

    public Rdv3SetBox(float sc)
    {
        Sc = sc;
        BorderStyle = BorderStyle.None;
        BackColor = Rdv3Skin.Surface;
        ForeColor = Rdv3Skin.Ink;
        Font = Rdv3Skin.F(13, FontStyle.Regular, sc);
        TextChanged += delegate { if (OnEdit != null) { OnEdit(); } };
    }

    // the reference's input padding is 6px 10px; the frame is the caller's
    public void PlaceIn(Rectangle frame)
    {
        int pad = (int)Math.Round(10.0 * Sc);
        int h = PreferredHeight;
        SetBounds(frame.X + pad, frame.Y + (frame.Height - h) / 2, Math.Max(1, frame.Width - 2 * pad), h);
    }
}

// ---------------------------------------------------------------------------
// a switch: a 15 px box with an accent tick and a 13 px label
// ---------------------------------------------------------------------------
internal sealed class Rdv3SetCheck : Control
{
    public float Sc = 1f;
    public bool Checked;
    public Action OnToggle;
    private bool over;

    public Rdv3SetCheck(string text, float sc)
    {
        Sc = sc;
        Text = text;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
        Font = Rdv3Skin.F(13, FontStyle.Regular, sc);
        Cursor = Cursors.Hand;
    }

    public int BoxSize { get { return (int)Math.Round(15.0 * Sc); } }
    public int Gap { get { return (int)Math.Round(8.0 * Sc); } }

    // the label's own box, so the dialog can size the control from the text
    public int Wanted
    {
        get { return BoxSize + Gap + Rdv3Skin.Measure(Text, Font).Width; }
    }

    public Rectangle BoxRect
    {
        get { int s = BoxSize; return new Rectangle(Left, Top + (Height - s) / 2, s, s); }
    }

    public Rectangle LabelRect
    {
        get { return new Rectangle(Left + BoxSize + Gap, Top, Width - BoxSize - Gap, Height); }
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
        int s = BoxSize;
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
            new Rectangle(box.Right + Gap, 0, Width - box.Width - Gap, Height), false);
    }
}

// ---------------------------------------------------------------------------
// one of N: the reference's segmented control, a hairline box divided evenly
// ---------------------------------------------------------------------------
internal sealed class Rdv3SetSeg : Control
{
    public float Sc = 1f;
    private readonly string[] items;
    private int hot = -1;
    public int Index;
    public Action OnPick;

    public Rdv3SetSeg(string[] options, float sc)
    {
        Sc = sc;
        items = options;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
        Font = Rdv3Skin.F(13, FontStyle.Regular, sc);
        Cursor = Cursors.Hand;
    }

    public int Count { get { return items.Length; } }

    // the option rectangles live inside the hairline, evenly divided
    public Rectangle Cell(int i)
    {
        int t = Rdv3Skin.Hair();
        int inner = Width - 2 * t;
        int x0 = t + (int)Math.Round(inner * (double)i / items.Length);
        int x1 = t + (int)Math.Round(inner * (double)(i + 1) / items.Length);
        return new Rectangle(x0, t, x1 - x0, Height - 2 * t);
    }

    private int At(int x)
    {
        for (int i = 0; i < items.Length; i++)
        {
            Rectangle c = Cell(i);
            if (x >= c.X && x < c.Right) { return i; }
        }
        return -1;
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
        if (i >= 0 && i != Index) { Index = i; Invalidate(); if (OnPick != null) { OnPick(); } }
        base.OnMouseClick(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        for (int i = 0; i < items.Length; i++)
        {
            Rectangle r = Cell(i);
            Color back = (i == Index) ? Rdv3Skin.Accent : ((i == hot) ? Rdv3Skin.Mix(0.07) : Rdv3Skin.Bg);
            using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, r); }
            Rdv3Skin.DrawIn(g, items[i], Font, (i == Index) ? Rdv3Skin.Bg : Rdv3Skin.Ink,
                Centre(r, items[i]), false);
            if (i > 0) { Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, Rdv3Skin.Hair(), r.Height); }
        }
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }

    private Rectangle Centre(Rectangle r, string s)
    {
        int w = Rdv3Skin.Measure(s, Font).Width;
        int x = r.X + Math.Max(0, (r.Width - w) / 2);
        return new Rectangle(x, r.Y, Math.Min(w, r.Width), r.Height);
    }
}

// ---------------------------------------------------------------------------
// the target list: each card says what the target is called, whether it is on,
// and what it actually points at
// ---------------------------------------------------------------------------
internal sealed class Rdv3SetCards : Control
{
    public float Sc = 1f;
    public List<string[]> Items = new List<string[]>();   // { name, tag, summary }
    public int Index;
    public Action OnPick;
    public Action OnView;
    public string Empty = "";
    private int hot = -1;
    private int top;
    private readonly List<Rectangle[]> parts = new List<Rectangle[]>();
    private Font fName, fTag, fSum, fEmpty;

    public Rdv3SetCards(float sc)
    {
        Sc = sc;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
        MakeFonts();
    }

    public void MakeFonts()
    {
        fName = Rdv3Skin.F(13, FontStyle.Bold, Sc);
        fTag = Rdv3Skin.F(10, FontStyle.Regular, Sc);
        fSum = Rdv3Skin.F(11, FontStyle.Regular, Sc);
        fEmpty = Rdv3Skin.F(11, FontStyle.Regular, Sc);
    }

    private int P(double v) { return (int)Math.Round(v * Sc); }

    public int RowH { get { return P(Rdv3SetGeom.CardH); } }
    public int Rows { get { return Math.Max(1, (Height - 2 * Rdv3Skin.Hair()) / Math.Max(1, RowH)); } }
    private bool NeedBar { get { return Items.Count > Rows; } }
    private int BarW { get { return P(10); } }

    public void ScrollToSelected()
    {
        if (Index < top) { top = Index; }
        else if (Index >= top + Rows) { top = Index - Rows + 1; }
        int max = Math.Max(0, Items.Count - Rows);
        if (top > max) { top = max; }
        if (top < 0) { top = 0; }
    }

    // the drawn rectangles, in the control's coordinates, so the dialog can
    // report them in the acceptance dump
    public void Measure()
    {
        parts.Clear();
        int t = Rdv3Skin.Hair();
        int w = Width - 2 * t - (NeedBar ? BarW : 0);
        int y = t;
        for (int i = top; i < Items.Count && y + RowH <= Height - t + 1; i++)
        {
            Rectangle card = new Rectangle(t, y, w, RowH);
            int x = card.X + P(Rdv3SetGeom.CardPadX);
            int nameW = Rdv3Skin.Measure(Items[i][0], fName).Width;
            int tagW = Rdv3Skin.Measure(Items[i][1], fTag).Width + P(12);
            // the tag keeps its room: a long name gives way (and is ellipsised)
            // rather than pushing "enabled / disabled" off the card
            int room = Math.Max(1, card.Right - P(Rdv3SetGeom.CardPadX) - x - tagW - P(Rdv3SetGeom.CardGap));
            Rectangle name = new Rectangle(x, y + P(Rdv3SetGeom.CardNameDy),
                Math.Min(nameW, room), P(Rdv3SetGeom.CardNameH));
            Rectangle tag = new Rectangle(name.Right + P(Rdv3SetGeom.CardGap),
                y + P(Rdv3SetGeom.CardTagDy), tagW, P(Rdv3SetGeom.CardTagH));
            Rectangle sum = new Rectangle(x, y + P(Rdv3SetGeom.CardSumDy),
                Math.Max(1, card.Right - P(Rdv3SetGeom.CardPadX) - x), P(Rdv3SetGeom.CardSumH));
            parts.Add(new Rectangle[] { card, name, tag, sum });
            y += RowH;
        }
    }

    public int Drawn { get { return parts.Count; } }
    public Rectangle[] Part(int i) { return (i >= 0 && i < parts.Count) ? parts[i] : null; }
    public int FirstDrawn { get { return top; } }

    private int RowAt(int y)
    {
        int t = Rdv3Skin.Hair();
        if (y < t) { return -1; }
        int i = top + (y - t) / Math.Max(1, RowH);
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
        int was = top;
        top = Math.Max(0, Math.Min(max, top + ((e.Delta > 0) ? -1 : 1)));
        if (top != was) { Measure(); if (OnView != null) { OnView(); } }
        Invalidate();
        base.OnMouseWheel(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        Measure();
        if (Items.Count == 0)
        {
            Rdv3Skin.DrawIn(g, Empty, fEmpty, Rdv3Skin.N600,
                new Rectangle(P(14), 0, Width - P(28), Math.Min(Height, P(80))), false);
            Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
            return;
        }
        for (int k = 0; k < parts.Count; k++)
        {
            int i = top + k;
            Rectangle[] p = parts[k];
            Rectangle card = p[0];
            Color back = Rdv3Skin.Bg;
            if (i == Index) { back = Rdv3Skin.Accent100; }
            else if (i == hot) { back = Rdv3Skin.Mix(0.05); }
            using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, card); }
            if (i == Index) { Rdv3Skin.Line(g, Rdv3Skin.Accent, card.X, card.Y, P(2), card.Height); }
            Rdv3Skin.DrawIn(g, Items[i][0], fName, Rdv3Skin.Ink, p[1], false);
            bool on = Items[i][1] != Rdv3Text.LblDisabled;
            using (SolidBrush b = new SolidBrush(on ? Rdv3Skin.Accent100 : Rdv3Skin.N100))
            {
                g.FillRectangle(b, p[2]);
            }
            Size tz = Rdv3Skin.Measure(Items[i][1], fTag);
            Rdv3Skin.Draw(g, Items[i][1], fTag, on ? Rdv3Skin.Accent800 : Rdv3Skin.N800,
                p[2].X + (p[2].Width - tz.Width) / 2, p[2].Y + (p[2].Height - tz.Height) / 2);
            Rdv3Skin.DrawIn(g, Items[i][2], fSum, Rdv3Skin.N700, p[3], false);
            if (i + 1 < Items.Count)
            {
                Rdv3Skin.Line(g, Rdv3Skin.Mix(back, 0.08), card.X, card.Bottom - Rdv3Skin.Hair(),
                    card.Width, Rdv3Skin.Hair());
            }
        }
        if (NeedBar) { PaintBar(g); }
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }

    private void PaintBar(Graphics g)
    {
        int t = Rdv3Skin.Hair();
        Rectangle track = new Rectangle(Width - t - BarW, t, BarW, Height - 2 * t);
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Mix(0.04))) { g.FillRectangle(b, track); }
        double frac = Rows / (double)Math.Max(1, Items.Count);
        int th = Math.Max(P(24), (int)(track.Height * frac));
        int max = Math.Max(0, Items.Count - Rows);
        int off = (max == 0) ? 0 : (int)((track.Height - th) * (top / (double)max));
        Rectangle thumb = new Rectangle(track.X + P(3), track.Y + off, BarW - P(6), th);
        using (SolidBrush b = new SolidBrush(Rdv3Skin.N400)) { g.FillRectangle(b, thumb); }
    }
}

// ---------------------------------------------------------------------------
// the settings modal
// ---------------------------------------------------------------------------
public sealed class Rdv3SettingsForm : Rdv3Dialog
{
    private readonly Rdv3Config cfg;
    private Rdv3SetCards cards;

    private Rdv3SetBox tName, tWinId, tWinClass, tWinLike, tWinProc;
    private Rdv3SetBox tFldId, tFldClass, tFldTypes, tFldIndex;
    private Rdv3SetBox tKeyLen, tPoll, tStable, tRebind, tCand;
    private Rdv3SetBox tData, tLedger, tLog;
    private Rdv3SetCheck cEnabled, cVp, cDigits, cFocus;
    private Rdv3SetSeg segRead, segScope;
    private Rdv3Btn btnClose, btnAdd, btnCopy, btnDel, btnPick, btnStepDel, btnSave, btnCancel;

    private readonly List<Control> page0 = new List<Control>();
    private readonly List<Control> page1 = new List<Control>();
    private readonly List<Control> page2 = new List<Control>();

    private Font fTitle, fSub, fTab, fLabel, fValue, fMeta, fNote;
    private int page;
    private int hotTab = -1;
    private int stepSel = -1;
    private bool loading;

    public Rdv3Config Result;

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
        Sc = FitScale(Rdv3SetGeom.ModalW, Rdv3SetGeom.ModalH);
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.SettingsTitle;
        ClientSize = new Size(PX(Rdv3SetGeom.ModalW), PX(Rdv3SetGeom.ModalH));
        MakeFonts();
        Build();
        LoadAll();
        ShowPage(0);
    }

    // The acceptance harness builds the dialog and renders it without ever
    // showing it, so it needs the constructor and the page switch by name. It
    // also drives the design scale itself, so the numbers do not depend on the
    // machine the check happens to run on.
    public static Rdv3SettingsForm ForCheck(Rdv3Config cfg) { return new Rdv3SettingsForm(cfg); }

    public void GoToPage(int i) { ShowPage(i); }

    public void SetDesignScale(float s)
    {
        Sc = Math.Max(0.5f, s);
        ClientSize = new Size(PX(Rdv3SetGeom.ModalW), PX(Rdv3SetGeom.ModalH));
        MakeFonts();
        Rescale();
        Layout1();
        Invalidate();
    }

    public int Page { get { return page; } }

    private void MakeFonts()
    {
        fTitle = Sf(21);
        fSub = Fn(10);
        fTab = Sf(16);
        fLabel = Fn(11);
        fValue = Fn(13);
        fMeta = Fn(10);
        fNote = Fn(11);
    }

    private void Rescale()
    {
        for (int i = 0; i < Controls.Count; i++)
        {
            Rdv3SetBox b = Controls[i] as Rdv3SetBox;
            if (b != null) { b.Sc = Sc; b.Font = Fn(13); continue; }
            Rdv3SetCheck c = Controls[i] as Rdv3SetCheck;
            if (c != null) { c.Sc = Sc; c.Font = Fn(13); continue; }
            Rdv3SetSeg s = Controls[i] as Rdv3SetSeg;
            if (s != null) { s.Sc = Sc; s.Font = Fn(13); continue; }
            Rdv3SetCards k = Controls[i] as Rdv3SetCards;
            if (k != null) { k.Sc = Sc; k.MakeFonts(); continue; }
            Rdv3Btn n = Controls[i] as Rdv3Btn;
            if (n != null) { n.Sc = Sc; n.Font = Sf(n.Ghost ? 12 : (n == btnSave || n == btnCancel) ? 14 : 13); }
        }
    }

    // ---- the controls ------------------------------------------------------
    private void Build()
    {
        btnClose = Ghost(5, "");
        btnClose.Click += delegate { Cancel(); };

        cards = new Rdv3SetCards(Sc);
        cards.Empty = Rdv3Text.NoteNoTargetShort;
        cards.OnPick = delegate { LoadTarget(); Layout1(); Invalidate(); };
        // scrolling moves the cards, so the reported rectangles move with them
        cards.OnView = delegate { Layout1(); };
        Add(page0, cards);

        btnAdd = Sec(Rdv3Text.BtnAdd, 13);
        btnAdd.Click += delegate { AddTarget(); };
        btnCopy = Sec(Rdv3Text.BtnCopy, 13);
        btnCopy.Click += delegate { CopyTarget(); };
        btnDel = Sec(Rdv3Text.BtnRemove, 13);
        btnDel.Click += delegate { RemoveTarget(); };
        btnPick = Sec(Rdv3Text.BtnInspect, 13);
        btnPick.Icon = 1;
        btnPick.Click += delegate { Pick(); };
        Add(page0, btnAdd); Add(page0, btnCopy); Add(page0, btnDel); Add(page0, btnPick);

        tName = Box(page0);
        cEnabled = Chk(page0, Rdv3Text.LblEnabled);
        tWinId = Box(page0); tWinClass = Box(page0); tWinLike = Box(page0); tWinProc = Box(page0);
        tFldId = Box(page0); tFldClass = Box(page0); tFldTypes = Box(page0); tFldIndex = Box(page0);
        cVp = Chk(page0, Rdv3Text.LblValuePattern);
        segRead = Seg(page0, new string[] { "value", "text", "name" });
        segScope = Seg(page0, new string[] { Rdv3Text.LblScopeDesc, Rdv3Text.LblScopeChildren });
        btnStepDel = Ghost(0, Rdv3Text.BtnRemove);
        btnStepDel.Click += delegate { RemoveStep(); };
        Add(page0, btnStepDel);

        tKeyLen = Box(page1);
        cDigits = Chk(page1, Rdv3Text.LblDigitsOnly);
        tPoll = Box(page1); tStable = Box(page1); tRebind = Box(page1); tCand = Box(page1);
        cFocus = Chk(page1, Rdv3Text.LblPreferFocus);

        tData = Box(page2); tLedger = Box(page2); tLog = Box(page2);

        btnSave = new Rdv3Btn();
        btnSave.Text = Rdv3Text.BtnSave;
        btnSave.Primary = true;
        btnSave.Icon = 2;
        btnSave.Sc = Sc;
        btnSave.Font = Sf(14);
        btnSave.BackColor = Rdv3Skin.N100;
        btnSave.Click += delegate { SaveAndClose(); };
        Controls.Add(btnSave);

        btnCancel = new Rdv3Btn();
        btnCancel.Text = Rdv3Text.BtnCancel;
        btnCancel.Sc = Sc;
        btnCancel.Font = Sf(14);
        btnCancel.BackColor = Rdv3Skin.N100;
        btnCancel.Click += delegate { Cancel(); };
        Controls.Add(btnCancel);
        CancelButton = btnCancel;
    }

    private Rdv3Btn Sec(string text, double px)
    {
        Rdv3Btn b = new Rdv3Btn();
        b.Text = text;
        b.Sc = Sc;
        b.Font = Sf(px);
        return b;
    }

    private Rdv3Btn Ghost(int icon, string text)
    {
        Rdv3Btn b = new Rdv3Btn();
        b.Text = text;
        b.Ghost = true;
        b.Icon = icon;
        b.Sc = Sc;
        b.Font = Sf(12);
        Controls.Add(b);
        return b;
    }

    private void Add(List<Control> page, Control c)
    {
        page.Add(c);
        if (!Controls.Contains(c)) { Controls.Add(c); }
    }

    private Rdv3SetBox Box(List<Control> page)
    {
        Rdv3SetBox b = new Rdv3SetBox(Sc);
        b.OnEdit = delegate { Store(); };
        b.GotFocus += delegate { Invalidate(); };
        b.LostFocus += delegate { Invalidate(); };
        Add(page, b);
        return b;
    }

    private Rdv3SetCheck Chk(List<Control> page, string text)
    {
        Rdv3SetCheck c = new Rdv3SetCheck(text, Sc);
        c.OnToggle = delegate { Store(); Invalidate(); };
        Add(page, c);
        return c;
    }

    private Rdv3SetSeg Seg(List<Control> page, string[] options)
    {
        Rdv3SetSeg s = new Rdv3SetSeg(options, Sc);
        s.OnPick = delegate { Store(); };
        Add(page, s);
        return s;
    }

    // ---- layout: straight off the measured table ---------------------------
    private void ShowPage(int i)
    {
        page = i;
        Vis(page0, i == 0);
        Vis(page1, i == 1);
        Vis(page2, i == 2);
        Layout1();
        Invalidate();
    }

    private static void Vis(List<Control> page, bool on)
    {
        for (int k = 0; k < page.Count; k++) { page[k].Visible = on; }
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Layout1();
    }

    private void Layout1()
    {
        if (fTitle == null) { return; }
        ClearRects();
        LayChrome();
        if (page == 0) { LayTargets(); }
        else if (page == 1) { LayBehaviour(); }
        else { LayFiles(); }
    }

    private double MW(string s, Font f)
    {
        return (Rdv3Skin.Measure(s, f).Width + 1) / (double)Sc;
    }

    private void LayChrome()
    {
        Rdv3SetGeom.R h = Rdv3SetGeom.Head;
        Put("head", h);
        Put("head.rule", Rdv3SetGeom.HeadRule);
        Rdv3SetGeom.R t = Rdv3SetGeom.HeadTitle;
        double tw = Math.Max(t.W, MW(Rdv3Text.SettingsTitle, fTitle));
        Put("head.title", t.X, t.Y, tw, t.H);
        Rdv3SetGeom.R s = Rdv3SetGeom.HeadSub;
        Put("head.sub", t.X + tw + 13.6, s.Y, Math.Max(s.W, MW(Rdv3Text.SetSubtitle, fSub)), s.H);
        Put("head.close", Rdv3SetGeom.HeadClose);
        Place(btnClose, "head.close");
        Rectangle cb = At("head.close");
        int ic = PX(14);
        PutR("head.close.icon", new Rectangle(cb.X + (cb.Width - ic) / 2, cb.Y + (cb.Height - ic) / 2, ic, ic));

        Put("tabs", Rdv3SetGeom.Tabs);
        Put("tabs.rule", Rdv3SetGeom.TabRule);
        Rdv3SetGeom.R t0 = Rdv3SetGeom.Tab0;
        string[] names = new string[] { Rdv3Text.SecTargets, Rdv3Text.SecBehaviour, Rdv3Text.SecPaths };
        double x = t0.X;
        for (int i = 0; i < 3; i++)
        {
            double w = MW(names[i], fTab) + 2 * Rdv3SetGeom.TabPad;
            Put("tabs.t" + i.ToString(CultureInfo.InvariantCulture), x, t0.Y, w, t0.H);
            if (i == page) { Put("tabs.mark", x, Rdv3SetGeom.TabMark.Y, w, Rdv3SetGeom.TabMark.H); }
            x += w + Rdv3SetGeom.TabGap;
        }
        Rdv3SetGeom.R m = Rdv3SetGeom.TabMeta;
        double mw = Math.Max(1.0, MW(MetaText(), fMeta));
        Put("tabs.meta", m.R2 - mw, m.Y, mw, m.H);

        Put("body", Rdv3SetGeom.Body);

        Put("foot", Rdv3SetGeom.Foot);
        Put("foot.rule", Rdv3SetGeom.FootRule);
        Put("foot.save", Rdv3SetGeom.FootSave);
        Put("foot.cancel", Rdv3SetGeom.FootCancel);
        Place(btnSave, "foot.save");
        Place(btnCancel, "foot.cancel");
        Rectangle sb = At("foot.save");
        int fi = PX(14);
        int lab = Rdv3Skin.Measure(Rdv3Text.BtnSave, btnSave.Font).Width;
        int all = fi + PX(6) + lab;
        PutR("foot.save.icon", new Rectangle(sb.X + (sb.Width - all) / 2,
            sb.Y + (sb.Height - fi) / 2, fi, fi));
        Rdv3SetGeom.R fn = Rdv3SetGeom.FootNote;
        Put("foot.note", fn.X, fn.Y, Rdv3SetGeom.FootSave.X - 20.4 - fn.X, fn.H);

        // the header is this window's title bar, so it drags it (the close
        // button is a child control and takes its own clicks)
        DragZone = At("head");
    }

    private string MetaText()
    {
        string file = cfg.SourcePath;
        int cut = file.LastIndexOf('\\');
        if (cut >= 0) { file = file.Substring(cut + 1); }
        if (file.Length == 0) { file = "ReaderDataViewer.json"; }
        return file.ToUpperInvariant() + Rdv3Text.SetMetaSchema
            + Rdv3Config.Schema.ToString(CultureInfo.InvariantCulture);
    }

    private void Place(Control c, string k)
    {
        Rectangle r = At(k);
        c.SetBounds(r.X, r.Y, r.Width, r.Height);
    }

    // a caption + its field, from two measured rectangles
    private void LayField(string k, Rdv3SetGeom.R label, Rdv3SetGeom.R box, Rdv3SetBox ctl, string caption)
    {
        Put(k + ".label", label.X, label.Y, Math.Max(label.W, MW(caption, fLabel)), label.H);
        Put(k, box);
        ctl.PlaceIn(At(k));
    }

    private void LayCheck(string k, Rdv3SetCheck c, double boxX, double labelY, double labelH)
    {
        int w = c.Wanted;
        c.SetBounds(PX(boxX), PX(labelY), w, PX(labelH));
        PutR(k + ".box", c.BoxRect);
        Rectangle lr = c.LabelRect;
        int lw = Rdv3Skin.Measure(c.Text, c.Font).Width;
        PutR(k + ".label", new Rectangle(lr.X, lr.Y, lw, lr.Height));
    }

    private void LaySeg(string k, Rdv3SetGeom.R label, Rdv3SetGeom.R seg, Rdv3SetSeg ctl, string caption)
    {
        Put(k + ".label", label.X, label.Y, Math.Max(label.W, MW(caption, fLabel)), label.H);
        Put(k + ".seg", seg);
        Place(ctl, k + ".seg");
        Rectangle sr = At(k + ".seg");
        for (int i = 0; i < ctl.Count; i++)
        {
            Rectangle c = ctl.Cell(i);
            PutR(k + ".seg" + i.ToString(CultureInfo.InvariantCulture),
                new Rectangle(sr.X + c.X, sr.Y + c.Y, c.Width, c.Height));
        }
    }

    private void LayTargets()
    {
        // ---- the list column -----------------------------------------------
        Rdv3SetGeom.R ll = Rdv3SetGeom.LeftLabel;
        Put("p0.left.label", ll.X, ll.Y, Math.Max(ll.W, MW(Rdv3Text.SecTargetList, fMeta)), ll.H);
        string count = cfg.Targets.Count.ToString(CultureInfo.InvariantCulture) + " " + Rdv3Text.UnitRows;
        Rdv3SetGeom.R lc = Rdv3SetGeom.LeftCount;
        double cw = Math.Max(lc.W, MW(count, fLabel));
        Put("p0.left.count", Rdv3SetGeom.List.R2 - cw, lc.Y, cw, lc.H);
        Put("p0.list", Rdv3SetGeom.List);
        Place(cards, "p0.list");
        cards.Measure();
        Rectangle lb = At("p0.list");
        for (int i = 0; i < cards.Drawn; i++)
        {
            Rectangle[] p = cards.Part(i);
            string key = "p0.card" + (cards.FirstDrawn + i).ToString(CultureInfo.InvariantCulture);
            PutR(key, Off(p[0], lb));
            PutR(key + ".name", Off(p[1], lb));
            PutR(key + ".tag", Off(p[2], lb));
            PutR(key + ".sum", Off(p[3], lb));
        }

        Rdv3SetGeom.R ba = Rdv3SetGeom.BtnAdd;
        Put("p0.btn.add", ba.X, ba.Y, ba.W, ba.H);
        Put("p0.btn.copy", ba.X + ba.W + Rdv3SetGeom.BtnGap, ba.Y, ba.W, ba.H);
        Put("p0.btn.del", ba.X + 2 * (ba.W + Rdv3SetGeom.BtnGap), ba.Y, ba.W, ba.H);
        Place(btnAdd, "p0.btn.add"); Place(btnCopy, "p0.btn.copy"); Place(btnDel, "p0.btn.del");
        Put("p0.btn.pick", Rdv3SetGeom.BtnPick);
        Place(btnPick, "p0.btn.pick");
        Rectangle pb = At("p0.btn.pick");
        int pi = PX(14);
        int plab = Rdv3Skin.Measure(Rdv3Text.BtnInspect, btnPick.Font).Width;
        int pall = pi + PX(6) + plab;
        PutR("p0.btn.pick.icon", new Rectangle(pb.X + (pb.Width - pall) / 2,
            pb.Y + (pb.Height - pi) / 2, pi, pi));
        Put("p0.hint", Rdv3SetGeom.PickHint);

        // ---- the detail column ----------------------------------------------
        // the "enabled" switch is anchored to the right edge; the name box ends
        // one gutter before it, exactly as the reference's grid does
        double right = Rdv3SetGeom.RightX + Rdv3SetGeom.RightW;
        double enW = cEnabled.Wanted / (double)Sc;
        LayCheck("p0.enabled", cEnabled, right - enW, Rdv3SetGeom.EnabledLbl.Y, Rdv3SetGeom.EnabledLbl.H);
        Rdv3SetGeom.R nb = Rdv3SetGeom.NameBox;
        double nameW = Math.Max(120.0, right - enW - Rdv3SetGeom.ColGap - nb.X);
        Put("p0.name.label", Rdv3SetGeom.NameLabel.X, Rdv3SetGeom.NameLabel.Y, nameW,
            Rdv3SetGeom.NameLabel.H);
        Put("p0.name.box", nb.X, nb.Y, nameW, nb.H);
        tName.PlaceIn(At("p0.name.box"));

        Put("p0.sec.window", Rdv3SetGeom.SecWindow);
        Put("p0.sec.window.rule", Rdv3SetGeom.SecWindow.X,
            Rdv3SetGeom.SecWindow.Y + Rdv3SetGeom.SecRuleDy, Rdv3SetGeom.SecWindow.W, 1.0);
        LayField("p0.win.id", Rdv3SetGeom.WinIdLabel, Rdv3SetGeom.WinIdBox, tWinId, Rdv3Text.LblAutomationId);
        LayField("p0.win.class", Col2(Rdv3SetGeom.WinIdLabel), Col2(Rdv3SetGeom.WinIdBox), tWinClass,
            Rdv3Text.LblClassName);
        LayField("p0.win.like", Rdv3SetGeom.WinLikeLabel, Rdv3SetGeom.WinLikeBox, tWinLike, Rdv3Text.LblNameLike);
        LayField("p0.win.proc", Col2(Rdv3SetGeom.WinLikeLabel), Col2(Rdv3SetGeom.WinLikeBox), tWinProc,
            Rdv3Text.LblProcess);

        Put("p0.sec.field", Rdv3SetGeom.SecField);
        Put("p0.sec.field.rule", Rdv3SetGeom.SecField.X,
            Rdv3SetGeom.SecField.Y + Rdv3SetGeom.SecRuleDy, Rdv3SetGeom.SecField.W, 1.0);
        LayField("p0.fld.id", Rdv3SetGeom.FldIdLabel, Rdv3SetGeom.FldIdBox, tFldId, Rdv3Text.LblAutomationId);
        LayField("p0.fld.class", Col2(Rdv3SetGeom.FldIdLabel), Col2(Rdv3SetGeom.FldIdBox), tFldClass,
            Rdv3Text.LblClassName);
        LayField("p0.fld.types", Rdv3SetGeom.FldTypesLabel, Rdv3SetGeom.FldTypesBox, tFldTypes,
            Rdv3Text.LblControlTypes);
        LayField("p0.fld.index", Rdv3SetGeom.FldIndexLabel, Rdv3SetGeom.FldIndexBox, tFldIndex,
            Rdv3Text.LblIndex);
        LayCheck("p0.vp", cVp, Rdv3SetGeom.FldIndexBox.R2 + 13.6, Rdv3SetGeom.VpLbl.Y, Rdv3SetGeom.VpLbl.H);

        LaySeg("p0.read", Rdv3SetGeom.ReadLabel, Rdv3SetGeom.ReadSeg, segRead, Rdv3Text.LblRead);
        LaySeg("p0.scope", Rdv3SetGeom.ScopeLabel, Rdv3SetGeom.ScopeSeg, segScope, Rdv3Text.LblScope);

        Put("p0.step.rule", Rdv3SetGeom.StepRule);
        Rdv3SetGeom.R sl = Rdv3SetGeom.StepLabel;
        double slw = Math.Max(sl.W, MW(Rdv3Text.LblPath, fLabel));
        Put("p0.step.label", sl.X, sl.Y, slw, sl.H);
        Put("p0.step.del", Rdv3SetGeom.StepDel);
        Place(btnStepDel, "p0.step.del");
        double noteX = sl.X + slw + Rdv3SetGeom.StepNoteGap;
        Put("p0.step.note", noteX, Rdv3SetGeom.StepNote.Y,
            Math.Max(40.0, Rdv3SetGeom.StepDel.X - 10.2 - noteX), Rdv3SetGeom.StepNote.H);
    }

    private static Rectangle Off(Rectangle r, Rectangle origin)
    {
        return new Rectangle(r.X + origin.X, r.Y + origin.Y, r.Width, r.Height);
    }

    private static Rdv3SetGeom.R Col2(Rdv3SetGeom.R r)
    {
        return new Rdv3SetGeom.R(Rdv3SetGeom.Col2X, r.Y, r.W, r.H);
    }

    private void LayBehaviour()
    {
        Put("p1.sec.key", Rdv3SetGeom.SecKey);
        Put("p1.sec.key.rule", Rdv3SetGeom.SecKey.X, Rdv3SetGeom.SecKey.Y + Rdv3SetGeom.SecRuleDy,
            Rdv3SetGeom.SecKey.W, 1.0);
        LayField("p1.keylen", Rdv3SetGeom.KeyLenLabel, Rdv3SetGeom.KeyLenBox, tKeyLen, Rdv3Text.LblKeyLen);
        Rdv3SetGeom.R ku = Rdv3SetGeom.KeyLenUnit;
        Put("p1.keylen.unit", ku.X, ku.Y, Math.Max(ku.W, MW(Rdv3Text.LblUnitDigits, fLabel)), ku.H);
        LayCheck("p1.digits", cDigits, Rdv3SetGeom.DigitsChk.X, Rdv3SetGeom.DigitsLbl.Y,
            Rdv3SetGeom.DigitsLbl.H);
        Put("p1.key.note", Rdv3SetGeom.KeyNote);

        Put("p1.sec.watch", Rdv3SetGeom.SecWatch);
        Put("p1.sec.watch.rule", Rdv3SetGeom.SecWatch.X, Rdv3SetGeom.SecWatch.Y + Rdv3SetGeom.SecRuleDy,
            Rdv3SetGeom.SecWatch.W, 1.0);
        LayField("p1.poll", Rdv3SetGeom.PollLabel, Rdv3SetGeom.PollBox, tPoll, Rdv3Text.LblPollMs);
        LayField("p1.stable", Rdv3SetGeom.StableLabel, Rdv3SetGeom.StableBox, tStable, Rdv3Text.LblStableMs);
        LayField("p1.rebind", Rdv3SetGeom.RebindLabel, Rdv3SetGeom.RebindBox, tRebind, Rdv3Text.LblRebindMs);
        LayField("p1.cand", Rdv3SetGeom.CandLabel, Rdv3SetGeom.CandBox, tCand, Rdv3Text.LblCandRows);
        Rdv3SetGeom.R cu = Rdv3SetGeom.CandUnit;
        Put("p1.cand.unit", cu.X, cu.Y, Math.Max(cu.W, MW(Rdv3Text.LblUnitRows, fLabel)), cu.H);
        Put("p1.watch.note", Rdv3SetGeom.WatchNote);
        LayCheck("p1.focus", cFocus, Rdv3SetGeom.FocusChk.X, Rdv3SetGeom.FocusLbl.Y, Rdv3SetGeom.FocusLbl.H);
    }

    private void LayFiles()
    {
        Put("p2.sec", Rdv3SetGeom.SecFiles);
        Put("p2.sec.rule", Rdv3SetGeom.SecFiles.X, Rdv3SetGeom.SecFiles.Y + Rdv3SetGeom.SecRuleDy,
            Rdv3SetGeom.SecFiles.W, 1.0);
        LayField("p2.data", Rdv3SetGeom.DataLabel, Rdv3SetGeom.DataBox, tData, Rdv3Text.LblDataDir);
        LayField("p2.ledger", Rdv3SetGeom.LedgerLabel, Rdv3SetGeom.LedgerBox, tLedger, Rdv3Text.LblLedger);
        LayField("p2.log", Rdv3SetGeom.LogLabel, Rdv3SetGeom.LogBox, tLog, Rdv3Text.LblLog);
        Put("p2.note", Rdv3SetGeom.FilesNote);
    }

    // ---- painting ----------------------------------------------------------
    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        ClearClip();
        // the layout is NOT redone here: it moves live controls, and moving a
        // control from inside its parent's paint is how a repaint loop starts.
        // Everything that can change the geometry calls Layout1 itself.
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }

        PaintHead(g);
        PaintTabs(g);
        if (page == 0) { PaintTargets(g); }
        else if (page == 1) { PaintBehaviour(g); }
        else { PaintFiles(g); }
        PaintFoot(g);
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }

    private void PaintHead(Graphics g)
    {
        T(g, "head.title", Rdv3Text.SettingsTitle, fTitle, Rdv3Skin.Ink);
        T(g, "head.sub", Rdv3Text.SetSubtitle, fSub, Rdv3Skin.N700);
        Rule(g, "head.rule");
    }

    private void PaintTabs(Graphics g)
    {
        string[] names = new string[] { Rdv3Text.SecTargets, Rdv3Text.SecBehaviour, Rdv3Text.SecPaths };
        for (int i = 0; i < 3; i++)
        {
            string k = "tabs.t" + i.ToString(CultureInfo.InvariantCulture);
            Color c = (i == page) ? Rdv3Skin.Accent800 : ((i == hotTab) ? Rdv3Skin.Ink : Rdv3Skin.N600);
            Rectangle r = At(k);
            Size z = Rdv3Skin.Measure(names[i], fTab);
            Rdv3Skin.Draw(g, names[i], fTab, c, r.X + (r.Width - z.Width) / 2, r.Y + (r.Height - z.Height) / 2);
        }
        Rectangle m = At("tabs.mark");
        if (m.Width > 0) { Rdv3Skin.Line(g, Rdv3Skin.Accent, m.X, m.Y, m.Width, m.Height); }
        TR(g, "tabs.meta", MetaText(), fMeta, Rdv3Skin.N600);
        Rule(g, "tabs.rule");
    }

    private void PaintFoot(Graphics g)
    {
        Rectangle f = At("foot");
        using (SolidBrush b = new SolidBrush(Rdv3Skin.N100)) { g.FillRectangle(b, f); }
        Rule(g, "foot.rule");
        T(g, "foot.note", Rdv3Text.NoteSavedTo + "  " + cfg.SourcePath, fNote, Rdv3Skin.N700);
    }

    private void PaintTargets(Graphics g)
    {
        T(g, "p0.left.label", Rdv3Text.SecTargetList, fMeta, Rdv3Skin.Accent);
        TR(g, "p0.left.count", cfg.Targets.Count.ToString(CultureInfo.InvariantCulture)
            + " " + Rdv3Text.UnitRows, fLabel, Rdv3Skin.N700);
        TWrap(g, "p0.hint", Rdv3Text.HintPick, fNote, Rdv3Skin.N700, 16.5);

        T(g, "p0.name.label", Rdv3Text.LblName, fLabel, Rdv3Skin.N700);
        Field(g, "p0.name.box", tName.Focused);

        Section(g, "p0.sec.window", Rdv3Text.SecWindowLong);
        Cap(g, "p0.win.id", Rdv3Text.LblAutomationId, tWinId);
        Cap(g, "p0.win.class", Rdv3Text.LblClassName, tWinClass);
        Cap(g, "p0.win.like", Rdv3Text.LblNameLike, tWinLike);
        Cap(g, "p0.win.proc", Rdv3Text.LblProcess, tWinProc);

        Section(g, "p0.sec.field", Rdv3Text.SecFieldLong);
        Cap(g, "p0.fld.id", Rdv3Text.LblAutomationId, tFldId);
        Cap(g, "p0.fld.class", Rdv3Text.LblClassName, tFldClass);
        Cap(g, "p0.fld.types", Rdv3Text.LblControlTypes, tFldTypes);
        Cap(g, "p0.fld.index", Rdv3Text.LblIndex, tFldIndex);

        T(g, "p0.read.label", Rdv3Text.LblRead, fLabel, Rdv3Skin.N700);
        T(g, "p0.scope.label", Rdv3Text.LblScope, fLabel, Rdv3Skin.N700);

        Rule(g, "p0.step.rule");
        T(g, "p0.step.label", Rdv3Text.LblPath, fLabel, Rdv3Skin.N700);
        T(g, "p0.step.note", StepNote(), fLabel, Rdv3Skin.N600);
    }

    private string StepNote()
    {
        Rdv3Target t = Current;
        if (t == null || t.Steps.Count == 0) { return Rdv3Text.LblStepNone; }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < t.Steps.Count; i++)
        {
            if (i > 0) { sb.Append("  >  "); }
            sb.Append(Step(t.Steps[i]));
        }
        return sb.ToString();
    }

    private static string Step(Rdv3Match m)
    {
        if (m.AutomationId.Length > 0) { return "#" + m.AutomationId; }
        if (m.ClassName.Length > 0) { return m.ClassName; }
        if (m.Name.Length > 0) { return m.Name; }
        return m.Describe();
    }

    private void Cap(Graphics g, string k, string caption, Rdv3SetBox box)
    {
        T(g, k + ".label", caption, fLabel, Rdv3Skin.N700);
        Field(g, k, box.Focused);
    }

    private void PaintBehaviour(Graphics g)
    {
        Section(g, "p1.sec.key", Rdv3Text.SecKeyLong);
        Cap(g, "p1.keylen", Rdv3Text.LblKeyLen, tKeyLen);
        T(g, "p1.keylen.unit", Rdv3Text.LblUnitDigits, fLabel, Rdv3Skin.N700);
        T(g, "p1.key.note", Rdv3Text.NoteKeyForm, fNote, Rdv3Skin.N700);

        Section(g, "p1.sec.watch", Rdv3Text.SecWatchLong);
        Cap(g, "p1.poll", Rdv3Text.LblPollMs, tPoll);
        Cap(g, "p1.stable", Rdv3Text.LblStableMs, tStable);
        Cap(g, "p1.rebind", Rdv3Text.LblRebindMs, tRebind);
        Cap(g, "p1.cand", Rdv3Text.LblCandRows, tCand);
        T(g, "p1.cand.unit", Rdv3Text.LblUnitRows, fLabel, Rdv3Skin.N700);
        T(g, "p1.watch.note", Rdv3Text.NoteWatchTiming, fNote, Rdv3Skin.N700);
    }

    private void PaintFiles(Graphics g)
    {
        Section(g, "p2.sec", Rdv3Text.SecFilesLong);
        Cap(g, "p2.data", Rdv3Text.LblDataDir, tData);
        Cap(g, "p2.ledger", Rdv3Text.LblLedger, tLedger);
        Cap(g, "p2.log", Rdv3Text.LblLog, tLog);
        T(g, "p2.note", Rdv3Text.NoteFilesBase, fNote, Rdv3Skin.N700);
    }

    // ---- the tab row is painted, so its hit test lives here ----------------
    protected override void OnMouseDown(MouseEventArgs e)
    {
        for (int i = 0; i < 3; i++)
        {
            if (At("tabs.t" + i.ToString(CultureInfo.InvariantCulture)).Contains(e.Location)
                && i != page)
            {
                ShowPage(i);
                break;
            }
        }
        base.OnMouseDown(e);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        int hit = -1;
        for (int i = 0; i < 3; i++)
        {
            if (At("tabs.t" + i.ToString(CultureInfo.InvariantCulture)).Contains(e.Location)) { hit = i; }
        }
        if (hit != hotTab) { hotTab = hit; Invalidate(At("tabs")); }
        base.OnMouseMove(e);
    }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        if (keyData == (Keys.Control | Keys.Tab)) { ShowPage((page + 1) % 3); return true; }
        if (keyData == (Keys.Control | Keys.Shift | Keys.Tab)) { ShowPage((page + 2) % 3); return true; }
        return base.ProcessCmdKey(ref msg, keyData);
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
            cards.Items.Add(new string[] {
                (t.Name.Length > 0) ? t.Name : Rdv3Text.SecTarget,
                t.Enabled ? Rdv3Text.LblEnabled : Rdv3Text.LblDisabled,
                Summary(t) });
        }
        if (cards.Index >= cfg.Targets.Count) { cards.Index = Math.Max(0, cfg.Targets.Count - 1); }
        cards.ScrollToSelected();
        cards.Invalidate();
    }

    private static string Summary(Rdv3Target t)
    {
        string win = t.Window.ProcessName.Length > 0 ? t.Window.ProcessName
            : (t.Window.ClassName.Length > 0 ? t.Window.ClassName : t.Window.NameLike);
        string fld = t.Field.AutomationId.Length > 0 ? ("#" + t.Field.AutomationId)
            : (t.Field.ControlTypes.Length > 0 ? string.Join("/", t.Field.ControlTypes) : t.Field.ClassName);
        if (win.Length == 0) { win = Rdv3Text.NoValue; }
        if (fld.Length == 0) { fld = Rdv3Text.NoValue; }
        return Rdv3Text.NoteTargetSummary.Replace("{win}", win).Replace("{field}", fld);
    }

    private void LoadTarget()
    {
        Rdv3Target t = Current;
        loading = true;
        stepSel = (t != null && t.Steps.Count > 0) ? 0 : -1;
        bool on = (t != null);
        if (t == null)
        {
            tName.Text = ""; cEnabled.Checked = false;
            tWinId.Text = ""; tWinClass.Text = ""; tWinLike.Text = ""; tWinProc.Text = "";
            tFldId.Text = ""; tFldClass.Text = ""; tFldTypes.Text = ""; tFldIndex.Text = "0";
            cVp.Checked = false;
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
            cVp.Checked = t.Field.RequireValuePattern;
            segRead.Index = t.ReadMode;
            segScope.Index = t.Field.Descendants ? 0 : 1;
        }
        tName.Enabled = on; tWinId.Enabled = on; tWinClass.Enabled = on;
        tWinLike.Enabled = on; tWinProc.Enabled = on;
        tFldId.Enabled = on; tFldClass.Enabled = on; tFldTypes.Enabled = on; tFldIndex.Enabled = on;
        btnCopy.Enabled = on; btnDel.Enabled = on;
        btnStepDel.Enabled = (t != null && t.Steps.Count > 0);
        loading = false;
        cEnabled.Invalidate(); cVp.Invalidate();
        segRead.Invalidate(); segScope.Invalidate();
        Invalidate();
    }

    private void Store()
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
        t.Field.ControlTypes = SplitList(tFldTypes.Text);
        t.Field.Index = ToInt(tFldIndex.Text, 0);
        t.Field.RequireValuePattern = cVp.Checked;
        t.Field.Descendants = (segScope.Index == 0);
        t.ReadMode = segRead.Index;
        RefreshCards();
        Invalidate();
    }

    private static string[] SplitList(string s)
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
        cards.Index = cfg.Targets.Count - 1;
        RefreshCards();
        LoadTarget();
        Layout1();
        Invalidate();
    }

    private void CopyTarget()
    {
        Rdv3Target t = Current;
        if (t == null) { return; }
        cfg.Targets.Add(t.Clone());
        cards.Index = cfg.Targets.Count - 1;
        RefreshCards();
        LoadTarget();
        Layout1();
        Invalidate();
    }

    private void RemoveTarget()
    {
        int i = cards.Index;
        if (i < 0 || i >= cfg.Targets.Count) { return; }
        cfg.Targets.RemoveAt(i);
        cards.Index = Math.Max(0, Math.Min(i, cfg.Targets.Count - 1));
        RefreshCards();
        LoadTarget();
        Layout1();
        Invalidate();
    }

    private void RemoveStep()
    {
        Rdv3Target t = Current;
        if (t == null || t.Steps.Count == 0) { return; }
        int at = (stepSel >= 0 && stepSel < t.Steps.Count) ? stepSel : (t.Steps.Count - 1);
        t.Steps.RemoveAt(at);
        stepSel = (t.Steps.Count > 0) ? 0 : -1;
        btnStepDel.Enabled = (t.Steps.Count > 0);
        RefreshCards();
        Invalidate();
    }

    private void Pick()
    {
        Rdv3Target picked = Rdv3PickerForm.Pick(this);
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
        LoadTarget();
        Layout1();
        Invalidate();
    }

    private void Cancel()
    {
        DialogResult = DialogResult.Cancel;
        Close();
    }

    private void SaveAndClose()
    {
        Store();
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
    private Rdv3Btn btnUse, btnClose;
    private Font fTitle, fHow, fKey, fVal, fValBold, fTag;

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
        Sc = FitScale(Rdv3SetGeom.PickW, Rdv3SetGeom.PickH);
        Text = Rdv3Text.AppTitle + " - " + Rdv3Text.PickTitle;
        ClientSize = new Size(PX(Rdv3SetGeom.PickW), PX(Rdv3SetGeom.PickH));
        TopMost = true;
        StartPosition = FormStartPosition.Manual;
        MakeFonts();

        Rectangle wa = Screen.PrimaryScreen.WorkingArea;
        Location = new Point(wa.Right - Width - PX(24), wa.Bottom - Height - PX(24));

        btnUse = new Rdv3Btn();
        btnUse.Text = Rdv3Text.BtnUseElement;
        btnUse.Primary = true;
        btnUse.Sc = Sc;
        btnUse.Font = Sf(13);
        btnUse.Enabled = false;
        btnUse.Click += delegate { Take(); };
        Controls.Add(btnUse);

        btnClose = new Rdv3Btn();
        btnClose.Text = Rdv3Text.BtnClose;
        btnClose.Sc = Sc;
        btnClose.Font = Sf(13);
        btnClose.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        Controls.Add(btnClose);
        CancelButton = btnClose;

        tick.Interval = 90;
        tick.Tick += delegate { Look(); };
        tick.Start();
    }

    // the acceptance harness renders the panel without a live cursor under it
    public static Rdv3PickerForm ForCheck()
    {
        Rdv3PickerForm f = new Rdv3PickerForm();
        f.tick.Stop();
        return f;
    }

    public void SetDesignScale(float s)
    {
        Sc = Math.Max(0.5f, s);
        ClientSize = new Size(PX(Rdv3SetGeom.PickW), PX(Rdv3SetGeom.PickH));
        MakeFonts();
        btnUse.Sc = Sc; btnUse.Font = Sf(13);
        btnClose.Sc = Sc; btnClose.Font = Sf(13);
        Layout1();
        Invalidate();
    }

    private void MakeFonts()
    {
        fTitle = Sf(18);
        fHow = Rdv3Skin.F(12, FontStyle.Bold, Sc);
        fKey = Fn(11);
        fVal = Fn(12);
        fValBold = Fb(12);
        fTag = Fn(10);
    }

    // shown for the acceptance render, which never has a live element under the
    // cursor: the rows say what they will say
    public void SetSample(string type, string id, string cls, string name, string process, string read)
    {
        ctrlType = type; autoId = id; className = cls; elName = name; proc = process;
        value = read; canRead = (read != null && read.Length > 0);
        btnUse.Enabled = true;
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

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); }
        base.OnKeyDown(e);
    }

    private static readonly string[] RowKeys = { "pk.k0", "pk.k1", "pk.k2", "pk.k3", "pk.k4", "pk.k5" };

    private void Layout1()
    {
        if (fTitle == null) { return; }
        ClearRects();
        Rdv3SetGeom.R t = Rdv3SetGeom.PkTitle;
        Put("pk.title", t.X, t.Y, Math.Max(t.W, Wid(Rdv3Text.PickTitle, fTitle)), t.H);
        Rdv3SetGeom.R g = Rdv3SetGeom.PkTag;
        double gw = Math.Max(g.W, Wid(Rdv3Text.TagTopMost, fTag) + 20.0);
        Put("pk.tag", g.R2 - gw, g.Y, gw, g.H);
        Put("pk.how", Rdv3SetGeom.PkHow);
        Put("pk.rule", Rdv3SetGeom.PkRule);
        Rdv3SetGeom.R k0 = Rdv3SetGeom.PkKey0;
        Rdv3SetGeom.R v0 = Rdv3SetGeom.PkVal0;
        for (int i = 0; i < 6; i++)
        {
            Put(RowKeys[i], k0.X, k0.Y + Rdv3SetGeom.PkRowH * i, k0.W, k0.H);
            Put("pk.v" + i.ToString(CultureInfo.InvariantCulture),
                v0.X, v0.Y + Rdv3SetGeom.PkRowH * i, v0.W, v0.H);
        }
        Put("pk.esc", Rdv3SetGeom.PkEsc);
        Rdv3SetGeom.R cl = Rdv3SetGeom.PkClose;
        double clw = Math.Max(cl.W, Wid(Rdv3Text.BtnClose, btnClose.Font) + 28.5);
        double clx = Rdv3SetGeom.PkFootRight.X - clw;
        Put("pk.close", clx, cl.Y, clw, cl.H);
        Rdv3SetGeom.R us = Rdv3SetGeom.PkUse;
        double usw = Math.Max(us.W, Wid(Rdv3Text.BtnUseElement, btnUse.Font) + 28.5);
        Put("pk.use", clx - Rdv3SetGeom.PkBtnGap - usw, us.Y, usw, us.H);
        Rectangle ur = At("pk.use"), cr = At("pk.close");
        btnUse.SetBounds(ur.X, ur.Y, ur.Width, ur.Height);
        btnClose.SetBounds(cr.X, cr.Y, cr.Width, cr.Height);
        // the panel sits over the application being inspected, so it has to be
        // movable: everything above the rule drags it
        DragZone = new Rectangle(0, 0, ClientSize.Width, At("pk.rule").Y);
    }

    private double Wid(string s, Font f)
    {
        return (Rdv3Skin.Measure(s, f).Width + 1) / (double)Sc;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics gr = e.Graphics;
        gr.SmoothingMode = SmoothingMode.AntiAlias;
        ClearClip();
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { gr.FillRectangle(b, ClientRectangle); }

        T(gr, "pk.title", Rdv3Text.PickTitle, fTitle, Rdv3Skin.Ink);
        Chip(gr, "pk.tag", Rdv3Text.TagTopMost, Rdv3Skin.Accent100, Rdv3Skin.Accent800);
        T(gr, "pk.how", Rdv3Text.PickHow, fHow, Rdv3Skin.Accent700);
        Rule(gr, "pk.rule");

        Row(gr, 0, Rdv3Text.LblControlTypes, ctrlType, false, false);
        Row(gr, 1, Rdv3Text.LblAutomationId, autoId, true, false);
        Row(gr, 2, Rdv3Text.LblClassName, className, false, false);
        Row(gr, 3, Rdv3Text.LblName, elName, false, false);
        Row(gr, 4, Rdv3Text.LblProcessOf, proc, false, false);
        Row(gr, 5, Rdv3Text.PickReading, canRead ? value : Rdv3Text.PickNoRead, false, !canRead);

        T(gr, "pk.esc", Rdv3Text.PickEsc, fKey, Rdv3Skin.N700);
        Rdv3Skin.Frame(gr, Rdv3Skin.Divider, ClientRectangle);
    }

    private void Row(Graphics g, int i, string k, string v, bool strong, bool muted)
    {
        string kk = "pk.k" + i.ToString(CultureInfo.InvariantCulture);
        string vk = "pk.v" + i.ToString(CultureInfo.InvariantCulture);
        T(g, kk, k, fKey, Rdv3Skin.N700);
        string shown = (v == null || v.Length == 0) ? Rdv3Text.NotYet : v;
        T(g, vk, shown, strong ? fValBold : fVal, muted ? Rdv3Skin.N600 : Rdv3Skin.Ink);
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
            btnUse.Enabled = true;
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
}
