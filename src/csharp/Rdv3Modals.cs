// ============================================================================
// Rdv3Modals.cs -- the overlays of the reference: a dimmed card with a modal
// on top (confirmations, the candidate list, the settings modal), the
// always-on-top picker, and the toast at the bottom right.
//
// A WinForms child control cannot be translucent over its siblings, so each
// overlay is its own borderless window owned by the main form: a backdrop
// form at 45% over the card, and the modal card above it. The modal draws
// its own chrome (title, close mark) -- the reference has no system title
// bar -- and is dragged by its header. Esc closes it.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

// ---------------------------------------------------------------------------
// Rounded window corners the clean way: the desktop window manager rounds
// them (Windows 11; DWMWA_WINDOW_CORNER_PREFERENCE) with anti-aliased edges
// and a matching shadow. Clipping a Region gives jagged, darkened edges, so
// that is not done. On Windows 10 the call fails quietly and the corners
// stay square.
// ---------------------------------------------------------------------------
internal static class Rdv3Corners
{
    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

    public static void Round(IWin32Window w)
    {
        try
        {
            int pref = 2;                                   // DWMWCP_ROUND
            DwmSetWindowAttribute(w.Handle, 33, ref pref, 4);
        }
        catch (Exception) { }
    }
}

// ---------------------------------------------------------------------------
// the 45% ink over the card while a modal is up
// ---------------------------------------------------------------------------
internal sealed class Rdv3Backdrop : Form
{
    public Rdv3Backdrop(Rectangle screenRect)
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = Rdv3Skin.N900;
        Opacity = 0.45;
        Bounds = screenRect;
    }

    protected override bool ShowWithoutActivation { get { return true; } }
}

// ---------------------------------------------------------------------------
// a modal card: rounded, shadowed, with its own header, laid out in CSS px
// at a scale of its own (the screen's, or less when the work area is small)
// ---------------------------------------------------------------------------
public abstract class Rdv3Dialog : Form
{
    protected float Sc = 1f;
    protected double CssW, CssH;
    protected Rectangle DragZone = Rectangle.Empty;

    private readonly Dictionary<string, Rectangle> rc = new Dictionary<string, Rectangle>();
    private readonly List<string> clip = new List<string>();
    private bool dragging;
    private Point dragFrom;
    private Rdv3Backdrop backdrop;

    protected Rdv3Dialog()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        FormBorderStyle = FormBorderStyle.None;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = Rdv3Skin.White;
        AutoScaleMode = AutoScaleMode.None;
        KeyPreview = true;
    }

    // CS_DROPSHADOW: the reference's shadow-lg, drawn by the window manager
    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams cp = base.CreateParams;
            cp.ClassStyle |= 0x00020000;
            return cp;
        }
    }

    // the design is a fixed shape; if the space it is shown in (the window's
    // client area, or the work area when there is no window) cannot hold it
    // at the screen's scale, the whole dialog scales down (never below 0.7)
    protected static float FitScale(double cssW, double cssH, float screenScale, Rdv3Form owner)
    {
        float sc = screenScale;
        Rectangle room = (owner == null) ? Screen.PrimaryScreen.WorkingArea : owner.CardBounds;
        double mw = Math.Max(360.0, room.Width - 24.0);
        double mh = Math.Max(280.0, room.Height - 24.0);
        if (cssW * sc > mw) { sc = (float)(mw / cssW); }
        if (cssH * sc > mh) { sc = (float)(mh / cssH); }
        if (sc > screenScale) { sc = screenScale; }
        return Math.Max(0.7f, sc);
    }

    protected void SizeTo(double cssW, double cssH)
    {
        CssW = cssW; CssH = cssH;
        ClientSize = new Size(PX(cssW), PX(cssH));
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Rdv3Corners.Round(this);
    }

    // centred over a rectangle on the screen (the card), kept on the work area
    protected void CentreOver(Rectangle screenRect)
    {
        Rectangle wa = Screen.FromRectangle(screenRect).WorkingArea;
        int x = screenRect.X + (screenRect.Width - Width) / 2;
        int y = screenRect.Y + (screenRect.Height - Height) / 2;
        x = Math.Max(wa.X, Math.Min(x, wa.Right - Width));
        y = Math.Max(wa.Y, Math.Min(y, wa.Bottom - Height));
        Location = new Point(x, y);
    }

    // show as a modal over the owner's card, dimming it
    protected DialogResult ShowOver(Rdv3Form owner)
    {
        Rectangle card = (owner == null) ? Screen.PrimaryScreen.WorkingArea : owner.CardBounds;
        CentreOver(card);
        if (owner != null)
        {
            backdrop = new Rdv3Backdrop(card);
            backdrop.Show(owner);
        }
        try { return ShowDialog(owner); }
        finally
        {
            if (backdrop != null) { backdrop.Close(); backdrop.Dispose(); backdrop = null; }
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

    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); e.Handled = true; return; }
        base.OnKeyDown(e);
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

    protected Rectangle At(string k)
    {
        Rectangle r;
        return rc.TryGetValue(k, out r) ? r : Rectangle.Empty;
    }

    protected bool Has(string k) { return rc.ContainsKey(k); }

    protected double MW(string s, Font f)
    {
        return (Rdv3Skin.Measure(s, f).Width + 1) / (double)Sc;
    }

    protected void ClearClip() { clip.Clear(); }

    private void Note(string k, string s, Font f, Rectangle r)
    {
        if (s == null || s.Length == 0) { return; }
        if (Rdv3Skin.Measure(s, f).Width > r.Width && !clip.Contains(k)) { clip.Add(k); }
    }

    protected void T(Graphics g, string k, string s, Font f, Color c, int align)
    {
        Rectangle r = At(k);
        Note(k, s, f, r);
        Rdv3Skin.DrawIn(g, s, f, c, r, align);
    }

    protected void Rule(Graphics g, string k)
    {
        Rectangle r = At(k);
        if (r.Width <= 0) { return; }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, r.Width, 1);
    }

    // a rounded, borderless input (the reference's inputs sit on neutral-200)
    protected void Field(Graphics g, string k, bool focused)
    {
        Rectangle r = At(k);
        if (r.Width <= 0) { return; }
        Rdv3Skin.FillRound(g, Rdv3Skin.N200, r, PX(10));
        if (focused) { Rdv3Skin.FrameRound(g, Rdv3Skin.Accent, r, PX(10)); }
    }

    protected Rdv3Btn MakeBtn(string text, int kind, double fontCss, string icon)
    {
        Rdv3Btn b = new Rdv3Btn();
        b.Text = text;
        b.Kind = kind;
        b.Icon = icon;
        b.Sc = Sc;
        b.Font = Sf(fontCss);
        b.BackColor = Rdv3Skin.White;
        Controls.Add(b);
        return b;
    }

    protected void Place(Control c, string k)
    {
        Rectangle r = At(k);
        c.SetBounds(r.X, r.Y, r.Width, r.Height);
    }

    // ---- the acceptance dump ------------------------------------------------
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
// yes / no: 440 wide, padding 14.6, title 22/600, body 16/400. Tell() is the
// same modal with one button, for a message that only needs acknowledging.
// ---------------------------------------------------------------------------
public sealed class Rdv3ConfirmForm : Rdv3Dialog
{
    private readonly string title;
    private readonly string body;
    private readonly bool single;
    private Rdv3Btn btnYes, btnNo;
    private Font fTitle, fBody;
    private const double Pad = 14.6;
    private const double LineH = 25.6;       // 16px at 1.6

    public static bool Ask(Rdv3Form owner, string title, string body)
    {
        Rdv3ConfirmForm f = new Rdv3ConfirmForm(owner, title, body, false);
        DialogResult dr = f.ShowOver(owner);
        f.Dispose();
        return dr == DialogResult.Yes;
    }

    public static void Tell(Rdv3Form owner, string title, string body)
    {
        Rdv3ConfirmForm f = new Rdv3ConfirmForm(owner, title, body, true);
        f.ShowOver(owner);
        f.Dispose();
    }

    private Rdv3ConfirmForm(Rdv3Form owner, string t, string b, bool one)
    {
        title = t;
        single = one;
        body = (b == null) ? "" : b.Replace("\r\n", "\n");
        Sc = FitScale(440, 200, (owner == null) ? Rdv3Skin.Scale : owner.Sc, owner);
        Text = title;
        fTitle = Sf(22);
        fBody = Fn(16);
        int lines = CountLines(body, 440 - 2 * Pad);
        double h = Pad + 26.4 + 8 + lines * LineH + 14 + 38 + Pad;
        SizeTo(440, h);
        btnYes = MakeBtn(single ? Rdv3Text.BtnClose : Rdv3Text.BtnYes, Rdv3Btn.Primary, 16, "");
        btnNo = MakeBtn(Rdv3Text.BtnNo, Rdv3Btn.Secondary, 16, "");
        btnYes.Click += delegate { DialogResult = single ? DialogResult.OK : DialogResult.Yes; Close(); };
        btnNo.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        btnNo.Visible = !single;
        AcceptButton = btnYes;
        CancelButton = single ? btnYes : btnNo;
        Layout1(lines);
    }

    // the headless check builds one without showing it
    public static Rdv3ConfirmForm ForCheck(string title, string body, float scale)
    {
        Rdv3ConfirmForm f = new Rdv3ConfirmForm(null, title, body, false);
        f.SetDesignScale(scale);
        return f;
    }

    public void SetDesignScale(float s)
    {
        Sc = Math.Max(0.5f, s);
        fTitle = Sf(22); fBody = Fn(16);
        btnYes.Sc = Sc; btnYes.Font = Sf(16);
        btnNo.Sc = Sc; btnNo.Font = Sf(16);
        int lines = CountLines(body, 440 - 2 * Pad);
        SizeTo(440, Pad + 26.4 + 8 + lines * LineH + 14 + 38 + Pad);
        Layout1(lines);
        Invalidate();
    }

    private int CountLines(string s, double widthCss)
    {
        return Rdv3Skin.CountLines(s, fBody, PX(widthCss));
    }

    private void Layout1(int lines)
    {
        ClearRects();
        Put("cf.title", Pad, Pad, CssW - 2 * Pad, 26.4);
        Put("cf.body", Pad, Pad + 26.4 + 8, CssW - 2 * Pad, lines * LineH);
        double by = Pad + 26.4 + 8 + lines * LineH + 14;
        double wNo = single ? 0 : Math.Max(84, MW(Rdv3Text.BtnNo, btnNo.Font) + 32);
        double wYes = Math.Max(84, MW(btnYes.Text, btnYes.Font) + 32);
        Put("cf.no", CssW - Pad - wNo, by, wNo, 38);
        Put("cf.yes", CssW - Pad - wNo - (single ? 0 : 6.8) - wYes, by, wYes, 38);
        Place(btnYes, "cf.yes");
        Place(btnNo, "cf.no");
        DragZone = new Rectangle(0, 0, ClientSize.Width, PX(Pad + 26.4));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        ClearClip();
        using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, ClientRectangle); }
        T(g, "cf.title", title, fTitle, Rdv3Skin.Ink, 0);
        Rdv3Skin.DrawWrapped(g, body, fBody, Rdv3Skin.N800, At("cf.body"), 25.6, Sc);
    }
}

// ---------------------------------------------------------------------------
// the candidate table: sticky header, hover, selection, wheel + slim scrollbar,
// columns from the screen definition
// ---------------------------------------------------------------------------
internal sealed class Rdv3CandTable : Control
{
    private readonly Rdv3CandidatesDef def;
    private readonly Rdv3Fields fields;
    private readonly Rdv3WorkState work;
    private readonly float sc;
    private List<Rdv3CandRow> rows = new List<Rdv3CandRow>();
    private int sel = -1;
    private int hot = -1;
    private int top;
    private bool dragging;
    private int dragGrab;
    public Action<int> OnPick;

    private Font fTh, fTd, fTdBold, fTag;
    private int[] colw;

    public Rdv3CandTable(Rdv3CandidatesDef d, Rdv3Fields f, Rdv3WorkState w, float scale)
    {
        def = d; fields = f; work = w; sc = scale;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.White;
        fTh = Rdv3Skin.S(13, sc);
        fTd = Rdv3Skin.F(16, FontStyle.Regular, sc);
        fTdBold = Rdv3Skin.F(16, FontStyle.Bold, sc);
        fTag = Rdv3Skin.F(13, FontStyle.Regular, sc);
    }

    private int P(double css) { return (int)Math.Round(css * sc); }

    public int HeaderH { get { return P(def.HeaderHeight); } }
    public int RowH { get { return P(def.RowHeight); } }
    public int Count { get { return rows.Count; } }
    public int Selected { get { return sel; } }

    public void SetRows(List<Rdv3CandRow> r, int selected)
    {
        rows = (r == null) ? new List<Rdv3CandRow>() : r;
        sel = selected;
        top = 0;
        hot = -1;
        colw = null;
        SyncBar();
        Invalidate();
    }

    // the column widths: the definition's, the last (0) takes the rest. A
    // cell is cut with an ellipsis rather than widening its column, as the
    // reference's table-layout: fixed does.
    private int[] Cols()
    {
        if (colw != null) { return colw; }
        int n = def.Columns.Count;
        int[] w = new int[n];
        int fixedSum = 0, flex = 0;
        for (int c = 0; c < n; c++)
        {
            if (def.Columns[c].Width > 0) { w[c] = P(def.Columns[c].Width); fixedSum += w[c]; }
            else { flex++; }
        }
        int avail = Math.Max(0, ClientSize.Width - (NeedBar ? BarW : 0) - fixedSum);
        for (int c = 0; c < n; c++)
        {
            if (def.Columns[c].Width <= 0) { w[c] = (flex > 0) ? avail / flex : 0; }
        }
        colw = w;
        return w;
    }

    private int[] Edges()
    {
        int[] w = Cols();
        int[] x = new int[w.Length + 1];
        x[0] = 0;
        for (int i = 0; i < w.Length; i++) { x[i + 1] = x[i] + w[i]; }
        return x;
    }

    private int VisibleRows()
    {
        int h = ClientSize.Height - HeaderH;
        return Math.Max(1, h / Math.Max(1, RowH));
    }

    private int BarW { get { return P(10); } }
    private bool NeedBar { get { return rows.Count > VisibleRows(); } }

    private void SyncBar()
    {
        int vis = VisibleRows();
        if (rows.Count > vis) { top = Math.Max(0, Math.Min(top, rows.Count - vis)); }
        else { top = 0; }
    }

    private Rectangle BarTrack()
    {
        return new Rectangle(ClientSize.Width - BarW, HeaderH, BarW, Math.Max(0, ClientSize.Height - HeaderH));
    }

    private Rectangle BarThumb()
    {
        Rectangle t = BarTrack();
        int vis = VisibleRows();
        if (rows.Count <= vis || t.Height <= 0) { return Rectangle.Empty; }
        int h = Math.Max(P(24), (int)((long)t.Height * vis / rows.Count));
        int span = t.Height - h;
        int y = t.Y + (int)((long)span * top / Math.Max(1, rows.Count - vis));
        int pad = P(3);
        return new Rectangle(t.X + pad, y, Math.Max(2, t.Width - 2 * pad), h);
    }

    private void ScrollTo(int t)
    {
        int vis = VisibleRows();
        int max = Math.Max(0, rows.Count - vis);
        int v = Math.Max(0, Math.Min(t, max));
        if (v != top) { top = v; Invalidate(); }
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); colw = null; SyncBar(); }

    protected override void OnMouseWheel(MouseEventArgs e)
    {
        if (NeedBar) { ScrollTo(top + ((e.Delta > 0) ? -1 : 1)); }
        base.OnMouseWheel(e);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (NeedBar && e.X >= BarTrack().X)
        {
            Rectangle th = BarThumb();
            if (th.Contains(e.Location)) { dragging = true; dragGrab = e.Y - th.Y; }
            else { ScrollTo(top + ((e.Y < th.Y) ? -VisibleRows() : VisibleRows())); }
            return;
        }
        base.OnMouseDown(e);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        dragging = false;
        base.OnMouseUp(e);
    }

    private int RowAt(int y)
    {
        if (y < HeaderH) { return -1; }
        int i = top + (y - HeaderH) / Math.Max(1, RowH);
        return (i >= 0 && i < rows.Count) ? i : -1;
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (dragging)
        {
            Rectangle t = BarTrack();
            int h = BarThumb().Height;
            int span = Math.Max(1, t.Height - h);
            int vis = VisibleRows();
            ScrollTo((int)Math.Round((double)(e.Y - dragGrab - t.Y) * Math.Max(0, rows.Count - vis) / span));
            return;
        }
        int i = (NeedBar && e.X >= BarTrack().X) ? -1 : RowAt(e.Y);
        if (i != hot) { hot = i; Invalidate(); }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        if (hot != -1) { hot = -1; Invalidate(); }
        base.OnMouseLeave(e);
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        if (NeedBar && e.X >= BarTrack().X) { return; }
        int i = RowAt(e.Y);
        if (i >= 0 && OnPick != null) { OnPick(i); }
        base.OnMouseClick(e);
    }

    public int[] ColumnEdges() { return Edges(); }

    // one cell's text and tone, evaluated through the same binding rules as
    // the main screen, with the row's own record and state in view
    private Rdv3Value Cell(Rdv3ColumnDef c, Rdv3View v)
    {
        return Rdv3Eval.Evaluate(c.Value, v, fields, work);
    }

    private int TagKindOf(Rdv3ColumnDef c, Rdv3View v, string raw)
    {
        if (c.Value != null && c.Value.IsState && (c.Value.State == "workState" || c.Value.State == "workStateShort"))
        {
            return Rdv3Skin.TagKind(Rdv3Eval.WorkStateLook(v, work));
        }
        string look;
        if (c.Looks.TryGetValue(raw, out look)) { return Rdv3Skin.TagKind(look); }
        if (c.Looks.TryGetValue("*", out look)) { return Rdv3Skin.TagKind(look); }
        return Rdv3Skin.TagNeutral;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, ClientRectangle); }
        int[] xs = Edges();
        int padL = P(6.8), padR = P(13.6);
        int w = ClientSize.Width - (NeedBar ? BarW : 0);
        Rdv3View v = new Rdv3View();

        int y = HeaderH;
        for (int i = top; i < rows.Count && y < ClientSize.Height; i++)
        {
            Rdv3CandRow r = rows[i];
            Rectangle rr = new Rectangle(0, y, w, RowH);
            if (i == sel)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent100)) { g.FillRectangle(b, rr); }
            }
            else if (i == hot)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Mix(Rdv3Skin.White, 0.04))) { g.FillRectangle(b, rr); }
            }
            Rdv3Skin.Line(g, Rdv3Skin.RowLine, 0, rr.Bottom - 1, w, 1);

            v.Record = Rdv3Ledger.SplitLine(r.Line);
            v.StoredState = r.Stored;
            v.RowNumber = i + 1;
            for (int c = 0; c < def.Columns.Count; c++)
            {
                Rdv3ColumnDef col = def.Columns[c];
                bool right = (col.Align == "right");
                Rectangle cell = new Rectangle(xs[c] + padL, y, Math.Max(0, xs[c + 1] - xs[c] - padL - (right ? padR : padL)), RowH);
                Rdv3Value val = Cell(col, v);
                if (col.Render == "tag")
                {
                    if (val.Text.Length == 0) { continue; }
                    Size tz = Rdv3Skin.TagSize(val.Text, fTag, sc);
                    int kind = (val.Tone == Rdv3Value.Error) ? Rdv3Skin.TagError : TagKindOf(col, v, val.Text);
                    Rdv3Skin.Tag(g, val.Text, fTag, kind, new Rectangle(cell.X, y + (RowH - tz.Height) / 2, Math.Min(tz.Width, cell.Width), tz.Height), sc);
                }
                else
                {
                    Font font = (col.Bold && val.Tone == Rdv3Value.Normal) ? fTdBold : fTd;
                    Color ink = (val.Tone == Rdv3Value.Error) ? Rdv3Skin.Danger
                        : (val.Tone == Rdv3Value.Muted) ? Rdv3Skin.N400 : (col.Muted ? Rdv3Skin.N600 : Rdv3Skin.Ink);
                    Rdv3Skin.DrawIn(g, val.Text, font, ink, cell, right ? 2 : 0);
                }
            }
            y += RowH;
        }

        // the sticky header
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, 0, 0, ClientSize.Width, HeaderH); }
        for (int c = 0; c < def.Columns.Count; c++)
        {
            Rdv3ColumnDef col = def.Columns[c];
            bool right = (col.Align == "right");
            Rectangle cell = new Rectangle(xs[c] + padL, 0, Math.Max(0, xs[c + 1] - xs[c] - padL - (right ? padR : padL)), HeaderH);
            Rdv3Skin.DrawIn(g, col.Header, fTh, Rdv3Skin.N600, cell, right ? 2 : 0);
        }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, 0, HeaderH - 1, ClientSize.Width, 1);

        if (NeedBar)
        {
            Rectangle tr = BarTrack(), th = BarThumb();
            using (SolidBrush b = new SolidBrush(Color.FromArgb(10, Rdv3Skin.Ink))) { g.FillRectangle(b, tr); }
            using (SolidBrush b = new SolidBrush(Color.FromArgb(dragging ? 92 : 56, Rdv3Skin.Ink))) { g.FillRectangle(b, th); }
        }
    }
}

// ---------------------------------------------------------------------------
// the candidate list modal: 980 wide, header row, rule, scrolling table
// ---------------------------------------------------------------------------
public sealed class Rdv3CandidatesForm : Rdv3Dialog
{
    private readonly Rdv3CandidatesDef def;
    private readonly Rdv3CandTable table;
    private readonly Rdv3Btn btnClose;
    private readonly int total;
    private Font fTitle, fHint, fTag;
    public int Picked = -1;

    private const double HeadH = 48.4;      // 12 + 26.4 + 10

    public static int Pick(Rdv3Form owner, Rdv3CandidatesDef def, List<Rdv3CandRow> rows, int total, int selected)
    {
        Rdv3CandidatesForm f = new Rdv3CandidatesForm(owner, def, owner.Fields, owner.Screen.Work, rows, total, selected, owner.Sc);
        f.ShowOver(owner);
        int picked = f.Picked;
        f.Dispose();
        return picked;
    }

    private Rdv3CandidatesForm(Rdv3Form owner, Rdv3CandidatesDef d, Rdv3Fields fields, Rdv3WorkState work,
                               List<Rdv3CandRow> rows, int totalHits, int selected, float screenScale)
    {
        def = d;
        total = totalHits;
        double tableH = Math.Min(d.MaxHeight, d.HeaderHeight + Math.Max(1, rows.Count) * d.RowHeight);
        Sc = FitScale(d.Width, HeadH + 1 + tableH + 0.6, screenScale, owner);
        Text = d.Title;
        fTitle = Sf(22);
        fHint = Fn(15);
        fTag = Fn(13);
        SizeTo(d.Width, HeadH + 1 + tableH + 0.6);

        table = new Rdv3CandTable(d, fields, work, Sc);
        table.OnPick = delegate(int i) { Picked = i; DialogResult = DialogResult.OK; Close(); };
        table.SetRows(rows, selected);
        Controls.Add(table);

        btnClose = MakeBtn("", Rdv3Btn.Round, 13, "close");
        btnClose.IconCss = 15;
        btnClose.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };
        Layout1();
    }

    public static Rdv3CandidatesForm ForCheck(Rdv3CandidatesDef d, Rdv3Fields fields, Rdv3WorkState work,
                                              List<Rdv3CandRow> rows, int total, int selected, float scale)
    {
        Rdv3CandidatesForm f = new Rdv3CandidatesForm(null, d, fields, work, rows, total, selected, scale);
        return f;
    }

    public int[] ColumnEdges() { return table.ColumnEdges(); }

    private void Layout1()
    {
        ClearRects();
        Put("cl.head", 0, 0, CssW, HeadH);
        double x = 14.2;
        double tw = MW(def.Title, fTitle);
        Put("cl.title", x, 12, tw, 26.4);
        x += tw + 10.2;
        string tag = Rdv3Text.CandCount.Replace("{n}", total.ToString("N0", CultureInfo.InvariantCulture));
        Size tz = Rdv3Skin.TagSize(tag, fTag, Sc);
        double tagW = tz.Width / (double)Sc, tagH = tz.Height / (double)Sc;
        double rx = CssW - 13.6;
        Put("cl.close", rx - 30, (HeadH - 26) / 2, 30, 26);
        rx -= 30 + 10.2;
        Put("cl.tag", rx - tagW, (HeadH - tagH) / 2, tagW, tagH);
        rx -= tagW + 10.2;
        double hintW = Math.Min(MW(def.Hint, fHint), Math.Max(0, rx - x));
        Put("cl.hint", x, (HeadH - 23.25) / 2, hintW, 23.25);
        Put("cl.rule", 0.6, HeadH, CssW - 1.2, 1);
        Put("cl.table", 0, HeadH + 1, CssW, CssH - HeadH - 1 - 0.6);
        Place(table, "cl.table");
        Place(btnClose, "cl.close");
        DragZone = new Rectangle(0, 0, ClientSize.Width, PX(HeadH));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        ClearClip();
        using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, ClientRectangle); }
        T(g, "cl.title", def.Title, fTitle, Rdv3Skin.Ink, 0);
        T(g, "cl.hint", def.Hint, fHint, Rdv3Skin.Mix(Rdv3Skin.White, 0.55), 0);
        string tag = Rdv3Text.CandCount.Replace("{n}", total.ToString("N0", CultureInfo.InvariantCulture));
        Rdv3Skin.Tag(g, tag, fTag, Rdv3Skin.TagNeutral, At("cl.tag"), Sc);
        Rule(g, "cl.rule");
    }
}

// ---------------------------------------------------------------------------
// the toast: bottom right of the card, a tag (done / error), bold text, a close
// mark; slides in, stays durationMs, slides out. One at a time.
// ---------------------------------------------------------------------------
public sealed class Rdv3Toast : Form
{
    private readonly Rdv3Form owner;
    private readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
    private readonly System.Windows.Forms.Timer anim = new System.Windows.Forms.Timer();
    private string text = "";
    private bool note;
    private int durationMs = 3600;
    private Font fTag, fText;
    private Rectangle closeRect;
    private bool closeHot;
    private int animStep;
    private bool leaving;
    private float sc = 1f;

    public Rdv3Toast(Rdv3Form f)
    {
        owner = f;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        BackColor = Rdv3Skin.White;
        TopMost = false;
        timer.Tick += delegate { timer.Stop(); Dismiss(); };
        anim.Interval = 25;
        anim.Tick += delegate { Step(); };
        owner.LocationChanged += delegate { if (Visible) { PlaceOnCard(); } };
        owner.Resize += delegate { if (Visible) { PlaceOnCard(); } };
    }

    protected override bool ShowWithoutActivation { get { return true; } }

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams cp = base.CreateParams;
            cp.ClassStyle |= 0x00020000;         // CS_DROPSHADOW
            cp.ExStyle |= 0x08000000;            // WS_EX_NOACTIVATE
            return cp;
        }
    }

    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        Rdv3Corners.Round(this);
    }

    public void Show(string message, bool completion, int ms)
    {
        text = (message == null) ? "" : message.Replace("\r", " ").Replace("\n", " ");
        note = completion;
        durationMs = Math.Max(500, ms);
        Text = text;
        sc = owner.Sc;
        fTag = Rdv3Skin.F(13, FontStyle.Regular, sc);
        fText = Rdv3Skin.F(16, FontStyle.Bold, sc);
        Measure1();
        PlaceOnCard();
        leaving = false;
        animStep = 0;
        Opacity = 0.0;
        if (!Visible) { Show(owner); }
        else { Invalidate(); }
        anim.Start();
        timer.Stop();
        timer.Interval = durationMs;
        timer.Start();
    }

    private int P(double css) { return (int)Math.Round(css * sc); }

    private void Measure1()
    {
        int pad = P(12), gap = P(10.2);
        Size tz = Rdv3Skin.TagSize(note ? Rdv3Text.TagDone : Rdv3Text.TagError, fTag, sc);
        int tw = Rdv3Skin.Measure(text, fText).Width;
        int maxW = P(440);
        int w = Math.Min(maxW, pad + tz.Width + gap + tw + gap + P(22) + pad);
        int h = P(10) * 2 + Math.Max(tz.Height, Rdv3Skin.Measure("Ag", fText).Height) + P(4);
        ClientSize = new Size(w, h);
        closeRect = new Rectangle(w - pad - P(22), (h - P(22)) / 2, P(22), P(22));
    }

    private void PlaceOnCard()
    {
        Rectangle card = owner.CardBounds;
        Location = new Point(card.Right - P(16) - Width, card.Bottom - P(52) - Height);
    }

    private void Step()
    {
        animStep++;
        double t = Math.Min(1.0, animStep / 10.0);
        Opacity = leaving ? (1.0 - t) : t;
        if (t >= 1.0)
        {
            anim.Stop();
            if (leaving) { Hide(); }
        }
    }

    public void Dismiss()
    {
        if (!Visible || leaving) { return; }
        leaving = true;
        animStep = 0;
        anim.Start();
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        bool h = closeRect.Contains(e.Location);
        if (h != closeHot) { closeHot = h; Invalidate(closeRect); }
        base.OnMouseMove(e);
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        if (closeHot) { closeHot = false; Invalidate(closeRect); }
        base.OnMouseLeave(e);
    }

    protected override void OnMouseClick(MouseEventArgs e)
    {
        if (closeRect.Contains(e.Location)) { timer.Stop(); Dismiss(); }
        base.OnMouseClick(e);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, ClientRectangle); }
        int pad = P(12), gap = P(10.2);
        string tag = note ? Rdv3Text.TagDone : Rdv3Text.TagError;
        Size tz = Rdv3Skin.TagSize(tag, fTag, sc);
        Rectangle tr = new Rectangle(pad, (Height - tz.Height) / 2, tz.Width, tz.Height);
        Rdv3Skin.Tag(g, tag, fTag, note ? Rdv3Skin.TagAccent : Rdv3Skin.TagOutline, tr, sc);
        int tx = tr.Right + gap;
        Rectangle textRect = new Rectangle(tx, 0, Math.Max(0, closeRect.X - gap - tx), Height);
        Rdv3Skin.DrawIn(g, text, fText, Rdv3Skin.Accent800, textRect, 0);
        Rdv3Skin.FillRound(g, Rdv3Skin.Mix(Rdv3Skin.White, closeHot ? 0.12 : 0.06), closeRect, closeRect.Height / 2);
        int m = P(7);
        Rdv3Skin.IconClose(g, new Rectangle(closeRect.X + m, closeRect.Y + m, closeRect.Width - 2 * m, closeRect.Height - 2 * m), Rdv3Skin.N700, sc);
    }
}
