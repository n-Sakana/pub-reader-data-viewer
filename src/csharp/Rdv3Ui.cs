// ============================================================================
// Rdv3Ui.cs -- the drawing kit and the main screen, built from the screen
// definition (Rdv3Screen) rather than from coordinates written into the code.
//
// The reference is "Reader Data Viewer v2 (standalone).html": a 1240 px card
// with white rounded cards, pill tags, rounded buttons, a status band and a
// status bar. In the browser the card floats on a grey page; here the window
// IS the card: its content fills the client area edge to edge under the
// ordinary window frame, nothing is drawn around it. Native WinForms only, no
// WebView. Everything is painted with GDI+ from tokens scaled to the window's
// DPI; the only live controls are the ones that must take keyboard input or
// scroll: the search box, the buttons and the long-text boxes.
//
// Layout: the card is a vertical stack of sections. Each section type knows
// its own height from its definition (title 48, key panel from its padding and
// type sizes, field lists from their rows, text boxes from their box height,
// the band and the bar from theirs); a columns section is as tall as its
// tallest child. The numbers that make this card 1240 wide at the shipped
// definition are the reference's own CSS values, so the screen lands on the
// reference's coordinates without a measured table.
//
// The screen follows the window (Rdv3Form.Layout1 -> Rdv3Card.Fit):
//   wider     every section stretches; narrower than a columns section's
//             stackBelow, its items stack; a key panel whose figure and input
//             group no longer fit side by side wraps the group below
//   taller    the spare height goes to the long-text boxes
//   shorter   the text boxes give height first (down to two lines), then the
//             field rows and the band shrink a little, and only what still does
//             not fit scrolls
//
// Every public method of the form marshals itself onto the UI thread: the
// worker calls them directly and never touches a control.
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
// the design tokens of the reference, and the drawing primitives
// ---------------------------------------------------------------------------
internal static class Rdv3Skin
{
    public static readonly Color Ink = Color.FromArgb(0x1D, 0x1F, 0x20);
    public static readonly Color White = Color.White;
    public static readonly Color Bg = Color.FromArgb(0xF2, 0xF2, 0xF3);        // --color-bg
    public static readonly Color Surface = Color.FromArgb(0xE9, 0xE9, 0xEA);   // --color-surface
    public static readonly Color N100 = Color.FromArgb(0xF5, 0xF5, 0xF8);
    public static readonly Color N200 = Color.FromArgb(0xE7, 0xE7, 0xEA);      // the card's ground
    public static readonly Color N300 = Color.FromArgb(0xD4, 0xD4, 0xD7);      // the page
    public static readonly Color N400 = Color.FromArgb(0xB7, 0xB7, 0xBA);
    public static readonly Color N500 = Color.FromArgb(0x98, 0x98, 0x9B);
    public static readonly Color N600 = Color.FromArgb(0x7A, 0x7A, 0x7D);
    public static readonly Color N700 = Color.FromArgb(0x5D, 0x5D, 0x60);
    public static readonly Color N800 = Color.FromArgb(0x42, 0x42, 0x44);
    public static readonly Color N900 = Color.FromArgb(0x2B, 0x2B, 0x2D);
    public static readonly Color Accent = Color.FromArgb(0x59, 0x80, 0xA6);
    public static readonly Color Accent100 = Color.FromArgb(0xEE, 0xF6, 0xFF);
    public static readonly Color Accent300 = Color.FromArgb(0xB5, 0xD9, 0xFD);
    public static readonly Color Accent600 = Color.FromArgb(0x59, 0x7E, 0xA3);
    public static readonly Color Accent700 = Color.FromArgb(0x41, 0x61, 0x80);
    public static readonly Color Accent800 = Color.FromArgb(0x2C, 0x45, 0x5D);
    // the one colour the reference does not have: the error tone of an
    // unresolved column or an unknown stored state, a muted brick red
    public static readonly Color Danger = Color.FromArgb(0xB0, 0x4A, 0x3E);
    public static readonly Color DangerBack = Color.FromArgb(0xFB, 0xEE, 0xEC);

    // color-mix(in srgb, ink N%, transparent) over a ground
    public static Color Mix(Color ground, double f)
    {
        return Color.FromArgb(
            (int)Math.Round(ground.R * (1 - f) + Ink.R * f),
            (int)Math.Round(ground.G * (1 - f) + Ink.G * f),
            (int)Math.Round(ground.B * (1 - f) + Ink.B * f));
    }

    public static Color MixWith(Color ground, Color tint, double f)
    {
        return Color.FromArgb(
            (int)Math.Round(ground.R * (1 - f) + tint.R * f),
            (int)Math.Round(ground.G * (1 - f) + tint.G * f),
            (int)Math.Round(ground.B * (1 - f) + tint.B * f));
    }

    public static readonly Color Divider = Mix(White, 0.16);
    public static readonly Color CardEdge = Mix(White, 0.07);
    public static readonly Color RowLine = Mix(White, 0.09);
    public static readonly Color BtnBack = Mix(White, 0.06);
    public static readonly Color BtnHover = Mix(White, 0.07);
    public static readonly Color BtnPress = Mix(White, 0.14);

    public static float Scale = 1f;
    public static int P(double cssPx) { return (int)Math.Round(cssPx * Scale); }
    public static int P(double cssPx, float sc) { return (int)Math.Round(cssPx * sc); }

    // ---- fonts ---------------------------------------------------------------
    // Yu Gothic UI is the face Windows itself is set in: familiar, and crisp
    // for kana, kanji and Latin alike under GDI's ClearType; its 600 weight is
    // the Semibold cut. The family is the screen definition's "card.font"
    // (Meiryo is the other Windows-standard choice); a machine without the
    // named face falls back down the list.
    private static string family = "Yu Gothic UI";
    private static string semiFamily = "Yu Gothic UI Semibold";
    private static bool semiIsBold;

    private static string FirstInstalled(string[] want, string fallback)
    {
        for (int i = 0; i < want.Length; i++)
        {
            try
            {
                using (FontFamily f = new FontFamily(want[i])) { return f.Name; }
            }
            catch (Exception) { }
        }
        return fallback;
    }

    public static void PickFamily(string preferred)
    {
        List<string> want = new List<string>();
        if (preferred != null && preferred.Length > 0) { want.Add(preferred); }
        want.Add("Yu Gothic UI"); want.Add("Meiryo"); want.Add("Segoe UI");
        family = FirstInstalled(want.ToArray(), FontFamily.GenericSansSerif.Name);
        string dedicated = FirstInstalled(new string[] { family + " Semibold" }, "");
        if (dedicated.Length > 0) { semiFamily = dedicated; semiIsBold = false; }
        else { semiFamily = family; semiIsBold = true; }
        ResetMetrics();
    }

    public static string Family { get { return family; } }

    // A Font here is a DESCRIPTOR: face, size in CSS px times the scale, and
    // weight. The live edit controls use it as is; everything painted goes
    // through the GDI font it describes (below).
    public static Font F(double cssPx, FontStyle st, float sc)
    {
        return new Font(family, (float)(cssPx * sc), st, GraphicsUnit.Pixel);
    }

    public static Font F(double cssPx, FontStyle st) { return F(cssPx, st, Scale); }

    // weight 600
    public static Font S(double cssPx, float sc)
    {
        return new Font(semiFamily, (float)(cssPx * sc), semiIsBold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
    }

    public static Font S(double cssPx) { return S(cssPx, Scale); }

    public static Font Px(double cssPx, float sc) { return F(cssPx, FontStyle.Regular, sc); }

    // ---- measuring and drawing text: GDI, at an exact pixel height --------------
    // Text is set with GDI, the way native controls are: ClearType, hinted,
    // crisp. The font is created straight from the descriptor with lfHeight =
    // the pixel size, so it is the same height on every surface regardless
    // of DPI context (TextRenderer re-derives the height from points and the
    // thread's DPI, which put the live window's text at the wrong size), and
    // GetTextExtentPoint32 measures the run exactly, with no padding to take
    // off. Measuring and drawing use the same HFONT, so they agree.
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct LOGFONT
    {
        public int lfHeight, lfWidth, lfEscapement, lfOrientation, lfWeight;
        public byte lfItalic, lfUnderline, lfStrikeOut, lfCharSet, lfOutPrecision, lfClipPrecision, lfQuality, lfPitchAndFamily;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string lfFaceName;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct SIZE { public int cx, cy; }

    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr CreateFontIndirectW(ref LOGFONT lf);
    [DllImport("gdi32.dll")]
    private static extern IntPtr SelectObject(IntPtr hdc, IntPtr h);
    [DllImport("gdi32.dll")]
    private static extern int SetBkMode(IntPtr hdc, int mode);
    [DllImport("gdi32.dll")]
    private static extern uint SetTextColor(IntPtr hdc, uint color);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern bool GetTextExtentPoint32W(IntPtr hdc, string s, int len, out SIZE size);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
    private static extern bool ExtTextOutW(IntPtr hdc, int x, int y, uint options, IntPtr rect, string s, uint len, IntPtr dx);

    private sealed class GdiFont
    {
        public IntPtr Handle;
        public int Height;       // the cell height (ascent + descent)
    }

    private static readonly Dictionary<string, GdiFont> fonts = new Dictionary<string, GdiFont>();
    private static readonly Dictionary<string, Size> memo = new Dictionary<string, Size>();
    private static Graphics scratch;
    private static IntPtr scratchDc;

    public static void ResetMetrics()
    {
        lock (memo) { memo.Clear(); }
    }

    private static IntPtr ScratchDc()
    {
        if (scratchDc == IntPtr.Zero)
        {
            scratch = Graphics.FromImage(new Bitmap(1, 1));
            scratchDc = scratch.GetHdc();
        }
        return scratchDc;
    }

    private static string FontKey(Font f)
    {
        return f.Name + "|" + f.Size.ToString("0.##", CultureInfo.InvariantCulture) + "|" + ((int)f.Style).ToString(CultureInfo.InvariantCulture);
    }

    private static GdiFont Gdi(Font f)
    {
        string k = FontKey(f);
        GdiFont g;
        lock (fonts)
        {
            if (fonts.TryGetValue(k, out g)) { return g; }
            LOGFONT lf = new LOGFONT();
            lf.lfHeight = -(int)Math.Round(f.Size);
            lf.lfWeight = f.Bold ? 700 : 400;
            lf.lfItalic = (byte)(f.Italic ? 1 : 0);
            lf.lfCharSet = 1;                                   // DEFAULT_CHARSET
            lf.lfQuality = 5;                                   // CLEARTYPE_QUALITY
            lf.lfFaceName = f.Name;
            g = new GdiFont();
            g.Handle = CreateFontIndirectW(ref lf);
            IntPtr dc = ScratchDc();
            IntPtr old = SelectObject(dc, g.Handle);
            SIZE z;
            GetTextExtentPoint32W(dc, "Ag", 2, out z);
            g.Height = z.cy;
            SelectObject(dc, old);
            fonts[k] = g;
            return g;
        }
    }

    public static Size Measure(string s, Font f)
    {
        GdiFont gf = Gdi(f);
        if (s == null || s.Length == 0) { return new Size(0, gf.Height); }
        string k = FontKey(f) + "|" + s;
        Size hit;
        lock (memo)
        {
            if (memo.TryGetValue(k, out hit)) { return hit; }
        }
        lock (fonts)
        {
            IntPtr dc = ScratchDc();
            IntPtr old = SelectObject(dc, gf.Handle);
            SIZE z;
            GetTextExtentPoint32W(dc, s, s.Length, out z);
            SelectObject(dc, old);
            hit = new Size(z.cx, gf.Height);
        }
        lock (memo)
        {
            if (memo.Count > 8000) { memo.Clear(); }
            memo[k] = hit;
        }
        return hit;
    }

    public static int HeightOf(Font f) { return Gdi(f).Height; }

    public static void Draw(Graphics g, string s, Font f, Color c, int x, int y)
    {
        if (s == null || s.Length == 0) { return; }
        GdiFont gf = Gdi(f);
        IntPtr dc = g.GetHdc();
        try
        {
            IntPtr old = SelectObject(dc, gf.Handle);
            SetBkMode(dc, 1);                                   // TRANSPARENT
            SetTextColor(dc, (uint)(c.R | (c.G << 8) | (c.B << 16)));
            ExtTextOutW(dc, x, y, 0, IntPtr.Zero, s, (uint)s.Length, IntPtr.Zero);
            SelectObject(dc, old);
        }
        finally { g.ReleaseHdc(dc); }
    }

    private const string Ellipsis = "\u2026";

    // the ellipsis, with the same metric as everything else
    public static string Fit(string s, Font f, int w)
    {
        if (s == null || s.Length == 0) { return ""; }
        if (Measure(s, f).Width <= w) { return s; }
        int lo = 0, hi = s.Length;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) / 2;
            if (Measure(s.Substring(0, mid) + Ellipsis, f).Width <= w) { lo = mid; } else { hi = mid - 1; }
        }
        return (lo <= 0) ? "" : (s.Substring(0, lo) + Ellipsis);
    }

    // inside a box, vertically centred, cut with an ellipsis rather than
    // spilling into the neighbour: align 0 left, 1 centre, 2 right
    public static void DrawIn(Graphics g, string s, Font f, Color c, Rectangle r, int align)
    {
        if (s == null || s.Length == 0 || r.Width <= 0 || r.Height <= 0) { return; }
        string t = Fit(s, f, r.Width);
        if (t.Length == 0) { return; }
        Size z = Measure(t, f);
        int x = (align == 2) ? (r.Right - z.Width) : (align == 1) ? (r.X + (r.Width - z.Width) / 2) : r.X;
        Draw(g, t, f, c, x, r.Y + (r.Height - z.Height) / 2);
    }

    // word-wrapped paragraph inside a box; returns the height it used
    public static int DrawWrapped(Graphics g, string s, Font f, Color c, Rectangle r, double lineCss, float sc)
    {
        if (s == null || s.Length == 0 || r.Width <= 0) { return 0; }
        int lh = P(lineCss, sc);
        int y = r.Y;
        string[] paras = s.Replace("\r\n", "\n").Split('\n');
        for (int p = 0; p < paras.Length; p++)
        {
            string para = paras[p];
            int at = 0;
            if (para.Length == 0) { y += lh; continue; }
            while (at < para.Length)
            {
                int take = FitCount(para, at, f, r.Width);
                if (y + lh > r.Bottom + 1) { return y - r.Y; }
                Draw(g, para.Substring(at, take), f, c, r.X, y + (lh - HeightOf(f)) / 2);
                at += take;
                y += lh;
            }
        }
        return y - r.Y;
    }

    // how many characters of s from `from` fit in w. A browser breaks
    // Japanese text anywhere and Latin words at spaces: only step back to a
    // space when the cut would split a Latin word.
    public static int FitCount(string s, int from, Font f, int w)
    {
        int lo = 1, hi = s.Length - from;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) / 2;
            if (Measure(s.Substring(from, mid), f).Width <= w) { lo = mid; } else { hi = mid - 1; }
        }
        int take = Math.Max(1, lo);
        if (from + take < s.Length && IsWordChar(s[from + take - 1]) && IsWordChar(s[from + take]))
        {
            int sp = s.LastIndexOf(' ', from + take - 1, take);
            if (sp > from) { take = sp - from + 1; }
        }
        return take;
    }

    private static bool IsWordChar(char c)
    {
        return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_' || c == '-';
    }

    // the lines a wrapped paragraph takes in w
    public static int CountLines(string s, Font f, int w)
    {
        if (s == null) { return 1; }
        int n = 0;
        string[] paras = s.Replace("\r\n", "\n").Split('\n');
        for (int p = 0; p < paras.Length; p++)
        {
            if (paras[p].Length == 0) { n++; continue; }
            int at = 0;
            while (at < paras[p].Length) { at += FitCount(paras[p], at, f, w); n++; }
        }
        return Math.Max(1, n);
    }

    // ---- shapes ------------------------------------------------------------
    public static GraphicsPath Round(Rectangle r, int radius)
    {
        GraphicsPath p = new GraphicsPath();
        int d = Math.Max(0, Math.Min(radius * 2, Math.Min(r.Width, r.Height)));
        if (d <= 1) { p.AddRectangle(r); return p; }
        p.AddArc(r.X, r.Y, d, d, 180, 90);
        p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        p.CloseFigure();
        return p;
    }

    public static void FillRound(Graphics g, Color c, Rectangle r, int radius)
    {
        if (r.Width <= 0 || r.Height <= 0) { return; }
        using (GraphicsPath p = Round(r, radius))
        using (SolidBrush b = new SolidBrush(c)) { g.FillPath(b, p); }
    }

    public static void FrameRound(Graphics g, Color c, Rectangle r, int radius)
    {
        if (r.Width <= 1 || r.Height <= 1) { return; }
        Rectangle inner = new Rectangle(r.X, r.Y, r.Width - 1, r.Height - 1);
        using (GraphicsPath p = Round(inner, radius))
        using (Pen pen = new Pen(c, 1f)) { g.DrawPath(pen, p); }
    }

    public static void Line(Graphics g, Color c, int x, int y, int w, int h)
    {
        if (w <= 0 || h <= 0) { return; }
        using (SolidBrush b = new SolidBrush(c)) { g.FillRectangle(b, x, y, w, h); }
    }

    // ---- tags: the reference's pills (13 px, padding 4 x 12, radius 999) ----
    public const int TagAccent = 0;
    public const int TagOutline = 1;
    public const int TagNeutral = 2;
    public const int TagFaded = 3;
    public const int TagError = 4;

    public static int TagKind(string look)
    {
        if (look == "accent") { return TagAccent; }
        if (look == "outline") { return TagOutline; }
        if (look == "faded") { return TagFaded; }
        if (look == "error") { return TagError; }
        return TagNeutral;
    }

    public static Size TagSize(string text, Font f, float sc)
    {
        Size s = Measure(text, f);
        return new Size(s.Width + P(24, sc), s.Height + P(8, sc));
    }

    public static void Tag(Graphics g, string text, Font f, int kind, Rectangle r, float sc)
    {
        Color back, ink, edge;
        if (kind == TagAccent) { back = Accent100; ink = Accent800; edge = Color.Empty; }
        else if (kind == TagOutline) { back = Color.Empty; ink = Accent; edge = Accent; }
        else if (kind == TagError) { back = DangerBack; ink = Danger; edge = Color.Empty; }
        else { back = N100; ink = N800; edge = Color.Empty; }
        if (kind == TagFaded)
        {
            // opacity .5 over white
            back = MixWith(White, back, 0.5);
            ink = MixWith(White, ink, 0.5);
        }
        int radius = r.Height / 2;
        if (back != Color.Empty) { FillRound(g, back, r, radius); }
        if (edge != Color.Empty) { FrameRound(g, edge, r, radius); }
        Size ts = Measure(text, f);
        Draw(g, text, f, ink, r.X + (r.Width - ts.Width) / 2, r.Y + (r.Height - ts.Height) / 2);
    }

    // ---- the line icons the reference draws as SVG ---------------------------
    private static Pen IconPen(Color c, float sc)
    {
        Pen p = new Pen(c, Math.Max(1f, 1.25f * sc));
        p.StartCap = LineCap.Round; p.EndCap = LineCap.Round; p.LineJoin = LineJoin.Round;
        return p;
    }

    public static void IconSearch(Graphics g, Rectangle r, Color c, float sc)
    {
        using (Pen p = IconPen(c, sc))
        {
            float d = r.Width * 0.66f;
            g.DrawEllipse(p, r.X + r.Width * 0.05f, r.Y + r.Height * 0.05f, d, d);
            g.DrawLine(p, r.X + r.Width * 0.62f, r.Y + r.Height * 0.62f, r.Right - r.Width * 0.08f, r.Bottom - r.Height * 0.08f);
        }
    }

    public static void IconRefresh(Graphics g, Rectangle r, Color c, float sc)
    {
        // lucide refresh-cw: two 135-degree arcs, each ending in an arrowhead
        using (Pen p = IconPen(c, sc))
        {
            RectangleF a = new RectangleF(r.X + r.Width * 0.12f, r.Y + r.Height * 0.12f, r.Width * 0.76f, r.Height * 0.76f);
            g.DrawArc(p, a, 180, 135);
            g.DrawArc(p, a, 0, 135);
            float t = r.Width * 0.2f;
            PointF tip1 = new PointF(a.Right, a.Y + a.Height * 0.33f);
            g.DrawLine(p, tip1, new PointF(tip1.X, tip1.Y - t));
            g.DrawLine(p, tip1, new PointF(tip1.X - t, tip1.Y));
            PointF tip2 = new PointF(a.X, a.Bottom - a.Height * 0.33f);
            g.DrawLine(p, tip2, new PointF(tip2.X, tip2.Y + t));
            g.DrawLine(p, tip2, new PointF(tip2.X + t, tip2.Y));
        }
    }

    public static void IconGear(Graphics g, Rectangle r, Color c, float sc)
    {
        float cx = r.X + r.Width / 2f, cy = r.Y + r.Height / 2f;
        float rad = Math.Min(r.Width, r.Height) / 2f;
        float ring = rad * 0.62f;
        using (Pen pen = IconPen(c, sc))
        {
            g.DrawEllipse(pen, cx - rad * 0.25f, cy - rad * 0.25f, rad * 0.5f, rad * 0.5f);
            g.DrawEllipse(pen, cx - ring, cy - ring, ring * 2f, ring * 2f);
            for (int i = 0; i < 8; i++)
            {
                double a = Math.PI * i / 4.0;
                float dx = (float)Math.Cos(a), dy = (float)Math.Sin(a);
                g.DrawLine(pen, cx + dx * ring, cy + dy * ring, cx + dx * rad, cy + dy * rad);
            }
        }
    }

    public static void IconClose(Graphics g, Rectangle r, Color c, float sc)
    {
        using (Pen p = IconPen(c, sc))
        {
            float m = r.Width * 0.25f;
            g.DrawLine(p, r.X + m, r.Y + m, r.Right - m, r.Bottom - m);
            g.DrawLine(p, r.Right - m, r.Y + m, r.X + m, r.Bottom - m);
        }
    }

    // the check-circle of an OK band
    public static void IconCheckCircle(Graphics g, Rectangle r, Color c, float sc)
    {
        using (Pen p = IconPen(c, sc))
        {
            p.Width = Math.Max(1f, 2f * sc);
            g.DrawEllipse(p, r.X + 1, r.Y + 1, r.Width - 2, r.Height - 2);
            g.DrawLines(p, new PointF[] {
                new PointF(r.X + r.Width * 0.30f, r.Y + r.Height * 0.50f),
                new PointF(r.X + r.Width * 0.44f, r.Y + r.Height * 0.64f),
                new PointF(r.X + r.Width * 0.70f, r.Y + r.Height * 0.38f) });
        }
    }
}

// ---------------------------------------------------------------------------
// A button with the reference's own states: rounded 10 px, no border.
//   primary    accent ground, white text; hover accent-600; active accent-700
//   secondary  ink 6% ground, accent-700 text; hover 7%; active 14%
//   soft       accent 12% ground (the settings modal's pick button); hover 20%
//   round      a circle (the modal close marks)
//   quiet      white ground with a hairline edge (the status bar's buttons,
//              which sit in a white bar); hover ink 4%, active 8%
//   :disabled  opacity .45
// ---------------------------------------------------------------------------
public sealed class Rdv3Btn : Button
{
    public const int Secondary = 0;
    public const int Primary = 1;
    public const int Soft = 2;
    public const int Round = 3;
    public const int Quiet = 4;

    public int Kind = Secondary;
    public string Icon = "";              // "" | search | refresh | gear | close
    public double IconCss = 13;
    public double PadCss = 16;            // the reference: 12 on the title bar, 16 in the key panel
    public double RadiusCss = 10;
    public float Sc;                      // 0 = the global scale
    private bool over, down;

    public Rdv3Btn()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        BackColor = Rdv3Skin.White;
        UseVisualStyleBackColor = false;
        Cursor = Cursors.Hand;
    }

    public float Sk { get { return (Sc > 0f) ? Sc : Rdv3Skin.Scale; } }

    // WinForms makes the focused button the form's default button and draws a
    // grey frame around it; this button has exactly one look per state
    public override void NotifyDefault(bool value) { base.NotifyDefault(false); }
    protected override bool ShowFocusCues { get { return false; } }

    protected override void OnMouseEnter(EventArgs e) { over = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { over = false; down = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseDown(MouseEventArgs e) { down = true; Invalidate(); base.OnMouseDown(e); }
    protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
    protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

    // the width the reference gives it: padding + icon + gap + text
    public int WantedWidth()
    {
        float sc = Sk;
        int pad = Rdv3Skin.P(PadCss, sc);
        int tw = Rdv3Skin.Measure(Text, Font).Width;
        int iw = (Icon.Length > 0) ? (Rdv3Skin.P(IconCss, sc) + Rdv3Skin.P((Kind == Primary) ? 7 : 6, sc)) : 0;
        return pad * 2 + iw + tw;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
        using (SolidBrush b = new SolidBrush(BackColor)) { g.FillRectangle(b, ClientRectangle); }
        if (Enabled) { Render(g); return; }
        using (Bitmap bmp = new Bitmap(Math.Max(1, Width), Math.Max(1, Height)))
        {
            using (Graphics gb = Graphics.FromImage(bmp))
            {
                gb.SmoothingMode = SmoothingMode.AntiAlias;
                gb.PixelOffsetMode = PixelOffsetMode.Half;
                using (SolidBrush b = new SolidBrush(BackColor)) { gb.FillRectangle(b, 0, 0, Width, Height); }
                Render(gb);
            }
            using (System.Drawing.Imaging.ImageAttributes ia = new System.Drawing.Imaging.ImageAttributes())
            {
                System.Drawing.Imaging.ColorMatrix m = new System.Drawing.Imaging.ColorMatrix();
                m.Matrix33 = 0.45f;
                ia.SetColorMatrix(m);
                g.DrawImage(bmp, new Rectangle(0, 0, Width, Height), 0, 0, Width, Height, GraphicsUnit.Pixel, ia);
            }
        }
    }

    private void Render(Graphics g)
    {
        float sc = Sk;
        Rectangle r = new Rectangle(0, 0, Width, Height);
        Color back, ink;
        if (Kind == Primary)
        {
            back = down ? Rdv3Skin.Accent700 : (over ? Rdv3Skin.Accent600 : Rdv3Skin.Accent);
            ink = Rdv3Skin.White;
        }
        else if (Kind == Soft)
        {
            back = Rdv3Skin.MixWith(Rdv3Skin.White, Rdv3Skin.Accent, down ? 0.24 : (over ? 0.20 : 0.12));
            ink = Rdv3Skin.Accent700;
        }
        else if (Kind == Round)
        {
            back = Rdv3Skin.Mix(Rdv3Skin.White, down ? 0.14 : (over ? 0.12 : 0.06));
            ink = Rdv3Skin.N700;
        }
        else if (Kind == Quiet)
        {
            back = down ? Rdv3Skin.Mix(Rdv3Skin.White, 0.08) : (over ? Rdv3Skin.Mix(Rdv3Skin.White, 0.04) : Rdv3Skin.White);
            ink = Rdv3Skin.Accent700;
        }
        else
        {
            back = down ? Rdv3Skin.BtnPress : (over ? Rdv3Skin.BtnHover : Rdv3Skin.BtnBack);
            ink = Rdv3Skin.Accent700;
        }
        int radius = (Kind == Round) ? (Math.Min(r.Width, r.Height) / 2) : Rdv3Skin.P(RadiusCss, sc);
        Rdv3Skin.FillRound(g, back, r, radius);
        if (Kind == Quiet) { Rdv3Skin.FrameRound(g, Rdv3Skin.Divider, r, radius); }

        int sz = Rdv3Skin.P(IconCss, sc), gap = Rdv3Skin.P((Kind == Primary) ? 7 : 6, sc);
        Size tz = Rdv3Skin.Measure(Text, Font);
        int all = (Icon.Length > 0) ? (sz + ((tz.Width > 0) ? (gap + tz.Width) : 0)) : tz.Width;
        int x = r.X + (r.Width - all) / 2;
        if (Icon.Length > 0)
        {
            Rectangle ir = new Rectangle(x, r.Y + (r.Height - sz) / 2, sz, sz);
            if (Icon == "search") { Rdv3Skin.IconSearch(g, ir, ink, sc); }
            else if (Icon == "refresh") { Rdv3Skin.IconRefresh(g, ir, ink, sc); }
            else if (Icon == "gear") { Rdv3Skin.IconGear(g, ir, ink, sc); }
            else if (Icon == "close") { Rdv3Skin.IconClose(g, ir, ink, sc); }
            x += sz + gap;
        }
        if (tz.Width > 0) { Rdv3Skin.Draw(g, Text, Font, ink, x, r.Y + (r.Height - tz.Height) / 2); }
    }
}

// ---------------------------------------------------------------------------
// one candidate row as the screen needs it: the ledger line and its state
// ---------------------------------------------------------------------------
public sealed class Rdv3CandRow
{
    public string Line;
    public string Stored;
}

// ---------------------------------------------------------------------------
// the main screen
// ---------------------------------------------------------------------------
public sealed class Rdv3Form : Form
{
    [DllImport("user32.dll")]
    private static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    // The .cmd runs inside PowerShell, whose process is already dpi-unaware by
    // the time this code exists, so SetProcessDPIAware cannot help: Windows
    // would stretch a 96-dpi bitmap and the screen would be soft. The THREAD
    // context can still be raised (Win10 1607+), and that is enough -- every
    // window this app creates belongs to this thread.
    [DllImport("user32.dll")]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr ctx);

    // EM_SETCUEBANNER: the edit control's own placeholder, shown while it is empty
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, string lParam);

    public readonly Rdv3Screen Screen;
    public readonly Rdv3View View = new Rdv3View();
    private Rdv3Fields fields = Rdv3Fields.Empty;

    // what the app reacts to
    public Action<string> OnSearch;
    public Action OnClear;
    public Action OnWorkState;
    public Action OnRefreshLedger;
    public Action OnSettings;
    public Action<int> OnPick;

    // live controls
    private readonly TextBox txtKey = new TextBox();
    private readonly List<Rdv3Btn> buttons = new List<Rdv3Btn>();
    private readonly Dictionary<Rdv3Btn, Rdv3ButtonDef> buttonDefs = new Dictionary<Rdv3Btn, Rdv3ButtonDef>();
    private Rdv3Btn btnWork;
    private readonly List<TextBox> boxes = new List<TextBox>();
    private readonly List<Rdv3Section> boxSections = new List<Rdv3Section>();
    private readonly ToolTip tips = new ToolTip();
    private readonly System.Windows.Forms.Timer clock = new System.Windows.Forms.Timer();
    private readonly Panel page = new Panel();          // scrolls when the card does not fit

    // the candidate list of the current search, for the modal and the app
    private List<Rdv3CandRow> cands = new List<Rdv3CandRow>();
    private int candTotal;

    // fonts at the current scale
    private Font fBrand, fTag, fTitleBtn, fCardTitle, fBig, fInput, fBtn, fRowLabel, fRowValue, fBox,
                 fBandLabel, fBandValue, fBandSub, fBar, fBarBold;

    private float sc = 1f;              // the card's scale (DPI, or less when the screen is small)
    private bool scaled;
    private bool inputHot;
    private string placeholder = "";

    public Rdv3Form(Rdv3Screen screen)
    {
        Screen = screen;
        MakeDpiAware();
        Rdv3Skin.PickFamily(screen.FontFamily);

        Text = Rdv3Text.AppTitle;
        StartPosition = FormStartPosition.Manual;
        BackColor = Rdv3Skin.N200;
        DoubleBuffered = true;
        AutoScaleMode = AutoScaleMode.None;
        KeyPreview = true;

        page.BackColor = Rdv3Skin.N200;
        page.AutoScroll = true;
        page.Dock = DockStyle.Fill;
        Controls.Add(page);

        card = new Rdv3Card(this, false);
        page.Controls.Add(card);
        // the status bar stays put under the scrolling body
        bar = new Rdv3Card(this, true);
        bar.Dock = DockStyle.Bottom;
        Controls.Add(bar);
        page.BringToFront();

        txtKey.BorderStyle = BorderStyle.None;
        txtKey.BackColor = Rdv3Skin.N200;
        txtKey.ForeColor = Rdv3Skin.Ink;
        txtKey.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Fire(OnSearch); }
        };
        txtKey.MouseEnter += delegate { inputHot = true; card.Invalidate(); };
        txtKey.MouseLeave += delegate { inputHot = false; card.Invalidate(); };
        txtKey.Enter += delegate { card.Invalidate(); };
        txtKey.Leave += delegate { card.Invalidate(); };
        card.Controls.Add(txtKey);

        BuildControls();
        UpdateWorkButton();

        clock.Interval = 1000;
        clock.Tick += delegate { bar.Invalidate(); };
        if (UsesClock()) { clock.Start(); }

        View.UserName = SafeEnv("USERNAME");
        View.HostName = SafeEnv("COMPUTERNAME");
        SetScale(1f);
    }

    private Rdv3Card card;
    private Rdv3Card bar;

    private static void MakeDpiAware()
    {
        IntPtr[] ctx = { new IntPtr(-4), new IntPtr(-3), new IntPtr(-2) };
        for (int i = 0; i < ctx.Length; i++)
        {
            try { if (SetThreadDpiAwarenessContext(ctx[i]) != IntPtr.Zero) { return; } }
            catch (Exception) { break; }
        }
        try { SetProcessDPIAware(); } catch (Exception) { }
    }

    private static string SafeEnv(string n)
    {
        try { string v = Environment.GetEnvironmentVariable(n); return (v == null) ? "" : v; }
        catch (Exception) { return ""; }
    }

    private void Fire(Action<string> a)
    {
        if (a != null) { a(txtKey.Text.Trim()); }
    }

    private bool UsesClock()
    {
        List<Rdv3Bind> all = Screen.AllBindings();
        for (int i = 0; i < all.Count; i++) { if (all[i].IsState && all[i].State == "clock") { return true; } }
        return false;
    }

    // ---- the live controls the definition asks for ------------------------------
    private void BuildControls()
    {
        for (int i = 0; i < Screen.Sections.Count; i++)
        {
            Rdv3Section s = Screen.Sections[i];
            if (s.Type == "titleBar" || s.Type == "keyPanel" || s.Type == "statusBar")
            {
                for (int k = 0; k < s.Buttons.Count; k++) { MakeButton(s.Buttons[k], s.Type != "keyPanel", s.Type == "statusBar"); }
                if (s.Type == "keyPanel") { txtKey.MaxLength = s.MaxLength; placeholder = s.Placeholder; }
            }
            else if (s.Type == "textBox") { MakeBox(s); }
            else if (s.Type == "columns")
            {
                for (int k = 0; k < s.Items.Count; k++) { if (s.Items[k].Type == "textBox") { MakeBox(s.Items[k]); } }
            }
        }
    }

    // bar: the compact cut of the title and status bars (13 px icon, 12 px
    // padding); the key panel's buttons are the larger cut
    private void MakeButton(Rdv3ButtonDef d, bool compact, bool onBar)
    {
        Rdv3Btn b = new Rdv3Btn();
        b.Kind = d.Primary ? Rdv3Btn.Primary : (compact ? Rdv3Btn.Quiet : Rdv3Btn.Secondary);
        b.Icon = d.Icon;
        b.IconCss = compact ? 13 : 15;
        b.PadCss = compact ? 12 : 16;
        b.Text = d.Text;
        b.TabStop = true;
        if (d.Tip.Length > 0) { tips.SetToolTip(b, d.Tip); }
        if (d.Action == "workState")
        {
            btnWork = b;
            if (Screen.Work != null && Screen.Work.ButtonTip.Length > 0) { tips.SetToolTip(b, Screen.Work.ButtonTip); }
        }
        b.Click += delegate { Act(d.Action); };
        buttons.Add(b);
        buttonDefs[b] = d;
        (onBar ? bar : card).Controls.Add(b);
    }

    private void Act(string action)
    {
        if (action == "search") { Fire(OnSearch); }
        else if (action == "clear") { if (OnClear != null) { OnClear(); } }
        else if (action == "workState") { if (OnWorkState != null) { OnWorkState(); } }
        else if (action == "refreshLedger") { if (OnRefreshLedger != null) { OnRefreshLedger(); } }
        else if (action == "settings") { if (OnSettings != null) { OnSettings(); } }
    }

    private void MakeBox(Rdv3Section s)
    {
        TextBox t = new TextBox();
        t.Multiline = true;
        t.ReadOnly = true;
        t.BorderStyle = BorderStyle.None;
        t.BackColor = Rdv3Skin.N100;
        t.ForeColor = Rdv3Skin.Ink;
        t.ScrollBars = ScrollBars.None;
        t.TabStop = false;
        boxes.Add(t);
        boxSections.Add(s);
        card.Controls.Add(t);
    }

    // WinForms shows a scrollbar for good, or never; the reference shows one
    // only when the text is longer than its box, so the decision is per text
    private void SetBox(TextBox t, string text, Color ink)
    {
        t.ForeColor = ink;
        t.Text = (text == null) ? "" : text;
        bool over = false;
        if (t.Text.Length > 0 && t.Width > 8)
        {
            // the lines the box needs at its width, with the same metric as the painter
            int lines = Rdv3Skin.CountLines(t.Text, t.Font, Math.Max(1, t.Width));
            over = lines * Rdv3Skin.HeightOf(t.Font) > t.Height;
        }
        ScrollBars want = over ? ScrollBars.Vertical : ScrollBars.None;
        if (t.ScrollBars != want) { t.ScrollBars = want; }
    }

    // ---- scale ------------------------------------------------------------------
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        try { SendMessage(txtKey.Handle, 0x1501, new IntPtr(1), placeholder); } catch (Exception) { }
        ApplyScale();
        FlushPending();
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        ApplyScale();
        FlushPending();
    }

    private void ApplyScale()
    {
        if (scaled || !IsHandleCreated) { return; }
        float s = 1f;
        try
        {
            uint d = GetDpiForWindow(Handle);
            if (d >= 72) { s = d / 96f; }
            else { using (Graphics g = Graphics.FromHwnd(Handle)) { s = g.DpiX / 96f; } }
        }
        catch (Exception)
        {
            try { using (Graphics g = Graphics.FromHwnd(Handle)) { s = g.DpiX / 96f; } }
            catch (Exception) { s = 1f; }
        }
        scaled = true;
        // the card keeps its proportions: when the work area cannot hold it at
        // the monitor's scale, the whole card is scaled down (never below 0.7)
        // and anything still over the edge scrolls
        Rectangle wa = System.Windows.Forms.Screen.FromControl(this).WorkingArea;
        Size frame = new Size(Math.Max(0, Width - ClientSize.Width), Math.Max(0, Height - ClientSize.Height));
        // the start-up size is the definition's; only if even that does not
        // fit the work area does the whole screen scale down
        double startW = Screen.StartWidth, startH = Screen.StartHeight;
        double needW = startW * s + frame.Width;
        double needH = startH * s + frame.Height;
        float fit = s;
        if (needW > wa.Width) { fit = Math.Min(fit, (float)((wa.Width - frame.Width) / startW)); }
        if (needH > wa.Height) { fit = Math.Min(fit, (float)((wa.Height - frame.Height) / startH)); }
        fit = Math.Max(0.7f, Math.Min(s, fit));
        SetScale(fit);
        Size want = new Size(PX(startW) + frame.Width, PX(startH) + frame.Height);
        want.Width = Math.Min(want.Width, wa.Width);
        want.Height = Math.Min(want.Height, wa.Height);
        Size = want;
        // the layout reflows down to MinCardCss wide; below that it scrolls
        MinimumSize = new Size(Math.Min(PX(MinCardCss) + frame.Width, wa.Width), Math.Min(PX(300) + frame.Height, wa.Height));
        Location = new Point(wa.X + Math.Max(0, (wa.Width - Width) / 2), wa.Y + Math.Max(0, (wa.Height - Height) / 2));
        Layout1();
    }

    // also the entry point the headless acceptance harness uses
    public void SetScale(float s)
    {
        sc = s;
        Rdv3Skin.Scale = s;
        MakeFonts();
        txtKey.Font = fInput;
        for (int i = 0; i < buttons.Count; i++)
        {
            Rdv3Btn b = buttons[i];
            b.Sc = s;
            b.Font = fTitleBtn;
        }
        for (int i = 0; i < boxes.Count; i++) { boxes[i].Font = fBox; }
        Layout1();
    }

    public float Sc { get { return sc; } }

    private int PX(double css) { return (int)Math.Round(css * sc); }

    private void MakeFonts()
    {
        Rdv3Skin.ResetMetrics();
        fBrand = Rdv3Skin.S(22, sc);
        fTag = Rdv3Skin.F(13, FontStyle.Regular, sc);
        fTitleBtn = Rdv3Skin.S(16, sc);
        fBtn = Rdv3Skin.S(16, sc);
        fCardTitle = Rdv3Skin.S(13, sc);
        fBig = Rdv3Skin.S(34, sc);
        fInput = Rdv3Skin.Px(16, sc);
        fRowLabel = Rdv3Skin.F(15, FontStyle.Regular, sc);
        fRowValue = Rdv3Skin.S(17, sc);
        fBox = Rdv3Skin.Px(16, sc);
        fBandLabel = Rdv3Skin.S(13, sc);
        fBandValue = Rdv3Skin.F(28, FontStyle.Bold, sc);
        fBandSub = Rdv3Skin.F(15, FontStyle.Regular, sc);
        fBar = Rdv3Skin.F(13, FontStyle.Regular, sc);
        fBarBold = Rdv3Skin.F(13, FontStyle.Bold, sc);
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); Layout1(); }

    // ---- layout ------------------------------------------------------------------
    // The card fills the window. A wider window widens every section (they
    // are laid out from the card's width); a taller one grows the long-text
    // boxes, which is where more room is useful; a smaller one scrolls.
    private void Layout1()
    {
        if (fBrand == null || card == null || bar == null) { return; }
        bar.Height = PX(bar.CardCssHeight);
        bar.Fit(Math.Max(MinCardCss, ClientSize.Width / (double)sc), bar.CardCssHeight);
        bar.Relayout();
        // The scroll decision is made here, from the page's outer size, rather
        // than left to AutoScroll's own inference from child bounds (which
        // cascades a vertical bar into a horizontal one and back). Sideways
        // scrolling only below the reflow floor; downwards only when the body
        // cannot compress enough.
        int sbW = PX(17), sbH = PX(17);
        try { sbW = SystemInformation.VerticalScrollBarWidth; sbH = SystemInformation.HorizontalScrollBarHeight; }
        catch (Exception) { }
        int availW = page.Width, availH = page.Height;
        bool hBar = availW < PX(MinCardCss);
        if (hBar) { availH -= sbH; }
        int cw = Math.Max(PX(MinCardCss), availW);
        card.Fit(cw / (double)sc, availH / (double)sc);
        bool vBar = PX(card.FittedCssHeight) > availH;
        if (vBar)
        {
            availW -= sbW;
            cw = Math.Max(PX(MinCardCss), availW);
            card.Fit(cw / (double)sc, availH / (double)sc);
        }
        int ch = Math.Max(PX(card.FittedCssHeight), availH);
        page.AutoScrollMinSize = new Size(hBar ? PX(MinCardCss) : 0, vBar ? ch : 0);
        Point off = page.AutoScrollPosition;
        card.SetBounds(off.X, off.Y, cw, ch);
        card.Relayout();
    }

    // below this the layout no longer reflows; it scrolls sideways instead
    private const double MinCardCss = 480;

    // the card: paints every section and places the live controls
    private sealed class Rdv3Card : Control
    {
        private readonly Rdv3Form f;
        // true: the status bar sections only (docked under the body); false: every other section
        private readonly bool isBar;
        public readonly Dictionary<string, Rectangle> Rc = new Dictionary<string, Rectangle>();
        public readonly List<string> Clipped = new List<string>();
        // the window's spare height, shared by the text boxes (CSS px); negative
        // while the boxes are giving height back
        public double ExtraCss;
        // how much every field row and the band give when the window is short
        public double RowGive;
        public double BandGive;
        // the width the layout was fitted to, for the stacking decisions
        private double fitW = 1240;
        // a text box is `lines` lines of the box font plus the box's own
        // padding (13 above, 11 below); under pressure it gives down to two
        // lines (one, when it was one)
        private const double BoxPadCss = 24;
        private double BoxLineCss { get { return Rdv3Skin.HeightOf(f.fBox) / (double)f.sc; } }
        private double NaturalTextCss(Rdv3Section s) { return BoxPadCss + s.Lines * BoxLineCss; }
        private double FloorTextCss(Rdv3Section s) { return BoxPadCss + Math.Min(2, s.Lines) * BoxLineCss; }
        private const double MinRowCss = 30;
        private const double MinBandCss = 60;

        public Rdv3Card(Rdv3Form form, bool statusBar)
        {
            f = form;
            isBar = statusBar;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
            BackColor = Rdv3Skin.N200;
        }

        // this card's sections, in order; their index in the definition is kept
        // for the rectangle keys
        private List<KeyValuePair<int, Rdv3Section>> Mine()
        {
            List<KeyValuePair<int, Rdv3Section>> list = new List<KeyValuePair<int, Rdv3Section>>();
            for (int i = 0; i < f.Screen.Sections.Count; i++)
            {
                Rdv3Section s = f.Screen.Sections[i];
                if ((s.Type == "statusBar") == isBar) { list.Add(new KeyValuePair<int, Rdv3Section>(i, s)); }
            }
            return list;
        }

        // ---- heights, in CSS px, straight from the definition -----------------
        private int TextBoxCount()
        {
            int n = 0;
            for (int i = 0; i < f.Screen.Sections.Count; i++)
            {
                Rdv3Section s = f.Screen.Sections[i];
                if (s.Type == "textBox") { n++; }
                for (int k = 0; k < s.Items.Count; k++) { if (s.Items[k].Type == "textBox") { n++; } }
            }
            return n;
        }

        private double ExtraPerBox()
        {
            int n = TextBoxCount();
            return (n == 0) ? 0 : ExtraCss / n;
        }

        // the definition's own height at the design width, without any fit
        public double CardCssHeight
        {
            get
            {
                double e = ExtraCss, r = RowGive, b = BandGive, w = fitW;
                ExtraCss = 0; RowGive = 0; BandGive = 0; fitW = f.Screen.CardWidth;
                double h = StackHeight();
                ExtraCss = e; RowGive = r; BandGive = b; fitW = w;
                return h;
            }
        }

        // the height after Fit, which is what the window shows
        public double FittedCssHeight { get { return StackHeight(); } }

        // Decide how the sections share a window of w x h (CSS px). Spare
        // height goes to the text boxes; a shortfall is taken from the text
        // boxes first, then the rows and the band, each down to its floor.
        public void Fit(double w, double h)
        {
            fitW = w;
            ExtraCss = 0; RowGive = 0; BandGive = 0;
            double natural = StackHeight();
            if (h >= natural) { ExtraCss = h - natural; return; }
            double deficit = natural - h;
            double textRoom = 0;
            Rdv3Screen s = f.Screen;
            for (int i = 0; i < s.Sections.Count; i++)
            {
                Rdv3Section sec = s.Sections[i];
                if (sec.Type == "textBox") { textRoom += Math.Max(0, NaturalTextCss(sec) - FloorTextCss(sec)); }
                for (int k = 0; k < sec.Items.Count; k++)
                {
                    if (sec.Items[k].Type == "textBox") { textRoom += Math.Max(0, NaturalTextCss(sec.Items[k]) - FloorTextCss(sec.Items[k])); }
                }
            }
            // a text box inside columns only shortens the stack when it is the
            // tallest item, so the give is applied and the result measured
            double give = Math.Min(deficit, textRoom);
            ExtraCss = -give;
            deficit = Math.Max(0, StackHeight() - h);
            for (int step = 0; step < 40 && deficit > 0.5; step++)
            {
                bool moved = false;
                if (RowGive < RowFloor()) { RowGive = Math.Min(RowFloor(), RowGive + 1); moved = true; }
                if (BandGive < BandFloor()) { BandGive = Math.Min(BandFloor(), BandGive + 2); moved = true; }
                if (!moved) { break; }
                deficit = Math.Max(0, StackHeight() - h);
            }
        }

        private double RowFloor()
        {
            double floor = 0;
            Rdv3Screen s = f.Screen;
            for (int i = 0; i < s.Sections.Count; i++)
            {
                Rdv3Section sec = s.Sections[i];
                if (sec.Type == "fieldList") { floor = Math.Max(floor, sec.RowHeight - MinRowCss); }
                for (int k = 0; k < sec.Items.Count; k++)
                {
                    if (sec.Items[k].Type == "fieldList") { floor = Math.Max(floor, sec.Items[k].RowHeight - MinRowCss); }
                }
            }
            return Math.Max(0, floor);
        }

        private double BandFloor()
        {
            double floor = 0;
            Rdv3Screen s = f.Screen;
            for (int i = 0; i < s.Sections.Count; i++)
            {
                if (s.Sections[i].Type == "statusBand") { floor = Math.Max(floor, s.Sections[i].Height - MinBandCss); }
            }
            return Math.Max(0, floor);
        }

        private double RowH(Rdv3Section s) { return Math.Max(MinRowCss, s.RowHeight - RowGive); }
        private double TextH(Rdv3Section s) { return Math.Max(FloorTextCss(s), NaturalTextCss(s) + ExtraPerBox()); }
        private bool Stacked(Rdv3Section columns) { return fitW - f.Screen.Padding[1] - f.Screen.Padding[3] < columns.StackBelow; }

        private double StackHeight()
        {
            double h = 0;
            List<KeyValuePair<int, Rdv3Section>> mine = Mine();
            for (int k = 0; k < mine.Count; k++)
            {
                double[] m = MarginOf(mine[k].Value, k, mine.Count);
                h += m[0] + HeightOf(mine[k].Value) + m[2];
            }
            return h;
        }

        // a section's margin: the bars are flush; the cards take the card's
        // padding at the top and bottom and its gap between each other
        private double[] MarginOf(Rdv3Section sec, int k, int n)
        {
            if (sec.Margin != null) { return sec.Margin; }
            Rdv3Screen s = f.Screen;
            if (sec.Type == "titleBar" || sec.Type == "statusBar") { return new double[] { 0, 0, 0, 0 }; }
            List<KeyValuePair<int, Rdv3Section>> mine = Mine();
            bool firstCard = (k == 0) || mine[k - 1].Value.Type == "titleBar";
            bool lastCard = (k == n - 1) || mine[k + 1].Value.Type == "titleBar";
            double top = firstCard ? s.Padding[0] : s.Gap;
            double bottom = lastCard ? s.Padding[2] : 0;
            return new double[] { top, s.Padding[1], bottom, s.Padding[3] };
        }

        // type sizes of the reference, as line heights: 13 px at 1.55 = 20.1
        private const double LabelH = 20.15;        // 13px / 1.55
        private const double RowLabelH = 23.25;     // 15px / 1.55
        private const double BigH = 41.0;           // 34px, line-height 1.2, min-height 41

        public double HeightOf(Rdv3Section s)
        {
            if (s.Type == "titleBar") { return 48; }
            if (s.Type == "keyPanel") { return 12 + LabelH + 2 + BigH + (KeyWraps(s) ? 8 + 44 : 0) + 12 + 2; }
            if (s.Type == "columns")
            {
                double h = 0;
                if (Stacked(s))
                {
                    for (int i = 0; i < s.Items.Count; i++) { h += HeightOf(s.Items[i]) + ((i > 0) ? s.Gap : 0); }
                    return h;
                }
                for (int i = 0; i < s.Items.Count; i++) { h = Math.Max(h, HeightOf(s.Items[i])); }
                return h;
            }
            if (s.Type == "fieldList") { return 14 + LabelH + 6 + s.Rows.Count * RowH(s) + 10 + 2; }
            if (s.Type == "textBox") { return 14 + LabelH + 9 + TextH(s) + 16 + 2; }
            if (s.Type == "statusBand") { return Math.Max(MinBandCss, s.Height - BandGive); }
            if (s.Type == "statusBar") { return s.Height; }
            return 0;
        }

        private int PX(double css) { return (int)Math.Round(css * f.sc); }

        private void Put(string k, double x, double y, double w, double h)
        {
            int x0 = PX(x), y0 = PX(y);
            Rc[k] = new Rectangle(x0, y0, PX(x + w) - x0, PX(y + h) - y0);
        }

        private void PutR(string k, Rectangle r) { Rc[k] = r; }

        public Rectangle At(string k)
        {
            Rectangle r;
            return Rc.TryGetValue(k, out r) ? r : Rectangle.Empty;
        }

        // ---- the layout: every rectangle the painter and the controls use -------
        public void Relayout()
        {
            Rc.Clear();
            double cw = Width / (double)f.sc;
            double y = 0;
            int nText = 0, nCol = 0;
            List<KeyValuePair<int, Rdv3Section>> mine = Mine();
            for (int idx = 0; idx < mine.Count; idx++)
            {
                int i = mine[idx].Key;
                Rdv3Section sec = mine[idx].Value;
                double[] m = MarginOf(sec, idx, mine.Count);
                y += m[0];
                double x = m[3], w = cw - m[1] - m[3], h = HeightOf(sec);
                string key = sec.Type + i.ToString(CultureInfo.InvariantCulture);
                if (sec.Type == "titleBar") { LayTitle(sec, key, x, y, w, h); }
                else if (sec.Type == "keyPanel") { LayKey(sec, key, x, y, w, h); }
                else if (sec.Type == "columns")
                {
                    double total = 0;
                    for (int k = 0; k < sec.Weights.Length; k++) { total += sec.Weights[k]; }
                    bool stacked = Stacked(sec);
                    double cx = x, cy = y;
                    double avail = w - sec.Gap * (sec.Items.Count - 1);
                    for (int k = 0; k < sec.Items.Count; k++)
                    {
                        double iw = stacked ? w : (avail * sec.Weights[k] / total);
                        double ih = stacked ? HeightOf(sec.Items[k]) : h;
                        string ik = key + "." + k.ToString(CultureInfo.InvariantCulture);
                        if (sec.Items[k].Type == "fieldList") { LayFieldList(sec.Items[k], ik, cx, cy, iw, ih, nCol++); }
                        else { LayTextBox(sec.Items[k], ik, cx, cy, iw, ih, nText++); }
                        if (stacked) { cy += ih + sec.Gap; } else { cx += iw + sec.Gap; }
                    }
                    Put(key, x, y, w, h);
                }
                else if (sec.Type == "fieldList") { LayFieldList(sec, key, x, y, w, h, nCol++); }
                else if (sec.Type == "textBox") { LayTextBox(sec, key, x, y, w, h, nText++); }
                else if (sec.Type == "statusBand") { LayBand(sec, key, x, y, w, h); }
                else if (sec.Type == "statusBar") { LayBar(sec, key, x, y, w, h); }
                y += h + m[2];
            }
            Invalidate();
        }

        private double MW(string s, Font font)
        {
            return (Rdv3Skin.Measure(s, font).Width + 1) / (double)f.sc;
        }

        private void LayTitle(Rdv3Section sec, string key, double x, double y, double w, double h)
        {
            Put(key, x, y, w, h);
            double bx = x + 16;
            double bw = MW(sec.Brand, f.fBrand);
            Put(key + ".brand", bx, y + (h - 34.1) / 2, bw, 34.1);
            bx += bw + 10.2;
            for (int i = 0; i < sec.Tags.Count; i++)
            {
                Size tz = Rdv3Skin.TagSize(sec.Tags[i].Text, f.fTag, f.sc);
                double tw = tz.Width / (double)f.sc, th = tz.Height / (double)f.sc;
                Put(key + ".tag" + i.ToString(CultureInfo.InvariantCulture), bx, y + (h - th) / 2, tw, th);
                bx += tw + 10.2;
            }
            // the buttons, right-anchored, 6.8 apart, 34 tall
            double rx = x + w - 16;
            int first = ButtonIndexOf(sec);
            for (int i = sec.Buttons.Count - 1; i >= 0; i--)
            {
                Rdv3Btn b = f.buttons[first + i];
                double bwid = b.WantedWidth() / (double)f.sc;
                rx -= bwid;
                Put(key + ".btn" + i.ToString(CultureInfo.InvariantCulture), rx, y + (h - 34) / 2, bwid, 34);
                b.Bounds = At(key + ".btn" + i.ToString(CultureInfo.InvariantCulture));
                b.BackColor = Rdv3Skin.White;
                rx -= 6.8;
            }
            Put(key + ".rule", x, y + h - 1, w, 1);
        }

        // which of the form's buttons is this section's first
        private int ButtonIndexOf(Rdv3Section sec)
        {
            int n = 0;
            for (int i = 0; i < f.Screen.Sections.Count; i++)
            {
                Rdv3Section s = f.Screen.Sections[i];
                if (s == sec) { return n; }
                if (s.Type == "titleBar" || s.Type == "keyPanel" || s.Type == "statusBar") { n += s.Buttons.Count; }
            }
            return n;
        }

        // the figure and the input group side by side, or the group below
        private double KeyGroupWidth(Rdv3Section sec)
        {
            double gw = sec.InputWidth;
            int first = ButtonIndexOf(sec);
            for (int i = 0; i < sec.Buttons.Count; i++)
            {
                Rdv3Btn b = f.buttons[first + i];
                b.Font = f.fBtn;
                gw += 8 + Math.Max(b.WantedWidth() / (double)f.sc, (sec.Buttons[i].Action == "workState") ? 80 : 0);
            }
            return gw;
        }

        private bool KeyWraps(Rdv3Section sec)
        {
            double inner = fitW - f.Screen.Padding[1] - f.Screen.Padding[3] - 2 - 32;
            double lw = Math.Max(MW(sec.Label, f.fCardTitle), MW(f.View.SearchKey, f.fBig));
            return lw + 20 + KeyGroupWidth(sec) > inner;
        }

        private void LayKey(Rdv3Section sec, string key, double x, double y, double w, double h)
        {
            Put(key, x, y, w, h);
            double ix = x + 1 + 16, iy = y + 1 + 12;
            double lw = Math.Max(MW(sec.Label, f.fCardTitle), MW(f.View.SearchKey, f.fBig));
            Put(key + ".label", ix, iy, lw, LabelH);
            Put(key + ".value", ix, iy + LabelH + 2, lw, BigH);
            // the input group, right-anchored, 8 apart, vertically centred
            // beside the figure, or on its own row below it
            bool wrap = KeyWraps(sec);
            double rx = x + w - 1 - 16;
            int first = ButtonIndexOf(sec);
            double by = wrap ? (iy + LabelH + 2 + BigH + 8 + 1) : (y + (h - 42) / 2);
            for (int i = sec.Buttons.Count - 1; i >= 0; i--)
            {
                Rdv3Btn b = f.buttons[first + i];
                b.Font = f.fBtn;
                double bwid = Math.Max(b.WantedWidth() / (double)f.sc, (sec.Buttons[i].Action == "workState") ? 80 : 0);
                rx -= bwid;
                string bk = key + ".btn" + i.ToString(CultureInfo.InvariantCulture);
                Put(bk, rx, by, bwid, 42);
                b.Bounds = At(bk);
                b.BackColor = Rdv3Skin.White;
                rx -= 8;
            }
            // the input takes its definition width, or what is left of the row
            // once the buttons have theirs -- it never runs out to the left
            double inW = Math.Max(60, Math.Min(sec.InputWidth, rx - (x + 1 + 16)));
            Put(key + ".input", rx - inW, by - 1, inW, 44);
            Rectangle box = At(key + ".input");
            int pad = PX(12);
            int th = Math.Max(f.txtKey.PreferredHeight, PX(18));
            f.txtKey.SetBounds(box.X + pad, box.Y + (box.Height - th) / 2, Math.Max(PX(20), box.Width - 2 * pad), th);
        }

        private void LayFieldList(Rdv3Section sec, string key, double x, double y, double w, double h, int n)
        {
            Put(key, x, y, w, h);
            double ix = x + 1 + 16, iw = w - 2 - 32;
            double ty = y + 1 + 14;
            Put(key + ".title", ix, ty, iw, LabelH);
            double ry = ty + LabelH + 6;
            double rh = RowH(sec);
            for (int i = 0; i < sec.Rows.Count; i++)
            {
                string rk = key + ".row" + i.ToString(CultureInfo.InvariantCulture);
                Put(rk, ix, ry, iw, rh);
                Put(rk + ".label", ix, ry + (rh - RowLabelH) / 2, sec.LabelWidth, RowLabelH);
                Put(rk + ".value", ix + sec.LabelWidth + 12, ry, iw - sec.LabelWidth - 12, rh);
                ry += rh;
            }
        }

        private void LayTextBox(Rdv3Section sec, string key, double x, double y, double w, double h, int n)
        {
            Put(key, x, y, w, h);
            double ix = x + 1 + 16, iw = w - 2 - 32;
            double ty = y + 1 + 14;
            Put(key + ".title", ix, ty, iw, LabelH);
            // the box is its lines, or the whole card when a columns row made
            // the card taller than that: no empty floor under a text box
            double boxH = Math.Max(TextH(sec), h - (14 + LabelH + 9 + 16 + 2));
            Put(key + ".box", ix, ty + LabelH + 9, iw, boxH);
            if (n < f.boxes.Count)
            {
                Rectangle box = At(key + ".box");
                int px = PX(13), py = PX(13);
                f.boxes[n].SetBounds(box.X + px, box.Y + py, Math.Max(PX(20), box.Width - 2 * px), Math.Max(PX(14), box.Height - py - PX(11)));
            }
        }

        private void LayBand(Rdv3Section sec, string key, double x, double y, double w, double h)
        {
            Put(key, x, y, w, h);
            double lw = MW(sec.Label, f.fBandLabel);
            Put(key + ".label", x + 16, y + (h - LabelH) / 2, lw, LabelH);
            // the centred group: icon 26 + 12 + value + 12 + sub. With nothing
            // to judge there is no value (no dash, no placeholder): the band
            // shows its label, and the sub-text when there is one (the
            // several-hits note)
            Rdv3Verdict v = f.Verdict(sec);
            string valueText = (v.Result == null) ? "" : v.Result.Text;
            bool icon = (v.Result != null && v.Result.Icon == "check");
            string sub = f.BandSub(sec);
            double vw = (valueText.Length > 0) ? MW(valueText, f.fBandValue) : 0;
            double sw = (sub.Length > 0) ? MW(sub, f.fBandSub) : 0;
            // the group may not run into the label on the left or out of the
            // band on the right: a long sub-text is shortened first (it is drawn
            // with an ellipsis), then the value itself
            double groupLeft = x + 16 + lw + 16;
            double room = (x + w - 16) - groupLeft;
            double subGap = (sub.Length > 0 && (vw > 0 || icon)) ? 12 : 0;
            double fixedW = (icon ? 26 + 12 : 0) + subGap;
            if (fixedW + vw + sw > room)
            {
                sw = Math.Max(0, room - fixedW - vw);
                if (sub.Length > 0 && sw < 30)
                {
                    // no room worth reading for the sub-text: leave it out
                    sub = ""; sw = 0; subGap = 0; fixedW = (icon ? 26 + 12 : 0);
                }
                if (fixedW + vw > room) { vw = Math.Max(24, room - fixedW); }
            }
            double all = fixedW + vw + sw;
            double left = Math.Max(groupLeft, x + (w - all) / 2);
            if (icon) { Put(key + ".icon", left, y + (h - 26) / 2, 26, 26); left += 26 + 12; }
            if (vw > 0) { Put(key + ".value", left, y + (h - 30.8) / 2, vw, 30.8); left += vw; }
            if (sub.Length > 0) { Put(key + ".sub", left + subGap, y + (h - RowLabelH) / 2 + 2, sw, RowLabelH); }
        }

        private void LayBar(Rdv3Section sec, string key, double x, double y, double w, double h)
        {
            Put(key, x, y, w, h);
            double rx = x + w - 16;
            int first = ButtonIndexOf(sec);
            double bh = Math.Min(34, h - 4);
            for (int i = sec.Buttons.Count - 1; i >= 0; i--)
            {
                Rdv3Btn b = f.buttons[first + i];
                b.Font = f.fTitleBtn;
                double bwid = b.WantedWidth() / (double)f.sc;
                rx -= bwid;
                string bk = key + ".btn" + i.ToString(CultureInfo.InvariantCulture);
                Put(bk, rx, y + (h - bh) / 2, bwid, bh);
                b.Bounds = At(bk);
                b.BackColor = Rdv3Skin.White;
                rx -= 6.8;
            }
            if (sec.Buttons.Count > 0) { rx -= 10.2; }
            double bx = x + 16;
            for (int i = 0; i < sec.Segments.Count; i++)
            {
                Rdv3SegmentDef d = sec.Segments[i];
                string text = d.Prefix + f.Eval(d.Value).Text;
                if (d.Dot)
                {
                    Put(key + ".dot", bx, y + (h - 8) / 2, 8, 8);
                    bx += 8 + 6;
                }
                double tw = MW(text, d.Bold ? f.fBarBold : f.fBar);
                double room = rx - bx;
                if (room <= 20) { break; }
                double use = Math.Min(tw, room);
                Put(key + ".seg" + i.ToString(CultureInfo.InvariantCulture), bx, y + (h - LabelH) / 2, use, LabelH);
                bx += use;
                if (i < sec.Segments.Count - 1)
                {
                    bx += 10.2;
                    if (bx + 1 >= rx) { break; }
                    Put(key + ".sep" + i.ToString(CultureInfo.InvariantCulture), bx, y + (h - 12) / 2, 1, 12);
                    bx += 1 + 10.2;
                }
            }
        }

        // ---- painting -----------------------------------------------------------
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
        g.PixelOffsetMode = PixelOffsetMode.Half;
            Clipped.Clear();
            using (SolidBrush b = new SolidBrush(Rdv3Skin.N200)) { g.FillRectangle(b, ClientRectangle); }
            int nCol = 0, nText = 0;
            List<KeyValuePair<int, Rdv3Section>> mine = Mine();
            for (int idx = 0; idx < mine.Count; idx++)
            {
                int i = mine[idx].Key;
                Rdv3Section sec = mine[idx].Value;
                string key = sec.Type + i.ToString(CultureInfo.InvariantCulture);
                if (sec.Type == "titleBar") { PaintTitle(g, sec, key, i == 0); }
                else if (sec.Type == "keyPanel") { PaintKey(g, sec, key); }
                else if (sec.Type == "columns")
                {
                    for (int k = 0; k < sec.Items.Count; k++)
                    {
                        string ik = key + "." + k.ToString(CultureInfo.InvariantCulture);
                        if (sec.Items[k].Type == "fieldList") { PaintFieldList(g, sec.Items[k], ik); nCol++; }
                        else { PaintTextBox(g, sec.Items[k], ik, nText++); }
                    }
                }
                else if (sec.Type == "fieldList") { PaintFieldList(g, sec, key); nCol++; }
                else if (sec.Type == "textBox") { PaintTextBox(g, sec, key, nText++); }
                else if (sec.Type == "statusBand") { PaintBand(g, sec, key); }
                else if (sec.Type == "statusBar") { PaintBar(g, sec, key, true); }
            }
        }

        private void Note(string k, string s, Font font, Rectangle r)
        {
            if (s == null || s.Length == 0) { return; }
            if (Rdv3Skin.Measure(s, font).Width > r.Width && !Clipped.Contains(k)) { Clipped.Add(k); }
        }

        private void T(Graphics g, string k, string s, Font font, Color c, int align)
        {
            Rectangle r = At(k);
            Note(k, s, font, r);
            Rdv3Skin.DrawIn(g, s, font, c, r, align);
        }

        private void Card(Graphics g, Rectangle r)
        {
            Rdv3Skin.FillRound(g, Rdv3Skin.White, r, PX(10));
            Rdv3Skin.FrameRound(g, Rdv3Skin.CardEdge, r, PX(10));
        }

        private void PaintTitle(Graphics g, Rdv3Section sec, string key, bool first)
        {
            Rectangle r = At(key);
            using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, r); }
            T(g, key + ".brand", sec.Brand, f.fBrand, Rdv3Skin.Ink, 0);
            for (int i = 0; i < sec.Tags.Count; i++)
            {
                Rdv3Skin.Tag(g, sec.Tags[i].Text, f.fTag, Rdv3Skin.TagKind(sec.Tags[i].Look),
                    At(key + ".tag" + i.ToString(CultureInfo.InvariantCulture)), f.sc);
            }
            Rectangle rule = At(key + ".rule");
            Rdv3Skin.Line(g, Rdv3Skin.Divider, rule.X, rule.Y, rule.Width, 1);
        }

        private void PaintKey(Graphics g, Rdv3Section sec, string key)
        {
            Card(g, At(key));
            T(g, key + ".label", sec.Label, f.fCardTitle, Rdv3Skin.N600, 0);
            Rdv3Value v = f.Eval(sec.Value);
            T(g, key + ".value", v.Text, f.fBig, Rdv3Skin.Ink, 0);
            Rectangle box = At(key + ".input");
            Rdv3Skin.FillRound(g, Rdv3Skin.N200, box, PX(10));
            if (f.txtKey.Focused) { Rdv3Skin.FrameRound(g, Rdv3Skin.Accent, box, PX(10)); }
            else if (f.inputHot) { Rdv3Skin.FrameRound(g, Rdv3Skin.Mix(Rdv3Skin.N200, 0.2), box, PX(10)); }
        }

        private void PaintFieldList(Graphics g, Rdv3Section sec, string key)
        {
            Card(g, At(key));
            T(g, key + ".title", sec.Title, f.fCardTitle, Rdv3Skin.N600, 0);
            for (int i = 0; i < sec.Rows.Count; i++)
            {
                string rk = key + ".row" + i.ToString(CultureInfo.InvariantCulture);
                Rectangle row = At(rk);
                T(g, rk + ".label", sec.Rows[i].Label, f.fRowLabel, Rdv3Skin.N700, 0);
                Rdv3Value v = f.Eval(sec.Rows[i].Value);
                Font vf = (v.Tone == Rdv3Value.Normal) ? f.fRowValue : f.fRowLabel;
                Color vc = (v.Tone == Rdv3Value.Error) ? Rdv3Skin.Danger : (v.Tone == Rdv3Value.Muted) ? Rdv3Skin.N400 : Rdv3Skin.Ink;
                T(g, rk + ".value", v.Text, vf, vc, 2);
                if (i < sec.Rows.Count - 1) { Rdv3Skin.Line(g, Rdv3Skin.RowLine, row.X, row.Bottom - 1, row.Width, 1); }
            }
        }

        private void PaintTextBox(Graphics g, Rdv3Section sec, string key, int n)
        {
            Card(g, At(key));
            T(g, key + ".title", sec.Title, f.fCardTitle, Rdv3Skin.N600, 0);
            Rectangle box = At(key + ".box");
            Rdv3Skin.FillRound(g, Rdv3Skin.N100, box, PX(8));
            Rdv3Skin.FrameRound(g, Rdv3Skin.CardEdge, box, PX(8));
            if (n < f.boxes.Count)
            {
                Rdv3Value v = f.Eval(sec.Value);
                // nothing is shown while there is no record: no placeholder sentence
                string text = f.View.HasRecord ? v.Text : "";
                Color ink = (v.Tone == Rdv3Value.Error) ? Rdv3Skin.Danger
                    : (v.Tone == Rdv3Value.Muted) ? Rdv3Skin.N400 : Rdv3Skin.Ink;
                TextBox t = f.boxes[n];
                // a joiner of "\n" stacks several columns: the EDIT control
                // only breaks a line at CR LF
                string boxText = text.Replace("\r\n", "\n").Replace("\n", "\r\n");
                if (t.Text != boxText || t.ForeColor != ink) { f.SetBox(t, boxText, ink); }
            }
        }

        private void PaintBand(Graphics g, Rdv3Section sec, string key)
        {
            Rectangle r = At(key);
            Rdv3Verdict v = f.Verdict(sec);
            string look = (v.Result == null) ? "" : v.Result.Look;
            Color back, edge, ink;
            if (look == "ok") { back = Rdv3Skin.Accent100; edge = Rdv3Skin.Accent300; ink = Rdv3Skin.Accent800; }
            else if (look == "ng") { back = Rdv3Skin.Mix(Rdv3Skin.White, 0.08); edge = Rdv3Skin.Mix(Rdv3Skin.White, 0.12); ink = Rdv3Skin.Ink; }
            else if (look == "error") { back = Rdv3Skin.DangerBack; edge = Rdv3Skin.MixWith(Rdv3Skin.White, Rdv3Skin.Danger, 0.35); ink = Rdv3Skin.Danger; }
            else if (look == "undefined") { back = Rdv3Skin.N100; edge = Rdv3Skin.Mix(Rdv3Skin.White, 0.12); ink = Rdv3Skin.N700; }
            else { back = Rdv3Skin.N100; edge = Rdv3Skin.Mix(Rdv3Skin.White, 0.12); ink = Rdv3Skin.N500; }
            Rdv3Skin.FillRound(g, back, r, PX(10));
            Rdv3Skin.FrameRound(g, edge, r, PX(10));
            T(g, key + ".label", sec.Label, f.fBandLabel, Rdv3Skin.MixWith(back, ink, 0.55), 0);
            if (Rc.ContainsKey(key + ".icon")) { Rdv3Skin.IconCheckCircle(g, At(key + ".icon"), ink, f.sc); }
            if (v.Result != null && Rc.ContainsKey(key + ".value")) { T(g, key + ".value", v.Result.Text, f.fBandValue, ink, 0); }
            string sub = f.BandSub(sec);
            if (sub.Length > 0 && Rc.ContainsKey(key + ".sub")) { T(g, key + ".sub", sub, f.fBandSub, Rdv3Skin.MixWith(back, ink, 0.75), 0); }
        }

        private void PaintBar(Graphics g, Rdv3Section sec, string key, bool last)
        {
            Rectangle r = At(key);
            using (SolidBrush b = new SolidBrush(Rdv3Skin.White)) { g.FillRectangle(b, r); }
            Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, r.Width, 1);
            if (Rc.ContainsKey(key + ".dot"))
            {
                Rectangle d = At(key + ".dot");
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent)) { g.FillEllipse(b, d); }
            }
            for (int i = 0; i < sec.Segments.Count; i++)
            {
                string sk = key + ".seg" + i.ToString(CultureInfo.InvariantCulture);
                if (!Rc.ContainsKey(sk)) { break; }
                Rdv3SegmentDef d = sec.Segments[i];
                string text = d.Prefix + f.Eval(d.Value).Text;
                Rdv3Skin.DrawIn(g, text, d.Bold ? f.fBarBold : f.fBar, Rdv3Skin.N700, At(sk), 0);
                string pk = key + ".sep" + i.ToString(CultureInfo.InvariantCulture);
                if (Rc.ContainsKey(pk))
                {
                    Rectangle p = At(pk);
                    Rdv3Skin.Line(g, Rdv3Skin.Mix(Rdv3Skin.White, 0.15), p.X, p.Y, 1, p.Height);
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
        }
    }

    // ---- evaluation helpers the card uses ------------------------------------------
    internal Rdv3Value Eval(Rdv3Bind b)
    {
        return Rdv3Eval.Evaluate(b, View, fields, Screen.Work);
    }

    internal Rdv3Verdict Verdict(Rdv3Section band)
    {
        return Rdv3Eval.Judge(Screen.JudgmentOf(band.Judgment), View, fields);
    }

    // the band's sub-line: its parts joined, or the multi-hit note while
    // candidates exist but none is chosen
    internal string BandSub(Rdv3Section band)
    {
        if (!View.HasRecord)
        {
            if (View.CandidateCount > 1)
            {
                return Rdv3Text.SubMulti.Replace("{n}", View.CandidateCount.ToString("N0", CultureInfo.InvariantCulture));
            }
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < band.Sub.Count; i++)
        {
            string t = Eval(band.Sub[i]).Text;
            if (t.Length == 0) { continue; }
            if (sb.Length > 0) { sb.Append(band.Joiner); }
            sb.Append(t);
        }
        return sb.ToString();
    }

    // ---- the public surface (all marshalled onto the UI thread) ------------------
    private readonly List<Action> pending = new List<Action>();

    private void Ui(Action a)
    {
        if (IsDisposed) { return; }
        if (!IsHandleCreated)
        {
            lock (pending) { pending.Add(a); }
            return;
        }
        try { if (InvokeRequired) { Invoke(a); } else { a(); } }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { }
    }

    private void FlushPending()
    {
        Action[] q;
        lock (pending)
        {
            if (pending.Count == 0) { return; }
            q = pending.ToArray();
            pending.Clear();
        }
        for (int i = 0; i < q.Length; i++)
        {
            try { q[i](); }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }
    }

    public void RunOnUi(Action a) { Ui(a); }

    // queued, not waited for: a worker must never sit inside a modal loop
    public void PostOnUi(Action a)
    {
        if (IsDisposed) { return; }
        if (!IsHandleCreated) { Ui(a); return; }
        try { BeginInvoke(a); }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { }
    }

    // A state change can change the stack's height (the figure's key decides
    // whether the input group wraps under it), so the whole fit is redone,
    // not just the placement: the card's height, the scroll decision and the
    // rectangles all follow from it.
    private void Refresh1() { Layout1(); }

    // the column names of the data, once the merge has read them
    public void SetFields(Rdv3Fields f)
    {
        Ui(delegate { fields = (f == null) ? Rdv3Fields.Empty : f; Refresh1(); });
    }

    public Rdv3Fields Fields { get { return fields; } }

    public void SetState(string text)
    {
        Ui(delegate { View.AppState = text; Refresh1(); });
    }

    public void SetWatch(string label, string detail)
    {
        Ui(delegate
        {
            View.WatchLabel = ((label == null || label.Length == 0) ? Rdv3Text.LabelNotepad : label)
                + ((detail == null || detail.Length == 0) ? "" : (" " + detail));
            View.WatchDetail = (detail == null) ? "" : detail;
            Refresh1();
        });
    }

    public void SetLedger(string file, string rows, string saved)
    {
        Ui(delegate
        {
            View.LedgerFile = file;
            View.LedgerRows = rows;
            View.LedgerSaved = saved;
            Refresh1();
        });
    }

    public void SetTimes(double mergeMs, double searchMs)
    {
        Ui(delegate
        {
            if (mergeMs >= 0) { View.MergeMs = Rdv3Clock.Fmt(mergeMs) + Rdv3Text.MsUnit; }
            if (searchMs >= 0) { View.SearchMs = Rdv3Clock.Fmt(searchMs) + Rdv3Text.MsUnit; }
            Refresh1();
        });
    }

    public void SetIdentity(string pid, string logName)
    {
        Ui(delegate { View.Pid = pid; View.LogName = logName; Refresh1(); });
    }

    public void EnableOps(bool on)
    {
        Ui(delegate
        {
            txtKey.Enabled = on;
            for (int i = 0; i < buttons.Count; i++)
            {
                Rdv3ButtonDef d = buttonDefs[buttons[i]];
                if (d.Action == "settings") { continue; }
                buttons[i].Enabled = on;
            }
        });
    }

    public void EnableWorkState(bool on)
    {
        Ui(delegate { if (btnWork != null) { btnWork.Enabled = on; } });
    }

    // ---- results ---------------------------------------------------------------------
    // a search came back: the list (possibly empty) and no selection yet
    public void ShowCandidates(string key, List<Rdv3CandRow> rows, int totalHits)
    {
        Ui(delegate
        {
            View.SearchKey = key;
            cands = (rows == null) ? new List<Rdv3CandRow>() : rows;
            candTotal = totalHits;
            View.CandidateCount = totalHits;
            View.SelectedIndex = -1;
            View.Record = null;
            View.StoredState = "";
            View.Saving = false;
            UpdateWorkButton();
            Refresh1();
        });
    }

    public void SelectCandidate(int index)
    {
        Ui(delegate
        {
            if (index < 0 || index >= cands.Count) { return; }
            View.SelectedIndex = index;
            View.Record = Rdv3Ledger.SplitLine(cands[index].Line);
            View.StoredState = cands[index].Stored;
            View.Saving = false;
            UpdateWorkButton();
            Refresh1();
        });
    }

    public List<Rdv3CandRow> Candidates { get { return cands; } }
    public int CandidateTotal { get { return candTotal; } }

    // the stored state of the record on show (and of its list row)
    public void SetStoredState(int index, string stored, bool saving)
    {
        Ui(delegate
        {
            if (index >= 0 && index < cands.Count) { cands[index].Stored = stored; }
            if (index == View.SelectedIndex && View.Record != null) { View.StoredState = stored; }
            View.Saving = saving && (index == View.SelectedIndex);
            UpdateWorkButton();
            Refresh1();
        });
    }

    private void UpdateWorkButton()
    {
        if (btnWork == null) { return; }
        Rdv3WorkState w = Screen.Work;
        string text;
        if (!View.HasRecord)
        {
            Rdv3StateDef init = (w == null) ? null : w.InitialState;
            text = (init == null) ? "" : init.Text;
        }
        else { text = Rdv3Eval.WorkStateValue(View, w, false).Text.Replace(Rdv3Text.SavingSuffix, ""); }
        string tpl = (w == null) ? "{state}" : w.ButtonText;
        btnWork.Text = tpl.Replace("{state}", text);
    }

    public void ClearResult()
    {
        Ui(delegate
        {
            txtKey.Text = "";
            View.SearchKey = "";
            cands = new List<Rdv3CandRow>();
            candTotal = 0;
            View.CandidateCount = 0;
            View.SelectedIndex = -1;
            View.Record = null;
            View.StoredState = "";
            View.Saving = false;
            UpdateWorkButton();
            Refresh1();
        });
    }

    public string KeyText { get { return txtKey.Text.Trim(); } }

    public void SetKeyText(string s)
    {
        Ui(delegate { txtKey.Text = (s == null) ? "" : s; card.Invalidate(); });
    }

    // ---- the acceptance dump ------------------------------------------------------
    public string GeometryDump()
    {
        Dictionary<string, Rectangle> all = new Dictionary<string, Rectangle>(card.Rc);
        int barY = card.Height;
        foreach (KeyValuePair<string, Rectangle> kv in bar.Rc)
        {
            all[kv.Key] = new Rectangle(kv.Value.X, kv.Value.Y + barY, kv.Value.Width, kv.Value.Height);
        }
        List<string> keys = new List<string>(all.Keys);
        keys.Sort(StringComparer.Ordinal);
        StringBuilder sb = new StringBuilder();
        sb.Append("{\"scale\":").Append(sc.ToString("0.###", CultureInfo.InvariantCulture));
        sb.Append(",\"card\":[").Append(Css(Math.Max(card.Width, bar.Width))).Append(",").Append(Css(card.Height + bar.Height)).Append("],\"el\":{");
        for (int i = 0; i < keys.Count; i++)
        {
            Rectangle r = all[keys[i]];
            if (i > 0) { sb.Append(","); }
            sb.Append("\"").Append(keys[i]).Append("\":[")
              .Append(Css(r.X)).Append(",").Append(Css(r.Y)).Append(",")
              .Append(Css(r.Width)).Append(",").Append(Css(r.Height)).Append("]");
        }
        sb.Append("},\"clipped\":[");
        List<string> clipped = new List<string>(card.Clipped);
        clipped.AddRange(bar.Clipped);
        for (int i = 0; i < clipped.Count; i++)
        {
            if (i > 0) { sb.Append(","); }
            sb.Append("\"").Append(clipped[i]).Append("\"");
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private string Css(int devicePx)
    {
        return (devicePx / (double)sc).ToString("0.#", CultureInfo.InvariantCulture);
    }

    // the client area on the screen, for the modals and the toast
    public Rectangle CardBounds
    {
        get { return RectangleToScreen(new Rectangle(0, 0, ClientSize.Width, ClientSize.Height)); }
    }

    // ---- overlays ------------------------------------------------------------------
    private Rdv3Toast toast;

    // a completion (TagDone) or an error (TagError), bottom right, for toast.durationMs
    public void Notice(string text) { Toast(text, true); }
    public void Error(string text) { Toast(text, false); }

    private void Toast(string text, bool done)
    {
        Ui(delegate
        {
            if (text == null || text.Length == 0) { return; }
            if (toast == null) { toast = new Rdv3Toast(this); }
            toast.Show(text, done, Screen.ToastMs);
        });
    }

    // UI thread only: the yes / no modal over the dimmed card
    public bool Ask(string title, string body)
    {
        return Rdv3ConfirmForm.Ask(this, title, body);
    }

    // UI thread only: the app cannot go on. One modal with the reason, and
    // the window closes when it is acknowledged.
    public void Fatal(string title, string body)
    {
        Rdv3ConfirmForm.Tell(this, title, body);
        Close();
    }

    // UI thread only: the candidate list modal; the index picked, or -1
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
            return "scale=" + sc.ToString("0.00", CultureInfo.InvariantCulture)
                + " client=" + ClientSize.Width + "x" + ClientSize.Height
                + " card=" + card.Width + "x" + card.Height
                + " work=" + wa.Width + "x" + wa.Height
                + " font=" + Rdv3Skin.Family;
        }
    }
}
