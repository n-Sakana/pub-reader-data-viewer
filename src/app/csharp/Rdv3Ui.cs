// ============================================================================
// Rdv3Ui.cs -- the WinForms screen of the practical build, rebuilt to the
// reference artifact "Reader Data Viewer_ver3.html" (docs\ui-spec.md carries
// the measured tokens and geometry; work\ui-ref\ref-*.png are the reference
// renders this was matched against).
//
// Native WinForms only: no WebView, no HTML. One card, six rows:
//
//   1 title bar    brand + method tag + ledger tag + "notepad re-detect"
//   2 summary      key 1, representative status, ledger size, the input and
//                  its three buttons, and the session identity block
//   3 candidates   the one-to-many candidate table (key 1 legitimately has
//                  many rows); clicking a row selects it
//   4 record       what the list does NOT show: table A's fields plus the two
//                  long-text columns
//   5 error        only while there is an error to show
//   6 status bar   state, notepad, ledger, merge time, search time, log, pid,
//                  clock -- the same figures as before, in the reference's bar
//
// Everything is drawn with GDI+ against tokens scaled from 96 dpi, so the
// screen keeps its proportions at 125/150/175% Windows scaling. The only
// live controls are the ones that must take keyboard input or scroll: the
// search box, four buttons, the two long-text boxes and the table's scrollbar.
//
// Every public method marshals itself onto the UI thread: the worker calls
// them directly and never touches a control.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;

// ---------------------------------------------------------------------------
// the design tokens, measured from the reference render
// ---------------------------------------------------------------------------
internal static class Rdv3Skin
{
    public static readonly Color Bg = Color.FromArgb(0xF2, 0xF2, 0xF3);
    public static readonly Color Page = Color.FromArgb(0xE7, 0xE7, 0xEA);
    public static readonly Color Surface = Color.FromArgb(0xE9, 0xE9, 0xEA);
    public static readonly Color Ink = Color.FromArgb(0x1D, 0x1F, 0x20);
    public static readonly Color Divider = Color.FromArgb(41, 0x1D, 0x1F, 0x20);      // 16%
    public static readonly Color RowLine = Color.FromArgb(12, 0x1D, 0x1F, 0x20);      // ~5%
    public static readonly Color Hover = Color.FromArgb(10, 0x1D, 0x1F, 0x20);        // 4%
    public static readonly Color Press = Color.FromArgb(36, 0x1D, 0x1F, 0x20);        // 14%
    public static readonly Color Corner = Color.FromArgb(168, 0x1D, 0x1F, 0x20);      // 66% (reference 55%)
    public static readonly Color ThInk = Color.FromArgb(153, 0x1D, 0x1F, 0x20);       // 60%

    public static readonly Color Accent = Color.FromArgb(0x59, 0x80, 0xA6);
    public static readonly Color Accent100 = Color.FromArgb(0xEE, 0xF6, 0xFF);
    public static readonly Color Accent300 = Color.FromArgb(0xB5, 0xD9, 0xFD);
    public static readonly Color Accent600 = Color.FromArgb(0x59, 0x7E, 0xA3);
    public static readonly Color Accent700 = Color.FromArgb(0x41, 0x61, 0x80);
    public static readonly Color Accent800 = Color.FromArgb(0x2C, 0x45, 0x5D);
    public static readonly Color Accent900 = Color.FromArgb(0x1D, 0x2D, 0x3D);

    public static readonly Color N100 = Color.FromArgb(0xF5, 0xF5, 0xF8);
    public static readonly Color N400 = Color.FromArgb(0xB7, 0xB7, 0xBA);
    public static readonly Color N600 = Color.FromArgb(0x7A, 0x7A, 0x7D);
    // the reference's next step down the same neutral ramp. Sub-lines, key
    // labels and box captions use it instead of N600: at 11-12 px the lighter
    // tone was legible on a browser's rendering and not on this one.
    public static readonly Color N700 = Color.FromArgb(0x5D, 0x5D, 0x60);
    // input validation, said quietly: a brick red that sits in the same
    // muted family as the rest of the palette rather than a signal red
    public static readonly Color Danger = Color.FromArgb(0xB0, 0x4A, 0x3E);
    public static readonly Color N800 = Color.FromArgb(0x42, 0x42, 0x44);

    public static float Scale = 1f;
    public static int P(double cssPx) { return (int)Math.Round(cssPx * Scale); }

    private static string family = "Yu Gothic UI";
    private static string headFamily = "Yu Gothic UI";
    private static string figFamily = "Yu Gothic UI";
    private static string semiFamily = "Yu Gothic UI";
    private static string lightFamily = "Yu Gothic UI";

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

    private static bool semiIsBold;

    public static void PickFamily()
    {
        // Yu Gothic UI is what Windows dresses its own dialogs in, but at the
        // 10-13 px this screen is set in it renders THIN: the operator's first
        // note on the rebuilt screen was that it was hard to read. Meiryo UI is
        // the heavier, denser Japanese UI face on the same machine and is a
        // straight improvement at these sizes (compared side by side at 12 and
        // 13 px against Yu Gothic UI, Noto Sans JP and BIZ UDP Gothic).
        family = FirstInstalled(new string[] { "Meiryo UI", "Yu Gothic UI", "Meiryo", "MS UI Gothic", "Segoe UI" },
            FontFamily.GenericSansSerif.Name);
        lightFamily = family;
        // The reference sets its headings and figures in Barlow Condensed, which
        // Windows does not ship. Bahnschrift's cuts are the closest thing that
        // IS on the machine -- but no single cut matches both halves of Barlow
        // Condensed: measured against the reference at its own sizes,
        //   "Reader Data Viewer" 19px  target 132.1  SemiConden 146.7  Condensed 123.9
        //   "00016168"           34px  target 131.2  SemiConden 116.9  SemiBold  132.9
        // so the display face is picked per role, by measured width, from the
        // same Bahnschrift SemiBold family: Condensed for display TEXT, the
        // regular width for the big FIGURES. Japanese glyphs fall back to the
        // body face per glyph, exactly as the browser does.
        headFamily = FirstInstalled(new string[] { "Bahnschrift SemiBold Condensed", "Bahnschrift Condensed",
            "Bahnschrift SemiBold SemiConden", "Segoe UI Semibold" }, family);
        figFamily = FirstInstalled(new string[] { "Bahnschrift SemiBold", "Bahnschrift",
            "Segoe UI Semibold" }, family);
        // weight 600: a dedicated semibold cut if the body face has one (Yu
        // Gothic UI does), otherwise the family's own bold (Meiryo UI)
        string dedicated = FirstInstalled(new string[] { family + " Semibold" }, "");
        if (dedicated.Length > 0) { semiFamily = dedicated; semiIsBold = false; }
        else { semiFamily = family; semiIsBold = true; }
    }

    public static Font F(double cssPx, FontStyle st)
    {
        return new Font(family, (float)(cssPx * Scale), st, GraphicsUnit.Pixel);
    }

    // headings and the three big figures: condensed where the machine has it
    public static Font H(double cssPx, FontStyle st)
    {
        return new Font(headFamily, (float)(cssPx * Scale), st, GraphicsUnit.Pixel);
    }

    // the three big figures: the digit-matched display cut
    public static Font D(double cssPx)
    {
        return new Font(figFamily, (float)(cssPx * Scale), FontStyle.Regular, GraphicsUnit.Pixel);
    }

    // the muted 10-12 px labels
    public static Font L(double cssPx)
    {
        return new Font(lightFamily, (float)(cssPx * Scale), FontStyle.Regular, GraphicsUnit.Pixel);
    }

    // weight 600
    public static Font S(double cssPx)
    {
        return new Font(semiFamily, (float)(cssPx * Scale),
            semiIsBold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
    }

    // color-mix(in srgb, var(--color-text) N%, transparent) laid over the page
    public static Color Mix(double f) { return Mix(Bg, f); }

    public static Color Mix(Color over, double f)
    {
        return Color.FromArgb(
            (int)Math.Round(over.R * (1 - f) + Ink.R * f),
            (int)Math.Round(over.G * (1 - f) + Ink.G * f),
            (int)Math.Round(over.B * (1 - f) + Ink.B * f));
    }

    // a rule stays a rule: one device pixel, whatever the scale
    public static int Hair() { return 1; }

    // TextRenderer.MeasureText adds a constant of roughly 0.57 em to whatever it
    // is given -- even with NoPadding, and even for one character. Measured on
    // this machine: a 13 px kanji comes back as 21, eleven of them as 151 (the
    // advance is a correct 13; the box is simply padded). Sizing boxes from that
    // number made every button, tag and figure block half an em too wide, and
    // the summary row then hit its own "give way" rule and clipped itself.
    //
    // GDI+ with the typographic format returns the true run width (13 and 143),
    // and its per-glyph advances are identical to the ones GDI then draws with,
    // so measuring here and drawing with TextRenderer stay in step.
    // A layout plus a repaint asks for the same few dozen strings every time,
    // and a tracked run measures one prefix per glyph on top of that. Measuring
    // is a GDI+ call each time, so the answers are remembered: the save job's
    // worker thread was measurably paying for the UI thread's arithmetic.
    private static readonly Dictionary<string, Size> memo = new Dictionary<string, Size>();

    public static void ResetMetrics()
    {
        lock (memo) { memo.Clear(); }
        lock (memoF) { memoF.Clear(); }
    }

    public static Size Measure(string s, Font f)
    {
        if (s == null || s.Length == 0) { return new Size(0, HeightTracked("A", f)); }
        string k = f.Name + "" + f.Size.ToString(CultureInfo.InvariantCulture)
            + "" + ((int)f.Style).ToString(CultureInfo.InvariantCulture) + "" + s;
        Size hit;
        lock (memo)
        {
            if (memo.TryGetValue(k, out hit)) { return hit; }
        }
        SizeF z = Scratch().MeasureString(s, f, PointF.Empty, Typo());
        hit = new Size((int)Math.Ceiling(z.Width), (int)Math.Ceiling(z.Height));
        lock (memo)
        {
            if (memo.Count > 8000) { memo.Clear(); }
            memo[k] = hit;
        }
        return hit;
    }

    // Drawn through GDI+ as well, so the width a box was sized from is exactly
    // the width the glyphs occupy. Mixing the two engines cost a couple of
    // pixels per run, which was enough to trip the ellipsis on titles that
    // actually fitted ("候補一覧" came out as "候...").
    public static void Draw(Graphics g, string s, Font f, Color c, int x, int y)
    {
        if (s == null || s.Length == 0) { return; }
        TextRenderingHint keep = g.TextRenderingHint;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        using (SolidBrush b = new SolidBrush(c)) { g.DrawString(s, f, b, x, y, Typo()); }
        g.TextRenderingHint = keep;
    }

    // inside a box, and cut with an ellipsis rather than spilling into the
    // neighbour: overlapping text is then structurally impossible
    public static void DrawIn(Graphics g, string s, Font f, Color c, Rectangle r, bool right)
    {
        if (s == null || s.Length == 0 || r.Width <= 0 || r.Height <= 0) { return; }
        string t = Fit(s, f, r.Width);
        if (t.Length == 0) { return; }
        Size z = Measure(t, f);
        int x = right ? (r.Right - z.Width) : r.X;
        Draw(g, t, f, c, x, r.Y + (r.Height - z.Height) / 2);
    }

    // the ellipsis, done here so it uses the same metric as everything else
    private static string Fit(string s, Font f, int w)
    {
        if (Measure(s, f).Width <= w) { return s; }
        int lo = 0, hi = s.Length;
        while (lo < hi)
        {
            int mid = (lo + hi + 1) / 2;
            if (Measure(s.Substring(0, mid) + "…", f).Width <= w) { lo = mid; } else { hi = mid - 1; }
        }
        return (lo <= 0) ? "" : (s.Substring(0, lo) + "…");
    }

    // GDI rounds every MeasureText call, so measuring glyph by glyph opened a
    // five pixel hole after the first character. Tracked runs are therefore
    // measured AND drawn through GDI+ with the typographic string format: the
    // offsets are sub-pixel and the two agree by construction.
    private static Graphics scratch;
    private static StringFormat typo;

    private static Graphics Scratch()
    {
        if (scratch == null)
        {
            scratch = Graphics.FromImage(new Bitmap(1, 1));
            scratch.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        }
        return scratch;
    }

    // GenericTypographic drops trailing spaces, which slid every glyph after an
    // internal space one space to the left ("ReaderD ataV iewer")
    private static StringFormat Typo()
    {
        if (typo == null)
        {
            typo = new StringFormat(StringFormat.GenericTypographic);
            typo.FormatFlags |= StringFormatFlags.MeasureTrailingSpaces;
        }
        return typo;
    }

    // Sub-pixel prefix widths. The tracked runs place one glyph at a time, so
    // taking each prefix from the ROUNDED cache accumulated up to a couple of
    // pixels over a short run -- enough to push the last digit of the key past
    // its box and have it dropped ("0001616" for 00016168).
    private static readonly Dictionary<string, float> memoF = new Dictionary<string, float>();

    private static float RunW(string s, Font f)
    {
        if (s == null || s.Length == 0) { return 0f; }
        string k = f.Name + "" + f.Size.ToString(CultureInfo.InvariantCulture)
            + "" + ((int)f.Style).ToString(CultureInfo.InvariantCulture) + "" + s;
        float hit;
        lock (memoF) { if (memoF.TryGetValue(k, out hit)) { return hit; } }
        hit = Scratch().MeasureString(s, f, PointF.Empty, Typo()).Width;
        lock (memoF)
        {
            if (memoF.Count > 8000) { memoF.Clear(); }
            memoF[k] = hit;
        }
        return hit;
    }

    public static int MeasureTracked(string s, Font f, double track)
    {
        if (s == null || s.Length == 0) { return 0; }
        return (int)Math.Ceiling(RunW(s, f) + track * Scale * (s.Length - 1));
    }

    private static float TrackPx(double track) { return (float)(track * Scale); }

    public static int HeightOf(Font f) { return HeightTracked("A", f); }

    public static int HeightTracked(string s, Font f)
    {
        if (s == null || s.Length == 0) { s = "A"; }
        return (int)Math.Ceiling(Scratch().MeasureString(s, f, PointF.Empty, Typo()).Height);
    }

    public static void DrawTracked(Graphics g, string s, Font f, Color c, int x, int y,
        double track, int rightLimit)
    {
        if (s == null || s.Length == 0) { return; }
        TextRenderingHint keep = g.TextRenderingHint;
        g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
        using (SolidBrush br = new SolidBrush(c))
        {
            float t = TrackPx(track);
            float pen = x;
            for (int i = 0; i < s.Length; i++)
            {
                float w = RunW(s.Substring(0, i + 1), f) - RunW(s.Substring(0, i), f);
                if (pen + w > rightLimit + 1f) { break; }
                g.DrawString(s.Substring(i, 1), f, br, pen, y, Typo());
                pen += w + t;
            }
        }
        g.TextRenderingHint = keep;
    }

    public static void DrawRight(Graphics g, string s, Font f, Color c, int right, int y)
    {
        Draw(g, s, f, c, right - Measure(s, f).Width, y);
    }

    public static void Line(Graphics g, Color c, int x, int y, int w, int h)
    {
        using (SolidBrush b = new SolidBrush(c)) { g.FillRectangle(b, x, y, w, h); }
    }

    public static void Frame(Graphics g, Color c, Rectangle r)
    {
        int t = Hair();
        Line(g, c, r.X, r.Y, r.Width, t);
        Line(g, c, r.X, r.Bottom - t, r.Width, t);
        Line(g, c, r.X, r.Y, t, r.Height);
        Line(g, c, r.Right - t, r.Y, t, r.Height);
    }

    // a panel frame: the reference's hairline box. Its registration marks are
    // deliberately NOT drawn (the operator asked for the corners to go).
    public static void Blueprint(Graphics g, Rectangle r)
    {
        Frame(g, Divider, r);
    }

    public const int TagAccent = 0;
    public const int TagOutline = 1;
    public const int TagNeutral = 2;

    public static Size TagSize(string text, Font f)
    {
        Size s = Measure(text, f);
        return new Size(s.Width + P(20), Math.Max(s.Height + P(6), P(23)));
    }

    public static void Tag(Graphics g, string text, Font f, int kind, Rectangle r, int alpha)
    {
        Color back, ink, edge;
        if (kind == TagAccent) { back = Accent100; ink = Accent800; edge = Color.Empty; }
        else if (kind == TagOutline) { back = Color.Empty; ink = Accent; edge = Accent; }
        else { back = N100; ink = N800; edge = Color.Empty; }
        if (alpha < 255)
        {
            back = Color.FromArgb(alpha, back);
            ink = Color.FromArgb(alpha, ink);
            if (edge != Color.Empty) { edge = Color.FromArgb(alpha, edge); }
        }
        if (back != Color.Empty)
        {
            using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, r); }
        }
        if (edge != Color.Empty) { Frame(g, edge, r); }
        Size ts = Measure(text, f);
        Draw(g, text, f, ink, r.X + (r.Width - ts.Width) / 2, r.Y + (r.Height - ts.Height) / 2);
    }

    // ---- the small line icons the reference draws as SVG -------------------
    public static void IconSearch(Graphics g, Rectangle r, Color c)
    {
        using (Pen p = new Pen(c, Math.Max(1f, 1.4f * Scale)))
        {
            p.StartCap = LineCap.Round; p.EndCap = LineCap.Round;
            int d = (int)(r.Width * 0.62);
            g.DrawEllipse(p, r.X + r.Width / 12, r.Y + r.Height / 12, d, d);
            g.DrawLine(p, r.X + r.Width / 12 + d - r.Width / 12, r.Y + r.Height / 12 + d - r.Height / 12,
                r.Right - r.Width / 10, r.Bottom - r.Height / 10);
        }
    }

    public static void IconCheck(Graphics g, Rectangle r, Color c)
    {
        using (Pen p = new Pen(c, Math.Max(1f, 1.4f * Scale)))
        {
            p.StartCap = LineCap.Round; p.EndCap = LineCap.Round;
            g.DrawLines(p, new Point[] {
                new Point(r.X + r.Width / 8, r.Y + r.Height / 2),
                new Point(r.X + r.Width * 2 / 5, r.Bottom - r.Height / 4),
                new Point(r.Right - r.Width / 8, r.Y + r.Height / 5) });
        }
    }

    public static void IconRefresh(Graphics g, Rectangle r, Color c)
    {
        using (Pen p = new Pen(c, Math.Max(1f, 1.4f * Scale)))
        {
            p.StartCap = LineCap.Round; p.EndCap = LineCap.Round;
            Rectangle a = new Rectangle(r.X + r.Width / 8, r.Y + r.Height / 8,
                r.Width - r.Width / 4, r.Height - r.Height / 4);
            g.DrawArc(p, a, 40, 250);
            int t = Math.Max(2, r.Width / 5);
            Point tip = new Point(a.Right - t / 2, a.Y + t / 3);
            g.DrawLine(p, tip, new Point(tip.X, tip.Y - t));
            g.DrawLine(p, tip, new Point(tip.X - t, tip.Y));
        }
    }

    public static void IconGear(Graphics g, Rectangle r, Color c)
    {
        float cx = r.X + r.Width / 2f, cy = r.Y + r.Height / 2f;
        float rad = Math.Min(r.Width, r.Height) / 2f;
        float ring = rad * 0.56f;
        using (Pen pen = new Pen(c, Math.Max(1f, rad * 0.20f)))
        {
            pen.StartCap = LineCap.Round;
            pen.EndCap = LineCap.Round;
            g.DrawEllipse(pen, cx - ring, cy - ring, ring * 2f, ring * 2f);
            for (int i = 0; i < 8; i++)
            {
                double a = Math.PI * i / 4.0;
                float dx = (float)Math.Cos(a), dy = (float)Math.Sin(a);
                g.DrawLine(pen, cx + dx * (ring + rad * 0.10f), cy + dy * (ring + rad * 0.10f),
                    cx + dx * rad, cy + dy * rad);
            }
        }
    }

    public static void IconPerson(Graphics g, Rectangle r, Color c)
    {
        using (Pen p = new Pen(c, Math.Max(1f, 1.4f * Scale)))
        {
            p.StartCap = LineCap.Round; p.EndCap = LineCap.Round;
            int hd = r.Width / 3;
            g.DrawEllipse(p, r.X + (r.Width - hd) / 2, r.Y + r.Height / 6, hd, hd);
            Rectangle sh = new Rectangle(r.X + r.Width / 6, r.Y + r.Height / 2,
                r.Width - r.Width / 3, r.Height * 3 / 4);
            g.DrawArc(p, sh, 200, 140);
        }
    }
}


// ---------------------------------------------------------------------------
// A button with the reference's own states. Taken from the artifact's CSS:
//   .btn            1px divider border, radius 0, icon and label centred with
//                   a 6 px gap, heading face at weight 600
//   .btn-primary    accent; :hover accent-600; :active accent-700
//   .btn-secondary  divider border; :hover text 7%; :active text 14%
//   .btn:disabled   opacity .45  (the WHOLE button, not just the label)
//   :focus-visible  2 px accent outline, 2 px outside -- painted by the form,
//                   since it falls outside these bounds
// ---------------------------------------------------------------------------
public sealed class Rdv3Btn : Button
{
    public bool Primary;
    public int Icon;                 // 0 none, 1 search, 2 check, 3 refresh
    private bool over, down;

    public Rdv3Btn()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        FlatStyle = FlatStyle.Flat;
        FlatAppearance.BorderSize = 0;
        BackColor = Rdv3Skin.Bg;
        UseVisualStyleBackColor = false;
    }

    protected override void OnMouseEnter(EventArgs e) { over = true; Invalidate(); base.OnMouseEnter(e); }
    protected override void OnMouseLeave(EventArgs e) { over = false; down = false; Invalidate(); base.OnMouseLeave(e); }
    protected override void OnMouseDown(MouseEventArgs e) { down = true; Invalidate(); base.OnMouseDown(e); }
    protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
    protected override void OnEnabledChanged(EventArgs e) { Invalidate(); base.OnEnabledChanged(e); }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        if (Enabled) { Render(g, 0, 0); return; }
        // disabled is opacity .45 over the page, so the whole face is composited
        using (Bitmap bmp = new Bitmap(Math.Max(1, Width), Math.Max(1, Height)))
        {
            using (Graphics gb = Graphics.FromImage(bmp))
            {
                gb.SmoothingMode = SmoothingMode.AntiAlias;
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { gb.FillRectangle(b, 0, 0, Width, Height); }
                Render(gb, 0, 0);
            }
            using (ImageAttributes ia = new ImageAttributes())
            {
                ColorMatrix m = new ColorMatrix();
                m.Matrix33 = 0.45f;
                ia.SetColorMatrix(m);
                g.DrawImage(bmp, new Rectangle(0, 0, Width, Height), 0, 0, Width, Height, GraphicsUnit.Pixel, ia);
            }
        }
    }

    private void Render(Graphics g, int dx, int dy)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        Rectangle r = new Rectangle(dx, dy, Width, Height);
        Color back, ink, edge;
        if (Primary)
        {
            back = down ? Rdv3Skin.Accent700 : (over ? Rdv3Skin.Accent600 : Rdv3Skin.Accent);
            ink = Rdv3Skin.Bg;
            edge = Rdv3Skin.Accent;
        }
        else
        {
            back = down ? Rdv3Skin.Mix(0.14) : (over ? Rdv3Skin.Mix(0.07) : Rdv3Skin.Bg);
            ink = Rdv3Skin.Ink;
            edge = Rdv3Skin.Divider;
        }
        using (SolidBrush b = new SolidBrush(back)) { g.FillRectangle(b, r); }
        Rdv3Skin.Frame(g, edge, r);

        int sz = Rdv3Skin.P(14), gap = Rdv3Skin.P(6);
        int tw = Rdv3Skin.Measure(Text, Font).Width;
        int th = Rdv3Skin.Measure(Text, Font).Height;
        int all = (Icon > 0) ? (sz + gap + tw) : tw;
        int x = r.X + (r.Width - all) / 2;
        if (Icon > 0)
        {
            Rectangle ir = new Rectangle(x, r.Y + (r.Height - sz) / 2, sz, sz);
            if (Icon == 1) { Rdv3Skin.IconSearch(g, ir, ink); }
            else if (Icon == 2) { Rdv3Skin.IconCheck(g, ir, ink); }
            else if (Icon == 4) { Rdv3Skin.IconGear(g, ir, ink); }
            else { Rdv3Skin.IconRefresh(g, ir, ink); }
            x += sz + gap;
        }
        Rdv3Skin.Draw(g, Text, Font, ink, x, r.Y + (r.Height - th) / 2);
    }
}

// ---------------------------------------------------------------------------
// one candidate row as the screen needs it
// ---------------------------------------------------------------------------
public sealed class Rdv3CandRow
{
    public string[] Cols;      // key2, line, slip, date, qty, status, item, maker
    public bool Processed;
    public string Line;        // the ledger content line, for the record panel
}

// ---------------------------------------------------------------------------
// the candidate table: sticky header, hover, selection, wheel + scrollbar
// ---------------------------------------------------------------------------
internal sealed class Rdv3CandTable : Panel
{
    // the reference's scrollbar is a slim painted track, not a Win32 one with
    // arrow buttons; it is drawn here and dragged with the mouse
    private bool dragging;
    private int dragGrab;
    private List<Rdv3CandRow> rows = new List<Rdv3CandRow>();
    private int sel = -1;
    private int hot = -1;
    private int top;                       // first visible row
    public Action<int> OnPick;

    private Font fTh, fTd, fTdBold, fTag;

    private static readonly double[] ColW = { 44, 94.7, 69.6, 103.8, 88.8, 57.8, 83.3, 85.9, 107.4, 80.5 };
    private static readonly string[] ColHead = { "#", "番号2", "行番号", "伝票番号",
        "日付", "数量", "状態", "品目コード", "メーカー",
        "処理済み" };

    public Rdv3CandTable()
    {
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);
        BackColor = Rdv3Skin.Bg;
        MakeFonts();
    }

    public void MakeFonts()
    {
        fTh = Rdv3Skin.F(11, FontStyle.Bold);
        fTd = Rdv3Skin.F(13, FontStyle.Regular);
        fTdBold = Rdv3Skin.F(13, FontStyle.Bold);
        fTag = Rdv3Skin.F(11, FontStyle.Regular);
        colw = null;
    }

    public int HeaderH { get { return Rdv3Skin.P(31); } }
    public int RowH { get { return Rdv3Skin.P(38.6); } }
    public int Count { get { return rows.Count; } }
    public int Selected { get { return sel; } }

    public void SetRows(List<Rdv3CandRow> r, int selected)
    {
        rows = (r == null) ? new List<Rdv3CandRow>() : r;
        sel = selected;
        top = 0;
        hot = -1;
        MeasureColumns();
        SyncBar();
        Invalidate();
    }

    private int[] colw;

    // the reference's column widths are a floor; what the rows actually need
    // decides the rest, so no cell ever runs into the next column
    private void MeasureColumns()
    {
        int padL = Rdv3Skin.P(6.8), padR = Rdv3Skin.P(6.8);   // the reference's cell padding
        int[] w = new int[ColW.Length];
        for (int c = 0; c < ColW.Length; c++)
        {
            w[c] = Math.Max(Rdv3Skin.P(ColW[c]), Rdv3Skin.Measure(ColHead[c], fTh).Width + padL + padR);
        }
        for (int i = 0; i < rows.Count; i++)
        {
            Rdv3CandRow r = rows[i];
            for (int c = 0; c < 10; c++)
            {
                int need;
                if (c == 0) { need = Rdv3Skin.Measure((i + 1).ToString(CultureInfo.InvariantCulture), fTd).Width; }
                else if (c == 6) { need = Rdv3Skin.TagSize(r.Cols[5], fTag).Width; }
                else if (c == 9) { need = Rdv3Skin.TagSize(r.Processed ? "済" : "未", fTag).Width; }
                else { need = Rdv3Skin.Measure(r.Cols[ColData[c]], (c == 1) ? fTdBold : fTd).Width; }
                need += padL + padR;
                if (need > w[c]) { w[c] = need; }
            }
        }
        colw = w;
    }

    // which Cols index each visible column takes (0 and 6/9 are special)
    private static readonly int[] ColData = { -1, 0, 1, 2, 3, 4, -1, 6, 7, -1 };

    public void Select(int i)
    {
        sel = i;
        if (i >= 0)
        {
            int vis = VisibleRows();
            if (i < top) { top = i; }
            else if (i >= top + vis) { top = Math.Max(0, i - vis + 1); }
            SyncBar();
        }
        Invalidate();
    }

    public void MarkProcessed(int i)
    {
        if (i >= 0 && i < rows.Count) { rows[i].Processed = true; Invalidate(); }
    }

    public Rdv3CandRow Row(int i) { return (i >= 0 && i < rows.Count) ? rows[i] : null; }

    private int VisibleRows()
    {
        int h = ClientSize.Height - HeaderH;
        return Math.Max(1, h / Math.Max(1, RowH));
    }

    private int BarW { get { return Rdv3Skin.P(10); } }

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
        int h = Math.Max(Rdv3Skin.P(24), (int)((long)t.Height * vis / rows.Count));
        int span = t.Height - h;
        int y = t.Y + (int)((long)span * top / Math.Max(1, rows.Count - vis));
        int pad = Rdv3Skin.P(3);
        return new Rectangle(t.X + pad, y, Math.Max(2, t.Width - 2 * pad), h);
    }

    private void ScrollTo(int t)
    {
        int vis = VisibleRows();
        int max = Math.Max(0, rows.Count - vis);
        int v = Math.Max(0, Math.Min(t, max));
        if (v != top) { top = v; Invalidate(); }
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); SyncBar(); }

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

    // the acceptance dump needs the same edges the painter uses
    public int[] ColumnEdges() { return ColEdges(); }

    private int[] ColEdges()
    {
        if (colw == null) { MeasureColumns(); }
        int[] x = new int[colw.Length + 1];
        x[0] = 0;
        for (int i = 0; i < colw.Length; i++) { x[i + 1] = x[i] + colw[i]; }
        return x;
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        int[] xs = ColEdges();
        int padL = Rdv3Skin.P(6.8), padR = Rdv3Skin.P(13.6);
        int w = ClientSize.Width - (NeedBar ? BarW : 0);

        // rows first, then the header over them (sticky)
        int y = HeaderH;
        for (int i = top; i < rows.Count && y < ClientSize.Height; i++)
        {
            Rdv3CandRow r = rows[i];
            Rectangle rr = new Rectangle(0, y, w, RowH);
            if (i == sel)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent100)) { g.FillRectangle(b, rr); }
                Rdv3Skin.Line(g, Rdv3Skin.Accent, 0, y, Rdv3Skin.P(2), RowH);
            }
            else if (i == hot)
            {
                using (SolidBrush b = new SolidBrush(Rdv3Skin.Hover)) { g.FillRectangle(b, rr); }
            }
            Rdv3Skin.Line(g, Rdv3Skin.RowLine, 0, rr.Bottom - Rdv3Skin.Hair(), w, Rdv3Skin.Hair());

            int ty = y + (RowH - Rdv3Skin.Measure("0", fTd).Height) / 2;
            Rdv3Skin.DrawRight(g, (i + 1).ToString(CultureInfo.InvariantCulture), fTd, Rdv3Skin.N600,
                xs[1] - padR, ty);
            Rdv3Skin.Draw(g, r.Cols[0], fTdBold, Rdv3Skin.Ink, xs[1] + padL, ty);
            Rdv3Skin.DrawRight(g, r.Cols[1], fTd, Rdv3Skin.Ink, xs[3] - padR, ty);
            Rdv3Skin.Draw(g, r.Cols[2], fTd, Rdv3Skin.Ink, xs[3] + padL, ty);
            Rdv3Skin.Draw(g, r.Cols[3], fTd, Rdv3Skin.Ink, xs[4] + padL, ty);
            Rdv3Skin.DrawRight(g, r.Cols[4], fTd, Rdv3Skin.Ink, xs[6] - padR, ty);

            string st = r.Cols[5];
            int kind = Rdv3Skin.TagNeutral;
            int alpha = 255;
            if (st == "DONE") { kind = Rdv3Skin.TagAccent; }
            else if (st == "HOLD") { kind = Rdv3Skin.TagOutline; }
            else if (st == "VOID") { alpha = 128; }
            Size tz = Rdv3Skin.TagSize(st, fTag);
            Rdv3Skin.Tag(g, st, fTag, kind, new Rectangle(xs[6] + padL, y + (RowH - tz.Height) / 2, tz.Width, tz.Height), alpha);

            Rdv3Skin.Draw(g, r.Cols[6], fTd, Rdv3Skin.Ink, xs[7] + padL, ty);
            Rdv3Skin.Draw(g, r.Cols[7], fTd, Rdv3Skin.Ink, xs[8] + padL, ty);

            string pt = r.Processed ? "済" : "未";
            Size pz = Rdv3Skin.TagSize(pt, fTag);
            Rdv3Skin.Tag(g, pt, fTag, r.Processed ? Rdv3Skin.TagAccent : Rdv3Skin.TagNeutral,
                new Rectangle(xs[9] + padL, y + (RowH - pz.Height) / 2, pz.Width, pz.Height), 255);

            y += RowH;
        }

        // header
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, 0, 0, ClientSize.Width, HeaderH); }
        int hy = (HeaderH - Rdv3Skin.Measure("A", fTh).Height) / 2;
        for (int i = 0; i < ColHead.Length; i++)
        {
            bool right = (i == 0 || i == 2 || i == 5);
            if (right) { Rdv3Skin.DrawRight(g, ColHead[i], fTh, Rdv3Skin.ThInk, xs[i + 1] - padR, hy); }
            else { Rdv3Skin.Draw(g, ColHead[i], fTh, Rdv3Skin.ThInk, xs[i] + padL, hy); }
        }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, 0, HeaderH - Rdv3Skin.Hair(), ClientSize.Width, Rdv3Skin.Hair());

        if (NeedBar)
        {
            Rectangle tr = BarTrack(), th = BarThumb();
            using (SolidBrush b = new SolidBrush(Color.FromArgb(10, Rdv3Skin.Ink))) { g.FillRectangle(b, tr); }
            using (SolidBrush b = new SolidBrush(Color.FromArgb(dragging ? 92 : 56, Rdv3Skin.Ink)))
            {
                g.FillRectangle(b, th);
            }
        }
    }
}

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

    // live controls
    private readonly TextBox txtKey = new TextBox();
    private readonly Rdv3Btn btnSearch = new Rdv3Btn();
    private readonly Rdv3Btn btnClear = new Rdv3Btn();
    private readonly Rdv3Btn btnProcessed = new Rdv3Btn();
    private readonly Rdv3Btn btnRebind = new Rdv3Btn();
    private readonly Rdv3Btn btnSettings = new Rdv3Btn();
    private readonly Rdv3CandTable table = new Rdv3CandTable();
    private readonly TextBox txtMemo = new TextBox();
    private readonly TextBox txtRemark = new TextBox();
    private readonly ToolTip tips = new ToolTip();
    private readonly System.Windows.Forms.Timer clock = new System.Windows.Forms.Timer();

    public Action<string> OnSearch;
    public Action OnClear;
    public Action OnProcessed;
    public Action OnRebind;
    public Action OnSettings;
    public Action<int> OnPick;

    // screen state
    private string sKey = "";
    private string sKeySub = "";
    private string sStatus = "";
    private string sStatusSub = "";
    private bool sStatusNg;
    private string sRows = "0";
    private string sSaved = "";
    private string sState = Rdv3Text.StateBoot;
    private string sNotepad = Rdv3Text.NotepadNone;
    // which target the status line names, and the input label with the key
    // length filled in: both come from the settings file
    private string sWatchLabel = Rdv3Text.LabelNotepad;
    private string sKeyLabel = Rdv3Text.LabelSearchBox.Replace("{n}",
        Rdv3Spec.KeyLen.ToString(CultureInfo.InvariantCulture));
    private string sLedger = "";
    private string sMerge = Rdv3Text.NotYet;
    private string sSearch = Rdv3Text.NotYet;
    private string sError = "";
    // a bad key is answered under the input box, not in the error row at the
    // bottom: it belongs to the thing the operator just typed, and moving the
    // rest of the screen for it would be worse than the mistake
    private string sInputError = "";
    private string sVerdict = "";
    private string sCandTag = "";
    private string sPid = "";
    private string sLog = "";
    private string sUser = "";
    private string sHost = "";
    private string sRole = "";
    private string[] rec;                 // the selected ledger line, split
    private bool recProc;
    private string sKey2 = "";
    private string sRecStatus = "";
    private string procSuffix = "";

    // fonts
    private Font fBrand, fTag, fBtn, fBtnSm, fLabel10, fBig, fBig13, fSub12, fH4, fHint,
                 fVerdict, fKv, fKvB, fBox, fBoxLabel, fStatus, fStatusB, fName, fName11;

    private static readonly string[] KvLabels = {
        "取引先名", "取引先コード", "グレード",
        "部門", "登録日", "金額", "レート ・ フラグ" };
    // ledger content columns behind them: a_name, a_code, a_grade, a_dept, a_date, a_amount, (a_rate . a_flag)
    private static readonly int[] KvCols = { 3, 2, 4, 9, 5, 6, 7 };

    private bool scaled;

    public Rdv3Form()
    {
        MakeDpiAware();
        Rdv3Skin.PickFamily();
        MakeFonts();

        Text = Rdv3Text.AppTitle;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Rdv3Skin.Page;
        DoubleBuffered = true;
        // every coordinate here is the skin's own; WinForms must not re-scale
        AutoScaleMode = AutoScaleMode.None;
        SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
            | ControlStyles.UserPaint | ControlStyles.ResizeRedraw, true);

        // ---- live controls -------------------------------------------------
        txtKey.BorderStyle = BorderStyle.None;
        txtKey.MaxLength = Rdv3Spec.KeyLen;
        txtKey.Font = Rdv3Skin.F(14, FontStyle.Regular);
        txtKey.BackColor = Rdv3Skin.Surface;
        txtKey.ForeColor = Rdv3Skin.Ink;
        txtKey.TextChanged += delegate
        {
            if (sInputError.Length > 0) { sInputError = ""; Layout1(); }
        };
        txtKey.KeyDown += delegate(object s, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Fire(OnSearch); }
        };
        Controls.Add(txtKey);

        StyleButton(btnSearch, Rdv3Text.BtnSearch, true, 1);
        StyleButton(btnClear, Rdv3Text.BtnClear, false, 0);
        StyleButton(btnProcessed, Rdv3Text.BtnProcessed, false, 2);
        StyleButton(btnRebind, Rdv3Text.BtnRebind, false, 3);
        StyleButton(btnSettings, Rdv3Text.BtnSettings, false, 4);
        btnRebind.Font = fBtnSm;
        btnSettings.Font = fBtnSm;
        btnSettings.Click += delegate { if (OnSettings != null) { OnSettings(); } };
        btnSearch.Click += delegate { Fire(OnSearch); };
        btnClear.Click += delegate { if (OnClear != null) { OnClear(); } };
        btnProcessed.Click += delegate { if (OnProcessed != null) { OnProcessed(); } };
        btnRebind.Click += delegate { if (OnRebind != null) { OnRebind(); } };
        Controls.Add(btnSearch); Controls.Add(btnClear); Controls.Add(btnProcessed);
        Controls.Add(btnRebind); Controls.Add(btnSettings);

        table.OnPick = delegate(int i) { if (OnPick != null) { OnPick(i); } };
        Controls.Add(table);

        StyleBox(txtMemo);
        StyleBox(txtRemark);
        Controls.Add(txtMemo); Controls.Add(txtRemark);

        tips.SetToolTip(btnSearch, Rdv3Text.TipSearch);
        tips.SetToolTip(btnClear, Rdv3Text.TipClear);
        tips.SetToolTip(btnProcessed, Rdv3Text.TipProcessed);

        txtKey.MouseEnter += delegate { inputHot = true; Invalidate(At("sum.input")); };
        txtKey.MouseLeave += delegate { inputHot = false; Invalidate(At("sum.input")); };
        txtKey.Enter += delegate { Invalidate(); };
        txtKey.Leave += delegate { Invalidate(); };
        table.Enter += delegate { Invalidate(); };
        table.Leave += delegate { Invalidate(); };
        KeyPreview = true;
        AcceptButton = btnSearch;
        clock.Interval = 1000;
        clock.Tick += delegate { RelayoutStatus(); };
        clock.Start();

        sUser = SafeEnv("USERNAME");
        sHost = SafeEnv("COMPUTERNAME");
        sRole = Rdv3Text.RoleNormal;
        ResetState();
        Layout1();
    }

    private static void MakeDpiAware()
    {
        // per-monitor v2, then per-monitor, then system aware; the process-wide
        // call is the last resort (it fails in a host that already has windows)
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

    // the sizes and weights the reference computes: 19/600 condensed brand,
    // 11/400 tags, 13/600 buttons, 10/400 labels, 34/600 condensed figures,
    // 12/400 sub-lines and hints, 18/600 panel titles, 13/700 verdicts and
    // record values, 11/700 table headers, 12/400 status bar.
    private void MakeFonts()
    {
        Rdv3Skin.ResetMetrics();
        fBrand = Rdv3Skin.H(19, FontStyle.Regular);
        fTag = Rdv3Skin.F(11, FontStyle.Regular);
        fBtn = Rdv3Skin.S(14);        // .btn
        fBtnSm = Rdv3Skin.S(13);      // the nav bar's smaller cut
        fLabel10 = Rdv3Skin.F(10, FontStyle.Regular);
        fBig = Rdv3Skin.D(34);
        fBig13 = Rdv3Skin.F(13, FontStyle.Regular);
        fSub12 = Rdv3Skin.F(12, FontStyle.Regular);
        fH4 = Rdv3Skin.S(18);
        fHint = Rdv3Skin.F(12, FontStyle.Regular);
        fVerdict = Rdv3Skin.F(13, FontStyle.Bold);
        fKv = Rdv3Skin.F(13, FontStyle.Regular);
        fKvB = Rdv3Skin.F(13, FontStyle.Bold);
        fBox = Rdv3Skin.F(13, FontStyle.Regular);
        fBoxLabel = Rdv3Skin.F(11, FontStyle.Regular);
        fStatus = Rdv3Skin.F(12, FontStyle.Regular);
        fStatusB = Rdv3Skin.F(12, FontStyle.Bold);
        fName = Rdv3Skin.F(13, FontStyle.Bold);
        fName11 = Rdv3Skin.F(11, FontStyle.Regular);
    }

    private void StyleButton(Rdv3Btn b, string text, bool primary, int icon)
    {
        b.Text = text;
        b.Font = fBtn;
        b.Primary = primary;
        b.Icon = icon;
        b.TabStop = true;
        b.Enter += delegate { Invalidate(); };
        b.Leave += delegate { Invalidate(); };
    }

    // WinForms shows a scrollbar for good, or never; the reference shows one
    // only when the memo is longer than its box, so the decision is made per
    // text and the box stays clean for the ordinary case.
    private void SetBox(TextBox t, string text)
    {
        t.Text = (text == null) ? "" : text;
        bool over = false;
        if (t.Text.Length > 0 && t.Width > 8)
        {
            Size z = TextRenderer.MeasureText(t.Text, t.Font, new Size(t.Width, int.MaxValue),
                TextFormatFlags.WordBreak | TextFormatFlags.NoPadding);
            over = z.Height > t.Height;
        }
        ScrollBars want = over ? ScrollBars.Vertical : ScrollBars.None;
        if (t.ScrollBars != want) { t.ScrollBars = want; }
    }

    private void StyleBox(TextBox t)
    {
        t.Multiline = true;
        t.ReadOnly = true;
        t.BorderStyle = BorderStyle.None;
        t.BackColor = Rdv3Skin.N100;
        t.ForeColor = Rdv3Skin.Ink;
        t.Font = fBox;
        t.ScrollBars = ScrollBars.None;
        t.TabStop = false;
    }

    private void Fire(Action<string> a)
    {
        if (a != null) { a(txtKey.Text.Trim()); }
    }

    // ---- layout: the reference's own coordinates ---------------------------
    // Rdv3Geom is the artifact measured rectangle by rectangle, card-relative,
    // in CSS px. Layout1 turns that table into device rectangles and stores
    // them in `rc`; the painter, the live controls and the acceptance dump all
    // read that one dictionary, so what is measured IS what gets drawn.
    //
    // Nothing here re-invents a flex row. An element may only be WIDER than the
    // reference when the real string is wider than the artifact's sample one,
    // and the elements after it then move by that difference. Overlap is
    // therefore structurally impossible, and at the sample strings every
    // rectangle lands on the reference's own coordinates.
    private readonly Dictionary<string, Rectangle> rc = new Dictionary<string, Rectangle>();
    private Rectangle rStatus, rErr;

    // extra space over the design: width, and the two panels' share of height
    private double exW, exList, exRec;
    private double kvRowH = Rdv3Geom.KvH;
    private double recLeftW = 513.667;
    private double recRightX = 569.4;
    private double recRightW = 642.094;
    private double recBoxH = 74.4;

    private static int PX(double v) { return (int)Math.Round(v * Rdv3Skin.Scale); }
    private static double CX(int devicePx) { return devicePx / (double)Rdv3Skin.Scale; }

    // One device pixel of slack. Widths are measured in device pixels, carried
    // as CSS px and rounded back, and that round trip can leave a box a pixel
    // short of its own text -- which the painter then ellipsises. At 125% the
    // longest sub-line ("HOLD / 未処理 (保存中...)") lost its tail to exactly
    // that pixel.
    private double MW(string s, Font f)
    {
        return (Rdv3Skin.Measure(s, f).Width + 1) / (double)Rdv3Skin.Scale;
    }

    private double TagW(string s)
    {
        return MW(s, fTag) + 20.0;
    }

    private double BtnCss(Button b, Font f)
    {
        double icon = (b == btnClear) ? 0.0 : 20.0;
        return MW(b.Text, f) + icon + 24.48;
    }

    private void Put(string k, double x, double y, double w, double h)
    {
        int x0 = PX(x), y0 = PX(y);
        rc[k] = new Rectangle(x0, y0, PX(x + w) - x0, PX(y + h) - y0);
    }

    private Rectangle At(string k)
    {
        Rectangle r;
        return rc.TryGetValue(k, out r) ? r : Rectangle.Empty;
    }

    private void Layout1()
    {
        if (fTag == null) { return; }
        double cw = CX(ClientSize.Width), ch = CX(ClientSize.Height);
        if (cw < 100 || ch < 100) { return; }
        exW = Math.Max(0, cw - Rdv3Geom.CardW);
        double errH = (sError.Length > 0) ? Rdv3Geom.ErrRow.H : 0.0;
        double exH = ch - Rdv3Geom.CardH - errH;
        if (exH >= 0) { exList = Math.Floor(exH / 2.0); exRec = exH - exList; }
        else { exList = 0.0; exRec = exH; }        // the record body gives way first

        rc.Clear();
        LayNav();
        LaySummary();
        LayList();
        LayRecord();
        LayFoot(cw, ch, errH);
        rStatus = At("status");
        rErr = At("err");
        PlaceControls();
        Invalidate();
        EnsureRoom();
    }

    // The reference's 1240 px is a minimum, not a promise: with a long user
    // name or a nine digit ledger the same row needs more, and a native window
    // can simply say so. The minimum width follows the content (never below the
    // design, never past the work area) and the window grows with it once.
    private bool sizing;
    private double needW = Rdv3Geom.CardW;

    private void EnsureRoom()
    {
        if (sizing || !IsHandleCreated) { return; }
        Rectangle wa = Screen.FromControl(this).WorkingArea;
        int frame = Math.Max(0, Width - ClientSize.Width);
        int want = Math.Min(PX(Math.Max(Rdv3Geom.CardW, needW)) + frame, Math.Max(400, wa.Width));
        if (want <= MinimumSize.Width) { return; }
        sizing = true;
        try
        {
            MinimumSize = new Size(want, MinimumSize.Height);
            if (Width < want) { Width = want; }
        }
        finally { sizing = false; }
        Layout1();
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); Layout1(); }

    // the browser shows its focus ring for keyboard moves only; so does this
    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        Keys k = keyData & Keys.KeyCode;
        if (k == Keys.Tab || k == Keys.Up || k == Keys.Down || k == Keys.Left || k == Keys.Right)
        {
            if (!keyFocus) { keyFocus = true; Invalidate(); }
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        if (keyFocus) { keyFocus = false; Invalidate(); }
        base.OnMouseDown(e);
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        bool hot = At("sum.input").Contains(e.Location);
        if (hot != inputHot) { inputHot = hot; Invalidate(At("sum.input")); }
        base.OnMouseMove(e);
    }

    private void LayNav()
    {
        Rdv3Geom.R n = Rdv3Geom.Nav;
        Put("nav", n.X, n.Y, n.W + exW, n.H);
        Rdv3Geom.R b = Rdv3Geom.NavBrand;
        double bw = Math.Max(b.W, MW(Rdv3Text.AppTitle, fBrand));
        Put("nav.brand", b.X, b.Y, bw, b.H);
        double x = b.X + bw + 10.2;
        double t1 = Math.Max(Rdv3Geom.NavTagMethod.W, TagW(Rdv3Text.Method));
        Put("nav.tagMethod", x, Rdv3Geom.NavTagMethod.Y, t1, Rdv3Geom.NavTagMethod.H);
        x += t1 + 10.2;
        double t2 = Math.Max(Rdv3Geom.NavTagLedger.W, TagW(Rdv3Text.TagLedger));
        Put("nav.tagLedger", x, Rdv3Geom.NavTagMethod.Y, t2, Rdv3Geom.NavTagMethod.H);
        Rdv3Geom.R rb = Rdv3Geom.NavBtnRebind;
        double edge = rb.X + rb.W + exW;
        double sw = BtnCss(btnSettings, fBtnSm);
        Put("nav.btnSettings", edge - sw, rb.Y, sw, rb.H);
        double rw = Math.Max(rb.W, BtnCss(btnRebind, fBtnSm));
        Put("nav.btnRebind", edge - sw - 6.8 - rw, rb.Y, rw, rb.H);
        Put("nav.rule", n.X, n.Y + n.H - 1.0, n.W + exW, 1.0);
    }

    // the three figure blocks, the input group and the identity block. The
    // reference's own widths are the floor; only a longer real string widens a
    // block, and then the divider after it moves too.
    private void LaySummary()
    {
        Rdv3Geom.R p = Rdv3Geom.SumPanel;
        Put("sum.panel", p.X, p.Y, p.W + exW, p.H);
        double panelR = p.X + p.W + exW;

        double wKey = Math.Max(Rdv3Geom.SumKeyLabel.W, FigW(Rdv3Text.LabelKeyNow, sKey, sKeySub, false));
        double wSt = Math.Max(Rdv3Geom.SumStatusLabel.W, FigW(Rdv3Text.LabelRepStatus, sStatus, sStatusSub, false));
        double wRows = Math.Max(Rdv3Geom.SumRowsLabel.W, FigW(Rdv3Text.LabelLedgerRows, sRows, SavedSub, true));

        // the identity block, right-anchored to the panel
        double idName = Math.Max(MW(sUser, fName), MW(sHost, fName11));
        double idTag = Math.Max(Rdv3Skin.P(0) + 0.0, TagW(sRole));
        double idW = Math.Max(Rdv3Geom.SumIdent.W, 26.0 + 6.7 + idName + 6.8 + idTag);
        double idX = panelR - 13.6 - idW;
        double idDiv = idX - 14.3;

        // the input group, right-anchored to the identity divider
        double wS = Math.Max(Rdv3Geom.SumBtnSearch.W, BtnCss(btnSearch, fBtn));
        double wC = Math.Max(Rdv3Geom.SumBtnClear.W, BtnCss(btnClear, fBtn));
        double wP = Math.Max(Rdv3Geom.SumBtnProcessed.W, BtnCss(btnProcessed, fBtn));
        double inW = Rdv3Geom.SumInput.W;
        double fieldW = inW + 6.8 + wS + 6.8 + wC + 6.8 + wP;
        double fieldR = idDiv - 13.6;
        double fieldX = fieldR - fieldW;

        // What the row needs at these fonts. The window's minimum width follows
        // it (EnsureRoom), so on a normal screen this never has to give way.
        double leftEnd = 28.5 + wKey + 13.6 + 14.3 + wSt + 13.6 + 14.3 + wRows;
        needW = leftEnd + 13.6 + fieldW + 13.6 + 14.3 + idW + 13.6 + 14.3;

        // Only if even the screen is too narrow does something give: the
        // sub-lines, never the figures. A truncated number is a wrong number.
        double over = leftEnd + 13.6 - fieldX;
        if (over > 0)
        {
            double[] w = { wKey, wSt, wRows };
            double[] floor = { FigFloor(Rdv3Text.LabelKeyNow, sKey, false),
                FigFloor(Rdv3Text.LabelRepStatus, sStatus, false),
                FigFloor(Rdv3Text.LabelLedgerRows, sRows, true) };
            for (int guard = 0; guard < 64 && over > 0; guard++)
            {
                int widest = (w[0] - floor[0] >= w[1] - floor[1] && w[0] - floor[0] >= w[2] - floor[2])
                    ? 0 : ((w[1] - floor[1] >= w[2] - floor[2]) ? 1 : 2);
                double cut = Math.Min(over, Math.Max(0, w[widest] - floor[widest]));
                if (cut <= 0.01) { break; }
                w[widest] -= cut;
                over -= cut;
            }
            wKey = w[0]; wSt = w[1]; wRows = w[2];
        }

        double x0 = Rdv3Geom.SumKeyLabel.X;
        PutFigure("sum.key", x0, wKey);
        double div1 = x0 + wKey + 13.6;
        Put("sum.status.div", div1, Rdv3Geom.SumStatusDiv.Y, 1.0, Rdv3Geom.SumStatusDiv.H);
        double x1 = div1 + 14.3;
        PutFigure("sum.status", x1, wSt);
        double div2 = x1 + wSt + 13.6;
        Put("sum.rows.div", div2, Rdv3Geom.SumRowsDiv.Y, 1.0, Rdv3Geom.SumRowsDiv.H);
        PutFigure("sum.rows", div2 + 14.3, wRows);

        Rdv3Geom.R fl = Rdv3Geom.SumFieldLabel;
        Put("sum.field.label", fieldX, fl.Y, Math.Max(inW, MW(sKeyLabel, fSub12)), fl.H);
        Put("sum.input", fieldX, Rdv3Geom.SumInput.Y, inW, Rdv3Geom.SumInput.H);
        double bx = fieldX + inW + 6.8, by = Rdv3Geom.SumBtnSearch.Y, bh = Rdv3Geom.SumBtnSearch.H;
        Put("sum.btnSearch", bx, by, wS, bh); bx += wS + 6.8;
        Put("sum.btnClear", bx, by, wC, bh); bx += wC + 6.8;
        Put("sum.btnProcessed", bx, by, wP, bh);

        Put("sum.field.error", fieldX, 130.0, fieldW, 16.0);

        Put("sum.ident.div", idDiv, Rdv3Geom.SumIdentDiv.Y, 1.0, Rdv3Geom.SumIdentDiv.H);
        Put("sum.ident.icon", idX, 93.6, 26.0, 26.0);
        Put("sum.ident.name", idX + 32.7, 90.3, idName, 15.6);
        Put("sum.ident.host", idX + 32.7, 105.8, idName, 17.0);
        Put("sum.ident.role", idX + 32.7 + idName + 6.8, 95.0, idTag, 23.0);
    }

    private string SavedSub
    {
        get { return Rdv3Text.LabelLastSaved + " " + sSaved; }
    }

    // label / figure / sub-line, the reference's three bands
    private void PutFigure(string k, double x, double w)
    {
        Put(k + ".label", x, Rdv3Geom.SumKeyLabel.Y, w, Rdv3Geom.SumKeyLabel.H);
        Put(k + ".value", x, Rdv3Geom.SumKeyValue.Y, w, Rdv3Geom.SumKeyValue.H);
        Put(k + ".sub", x, Rdv3Geom.SumKeySub.Y, w, Rdv3Geom.SumKeySub.H);
    }

    // the block is as wide as its widest band, measured exactly as it is drawn
    // (tracking included) -- otherwise the painter would clip its own figure
    private double FigW(string label, string value, string sub, bool rows)
    {
        double v = (Rdv3Skin.MeasureTracked(value, fBig, ValueTrack) + 1) / (double)Rdv3Skin.Scale;
        if (rows) { v += 4.0 + MW(Rdv3Text.UnitRows, fBig13); }
        double w = Math.Max((Rdv3Skin.MeasureTracked(label, fLabel10, LabelTrack) + 1) / (double)Rdv3Skin.Scale, v);
        return Math.Max(w, MW(sub, fSub12));
    }

    // the label and the figure: the part of a block that may not be cut
    private double FigFloor(string label, string value, bool rows)
    {
        double v = (Rdv3Skin.MeasureTracked(value, fBig, ValueTrack) + 1) / (double)Rdv3Skin.Scale;
        if (rows) { v += 4.0 + MW(Rdv3Text.UnitRows, fBig13); }
        return Math.Max((Rdv3Skin.MeasureTracked(label, fLabel10, LabelTrack) + 1) / (double)Rdv3Skin.Scale, v);
    }

    // The reference tracks its 10 px section labels at 0.1em, which is what
    // makes them read as labels rather than text, and its 34 px figures at
    // 0.06em. The figures DO NOT get it here: Barlow Condensed is narrow, so a
    // little air suits it, while the Bahnschrift cut standing in for it is
    // already wide and the same tracking just looks stretched -- and the key
    // will hold letters as well as digits, where per-glyph spacing is exactly
    // where unevenness would show. Labels are fixed Japanese words and keep it.
    private const double LabelTrack = 1.0;
    private const double ValueTrack = 0.0;

    private void LayList()
    {
        Rdv3Geom.R p = Rdv3Geom.ListPanel;
        double panelR = p.X + p.W + exW;
        Put("list.panel", p.X, p.Y, p.W + exW, p.H + exList);
        double hw = Math.Max(Rdv3Geom.ListH4.W, MW(Rdv3Text.PanelCand, fH4));
        Put("list.h4", Rdv3Geom.ListH4.X, Rdv3Geom.ListH4.Y, hw, Rdv3Geom.ListH4.H);
        double tw = Math.Max(Rdv3Geom.ListTag.W, TagW(sCandTag));
        double tagX = panelR - 13.6 - tw;
        Put("list.tag", tagX, Rdv3Geom.ListTag.Y, tw, Rdv3Geom.ListTag.H);
        double hx = Rdv3Geom.ListH4.X + hw + 10.2;
        double room = tagX - 13.6 - hx;
        double hintW = Math.Min(MW(Rdv3Text.HintCand, fHint), Math.Max(0, room));
        Put("list.hint", hx, Rdv3Geom.ListHint.Y, hintW, Rdv3Geom.ListHint.H);
        double vx = hx + hintW + 20.4;
        double vw = Math.Min(MW(sVerdict, fVerdict), Math.Max(0, tagX - 13.6 - vx));
        Put("list.verdict", vx, 185.0, vw, 20.1);
        Put("list.rule", Rdv3Geom.ListHeadRule.X, Rdv3Geom.ListHeadRule.Y, Rdv3Geom.ListHeadRule.W + exW, 1.0);
        Rdv3Geom.R sc = Rdv3Geom.ListScroll;
        Put("list.scroll", sc.X, sc.Y, sc.W + exW, sc.H + exList);
    }

    private void LayRecord()
    {
        double dy = exList;
        Rdv3Geom.R p = Rdv3Geom.RecPanel;
        double panelR = p.X + p.W + exW;
        Put("rec.panel", p.X, p.Y + dy, p.W + exW, p.H + exRec);
        double hw = Math.Max(Rdv3Geom.RecH4.W, MW(Rdv3Text.PanelRec, fH4));
        Put("rec.h4", Rdv3Geom.RecH4.X, Rdv3Geom.RecH4.Y + dy, hw, Rdv3Geom.RecH4.H);

        string[] tags = RecTags();
        double tx = panelR - 13.6;
        for (int i = 2; i >= 0; i--)
        {
            if (tags[i].Length == 0) { continue; }
            double tw = Math.Max(27.8, TagW(tags[i]));
            tx -= tw;
            Put("rec.tag" + i.ToString(CultureInfo.InvariantCulture), tx, Rdv3Geom.RecTag0.Y + dy, tw, 23.0);
            tx -= 10.2;
        }
        double hx = Rdv3Geom.RecH4.X + hw + 10.2;
        double hintW = Math.Min(MW(Rdv3Text.HintRec, fHint), Math.Max(0, tx - hx));
        Put("rec.hint", hx, 399.7 + dy, hintW, 18.6);
        Put("rec.rule", Rdv3Geom.RecHeadRule.X, Rdv3Geom.RecHeadRule.Y + dy, Rdv3Geom.RecHeadRule.W + exW, 1.0);

        // the body is the reference's two-column grid, 513.667 : 642.094 with a
        // 27.2 gutter; extra width is shared in the same proportion
        double innerX = Rdv3Geom.RecKv0.X;
        double innerW = 1182.9 + exW;
        recLeftW = 513.667 + exW * (513.667 / 1155.761);
        recRightX = innerX + recLeftW + 27.2;
        recRightW = innerW - recLeftW - 27.2;

        double bodyTop = Rdv3Geom.RecKv0.Y + dy;
        double bodyBottom = p.Y + dy + p.H + exRec - 14.2;
        kvRowH = Math.Min(Rdv3Geom.KvH, Math.Max(16.0, (bodyBottom - bodyTop) / KvLabels.Length));
        for (int i = 0; i < KvLabels.Length; i++)
        {
            Put("rec.kv" + i.ToString(CultureInfo.InvariantCulture),
                innerX, bodyTop + i * kvRowH, recLeftW, kvRowH);
        }

        recBoxH = Math.Max(24.0, 74.4 + exRec / 2.0);
        double shift = recBoxH - 74.4;
        Put("rec.memoLabel", recRightX, Rdv3Geom.RecMemoLabel.Y + dy, recRightW, 17.0);
        Put("rec.memoBox", recRightX, Rdv3Geom.RecMemoBox.Y + dy, recRightW, recBoxH);
        Put("rec.remarkLabel", recRightX, Rdv3Geom.RecRemarkLabel.Y + dy + shift, recRightW, 17.0);
        Put("rec.remarkBox", recRightX, Rdv3Geom.RecRemarkBox.Y + dy + shift, recRightW, recBoxH);
    }

    private string[] RecTags()
    {
        if (rec == null) { return new string[] { "", "", "" }; }
        return new string[] {
            sRecStatus,
            recProc ? Rdv3Text.LabelProcessed : Rdv3Text.LabelUnprocessed,
            Rdv3Text.TagKey2 + " = " + sKey2 };
    }

    private void LayFoot(double cw, double ch, double errH)
    {
        double b = 0.7, stH = Rdv3Geom.Status.H;
        double top = ch - stH - b;
        Put("status", b, top, cw - 2 * b, stH);
        if (errH > 0) { Put("err", b, top - errH, cw - 2 * b, errH); }
        LayStatusFlow(cw, top, stH);
    }

    // The status bar is the one region the reference lays out by flow: segments
    // run left to right, the log/pid/clock cluster is right-aligned, and the
    // left cluster stops where the right one starts. Doing it here rather than
    // in the painter means every segment is in the dump and gets checked for
    // overlap like everything else. A segment that would not fit whole is left
    // out of the dictionary entirely -- it is never half drawn.
    private void LayStatusFlow(double cw, double top, double stH)
    {
        double pad = 13.6, gap = 10.2, band = 18.6;
        double ty = top + (stH - band) / 2.0;
        double sy = top + (stH - 12.0) / 2.0;
        string now = ClockText;
        double rx = cw - 0.7 - pad;
        double w = MW(now, fStatus);
        rx -= w; Put("status.clock", rx, ty, w, band);
        rx -= gap; Put("status.sep3", rx, sy, 1.0, 12.0); rx -= gap;
        if (sPid.Length > 0)
        {
            w = MW(sPid, fStatus); rx -= w; Put("status.pid", rx, ty, w, band);
            rx -= gap; Put("status.sep2", rx, sy, 1.0, 12.0); rx -= gap;
        }
        if (sLog.Length > 0)
        {
            w = MW(sLog, fStatus); rx -= w; Put("status.log", rx, ty, w, band);
        }

        double limit = rx - gap;
        double x = 0.7 + pad;
        Put("status.dot", x, top + (stH - 7.0) / 2.0, 7.0, 7.0);
        x += 7.0 + 6.0;
        string[] seg = { sState, sWatchLabel + " " + sNotepad,
            Rdv3Text.LabelLedger + " " + sLedger,
            Rdv3Text.LabelMergeMs + " " + sMerge, Rdv3Text.LabelSearchMs + " " + sSearch };
        string[] key = { "status.state", "status.notepad", "status.ledger", "status.merge", "status.search" };
        // a time that has not been measured yet is not a value: no segment
        bool[] show = { true, true, true, sMerge != Rdv3Text.NotYet, sSearch != Rdv3Text.NotYet };
        double[] wid = new double[seg.Length];
        double need = x;
        for (int i = 0; i < seg.Length; i++)
        {
            if (!show[i]) { continue; }
            wid[i] = (i >= 3)
                ? MW(SegLabel(i) + " ", fStatus) + MW(SegValue(i), fStatusB)
                : MW(seg[i], (i == 0) ? fStatusB : fStatus);
            need += wid[i] + gap + ((i < seg.Length - 1) ? (1.0 + gap) : 0.0);
        }
        // The one variable-length string in the bar is the notepad's window
        // TITLE, and it is the least load-bearing: the label stays either way.
        // So when the bar is tight it gives up its tail rather than letting a
        // whole segment (the search time) disappear off the end.
        if (need > limit && wid[1] > 100.0)
        {
            wid[1] = Math.Max(100.0, wid[1] - (need - limit));
        }
        for (int i = 0; i < seg.Length; i++)
        {
            if (!show[i]) { continue; }
            if (x + wid[i] > limit) { break; }
            Put(key[i], x, ty, wid[i], band);
            x += wid[i] + gap;
            if (i < seg.Length - 1)
            {
                if (x + 1.0 > limit) { break; }
                Put("status.sepL" + i.ToString(CultureInfo.InvariantCulture), x, sy, 1.0, 12.0);
                x += gap;
            }
        }
    }

    private string SegLabel(int i)
    {
        return (i == 3) ? Rdv3Text.LabelMergeMs : Rdv3Text.LabelSearchMs;
    }

    private string SegValue(int i)
    {
        return (i == 3) ? sMerge : sSearch;
    }

    private string ClockText
    {
        get { return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
    }

    private void PlaceControls()
    {
        Rectangle box = At("sum.input");
        int pad = PX(10.2);
        int th = Math.Max(txtKey.PreferredHeight, PX(16));
        txtKey.SetBounds(box.X + pad, box.Y + (box.Height - th) / 2,
            Math.Max(PX(20), box.Width - 2 * pad), th);
        btnSearch.Bounds = At("sum.btnSearch");
        btnClear.Bounds = At("sum.btnClear");
        btnProcessed.Bounds = At("sum.btnProcessed");
        btnRebind.Bounds = At("nav.btnRebind");
        btnSettings.Bounds = At("nav.btnSettings");
        table.Bounds = At("list.scroll");
        int bp = PX(10.2), bq = PX(6.8), hair = Rdv3Skin.Hair();
        Rectangle m = At("rec.memoBox");
        txtMemo.SetBounds(m.X + bp, m.Y + bq, Math.Max(PX(20), m.Width - 2 * bp),
            Math.Max(PX(14), m.Height - 2 * bq));
        Rectangle k = At("rec.remarkBox");
        txtRemark.SetBounds(k.X + bp, k.Y + bq, Math.Max(PX(20), k.Width - 2 * bp),
            Math.Max(PX(14), k.Height - 2 * bq));
        if (hair < 1) { hair = 1; }
        SetBox(txtMemo, txtMemo.Text);
        SetBox(txtRemark, txtRemark.Text);
    }

    // ---- painting ----------------------------------------------------------
    // every string is drawn INSIDE its measured rectangle, vertically centred
    // and cut with an ellipsis if the real value is longer than the box.
    // An element whose string does not fit its rectangle is the one failure the
    // eye forgives and the operator does not: the information is simply gone.
    // Every draw records it, and the acceptance dump reports the list.
    private readonly List<string> clipped = new List<string>();

    private void T(Graphics g, string k, string s, Font f, Color c)
    {
        Note(k, s, f, At(k));
        Rdv3Skin.DrawIn(g, s, f, c, At(k), false);
    }

    private void TR(Graphics g, string k, string s, Font f, Color c)
    {
        Note(k, s, f, At(k));
        Rdv3Skin.DrawIn(g, s, f, c, At(k), true);
    }

    // the notepad segment is deliberately elastic (see LayStatusFlow)
    private static readonly string[] Elastic = { "status.notepad" };

    private void Note(string k, string s, Font f, Rectangle r)
    {
        if (s == null || s.Length == 0) { return; }
        for (int i = 0; i < Elastic.Length; i++) { if (Elastic[i] == k) { return; } }
        if (Rdv3Skin.Measure(s, f).Width > r.Width && !clipped.Contains(k)) { clipped.Add(k); }
    }

    private void TT(Graphics g, string k, string s, Font f, Color c, double track)
    {
        Rectangle r = At(k);
        if (s != null && s.Length > 0 && Rdv3Skin.MeasureTracked(s, f, track) > r.Width
            && !clipped.Contains(k)) { clipped.Add(k); }
        int h = Rdv3Skin.HeightTracked(s, f);
        Rdv3Skin.DrawTracked(g, s, f, c, r.X, r.Y + (r.Height - h) / 2, track, r.Right);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        clipped.Clear();
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Bg)) { g.FillRectangle(b, ClientRectangle); }
        PaintNav(g);
        PaintSummary(g);
        PaintList(g);
        PaintRecord(g);
        if (sError.Length > 0) { PaintError(g); }
        PaintStatus(g);
        PaintFocus(g);
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, ClientRectangle);
    }

    private void PaintNav(Graphics g)
    {
        T(g, "nav.brand", Rdv3Text.AppTitle, fBrand, Rdv3Skin.Ink);
        Rdv3Skin.Tag(g, Rdv3Text.Method, fTag, Rdv3Skin.TagAccent, At("nav.tagMethod"), 255);
        Rdv3Skin.Tag(g, Rdv3Text.TagLedger, fTag, Rdv3Skin.TagOutline, At("nav.tagLedger"), 255);
        Rectangle r = At("nav.rule");
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, r.Width, Rdv3Skin.Hair());
    }

    // :focus-visible is a 2 px accent outline 2 px outside the control -- which
    // lands on the form, not on the control, so the form draws it. Like the
    // browser, it only appears for keyboard focus, not after a mouse click.
    private bool keyFocus;
    private bool inputHot;

    private void PaintFocus(Graphics g)
    {
        if (!keyFocus) { return; }
        Control c = ActiveControl;
        if (c == null || !c.Visible) { return; }
        int off = (c == txtKey) ? 0 : PX(2);
        Rectangle r = (c == txtKey) ? At("sum.input") : c.Bounds;
        r.Inflate(off, off);
        int t = PX(2);
        using (Pen pen = new Pen(Rdv3Skin.Accent, t))
        {
            pen.Alignment = PenAlignment.Inset;
            g.DrawRectangle(pen, r.X - t, r.Y - t, r.Width + 2 * t - 1, r.Height + 2 * t - 1);
        }
    }

    private void PaintSummary(Graphics g)
    {
        Rdv3Skin.Blueprint(g, At("sum.panel"));
        Figure(g, "sum.key", Rdv3Text.LabelKeyNow, sKey, sKeySub, Rdv3Skin.Ink, false);
        Div(g, "sum.status.div");
        Color sc = (sStatus.Length == 0) ? Rdv3Skin.N400 : (sStatusNg ? Rdv3Skin.Ink : Rdv3Skin.Accent700);
        Figure(g, "sum.status", Rdv3Text.LabelRepStatus, sStatus, sStatusSub, sc, false);
        Div(g, "sum.rows.div");
        Figure(g, "sum.rows", Rdv3Text.LabelLedgerRows, sRows, SavedSub, Rdv3Skin.Ink, true);

        T(g, "sum.field.label", sKeyLabel, fSub12, Rdv3Skin.N700);
        Rectangle box = At("sum.input");
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Surface)) { g.FillRectangle(b, box); }
        // .input / .input:hover (text 45%) / .input:focus-visible (accent)
        Color edge = txtKey.Focused ? Rdv3Skin.Accent
            : (inputHot ? Rdv3Skin.Mix(Rdv3Skin.Surface, 0.45) : Rdv3Skin.Divider);
        Rdv3Skin.Frame(g, edge, box);

        if (sInputError.Length > 0)
        {
            T(g, "sum.field.error", sInputError, fBoxLabel, Rdv3Skin.Danger);
        }

        Div(g, "sum.ident.div");
        Rdv3Skin.IconPerson(g, At("sum.ident.icon"), Rdv3Skin.Accent700);
        T(g, "sum.ident.name", sUser, fName, Rdv3Skin.Ink);
        T(g, "sum.ident.host", sHost, fName11, Rdv3Skin.N700);
        Rdv3Skin.Tag(g, sRole, fTag, Rdv3Skin.TagNeutral, At("sum.ident.role"), 255);
    }

    private void Div(Graphics g, string k)
    {
        Rectangle r = At(k);
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, Rdv3Skin.Hair(), r.Height);
    }

    private void Figure(Graphics g, string k, string label, string value, string sub, Color vc, bool rows)
    {
        TT(g, k + ".label", label, fLabel10, Rdv3Skin.Accent, LabelTrack);
        TT(g, k + ".value", value, fBig, vc, ValueTrack);
        if (rows && value.Length > 0)
        {
            Rectangle r = At(k + ".value");
            int w = Rdv3Skin.MeasureTracked(value, fBig, ValueTrack);
            int hb = Rdv3Skin.Measure(value, fBig).Height;
            int hu = Rdv3Skin.Measure(Rdv3Text.UnitRows, fBig13).Height;
            int ux = r.X + w + PX(4);
            if (ux + Rdv3Skin.Measure(Rdv3Text.UnitRows, fBig13).Width <= r.Right)
            {
                Rdv3Skin.Draw(g, Rdv3Text.UnitRows, fBig13, Rdv3Skin.Ink, ux,
                    r.Y + (r.Height + hb) / 2 - hu - PX(3));
            }
        }
        T(g, k + ".sub", sub, fSub12, Rdv3Skin.N700);
    }

    private void PaintList(Graphics g)
    {
        Rdv3Skin.Blueprint(g, At("list.panel"));
        T(g, "list.h4", Rdv3Text.PanelCand, fH4, Rdv3Skin.Ink);
        T(g, "list.hint", Rdv3Text.HintCand, fHint, Rdv3Skin.Corner);
        T(g, "list.verdict", sVerdict, fVerdict, Rdv3Skin.Accent700);
        Rdv3Skin.Tag(g, sCandTag, fTag, Rdv3Skin.TagNeutral, At("list.tag"), 255);
        Rectangle r = At("list.rule");
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, r.Width, Rdv3Skin.Hair());
    }

    private void PaintRecord(Graphics g)
    {
        Rdv3Skin.Blueprint(g, At("rec.panel"));
        T(g, "rec.h4", Rdv3Text.PanelRec, fH4, Rdv3Skin.Ink);
        T(g, "rec.hint", Rdv3Text.HintRec, fHint, Rdv3Skin.Corner);
        string[] tags = RecTags();
        int[] kinds = { StatusKind(sRecStatus),
            (rec != null && recProc) ? Rdv3Skin.TagAccent : Rdv3Skin.TagNeutral, Rdv3Skin.TagAccent };
        for (int i = 0; i < 3; i++)
        {
            string tk = "rec.tag" + i.ToString(CultureInfo.InvariantCulture);
            if (tags[i].Length == 0 || !rc.ContainsKey(tk)) { continue; }
            Rdv3Skin.Tag(g, tags[i], fTag, kinds[i], At(tk), 255);
        }
        Rectangle rr = At("rec.rule");
        Rdv3Skin.Line(g, Rdv3Skin.Divider, rr.X, rr.Y, rr.Width, Rdv3Skin.Hair());

        for (int i = 0; i < KvLabels.Length; i++)
        {
            Rectangle row = At("rec.kv" + i.ToString(CultureInfo.InvariantCulture));
            string v = "";
            bool none = false;
            if (rec != null)
            {
                v = (i == 6) ? (Get(rec, 7) + " \u30fb " + Get(rec, 8)).Trim() : Get(rec, KvCols[i]);
                if (v.Length == 0 || v == "\u30fb") { v = Rdv3Text.NoValue; none = true; }
            }
            int lw = Rdv3Skin.Measure(KvLabels[i], fKv).Width;
            Note("rec.kv" + i.ToString(CultureInfo.InvariantCulture),
                KvLabels[i] + v, fKvB, new Rectangle(row.X, row.Y, row.Width - PX(10.2), row.Height));
            Rdv3Skin.DrawIn(g, KvLabels[i], fKv, Rdv3Skin.N700,
                new Rectangle(row.X, row.Y, Math.Min(lw, row.Width), row.Height), false);
            int vx = row.X + lw + PX(10.2);
            Rdv3Skin.DrawIn(g, v, none ? fKv : fKvB, none ? Rdv3Skin.N400 : Rdv3Skin.Ink,
                new Rectangle(vx, row.Y, Math.Max(0, row.Right - vx), row.Height), true);
            if (i < KvLabels.Length - 1)
            {
                Rdv3Skin.Line(g, Rdv3Skin.RowLine, row.X, row.Bottom - Rdv3Skin.Hair(),
                    row.Width, Rdv3Skin.Hair());
            }
        }

        BoxFrame(g, "rec.memoLabel", "rec.memoBox", Rdv3Text.LabelMemo);
        BoxFrame(g, "rec.remarkLabel", "rec.remarkBox", Rdv3Text.LabelRemark);
    }

    private void BoxFrame(Graphics g, string kLabel, string kBox, string label)
    {
        T(g, kLabel, label, fBoxLabel, Rdv3Skin.N700);
        Rectangle box = At(kBox);
        using (SolidBrush b = new SolidBrush(Rdv3Skin.N100)) { g.FillRectangle(b, box); }
        Rdv3Skin.Frame(g, Rdv3Skin.Divider, box);
    }

    private static int StatusKind(string st)
    {
        if (st == "DONE") { return Rdv3Skin.TagAccent; }
        if (st == "HOLD") { return Rdv3Skin.TagOutline; }
        return Rdv3Skin.TagNeutral;
    }

    private static string Get(string[] f, int i)
    {
        return (f != null && i >= 0 && i < f.Length && f[i] != null) ? f[i] : "";
    }

    // a value that IS on screen but empty in the ledger
    private static string Val(string s)
    {
        return (s == null || s.Length == 0) ? Rdv3Text.NoValue : s;
    }

    private void PaintError(Graphics g)
    {
        Rectangle r = At("err");
        using (SolidBrush b = new SolidBrush(Rdv3Skin.N100)) { g.FillRectangle(b, r); }
        Rdv3Skin.Line(g, Rdv3Skin.Divider, r.X, r.Y, r.Width, Rdv3Skin.Hair());
        Size z = Rdv3Skin.TagSize(Rdv3Text.VerdictErr, fTag);
        Rdv3Skin.Tag(g, Rdv3Text.VerdictErr, fTag, Rdv3Skin.TagOutline,
            new Rectangle(r.X + PX(13.6), r.Y + (r.Height - z.Height) / 2, z.Width, z.Height), 255);
        int tx = r.X + PX(13.6) + z.Width + PX(10.2);
        Rdv3Skin.DrawIn(g, sError, fVerdict, Rdv3Skin.Accent800,
            new Rectangle(tx, r.Y, Math.Max(0, r.Right - PX(13.6) - tx), r.Height), false);
    }

    private void PaintStatus(Graphics g)
    {
        Rectangle r = At("status");
        using (SolidBrush b = new SolidBrush(Rdv3Skin.Accent900)) { g.FillRectangle(b, r); }
        Rectangle dot = At("status.dot");
        Rdv3Skin.Line(g, Rdv3Skin.Accent300, dot.X, dot.Y, dot.Width, dot.Height);
        StatusSeg(g, "status.state", sState, fStatusB, Rdv3Skin.Bg);
        StatusSeg(g, "status.notepad", sWatchLabel + " " + sNotepad, fStatus, Rdv3Skin.Bg);
        StatusSeg(g, "status.ledger", Rdv3Text.LabelLedger + " " + sLedger, fStatus, Rdv3Skin.Bg);
        StatusPair(g, "status.merge", Rdv3Text.LabelMergeMs, sMerge);
        StatusPair(g, "status.search", Rdv3Text.LabelSearchMs, sSearch);
        StatusSeg(g, "status.log", sLog, fStatus, Color.FromArgb(191, Rdv3Skin.Bg));
        StatusSeg(g, "status.pid", sPid, fStatus, Rdv3Skin.Bg);
        StatusSeg(g, "status.clock", ClockText, fStatus, Rdv3Skin.Bg);
        Color sepc = Color.FromArgb(76, Rdv3Skin.Bg);
        foreach (KeyValuePair<string, Rectangle> kv in rc)
        {
            if (kv.Key.StartsWith("status.sep"))
            {
                Rdv3Skin.Line(g, sepc, kv.Value.X, kv.Value.Y, Rdv3Skin.Hair(), kv.Value.Height);
            }
        }
    }

    private void StatusSeg(Graphics g, string k, string s, Font f, Color c)
    {
        if (!rc.ContainsKey(k)) { return; }
        Rdv3Skin.DrawIn(g, s, f, c, At(k), false);
    }

    private void StatusPair(Graphics g, string k, string label, string value)
    {
        if (!rc.ContainsKey(k)) { return; }
        Rectangle r = At(k);
        int lw = Rdv3Skin.Measure(label + " ", fStatus).Width;
        Rdv3Skin.DrawIn(g, label + " ", fStatus, Rdv3Skin.Bg, r, false);
        Rdv3Skin.DrawIn(g, value, fStatusB, Rdv3Skin.Bg,
            new Rectangle(r.X + lw, r.Y, Math.Max(0, r.Width - lw), r.Height), false);
    }

    // the status bar's segments are laid out from live strings, so changing one
    // has to re-run that flow -- but only that flow, once a second at most
    private void RelayoutStatus()
    {
        double cw = CX(ClientSize.Width), ch = CX(ClientSize.Height);
        if (cw < 100 || ch < 100) { return; }
        List<string> drop = new List<string>();
        foreach (KeyValuePair<string, Rectangle> kv in rc)
        {
            if (kv.Key.StartsWith("status")) { drop.Add(kv.Key); }
        }
        for (int i = 0; i < drop.Count; i++) { rc.Remove(drop[i]); }
        LayFoot(cw, ch, (sError.Length > 0) ? Rdv3Geom.ErrRow.H : 0.0);
        rStatus = At("status");
        Invalidate(rStatus);
    }

    // ---- the acceptance dump ----------------------------------------------
    // the same dictionary the painter reads, in CSS px, so the screen can be
    // compared with the reference table element by element
    public string GeometryDump()
    {
        int[] cx = table.ColumnEdges();
        Rectangle tb = table.Bounds;
        for (int i = 0; i + 1 < cx.Length; i++)
        {
            rc["list.col" + i.ToString(CultureInfo.InvariantCulture)] =
                new Rectangle(tb.X + cx[i], tb.Y, cx[i + 1] - cx[i], table.HeaderH);
        }
        List<string> keys = new List<string>(rc.Keys);
        keys.Sort(StringComparer.Ordinal);
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        sb.Append("{\"scale\":").Append(Rdv3Skin.Scale.ToString("0.###", CultureInfo.InvariantCulture));
        sb.Append(",\"client\":[").Append(CssStr(ClientSize.Width)).Append(",")
          .Append(CssStr(ClientSize.Height)).Append("],\"el\":{");
        for (int i = 0; i < keys.Count; i++)
        {
            Rectangle r = rc[keys[i]];
            if (i > 0) { sb.Append(","); }
            sb.Append("\"").Append(keys[i]).Append("\":[")
              .Append(CssStr(r.X)).Append(",").Append(CssStr(r.Y)).Append(",")
              .Append(CssStr(r.Width)).Append(",").Append(CssStr(r.Height)).Append("]");
        }
        sb.Append("},\"clipped\":[");
        for (int i = 0; i < clipped.Count; i++)
        {
            if (i > 0) { sb.Append(","); }
            sb.Append("\"").Append(clipped[i]).Append("\"");
        }
        sb.Append("]}");
        return sb.ToString();
    }

    private static string CssStr(int devicePx)
    {
        return (devicePx / (double)Rdv3Skin.Scale).ToString("0.#", CultureInfo.InvariantCulture);
    }

    // The real scale is only knowable once there IS a window: a process that
    // becomes dpi-aware in its own constructor still reports 96 from a Graphics
    // made before the handle exists. So the skin is scaled here, the fonts are
    // rebuilt, and the window is sized once for real. Creating the handle in
    // the constructor is what used to leave half the controls unscaled.
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
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
        SetScale(s);
    }

    // also the entry point the headless acceptance harness uses: it renders
    // this form at each Windows scaling factor without ever showing a window
    public void SetScale(float s)
    {
        Rdv3Skin.Scale = s;
        MakeFonts();
        ApplyFonts();
        Rectangle wa = Screen.FromControl(this).WorkingArea;
        Size nc = new Size(Math.Max(0, Width - ClientSize.Width), Math.Max(0, Height - ClientSize.Height));
        // the reference refuses to go below its design width (min-width 1240),
        // so the window does too: it may only grow, and growth goes to the two
        // panels. That removes the whole class of "shrunk until it overlaps".
        int minW = Math.Min(PX(Rdv3Geom.CardW) + nc.Width, Math.Max(400, wa.Width));
        int minH = Math.Min(PX(Rdv3Geom.CardH) + nc.Height, Math.Max(300, wa.Height));
        MinimumSize = new Size(minW, minH);
        if (Width < minW || Height < minH) { Size = new Size(Math.Max(Width, minW), Math.Max(Height, minH)); }
        Layout1();
    }

    private void ApplyFonts()
    {
        txtKey.Font = Rdv3Skin.F(14, FontStyle.Regular);
        btnSearch.Font = fBtn; btnClear.Font = fBtn; btnProcessed.Font = fBtn;
        btnRebind.Font = fBtnSm; btnSettings.Font = fBtnSm;
        txtMemo.Font = fBox; txtRemark.Font = fBox;
        table.MakeFonts();
    }


    // ---- the public surface (all marshalled onto the UI thread) ------------
    // Everything the worker sets before there is a window used to be dropped on
    // the floor: SetIdentity runs during start-up, so the PID and the log name
    // never reached the status bar (found by looking at the real window). They
    // are queued instead and replayed the moment the handle exists.
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

    public void SetIdentity(string user, string host, string role, string pid, string log)
    {
        Ui(delegate
        {
            if (user != null && user.Length > 0) { sUser = user; }
            if (host != null && host.Length > 0) { sHost = host; }
            if (role != null && role.Length > 0) { sRole = role; }
            sPid = (pid == null) ? "" : pid;
            sLog = (log == null) ? "" : log;
            Layout1();
        });
    }

    public void SetState(string text, int tone)
    {
        Ui(delegate { sState = text; RelayoutStatus(); });
    }

    public void SetNotepad(string detail)
    {
        Ui(delegate { sNotepad = detail; RelayoutStatus(); });
    }

    // the watch target's name, from ReaderDataViewer.json
    public void SetWatchLabel(string label)
    {
        Ui(delegate
        {
            sWatchLabel = (label == null || label.Length == 0) ? Rdv3Text.LabelNotepad : label;
            RelayoutStatus();
        });
    }

    // the key rule, so the input label and the box agree with the settings
    public void SetKeyRule(int length)
    {
        Ui(delegate
        {
            txtKey.MaxLength = length;
            sKeyLabel = Rdv3Text.LabelSearchBox.Replace("{n}",
                length.ToString(CultureInfo.InvariantCulture));
            Layout1();
        });
    }

    public void SetLedgerInfo(string text)
    {
        Ui(delegate { sLedger = text; RelayoutStatus(); });
    }

    public void SetLedgerSummary(string rows, string saved)
    {
        Ui(delegate { sRows = rows; sSaved = saved; Layout1(); });
    }

    public void SetMergeMs(double ms)
    {
        Ui(delegate { sMerge = Rdv3Log.F(ms) + Rdv3Text.MsUnit; RelayoutStatus(); });
    }

    public void SetSearchMs(double ms)
    {
        Ui(delegate { sSearch = Rdv3Log.F(ms) + Rdv3Text.MsUnit; RelayoutStatus(); });
    }

    // the answer to a bad key: under the box, in the gap that is already there
    public void SetInputError(string text)
    {
        Ui(delegate
        {
            string t = (text == null) ? "" : text;
            if (t == sInputError) { return; }
            sInputError = t;
            Layout1();
        });
    }

    public void SetError(string text)
    {
        Ui(delegate
        {
            bool was = sError.Length > 0;
            sError = (text == null) ? "" : text;
            Layout1();
        });
    }

    // the reference has no overlay: the update check is shown in the status bar
    public void ShowOverlay(string baseText) { SetState(baseText, 2); }
    public void HideOverlay() { }

    public void EnableOps(bool on)
    {
        Ui(delegate
        {
            txtKey.Enabled = on;
            btnSearch.Enabled = on;
            btnClear.Enabled = on;
            btnProcessed.Enabled = on;
            table.Enabled = on;
        });
    }

    public void EnableProcessed(bool on)
    {
        Ui(delegate { btnProcessed.Enabled = on; });
    }

    // ---- results -----------------------------------------------------------
    public void ShowCandidates(string key, List<Rdv3CandRow> rows, int totalHits, int shown)
    {
        Ui(delegate
        {
            sKey = key;
            table.SetRows(rows, -1);
            rec = null;
            recProc = false;
            sKey2 = "";
            sRecStatus = "";
            sStatus = "";
            sStatusNg = false;
            SetBox(txtMemo, Rdv3Text.PickToSee);
            SetBox(txtRemark, Rdv3Text.PickToSee);
            if (totalHits <= 0)
            {
                sVerdict = Rdv3Text.VerdictNone;
                sCandTag = Rdv3Text.CandCount.Replace("{n}", "0");
                sKeySub = Rdv3Text.SubNone;
                sStatusSub = "";
            }
            else
            {
                sVerdict = (shown < totalHits)
                    ? Rdv3Text.VerdictManyCut.Replace("{n}", totalHits.ToString(CultureInfo.InvariantCulture))
                                             .Replace("{m}", shown.ToString(CultureInfo.InvariantCulture))
                    : Rdv3Text.VerdictMany.Replace("{n}", totalHits.ToString(CultureInfo.InvariantCulture));
                sCandTag = Rdv3Text.CandCount.Replace("{n}", totalHits.ToString(CultureInfo.InvariantCulture));
                sKeySub = Rdv3Text.SubUnselected.Replace("{n}", totalHits.ToString(CultureInfo.InvariantCulture));
                sStatusSub = Rdv3Text.SubMulti.Replace("{n}", totalHits.ToString(CultureInfo.InvariantCulture));
            }
            Layout1();
        });
    }

    public void SelectCandidate(int index, string line, bool processed, int totalHits)
    {
        Ui(delegate
        {
            table.Select(index);
            rec = (line == null) ? null : Rdv3Ledger.SplitLine(line);
            recProc = processed;
            if (rec != null)
            {
                sKey2 = Val(Get(rec, 1));
                sRecStatus = Val(Get(rec, 16));
                bool ng = (sRecStatus == "HOLD" || sRecStatus == "VOID");
                sStatus = ng ? "NG" : "OK";
                sStatusNg = ng;
                sStatusSub = sRecStatus + " ・ " + (processed ? Rdv3Text.LabelProcessed : Rdv3Text.LabelUnprocessed);
                SetBox(txtMemo, Get(rec, 18));
                SetBox(txtRemark, Get(rec, 27));
                sKeySub = (table.Count <= 1)
                    ? Rdv3Text.SubOne.Replace("{key2}", sKey2)
                    : Rdv3Text.SubPicked.Replace("{i}", (index + 1).ToString(CultureInfo.InvariantCulture))
                                        .Replace("{n}", table.Count.ToString(CultureInfo.InvariantCulture))
                                        .Replace("{key2}", sKey2);
                sVerdict = (table.Count <= 1)
                    ? Rdv3Text.VerdictOne.Replace("{key2}", sKey2)
                    : Rdv3Text.VerdictPicked.Replace("{key2}", sKey2)
                        .Replace("{n}", totalHits.ToString(CultureInfo.InvariantCulture))
                        .Replace("{i}", (index + 1).ToString(CultureInfo.InvariantCulture));
            }
            Layout1();
        });
    }

    // the idle screen, without the UI-thread guard (the constructor runs
    // before there is a handle to marshal onto)
    private void ResetState()
    {
        sKey = "";
        sKeySub = Rdv3Text.SubIdle;
        sStatus = "";
        sStatusSub = "";
        sStatusNg = false;
        sVerdict = "";
        sCandTag = Rdv3Text.CandCount.Replace("{n}", "0");
        sKey2 = "";
        sRecStatus = "";
        sInputError = "";
        rec = null;
        recProc = false;
        table.SetRows(new List<Rdv3CandRow>(), -1);
        SetBox(txtMemo, Rdv3Text.PickToSee);
        SetBox(txtRemark, Rdv3Text.PickToSee);
        txtKey.Text = "";
    }

    public void ClearResult()
    {
        Ui(delegate { ResetState(); Layout1(); });
    }

    public void ShowProcessedState(string processedText)
    {
        Ui(delegate
        {
            // the text carries the label; the screen shows it as a tag plus
            // the "(saving...)" suffix in the status line
            procSuffix = (processedText != null && processedText.EndsWith(Rdv3Text.SavingSuffix)) ? Rdv3Text.SavingSuffix : "";
            if (processedText != null && processedText.Length > 0)
            {
                recProc = processedText.Contains(Rdv3Ledger.ProcessedTrue);
                if (recProc && table.Selected >= 0) { table.MarkProcessed(table.Selected); }
                if (rec != null)
                {
                    sStatusSub = sRecStatus + " ・ " + (recProc ? Rdv3Text.LabelProcessed : Rdv3Text.LabelUnprocessed) + procSuffix;
                }
            }
            Layout1();
        });
    }

    // what the screen actually resolved to, for the log
    // what the screen actually resolved to, for the log
    public string Diag
    {
        get
        {
            Rectangle wa = Screen.FromControl(this).WorkingArea;
            return "scale=" + Rdv3Skin.Scale.ToString("0.00", CultureInfo.InvariantCulture)
                + " client=" + ClientSize.Width + "x" + ClientSize.Height
                + " min=" + MinimumSize.Width + "x" + MinimumSize.Height
                + " work=" + wa.Width + "x" + wa.Height
                + " logical=" + Math.Round(CX(ClientSize.Width)) + "x" + Math.Round(CX(ClientSize.Height))
                + " tag=[" + sCandTag + "|" + sKeySub + "]"
                + " grow=[w " + Math.Round(exW) + " list " + Math.Round(exList)
                + " rec " + Math.Round(exRec) + "]";
        }
    }

    public string KeyText { get { return txtKey.Text.Trim(); } }

    public void SetKeyText(string s)
    {
        Ui(delegate { txtKey.Text = (s == null) ? "" : s; });
    }

    public bool Ask(string title, string body)
    {
        return MessageBox.Show(this, body, title, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            == DialogResult.Yes;
    }

    public void Tell(string title, string body)
    {
        Ui(delegate { MessageBox.Show(this, body, title, MessageBoxButtons.OK, MessageBoxIcon.Information); });
    }
}
