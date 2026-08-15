// ============================================================================
// Rdv3Geom.cs -- the reference geometry, measured from the artifact.
//
// Every rectangle below was read out of "Reader Data Viewer_ver3.html" rendered
// in a browser at its design width (card 1240 x 689 CSS px) and is stored
// CARD-RELATIVE, in CSS px. The screen is laid out FROM THIS TABLE: no
// hand-rolled flex, no "drop it if it does not fit". If something does not fit
// at another font, that is reported as a difference -- it is never silently
// re-arranged.
//
// Anchoring, and only this, is allowed to move a rectangle when the window is
// larger than the design:
//   Fixed        stays where the reference put it
//   Right        keeps its distance to the right edge
//   StretchW     keeps its left edge, grows with the width
// Vertical growth is shared by the two list panels; everything below them
// shifts by the same amount, which is what the reference's grid does.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Drawing;

public static class Rdv3Geom
{
    public const double CardW = 1240.0;
    public const double CardH = 689.0;

    public struct R
    {
        public double X, Y, W, H;
        public R(double x, double y, double w, double h) { X = x; Y = y; W = w; H = h; }
    }

    // ---- nav ---------------------------------------------------------------
    public static readonly R Nav = new R(0.7, 0.7, 1238.7, 43.7);
    public static readonly R NavBrand = new R(14.3, 7.5, 138.6, 29.4);
    public static readonly R NavTagMethod = new R(163.0, 10.7, 112.9, 23.0);
    public static readonly R NavTagLedger = new R(286.1, 10.0, 66.2, 24.4);
    public static readonly R NavBtnRebind = new R(1099.0, 8.7, 126.7, 26.9);   // right-anchored

    // ---- summary -----------------------------------------------------------
    public static readonly R SumPanel = new R(14.3, 58.0, 1211.5, 97.2);
    public static readonly R SumKeyLabel = new R(28.5, 68.8, 145.5, 15.5);
    public static readonly R SumKeyValue = new R(28.5, 86.3, 145.5, 37.4);
    public static readonly R SumKeySub = new R(28.5, 125.7, 145.5, 18.6);
    public static readonly R SumStatusDiv = new R(187.6, 64.8, 1.0, 84.0);
    public static readonly R SumStatusLabel = new R(201.9, 68.8, 119.6, 15.5);
    public static readonly R SumStatusValue = new R(201.9, 86.3, 119.6, 37.4);
    public static readonly R SumStatusSub = new R(201.9, 125.7, 119.6, 18.6);
    public static readonly R SumRowsDiv = new R(335.0, 64.8, 1.0, 84.0);
    public static readonly R SumRowsLabel = new R(349.3, 68.8, 123.5, 15.5);
    public static readonly R SumRowsValue = new R(349.3, 86.3, 123.5, 37.4);
    public static readonly R SumRowsSub = new R(349.3, 125.7, 123.5, 18.6);
    public static readonly R SumFieldLabel = new R(633.0, 68.8, 110.0, 18.6);   // right-anchored group
    public static readonly R SumInput = new R(633.0, 92.4, 110.0, 36.0);
    public static readonly R SumBtnSearch = new R(749.8, 96.7, 73.8, 31.7);
    public static readonly R SumBtnClear = new R(830.4, 96.7, 85.8, 31.7);
    public static readonly R SumBtnProcessed = new R(923.0, 96.7, 100.5, 31.7);
    public static readonly R SumIdentDiv = new R(1037.1, 64.8, 1.0, 84.0);
    public static readonly R SumIdent = new R(1051.4, 68.8, 160.1, 75.5);

    // ---- candidate list ----------------------------------------------------
    public static readonly R ListPanel = new R(14.3, 172.2, 1211.5, 195.7);
    public static readonly R ListH4 = new R(28.5, 183.0, 70.9, 20.2);
    public static readonly R ListHint = new R(109.6, 187.0, 200.3, 18.6);
    public static readonly R ListVerdict = new R(320.0, 186.0, 700.0, 20.0);
    public static readonly R ListTag = new R(1151.8, 184.7, 59.7, 23.0);       // right-anchored
    public static readonly R ListHeadRule = new R(14.9, 214.6, 1210.1, 1.0);
    public static readonly R ListScroll = new R(14.9, 215.2, 1210.1, 152.0);
    public const double ThH = 31.0;
    public const double RowH = 38.6;
    public static readonly double[] Col = { 44.0, 94.7, 69.6, 103.8, 88.8, 57.8, 83.3, 85.9, 107.4, 80.5, 379.0 };

    // ---- record ------------------------------------------------------------
    public static readonly R RecPanel = new R(14.3, 384.9, 1211.5, 266.8);
    public static readonly R RecH4 = new R(28.5, 395.7, 85.3, 20.2);
    public static readonly R RecHint = new R(123.0, 399.0, 430.0, 18.6);
    public static readonly R RecTag0 = new R(1071.2, 395.7, 27.8, 23.0);       // right-anchored
    public static readonly R RecTag1 = new R(1109.2, 395.7, 27.8, 23.0);
    public static readonly R RecTag2 = new R(1147.1, 395.7, 64.4, 23.0);
    public static readonly R RecHeadRule = new R(14.9, 425.6, 1210.1, 1.0);
    public static readonly R RecKv0 = new R(28.5, 436.4, 513.7, 28.8);
    public const double KvH = 28.8;
    public static readonly R RecMemoLabel = new R(569.4, 436.4, 642.1, 17.0);
    public static readonly R RecMemoBox = new R(569.4, 457.5, 642.1, 74.4);
    public static readonly R RecRemarkLabel = new R(569.4, 542.0, 642.1, 17.0);
    public static readonly R RecRemarkBox = new R(569.4, 563.1, 642.1, 74.4);

    // ---- error row + status bar -------------------------------------------
    public static readonly R ErrRow = new R(0.7, 651.0, 1238.7, 38.0);         // only when shown
    public static readonly R Status = new R(0.7, 659.7, 1238.7, 28.6);
}
