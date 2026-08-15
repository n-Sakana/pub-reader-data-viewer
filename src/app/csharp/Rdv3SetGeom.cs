// ============================================================================
// Rdv3SetGeom.cs -- the settings modal's reference geometry, measured.
//
// Same method as Rdv3Geom.cs, and the same promise: every rectangle below was
// read out of "Reader Data Viewer_ver4.html" rendered in a browser, with the
// settings modal open, and is stored MODAL-RELATIVE in CSS px. The dialog is
// laid out FROM THIS TABLE.
//
// The modal is 1060 x 764 and it is FULL: the left column runs list -> three
// buttons -> pick -> hint down to the last pixel of the body, and the right
// column runs name -> window -> field -> read/scope -> path down to the same
// line. There is no slack to give away, so the dialog does not stretch: it
// keeps the reference's proportions and scales as a whole. If the work area
// cannot hold the design at the current Windows scaling, the WHOLE dialog is
// scaled down to fit (Rdv3SettingsForm.FitScale) rather than re-arranged.
//
// Reading the table:
//   x, y are the modal-relative top-left corner; w, h the size.
//   A ".rule" is a hairline (the reference's 0.67 px border) and is stored at
//   the edge it sits on, one CSS px tall.
// ============================================================================

using System;

public static class Rdv3SetGeom
{
    public const double ModalW = 1060.0;
    public const double ModalH = 764.0;

    public struct R
    {
        public double X, Y, W, H;
        public R(double x, double y, double w, double h) { X = x; Y = y; W = w; H = h; }
        public double R2 { get { return X + W; } }
        public double B { get { return Y + H; } }
    }

    // ---- chrome ------------------------------------------------------------
    public static readonly R Head = new R(0.7, 0.7, 1058.7, 60.0);
    public static readonly R HeadTitle = new R(21.1, 14.1, 42.8, 32.5);          // 21 px / 600
    public static readonly R HeadSub = new R(77.5, 22.6, 162.8, 15.5);           // 10 px / 400
    public static readonly R HeadClose = new R(1006.6, 18.2, 32.3, 24.3);        // ghost
    public static readonly R HeadCloseIcon = new R(1015.3, 22.8, 15.0, 15.0);
    public static readonly R HeadRule = new R(0.7, 60.0, 1058.7, 1.0);

    public static readonly R Tabs = new R(0.7, 60.7, 1058.7, 46.0);
    public static readonly R Tab0 = new R(21.1, 60.7, 69.9, 45.3);               // 16 px / 600
    public const double TabPad = 2.0;                                            // padding: 0 2px
    public const double TabGap = 20.4;
    public static readonly R TabLabel0 = new R(23.1, 70.9, 65.9, 24.8);
    public static readonly R TabMark = new R(21.1, 104.0, 69.9, 2.0);            // accent underline
    public static readonly R TabMeta = new R(837.8, 75.6, 201.1, 15.5);          // 10 px, uppercase
    public static readonly R TabRule = new R(0.7, 106.0, 1058.7, 1.0);

    public static readonly R Body = new R(0.7, 106.7, 1058.7, 596.7);
    public static readonly R Content = new R(21.1, 127.1, 1017.9, 558.0);

    public static readonly R Foot = new R(0.7, 703.3, 1058.7, 60.0);             // N100
    public static readonly R FootRule = new R(0.7, 703.3, 1058.7, 1.0);
    public static readonly R FootNote = new R(21.1, 725.1, 282.6, 17.0);         // 11 px / 400
    public static readonly R FootSave = new R(768.7, 717.8, 130.0, 31.7);        // primary
    public static readonly R FootCancel = new R(908.9, 717.8, 130.0, 31.7);

    // ---- page 0: the watch targets -----------------------------------------
    public static readonly R LeftLabel = new R(21.1, 128.4, 46.5, 15.5);         // 10 px accent
    public static readonly R LeftCount = new R(302.1, 127.1, 19.0, 17.0);        // right aligned
    public static readonly R List = new R(21.1, 154.1, 300.0, 384.0);
    public const double CardH = 58.9;
    public const double CardPadX = 14.0;
    public const double CardNameDy = 10.0;                                       // 13 px / 700
    public const double CardTagDy = 12.3;                                        // 10 px tag
    public const double CardSumDy = 32.5;                                        // 11 px / 400
    public const double CardNameH = 20.1;
    public const double CardTagH = 15.5;
    public const double CardSumH = 13.3;
    public const double CardGap = 8.0;                                           // name -> tag

    public static readonly R BtnAdd = new R(21.1, 551.7, 94.7, 28.9);
    public const double BtnGap = 7.95;
    public static readonly R BtnPick = new R(21.1, 608.5, 300.0, 30.5);
    public static readonly R PickHint = new R(21.1, 649.0, 300.0, 33.0);         // 11 px, 2 lines

    public const double RightX = 348.3;
    public const double RightW = 690.7;
    public const double ColW = 335.1;                                            // (RightW - 20.4) / 2
    public const double ColGap = 20.4;
    public const double Col2X = 703.8;

    public static readonly R NameLabel = new R(348.3, 127.1, 621.3, 17.0);
    public static readonly R NameBox = new R(348.3, 150.1, 621.3, 36.0);
    public static readonly R EnabledChk = new R(989.9, 160.5, 15.0, 15.0);
    public static readonly R EnabledLbl = new R(1012.9, 158.0, 26.0, 20.1);

    public static readonly R SecWindow = new R(348.3, 206.5, 690.7, 22.2);       // 10 px accent
    public const double SecRuleDy = 21.5;                                        // hairline inside
    public static readonly R WinIdLabel = new R(348.3, 242.3, 335.1, 17.0);
    public static readonly R WinIdBox = new R(348.3, 265.3, 335.1, 36.0);
    public static readonly R WinLikeLabel = new R(348.3, 314.9, 335.1, 17.0);
    public static readonly R WinLikeBox = new R(348.3, 337.9, 335.1, 36.0);

    public static readonly R SecField = new R(348.3, 394.3, 690.7, 22.2);
    public static readonly R FldIdLabel = new R(348.3, 430.1, 335.1, 17.0);
    public static readonly R FldIdBox = new R(348.3, 453.1, 335.1, 36.0);
    public static readonly R FldTypesLabel = new R(348.3, 502.7, 335.1, 17.0);
    public static readonly R FldTypesBox = new R(348.3, 525.8, 335.1, 36.0);
    public static readonly R FldIndexLabel = new R(703.8, 502.7, 70.0, 17.0);
    public static readonly R FldIndexBox = new R(703.8, 525.8, 70.0, 36.0);
    public static readonly R VpChk = new R(787.4, 536.2, 15.0, 15.0);
    public static readonly R VpLbl = new R(810.4, 533.6, 101.4, 20.1);

    public static readonly R ReadLabel = new R(348.3, 575.4, 335.1, 17.0);
    public static readonly R ReadSeg = new R(348.3, 598.4, 210.0, 28.0);
    public static readonly R ScopeLabel = new R(703.8, 575.4, 335.1, 17.0);
    public static readonly R ScopeSeg = new R(703.8, 598.4, 210.0, 28.0);

    public static readonly R StepRule = new R(348.3, 646.8, 690.7, 1.0);
    public static readonly R StepLabel = new R(348.3, 664.5, 39.3, 17.0);
    public static readonly R StepNote = new R(397.7, 664.5, 641.3, 17.0);
    public const double StepNoteGap = 10.1;
    public static readonly R StepDel = new R(993.6, 662.2, 45.3, 21.7);          // ghost, 12 px

    // ---- page 1: how it behaves --------------------------------------------
    public const double OneColX = 21.1;
    public const double OneColW = 620.0;

    public static readonly R SecKey = new R(21.1, 127.1, 620.0, 22.2);
    public static readonly R KeyLenLabel = new R(21.1, 162.8, 90.0, 17.0);
    public static readonly R KeyLenBox = new R(21.1, 185.9, 90.0, 36.0);
    public static readonly R KeyLenUnit = new R(124.7, 195.8, 11.0, 26.0);
    public static readonly R DigitsChk = new R(149.3, 196.3, 15.0, 15.0);
    public static readonly R DigitsLbl = new R(172.3, 193.7, 48.1, 20.1);
    public static readonly R KeyNote = new R(21.1, 231.9, 620.0, 17.6);

    public static readonly R SecWatch = new R(21.1, 276.7, 620.0, 22.2);
    public static readonly R PollLabel = new R(21.1, 312.4, 110.0, 17.0);
    public static readonly R PollBox = new R(21.1, 335.5, 110.0, 36.0);
    public static readonly R StableLabel = new R(158.3, 312.4, 131.2, 17.0);
    public static readonly R StableBox = new R(158.3, 335.5, 110.0, 36.0);
    public static readonly R RebindLabel = new R(21.1, 385.1, 110.0, 17.0);
    public static readonly R RebindBox = new R(21.1, 408.1, 110.0, 36.0);
    public static readonly R CandLabel = new R(158.3, 385.1, 110.0, 17.0);
    public static readonly R CandBox = new R(158.3, 408.1, 110.0, 36.0);
    public static readonly R CandUnit = new R(278.5, 418.1, 11.0, 26.0);
    public static readonly R WatchNote = new R(21.1, 454.1, 620.0, 17.6);
    public static readonly R FocusChk = new R(21.1, 494.7, 15.0, 15.0);
    public static readonly R FocusLbl = new R(44.1, 492.1, 118.3, 20.1);

    // ---- page 2: where the files are ---------------------------------------
    public static readonly R SecFiles = new R(21.1, 127.1, 620.0, 22.2);
    public static readonly R DataLabel = new R(21.1, 162.8, 620.0, 17.0);
    public static readonly R DataBox = new R(21.1, 185.9, 620.0, 36.0);
    public static readonly R LedgerLabel = new R(21.1, 235.5, 620.0, 17.0);
    public static readonly R LedgerBox = new R(21.1, 258.5, 620.0, 36.0);
    public static readonly R LogLabel = new R(21.1, 308.1, 620.0, 17.0);
    public static readonly R LogBox = new R(21.1, 331.1, 620.0, 36.0);
    public static readonly R FilesNote = new R(21.1, 380.7, 620.0, 17.6);

    // ---- the picker panel --------------------------------------------------
    public const double PickW = 460.0;
    public const double PickH = 292.8;
    public static readonly R PkTitle = new R(14.3, 14.9, 95.8, 20.2);            // 18 px / 600
    public static readonly R PkTag = new R(395.1, 14.3, 50.6, 21.5);             // accent tag
    public static readonly R PkHow = new R(14.3, 41.8, 431.5, 18.6);             // 12 px / 700
    public static readonly R PkRule = new R(14.3, 70.6, 431.5, 1.0);
    public static readonly R PkKey0 = new R(14.3, 82.1, 120.0, 17.0);            // 11 px label
    public static readonly R PkVal0 = new R(134.3, 81.4, 311.5, 18.6);           // 12 px value
    public const double PkRowH = 27.6;
    public static readonly R PkEsc = new R(14.3, 256.5, 50.7, 17.0);
    public static readonly R PkUse = new R(234.7, 251.6, 142.3, 26.9);           // primary
    public static readonly R PkClose = new R(383.8, 251.6, 61.9, 26.9);
    public const double PkBtnGap = 6.8;
    public static readonly R PkFootRight = new R(445.7, 251.6, 0.0, 26.9);       // buttons end here
}
