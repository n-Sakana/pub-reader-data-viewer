// ============================================================================
// ZipWorker.cs -- ZipBench の C# ワーカー (UI Automation の PROVIDER 側)
//
// 役割分担
//   Excel VBA  = UIA CLIENT   (このプロセスの要素を探し、読み書きする)
//   この .exe  = UIA PROVIDER (要素を公開する)
//
// 制御チャネル (command / start / progress / result / cancel / shutdown) は
// UI Automation だけを通る。ファイル・COM・名前付きパイプ・ソケット・レジストリ
// ・ウィンドウメッセージは制御に使わない。
//
// ただし「郵便番号100万件」という大量データそのものは UIA で運ばない。
// ValuePattern の往復を 100 万回行うのは非現実的であり、計測の意味も失われる。
// よって
//     制御 (どのファイルを、いつ、変換するか / 進捗 / 結果 / 取消) = UIA
//     データ (入力100万行 / 出力100万行 / 辞書12万行)              = ファイル
// と明確に分ける。これは全方式で共通の取り決めで、方式間の比較条件を揃える。
//
// UIA プロバイダの実装方針は poc-reference/worker-reference/UiaWorker.cs で
// 実測済みの構成をそのまま踏襲している:
//   * 可視でないウィンドウは UIA ツリーに現れないため、最小のホストウィンドウを
//     1 枚だけ持ち、仮想画面の外に置く (offscreen)
//   * WinForms コントロールは Win11 26200 では ValuePattern を持たない Pane に
//     なるため使わず、サーバ側プロバイダを手書きする
//   * ServerSideProvider を UseComThreading なしで登録するので UIA 呼び出しは
//     RPC スレッドに来る。よって重い変換中でも Excel からの読み取りに即答する
//
// ---------------------------------------------------------------------------
// 変換規則 (ZipBench rule v1) -- VBA 側 modZipRule.bas と 1 文字も違わないこと
// ---------------------------------------------------------------------------
//   1. 郵便番号の正規化: 全角数字→半角、ハイフン類・空白・〒 を除去。
//      結果がちょうど半角数字 7 桁でなければ「不正」とする。
//   2. 辞書 zip_dict.csv (zip7,住所) を引く。
//   3. 見つからない/不正なら "\u8A72\u5F53\u306A\u3057"。
//   辞書そのもの (KEN_ALL からの町域整形) は Excel 側で 1 度だけ作る。
//   全方式が同じ zip_dict.csv を読むので、変換規則の差は原理的に生じない。
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ZipWorkerApp
{
    // =======================================================================
    // UIA interop. vtable 順は UIAutomationCore.idl と一致させること。
    // 間違えると「要素は見えるが何も答えない」状態になる。
    // =======================================================================
    #region UIA interop

    [Flags]
    public enum ProviderOptions
    {
        ClientSideProvider    = 0x01,
        ServerSideProvider    = 0x02,
        NonClientAreaProvider = 0x04,
        OverrideProvider      = 0x08,
        ProviderOwnsSetFocus  = 0x10,
        UseComThreading       = 0x20
    }

    public enum NavigateDirection
    {
        Parent = 0, NextSibling = 1, PreviousSibling = 2, FirstChild = 3, LastChild = 4
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct UiaRect { public double left, top, width, height; }

    [ComVisible(true), Guid("d6dd68d1-86fd-4332-8666-9abedea2d24c")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRawElementProviderSimple
    {
        ProviderOptions ProviderOptions { get; }
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object GetPatternProvider(int patternId);
        [return: MarshalAs(UnmanagedType.Struct)]
        object GetPropertyValue(int propertyId);
        IRawElementProviderSimple HostRawElementProvider { get; }
    }

    [ComVisible(true), Guid("f7063da8-8359-439c-9297-bbc5299a7d87")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRawElementProviderFragment : IRawElementProviderSimple
    {
        IRawElementProviderFragment Navigate(NavigateDirection direction);
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_I4)]
        int[] GetRuntimeId();
        UiaRect BoundingRectangle { get; }
        [return: MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_UNKNOWN)]
        IRawElementProviderSimple[] GetEmbeddedFragmentRoots();
        void SetFocus();
        IRawElementProviderFragmentRoot FragmentRoot { get; }
    }

    [ComVisible(true), Guid("620ce2a5-ab8f-40a9-86cb-de3c75599b58")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IRawElementProviderFragmentRoot : IRawElementProviderFragment
    {
        IRawElementProviderFragment ElementProviderFromPoint(double x, double y);
        IRawElementProviderFragment GetFocus();
    }

    [ComVisible(true), Guid("c7935180-6fb3-4201-b174-7df73adbf64a")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IValueProvider
    {
        void SetValue([MarshalAs(UnmanagedType.LPWStr)] string value);
        string Value { [return: MarshalAs(UnmanagedType.BStr)] get; }
        bool IsReadOnly { [return: MarshalAs(UnmanagedType.Bool)] get; }
    }

    [ComVisible(true), Guid("54fcb24b-e18e-47a2-b4d3-eccbe77599a2")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IInvokeProvider { void Invoke(); }

    internal static class Uia
    {
        public const int ProcessId = 30002, ControlType = 30003, LocalizedControlType = 30004;
        public const int Name = 30005, IsKeyboardFocusable = 30009, IsEnabled = 30010;
        public const int AutomationId = 30011, ClassName = 30012, HelpText = 30013;
        public const int IsControlElement = 30016, IsContentElement = 30017;
        public const int IsOffscreen = 30022, FrameworkId = 30024;
        public const int ValueValue = 30045;

        public const int CtButton = 50000, CtEdit = 50004;

        public const int PatInvoke = 10000, PatValue = 10002;

        public const int UiaRootObjectId = -25;
        public const int WM_GETOBJECT = 0x003D;
        public const int UiaAppendRuntimeId = 3;

        [DllImport("UIAutomationCore.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr UiaReturnRawElementProvider(
            IntPtr hwnd, IntPtr wParam, IntPtr lParam, IRawElementProviderSimple el);

        [DllImport("UIAutomationCore.dll")]
        public static extern int UiaHostProviderFromHwnd(
            IntPtr hwnd, [MarshalAs(UnmanagedType.Interface)] out IRawElementProviderSimple provider);

        [DllImport("UIAutomationCore.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UiaClientsAreListening();

        [DllImport("UIAutomationCore.dll")]
        public static extern int UiaRaiseAutomationPropertyChangedEvent(
            IRawElementProviderSimple provider, int id,
            [MarshalAs(UnmanagedType.Struct)] object oldValue,
            [MarshalAs(UnmanagedType.Struct)] object newValue);

        [DllImport("UIAutomationCore.dll")]
        public static extern int UiaDisconnectAllProviders();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        // 「起動側から渡された表示状態」を実際に読むための宣言。
        // 推測ではなく worker.log に事実として残す。
        // 文字列 3 つは IntPtr で受ける。string にすると、CLR が構造体の後始末で
        // ネイティブ側のバッファを解放しようとしてプロセスごと落ちる。
        // (実測: string にした版はログを 1 行も残さず即死した)
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct STARTUPINFO
        {
            public int cb;
            public IntPtr lpReserved, lpDesktop, lpTitle;
            public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
            public short wShowWindow, cbReserved2;
            public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern void GetStartupInfo(out STARTUPINFO lpStartupInfo);

        public const int STARTF_USESHOWWINDOW = 0x00000001;

        public const int SW_SHOWNOACTIVATE  = 4;   // show, do not steal focus
        public const int SW_SHOWMINNOACTIVE = 7;   // minimize, do not steal focus

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }
    }

    #endregion

    // =======================================================================
    // チャネル定義。クライアントは AutomationId で参照する。
    // poc-reference と同じ名前にしてあるので VBA 側の UIA クライアントは共通。
    // =======================================================================
    internal static class Chan
    {
        public const string Command  = "UIAW_CMD";
        public const string Submit   = "UIAW_SUBMIT";
        public const string Ack      = "UIAW_ACK";
        public const string Status   = "UIAW_STATUS";
        public const string Progress = "UIAW_PROGRESS";
        public const string Result   = "UIAW_RESULT";
        public const string Heart    = "UIAW_HEART";
        public const string Info     = "UIAW_INFO";

        public static readonly string[] All =
            { Command, Submit, Ack, Status, Progress, Result, Heart, Info };
    }

    // =======================================================================
    // Excel への直接書き込み。
    //
    // 変換結果 100 万件は、ファイルに書いて Excel に読み直させるのではなく、
    // このプロセスから COM で対象 Excel のセルへ直接置く。実測 (1,000,000 件):
    //     ファイルに書いて Excel が VBA 配列へ取り込む   9,500 ms (しかもセルには入らない)
    //     COM で object[,] を Range.Value2 へ            7,200 ms
    //     COM で string[,] を Range.Value2 へ            5,030 ms
    // 分割しても速くならない。costは呼び出し回数ではなく、100 万個の BSTR を
    // プロセス境界で運ぶこと自体。だから string[,] (VT_BSTR SAFEARRAY) を使う。
    // object[,] は要素ごとに 16 バイトの VARIANT ヘッダが付くぶん重い。
    //
    // 相手の Excel は「ウィンドウハンドル」で特定する。GetActiveObject や
    // ROT からの取得は、他人の Excel を掴む可能性がある。所有していない
    // インスタンスには触らない、という約束を守れるのはこの経路だけ。
    // =======================================================================
    #region Excel COM bridge

    // VBA が走っている最中の外部からの COM 呼び出しを Excel は拒否できる
    // (RPC_E_SERVERCALL_RETRYLATER)。既定では即座に例外になるので、
    // 「少し待って再送」を COM に指示するフィルタを入れておく。
    // これが無いと、混んだときだけ落ちる構成になる。
    [ComImport, Guid("00000016-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }

    internal sealed class RetryFilter : IOleMessageFilter
    {
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

        internal static void Register()
        {
            IOleMessageFilter old;
            CoRegisterMessageFilter(new RetryFilter(), out old);
        }

        internal static void Revoke()
        {
            IOleMessageFilter old;
            CoRegisterMessageFilter(null, out old);
        }

        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
        {
            return 0;   // SERVERCALL_ISHANDLED
        }

        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            // 2 = SERVERCALL_RETRYLATER。100 ms 後に再送させる。
            // -1 を返すと呼び出しは失敗として返る。
            if (dwRejectType == 2) return 100;
            return -1;
        }

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
        {
            return 2;   // PENDINGMSG_WAITDEFPROCESS
        }
    }

    internal static class Xl
    {
        private const int OBJID_NATIVEOM = unchecked((int)0xFFFFFFF0);
        private const BindingFlags GetOrCall =
            BindingFlags.GetProperty | BindingFlags.InvokeMethod;

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd, int dwObjectID, ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);

        // Cells も Range も「引数付きプロパティ」なので、GetProperty と
        // InvokeMethod の両方を許して IDispatch に選ばせる。片方だけだと
        // DISP_E_MEMBERNOTFOUND になる。
        internal static object Get(object o, string name, params object[] args)
        {
            return o.GetType().InvokeMember(name, GetOrCall, null, o, args);
        }

        internal static void Set(object o, string name, params object[] args)
        {
            o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, args);
        }

        internal static void Release(object o)
        {
            try { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); }
            catch { }
        }

        // Excel のオブジェクトモデルは XLMAIN > XLDESK > EXCEL7 の
        // EXCEL7 ペインからしか取れない。ここを辿ると、渡されたハンドルの
        // インスタンスだけに確実に着く。
        internal static object AppFromHwnd(IntPtr xlMain)
        {
            IntPtr desk = FindWindowEx(xlMain, IntPtr.Zero, "XLDESK", null);
            if (desk == IntPtr.Zero) throw new Exception("XLDESK not found under hwnd 0x" + xlMain.ToInt64().ToString("X"));
            IntPtr pane = FindWindowEx(desk, IntPtr.Zero, "EXCEL7", null);
            if (pane == IntPtr.Zero) throw new Exception("EXCEL7 not found; that Excel has no open workbook window");

            Guid iid = new Guid("00020400-0000-0000-C000-000000000046");   // IID_IDispatch
            object win;
            int hr = AccessibleObjectFromWindow(pane, OBJID_NATIVEOM, ref iid, out win);
            if (hr != 0 || win == null)
                throw new Exception("AccessibleObjectFromWindow failed hr=0x" + hr.ToString("X8"));

            object app = Get(win, "Application");
            Release(win);
            return app;
        }

        internal static object WorkbookByName(object app, string name)
        {
            object books = Get(app, "Workbooks");
            int count = Convert.ToInt32(Get(books, "Count"), CultureInfo.InvariantCulture);
            for (int i = 1; i <= count; i++)
            {
                object wb = Get(books, "Item", i);
                string n = Convert.ToString(Get(wb, "Name"), CultureInfo.InvariantCulture);
                if (string.Equals(n, name, StringComparison.OrdinalIgnoreCase))
                {
                    Release(books);
                    return wb;
                }
                Release(wb);
            }
            Release(books);
            throw new Exception("workbook not found in that Excel: " + name);
        }

        internal static object SheetByName(object wb, string name)
        {
            object sheets = Get(wb, "Worksheets");
            object ws = Get(sheets, "Item", name);
            Release(sheets);
            return ws;
        }

        // 1 列ぶんを 10 万件ずつ書く。粒度は実測で決めた
        // (一括でも 10 万でも同速。細かいほうが取消と再試行の単位が細かい)。
        internal const int WriteChunk = 100000;

        internal static void WriteColumn(object ws, int firstRow, int col, string[] values, int count,
                                         Func<bool> cancelled)
        {
            int written = 0;
            while (written < count)
            {
                if (cancelled != null && cancelled()) return;

                int take = Math.Min(WriteChunk, count - written);
                string[,] buf = new string[take, 1];
                for (int i = 0; i < take; i++) buf[i, 0] = values[written + i];

                object c1 = Get(ws, "Cells", firstRow + written, col);
                object c2 = Get(ws, "Cells", firstRow + written + take - 1, col);
                object rng = Get(ws, "Range", c1, c2);
                Set(rng, "Value2", buf);
                Release(rng); Release(c1); Release(c2);

                written += take;
            }
        }

        internal static void WriteCell(object ws, string a1, string value)
        {
            object rng = Get(ws, "Range", a1);
            Set(rng, "Value2", value);
            Release(rng);
        }

        // 1 列を 1 回で読む。Excel は 2 次元の VARIANT 配列を返す (1 起点)。
        internal static object[,] ReadColumn(object ws, int firstRow, int lastRow, int col)
        {
            object c1 = Get(ws, "Cells", firstRow, col);
            object c2 = Get(ws, "Cells", lastRow, col);
            object rng = Get(ws, "Range", c1, c2);
            object v = Get(rng, "Value2");
            Release(rng); Release(c1); Release(c2);
            return (object[,])v;
        }

        // 2 列を 1 回で読む (MASTER の 郵便番号 + 住所)。
        internal static object[,] ReadBlock(object ws, int firstRow, int lastRow, int firstCol, int lastCol)
        {
            object c1 = Get(ws, "Cells", firstRow, firstCol);
            object c2 = Get(ws, "Cells", lastRow, lastCol);
            object rng = Get(ws, "Range", c1, c2);
            object v = Get(rng, "Value2");
            Release(rng); Release(c1); Release(c2);
            return (object[,])v;
        }

        // 1 列を 1 回で書く。分割しない。
        internal static void WriteColumnOnce(object ws, int firstRow, int col, string[,] values)
        {
            int n = values.GetLength(0);
            object c1 = Get(ws, "Cells", firstRow, col);
            object c2 = Get(ws, "Cells", firstRow + n - 1, col);
            object rng = Get(ws, "Range", c1, c2);
            Set(rng, "Value2", values);
            Release(rng); Release(c1); Release(c2);
        }
    }

    #endregion

    // =======================================================================
    // 変換規則。VBA の modZipRule.NormalizeZip / FnvHashOfArray と同一。
    // =======================================================================
    internal static class ZipRule
    {
        // "\u8A72\u5F53\u306A\u3057"。ソースの文字コードが何であっても意味が変わらないよう、
        // 非 ASCII はすべて \u エスケープで書く (コメントの文字化けは無害)。
        public const string NotFound = "\u8A72\u5F53\u306A\u3057";   // 該当なし

        // 全角数字→半角、区切り文字の除去。ちょうど 7 桁でなければ "" を返す。
        // 除去する文字は VBA の modZipRule.NormalizeZip と同一集合。
        internal static string NormalizeZip(string s)
        {
            if (s == null) return "";
            char[] buf = new char[8];
            int n = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c >= '\uFF10' && c <= '\uFF19') c = (char)('0' + (c - '\uFF10'));  // 全角数字
                if (c >= '0' && c <= '9')
                {
                    if (n >= 8) return "";
                    buf[n++] = c;
                    continue;
                }
                switch (c)
                {
                    case '-':        // ASCII HYPHEN-MINUS
                    case '\u2010':   // HYPHEN
                    case '\u2012':   // FIGURE DASH
                    case '\u2013':   // EN DASH
                    case '\u2014':   // EM DASH
                    case '\u2015':   // HORIZONTAL BAR
                    case '\u2212':   // MINUS SIGN
                    case '\uFF0D':   // FULLWIDTH HYPHEN-MINUS
                    case '\u30FC':   // KATAKANA-HIRAGANA PROLONGED SOUND MARK
                    case '\u3012':   // POSTAL MARK
                    case ' ':
                    case '\u3000':   // IDEOGRAPHIC SPACE
                    case '\t':
                    case '\r':
                    case '\n':
                        continue;
                    default:
                        return "";   // それ以外の文字が混ざったら不正
                }
            }
            return (n == 7) ? new string(buf, 0, 7) : "";
        }

        // FNV-1a 32bit を UTF-16 コードユニット列に対して回す。
        // 各出力行のあとに '\n' を 1 つ混ぜる。VBA 側と完全に同じ手順。
        // 出力が 2 次元 (n,1) のときの同じハッシュ。VBA 側・PowerShell 参照実装と
        // 同じ手順 (各行のあとに '\n' を 1 つ混ぜる)。
        internal static uint FnvOf2D(string[,] values, int count)
        {
            unchecked
            {
                uint h = 2166136261u;
                for (int i = 0; i < count; i++)
                {
                    string s = values[i, 0];
                    if (s != null)
                    {
                        for (int j = 0; j < s.Length; j++)
                        {
                            h = (h ^ (uint)(s[j] & 0xFFFF)) * 16777619u;
                        }
                    }
                    h = (h ^ (uint)'\n') * 16777619u;
                }
                return h;
            }
        }

        internal static uint FnvOfLines(string[] lines, int count)
        {
            unchecked
            {
                uint h = 2166136261u;
                for (int i = 0; i < count; i++)
                {
                    string s = lines[i];
                    for (int j = 0; j < s.Length; j++)
                    {
                        h = (h ^ (uint)(s[j] & 0xFFFF)) * 16777619u;
                    }
                    h = (h ^ (uint)'\n') * 16777619u;
                }
                return h;
            }
        }
    }

    // =======================================================================
    // 1 つの仮想要素。HWND も WinForms コントロールも持たない。
    // =======================================================================
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class FieldProvider :
        IRawElementProviderSimple, IRawElementProviderFragment, IValueProvider, IInvokeProvider
    {
        private readonly RootProvider _root;
        private readonly string _id;
        private readonly int _index;
        private readonly bool _isButton;
        private readonly bool _readOnly;

        internal FieldProvider(RootProvider root, string id, int index, bool isButton, bool readOnly)
        {
            _root = root; _id = id; _index = index; _isButton = isButton; _readOnly = readOnly;
        }

        internal string Id { get { return _id; } }

        public ProviderOptions ProviderOptions { get { return ProviderOptions.ServerSideProvider; } }

        public object GetPatternProvider(int patternId)
        {
            if (_isButton && patternId == Uia.PatInvoke) return this;
            if (!_isButton && patternId == Uia.PatValue) return this;
            return null;
        }

        public object GetPropertyValue(int propertyId)
        {
            switch (propertyId)
            {
                case Uia.Name:                 return _id;
                case Uia.AutomationId:         return _id;
                case Uia.ControlType:          return _isButton ? Uia.CtButton : Uia.CtEdit;
                case Uia.LocalizedControlType: return _isButton ? "button" : "edit";
                case Uia.ClassName:            return "ZipWorkerField";
                case Uia.FrameworkId:          return "ZipBenchWorker";
                case Uia.IsControlElement:     return true;
                case Uia.IsContentElement:     return true;
                case Uia.IsEnabled:            return true;
                case Uia.IsKeyboardFocusable:  return false;
                case Uia.IsOffscreen:          return false;
                case Uia.HelpText:             return "ZipBench UIA channel field";
                default:                       return null;
            }
        }

        public IRawElementProviderSimple HostRawElementProvider { get { return null; } }

        public IRawElementProviderFragment Navigate(NavigateDirection direction)
        {
            switch (direction)
            {
                case NavigateDirection.Parent:          return _root;
                case NavigateDirection.NextSibling:     return _root.ChildAt(_index + 1);
                case NavigateDirection.PreviousSibling: return _root.ChildAt(_index - 1);
                default:                                return null;
            }
        }

        public int[] GetRuntimeId() { return new int[] { Uia.UiaAppendRuntimeId, _index + 1 }; }

        public UiaRect BoundingRectangle { get { return _root.WindowRect(); } }

        public IRawElementProviderSimple[] GetEmbeddedFragmentRoots() { return null; }

        public void SetFocus() { }

        public IRawElementProviderFragmentRoot FragmentRoot { get { return _root; } }

        public string Value { get { return _root.Core.Get(_id); } }

        public bool IsReadOnly { get { return _readOnly || _isButton; } }

        public void SetValue(string value)
        {
            if (IsReadOnly) throw new InvalidOperationException("field is read-only: " + _id);
            _root.Core.Set(_id, value ?? "");
        }

        // UIA の RPC スレッドから呼ばれる。すぐ返すこと。重い処理は別スレッドへ。
        public void Invoke() { _root.Core.ExecuteCurrentCommand(); }
    }

    // =======================================================================
    // ホストウィンドウの HWND に結びつく fragment root。
    // =======================================================================
    [ComVisible(true), ClassInterface(ClassInterfaceType.None)]
    public sealed class RootProvider :
        IRawElementProviderSimple, IRawElementProviderFragment, IRawElementProviderFragmentRoot
    {
        private readonly FieldProvider[] _children;
        private IRawElementProviderSimple _hostProvider;
        private readonly object _hostGate = new object();
        private volatile IntPtr _hwnd = IntPtr.Zero;

        internal WorkerCore Core { get; private set; }

        internal void BindHwnd(IntPtr h) { _hwnd = h; }

        internal RootProvider(WorkerCore core)
        {
            Core = core;
            _children = new FieldProvider[Chan.All.Length];
            for (int i = 0; i < Chan.All.Length; i++)
            {
                string id = Chan.All[i];
                bool isButton = (id == Chan.Submit);
                bool ro = !(id == Chan.Command);
                _children[i] = new FieldProvider(this, id, i, isButton, ro);
            }
        }

        internal FieldProvider ChildAt(int i)
        {
            if (i < 0 || i >= _children.Length) return null;
            return _children[i];
        }

        internal FieldProvider ChildById(string id)
        {
            foreach (FieldProvider f in _children) { if (f.Id == id) return f; }
            return null;
        }

        internal UiaRect WindowRect()
        {
            UiaRect r = new UiaRect();
            try
            {
                IntPtr h = _hwnd;
                if (h == IntPtr.Zero) return r;
                Uia.RECT n;
                if (Uia.GetWindowRect(h, out n))
                {
                    r.left = n.Left; r.top = n.Top;
                    r.width = n.Right - n.Left; r.height = n.Bottom - n.Top;
                }
            }
            catch { }
            return r;
        }

        public ProviderOptions ProviderOptions { get { return ProviderOptions.ServerSideProvider; } }

        public object GetPatternProvider(int patternId) { return null; }

        public object GetPropertyValue(int propertyId)
        {
            switch (propertyId)
            {
                case Uia.AutomationId: return "ZIPWORKER_ROOT";
                case Uia.FrameworkId:  return "ZipBenchWorker";
                case Uia.ClassName:    return "ZipWorkerRoot";
                default:               return null;
            }
        }

        public IRawElementProviderSimple HostRawElementProvider
        {
            get
            {
                lock (_hostGate)
                {
                    if (_hostProvider == null && _hwnd != IntPtr.Zero)
                    {
                        IRawElementProviderSimple p;
                        Uia.UiaHostProviderFromHwnd(_hwnd, out p);
                        _hostProvider = p;
                    }
                    return _hostProvider;
                }
            }
        }

        public IRawElementProviderFragment Navigate(NavigateDirection direction)
        {
            switch (direction)
            {
                case NavigateDirection.FirstChild: return _children[0];
                case NavigateDirection.LastChild:  return _children[_children.Length - 1];
                default:                           return null;
            }
        }

        public int[] GetRuntimeId() { return null; }

        public UiaRect BoundingRectangle { get { return WindowRect(); } }

        public IRawElementProviderSimple[] GetEmbeddedFragmentRoots() { return null; }

        public void SetFocus() { }

        public IRawElementProviderFragmentRoot FragmentRoot { get { return this; } }

        public IRawElementProviderFragment ElementProviderFromPoint(double x, double y) { return null; }
        public IRawElementProviderFragment GetFocus() { return null; }

        internal void RaiseValueChanged(string id, string oldValue, string newValue)
        {
            try
            {
                if (!Uia.UiaClientsAreListening()) return;
                FieldProvider f = ChildById(id);
                if (f == null) return;
                Uia.UiaRaiseAutomationPropertyChangedEvent(f, Uia.ValueValue, oldValue, newValue);
            }
            catch { }
        }
    }

    // =======================================================================
    // チャネル状態 + コマンド処理 + 郵便番号変換ジョブ本体
    // =======================================================================
    internal sealed class WorkerCore
    {
        private readonly object _gate = new object();
        private readonly Dictionary<string, string> _state = new Dictionary<string, string>();
        private readonly string _logPath;

        internal RootProvider Root;
        internal Form Host;

        private Thread _jobThread;
        private volatile bool _cancelRequested;
        private volatile bool _jobRunning;
        internal DateTime LastCommandUtc = DateTime.UtcNow;

        // 辞書はキャッシュしない。毎回 Master シートから読んで組み直す。
        // キャッシュがあると 2 回目以降だけ速くなり、繰り返し計測の意味が失われる。

        internal WorkerCore(string logPath)
        {
            _logPath = logPath;
            foreach (string k in Chan.All) { _state[k] = ""; }
            _state[Chan.Status]   = "IDLE";
            _state[Chan.Progress] = "0";
            _state[Chan.Ack]      = "0";
            _state[Chan.Heart]    = "0";
        }

        internal string Get(string id)
        {
            lock (_gate) { string v; return _state.TryGetValue(id, out v) ? v : ""; }
        }

        internal void Set(string id, string value)
        {
            string old;
            lock (_gate)
            {
                _state.TryGetValue(id, out old);
                if (old == value) return;
                _state[id] = value;
            }
            RootProvider r = Root;
            if (r != null) r.RaiseValueChanged(id, old, value);
        }

        internal bool JobRunning { get { return _jobRunning; } }

        // ===================================================================
        // 唯一の入口。IInvokeProvider.Invoke すなわち UIA の RPC スレッドから来る。
        // 電文: <seq>|<VERB>|<k=v>;<k=v>
        // ===================================================================
        internal void ExecuteCurrentCommand()
        {
            string raw = Get(Chan.Command);
            LastCommandUtc = DateTime.UtcNow;
            Log("command <- " + raw);

            string[] parts = raw.Split(new char[] { '|' }, 3);
            long seq = 0;
            if (parts.Length >= 1)
                long.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out seq);
            string verb = parts.Length >= 2 ? parts[1].Trim().ToUpperInvariant() : "";
            string arg = parts.Length >= 3 ? parts[2] : "";

            try
            {
                switch (verb)
                {
                    case "CONVERT":  DoConvert(arg); break;
                    case "CANCEL":   DoCancel();     break;
                    case "PING":
                        Set(Chan.Result, "pong " + DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture));
                        break;
                    case "SHUTDOWN": DoShutdown();   break;
                    default:
                        Set(Chan.Status, "ERROR");
                        Set(Chan.Result, "unknown verb: " + verb);
                        break;
                }
            }
            catch (Exception ex)
            {
                Set(Chan.Status, "ERROR");
                Set(Chan.Result, ex.GetType().Name + ": " + ex.Message);
                Log("command failed: " + ex);
            }

            // ack は最後。クライアントは「ack == 自分の seq」を
            // 「コマンドを処理し終えた」と見なせる。
            Set(Chan.Ack, seq.ToString(CultureInfo.InvariantCulture));
        }

        private static Dictionary<string, string> ParseArgs(string arg)
        {
            Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string kv in arg.Split(';'))
            {
                int eq = kv.IndexOf('=');
                if (eq <= 0) continue;
                d[kv.Substring(0, eq).Trim()] = kv.Substring(eq + 1).Trim();
            }
            return d;
        }

        // CONVERT が要求するもの。結果はファイルではなく、呼び出し元の Excel の
        // セルへ直接置く。だから「どの Excel の、どのブックの、どのシートの、
        // 何列目の、何行目から」が要る。
        internal sealed class ConvertJob
        {
            internal string JobId;
            internal IntPtr XlHwnd;
            internal string WbName;
            internal string MasterSheet, InputSheet, OutputSheet;
            internal int MasterRows, Rows;
            internal string SigWbName, SigSheetName, SigCell;

            // 結果の届け方。ここだけが候補2と候補3の違い。
            //   "com" : Output.Value2 へ 1 回だけ一括書込 (候補3)
            //   "tsv" : 一時名で TSV を書き、完成後に atomic rename (候補2)
            //           Excel 側が QueryTable でネイティブに取り込む
            internal string Deliver;
            internal string TsvPath;
        }

        private void DoConvert(string arg)
        {
            if (_jobRunning) { Set(Chan.Result, "already running"); return; }

            // UIA で受け取るのは「どこにあるか」だけ。100 万件そのものは
            // UIA にも COM の引数にも載せない。ワーカーが Excel から直接読む。
            Dictionary<string, string> a = ParseArgs(arg);
            ConvertJob j = new ConvertJob();
            string v;

            a.TryGetValue("jobid", out j.JobId);
            a.TryGetValue("wb", out j.WbName);
            a.TryGetValue("master", out j.MasterSheet);
            a.TryGetValue("input", out j.InputSheet);
            a.TryGetValue("output", out j.OutputSheet);
            a.TryGetValue("sigwb", out j.SigWbName);
            a.TryGetValue("sigsheet", out j.SigSheetName);
            a.TryGetValue("sigcell", out j.SigCell);
            a.TryGetValue("tsv", out j.TsvPath);
            if (!a.TryGetValue("deliver", out j.Deliver) || j.Deliver.Length == 0) j.Deliver = "com";
            j.Deliver = j.Deliver.ToLowerInvariant();

            long h = 0;
            if (a.TryGetValue("xlhwnd", out v)) long.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out h);
            j.XlHwnd = new IntPtr(h);
            j.MasterRows = a.TryGetValue("masterrows", out v) ? int.Parse(v, CultureInfo.InvariantCulture) : 0;
            j.Rows = a.TryGetValue("rows", out v) ? int.Parse(v, CultureInfo.InvariantCulture) : 0;

            if (h == 0 || string.IsNullOrEmpty(j.WbName) ||
                string.IsNullOrEmpty(j.MasterSheet) || string.IsNullOrEmpty(j.InputSheet) ||
                string.IsNullOrEmpty(j.OutputSheet) || j.MasterRows < 1 || j.Rows < 1)
            {
                Set(Chan.Status, "ERROR");
                Set(Chan.Result, "CONVERT needs xlhwnd=, wb=, master=, input=, output=, masterrows=, rows=");
                return;
            }

            _cancelRequested = false;
            _jobRunning = true;
            Set(Chan.Progress, "0");
            Set(Chan.Result, "");
            Set(Chan.Status, "RUNNING");

            ConvertJob captured = j;
            _jobThread = new Thread(delegate () { RunConvert(captured); });
            _jobThread.IsBackground = true;
            // STA でなければ COM の呼び出しごとにマーシャリングが挟まり、
            // メッセージフィルタ (RetryRejectedCall) も効かない。
            _jobThread.SetApartmentState(ApartmentState.STA);
            _jobThread.Start();
        }

        private void DoCancel()
        {
            if (!_jobRunning) { Set(Chan.Result, "nothing to cancel"); return; }
            _cancelRequested = true;
            Log("cancel requested");
        }

        private void DoShutdown()
        {
            _cancelRequested = true;
            Set(Chan.Status, "SHUTDOWN");
            Log("shutdown requested");
            Form h = Host;
            if (h == null) return;
            try
            {
                h.BeginInvoke((MethodInvoker)delegate ()
                {
                    try { Uia.UiaDisconnectAllProviders(); } catch { }
                    Application.ExitThread();
                });
            }
            catch (InvalidOperationException) { }
        }

        // ===================================================================
        // 変換本体。
        //
        //   変換秒 = 辞書読み込み + 入力読み込み + 変換        (このプロセスの中)
        //   書込秒 = 前面ブックのセルへ COM で一括書き込み      (プロセス境界を越える)
        //
        // 最後に通知セルを 1 つだけ書く。Excel 側はその Worksheet_Change で
        // 完了を知る。Excel にポーリングはさせない。
        //
        // 自分専用の STA スレッドで走るので、重い書き込み中でも UIA の
        // 読み取り (RPC スレッド) は止まらない。
        // ===================================================================
        private void RunConvert(ConvertJob job)
        {
            long bindMs = 0, masterMs = 0, inputMs = 0, dictMs = 0, convMs = 0, writeMs = 0, sigMs = 0;
            int rows = 0, dictRows = 0, notFound = 0;
            uint hash = 0;

            object app = null, wb = null, wsM = null, wsI = null, wsO = null, sigWb = null, sigWs = null;
            bool filterOn = false;
            Stopwatch e2e = Stopwatch.StartNew();

            try
            {
                // Excel が VBA を実行している最中でも外から呼べるように、
                // 拒否されたら再送するフィルタを先に入れる。
                RetryFilter.Register();
                filterOn = true;

                // ---- 1. 対象 Excel へ接続 ----
                Stopwatch sw = Stopwatch.StartNew();
                app = Xl.AppFromHwnd(job.XlHwnd);
                wb = Xl.WorkbookByName(app, job.WbName);
                wsM = Xl.SheetByName(wb, job.MasterSheet);
                wsI = Xl.SheetByName(wb, job.InputSheet);
                wsO = Xl.SheetByName(wb, job.OutputSheet);
                sw.Stop(); bindMs = sw.ElapsedMilliseconds;

                // ---- 2. Master を 1 回で読む ----
                sw = Stopwatch.StartNew();
                object[,] master = Xl.ReadBlock(wsM, 1, job.MasterRows, 1, 2);
                sw.Stop(); masterMs = sw.ElapsedMilliseconds;

                // ---- 3. Input を 1 回で読む ----
                sw = Stopwatch.StartNew();
                object[,] input = Xl.ReadColumn(wsI, 1, job.Rows, 1);
                sw.Stop(); inputMs = sw.ElapsedMilliseconds;

                // ---- 4. 辞書構築 ----
                sw = Stopwatch.StartNew();
                Dictionary<string, string> dict =
                    new Dictionary<string, string>(job.MasterRows + 16, StringComparer.Ordinal);
                for (int i = 1; i <= job.MasterRows; i++)
                {
                    object k = master[i, 1];
                    if (k == null) continue;
                    string key = Convert.ToString(k, CultureInfo.InvariantCulture);
                    if (key.Length == 0 || dict.ContainsKey(key)) continue;
                    object a2 = master[i, 2];
                    dict.Add(key, a2 == null ? "" : Convert.ToString(a2, CultureInfo.InvariantCulture));
                }
                dictRows = dict.Count;
                sw.Stop(); dictMs = sw.ElapsedMilliseconds;

                // ---- 5. 変換 + 出力配列作成 ----
                // 出力は string[,] にする。VT_BSTR の SAFEARRAY になるので、
                // object[,] (要素ごとに VARIANT ヘッダが付く) より 30% 速い。
                sw = Stopwatch.StartNew();
                int n = job.Rows;
                string[,] output = new string[n, 1];
                for (int i = 0; i < n; i++)
                {
                    object cell = input[i + 1, 1];
                    string key = ZipRule.NormalizeZip(
                        cell == null ? "" : Convert.ToString(cell, CultureInfo.InvariantCulture));
                    string addr;
                    if (key.Length == 0 || !dict.TryGetValue(key, out addr))
                    {
                        addr = ZipRule.NotFound;
                        notFound++;
                    }
                    output[i, 0] = addr;
                    if (_cancelRequested) break;
                }
                rows = n;
                sw.Stop(); convMs = sw.ElapsedMilliseconds;

                if (_cancelRequested)
                {
                    _jobRunning = false;
                    Set(Chan.Result, "cancelled before write");
                    Set(Chan.Status, "CANCELLED");
                    Log("convert CANCELLED before write");
                    return;
                }

                // 照合用ハッシュ。どの計測区間にも入れない。
                hash = ZipRule.FnvOf2D(output, n);

                // ---- 6. 結果を届ける ----
                // ここだけが候補2と候補3の違い。上の 1～5 は完全に同じコード。
                sw = Stopwatch.StartNew();
                if (job.Deliver == "tsv")
                {
                    // 一時名へ書いてから rename する。読み手が半端なファイルを
                    // 掴む余地を無くすため。同一ボリューム内の rename なので
                    // 原子的に置き換わる。
                    string tmp = job.TsvPath + ".tmp";
                    // BOM は付けない。QueryTable も OpenText も符号化を明示して渡すので
                // BOM を必要としないが、ADO と DAO のテキスト ドライバは BOM を
                // 1 列目のデータとして読んでしまい、先頭行だけが化ける。
                // 4 経路が同じ 1 本のファイルを読めるように、無い側へ揃える。
                using (StreamWriter w = new StreamWriter(tmp, false, new UTF8Encoding(false), 1 << 20))
                    {
                        w.NewLine = "\r\n";
                        for (int i = 0; i < n; i++) w.WriteLine(output[i, 0]);
                    }
                    if (File.Exists(job.TsvPath)) File.Delete(job.TsvPath);
                    File.Move(tmp, job.TsvPath);
                }
                else
                {
                    Xl.WriteColumnOnce(wsO, 1, 1, output);
                }
                sw.Stop(); writeMs = sw.ElapsedMilliseconds;

                e2e.Stop();

                // UIA 側の結果を先に埋める。通知セルを書いた瞬間に Excel は
                // 動き出すので、結果を後から書くと読みに来たときにまだ空になる。
                Set(Chan.Progress, "100");
                Set(Chan.Result, BuildPayload(rows, dictRows, notFound, bindMs, masterMs,
                                              inputMs, dictMs, convMs, writeMs,
                                              e2e.ElapsedMilliseconds, hash));
                Set(Chan.Status, "DONE");

                // ---- 7. 通知セルを 1 回だけ更新。これが Excel の完了合図 ----
                if (!string.IsNullOrEmpty(job.SigCell))
                {
                    sw = Stopwatch.StartNew();
                    sigWb = string.Equals(job.SigWbName, job.WbName, StringComparison.OrdinalIgnoreCase)
                          ? wb : Xl.WorkbookByName(app, job.SigWbName);
                    sigWs = Xl.SheetByName(sigWb, job.SigSheetName);
                    Xl.WriteCell(sigWs, job.SigCell, SigText(job.JobId, rows, hash, bindMs,
                                                             masterMs, inputMs, dictMs, convMs,
                                                             writeMs, e2e.ElapsedMilliseconds));
                    sw.Stop(); sigMs = sw.ElapsedMilliseconds;
                }
            }
            catch (Exception ex)
            {
                e2e.Stop();
                _jobRunning = false;
                Set(Chan.Result, ex.GetType().Name + ": " + ex.Message);
                Set(Chan.Status, "ERROR");
                Log("convert failed: " + ex);
                return;
            }
            finally
            {
                if (sigWs != null && !ReferenceEquals(sigWs, wsO)) Xl.Release(sigWs);
                if (sigWb != null && !ReferenceEquals(sigWb, wb)) Xl.Release(sigWb);
                Xl.Release(wsO); Xl.Release(wsI); Xl.Release(wsM);
                Xl.Release(wb); Xl.Release(app);
                if (filterOn) { try { RetryFilter.Revoke(); } catch { } }
            }

            _jobRunning = false;
            Log("convert DONE :: " + Get(Chan.Result) + ";sigMs=" + sigMs.ToString(CultureInfo.InvariantCulture));
        }

        // 通知セルの電文。Excel はこれ 1 つで全工程の値を受け取る。
        //   DONE|jobId|rows|hash|bindMs|masterMs|inputMs|dictMs|convMs|writeMs|workerE2eMs
        private static string SigText(string jobId, int rows, uint hash, long bindMs,
                                      long masterMs, long inputMs, long dictMs, long convMs,
                                      long writeMs, long workerE2eMs)
        {
            return "DONE|" + (jobId ?? "") +
                   "|" + rows.ToString(CultureInfo.InvariantCulture) +
                   "|" + hash.ToString("X8", CultureInfo.InvariantCulture) +
                   "|" + bindMs.ToString(CultureInfo.InvariantCulture) +
                   "|" + masterMs.ToString(CultureInfo.InvariantCulture) +
                   "|" + inputMs.ToString(CultureInfo.InvariantCulture) +
                   "|" + dictMs.ToString(CultureInfo.InvariantCulture) +
                   "|" + convMs.ToString(CultureInfo.InvariantCulture) +
                   "|" + writeMs.ToString(CultureInfo.InvariantCulture) +
                   "|" + workerE2eMs.ToString(CultureInfo.InvariantCulture);
        }

        private static string BuildPayload(int rows, int dictRows, int notFound,
                                           long bindMs, long masterMs, long inputMs,
                                           long dictMs, long convMs, long writeMs,
                                           long workerE2eMs, uint hash)
        {
            return "rows=" + rows.ToString(CultureInfo.InvariantCulture) +
                   ";dictRows=" + dictRows.ToString(CultureInfo.InvariantCulture) +
                   ";notFound=" + notFound.ToString(CultureInfo.InvariantCulture) +
                   ";bindMs=" + bindMs.ToString(CultureInfo.InvariantCulture) +
                   ";masterMs=" + masterMs.ToString(CultureInfo.InvariantCulture) +
                   ";inputMs=" + inputMs.ToString(CultureInfo.InvariantCulture) +
                   ";dictMs=" + dictMs.ToString(CultureInfo.InvariantCulture) +
                   ";convertMs=" + convMs.ToString(CultureInfo.InvariantCulture) +
                   ";writeMs=" + writeMs.ToString(CultureInfo.InvariantCulture) +
                   ";workerE2eMs=" + workerE2eMs.ToString(CultureInfo.InvariantCulture) +
                   ";hash=" + hash.ToString("X8", CultureInfo.InvariantCulture);
        }

        internal void Log(string msg)
        {
            if (string.IsNullOrEmpty(_logPath)) return;
            try
            {
                lock (_logPath)
                {
                    File.AppendAllText(_logPath,
                        DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture) + "  " + msg + Environment.NewLine,
                        new UTF8Encoding(false));
                }
            }
            catch { }
        }
    }

    // =======================================================================
    // 最小のホストウィンドウ。HWND を持って UIA ツリーに参加するためだけに在る。
    // =======================================================================
    internal sealed class WorkerForm : Form
    {
        private readonly WorkerCore _core;
        private readonly RootProvider _root;
        private readonly string _mode;
        private readonly int _parentPid;
        private readonly System.Windows.Forms.Timer _tick = new System.Windows.Forms.Timer();
        private int _heartCount;

        private static readonly TimeSpan IdleKill = TimeSpan.FromMinutes(20);

        internal WorkerForm(string token, string mode, int parentPid, string logPath)
        {
            _mode = mode;
            _parentPid = parentPid;

            _core = new WorkerCore(logPath);
            _root = new RootProvider(_core);
            _core.Root = _root;
            _core.Host = this;

            // ウィンドウタイトルが待ち合わせ鍵。トークンが違えば互いを見つけない。
            Text = "ZIPWORKER::" + token;

            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            ControlBox = false;
            StartPosition = FormStartPosition.Manual;
            ClientSize = new Size(320, 90);

            Process me = Process.GetCurrentProcess();
            _core.Set(Chan.Info,
                "pid=" + me.Id +
                ";mode=" + mode +
                ";clr=" + Environment.Version +
                ";bits=" + (IntPtr.Size * 8) +
                ";startedAt=" + me.StartTime.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture));

            _tick.Interval = 1000;
            _tick.Tick += OnTick;
            _tick.Start();
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(_mode == "hidden" ? false : value);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == Uia.WM_GETOBJECT && m.LParam.ToInt64() == Uia.UiaRootObjectId)
            {
                m.Result = Uia.UiaReturnRawElementProvider(m.HWnd, m.WParam, m.LParam, _root);
                return;
            }
            base.WndProc(ref m);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            _root.BindHwnd(Handle);
            ApplyMode();

            Uia.STARTUPINFO si;
            Uia.GetStartupInfo(out si);
            _core.Log("worker ready mode=" + _mode +
                      " pid=" + Process.GetCurrentProcess().Id +
                      " hwnd=0x" + Handle.ToInt64().ToString("X") +
                      " clientsListening=" + Uia.UiaClientsAreListening() +
                      " startupShow=" + si.wShowWindow +
                      " useShowWindow=" + ((si.dwFlags & Uia.STARTF_USESHOWWINDOW) != 0) +
                      " desktop=" + (si.lpDesktop == IntPtr.Zero
                                       ? "<null>" : Marshal.PtrToStringUni(si.lpDesktop)));
        }

        private void ApplyMode()
        {
            Rectangle vs = SystemInformation.VirtualScreen;
            switch (_mode)
            {
                case "normal":
                case "minimized":
                    Location = new Point(vs.Left + 40, vs.Top + 40);
                    break;
                case "hidden":
                case "offscreen":
                default:
                    Location = new Point(vs.Right + 200, vs.Bottom + 200);
                    break;
            }
        }

        // UIA ツリーに出るかどうかは「ウィンドウが可視か」で決まる。
        // 可視でないなら、そう記録する。後から推測しなくて済むように。
        internal void LogVisibility(string when)
        {
            _core.Log(when + ": IsWindowVisible=" + Uia.IsWindowVisible(Handle) +
                      " clientsListening=" + Uia.UiaClientsAreListening());
        }

        private void OnTick(object sender, EventArgs e)
        {
            if (_heartCount == 0) LogVisibility("first tick");
            _heartCount++;
            _core.Set(Chan.Heart, _heartCount.ToString(CultureInfo.InvariantCulture));

            // 親 (Excel) が消えたら道連れで終わる。孤児プロセスを残さない。
            if (_parentPid > 0)
            {
                try { Process.GetProcessById(_parentPid); }
                catch (ArgumentException)
                {
                    _core.Log("parent pid " + _parentPid + " gone; exiting");
                    _tick.Stop();
                    Application.ExitThread();
                    return;
                }
            }

            if (!_core.JobRunning && DateTime.UtcNow - _core.LastCommandUtc > IdleKill)
            {
                _core.Log("idle timeout; exiting");
                _tick.Stop();
                Application.ExitThread();
            }
        }
    }

    internal static class Program
    {
        [STAThread]
        private static int Main(string[] argv)
        {
            string token = "", mode = "offscreen", log = "";
            int parentPid = 0;

            for (int i = 0; i < argv.Length - 1; i++)
            {
                switch (argv[i].ToLowerInvariant())
                {
                    case "--token":     token = argv[++i]; break;
                    case "--mode":      mode = argv[++i].ToLowerInvariant(); break;
                    case "--log":       log = argv[++i]; break;
                    case "--parentpid": int.TryParse(argv[++i], out parentPid); break;
                }
            }
            if (token.Length == 0) return 2;

            Application.EnableVisualStyles();
            WorkerForm f = new WorkerForm(token, mode, parentPid, log);

            if (mode == "hidden")
            {
                IntPtr force = f.Handle;
                GC.KeepAlive(force);
                Application.Run();
            }
            else
            {
                f.Show();

                // 起動側から SW_HIDE を渡されていても、ここで見える状態に戻す。
                //
                // Windows は STARTUPINFO の wShowWindow を「そのプロセスが最初に出す
                // ウィンドウ」に適用する。Shell.Application.ShellExecute を vShow=0 で
                // 呼ぶと .bat が SW_HIDE で走り、cmd の start /b がそれを子へ引き継ぐ。
                // すると WinForms の Show() は効かず、ハンドルはあるのに一度も可視に
                // ならない。可視にならなかったウィンドウは UIA ツリーに出ない
                // (poc-reference の verification でも同じことを実測している)。
                //
                // 起動手段を比べるベンチマークで「どう起こしたか」が「チャネルに
                // 届くかどうか」を左右してしまっては、比較にならない。だから
                // 表示状態はワーカー自身が決める。
                Uia.ShowWindow(f.Handle,
                    mode == "minimized" ? Uia.SW_SHOWMINNOACTIVE : Uia.SW_SHOWNOACTIVATE);

                f.LogVisibility("after Show+ShowWindow");
                Application.Run(f);
            }
            return 0;
        }
    }
}
