// ============================================================================
// ComWriteProbe.cs -- DEVELOPMENT TOOL, not part of the benchmark.
//
// Answers one question before anything is designed around it:
//
//     how long does it take a separate process to push 1,000,000 strings into
//     an Excel column with Range.Value2, and what chunk size is fastest?
//
// The old design had the worker write a CSV and Excel read it back, which cost
// 9.5 s at 1M rows. Replacing that with a cross-process COM write is only worth
// doing if the COM write is materially faster. So measure it first.
//
// It also proves the instance-binding approach the real worker will use:
// bind to ONE named Excel instance through its window handle, never
// GetActiveObject / ROT, which would be free to hand back somebody else's Excel.
//
//   ComWriteProbe.exe --hwnd <Application.Hwnd> --rows 1000000
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;

internal static class ComWriteProbe
{
    private const int OBJID_NATIVEOM = unchecked((int)0xFFFFFFF0);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, int dwObjectID, ref Guid riid,
        [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string cls, string title);

    // Excel's object model hangs off the EXCEL7 pane inside XLDESK inside the
    // main XLMAIN window. Walking that chain reaches exactly the instance whose
    // handle we were given.
    private static object ExcelAppFromHwnd(IntPtr xlMain)
    {
        IntPtr desk = FindWindowEx(xlMain, IntPtr.Zero, "XLDESK", null);
        if (desk == IntPtr.Zero) throw new Exception("XLDESK not found under the given hwnd");
        IntPtr pane = FindWindowEx(desk, IntPtr.Zero, "EXCEL7", null);
        if (pane == IntPtr.Zero) throw new Exception("EXCEL7 not found; the instance has no open workbook window");

        Guid iid = new Guid("00020400-0000-0000-C000-000000000046");   // IID_IDispatch
        object win;
        int hr = AccessibleObjectFromWindow(pane, OBJID_NATIVEOM, ref iid, out win);
        if (hr != 0 || win == null) throw new Exception("AccessibleObjectFromWindow failed hr=0x" + hr.ToString("X8"));

        // win is an Excel.Window; .Application is the instance we want
        return win.GetType().InvokeMember("Application",
            System.Reflection.BindingFlags.GetProperty, null, win, null);
    }

    // Excel exposes Cells/Range as parameterised PROPERTIES, not methods, so a
    // plain InvokeMethod gives DISP_E_MEMBERNOTFOUND. Asking for both at once
    // lets IDispatch pick whichever the member really is.
    private static object Get(object o, string name, params object[] args)
    {
        return o.GetType().InvokeMember(name,
            System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.InvokeMethod,
            null, o, args);
    }
    private static object Call(object o, string name, params object[] args)
    {
        return o.GetType().InvokeMember(name,
            System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.GetProperty,
            null, o, args);
    }
    private static void Set(object o, string name, params object[] args)
    {
        o.GetType().InvokeMember(name, System.Reflection.BindingFlags.SetProperty, null, o, args);
    }

    private static int Main(string[] argv)
    {
        long hwnd = 0;
        int rows = 1000000;
        for (int i = 0; i < argv.Length - 1; i++)
        {
            if (argv[i] == "--hwnd") long.TryParse(argv[++i], out hwnd);
            else if (argv[i] == "--rows") int.TryParse(argv[++i], out rows);
        }
        if (hwnd == 0) { Console.Error.WriteLine("--hwnd is required"); return 2; }

        object app = ExcelAppFromHwnd(new IntPtr(hwnd));
        Console.WriteLine("bound to Excel: version=" + Get(app, "Version") + " hwnd=" + hwnd);

        // make a scratch workbook inside that instance
        object books = Get(app, "Workbooks"); Console.WriteLine("  got Workbooks");
        object wb = Call(books, "Add");       Console.WriteLine("  added workbook");
        object sheets = Get(wb, "Worksheets");
        object ws = Get(sheets, "Item", 1);   Console.WriteLine("  got sheet");
        Set(app, "ScreenUpdating", false);
        Set(app, "Calculation", -4135);      // xlCalculationManual
        Set(app, "EnableEvents", false);
        Console.WriteLine("  app state set");

        string[] data = new string[rows];
        for (int i = 0; i < rows; i++) data[i] = "北海道札幌市中央区大通西" + (i % 1000);

        // object[,] marshals as a VT_VARIANT SAFEARRAY: every element carries a
        // 16-byte VARIANT header around the BSTR. string[,] marshals as VT_BSTR
        // and skips all of that. Whether Excel accepts it, and whether it is
        // actually faster, is the thing to find out.
        foreach (bool typed in new bool[] { false, true })
        {
            foreach (int chunk in new int[] { rows, 100000, 25000 })
            {
                Call(Get(ws, "Cells"), "Clear");
                Stopwatch sw = Stopwatch.StartNew();
                int written = 0;
                bool ok = true;
                while (written < rows)
                {
                    int take = Math.Min(chunk, rows - written);
                    object buf;
                    if (typed)
                    {
                        string[,] b = new string[take, 1];
                        for (int i = 0; i < take; i++) b[i, 0] = data[written + i];
                        buf = b;
                    }
                    else
                    {
                        object[,] b = new object[take, 1];
                        for (int i = 0; i < take; i++) b[i, 0] = data[written + i];
                        buf = b;
                    }

                    object c1 = Get(ws, "Cells", written + 1, 1);
                    object c2 = Get(ws, "Cells", written + take, 1);
                    object rng = Get(ws, "Range", c1, c2);
                    try { Set(rng, "Value2", buf); }
                    catch (Exception ex) { ok = false; Console.WriteLine("    " + (typed ? "string[,]" : "object[,]") + " rejected: " + ex.GetType().Name); }
                    Marshal.ReleaseComObject(rng);
                    Marshal.ReleaseComObject(c1);
                    Marshal.ReleaseComObject(c2);
                    if (!ok) break;
                    written += take;
                }
                sw.Stop();
                if (!ok) break;
                Console.WriteLine(string.Format(CultureInfo.InvariantCulture,
                    "{0,-10} chunk {1,7} -> {2,7} ms  ({3} calls)",
                    typed ? "string[,]" : "object[,]", chunk, sw.ElapsedMilliseconds, (rows + chunk - 1) / chunk));
            }
        }

        // read one value back so the write is proven, not just timed
        object probe = Call(ws, "Cells", rows, 1);
        Console.WriteLine("last cell = " + Get(probe, "Value2"));

        Set(wb, "Saved", true);
        Call(wb, "Close", false);
        Set(app, "ScreenUpdating", true);
        Set(app, "EnableEvents", true);
        return 0;
    }
}
