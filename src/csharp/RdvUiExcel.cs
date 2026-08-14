// ============================================================================
// RdvUiExcel.cs -- the Excel front for the hybrid build, driven over COM from
// the same engine the C#-only build uses.
//
// Ownership rule, taken from benchmarks\excel-background-bench: this process
// starts its OWN Excel and touches nothing else. No GetActiveObject, no running
// object table. The instance is identified by the window handle Excel itself
// reports, so even an identically named Excel started by the user next door is
// invisible to this code, and only the pid behind that handle is ever closed.
//
// The workbook is opened read-only. Cells are written for display, which a
// read-only workbook allows, but the file itself can never be saved over.
//
// Late binding through reflection on purpose: no interop assembly is shipped,
// and Add-Type has none to reference.
// ============================================================================

using System;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;

public static class Com
{
    // Excel says "call me later" while it is repainting or a cell is in edit
    // mode. Retry the same call rather than pretend the write happened.
    private const int RetryLater = unchecked((int)0x8001010A);
    private const int CallRejected = unchecked((int)0x80010001);

    private static object Invoke(object o, string name, BindingFlags f, object[] args)
    {
        for (int attempt = 0; ; attempt++)
        {
            try
            {
                return o.GetType().InvokeMember(name, f, null, o, args);
            }
            catch (TargetInvocationException tie)
            {
                COMException ce = tie.InnerException as COMException;
                if (ce != null && (ce.ErrorCode == RetryLater || ce.ErrorCode == CallRejected) && attempt < 100)
                {
                    Thread.Sleep(20);
                    continue;
                }
                // name the member: "member not found" with no member named is
                // the least useful error message in COM automation
                throw new InvalidOperationException("Excel." + name + " -> "
                    + (tie.InnerException != null ? tie.InnerException.Message : tie.Message), tie);
            }
            catch (COMException ce)
            {
                if ((ce.ErrorCode == RetryLater || ce.ErrorCode == CallRejected) && attempt < 100)
                {
                    Thread.Sleep(20);
                    continue;
                }
                throw new InvalidOperationException("Excel." + name + " -> " + ce.Message, ce);
            }
            catch (MissingMemberException mme)
            {
                throw new InvalidOperationException("Excel." + name + " -> " + mme.Message, mme);
            }
        }
    }

    public static object Get(object o, string name, params object[] args)
    {
        return Invoke(o, name, BindingFlags.GetProperty, args);
    }

    public static void Set(object o, string name, params object[] args)
    {
        Invoke(o, name, BindingFlags.SetProperty, args);
    }

    // Method and parameterised-property get in one flag set, the way VB's own
    // late binding does it. Excel's Worksheets("VIEW") is propget only, and
    // asking for DISPATCH_METHOD alone comes back DISP_E_MEMBERNOTFOUND.
    public static object Call(object o, string name, params object[] args)
    {
        return Invoke(o, name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, args);
    }

    public static void Release(object o)
    {
        if (o == null) { return; }
        try { Marshal.ReleaseComObject(o); }
        catch (Exception) { }
    }
}

public sealed class RdvExcelFront
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int pid);

    private object app;
    private object books;
    private object book;
    private object sheet;
    private int excelPid = -1;

    public int ExcelPid { get { return excelPid; } }

    public void Start(string workbookPath)
    {
        Type t = Type.GetTypeFromProgID("Excel.Application");
        if (t == null) { throw new InvalidOperationException("Excel.Application"); }
        app = Activator.CreateInstance(t);

        // find out which process this is, from the handle Excel hands back
        object h = Com.Get(app, "Hwnd");
        int pid;
        GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(h, CultureInfo.InvariantCulture)), out pid);
        excelPid = pid;

        Com.Set(app, "DisplayAlerts", false);
        Com.Set(app, "AskToUpdateLinks", false);
        Com.Set(app, "ScreenUpdating", true);

        books = Com.Get(app, "Workbooks");
        book = Com.Call(books, "Open", workbookPath, 0, true);
        object sheets = Com.Get(book, "Worksheets");
        sheet = Com.Get(sheets, "Item", RdvSheet.Name);
        Com.Release(sheets);
        Com.Call(sheet, "Activate");
        Com.Set(app, "Visible", true);
        // WindowState is a property, not a method: calling it as one throws.
        // Cosmetic either way, so it never fails the startup.
        try { Com.Set(app, "WindowState", -4137); }
        catch (Exception) { }
    }

    public static string Explain(Exception ex)
    {
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        Exception e = ex;
        int depth = 0;
        while (e != null && depth < 5)
        {
            if (sb.Length > 0) { sb.Append(" <- "); }
            sb.Append(e.GetType().Name).Append(": ").Append(e.Message);
            e = e.InnerException;
            depth++;
        }
        return sb.ToString();
    }

    public bool Alive()
    {
        try
        {
            object v = Com.Get(app, "Visible");
            return v != null;
        }
        catch (Exception) { return false; }
    }

    private void Put(string a1, object[,] v)
    {
        object r = Com.Get(sheet, "Range", a1);
        Com.Set(r, "Value2", v);
        Com.Release(r);
    }

    private void Put1(string a1, object v)
    {
        object r = Com.Get(sheet, "Range", a1);
        Com.Set(r, "Value2", v);
        Com.Release(r);
    }

    public void Header(string state, string notepad, string data)
    {
        object[,] v = new object[3, 1];
        v[0, 0] = state; v[1, 0] = notepad; v[2, 0] = data;
        Put(RdvSheet.HeaderRange, v);
    }

    // stage 7. Two range writes and one read back: the read is what proves the
    // values are in Excel's own model rather than sitting in an RPC queue.
    public void ShowRecord(RdvRun run)
    {
        RdvHit h = run.Hit;
        object[,] top = new object[1, 3];
        top[0, 0] = run.Key;
        top[0, 1] = "";
        if (run.Error.Length > 0) { top[0, 2] = RdvText.VerdictError; }
        else if (h != null && h.Found) { top[0, 2] = RdvText.VerdictHit + "   key2 = " + h.Key2; }
        else { top[0, 2] = RdvText.VerdictMiss; }

        object[,] rec = new object[RdvSpec.Fields, 7];
        for (int i = 0; i < RdvSpec.Fields; i++)
        {
            rec[i, 0] = i + 1;
            rec[i, 1] = Pick(h == null ? null : h.NameA, i);
            rec[i, 2] = Pick(h == null ? null : h.ValA, i);
            rec[i, 3] = Pick(h == null ? null : h.NameB, i);
            rec[i, 4] = Pick(h == null ? null : h.ValB, i);
            rec[i, 5] = Pick(h == null ? null : h.NameC, i);
            rec[i, 6] = Pick(h == null ? null : h.ValC, i);
        }
        Put(RdvSheet.KeyRange, top);
        Put(RdvSheet.RecordRange, rec);
        object probe = Com.Get(sheet, "Range", RdvSheet.CommitCell);
        object back = Com.Get(probe, "Value2");
        Com.Release(probe);
        GC.KeepAlive(back);
    }

    private static string Pick(string[] a, int i)
    {
        if (a == null || i >= a.Length) { return ""; }
        return a[i];
    }

    public void ShowStats(RdvRun run)
    {
        int n = RdvSpec.StageCount + 5;
        object[,] st = new object[n, 3];
        double total = run.TotalMs;
        int row = 0;
        for (int i = 0; i < RdvSpec.StageCount; i++)
        {
            st[row, 0] = RdvText.StageName[i];
            st[row, 1] = run.Stage[i];
            st[row, 2] = total > 0 ? (100.0 * run.Stage[i] / total) : 0.0;
            row++;
        }
        st[row, 0] = RdvText.StageOther;
        st[row, 1] = total - run.StageSum();
        st[row, 2] = total > 0 ? (100.0 * (total - run.StageSum()) / total) : 0.0;
        row++;
        st[row, 0] = RdvText.StageTotal; st[row, 1] = total; st[row, 2] = 100.0; row++;
        st[row, 0] = RdvText.StageDetect; st[row, 1] = run.DetectMs; st[row, 2] = ""; row++;
        st[row, 0] = RdvText.StageRows; st[row, 1] = run.Rows; st[row, 2] = ""; row++;
        st[row, 0] = RdvText.StageProbes; st[row, 1] = run.Probes; st[row, 2] = "";
        Put(RdvSheet.StageRange, st);

        object[,] hist = new object[1, 6];
        hist[0, 0] = run.Seq;
        hist[0, 1] = run.When.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
        hist[0, 2] = run.Key;
        hist[0, 3] = run.TotalMs;
        hist[0, 4] = run.DetectMs;
        hist[0, 5] = (run.Error.Length > 0) ? RdvText.VerdictError
                   : (run.OracleOk ? RdvText.OracleOk : (RdvText.OracleBad + " " + run.OracleNote));
        int line = ((run.Seq - 1) % RdvSheet.HistoryRows) + RdvSheet.HistoryTop;
        Put(RdvSheet.HistoryCol + line.ToString(CultureInfo.InvariantCulture) + ":"
            + RdvSheet.HistoryColEnd + line.ToString(CultureInfo.InvariantCulture), hist);

        Put1(RdvSheet.ErrorCell, run.Error);
    }

    // close down only what this process started
    public void Shutdown()
    {
        try
        {
            if (book != null)
            {
                try { Com.Set(book, "Saved", true); }
                catch (Exception) { }
                Com.Call(book, "Close", false);
            }
        }
        catch (Exception) { }
        try { if (app != null) { Com.Call(app, "Quit"); } }
        catch (Exception) { }
        Com.Release(sheet); Com.Release(book); Com.Release(books); Com.Release(app);
        sheet = null; book = null; books = null; app = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();

        if (excelPid > 0)
        {
            for (int i = 0; i < 25; i++)
            {
                try
                {
                    System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(excelPid);
                    if (p.HasExited) { return; }
                    Thread.Sleep(200);
                    if (i == 24) { p.Kill(); }     // only ever this pid
                }
                catch (ArgumentException) { return; }
                catch (Exception) { return; }
            }
        }
    }
}

// The one place the sheet layout is written down. build\build_workbooks.ps1
// paints the same addresses; if they drift apart the numbers land in the wrong
// boxes, so both sides read this list.
public static class RdvSheet
{
    public const string Name = "VIEW";
    public const string HeaderRange = "C2:C4";
    public const string KeyRange = "C6:E6";
    public const string RecordRange = "B9:H18";
    public const string CommitCell = "B9";
    public const string StageRange = "J9:L20";
    public const string HistoryCol = "B";
    public const string HistoryColEnd = "G";
    public const int HistoryTop = 22;
    public const int HistoryRows = 20;
    public const string ErrorCell = "C44";
}
