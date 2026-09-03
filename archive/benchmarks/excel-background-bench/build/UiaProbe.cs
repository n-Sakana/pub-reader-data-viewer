// ============================================================================
// UiaProbe.cs -- DEVELOPMENT TOOL, not part of the benchmark.
//
// A tiny UI Automation CLIENT that talks to a running ZipWorker.exe exactly the
// way the Excel VBA client does: find the window by Name, write UIAW_CMD with
// ValuePattern, fire UIAW_SUBMIT with InvokePattern, then read UIAW_STATUS and
// UIAW_RESULT back.
//
// It exists so the worker and the channel can be proven from a console, in
// seconds, before a single line of VBA is written. Nothing in ZipBench depends
// on it at run time.
//
//   UiaProbe.exe --token T [--wait 120]
//   UiaProbe.exe --token T --dict D --in I --out O
//   UiaProbe.exe --token T --verb PING|CANCEL|SHUTDOWN
// ============================================================================

using System;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Windows.Automation;

internal static class UiaProbe
{
    private const string WindowPrefix = "ZIPWORKER::";

    private static int Main(string[] argv)
    {
        string token = "", dict = "", inp = "", outp = "", verb = "CONVERT";
        int waitSec = 180;

        for (int i = 0; i < argv.Length - 1; i++)
        {
            switch (argv[i].ToLowerInvariant())
            {
                case "--token": token = argv[++i]; break;
                case "--dict":  dict = argv[++i]; break;
                case "--in":    inp = argv[++i]; break;
                case "--out":   outp = argv[++i]; break;
                case "--verb":  verb = argv[++i].ToUpperInvariant(); break;
                case "--wait":  int.TryParse(argv[++i], out waitSec); break;
            }
        }
        if (token.Length == 0) { Console.Error.WriteLine("--token is required"); return 2; }

        Stopwatch findSw = Stopwatch.StartNew();
        AutomationElement win = null;
        while (findSw.Elapsed.TotalSeconds < waitSec)
        {
            win = AutomationElement.RootElement.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.NameProperty, WindowPrefix + token));
            if (win != null) break;
            Thread.Sleep(20);
        }
        findSw.Stop();
        if (win == null) { Console.Error.WriteLine("worker not found within " + waitSec + "s"); return 3; }
        Console.WriteLine("found      : " + findSw.ElapsedMilliseconds + " ms");

        AutomationElement cmd = Field(win, "UIAW_CMD");
        AutomationElement sub = Field(win, "UIAW_SUBMIT");
        AutomationElement ack = Field(win, "UIAW_ACK");
        AutomationElement st = Field(win, "UIAW_STATUS");
        AutomationElement res = Field(win, "UIAW_RESULT");
        AutomationElement info = Field(win, "UIAW_INFO");
        AutomationElement prog = Field(win, "UIAW_PROGRESS");
        if (cmd == null || sub == null || ack == null || st == null || res == null)
        {
            Console.Error.WriteLine("channel fields missing"); return 4;
        }
        Console.WriteLine("info       : " + Read(info));

        string arg = "";
        if (verb == "CONVERT")
        {
            if (dict.Length == 0 || inp.Length == 0 || outp.Length == 0)
            {
                Console.Error.WriteLine("CONVERT needs --dict --in --out"); return 5;
            }
            arg = "dict=" + dict + ";in=" + inp + ";out=" + outp;
        }

        Stopwatch sw = Stopwatch.StartNew();
        long seq = 1;
        ((ValuePattern)cmd.GetCurrentPattern(ValuePattern.Pattern))
            .SetValue(seq.ToString(CultureInfo.InvariantCulture) + "|" + verb + "|" + arg);
        ((InvokePattern)sub.GetCurrentPattern(InvokePattern.Pattern)).Invoke();

        while (sw.Elapsed.TotalSeconds < 5 && Read(ack) != seq.ToString(CultureInfo.InvariantCulture)) { }
        Console.WriteLine("acked      : " + sw.ElapsedMilliseconds + " ms");

        if (verb != "CONVERT")
        {
            Console.WriteLine("status     : " + Read(st));
            Console.WriteLine("result     : " + Read(res));
            return 0;
        }

        string status = "";
        while (sw.Elapsed.TotalSeconds < waitSec)
        {
            status = Read(st);
            if (status == "DONE" || status == "ERROR" || status == "CANCELLED") break;
            Thread.Sleep(5);
        }
        sw.Stop();
        Console.WriteLine("progress   : " + Read(prog));
        Console.WriteLine("status     : " + status);
        Console.WriteLine("result     : " + Read(res));
        Console.WriteLine("wallMs     : " + sw.ElapsedMilliseconds);
        return status == "DONE" ? 0 : 1;
    }

    private static AutomationElement Field(AutomationElement win, string id)
    {
        return win.FindFirst(TreeScope.Descendants,
            new PropertyCondition(AutomationElement.AutomationIdProperty, id));
    }

    private static string Read(AutomationElement el)
    {
        if (el == null) return "";
        try { return ((ValuePattern)el.GetCurrentPattern(ValuePattern.Pattern)).Current.Value; }
        catch (Exception ex) { return "<read error " + ex.GetType().Name + ">"; }
    }
}
