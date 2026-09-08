using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

public static class LoggingFaults
{
    [DllImport("kernel32.dll")]
    private static extern uint SetErrorMode(uint mode);

    public static void Main(string[] args)
    {
        SetErrorMode(2);
        string mode = args[0];
        string token = args[1];
        if (mode == "append")
        {
            for (int i = 0; i < 250; i++) { Rdv3Log.AppendBounded(args[2], token + ":" + i + "\r\n"); }
            return;
        }
        if (mode == "rotate")
        {
            for (int i = 0; i < 10; i++)
            { Rdv3Log.AppendBounded(args[2], token + ":" + i + ":" + new string('x', Rdv3Log.MaxBytes / 2) + "\r\n"); }
            return;
        }
        if (mode == "fallback")
        {
            using (FileStream held = new FileStream(Rdv3Log.FeedbackPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
            { Rdv3Log.Feedback("FIXTURE FALLBACK", token); }
            return;
        }
        if (mode == "bad-destination") { new Rdv3Log(args[2]).Write(token, "test", "configured destination is a directory"); return; }
        Rdv3Log.Begin(mode, token);
        Rdv3Log.Phase(token + " " + mode);
        if (mode == "silent") { Environment.Exit(0); }
        if (mode == "throw")
        {
            new Thread(delegate() { throw new InvalidOperationException(token + " worker crash"); }).Start();
            Thread.Sleep(Timeout.Infinite);
        }
        if (mode == "wait") { Thread.Sleep(12000); }
        Rdv3Log.End(0);
    }
}
