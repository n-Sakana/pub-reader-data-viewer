// Execution log: serialize writes and report a write failure once.
using System;
using System.Globalization;
using System.IO;
using System.Text;

public sealed class Rdv3Log
{
    private readonly string path;
    private readonly object gate = new object();
    private bool failed;
    public Action<string> OnFail;

    public Rdv3Log(string p)
    {
        path = p;
    }

    public string Path { get { return path; } }

    public void Write(string runId, string section, string detail)
    {
        StringBuilder sb = new StringBuilder(160);
        sb.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture));
        sb.Append('\t').Append(runId);
        sb.Append('\t').Append(section);
        sb.Append('\t').Append(detail.Replace('\t', ' ').Replace('\r', ' ').Replace('\n', ' '));
        sb.Append("\r\n");
        lock (gate)
        {
            try
            {
                File.AppendAllText(path, sb.ToString(), new UTF8Encoding(false));
                failed = false;
            }
            catch (Exception ex)
            {
                // a log that cannot be written must not crash the app, but it
                // must not fail silently either: surface it once
                if (!failed)
                {
                    failed = true;
                    Action<string> h = OnFail;
                    if (h != null) { h(ex.Message); }
                }
            }
        }
    }

    public static string F(double ms)
    {
        return ms.ToString("F2", CultureInfo.InvariantCulture);
    }
}
