// ============================================================================
// Rdv3Settings.cs -- WebView settings dialog and UI Automation target picker.
// C# 5, ASCII source.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Threading;

public static class Rdv3SettingsForm
{
    public static Rdv3Config Edit(Rdv3Form owner, Rdv3Config current)
    {
        Rdv3Config working = current.Clone();
        Rdv3Target before = working.Targets.Count == 0 ? null : working.Targets[0];
        owner.TakePickedTarget();
        owner.SetPickerPollMs(working.PollMs);
        StringBuilder sb = new StringBuilder(4096);
        sb.Append("{\"title\":").Append(Rdv3WebJson.Q(Rdv3Text.SettingsTitle));
        sb.Append(",\"hint\":").Append(Rdv3WebJson.Q(Rdv3Text.SettingsHint));
        sb.Append(",\"dataDir\":").Append(Rdv3WebJson.Q(working.DataDir));
        sb.Append(",\"ledger\":").Append(Rdv3WebJson.Q(working.Ledger));
        sb.Append(",\"log\":").Append(Rdv3WebJson.Q(working.Log));
        sb.Append(",\"pattern\":").Append(Rdv3WebJson.Q(working.KeyPattern));
        sb.Append(",\"candidateRows\":").Append(working.CandidateRowsShown.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"target\":").Append(TargetJson(before, working.PollMs)).Append('}');
        Rdv3Json result = owner.ShowModal("settings", sb.ToString());
        if (!Rdv3Form.Flag(result, "ok", false))
        {
            owner.TakePickedTarget();
            return null;
        }

        string dataDir = Rdv3Form.Text(result, "dataDir").Trim();
        string ledger = Rdv3Form.Text(result, "ledger").Trim();
        string log = Rdv3Form.Text(result, "log").Trim();
        string pattern = Rdv3Form.Text(result, "pattern").Trim();
        int candidateRows = Rdv3Form.Number(result, "candidateRows", -1);
        string patternError = Rdv3Config.PatternError(pattern);
        if (patternError != null)
        {
            owner.Error(Rdv3Text.ErrPatternTyped + patternError);
            owner.TakePickedTarget();
            return null;
        }
        if (dataDir.Length == 0 || ledger.Length == 0 || log.Length == 0)
        {
            owner.Error(Rdv3Text.ErrPathBlank);
            owner.TakePickedTarget();
            return null;
        }
        if (candidateRows < 1 || candidateRows > 1000)
        {
            owner.Error(Rdv3Text.LblCandidateRows);
            owner.TakePickedTarget();
            return null;
        }
        working.DataDir = dataDir;
        working.Ledger = ledger;
        working.Log = log;
        working.KeyPattern = pattern;
        working.CandidateRowsShown = candidateRows;

        Rdv3Target picked = owner.TakePickedTarget();
        if (picked != null)
        {
            if (before != null)
            {
                if (before.Name.Length > 0 && !before.Name.StartsWith(Rdv3Text.SecTarget))
                {
                    picked.Name = before.Name;
                }
                picked.Enabled = before.Enabled;
                working.Targets[0] = picked;
            }
            else { working.Targets.Add(picked); }
        }
        return working;
    }

    public static string TargetJson(Rdv3Target target)
    {
        return TargetJson(target, 40);
    }

    public static string TargetJson(Rdv3Target target, int pollMs)
    {
        string summary;
        string read;
        if (target == null)
        {
            summary = Rdv3Text.NoteNoTargetShort;
            read = Rdv3Text.NoValue;
        }
        else
        {
            string name = target.Name.Length > 0 ? target.Name : Rdv3Text.NoValue;
            string kind = target.Window.ClassName.Length > 0 ? "className"
                : target.Window.AutomationId.Length > 0 ? "automationId" : "name";
            string value = target.Window.ClassName.Length > 0 ? target.Window.ClassName
                : target.Window.AutomationId.Length > 0 ? target.Window.AutomationId : target.Window.NameLike;
            if (value.Length == 0) { value = Rdv3Text.NoValue; }
            summary = Rdv3Text.NoteTargetSummary.Replace("{name}", name)
                .Replace("{kind}", kind).Replace("{value}", value);
            string mode = target.ReadMode == Rdv3Uia.ReadValue ? Rdv3Text.ReadValuePattern
                : target.ReadMode == Rdv3Uia.ReadText ? Rdv3Text.ReadTextPattern
                : Rdv3Text.ReadNameProperty;
            read = Rdv3Text.ReadSummaryFmt.Replace("{mode}", mode)
                .Replace("{poll}", pollMs.ToString(CultureInfo.InvariantCulture));
        }
        return "{\"summary\":" + Rdv3WebJson.Q(summary) +
            ",\"read\":" + Rdv3WebJson.Q(read) + "}";
    }
}

public sealed class Rdv3PickerForm
{
    private const int VkControl = 0x11;
    private const int VkShift = 0x10;
    private const int VkEscape = 0x1b;

    private readonly Rdv3Form owner;
    private readonly DispatcherTimer timer = new DispatcherTimer();
    private readonly DispatcherFrame frame = new DispatcherFrame();
    private AutomationElement hover;
    private bool armed;
    private Rdv3Target result;
    private static Rdv3PickerForm current;

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out NativePoint point);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int key);

    private Rdv3PickerForm(Rdv3Form form)
    {
        owner = form;
        timer.Interval = TimeSpan.FromMilliseconds(90);
        timer.Tick += OnTick;
    }

    public static Rdv3Target Pick(Rdv3Form owner)
    {
        if (ReaderDataViewer.App.IsProbe) { return null; }
        Rdv3PickerForm picker = new Rdv3PickerForm(owner);
        current = picker;
        picker.timer.Start();
        Dispatcher.PushFrame(picker.frame);
        picker.timer.Stop();
        picker.timer.Tick -= picker.OnTick;
        current = null;
        return picker.result;
    }

    public static void CancelCurrent()
    {
        if (current != null) { current.frame.Continue = false; }
    }

    private void OnTick(object sender, EventArgs eventArgs)
    {
        if (Down(VkEscape))
        {
            frame.Continue = false;
            return;
        }
        NativePoint native;
        if (!GetCursorPos(out native)) { return; }
        Point point = new Point(native.X, native.Y);
        if (owner.CardBounds.Contains(point)) { return; }
        AutomationElement element = null;
        try { element = AutomationElement.FromPoint(point); }
        catch (Exception) { }
        if (element != null && !Rdv3Uia.Same(element, hover))
        {
            hover = element;
            SendPreview(element);
        }
        bool chord = Down(VkControl) && Down(VkShift);
        if (chord && !armed && hover != null)
        {
            armed = true;
            result = Take(hover);
            if (result != null) { frame.Continue = false; }
        }
        else if (!chord) { armed = false; }
    }

    private void SendPreview(AutomationElement element)
    {
        try
        {
            AutomationElement.AutomationElementInformation current = element.Current;
            string read = Rdv3Uia.ReadValueOf(element, Rdv3Uia.ReadValue);
            string mode = Rdv3Text.ReadValuePattern;
            if (read == null)
            {
                read = Rdv3Uia.ReadValueOf(element, Rdv3Uia.ReadText);
                mode = Rdv3Text.ReadTextPattern;
            }
            if (read == null)
            {
                read = current.Name;
                mode = Rdv3Text.ReadNameProperty;
            }
            string json = "{\"type\":" + Rdv3WebJson.Q(Rdv3Uia.ControlTypeName(current.ControlType)) +
                ",\"automationId\":" + Rdv3WebJson.Q(current.AutomationId) +
                ",\"className\":" + Rdv3WebJson.Q(current.ClassName) +
                ",\"name\":" + Rdv3WebJson.Q(Cut(current.Name, 46)) +
                ",\"process\":" + Rdv3WebJson.Q(ProcessName(current.ProcessId)) +
                ",\"read\":" + Rdv3WebJson.Q(mode + " / " + Cut(Rdv3Watch.Candidate(read), 46)) + "}";
            owner.PostPickerPreview(json);
        }
        catch (Exception) { }
    }

    private static bool Down(int key)
    {
        return (GetAsyncKeyState(key) & 0x8000) != 0;
    }

    private static Rdv3Target Take(AutomationElement element)
    {
        try
        {
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement root = AutomationElement.RootElement;
            List<AutomationElement> chain = new List<AutomationElement>();
            AutomationElement at = element;
            for (int guard = 0; guard < 64 && at != null; guard++)
            {
                AutomationElement parent = walker.GetParent(at);
                if (parent == null || Rdv3Uia.Same(parent, root)) { break; }
                chain.Add(parent);
                at = parent;
            }
            chain.Reverse();
            AutomationElement window = chain.Count > 0 ? chain[0] : element;
            Rdv3Target target = new Rdv3Target();
            target.Window = new Rdv3Match();
            target.Window.Descendants = false;
            target.Window.ClassName = window.Current.ClassName;
            target.Window.ProcessName = ProcessName(window.Current.ProcessId);
            if (window.Current.AutomationId.Length > 0)
            {
                target.Window.AutomationId = window.Current.AutomationId;
            }
            target.Name = window.Current.Name.Length > 0
                ? Cut(window.Current.Name, 24) : target.Window.ProcessName;
            for (int i = 1; i < chain.Count; i++)
            {
                string id = chain[i].Current.AutomationId;
                if (id.Length == 0) { continue; }
                Rdv3Match step = new Rdv3Match();
                step.AutomationId = id;
                step.Descendants = true;
                target.Steps.Add(step);
            }
            target.Field = new Rdv3Match();
            target.Field.Descendants = true;
            AutomationElement.AutomationElementInformation current = element.Current;
            if (current.AutomationId.Length > 0) { target.Field.AutomationId = current.AutomationId; }
            else if (current.ClassName.Length > 0) { target.Field.ClassName = current.ClassName; }
            string controlType = Rdv3Uia.ControlTypeName(current.ControlType);
            if (controlType.Length > 0) { target.Field.ControlTypes = new string[] { controlType }; }
            if (target.Field.IsEmpty && current.Name.Length > 0) { target.Field.Name = current.Name; }
            target.ReadMode = Rdv3Uia.ReadValue;
            if (Rdv3Uia.ReadValueOf(element, Rdv3Uia.ReadValue) == null)
            {
                target.ReadMode = Rdv3Uia.ReadValueOf(element, Rdv3Uia.ReadText) != null
                    ? Rdv3Uia.ReadText : Rdv3Uia.ReadName;
            }
            if (target.ReadMode == Rdv3Uia.ReadValue) { target.Field.RequireValuePattern = true; }
            AutomationElement anchor = window;
            for (int i = 1; i < chain.Count; i++)
            {
                if (chain[i].Current.AutomationId.Length > 0) { anchor = chain[i]; }
            }
            List<AutomationElement> same = Rdv3Uia.FindAll(
                anchor,
                target.Field,
                target.Field.Descendants);
            for (int i = 0; i < same.Count; i++)
            {
                if (Rdv3Uia.Same(same[i], element)) { target.Field.Index = i; break; }
            }
            return target;
        }
        catch (Exception) { return null; }
    }

    private static string ProcessName(int pid)
    {
        try { return Process.GetProcessById(pid).ProcessName; }
        catch (Exception) { return ""; }
    }

    private static string Cut(string text, int count)
    {
        if (text == null) { return ""; }
        text = text.Replace("\r", " ").Replace("\n", " ");
        return text.Length <= count ? text : text.Substring(0, count) + "...";
    }
}
