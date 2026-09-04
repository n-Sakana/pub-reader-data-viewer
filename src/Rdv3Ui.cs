// ============================================================================
// Rdv3Ui.cs -- settings-driven WebView2 UI bridge.
//
// The browser owns pixels and focus.  Rdv3App still owns every operation and
// state transition; this class only serializes screen state and turns browser
// messages back into the same actions used by the former desktop surface.
// C# 5, ASCII source.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Win32;

public sealed class Rdv3CandRow
{
    public string Line = "";
    public string Stored = "";
}

public static class Rdv3WebJson
{
    public static string Q(string value)
    {
        StringBuilder sb = new StringBuilder();
        sb.Append('"');
        if (value != null)
        {
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == '"' || c == '\\') { sb.Append('\\').Append(c); }
                else if (c == '\b') { sb.Append("\\b"); }
                else if (c == '\f') { sb.Append("\\f"); }
                else if (c == '\n') { sb.Append("\\n"); }
                else if (c == '\r') { sb.Append("\\r"); }
                else if (c == '\t') { sb.Append("\\t"); }
                else if (c < ' ') { sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture)); }
                else { sb.Append(c); }
            }
        }
        sb.Append('"');
        return sb.ToString();
    }

    public static string B(bool value) { return value ? "true" : "false"; }

    public static string N(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    public static string A(double[] values)
    {
        StringBuilder sb = new StringBuilder("[");
        if (values != null)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (i > 0) { sb.Append(','); }
                sb.Append(N(values[i]));
            }
        }
        return sb.Append(']').ToString();
    }

    public static string S(string[] values)
    {
        StringBuilder sb = new StringBuilder("[");
        if (values != null)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (i > 0) { sb.Append(','); }
                sb.Append(Q(values[i]));
            }
        }
        return sb.Append(']').ToString();
    }
}

public sealed class Rdv3Form
{
    private readonly ReaderDataViewer.MainWindow host;
    private Rdv3Fields fields = Rdv3Fields.Empty;
    private List<Rdv3CandRow> candidates = new List<Rdv3CandRow>();
    private int candidateTotal;
    private bool pageReady;
    private bool shown;
    private bool opsEnabled;
    private bool workEnabled;
    private string keyText = "";
    private string notice = "";
    private bool noticeError;
    private readonly DispatcherTimer noticeTimer = new DispatcherTimer();
    private Rdv3Data exportFilterData;
    private int modalToken;
    private int waitingToken;
    private Rdv3Json modalResult;
    private DispatcherFrame modalFrame;
    private Rdv3Target pickedTarget;
    private int pickerPollMs = 40;

    public readonly Rdv3Screen Screen;
    public readonly Rdv3View View = new Rdv3View();

    public Action<string> OnSearch;
    public Action OnClear;
    public Action OnWorkState;
    public Action OnRefreshLedger;
    public Action OnTableExport;
    public Action<string> OnUpdateRecords;
    public Action<string> OnDeleteRecords;
    public Action OnSendChanges;
    public Action OnSettings;

    public event EventHandler Shown;
    public event EventHandler<ReaderDataViewer.Rdv3FormClosingEventArgs> FormClosing;

    public Rdv3Form(ReaderDataViewer.MainWindow window, Rdv3Screen screen)
    {
        host = window;
        Screen = screen;
        View.UserName = Environment.UserName;
        View.HostName = Environment.MachineName;
        host.PageLoaded += OnPageLoaded;
        host.WebMessage += OnWebMessage;
        host.ClosingRequested += OnClosingRequested;
        noticeTimer.Interval = TimeSpan.FromMilliseconds(3600);
        noticeTimer.Tick += delegate
        {
            noticeTimer.Stop();
            notice = "";
            noticeError = false;
            RefreshValues();
        };
    }

    public ReaderDataViewer.MainWindow Host { get { return host; } }
    public Rdv3Fields Fields { get { return fields; } }
    public List<Rdv3CandRow> Candidates { get { return candidates; } }
    public int CandidateTotal { get { return candidateTotal; } }
    public Rect CardBounds { get { return host.ScreenBounds; } }

    public void RunOnUi(Action action) { Ui(action); }

    public void PostOnUi(Action action)
    {
        if (host.Dispatcher.HasShutdownStarted) { return; }
        host.Dispatcher.BeginInvoke(action, DispatcherPriority.Background);
    }

    public void SetFields(Rdv3Fields value)
    {
        Ui(delegate
        {
            fields = value == null ? Rdv3Fields.Empty : value;
            RefreshValues();
        });
    }

    public void SetState(string text)
    {
        Ui(delegate { View.AppState = text ?? ""; RefreshValues(); });
    }

    public void SetWatch(string label, string detail)
    {
        Ui(delegate
        {
            View.WatchLabel = ((label == null || label.Length == 0)
                ? Rdv3Text.LabelWatch : label) +
                ((detail == null || detail.Length == 0) ? "" : " " + detail);
            View.WatchDetail = detail ?? "";
            RefreshValues();
        });
    }

    public void SetLedger(string file, string rows, string saved)
    {
        Ui(delegate
        {
            View.LedgerFile = file ?? "";
            View.LedgerRows = rows ?? "";
            View.LedgerSaved = saved ?? "";
            RefreshValues();
        });
    }

    public void SetPendingCount(int count)
    {
        Ui(delegate { View.PendingCount = Math.Max(0, count); RefreshValues(); });
    }

    public void SetTimes(double mergeMs, double searchMs)
    {
        Ui(delegate
        {
            if (mergeMs >= 0) { View.MergeMs = Rdv3Clock.Fmt(mergeMs) + Rdv3Text.MsUnit; }
            if (searchMs >= 0) { View.SearchMs = Rdv3Clock.Fmt(searchMs) + Rdv3Text.MsUnit; }
            RefreshValues();
        });
    }

    public void SetIdentity(string pid, string logName)
    {
        Ui(delegate
        {
            View.Pid = pid ?? "";
            View.LogName = logName ?? "";
            RefreshValues();
        });
    }

    public void EnableOps(bool on)
    {
        Ui(delegate { opsEnabled = on; RefreshValues(); });
    }

    public void EnableWorkState(bool on)
    {
        Ui(delegate { workEnabled = on; RefreshValues(); });
    }

    public void ShowCandidates(string key, List<Rdv3CandRow> rows, int totalHits)
    {
        Ui(delegate
        {
            View.SearchKey = key ?? "";
            keyText = View.SearchKey;
            candidates = rows ?? new List<Rdv3CandRow>();
            candidateTotal = totalHits;
            View.CandidateCount = totalHits;
            View.SelectedIndex = -1;
            View.RowNumber = 0;
            View.Record = null;
            View.StoredState = "";
            View.Saving = false;
            RefreshValues();
        });
    }

    public void SelectCandidate(int index)
    {
        Ui(delegate
        {
            if (index < 0 || index >= candidates.Count) { return; }
            View.SelectedIndex = index;
            View.RowNumber = index + 1;
            View.Record = Rdv3Ledger.SplitLine(candidates[index].Line);
            View.StoredState = candidates[index].Stored ?? "";
            View.Saving = false;
            RefreshValues();
        });
    }

    public void SetStoredState(int index, string stored, bool saving)
    {
        Ui(delegate
        {
            if (index >= 0 && index < candidates.Count)
            {
                candidates[index].Stored = stored ?? "";
            }
            if (index == View.SelectedIndex && View.Record != null)
            {
                View.StoredState = stored ?? "";
            }
            View.Saving = saving && index == View.SelectedIndex;
            RefreshValues();
        });
    }

    public void ClearResult()
    {
        Ui(delegate
        {
            keyText = "";
            View.SearchKey = "";
            candidates = new List<Rdv3CandRow>();
            candidateTotal = 0;
            View.CandidateCount = 0;
            View.SelectedIndex = -1;
            View.RowNumber = 0;
            View.Record = null;
            View.StoredState = "";
            View.Saving = false;
            RefreshValues();
        });
    }

    public string KeyText { get { return (keyText ?? "").Trim(); } }

    public void SetKeyText(string value)
    {
        Ui(delegate
        {
            keyText = value ?? "";
            if (pageReady)
            {
                host.PostJson("{\"type\":\"keySet\",\"value\":" +
                    Rdv3WebJson.Q(keyText) + "}");
            }
        });
    }

    public void CaptureToFile(string path) { Ui(delegate { host.CaptureToFile(path); }); }

    public void Notice(string text)
    {
        Ui(delegate
        {
            noticeTimer.Stop();
            notice = text ?? "";
            noticeError = false;
            if (notice.Length > 0) { noticeTimer.Start(); }
            RefreshValues();
        });
    }

    public void Error(string text)
    {
        if (string.IsNullOrEmpty(text)) { return; }
        Ui(delegate
        {
            Rdv3ConfirmForm.Tell(this, Rdv3Text.AppTitle, text);
        });
    }

    public void SharedNotice(string text) { Notice(text); }
    public bool Ask(string title, string body) { return Rdv3ConfirmForm.Ask(this, title, body); }
    public void Tell(string title, string body) { Rdv3ConfirmForm.Tell(this, title, body); }
    public bool AskLedgerSwitch(List<Rdv3CandRow> rows) { return Rdv3LedgerUpdateForm.Ask(this, rows); }
    public void TellResetRows(List<Rdv3CandRow> rows) { Rdv3LedgerUpdateForm.TellReset(this, rows); }
    public void TellUnmatched(List<Rdv3UnmatchedChange> rows) { Rdv3UnmatchedForm.Tell(this, rows); }

    public void Fatal(string title, string body)
    {
        Rdv3ConfirmForm.Tell(this, title, body);
        host.WindowCommand("close");
    }

    public int PickFromList()
    {
        if (candidates.Count == 0) { return -1; }
        return Rdv3CandidatesForm.Pick(
            this,
            Screen.Candidates,
            candidates,
            candidateTotal,
            View.SelectedIndex);
    }

    public string Diag
    {
        get
        {
            return "client=" + ((int)Math.Round(host.ActualWidth)).ToString(CultureInfo.InvariantCulture)
                + "x" + ((int)Math.Round(host.ActualHeight)).ToString(CultureInfo.InvariantCulture)
                + " card=" + Screen.CardWidth.ToString(CultureInfo.InvariantCulture)
                + " font=" + Screen.FontFamily
                + " sections=" + Screen.Sections.Count.ToString(CultureInfo.InvariantCulture)
                + " bridge=webview2";
        }
    }

    internal Rdv3Json ShowModal(string modal, string content)
    {
        string modalCapture = Environment.GetEnvironmentVariable(
            "RDV_WEBVIEW2_PROBE_MODAL_CAPTURE");
        if (ReaderDataViewer.App.IsProbe && string.IsNullOrWhiteSpace(modalCapture))
        {
            if (modal == "candidates") { return Rdv3Json.Parse("{\"ok\":true,\"index\":0}"); }
            return Rdv3Json.Parse("{\"ok\":true}");
        }
        if (!pageReady) { return Rdv3Json.Parse("{\"ok\":false}"); }
        modalToken++;
        waitingToken = modalToken;
        modalResult = null;
        modalFrame = new DispatcherFrame();
        host.PostJson("{\"type\":\"modalOpen\",\"token\":" +
            waitingToken.ToString(CultureInfo.InvariantCulture) +
            ",\"modal\":" + Rdv3WebJson.Q(modal) +
            ",\"content\":" + content + "}");
        Dispatcher.PushFrame(modalFrame);
        Rdv3Json result = modalResult;
        modalFrame = null;
        modalResult = null;
        waitingToken = 0;
        return result ?? Rdv3Json.Parse("{\"ok\":false}");
    }

    public void TriggerProbeAction(string action)
    {
        if (action == "tableExport" && OnTableExport != null) { OnTableExport(); }
        else if (action == "settings" && OnSettings != null) { OnSettings(); }
        else if (action == "sendChanges" && OnSendChanges != null) { OnSendChanges(); }
        else if (action == "updateRecords" && OnUpdateRecords != null) { OnUpdateRecords(JobOf(action)); }
        else if (action == "deleteRecords" && OnDeleteRecords != null) { OnDeleteRecords(JobOf(action)); }
    }

    private string JobOf(string action)
    {
        for (int i = 0; i < Screen.Sections.Count; i++)
        {
            for (int k = 0; k < Screen.Sections[i].Buttons.Count; k++)
            {
                Rdv3ButtonDef button = Screen.Sections[i].Buttons[k];
                if (button.Action == action) { return button.Job; }
            }
        }
        return "";
    }

    internal void PatchModal(string field, string value)
    {
        host.PostJson("{\"type\":\"modalPatch\",\"token\":" +
            waitingToken.ToString(CultureInfo.InvariantCulture) +
            ",\"field\":" + Rdv3WebJson.Q(field) +
            ",\"value\":" + Rdv3WebJson.Q(value) + "}");
    }

    internal void PostPickerPreview(string json)
    {
        host.PostJson("{\"type\":\"pickerPreview\",\"preview\":" + json + "}");
    }

    internal void SetPickedTarget(Rdv3Target target)
    {
        pickedTarget = target;
    }

    internal Rdv3Target TakePickedTarget()
    {
        Rdv3Target value = pickedTarget;
        pickedTarget = null;
        return value;
    }

    internal void SetPickerPollMs(int value)
    {
        pickerPollMs = value;
    }

    internal void SetExportFilterData(Rdv3Data value)
    {
        exportFilterData = value;
    }

    private void Ui(Action action)
    {
        if (host.Dispatcher.HasShutdownStarted) { return; }
        if (host.Dispatcher.CheckAccess()) { action(); }
        else { host.Dispatcher.Invoke(action); }
    }

    private void OnPageLoaded(object sender, EventArgs eventArgs)
    {
        host.PostJson(BuildInitJson());
    }

    private void OnClosingRequested(
        object sender,
        ReaderDataViewer.Rdv3FormClosingEventArgs eventArgs)
    {
        EventHandler<ReaderDataViewer.Rdv3FormClosingEventArgs> handler = FormClosing;
        if (handler != null) { handler(this, eventArgs); }
    }

    private void OnWebMessage(string json)
    {
        try
        {
            Rdv3Json root = Rdv3Json.Parse(json);
            string type = Text(root, "type");
            if (type == "ready")
            {
                pageReady = true;
                host.ShowSurface();
                RefreshValues();
                if (!shown)
                {
                    shown = true;
                    EventHandler handler = Shown;
                    if (handler != null) { handler(this, EventArgs.Empty); }
                }
            }
            else if (type == "key") { keyText = Text(root, "value"); }
            else if (type == "action")
            {
                Rdv3Json action = root;
                PostOnUi(delegate { DispatchAction(action); });
            }
            else if (type == "window") { host.WindowCommand(Text(root, "command")); }
            else if (type == "modalResult") { CompleteModal(root); }
            else if (type == "modalShown")
            {
                PostOnUi(delegate { CaptureModal(root); });
            }
            else if (type == "settingsSubmit") { ValidateSettings(root); }
            else if (type == "validateExportFilter") { ValidateExportFilter(root); }
            else if (type == "browse") { Browse(root); }
            else if (type == "picker") { PickTarget(); }
            else if (type == "pickerCancel") { Rdv3PickerForm.CancelCurrent(); }
        }
        catch (Exception exception)
        {
            Error(exception.Message);
        }
    }

    private void ValidateSettings(Rdv3Json root)
    {
        int token = Number(root, "token", 0);
        if (token != waitingToken || modalFrame == null) { return; }
        string dataDir = Text(root, "dataDir").Trim();
        string ledger = Text(root, "ledger").Trim();
        string log = Text(root, "log").Trim();
        string pattern = Text(root, "pattern").Trim();
        int candidateRows = Number(root, "candidateRows", -1);
        string error = "";
        string field = "";
        string patternError = Rdv3Config.PatternError(pattern);
        if (patternError != null)
        {
            error = Rdv3Text.ErrPatternTyped + patternError;
            field = "pattern";
        }
        else if (dataDir.Length == 0 || ledger.Length == 0 || log.Length == 0)
        {
            error = Rdv3Text.ErrPathBlank;
            field = dataDir.Length == 0 ? "dataDir" : ledger.Length == 0 ? "ledger" : "log";
        }
        else if (candidateRows < 1 || candidateRows > 1000)
        {
            error = Rdv3Text.LblCandidateRows;
            field = "candidateRows";
        }
        host.PostJson("{\"type\":\"settingsValidation\",\"token\":" +
            token.ToString(CultureInfo.InvariantCulture) +
            ",\"ok\":" + Rdv3WebJson.B(error.Length == 0) +
            ",\"error\":" + Rdv3WebJson.Q(error) +
            ",\"field\":" + Rdv3WebJson.Q(field) + "}");
    }

    private void ValidateExportFilter(Rdv3Json root)
    {
        int token = Number(root, "token", 0);
        if (token != waitingToken || modalFrame == null || exportFilterData == null) { return; }
        string reference = Text(root, "field");
        string firstText = Text(root, "first");
        string lastText = Text(root, "last");
        string error = "";
        Rdv3ColumnTypeDef type = exportFilterData.TypeOf(reference);
        if (type == null)
        {
            if (firstText.Trim().Length == 0) { error = Rdv3Text.ExportFilterNeedValue; }
            lastText = "";
        }
        else if (type.Type == "date")
        {
            DateTime first;
            DateTime last;
            if (!type.TryDate(firstText, out first) || !type.TryDate(lastText, out last))
            {
                error = Rdv3Text.ExportFilterNeedValue;
            }
            else if (first.Date > last.Date) { error = Rdv3Text.ExportFilterOrder; }
            else
            {
                firstText = first.ToString(type.Format, CultureInfo.InvariantCulture);
                lastText = last.ToString(type.Format, CultureInfo.InvariantCulture);
            }
        }
        else
        {
            decimal first;
            decimal last;
            if (!type.TryNumber(firstText, out first) || !type.TryNumber(lastText, out last))
            {
                error = Rdv3Text.ExportFilterNeedNumber;
            }
            else if (first > last) { error = Rdv3Text.ExportFilterOrder; }
            else
            {
                firstText = first.ToString(CultureInfo.InvariantCulture);
                lastText = last.ToString(CultureInfo.InvariantCulture);
            }
        }
        host.PostJson("{\"type\":\"exportFilterValidation\",\"token\":" +
            token.ToString(CultureInfo.InvariantCulture) +
            ",\"ok\":" + Rdv3WebJson.B(error.Length == 0) +
            ",\"error\":" + Rdv3WebJson.Q(error) +
            ",\"first\":" + Rdv3WebJson.Q(firstText) +
            ",\"last\":" + Rdv3WebJson.Q(lastText) + "}");
    }

    private void DispatchAction(Rdv3Json root)
    {
        string action = Text(root, "name");
        string job = Text(root, "job");
        string key = Text(root, "key").Trim();
        if (action == "search" && OnSearch != null) { keyText = key; OnSearch(key); }
        else if (action == "clear" && OnClear != null) { OnClear(); }
        else if (action == "workState" && OnWorkState != null) { OnWorkState(); }
        else if (action == "refreshLedger" && OnRefreshLedger != null) { OnRefreshLedger(); }
        else if (action == "tableExport" && OnTableExport != null) { OnTableExport(); }
        else if (action == "updateRecords" && OnUpdateRecords != null) { OnUpdateRecords(job); }
        else if (action == "deleteRecords" && OnDeleteRecords != null) { OnDeleteRecords(job); }
        else if (action == "sendChanges" && OnSendChanges != null) { OnSendChanges(); }
        else if (action == "settings" && OnSettings != null) { OnSettings(); }
    }

    private void CompleteModal(Rdv3Json root)
    {
        int token = Number(root, "token", 0);
        if (modalFrame == null || token != waitingToken) { return; }
        Rdv3Json result = root.Member("result");
        modalResult = result != null && result.Kind == Rdv3Json.TObject
            ? result : Rdv3Json.Parse("{\"ok\":false}");
        modalFrame.Continue = false;
    }

    private void CaptureModal(Rdv3Json root)
    {
        if (!ReaderDataViewer.App.IsProbe || modalFrame == null ||
            Number(root, "token", 0) != waitingToken) { return; }
        string path = Environment.GetEnvironmentVariable(
            "RDV_WEBVIEW2_PROBE_MODAL_CAPTURE");
        if (string.IsNullOrWhiteSpace(path)) { return; }
        host.CaptureToFile(path);
        modalResult = Rdv3Json.Parse("{\"ok\":false}");
        modalFrame.Continue = false;
    }

    private void Browse(Rdv3Json root)
    {
        if (Number(root, "token", 0) != waitingToken) { return; }
        string field = Text(root, "field");
        string kind = Text(root, "kind");
        string initial = Text(root, "value");
        string value = kind == "folder" ? BrowseFolder(initial)
            : BrowseFile(initial, kind == "export" ? "csv"
                : kind == "log" ? "log" : "xlsx");
        if (value.Length > 0) { PatchModal(field, value); }
    }

    private string BrowseFile(string initial, string extension)
    {
        SaveFileDialog dialog = new SaveFileDialog();
        dialog.AddExtension = true;
        dialog.DefaultExt = extension;
        dialog.Filter = extension == "csv" ? "CSV (*.csv)|*.csv|All files (*.*)|*.*"
            : extension == "log" ? "Log (*.log)|*.log|All files (*.*)|*.*"
            : "Excel (*.xlsx)|*.xlsx|All files (*.*)|*.*";
        try
        {
            dialog.InitialDirectory = Path.GetDirectoryName(initial);
            dialog.FileName = Path.GetFileName(initial);
        }
        catch (Exception) { }
        return dialog.ShowDialog(host) == true ? dialog.FileName : "";
    }

    private string BrowseFolder(string initial)
    {
        object shell = null;
        object folder = null;
        object self = null;
        try
        {
            Type type = Type.GetTypeFromProgID("Shell.Application");
            if (type == null) { return ""; }
            shell = Activator.CreateInstance(type);
            folder = type.InvokeMember(
                "BrowseForFolder",
                BindingFlags.InvokeMethod,
                null,
                shell,
                new object[] { 0, Rdv3Text.LblDataShort, 0, initial });
            if (folder == null) { return ""; }
            self = folder.GetType().InvokeMember(
                "Self",
                BindingFlags.GetProperty,
                null,
                folder,
                null);
            object path = self.GetType().InvokeMember(
                "Path",
                BindingFlags.GetProperty,
                null,
                self,
                null);
            return path == null ? "" : path.ToString();
        }
        catch (Exception) { return ""; }
        finally
        {
            ReleaseCom(self);
            ReleaseCom(folder);
            ReleaseCom(shell);
        }
    }

    private static void ReleaseCom(object value)
    {
        try { if (value != null && Marshal.IsComObject(value)) { Marshal.FinalReleaseComObject(value); } }
        catch (Exception) { }
    }

    private void PickTarget()
    {
        if (waitingToken == 0) { return; }
        Rdv3Target picked = Rdv3PickerForm.Pick(this);
        if (picked == null)
        {
            host.PostJson("{\"type\":\"pickerResult\",\"token\":" +
                waitingToken.ToString(CultureInfo.InvariantCulture) +
                ",\"target\":null}");
            return;
        }
        pickedTarget = picked;
        host.PostJson("{\"type\":\"pickerResult\",\"token\":" +
            waitingToken.ToString(CultureInfo.InvariantCulture) +
            ",\"target\":" + Rdv3SettingsForm.TargetJson(picked, pickerPollMs) + "}");
    }

    private void RefreshValues()
    {
        if (!pageReady) { return; }
        host.PostJson(BuildStateJson());
    }

    private string BuildInitJson()
    {
        StringBuilder sb = new StringBuilder(32768);
        sb.Append("{\"type\":\"init\",\"screen\":{");
        sb.Append("\"card\":{");
        sb.Append("\"width\":").Append(Rdv3WebJson.N(Screen.CardWidth));
        sb.Append(",\"startWidth\":").Append(Rdv3WebJson.N(Screen.StartWidth));
        sb.Append(",\"startHeight\":").Append(Rdv3WebJson.N(Screen.StartHeight));
        sb.Append(",\"gap\":").Append(Rdv3WebJson.N(Screen.Gap));
        sb.Append(",\"padding\":").Append(Rdv3WebJson.A(Screen.Padding));
        sb.Append(",\"font\":").Append(Rdv3WebJson.Q(Screen.FontFamily));
        sb.Append(",\"fontSize\":").Append(Rdv3WebJson.N(Screen.FontSize));
        sb.Append(",\"keyValueFontSize\":").Append(Rdv3WebJson.N(Screen.KeyValueFontSize));
        sb.Append(",\"judgmentFontSize\":").Append(Rdv3WebJson.N(Screen.JudgmentFontSize));
        sb.Append(",\"unsearchedFontSize\":").Append(Rdv3WebJson.N(Screen.UnsearchedFontSize));
        sb.Append("},\"sections\":[");
        for (int i = 0; i < Screen.Sections.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            AppendSection(sb, Screen.Sections[i], "s" + i.ToString(CultureInfo.InvariantCulture));
        }
        sb.Append("],\"candidates\":");
        AppendCandidatesDef(sb, Screen.Candidates);
        sb.Append("},\"state\":").Append(BuildStateBody()).Append('}');
        return sb.ToString();
    }

    private string BuildStateJson()
    {
        return "{\"type\":\"state\",\"state\":" + BuildStateBody() + "}";
    }

    private string BuildStateBody()
    {
        StringBuilder sb = new StringBuilder(16384);
        sb.Append("{\"values\":{");
        bool comma = false;
        for (int i = 0; i < Screen.Sections.Count; i++)
        {
            AppendSectionValues(
                sb,
                Screen.Sections[i],
                "s" + i.ToString(CultureInfo.InvariantCulture),
                ref comma);
        }
        sb.Append("},\"judgments\":{");
        comma = false;
        for (int i = 0; i < Screen.Sections.Count; i++)
        {
            Rdv3Section section = Screen.Sections[i];
            if (section.Type != "statusBand") { continue; }
            if (comma) { sb.Append(','); }
            comma = true;
            string id = "s" + i.ToString(CultureInfo.InvariantCulture);
            Rdv3Verdict verdict = Rdv3Eval.Judge(
                Screen.JudgmentOf(section.Judgment),
                View,
                fields);
            sb.Append(Rdv3WebJson.Q(id)).Append(":{");
            if (!View.HasRecord || verdict.Result == null)
            {
                sb.Append("\"text\":").Append(Rdv3WebJson.Q(Rdv3Text.Unsearched));
                sb.Append(",\"look\":\"unsearched\",\"icon\":\"\"");
            }
            else
            {
                sb.Append("\"text\":").Append(Rdv3WebJson.Q(verdict.Result.Text));
                sb.Append(",\"look\":").Append(Rdv3WebJson.Q(verdict.Result.Look));
                sb.Append(",\"icon\":").Append(Rdv3WebJson.Q(verdict.Result.Icon));
            }
            List<string> subs = new List<string>();
            for (int k = 0; k < section.Sub.Count; k++)
            {
                string value = Rdv3Eval.Evaluate(section.Sub[k], View, fields, Screen.Work).Text;
                if (!string.IsNullOrEmpty(value)) { subs.Add(value); }
            }
            sb.Append(",\"sub\":").Append(Rdv3WebJson.Q(string.Join(section.Joiner, subs.ToArray()))).Append('}');
        }
        Rdv3StateDef workState = View.HasRecord
            ? Screen.Work.ByStored(View.StoredState) : Screen.Work.InitialState;
        string stateText = workState == null ? "" : workState.Text;
        string buttonText = Screen.Work.ButtonText.Replace("{state}", stateText);
        bool down = View.HasRecord && workState != null && workState.Id != Screen.Work.Initial;
        sb.Append("},\"key\":").Append(Rdv3WebJson.Q(keyText));
        sb.Append(",\"opsEnabled\":").Append(Rdv3WebJson.B(opsEnabled));
        sb.Append(",\"workEnabled\":").Append(Rdv3WebJson.B(workEnabled));
        sb.Append(",\"workText\":").Append(Rdv3WebJson.Q(buttonText));
        sb.Append(",\"workDown\":").Append(Rdv3WebJson.B(down));
        sb.Append(",\"pending\":").Append(View.PendingCount.ToString(CultureInfo.InvariantCulture));
        sb.Append(",\"notice\":").Append(Rdv3WebJson.Q(notice));
        sb.Append(",\"noticeError\":").Append(Rdv3WebJson.B(noticeError));
        sb.Append('}');
        return sb.ToString();
    }

    private void AppendSectionValues(
        StringBuilder sb,
        Rdv3Section section,
        string path,
        ref bool comma)
    {
        if (section.Value != null) { AppendValue(sb, path + ".value", section.Value, ref comma); }
        for (int i = 0; i < section.Rows.Count; i++)
        {
            AppendValue(sb, path + ".row" + i.ToString(CultureInfo.InvariantCulture), section.Rows[i].Value, ref comma);
        }
        for (int i = 0; i < section.Sub.Count; i++)
        {
            AppendValue(sb, path + ".sub" + i.ToString(CultureInfo.InvariantCulture), section.Sub[i], ref comma);
        }
        for (int i = 0; i < section.Segments.Count; i++)
        {
            AppendValue(sb, path + ".segment" + i.ToString(CultureInfo.InvariantCulture), section.Segments[i].Value, ref comma);
        }
        for (int i = 0; i < section.Items.Count; i++)
        {
            AppendSectionValues(sb, section.Items[i], path + ".item" + i.ToString(CultureInfo.InvariantCulture), ref comma);
        }
    }

    private void AppendValue(StringBuilder sb, string id, Rdv3Bind bind, ref bool comma)
    {
        if (comma) { sb.Append(','); }
        comma = true;
        Rdv3Value value = Rdv3Eval.Evaluate(bind, View, fields, Screen.Work);
        sb.Append(Rdv3WebJson.Q(id)).Append(":{");
        sb.Append("\"text\":").Append(Rdv3WebJson.Q(value.Text));
        sb.Append(",\"tone\":").Append(value.Tone.ToString(CultureInfo.InvariantCulture));
        sb.Append('}');
    }

    private static void AppendSection(StringBuilder sb, Rdv3Section section, string path)
    {
        sb.Append('{');
        sb.Append("\"type\":").Append(Rdv3WebJson.Q(section.Type));
        sb.Append(",\"id\":").Append(Rdv3WebJson.Q(path));
        if (section.Margin != null) { sb.Append(",\"margin\":").Append(Rdv3WebJson.A(section.Margin)); }
        if (section.Type == "titleBar")
        {
            sb.Append(",\"brand\":").Append(Rdv3WebJson.Q(section.Brand));
            sb.Append(",\"tags\":[");
            for (int i = 0; i < section.Tags.Count; i++)
            {
                if (i > 0) { sb.Append(','); }
                sb.Append("{\"text\":").Append(Rdv3WebJson.Q(section.Tags[i].Text));
                sb.Append(",\"look\":").Append(Rdv3WebJson.Q(section.Tags[i].Look)).Append('}');
            }
            sb.Append(']');
        }
        else if (section.Type == "keyPanel")
        {
            sb.Append(",\"title\":").Append(Rdv3WebJson.Q(section.Title));
            sb.Append(",\"label\":").Append(Rdv3WebJson.Q(section.Label));
            sb.Append(",\"value\":").Append(Rdv3WebJson.Q(path + ".value"));
            sb.Append(",\"inputLabel\":").Append(Rdv3WebJson.Q(section.InputLabel));
            sb.Append(",\"placeholder\":").Append(Rdv3WebJson.Q(section.Placeholder));
            sb.Append(",\"inputWidth\":").Append(Rdv3WebJson.N(section.InputWidth));
            sb.Append(",\"maxLength\":").Append(section.MaxLength.ToString(CultureInfo.InvariantCulture));
        }
        else if (section.Type == "columns")
        {
            sb.Append(",\"gap\":").Append(Rdv3WebJson.N(section.Gap));
            sb.Append(",\"stackBelow\":").Append(Rdv3WebJson.N(section.StackBelow));
            sb.Append(",\"weights\":").Append(Rdv3WebJson.A(section.Weights));
            sb.Append(",\"items\":[");
            for (int i = 0; i < section.Items.Count; i++)
            {
                if (i > 0) { sb.Append(','); }
                AppendSection(sb, section.Items[i], path + ".item" + i.ToString(CultureInfo.InvariantCulture));
            }
            sb.Append(']');
        }
        else if (section.Type == "fieldList")
        {
            sb.Append(",\"title\":").Append(Rdv3WebJson.Q(section.Title));
            sb.Append(",\"labelWidth\":").Append(Rdv3WebJson.N(section.LabelWidth));
            sb.Append(",\"rowHeight\":").Append(Rdv3WebJson.N(section.RowHeight));
            sb.Append(",\"rows\":[");
            for (int i = 0; i < section.Rows.Count; i++)
            {
                if (i > 0) { sb.Append(','); }
                sb.Append("{\"label\":").Append(Rdv3WebJson.Q(section.Rows[i].Label));
                sb.Append(",\"value\":").Append(Rdv3WebJson.Q(path + ".row" + i.ToString(CultureInfo.InvariantCulture))).Append('}');
            }
            sb.Append(']');
        }
        else if (section.Type == "textBox")
        {
            sb.Append(",\"title\":").Append(Rdv3WebJson.Q(section.Title));
            sb.Append(",\"lines\":").Append(section.Lines.ToString(CultureInfo.InvariantCulture));
            sb.Append(",\"value\":").Append(Rdv3WebJson.Q(path + ".value"));
        }
        else if (section.Type == "statusBand")
        {
            sb.Append(",\"label\":").Append(Rdv3WebJson.Q(section.Label));
            sb.Append(",\"height\":").Append(Rdv3WebJson.N(section.Height));
            sb.Append(",\"judgment\":").Append(Rdv3WebJson.Q(path));
        }
        else if (section.Type == "sendBar")
        {
            sb.Append(",\"height\":").Append(Rdv3WebJson.N(section.Height));
            sb.Append(",\"value\":").Append(Rdv3WebJson.Q(path + ".value"));
        }
        else if (section.Type == "statusBar")
        {
            sb.Append(",\"height\":").Append(Rdv3WebJson.N(section.Height));
            sb.Append(",\"segments\":[");
            for (int i = 0; i < section.Segments.Count; i++)
            {
                if (i > 0) { sb.Append(','); }
                Rdv3SegmentDef segment = section.Segments[i];
                sb.Append("{\"prefix\":").Append(Rdv3WebJson.Q(segment.Prefix));
                sb.Append(",\"value\":").Append(Rdv3WebJson.Q(path + ".segment" + i.ToString(CultureInfo.InvariantCulture)));
                sb.Append(",\"bold\":").Append(Rdv3WebJson.B(segment.Bold));
                sb.Append(",\"dot\":").Append(Rdv3WebJson.B(segment.Dot));
                sb.Append(",\"clock\":").Append(Rdv3WebJson.B(
                    segment.Value != null && segment.Value.IsState && segment.Value.State == "clock")).Append('}');
            }
            sb.Append(']');
        }
        if (section.Buttons.Count > 0)
        {
            sb.Append(",\"buttons\":[");
            for (int i = 0; i < section.Buttons.Count; i++)
            {
                if (i > 0) { sb.Append(','); }
                AppendButton(sb, section.Buttons[i]);
            }
            sb.Append(']');
        }
        sb.Append('}');
    }

    private static void AppendButton(StringBuilder sb, Rdv3ButtonDef button)
    {
        sb.Append("{\"action\":").Append(Rdv3WebJson.Q(button.Action));
        sb.Append(",\"text\":").Append(Rdv3WebJson.Q(button.Text));
        sb.Append(",\"icon\":").Append(Rdv3WebJson.Q(button.Icon));
        sb.Append(",\"tip\":").Append(Rdv3WebJson.Q(button.Tip));
        sb.Append(",\"job\":").Append(Rdv3WebJson.Q(button.Job));
        sb.Append(",\"primary\":").Append(Rdv3WebJson.B(button.Primary)).Append('}');
    }

    private static void AppendCandidatesDef(StringBuilder sb, Rdv3CandidatesDef def)
    {
        if (def == null) { sb.Append("null"); return; }
        sb.Append('{');
        sb.Append("\"title\":").Append(Rdv3WebJson.Q(def.Title));
        sb.Append(",\"hint\":").Append(Rdv3WebJson.Q(def.Hint));
        sb.Append(",\"width\":").Append(Rdv3WebJson.N(def.Width));
        sb.Append(",\"maxHeight\":").Append(Rdv3WebJson.N(def.MaxHeight));
        sb.Append(",\"rowHeight\":").Append(Rdv3WebJson.N(def.RowHeight));
        sb.Append(",\"headerHeight\":").Append(Rdv3WebJson.N(def.HeaderHeight));
        sb.Append(",\"columns\":[");
        for (int i = 0; i < def.Columns.Count; i++)
        {
            if (i > 0) { sb.Append(','); }
            Rdv3ColumnDef column = def.Columns[i];
            sb.Append("{\"header\":").Append(Rdv3WebJson.Q(column.Header));
            sb.Append(",\"width\":").Append(Rdv3WebJson.N(column.Width));
            sb.Append(",\"align\":").Append(Rdv3WebJson.Q(column.Align));
            sb.Append(",\"bold\":").Append(Rdv3WebJson.B(column.Bold));
            sb.Append(",\"muted\":").Append(Rdv3WebJson.B(column.Muted));
            sb.Append(",\"render\":").Append(Rdv3WebJson.Q(column.Render)).Append('}');
        }
        sb.Append("]}");
    }

    internal static string Text(Rdv3Json root, string name)
    {
        Rdv3Json value = root == null ? null : root.Member(name);
        return value != null && value.Kind == Rdv3Json.TString ? value.Str : "";
    }

    internal static int Number(Rdv3Json root, string name, int fallback)
    {
        Rdv3Json value = root == null ? null : root.Member(name);
        return value != null && value.Kind == Rdv3Json.TNumber ? (int)value.Num : fallback;
    }

    internal static bool Flag(Rdv3Json root, string name, bool fallback)
    {
        Rdv3Json value = root == null ? null : root.Member(name);
        return value != null && value.Kind == Rdv3Json.TBool ? value.Flag : fallback;
    }
}
