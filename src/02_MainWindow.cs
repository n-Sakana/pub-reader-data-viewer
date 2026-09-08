using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;

namespace ReaderDataViewer
{
    public sealed class Rdv3FormClosingEventArgs : EventArgs
    {
        public bool Cancel;
    }


    // Windows 11 lets an app colour its own caption. Without it the light
    // system caption sits on top of the 98 grey face and the seam shows.
    // There is no managed wrapper for it, so this is the documented route.
    internal static class WindowCaption
    {
        private const int BorderColour = 34;
        private const int CaptionColour = 35;
        private const int TextColour = 36;

        // On older Windows, unsupported colour attributes are ignored.

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(
            IntPtr window, int attribute, ref int value, int size);

        public static void Apply(Window window)
        {
            IntPtr handle = new WindowInteropHelper(window).Handle;
            if (handle == IntPtr.Zero) { return; }
            int face = Win98.Caption;
            int ink = Win98.CaptionText;
            int edge = Win98.Border;
            try
            {
                // Older Windows simply reports the attribute as unsupported.
                DwmSetWindowAttribute(handle, CaptionColour, ref face, sizeof(int));
                DwmSetWindowAttribute(handle, TextColour, ref ink, sizeof(int));
                DwmSetWindowAttribute(handle, BorderColour, ref edge, sizeof(int));
            }
            catch (EntryPointNotFoundException) { }
            catch (DllNotFoundException) { }
        }
    }

    public sealed class MainWindow : Window
    {
        private const string TrustedHost = "reader-data-viewer.local";
        private const string StartPage =
            "https://reader-data-viewer.local/index.html";

        private readonly WebView2 webView;
        private readonly double targetClientWidth;
        private readonly double targetClientHeight;
        private bool forceClose;
        private bool pageLoaded;
        private string lastBrowserMetrics = "{}";
        private CoreWebView2Environment webViewEnvironment;
        private DialogWindow dialogWindow;
        private DispatcherFrame dialogFrame;

        public event EventHandler PageLoaded;
        public event Action<string> WebMessage;
        public event EventHandler<Rdv3FormClosingEventArgs> ClosingRequested;

        public int ExitCode { get; private set; }

        public MainWindow(Rdv3Screen screen)
        {
            targetClientWidth = Math.Max(480.0, screen.StartWidth);
            targetClientHeight = Math.Max(300.0, screen.StartHeight);
            Title = "Reader Data Viewer";
            Width = targetClientWidth;
            Height = targetClientHeight;
            MinWidth = 480;
            MinHeight = 300;
            // Windows draws the frame and the caption. Moving, resizing,
            // snapping, minimise/maximise and the system menu all come from
            // the window manager; the page only draws the client area.
            WindowStyle = WindowStyle.SingleBorderWindow;
            ResizeMode = ResizeMode.CanResize;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(Color.FromRgb(
                Win98.Background.R,
                Win98.Background.G,
                Win98.Background.B));
            UseLayoutRounding = true;
            SnapsToDevicePixels = true;

            Grid root = new Grid();
            webView = new WebView2();
            webView.DefaultBackgroundColor =
                Win98.Background;
            webView.SetValue(UIElement.OpacityProperty, 0.0);
            root.Children.Add(webView);
            Content = root;

            // Fitting the client to the configured size resizes the window
            // once the page has loaded. Sit off screen until that is done so
            // the window is never seen at the wrong size.
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = -32000;
            Top = -32000;

            if (App.IsProbe)
            {
                Opacity = 0;
                ShowInTaskbar = false;
                ShowActivated = false;
            }

            Loaded += OnLoaded;
            Closing += OnClosing;
            Closed += OnClosed;
        }

        protected override void OnSourceInitialized(EventArgs eventArgs)
        {
            base.OnSourceInitialized(eventArgs);
            WindowCaption.Apply(this);
        }

        public Rect ScreenBounds
        {
            get
            {
                Point p = PointToScreen(new Point(0, 0));
                return new Rect(p.X, p.Y, ActualWidth, ActualHeight);
            }
        }

        public void PostJson(string json)
        {
            if (!pageLoaded || webView.CoreWebView2 == null) { return; }
            webView.CoreWebView2.PostWebMessageAsJson(json);
        }

        public void ShowSurface()
        {
            webView.SetValue(UIElement.OpacityProperty, 1.0);
        }

        // One owned window holds the dialog surface for the life of the app.
        // Building a WebView2 per dialog cost a browser start-up every time,
        // so the window is warmed once and afterwards only shown and hidden.
        public void PrewarmDialogSurface(string initJson)
        {
            if (webViewEnvironment == null || dialogWindow != null) { return; }
            DialogWindow dialog = new DialogWindow(
                this,
                webViewEnvironment,
                TrustedHost,
                StartPage + "#dialog",
                Path.Combine(App.BaseDirectory, "web"),
                initJson);
            dialogWindow = dialog;
            dialog.WebMessage += RaiseWebMessage;
            dialog.Failed += delegate { DiscardDialogSurface(); };
            dialog.Warm();
        }

        public void ShowDialogSurface(string initJson, string openJson)
        {
            PrewarmDialogSurface(initJson);
            DialogWindow dialog = dialogWindow;
            if (dialog == null) { return; }
            dialogFrame = new DispatcherFrame();
            IsEnabled = false;
            try
            {
                dialog.OpenWith(openJson);
                Dispatcher.PushFrame(dialogFrame);
            }
            finally
            {
                // Never leave the main window disabled, whatever went wrong.
                dialogFrame = null;
                IsEnabled = true;
                Activate();
            }
        }

        public void CloseDialogSurface()
        {
            // The picker has an inner dispatcher frame. Closing only the
            // dialog frame leaves that frame running and the owner disabled.
            Rdv3PickerForm.CancelCurrent();
            if (dialogWindow != null) { dialogWindow.HideSurface(); }
            if (dialogFrame != null) { dialogFrame.Continue = false; }
        }

        private void DiscardDialogSurface()
        {
            Rdv3PickerForm.CancelCurrent();
            DialogWindow dialog = dialogWindow;
            dialogWindow = null;
            if (dialog != null)
            {
                dialog.WebMessage -= RaiseWebMessage;
                dialog.CloseSurface();
            }
            if (dialogFrame != null) { dialogFrame.Continue = false; }
        }

        public void SizeDialogSurface(double width, double height, string title)
        {
            if (dialogWindow != null) { dialogWindow.FitTo(width, height, title); }
        }

        // Messages that belong to an open dialog go to that surface; everything
        // else keeps flowing to the main page so it stays current underneath.
        public void PostSurfaceJson(string json)
        {
            if (dialogWindow != null && dialogWindow.IsVisible) { dialogWindow.PostJson(json); }
            else { PostJson(json); }
        }

        private void RaiseWebMessage(string json)
        {
            Action<string> handler = WebMessage;
            if (handler != null) { handler(json); }
        }

        public void WindowCommand(string command)
        {
            if (command == "minimize") { WindowState = WindowState.Minimized; }
            else if (command == "maximize")
            {
                WindowState = WindowState == WindowState.Maximized
                    ? WindowState.Normal : WindowState.Maximized;
            }
            else if (command == "close") { Close(); }
        }

        public void ForceClose(int exitCode)
        {
            ExitCode = exitCode;
            forceClose = true;
            Close();
        }

        public void CaptureToFile(string path)
        {
            if (dialogWindow != null && dialogWindow.IsVisible)
            {
                dialogWindow.CaptureToFile(path);
                return;
            }
            string fullPath = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory)) { Directory.CreateDirectory(directory); }

            DispatcherFrame frame = new DispatcherFrame();
            Exception failure = null;
            CaptureAsync(fullPath, frame, delegate(Exception error)
            {
                failure = error;
            });
            Dispatcher.PushFrame(frame);
            if (failure != null) { throw failure; }

            if (App.IsProbe && !string.Equals(
                Environment.GetEnvironmentVariable("RDV_WEBVIEW2_PROBE_KEEP_OPEN"),
                "1",
                StringComparison.Ordinal))
            {
                string metrics = string.Format(
                    CultureInfo.InvariantCulture,
                    "window={0}x{1}; webview={2}x{3}",
                    ActualWidth,
                    ActualHeight,
                    webView.ActualWidth,
                    webView.ActualHeight) + "; browser=" + lastBrowserMetrics;
                File.WriteAllText(
                    fullPath + ".txt",
                    metrics,
                    new System.Text.UTF8Encoding(false));
                Dispatcher.BeginInvoke(new Action(delegate
                {
                    ExitCode = 0;
                    WindowCommand("close");
                }), DispatcherPriority.Background);
            }
        }

        private async void CaptureAsync(
            string path,
            DispatcherFrame frame,
            Action<Exception> completed)
        {
            Exception failure = null;
            try
            {
                await Task.Delay(120);
                lastBrowserMetrics = await webView.CoreWebView2.ExecuteScriptAsync(
                    "({viewportWidth:innerWidth,viewportHeight:innerHeight," +
                    "stageWidth:document.querySelector('.stage').offsetWidth," +
                    "stageHeight:document.querySelector('.stage').offsetHeight," +
                    "windowWidth:document.querySelector('.win').offsetWidth," +
                    "windowHeight:document.querySelector('.win').offsetHeight," +
                    "gap:getComputedStyle(document.querySelector('.stage')).getPropertyValue('--card-gap').trim()," +
                    "padding:getComputedStyle(document.querySelector('.client')).padding," +
                    "sections:document.querySelectorAll('.stack>:not(.sep)').length," +
                    "modal:(document.querySelector('.veil.show')||{}).id||''})");
                using (FileStream stream = new FileStream(
                    path,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.None))
                {
                    await webView.CoreWebView2.CapturePreviewAsync(
                        CoreWebView2CapturePreviewImageFormat.Png,
                        stream);
                }
            }
            catch (Exception exception) { failure = exception; }
            completed(failure);
            frame.Continue = false;
        }

        private async void OnLoaded(object sender, RoutedEventArgs eventArgs)
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(
                        Environment.SpecialFolder.LocalApplicationData),
                    "ReaderDataViewer",
                    "WebView2Cache");
                CoreWebView2Environment environment =
                    await CoreWebView2Environment.CreateAsync(
                        null,
                        userDataFolder,
                        null);
                webViewEnvironment = environment;
                await webView.EnsureCoreWebView2Async(environment);
                webView.ZoomFactor = 1.0;

                string webDirectory = Path.Combine(App.BaseDirectory, "web");
                string indexPath = Path.Combine(webDirectory, "index.html");
                string scriptPath = Path.Combine(webDirectory, "app.js");
                if (!File.Exists(indexPath) || !File.Exists(scriptPath))
                {
                    throw new FileNotFoundException(
                        "The Reader Data Viewer web surface is incomplete.");
                }

                webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    TrustedHost,
                    webDirectory,
                    CoreWebView2HostResourceAccessKind.Allow);
                webView.CoreWebView2.NavigationStarting += OnNavigationStarting;
                webView.CoreWebView2.NewWindowRequested += OnNewWindowRequested;
                webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
                webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                webView.CoreWebView2.Navigate(StartPage);
            }
            catch (Exception exception)
            {
                Fail("The WebView2 surface could not be initialized.", exception);
            }
        }

        private void OnNavigationStarting(
            object sender,
            CoreWebView2NavigationStartingEventArgs eventArgs)
        {
            if (!eventArgs.Uri.StartsWith(
                "https://" + TrustedHost + "/",
                StringComparison.OrdinalIgnoreCase))
            {
                eventArgs.Cancel = true;
            }
        }

        private async void OnNavigationCompleted(
            object sender,
            CoreWebView2NavigationCompletedEventArgs eventArgs)
        {
            try
            {
                if (!eventArgs.IsSuccess)
                {
                    throw new InvalidOperationException(
                        "The Reader Data Viewer page did not load: " +
                        eventArgs.WebErrorStatus.ToString());
                }
                await FitConfiguredClient();
                if (!App.IsProbe) { CentreOnScreen(); }
                string result = await webView.CoreWebView2.ExecuteScriptAsync(
                    "Boolean(document.querySelector('.stage .win') && " +
                    "!document.querySelector('.demo') && " +
                    "!document.querySelector('.notes') && " +
                    "window.rdvBridge)");
                if (!string.Equals(result, "true", StringComparison.Ordinal))
                {
                    throw new InvalidDataException(
                        "The approved v13 product surface was not found.");
                }
                pageLoaded = true;
                EventHandler handler = PageLoaded;
                if (handler != null) { handler(this, EventArgs.Empty); }
            }
            catch (Exception exception)
            {
                Fail("The v13 page could not be verified.", exception);
            }
        }

        private void CentreOnScreen()
        {
            Rect area = SystemParameters.WorkArea;
            Left = area.Left + (area.Width - ActualWidth) / 2;
            Top = area.Top + (area.Height - ActualHeight) / 2;
        }

        private async Task FitConfiguredClient()
        {
            Width += targetClientWidth - webView.ActualWidth;
            Height += targetClientHeight - webView.ActualHeight;
            await Task.Delay(60);

            // WebView2 rounds its composition bounds to physical pixels.  On
            // fractional DPI that can leave innerHeight one CSS pixel away
            // from the value declared in settings.json, even when the WPF
            // control's desired size is exact.  Correct against the browser
            // viewport as well so the settings value remains authoritative.
            for (int attempt = 0; attempt < 2; attempt++)
            {
                string widthText = await webView.CoreWebView2.ExecuteScriptAsync(
                    "window.innerWidth");
                string heightText = await webView.CoreWebView2.ExecuteScriptAsync(
                    "window.innerHeight");
                double browserWidth;
                double browserHeight;
                if (!double.TryParse(
                    widthText,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out browserWidth) ||
                    !double.TryParse(
                    heightText,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out browserHeight))
                {
                    return;
                }

                double widthDelta = targetClientWidth - browserWidth;
                double heightDelta = targetClientHeight - browserHeight;
                if (Math.Abs(widthDelta) < 0.1 &&
                    Math.Abs(heightDelta) < 0.1)
                {
                    return;
                }
                Width += widthDelta;
                Height += heightDelta;
                await Task.Delay(60);
            }
        }

        private void OnNewWindowRequested(object sender, CoreWebView2NewWindowRequestedEventArgs eventArgs)
        { eventArgs.Handled = true; }

        private void OnWebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            if (!eventArgs.Source.StartsWith("https://" + TrustedHost + "/", StringComparison.OrdinalIgnoreCase)) { return; }
            Action<string> handler = WebMessage;
            if (handler != null) { handler(eventArgs.WebMessageAsJson); }
        }

        private void OnClosing(object sender, CancelEventArgs eventArgs)
        {
            if (forceClose) { return; }
            EventHandler<Rdv3FormClosingEventArgs> handler = ClosingRequested;
            if (handler == null) { return; }
            Rdv3FormClosingEventArgs args = new Rdv3FormClosingEventArgs();
            handler(this, args);
            eventArgs.Cancel = args.Cancel;
        }

        private void Fail(string message, Exception exception)
        {
            ExitCode = App.IsProbe ? 4 : 3;
            App.ShowStartupMessage(message, exception);
            forceClose = true;
            Close();
        }

        private void OnClosed(object sender, EventArgs eventArgs)
        {
            pageLoaded = false;
            DiscardDialogSurface();
            if (webView.CoreWebView2 != null)
            {
                webView.CoreWebView2.NavigationStarting -= OnNavigationStarting;
                webView.CoreWebView2.NewWindowRequested -= OnNewWindowRequested;
                webView.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
                webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
            }
            webView.Dispose();
        }
    }

    // An ordinary owned window that hosts the same page in dialog mode.
    // Windows draws the frame and the caption and enforces the modal
    // behaviour; the page only draws the dialog body.
    public sealed class DialogWindow : Window
    {
        private readonly WebView2 webView = new WebView2();
        private readonly CoreWebView2Environment environment;
        private readonly string trustedHost;
        private readonly string startPage;
        private readonly string webDirectory;
        private bool pageLoaded;
        private bool sized;
        private bool closing;
        private double frameWidth = -1;
        private double frameHeight = -1;
        private readonly string initJson;
        private string pendingOpen;

        public event Action<string> WebMessage;
        public event Action Failed;

        public DialogWindow(
            Window owner,
            CoreWebView2Environment environment,
            string trustedHost,
            string startPage,
            string webDirectory,
            string initJson)
        {
            this.initJson = initJson;
            this.environment = environment;
            this.trustedHost = trustedHost;
            this.startPage = startPage;
            this.webDirectory = webDirectory;

            Owner = owner;
            Title = owner.Title;
            WindowStyle = WindowStyle.SingleBorderWindow;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            UseLayoutRounding = true;
            SnapsToDevicePixels = true;
            Background = new SolidColorBrush(Color.FromRgb(
                Win98.Background.R,
                Win98.Background.G,
                Win98.Background.B));

            // Lay the page out at a generous size off screen, then fit the
            // window to what the dialog actually measured before showing it
            // over the owner.
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = -32000;
            Top = -32000;
            Width = 1000;
            Height = 820;

            Grid root = new Grid();
            webView.DefaultBackgroundColor =
                Win98.Background;
            webView.SetValue(UIElement.OpacityProperty, 0.0);
            root.Children.Add(webView);
            Content = root;

            Loaded += OnLoaded;
            Closed += OnClosed;
        }

        protected override void OnSourceInitialized(EventArgs eventArgs)
        {
            base.OnSourceInitialized(eventArgs);
            WindowCaption.Apply(this);
        }

        public void PostJson(string json)
        {
            if (!pageLoaded || webView.CoreWebView2 == null) { return; }
            webView.CoreWebView2.PostWebMessageAsJson(json);
        }

        // Show the window off screen once so the WebView2 starts and the page
        // loads while nobody is waiting for it.
        public void Warm()
        {
            if (IsVisible) { return; }
            Left = -32000;
            Top = -32000;
            ShowActivated = false;
            Show();
        }

        public void OpenWith(string openJson)
        {
            sized = false;
            webView.SetValue(UIElement.OpacityProperty, 0.0);
            Left = -32000;
            Top = -32000;
            if (pageLoaded) { PostJson(openJson); } else { pendingOpen = openJson; }
            ShowActivated = true;
            if (!IsVisible) { Show(); }
        }

        public void HideSurface()
        {
            webView.SetValue(UIElement.OpacityProperty, 0.0);
            Hide();
        }

        // The frame around the page never changes for this window style, so
        // measure it once and set the size outright afterwards. Adding a delta
        // each time drifts, because the layout has not caught up with the
        // previous change when the next size is reported.
        public void FitTo(double clientWidth, double clientHeight, string title)
        {
            if (clientWidth < 1 || clientHeight < 1) { return; }
            if (!string.IsNullOrEmpty(title)) { Title = title; }
            if (frameWidth < 0)
            {
                UpdateLayout();
                if (webView.ActualWidth < 1 || webView.ActualHeight < 1) { return; }
                frameWidth = ActualWidth - webView.ActualWidth;
                frameHeight = ActualHeight - webView.ActualHeight;
            }
            Width = clientWidth + frameWidth;
            Height = clientHeight + frameHeight;
            Left = Owner.Left + (Owner.ActualWidth - Width) / 2;
            Top = Owner.Top + (Owner.ActualHeight - Height) / 2;
            sized = true;
            webView.SetValue(UIElement.OpacityProperty, 1.0);
            Activate();
        }

        public void CloseSurface()
        {
            closing = true;
            Close();
        }

        public void CaptureToFile(string path)
        {
            if (!pageLoaded || webView.CoreWebView2 == null) { return; }
            string fullPath = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory)) { Directory.CreateDirectory(directory); }
            DispatcherFrame frame = new DispatcherFrame();
            CaptureAsync(fullPath, frame);
            Dispatcher.PushFrame(frame);
        }

        private async void CaptureAsync(string path, DispatcherFrame frame)
        {
            string probe = null;
            try
            {
                await Task.Delay(120);
                // Headless checks need a way to drive the dialog itself.
                string script = Environment.GetEnvironmentVariable(
                    "RDV_HEADLESS_DIALOG_SCRIPT");
                if (App.IsProbe && !string.IsNullOrWhiteSpace(script))
                {
                    probe = await webView.CoreWebView2.ExecuteScriptAsync(script);
                    await Task.Delay(500);
                    probe += " | " + await webView.CoreWebView2.ExecuteScriptAsync(
                        "(typeof window.__readback==='function')?String(window.__readback()):"
                        + "((typeof window.__probe==='undefined')?'-':String(window.__probe))");
                }
                using (FileStream stream = new FileStream(
                    path, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await webView.CoreWebView2.CapturePreviewAsync(
                        CoreWebView2CapturePreviewImageFormat.Png, stream);
                }
                File.WriteAllText(
                    path + ".txt",
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "dialog window={0}x{1}; webview={2}x{3}; sized={4}; title={5}; probe={6}",
                        ActualWidth, ActualHeight,
                        webView.ActualWidth, webView.ActualHeight, sized,
                        Title, probe ?? "-"),
                    new System.Text.UTF8Encoding(false));
            }
            catch { }
            frame.Continue = false;
        }

        private async void OnLoaded(object sender, RoutedEventArgs eventArgs)
        {
            try
            {
                await webView.EnsureCoreWebView2Async(environment);
                webView.ZoomFactor = 1.0;
                webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    trustedHost,
                    webDirectory,
                    CoreWebView2HostResourceAccessKind.Allow);
                webView.CoreWebView2.NavigationStarting += OnNavigationStarting;
                webView.CoreWebView2.NewWindowRequested += OnNewWindowRequested;
                webView.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
                webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
                webView.CoreWebView2.Navigate(startPage);
            }
            catch
            {
                RaiseFailed();
            }
        }

        private void OnNavigationCompleted(
            object sender,
            CoreWebView2NavigationCompletedEventArgs eventArgs)
        {
            if (!eventArgs.IsSuccess) { RaiseFailed(); return; }
            pageLoaded = true;
            PostJson(initJson);
            if (pendingOpen != null)
            {
                string open = pendingOpen;
                pendingOpen = null;
                PostJson(open);
            }
            else
            {
                Hide();
            }
        }

        private void OnNavigationStarting(object sender, CoreWebView2NavigationStartingEventArgs eventArgs)
        {
            if (!eventArgs.Uri.StartsWith("https://" + trustedHost + "/", StringComparison.OrdinalIgnoreCase)) { eventArgs.Cancel = true; }
        }

        private void OnNewWindowRequested(object sender, CoreWebView2NewWindowRequestedEventArgs eventArgs)
        { eventArgs.Handled = true; }

        private void OnWebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            if (!eventArgs.Source.StartsWith("https://" + trustedHost + "/", StringComparison.OrdinalIgnoreCase)) { return; }
            Action<string> handler = WebMessage;
            if (handler != null) { handler(eventArgs.WebMessageAsJson); }
        }

        private void OnClosed(object sender, EventArgs eventArgs)
        {
            if (webView.CoreWebView2 != null)
            {
                webView.CoreWebView2.NavigationStarting -= OnNavigationStarting;
                webView.CoreWebView2.NewWindowRequested -= OnNewWindowRequested;
                webView.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
                webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
            }
            webView.Dispose();
            if (!closing) { RaiseFailed(); }
        }

        private void RaiseFailed()
        {
            Action handler = Failed;
            if (handler != null) { handler(); }
        }
    }
}
