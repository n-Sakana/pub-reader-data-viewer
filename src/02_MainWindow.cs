using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
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
            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.CanResize;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            Background = new SolidColorBrush(Color.FromRgb(212, 208, 200));
            UseLayoutRounding = true;
            SnapsToDevicePixels = true;

            Grid root = new Grid();
            webView = new WebView2();
            webView.DefaultBackgroundColor =
                System.Drawing.Color.FromArgb(212, 208, 200);
            webView.SetValue(UIElement.OpacityProperty, 0.0);
            root.Children.Add(webView);
            Content = root;

            if (App.IsProbe)
            {
                WindowStartupLocation = WindowStartupLocation.Manual;
                Left = -32000;
                Top = -32000;
                Opacity = 0;
                ShowInTaskbar = false;
                ShowActivated = false;
            }

            Loaded += OnLoaded;
            Closing += OnClosing;
            Closed += OnClosed;
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

        public void WindowCommand(string command)
        {
            if (command == "minimize") { WindowState = WindowState.Minimized; }
            else if (command == "maximize")
            {
                WindowState = WindowState == WindowState.Maximized
                    ? WindowState.Normal : WindowState.Maximized;
            }
            else if (command == "close") { Close(); }
            else if (command == "drag")
            {
                try { DragMove(); } catch (InvalidOperationException) { }
            }
        }

        public void ForceClose(int exitCode)
        {
            ExitCode = exitCode;
            forceClose = true;
            Close();
        }

        public void CaptureToFile(string path)
        {
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

        private void OnWebMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
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
            if (webView.CoreWebView2 != null)
            {
                webView.CoreWebView2.NavigationStarting -= OnNavigationStarting;
                webView.CoreWebView2.NavigationCompleted -= OnNavigationCompleted;
                webView.CoreWebView2.WebMessageReceived -= OnWebMessageReceived;
            }
            webView.Dispose();
        }
    }
}
