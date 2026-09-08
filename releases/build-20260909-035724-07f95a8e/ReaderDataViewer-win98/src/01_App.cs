using System;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace ReaderDataViewer
{
    public static class App
    {
        public static string BaseDirectory { get; private set; }

        public static string ProbeOutputPath
        {
            get
            {
                return Environment.GetEnvironmentVariable(
                    "RDV_WEBVIEW2_PROBE_OUTPUT");
            }
        }

        public static bool IsProbe
        {
            get { return !string.IsNullOrWhiteSpace(ProbeOutputPath); }
        }

        [STAThread]
        public static int Run(string baseDirectory)
        {
            BaseDirectory = Path.GetFullPath(baseDirectory);
            InstallAssemblyResolver();

            if (!IsWebView2Available())
            {
                ShowStartupMessage(
                    "Microsoft Edge WebView2 Runtime is not available.",
                    null);
                return 3;
            }

            if (IsProbe && string.IsNullOrWhiteSpace(
                Environment.GetEnvironmentVariable(
                    "RDV_HEADLESS_CAPTURE_PATH")))
            {
                Environment.SetEnvironmentVariable(
                    "RDV_HEADLESS_CAPTURE_PATH",
                    ProbeOutputPath);
            }

            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                return RunCore();
            }

            Exception startupError = null;
            int exitCode = 3;
            Thread thread = new Thread(delegate()
            {
                try { exitCode = RunCore(); }
                catch (Exception exception) { startupError = exception; }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (startupError != null)
            {
                ShowStartupMessage(
                    "Reader Data Viewer could not open its window.",
                    startupError);
            }
            return exitCode;
        }

        private static int RunCore()
        {
            string settingsPath = Path.Combine(BaseDirectory, "settings.json");
            return Rdv3Program.Run(
                settingsPath,
                BaseDirectory,
                "",
                "",
                "",
                0.0);
        }

        public static void ShowStartupMessage(string message, Exception error)
        {
            WriteStartupError(message, error);
            if (IsProbe) { return; }
            if (error != null)
            {
                message = message + Environment.NewLine +
                    Environment.NewLine + error.Message;
            }
            MessageBox.Show(
                message,
                "Reader Data Viewer",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private static void InstallAssemblyResolver()
        {
            string libraryDirectory = Path.Combine(BaseDirectory, "lib");
            AppDomain.CurrentDomain.AssemblyResolve += delegate(
                object sender,
                ResolveEventArgs args)
            {
                string assemblyName = new AssemblyName(args.Name).Name;
                Assembly[] loaded = AppDomain.CurrentDomain.GetAssemblies();
                for (int index = 0; index < loaded.Length; index++)
                {
                    if (string.Equals(
                        loaded[index].GetName().Name,
                        assemblyName,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return loaded[index];
                    }
                }

                string assemblyPath = Path.Combine(
                    libraryDirectory,
                    assemblyName + ".dll");
                if (File.Exists(assemblyPath))
                {
                    return Assembly.Load(File.ReadAllBytes(assemblyPath));
                }
                return null;
            };
        }

        private static bool IsWebView2Available()
        {
            try
            {
                string version =
                    CoreWebView2Environment.GetAvailableBrowserVersionString();
                return !string.IsNullOrEmpty(version);
            }
            catch { return false; }
        }

        private static void WriteStartupError(string message, Exception error)
        {
            try
            {
                string localData = Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData);
                if (string.IsNullOrWhiteSpace(localData)) { return; }
                string logDirectory = Path.Combine(
                    localData,
                    "ReaderDataViewer",
                    "logs");
                Directory.CreateDirectory(logDirectory);
                string logPath = Path.Combine(
                    logDirectory,
                    "reader-data-viewer_" +
                        DateTime.Now.ToString("yyyyMMdd") + ".log");
                string detail = error == null ? message : error.ToString();
                string line = string.Format(
                    "[{0:HH:mm:ss}] [ERROR] desktop startup error: {1}{2}",
                    DateTime.Now,
                    detail,
                    Environment.NewLine);
                File.AppendAllText(
                    logPath,
                    line,
                    new UTF8Encoding(false));
            }
            catch { }
        }
    }
}
