using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace ReaderDataViewer
{
    public static class App
    {
        public static string BaseDirectory { get; private set; }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern IntPtr LoadLibraryW(string path);

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
            Rdv3Log.Begin("window", "app=" + baseDirectory);
            int code = 3;
            try { code = RunApplication(baseDirectory); return code; }
            catch (Exception error) { Rdv3Log.Error("window lifetime", error); throw; }
            finally { Rdv3Log.End(code); }
        }

        private static int RunApplication(string baseDirectory)
        {
            BaseDirectory = Path.GetFullPath(baseDirectory);
            InstallAssemblyResolver();

            Exception runtimeError;
            if (!IsWebView2Available(out runtimeError))
            {
                ShowStartupMessage(
                    "画面の起動に必要なWebView2を読み込めませんでした。",
                    runtimeError);
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

        private static bool IsWebView2Available(out Exception error)
        {
            error = null;
            try
            {
                // PATH uses semicolons as separators even when a directory
                // name contains one. Keep the bundled loader loaded by its
                // literal full path for the lifetime of the WebView2 process.
                string loader = Path.Combine(BaseDirectory, "lib", "WebView2Loader.dll");
                if (LoadLibraryW(loader) == IntPtr.Zero)
                { throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error(), "同梱の画面起動部品を読み込めません: " + loader); }
                string version =
                    CoreWebView2Environment.GetAvailableBrowserVersionString();
                return !string.IsNullOrEmpty(version);
            }
            catch (Exception failure) { error = failure; return false; }
        }

        private static void WriteStartupError(string message, Exception error)
        {
            if (error == null) { Rdv3Log.Feedback("ERROR", "startup: " + message); }
            else { Rdv3Log.Error("startup: " + message, error); }
        }
    }
}
