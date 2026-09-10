using System;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Runtime.InteropServices;

// A real UI Automation ValuePattern provider, isolated from user applications.
public static class ReaderWatchSource
{
    [DllImport("user32.dll")] private static extern bool AllowSetForegroundWindow(uint processId);
    [STAThread]
    public static void Main(string[] args)
    {
        Window form = new Window { Title = args[0], Left = 100, Top = 100, Width = 620, Height = 180 };
        StackPanel panel = new StackPanel { Margin = new Thickness(12) };
        panel.Children.Add(new TextBlock { Text = "Reader UI Automation verification source", Height = 35 });
        TextBox input = new TextBox { Name = "ReaderNumber", Height = 30 };
        AutomationProperties.SetAutomationId(input, "ReaderNumber");
        panel.Children.Add(input); form.Content = panel;
        Button grant = new Button { Content = "Allow Reader foreground", Height = 30 };
        AutomationProperties.SetAutomationId(grant, "AllowReaderForeground");
        grant.Click += delegate { grant.Content = AllowSetForegroundWindow(uint.Parse(args[1])) ? "Permission granted" : "Permission refused"; };
        panel.Children.Add(grant);
        new Application().Run(form);
    }
}
