using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public static class ProtectionNativeProbe
{
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr window, int command);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr window);
    [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr window, int index);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll")] private static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);
    private struct Rect { public int Left, Top, Right, Bottom; }
    public static void FocusSource(IntPtr window)
    {
        if (GetForegroundWindow() == window || SetForegroundWindow(window)) { return; }
        // Establish the test precondition with an Alt key release. Never call
        // this for Reader or after setting the simulated number: the product
        // must perform its own activation, measured by GetForegroundWindow.
        keybd_event(0x12, 0, 0, UIntPtr.Zero);
        keybd_event(0x12, 0, 2, UIntPtr.Zero);
        SetForegroundWindow(window);
    }
    public static bool Topmost(IntPtr window) { return (GetWindowLong(window, -20) & 8) != 0; }
    public static bool OnScreen(ReaderWindowInfo window)
    {
        foreach (Screen screen in Screen.AllScreens)
        {
            Rectangle b = screen.Bounds;
            // Maximized windows include the invisible resize border.
            if (window.left >= b.Left - 10 && window.top >= b.Top - 10
                && window.right <= b.Right + 10 && window.bottom <= b.Bottom + 10) { return true; }
        }
        return false;
    }
    public static void Capture(long handle, string path)
    {
        IntPtr window = new IntPtr(handle); Rect rect;
        if (!GetWindowRect(window, out rect)) { throw new Exception("GetWindowRect failed"); }
        using (Bitmap bitmap = new Bitmap(rect.Right - rect.Left, rect.Bottom - rect.Top))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            IntPtr dc = graphics.GetHdc(); bool captured;
            try { captured = ReaderWindowProbe.PrintWindow(window, dc, 2); }
            finally { graphics.ReleaseHdc(dc); }
            if (!captured) { throw new Exception("PrintWindow failed"); }
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
