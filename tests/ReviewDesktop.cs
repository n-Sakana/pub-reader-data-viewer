using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

public static class ReviewDesktop
{
    public delegate bool EnumCallback(IntPtr hwnd, IntPtr state);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumCallback callback, IntPtr state);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] static extern void SwitchToThisWindow(IntPtr hwnd, bool altTab);
    [DllImport("user32.dll")] static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int w, int h, uint flags);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] static extern void mouse_event(uint flags, uint x, uint y, uint data, UIntPtr extra);
    [DllImport("user32.dll")] static extern void keybd_event(byte key, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out Rect rect);
    [DllImport("user32.dll")] static extern bool MoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
    struct Rect { public int Left, Top, Right, Bottom; }
    public static void Physical() { SetThreadDpiAwarenessContext(new IntPtr(-4)); }
    public static bool IsOwnedVisible(IntPtr hwnd, int owner)
    { uint pid; GetWindowThreadProcessId(hwnd, out pid); return pid == owner && IsWindowVisible(hwnd); }
    public static long[] Windows(int owner)
    {
        List<long> result = new List<long>();
        EnumWindows(delegate(IntPtr hwnd, IntPtr unused) {
            uint pid; GetWindowThreadProcessId(hwnd, out pid);
            if (pid == owner && IsWindowVisible(hwnd)) { result.Add(hwnd.ToInt64()); }
            return true;
        }, IntPtr.Zero);
        return result.ToArray();
    }
    public static void Focus(IntPtr hwnd)
    {
        uint unused; uint other = GetWindowThreadProcessId(GetForegroundWindow(), out unused);
        uint self = GetCurrentThreadId(); bool attached = other != self && AttachThreadInput(self, other, true);
        try { SetForegroundWindow(hwnd); } finally { if (attached) { AttachThreadInput(self, other, false); } }
        Thread.Sleep(100);
        if (GetForegroundWindow() != hwnd) { SwitchToThisWindow(hwnd, true); Thread.Sleep(100); }
        if (GetForegroundWindow() != hwnd) {
            SetWindowPos(hwnd, new IntPtr(-1), 0, 0, 0, 0, 0x13);
            Rect r; GetWindowRect(hwnd, out r);
            SetCursorPos(r.Left + 100, r.Top + 20);
            mouse_event(2, 0, 0, 0, UIntPtr.Zero); mouse_event(4, 0, 0, 0, UIntPtr.Zero);
            Thread.Sleep(100);
            SetWindowPos(hwnd, new IntPtr(-2), 0, 0, 0, 0, 0x13);
        }
        if (GetForegroundWindow() != hwnd) { throw new Exception("Could not foreground the owned test window; actual=" + GetForegroundWindow().ToInt64()); }
    }
    public static void Move(IntPtr hwnd, int x, int y, int w, int h) { MoveWindow(hwnd, x, y, w, h, true); }
    public static void Chord(byte[] keys)
    {
        try { foreach (byte key in keys) { keybd_event(key, 0, 0, UIntPtr.Zero); } Thread.Sleep(350); }
        finally { for (int i=keys.Length-1;i>=0;i--) { keybd_event(keys[i], 0, 2, UIntPtr.Zero); } }
    }
    public static void Click(AutomationElement element)
    {
        var p = element.Current;
        if (p.IsOffscreen || !p.IsEnabled || p.BoundingRectangle.IsEmpty || p.BoundingRectangle.Width < 1 || p.BoundingRectangle.Height < 1)
        { throw new Exception("Control is not visible/enabled: " + p.Name); }
        int x = (int)(p.BoundingRectangle.Left + p.BoundingRectangle.Width / 2);
        int y = (int)(p.BoundingRectangle.Top + p.BoundingRectangle.Height / 2);
        bool visible = false;
        foreach (Screen screen in Screen.AllScreens) { if (screen.Bounds.Contains(x, y)) { visible = true; } }
        if (!visible) { throw new Exception("Control is outside the actual displays"); }
        SetCursorPos(x, y); mouse_event(2, 0, 0, 0, UIntPtr.Zero); mouse_event(4, 0, 0, 0, UIntPtr.Zero);
    }
    public static void Capture(IntPtr hwnd, string path)
    {
        Rect r; GetWindowRect(hwnd, out r);
        using (Bitmap bitmap = new Bitmap(r.Right-r.Left, r.Bottom-r.Top))
        using (Graphics graphics = Graphics.FromImage(bitmap)) {
            graphics.CopyFromScreen(r.Left, r.Top, 0, 0, bitmap.Size);
            bitmap.Save(path, ImageFormat.Png);
        }
    }
}
