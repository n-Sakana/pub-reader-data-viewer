using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public sealed class ReaderWindowInfo {
 public long handle; public uint pid; public uint tid; public bool enabled; public bool visible; public bool hung; public string title; public int left; public int top; public int right; public int bottom;
}
public static class ReaderWindowProbe {
 [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h,IntPtr dc,uint flags);
 [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr value);
 public static void PhysicalPixels(){SetProcessDpiAwarenessContext(new IntPtr(-4));}
 public delegate bool EnumCallback(IntPtr h,IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumCallback callback,IntPtr p);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
 [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr h);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] static extern bool IsHungAppWindow(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
 [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
 struct Rect {public int L,T,R,B;}
 public static ReaderWindowInfo[] Read(int[] owners){
  HashSet<int> ids=new HashSet<int>(owners);var rows=new List<ReaderWindowInfo>();
  EnumWindows(delegate(IntPtr h,IntPtr unused){uint pid;uint tid=GetWindowThreadProcessId(h,out pid);if(!ids.Contains((int)pid))return true;StringBuilder text=new StringBuilder(256);GetWindowText(h,text,text.Capacity);Rect r;GetWindowRect(h,out r);rows.Add(new ReaderWindowInfo{handle=h.ToInt64(),pid=pid,tid=tid,enabled=IsWindowEnabled(h),visible=IsWindowVisible(h),hung=IsHungAppWindow(h),title=text.ToString(),left=r.L,top=r.T,right=r.R,bottom=r.B});return true;},IntPtr.Zero);
  return rows.ToArray();
 }
}
