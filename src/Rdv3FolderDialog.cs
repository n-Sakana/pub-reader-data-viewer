// The folder picker the operating system provides since Vista: the file
// dialog with FOS_PICKFOLDERS, owned by the window that asked for it. The
// shell's tree-style BrowseForFolder looked nothing like the file dialogs next
// to it and, having no owner, sank behind the app the moment the app was
// clicked -- leaving a window that could not be found and an app that would
// not answer. C# 5, no external libraries: the two COM interfaces are
// declared here.
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

public static class Rdv3FolderDialog
{
    private const uint FOS_NOCHANGEDIR = 0x00000008;
    private const uint FOS_PICKFOLDERS = 0x00000020;
    private const uint FOS_FORCEFILESYSTEM = 0x00000040;
    private const uint FOS_PATHMUSTEXIST = 0x00000800;
    private const uint SIGDN_FILESYSPATH = 0x80058000;
    private const int ERROR_CANCELLED = unchecked((int)0x800704C7);

    [ComImport, Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogRCW { }

    [ComImport, Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        [PreserveSig] int BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        [PreserveSig] int GetParent(out IShellItem ppsi);
        [PreserveSig] int GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        [PreserveSig] int GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        [PreserveSig] int Compare(IShellItem psi, uint hint, out int piOrder);
    }

    // IModalWindow, IFileDialog and IFileOpenDialog in vtable order.
    [ComImport, Guid("d57c7288-d4ad-4768-be02-9d969532d960"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr hwndOwner);
        [PreserveSig] int SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        [PreserveSig] int SetFileTypeIndex(uint iFileType);
        [PreserveSig] int GetFileTypeIndex(out uint piFileType);
        [PreserveSig] int Advise(IntPtr pfde, out uint pdwCookie);
        [PreserveSig] int Unadvise(uint dwCookie);
        [PreserveSig] int SetOptions(uint fos);
        [PreserveSig] int GetOptions(out uint pfos);
        [PreserveSig] int SetDefaultFolder(IShellItem psi);
        [PreserveSig] int SetFolder(IShellItem psi);
        [PreserveSig] int GetFolder(out IShellItem ppsi);
        [PreserveSig] int GetCurrentSelection(out IShellItem ppsi);
        [PreserveSig] int SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        [PreserveSig] int GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        [PreserveSig] int SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        [PreserveSig] int SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        [PreserveSig] int SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        [PreserveSig] int GetResult(out IShellItem ppsi);
        [PreserveSig] int AddPlace(IShellItem psi, int fdap);
        [PreserveSig] int SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        [PreserveSig] int Close(int hr);
        [PreserveSig] int SetClientGuid(ref Guid guid);
        [PreserveSig] int ClearClientData();
        [PreserveSig] int SetFilter(IntPtr pFilter);
        [PreserveSig] int GetResults(out IntPtr ppenum);
        [PreserveSig] int GetSelectedItems(out IntPtr ppsai);
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
    private static extern int SHCreateItemFromParsingName([MarshalAs(UnmanagedType.LPWStr)] string pszPath,
        IntPtr pbc, ref Guid riid, out IShellItem ppv);

    // The folder the operator chose, or "" when the dialog was cancelled. Must
    // be called on the UI thread; the dialog is modal to the owner it is given.
    public static string Pick(Window owner, string title, string initial)
    {
        IFileOpenDialog dialog = (IFileOpenDialog)new FileOpenDialogRCW();
        try
        {
            uint options;
            dialog.GetOptions(out options);
            dialog.SetOptions(options | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM | FOS_PATHMUSTEXIST | FOS_NOCHANGEDIR);
            if (!string.IsNullOrEmpty(title)) { dialog.SetTitle(title); }
            string start = StartFolder(initial);
            if (start != null)
            {
                Guid shellItem = typeof(IShellItem).GUID;
                IShellItem item;
                if (SHCreateItemFromParsingName(start, IntPtr.Zero, ref shellItem, out item) == 0 && item != null)
                {
                    try { dialog.SetFolder(item); }
                    finally { Marshal.ReleaseComObject(item); }
                }
            }
            IntPtr hwnd = (owner == null) ? IntPtr.Zero : new WindowInteropHelper(owner).Handle;
            int hr = dialog.Show(hwnd);
            if (hr == ERROR_CANCELLED) { return ""; }
            if (hr != 0) { Marshal.ThrowExceptionForHR(hr); }
            IShellItem result;
            if (dialog.GetResult(out result) != 0 || result == null) { return ""; }
            try
            {
                string path;
                return (result.GetDisplayName(SIGDN_FILESYSPATH, out path) == 0 && path != null) ? path : "";
            }
            finally { Marshal.ReleaseComObject(result); }
        }
        finally { Marshal.ReleaseComObject(dialog); }
    }

    // The existing folder nearest to what the field holds, or null.
    private static string StartFolder(string initial)
    {
        if (string.IsNullOrEmpty(initial)) { return null; }
        try
        {
            string folder = Path.GetFullPath(initial);
            if (Directory.Exists(folder)) { return folder; }
            folder = Path.GetDirectoryName(folder);
            return (folder != null && Directory.Exists(folder)) ? folder : null;
        }
        catch (Exception) { return null; }
    }
}
