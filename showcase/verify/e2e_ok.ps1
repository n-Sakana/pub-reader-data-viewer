# e2e_ok.ps1 -- VERIFICATION ONLY. Waits for a modal dialog to appear in the
# given Excel process and presses its default button (OK).
#
# Legacy helper for modal-dialog probes. The current 30-second benchmark has no
# completion dialog, so this is not part of its acceptance path. It polls
# because older OnTime-driven probes raised their dialog up to a second after
# the button was pressed.
#
# Win32 is used HERE, in a verification script, never in the shipped VBA.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][int] $ExcelPid,
    [int] $TimeoutSec = 30,
    [switch] $Cancel
)
$ErrorActionPreference = 'Stop'

# VBA の MsgBox はダイアログの標準クラス #32770 ではない（実測：画面には
# 確かに出ているのに FindWindowEx(..,'#32770',..) が 1 つも返さなかった）。
# だからクラスでは探さない。プロセスが同じで、Excel の本体窓ではなく、
# 見えている最上位の窓、で拾う。
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class OkW {
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern void keybd_event(byte k, byte s, uint f, IntPtr e);
    public static string Text(IntPtr h) { var b = new StringBuilder(512); GetWindowText(h, b, 512); return b.ToString(); }
    public static string Cls(IntPtr h) { var b = new StringBuilder(256); GetClassName(h, b, 256); return b.ToString(); }
    public static IntPtr[] Dialogs(int pid) {
        var found = new List<IntPtr>();
        EnumWindows((h, l) => {
            int p; GetWindowThreadProcessId(h, out p);
            if (p == pid && IsWindowVisible(h) && Cls(h) != "XLMAIN") found.Add(h);
            return true;
        }, IntPtr.Zero);
        return found.ToArray();
    }
}
'@

$cmd = if ($Cancel) { 2 } else { 1 }        # IDOK = 1, IDCANCEL = 2
$deadline = (Get-Date).AddSeconds($TimeoutSec)
while ((Get-Date) -lt $deadline) {
    foreach ($hw in [OkW]::Dialogs($ExcelPid)) {
        $cls = [OkW]::Cls($hw)
        $txt = [OkW]::Text($hw)
        if ($cls -eq 'XLMAIN' -or $txt -eq '') { continue }
        Write-Host "dialog [$txt] class=$cls -> pressing $(if ($Cancel) {'Cancel'} else {'OK'})"
        # 実キーで押す。WM_COMMAND(IDOK) を送るだけでは VBA の MsgBox が閉じ
        # ないことがある（実測：送信は成功したのにダイアログが残り続けた）。
        # 前面に出して Enter / Esc を打つのが、人がやることと同じで確実。
        [void][OkW]::SetForegroundWindow($hw)
        Start-Sleep -Milliseconds 350
        $vk = if ($Cancel) { 0x1B } else { 0x0D }      # VK_ESCAPE / VK_RETURN
        [OkW]::keybd_event($vk, 0, 0, [System.IntPtr]::Zero)
        Start-Sleep -Milliseconds 60
        [OkW]::keybd_event($vk, 0, 2, [System.IntPtr]::Zero)
        Start-Sleep -Milliseconds 400
        # それでも残っていたら WM_COMMAND も試す
        if ([OkW]::Dialogs($ExcelPid) -contains $hw) {
            [void][OkW]::SendMessage($hw, 0x0111, [System.IntPtr]$cmd, [System.IntPtr]::Zero)
        }
        return
    }
    Start-Sleep -Milliseconds 200
}
Write-Host "no dialog appeared within $TimeoutSec s"
