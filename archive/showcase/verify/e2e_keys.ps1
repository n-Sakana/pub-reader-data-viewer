# e2e_keys.ps1 -- VERIFICATION ONLY. Brings one owned window to the front and
# sends a keystroke, the way a person would press it. Used to exercise the
# Application.OnKey shortcuts the screen advertises.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][int] $ProcessId,
    [Parameter(Mandatory=$true)][string] $Keys,
    [int] $SettleMs = 800
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -Namespace Fg -Name W -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
'@
$p = Get-Process -Id $ProcessId -ErrorAction Stop
$h = $p.MainWindowHandle
if ($h -eq [IntPtr]::Zero) { throw "no main window on $ProcessId" }
[void][Fg.W]::ShowWindow($h, 9)
[void][Fg.W]::SetForegroundWindow($h)
Start-Sleep -Milliseconds $SettleMs
Write-Host "sending [$Keys] to pid $ProcessId"
[System.Windows.Forms.SendKeys]::SendWait($Keys)
Start-Sleep -Milliseconds $SettleMs
