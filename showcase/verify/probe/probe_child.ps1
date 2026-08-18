# probe_child.ps1 -- VERIFICATION ONLY. Runs every COM call, so that a modal
# (a VBA compile error is one) blocks THIS process and not the watchdog.
# The parent kills what the pid file lists if this does not come back.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Bas,
    [Parameter(Mandatory=$true)][string] $Res,
    [Parameter(Mandatory=$true)][string] $PidFile,
    [string] $Skip = '',
    [int] $WaitSec = 200
)
$ErrorActionPreference = 'Stop'

Add-Type -Namespace Fg -Name W -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
'@

function Read-Shared([string]$p) {
    if (-not (Test-Path $p)) { return '' }
    try {
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t  = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return '' }
}

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)

# INVISIBLE first. Everything that can raise a modal happens while there is no
# window for it to appear in front of the operator.
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
Start-Sleep -Milliseconds 400
$mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
          Where-Object { $before -notcontains $_ })
Set-Content -Path $PidFile -Value ($mine -join ',') -Encoding ascii
Write-Host "owned excel pid: $($mine -join ',')"

$wb = $xl.Workbooks.Add()
$wb.Worksheets.Item(1).Range('A1').Value2 = $Res
$wb.Worksheets.Item(1).Range('A2').Value2 = $Skip

# UIAutomationClient has to be an EARLY binding: its interfaces derive from
# IUnknown only, so CreateObject would assign but never call.
[void]$wb.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)
[void]$wb.VBProject.VBComponents.Import($Bas)

# THE COMPILE CHECK. If the project does not compile this call never returns and
# the parent kills this process -- with the Excel still invisible.
Write-Host 'compile check (invisible) ...'
$xl.Run('PbProbePing')
Write-Host 'compile check passed'

# only now is it safe to show anything
$xl.Visible = $true
[void][Fg.W]::ShowWindow([IntPtr]$xl.Hwnd, 9)
[void][Fg.W]::SetForegroundWindow([IntPtr]$xl.Hwnd)
Start-Sleep -Milliseconds 300

$xl.Run('PbProbeArm')
Write-Host 'armed; watching the result file only (no further COM calls)'

$deadline = (Get-Date).AddSeconds($WaitSec)
$done = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 400
    $txt = Read-Shared $Res
    if ($txt -match '=== probe done ===') { $done = $true; break }
}
if ($done) {
    Write-Host 'probe finished; quitting the owned Excel'
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    Start-Sleep -Milliseconds 600
}
foreach ($id in $mine) {
    if (Get-Process -Id $id -ErrorAction SilentlyContinue) {
        Write-Host "killing owned excel pid $id"
        Stop-Process -Id $id -Force
    }
}
if (-not $done) { exit 2 }
exit 0
