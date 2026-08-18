# mod_child.ps1 -- VERIFICATION ONLY. Generic runner: imports one .bas into an
# Excel this process owns, compile-checks it while the instance is still
# INVISIBLE, then arms it and watches the result file. Every COM call is here so
# that a modal blocks this process and not the watchdog.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Bas,
    [Parameter(Mandatory=$true)][string] $Res,
    [Parameter(Mandatory=$true)][string] $PidFile,
    [Parameter(Mandatory=$true)][string] $PingProc,
    [Parameter(Mandatory=$true)][string] $ArmProc,
    [Parameter(Mandatory=$true)][string] $DoneMark,
    [string] $A2 = '',
    [switch] $Visible,
    [switch] $Uia,
    [int] $WaitSec = 120
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
$wb.Worksheets.Item(1).Range('A2').Value2 = $A2
if ($Uia) {
    [void]$wb.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)
}
[void]$wb.VBProject.VBComponents.Import($Bas)

$xl.Run($PingProc)                      # compile check, still invisible
Write-Host 'compile check passed'
if ($Visible) {
    $xl.Visible = $true
    [void][Fg.W]::ShowWindow([IntPtr]$xl.Hwnd, 9)
    [void][Fg.W]::SetForegroundWindow([IntPtr]$xl.Hwnd)
    Start-Sleep -Milliseconds 300
}
$xl.Run($ArmProc)
Write-Host 'armed; watching the result file only'

$deadline = (Get-Date).AddSeconds($WaitSec)
$done = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 300
    if ((Read-Shared $Res) -match [regex]::Escape($DoneMark)) { $done = $true; break }
}
if ($done) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    Start-Sleep -Milliseconds 600
}
foreach ($id in $mine) {
    if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force }
}
if (-not $done) { exit 2 }
exit 0
