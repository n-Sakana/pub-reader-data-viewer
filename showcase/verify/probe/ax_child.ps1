# ax_child.ps1 -- VERIFICATION ONLY. Same watchdog shape as probe_child.ps1:
# every COM call is here, Excel starts invisible, and the module is compile
# checked before anything is shown.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Bas,
    [Parameter(Mandatory=$true)][string] $Res,
    [Parameter(Mandatory=$true)][string] $PidFile,
    [switch] $Visible,
    [switch] $SaveFirst,
    [int] $WaitSec = 40
)
$ErrorActionPreference = 'Stop'

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
Write-Host "owned excel pid: $($mine -join ',')  visible=$Visible saveFirst=$SaveFirst"

$wb = $xl.Workbooks.Add()
$wb.Worksheets.Item(1).Range('A1').Value2 = $Res
[void]$wb.VBProject.VBComponents.Import($Bas)

if ($SaveFirst) {
    $tmp = Join-Path $env:TEMP ("pbax-" + [System.Guid]::NewGuid().ToString('N') + '.xlsm')
    $wb.SaveAs($tmp, 52)   # xlOpenXMLWorkbookMacroEnabled
    Write-Host "saved as $tmp"
}

$xl.Run('PbAxPing')
Write-Host 'compile check passed'
if ($Visible) { $xl.Visible = $true; Start-Sleep -Milliseconds 300 }

$xl.Run('PbAxArm')
$deadline = (Get-Date).AddSeconds($WaitSec)
$done = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 300
    if ((Read-Shared $Res) -match '=== ax probe done ===') { $done = $true; break }
}
if ($done) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    Start-Sleep -Milliseconds 500
}
foreach ($id in $mine) {
    if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force }
}
if (-not $done) { exit 2 }
exit 0
