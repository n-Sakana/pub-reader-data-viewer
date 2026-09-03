# run_shell.ps1 -- VERIFICATION ONLY. Runs a probe module the way the real app
# runs: a workbook on disk, opened by the SHELL. That is the one condition the
# COM-started probes cannot reproduce -- a shell start loads the user's add-ins
# and XLSTART, an automation start does not.
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Bas,
    [Parameter(Mandatory=$true)][string] $DoneMark,
    [string] $Tag = 'shell',
    [int] $TimeoutSec = 300
)
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$res  = Join-Path $here "res-$Tag.txt"
$book = Join-Path $here "probe-$Tag.xlsm"
foreach ($f in @($res, $book)) { if (Test-Path $f) { Remove-Item $f -Force } }

function Read-Shared([string]$p) {
    if (-not (Test-Path $p)) { return '' }
    try {
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t  = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return '' }
}

# 1. make the workbook with COM (this Excel is ours and is closed right after)
$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
Start-Sleep -Milliseconds 300
$mineA = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
           Where-Object { $before -notcontains $_ })
$xl.SheetsInNewWorkbook = 1
$wb = $xl.Workbooks.Add()
$wb.Worksheets.Item(1).Range('A1').Value2 = $res
[void]$wb.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)
[void]$wb.VBProject.VBComponents.Import($Bas)
$wb.SaveAs($book, 52)
$wb.Close($false)
$xl.Quit()
Start-Sleep -Milliseconds 600
foreach ($id in $mineA) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force } }
Write-Host "made $book"

# 2. open it the way a person does
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Start-Process -FilePath $book | Out-Null
$mine = @()
$dl = (Get-Date).AddSeconds(40)
while ((Get-Date) -lt $dl) {
    Start-Sleep -Milliseconds 400
    $mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
              Where-Object { $before -notcontains $_ })
    if ($mine.Count -gt 0) { break }
}
Write-Host "shell excel: $($mine -join ',')"
Start-Sleep -Seconds 4
$root = [System.Windows.Automation.AutomationElement]::RootElement
foreach ($id in $mine) {
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $id)
    $win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
    if ($null -eq $win) { continue }
    $bc = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)
    foreach ($b in $win.FindAll([System.Windows.Automation.TreeScope]::Descendants, $bc)) {
        if ($b.Current.Name -match 'コンテンツの有効化|Enable Content') {
            Write-Host "pressing [$($b.Current.Name)]"
            $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
            break
        }
    }
}

$dl = (Get-Date).AddSeconds($TimeoutSec)
$done = $false
while ((Get-Date) -lt $dl) {
    Start-Sleep -Milliseconds 500
    if ((Read-Shared $res) -match [regex]::Escape($DoneMark)) { $done = $true; break }
}
foreach ($id in $mine) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force } }
if (-not $done) { Write-Host 'DID NOT FINISH' }
Write-Host "--- result ($Tag) ---"
$t = Read-Shared $res
if ($t) { $t -split "`r?`n" | Where-Object { $_ } | ForEach-Object { "  $_" } } else { '  (empty)' }
