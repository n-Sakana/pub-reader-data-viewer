# e2e_open.ps1 -- VERIFICATION ONLY. Opens the distributable the way a person
# does: the shell association (what a double click runs). Then, if Excel shows
# the macro security bar, presses "Enable Content" through UI Automation -- the
# same gesture a person would make, not a settings change.
#
# It records which Excel processes it started, so cleanup only ever touches
# those. Nothing that was already running is used or closed.
[CmdletBinding()]
param(
    [string] $Book = 'C:\repos\pub\reader-data-viewer\showcase\dist\VBA Pixel Bridge.xlsm',
    [int] $WaitSec = 25,
    [string] $Shot = ''
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $here 'e2e-owned.pid'

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Write-Host "excel before: $($before -join ',')"

Start-Process -FilePath $Book | Out-Null
$deadline = (Get-Date).AddSeconds(40)
$mine = @()
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    $mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
              Where-Object { $before -notcontains $_ })
    if ($mine.Count -gt 0) { break }
}
if ($mine.Count -eq 0) { throw 'excel did not start' }
Add-Content -Path $pidFile -Value ($mine -join ',') -Encoding ascii
Write-Host "excel started by this script: $($mine -join ',')"

# Excel's "this document caused a serious error last time" prompt, if a probe
# run had to be killed with the book open. Answering it is what a person does.
Start-Sleep -Seconds 4
& (Join-Path $here 'e2e_dialog.ps1') -Pids $mine -ButtonPattern '^はい' -WindowTextPattern '重大なエラー'
Start-Sleep -Seconds 4

# the macro security bar, if it is there
$root = [System.Windows.Automation.AutomationElement]::RootElement
$scope = [System.Windows.Automation.TreeScope]::Descendants
$pressed = $false
foreach ($id in $mine) {
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $id)
    $win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
    if ($null -eq $win) { continue }
    $btnCond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        [System.Windows.Automation.ControlType]::Button)
    $btns = $win.FindAll($scope, $btnCond)
    foreach ($b in $btns) {
        $n = $b.Current.Name
        if ($n -match 'コンテンツの有効化|Enable Content|有効にする') {
            Write-Host "pressing the security bar button: [$n]"
            $p = $b.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
            $p.Invoke()
            $pressed = $true
            break
        }
    }
    if ($pressed) { break }
}
if (-not $pressed) { Write-Host 'no security bar found (macros were already allowed)' }

Write-Host "waiting $WaitSec s for the screen to build ..."
Start-Sleep -Seconds $WaitSec

# where the excel window is, in screen pixels
foreach ($id in $mine) {
    $cond = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty, $id)
    $win = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $cond)
    if ($null -ne $win) {
        $r = $win.Current.BoundingRectangle
        Write-Host ("excel window rect: {0},{1} {2}x{3}  name=[{4}]" -f `
            [int]$r.X, [int]$r.Y, [int]$r.Width, [int]$r.Height, $win.Current.Name)
    }
}
if ($Shot) {
    & (Join-Path $here 'e2e_shot.ps1') -Out $Shot
}
Write-Host 'done'
