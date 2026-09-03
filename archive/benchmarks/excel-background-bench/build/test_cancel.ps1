# ============================================================================
# test_cancel.ps1 -- prove the safe cancel actually stops a run.
#
# There are two cancel paths and they are cancelled for different reasons:
#
#   * a VBA conversion loop (methods 1 and 2) blocks Excel, so the only thing
#     the user can press is Esc. Application.EnableCancelKey = xlErrorHandler
#     turns that into a trappable error 18, which the method records as a
#     cancelled run instead of a crash.
#
#   * a worker or hidden-Excel run yields between polls / chunks, so the 取消
#     button on the sheet works there; it sets a flag the loops check.
#
# This script tests the Esc path, because it is the one that has to work while
# Excel is busy. It starts a 1,000,000 row run of method 1 in this process,
# and a second process presses Esc a few seconds in.
#
#   .\test_cancel.ps1
# ============================================================================
[CmdletBinding()]
param([string] $Root = "", [int] $Count = 1000000, [int] $AfterSeconds = 5)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$log = Join-Path $Root 'work\test_cancel.log'
Set-Content -LiteralPath $log -Value '' -Encoding utf8 -Force
function Step([string] $m) {
  $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $m
  Add-Content -LiteralPath $log -Value $line -Encoding utf8
  Write-Output $line
}

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$mine = @($after | Where-Object { $before -notcontains $_ })
Step ("excel pid: " + ($mine -join ','))

try {
  $wb = $xl.Workbooks.Open((Join-Path $Root 'ZipBench.xlsm'), 0)
  $ws = $wb.Worksheets | Where-Object { $_.Name -eq 'BENCH' } | Select-Object -First 1
  $ws.Range('B3').Value2 = [double]$Count
  Step 'ZB_Prepare ...'
  $xl.Run("'$($wb.Name)'!ZB_Prepare")
  Step ("prepared: " + $ws.Range('B6').Text)

  # second process: wait, focus this Excel, press Esc
  $presser = Start-Job -ScriptBlock {
    param($pid2, $delay)
    Start-Sleep -Seconds $delay
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type @"
using System;using System.Runtime.InteropServices;
public class FG {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int n);
}
"@
    $p = Get-Process -Id $pid2 -ErrorAction SilentlyContinue
    if (-not $p) { return 'excel gone' }
    [void][FG]::ShowWindow($p.MainWindowHandle, 9)     # SW_RESTORE
    [void][FG]::SetForegroundWindow($p.MainWindowHandle)
    Start-Sleep -Milliseconds 400
    [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
    return ("pressed Esc at " + (Get-Date -Format 'HH:mm:ss.fff'))
  } -ArgumentList ([int]$mine[0]), $AfterSeconds

  Step "ZB_Run1 (Esc will be pressed after ${AfterSeconds}s) ..."
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $xl.Run("'$($wb.Name)'!ZB_Run1")
  $sw.Stop()
  Step ("ZB_Run1 returned in {0:N1}s" -f ($sw.ElapsedMilliseconds / 1000))
  Step ("  presser: " + ((Receive-Job $presser -Wait) -join '; '))
  Remove-Job $presser

  $row = 12
  Step ("  status bar : " + $xl.StatusBar)
  Step ("  変換={0} Err={1} E2E={2}s 行数={3} 一致={4}" -f `
        $ws.Cells($row,7).Text, $ws.Cells($row,8).Text, $ws.Cells($row,12).Text, `
        $ws.Cells($row,14).Text, $ws.Cells($row,15).Text)
  Step ("  完了結果   : " + $ws.Cells($row,18).Text)
  Step ("  メモ       : " + $ws.Cells($row,19).Text)

  # after a cancel, Excel must be left usable: settings restored, no stuck state
  Step ("  after cancel -> ScreenUpdating={0} EnableEvents={1} Calculation={2} Interactive={3}" -f `
        $xl.ScreenUpdating, $xl.EnableEvents, $xl.Calculation, $xl.Interactive)

  $xl.Run("'$($wb.Name)'!ZB_Cleanup")
  Step 'ZB_Cleanup done'
}
catch { Step ("FAILED: " + $_.Exception.Message); throw }
finally {
  try { $wb.Close($false) } catch { }
  try { $xl.Quit() } catch { }
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
Step 'done'
