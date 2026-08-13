# ============================================================================
# run_bench.ps1 -- open ZipBench.xlsm in a NORMAL, visible Excel window and
#                  drive the same buttons a person would press.
#
# This is how the benchmark is verified without a human at the keyboard. It
# does not bypass anything: it opens the workbook the ordinary way, calls the
# very macros the sheet's buttons call, and reads the result table back off
# the sheet afterwards.
#
#   .\run_bench.ps1 -Count 2000 -Methods 7
#   .\run_bench.ps1 -Count 1000000 -Methods 1,2,4,5,6,7,8
#   .\run_bench.ps1 -Count 2000 -Methods 3 -ManualBat      # opens the BAT for method 3
#
# It only ever touches the Excel instance it started itself.
# ============================================================================
[CmdletBinding()]
param(
  [int]      $Count      = 2000,
  [string]   $Methods    = '7',   # comma separated, e.g. "1,2,4,5,6,7,8"
  [string]   $WorkerKind = 'prebuilt',
  [switch]   $ManualBat,
  [switch]   $EmitOnly,
  [int]      $Repeat     = 1,
  [switch]   $KeepOpen,
  [string]   $Root       = ""
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$log = Join-Path $Root 'work\run_bench.log'
New-Item -ItemType Directory -Path (Join-Path $Root 'work') -Force | Out-Null
Set-Content -LiteralPath $log -Value '' -Encoding utf8 -Force
function Step([string] $m) {
  $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $m
  Add-Content -LiteralPath $log -Value $line -Encoding utf8
  Write-Output $line
}

$benchPath = Join-Path $Root 'ZipBench.xlsm'
if (-not (Test-Path -LiteralPath $benchPath)) { throw "not built yet: $benchPath  (run build\build_workbooks.ps1)" }

$before = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $true                 # a normal Excel window, not a hidden automation instance
$xl.DisplayAlerts = $false
$xl.AutomationSecurity = 1          # macros enabled, same as a user clicking "Enable content"
$after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$mine = @($after | Where-Object { $before -notcontains $_ })
Step ("excel pid: " + ($mine -join ','))
if ($mine.Count) { Set-Content -LiteralPath "$log.pid" -Value ($mine -join ',') }

$manualJob = $null
try {
  $wb = $xl.Workbooks.Open($benchPath, 0)
  Step "opened $($wb.Name)"
  Step ("sheets: " + ((@($wb.Worksheets) | ForEach-Object { $_.Name }) -join ', '))
  $ws = $wb.Worksheets | Where-Object { $_.Name -eq 'BENCH' } | Select-Object -First 1
  if (-not $ws) { throw 'BENCH sheet not found' }
  Step 'got BENCH'

  $ws.Range('B3').Value2 = [double]$Count
  $ws.Range('B4').Value2 = [string]$WorkerKind
  $ws.Range('B8').Value2 = [string]$(if ($ManualBat) { '非表示' } else { '表示' })
  $methodList = @($Methods -split ',' | Where-Object { $_ -match '^\d+$' } | ForEach-Object { [int]$_ })
  Step "count=$Count workerKind=$WorkerKind methods=$($methodList -join ',')"

  if ($EmitOnly) {
    Step 'ZB_EmitOnlyPrebuilt'; $xl.Run("'$($wb.Name)'!ZB_EmitOnlyPrebuilt")
    Step 'ZB_EmitOnlyEmitted';  $xl.Run("'$($wb.Name)'!ZB_EmitOnlyEmitted")
  }

  Step 'ZB_Prepare ...'
  $sw = [Diagnostics.Stopwatch]::StartNew()
  $xl.Run("'$($wb.Name)'!ZB_Prepare")
  $sw.Stop()
  Step ("ZB_Prepare done in {0:N1}s  -> {1}" -f ($sw.ElapsedMilliseconds / 1000), $ws.Range('B6').Text)
  if ($ws.Range('B6').Text -notlike '準備済*') { throw "prepare failed: $($ws.Range('B6').Text)" }

  for ($rep = 1; $rep -le $Repeat; $rep++) {
   Step "===== 実行 $rep / $Repeat ====="
   foreach ($m in $methodList) {
    # Method 3 is the manual one: a human opens the BAT. Stand in for that human
    # by watching for the file the macro writes and opening it from outside Excel,
    # which is exactly what a double-click in Explorer does.
    if ($m -eq 3 -and $ManualBat) {
      $bat = Join-Path $Root 'work\manual_run\launch.bat'
      if (Test-Path -LiteralPath $bat) { Remove-Item -LiteralPath $bat -Force }
      $manualJob = Start-Job -ScriptBlock {
        param($b)
        $deadline = (Get-Date).AddSeconds(240)
        while ((Get-Date) -lt $deadline -and -not (Test-Path -LiteralPath $b)) { Start-Sleep -Milliseconds 50 }
        if (Test-Path -LiteralPath $b) {
          Start-Sleep -Milliseconds 250      # let the macro finish writing the file
          Start-Process -FilePath $b -WindowStyle Minimized
          return "opened $b"
        }
        return "bat never appeared"
      } -ArgumentList $bat
      Step 'method 3: watcher armed (will open the BAT from outside Excel)'
    }

    Step "ZB_Run$m ..."
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $xl.Run("'$($wb.Name)'!ZB_Run$m")
    $sw.Stop()
    $row = 11 + $m
    Step ("ZB_Run$m returned in {0:N1}s" -f ($sw.ElapsedMilliseconds / 1000))
    Step ("  status bar : " + $xl.StatusBar)
    Step ("  作成={0} 起動={1} UIA={2} 変換={3} Err={4}" -f `
          $ws.Cells($row,4).Text, $ws.Cells($row,5).Text, $ws.Cells($row,6).Text, $ws.Cells($row,7).Text, $ws.Cells($row,8).Text)
    Step ("  中央値 起動={0} M={1} I={2} 辞書={3} 変換={4} 出力={5} 反映={6} 通知={7} E2E={8} 照合={9}" -f `
          $ws.Cells($row,9).Text,  $ws.Cells($row,10).Text, $ws.Cells($row,11).Text, `
          $ws.Cells($row,12).Text, $ws.Cells($row,13).Text, $ws.Cells($row,14).Text, `
          $ws.Cells($row,15).Text, $ws.Cells($row,16).Text, $ws.Cells($row,17).Text, $ws.Cells($row,18).Text)
    Step ("  回数={0} E2E最小={1} E2E最大={2}" -f $ws.Cells($row,19).Text, $ws.Cells($row,20).Text, $ws.Cells($row,21).Text)
    Step ("  行数={0} 一致={1} 不一致行={2} hash={3}" -f `
          $ws.Cells($row,22).Text, $ws.Cells($row,23).Text, $ws.Cells($row,24).Text, $ws.Cells($row,25).Text)
    Step ("  完了結果   : " + $ws.Cells($row,26).Text)
    Step ("  メモ       : " + $ws.Cells($row,27).Text)

    if ($manualJob) { Step ("  method 3 watcher: " + ((Receive-Job $manualJob -Wait) -join '; ')); Remove-Job $manualJob; $manualJob = $null }
   }
  }

  Step 'ZB_Cleanup'
  $xl.Run("'$($wb.Name)'!ZB_Cleanup")

  Step '--- LOG sheet ---'
  $wsLog = $wb.Worksheets.Item('LOG')
  $lastLog = $wsLog.Cells($wsLog.Rows.Count, 1).End(-4162).Row
  for ($r = 2; $r -le $lastLog; $r++) {
    Step ("  " + $wsLog.Cells($r,1).Text + "  " + $wsLog.Cells($r,2).Text)
  }
}
catch {
  Step ("FAILED: " + $_.Exception.Message)
  throw
}
finally {
  if ($manualJob) { Remove-Job $manualJob -Force -ErrorAction SilentlyContinue }
  if (-not $KeepOpen) {
    try { $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}
Step 'done'
