# ============================================================================
# bench_app.ps1 (v2) -- measure the two practical builds over the real
# product path, END TO END, with every run recorded (cold and failed included).
#
#   powershell -File build\bench_app.ps1 -Method vba    -Launches 7
#   powershell -File build\bench_app.ps1 -Method csharp -Launches 7
#
# The bench copies dist\app-* into a scratch folder under work\ and runs
# there: the shipped dist stays byte-identical. Per launch it records
#
#   full_e2e_ms      process start (CreateObject Excel / Start-Process cmd)
#                    -> the app's own "ready" log line, seen by this script.
#                    Includes Excel/console/csc startup and <=200 ms log-poll
#                    latency. The VBA launch goes through COM (AutomationSecurity
#                    for the unsigned macro); a plain double-click additionally
#                    needs a trusted location, which this bench does not touch.
#   boot_to_ready_ms the app's own startup line (QPC, boot -> operable)
#   merge stages     read/index/join sum (the on-screen merge figure),
#                    compose separately, and their total
#   load_ms (+src)   the saved-state load (sidecar / book / xlsx)
#   per key          the app's detect latency AND the search figure
#                    (detect t0 -> render complete), candidate counts
#   processed        a real mouse/button click + the real confirm dialog:
#                    persist figure and confirm-to-persisted e2e from the log
#   apply cycle      (once, after the plain launches) a CSV edit, the real
#                    confirm dialog approved, approve-click -> ready wall
#
# Every run lands in the TSV -- failures as status=FAIL rows with a reason --
# and the summary prints min/median/max/n over ALL runs (cold included), with
# launch 1 also shown alone as the cold run.
#
# Windows this script starts, it closes (WM_CLOSE to the window; never a
# process kill on Notepad -- Win11 Notepad is one shared process). It types
# ONLY into the Notepad window it opened itself. If the app binds to some
# other window the run fails loudly rather than touching that window.
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('vba', 'csharp')] [string] $Method = 'csharp',
  [int] $Launches = 7,
  [string[]] $Keys = @(),
  [string] $Root = "",
  [int] $TypeDelayMs = 15,
  [int] $ReadyTimeoutSec = 180,
  [int] $SearchTimeoutSec = 30
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class RdvBenchWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr root, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);

  public class Kid { public long Hwnd; public string Cls; public string Text; }
  // the child enumeration the scenario suites use successfully
  public static List<Kid> Kids(IntPtr root) {
    List<Kid> outp = new List<Kid>();
    EnumChildWindows(root, delegate(IntPtr h, IntPtr l) {
      StringBuilder cb = new StringBuilder(256); GetClassNameW(h, cb, 256);
      StringBuilder tb = new StringBuilder(512); GetWindowTextW(h, tb, 512);
      Kid k = new Kid(); k.Hwnd = h.ToInt64(); k.Cls = cb.ToString(); k.Text = tb.ToString();
      outp.Add(k);
      return true;
    }, IntPtr.Zero);
    return outp;
  }
}
"@
[void][RdvBenchWin]::SetProcessDPIAware()   # Excel hands back physical pixels
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$VP = [System.Windows.Automation.ValuePattern]
$Clock = [System.Diagnostics.Stopwatch]::StartNew()

# an Excel that got force-killed leaves this path blocklisted; the entry is
# (re)written when the NEXT Excel starts, so it is cleared after creation.
# This is a HARNESS-ONLY concern: the product never kills its own Excel.
function Clear-OurDisabledItems {
  foreach ($v in @('16.0', '15.0')) {
    $key = "HKCU:\Software\Microsoft\Office\$v\Excel\Resiliency\DisabledItems"
    if (-not (Test-Path $key)) { continue }
    $props = Get-ItemProperty $key
    foreach ($pp in $props.PSObject.Properties) {
      if ($pp.Name -notlike 'PS*' -and $pp.Value -is [byte[]]) {
        $ss = [Text.Encoding]::Unicode.GetString($pp.Value)
        if ($ss -like '*reader-data-viewer*' -or $ss -like '*bench-run*' -or $ss -like '*rdv3*') {
          Remove-ItemProperty -Path $key -Name $pp.Name -ErrorAction SilentlyContinue
        }
      }
    }
  }
}

if ($Keys.Count -eq 0) {
  $Keys = @('00000001', '00001001', '00006001', '00021001')
} else {
  $Keys = @($Keys -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 0 })
}
$SingleKey = '00021001'    # the 1-candidate key: its record hosts the processed click

$isVba = ($Method -eq 'vba')
$distRoot = Join-Path $Root $(if ($isVba) { 'dist\app-vba' } else { 'dist\app-csharp' })
if ($isVba) {
  if (-not (Test-Path -LiteralPath (Join-Path $distRoot 'ReaderDataViewer.xlsm'))) { throw "not built: run build_app.ps1" }
} else {
  if (-not (Test-Path -LiteralPath (Join-Path $distRoot 'ReaderDataViewer.cmd'))) { throw "not built: run build_app.ps1" }
}

# --- scratch copy: the dist stays untouched ---------------------------------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$benchRoot = Join-Path $Root ("work\bench-run\{0}-{1}" -f $Method, $stamp)
New-Item -ItemType Directory -Path $benchRoot -Force | Out-Null
Copy-Item -Path (Join-Path $distRoot '*') -Destination $benchRoot -Recurse -Force
$appLog = Join-Path $benchRoot 'ReaderDataViewer.log'
$beLog = $appLog + '.worker.log'            # the VBA BE logs its detect line here
$logEnc = if ($isVba) { [Text.Encoding]::GetEncoding(932) } else { (New-Object Text.UTF8Encoding($false)) }
if (Test-Path -LiteralPath $appLog) { Remove-Item -LiteralPath $appLog -Force }
if (Test-Path -LiteralPath $beLog) { Remove-Item -LiteralPath $beLog -Force }
Write-Output ("bench scratch: " + $benchRoot)

# --- refuse to start on top of a previous run -------------------------------
$staleBooks = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*ReaderDataViewer*' })
if ($staleBooks.Count -gt 0) {
  throw ("a Reader Data Viewer workbook is still open in Excel (pid " + (($staleBooks | ForEach-Object { $_.Id }) -join ', ') + "). Close it before measuring.")
}
$stalePs = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object { $_.CommandLine -like '*RDV_SELF*' })
if ($stalePs.Count -gt 0) {
  throw ("a Reader Data Viewer .cmd app is still running (pid " + (($stalePs | ForEach-Object { $_.ProcessId }) -join ', ') + "). Close it before measuring.")
}

# --- my own Notepad -----------------------------------------------------------
# A PLAIN notepad.exe launch opens a NEW top-level window even when other
# notepad windows exist (a FILE argument would be absorbed as a tab of an
# existing window instead -- measured). Pre-existing windows are located by
# handle first and never touched; the new window is identified by handle
# difference, and the app-side bound-hwnd check below aborts the run rather
# than let anything type into a window that is not ours. The tab stays
# untitled: it is cleared before WM_CLOSE (a dirty tab ignores WM_CLOSE).
function Find-NotepadWindows {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'Notepad')
  return $AE::RootElement.FindAll($TS::Children, $c)
}
$beforeNp = @()
foreach ($w in (Find-NotepadWindows)) { $beforeNp += [long]$w.Current.NativeWindowHandle }
Start-Process -FilePath "$env:WINDIR\System32\notepad.exe" | Out-Null
$myNp = $null
for ($i = 0; $i -lt 60; $i++) {
  Start-Sleep -Milliseconds 250
  foreach ($w in (Find-NotepadWindows)) {
    if ($beforeNp -notcontains [long]$w.Current.NativeWindowHandle) { $myNp = $w; break }
  }
  if ($null -ne $myNp) { break }
}
if ($null -eq $myNp) { throw 'my Notepad window did not appear (a file-less launch should always open one)' }
$myNpHwnd = [IntPtr]$myNp.Current.NativeWindowHandle
[void][RdvBenchWin]::SetForegroundWindow($myNpHwnd)
Write-Output ("my notepad hwnd " + $myNpHwnd + " (existing notepad windows untouched: " + $beforeNp.Count + ")")

$edit = $myNp.FindFirst($TS::Descendants,
  (New-Object System.Windows.Automation.AndCondition(
    (New-Object System.Windows.Automation.PropertyCondition($AE::IsValuePatternAvailableProperty, $true)),
    (New-Object System.Windows.Automation.OrCondition(
      (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Document)),
      (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)))))))
if ($null -eq $edit) { throw 'my notepad text host not found' }
$editHwnd = [IntPtr]$edit.Current.NativeWindowHandle
$vp = $edit.GetCurrentPattern($VP::Pattern)

function Type-Key([string] $k) {
  $vp.SetValue('')          # my own notepad only
  Start-Sleep -Milliseconds 400
  foreach ($ch in $k.ToCharArray()) {
    [void][RdvBenchWin]::PostMessage($editHwnd, 0x0102, [IntPtr][int]$ch, [IntPtr]0)
    Start-Sleep -Milliseconds $TypeDelayMs
  }
}

# --- input helpers (real mouse; a synthetic WM_LBUTTON pair does not fire
# --- WinForms MouseClick or Excel FollowHyperlink) --------------------------
function Click-Screen([int] $x, [int] $y) {
  [void][RdvBenchWin]::SetCursorPos($x, $y)
  Start-Sleep -Milliseconds 120
  [RdvBenchWin]::mouse_event(2, 0, 0, 0, [UIntPtr]::Zero)   # LEFTDOWN
  Start-Sleep -Milliseconds 60
  [RdvBenchWin]::mouse_event(4, 0, 0, 0, [UIntPtr]::Zero)   # LEFTUP
}

# waits for a native dialog with the given title and answers it. Preferred:
# BM_CLICK POSTED to the button found by caption (a SendMessage would deadlock
# against the modal loop). Fallback when child enumeration yields nothing:
# WM_COMMAND with the standard dialog ID (はい=IDYES 6 / いいえ=IDNO 7 / OK=1)
# posted to the dialog itself -- MessageBox answers those unconditionally.
function Answer-Dialog([string] $title, [string] $buttonPrefix, [int] $sec) {
  $cmdId = 1
  if ($buttonPrefix -like 'はい*') { $cmdId = 6 }
  elseif ($buttonPrefix -like 'いいえ*') { $cmdId = 7 }
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $dlg = [RdvBenchWin]::FindWindowW('#32770', $title)
    if ($dlg -ne [IntPtr]::Zero) {
      Start-Sleep -Milliseconds 250
      $btn = [IntPtr]::Zero
      foreach ($k in [RdvBenchWin]::Kids($dlg)) {
        if ($k.Cls -eq 'Button' -and ($k.Text.Replace('&', '') -like ($buttonPrefix + '*'))) {
          $btn = [IntPtr]$k.Hwnd
          break
        }
      }
      if ($btn -ne [IntPtr]::Zero) {
        [void][RdvBenchWin]::PostMessage($btn, 0x00F5, [IntPtr]0, [IntPtr]0)  # BM_CLICK
        # Write-Host, not Write-Output: this function RETURNS $true/$false and
        # anything written to the output stream joins that return value, so the
        # caller's "if (-not (Answer-Dialog ...))" would be testing an array
        Write-Host ("  dialog '" + $title + "' answered (button click)")
      } else {
        [void][RdvBenchWin]::PostMessage($dlg, 0x0111, [IntPtr]$cmdId, [IntPtr]0)  # WM_COMMAND
        Write-Host ("  dialog '" + $title + "' answered (WM_COMMAND id=" + $cmdId + ")")
      }
      return $true
    }
    Start-Sleep -Milliseconds 200
  }
  return $false
}

# --- log helpers ------------------------------------------------------------
function Log-Lines {
  if (-not (Test-Path -LiteralPath $appLog)) { return @() }
  # the app appends with a brief exclusive lock; a colliding read retries
  for ($ri = 0; $ri -lt 20; $ri++) {
    try { return @([IO.File]::ReadAllLines($appLog, $logEnc)) } catch { Start-Sleep -Milliseconds 150 }
  }
  throw "app log stayed locked: $appLog"
}
function Wait-LogLine([string] $pattern, [int] $fromCount, [int] $sec) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $all = Log-Lines
    for ($i = $fromCount; $i -lt $all.Count; $i++) {
      if ($all[$i] -match $pattern) { return , @($i, $all[$i]) }
    }
    Start-Sleep -Milliseconds 200
  }
  return $null
}
function Extract([string] $line, [string] $re) {
  if ($null -ne $line -and $line -match $re) { return $Matches[1] }
  return ''
}
function BeLog-Lines {
  if (-not (Test-Path -LiteralPath $beLog)) { return @() }
  for ($ri = 0; $ri -lt 10; $ri++) {
    try { return @([IO.File]::ReadAllLines($beLog, $logEnc)) } catch { Start-Sleep -Milliseconds 150 }
  }
  return @()
}

# --- result recording: EVERY run lands in the TSV, failures included --------
$outTsv = Join-Path $Root ("work\bench-app-{0}-{1}.tsv" -f $Method, $stamp)
$workDir = Split-Path -Parent $outTsv
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }
Set-Content -LiteralPath $outTsv -Value ("launch`tphase`tstatus`tmetric`tvalue_ms`textra") -Encoding utf8
$rows = New-Object System.Collections.ArrayList
function Rec([int] $launch, [string] $phase, [string] $status, [string] $metric, [string] $value, [string] $extra) {
  Add-Content -LiteralPath $outTsv -Value ("{0}`t{1}`t{2}`t{3}`t{4}`t{5}" -f $launch, $phase, $status, $metric, $value, $extra) -Encoding utf8
  [void]$rows.Add([pscustomobject]@{ Launch = $launch; Phase = $phase; Status = $status; Metric = $metric; Value = $value; Extra = $extra })
}


# --- one launch: start, measure, operate, close -----------------------------
function Close-App([object] $ctx) {
  if ($isVba) {
    # the app defers a close while a pump tick is armed (it closes itself
    # right after the tick); wait, then Quit. A kill would poison the path
    # via Excel's DisabledItems, so it stays the last resort.
    if ($null -ne $ctx.Wb) { try { $ctx.Wb.Close($false) } catch { } }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.Wb) } catch { }
    $ctx.Wb = $null
    if ($null -ne $ctx.Xl) {
      for ($ri = 0; $ri -lt 20; $ri++) {
        $cnt = 1
        try { $cnt = $ctx.Xl.Workbooks.Count } catch { $cnt = -1 }
        if ($cnt -eq 0) { break }
        Start-Sleep -Milliseconds 600
      }
      for ($ri = 0; $ri -lt 12; $ri++) {
        try { $ctx.Xl.Quit(); break } catch { Start-Sleep -Milliseconds 800 }
      }
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ctx.Xl) } catch { }
      $ctx.Xl = $null
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $gone = $false
    for ($ri = 0; $ri -lt 12; $ri++) {
      if (@($ctx.MyExcel | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue }).Count -eq 0) { $gone = $true; break }
      Start-Sleep -Milliseconds 700
    }
    if (-not $gone -and $ctx.XlHwnd -ne [IntPtr]::Zero) {
      # a REAL mouse click into the window leaves Excel ignoring COM Quit
      # (measured); a user-style WM_CLOSE on the main window ends it cleanly
      [void][RdvBenchWin]::PostMessage($ctx.XlHwnd, 0x0010, [IntPtr]0, [IntPtr]0)
      for ($ri = 0; $ri -lt 15; $ri++) {
        if (@($ctx.MyExcel | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue }).Count -eq 0) { $gone = $true; break }
        Start-Sleep -Milliseconds 700
      }
    }
    foreach ($p in $ctx.MyExcel) {
      $q = Get-Process -Id $p -ErrorAction SilentlyContinue
      if ($q) { Write-Output "  killing my excel $p (last resort)"; $q.Kill() }
    }
  } else {
    $psHost = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object { $_.CommandLine -like '*RDV_SELF*' })
    foreach ($ph in $psHost) {
      $h = [IntPtr]::Zero
      $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
      $w = $walker.GetFirstChild($AE::RootElement)
      while ($null -ne $w) {
        if ($w.Current.ProcessId -eq $ph.ProcessId -and $w.Current.Name -like 'Reader Data Viewer*') { $h = [IntPtr]$w.Current.NativeWindowHandle; break }
        $w = $walker.GetNextSibling($w)
      }
      if ($h -ne [IntPtr]::Zero) { [void][RdvBenchWin]::PostMessage($h, 0x0010, [IntPtr]0, [IntPtr]0) }
    }
    Start-Sleep -Seconds 3
    foreach ($ph in $psHost) {
      $q = Get-Process -Id $ph.ProcessId -ErrorAction SilentlyContinue
      if ($q) { Write-Output ("  killing my app powershell " + $ph.ProcessId); $q.Kill() }
    }
    if ($null -ne $ctx.Proc -and -not $ctx.Proc.HasExited) {
      Stop-Process -Id $ctx.Proc.Id -Force -ErrorAction SilentlyContinue
    }
  }
  Start-Sleep -Milliseconds 800
}

# the whole launch body runs under try/finally: whatever throws, THIS ctx is
# closed (the failed-run cleanup can never orphan an app host again)
function Run-Launch([int] $launch, [string] $phase, [bool] $doApply) {
  $ctx = [pscustomobject]@{ Xl = $null; Wb = $null; Proc = $null; MyExcel = @(); XlHwnd = [IntPtr]::Zero }
  try {
    Run-LaunchBody $launch $phase $doApply $ctx
  } finally {
    # a dangling modal would block the close; answer either confirm with NO
    [void](Answer-Dialog '更新の確認' 'いいえ' 1)
    [void](Answer-Dialog '処理済みの確認' 'いいえ' 1)
    Close-App $ctx
  }
}

function Run-LaunchBody([int] $launch, [string] $phase, [bool] $doApply, [object] $ctx) {
  $lineBase = (Log-Lines).Count
  [void][RdvBenchWin]::SetForegroundWindow($myNpHwnd)

  # ---- start (the full-E2E clock runs across the whole process start) ----
  $tStart = $Clock.Elapsed.TotalMilliseconds
  if ($isVba) {
    # identity is settled before anything is done to it (excel_own.ps1):
    # a reused or unidentifiable instance throws here and is never driven
    $rdvOwn = New-OwnedExcel
    $ctx.Xl = $rdvOwn.App
    $ctx.MyExcel = @($rdvOwn.Pid)
    $ctx.Xl.Visible = $false                # invisible: my notepad stays foreground
    $ctx.Xl.DisplayAlerts = $false
    Clear-OurDisabledItems
    # a just-started Excel can be briefly busy (0x800AC472); retry the property
    for ($ri = 0; $ri -lt 10; $ri++) {
      try { $ctx.Xl.AutomationSecurity = 1; break } catch { Start-Sleep -Milliseconds 800 }
    }
    for ($ri = 0; $ri -lt 10; $ri++) {
      try { $ctx.Wb = $ctx.Xl.Workbooks.Open((Join-Path $benchRoot 'ReaderDataViewer.xlsm')); break }
      catch { Clear-OurDisabledItems; Start-Sleep -Milliseconds 1000 }
    }
    if ($null -eq $ctx.Wb) { throw "workbook open failed" }
    $ctx.XlHwnd = [IntPtr]$ctx.Xl.Hwnd
  } else {
    $ctx.Proc = Start-Process -FilePath (Join-Path $benchRoot 'ReaderDataViewer.cmd') -PassThru -WindowStyle Minimized
  }

  # ---- apply path: the real confirm dialog, approved -----------------------
  $applyClickMs = -1
  if ($doApply) {
    if (-not (Answer-Dialog '更新の確認' 'はい' 150)) { throw 'update-confirm dialog never appeared' }
    $applyClickMs = $Clock.Elapsed.TotalMilliseconds
    Write-Output '  update-confirm dialog approved'
  }

  $ready = Wait-LogLine "`tdecision`tready " $lineBase $ReadyTimeoutSec
  if ($null -eq $ready) { throw "app never reached READY (see $appLog)" }
  $fullE2e = $Clock.Elapsed.TotalMilliseconds - $tStart
  Rec $launch $phase 'ok' 'full_e2e' ('{0:0.0}' -f $fullE2e) $(if ($launch -eq 1 -and $phase -eq 'nodiff') { 'cold' } else { '' })
  if ($doApply -and $applyClickMs -ge 0) {
    Rec $launch $phase 'ok' 'approve_to_ready' ('{0:0.0}' -f ($Clock.Elapsed.TotalMilliseconds - $applyClickMs)) 'bench clock, incl <=200ms log poll'
  }

  $all = Log-Lines
  $seg = $all[$lineBase..($all.Count - 1)]
  $mergeLine = ($seg | Where-Object { $_ -match "`tmerge`t" } | Select-Object -Last 1)
  $mergeMs = Extract $mergeLine ' ms=([0-9.]+)'
  $composeMs = Extract $mergeLine 'compose_ms=([0-9.]+)'
  $checksum = Extract $mergeLine 'checksum=(\d+)'
  if ($checksum -eq '') {
    $checksum = Extract (($seg | Where-Object { $_ -match "`tjoin`t.*checksum=" } | Select-Object -Last 1)) 'checksum=(\d+)'
  }
  $loadLine = ($seg | Where-Object { $_ -match "`tload`t" } | Select-Object -Last 1)
  $loadMs = Extract $loadLine ' ms=([0-9.]+)'
  $loadSrc = Extract $loadLine ' src=([a-z]+)'
  $startup = Extract (($seg | Where-Object { $_ -match "`tstartup`t" } | Select-Object -Last 1)) 'boot_to_ready_ms=([0-9.]+)'
  $mergeTotal = ''
  if ($mergeMs -ne '' -and $composeMs -ne '') { $mergeTotal = '{0:0.0}' -f ([double]$mergeMs + [double]$composeMs) }
  Rec $launch $phase 'ok' 'boot_to_ready' $startup ''
  Rec $launch $phase 'ok' 'merge' $mergeMs ("checksum=" + $checksum)
  Rec $launch $phase 'ok' 'compose' $composeMs ''
  Rec $launch $phase 'ok' 'merge_total' $mergeTotal 'merge+compose'
  Rec $launch $phase 'ok' 'load' $loadMs ("src=" + $loadSrc)
  Write-Output ("  ready; full_e2e {0:0.0} ms  boot_to_ready {1} ms  merge {2} ms (+compose {3})  load {4} ms ({5})  checksum {6}" -f `
    $fullE2e, $startup, $mergeMs, $composeMs, $loadMs, $loadSrc, $checksum)

  # ---- the VBA build logs the bound hwnd; make sure it watches MY notepad --
  if ($isVba) {
    $bound = Wait-LogLine "`twatch`tbound hwnd=" $lineBase 30
    if ($null -eq $bound) { throw "watch never bound" }
    if ($bound[1] -match 'hwnd=(\d+)') {
      if ([long]$Matches[1] -ne [long]$myNpHwnd) {
        throw ("the app bound notepad hwnd " + $Matches[1] + ", not mine (" + [long]$myNpHwnd + "). Not typing into someone else's window.")
      }
    }
  }

  # ---- searches through the real notepad detect path -----------------------
  $from = (Log-Lines).Count
  foreach ($k in $Keys) {
    $tType = $Clock.Elapsed.TotalMilliseconds
    Type-Key $k
    $hit = Wait-LogLine ("`tsearch`tkey=" + $k + " source=detect ") $from $SearchTimeoutSec
    if ($null -eq $hit) {
      Write-Output ("  key {0}: TIMEOUT (no detect-search in log)" -f $k)
      Rec $launch $phase 'FAIL' ('search_' + $k) '' 'no detect-search in log'
      continue
    }
    $line = $hit[1]
    $sMs = Extract $line ' ms=([0-9.]+)'
    $cands = Extract $line ' hits=(\d+)'
    # detect latency: the C# app logs it in the main log (tab-separated
    # section), the VBA BE in its own log (space-separated)
    $detSrc = if ($isVba) { BeLog-Lines } else { Log-Lines }
    $detLine = ($detSrc | Where-Object { $_ -match ("detect[`t ]key=" + $k) } | Select-Object -Last 1)
    $detMs = Extract $detLine 'latency_ms=([0-9.]+)'
    $wall = $Clock.Elapsed.TotalMilliseconds - $tType
    Rec $launch $phase 'ok' ('search_' + $k) $sMs ("cands=" + $cands)
    if ($detMs -ne '') {
      Rec $launch $phase 'ok' ('detect_' + $k) $detMs 'typing settle -> confirmed'
    } else {
      Rec $launch $phase 'FAIL' ('detect_' + $k) '' 'latency line not found'
    }
    Rec $launch $phase 'ok' ('typewall_' + $k) ('{0:0.0}' -f $wall) 'bench clock: first keystroke -> search line seen'
    Write-Output ("  key {0}: cands {1}  search {2} ms  detect {3} ms  type-wall {4:0.0} ms" -f $k, $cands, $sMs, $detMs, $wall)
    $from = $hit[0] + 1
  }

  # ---- processed: a real click on the real UI + the real dialog ------------
  # the last key ($SingleKey) left its single record on screen
  $from = (Log-Lines).Count
  $clicked = $false
  if ($isVba) {
    $ctx.Xl.Visible = $true
    $ws = $ctx.Wb.Worksheets.Item('UI')
    $ws.Activate()
    [void][RdvBenchWin]::SetForegroundWindow($ctx.XlHwnd)
    Start-Sleep -Milliseconds 500
    # PointsToScreenPixels supplies the pane ORIGIN; the points->pixels scale
    # is the window DPI / 72 (measured: P2SP's own scaling is a 1:1 lie here)
    $win = $ctx.Xl.ActiveWindow
    $x0 = $win.PointsToScreenPixelsX(0)
    $y0 = $win.PointsToScreenPixelsY(0)
    $dpi = [RdvBenchWin]::GetDpiForWindow($ctx.XlHwnd)
    if ($dpi -eq 0) { $dpi = 96 }
    $s = $dpi / 72.0
    $r = $ws.Range('G9')
    $x = [int]($x0 + ($r.Left + $r.Width / 2) * $s)
    $y = [int]($y0 + ($r.Top + $r.Height / 2) * $s)
    Click-Screen $x $y
    $clicked = $true
  } else {
    $main = [IntPtr]::Zero
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $w = $walker.GetFirstChild($AE::RootElement)
    while ($null -ne $w) {
      if ($w.Current.Name -like 'Reader Data Viewer*' -and $w.Current.ProcessId -ne $PID) { $main = [IntPtr]$w.Current.NativeWindowHandle; break }
      $w = $walker.GetNextSibling($w)
    }
    if ($main -ne [IntPtr]::Zero) {
      # the WinForms button, by caption via child enumeration (the same path
      # the scenario suite uses successfully), then a posted BM_CLICK
      foreach ($k in [RdvBenchWin]::Kids($main)) {
        if ($k.Text -eq '処理済み') {
          [void][RdvBenchWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0)
          $clicked = $true
          break
        }
      }
    }
  }
  if (-not $clicked) {
    Rec $launch $phase 'FAIL' 'processed' '' 'processed button not reachable'
  } else {
    if (-not (Answer-Dialog '処理済みの確認' 'はい' 20)) {
      Rec $launch $phase 'FAIL' 'processed' '' 'confirm dialog never appeared'
    } else {
      $procHit = Wait-LogLine "`tprocessed`tkey2=.*value=TRUE" $from 90
      if ($null -eq $procHit) {
        Rec $launch $phase 'FAIL' 'processed' '' 'no processed line in log'
      } else {
        $pl = $procHit[1]
        $persist = Extract $pl '(?:persist_ms|save_ms)=([0-9.]+)'
        $e2e = Extract $pl 'e2e_ms=([0-9.]+)'
        Rec $launch $phase 'ok' 'processed_persist' $persist ''
        Rec $launch $phase 'ok' 'processed_e2e' $e2e 'confirm -> persisted (+screen)'
        Write-Output ("  processed: persist {0} ms  e2e {1} ms" -f $persist, $e2e)
      }
    }
    if ($isVba) { $ctx.Xl.Visible = $false }
  }
  [void][RdvBenchWin]::SetForegroundWindow($myNpHwnd)
}

# --- phase 1: plain launches (cold = launch 1, no warm-up excluded) ---------
for ($launch = 1; $launch -le $Launches; $launch++) {
  Write-Output ("=== launch {0}/{1} ({2})" -f $launch, $Launches, $Method)
  try {
    Run-Launch $launch 'nodiff' $false
  } catch {
    Write-Output ("  LAUNCH FAILED: " + $_.Exception.Message)
    Rec $launch 'nodiff' 'FAIL' 'launch' '' $_.Exception.Message.Replace("`t", ' ').Replace("`n", ' ')
    Clear-OurDisabledItems   # Run-Launch already closed its own ctx (finally)
  }
}

# --- phase 2: one apply cycle (CSV edit -> approve -> operable) -------------
Write-Output ("=== apply cycle ({0})" -f $Method)
try {
  $fA = Join-Path $benchRoot 'data\tableA.csv'
  $enc = New-Object Text.UTF8Encoding($false)
  $linesA = [IO.File]::ReadAllLines($fA, $enc)
  for ($i = 0; $i -lt $linesA.Count; $i++) {
    if ($linesA[$i] -like ($SingleKey + ',*')) { $linesA[$i] = ($linesA[$i] -replace 'NOTE-[A-Z0-9]+$', 'NOTE-BENCH2'); break }
  }
  [IO.File]::WriteAllLines($fA, $linesA, $enc)
  Run-Launch ($Launches + 1) 'apply' $true
} catch {
  Write-Output ("  APPLY FAILED: " + $_.Exception.Message)
  Rec ($Launches + 1) 'apply' 'FAIL' 'launch' '' $_.Exception.Message.Replace("`t", ' ').Replace("`n", ' ')
}

# --- close my notepad (WM_CLOSE to MY window; never a process kill). A dirty
# --- untitled tab silently IGNORES WM_CLOSE (measured), so clear it first.
if ([RdvBenchWin]::IsWindow($myNpHwnd)) {
  try { $vp.SetValue('') } catch { }
  Start-Sleep -Milliseconds 800
  [void][RdvBenchWin]::PostMessage($myNpHwnd, 0x0010, [IntPtr]0, [IntPtr]0)
  for ($ri = 0; $ri -lt 10; $ri++) {
    Start-Sleep -Milliseconds 800
    if (-not [RdvBenchWin]::IsWindow($myNpHwnd)) { break }
    if ($ri -eq 4) { [void][RdvBenchWin]::PostMessage($myNpHwnd, 0x0010, [IntPtr]0, [IntPtr]0) }
  }
  Write-Output ("my notepad closed: " + (-not [RdvBenchWin]::IsWindow($myNpHwnd)))
}

# --- summary: min / median / max / n over ALL runs, cold included -----------
function Med([double[]] $v) {
  $s = @($v | Sort-Object)
  $n = $s.Count
  if ($n -eq 0) { return 0 }
  if ($n % 2 -eq 1) { return $s[[int](($n - 1) / 2)] }
  return ($s[$n / 2 - 1] + $s[$n / 2]) / 2
}
function Line([string] $name, [object[]] $set) {
  $v = @($set | ForEach-Object { [double]$_.Value })
  if ($v.Count -eq 0) { Write-Output ("{0,-22} (no data)" -f $name); return }
  Write-Output ("{0,-22} min {1,10:N1}   median {2,10:N1}   max {3,10:N1} ms   (n={4})" -f $name,
    (($v | Measure-Object -Minimum).Minimum), (Med $v), (($v | Measure-Object -Maximum).Maximum), $v.Count)
}

Write-Output ""
Write-Output ("=== summary ({0}) -> {1}" -f $Method, $outTsv)
$ok = @($rows | Where-Object { $_.Status -eq 'ok' })
$fails = @($rows | Where-Object { $_.Status -eq 'FAIL' })
$chks = @($ok | Where-Object { $_.Metric -eq 'merge' } | ForEach-Object { ($_.Extra -replace 'checksum=', '') } | Sort-Object -Unique)
Write-Output ("checksums seen: " + (($chks | Where-Object { $_ -ne '' }) -join ', '))
Write-Output ("failed runs recorded: " + $fails.Count + $(if ($fails.Count) { '  (see FAIL rows in the TSV)' } else { '' }))
$nod = @($ok | Where-Object { $_.Phase -eq 'nodiff' })
Line 'full_e2e'        @($nod | Where-Object { $_.Metric -eq 'full_e2e' })
$cold = @($nod | Where-Object { $_.Metric -eq 'full_e2e' -and $_.Extra -eq 'cold' })
if ($cold.Count) { Write-Output ("{0,-22} {1,10:N1} ms   (launch 1, cold, included above)" -f 'full_e2e cold', [double]$cold[0].Value) }
Line 'boot_to_ready'   @($nod | Where-Object { $_.Metric -eq 'boot_to_ready' })
Line 'merge'           @($nod | Where-Object { $_.Metric -eq 'merge' })
Line 'compose'         @($nod | Where-Object { $_.Metric -eq 'compose' })
Line 'merge_total'     @($nod | Where-Object { $_.Metric -eq 'merge_total' })
Line 'load'            @($nod | Where-Object { $_.Metric -eq 'load' })
foreach ($k in $Keys) {
  Line ('search ' + $k)  @($nod | Where-Object { $_.Metric -eq ('search_' + $k) })
}
Line 'search all keys' @($nod | Where-Object { $_.Metric -like 'search_*' })
Line 'detect all keys' @($nod | Where-Object { $_.Metric -like 'detect_*' })
Line 'processed_persist' @($nod | Where-Object { $_.Metric -eq 'processed_persist' })
Line 'processed_e2e'   @($nod | Where-Object { $_.Metric -eq 'processed_e2e' })
$ap = @($ok | Where-Object { $_.Phase -eq 'apply' })
Line 'apply full_e2e'  @($ap | Where-Object { $_.Metric -eq 'full_e2e' })
Line 'approve_to_ready' @($ap | Where-Object { $_.Metric -eq 'approve_to_ready' })
Write-Output ("bench scratch kept for inspection: " + $benchRoot)
