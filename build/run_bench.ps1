# ============================================================================
# run_bench.ps1 -- drive one method through N merge-selects and collect the
# per-stage numbers.
#
# The trigger is the real one. Nothing calls the engine directly: the script
# types digits into Notepad one WM_CHAR at a time, the way a card reader does,
# and the app has to notice by itself. That keeps the measured path identical
# to the one an operator would use.
#
# Windows this script starts, it closes. A Notepad window that was already open
# is used as-is and is never closed or killed; the same goes for any Excel that
# was already running.
#
#   pwsh -File build\run_bench.ps1 -Method csharp -Repeat 5
#   pwsh -File build\run_bench.ps1 -Method hybrid -Repeat 5
#   pwsh -File build\run_bench.ps1 -Method vba    -Repeat 5
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('csharp', 'hybrid', 'vba')] [string] $Method = 'csharp',
  [int]      $Repeat = 5,
  [string[]] $Keys = @(),
  [string]   $Root = "",
  [int]      $TypeDelayMs = 15,
  [int]      $TimeoutSec = 180,
  [switch]   $KeepOpen
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class RdvWin {
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
}
"@

$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$VP = [System.Windows.Automation.ValuePattern]

function Find-NotepadWindows {
  $c = New-Object System.Windows.Automation.PropertyCondition($AE::ClassNameProperty, 'Notepad')
  return $AE::RootElement.FindAll($TS::Children, $c)
}

# same rule the app uses: the foreground Notepad if there is one, else the last
function Pick-Notepad($wins) {
  if ($wins.Count -eq 0) { return $null }
  $fg = [RdvWin]::GetForegroundWindow()
  foreach ($w in $wins) { if ([IntPtr]$w.Current.NativeWindowHandle -eq $fg) { return $w } }
  return $wins[$wins.Count - 1]
}

function Get-TextHost($win) {
  $c = New-Object System.Windows.Automation.AndCondition(
    (New-Object System.Windows.Automation.PropertyCondition($AE::IsValuePatternAvailableProperty, $true)),
    (New-Object System.Windows.Automation.OrCondition(
      (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Document)),
      (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)))))
  return $win.FindFirst($TS::Descendants, $c)
}

# An error MessageBox carries the app's own title, so a plain name match would
# "find" the app when what is really on screen is its failure. Dialogs are
# reported instead, with their text, and never mistaken for the window.
function Wait-Window([string] $namePrefix, [int] $sec) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $w = $walker.GetFirstChild($AE::RootElement)
    while ($null -ne $w) {
      if ($w.Current.Name -like ($namePrefix + '*')) {
        if ($w.Current.ClassName -eq '#32770') {
          $msg = @()
          $c = $walker.GetFirstChild($w)
          while ($null -ne $c) { if ($c.Current.Name) { $msg += $c.Current.Name }; $c = $walker.GetNextSibling($c) }
          throw ("the app put up a dialog instead of a window: " + ($msg -join ' | '))
        }
        return $w
      }
      $w = $walker.GetNextSibling($w)
    }
    Start-Sleep -Milliseconds 200
  }
  return $null
}

# --- refuse to start on top of a previous run -------------------------------
# A left-over app window (or its error dialog) carries the same title, and the
# script would happily attach to it and measure the wrong process.
$stale = $null
try { $stale = Wait-Window 'Reader Data Viewer' 0 } catch { throw ("a Reader Data Viewer dialog is still on screen from an earlier run: " + $_.Exception.Message) }
if ($null -ne $stale) {
  throw "a Reader Data Viewer window is still open (pid $($stale.Current.ProcessId)). Close it before measuring."
}

# ---------------------------------------------------------------------------
$dist = Join-Path $Root 'dist'
$data = Join-Path $Root 'data'
if (-not (Test-Path -LiteralPath (Join-Path $data 'tableA.csv'))) {
  throw "data not generated: run build\gen_data.ps1 first"
}
if ($Keys.Count -eq 0) {
  $Keys = @()
  $pool = @('00000001', '00000042', '00500000', '01000000', '00123456', '00777777', '00000999', '00654321')
  for ($i = 0; $i -lt $Repeat; $i++) { $Keys += $pool[$i % $pool.Count] }
}
$Repeat = $Keys.Count

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$log = Join-Path $Root ("work\bench-{0}-{1}.tsv" -f $Method, $stamp)
$workDir = Split-Path -Parent $log
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }

# --- make sure there is a Notepad to read from ------------------------------
$myNotepadHwnd = [IntPtr]::Zero
$wins = Find-NotepadWindows
if ($wins.Count -eq 0) {
  Write-Output 'no Notepad window open: starting one'
  $before = @()
  foreach ($w in (Find-NotepadWindows)) { $before += $w.Current.NativeWindowHandle }
  Start-Process -FilePath "$env:WINDIR\System32\notepad.exe" | Out-Null
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 250
    $wins = Find-NotepadWindows
    if ($wins.Count -gt 0) { break }
  }
  if ($wins.Count -eq 0) { throw 'Notepad did not appear' }
  $myNotepadHwnd = [IntPtr]$wins[$wins.Count - 1].Current.NativeWindowHandle
}

# --- start the app ----------------------------------------------------------
$appProc = $null
$excelPidBefore = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$myExcel = @()

if ($Method -eq 'vba') {
  $book = Join-Path $dist 'ReaderDataViewer-VBA.xlsm'
  if (-not (Test-Path -LiteralPath $book)) { throw "not built: $book" }
  $xl = New-Object -ComObject Excel.Application
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  $after = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
  $myExcel = @($after | Where-Object { $excelPidBefore -notcontains $_ })
  Write-Output ("started Excel pid " + ($myExcel -join ','))
  $sec = $xl.AutomationSecurity
  $xl.AutomationSecurity = 1     # msoAutomationSecurityLow: this book only
  $wb = $xl.Workbooks.Open($book, 0, $true)   # read-only: the app is a viewer
  $xl.AutomationSecurity = $sec
  $ws = $wb.Worksheets.Item('VIEW')
  $ws.Range('C46').Value2 = $log            # log path cell
  $ws.Range('C45').Value2 = ''              # clear the stop flag
  $ws.Range('C47').Value2 = $data
  # OnTime, not a direct call: RDV_StartMonitor does not return until the
  # monitor stops, and a blocking Application.Run would deadlock this script
  $xl.Run("'" + (Split-Path -Leaf $book) + "'!RDV_StartMonitorAsync")
  Start-Sleep -Seconds 3
} else {
  $cmd = if ($Method -eq 'csharp') { Join-Path $dist 'ReaderDataViewer-CSharp.cmd' } else { Join-Path $dist 'ReaderDataViewer-Hybrid.cmd' }
  if (-not (Test-Path -LiteralPath $cmd)) { throw "not built: $cmd" }
  # launch the .cmd itself, not "cmd /c <quoted line>": with more than one pair
  # of quotes cmd strips the outermost pair and the rest of the line falls apart
  $appProc = Start-Process -FilePath $cmd -ArgumentList @(('"' + $data + '"'), '-log', ('"' + $log + '"')) -PassThru -WindowStyle Minimized
  Write-Output ("started app pid " + $appProc.Id)
}

# The VBA build has no window of its own: it lives in the Excel this script
# started, and it reports the Notepad it attached to in cell C3.
$appWin = $null
$boundHwnd = [IntPtr]::Zero
if ($Method -eq 'vba') {
  for ($i = 0; $i -lt 40; $i++) {
    $c3 = [string]$ws.Range('C3').Value2
    if ($c3 -match 'hwnd\s+(\d+)') { $boundHwnd = [IntPtr][int]$Matches[1]; break }
    Start-Sleep -Milliseconds 400
  }
  Write-Output ("app: " + $ws.Range('C2').Value2 + " / " + $ws.Range('C3').Value2)
} else {
  $appWin = Wait-Window 'Reader Data Viewer' 120
  if ($null -eq $appWin) { Write-Output 'WARNING: app window not found by name' }
  else { Write-Output ("app window: " + $appWin.Current.Name) }
}

# which Notepad did the app bind to? the app shows the handle, so read it back
if ($null -ne $appWin) {
  for ($i = 0; $i -lt 40; $i++) {
    $texts = $appWin.FindAll($TS::Descendants,
      (New-Object System.Windows.Automation.PropertyCondition($AE::ControlTypeProperty, [System.Windows.Automation.ControlType]::Text)))
    foreach ($t in $texts) {
      if ($t.Current.Name -match 'hwnd\s+(\d+)') { $boundHwnd = [IntPtr][int]$Matches[1]; break }
    }
    if ($boundHwnd -ne [IntPtr]::Zero) { break }
    Start-Sleep -Milliseconds 400
  }
}

$wins = Find-NotepadWindows
$target = $null
foreach ($w in $wins) { if ([IntPtr]$w.Current.NativeWindowHandle -eq $boundHwnd) { $target = $w } }
if ($null -eq $target) { $target = Pick-Notepad $wins }
if ($null -eq $target) { throw 'no Notepad window to type into' }
$targetHwnd = [IntPtr]$target.Current.NativeWindowHandle
Write-Output ("typing into Notepad hwnd " + $targetHwnd + "  (app reported " + $boundHwnd + ")")

$host_ = Get-TextHost $target
if ($null -eq $host_) { throw 'Notepad text host not found' }
$editHwnd = [IntPtr]$host_.Current.NativeWindowHandle
$vp = $host_.GetCurrentPattern($VP::Pattern)

function Type-Key([string] $k) {
  $vp.SetValue('')                       # empty field: a repeat of the same number is allowed again
  Start-Sleep -Milliseconds 350
  foreach ($ch in $k.ToCharArray()) {
    [void][RdvWin]::PostMessage($editHwnd, 0x0102, [IntPtr][int]$ch, [IntPtr]0)
    Start-Sleep -Milliseconds $TypeDelayMs
  }
}

function Log-Lines {
  if (-not (Test-Path -LiteralPath $log)) { return 0 }
  return (@(Get-Content -LiteralPath $log -ErrorAction SilentlyContinue)).Count
}

# --- the runs ---------------------------------------------------------------
$ok = 0
for ($i = 0; $i -lt $Repeat; $i++) {
  $before = Log-Lines
  $k = $Keys[$i]
  Write-Output ("--- run {0}/{1}  key {2}" -f ($i + 1), $Repeat, $k)
  Type-Key $k
  $t0 = Get-Date
  $done = $false
  while (((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec) {
    if ((Log-Lines) -gt $before) { $done = $true; break }
    Start-Sleep -Milliseconds 200
  }
  if ($done) {
    $ok++
    $last = (Get-Content -LiteralPath $log)[-1]
    $f = $last -split "`t"
    Write-Output ("    total {0} ms   detect {1} ms   {2}" -f $f[11], $f[12], $f[18])
  } else {
    Write-Output '    TIMEOUT'
  }
}

# --- shut down what this script started -------------------------------------
if (-not $KeepOpen) {
  if ($Method -eq 'vba') {
    try { $ws.Range('C45').Value2 = 'STOP' } catch { }
    Start-Sleep -Seconds 2
    try { $wb.Saved = $true; $wb.Close($false) } catch { }
    try { $xl.Quit() } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 1
    foreach ($p in $myExcel) {
      $q = Get-Process -Id $p -ErrorAction SilentlyContinue
      if ($q) { Write-Output "killing my excel $p"; $q.Kill() }
    }
  } else {
    if ($null -ne $appWin) {
      $wp = $appWin.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
      try { $wp.Close() } catch { }
    }
    Start-Sleep -Seconds 3
    if ($null -ne $appProc -and -not $appProc.HasExited) {
      Get-CimInstance Win32_Process -Filter ("ParentProcessId=" + $appProc.Id) |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
      Stop-Process -Id $appProc.Id -Force -ErrorAction SilentlyContinue
    }
  }
  if ($myNotepadHwnd -ne [IntPtr]::Zero -and [RdvWin]::IsWindow($myNotepadHwnd)) {
    [void][RdvWin]::PostMessage($myNotepadHwnd, 0x0010, [IntPtr]0, [IntPtr]0)   # WM_CLOSE, only the one we opened
  }
}

Write-Output ""
Write-Output ("{0}/{1} runs completed -> {2}" -f $ok, $Repeat, $log)
if ($ok -gt 0) { & (Join-Path $Root 'build\summarize.ps1') -Log $log }
