# ============================================================================
# run_bench2.ps1 -- drive one of the four one-to-many methods through N
# merge-selects and collect the per-stage numbers.
#
#   chash   dist\ReaderDataViewer-CSharp-Hash.cmd    C# hand-built hash
#   cdict   dist\ReaderDataViewer-CSharp-Dict.cmd    C# standard Dictionary
#   vhash   dist\ReaderDataViewer-VBA-Hash.xlsm      VBA hand-built hash
#   vdict   dist\ReaderDataViewer-VBA-Dict.xlsm      VBA Scripting.Dictionary
#
# The trigger is the real one. Nothing calls the engine directly: the script
# types digits into Notepad one WM_CHAR at a time, the way a card reader does,
# and the app has to notice by itself.
#
# All four get the SAME key sequence, so they meet the same candidate counts in
# the same order. The default keys are chosen from the generated shape:
#
#   00000001  5 candidates      00006001  2 candidates
#   00001001  3 candidates      00021001  1 candidate
#
# -AutoPick answers the candidate dialog so a run needs nobody present. It acts
# after the measured region has already ended, at exactly the point a person
# would have clicked, and the time it takes is reported separately as "pick".
#
# Windows this script starts, it closes. A Notepad window that was already open
# is used as-is and is never closed or killed; the same goes for any Excel that
# was already running.
#
#   pwsh -File build\run_bench2.ps1 -Method chash -Repeat 8
#   pwsh -File build\run_bench2.ps1 -Method vdict -Repeat 8
#   pwsh -File build\run_bench2.ps1 -Method vhash -Data data-tiny -Keys 00000001
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('chash', 'cdict', 'vhash', 'vdict')] [string] $Method = 'chash',
  [int]      $Repeat = 8,
  [string[]] $Keys = @(),
  [string]   $Data = 'data-100k',
  [int]      $AutoPick = 1,
  [string]   $Root = "",
  [int]      $TypeDelayMs = 15,
  [int]      $TimeoutSec = 180,
  [int]      $RetryAfterSec = 45,
  [switch]   $KeepOpen
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class Rdv2Win {
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
  $fg = [Rdv2Win]::GetForegroundWindow()
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
# "find" the app when what is really on screen is its failure.
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

$isVba = ($Method -eq 'vhash' -or $Method -eq 'vdict')

# --- refuse to start on top of a previous run -------------------------------
$stale = $null
try { $stale = Wait-Window 'Reader Data Viewer' 0 } catch { throw ("a Reader Data Viewer dialog is still on screen from an earlier run: " + $_.Exception.Message) }
if ($null -ne $stale) {
  throw "a Reader Data Viewer window is still open (pid $($stale.Current.ProcessId)). Close it before measuring."
}
# The workbooks have no window of their own -- they are an Excel with a title
# like "ReaderDataViewer-VBA-Hash.xlsm - Excel", which the check above cannot
# see. One left monitoring from an earlier run keeps reading the same Notepad
# and runs its own merge-select against every key typed here, which shows up as
# nothing but slower numbers for whatever is being measured. Refuse, and say
# which process to close: this script never kills an Excel it did not start.
$staleBooks = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -like '*ReaderDataViewer*' })
if ($staleBooks.Count -gt 0) {
  throw ("a Reader Data Viewer workbook is still open in Excel (pid " +
         (($staleBooks | ForEach-Object { $_.Id }) -join ', ') + "). Close it before measuring.")
}

# ---------------------------------------------------------------------------
$dist = Join-Path $Root 'dist'
$dataDir = $Data
if (-not [IO.Path]::IsPathRooted($dataDir)) { $dataDir = Join-Path $Root $Data }
if (-not (Test-Path -LiteralPath (Join-Path $dataDir 'tableA.csv'))) {
  throw "data not generated: run build\gen_data2.ps1 first (looked in $dataDir)"
}
if ($Keys.Count -eq 0) {
  $Keys = @()
  $pool = @('00000001', '00001001', '00006001', '00021001')
  for ($i = 0; $i -lt $Repeat; $i++) { $Keys += $pool[$i % $pool.Count] }
} else {
  # powershell -File hands "a,b,c" over as one string, so accept either that or
  # a real array. Getting this wrong types a 35-character "key" that no rule can
  # match, and the run just times out.
  $Keys = @($Keys -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_.Length -gt 0 })
}
$Repeat = $Keys.Count

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$log = Join-Path $Root ("work\bench2-{0}-{1}.tsv" -f $Method, $stamp)
$workDir = Split-Path -Parent $log
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }

# --- make sure there is a Notepad to read from ------------------------------
$myNotepadHwnd = [IntPtr]::Zero
$wins = Find-NotepadWindows
if ($wins.Count -eq 0) {
  Write-Output 'no Notepad window open: starting one'
  Start-Process -FilePath "$env:WINDIR\System32\notepad.exe" | Out-Null
  for ($i = 0; $i -lt 60; $i++) {
    Start-Sleep -Milliseconds 250
    $wins = Find-NotepadWindows
    if ($wins.Count -gt 0) { break }
  }
  if ($wins.Count -eq 0) { throw 'Notepad did not appear' }
  $myNotepadHwnd = [IntPtr]$wins[$wins.Count - 1].Current.NativeWindowHandle
}

# Empty every Notepad field BEFORE the app starts. A number left over from an
# earlier run is a valid reading, and the app is right to fire on it the moment
# it attaches -- but that run is not one of the ones being asked for, and it
# lands in the log ahead of them.
foreach ($w in (Find-NotepadWindows)) {
  $h = Get-TextHost $w
  if ($null -ne $h) {
    try { $h.GetCurrentPattern($VP::Pattern).SetValue('') } catch { }
  }
}
Start-Sleep -Milliseconds 400

# --- start the app ----------------------------------------------------------
$appProc = $null
$myExcel = @()
$ws = $null
$wb = $null
$xl = $null

if ($isVba) {
  $book = Join-Path $dist $(if ($Method -eq 'vhash') { 'ReaderDataViewer-VBA-Hash.xlsm' } else { 'ReaderDataViewer-VBA-Dict.xlsm' })
  if (-not (Test-Path -LiteralPath $book)) { throw "not built: $book" }
  # identity is settled before anything is done to it (excel_own.ps1):
  # a reused or unidentifiable instance throws here and is never driven
  $rdvOwn = New-OwnedExcel
  $xl = $rdvOwn.App
  $myExcel = @($rdvOwn.Pid)
  $xl.Visible = $true
  $xl.DisplayAlerts = $false
  Write-Output ("started Excel pid " + ($myExcel -join ','))
  $sec = $xl.AutomationSecurity
  $xl.AutomationSecurity = 1     # msoAutomationSecurityLow: this book only
  $wb = $xl.Workbooks.Open($book, 0, $true)   # read-only: the app is a viewer
  $xl.AutomationSecurity = $sec
  $ws = $wb.Worksheets.Item('VIEW')
  $ws.Range('C59').Value2 = $log          # log path
  $ws.Range('C58').Value2 = ''            # clear the stop flag
  $ws.Range('C60').Value2 = $dataDir
  $ws.Range('C61').Value2 = $AutoPick
  $xl.Run("'" + (Split-Path -Leaf $book) + "'!RDV2_StartMonitorAsync")
  Start-Sleep -Seconds 3
} else {
  $cmd = Join-Path $dist $(if ($Method -eq 'chash') { 'ReaderDataViewer-CSharp-Hash.cmd' } else { 'ReaderDataViewer-CSharp-Dict.cmd' })
  if (-not (Test-Path -LiteralPath $cmd)) { throw "not built: $cmd" }
  # launch the .cmd itself, not "cmd /c <quoted line>": with more than one pair
  # of quotes cmd strips the outermost pair and the rest of the line falls apart
  $appProc = Start-Process -FilePath $cmd -ArgumentList @(('"' + $dataDir + '"'), '-log', ('"' + $log + '"'), '-autopick', "$AutoPick") -PassThru -WindowStyle Minimized
  Write-Output ("started app pid " + $appProc.Id)
}

# The VBA builds have no window of their own: they live in the Excel this
# script started, and report the Notepad they attached to in cell C3.
$appWin = $null
$boundHwnd = [IntPtr]::Zero
if ($isVba) {
  for ($i = 0; $i -lt 60; $i++) {
    $c3 = [string]$ws.Range('C3').Value2
    if ($c3 -match 'hwnd\s+(\d+)') { $boundHwnd = [IntPtr][int]$Matches[1]; break }
    Start-Sleep -Milliseconds 400
  }
  Write-Output ("app: " + $ws.Range('C2').Value2 + " / " + $ws.Range('C3').Value2 + " / index=" + $ws.Range('C5').Value2)
} else {
  $appWin = Wait-Window 'Reader Data Viewer' 180
  if ($null -eq $appWin) { Write-Output 'WARNING: app window not found by name' }
  else { Write-Output ("app window: " + $appWin.Current.Name) }
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
    [void][Rdv2Win]::PostMessage($editHwnd, 0x0102, [IntPtr][int]$ch, [IntPtr]0)
    Start-Sleep -Milliseconds $TypeDelayMs
  }
}

function Log-Lines {
  if (-not (Test-Path -LiteralPath $log)) { return 0 }
  return (@(Get-Content -LiteralPath $log -ErrorAction SilentlyContinue)).Count
}

# --- the runs ---------------------------------------------------------------
# A reading can be missed: if the app is still finishing the previous run when
# the digits arrive, or has just rebound to another Notepad, the field changes
# while nothing is polling it and there is no second edge to notice. The key is
# typed again rather than counted as a failed run, and the retry is printed so
# it is never invisible.
$ok = 0
$retried = 0
for ($i = 0; $i -lt $Repeat; $i++) {
  $before = Log-Lines
  $k = $Keys[$i]
  Write-Output ("--- run {0}/{1}  key {2}" -f ($i + 1), $Repeat, $k)
  Type-Key $k
  $t0 = Get-Date
  $done = $false
  $sentAgain = $false
  while (((Get-Date) - $t0).TotalSeconds -lt $TimeoutSec) {
    if ((Log-Lines) -gt $before) { $done = $true; break }
    if (-not $sentAgain -and ((Get-Date) - $t0).TotalSeconds -gt $RetryAfterSec) {
      $sentAgain = $true
      $retried++
      if ($isVba) { Write-Output ("    no reading after {0}s; app says: {1} / {2}" -f $RetryAfterSec, $ws.Range('C2').Value2, $ws.Range('C3').Value2) }
      else { Write-Output ("    no reading after {0}s; typing it again" -f $RetryAfterSec) }
      Type-Key $k
    }
    Start-Sleep -Milliseconds 200
  }
  if ($done) {
    $ok++
    $last = (Get-Content -LiteralPath $log)[-1]
    $f = $last -split "`t"
    # 0 seq 1 time 2 key 3 index 4..13 stages 14 other 15 total 16 pick
    # 17 detect 18 polls 19 cands 20 rows 21 probes 22 checksum 23 matchedAB
    # 24 matchedBC 25 oracle 26 error
    Write-Output ("    cands {0}   total {1} ms   pick {2} ms   detect {3} ms   {4}" -f $f[19], $f[15], $f[16], $f[17], $f[25])
  } else {
    Write-Output '    TIMEOUT'
  }
}

# --- shut down what this script started -------------------------------------
if (-not $KeepOpen) {
  if ($isVba) {
    try { $ws.Range('C58').Value2 = 'STOP' } catch { }
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
  if ($myNotepadHwnd -ne [IntPtr]::Zero -and [Rdv2Win]::IsWindow($myNotepadHwnd)) {
    [void][Rdv2Win]::PostMessage($myNotepadHwnd, 0x0010, [IntPtr]0, [IntPtr]0)   # WM_CLOSE, only the one we opened
  }
}

Write-Output ""
Write-Output ("{0}/{1} runs completed ({2} needed the key typed again) -> {3}" -f $ok, $Repeat, $retried, $log)
if ($ok -gt 0) { & (Join-Path $Root 'build\summarize2.ps1') -Log $log }
