# ============================================================================
# test_exit_guard.ps1 -- exercise the local-pending/send boundary on the real
# packed product. A state change must leave the shared xlsx untouched. Send
# must discard an abandoned stale lock, then reread and update the xlsx. A
# close is allowed while a different, live lock still has the send waiting.
#
#   powershell -File build\test_exit_guard.ps1
#
# The test runs on its own scratch copy of dist\; shipped files are never
# touched. Only the process this script started is closed.
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [int] $ReadyTimeoutSec = 240,
  [int] $SaveTimeoutSec = 180
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type @"
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class RdvGuardWin {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll", EntryPoint="SendMessageW")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll", EntryPoint="SendMessageW", CharSet = CharSet.Unicode)] public static extern IntPtr SetText(IntPtr h, uint m, IntPtr w, string l);
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
  [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowLongW(IntPtr h, int index);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr root, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  public class Kid { public long Hwnd; public string Cls; public string Text; public int Style; }
  public static List<Kid> Kids(IntPtr root) {
    List<Kid> outp = new List<Kid>();
    EnumChildWindows(root, delegate(IntPtr h, IntPtr l) {
      StringBuilder cb = new StringBuilder(256); GetClassNameW(h, cb, 256);
      StringBuilder tb = new StringBuilder(512); GetWindowTextW(h, tb, 512);
      Kid k = new Kid(); k.Hwnd = h.ToInt64(); k.Cls = cb.ToString(); k.Text = tb.ToString(); k.Style = GetWindowLongW(h, -16);
      outp.Add(k);
      return true;
    }, IntPtr.Zero);
    return outp;
  }
}
"@

# Load the shipping types into this test process as well. They are used only
# to act as a second writer against the scratch ledger below; the running app
# is still the packed dist product in its own process.
. (Join-Path $Root 'build\sources.ps1')
$productUsings = New-Object System.Collections.Specialized.OrderedDictionary
$productBodies = New-Object System.Text.StringBuilder
foreach ($source in $RdvSources) {
  $sourceText = [IO.File]::ReadAllText((Join-Path $Root "src\csharp\$source"), [Text.Encoding]::UTF8)
  foreach ($line in ($sourceText -split "`r?`n")) {
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $using = $line.Trim()
      if (-not $productUsings.Contains($using)) { $productUsings.Add($using, $true) }
    } else { [void]$productBodies.AppendLine($line) }
  }
}
$productCode = (($productUsings.Keys | ForEach-Object { $_ }) -join "`r`n") + "`r`n`r`n" + $productBodies.ToString()
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml
$productRefs = @(
  [System.Diagnostics.Process].Assembly.Location,
  [System.Windows.Forms.Form].Assembly.Location,
  [System.Drawing.Point].Assembly.Location,
  [System.Windows.Automation.AutomationElement].Assembly.Location,
  [System.Windows.Automation.AutomationElementIdentifiers].Assembly.Location,
  [System.Windows.DependencyObject].Assembly.Location,
  [System.IO.Compression.ZipArchive].Assembly.Location,
  [System.Xml.XmlReader].Assembly.Location
)
Add-Type -TypeDefinition $productCode -ReferencedAssemblies $productRefs -Language CSharp

# The C# build receives its search key through the headless-test hook and its
# button clicks through posted messages. Neither needs keyboard focus, so the
# window stays off-screen and at the back for the whole run.
function Park-Window([IntPtr] $hwnd) {
  # HWND_BOTTOM = 1, SWP_NOSIZE|SWP_NOACTIVATE = 0x0001|0x0010
  [void][RdvGuardWin]::SetWindowPos($hwnd, [IntPtr]1, -32000, -32000, 0, 0, 0x0011)
}
[void][RdvGuardWin]::SetProcessDPIAware()
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$scratch = Join-Path $Root ("work\guard-run\{0}" -f $stamp)
New-Item -ItemType Directory -Path $scratch -Force | Out-Null
$outTsv = Join-Path $Root ("work\guard-{0}.tsv" -f $stamp)
Set-Content -LiteralPath $outTsv -Value "build`tcheck`tresult`tdetail" -Encoding utf8
$checks = New-Object System.Collections.ArrayList
function Check([string] $b, [string] $name, [bool] $ok, [string] $detail) {
  $r = if ($ok) { 'PASS' } else { 'FAIL' }
  Add-Content -LiteralPath $outTsv -Value ("{0}`t{1}`t{2}`t{3}" -f $b, $name, $r, ($detail -replace "`t", ' ')) -Encoding utf8
  [void]$checks.Add([pscustomobject]@{ Build = $b; Name = $name; Result = $r; Detail = $detail })
  Write-Output ("  [{0}] {1}  {2}" -f $r, $name, $detail)
}
function Say([string] $s) { Write-Output ("{0}  {1}" -f (Get-Date -Format 'HH:mm:ss'), $s) }

# key1 00021001 owns exactly ONE ledger row in this dataset (the generator's
# 1-row block runs from 21001 to 71000), so a search on it puts a single record
# on screen -- which is what the processed button needs.
$TargetKey1 = '00021001'

# the app's modals are its own windows (the caption is the title); their
# buttons are WinForms buttons that take BM_CLICK
function Answer-Dialog([string] $title, [string] $buttonPrefix, [int] $sec) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $dlg = [RdvGuardWin]::FindWindowW([NullString]::Value, $title)
    if ($dlg -ne [IntPtr]::Zero) {
      Start-Sleep -Milliseconds 200
      foreach ($k in [RdvGuardWin]::Kids($dlg)) {
        if ($k.Cls -like '*BUTTON*' -and ($k.Text -like ($buttonPrefix + '*'))) {
          [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0)
          return $true
        }
      }
    }
    Start-Sleep -Milliseconds 100
  }
  return $false
}
function Dialog-Present([string] $title) {
  return ([RdvGuardWin]::FindWindowW([NullString]::Value, $title) -ne [IntPtr]::Zero)
}
function Wait-DialogText([string] $title, [string] $pattern, [int] $sec) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $dlg = [RdvGuardWin]::FindWindowW([NullString]::Value, $title)
    if ($dlg -ne [IntPtr]::Zero) {
      $text = (([RdvGuardWin]::Kids($dlg) | ForEach-Object { $_.Text }) -join "`n")
      if ($text -match $pattern) { return $true }
    }
    Start-Sleep -Milliseconds 100
  }
  return $false
}
function Write-TestMarker([string] $path, [long] $version, [string] $kind, [int] $done, [int] $todo) {
  $host64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('REMOTE-HOST'))
  $user64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('remote-user'))
  $writer64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('remote-writer'))
  $text = "RDV-MARKER-1`t$version`t$host64`t$user64`t$writer64`t$([DateTime]::UtcNow.Ticks)`t100000`t$kind`t$done`t$todo`r`n"
  $temp = $path + '.test-tmp'
  [IO.File]::WriteAllText($temp, $text, (New-Object Text.UTF8Encoding($false)))
  [IO.File]::Replace($temp, $path, [NullString]::Value)
}

function Write-ChangedLedger([string] $ledger, [string] $settings, [string] $dataDir, [string] $identity) {
  $site = [Rdv3Config]::Load($settings)
  $heads = New-Object 'string[][]' $site.Data.Tables.Count
  for ($i = 0; $i -lt $site.Data.Tables.Count; $i++) {
    $heads[$i] = [Rdv3Table]::ReadHead((Join-Path $dataDir $site.Data.Tables[$i].File), $site.Data.Enc)
  }
  $site.Data.Bind($heads)
  $lines = $null
  $states = $null
  [Rdv3Xlsx]::Read($ledger, $site.Data.Head, $site.Screen.Work.Column, [ref]$lines, [ref]$states)
  $row = -1
  for ($i = 0; $i -lt $lines.Length; $i++) {
    if ([Rdv3Ledger]::FieldOf($lines[$i], $site.Data.IdentityCol) -eq $identity) { $row = $i; break }
  }
  if ($row -lt 0) { throw "identity not found in scratch ledger: $identity" }
  $cells = [Rdv3Ledger]::SplitLine($lines[$row])
  $cells[2] = $cells[2] + '-REMOTE'
  $lines[$row] = [string]::Join("`t", $cells)
  $states[$row] = $site.Screen.Work.InitialStored

  $remote = New-Object Rdv3SharedFiles $ledger, 'REMOTE-HOST', 'remote-user', 'remote-writer'
  $owner = $null
  $lease = $remote.TryAcquire([ref]$owner)
  if ($null -eq $lease) { throw 'the remote test writer could not acquire the scratch lock' }
  try {
    [Rdv3Xlsx]::Write($ledger, $site.Data.Head, $site.Screen.Work.Column, $lines, $states, 'remote-update')
    return $remote.WriteMarker('update', $lines.Length, 0, 0)
  } finally {
    $lease.Release()
  }
}

function Read-Log([string] $p, [object] $enc) {
  if (-not (Test-Path -LiteralPath $p)) { return @() }
  for ($i = 0; $i -lt 20; $i++) {
    try { return @([IO.File]::ReadAllLines($p, $enc)) } catch { Start-Sleep -Milliseconds 100 }
  }
  return @()
}
function Wait-Log([string] $p, [object] $enc, [string] $pattern, [int] $from, [int] $sec, [int] $pollMs) {
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $all = Read-Log $p $enc
    for ($i = $from; $i -lt $all.Count; $i++) { if ($all[$i] -match $pattern) { return , @($i, $all[$i]) } }
    Start-Sleep -Milliseconds $pollMs
  }
  return $null
}

if (-not (Test-Path -LiteralPath (Join-Path $Root 'dist\app-csharp\ReaderDataViewer-Ledger.xlsx'))) {
  throw "not built: run build\build_app.ps1 first"
}
Say ("target key1 = " + $TargetKey1)

# ============================================================================
# C# build
# ============================================================================
function Test-CSharp {
  $b = 'csharp'
  $dir = Join-Path $scratch 'csharp'
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  Copy-Item -Path (Join-Path $Root 'dist\app-csharp\*') -Destination $dir -Recurse -Force
  # this test types its own key: a Notepad window that happens to be open on
  # the machine must not feed the watcher a number of its own, so the scratch
  # copy watches nothing (the settings dialog's "監視対象なし" state)
  $settings = Join-Path $dir 'settings.json'
  $text = [IO.File]::ReadAllText($settings, [Text.Encoding]::UTF8)
  $text = [regex]::Replace($text, '"targets": \[[\s\S]*?\n    \]', '"targets": []')
  [IO.File]::WriteAllText($settings, $text, (New-Object Text.UTF8Encoding($false)))
  $log = Join-Path $dir 'ReaderDataViewer.log'
  $ledger = Join-Path $dir 'ReaderDataViewer-Ledger.xlsx'
  $enc = New-Object Text.UTF8Encoding($false)
  if (Test-Path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force }

  $proc = $null
  $appHostPid = 0
  $main = [IntPtr]::Zero
  $pendingFile = ''
  $fakeLock = $ledger + '.lock'
  try {
    Say 'C#: starting'
    $oldHeadless = $env:RDV_HEADLESS_TEST
    $oldHeadlessKey = $env:RDV_HEADLESS_TEST_KEY
    $env:RDV_HEADLESS_TEST = '1'
    $env:RDV_HEADLESS_TEST_KEY = $TargetKey1
    try { $proc = Start-Process -FilePath (Join-Path $dir 'ReaderDataViewer.cmd') -WindowStyle Hidden -PassThru }
    finally {
      $env:RDV_HEADLESS_TEST = $oldHeadless
      $env:RDV_HEADLESS_TEST_KEY = $oldHeadlessKey
    }
    for ($i = 0; $i -lt 100 -and $appHostPid -eq 0; $i++) {
      $child = @(Get-CimInstance Win32_Process -Filter ("ParentProcessId = " + $proc.Id) | Where-Object { $_.Name -ieq 'powershell.exe' }) | Select-Object -First 1
      if ($null -ne $child) { $appHostPid = [int]$child.ProcessId; break }
      Start-Sleep -Milliseconds 50
    }
    if ($appHostPid -eq 0) { Check $b 'host' $false 'the app host process was not found'; return }
    $ready = Wait-Log $log $enc "`tdecision`tready " 0 $ReadyTimeoutSec 200
    if ($null -eq $ready) { Check $b 'ready' $false 'never reached READY'; return }
    Check $b 'ready' $true $ready[1]
    $pendingLine = Wait-Log $log $enc "`tpending`tpath=.* count=" 0 10 100
    if ($null -ne $pendingLine -and $pendingLine[1] -match "`tpending`tpath=(.*) count=[0-9]+$") { $pendingFile = $Matches[1] }
    Check $b 'pending_path' (-not [string]::IsNullOrEmpty($pendingFile)) $pendingFile

    # the window, and the controls inside it
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $w = $walker.GetFirstChild($AE::RootElement)
    while ($null -ne $w) {
      if ($w.Current.Name -like 'Reader Data Viewer*' -and $w.Current.ProcessId -eq $appHostPid) { $main = [IntPtr]$w.Current.NativeWindowHandle; break }
      $w = $walker.GetNextSibling($w)
    }
    if ($main -eq [IntPtr]::Zero) { Check $b 'window' $false 'main window not found'; return }
    Check $b 'window' $true ("hwnd " + $main)

    # In headless-test mode the app feeds this key through the same manual
    # search path after READY. This avoids focus, activation and keyboard input
    # on the desktop while still exercising the real search and selection.
    Park-Window $main
    $from = [Math]::Max(0, [int]$ready[0])
    $hit = Wait-Log $log $enc ("`tsearch`tkey=" + $TargetKey1 + " ") $from 30 100
    if ($null -eq $hit) { Check $b 'search' $false 'no search line'; return }
    Check $b 'search' $true $hit[1]

    # Move it to the next state. This must write only the local pending file.
    $from = (Read-Log $log $enc).Count
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Text -like '未処理*') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    if (-not (Answer-Dialog '処理済の確認' 'はい' 20)) { Check $b 'confirm' $false 'confirm dialog never appeared'; return }
    Check $b 'confirm' $true 'answered はい'

    $started = Wait-Log $log $enc 'save started .*exit held' $from 20 15
    if ($null -eq $started) { Check $b 'save_started' $false 'no save-started line'; return }
    Check $b 'save_started' $true $started[1]
    $done = Wait-Log $log $enc "`tstate`tkey2=.*value=TRUE" $from $SaveTimeoutSec 100
    Check $b 'save_completed' ($null -ne $done) $(if ($null -ne $done) { $done[1] } else { 'no state line' })
    $released = Wait-Log $log $enc 'write decided .*exit released' $from 20 100
    Check $b 'exit_released' ($null -ne $released) $(if ($null -ne $released) { $released[1] } else { 'no release line' })
    $onePending = @([RdvGuardWin]::Kids($main) | Where-Object { $_.Text -eq '未送信 1 件' }).Count -eq 1
    Check $b 'pending_count_one' $onePending 'the send band reads 未送信 1 件'
    Check $b 'shared_untouched_before_send' ((Count-TrueRows $ledger) -eq 0) 'processed=TRUE rows before send: expected 0'
    Check $b 'small_local_pending' ((Test-Path -LiteralPath $pendingFile) -and ((Get-Item -LiteralPath $pendingFile).Length -lt 512)) $(if (Test-Path -LiteralPath $pendingFile) { (Get-Item -LiteralPath $pendingFile).Length } else { 'missing' })

    # Hold a synthetic, well-formed lock so the real send has to wait and name
    # its owner. The lock lives beside this scratch ledger only.
    $host64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('TEST-HOST'))
    $user64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('lock-owner'))
    $lockText = "RDV-LOCK-1`t$host64`t$user64`t$([DateTime]::UtcNow.Ticks)`r`n"
    [IO.File]::WriteAllText($fakeLock, $lockText, $enc)
    [IO.File]::SetLastWriteTimeUtc($fakeLock, [DateTime]::UtcNow.AddMinutes(-11))
    $sendFrom = (Read-Log $log $enc).Count
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Text -like '送信*') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    $sendBody = Wait-DialogText '送信' '未送信の 1 件を送信します。よろしいですか' 20
    Check $b 'send_confirm_body' $sendBody 'the confirmation includes the pending count'
    $sendEffect = Wait-DialogText '送信' '送信した行は未送信から外れます。取り込んだデータは書き換えません。' 5
    Check $b 'send_confirm_effect' $sendEffect 'the confirmation distinguishes the ledger mark from imported data'
    if (-not (Answer-Dialog '送信' 'はい' 20)) { Check $b 'send_confirm' $false 'send dialog never appeared'; return }
    Check $b 'send_confirm' $true 'answered はい'
    $stale = Wait-Log $log $enc "`tlock`tremoved stale lock age_ms=" $sendFrom 20 100
    Check $b 'stale_lock_removed' ($null -ne $stale) $(if ($null -ne $stale) { $stale[1] } else { 'no stale-removal line' })
    $staleAge = 0
    if ($null -ne $stale -and $stale[1] -match 'age_ms=([0-9]+)') { $staleAge = [long]$Matches[1] }
    Check $b 'stale_lock_threshold' ($staleAge -ge 600000) ("age_ms=" + $staleAge)
    $acquired = Wait-Log $log $enc "`tlock`tacquired " $sendFrom 20 100
    Check $b 'lock_reacquired_after_stale' ($null -ne $acquired) $(if ($null -ne $acquired) { $acquired[1] } else { 'no acquired line' })
    Check $b 'window_alive_after_stale' ([RdvGuardWin]::IsWindow($main)) 'the send continued in the same window'

    $sent = Wait-Log $log $enc "`tmarker`tversion=.* kind=send" $sendFrom $SaveTimeoutSec 100
    Check $b 'send_completed' ($null -ne $sent) $(if ($null -ne $sent) { $sent[1] } else { 'no marker line' })
    $sendReady = Wait-Log $log $enc "`tshared`tready .* pending=0" $sendFrom 30 100
    Check $b 'pending_cleared' ($null -ne $sendReady) $(if ($null -ne $sendReady) { $sendReady[1] } else { 'pending did not clear' })
    $afterSendKids = @([RdvGuardWin]::Kids($main))
    $zeroPending = @($afterSendKids | Where-Object { $_.Text -eq '未送信 0 件' }).Count -eq 1
    $repeatFrom = (Read-Log $log $enc).Count
    foreach ($k in $afterSendKids) {
      if ($k.Cls -like '*BUTTON*' -and $k.Text -like '検索*') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    $repeatHit = Wait-Log $log $enc ("`tsearch`tkey=" + $TargetKey1 + " ") $repeatFrom 30 100
    $afterRepeatKids = @([RdvGuardWin]::Kids($main))
    $stateButtons = @($afterRepeatKids | Where-Object { $_.Cls -like '*BUTTON*' -and $_.Text -match '処理済' })
    $stateButton = $stateButtons | Select-Object -First 1
    $stateDetail = if ($null -eq $stateButton) {
      'no processed-state button; buttons=' + ((@($afterRepeatKids | Where-Object { $_.Cls -like '*BUTTON*' } | ForEach-Object { $_.Text })) -join '|')
    } else {
      'text=' + $stateButton.Text + ' style=0x' + ([uint32]$stateButton.Style).ToString('X8')
    }
    Check $b 'pending_count_zero' $zeroPending 'the send band reads 未送信 0 件'
    Check $b 'sent_row_not_frozen' (($null -ne $repeatHit) -and ($null -ne $stateButton) -and (($stateButton.Style -band 0x08000000) -eq 0)) ('after a repeat lookup: ' + $stateDetail)
    Check $b 'shared_after_send' ((Count-TrueRows $ledger) -eq 1) 'processed=TRUE rows after send: expected 1'
    Check $b 'marker_written' ((Test-Path -LiteralPath ($ledger + '.version')) -and (([IO.File]::ReadAllLines($ledger + '.version')).Count -eq 1)) 'one-line marker'
    Check $b 'lock_released' (-not (Test-Path -LiteralPath $fakeLock)) 'no lock file remains'

    # A marker from another writer must be noticed without polling the xlsx.
    # Consecutive send notices replace one StatusStrip item without creating a
    # dialog; update notifications still ask whether to switch.
    $remoteFrom = (Read-Log $log $enc).Count
    Write-TestMarker ($ledger + '.version') 2 'send' 12 3
    $remoteFirst = Wait-Log $log $enc "`tnotice`ttarget=status text=remote-user が 12 件を処理済、3 件を未処理にしました" $remoteFrom 15 100
    Write-TestMarker ($ledger + '.version') 3 'send' 7 4
    $remoteSecond = Wait-Log $log $enc "`tnotice`ttarget=status text=remote-user が 7 件を処理済、4 件を未処理にしました" $remoteFrom 15 100
    $noSendDialog = -not (Dialog-Present '台帳が更新されました')
    Check $b 'remote_send_status' (($null -ne $remoteFirst) -and ($null -ne $remoteSecond) -and $noSendDialog) ('first=' + $(if ($null -ne $remoteFirst) { $remoteFirst[1] } else { 'missing' }) + '; second=' + $(if ($null -ne $remoteSecond) { $remoteSecond[1] } else { 'missing' }) + '; noDialog=' + $noSendDialog)
    $remoteReload = Wait-Log $log $enc "`treload`tversion=3 .*rows=100000" $remoteFrom 30 100
    Check $b 'remote_send_reload' ($null -ne $remoteReload) $(if ($null -ne $remoteReload) { $remoteReload[1] } else { 'no reload line' })
    $remoteReady = Wait-Log $log $enc "`tshared`tready .*note=marker-3" $remoteFrom 30 100
    Check $b 'remote_send_ready' ($null -ne $remoteReady) $(if ($null -ne $remoteReady) { $remoteReady[1] } else { 'marker-3 was not adopted' })

    $updateFrom = (Read-Log $log $enc).Count
    $changedMarker = Write-ChangedLedger $ledger $settings (Join-Path $dir 'data') '00089897'
    Check $b 'remote_update_written' (($changedMarker.Version -eq 4) -and ((Count-TrueRows $ledger) -eq 0)) ('version ' + $changedMarker.Version + ', processed=TRUE rows 0')
    $updateBody = Wait-DialogText '台帳の更新' '台帳が更新されました。切り替えますか' 15
    Check $b 'remote_update_dialog' $updateBody 'the update asks before switching'
    $resetBody = Wait-DialogText '台帳の更新' '中身が変わったため未処理に戻ったレコード: 1 件' 5
    $updateDialog = [RdvGuardWin]::FindWindowW([NullString]::Value, '台帳の更新')
    $resetList = if ($updateDialog -eq [IntPtr]::Zero) { $false } else { @([RdvGuardWin]::Kids($updateDialog) | Where-Object { $_.Cls -like '*SysListView32*' }).Count -eq 1 }
    Check $b 'remote_update_reset_count' $resetBody 'the update dialog reports one reset record'
    Check $b 'remote_update_reset_list' $resetList 'the reset records use one standard Details ListView'
    [void](Answer-Dialog '台帳の更新' 'OK' 5)
    $updateReady = Wait-Log $log $enc "`tshared`tready .*note=marker-4" $updateFrom 30 100
    Check $b 'remote_update_switch' ($null -ne $updateReady) $(if ($null -ne $updateReady) { $updateReady[1] } else { 'no marker-4 ready line' })

    # A fresh lock is live, so it is not removed. Closing while this send is
    # still waiting must end the app without a write-in-flight refusal.
    $closeFrom = (Read-Log $log $enc).Count
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Cls -like '*BUTTON*' -and $k.Text -like '検索*') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    $closeHit = Wait-Log $log $enc ("`tsearch`tkey=" + $TargetKey1 + " ") $closeFrom 30 100
    if ($null -eq $closeHit) { Check $b 'closes_while_lock_waiting' $false 'the setup search did not finish'; return }
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Text -like '未処理*') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    if (-not (Answer-Dialog '処理済の確認' 'はい' 20)) { Check $b 'closes_while_lock_waiting' $false 'the setup state dialog never appeared'; return }
    $closeSaved = Wait-Log $log $enc "`tstate`tkey2=.*value=TRUE" $closeFrom $SaveTimeoutSec 100
    if ($null -eq $closeSaved) { Check $b 'closes_while_lock_waiting' $false 'the setup state did not save'; return }
    $lockText = "RDV-LOCK-1`t$host64`t$user64`t$([DateTime]::UtcNow.Ticks)`r`n"
    [IO.File]::WriteAllText($fakeLock, $lockText, $enc)
    $closeSendFrom = (Read-Log $log $enc).Count
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Text -like '送信*') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    if (-not (Answer-Dialog '送信' 'はい' 20)) { Check $b 'closes_while_lock_waiting' $false 'the setup send dialog never appeared'; return }
    $closeWaiting = Wait-Log $log $enc "`tlock`twaiting owner=lock-owner host=TEST-HOST" $closeSendFrom 20 100
    if ($null -eq $closeWaiting) { Check $b 'closes_while_lock_waiting' $false 'the send did not wait behind the live lock'; return }
    [void][RdvGuardWin]::PostMessage($main, 0x0010, [IntPtr]0, [IntPtr]0)
    $closed = $false
    for ($i = 0; $i -lt 40; $i++) {
      if (-not [RdvGuardWin]::IsWindow($main)) { $closed = $true; break }
      Start-Sleep -Milliseconds 250
    }
    $closingLog = Wait-Log $log $enc "`texit`tclosing" $closeSendFrom 5 100
    Check $b 'closes_while_lock_waiting' ($closed -and ($null -ne $closingLog)) ('closed=' + $closed + '; wait=' + $closeWaiting[1] + '; closing=' + $(if ($null -ne $closingLog) { $closingLog[1] } else { 'missing' }))
    Start-Sleep -Seconds 2

    # The pending state reached the ledger through send, then the simulated
    # external content update reset that changed record to the initial state.
    $n = Count-TrueRows $ledger
    Check $b 'changed_record_reset' ($n -eq 0) ("processed=TRUE rows after the changed-record update: " + $n)
  } finally {
    if ($main -ne [IntPtr]::Zero -and [RdvGuardWin]::IsWindow($main)) {
      [void](Answer-Dialog '処理済の確認' 'いいえ' 1)
      [void][RdvGuardWin]::PostMessage($main, 0x0010, [IntPtr]0, [IntPtr]0)
      Start-Sleep -Seconds 3
    }
    if ($appHostPid -ne 0) {
      $q = Get-Process -Id $appHostPid -ErrorAction SilentlyContinue
      if ($q) { Say ("  closing my app host " + $appHostPid); $q.Kill() }
    }
    if ($null -ne $proc -and -not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not [string]::IsNullOrEmpty($pendingFile) -and (Test-Path -LiteralPath $pendingFile)) {
      Remove-Item -LiteralPath $pendingFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $fakeLock) { Remove-Item -LiteralPath $fakeLock -Force -ErrorAction SilentlyContinue }
  }
}

# how many ledger rows carry processed=TRUE, read straight out of the package
function Count-TrueRows([string] $path) {
  $z = [IO.Compression.ZipFile]::OpenRead($path)
  try {
    $sheet = $null
    foreach ($e in $z.Entries) {
      if ($e.FullName -like 'xl/worksheets/*.xml' -and ($null -eq $sheet -or $e.Length -gt $sheet.Length)) { $sheet = $e }
    }
    if ($null -eq $sheet) { return -1 }
    $sr = New-Object IO.StreamReader($sheet.Open(), [Text.Encoding]::UTF8)
    $n = 0
    $buf = New-Object char[] 262144
    $carry = ''
    while (($r = $sr.Read($buf, 0, $buf.Length)) -gt 0) {
      $s = $carry + (New-Object string($buf, 0, $r))
      # The processed cell is the first cell of a row. The C# writer stores
      # TRUE as an inline string.
      $n += ([regex]::Matches($s, '<row[^>]*><c[^>]*><is><t>TRUE</t>')).Count
      if ($s.Length -gt 64) { $carry = $s.Substring($s.Length - 64) } else { $carry = $s }
    }
    $sr.Close()
    return $n
  } finally { $z.Dispose() }
}

# ============================================================================
# Run
# ============================================================================
Write-Output ''
Write-Output '=== C# build'
try {
  Test-CSharp
} catch {
  Check 'csharp' 'harness' $false ("harness error: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber + " | " + ($_.ScriptStackTrace -replace "`r?`n", ' <- '))
}

Write-Output ''
$pass = @($checks | Where-Object { $_.Result -eq 'PASS' }).Count
$fail = @($checks | Where-Object { $_.Result -eq 'FAIL' }).Count
Write-Output ("=== {0} checks: {1} PASS, {2} FAIL   -> {3}" -f $checks.Count, $pass, $fail, $outTsv)
foreach ($c in $checks) {
  if ($c.Result -eq 'FAIL') {
    Write-Output ("  FAIL {0}/{1}: {2}" -f $c.Build, $c.Name, $c.Detail)
  }
}
Write-Output ("scratch: {0}" -f $scratch)
if ($fail -gt 0) { exit 1 }
