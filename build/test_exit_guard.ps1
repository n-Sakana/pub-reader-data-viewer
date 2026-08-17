# ============================================================================
# test_exit_guard.ps1 -- the one thing the exit guard exists for, tried for
# real on both builds:
#
#   mark a record processed, and ask the app to CLOSE while that one save is
#   still in flight. The close must be refused with the reason on screen, the
#   record must reach the ledger file anyway, and only then may the app close.
#
#   powershell -File build\test_exit_guard.ps1              (both builds)
#   powershell -File build\test_exit_guard.ps1 -Build vba
#
# Both builds run on their own scratch copy of dist\ -- the shipped files are
# never touched. Only processes this script started are ever closed, and the
# close is always a WM_CLOSE to a window, never a process kill (a killed Excel
# blocklists the path in Resiliency\DisabledItems, measured in docs\app.md).
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('both', 'csharp', 'vba')] [string] $Build = 'both',
  [string] $Root = "",
  [int] $ReadyTimeoutSec = 240,
  [int] $SaveTimeoutSec = 180
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
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
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string cls, string title);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr root, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
  [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
  [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
  public class Kid { public long Hwnd; public string Cls; public string Text; }
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

# The C# build is driven entirely by posted messages (WM_CHAR to the edit,
# BM_CLICK to the buttons), which do not need the keyboard focus. So the window
# is pushed to the BACK of the z-order as soon as it appears and whatever the
# operator was using keeps the foreground: a test run no longer interrupts them.
function Park-Window([IntPtr] $hwnd) {
  $keep = [RdvGuardWin]::GetForegroundWindow()
  # HWND_BOTTOM = 1, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE = 0x0002|0x0001|0x0010
  [void][RdvGuardWin]::SetWindowPos($hwnd, [IntPtr]1, 0, 0, 0, 0, 0x0013)
  if ($keep -ne [IntPtr]::Zero -and $keep -ne $hwnd) {
    [void][RdvGuardWin]::SetForegroundWindow($keep)
  }
}
[void][RdvGuardWin]::SetProcessDPIAware()
$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$VP = [System.Windows.Automation.ValuePattern]

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
# on screen -- which is what the processed button needs. Its ledger row is
# found by scanning the sidecar mirror, because table B is written shuffled.
$TargetKey1 = '00021001'
$TargetRow = -1             # filled from the sidecar below

function Answer-Dialog([string] $title, [string] $buttonPrefix, [int] $sec) {
  $cmdId = 1
  if ($buttonPrefix -like 'はい*') { $cmdId = 6 }
  elseif ($buttonPrefix -like 'いいえ*') { $cmdId = 7 }
  $t0 = Get-Date
  while (((Get-Date) - $t0).TotalSeconds -lt $sec) {
    $dlg = [RdvGuardWin]::FindWindowW('#32770', $title)
    if ($dlg -ne [IntPtr]::Zero) {
      Start-Sleep -Milliseconds 200
      $btn = [IntPtr]::Zero
      foreach ($k in [RdvGuardWin]::Kids($dlg)) {
        if ($k.Cls -eq 'Button' -and ($k.Text.Replace('&', '') -like ($buttonPrefix + '*'))) { $btn = [IntPtr]$k.Hwnd; break }
      }
      if ($btn -ne [IntPtr]::Zero) { [void][RdvGuardWin]::PostMessage($btn, 0x00F5, [IntPtr]0, [IntPtr]0) }
      else { [void][RdvGuardWin]::PostMessage($dlg, 0x0111, [IntPtr]$cmdId, [IntPtr]0) }
      return $true
    }
    Start-Sleep -Milliseconds 100
  }
  return $false
}
function Dialog-Present([string] $title) {
  return ([RdvGuardWin]::FindWindowW('#32770', $title) -ne [IntPtr]::Zero)
}

# Answer a modal from ANOTHER process. Needed when the call that raises the
# modal does not return until it is answered (Hyperlink.Follow on a button whose
# handler shows a confirm), which no single thread can do for itself. The child
# clicks the real button through UI Automation; nothing is posted blindly.
function Arm-Answer([string] $title, [string] $buttonPrefix, [int] $sec) {
  $ps = Join-Path $scratch ('answer-' + [Guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
  # Win32, not UI Automation: this MsgBox belongs to Excel, and UIA does not
  # list it under the desktop (measured: FindFirst by name at Children scope
  # returns nothing while FindWindowW('#32770', title) finds it at once). Same
  # mechanism Answer-Dialog above already uses.
  # Template is single-quoted so nothing in it is interpolated here; the three
  # placeholders are substituted afterwards.
  $tpl = @'
Add-Type @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class RdvAns {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr FindWindowW(string c, string t);
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr r, EnumProc p, IntPtr l);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassNameW(IntPtr h, StringBuilder b, int n);
}
"@
$t0 = Get-Date
while (((Get-Date) - $t0).TotalSeconds -lt __SECS__) {
  $d = [RdvAns]::FindWindowW('#32770', '__TITLE__')
  if ($d -ne [IntPtr]::Zero) {
    Start-Sleep -Milliseconds 250
    $script:hit = [IntPtr]::Zero
    $cb = [RdvAns+EnumProc]{
      param($h, $l)
      $c = New-Object System.Text.StringBuilder 256
      [void][RdvAns]::GetClassNameW($h, $c, 256)
      $t = New-Object System.Text.StringBuilder 256
      [void][RdvAns]::GetWindowTextW($h, $t, 256)
      if ($c.ToString() -eq 'Button' -and $t.ToString().Replace('&', '') -like '__BTN__*') { $script:hit = $h }
      return $true
    }
    [void][RdvAns]::EnumChildWindows($d, $cb, [IntPtr]::Zero)
    if ($script:hit -ne [IntPtr]::Zero) {
      [void][RdvAns]::PostMessage($script:hit, 0x00F5, [IntPtr]0, [IntPtr]0)
      exit 0
    }
  }
  Start-Sleep -Milliseconds 200
}
exit 1
'@
  $body = $tpl.Replace('__SECS__', [string]$sec).Replace('__TITLE__', $title).Replace('__BTN__', $buttonPrefix)
  Set-Content -LiteralPath $ps -Value $body -Encoding UTF8
  return (Start-Process powershell -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $ps `
          -PassThru -WindowStyle Hidden)
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

# the ledger row that carries this key1, found in the sidecar mirror
function Find-TargetRow([string] $statePath, [string] $key1) {
  $fs = [IO.File]::OpenRead($statePath)
  try {
    $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::Unicode)
    [void]$sr.ReadLine()
    $i = 0
    while ($null -ne ($line = $sr.ReadLine())) {
      $t1 = $line.IndexOf("`t")
      if ($t1 -ge 0) {
        $t2 = $line.IndexOf("`t", $t1 + 1)
        if ($t2 -gt $t1 -and $line.Substring($t1 + 1, $t2 - $t1 - 1) -eq $key1) { return $i }
      }
      $i++
    }
    return -1
  } finally { $fs.Dispose() }
}

$srcState = Join-Path $Root 'dist\app-vba\ReaderDataViewer-Ledger.state'
if (-not (Test-Path -LiteralPath $srcState)) { throw "not built: run build\build_app.ps1 first" }
$TargetRow = Find-TargetRow $srcState $TargetKey1
if ($TargetRow -lt 0) { throw "key1 $TargetKey1 is not in the ledger" }
Say ("target key1 = {0} -> ledger row {1} (sheet row {2})" -f $TargetKey1, $TargetRow, ($TargetRow + 2))

# ============================================================================
# C# build
# ============================================================================
function Test-CSharp {
  $b = 'csharp'
  $dir = Join-Path $scratch 'csharp'
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  Copy-Item -Path (Join-Path $Root 'dist\app-csharp\*') -Destination $dir -Recurse -Force
  $log = Join-Path $dir 'ReaderDataViewer.log'
  $ledger = Join-Path $dir 'ReaderDataViewer-Ledger.xlsx'
  $enc = New-Object Text.UTF8Encoding($false)
  if (Test-Path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force }

  $proc = $null
  $main = [IntPtr]::Zero
  try {
    Say 'C#: starting'
    $proc = Start-Process -FilePath (Join-Path $dir 'ReaderDataViewer.cmd') -PassThru
    $ready = Wait-Log $log $enc "`tdecision`tready " 0 $ReadyTimeoutSec 200
    if ($null -eq $ready) { Check $b 'ready' $false 'never reached READY'; return }
    Check $b 'ready' $true $ready[1]

    # the window, and the controls inside it
    $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
    $w = $walker.GetFirstChild($AE::RootElement)
    while ($null -ne $w) {
      if ($w.Current.Name -like 'Reader Data Viewer*' -and $w.Current.ProcessId -ne $PID) { $main = [IntPtr]$w.Current.NativeWindowHandle; break }
      $w = $walker.GetNextSibling($w)
    }
    if ($main -eq [IntPtr]::Zero) { Check $b 'window' $false 'main window not found'; return }
    Check $b 'window' $true ("hwnd " + $main)

    # search: put the key in the box, then click the button. The edit is found
    # by child-window class first (child enumeration is what works reliably on
    # this WinForms window) and bound through UIA from its handle.
    Park-Window $main
    Start-Sleep -Milliseconds 400
    Say ("C#: looking for the search box in hwnd " + $main)
    # typed in, one WM_CHAR per digit, the way a person types: the WinForms
    # element does not offer ValuePattern here, and a posted keystroke is what
    # the control was built to take anyway
    $editH = [IntPtr]::Zero
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Cls -like '*.EDIT.*') { $editH = [IntPtr]$k.Hwnd; break }
    }
    if ($editH -eq [IntPtr]::Zero) {
      $cls = (([RdvGuardWin]::Kids($main)) | ForEach-Object { $_.Cls } | Sort-Object -Unique) -join ','
      Check $b 'searchbox' $false ("no search box; child classes: " + $cls)
      return
    }
    foreach ($ch in $TargetKey1.ToCharArray()) {
      [void][RdvGuardWin]::PostMessage($editH, 0x0102, [IntPtr][int]$ch, [IntPtr]0)
      Start-Sleep -Milliseconds 20
    }
    Start-Sleep -Milliseconds 300
    Say ("C#: typed " + $TargetKey1 + " into the search box (hwnd " + $editH + ")")
    $from = (Read-Log $log $enc).Count
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Text -eq '検索') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    $hit = Wait-Log $log $enc ("`tsearch`tkey=" + $TargetKey1 + " ") $from 30 100
    if ($null -eq $hit) { Check $b 'search' $false 'no search line'; return }
    Check $b 'search' $true $hit[1]

    # mark it processed, answer the confirm, then ask to close DURING the save
    $from = (Read-Log $log $enc).Count
    foreach ($k in [RdvGuardWin]::Kids($main)) {
      if ($k.Text -eq '処理済み') { [void][RdvGuardWin]::PostMessage([IntPtr]$k.Hwnd, 0x00F5, [IntPtr]0, [IntPtr]0); break }
    }
    if (-not (Answer-Dialog '処理済みの確認' 'はい' 20)) { Check $b 'confirm' $false 'confirm dialog never appeared'; return }
    Check $b 'confirm' $true 'answered はい'

    # the app logs the save start before it posts the job; the write itself is
    # sub-second, so this polls fast and closes the moment it is in flight
    $started = Wait-Log $log $enc 'save started .*exit held' $from 20 15
    if ($null -eq $started) { Check $b 'save_started' $false 'no save-started line'; return }
    Check $b 'save_started' $true $started[1]
    [void][RdvGuardWin]::PostMessage($main, 0x0010, [IntPtr]0, [IntPtr]0)   # WM_CLOSE
    Say 'C#: WM_CLOSE posted while the save was in flight'

    # the close must be refused, said out loud, and the window must survive
    $blocked = $false
    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt 10) {
      if (Dialog-Present '保存中のため終了できません') { $blocked = $true; break }
      Start-Sleep -Milliseconds 50
    }
    Check $b 'close_refused_dialog' $blocked 'the close was refused with a dialog naming the reason'
    if ($blocked) { [void](Answer-Dialog '保存中のため終了できません' 'OK' 5) }
    Check $b 'window_alive' ([RdvGuardWin]::IsWindow($main)) 'the window is still open after the refused close'

    $refused = Wait-Log $log $enc 'close refused: a processed save is still in flight' $from 10 100
    Check $b 'close_refused_log' ($null -ne $refused) $(if ($null -ne $refused) { $refused[1] } else { 'no refusal line' })

    # the save must still finish, and be reported as decided
    $done = Wait-Log $log $enc "`tprocessed`tkey2=.*value=TRUE" $from $SaveTimeoutSec 100
    Check $b 'save_completed' ($null -ne $done) $(if ($null -ne $done) { $done[1] } else { 'no processed line' })
    $released = Wait-Log $log $enc 'processed save decided .*exit released' $from 20 100
    Check $b 'exit_released' ($null -ne $released) $(if ($null -ne $released) { $released[1] } else { 'no release line' })

    # and only now may it close
    Start-Sleep -Milliseconds 500
    [void][RdvGuardWin]::PostMessage($main, 0x0010, [IntPtr]0, [IntPtr]0)
    $closed = $false
    for ($i = 0; $i -lt 40; $i++) {
      if (-not [RdvGuardWin]::IsWindow($main)) { $closed = $true; break }
      Start-Sleep -Milliseconds 250
    }
    Check $b 'closes_after' $closed 'the second close request ended the app'
    Start-Sleep -Seconds 2

    # the record is in the file that survived
    $n = Count-TrueRows $ledger
    Check $b 'persisted' ($n -eq 1) ("processed=TRUE rows in the saved ledger: " + $n)
  } finally {
    if ($main -ne [IntPtr]::Zero -and [RdvGuardWin]::IsWindow($main)) {
      [void](Answer-Dialog '処理済みの確認' 'いいえ' 1)
      [void][RdvGuardWin]::PostMessage($main, 0x0010, [IntPtr]0, [IntPtr]0)
      Start-Sleep -Seconds 3
    }
    $psHost = @(Get-CimInstance Win32_Process -Filter "Name = 'powershell.exe'" | Where-Object { $_.CommandLine -like '*RDV_SELF*' })
    foreach ($ph in $psHost) {
      $q = Get-Process -Id $ph.ProcessId -ErrorAction SilentlyContinue
      if ($q) { Say ("  closing my app host " + $ph.ProcessId); $q.Kill() }
    }
    if ($null -ne $proc -and -not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
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
      # the processed cell is the FIRST cell of a row, and both writers are
      # covered: an inline string "TRUE" (the C# ledger) and a boolean cell,
      # with or without the r=/s= attributes Excel adds (the VBA ledger)
      $n += ([regex]::Matches($s, '<row[^>]*><c[^>]*><is><t>TRUE</t>')).Count
      $n += ([regex]::Matches($s, '<row[^>]*><c[^>]*t="b"[^>]*><v>1</v></c>')).Count
      if ($s.Length -gt 64) { $carry = $s.Substring($s.Length - 64) } else { $carry = $s }
    }
    $sr.Close()
    return $n
  } finally { $z.Dispose() }
}

# ============================================================================
# VBA build
# ============================================================================
function Test-Vba {
  $b = 'vba'
  $dir = Join-Path $scratch 'vba'
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
  Copy-Item -Path (Join-Path $Root 'dist\app-vba\*') -Destination $dir -Recurse -Force
  $log = Join-Path $dir 'ReaderDataViewer.log'
  $ledger = Join-Path $dir 'ReaderDataViewer-Ledger.xlsx'
  $enc = [Text.Encoding]::GetEncoding(932)
  if (Test-Path -LiteralPath $log) { Remove-Item -LiteralPath $log -Force }

  $xl = $null
  $wb = $null
  $mine = @()
  $xlHwnd = [IntPtr]::Zero
  try {
    Say 'VBA: starting Excel'
    # identity is settled before anything is done to it (excel_own.ps1):
    # a reused or unidentifiable instance throws here and is never driven
    $rdvOwn = New-OwnedExcel
    $xl = $rdvOwn.App
    $mine = @($rdvOwn.Pid)
    $xl.Visible = $true
    $xl.DisplayAlerts = $false
    for ($i = 0; $i -lt 10; $i++) { try { $xl.AutomationSecurity = 1; break } catch { Start-Sleep -Milliseconds 600 } }
    $wb = $xl.Workbooks.Open((Join-Path $dir 'ReaderDataViewer.xlsm'))
    $xlHwnd = [IntPtr]$xl.Hwnd
    Say ("VBA: excel pid {0}, hwnd {1}" -f ($mine -join ','), $xlHwnd)

    $ready = Wait-Log $log $enc "`tdecision`tready " 0 $ReadyTimeoutSec 250
    if ($null -eq $ready) { Check $b 'ready' $false 'never reached READY'; return }
    Check $b 'ready' $true $ready[1]

    $ws = $wb.Worksheets.Item('UI')
    $ws.Activate()
    [void][RdvGuardWin]::SetForegroundWindow($xlHwnd)
    Start-Sleep -Milliseconds 600

    # The screen is a pseudo-pixel grid, so a control is a NAMED RANGE spanning
    # many narrow cells -- not one fixed address. The app itself only ever asks
    # for the name (modRdv3App.Rdv3HandleClick -> HitsName -> Rdv3UiRangeOf), so
    # this drives it the same way. Fixed addresses were the old layout: after the
    # screen was rebuilt, "C9"/"E9"/"G9" stopped naming the input and the
    # buttons, and clicking them did nothing at all.
    # Resolve a name to a plain ADDRESS STRING and then use $ws.Range(<address>),
    # which is the one form that reliably carries .Value2 / .Left / .Top through
    # PowerShell's COM adapter. Name.RefersTo is a string ("=UI!$FD$25:$GE$33"),
    # so nothing here depends on a parameterised COM member.
    function Addr-Of([string] $nm) {
      try {
        $r = [string]$wb.Names.Item($nm).RefersTo      # =UI!$FD$25:$GE$33
        $r = $r.TrimStart('=')
        if ($r.Contains('!')) { $r = $r.Substring($r.IndexOf('!') + 1) }
        return $r.Replace('$', '')                     # FD25:GE33
      } catch { return $null }
    }
    $inAddr = Addr-Of 'rdvInput'
    if ([string]::IsNullOrEmpty($inAddr)) { Check $b 'ui_names' $false 'rdvInput is not a name in the workbook'; return }
    Say ("VBA: rdvInput = " + $inAddr)
    $rIn = $ws.Range(($inAddr -split ':')[0])          # the merged anchor
    # Write the merged range itself: Excel puts the value on its anchor. Do NOT
    # go through Range.Address / Range.Cells(r,c) / Range.Row here -- those are
    # parameterised members that do not bind through PowerShell's COM adapter
    # (Address comes back as the method signature, and the rest raise 0x800A03EC
    # / DISP_E_TYPEMISMATCH). Read it back, because a write that lands on a
    # merged cell's non-anchor is dropped in silence.
    $rIn.Value2 = $TargetKey1
    $echo = [string]$rIn.Value2
    if ($echo -ne $TargetKey1) {
      Check $b 'input' $false ("the key did not land in rdvInput (read back '" + $echo + "')")
      return
    }
    Check $b 'input' $true ("rdvInput now holds " + $echo)

    $win = $xl.ActiveWindow
    $x0 = $win.PointsToScreenPixelsX(0)
    $y0 = $win.PointsToScreenPixelsY(0)
    $dpi = [RdvGuardWin]::GetDpiForWindow($xlHwnd)
    if ($dpi -eq 0) { $dpi = 96 }
    # Range.Left/Top are POINTS on the sheet, so screen pixels are
    #   origin + points * (zoom/100) * (dpi/72)
    # PointsToScreenPixels is only trustworthy for the origin itself. The ZOOM
    # factor matters: this screen is a pseudo-pixel grid that Excel fits at well
    # under 100% (69% measured here), and the controls sit far to the right --
    # at column GG a missing zoom factor puts the pointer hundreds of pixels off
    # the button, so the click lands on empty grid and nothing happens.
    $zoom = 100.0
    try { $zoom = [double]$win.Zoom } catch { }
    $s = ($zoom / 100.0) * ($dpi / 72.0)
    Say ("VBA: zoom {0}% dpi {1} -> scale {2}" -f $zoom, $dpi, [Math]::Round($s, 4))
    # Press a button by FOLLOWING its hyperlink.
    #
    # Every button cell on this screen is a hyperlink to itself, and the sheet's
    # single click entry point is Worksheet_FollowHyperlink ->
    # modRdv3App.Rdv3HandleClick. Hyperlink.Follow raises that same event with
    # that same Hyperlink as Target, so the product's own handler runs exactly as
    # it does under a person's click. Selecting the cell does NOT press the
    # button -- there is no SelectionChange handler -- and a synthetic mouse
    # click at computed coordinates does not land on it either: the screen is a
    # pseudo-pixel grid Excel fits at 69% here, and points -> screen pixels
    # through PointsToScreenPixels plus zoom plus dpi missed by whole columns
    # (measured: aiming at GG26:GY33 selected HT34).
    function Click-Named([string] $nm) {
      $a = Addr-Of $nm
      if ([string]::IsNullOrEmpty($a)) { throw ("no such name in the workbook: " + $nm) }
      $anchor = ($a -split ':')[0]
      $cell = $ws.Range($anchor)
      if ($cell.Hyperlinks.Count -lt 1) { throw ("no hyperlink on " + $nm + " at " + $anchor) }
      $cell.Hyperlinks.Item(1).Follow()
    }

    $from = (Read-Log $log $enc).Count
    Click-Named 'rdvBtnSearch'
    # Where did the pointer actually land? The app reacts to the SelectionChange
    # of the button's own cell, so if the selection is not on the button the
    # click missed and "no search line" says nothing about the product.
    Start-Sleep -Milliseconds 400
    $landed = ''
    try { $landed = [string]$xl.ActiveCell.Address(0, 0, 1, 0) } catch { try { $landed = [string]$ws.Application.ActiveCell.Address } catch { $landed = '?' } }
    $btnAddr = Addr-Of 'rdvBtnSearch'
    Say ("VBA: click landed on " + $landed + " (rdvBtnSearch = " + $btnAddr + ")")
    $hit = Wait-Log $log $enc ("`tsearch`tkey=" + $TargetKey1 + " ") $from 60 200
    if ($null -eq $hit) {
      Check $b 'search' $false ("no search line; the click selected " + $landed + " and rdvBtnSearch is " + $btnAddr)
      return
    }
    Check $b 'search' $true $hit[1]

    # 処理済み puts a modal confirm up from INSIDE the click handler, and
    # Hyperlink.Follow does not return until that modal is answered. A single
    # thread therefore cannot press the button and then answer its own dialog --
    # it deadlocks (measured: the run sat on 処理済みの確認 for 30 minutes). So
    # the answer is armed in a separate process FIRST, and the button is pressed
    # afterwards. Nothing about the product changes: it still gets a real click
    # on the hyperlink and a real click on はい.
    $from = (Read-Log $log $enc).Count
    $answerer = Arm-Answer '処理済みの確認' 'はい' 90
    Click-Named 'rdvBtnProcessed'
    Start-Sleep -Milliseconds 400
    if ($null -ne $answerer -and -not $answerer.HasExited) { Start-Sleep -Seconds 2 }
    $answered = ($null -ne $answerer -and $answerer.HasExited -and $answerer.ExitCode -eq 0)
    if (-not $answered) { Check $b 'confirm' $false 'the armed answerer never found 処理済みの確認'; return }
    Check $b 'confirm' $true 'answered はい (armed before the click, because Follow blocks on the modal)'

    $started = Wait-Log $log $enc 'dispatched key2=.*exit held' $from 30 100
    if ($null -eq $started) { Check $b 'save_started' $false 'no dispatch line'; return }
    Check $b 'save_started' $true $started[1]

    # a user-style close of the whole application, while the save is in flight
    Start-Sleep -Milliseconds 300
    [void][RdvGuardWin]::PostMessage($xlHwnd, 0x0010, [IntPtr]0, [IntPtr]0)
    Say 'VBA: WM_CLOSE posted while the save was in flight'

    $blocked = $false
    $t0 = Get-Date
    while (((Get-Date) - $t0).TotalSeconds -lt 20) {
      if (Dialog-Present '保存中のため終了できません') { $blocked = $true; break }
      Start-Sleep -Milliseconds 50
    }
    Check $b 'close_refused_dialog' $blocked 'the close was refused with a dialog naming the reason'
    if ($blocked) { [void](Answer-Dialog '保存中のため終了できません' 'OK' 5) }

    $stillOpen = $false
    try { $stillOpen = ($wb.Name -ne '') } catch { $stillOpen = $false }
    Check $b 'window_alive' $stillOpen 'the workbook is still open after the refused close'

    $refused = Wait-Log $log $enc 'close refused: a processed save is still in flight' $from 20 200
    Check $b 'close_refused_log' ($null -ne $refused) $(if ($null -ne $refused) { $refused[1] } else { 'no refusal line' })

    $done = Wait-Log $log $enc "`tprocessed`tkey2=.*value=TRUE" $from $SaveTimeoutSec 250
    Check $b 'save_completed' ($null -ne $done) $(if ($null -ne $done) { $done[1] } else { 'no processed line' })
    $released = Wait-Log $log $enc 'processed save decided .*exit released' $from 30 250
    Check $b 'exit_released' ($null -ne $released) $(if ($null -ne $released) { $released[1] } else { 'no release line' })

    Start-Sleep -Seconds 1
    $wb.Close($false)
    $wb = $null
    $gone = $false
    for ($i = 0; $i -lt 40; $i++) {
      $cnt = -1
      try { $cnt = $xl.Workbooks.Count } catch { $cnt = -1 }
      if ($cnt -le 0) { $gone = $true; break }
      Start-Sleep -Milliseconds 500
    }
    Check $b 'closes_after' $gone 'the close after the save was accepted'
    Start-Sleep -Seconds 3

    $n = Count-TrueRows $ledger
    Check $b 'persisted' ($n -eq 1) ("processed=TRUE rows in the saved ledger: " + $n)
  } finally {
    Say 'VBA: closing my Excel'
    try { if ($null -ne $wb) { $wb.Close($false) } } catch { }
    try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb) } catch { }
    $wb = $null
    if ($null -ne $xl) {
      for ($i = 0; $i -lt 10; $i++) { try { $xl.Quit(); break } catch { Start-Sleep -Milliseconds 700 } }
      try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    for ($i = 0; $i -lt 12; $i++) {
      if (@($mine | Where-Object { $_ } | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue }).Count -eq 0) { break }
      Start-Sleep -Milliseconds 700
    }
    if ($xlHwnd -ne [IntPtr]::Zero -and [RdvGuardWin]::IsWindow($xlHwnd)) {
      [void][RdvGuardWin]::PostMessage($xlHwnd, 0x0010, [IntPtr]0, [IntPtr]0)
      Start-Sleep -Seconds 3
    }
    foreach ($p in @($mine | Where-Object { $_ })) {
      $q = Get-Process -Id $p -ErrorAction SilentlyContinue
      if ($q) { Say ("  killing my excel " + $p + " (last resort)"); $q.Kill() }
    }
  }
}

# each build is isolated: a failure in one must not take the other with it
if ($Build -eq 'both' -or $Build -eq 'csharp') {
  Write-Output ''
  Write-Output '=== C# build'
  try { Test-CSharp } catch { Check 'csharp' 'harness' $false ("harness error: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber + " | " + ($_.ScriptStackTrace -replace "`r?`n", ' <- ')) }
}
if ($Build -eq 'both' -or $Build -eq 'vba') {
  Write-Output ''
  Write-Output '=== VBA build'
  try { Test-Vba } catch { Check 'vba' 'harness' $false ("harness error: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber + " | " + ($_.ScriptStackTrace -replace "`r?`n", ' <- ')) }
}

Write-Output ''
$pass = @($checks | Where-Object { $_.Result -eq 'PASS' }).Count
$fail = @($checks | Where-Object { $_.Result -eq 'FAIL' }).Count
Write-Output ("=== {0} checks: {1} PASS, {2} FAIL   -> {3}" -f $checks.Count, $pass, $fail, $outTsv)
foreach ($c in $checks) { if ($c.Result -eq 'FAIL') { Write-Output ("  FAIL {0}/{1}: {2}" -f $c.Build, $c.Name, $c.Detail) } }
Write-Output ("scratch: {0}" -f $scratch)
if ($fail -gt 0) { exit 1 }
