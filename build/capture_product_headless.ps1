# Launch the shipped VBS off-screen and capture its real top-level window.
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $OutPath = ""
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($OutPath)) { $OutPath = Join-Path $Root 'work\ui-v3\app-vbs-headless.png' }
$folder = Split-Path -Parent $OutPath
if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }

$captureSource = @"
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class RdvProductShot {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr root, EnumProc p, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowTextW(IntPtr h, StringBuilder b, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr h, out RECT r);
  [DllImport("user32.dll")] static extern uint GetDpiForWindow(IntPtr h);
  [DllImport("user32.dll")] static extern IntPtr GetWindowDpiAwarenessContext(IntPtr h);
  [DllImport("user32.dll")] static extern bool AreDpiAwarenessContextsEqual(IntPtr a, IntPtr b);
  [DllImport("user32.dll")] public static extern IntPtr PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [StructLayout(LayoutKind.Sequential)] struct RECT { public int Left, Top, Right, Bottom; }

  public static IntPtr Find(int processId) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid != (uint)processId) return true;
      if (!IsWindowVisible(h)) return true;
      StringBuilder b = new StringBuilder(256); GetWindowTextW(h, b, b.Capacity);
      if (b.ToString().StartsWith("Reader Data Viewer", StringComparison.Ordinal)) { found = h; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }

  public static IntPtr FindModal(int processId) {
    IntPtr found = IntPtr.Zero;
    EnumWindows(delegate(IntPtr h, IntPtr l) {
      uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid != (uint)processId || !IsWindowVisible(h)) return true;
      StringBuilder b = new StringBuilder(256); GetWindowTextW(h, b, b.Capacity);
      if (!b.ToString().StartsWith("Reader Data Viewer", StringComparison.Ordinal)) { found = h; return false; }
      return true;
    }, IntPtr.Zero);
    return found;
  }

  public static bool ClickYes(IntPtr dialog) {
    bool clicked = false;
    EnumChildWindows(dialog, delegate(IntPtr h, IntPtr l) {
      StringBuilder b = new StringBuilder(64); GetWindowTextW(h, b, b.Capacity);
      if (String.Equals(b.ToString(), "\u306f\u3044", StringComparison.Ordinal)) {
        PostMessage(h, 0x00F5, IntPtr.Zero, IntPtr.Zero);
        clicked = true;
        return false;
      }
      return true;
    }, IntPtr.Zero);
    return clicked;
  }

  public static string Measure(IntPtr h) {
    RECT outer; RECT client;
    GetWindowRect(h, out outer); GetClientRect(h, out client);
    IntPtr awareness = GetWindowDpiAwarenessContext(h);
    string label = AreDpiAwarenessContextsEqual(awareness, new IntPtr(-2)) ? "system-aware"
      : AreDpiAwarenessContextsEqual(awareness, new IntPtr(-1)) ? "unaware"
      : AreDpiAwarenessContextsEqual(awareness, new IntPtr(-3)) ? "per-monitor-aware"
      : AreDpiAwarenessContextsEqual(awareness, new IntPtr(-4)) ? "per-monitor-aware-v2"
      : "other";
    return "outer=" + (outer.Right - outer.Left) + "x" + (outer.Bottom - outer.Top)
      + " client=" + (client.Right - client.Left) + "x" + (client.Bottom - client.Top)
      + " dpi=" + GetDpiForWindow(h) + " awareness=" + label;
  }

}
"@
Add-Type -TypeDefinition $captureSource

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$vbs = Join-Path $Root 'dist\app-csharp\ReaderDataViewer.vbs'
if (-not (Test-Path -LiteralPath $vbs)) { throw "missing product: $vbs" }
$runFolder = Join-Path $folder ("capture-run-$stamp")
New-Item -ItemType Directory -Path $runFolder | Out-Null
$ledger = Join-Path $runFolder 'ReaderDataViewer-Ledger.xlsx'
$log = Join-Path $folder ("capture-vbs-$stamp.log")
if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }

$oldHeadless = $env:RDV_HEADLESS_TEST
$oldHeadlessKey = $env:RDV_HEADLESS_TEST_KEY
$oldCapturePath = $env:RDV_HEADLESS_CAPTURE_PATH
$env:RDV_HEADLESS_TEST = '1'
$env:RDV_HEADLESS_TEST_KEY = $null
$env:RDV_HEADLESS_CAPTURE_PATH = $OutPath
$launcher = $null
$appPid = 0
$hwnd = [IntPtr]::Zero
$measureClock = [Diagnostics.Stopwatch]::StartNew()
$samples = New-Object 'System.Collections.Generic.List[string]'
$lastMetric = ''
$createApproved = $false
try {
  $arguments = '//B //Nologo "' + $vbs + '" -ledger "' + $ledger + '" -log "' + $log + '"'
  $launcher = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\wscript.exe') -ArgumentList $arguments -WindowStyle Hidden -PassThru
} finally {
  $env:RDV_HEADLESS_TEST = $oldHeadless
  $env:RDV_HEADLESS_TEST_KEY = $oldHeadlessKey
  $env:RDV_HEADLESS_CAPTURE_PATH = $oldCapturePath
}

try {
  $deadline = (Get-Date).AddSeconds(240)
  $ready = $false
  while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $log) {
      $lines = @([IO.File]::ReadAllLines($log, (New-Object Text.UTF8Encoding($false))))
      foreach ($line in $lines) {
        if ($appPid -eq 0 -and $line -match "`tboot`tpid=(\d+)") { $appPid = [int]$Matches[1] }
        if ($line -match "`tdecision`tready rows=") { $ready = $true }
      }
      if ($appPid -ne 0 -and -not $createApproved) {
        $createDialog = [RdvProductShot]::FindModal($appPid)
        if ($createDialog -ne [IntPtr]::Zero -and [RdvProductShot]::ClickYes($createDialog)) {
          # Keep the whole interaction off-screen and let the product create
          # its scratch ledger from the shipped CSVs.
          $createApproved = $true
        }
      }
      if ($appPid -ne 0 -and $hwnd -eq [IntPtr]::Zero) { $hwnd = [RdvProductShot]::Find($appPid) }
      if ($hwnd -ne [IntPtr]::Zero) {
        $metric = [RdvProductShot]::Measure($hwnd)
        if ($metric -ne $lastMetric) {
          $samples.Add(("t={0}ms {1}" -f $measureClock.ElapsedMilliseconds, $metric))
          $lastMetric = $metric
        }
      }
      if ($appPid -ne 0 -and $ready) { break }
    }
    Start-Sleep -Milliseconds 100
  }
  if ($appPid -eq 0 -or -not $ready) { throw "the VBS product did not reach READY; log=$log" }
  if (-not $createApproved -or -not (Test-Path -LiteralPath $ledger)) { throw "the first-run ledger was not created; log=$log" }
  $stableUntil = (Get-Date).AddMilliseconds(1500)
  while ((Get-Date) -lt $stableUntil) {
    if ($hwnd -eq [IntPtr]::Zero) { $hwnd = [RdvProductShot]::Find($appPid) }
    if ($hwnd -ne [IntPtr]::Zero) {
      $metric = [RdvProductShot]::Measure($hwnd)
      if ($metric -ne $lastMetric) {
        $samples.Add(("t={0}ms {1}" -f $measureClock.ElapsedMilliseconds, $metric))
        $lastMetric = $metric
      }
    }
    Start-Sleep -Milliseconds 50
  }
  for ($i = 0; $i -lt 100 -and -not (Test-Path -LiteralPath $OutPath); $i++) { Start-Sleep -Milliseconds 50 }
  if (-not (Test-Path -LiteralPath $OutPath)) { throw "the READY form did not write its DrawToBitmap capture; log=$log" }
  for ($i = 0; $i -lt 100 -and $hwnd -eq [IntPtr]::Zero; $i++) {
    $hwnd = [RdvProductShot]::Find($appPid)
    if ($hwnd -eq [IntPtr]::Zero) { Start-Sleep -Milliseconds 50 }
  }
  if ($hwnd -eq [IntPtr]::Zero) { throw "the READY process has no viewer window; pid=$appPid" }
  if ($samples.Count -eq 0) { throw "the viewer window produced no DPI samples; pid=$appPid" }
  foreach ($sample in $samples) { Write-Output ("sample " + $sample) }
  if (@($samples | Where-Object { $_ -notmatch 'dpi=144 awareness=system-aware$' }).Count -ne 0) {
    throw "the viewer was not system-aware at 144 DPI for every visible sample"
  }
  $sizes = @($samples | ForEach-Object { ($_ -replace '^.* outer=', 'outer=') -replace ' dpi=.*$', '' } | Select-Object -Unique)
  if ($sizes.Count -ne 1) { throw ("the visible window changed size: " + ($sizes -join ', ')) }
  Write-Output $OutPath
  Write-Output ("pid={0} hwnd={1} log={2}" -f $appPid, $hwnd, $log)
} finally {
  if ($hwnd -ne [IntPtr]::Zero) { [void][RdvProductShot]::PostMessage($hwnd, 0x0010, [IntPtr]0, [IntPtr]0) }
  if ($appPid -ne 0) {
    for ($i = 0; $i -lt 40; $i++) {
      if (-not (Get-Process -Id $appPid -ErrorAction SilentlyContinue)) { break }
      Start-Sleep -Milliseconds 100
    }
    $owned = Get-Process -Id $appPid -ErrorAction SilentlyContinue
    if ($owned) { Stop-Process -Id $appPid -Force }
  }
}
