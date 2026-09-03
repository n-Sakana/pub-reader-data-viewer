# ============================================================================
# build_dist.ps1 -- the PRODUCT, from the sources in this repository.
# build.bat calls this; it is also fine to run directly.
#
#   powershell -ExecutionPolicy Bypass -File build\build_dist.ps1
#
# It builds the C# product and NOTHING else. No archived build, benchmark,
# test fixture, sample or showcase may reach dist\ through this path.
#
# What it produces, all of it from the active C# / PowerShell sources:
#
#   dist\app-csharp\ReaderDataViewer.vbs            the C# product (entry point)
#   dist\app-csharp\ReaderDataViewer.cmd            the same payload with a console
#   dist\app-csharp\settings.json                   its settings (paths, watch, data, screen)
#   dist\app-csharp\ReaderDataViewer-Ledger.xlsx    its initial ledger
#   dist\app-csharp\data\table{A,B,C}.csv plus the two delete-job inputs
# WHAT IT NEEDS, and what it does NOT need
#   needs   Windows PowerShell 5.1 and .NET Framework csc (both in box).
#   does NOT need administrator rights, and never asks for elevation.
#   does NOT need Excel and does not read or write its registry settings.
#   does NOT write to the registry.
#   does NOT change the machine's execution policy: build.bat passes
#           -ExecutionPolicy Bypass, which applies to that one process.
#   does NOT start or touch Excel.
#
# The 100,000-row data set (data-100k\) is generated only if it is absent; the
# practical build needs it. The archived comparison data set is not generated.
#
# Exit codes: 0 = every artifact listed above is present, 1 = a step failed
# (the failing step and its message are printed).
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = ""
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

function Head([string] $t) {
  Write-Output ""
  Write-Output ("=== {0} {1}" -f $t, ('=' * [Math]::Max(4, 58 - $t.Length)))
}

$started = Get-Date

# --- preflight ---------------------------------------------------------------
Head 'preflight'
Write-Output ("  repository     : " + $Root)
Write-Output ("  PowerShell     : " + $PSVersionTable.PSVersion.ToString())
Write-Output ("  .NET (this ps) : " + [Environment]::Version.ToString())

Write-Output ("  will build     : app-csharp only")

try {
  # --- data: only the set a distributable is actually built from ------------
  Head 'data-100k (generated only if absent)'
  if ((Test-Path -LiteralPath (Join-Path $Root 'data-100k\tableA.csv')) -and
      (Test-Path -LiteralPath (Join-Path $Root 'data-100k\delete.csv')) -and
      (Test-Path -LiteralPath (Join-Path $Root 'data-100k\delete-ref.csv'))) {
    Write-Output '  data-100k already present, left as it is'
  } else {
    & (Join-Path $Root 'build\gen_data2.ps1')
  }

  # The product, and only the product. Comparison builds, benchmarks, test
  # fixtures are deliberately NOT called from here: nothing that
  # is not a shipped product may reach dist\ through build.bat.
  Head 'product -> dist\app-csharp'
  & (Join-Path $Root 'build\build_app.ps1') -Root $Root
}
catch {
  Write-Output ""
  Write-Output "=== FAILED ================================================="
  Write-Output ("  " + $_.Exception.Message)
  if ($_.InvocationInfo) {
    Write-Output ("  at " + $_.InvocationInfo.ScriptName + " line " + $_.InvocationInfo.ScriptLineNumber)
    if ($_.InvocationInfo.Line) { Write-Output ("     " + $_.InvocationInfo.Line.Trim()) }
  }
  exit 1
}

# --- what is on disk now, checked against what was supposed to be built -----
# The .vbs is the entry point people actually double-click and the .json is
# half of the distribution, so both are checked. They were missing from this
# list while the list still described the comparison builds it no longer makes.
$app = @('dist\app-csharp\ReaderDataViewer.vbs',
         'dist\app-csharp\ReaderDataViewer.cmd',
         'dist\app-csharp\ReaderDataViewer-Ledger.xlsx')
# the shipped CSVs and settings are copies, so "current" means "identical to
# the source", not "written after this run started" (Copy-Item keeps the
# source's timestamp, and re-copying an unchanged file changes nothing)
$copies = @()
foreach ($d in 'dist\app-csharp') {
  $copies += @{ dest = (Join-Path $d 'settings.json')
                src  = 'src\config\settings.json' }
  foreach ($n in 'tableA.csv', 'tableB.csv', 'tableC.csv', 'delete.csv', 'delete-ref.csv') {
    $copies += @{ dest = (Join-Path (Join-Path $d 'data') $n); src = (Join-Path 'data-100k' $n) }
  }
}
$want = @() + $app

Head 'produced'
$missing = @()
foreach ($p in $want) {
  $f = Join-Path $Root $p
  if (Test-Path -LiteralPath $f) {
    $i = Get-Item -LiteralPath $f
    $fresh = if ($i.LastWriteTime -ge $started) { 'built now' } else { 'NOT REWRITTEN' }
    Write-Output ("  {0,-46} {1,12:N0} bytes  {2}  {3}" -f $p, $i.Length, $i.LastWriteTime.ToString('MM-dd HH:mm'), $fresh)
    if ($i.LastWriteTime -lt $started) { $missing += ($p + ' (not rewritten by this run)') }
  } else {
    Write-Output ("  {0,-46} MISSING" -f $p)
    $missing += $p
  }
}
foreach ($c in $copies) {
  $f = Join-Path $Root $c.dest
  $s = Join-Path $Root $c.src
  if (-not (Test-Path -LiteralPath $f)) {
    Write-Output ("  {0,-46} MISSING" -f $c.dest)
    $missing += $c.dest
    continue
  }
  $fi = Get-Item -LiteralPath $f
  $si = Get-Item -LiteralPath $s
  $same = ($fi.Length -eq $si.Length -and $fi.LastWriteTime -eq $si.LastWriteTime)
  Write-Output ("  {0,-46} {1,12:N0} bytes  {2}  {3}" -f $c.dest, $fi.Length,
                $fi.LastWriteTime.ToString('MM-dd HH:mm'),
                $(if ($same) { 'copy of ' + $c.src } else { 'DIFFERS FROM ' + $c.src }))
  if (-not $same) { $missing += ($c.dest + ' (differs from ' + $c.src + ')') }
}
$want += ($copies | ForEach-Object { $_.dest })

# --- dist holds the product and nothing else --------------------------------
# Two things end up here that are not products, and BOTH have to go without the
# operator doing anything: artifacts an older version of this script used to
# build (the comparison files, which are no longer made but are
# still on disk from a previous run), and runtime leftovers from the workbook
# self tests (.log, .lock). A build that only stops MAKING them would leave a
# dist\ that still ships them, so they are removed here and the result is
# asserted to be exactly the allowed set.
$allowed = @{}
foreach ($p in $want) { $allowed[$p.ToLowerInvariant()] = $true }

$distRoot = Join-Path $Root 'dist'
$removed = @()
$stuck = @()
if (Test-Path -LiteralPath $distRoot) {
  foreach ($f in @(Get-ChildItem -LiteralPath $distRoot -Recurse -File)) {
    $rel = $f.FullName.Substring($Root.Length + 1)
    if ($allowed.ContainsKey($rel.ToLowerInvariant())) { continue }
    try {
      Remove-Item -LiteralPath $f.FullName -Force
      $removed += $rel
    } catch {
      $stuck += ($rel + '  (could not remove: ' + $_.Exception.Message + ')')
    }
  }
  # directories the removals emptied (an old comparison folder, say)
  foreach ($d in @(Get-ChildItem -LiteralPath $distRoot -Recurse -Directory |
                   Sort-Object { $_.FullName.Length } -Descending)) {
    if (-not @(Get-ChildItem -LiteralPath $d.FullName -Recurse -File)) {
      try { Remove-Item -LiteralPath $d.FullName -Force -Recurse } catch { }
    }
  }
}

Head 'dist holds the product only'
if ($removed.Count -eq 0) {
  Write-Output '  nothing to remove'
} else {
  Write-Output ("  removed {0} file(s) that are not products:" -f $removed.Count)
  foreach ($r in $removed) { Write-Output ('    ' + $r) }
}
foreach ($e in $stuck) { Write-Output ('  STILL THERE  ' + $e) }
$missing += $stuck

# and now say what is actually on disk, so "only the product" is a list, not a claim
$onDisk = @()
if (Test-Path -LiteralPath $distRoot) {
  $onDisk = @(Get-ChildItem -LiteralPath $distRoot -Recurse -File |
              ForEach-Object { $_.FullName.Substring($Root.Length + 1) } | Sort-Object)
}
Write-Output ("  dist now holds {0} file(s):" -f $onDisk.Count)
foreach ($f in $onDisk) { Write-Output ('    ' + $f) }
$extra = @($onDisk | Where-Object { -not $allowed.ContainsKey($_.ToLowerInvariant()) })
Write-Output ("  extra (not a product): {0}" -f $extra.Count)
$missing += $extra

Head 'result'
if ($missing.Count -eq 0) {
  Write-Output ("  {0} artifacts: every built file was rewritten by this run, and every" -f $want.Count)
  Write-Output "  shipped CSV matches its source in data-100k."
  exit 0
}
Write-Output ("  {0} of {1} artifacts are missing or stale:" -f $missing.Count, $want.Count)
foreach ($m in $missing) { Write-Output ("    " + $m) }
exit 1
