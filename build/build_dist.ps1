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
#   dist\app-csharp\docs\settings.md                settings reference
#   dist\app-csharp\docs\ui-spec.md                 screen reference
#   dist\app-csharp\output\                          table exports (preserved on rebuild)
#   dist\app-csharp\data\table{A,B,C}.csv plus one paired delete-job input
# WHAT IT NEEDS, and what it does NOT need
#   needs   Windows PowerShell 5.1 and .NET Framework csc (both in box).
#   does NOT need administrator rights, and never asks for elevation.
#   does NOT need Excel and does not read or write its registry settings.
#   does NOT write to the registry.
#   does NOT change the machine's execution policy: build.bat passes
#           -ExecutionPolicy Bypass, which applies to that one process.
#   does NOT start or touch Excel.
#
# The 1,000-row data set (data-1k\) is generated if absent or stale; the
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
  Head 'data-1k (generated if absent or stale)'
  & (Join-Path $Root 'build\gen_data2.ps1')

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
         'dist\app-csharp\ReaderDataViewer.cmd')
# the shipped CSVs and settings are copies, so "current" means "identical to
# the source", not "written after this run started" (Copy-Item keeps the
# source's timestamp, and re-copying an unchanged file changes nothing)
$copies = @()
foreach ($d in 'dist\app-csharp') {
  $copies += @{ dest = (Join-Path $d 'settings.json')
                src  = 'src\config\settings.json' }
  foreach ($n in 'settings.md', 'ui-spec.md') {
    $copies += @{ dest = (Join-Path (Join-Path $d 'docs') $n); src = (Join-Path 'docs' $n) }
  }
  foreach ($n in 'tableA.csv', 'tableB.csv', 'tableC.csv', 'delete.csv') {
    $copies += @{ dest = (Join-Path (Join-Path $d 'data') $n); src = (Join-Path 'data-1k' $n) }
  }
}
$want = @() + $app
$requiredDirs = @('dist\app-csharp\output')

Head 'produced'
$missing = @()
foreach ($p in $requiredDirs) {
  $d = Join-Path $Root $p
  if (Test-Path -LiteralPath $d -PathType Container) {
    Write-Output ("  {0,-46} directory" -f ($p + '\'))
  } else {
    Write-Output ("  {0,-46} MISSING" -f ($p + '\'))
    $missing += ($p + '\')
  }
}
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

# --- dist holds the product and preserved table exports ---------------------
# Old artifacts and runtime test leftovers (.log, .lock) have to go without the
# operator doing anything. Files under app-csharp\output belong to the user and
# are preserved across builds. Everything else is checked against the allowed
# product set below.
$allowed = @{}
foreach ($p in $want) { $allowed[$p.ToLowerInvariant()] = $true }
$outputPrefix = 'dist\app-csharp\output\'
$outputRoot = Join-Path $Root 'dist\app-csharp\output'

$distRoot = Join-Path $Root 'dist'
$removed = @()
$stuck = @()
if (Test-Path -LiteralPath $distRoot) {
  foreach ($f in @(Get-ChildItem -LiteralPath $distRoot -Recurse -File)) {
    $rel = $f.FullName.Substring($Root.Length + 1)
    if ($allowed.ContainsKey($rel.ToLowerInvariant())) { continue }
    if ($rel.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
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
    if ([string]::Equals($d.FullName, $outputRoot, [StringComparison]::OrdinalIgnoreCase)) { continue }
    if (-not @(Get-ChildItem -LiteralPath $d.FullName -Recurse -File)) {
      try { Remove-Item -LiteralPath $d.FullName -Force -Recurse } catch { }
    }
  }
}

Head 'dist holds the product and preserves output'
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
$preservedOutput = @($extra | Where-Object { $_.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase) })
$extra = @($extra | Where-Object { -not $_.StartsWith($outputPrefix, [StringComparison]::OrdinalIgnoreCase) })
Write-Output ("  preserved output file(s): {0}" -f $preservedOutput.Count)
Write-Output ("  extra (not a product): {0}" -f $extra.Count)
$missing += $extra

Head 'result'
if ($missing.Count -eq 0) {
  Write-Output ("  {0} files and the output directory: every built file was rewritten by this run, and every" -f $want.Count)
  Write-Output "  shipped CSV matches its source in data-1k."
  exit 0
}
Write-Output ("  {0} of {1} artifacts are missing or stale:" -f $missing.Count, $want.Count)
foreach ($m in $missing) { Write-Output ("    " + $m) }
exit 1
