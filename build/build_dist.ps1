# ============================================================================
# build_dist.ps1 -- the PRODUCT, as a folder that can be copied to a machine
# and double-clicked.
#
#   powershell -ExecutionPolicy Bypass -File build\build_dist.ps1
#
# The WebView2 version has no packing step: ReaderDataViewer.ps1 compiles
# src\*.cs with the in-box csc at launch, and the window needs the WebView2
# DLLs in lib\ and the page in web\ on disk. A distribution is therefore the
# runnable tree, copied and then verified by running it.
#
# It is written to build\out\ReaderDataViewer:
#
#   ReaderDataViewer.vbs   entry point (no console)
#   ReaderDataViewer.cmd   the same, with a console
#   settings.json          data, processing and screen
#   src\*.cs               the sources it compiles
#   src\ReaderDataViewer.ps1  launcher / compiler
#   web\                   index.html, app.js, app.css
#   lib\                   WebView2 DLLs and notices
#   data\                  the 1,000-row sample tables
#   output\                CSV exports (kept on rebuild)
#   LICENSE, THIRD-PARTY-NOTICES.md, README.md
#
# It builds the product and NOTHING else: no archived build, benchmark, test
# fixture or sample reaches dist\ through this path. No ledger xlsx is
# shipped; the app asks before creating one on first launch.
#
#   needs   Windows PowerShell 5.1 and the in-box .NET Framework csc.
#   does NOT need administrator rights, Excel, or the registry.
#
# Exit codes: 0 = every artifact listed above is present and verified,
#             1 = a step failed (the failing step and its message are printed).
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
$dest = Join-Path $Root 'build\out\ReaderDataViewer'
$dataSrc = Join-Path $Root 'data'

# What is copied, and from where. Everything else in the repository stays out.
$files = @(
  'ReaderDataViewer.vbs', 'ReaderDataViewer.cmd',
  'settings.json', 'LICENSE', 'THIRD-PARTY-NOTICES.md', 'README.md')
$trees = @(
  @{ from = 'src'; filter = '*.cs' },
  @{ from = 'src'; filter = '*.ps1' },
  @{ from = 'web'; filter = '*' },
  @{ from = 'lib'; filter = '*' })
$dataFiles = @('tableA.csv', 'tableB.csv', 'tableC.csv', 'delete.csv')

try {
  Head 'preflight'
  Write-Output ("  repository     : " + $Root)
  Write-Output ("  PowerShell     : " + $PSVersionTable.PSVersion.ToString())
  Write-Output ("  .NET (this ps) : " + [Environment]::Version.ToString())
  Write-Output ("  destination    : " + $dest)

  # A .ps1 with non-ASCII bytes must carry a UTF-8 BOM: Windows PowerShell 5.1
  # reads a BOM-less file in the ANSI code page. The .vbs and the C# sources
  # must be ASCII outright -- the in-box csc reads a BOM-less temp file in that
  # same code page, and the .vbs is read by wscript.
  $asciiOnly = @('ReaderDataViewer.vbs') + @(
    Get-ChildItem -LiteralPath (Join-Path $Root 'src') -Filter '*.cs' -File |
      ForEach-Object { 'src\' + $_.Name })
  $bomIfNonAscii = @(
    Get-ChildItem -LiteralPath (Join-Path $Root 'build') -Filter '*.ps1' -File |
      ForEach-Object { 'build\' + $_.Name }) + @('src\ReaderDataViewer.ps1')
  foreach ($rel in $asciiOnly) {
    $bytes = [IO.File]::ReadAllBytes((Join-Path $Root $rel))
    foreach ($b in $bytes) { if ($b -gt 127) { throw "$rel must be ASCII only" } }
  }
  foreach ($rel in $bomIfNonAscii) {
    $p = Join-Path $Root $rel
    if (-not (Test-Path -LiteralPath $p)) { continue }
    $bytes = [IO.File]::ReadAllBytes($p)
    $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
    $nonAscii = $false
    foreach ($b in $bytes) { if ($b -gt 127) { $nonAscii = $true; break } }
    if ($nonAscii -and -not $hasBom) { throw "$rel has non-ASCII bytes but no UTF-8 BOM" }
  }
  Write-Output ("  encoding       : {0} ASCII-only files, {1} .ps1 checked for a BOM" -f $asciiOnly.Count, $bomIfNonAscii.Count)

  # --- data: the verified 1,000-row set, generated if absent ----------------
  Head 'data\ (the 1,000-row set, generated if absent)'
  & (Join-Path $Root 'build\gen_data2.ps1')

  # --- the tree -------------------------------------------------------------
  # Everything but output\ is replaced, so a source file deleted since the last
  # build cannot survive in the distribution.
  Head 'copy -> build\out\ReaderDataViewer'
  if (Test-Path -LiteralPath $dest) {
    foreach ($item in Get-ChildItem -LiteralPath $dest -Force) {
      if ($item.Name -eq 'output') { continue }
      Remove-Item -LiteralPath $item.FullName -Recurse -Force
    }
  } else {
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
  }
  foreach ($name in $files) {
    Copy-Item -LiteralPath (Join-Path $Root $name) -Destination (Join-Path $dest $name) -Force
  }
  Write-Output ("  files : {0}" -f ($files -join ', '))
  foreach ($tree in $trees) {
    $d = Join-Path $dest $tree.from
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
    $copied = @(Get-ChildItem -LiteralPath (Join-Path $Root $tree.from) -Filter $tree.filter -File)
    foreach ($f in $copied) { Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $d $f.Name) -Force }
    Write-Output ("  {0,-6}: {1} files" -f $tree.from, $copied.Count)
  }
  $dataDst = Join-Path $dest 'data'
  New-Item -ItemType Directory -Path $dataDst | Out-Null
  foreach ($f in $dataFiles) {
    Copy-Item -LiteralPath (Join-Path $dataSrc $f) -Destination (Join-Path $dataDst $f) -Force
  }
  Write-Output ('  data  : {0} CSVs from data\' -f $dataFiles.Count)
  $outputDst = Join-Path $dest 'output'
  if (-not (Test-Path -LiteralPath $outputDst)) { New-Item -ItemType Directory -Path $outputDst | Out-Null }

  # --- verification: the copy, compiled and merged --------------------------
  # The distribution is checked, not the repository: this compiles the sources
  # that were copied, loads the settings.json that was copied, and merges the
  # CSVs that were copied.
  Head 'verify the copy'
  . (Join-Path $Root 'build\test_support.ps1')
  Import-RdvProduct -Root $dest

  $cfg = [Rdv3Config]::Load((Join-Path $dest 'settings.json'))
  $mr = [Rdv3Ledger]::BuildFromCsv($cfg.Data, $dataDst)

  $expected = @{}
  foreach ($line in [IO.File]::ReadAllLines((Join-Path $dataSrc 'expected.txt'))) {
    $eq = $line.IndexOf('=')
    if ($eq -gt 0) { $expected[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
  }
  foreach ($key in 'rows', 'tableA.rows', 'tableB.rows', 'tableC.rows', 'delete.rows',
                   'delete.columns', 'ledger.rows', 'joinchecksum') {
    if (-not $expected.ContainsKey($key)) { throw ("expected.txt lacks " + $key) }
  }
  foreach ($table in 'tableA', 'tableB', 'tableC') {
    $actualRows = [IO.File]::ReadAllLines((Join-Path $dataDst ($table + '.csv'))).Length - 1
    if ($actualRows -ne [int]$expected[$table + '.rows']) {
      throw ("{0}.csv rows {1} do not match expected.txt {2}" -f $table, $actualRows, $expected[$table + '.rows'])
    }
  }
  if ([long]$expected['ledger.rows'] -ne $mr.Rows -or [long]$expected['joinchecksum'] -ne $mr.Checksum) {
    throw ("merge does not match expected.txt: rows {0}/{1} checksum {2}/{3}" -f
      $mr.Rows, $expected['ledger.rows'], $mr.Checksum, $expected['joinchecksum'])
  }
  if ($mr.Rows -ge [long]$expected['rows']) {
    throw ("the ledger must be smaller than each input table: ledger {0}, tables {1}" -f $mr.Rows, $expected['rows'])
  }
  $deleteLines = [IO.File]::ReadAllLines((Join-Path $dataDst 'delete.csv'))
  $deleteHead = '<empty>'
  if ($deleteLines.Length -gt 0) { $deleteHead = $deleteLines[0] }
  if ($deleteLines.Length -lt 1 -or $deleteHead -ne $expected['delete.columns']) {
    throw ("delete.csv columns do not match expected.txt: {0}/{1}" -f $deleteHead, $expected['delete.columns'])
  }
  if (($deleteLines.Length - 1) -ne [int]$expected['delete.rows']) {
    throw ("delete.csv rows {0} do not match expected.txt {1}" -f ($deleteLines.Length - 1), $expected['delete.rows'])
  }

  # The shipped paired delete job, run against the shipped data.
  $states = [Rdv3Ledger]::FreshStates($mr.Lines.Length, $cfg.Screen.Work.InitialStored)
  $deleteJob = $cfg.Data.JobOf('delete-listed-records')
  $deleted = [Rdv3Ledger]::ApplyDelete($cfg.Data, $deleteJob, $dataDst, $mr.Lines, $states, $cfg.Screen.Work.InitialStored)
  if ($deleted.Deleted -ne [int]$expected['delete.rows'] -or
      $deleted.Lines.Length -ne ($mr.Rows - [int]$expected['delete.rows'])) {
    throw ("delete preview does not match expected.txt: deleted {0}/{1}, rows left {2}/{3}" -f
      $deleted.Deleted, $expected['delete.rows'], $deleted.Lines.Length, ($mr.Rows - [int]$expected['delete.rows']))
  }
  Write-Output ("  ledger {0} rows, checksum {1}; paired delete preview {2} rows" -f $mr.Rows, $mr.Checksum, $deleted.Deleted)
  Write-Output ("  screen: " + $cfg.Screen.Describe())

  # No ledger workbook ships: the app asks before it creates one.
  $ledger = Join-Path $dest $cfg.Ledger
  if (Test-Path -LiteralPath $ledger) { Remove-Item -LiteralPath $ledger -Force }
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

# --- what is on disk now, against what was supposed to be built -------------
Head 'produced'
$missing = @()
$total = 0
foreach ($rel in $files) {
  $p = Join-Path $dest $rel
  if (Test-Path -LiteralPath $p) {
    $i = Get-Item -LiteralPath $p
    $total += $i.Length
    Write-Output ("  {0,-40} {1,12:N0} bytes" -f $rel, $i.Length)
  } else {
    Write-Output ("  {0,-40} MISSING" -f $rel); $missing += $rel
  }
}
foreach ($dir in @('src', 'web', 'lib', 'data', 'output')) {
  $p = Join-Path $dest $dir
  if (Test-Path -LiteralPath $p -PathType Container) {
    $inDir = @(Get-ChildItem -LiteralPath $p -File)
    $bytes = 0
    foreach ($f in $inDir) { $bytes += $f.Length }
    $total += $bytes
    Write-Output ("  {0,-40} {1,12:N0} bytes  ({2} files)" -f ($dir + '\'), $bytes, $inDir.Count)
  } else {
    Write-Output ("  {0,-40} MISSING" -f ($dir + '\')); $missing += ($dir + '\')
  }
}
# The three files without which the window cannot open at all.
foreach ($needed in 'web\index.html', 'web\app.js', 'web\app.css',
                    'lib\Microsoft.Web.WebView2.Core.dll',
                    'lib\Microsoft.Web.WebView2.Wpf.dll',
                    'lib\WebView2Loader.dll') {
  if (-not (Test-Path -LiteralPath (Join-Path $dest $needed))) { $missing += $needed }
}

Write-Output ""
if ($missing.Count -gt 0) {
  Write-Output ("INCOMPLETE: " + ($missing -join ', '))
  exit 1
}
Write-Output ("dist\ReaderDataViewer is complete and verified ({0:N0} bytes, {1:F1} s)" -f
  $total, ((Get-Date) - $started).TotalSeconds)
