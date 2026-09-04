# ============================================================================
# build_app.ps1 -- the C# distribution root, from sources.
#
#   dist\app-csharp\ReaderDataViewer.cmd            self-contained
#   dist\app-csharp\settings.json                   the one settings file
#   dist\app-csharp\docs\settings.md                settings reference
#   dist\app-csharp\docs\ui-spec.md                 screen reference
#   dist\app-csharp\output\                          table exports (initially empty)
#   dist\app-csharp\data\table{A,B,C}.csv plus one paired delete-job input
#
# It never touches archive\ and writes only under
# dist\app-csharp.
#
# The same merge code the app runs is compiled and checked against expected.txt
# here, but no ledger xlsx is shipped. On first launch the app asks before it
# creates the ledger from the CSVs.
#
#   powershell -File build\build_app.ps1
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
  Write-Output ("=== {0} {1}" -f $t, ('=' * [Math]::Max(4, 60 - $t.Length)))
}

# --- data: the verified 1k set, generated if absent or stale ----------------
$dataSrc = Join-Path $Root 'data-1k'
Head 'data (1,000 rows x 3 tables -- gen_data2.ps1, the verified set)'
& (Join-Path $Root 'build\gen_data2.ps1')

# --- preflight ---------------------------------------------------------------
# .ps1 files of the practical build that contain non-ASCII must carry a UTF-8
# BOM (Windows PowerShell 5.1 reads BOM-less files in the ANSI code page).
Head 'preflight: ps1 encoding'
foreach ($ps in @('build\build_app.ps1', 'build\pack_app.ps1', 'build\build_dist.ps1', 'build\sources.ps1',
                  'build\compile_check.ps1', 'build\gen_data2.ps1', 'build\test_exit_guard.ps1',
                  'build\test_settings_geometry.ps1',
                  'build\test_settings_contract.ps1', 'build\gen_samples.ps1', 'build\test_samples.ps1',
                  'src\launcher\boot-app.ps1')) {
  $p = Join-Path $Root $ps
  if (-not (Test-Path -LiteralPath $p)) { continue }
  $bytes = [IO.File]::ReadAllBytes($p)
  $hasBom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)
  $nonAscii = $false
  foreach ($b in $bytes) { if ($b -gt 127) { $nonAscii = $true; break } }
  if ($nonAscii -and -not $hasBom) { throw "$ps has non-ASCII bytes but no UTF-8 BOM" }
  Write-Output ("  {0}: {1}" -f $ps, $(if ($nonAscii) { 'non-ASCII, BOM ok' } else { 'ASCII' }))
}

function Copy-Data([string] $destRoot) {
  $d = Join-Path $destRoot 'data'
  if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
  foreach ($f in 'tableA.csv', 'tableB.csv', 'tableC.csv', 'delete.csv') {
    Copy-Item -LiteralPath (Join-Path $dataSrc $f) -Destination (Join-Path $d $f) -Force
  }
  $legacyDelete = Join-Path $d 'delete-ref.csv'
  if (Test-Path -LiteralPath $legacyDelete) { Remove-Item -LiteralPath $legacyDelete -Force }
  Write-Output ("  data: 4 CSVs -> {0}" -f $d)
}

# --- C# product --------------------------------------------------------------
Head 'C# practical build -> dist\app-csharp'
$destC = Join-Path $Root 'dist\app-csharp'
if (-not (Test-Path -LiteralPath $destC)) { New-Item -ItemType Directory -Path $destC | Out-Null }
$outputDst = Join-Path $destC 'output'
if (-not (Test-Path -LiteralPath $outputDst)) { New-Item -ItemType Directory -Path $outputDst | Out-Null }

& (Join-Path $Root 'build\pack_app.ps1') -Root $Root
Copy-Data $destC

# The entry point people double-click is the .vbs (no console window); the
# .cmd is packed next to it for when a console is wanted.
$cfgSrc = Join-Path $Root 'src\config\settings.json'
$cfgDst = Join-Path $destC 'settings.json'
Copy-Item -LiteralPath $cfgSrc -Destination $cfgDst -Force
Write-Output ('  settings: ' + $cfgDst)

$docsDst = Join-Path $destC 'docs'
if (-not (Test-Path -LiteralPath $docsDst)) { New-Item -ItemType Directory -Path $docsDst | Out-Null }
foreach ($doc in 'settings.md', 'ui-spec.md') {
  Copy-Item -LiteralPath (Join-Path (Join-Path $Root 'docs') $doc) -Destination (Join-Path $docsDst $doc) -Force
}
Write-Output ('  references: settings.md, ui-spec.md -> ' + $docsDst)

Write-Output '  sample verification: compiling the app sources and merging...'
$srcDir = Join-Path $Root 'src\csharp'
# the product sources, as the packer compiles them (the settings loader pulls
# in the UI Automation names, so the whole list is needed)
. (Join-Path $Root 'build\sources.ps1')
$sources = $RdvSources
$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($f in $sources) {
  $text = [IO.File]::ReadAllText((Join-Path $srcDir $f), [Text.Encoding]::UTF8)
  foreach ($line in ($text -split "`r?`n")) {
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $k = $line.Trim()
      if (-not $usings.Contains($k)) { $usings.Add($k, $true) }
    } else { [void]$bodies.AppendLine($line) }
  }
}
$cs = New-Object System.Text.StringBuilder
foreach ($u in $usings.Keys) { [void]$cs.AppendLine($u) }
[void]$cs.AppendLine()
[void]$cs.Append($bodies.ToString())
$esc = New-Object System.Text.StringBuilder
foreach ($ch in $cs.ToString().ToCharArray()) {
  if ([int]$ch -gt 127) { [void]$esc.Append('\u' + ('{0:x4}' -f [int]$ch)) }
  else { [void]$esc.Append($ch) }
}
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.Xml
$refs = @(
  [System.Diagnostics.Process].Assembly.Location,
  [System.Windows.Forms.Form].Assembly.Location,
  [System.Drawing.Point].Assembly.Location,
  [System.Windows.Automation.AutomationElement].Assembly.Location,
  [System.Windows.Automation.AutomationElementIdentifiers].Assembly.Location,
  [System.Windows.DependencyObject].Assembly.Location,
  [System.IO.Compression.ZipArchive].Assembly.Location,
  [System.Xml.XmlReader].Assembly.Location
)
Add-Type -TypeDefinition $esc.ToString() -ReferencedAssemblies $refs -Language CSharp

# the shipped settings.json, read the way the app reads it (strictly): a
# definition that would not start the app does not build a ledger either
$cfg = [Rdv3Config]::Load($cfgSrc)
$mr = [Rdv3Ledger]::BuildFromCsv($cfg.Data, (Join-Path $destC 'data'))
# the data generator's own figures (data-1k\expected.txt: table rows,
# ledger.rows and joinchecksum) have to agree with the merge, or the ledger is
# not built
$expected = @{}
$expectedFile = Join-Path $dataSrc 'expected.txt'
if (Test-Path -LiteralPath $expectedFile) {
  foreach ($line in [IO.File]::ReadAllLines($expectedFile)) {
    $eq = $line.IndexOf('=')
    if ($eq -gt 0) { $expected[$line.Substring(0, $eq)] = $line.Substring($eq + 1) }
  }
}
foreach ($key in 'rows', 'tableA.rows', 'tableB.rows', 'tableC.rows', 'delete.rows', 'delete.columns', 'ledger.rows', 'joinchecksum') {
  if (-not $expected.ContainsKey($key)) { throw ("expected.txt lacks " + $key) }
}
foreach ($table in 'tableA', 'tableB', 'tableC') {
  $actualRows = [IO.File]::ReadAllLines((Join-Path $dataSrc ($table + '.csv'))).Length - 1
  if ($actualRows -ne [int]$expected[$table + '.rows']) {
    throw ("{0}.csv rows {1} do not match expected.txt {2}" -f $table, $actualRows, $expected[$table + '.rows'])
  }
}
if ([long]$expected['ledger.rows'] -ne $mr.Rows -or [long]$expected['joinchecksum'] -ne $mr.Checksum) {
  throw ("initial merge does not match expected.txt: rows {0}/{1} checksum {2}/{3}" -f $mr.Rows, $expected['ledger.rows'], $mr.Checksum, $expected['joinchecksum'])
}
if ($mr.Rows -ge [long]$expected['rows']) {
  throw ("initial ledger must be smaller than each input table: ledger {0}, tables {1}" -f $mr.Rows, $expected['rows'])
}
$deleteLines = [IO.File]::ReadAllLines((Join-Path $dataSrc 'delete.csv'))
if ($deleteLines.Length -lt 1 -or $deleteLines[0] -ne $expected['delete.columns']) {
  throw ("delete.csv columns do not match expected.txt: {0}/{1}" -f $(if ($deleteLines.Length -gt 0) { $deleteLines[0] } else { '<empty>' }), $expected['delete.columns'])
}
if (($deleteLines.Length - 1) -ne [int]$expected['delete.rows']) {
  throw ("delete.csv rows {0} do not match expected.txt {1}" -f ($deleteLines.Length - 1), $expected['delete.rows'])
}

# Exercise the shipped one-file/two-column intersection as part of the sample
# oracle. Every generated pair names the same ledger record, so the preview
# must delete exactly delete.rows records.
$states = [Rdv3Ledger]::FreshStates($mr.Lines.Length, $cfg.Screen.Work.InitialStored)
$deleteJob = $cfg.Data.JobOf('delete-listed-records')
$deleted = [Rdv3Ledger]::ApplyDelete($cfg.Data, $deleteJob, (Join-Path $destC 'data'), $mr.Lines, $states, $cfg.Screen.Work.InitialStored)
if ($deleted.Deleted -ne [int]$expected['delete.rows'] -or $deleted.Lines.Length -ne ($mr.Rows - [int]$expected['delete.rows'])) {
  throw ("delete preview does not match expected.txt: deleted {0}/{1}, rows left {2}/{3}" -f $deleted.Deleted, $expected['delete.rows'], $deleted.Lines.Length, ($mr.Rows - [int]$expected['delete.rows']))
}

# A previous build may have produced these two now-retired artifacts. A direct
# build_app run must leave the same distribution shape as build_dist.
$ledger = Join-Path $destC 'ReaderDataViewer-Ledger.xlsx'
if (Test-Path -LiteralPath $ledger) { Remove-Item -LiteralPath $ledger -Force }
Write-Output ("  sample verified: ledger {0} rows, checksum {1}; paired delete preview {2} rows" -f $mr.Rows, $mr.Checksum, $deleted.Deleted)
Write-Output '  ledger xlsx: not shipped (the app asks before creating it on first launch)'

Head 'done'
foreach ($d in 'dist\app-csharp') {
  $p = Join-Path $Root $d
  if (Test-Path -LiteralPath $p) {
    Get-ChildItem $p -Recurse -File | ForEach-Object {
      Write-Output ("  {0,-60} {1,12:N0} bytes" -f $_.FullName.Substring($Root.Length + 1), $_.Length)
    }
  }
}
