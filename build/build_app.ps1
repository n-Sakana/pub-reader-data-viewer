# ============================================================================
# build_app.ps1 -- the C# distribution root, from sources.
#
#   dist\app-csharp\ReaderDataViewer.cmd            self-contained
#   dist\app-csharp\settings.json                   the one settings file
#   dist\app-csharp\ReaderDataViewer-Ledger.xlsx    initial ledger (all FALSE)
#   dist\app-csharp\data\table{A,B,C}.csv
#
# It never touches archive\, benchmarks\ or showcase\ and writes only under
# dist\app-csharp.
#
# The initial xlsx ledger is produced by the SAME code the app runs (the
# Rdv3 sources are compiled here and Rdv3Ledger/Rdv3Xlsx are called directly),
# so what ships is what the app would itself have written.
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

# --- data: the verified 100k set, generated only if absent ------------------
$dataSrc = Join-Path $Root 'data-100k'
if (-not (Test-Path -LiteralPath (Join-Path $dataSrc 'tableA.csv'))) {
  Head 'data (100,000 rows x 3 tables -- gen_data2.ps1, the verified set)'
  & (Join-Path $Root 'build\gen_data2.ps1')
}

# --- preflight ---------------------------------------------------------------
# .ps1 files of the practical build that contain non-ASCII must carry a UTF-8
# BOM (Windows PowerShell 5.1 reads BOM-less files in the ANSI code page).
Head 'preflight: ps1 encoding'
foreach ($ps in @('build\build_app.ps1', 'build\pack_app.ps1', 'build\build_dist.ps1', 'build\sources.ps1',
                  'build\compile_check.ps1', 'build\gen_data2.ps1', 'build\test_exit_guard.ps1',
                  'build\test_ui_geometry.ps1', 'build\test_settings_geometry.ps1',
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
  foreach ($f in 'tableA.csv', 'tableB.csv', 'tableC.csv') {
    Copy-Item -LiteralPath (Join-Path $dataSrc $f) -Destination (Join-Path $d $f) -Force
  }
  Write-Output ("  data: 3 CSVs -> {0}" -f $d)
}

# --- C# product --------------------------------------------------------------
Head 'C# practical build -> dist\app-csharp'
$destC = Join-Path $Root 'dist\app-csharp'
if (-not (Test-Path -LiteralPath $destC)) { New-Item -ItemType Directory -Path $destC | Out-Null }

& (Join-Path $Root 'build\pack_app.ps1') -Root $Root
Copy-Data $destC

# The entry point people double-click is the .vbs (no console window); the
# .cmd is packed next to it for when a console is wanted.
$cfgSrc = Join-Path $Root 'src\config\settings.json'
$cfgDst = Join-Path $destC 'settings.json'
Copy-Item -LiteralPath $cfgSrc -Destination $cfgDst -Force
Write-Output ('  settings: ' + $cfgDst)

Write-Output '  initial ledger: compiling the app sources and merging...'
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
$exp = [Rdv3Expected]::Read($dataSrc)
if ($exp.Loaded -and ($exp.Rows -ne $mr.Rows -or $exp.Checksum -ne $mr.Checksum)) {
  throw ("initial merge does not match the oracle: rows {0}/{1} checksum {2}/{3}" -f $mr.Rows, $exp.Rows, $mr.Checksum, $exp.Checksum)
}
# every record starts in the definition's initial work state, stored under
# the column heading the definition names
$states = [Rdv3Ledger]::FreshStates($mr.Lines.Length, $cfg.Screen.Work.InitialStored)
$ledger = Join-Path $destC 'ReaderDataViewer-Ledger.xlsx'
if (Test-Path -LiteralPath $ledger) { Remove-Item -LiteralPath $ledger -Force }
[Rdv3Xlsx]::Write($ledger, $mr.Head, $cfg.Screen.Work.Column, $mr.Lines, $states, 'build')
Write-Output ("  initial ledger: {0} rows, checksum {1} (oracle ok) -> {2}" -f $mr.Rows, $mr.Checksum, $ledger)

Head 'done'
foreach ($d in 'dist\app-csharp') {
  $p = Join-Path $Root $d
  if (Test-Path -LiteralPath $p) {
    Get-ChildItem $p -Recurse -File | ForEach-Object {
      Write-Output ("  {0,-60} {1,12:N0} bytes" -f $_.FullName.Substring($Root.Length + 1), $_.Length)
    }
  }
}
