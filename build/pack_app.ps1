# ============================================================================
# pack_app.ps1 -- assemble the practical build's self-contained .cmd from
# src\app\csharp and src\app\cmd.
#
#   dist\app-csharp\ReaderDataViewer.cmd
#
# Same three rules as build\pack_cmd.ps1 (which packs the frozen comparison
# builds and is not touched by this script):
#
#  1. usings are hoisted to the top and de-duplicated
#  2. every character above U+007F is rewritten as \uXXXX, because Add-Type
#     hands the source to the in-box csc which reads a BOM-less temp file in
#     the system ANSI code page; verbatim strings are rejected outright
#  3. the batch header finds its sections by marker strings, so neither
#     section may contain them
#
#   powershell -File build\pack_app.ps1
#   powershell -File build\pack_app.ps1 -Out <path>
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $Out = ""
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$srcDir = 'src\app\csharp'
$sources = @('Rdv3Core.cs', 'Rdv3Index.cs', 'Rdv3Ledger.cs', 'Rdv3Xlsx.cs', 'Rdv3Jobs.cs',
             'Rdv3Watch.cs', 'Rdv3Text.cs', 'Rdv3Ui.cs', 'Rdv3App.cs')
$title = 'Reader Data Viewer -- practical build, C# standard Dictionary'
$usage = 'ReaderDataViewer.cmd [dataDir] [-ledger <file.xlsx>] [-log <file>]'
$name = 'ReaderDataViewer.cmd'

$usings = New-Object System.Collections.Specialized.OrderedDictionary
$bodies = New-Object System.Text.StringBuilder
foreach ($f in $sources) {
  $p = Join-Path $Root "$srcDir\$f"
  if (-not (Test-Path -LiteralPath $p)) { throw "missing source: $p" }
  $text = [IO.File]::ReadAllText($p, [Text.Encoding]::UTF8)
  if ($text.Contains('@"')) { throw "$f contains a verbatim string; the \u escaping below would break it" }
  foreach ($line in ($text -split "`r?`n")) {
    if ($line -match '^\s*using\s+[A-Za-z_][A-Za-z0-9_.]*\s*;\s*$') {
      $k = $line.Trim()
      if (-not $usings.Contains($k)) { $usings.Add($k, $true) }
    } else {
      [void]$bodies.AppendLine($line)
    }
  }
  [void]$bodies.AppendLine()
}

$cs = New-Object System.Text.StringBuilder
foreach ($u in $usings.Keys) { [void]$cs.AppendLine($u) }
[void]$cs.AppendLine()
[void]$cs.Append($bodies.ToString())

# escape everything above ASCII
$esc = New-Object System.Text.StringBuilder
foreach ($ch in $cs.ToString().ToCharArray()) {
  if ([int]$ch -gt 127) { [void]$esc.Append('\u' + ('{0:x4}' -f [int]$ch)) }
  else { [void]$esc.Append($ch) }
}
$csText = $esc.ToString()

$psText = [IO.File]::ReadAllText((Join-Path $Root 'src\app\cmd\boot-app.ps1'), [Text.Encoding]::UTF8)

foreach ($pair in @(@('PS section', $psText), @('C# section', $csText))) {
  foreach ($m in @(('#RDV' + '-PS-BEGIN'), ('#RDV' + '-CS-BEGIN'))) {
    if ($pair[1].Contains($m)) { throw ("{0} contains the marker {1}" -f $pair[0], $m) }
  }
}
foreach ($ch in $psText.ToCharArray()) {
  if ([int]$ch -gt 127) { throw "the PowerShell bootstrap must stay ASCII (found U+{0:X4})" -f [int]$ch }
}

$head = [IO.File]::ReadAllText((Join-Path $Root 'src\cmd\header.cmd'), [Text.Encoding]::UTF8)
$head = $head.Replace('@@TITLE@@', $title).Replace('@@USAGE@@', $usage)

$all = $head + "`r`n" + ('#RDV' + '-PS-BEGIN') + "`r`n" + $psText + "`r`n" +
       ('#RDV' + '-CS-BEGIN') + "`r`n" + $csText

if ([string]::IsNullOrEmpty($Out)) {
  $Out = Join-Path $Root "dist\app-csharp\$name"
}
$dir = Split-Path -Parent $Out
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

# UTF-8 without BOM: cmd.exe reads the head, PowerShell reads the whole file as
# UTF-8, and nothing in the head is outside ASCII anyway
[IO.File]::WriteAllText($Out, $all, (New-Object Text.UTF8Encoding($false)))

$kb = [Math]::Round((Get-Item -LiteralPath $Out).Length / 1KB, 1)
Write-Output ("packed {0}  ({1} sources, {2} KB)" -f $Out, $sources.Count, $kb)
