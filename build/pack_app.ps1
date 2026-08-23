# ============================================================================
# pack_app.ps1 -- assemble the practical build's self-contained .cmd from
# src\csharp and src\launcher.
#
#   dist\app-csharp\ReaderDataViewer.cmd
#
# Packing follows three rules:
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

$srcDir = 'src\csharp'
. (Join-Path $Root 'build\sources.ps1')
$sources = $RdvSources
$title = 'Reader Data Viewer -- practical build, C# standard Dictionary'
$usage = 'ReaderDataViewer.cmd [dataDir] [-config <settings.json>] [-ledger <file.xlsx>] [-log <file>]'
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

$psText = [IO.File]::ReadAllText((Join-Path $Root 'src\launcher\boot-app.ps1'), [Text.Encoding]::UTF8)

foreach ($pair in @(@('PS section', $psText), @('C# section', $csText))) {
  foreach ($m in @(('#RDV' + '-PS-BEGIN'), ('#RDV' + '-CS-BEGIN'))) {
    if ($pair[1].Contains($m)) { throw ("{0} contains the marker {1}" -f $pair[0], $m) }
  }
}
foreach ($ch in $psText.ToCharArray()) {
  if ([int]$ch -gt 127) { throw "the PowerShell bootstrap must stay ASCII (found U+{0:X4})" -f [int]$ch }
}

$psBegin = '#RDV' + '-PS-BEGIN'
$csBegin = '#RDV' + '-CS-BEGIN'

# ---- the .cmd: one file, a console window comes with it ---------------------
$head = [IO.File]::ReadAllText((Join-Path $Root 'src\launcher\header.cmd'), [Text.Encoding]::UTF8)
$head = $head.Replace('@@TITLE@@', $title).Replace('@@USAGE@@', $usage)
$all = $head + "`r`n" + $psBegin + "`r`n" + $psText + "`r`n" + $csBegin + "`r`n" + $csText

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

# ---- the .vbs: the same payload, run by a host that owns no console ---------
# VBScript parses the whole file before running anything, so every payload line
# is commented out with a leading quote; the bootstrap strips it back off.
function Rdv-Comment([string] $text) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($line in ($text -split "`r?`n")) {
    [void]$sb.Append("'").Append($line).Append("`r`n")
  }
  return $sb.ToString().TrimEnd("`r", "`n")
}

$vbsHead = [IO.File]::ReadAllText((Join-Path $Root 'src\launcher\header.vbs'), [Text.Encoding]::UTF8)
$vbsHead = $vbsHead.Replace('@@TITLE@@', $title).Replace('@@USAGE@@', ($usage -replace '\.cmd', '.vbs'))
foreach ($ch in $vbsHead.ToCharArray()) {
  if ([int]$ch -gt 127) { throw "the VBScript head must stay ASCII (found U+{0:X4})" -f [int]$ch }
}
$vbsAll = $vbsHead + "`r`n" + "'" + $psBegin + "`r`n" + (Rdv-Comment $psText) + "`r`n" +
          "'" + $csBegin + "`r`n" + (Rdv-Comment $csText) + "`r`n"

$vbsOut = [IO.Path]::ChangeExtension($Out, '.vbs')
[IO.File]::WriteAllText($vbsOut, $vbsAll, (New-Object Text.UTF8Encoding($false)))
$kb2 = [Math]::Round((Get-Item -LiteralPath $vbsOut).Length / 1KB, 1)
Write-Output ("packed {0}  ({1} sources, {2} KB)" -f $vbsOut, $sources.Count, $kb2)
