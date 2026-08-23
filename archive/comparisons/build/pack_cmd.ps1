# ============================================================================
# pack_cmd.ps1 -- assemble one self-contained .cmd from src\csharp and src\cmd.
#
#   dist\ReaderDataViewer-CSharp.cmd   = engine + WinForms front
#
# Three things it has to get right:
#
#  1. usings. The .cs files stay separately compilable and readable, so each
#     one carries its own using block. C# wants them first, so they are hoisted
#     and de-duplicated here.
#
#  2. non-ASCII. Add-Type writes the source to a temp file and the in-box csc
#     reads a BOM-less file in the system ANSI code page, so a UTF-8 Japanese
#     literal arrives as mojibake (measured: "監視中" -> 3 CP932 pairs). Every
#     character above U+007F is therefore rewritten as \uXXXX, which csc turns
#     back into the right character whatever the code page. A verbatim string
#     would keep the backslash instead, so @"..." is rejected outright.
#
#  3. markers. The batch header finds the two sections by searching its own
#     text, so neither section may contain the marker strings.
#
#   pwsh -File build\pack_cmd.ps1 -Variant csharp    src\csharp     (1:1, frozen)
#   pwsh -File build\pack_cmd.ps1 -Variant v2hash    src\v2\csharp  (one-to-many)
#   pwsh -File build\pack_cmd.ps1 -Variant v2dict    src\v2\csharp  (one-to-many)
# ============================================================================
[CmdletBinding()]
param(
  [ValidateSet('csharp', 'v2hash', 'v2dict')] [string] $Variant = 'csharp',
  [string] $Root = "",
  [string] $Out = ""
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

# src\csharp is the frozen 1:1 build. src\v2\csharp is the one-to-many build,
# whose two variants differ by exactly one file: the index.
$srcDir = 'src\csharp'
$cmdDir = 'src\cmd'
$bootCommon = 'boot-common.ps1'
$common = @('RdvCore.cs', 'RdvWatch.cs', 'RdvText.cs', 'RdvRunner.cs')
if ($Variant -eq 'csharp') {
  $sources = $common + @('RdvUiForms.cs', 'RdvApp.cs')
  $title = 'Reader Data Viewer -- method 2: C# only (UI Automation + WinForms)'
  $usage = 'ReaderDataViewer-CSharp.cmd [dataDir] [-log <file>]'
  $name = 'ReaderDataViewer-CSharp.cmd'
  $boot = 'boot-csharp.ps1'
} else {
  $srcDir = 'src\v2\csharp'
  $cmdDir = 'src\v2\cmd'
  $bootCommon = 'boot-common2.ps1'
  if ($Variant -eq 'v2hash') {
    $idx = 'Rdv2IdxHash.cs'
    $title = 'Reader Data Viewer -- one-to-many, C# hand-built hash index'
    $name = 'ReaderDataViewer-CSharp-Hash.cmd'
    $boot = 'boot-hash.ps1'
  } else {
    $idx = 'Rdv2IdxDict.cs'
    $title = 'Reader Data Viewer -- one-to-many, C# standard Dictionary index'
    $name = 'ReaderDataViewer-CSharp-Dict.cmd'
    $boot = 'boot-dict.ps1'
  }
  $sources = @('Rdv2Core.cs', $idx, 'Rdv2Watch.cs', 'Rdv2Text.cs', 'Rdv2Ui.cs', 'Rdv2App.cs')
  $usage = "$name [dataDir] [-log <file>] [-autopick <n>]"
}

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

$psText = [IO.File]::ReadAllText((Join-Path $Root "$cmdDir\$bootCommon"), [Text.Encoding]::UTF8) + "`r`n" +
          [IO.File]::ReadAllText((Join-Path $Root "$cmdDir\$boot"), [Text.Encoding]::UTF8)

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
  $Out = Join-Path $Root "dist\$name"
}
$dir = Split-Path -Parent $Out
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir | Out-Null }

# UTF-8 without BOM: cmd.exe reads the head, PowerShell reads the whole file as
# UTF-8, and nothing in the head is outside ASCII anyway
[IO.File]::WriteAllText($Out, $all, (New-Object Text.UTF8Encoding($false)))

$kb = [Math]::Round((Get-Item -LiteralPath $Out).Length / 1KB, 1)
Write-Output ("packed {0}  ({1} sources, {2} KB)" -f $Out, $sources.Count, $kb)
