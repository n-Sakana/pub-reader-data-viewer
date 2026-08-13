# ============================================================================
# make_dict_reference.ps1 -- REFERENCE implementation of the ZipBench dictionary
#                            build, used only to cross-check the VBA one.
#
# The dictionary that the benchmark actually uses is built by Excel
# (modZipRule.BuildDictionary). This script implements the same rule
# independently so the two can be diffed byte for byte; if they ever disagree,
# the rule has drifted and the benchmark is no longer comparing like with like.
#
# ZipBench rule v1 (town normalisation), applied once here and once in VBA:
#   1. town = KEN_ALL field 9 (kanji town)
#   2. cut the town at the first fullwidth open paren U+FF08, if any
#   3. if what remains ends with U+5834 U+5408 ("...baai"), the town is one of
#      Japan Post's marker rows, not a real place name -> drop it entirely
#      (this covers "ika ni keisai ga nai baai" and every "... no tsugi ni
#      banchi ga kuru baai" variant)
#   4. address = prefecture + city + town
#   5. the FIRST row wins when a postal code appears more than once
# ============================================================================
[CmdletBinding()]
param(
  [string] $KenAll = "",
  [string] $Out    = ""
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
if ([string]::IsNullOrEmpty($KenAll)) { $KenAll = Join-Path $root 'data\KEN_ALL.CSV' }
if ([string]::IsNullOrEmpty($Out))    { $Out    = Join-Path $root 'data\zip_dict_reference.csv' }

$sw = [Diagnostics.Stopwatch]::StartNew()
$lines = [IO.File]::ReadAllLines($KenAll, [Text.Encoding]::GetEncoding(932))

$PAREN = [char]0xFF08          # fullwidth (
$BA    = [char]0x5834          # kanji "ba"
$AI    = [char]0x5408          # kanji "ai"
$suffix = "$BA$AI"

$seen = New-Object 'System.Collections.Generic.HashSet[string]'
$sb   = New-Object System.Text.StringBuilder
$dropMarker = 0
$dropParen  = 0

foreach ($line in $lines) {
  if ($line.Length -eq 0) { continue }
  $f = $line -split ','
  if ($f.Count -lt 9) { continue }
  $zip  = $f[2].Trim('"')
  if (-not $seen.Add($zip)) { continue }
  $pref = $f[6].Trim('"')
  $city = $f[7].Trim('"')
  $town = $f[8].Trim('"')
  $p = $town.IndexOf($PAREN)
  if ($p -ge 0) { $town = $town.Substring(0, $p); $dropParen++ }
  if ($town.EndsWith($suffix)) { $town = ''; $dropMarker++ }
  [void]$sb.Append($zip).Append(',').Append($pref).Append($city).Append($town).Append("`r`n")
}

[IO.File]::WriteAllText($Out, $sb.ToString(), [Text.UTF8Encoding]::new($false))
$sw.Stop()

Write-Output ("read      : {0} rows from {1}" -f $lines.Count, $KenAll)
Write-Output ("unique    : {0} postal codes" -f $seen.Count)
Write-Output ("paren cut : {0} rows" -f $dropParen)
Write-Output ("marker    : {0} rows had the whole town dropped" -f $dropMarker)
Write-Output ("wrote     : {0}  ({1} bytes, {2} ms)" -f $Out, (Get-Item -LiteralPath $Out).Length, $sw.ElapsedMilliseconds)
