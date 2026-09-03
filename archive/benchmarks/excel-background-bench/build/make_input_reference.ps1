# ============================================================================
# make_input_reference.ps1 -- REFERENCE implementation of the ZipBench input
#                             generator and expected-result oracle.
#
# Like make_dict_reference.ps1 this is a cross-check, not the real thing: the
# benchmark generates its own input and its own oracle inside Excel
# (modZipRule.GenerateInput / BuildExpected). Running both and diffing the
# files proves the VBA and the C# worker really are fed identical bytes.
#
# Generator (deterministic, no randomness - the same N always yields the same
# file, on any machine, in any language):
#
#   for i = 0 .. N-1
#     if  i mod 997  == 996 -> "9999999"   a code absent from the dictionary
#     elif i mod 1499 == 1498 -> "ABC-DEFG" not a number at all
#     else z = zips[i mod M]   (dictionary order = KEN_ALL order)
#          i mod 5 == 0 -> z                      1000001
#          i mod 5 == 1 -> z[0..2] - z[3..6]      100-0001
#          i mod 5 == 2 -> U+3012 + hyphenated    <postal mark>100-0001
#          i mod 5 == 3 -> fullwidth digits       U+FF11 U+FF10 ...
#          i mod 5 == 4 -> z + " "                1000001 with a trailing space
#
# The five shapes exercise the normaliser; the two escape hatches guarantee
# that "not found" is exercised too, so every method has to agree about it.
# ============================================================================
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][int] $Count,
  [string] $Dict     = "",
  [string] $OutInput = "",
  [string] $OutExpected = ""
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
if ([string]::IsNullOrEmpty($Dict))        { $Dict        = Join-Path $root 'data\zip_dict_reference.csv' }
if ([string]::IsNullOrEmpty($OutInput))    { $OutInput    = Join-Path $root ("work\input_ref_{0}.csv" -f $Count) }
if ([string]::IsNullOrEmpty($OutExpected)) { $OutExpected = Join-Path $root ("work\expected_ref_{0}.csv" -f $Count) }

$POSTAL = [char]0x3012      # postal mark
$FW0    = 0xFF10            # fullwidth digit zero
$NOTFOUND = "$([char]0x8A72)$([char]0x5F53)$([char]0x306A)$([char]0x3057)"   # "gaitou nashi"

# Same normaliser as ZipWorker.ZipRule.NormalizeZip and modZipRule.NormalizeZip:
# fullwidth digits fold to ASCII, a fixed set of separators is dropped, anything
# else makes the whole input invalid, and the result must be exactly 7 digits.
$DROP = [int[]] @(0x2D, 0x2010, 0x2012, 0x2013, 0x2014, 0x2015, 0x2212, 0xFF0D,
                  0x30FC, 0x3012, 0x20, 0x3000, 0x09, 0x0D, 0x0A)
function Normalize-Zip([string] $s) {
  if ($null -eq $s) { return '' }
  $out = New-Object System.Text.StringBuilder
  foreach ($ch in $s.ToCharArray()) {
    $c = [int]$ch
    if ($c -ge 0xFF10 -and $c -le 0xFF19) { $c = 48 + ($c - 0xFF10) }
    if ($c -ge 48 -and $c -le 57) {
      if ($out.Length -ge 8) { return '' }
      [void]$out.Append([char]$c)
      continue
    }
    if ($DROP -contains $c) { continue }
    return ''
  }
  if ($out.Length -ne 7) { return '' }
  return $out.ToString()
}

# ---- load the dictionary, keeping KEN_ALL order ----
$sw = [Diagnostics.Stopwatch]::StartNew()
$dictLines = [IO.File]::ReadAllLines($Dict, [Text.UTF8Encoding]::new($false))
$zips = New-Object 'System.Collections.Generic.List[string]'
$map  = New-Object 'System.Collections.Generic.Dictionary[string,string]'
foreach ($l in $dictLines) {
  if ($l.Length -eq 0) { continue }
  $c = $l.IndexOf(',')
  if ($c -le 0) { continue }
  $k = $l.Substring(0, $c)
  if (-not $map.ContainsKey($k)) { $map.Add($k, $l.Substring($c + 1)); [void]$zips.Add($k) }
}
$M = $zips.Count
if ($M -eq 0) { throw "empty dictionary: $Dict" }

$sbIn  = New-Object System.Text.StringBuilder
$sbExp = New-Object System.Text.StringBuilder

for ($i = 0; $i -lt $Count; $i++) {
  if (($i % 997) -eq 996) {
    $s = '9999999'
  } elseif (($i % 1499) -eq 1498) {
    $s = 'ABC-DEFG'
  } else {
    $z = $zips[$i % $M]
    switch ($i % 5) {
      0 { $s = $z }
      1 { $s = $z.Substring(0,3) + '-' + $z.Substring(3) }
      2 { $s = "$POSTAL" + $z.Substring(0,3) + '-' + $z.Substring(3) }
      3 { $t = New-Object System.Text.StringBuilder
          foreach ($ch in $z.ToCharArray()) { [void]$t.Append([char]($FW0 + ([int]$ch - 48))) }
          $s = $t.ToString() }
      4 { $s = $z + ' ' }
    }
  }
  [void]$sbIn.Append($s).Append("`r`n")

  # oracle: normalise then look up, exactly like every engine must
  $key = Normalize-Zip $s
  $addr = $NOTFOUND
  if ($key.Length -eq 7) { $v = $null; if ($map.TryGetValue($key, [ref]$v)) { $addr = $v } }
  [void]$sbExp.Append($addr).Append("`r`n")
}

[IO.File]::WriteAllText($OutInput,    $sbIn.ToString(),  [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText($OutExpected, $sbExp.ToString(), [Text.UTF8Encoding]::new($false))
$sw.Stop()
Write-Output ("dict   : {0} codes" -f $M)
Write-Output ("input  : {0}  ({1} bytes)" -f $OutInput, (Get-Item -LiteralPath $OutInput).Length)
Write-Output ("expect : {0}  ({1} bytes)" -f $OutExpected, (Get-Item -LiteralPath $OutExpected).Length)
Write-Output ("took   : {0} ms" -f $sw.ElapsedMilliseconds)
