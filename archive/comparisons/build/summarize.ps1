# ============================================================================
# summarize.ps1 -- median / min / max per stage from one bench log.
#   pwsh -File build\summarize.ps1 -Log work\bench-csharp-20260813-101500.tsv
# ============================================================================
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string] $Log, [switch] $Markdown)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$rows = Import-Csv -LiteralPath $Log -Delimiter "`t"
if ($rows.Count -eq 0) { Write-Output "empty log"; return }

$cols = @('loadA', 'loadB', 'loadC', 'joinAB', 'joinBC', 'select', 'display', 'other', 'total', 'detect')

function Med([double[]] $v) {
  $s = $v | Sort-Object
  $n = $s.Count
  if ($n -eq 0) { return 0 }
  if ($n % 2 -eq 1) { return $s[[int](($n - 1) / 2)] }
  return ($s[$n / 2 - 1] + $s[$n / 2]) / 2
}

Write-Output ("runs: " + $rows.Count + "   file: " + (Split-Path -Leaf $Log))
$bad = @($rows | Where-Object { $_.oracle -ne 'ok' -or $_.error -ne '' })
if ($bad.Count -gt 0) { Write-Output ("*** " + $bad.Count + " run(s) failed verification") }
else { Write-Output ("verification: all " + $rows.Count + " runs ok (checksum " + $rows[0].checksum + ", probes " + $rows[0].probes + ")") }

if ($Markdown) {
  Write-Output ""
  Write-Output "| stage | median ms | min | max |"
  Write-Output "|---|---:|---:|---:|"
}
foreach ($c in $cols) {
  $v = @($rows | ForEach-Object { [double]$_.$c })
  $m = Med $v
  $mn = ($v | Measure-Object -Minimum).Minimum
  $mx = ($v | Measure-Object -Maximum).Maximum
  if ($Markdown) {
    Write-Output ("| {0} | {1:N1} | {2:N1} | {3:N1} |" -f $c, $m, $mn, $mx)
  } else {
    Write-Output ("  {0,-8} median {1,9:N1} ms    min {2,9:N1}    max {3,9:N1}" -f $c, $m, $mn, $mx)
  }
}
