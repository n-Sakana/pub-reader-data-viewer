# ============================================================================
# summarize2.ps1 -- median / min / max per stage from one one-to-many bench log.
#
# "total" is the measured region: confirmed number -> record on screen, or ->
# candidate list built. It never contains the time a person spent choosing.
# "pick" is what happened after the choice, and is reported on its own.
#
#   pwsh -File build\summarize2.ps1 -Log work\bench2-chash-20260814-021500.tsv
#   pwsh -File build\summarize2.ps1 -Log ... -Markdown
# ============================================================================
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string] $Log, [switch] $Markdown)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$rows = Import-Csv -LiteralPath $Log -Delimiter "`t"
if ($rows.Count -eq 0) { Write-Output "empty log"; return }

$cols = @('readA', 'readB', 'readC', 'idxA', 'idxB', 'idxC', 'joinAB', 'joinBC', 'search', 'present', 'other', 'total')

function Med([double[]] $v) {
  $s = $v | Sort-Object
  $n = $s.Count
  if ($n -eq 0) { return 0 }
  if ($n % 2 -eq 1) { return $s[[int](($n - 1) / 2)] }
  return ($s[$n / 2 - 1] + $s[$n / 2]) / 2
}

Write-Output ("runs: " + $rows.Count + "   index: " + $rows[0].index + "   file: " + (Split-Path -Leaf $Log))
$bad = @($rows | Where-Object { $_.oracle -ne 'ok' -or $_.error -ne '' })
if ($bad.Count -gt 0) {
  Write-Output ("*** " + $bad.Count + " run(s) failed verification")
  foreach ($b in $bad) { Write-Output ("    seq " + $b.seq + " key " + $b.key + " oracle=" + $b.oracle + " error=" + $b.error) }
} else {
  Write-Output ("verification: all " + $rows.Count + " runs ok (checksum " + $rows[0].checksum +
                ", rows " + $rows[0].rows + ", matchedAB " + $rows[0].matchedAB + ", matchedBC " + $rows[0].matchedBC + ")")
}

# candidate counts per key: the four methods must agree on these
$byKey = $rows | Group-Object key | Sort-Object Name
$parts = @()
foreach ($g in $byKey) {
  $counts = @($g.Group | ForEach-Object { $_.cands } | Sort-Object -Unique)
  $parts += ("{0}={1}" -f $g.Name, ($counts -join '/'))
}
Write-Output ("candidates by key: " + ($parts -join '  '))

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

# pick and detect are outside the measured region and are reported apart from it
$pick = @($rows | Where-Object { [double]$_.pick -ge 0 } | ForEach-Object { [double]$_.pick })
if ($pick.Count -gt 0) {
  Write-Output ("  {0,-8} median {1,9:N1} ms    min {2,9:N1}    max {3,9:N1}   ({4} runs with a dialog)" -f 'pick',
    (Med $pick), ($pick | Measure-Object -Minimum).Minimum, ($pick | Measure-Object -Maximum).Maximum, $pick.Count)
} else {
  Write-Output "  pick     -- no run had more than one candidate"
}
$det = @($rows | ForEach-Object { [double]$_.detect })
Write-Output ("  {0,-8} median {1,9:N1} ms    min {2,9:N1}    max {3,9:N1}   (reference, outside the measured region)" -f 'detect',
  (Med $det), ($det | Measure-Object -Minimum).Minimum, ($det | Measure-Object -Maximum).Maximum)
