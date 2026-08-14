# ============================================================================
# build_all.ps1 -- everything, from a clean checkout to the four files in dist\.
#
#   data\tableA.csv  tableB.csv  tableC.csv  expected.txt   (generated, ~241 MB)
#   dist\ReaderDataViewer-VBA.xlsm
#   dist\ReaderDataViewer-CSharp.cmd
#   dist\ReaderDataViewer-Hybrid.cmd
#   dist\ReaderDataViewer-Hybrid.xlsm
#
# None of those are in git: all four are build output and the data is 241 MB.
# This script is the whole recipe.
#
#   powershell -ExecutionPolicy Bypass -File build\build_all.ps1
#   powershell -ExecutionPolicy Bypass -File build\build_all.ps1 -SkipData
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [int]    $Rows = 1000000,
  [switch] $SkipData,
  [switch] $Force
)
$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

function Section([string] $s) {
  Write-Output ""
  Write-Output ("=== " + $s + " " + ("=" * [Math]::Max(0, 60 - $s.Length)))
}

Section "1/4  synthetic data"
if ($SkipData) {
  Write-Output "skipped (-SkipData)"
} else {
  & (Join-Path $Root 'build\gen_data.ps1') -Rows $Rows -Force:$Force
}

Section "2/4  dist\ReaderDataViewer-CSharp.cmd"
& (Join-Path $Root 'build\pack_cmd.ps1') -Variant csharp -Root $Root

Section "3/4  dist\ReaderDataViewer-Hybrid.cmd"
& (Join-Path $Root 'build\pack_cmd.ps1') -Variant hybrid -Root $Root

Section "4/4  dist\ReaderDataViewer-VBA.xlsm + dist\ReaderDataViewer-Hybrid.xlsm"
& (Join-Path $Root 'build\build_workbooks.ps1') -Root $Root

Section "done"
Get-ChildItem (Join-Path $Root 'dist') | Select-Object Name, @{n = 'KB'; e = { [Math]::Round($_.Length / 1KB, 1) } } |
  Format-Table -AutoSize | Out-String | Write-Output
