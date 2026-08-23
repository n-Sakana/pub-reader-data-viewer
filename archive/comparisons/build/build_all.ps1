# ============================================================================
# build_all.ps1 -- the frozen 1:1 comparison. NOT a product and NOT called by
# build.bat: it exists so docs\results.md stays reproducible.
#
#   data\tableA.csv  tableB.csv  tableC.csv  expected.txt   (generated, ~241 MB)
#   dist\ReaderDataViewer-VBA.xlsm
#   dist\ReaderDataViewer-CSharp.cmd
#
# None of those are in git: they are build output and the data is 241 MB.
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

Section "1/3  synthetic data"
if ($SkipData) {
  Write-Output "skipped (-SkipData)"
} else {
  & (Join-Path $Root 'build\gen_data.ps1') -Rows $Rows -Force:$Force
}

Section "2/3  dist\ReaderDataViewer-CSharp.cmd"
& (Join-Path $Root 'build\pack_cmd.ps1') -Variant csharp -Root $Root

Section "3/3  dist\ReaderDataViewer-VBA.xlsm"
& (Join-Path $Root 'build\build_workbooks.ps1') -Root $Root

Section "done"
Get-ChildItem (Join-Path $Root 'dist') | Select-Object Name, @{n = 'KB'; e = { [Math]::Round($_.Length / 1KB, 1) } } |
  Format-Table -AutoSize | Out-String | Write-Output
