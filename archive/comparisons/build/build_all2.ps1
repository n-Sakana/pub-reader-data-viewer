# ============================================================================
# build_all2.ps1 -- the four one-to-many distributables, from sources.
#
#   dist\ReaderDataViewer-CSharp-Hash.cmd   C# hand-built hash
#   dist\ReaderDataViewer-CSharp-Dict.cmd   C# standard Dictionary
#   dist\ReaderDataViewer-VBA-Hash.xlsm     VBA hand-built hash
#   dist\ReaderDataViewer-VBA-Dict.xlsm     VBA Scripting.Dictionary
#
# It does NOT touch the four 1:1 distributables built by build_all.ps1. Those
# and their sources are frozen; this script only ever writes the four names
# above.
#
#   pwsh -File build\build_all2.ps1
#   pwsh -File build\build_all2.ps1 -SkipData
# ============================================================================
[CmdletBinding()]
param([string] $Root = "", [switch] $SkipData)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

function Head([string] $t) {
  Write-Output ""
  Write-Output ("=== {0} {1}" -f $t, ('=' * [Math]::Max(4, 60 - $t.Length)))
}

if (-not $SkipData) {
  Head 'data (100,000 rows x 3 tables, key 1 one-to-many)'
  & (Join-Path $Root 'build\gen_data2.ps1')
  Head 'data (200 rows, the same shape in miniature, for the behaviour tests)'
  & (Join-Path $Root 'build\gen_data2.ps1') -Rows 200 -OutDir data-tiny
}

Head 'C# hand-built hash -> ReaderDataViewer-CSharp-Hash.cmd'
& (Join-Path $Root 'build\pack_cmd.ps1') -Variant v2hash

Head 'C# standard Dictionary -> ReaderDataViewer-CSharp-Dict.cmd'
& (Join-Path $Root 'build\pack_cmd.ps1') -Variant v2dict

Head 'the two workbooks (Excel is opened and closed by the build)'
& (Join-Path $Root 'build\build_workbooks2.ps1')

Head 'done'
Get-ChildItem (Join-Path $Root 'dist') -Filter 'ReaderDataViewer-*-*.*' |
  Where-Object { $_.Name -match '(Hash|Dict)' } |
  Sort-Object Name |
  Select-Object Name, @{n = 'KB'; e = { [Math]::Round($_.Length / 1KB, 1) } } |
  Format-Table -AutoSize
