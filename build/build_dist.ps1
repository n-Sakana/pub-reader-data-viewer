# ============================================================================
# build_dist.ps1 -- every distributable in dist\, from the sources in this
# repository. build.bat calls this; it is also fine to run directly.
#
#   powershell -ExecutionPolicy Bypass -File build\build_dist.ps1
#   ... -SkipFrozen     only the practical build (dist\app-csharp, dist\app-vba)
#   ... -SkipApp        only the two frozen comparisons
#
# What it produces, all of it from the .cs / .bas / .cls / .ps1 sources:
#
#   dist\ReaderDataViewer-CSharp.cmd        1:1 comparison, C# only
#   dist\ReaderDataViewer-Hybrid.cmd        1:1 comparison, C# engine + Excel UI
#   dist\ReaderDataViewer-VBA.xlsm          1:1 comparison, VBA
#   dist\ReaderDataViewer-Hybrid.xlsm       1:1 comparison, the hybrid's book
#   dist\ReaderDataViewer-CSharp-Hash.cmd   one-to-many, hand-built hash
#   dist\ReaderDataViewer-CSharp-Dict.cmd   one-to-many, standard Dictionary
#   dist\ReaderDataViewer-VBA-Hash.xlsm     one-to-many, hand-built hash
#   dist\ReaderDataViewer-VBA-Dict.xlsm     one-to-many, standard Dictionary
#   dist\app-csharp\ReaderDataViewer.cmd            practical build
#   dist\app-csharp\ReaderDataViewer-Ledger.xlsx    its initial ledger
#   dist\app-csharp\data\table{A,B,C}.csv
#   dist\app-vba\ReaderDataViewer.xlsm              practical build (FE)
#   dist\app-vba\ReaderDataViewer-Ledger.xlsx       the BE-owned ledger
#   dist\app-vba\ReaderDataViewer-Ledger.state      its sidecar mirror
#   dist\app-vba\data\table{A,B,C}.csv
#
# WHAT IT NEEDS, and what it does NOT need
#   needs   Windows PowerShell 5.1 (in box), .NET Framework csc (in box),
#           Excel, and Excel's per-user setting "Trust access to the VBA
#           project object model" -- the only way this repository can put VBA
#           into a workbook is VBProject.VBComponents.Import.
#   does NOT need administrator rights, and never asks for elevation.
#   does NOT write to the registry (it only READS the AccessVBOM value).
#   does NOT change the machine's execution policy: build.bat passes
#           -ExecutionPolicy Bypass, which applies to that one process.
#   does NOT touch any Excel instance it did not start itself.
#
# The 100,000-row data set (data-100k\) is generated only if it is absent; the
# practical build needs it. The 1,000,000-row set (data\, 241 MB) is NOT
# generated: no distributable is built from it -- it is what the 1:1
# comparison READS at run time -- and the note at the end says so.
#
# Exit codes: 0 = every artifact listed above is present, 1 = a step failed
# (the failing step and its message are printed).
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [switch] $SkipFrozen,
  [switch] $SkipApp
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

function Head([string] $t) {
  Write-Output ""
  Write-Output ("=== {0} {1}" -f $t, ('=' * [Math]::Max(4, 58 - $t.Length)))
}

$started = Get-Date

# --- preflight: say up front what is missing, never half-build silently -----
Head 'preflight'
Write-Output ("  repository     : " + $Root)
Write-Output ("  PowerShell     : " + $PSVersionTable.PSVersion.ToString())
Write-Output ("  .NET (this ps) : " + [Environment]::Version.ToString())

$needExcel = -not ($SkipFrozen -and $SkipApp)
if ($needExcel) {
  $excel = $null
  foreach ($k in 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
                 'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe') {
    if (Test-Path $k) { $excel = (Get-ItemProperty $k).'(default)'; break }
  }
  if ($null -eq $excel) {
    throw "Excel が見つかりません (App Paths\excel.exe が登録されていません)。ワークブックのビルドには Excel が要ります。"
  }
  Write-Output ("  Excel          : " + $excel)

  $vbom = 0
  foreach ($v in '16.0', '15.0', '14.0') {
    $p = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
    if (Test-Path $p) { $x = (Get-ItemProperty $p).AccessVBOM; if ($x) { $vbom = $x } }
  }
  if ($vbom -ne 1) {
    throw @"
Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無効です。
このリポジトリが .xlsm に VBA を入れる唯一の経路が VBProject.VBComponents.Import
なので、この設定なしではワークブックを作れません。管理者権限は不要です:

  Excel > ファイル > オプション > トラスト センター > トラスト センターの設定
        > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する

にチェックを入れて Excel を閉じ、build.bat をもう一度実行してください。
(この設定はユーザー単位 HKCU の値です。build.bat はレジストリを書き換えません。)
"@
  }
  Write-Output ("  AccessVBOM     : 1 (trusted)")
}
Write-Output ("  will build     : " +
  $(if ($SkipFrozen) { '' } else { 'frozen 1:1, frozen one-to-many, ' }) +
  $(if ($SkipApp) { '' } else { 'practical (app-csharp, app-vba)' }))

try {
  # --- data: only the set a distributable is actually built from ------------
  Head 'data-100k (generated only if absent)'
  if (Test-Path -LiteralPath (Join-Path $Root 'data-100k\tableA.csv')) {
    Write-Output '  data-100k already present, left as it is'
  } else {
    & (Join-Path $Root 'build\gen_data2.ps1')
  }

  if (-not $SkipFrozen) {
    Head 'frozen 1:1 comparison -> 2 .cmd + 2 .xlsm'
    & (Join-Path $Root 'build\build_all.ps1') -Root $Root -SkipData

    Head 'frozen one-to-many comparison -> 2 .cmd + 2 .xlsm'
    & (Join-Path $Root 'build\build_all2.ps1') -Root $Root -SkipData
  }

  if (-not $SkipApp) {
    Head 'practical build -> dist\app-csharp + dist\app-vba'
    & (Join-Path $Root 'build\build_app.ps1') -Root $Root
  }
}
catch {
  Write-Output ""
  Write-Output "=== FAILED ================================================="
  Write-Output ("  " + $_.Exception.Message)
  if ($_.InvocationInfo) {
    Write-Output ("  at " + $_.InvocationInfo.ScriptName + " line " + $_.InvocationInfo.ScriptLineNumber)
    if ($_.InvocationInfo.Line) { Write-Output ("     " + $_.InvocationInfo.Line.Trim()) }
  }
  exit 1
}

# --- what is on disk now, checked against what was supposed to be built -----
$frozen = @('dist\ReaderDataViewer-CSharp.cmd',
            'dist\ReaderDataViewer-Hybrid.cmd',
            'dist\ReaderDataViewer-VBA.xlsm',
            'dist\ReaderDataViewer-Hybrid.xlsm',
            'dist\ReaderDataViewer-CSharp-Hash.cmd',
            'dist\ReaderDataViewer-CSharp-Dict.cmd',
            'dist\ReaderDataViewer-VBA-Hash.xlsm',
            'dist\ReaderDataViewer-VBA-Dict.xlsm')
$app = @('dist\app-csharp\ReaderDataViewer.cmd',
         'dist\app-csharp\ReaderDataViewer-Ledger.xlsx',
         'dist\app-vba\ReaderDataViewer.xlsm',
         'dist\app-vba\ReaderDataViewer-Ledger.xlsx',
         'dist\app-vba\ReaderDataViewer-Ledger.state')
# the shipped CSVs are copies of data-100k\, so "current" means "identical to
# the source", not "written after this run started" (Copy-Item keeps the
# source's timestamp, and re-copying an unchanged file changes nothing)
$copies = @()
if (-not $SkipApp) {
  foreach ($d in 'dist\app-csharp\data', 'dist\app-vba\data') {
    foreach ($n in 'tableA.csv', 'tableB.csv', 'tableC.csv') {
      $copies += @{ dest = (Join-Path $d $n); src = (Join-Path 'data-100k' $n) }
    }
  }
}
$want = @()
if (-not $SkipFrozen) { $want += $frozen }
if (-not $SkipApp) { $want += $app }

Head 'produced'
$missing = @()
foreach ($p in $want) {
  $f = Join-Path $Root $p
  if (Test-Path -LiteralPath $f) {
    $i = Get-Item -LiteralPath $f
    $fresh = if ($i.LastWriteTime -ge $started) { 'built now' } else { 'NOT REWRITTEN' }
    Write-Output ("  {0,-46} {1,12:N0} bytes  {2}  {3}" -f $p, $i.Length, $i.LastWriteTime.ToString('MM-dd HH:mm'), $fresh)
    if ($i.LastWriteTime -lt $started) { $missing += ($p + ' (not rewritten by this run)') }
  } else {
    Write-Output ("  {0,-46} MISSING" -f $p)
    $missing += $p
  }
}
foreach ($c in $copies) {
  $f = Join-Path $Root $c.dest
  $s = Join-Path $Root $c.src
  if (-not (Test-Path -LiteralPath $f)) {
    Write-Output ("  {0,-46} MISSING" -f $c.dest)
    $missing += $c.dest
    continue
  }
  $fi = Get-Item -LiteralPath $f
  $si = Get-Item -LiteralPath $s
  $same = ($fi.Length -eq $si.Length -and $fi.LastWriteTime -eq $si.LastWriteTime)
  Write-Output ("  {0,-46} {1,12:N0} bytes  {2}  {3}" -f $c.dest, $fi.Length,
                $fi.LastWriteTime.ToString('MM-dd HH:mm'),
                $(if ($same) { 'copy of ' + $c.src } else { 'DIFFERS FROM ' + $c.src }))
  if (-not $same) { $missing += ($c.dest + ' (differs from ' + $c.src + ')') }
}
$want += ($copies | ForEach-Object { $_.dest })

Head 'note'
Write-Output '  data\ (1,000,000 rows x 3 tables, ~241 MB) is NOT built here: no'
Write-Output '  distributable is made from it. It is what the 1:1 comparison READS at'
Write-Output '  run time. Generate it when you want to run those four:'
Write-Output '      powershell -ExecutionPolicy Bypass -File build\gen_data.ps1'

Head 'result'
if ($missing.Count -eq 0) {
  Write-Output ("  {0} artifacts: every built file was rewritten by this run, and every" -f $want.Count)
  Write-Output "  shipped CSV matches its source in data-100k."
  exit 0
}
Write-Output ("  {0} of {1} artifacts are missing or stale:" -f $missing.Count, $want.Count)
foreach ($m in $missing) { Write-Output ("    " + $m) }
exit 1
