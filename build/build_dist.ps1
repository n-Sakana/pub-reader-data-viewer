# ============================================================================
# build_dist.ps1 -- the PRODUCT, from the sources in this repository.
# build.bat calls this; it is also fine to run directly.
#
#   powershell -ExecutionPolicy Bypass -File build\build_dist.ps1
#
# It builds the two authorised products and NOTHING else. No comparison build,
# no benchmark, no test fixture, no sample, no comparison: those are not products
# and must never reach dist\ through this path.
#
# What it produces, all of it from the .cs / .bas / .cls / .ps1 sources:
#
#   dist\app-csharp\ReaderDataViewer.vbs            the C# product (entry point)
#   dist\app-csharp\ReaderDataViewer.cmd            the same payload with a console
#   dist\app-csharp\ReaderDataViewer.json           its settings
#   dist\app-csharp\ReaderDataViewer-Ledger.xlsx    its initial ledger
#   dist\app-csharp\data\table{A,B,C}.csv
#   dist\app-vba\ReaderDataViewer.xlsm              the VBA product (FE)
#   dist\app-vba\ReaderDataViewer.json              its settings
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
  [string] $Root = ""
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

$needExcel = $true
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
Write-Output ("  will build     : the product only (app-csharp, app-vba)")

try {
  # --- data: only the set a distributable is actually built from ------------
  Head 'data-100k (generated only if absent)'
  if (Test-Path -LiteralPath (Join-Path $Root 'data-100k\tableA.csv')) {
    Write-Output '  data-100k already present, left as it is'
  } else {
    & (Join-Path $Root 'build\gen_data2.ps1')
  }

  # The product, and only the product. Comparison builds, benchmarks, test
  # fixtures are deliberately NOT called from here: nothing that
  # is not a shipped product may reach dist\ through build.bat.
  Head 'product -> dist\app-csharp + dist\app-vba'
  & (Join-Path $Root 'build\build_app.ps1') -Root $Root
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
# The .vbs is the entry point people actually double-click and the .json is
# half of the distribution, so both are checked. They were missing from this
# list while the list still described the comparison builds it no longer makes.
$app = @('dist\app-csharp\ReaderDataViewer.vbs',
         'dist\app-csharp\ReaderDataViewer.cmd',
         'dist\app-csharp\ReaderDataViewer-Ledger.xlsx',
         'dist\app-vba\ReaderDataViewer.xlsm',
         'dist\app-vba\ReaderDataViewer-Ledger.xlsx',
         'dist\app-vba\ReaderDataViewer-Ledger.state')
# the shipped CSVs and settings are copies, so "current" means "identical to
# the source", not "written after this run started" (Copy-Item keeps the
# source's timestamp, and re-copying an unchanged file changes nothing)
$copies = @()
foreach ($d in 'dist\app-csharp', 'dist\app-vba') {
  $copies += @{ dest = (Join-Path $d 'ReaderDataViewer.json')
                src  = 'src\app\config\ReaderDataViewer.json' }
  foreach ($n in 'tableA.csv', 'tableB.csv', 'tableC.csv') {
    $copies += @{ dest = (Join-Path (Join-Path $d 'data') $n); src = (Join-Path 'data-100k' $n) }
  }
}
$want = @() + $app

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

# --- dist holds the product and nothing else --------------------------------
# Two things end up here that are not products, and BOTH have to go without the
# operator doing anything: artifacts an older version of this script used to
# build (the comparison files, which are no longer made but are
# still on disk from a previous run), and runtime leftovers from the workbook
# self tests (.log, .lock). A build that only stops MAKING them would leave a
# dist\ that still ships them, so they are removed here and the result is
# asserted to be exactly the allowed set.
$allowed = @{}
foreach ($p in $want) { $allowed[$p.ToLowerInvariant()] = $true }

$distRoot = Join-Path $Root 'dist'
$removed = @()
$stuck = @()
if (Test-Path -LiteralPath $distRoot) {
  foreach ($f in @(Get-ChildItem -LiteralPath $distRoot -Recurse -File)) {
    $rel = $f.FullName.Substring($Root.Length + 1)
    if ($allowed.ContainsKey($rel.ToLowerInvariant())) { continue }
    try {
      Remove-Item -LiteralPath $f.FullName -Force
      $removed += $rel
    } catch {
      $stuck += ($rel + '  (could not remove: ' + $_.Exception.Message + ')')
    }
  }
  # directories the removals emptied (an old comparison folder, say)
  foreach ($d in @(Get-ChildItem -LiteralPath $distRoot -Recurse -Directory |
                   Sort-Object { $_.FullName.Length } -Descending)) {
    if (-not @(Get-ChildItem -LiteralPath $d.FullName -Recurse -File)) {
      try { Remove-Item -LiteralPath $d.FullName -Force -Recurse } catch { }
    }
  }
}

Head 'dist holds the product only'
if ($removed.Count -eq 0) {
  Write-Output '  nothing to remove'
} else {
  Write-Output ("  removed {0} file(s) that are not products:" -f $removed.Count)
  foreach ($r in $removed) { Write-Output ('    ' + $r) }
}
foreach ($e in $stuck) { Write-Output ('  STILL THERE  ' + $e) }
$missing += $stuck

# and now say what is actually on disk, so "only the product" is a list, not a claim
$onDisk = @()
if (Test-Path -LiteralPath $distRoot) {
  $onDisk = @(Get-ChildItem -LiteralPath $distRoot -Recurse -File |
              ForEach-Object { $_.FullName.Substring($Root.Length + 1) } | Sort-Object)
}
Write-Output ("  dist now holds {0} file(s):" -f $onDisk.Count)
foreach ($f in $onDisk) { Write-Output ('    ' + $f) }
$extra = @($onDisk | Where-Object { -not $allowed.ContainsKey($_.ToLowerInvariant()) })
Write-Output ("  extra (not a product): {0}" -f $extra.Count)
$missing += $extra

Head 'result'
if ($missing.Count -eq 0) {
  Write-Output ("  {0} artifacts: every built file was rewritten by this run, and every" -f $want.Count)
  Write-Output "  shipped CSV matches its source in data-100k."
  exit 0
}
Write-Output ("  {0} of {1} artifacts are missing or stale:" -f $missing.Count, $want.Count)
foreach ($m in $missing) { Write-Output ("    " + $m) }
exit 1
