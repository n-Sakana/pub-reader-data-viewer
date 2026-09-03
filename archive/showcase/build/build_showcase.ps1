# build_showcase.ps1
#
# 成果物をつくる。出来るのは 2 つだけ。
#   dist\<component>_<N>px.bas / .cls   ANSI(CP932) に直したソース一式
#   dist\VBA Pixel Bridge <N>px.xlsm    それを取り込み、UIAutomationClient の
#                                       参照を付けたブック。ふつうに開けば動く。
#
# 製品は責務ごとに複数のモジュール／クラスへ分かれている。ビルドは src\ にある
# .bas と .cls をすべて取り込む。取り込む順序は問わない（VBA は参照を解決して
# からコンパイルする）が、ここでは名前順にして出力を安定させる。
#
# ここでやる COM は全部この子プロセス側（build_child.ps1）に置いてある。VBA の
# コンパイルエラーは「見えないモーダル」になり得るので、ブックは不可視のまま
# つくり、保存前に PbPing を実際に呼んで公開入口を確かめる。返ってこなければ、
# このスクリプトが自分の持ち物だけを落とす。
[CmdletBinding()]
param(
    # 1 セルが受け持つ設計 px。利用者が選ぶもので、既定値は置かない。
    # 既定値を置くと「選ばなかった」と「4 を選んだ」が区別できなくなる。
    [Parameter(Mandatory=$true)][ValidateSet('1','2','4')][string] $Unit,
    [int] $TimeoutSec = 240
)
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$srcDir = Join-Path $root 'src'
$dist = Join-Path $root 'dist'
$child= Join-Path $here 'build_child.ps1'
if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

$sources = @(Get-ChildItem $srcDir -File | Where-Object { $_.Extension -in '.bas', '.cls' } |
             Sort-Object Name)
if ($sources.Count -eq 0) { throw "src にソースがありません: $srcDir" }

# .bas / .cls は取り込まれるときシステムの ANSI コードページで読まれる（日本語
# Windows なら CP932）。ソースは UTF-8 で持ち、成果物は CP932 で出す。
$cp932 = [System.Text.Encoding]::GetEncoding(932)
$decl  = 'Public Const PB_UNIT As Long ='
$note  = "              ' 1 セルが何設計 px 分か（ビルド時に選ぶ: 1 / 2 / 4）"
$cr = [char]13
$lf = [char]10
$unitHits = 0
$outFiles = @()

foreach ($s in $sources) {
    $text = [System.IO.File]::ReadAllText($s.FullName, (New-Object System.Text.UTF8Encoding($false)))

    # .cls の先頭のクラスヘッダは CRLF で書かれていないと取り込み側が
    # ヘッダとして認識せず、そのままコード 1 行目に流れ込む。取り込みは
    # 成功して見えるので、あとから「1 行目が VERSION 1.0 CLASS」という
    # コンパイル失敗になるだけになる。ここで必ず CRLF へそろえる。
    $text = $text.Replace("$cr$lf", "$lf").Replace("$lf", "$cr$lf")

    # 選んだピクセル値を差し込む。宣言はプロジェクト全体で 1 つだけ。
    # 置換できなかったら黙って既定のまま作らずに止める。「選んだのに効いて
    # いない成果物」がいちばんたちが悪いので、作れないほうがまし。
    $lines = $text.Split($lf)
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].TrimEnd($cr).StartsWith($decl)) {
            $tail = ''
            if ($lines[$i].EndsWith($cr)) { $tail = $cr }
            $lines[$i] = "$decl $Unit$note$tail"
            $unitHits++
        }
    }
    $text = [string]::Join($lf, $lines)

    $outPath = Join-Path $dist ("{0}_{1}px{2}" -f $s.BaseName, $Unit, $s.Extension)
    [System.IO.File]::WriteAllBytes($outPath, $cp932.GetBytes($text))
    $back = $cp932.GetString([System.IO.File]::ReadAllBytes($outPath))
    if ($back -ne $text) {
        throw "CP932 round trip changed $($s.Name). Fix the characters that do not survive it."
    }
    $outFiles += $outPath
}

if ($unitHits -ne 1) { throw "PB_UNIT の宣言が $unitHits 個ありました。1 個であるべきです" }
Write-Host "PB_UNIT = $Unit"
Write-Host "sources: $($outFiles.Count) files -> $dist"

$xlsm    = Join-Path $dist "VBA Pixel Bridge ${Unit}px.xlsm"
$pidFile = Join-Path $here "build-owned-${Unit}px.pid"
$log     = Join-Path $here "build-child-${Unit}px.log"
foreach ($f in @($pidFile, $log, "$log.err")) { if (Test-Path $f) { Remove-Item $f -Force } }
# 前回の成果物と、置き去りの BE コピーを消してから作り直す
foreach ($f in @($xlsm, (Join-Path $dist 'pixelbridge_be.xlsm'))) {
    if (Test-Path $f) { Remove-Item $f -Force }
}

function Read-Shared([string]$p) {
    if (-not (Test-Path $p)) { return '' }
    try {
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'ReadWrite')
        $sr = New-Object System.IO.StreamReader($fs)
        $t  = $sr.ReadToEnd(); $sr.Close(); $fs.Close(); return $t
    } catch { return '' }
}

# 取り込むファイルは 1 行 1 つのリストで渡す。引数へ並べるとパスの空白や
# 括弧で崩れるため。
$listFile = Join-Path $here "build-files-${Unit}px.txt"
Set-Content -Path $listFile -Value $outFiles -Encoding UTF8

$childArgs = @('-ExecutionPolicy','Bypass','-NoProfile','-File',"`"$child`"",
               '-FileList',"`"$listFile`"",'-Out',"`"$xlsm`"",'-PidFile',"`"$pidFile`"")
$p = Start-Process -FilePath 'powershell.exe' -ArgumentList $childArgs -PassThru `
        -RedirectStandardOutput $log -RedirectStandardError "$log.err" -WindowStyle Hidden
$deadline = (Get-Date).AddSeconds($TimeoutSec)
while (-not $p.HasExited -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 400 }
if (-not $p.HasExited) {
    Write-Host 'DEADLINE: the build child is stuck (a VBA compile error is a modal)'
    foreach ($id in ((Read-Shared $pidFile).Trim() -split ',' | Where-Object { $_ })) {
        if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) { Stop-Process -Id ([int]$id) -Force }
    }
    Start-Sleep -Milliseconds 400
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force }
}
foreach ($id in ((Read-Shared $pidFile).Trim() -split ',' | Where-Object { $_ })) {
    if (Get-Process -Id ([int]$id) -ErrorAction SilentlyContinue) { Stop-Process -Id ([int]$id) -Force }
}
if (Test-Path $log) { Get-Content $log | ForEach-Object { "  $_" } }
if ((Test-Path "$log.err") -and (Get-Item "$log.err").Length -gt 0) {
    Get-Content "$log.err" | ForEach-Object { "  ERR $_" }
}
if (-not (Test-Path $xlsm)) { throw 'build failed: the workbook was not produced' }
Write-Host "wrote $xlsm ($((Get-Item $xlsm).Length) bytes)"
