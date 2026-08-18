# build_showcase.ps1
#
# 成果物をつくる。出来るのは 2 つだけ。
#   dist\modPixelBridge.bas        ANSI(CP932) の .bas 1 枚。これが仕様上の成果物。
#   dist\VBA Pixel Bridge.xlsm     その .bas を取り込み、UIAutomationClient の
#                                  参照を付けたブック。ふつうに開けば動く。
#
# ここでやる COM は全部この子プロセス側（build_child.ps1）に置いてある。VBA の
# コンパイルエラーは「見えないモーダル」になり得るので、ブックは不可視のまま
# つくり、保存前に PbPing で必ずコンパイルを通す。返ってこなければ、この
# スクリプトが自分の持ち物だけを落とす。
[CmdletBinding()]
param(
    # 1 セルが受け持つ設計 px。利用者が選ぶもので、既定値は置かない。
    # 既定値を置くと「選ばなかった」と「4 を選んだ」が区別できなくなる。
    [Parameter(Mandatory=$true)][ValidateSet('1','2','4')][string] $Unit,
    [int] $TimeoutSec = 180
)
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$root = Split-Path -Parent $here
$src  = Join-Path $root 'src\modPixelBridge.bas'
$dist = Join-Path $root 'dist'
$child= Join-Path $here 'build_child.ps1'
if (-not (Test-Path $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

# .bas は取り込まれるときシステムの ANSI コードページで読まれる（日本語 Windows
# なら CP932）。ソースは UTF-8 で持ち、成果物は CP932 で出す。
$utf8 = [System.IO.File]::ReadAllText($src, (New-Object System.Text.UTF8Encoding($false)))

# 選んだピクセル値をソースへ埋める。ここが唯一の差し込み口。置換できなかったら
# 黙って既定のまま作らずに止める。「選んだのに効いていない成果物」がいちばん
# たちが悪いので、作れないほうがまし。
$cr = [char]13
$lf = [char]10
$decl = 'Private Const PB_UNIT As Long ='
$note = "              ' 1 セルが何設計 px 分か（ビルド時に選ぶ: 1 / 2 / 4）"
$lines = $utf8.Split($lf)
$hit = 0
for ($i = 0; $i -lt $lines.Count; $i++) {
    if ($lines[$i].TrimEnd($cr).StartsWith($decl)) {
        $tail = ''
        if ($lines[$i].EndsWith($cr)) { $tail = $cr }
        $lines[$i] = "$decl $Unit$note$tail"
        $hit++
    }
}
if ($hit -ne 1) { throw "PB_UNIT の宣言がソースに $hit 個ありました。1 個であるべきです" }
$utf8 = [string]::Join($lf, $lines)
if ($utf8.IndexOf("$decl $Unit") -lt 0) { throw "PB_UNIT を $Unit に差し替えられませんでした" }
Write-Host "PB_UNIT = $Unit"
$cp932 = [System.Text.Encoding]::GetEncoding(932)
$basOut = Join-Path $dist "modPixelBridge_${Unit}px.bas"
[System.IO.File]::WriteAllBytes($basOut, $cp932.GetBytes($utf8))
$back = $cp932.GetString([System.IO.File]::ReadAllBytes($basOut))
if ($back -ne $utf8) {
    throw "CP932 round trip changed the source. Fix the characters that do not survive it."
}
Write-Host "wrote $basOut ($((Get-Item $basOut).Length) bytes, CP932)"

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

$childArgs = @('-ExecutionPolicy','Bypass','-NoProfile','-File',"`"$child`"",
               '-Bas',"`"$basOut`"",'-Out',"`"$xlsm`"",'-PidFile',"`"$pidFile`"")
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
