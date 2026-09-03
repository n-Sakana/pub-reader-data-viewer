# e2e_open.ps1 -- VERIFICATION ONLY. Opens the distributable the way a person
# does: the shell association (what a double click runs). If Excel blocks the
# macros, it presses "Enable Content" through UI Automation -- the same gesture
# a person would make, not a settings change.
#
# It records which Excel processes it started, so cleanup only ever touches
# those. Nothing that was already running is used or closed.
#
# 【重要】**この Excel の UIA ツリーを、必要が無いのに歩かない。**
# FindAll(Descendants) を 1 回かけるだけで、その Excel の中で自身の UIA
# プロバイダが動き出し、以後セルへの書き込みが 1 回 180ms 級になる（実測：
# 歩く前のティックは 7～15ms、歩いたあとは同じ盤で 1.8～2.0 秒。歩いている
# あいだ Excel は 44 秒返らなかった）。閉じ方を UIA から WM_CLOSE へ替えた
# のと同じ理由で、開き方からも UIA を外す。**測っているものを歪めない。**
#
# 代わりに「マクロが走ったか」を製品自身の足跡で見る。FE は起動のごく初めに
# %TEMP%\pixelbridge\ を作って生存の印を書くので、それが出れば有効化は要らない。
# 出なければ人が押すはずのボタンが出ているということなので、そこではじめて
# UIA を使い、以後の時間はもう清潔ではないと画面に出す。
[CmdletBinding()]
param(
    [string] $Book = 'C:\repos\pub\reader-data-viewer\showcase\dist\VBA Pixel Bridge.xlsm',
    [int] $WaitSec = 25,
    [string] $Shot = ''
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $here 'e2e-owned.pid'

Add-Type -Namespace Op -Name W -MemberDefinition @'
[StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
[DllImport("user32.dll")] public static extern bool GetWindowRect(System.IntPtr h, out RECT r);
[DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
'@
[void][Op.W]::SetProcessDPIAware()

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
Write-Host "excel before: $($before -join ',')"

$mark = Join-Path $env:TEMP 'pixelbridge'
Start-Process -FilePath $Book | Out-Null
$deadline = (Get-Date).AddSeconds(40)
$mine = @()
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 500
    $mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
              Where-Object { $before -notcontains $_ })
    if ($mine.Count -gt 0) { break }
}
if ($mine.Count -eq 0) { throw 'excel did not start' }
Add-Content -Path $pidFile -Value ($mine -join ',') -Encoding ascii
Write-Host "excel started by this script: $($mine -join ',')"

# マクロが走ったか。製品が %TEMP%\pixelbridge\ を作れば走っている。
$deadline = (Get-Date).AddSeconds(20)
$ran = $false
while ((Get-Date) -lt $deadline) {
    if (Test-Path -LiteralPath $mark) { $ran = $true; break }
    Start-Sleep -Milliseconds 400
}

if ($ran) {
    Write-Host 'macros ran (the product created %TEMP%\pixelbridge) -- no UIA was used on this Excel'
} else {
    Write-Host 'WARNING: macros did not run within 20 s; pressing the dialogs through UIA.'
    Write-Host 'WARNING: from here the timings of this instance are NOT clean (UIA touched the process).'
    # Excel の「前回は重大なエラーが発生しました」と、マクロのセキュリティバー。
    & (Join-Path $here 'e2e_dialog.ps1') -Pids $mine -ButtonPattern '^はい' -WindowTextPattern '重大なエラー'
    Start-Sleep -Seconds 2
    & (Join-Path $here 'e2e_dialog.ps1') -Pids $mine `
        -ButtonPattern 'コンテンツの有効化|Enable Content|有効にする'
}

Write-Host "waiting $WaitSec s for the screen to build ..."
Start-Sleep -Seconds $WaitSec

# 窓の位置は Win32 で読む（UIA だと、それだけで上の毒を仕込むことになる）
foreach ($id in $mine) {
    $p = Get-Process -Id $id -ErrorAction SilentlyContinue
    if ($null -eq $p -or $p.MainWindowHandle -eq [System.IntPtr]::Zero) { continue }
    $r = New-Object Op.W+RECT
    [void][Op.W]::GetWindowRect($p.MainWindowHandle, [ref]$r)
    Write-Host ("excel window rect: {0},{1} {2}x{3}  name=[{4}]" -f `
        $r.L, $r.T, ($r.R - $r.L), ($r.B - $r.T), $p.MainWindowTitle)
}
if ($Shot) {
    & (Join-Path $here 'e2e_shot.ps1') -Out $Shot
}
Write-Host 'done'
