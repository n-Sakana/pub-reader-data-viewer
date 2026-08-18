# e2e_close.ps1 -- VERIFICATION ONLY.
#
# この検証が自分で起こした Excel だけを、人が ✕ を押すのと同じやり方で閉じる。
# Auto_Close を走らせて後始末（BE の終了、Temp の削除、表示設定の復帰）を
# させたいので、kill は最後の手段にする。
#
# 【重要】走っている Excel を名前で拾って全部落とすことは絶対にしない。
# 利用者が開いているブックを巻き込む。所有の証拠は e2e_open.ps1 が書き出した
# PID 台帳だけで、そこに無い PID には Quit も kill も送らない。
#
# 閉じ方は WM_CLOSE。以前は UIA の WindowPattern.Close() を使っていたが、
# 相手は自分も UIA クライアントなので UIA 越しに触ると刺さる。実測：Close()
# から Auto_Close が走り出すまで毎回 35 秒。WM_CLOSE なら 5 秒で後始末まで
# 終わる。測っているものを歪めない閉じ方を使う。
[CmdletBinding()]
param(
    [int[]] $Pids = @(),
    [int] $GraceSec = 45
)
$ErrorActionPreference = 'Stop'

if ($Pids.Count -eq 0) {
    $ledger = Join-Path $PSScriptRoot 'e2e-owned.pid'
    if (-not (Test-Path -LiteralPath $ledger)) {
        Write-Host 'no owned-pid ledger; this run started nothing, so nothing to close'
        exit 0
    }
    $Pids = @(Get-Content -LiteralPath $ledger |
              ForEach-Object { $_.Trim() } |
              Where-Object { $_ -match '^\d+$' } |
              ForEach-Object { [int]$_ })
}

# 台帳にあっても、いま生きていて、かつ本当に Excel のものだけ。PID は使い回される。
$Pids = @($Pids | Where-Object {
    $p = Get-Process -Id $_ -ErrorAction SilentlyContinue
    $p -and $p.ProcessName -eq 'EXCEL'
})
if ($Pids.Count -eq 0) { Write-Host 'no owned excel is running'; exit 0 }
Write-Host ("owned excel to close: " + ($Pids -join ','))

Add-Type -Namespace Cl -Name W -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool PostMessage(System.IntPtr h, uint m, System.IntPtr w, System.IntPtr l);
'@

foreach ($id in $Pids) {
    $p = Get-Process -Id $id -ErrorAction SilentlyContinue
    if (-not $p) { continue }
    if ($p.MainWindowHandle -eq [System.IntPtr]::Zero) {
        Write-Host "pid $id has no window (BE); leaving it to the FE's own shutdown"
        continue
    }
    Write-Host "closing window [$($p.MainWindowTitle)] of pid $id"
    [void][Cl.W]::PostMessage($p.MainWindowHandle, 0x0010, [System.IntPtr]::Zero, [System.IntPtr]::Zero)
}

$deadline = (Get-Date).AddSeconds($GraceSec)
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 700
    $alive = @($Pids | Where-Object { Get-Process -Id $_ -ErrorAction SilentlyContinue })
    if ($alive.Count -eq 0) { Write-Host 'all owned excel exited cleanly'; exit 0 }
}
foreach ($id in $Pids) {
    if (Get-Process -Id $id -ErrorAction SilentlyContinue) {
        Write-Host "still alive after $GraceSec s; killing owned pid $id"
        Stop-Process -Id $id -Force
    }
}
