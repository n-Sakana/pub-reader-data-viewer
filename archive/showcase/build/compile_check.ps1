# compile_check.ps1
#
# VBA は既定で「必要な手続きだけを都度コンパイル」する。だからビルド時に
# PbPing が通っても、まだ呼ばれていない手続きに構文エラーが残っていれば、
# 実行時にそこで初めて「コンパイル エラー」のモーダルが出る。しかもそれが
# 不可視の Excel（BE）で起きると、誰にも見えない停止になる。
#
# そこで VBE の「VBAProject のコンパイル」を明示的に実行して、全手続きを
# 一度にコンパイルさせる。失敗するとモーダルが出るので、この処理は
# 監視付きの子プロセス（このスクリプト自体）で行い、親が期限で落とす。
#
# 【警告】このスクリプトを、そのまま動かし続ける Excel に対して使わないこと。
# 自動化（CreateObject）で起こした Excel の VBE を触ると、そのインスタンスは
# 壊れる。実測：VBE のコンパイルを実行したあとの BE は、Application.OnTime で
# 起動した常駐ループが最初の 1 周を書いた直後に無応答になり、180 秒待っても
# 戻らなかった。同じ手順を VBE に触らずに走らせると最後まで通る。
# しかも FindControl(1,578) は enabled=False（＝押せない）で返ることがあり、
# その場合 Execute() は何も検査していない。つまり「壊すだけで、確かめない」。
# 使うなら、そのあと捨てる使い捨ての Excel に対してだけにすること。
# ビルド本体（build_showcase.ps1）はこれを使わず、Application.Run PbPing で
# 取り込みと公開入口だけを確かめる。未実行手続きまで含む全体コンパイルの証明では
# ないため、最後は実機の一連の操作で確かめる。
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Bas,
    [Parameter(Mandatory=$true)][string] $PidFile,
    [Parameter(Mandatory=$true)][string] $Res
)
$ErrorActionPreference = 'Stop'
function Note([string]$s) { Add-Content -LiteralPath $Res -Value $s -Encoding utf8; Write-Host $s }

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
Start-Sleep -Milliseconds 300
$mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
          Where-Object { $before -notcontains $_ })
Set-Content -LiteralPath $PidFile -Value ($mine -join ',') -Encoding ascii

$wb = $xl.Workbooks.Add()
[void]$wb.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)
[void]$wb.VBProject.VBComponents.Import($Bas)
Note "imported"

# VBE の Compile VBAProject（コントロール ID 578）
$vbe = $xl.VBE
$vbe.ActiveVBProject = $wb.VBProject
$ctl = $vbe.CommandBars.FindControl(1, 578)
if ($null -eq $ctl) { Note 'NOCONTROL'; }
else {
    $ctl.Execute()
    Note "compiled (control enabled=$($ctl.Enabled))"
}
Note 'DONE'
try { $wb.Close($false) } catch { }
try { $xl.Quit() } catch { }
Start-Sleep -Milliseconds 500
foreach ($id in $mine) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force } }
