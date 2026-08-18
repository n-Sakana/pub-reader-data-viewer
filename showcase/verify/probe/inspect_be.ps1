# inspect_be.ps1 -- VERIFICATION ONLY.
# BE 用コピーを不可視の Excel で開き、参照設定を並べ、VBAProject を明示的に
# コンパイルしてみる。コンパイルで止まれば「見えないモーダル」が原因と分かる。
#
# 【警告・実測】下の VBE コンパイル（30-34 行）は、その Excel を壊す。これを
# 実行したインスタンスでは Application.OnTime が発火しなくなり、BE の常駐
# ループも 1 周で無応答になる。私はこれに丸一日を溶かした。「OnTime が
# 発火しない」という観測は全部このスクリプトから採ったもので、VBE を触らない
# probe_be_loop.ps1 では同じブックが最後まで通る（再現手順は
# probe_be_loop.ps1 -Compile）。
# BE の挙動を測る目的でこのスクリプトを使わないこと。参照設定と
# コンポーネント一覧を見るだけに使い、見たら捨てること。
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string] $Book,
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

$xl.AutomationSecurity = 1
$wb = $xl.Workbooks.Open($Book)
Note "opened $($wb.Name)"
foreach ($r in $wb.VBProject.References) {
    Note ("  ref {0}  broken={1}  {2}" -f $r.Name, $r.IsBroken, $r.Description)
}
foreach ($c in $wb.VBProject.VBComponents) { Note ("  comp {0} lines={1}" -f $c.Name, $c.CodeModule.CountOfLines) }

Note 'compiling...'
$vbe = $xl.VBE
$vbe.ActiveVBProject = $wb.VBProject
$ctl = $vbe.CommandBars.FindControl(1, 578)
if ($null -eq $ctl) { Note 'NOCONTROL' } else { $ctl.Execute(); Note "compiled enabled=$($ctl.Enabled)" }

Note 'calling PbPing...'
Note ("PbPing -> " + $xl.Run("'" + $wb.Name + "'!PbPing"))
foreach ($t in @('PbArmCount')) {
  try {
    $r = $xl.Run("'" + $wb.Name + "'!" + $t)
    Note ("$t -> $r")
  } catch {
    Note ("$t -> EXCEPTION " + $_.Exception.Message)
    break
  }
}
Note 'calling PbBeBootstrap...'
$sw=[System.Diagnostics.Stopwatch]::StartNew()
$xl.Run("'" + $wb.Name + "'!PbBeBootstrap")
Note ("PbBeBootstrap returned in " + $sw.ElapsedMilliseconds + " ms")
Note 'waiting 6s without touching Excel...'
Start-Sleep -Seconds 6
Note ('fired count -> ' + $xl.Run("'" + $wb.Name + "'!PbFiredCount"))

$pf = [System.IO.Path]::Combine($env:TEMP, 'pixelbridge', 'progress.json')
$dir = [System.IO.Path]::Combine($env:TEMP, 'pixelbridge')
foreach ($m in @('mark-plain.txt','mark-live.txt','mark-bare.txt')) {
  Note ('  ' + $m + ' -> ' + (Test-Path ([System.IO.Path]::Combine($dir,$m))))
}
Note ("progress.json after 4s: " + (Test-Path $pf))
if (Test-Path $pf) { Note ("  content: " + (Get-Content $pf -Raw).Trim()) }
Note 'DONE'
try { $wb.Close($false) } catch { }
try { $xl.Quit() } catch { }
Start-Sleep -Milliseconds 500
foreach ($id in $mine) { if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force } }
