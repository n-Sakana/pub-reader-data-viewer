# build_child.ps1 -- every COM call of the build lives here, so that a VBA
# compile error (which is a modal) blocks this process and not the watchdog.
[CmdletBinding()]
param(
    # 取り込むファイルの一覧（1 行 1 パス）。製品は責務ごとに複数の
    # モジュール／クラスへ分かれているので、まとめて取り込む。
    [Parameter(Mandatory=$true)][string] $FileList,
    [Parameter(Mandatory=$true)][string] $Out,
    [Parameter(Mandatory=$true)][string] $PidFile
)
$ErrorActionPreference = 'Stop'

$before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
Start-Sleep -Milliseconds 400
$mine = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id |
          Where-Object { $before -notcontains $_ })
Set-Content -Path $PidFile -Value ($mine -join ',') -Encoding ascii
Write-Host "owned excel pid: $($mine -join ',')"

$prevSheets = $xl.SheetsInNewWorkbook
$xl.SheetsInNewWorkbook = 1
$wb = $xl.Workbooks.Add()
$xl.SheetsInNewWorkbook = $prevSheets

# 参照は 1 つだけ: UIAutomationClient。IUnknown しか持たない型なので遅延束縛は
# 使えず、参照設定が要る。
[void]$wb.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)
foreach ($f in (Get-Content -Path $FileList -Encoding UTF8 | Where-Object { $_.Trim() -ne '' })) {
    [void]$wb.VBProject.VBComponents.Import($f.Trim())
    Write-Host "imported $(Split-Path -Leaf $f)"
}

# ThisWorkbook にポンプの見張りを 1 つだけ置く。
#
# 1 秒ごとの OnTime は、何かの拍子に一度途切れるとそれきりになる。実測：
# Excel の「ドキュメントの回復」ウィンドウを閉じた直後にティックが止まり、
# 以後 1 度も来なかった。そうなると時計もミニマップも UIA 同期も死んだ画面に
# なり、しかも閉じても Auto_Close まで届かず後始末が走らない。
#
# 直し方は実証済み実装（src\app\vba\modRdv3App.bas の Rdv3PumpEnsureArmed）と
# 同じ：ブック自身のイベントを合図に、鎖が切れていたら張り直す。利用者が次に
# どこかを選べば戻る。標準モジュールからはブックのイベントを取れないので、
# ここで 1 つだけ注入する。中身は呼び出し 1 行で、判断はすべて .bas 側にある。
$tw = $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
#
# BeforeClose も要る。閉じる要求が来た時点で鎖を畳まないと、その間に予約が
# 発火して次を仕掛け直し、Excel は閉じられないまま待たされる（実測：閉じる
# 要求から Auto_Close が走り出すまで 35 秒）。実証済み実装の
# Rdv3AppPrepareClose と同じ。
# 完了ダイアログは廃止した。ベンチの結果（経過・計算桁・表示桁）は盤面に
# 残るし、モーダルは次のクリックを飲み込んでしまう（実機で実測：ベンチ直後の
# 「最大化」がダイアログに吸われて効かなかった）。ここへ注入するのは、ポンプの
# 見張りと BeforeClose の後始末だけ。
$tw.AddFromString(@'
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    PbEnsureArmed
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    PbEnsureArmed
End Sub

' メモ帳を触ってから Excel へ戻ってきたとき。このデモはその往復が多いので、
' 戻った時点で鎖が切れていれば、ここで気づいて張り直す。
Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    On Error Resume Next
    PbEnsureArmed
End Sub

Private Sub Workbook_Activate()
    On Error Resume Next
    PbEnsureArmed
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    PbPrepareClose
End Sub
'@)

# 保存する前に公開入口を実際に呼び、取り込みと入口の解決を確かめる。
# VBA 全体を事前コンパイルする検査ではない。
$ping = $xl.Run('PbPing')
Write-Host "PbPing -> $ping"

# グリッドは保存しない。66 万セルぶんの寸法を持ったシートは、読み込んだあとで
# 前に出す一手だけで Excel を応答不能にする（実測）。その場で作れば行高 2 ms・
# 列幅 0 ms なので、開くたびに作るほうが速い。
$wb.SaveAs($Out, 52)          # xlOpenXMLWorkbookMacroEnabled
Write-Host "saved $Out"
$wb.Close($false)
$xl.Quit()
Start-Sleep -Milliseconds 600
foreach ($id in $mine) {
    if (Get-Process -Id $id -ErrorAction SilentlyContinue) { Stop-Process -Id $id -Force }
}
exit 0
