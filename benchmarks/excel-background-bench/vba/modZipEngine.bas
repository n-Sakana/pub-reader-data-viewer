Attribute VB_Name = "modZipEngine"
'==============================================================================
' modZipEngine -- ZipEngine.xlsm に入る側。方式7 で不可視の別 Excel の中で動く。
'
' 入口は 1 つだけ。前面が Application.Run で 1 回呼ぶ。
'
'   ZipEngine_Run(masterSheet, inputSheet, outputSheet, masterRows, rows)
'       -> "masterMs=..;inputMs=..;dictMs=..;convertMs=..;writeMs=..;rows=.."
'
' 3 つの引数は前面プロセスにある Worksheet への参照そのもの。
' CaseDesk の StartWorker が worker へ ThisWorkbook を渡しているのと同じ形で、
' 前面から渡してもらう。だから中間ファイルも分割呼び出しも要らない。
'
' 中身は方式2 (高速VBA) と**同じモジュールの同じ関数**:
'     modZipRule.BuildDictFromMaster   辞書構築
'     modZipRule.ConvertBlock          変換と出力配列作成
' 違いは、これが別プロセスで走ることと、Range の読み書きがプロセス境界を
' 越えることだけ。
'
' 待機・ポーリング・チャンク分割はどこにも無い。Application.Run が戻ることが
' そのまま完了で、前面はそれを受けて通知セルを 1 回更新する。
'==============================================================================
Option Explicit

Public Function ZipEngine_Run(ByVal wsMaster As Object, ByVal wsInput As Object, _
                              ByVal wsOutput As Object, ByVal firstRow As Long, _
                              ByVal masterRows As Long, ByVal rows As Long) As String
    Dim master As Variant, src As Variant
    Dim dst() As String
    Dim d As Object
    Dim tk As Currency
    Dim masterMs As Double, inputMs As Double, dictMs As Double
    Dim convertMs As Double, writeMs As Double

    ' Master を 1 回で読む
    tk = modZipRule.QpcNow()
    master = wsMaster.Range(wsMaster.Cells(firstRow, 1), _
                            wsMaster.Cells(firstRow + masterRows - 1, 2)).Value
    masterMs = modZipRule.QpcSince(tk) * 1000#

    ' Input を 1 回で読む
    tk = modZipRule.QpcNow()
    src = wsInput.Range(wsInput.Cells(firstRow, 1), _
                        wsInput.Cells(firstRow + rows - 1, 1)).Value
    inputMs = modZipRule.QpcSince(tk) * 1000#

    ' 辞書構築 (方式2 と同一関数)
    tk = modZipRule.QpcNow()
    Set d = modZipRule.BuildDictFromMaster(master, masterRows)
    dictMs = modZipRule.QpcSince(tk) * 1000#

    ' 変換 + 出力配列作成 (方式2 と同一関数)
    tk = modZipRule.QpcNow()
    modZipRule.ConvertBlock src, d, dst, rows
    convertMs = modZipRule.QpcSince(tk) * 1000#

    ' Output へ 1 回で書く
    tk = modZipRule.QpcNow()
    wsOutput.Range(wsOutput.Cells(firstRow, 1), _
                   wsOutput.Cells(firstRow + rows - 1, 1)).Value = dst
    writeMs = modZipRule.QpcSince(tk) * 1000#

    ZipEngine_Run = "rows=" & rows & _
                    ";masterMs=" & Format$(masterMs, "0.000") & _
                    ";inputMs=" & Format$(inputMs, "0.000") & _
                    ";dictMs=" & Format$(dictMs, "0.000") & _
                    ";convertMs=" & Format$(convertMs, "0.000") & _
                    ";writeMs=" & Format$(writeMs, "0.000")
End Function

' 前面から生存確認したいとき用
Public Function ZipEngine_Ping() As String
    ZipEngine_Ping = "ZipEngine ok"
End Function
