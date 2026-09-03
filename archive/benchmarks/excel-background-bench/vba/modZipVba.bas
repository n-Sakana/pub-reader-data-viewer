Attribute VB_Name = "modZipVba"
'==============================================================================
' modZipVba -- 方式 1 と 方式 2。どちらも前面の Excel の中だけで完結する。
'
' 方式2 が「高速VBA基準」。正本の 3 シートに対して
'     Master を 1 回読む → Input を 1 回読む → 辞書を組む
'     → 変換して出力配列を作る → Output へ 1 回で書く → 通知セルを 1 回更新
' 辞書構築と変換ループは modZipRule.BuildDictFromMaster / ConvertBlock。
' 方式7 (別Excel) は**同じモジュールの同じ関数**を呼ぶので、
' 両者の違いは「どのプロセスで走るか」だけになる。
'
' 方式1 は意図的に遅い書き方。Input を 1 セルずつ読み、Output へ 1 セルずつ書く。
' 読みと書きが同じループに混ざるため工程を分けられない。変換秒に合算し、
' Input読込秒と出力書込秒は 0 として、メモにその旨を書く。
' (Master 読込と辞書構築は方式2 と同じ手順なので、そこは分けて計れる。)
'
' 起動秒は 0。起こすプロセスが無い。
' 処理E2E は「仕事を始めてから通知セルの Worksheet_Change を受けるまで」。
'==============================================================================
Option Explicit

Public Function RunVbaMethod(ByVal methodNo As Long) As ZbResult
    Dim r As ZbResult
    Dim wsM As Worksheet, wsI As Worksheet, wsO As Worksheet
    Dim d As Object
    Dim master As Variant, src As Variant
    Dim dst() As String
    Dim n As Long, mn As Long
    Dim t0 As Currency, tk As Currency

    r = modZipBench.ZB_NewResult(methodNo)
    r.Created = "-": r.Launched = "-": r.UiaOk = "-"
    n = g_N
    mn = g_MasterN

    ' セルへ書く方式なのでイベントは切る。方式1 は 100 万回の変更イベントを
    ' 起こしてしまい、測っているものが変わる。通知のときだけ一時的に入れる。
    modZipBench.ZB_EnterMeasure False
    modZipBench.ZbArmSignal

    On Error GoTo Fail
    Set wsM = modZipBench.DataSheet(DATA_MASTER)
    Set wsI = modZipBench.DataSheet(DATA_INPUT)
    Set wsO = modZipBench.DataSheet(DATA_OUTPUT)
    modZipBench.ClearOutput

    t0 = modZipRule.QpcNow()                        ' ← 処理E2E 開始

    ' --- Master を 1 回で読む ---
    tk = modZipRule.QpcNow()
    master = wsM.Range(wsM.Cells(FIRST_ROW, 1), wsM.Cells(FIRST_ROW + mn - 1, 2)).Value
    r.MasterSec = modZipRule.QpcSince(tk)

    ' --- 辞書構築 (方式7 と同じ関数) ---
    tk = modZipRule.QpcNow()
    Set d = modZipRule.BuildDictFromMaster(master, mn)
    r.DictSec = modZipRule.QpcSince(tk)

    If methodNo = 1 Then
        ' 1 セルずつ。読み・変換・書きが 1 つのループに混ざる。
        tk = modZipRule.QpcNow()
        ConvertCellByCell wsI, wsO, d, n
        r.ConvertSec = modZipRule.QpcSince(tk)
        r.InputSec = 0
        r.WriteSec = 0
    Else
        ' --- Input を 1 回で読む ---
        tk = modZipRule.QpcNow()
        src = wsI.Range(wsI.Cells(FIRST_ROW, 1), wsI.Cells(FIRST_ROW + n - 1, 1)).Value
        r.InputSec = modZipRule.QpcSince(tk)

        ' --- 変換 + 出力配列作成 (方式7 と同じ関数) ---
        tk = modZipRule.QpcNow()
        modZipRule.ConvertBlock src, d, dst, n
        r.ConvertSec = modZipRule.QpcSince(tk)

        ' --- Output へ 1 回で書く ---
        tk = modZipRule.QpcNow()
        wsO.Range(wsO.Cells(FIRST_ROW, 1), wsO.Cells(FIRST_ROW + n - 1, 1)).Value = dst
        r.WriteSec = modZipRule.QpcSince(tk)
    End If

    ' --- 通知セルを 1 回更新。完了検知は Worksheet_Change ---
    tk = modZipRule.QpcNow()
    modZipBench.ZbFireLocalSignal "DONE|" & methodNo & "|" & n
    r.NotifySec = modZipRule.QpcSince(tk)

    If g_SigFired Then
        r.TotalSec = modZipRule.QpcSec(t0, g_SigStamp)
    Else
        r.TotalSec = modZipRule.QpcSince(t0)
    End If
    r.LaunchSec = 0
    r.Converted = "OK"

    If g_Cancel Then
        r.Converted = "NG"
        r.MatchText = "取消"
        r.Note = "取消により中断"
        GoTo Finish
    End If

    ' --- ここから計測の外 ---
    modZipBench.ZB_VerifyFromCells r
    If methodNo = 1 Then
        r.Note = "セルを1つずつ / Input・Output 往復 " & Format$(2& * n, "#,##0") & " 回 / " & _
                 "読み書きが同じループなので Input読込と出力書込は変換秒に合算"
    Else
        r.Note = "Master1回・Input1回・Output1回 / 変換は modZipRule.ConvertBlock (方式7と同一関数)"
    End If

Finish:
    modZipBench.ZbDisarmSignal
    modZipBench.ZB_LeaveMeasure
    RunVbaMethod = r
    Exit Function

Fail:
    r.ErrNumber = Err.Number
    If r.TotalSec = 0 Then r.TotalSec = modZipRule.QpcSince(t0)
    r.Converted = "NG"
    If Err.Number = 18 Then
        r.Note = "Esc で取消"
        r.MatchText = "取消"
        g_Cancel = True
    Else
        r.Note = "Err " & Err.Number & ": " & Err.Description
    End If
    Resume Finish
End Function

'------------------------------------------------------------------------------
' 方式1: Input を 1 セルずつ読み、Output へ 1 セルずつ書く。
' 辞書引きと正規化は方式2 と同じものを使う。差はセルとの往復回数だけ。
'------------------------------------------------------------------------------
Private Sub ConvertCellByCell(ByVal wsI As Worksheet, ByVal wsO As Worksheet, _
                              ByVal d As Object, ByVal n As Long)
    Dim i As Long
    Dim k As String

    For i = 0 To n - 1
        k = modZipRule.NormalizeZip(CStr(wsI.Cells(FIRST_ROW + i, 1).Value))
        If Len(k) = 7 Then
            If d.Exists(k) Then
                wsO.Cells(FIRST_ROW + i, 1).Value = d(k)
            Else
                wsO.Cells(FIRST_ROW + i, 1).Value = ZB_NOTFOUND
            End If
        Else
            wsO.Cells(FIRST_ROW + i, 1).Value = ZB_NOTFOUND
        End If
    Next i
End Sub
