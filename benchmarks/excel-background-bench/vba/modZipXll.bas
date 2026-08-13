Attribute VB_Name = "modZipXll"
'==============================================================================
' modZipXll -- 方式16。Excel-DNA の XLL を Excel 自身のプロセスへ読み込ませ、
'              その中で変換する。
'
' 他の C# 方式との違いはプロセス境界の有無だけ。
'   方式 3-6, 8-10, 13-15  別プロセスの ZipWorker.exe が COM 越しに Excel を触る
'   方式 16                Excel のプロセスに .NET が入り、XLL の C API で直接触る
' 変換規則・辞書・ループは ZipWorker.cs と同じものを xll\ZbXll.cs に置いてある。
'
' 通知セルも Worksheet_Change も使わない。呼び出しが同期で返るので、
' 戻ってきた時点が完了時点。待つ相手がいないところに通知を挟んでも、
' 測っているものが増えるだけで意味がない。
' そのぶん通知秒は「E2E から各工程を引いた残り」= XLL 呼び出しの往復になる。
'
' 起動秒は RegisterXLL にかかった時間。初回は CLR とアドインの読み込みが入るので
' 大きく、2 回目以降はほぼ 0 になる。実測値をそのまま各回に記録する。
'==============================================================================
Option Explicit

Private Const XLL_FUNC As String = "ZbXllConvert"

Public Function RunXllMethod() As ZbResult
    Dim r As ZbResult
    Dim wsO As Worksheet
    Dim xllPath As String
    Dim ret As Variant
    Dim parts() As String
    Dim n As Long, mn As Long
    Dim t0 As Currency, tk As Currency

    r = modZipBench.ZB_NewResult(16)
    r.WorkerKind = "xll"
    r.Created = "-": r.UiaOk = "-"
    n = g_N
    mn = g_MasterN

    modZipBench.ZB_EnterMeasure False

    On Error GoTo Fail
    Set wsO = modZipBench.DataSheet(DATA_OUTPUT)
    modZipBench.ClearOutput

    xllPath = modZipBench.ZB_Root() & "\prebuilt\ZbXll64.xll"
    If Not modZipRule.FileExists(xllPath) Then
        Err.Raise 53, "RunXllMethod", "prebuilt\ZbXll64.xll が無い (xll\build_xll.ps1 を先に実行)"
    End If

    '--- 起動: XLL を Excel のプロセスへ読み込む ------------------------------
    tk = modZipRule.QpcNow()
    If Not Application.RegisterXLL(xllPath) Then
        r.LaunchSec = modZipRule.QpcSince(tk)
        r.Launched = "NG"
        Err.Raise 1004, "RunXllMethod", "RegisterXLL が False を返した"
    End If
    r.LaunchSec = modZipRule.QpcSince(tk)
    r.Launched = "OK"

    '--- 処理: Excel の中で走る -----------------------------------------------
    t0 = modZipRule.QpcNow()
    ret = Application.Run(XLL_FUNC, _
                          wsO.Parent.Name, DATA_MASTER, DATA_INPUT, DATA_OUTPUT, _
                          CDbl(mn), CDbl(n), CDbl(FIRST_ROW))
    r.TotalSec = modZipRule.QpcSince(t0)

    '--- 戻り値を分解する ------------------------------------------------------
    ' OK|rows|hash|master|input|dict|convert|write|notFound
    ' NG|段階|実エラー
    parts = Split(CStr(ret), "|")
    If UBound(parts) < 1 Then
        r.Converted = "NG"
        r.Note = "XLL の戻り値が読めない: " & CStr(ret)
        GoTo Finish
    End If

    If parts(0) <> "OK" Then
        r.Converted = "NG"
        r.Note = "XLL 内で失敗 (" & parts(1) & "): " & _
                 IIf(UBound(parts) >= 2, parts(2), "")
        GoTo Finish
    End If

    r.Hash = parts(2)
    r.MasterSec = CDbl(parts(3)) / 1000#
    r.InputSec = CDbl(parts(4)) / 1000#
    r.DictSec = CDbl(parts(5)) / 1000#
    r.ConvertSec = CDbl(parts(6)) / 1000#
    r.WriteSec = CDbl(parts(7)) / 1000#
    r.ImportSec = 0
    r.NotifySec = r.TotalSec - (r.MasterSec + r.InputSec + r.DictSec + r.ConvertSec + r.WriteSec)
    If r.NotifySec < 0 Then r.NotifySec = 0
    r.Converted = "OK"

    If g_Cancel Then
        r.Converted = "NG"
        r.MatchText = "取消"
        r.Note = "取消により中断"
        GoTo Finish
    End If

    '--- ここから計測の外 ------------------------------------------------------
    modZipBench.ZB_VerifyFromCells r
    r.Note = "Excel-DNA XLL / Excel内プロセス / XLL側 rows=" & parts(1) & _
             " hash=" & parts(2) & " master=" & parts(3) & "ms input=" & parts(4) & _
             "ms dict=" & parts(5) & "ms convert=" & parts(6) & "ms write=" & parts(7) & _
             "ms notFound=" & parts(8) & " / 通知セルは使わない (同期呼び出し)"

Finish:
    modZipBench.ZB_LeaveMeasure
    RunXllMethod = r
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
