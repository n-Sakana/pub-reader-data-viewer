Attribute VB_Name = "modZipExcel"
'==============================================================================
' modZipExcel -- 方式7: 不可視の別 Excel プロセスで処理する。
'
' 起動と FE/BE 分離は CaseDesk の StartWorker をそのまま踏襲する:
'   前後の EXCEL.EXE の PID を控える
'   → CreateObject("Excel.Application")
'   → Visible=False / DisplayAlerts=False
'   → AutomationSecurity を退避して 1 (Low) にする
'   → Workbooks.Open(ReadOnly:=True, UpdateLinks:=0)
'   → AutomationSecurity を戻す
'   → Workbooks(Count) を保持
'   → workerApp.Run "Module.Proc", ... と**モジュール修飾で 1 回だけ呼ぶ**
'   → 失敗したら Quit
'     (casedesk の CaseDeskMain.bas StartWorker、
'      xltoolrack の JobHost.cls StartJob と同じ組み立て)
'
' CaseDesk が worker へ ThisWorkbook を渡しているのと同じく、正本の 3 シートへの
' 参照をそのまま相手へ渡す。相手はその参照に対して読み書きするので、
' 中間ファイルも分割呼び出しも要らない。
'
' 持ち込まないもの: 既存実装の OnTime による「1秒処理して1秒譲る」ループと
' Application.Wait。固定待機・ポーリング・チャンク分割はどこにも無い。
' Application.Run が戻ることが完了で、前面がそれを受けて通知セルを 1 回更新する。
'
' 触るのは自分が CreateObject で起こしたインスタンスだけ。PID を控えるのは
' 後始末で本当に消えたかを確かめるためで、強制終了はしない (終了は COM の Quit)。
'==============================================================================
Option Explicit

Private Const msoAutomationSecurityLow As Long = 1
Private Const xlCalculationManual      As Long = -4135

Private m_App2 As Object
Private m_Wb2  As Object
Private m_Pid2 As Long

Public Function RunExcelMethod() As ZbResult
    Dim r As ZbResult
    Dim enginePath As String
    Dim wsM As Worksheet, wsI As Worksheet, wsO As Worksheet
    Dim t0 As Currency, tk As Currency, tLaunch As Currency
    Dim payload As String

    r = modZipBench.ZB_NewResult(7)
    r.UiaOk = "-"                       ' この方式は UIA を使わない (COM で直接話す)
    enginePath = modZipBench.ZB_Root() & "\engine\ZipEngine.xlsm"

    ' 相手が前面のセルへ書くので、前面のイベントは切る。
    ' 書込は 1 回なので変更イベントも 1 回だけ。通知のときだけ入れる。
    modZipBench.ZB_EnterMeasure False
    modZipBench.ZbArmSignal

    On Error GoTo Fail
    If Not modZipRule.FileExists(enginePath) Then
        Err.Raise 53, "RunExcelMethod", "engine\ZipEngine.xlsm が無い"
    End If

    Set wsM = modZipBench.DataSheet(DATA_MASTER)
    Set wsI = modZipBench.DataSheet(DATA_INPUT)
    Set wsO = modZipBench.DataSheet(DATA_OUTPUT)
    modZipBench.ClearOutput

    '--- 起動 (処理E2Eとは別枠) -----------------------------------------------
    tLaunch = modZipRule.QpcNow()

    Dim beforePids As Object, prevSec As Long
    Set beforePids = ExcelPids()

    Set m_App2 = CreateObject("Excel.Application")
    m_App2.Visible = False
    m_App2.DisplayAlerts = False
    m_App2.ScreenUpdating = False
    m_App2.EnableEvents = False
    m_Pid2 = NewPidSince(beforePids)
    r.Created = "OK"
    ' Calculation はブックが 1 冊も開いていないと設定できない (1004)。開いた後。

    prevSec = m_App2.AutomationSecurity
    m_App2.AutomationSecurity = msoAutomationSecurityLow
    m_App2.Workbooks.Open enginePath, ReadOnly:=True, UpdateLinks:=0
    m_App2.AutomationSecurity = prevSec
    Set m_Wb2 = m_App2.Workbooks(m_App2.Workbooks.Count)
    m_App2.Calculation = xlCalculationManual
    r.Launched = "OK"
    r.LaunchSec = modZipRule.QpcSince(tLaunch)

    '--- 処理 (別プロセスの中で 1 回の呼び出し) -------------------------------
    t0 = modZipRule.QpcNow()
    payload = m_App2.Run("modZipEngine.ZipEngine_Run", wsM, wsI, wsO, _
                         FIRST_ROW, g_MasterN, g_N)
    r.Converted = "OK"

    r.MasterSec = MsOf(payload, "masterMs")
    r.InputSec = MsOf(payload, "inputMs")
    r.DictSec = MsOf(payload, "dictMs")
    r.ConvertSec = MsOf(payload, "convertMs")
    r.WriteSec = MsOf(payload, "writeMs")

    '--- 通知セルを 1 回更新。完了検知は Worksheet_Change ---------------------
    tk = modZipRule.QpcNow()
    modZipBench.ZbFireLocalSignal "DONE|7|" & g_N
    r.NotifySec = modZipRule.QpcSince(tk)

    If g_SigFired Then
        r.TotalSec = modZipRule.QpcSec(t0, g_SigStamp)
    Else
        r.TotalSec = modZipRule.QpcSince(t0)
    End If

    '--- ここから計測の外 -----------------------------------------------------
    modZipBench.ZB_VerifyFromCells r
    r.Note = "別Excelプロセス pid=" & m_Pid2 & " / Application.Run は 1 回 / " & _
             "変換は modZipRule.ConvertBlock (方式2と同一関数) / 中間ファイルなし・待機なし"

Finish:
    On Error Resume Next
    QuitOwnExcel
    modZipBench.ZbDisarmSignal
    modZipBench.ZB_LeaveMeasure
    RunExcelMethod = r
    Exit Function

Fail:
    r.ErrNumber = Err.Number
    If r.TotalSec = 0 Then r.TotalSec = modZipRule.QpcSince(t0)
    If Err.Number = 18 Then
        r.Note = "Esc で取消"
        r.MatchText = "取消"
        g_Cancel = True
    Else
        r.Note = "Err " & Err.Number & ": " & Err.Description
    End If
    If r.Created = "-" Then r.Created = "NG"
    If r.Launched = "-" Then r.Launched = "NG"
    If r.Converted = "-" Then r.Converted = "NG"
    Resume Finish
End Function

Private Function MsOf(ByVal payload As String, ByVal key As String) As Double
    Dim s As String
    s = modZipUia.PayloadValue(payload, key)
    If Len(s) > 0 Then If IsNumeric(s) Then MsOf = CDbl(s) / 1000#
End Function

' 自分が起こしたインスタンスだけを閉じる。終了は COM の Quit だけ。
Public Sub QuitOwnExcel()
    On Error Resume Next
    If Not m_Wb2 Is Nothing Then
        m_Wb2.Close SaveChanges:=False
        Set m_Wb2 = Nothing
    End If
    If Not m_App2 Is Nothing Then
        m_App2.DisplayAlerts = False
        m_App2.Quit
        Set m_App2 = Nothing
    End If
    Err.Clear
End Sub

Public Function OwnExcelStillAlive() As Boolean
    Dim d As Object
    If m_Pid2 = 0 Then Exit Function
    Set d = ExcelPids()
    OwnExcelStillAlive = d.Exists(CStr(m_Pid2))
End Function

Public Function OwnExcelPid() As Long
    OwnExcelPid = m_Pid2
End Function

' EXCEL.EXE の PID 一覧。WMI で読むだけで、何も変更しない。
Private Function ExcelPids() As Object
    Dim d As Object, wmi As Object, proc As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    For Each proc In wmi.ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
        d(CStr(proc.ProcessId)) = True
    Next proc
    Err.Clear
    Set ExcelPids = d
End Function

Private Function NewPidSince(ByVal beforePids As Object) As Long
    Dim afterPids As Object, k As Variant
    On Error Resume Next
    Set afterPids = ExcelPids()
    For Each k In afterPids.keys
        If Not beforePids.Exists(k) Then NewPidSince = CLng(k): Exit Function
    Next k
End Function
