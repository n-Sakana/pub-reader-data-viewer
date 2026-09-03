Attribute VB_Name = "modZipUia"
'==============================================================================
' modZipUia -- Excel VBA を UI Automation の CLIENT として使う部分。
'
' 必須の参照設定 (遅延バインドは原理的に不可能。好みの問題ではない)
'     ツール > 参照設定 > "UIAutomationClient"
'     TypeLib {944DE083-8FB8-45CF-BCB7-C477ACB2F897} v1.0
'     -> C:\Windows\System32\UIAutomationCore.dll
'
' UIA のクライアント側インタフェースはすべて IUnknown 派生で IDispatch を
' 持たない。VBA の Object は IDispatch のことなので、代入はできても呼び出しが
' できない。だから早期バインドしか道がない。
' (遅延バインドでは実際にどこで落ちるかは build\UiaProbe.cs で確かめられる。)
'
' ワーカー側 (C#) は provider。こちらは client。運ぶのは
'     コマンド / 進捗 / 結果 / 取消 / 終了
' だけで、100万件のデータ本体は運ばない。データはファイル、制御は UIA。
'==============================================================================
Option Explicit

'--- UIA プロパティ id ---------------------------------------------------------
Private Const UIA_NamePropertyId         As Long = 30005
Private Const UIA_AutomationIdPropertyId As Long = 30011

'--- UIA パターン id -----------------------------------------------------------
Private Const UIA_InvokePatternId        As Long = 10000
Private Const UIA_ValuePatternId         As Long = 10002

'--- チャネル定義 (worker\ZipWorker.cs と一致させること) ----------------------
Public Const ZW_WINDOW_PREFIX As String = "ZIPWORKER::"
Private Const F_CMD      As String = "UIAW_CMD"
Private Const F_SUBMIT   As String = "UIAW_SUBMIT"
Private Const F_ACK      As String = "UIAW_ACK"
Private Const F_STATUS   As String = "UIAW_STATUS"
Private Const F_PROGRESS As String = "UIAW_PROGRESS"
Private Const F_RESULT   As String = "UIAW_RESULT"
Private Const F_HEART    As String = "UIAW_HEART"
Private Const F_INFO     As String = "UIAW_INFO"

#If VBA7 Then
    Public Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Public Declare PtrSafe Sub SleepMs Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#Else
    Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Public Declare Sub SleepMs Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
#End If

Private m_Uia      As UIAutomationClient.IUIAutomation
Private m_Win      As UIAutomationClient.IUIAutomationElement
Private m_Cmd      As UIAutomationClient.IUIAutomationElement
Private m_Submit   As UIAutomationClient.IUIAutomationElement
Private m_Ack      As UIAutomationClient.IUIAutomationElement
Private m_Status   As UIAutomationClient.IUIAutomationElement
Private m_Progress As UIAutomationClient.IUIAutomationElement
Private m_Result   As UIAutomationClient.IUIAutomationElement
Private m_Heart    As UIAutomationClient.IUIAutomationElement
Private m_Info     As UIAutomationClient.IUIAutomationElement
Private m_Seq      As Long

Public Function UiaBound() As Boolean
    UiaBound = Not (m_Win Is Nothing)
End Function

Private Function Uia() As UIAutomationClient.IUIAutomation
    If m_Uia Is Nothing Then Set m_Uia = New UIAutomationClient.CUIAutomation
    Set Uia = m_Uia
End Function

' 最上位ウィンドウはデスクトップ直下なので Name の一致 1 本で足りる。
' Name にトークンが入っているので、同時に走る別の実行と取り違えない。
Public Function UiaFindWorker(ByVal token As String) As UIAutomationClient.IUIAutomationElement
    On Error Resume Next
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Set cond = Uia().CreatePropertyCondition(UIA_NamePropertyId, ZW_WINDOW_PREFIX & token)
    Set UiaFindWorker = Uia().GetRootElement.FindFirst(TreeScope_Children, cond)
End Function

Private Function FindField(ByVal win As UIAutomationClient.IUIAutomationElement, _
                           ByVal autoId As String) As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Set cond = Uia().CreatePropertyCondition(UIA_AutomationIdPropertyId, autoId)
    Set FindField = win.FindFirst(TreeScope_Descendants, cond)
End Function

' 一度つないだら、以後のやりとりは同じ要素参照を使い回す。
Public Function UiaBind(ByVal token As String) As Boolean
    On Error GoTo Fail
    Dim w As UIAutomationClient.IUIAutomationElement

    Set w = UiaFindWorker(token)
    If w Is Nothing Then UiaBind = False: Exit Function

    Set m_Win = w
    Set m_Cmd = FindField(w, F_CMD)
    Set m_Submit = FindField(w, F_SUBMIT)
    Set m_Ack = FindField(w, F_ACK)
    Set m_Status = FindField(w, F_STATUS)
    Set m_Progress = FindField(w, F_PROGRESS)
    Set m_Result = FindField(w, F_RESULT)
    Set m_Heart = FindField(w, F_HEART)
    Set m_Info = FindField(w, F_INFO)
    m_Seq = 0

    UiaBind = Not (m_Cmd Is Nothing Or m_Submit Is Nothing Or m_Ack Is Nothing _
                Or m_Status Is Nothing Or m_Result Is Nothing Or m_Info Is Nothing)
    Exit Function
Fail:
    UiaBind = False
End Function

' ワーカーが UIA ツリーに現れるまで待つ。
' 固定長の待機は入れない。1 ms 刻みで見に行き、見つかった瞬間に返す。
' 見つからないうちは Excel を固めないよう DoEvents を回す。
Public Function UiaWaitForWorker(ByVal token As String, ByVal timeoutSec As Double, _
                                 ByRef waitedSec As Double) As Boolean
    Dim t0 As Double
    t0 = Timer
    Do
        If UiaBind(token) Then
            waitedSec = Timer - t0
            UiaWaitForWorker = True
            Exit Function
        End If
        DoEvents
        If g_Cancel Then Exit Do
        SleepMs 1
    Loop While Timer - t0 < timeoutSec
    waitedSec = Timer - t0
    UiaWaitForWorker = False
End Function

' worker -> client: ValuePattern を 1 つ読む
Public Function UiaRead(ByVal el As UIAutomationClient.IUIAutomationElement) As String
    On Error GoTo Fail
    If el Is Nothing Then UiaRead = "": Exit Function
    Dim vp As UIAutomationClient.IUIAutomationValuePattern
    Set vp = el.GetCurrentPattern(UIA_ValuePatternId)
    If vp Is Nothing Then UiaRead = "": Exit Function
    UiaRead = vp.CurrentValue
    Exit Function
Fail:
    UiaRead = "<read error " & Err.Number & ">"
End Function

Public Function UiaStatus() As String
    UiaStatus = UiaRead(m_Status)
End Function

Public Function UiaResult() As String
    UiaResult = UiaRead(m_Result)
End Function

Public Function UiaInfo() As String
    UiaInfo = UiaRead(m_Info)
End Function

Public Function UiaProgress() As String
    UiaProgress = UiaRead(m_Progress)
End Function

Public Function UiaHeart() As String
    UiaHeart = UiaRead(m_Heart)
End Function

' client -> worker: コマンド欄に書いて、submit を叩く。この 2 回が送信の全部。
' ワーカーは処理を終えてから ack に seq を書くので、ack の一致は本物の握手になる。
Public Function UiaSend(ByVal verb As String, ByVal argstr As String, _
                        Optional ByVal ackTimeoutSec As Double = 10) As Boolean
    On Error GoTo Fail
    If m_Cmd Is Nothing Or m_Submit Is Nothing Then UiaSend = False: Exit Function

    m_Seq = m_Seq + 1

    Dim vp As UIAutomationClient.IUIAutomationValuePattern
    Set vp = m_Cmd.GetCurrentPattern(UIA_ValuePatternId)
    vp.SetValue CStr(m_Seq) & "|" & verb & "|" & argstr

    Dim ip As UIAutomationClient.IUIAutomationInvokePattern
    Set ip = m_Submit.GetCurrentPattern(UIA_InvokePatternId)
    ip.Invoke

    Dim t0 As Double
    t0 = Timer
    Do While Timer - t0 < ackTimeoutSec
        If UiaRead(m_Ack) = CStr(m_Seq) Then UiaSend = True: Exit Function
        SleepMs 1
    Loop
    UiaSend = False
    Exit Function
Fail:
    UiaSend = False
End Function

' 終了状態になるまで STATUS だけを見る。
' 固定長の待機は入れない。1 ms 刻み。取消が来たら CANCEL を送って待ち続ける。
Public Function UiaWaitTerminal(ByVal timeoutSec As Double) As String
    Dim t0 As Double, st As String
    Dim cancelSent As Boolean

    t0 = Timer
    Do While Timer - t0 < timeoutSec
        st = UiaStatus()
        If st = "DONE" Or st = "ERROR" Or st = "CANCELLED" Then
            UiaWaitTerminal = st
            Exit Function
        End If
        If g_Cancel And Not cancelSent Then
            cancelSent = True
            UiaSend "CANCEL", ""
        End If
        DoEvents
        SleepMs 1
    Loop
    UiaWaitTerminal = "TIMEOUT"
End Function

Public Sub UiaShutdown()
    On Error Resume Next
    If m_Win Is Nothing Then Exit Sub
    UiaSend "SHUTDOWN", "", 5
End Sub

Public Sub UiaReset()
    Set m_Win = Nothing:      Set m_Cmd = Nothing:    Set m_Submit = Nothing
    Set m_Ack = Nothing:      Set m_Status = Nothing: Set m_Progress = Nothing
    Set m_Result = Nothing:   Set m_Heart = Nothing:  Set m_Info = Nothing
    m_Seq = 0
End Sub

'==============================================================================
' 結果電文 "k=v;k=v" から 1 つ取り出す
'==============================================================================
Public Function PayloadValue(ByVal payload As String, ByVal key As String) As String
    Dim parts() As String, i As Long, p As Long
    parts = Split(payload, ";")
    For i = LBound(parts) To UBound(parts)
        p = InStr(1, parts(i), "=", vbBinaryCompare)
        If p > 1 Then
            If LCase$(Left$(parts(i), p - 1)) = LCase$(key) Then
                PayloadValue = Mid$(parts(i), p + 1)
                Exit Function
            End If
        End If
    Next i
    PayloadValue = ""
End Function
