VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbBeSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbBeSession - 非表示の別プロセス Excel（BE）の一生
'
' 責務：BE を起こし、終われと言い、終わったことを見届ける。BE が何を計算するか
' は知らない（それは相手側の modPbBackend）。
'
' 起こし方は、BE の起動後に COM 参照を捨てる形にしてある。
'   CreateObject → Visible / DisplayAlerts を決める → AutomationSecurity を
'   落として開く → 戻す → 入口を名前で呼ぶ → 参照を捨てる
'
' 4 つの判断がここに埋まっている。
'
' 1. **FE は BE への COM 参照を持ち続けない。** 忙しい相手を COM で叩くと、
'    FE がサーバービジーで固まる。起動して入口を呼んだら、その場で手放す。
'    以後のやりとりはファイルだけ。
'
' 2. **入口は名前で直接呼ぶ。** RunAutoMacros で Auto_Open を起こす手もあるが、
'    相手側で何かあったときに返ってこないまま止まった。名前で呼べば失敗が
'    COM のエラーとして返る。入口は OnTime を仕込むだけなのですぐ返る。
'
' 3. **BE 用のコピーは画面を作る前に取る。** 組み上がった盤面ごと複製すると
'    94KB が 557KB になり、その読み込みで FE が詰まる。
'
' 4. **状態の文字列は 3 文字まで。** 盤面の状態欄（ProcBox の右側）は幅 54 設計
'    px の結合セルで、右揃え。溢れた分は隣へ流れず **先頭から** 切れる。実測：
'    「待機 非表示」は先頭の「待」が欠けて「寺機 非表示」に見えていた（3 文字の
'    「動作中」は余裕をもって収まる）。詳しい事情は記録・ScreenTip・ジョブの行が
'    持つので、ここは短い名札に徹する。状態は「待機」と「動作中」の 2 つだけに
'    そろえる（起動中・終了中・作成といった途中の名前は、見る側には区別が
'    要らない上に 3 つの箱で言葉がばらつく）。例外は「失敗」だけで、これは
'    握りつぶさず出す。
'
' 終わらせるときは quit をファイルで送るだけ。プロセスは殺さない。返事（bye）
' は次のティック以降で待つ。ここでブロックすると FE が数秒固まる。
'
' Temp のやりとり（置き場所・読み書き・掃除）は PbBenchmarkRun が持っている。
' 1 回の run と同じ寿命の資源だからで、BE の一生を見るこのクラスは、そこを
' 通してだけファイルへ触る。
'==============================================================================
Option Explicit

Private Const STOP_TICKS_MAX As Long = 8

Private m_run As PbBenchmarkRun
Private m_log As PbLog
Private m_started As Boolean
Private m_status As String
Private m_stopping As Boolean
Private m_stopTicks As Long
Private m_lastProg As String

Public Sub Init(ByVal run As PbBenchmarkRun, ByVal logger As PbLog)
    Set m_run = run
    Set m_log = logger
    m_status = "待機"
End Sub

Public Property Get IsStarted() As Boolean
    IsStarted = m_started
End Property

Public Property Get IsStopping() As Boolean
    IsStopping = m_stopping
End Property

Public Property Get StatusText() As String
    StatusText = m_status
End Property

Public Property Let StatusText(ByVal s As String)
    m_status = s
End Property

' reuseCopy: 途中で起こし直すとき True。起動時に取ったコピーをそのまま使う。
Public Function Start(Optional ByVal reuseCopy As Boolean = False) As Boolean
    Dim beApp As Object
    Dim beBook As Object
    Dim prevSec As Long
    Dim copyPath As String

    On Error GoTo Failed
    m_run.EnsureFolder
    m_run.KillQuiet m_run.CmdPath
    m_run.KillQuiet m_run.PiPath
    m_run.KillQuiet m_run.ProgPath
    m_run.Heartbeat

    copyPath = m_run.BeBookPath
    If Not (reuseCopy And Len(Dir$(copyPath)) > 0) Then
        m_run.KillQuiet copyPath
        ThisWorkbook.SaveCopyAs copyPath
    End If

    Set beApp = CreateObject("Excel.Application")
    beApp.Visible = False
    beApp.DisplayAlerts = False
    beApp.ScreenUpdating = False
    beApp.EnableEvents = False
    prevSec = beApp.AutomationSecurity
    beApp.AutomationSecurity = 1                  ' msoAutomationSecurityLow
    Set beBook = beApp.Workbooks.Open(copyPath, 0, True)
    beApp.AutomationSecurity = prevSec
    m_log.Note "be: opened, calling bootstrap"
    beApp.Run "'" & Replace(beBook.Name, "'", "''") & "'!PbBeBootstrap"
    m_log.Note "be: bootstrap returned"
    ' ここで手放す（上の 1 を参照）
    Set beBook = Nothing
    Set beApp = Nothing
    m_started = True
    m_stopping = False
    m_status = "待機"
    Start = True
    Exit Function
Failed:
    m_started = False
    m_status = "失敗"
    m_log.Note "be: FAILED " & Err.Number & " " & Err.Description
    Start = False
End Function

' 終われと言う。返事は待たない。
Public Sub RequestQuit()
    If Not m_started Then Exit Sub
    m_run.SendCommand "quit", ""
    m_stopping = True
    m_stopTicks = 0
    m_status = "待機"
End Sub

' ベンチ以外のティックで呼ぶ。生存の印を置き、BE の様子を見る。
' 戻り値 True = 片づけまで済んだ（Temp を消した）。
Public Function PumpIdle() As Boolean
    Dim js As String

    On Error Resume Next
    If m_started Then m_run.Heartbeat

    js = m_run.ReadProgress()
    If Len(js) > 0 And js <> m_lastProg Then
        m_lastProg = js
        ' BE は FE が黙って見えると自分から終わる（置き去り防止）。それを見たら、
        ' 次のベンチで起こし直せるようにこちら側の印も倒す。
        If JVal(js, "state") = "bye" Then
            m_started = False
            If Not m_stopping Then m_status = "待機"
        End If
    End If

    If m_stopping Then
        m_stopTicks = m_stopTicks + 1
        ' bye が見えたか、待っても来ないなら、そこで片づける。
        If Not m_started Or m_stopTicks > STOP_TICKS_MAX Then
            m_stopping = False
            m_started = False
            m_run.CleanAll
            m_lastProg = ""
            m_status = "待機"
            PumpIdle = True
        End If
    End If
End Function

' 閉じるときだけ、少しは待つ。相手のプロセスは殺さない。
Public Sub WaitForBye(ByVal maxSeconds As Long)
    Dim k As Long
    On Error Resume Next
    If Not m_started Then Exit Sub
    For k = 1 To maxSeconds
        If InStr(1, m_run.ReadProgress(), """bye""") > 0 Then Exit For
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next k
    m_log.Note "quit: be waited " & k
    m_started = False
End Sub

Public Sub ForgetProgress()
    m_lastProg = ""
End Sub
