Attribute VB_Name = "modZipWorker"
'==============================================================================
' modZipWorker -- C# ワーカーを使う 5 方式 (3,4,5,6,8) と、ワーカー一式の用意。
'
' 5 方式が違うのは「どうやってプロセスを起こすか」だけ。
' 作ったファイル・起動する 1 本 (launch.bat)・UIA のやりとり・変換の中身・
' 計測範囲は同一。
' だから表の中で意味を持つのは起動手段の差だけになる。
'
'   3  launch.bat を人が開く                (Excel は一切起動に関与しない)
'   4  WScript.Shell.Run                    (CreateProcess 系)
'   5  VBA の Shell 関数                    (CreateProcess 系、Excel 組み込み)
'   6  Task Scheduler の COM (Schedule.Service)
'   8  Shell.Application.ShellExecute       (シェル経由・関連付け起動)
'
' 4/5 と 8 は別物として数えてよい。4/5 は Excel が直接子プロセスを作る。
' 8 はシェル (explorer) に「この .bat を open してくれ」と頼む道で、
' 起動元も、拡張子の関連付けを経由する点も違う。
'
' ワーカー一式には 2 系統ある
'   prebuilt  あらかじめ用意した prebuilt\ZipWorker.exe と worker\*.bat/.ps1 を
'             実行フォルダへコピーする。ビルドは起きない。
'   emitted   Excel が ZipWorker.cs / run_worker.bat / build_worker.ps1 を
'             その場で書き出す。exe が無いので run_worker.bat が csc.exe で
'             建ててから起こす。
' さらに「ファイル作成だけ」(起動も変換もしない) を別の入口として分けてある。
'==============================================================================
Option Explicit

Private Const LAUNCH_TIMEOUT_SEC As Double = 120#
Private Const MANUAL_TIMEOUT_SEC As Double = 300#
Private Const JOB_TIMEOUT_SEC    As Double = 900#

' run_worker.bat が引数なしで開かれたときの既定トークン。
' ZipBench 自身はもう使わない (どの方式も launch.bat 経由でトークンを渡す) が、
' 誰かが run_worker.bat を直接叩いたときの後片づけのために覚えておく。
Private Const MANUAL_TOKEN As String = "zipbench"

Private m_OwnToken  As String     ' 自分が起こしたワーカーだけを覚える
Private m_OwnTask   As String     ' 自分が登録したタスクだけを覚える

'==============================================================================
' 入口
'==============================================================================
Public Function RunWorkerMethod(ByVal methodNo As Long) As ZbResult
    Dim r As ZbResult
    Dim kind As String, mode As String
    Dim token As String, runDir As String, bat As String
    Dim tLaunch0 As Currency, tDispatch As Currency
    Dim waited As Double, manualWait As Double
    Dim st As String, payload As String
    Dim got As Variant
    Dim rowCount As Long
    Dim timeoutSec As Double

    r = modZipBench.ZB_NewResult(methodNo)
    kind = modZipBench.ZB_WorkerKind()
    mode = modZipBench.ZB_WorkerMode()
    r.WorkerKind = kind

    If methodNo = 3 Then
        token = NewToken()
        runDir = modZipBench.ZB_Work() & "\manual_run"
        timeoutSec = MANUAL_TIMEOUT_SEC
    Else
        token = NewToken()
        runDir = modZipBench.ZB_Work() & "\run_" & token
        timeoutSec = LAUNCH_TIMEOUT_SEC
    End If
    m_OwnToken = token
    bat = runDir & "\launch.bat"

    modZipUia.UiaReset

    ' この方式群は完了を Worksheet_Change で知る。だからイベントは入れておく。
    ' VBA からセルへは 1 つも書かないので、余計なイベントは起きない。
    modZipBench.ZB_EnterMeasure True
    modZipBench.ZbArmSignal

    On Error GoTo Fail
    tLaunch0 = modZipRule.QpcNow()

    '--- 作成 -----------------------------------------------------------------
    ProvisionRunDir runDir, kind, token, mode
    r.Created = "OK"

    ' 出力先を空にしておく (前回の結果が残っていると照合が意味を失う)
    modZipBench.ClearOutput

    '--- 起動 -----------------------------------------------------------------
    If methodNo = 3 Then
        ' Excel は起動に関わらない。人が開くのを待つ。
        ' 人の反応時間を計っても意味がないので、方式3 だけは
        ' 「UIA ツリーにワーカーが現れた瞬間」から総経過秒を計り直す。
        ' 待った時間はメモ欄に別に出す。
        modZipBench.ZB_LeaveMeasure
        If ManualPrompt() Then
            MsgBox "方式3 は Excel から起動しません。" & vbCrLf & vbCrLf & _
                   "このダイアログを閉じてから、下のファイルを手で開いてください。" & vbCrLf & vbCrLf & _
                   bat & vbCrLf & vbCrLf & _
                   "閉じると検知待ちに入ります (最大 " & CLng(timeoutSec) & " 秒)。", _
                   vbInformation, "ZipBench 方式3"
        End If
        modZipBench.ZB_EnterMeasure True
        r.Launched = "手動"
        If Not modZipUia.UiaWaitForWorker(token, timeoutSec, manualWait) Then
            r.UiaOk = "NG"
            r.Note = "手動起動を " & Format$(manualWait, "0.0") & " 秒待ったがワーカーが現れなかった"
            GoTo Finish
        End If
        r.UiaOk = "OK"
        manualWait = manualWait
        tLaunch0 = modZipRule.QpcNow()       ' ← 人待ちを除くためここから計り直す
        waited = 0
    Else
        LaunchWorker methodNo, bat, token, mode, runDir
        r.Launched = "OK"
        If Not modZipUia.UiaWaitForWorker(token, timeoutSec, waited) Then
            r.UiaOk = "NG"
            r.Note = "起動は成功したが " & Format$(waited, "0.0") & " 秒でワーカーが UIA ツリーに現れなかった"
            GoTo Finish
        End If
        r.UiaOk = "OK"
    End If

    ' ここまでが起動。処理E2E とは別枠。
    r.LaunchSec = modZipRule.QpcSince(tLaunch0)

    '--- 変換の指示 -----------------------------------------------------------
    ' UIA で渡すのは「どこにあるか」だけ。100 万件は UIA にも COM の引数にも
    ' 載せない。ワーカーが対象 Excel へ COM で接続して自分で読む。
    tDispatch = modZipRule.QpcNow()                 ' ← 処理E2E 開始
    If Not modZipUia.UiaSend("CONVERT", _
            "jobid=" & token & _
            ";xlhwnd=" & CStr(Application.Hwnd) & _
            ";wb=" & g_DataWb.Name & _
            ";master=" & DATA_MASTER & _
            ";input=" & DATA_INPUT & _
            ";output=" & DATA_OUTPUT & _
            ";masterrows=" & CStr(g_MasterN) & _
            ";rows=" & CStr(g_N) & _
            ";sigwb=" & ThisWorkbook.Name & _
            ";sigsheet=" & SHEET_SIGNAL & _
            ";sigcell=" & SIGNAL_CELL & _
            ";deliver=" & DeliverOf(methodNo) & _
            ";tsv=" & TsvPathOf(token)) Then
        r.UiaOk = "NG"
        r.Note = "CONVERT が ack されなかった"
        GoTo Finish
    End If

    '--- 完了を待つ -----------------------------------------------------------
    ' ここはワーカーへの問い合わせではない。逆で、ワーカーから飛んでくる
    ' COM の書き込みを Excel が処理できるように、メッセージを回しているだけ。
    ' ポンプを止めると相手の書き込みそのものが進まない。
    ' 完了は Worksheet_Change が教えてくれる。E2E の終了時刻もそこで取るので、
    ' このループの粗さは計測値に入らない。
    st = WaitForSignal(JOB_TIMEOUT_SEC)

    '--- ここで計測を確定する -------------------------------------------------
    If st = "DONE" Then
        r.TotalSec = modZipRule.QpcSec(tDispatch, g_SigStamp)
    Else
        r.TotalSec = modZipRule.QpcSince(tDispatch)
    End If

    ' 通知セルの電文: DONE|jobId|rows|hash|bind|master|input|dict|convert|write|workerE2E
    payload = g_SigPayload
    r.Rows = CLng(SigLong(payload, 3))
    r.Hash = SigPart(payload, 4)
    r.MasterSec = SigLong(payload, 6) / 1000#
    r.InputSec = SigLong(payload, 7) / 1000#
    r.DictSec = SigLong(payload, 8) / 1000#
    r.ConvertSec = SigLong(payload, 9) / 1000#
    r.WriteSec = SigLong(payload, 10) / 1000#
    ' 通知秒 = 処理E2E から工程の合計を引いた残り。通知セル書込のほか、
    ' UIA往復・STAスレッド起動・Excel接続・ハッシュ計算を含む未分解の区間。
    r.NotifySec = r.TotalSec - (r.MasterSec + r.InputSec + r.DictSec + r.ConvertSec + r.WriteSec)
    If r.NotifySec < 0 Then r.NotifySec = 0

    Select Case st
        Case "DONE"
            r.Converted = "OK"
        Case "CANCELLED"
            r.Converted = "NG"
            r.MatchText = "取消"
            r.Note = "取消: " & modZipUia.UiaResult()
            GoTo Finish
        Case Else
            r.Converted = "NG"
            r.Note = st & ": " & modZipUia.UiaResult()
            GoTo Finish
    End Select

    ' 候補2 はここで Excel への反映を行う。ワーカーは atomic rename 済みなので
    ' ファイルは必ず完成している。存在確認も待機もしない。
    ' ファイル経由の 4 方式はここで Excel への反映を行う。ワーカーは atomic rename
    ' 済みなのでファイルは必ず完成している。存在確認も待機もしない。
    ' 取り込みが失敗した場合も別方式へ切り替えず、その場の実エラーを記録して終わる。
    If DeliverOf(methodNo) = "tsv" Then
        On Error Resume Next
        Err.Clear
        Select Case methodNo
            Case 9:  r.ImportSec = modZipBench.ZB_ImportTsv(TsvPathOf(token))
            Case 13: r.ImportSec = modZipBench.ZB_ImportAdo(TsvPathOf(token))
            Case 14: r.ImportSec = modZipBench.ZB_ImportDao(TsvPathOf(token))
            Case 15: r.ImportSec = modZipBench.ZB_ImportOpenText(TsvPathOf(token))
        End Select
        If Err.Number <> 0 Then
            r.ErrNumber = Err.Number
            r.Converted = "NG"
            r.Note = "取込失敗 (" & ImportStageName(methodNo) & ") Err " & _
                     Err.Number & ": " & Err.Description
            Err.Clear
            On Error GoTo 0
            r.TotalSec = modZipRule.QpcSince(tDispatch)
            GoTo Finish
        End If
        On Error GoTo 0
        r.TotalSec = modZipRule.QpcSince(tDispatch)
        r.NotifySec = r.TotalSec - (r.MasterSec + r.InputSec + r.DictSec + r.ConvertSec + r.WriteSec + r.ImportSec)
        If r.NotifySec < 0 Then r.NotifySec = 0
    End If

    '--- ここから計測の外: OUTPUT を読んで正解と突き合わせる ------------------
    modZipBench.ZB_VerifyFromCells r

    r.Note = Trim$(LaunchName(methodNo) & " / " & kind & " / " & _
             "ワーカー内訳 " & Brief(modZipUia.UiaResult()))
    If methodNo = 3 Then
        r.Note = r.Note & " / 手動起動を待った時間 " & Format$(manualWait, "0.0") & " 秒 (起動秒にもE2Eにも含めない)"
    ElseIf waited > 0 Then
        r.Note = r.Note & " / UIA検知まで " & Format$(waited, "0.00") & " 秒"
    End If

Finish:
    If r.TotalSec = 0 Then r.TotalSec = modZipRule.QpcSince(tDispatch)
    On Error Resume Next
    modZipBench.ZbDisarmSignal
    modZipUia.UiaShutdown
    modZipUia.UiaReset
    m_OwnToken = ""
    DeleteOwnTask
    modZipBench.ZB_LeaveMeasure
    RunWorkerMethod = r
    Exit Function

Fail:
    r.ErrNumber = Err.Number
    If r.TotalSec = 0 Then r.TotalSec = modZipRule.QpcSince(tDispatch)
    If Err.Number = 18 Then
        r.Note = "Esc で取消"
        r.MatchText = "取消"
        g_Cancel = True
    Else
        r.Note = "Err " & Err.Number & ": " & Err.Description
    End If
    If r.Created = "-" Then r.Created = "NG"
    If r.Launched = "-" Then r.Launched = "NG"
    Resume Finish
End Function

'==============================================================================
' ファイル作成だけ。起動も変換もしない。
'   methodNo 9  = prebuilt 一式をコピーするだけ
'   methodNo 10 = Excel が BAT・PS1・C#ソースを書くだけ
'==============================================================================
Public Sub ZbEmitOnlyPrebuilt()
    RunEmitOnly 11, "prebuilt"
End Sub

Public Sub ZbEmitOnlyEmitted()
    RunEmitOnly 12, "emitted"
End Sub

Private Sub RunEmitOnly(ByVal methodNo As Long, ByVal kind As String)
    Dim r As ZbResult
    Dim runDir As String
    Dim t0 As Double, t1 As Double
    Dim names As String

    modZipBench.ZB_ClearRow methodNo
    r = modZipBench.ZB_NewResult(methodNo)
    r.WorkerKind = kind
    runDir = modZipBench.ZB_Work() & "\emit_" & kind

    modZipBench.ZB_EnterMeasure
    On Error GoTo Fail
    t0 = Timer
    names = ProvisionRunDir(runDir, kind, "sample", "offscreen")
    t1 = Timer

    r.Created = "OK"
    r.Launched = "-"
    r.UiaOk = "-"
    r.Converted = "-"
    r.TotalSec = t1 - t0
    r.ConvertSec = 0
    r.Rows = 0
    r.MatchText = "-"
    r.Outcome = "ファイル作成のみ完了"
    r.Note = runDir & " : " & names
    GoTo Finish

Fail:
    t1 = Timer
    r.ErrNumber = Err.Number
    r.TotalSec = t1 - t0
    r.Created = "NG"
    r.Outcome = "失敗 Err " & Err.Number
    r.Note = "Err " & Err.Number & ": " & Err.Description

Finish:
    modZipBench.ZB_LeaveMeasure
    modZipBench.ZB_WriteRow r
    modZipBench.ZB_FinalStatus r
    modZipBench.ZB_Log "ファイル作成だけ (" & kind & "): " & r.Outcome & " / " & r.Note
End Sub

'==============================================================================
' 実行フォルダの用意。戻り値は作ったファイル名の一覧。
'==============================================================================
' 実行フォルダには launch.bat も置く。中身は 1 行で、run_worker.bat を
' トークン・表示モード・親PID つきで呼ぶだけ。
'
' なぜ引数を直接渡さず、こんな 1 行を挟むのか
' ------------------------------------------------------------------
' Shell.Application.ShellExecute は .bat に対して第2引数(vArgs)を渡さない。
' 実測すると、ワーカーは引数を 1 つも受け取らず run_worker.bat の既定値で
' 起動し、ウィンドウ名が ZIPWORKER::zipbench になっていた。Excel はその実行の
' トークンで探しているので、当然ながら永久に見つからない。
' (worker.log には「起動した」と出るのに UIA では見つからない、という
'  紛らわしい失敗の正体がこれだった。)
'
' 起動手段ごとに引数の渡り方が違うのを許すと、比べているのが
' 「起動できるか」ではなく「引数が届くか」になってしまう。だから
' 引数はファイルに焼き込み、どの起動手段も「この 1 本を起こすだけ」に揃える。
Private Function LaunchBatText(ByVal token As String, ByVal mode As String, ByVal parentPid As String) As String
    Dim s As String
    s = "@echo off" & vbCrLf
    s = s & "rem ZipBench launcher -- generated per run. ASCII only." & vbCrLf
    s = s & "rem The run parameters are baked in because not every launch" & vbCrLf
    s = s & "rem mechanism forwards arguments to a .bat (Shell.Application" & vbCrLf
    s = s & "rem .ShellExecute does not), and every method must start the" & vbCrLf
    s = s & "rem worker with exactly the same parameters to be comparable." & vbCrLf
    s = s & "call ""%~dp0run_worker.bat"" " & token & " " & mode & " " & parentPid & vbCrLf
    LaunchBatText = s
End Function

Private Function ProvisionRunDir(ByVal runDir As String, ByVal kind As String, _
                                 ByVal token As String, ByVal mode As String) As String
    Dim src As String

    modZipRule.EnsureFolder runDir

    If kind = "prebuilt" Then
        ' あらかじめ用意した一式をコピーする。ここでビルドは起きない。
        src = modZipBench.ZB_Root()
        If Not modZipRule.FileExists(src & "\prebuilt\ZipWorker.exe") Then
            Err.Raise 53, "ProvisionRunDir", "prebuilt\ZipWorker.exe が無い"
        End If
        FileCopy src & "\prebuilt\ZipWorker.exe", runDir & "\ZipWorker.exe"
        FileCopy src & "\worker\run_worker.bat", runDir & "\run_worker.bat"
        FileCopy src & "\worker\build_worker.ps1", runDir & "\build_worker.ps1"
        modZipRule.WriteTextAscii runDir & "\launch.bat", _
            LaunchBatText(token, mode, CStr(modZipUia.GetCurrentProcessId()))
        ProvisionRunDir = "ZipWorker.exe, run_worker.bat, build_worker.ps1 (コピー) + launch.bat"
    Else
        ' Excel が書く。exe は置かないので run_worker.bat が csc.exe で建てる。
        If modZipRule.FileExists(runDir & "\ZipWorker.exe") Then Kill runDir & "\ZipWorker.exe"
        modZipRule.WriteTextUtf8Bom runDir & "\ZipWorker.cs", modZipEmit.WorkerCsText()
        modZipRule.WriteTextAscii runDir & "\run_worker.bat", modZipEmit.RunBatText()
        modZipRule.WriteTextAscii runDir & "\build_worker.ps1", modZipEmit.BuildPs1Text()
        modZipRule.WriteTextAscii runDir & "\launch.bat", _
            LaunchBatText(token, mode, CStr(modZipUia.GetCurrentProcessId()))
        ProvisionRunDir = "ZipWorker.cs, run_worker.bat, build_worker.ps1, launch.bat (Excel が生成)"
    End If
End Function

'==============================================================================
' 起動手段。ここだけが方式ごとに違う。
' 失敗しても他の手段へ切り替えない。失敗は失敗として持ち帰る。
'==============================================================================
Private Sub LaunchWorker(ByVal methodNo As Long, ByVal bat As String, _
                         ByVal token As String, ByVal mode As String, ByVal runDir As String)
    Dim cmdline As String

    Select Case methodNo
        Case 4, 9, 10, 13, 14, 15
            ' WScript.Shell.Run  第2引数 0 = 非表示、第3引数 False = 待たない
            ' 結果の届け方を比べる方式 (9, 10, 13, 14, 15) は起動手段を揃えてある。
            ' 比べたいのは届け方だけなので、起動の差を比較から外す。
            CreateObject("WScript.Shell").Run _
                """" & bat & """", 0, False

        Case 5
            ' VBA の Shell。.bat は実行可能ファイルではないので cmd.exe を通す。
            cmdline = "cmd.exe /c """"" & bat & """"""
            Shell cmdline, vbHide

        Case 6
            LaunchViaTaskScheduler bat, runDir

        Case 8
            ' シェルに「この .bat を open してくれ」と頼む。
            ' 起動元が explorer 側になる点が 4/5 と違う。0 = 非表示。
            CreateObject("Shell.Application").ShellExecute _
                bat, "", runDir, "open", 0

        Case Else
            Err.Raise 5, "LaunchWorker", "起動手段が未定義: " & methodNo
    End Select
End Sub

'------------------------------------------------------------------------------
' Task Scheduler COM。
' 登録 -> 実行 -> (後で) 削除。トリガーの無い「要求時のみ」タスクにする。
' TASK_LOGON_INTERACTIVE_TOKEN で登録するので、ログオン中の対話セッションで
' 起動される。そうでないと別セッションになり UIA ツリーから見えない。
' 管理者権限は要らない (自分のユーザーのタスクだから)。
'------------------------------------------------------------------------------
Private Sub LaunchViaTaskScheduler(ByVal bat As String, ByVal runDir As String)
    Const TASK_CREATE_OR_UPDATE      As Long = 6
    Const TASK_LOGON_INTERACTIVE_TOKEN As Long = 3
    Const TASK_ACTION_EXEC           As Long = 0

    Dim svc As Object, folder As Object, td As Object, act As Object, task As Object
    Dim taskName As String

    taskName = "ZipBench_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(modZipUia.GetCurrentProcessId())

    Set svc = CreateObject("Schedule.Service")
    svc.Connect
    Set folder = svc.GetFolder("\")

    Set td = svc.NewTask(0)
    td.RegistrationInfo.Description = "ZipBench temporary launcher (auto-deleted)"
    td.RegistrationInfo.Author = "ZipBench"
    td.Settings.Enabled = True
    td.Settings.Hidden = False
    td.Settings.DisallowStartIfOnBatteries = False
    td.Settings.StopIfGoingOnBatteries = False
    td.Settings.StartWhenAvailable = True
    td.Settings.ExecutionTimeLimit = "PT30M"

    Set act = td.Actions.Create(TASK_ACTION_EXEC)
    act.Path = "cmd.exe"
    act.Arguments = "/c """"" & bat & """"""
    act.WorkingDirectory = runDir

    folder.RegisterTaskDefinition taskName, td, TASK_CREATE_OR_UPDATE, _
        vbNullString, vbNullString, TASK_LOGON_INTERACTIVE_TOKEN
    m_OwnTask = taskName

    Set task = folder.GetTask(taskName)
    task.Run vbNullString
End Sub

' 自分が登録したタスクだけを消す。他人のタスクには触らない。
Private Sub DeleteOwnTask()
    Dim svc As Object, folder As Object
    On Error Resume Next
    If Len(m_OwnTask) = 0 Then Exit Sub
    Set svc = CreateObject("Schedule.Service")
    svc.Connect
    Set folder = svc.GetFolder("\")
    folder.DeleteTask m_OwnTask, 0
    m_OwnTask = ""
    Err.Clear
End Sub

'==============================================================================
' 後始末。自分が起こしたワーカーだけを、UIA 経由で終わらせる。
' プロセス一覧を舐めて名前で殺すようなことはしない。
'==============================================================================
Public Sub ShutdownOwnWorker()
    On Error Resume Next
    If Len(m_OwnToken) > 0 Then
        If modZipUia.UiaBind(m_OwnToken) Then modZipUia.UiaShutdown
    End If
    ' 方式3 は固定トークンなので、残っていれば同じく畳む
    If Len(MANUAL_TOKEN) > 0 Then If modZipUia.UiaBind(MANUAL_TOKEN) Then modZipUia.UiaShutdown
    modZipUia.UiaReset
    m_OwnToken = ""
    DeleteOwnTask
    Err.Clear
End Sub

'==============================================================================
' 完了待ち。
'
' ワーカーは 100 万件をこの Excel のセルへ COM で書き込み、最後に通知セルを
' 1 つ書く。その書き込みは Excel のメインスレッドが処理するので、こちらが
' DoEvents を回していないと相手の書き込み自体が進まない。つまりこのループは
' 「相手の様子を見に行く」ためではなく「相手を通す」ために回している。
'
' 完了の判定は Worksheet_Change が立てるフラグだけを見る。
' UIA へは一切問い合わせない。
'==============================================================================
Private Function WaitForSignal(ByVal timeoutSec As Double) As String
    Dim t0 As Double
    Dim cancelSent As Boolean

    t0 = Timer
    Do
        If g_SigFired Then WaitForSignal = "DONE": Exit Function
        If g_Cancel And Not cancelSent Then
            cancelSent = True
            modZipUia.UiaSend "CANCEL", ""
        End If
        DoEvents
    Loop While Timer - t0 < timeoutSec

    If cancelSent Then WaitForSignal = "CANCELLED" Else WaitForSignal = "TIMEOUT"
End Function

' 通知セルの電文は "DONE|rows|hash|convertMs|writeMs|bindMs"
Private Function SigPart(ByVal payload As String, ByVal index As Long) As String
    Dim p() As String
    If Len(payload) = 0 Then Exit Function
    p = Split(payload, "|")
    If index - 1 <= UBound(p) Then SigPart = p(index - 1)
End Function

Private Function SigLong(ByVal payload As String, ByVal index As Long) As Double
    Dim s As String
    s = SigPart(payload, index)
    If Len(s) > 0 And IsNumeric(s) Then SigLong = CDbl(s)
End Function

'==============================================================================
' 候補2 と候補3 の違いはここだけ。
' 接続・Master読込・Input読込・辞書構築・変換は同じコードが走る。
'   候補2 (方式9)  ワーカーが TSV を一時名で書き、完成後 atomic rename。
'                  Excel は QueryTable でネイティブに取り込む。
'   候補3 (方式10) ワーカーが Output.Value2 へ 1 回だけ一括書込。
'==============================================================================
' 方式 9・13・14・15 はワーカーに同じファイルを書かせ、Excel 側の取り込み経路だけを
' 変える。方式 10 だけがファイルを介さず COM で直接セルへ書く。
Private Function DeliverOf(ByVal methodNo As Long) As String
    Select Case methodNo
        Case 9, 13, 14, 15: DeliverOf = "tsv"
        Case Else:          DeliverOf = "com"
    End Select
End Function

' 拡張子が .txt なのは ACE のテキスト ドライバが .tsv を認識しないため。
' 中身はタブ区切りのままで、方式 9・13・14・15 で 1 バイトも違わない。
Private Function TsvPathOf(ByVal token As String) As String
    TsvPathOf = modZipBench.ZB_Work() & "\out_" & token & ".txt"
End Function

Private Function ImportStageName(ByVal methodNo As Long) As String
    Select Case methodNo
        Case 9:  ImportStageName = "QueryTable"
        Case 13: ImportStageName = "ADODB + CopyFromRecordset"
        Case 14: ImportStageName = "DAO + CopyFromRecordset"
        Case 15: ImportStageName = "Workbooks.OpenText"
        Case Else: ImportStageName = "?"
    End Select
End Function

'==============================================================================
' 小物
'==============================================================================
Private Function NewToken() As String
    Randomize
    NewToken = "Z" & Format$(Now, "yyyymmddhhnnss") & _
               Right$("00000" & CStr(Int(Rnd() * 100000)), 5)
End Function

Private Function LaunchName(ByVal methodNo As Long) As String
    Select Case methodNo
        Case 3: LaunchName = "手動でBATを開く"
        Case 4: LaunchName = "WScript.Shell.Run"
        Case 5: LaunchName = "VBA Shell"
        Case 6: LaunchName = "Task Scheduler COM"
        Case 8: LaunchName = "Shell.Application.ShellExecute"
        ' 9・13・14・15 はワーカーが書いた同一ファイルを、Excel 側の別経路で取り込む。
        ' 10 だけがファイルを介さない。起動手段はこの 5 方式とも揃えてある。
        Case 9:  LaunchName = "TSV+rename → QueryTable (起動は WScript.Shell.Run)"
        Case 10: LaunchName = "COM Range.Value2 一括 (起動は WScript.Shell.Run)"
        Case 13: LaunchName = "TSV+rename → ADO + CopyFromRecordset (起動は WScript.Shell.Run)"
        Case 14: LaunchName = "TSV+rename → DAO + CopyFromRecordset (起動は WScript.Shell.Run)"
        Case 15: LaunchName = "TSV+rename → Workbooks.OpenText (起動は WScript.Shell.Run)"
        Case Else: LaunchName = "?"
    End Select
End Function

' 結果電文から表に載せたい内訳だけ拾う
Private Function Brief(ByVal payload As String) As String
    Brief = "rows=" & modZipUia.PayloadValue(payload, "rows") & _
            " bind=" & modZipUia.PayloadValue(payload, "bindMs") & "ms" & _
            " master=" & modZipUia.PayloadValue(payload, "masterMs") & "ms" & _
            " input=" & modZipUia.PayloadValue(payload, "inputMs") & "ms" & _
            " dict=" & modZipUia.PayloadValue(payload, "dictMs") & "ms" & _
            " convert=" & modZipUia.PayloadValue(payload, "convertMs") & "ms" & _
            " write=" & modZipUia.PayloadValue(payload, "writeMs") & "ms" & _
            " workerE2E=" & modZipUia.PayloadValue(payload, "workerE2eMs") & "ms" & _
            " notFound=" & modZipUia.PayloadValue(payload, "notFound")
End Function

' 方式3 の案内ダイアログを出すか (無人実行では切れるようにしてある)
Private Function ManualPrompt() As Boolean
    Dim v As String
    On Error Resume Next
    v = Trim$(CStr(modZipBench.ZB_Sheet().Range("B8").Value))
    ManualPrompt = (v <> "非表示")
End Function
