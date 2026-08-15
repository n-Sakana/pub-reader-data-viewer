Attribute VB_Name = "modRdv3App"
'==============================================================================
' modRdv3App -- the FE: a small UI book and nothing else. Modeled on a proven
' Excel pump from an earlier project (the async skeleton only; none of its
' cursor work).
'
' The FE owns NO ledger data. The 100k-row ledger lives in its own workbook
' next to this one, owned entirely by the BE process; what reaches this book
' is display-sized: a record, a candidate list, a status line. The FE never
' materializes rows and never saves itself (there is nothing to save; close
' marks the book clean). What runs here:
'   - Workbook_Open boot (short: paint the UI, arm the pump, return). The
'     BE is started by the first tick: CreateObject on an invisible Excel,
'     open the worker book, call its bootstrap -- which only arms the BE's
'     own OnTime and returns (~1 s of FE occupancy, logged as spawn_ms)
'   - the OnTime pump: one short, bounded tick per second that reads the
'     aggregate channel file once and does at most one small render, then
'     returns. Between ticks the VBA stack is fully released. Excel natively
'     defers OnTime during cell edits and modal dialogs, so the pump never
'     collides with the user and simply catches up afterwards -- that is
'     normal behavior, not an error.
'   - click dispatch (Worksheet_FollowHyperlink): validates, writes a small
'     request file for the BE, arms a few fast follow-up ticks, returns
'   - the pump watchdog (Rdv3PumpEnsureArmed), called by ThisWorkbook's own
'     SheetChange / SheetActivate
'
' The event entry points live in the workbook's document modules (ThisWorkbook
' and the UI sheet) and do nothing but call in here; all of the logic, and the
' OnTime pump itself, stay in this module.
'
' Everything heavy -- CSV read, merge, comparison, carry-over, ledger
' write/save, "processed" persistence, search, the Notepad watch -- runs in
' the BE process (modRdv3Be). The BE never calls into this process; results
' arrive only through the file channel.
'
' The two figures the screen shows keep their boundaries:
'   マージ時間  = the 8 merge stages, measured inside the BE's rebuild
'                 (compose is logged separately and shown in the log total)
'   検索時間    = search confirmed (click t0 / BE detect t0) -> the FE has
'                 rendered the record or the candidate list. The cross-process
'                 hop and the pump cadence are honestly inside this figure.
'==============================================================================
Option Explicit

Private Const ST_BOOT As Long = 0
Private Const ST_CHECKING As Long = 1          ' BE spawning/merging/comparing
Private Const ST_APPLY_WAIT As Long = 2        ' approved; BE carries + writes + saves
Private Const ST_READY As Long = 3
Private Const ST_BLOCKED As Long = 4
Private Const ST_DEAD As Long = 5

Private Const PHASE_TIMEOUT_S As Double = 300#
Private Const PUMP_GRACE_S As Long = 5

' How long a single "processed" save may stay undecided before the FE calls it
' a failure. The BE persists one record per request (cell + Save + sidecar) and
' the worst measured case is the first mark of a session, which also opens the
' ledger workbook: ~21 s. This ceiling only exists so a BE that died mid-save
' cannot leave the workbook permanently unclosable.
Private Const MARK_TIMEOUT_S As Double = 180#

Private m_started As Boolean
Private m_state As Long
Private m_sid As String
Private m_dataDir As String
Private m_ledgerPath As String
Private m_logPath As String
Private m_beLogPath As String
Private m_readOnly As Boolean
Private m_logBroken As Boolean
Private m_spawnDone As Boolean            ' the first tick has started the BE

' last consumed channel version per kind
Private m_verCheck As Long
Private m_verReady As Long
Private m_verResult As Long
Private m_verMark As Long
Private m_verState As Long
Private m_verApply As Long
Private m_verErr As Long
Private m_reqVer As Long

' current display state
Private m_shownRow As Long
Private m_shownKey As String
Private m_shownKey2 As String
Private m_candRows(0 To 15) As Long
Private m_candCount As Long
Private m_candTotal As Long
Private m_watchOn As Boolean
Private m_lastMergeMs As Double

' pump state
Private m_pumpNext As Date
Private m_pumpArmed As Boolean
Private m_inTick As Boolean
Private m_fastFollow As Long
Private m_phaseStart As Double            ' Timer at phase entry, for timeouts
Private m_animStep As Long
Private m_tickCount As Long
Private m_tickMaxMs As Double
Private m_tickWorkCount As Long

Private m_readyNote As String
Private m_readyRows As Long

Private m_searchSeq As Long
Private m_procSeq As Long
Private m_lastWatchSt As String
Private m_closePrepared As Boolean
Private m_closePending As Boolean

' ---- the exit guard around one unfinished "processed" save ------------------
' The BE persists a mark the moment it is requested -- one record, its own
' workbook Save, its own sidecar write -- and answers on the MARK slot (marked
' or markerr), naming the request it answers. Its OWN slot is what makes that
' answer undroppable: a search answered in the same second lands in RESULT and
' cannot take its place (measured before the split: a confirmation was
' overwritten 441 ms later, the screen never saw it, and this guard then held
' the book for its full 180 s ceiling and called a SUCCESSFUL save undecided --
' work\race-evidence-before\).
' Between the request and that answer the outcome is not known, and
' closing the book in that window would tear the session down (stop flag,
' session files) without the operator ever learning whether the record was
' saved. So in that window the FE refuses a second mark AND refuses the close,
' both with the reason on screen. Nothing is queued or held back: this guards
' the single save that is already running on disk.
Private m_markPending As Boolean
Private m_markSince As Double
Private m_markKey2 As String
Private m_markReqVer As Long              ' the request this FE is waiting on
Private m_closeAskedWhileSaving As Boolean
Private m_bootT0 As Double                ' boot instant, for startup_ms
Private m_startupLogged As Boolean

'------------------------------------------------------------------------------
' log
'------------------------------------------------------------------------------
Private Sub AppLog(ByVal runId As String, ByVal section As String, ByVal detail As String)
    Dim f As Integer
    On Error GoTo Fail
    If Len(m_logPath) = 0 Then Exit Sub
    f = FreeFile
    Open m_logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & runId & vbTab & section & vbTab & _
        Replace(Replace(Replace(detail, vbTab, " "), vbCr, " "), vbLf, " ")
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
    If Not m_logBroken Then
        m_logBroken = True
        Rdv3UiError "ログを書き込めません: " & m_logPath
    End If
End Sub

Private Function FmtF(ByVal v As Double) As String
    FmtF = Format$(v, "0.00")
End Function

' diagnostic entry for responsiveness tests; no side effects, no pump calls
Public Function Rdv3Ping() As Double
    Rdv3Ping = Timer
End Function

' build-time compile probe: touches every module so a compile error anywhere
' fails the BUILD, not the first user launch
Public Function Rdv3BuildTouch() As String
    Dim s As String
    s = Rdv3ChRoot()
    s = s & "|" & Rdv3HostBeId()
    s = s & "|" & Rdv3UiInputKey()
    s = s & "|" & CStr(Rdv3IsKey("00000000"))
    s = s & "|" & Rdv3SidecarPath("x.xlsx")
    Rdv3BuildTouch = s
End Function

'------------------------------------------------------------------------------
' boot (Workbook_Open -> paint, arm the pump, return; the tick spawns the BE)
'------------------------------------------------------------------------------
Public Sub Rdv3AppStart()
    Dim n As Long

    If m_started Then Exit Sub
    m_started = True
    m_state = ST_BOOT
    m_shownRow = -1
    m_lastMergeMs = -1
    m_bootT0 = Rdv3Ticks()
    m_startupLogged = False

    m_sid = Format$(Now, "yyyymmddhhnnss") & Format$(CLng(Timer * 1000!) Mod 100000, "00000")
    m_readOnly = ThisWorkbook.ReadOnly
    ResolvePaths
    AppLog "-", "boot", "fe=" & Rdv3SelfId() & " method=vba-dict sid=" & m_sid & _
        " book=" & ThisWorkbook.FullName & " data=" & m_dataDir & _
        " ledger=" & m_ledgerPath & " readonly=" & CStr(m_readOnly)

    n = Rdv3ChSweepStale()
    If n > 0 Then AppLog "-", "worker", "stale_files_cleaned=" & n

    Rdv3UiClearResult
    Rdv3UiMergeMs -1
    Rdv3UiSearchMs -1
    Rdv3UiError ""
    Rdv3UiNotepad "未接続 -- メモ帳の入力欄をクリックすると接続します"
    Rdv3UiWatchButton True
    m_watchOn = True
    If m_readOnly Then
        Rdv3UiError "読み取り専用で開かれています (既に開いていませんか)。更新の承認と処理済み登録はできません"
    End If

    m_state = ST_CHECKING
    m_phaseStart = Timer
    m_spawnDone = False
    Rdv3UiState "更新を確認中"
    AppLog "R1", "decision", "check started"

    If Dir$(m_dataDir & "\tableA.csv") = "" Then
        AppLog "R1", "error", "stage=check msg=CSV not found in " & m_dataDir
        EnterDead "CSV が見つかりません: " & m_dataDir
        Exit Sub
    End If

    ' The BE is started by the FIRST PUMP TICK, not here: starting it means
    ' CreateObject + Workbooks.Open on this thread (~1 s, logged as spawn_ms),
    ' and doing that inside Workbook_Open would hold the book open with nothing
    ' on screen. By the tick the UI is painted and the window is up.
    PumpSchedule True
End Sub

Private Sub ResolvePaths()
    Dim meta As Object
    Dim v As String
    Set meta = ThisWorkbook.Worksheets(RDV3_SHEET_META)
    v = Trim$(CStr(meta.Range(RDV3_M_DATADIR).Value2))
    If Len(v) > 0 And Dir$(v & "\tableA.csv") <> "" Then
        m_dataDir = v
    Else
        m_dataDir = ThisWorkbook.Path & "\data"
    End If
    v = Trim$(CStr(meta.Range(RDV3_M_LOGPATH).Value2))
    If Len(v) > 0 Then
        m_logPath = v
    Else
        m_logPath = ThisWorkbook.Path & "\ReaderDataViewer.log"
    End If
    m_beLogPath = m_logPath & ".worker.log"
    m_ledgerPath = ThisWorkbook.Path & "\ReaderDataViewer-Ledger.xlsx"
End Sub

'------------------------------------------------------------------------------
' the pump (guard, one bounded body, reschedule, return)
'------------------------------------------------------------------------------
Private Sub PumpSchedule(ByVal firstTick As Boolean)
    If m_state = ST_DEAD Or m_state = ST_BLOCKED Then Exit Sub
    ' single-chain invariant: never schedule on top of a live schedule
    If m_pumpArmed Then Exit Sub
    If firstTick Then
        ' the first tick starts the BE, so it must not wait a second for it;
        ' OnTime with a due time of now fires as soon as Excel pumps, which is
        ' immediately after Workbook_Open returns
        m_pumpNext = Now
    ElseIf m_fastFollow > 0 Then
        m_fastFollow = m_fastFollow - 1
        m_pumpNext = Now
    Else
        m_pumpNext = Now + TimeSerial(0, 0, 1)
    End If
    Application.OnTime m_pumpNext, PumpQualified()
    m_pumpArmed = True
End Sub

Private Function PumpQualified() As String
    PumpQualified = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modRdv3App.Rdv3PumpTick"
End Function

' NOTE: there is deliberately NO OnTime cancellation anywhere in this module.
' Near the due time a cancel can fail -- or report success while the tick
' still fires -- and a tick that fires after the workbook closed makes Excel
' REOPEN the book to run the macro, behind a macro security notice that pins
' the process forever (measured). The chain is therefore only ever ended by
' letting the armed tick fire into a state that refuses to reschedule, and a
' close hands itself off to that tick (Rdv3AppPrepareClose / Rdv3FinishClose).

' Watchdog: if the chain ended without a schedule -- a tick error before its
' reschedule -- re-arm it. armed=True always means a real schedule exists
' (ticks clear the flag on entry), so re-arming is only ever done from the
' unarmed state; that keeps the chain strictly single.
'
' The trigger is THIS workbook's own SheetChange / SheetActivate, handled in
' ThisWorkbook and dispatched straight here (there used to be a WithEvents
' class listening to Application-level events and filtering them down to this
' book; the workbook's own events are already that filter). It is NOT a
' result-notification path: BE results arrive only through the file channel,
' pulled by the pump. Duplicate ticks are harmless -- every channel record
' dedups by version.
Public Sub Rdv3PumpEnsureArmed()
    If m_inTick Then Exit Sub
    If Not m_started Then Exit Sub
    If m_state = ST_DEAD Or m_state = ST_BLOCKED Then Exit Sub
    If m_pumpArmed Then Exit Sub
    On Error Resume Next
    PumpSchedule False
    On Error GoTo 0
End Sub

Public Sub Rdv3PumpTick()
    Dim t As Double
    Dim worked As Boolean
    Dim ms As Double

    If m_inTick Then Exit Sub
    m_inTick = True
    m_pumpArmed = False
    ' a close was intercepted because this very tick could not be canceled:
    ' now that it has fired the book can close for real, right after this call
    If m_closePending Then
        m_closePending = False
        m_inTick = False
        On Error Resume Next
        Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!modRdv3App.Rdv3FinishClose"
        Exit Sub
    End If
    On Error GoTo Done

    t = Rdv3Ticks()
    worked = TickBody()
    ms = Rdv3MsSince(t)
    m_tickCount = m_tickCount + 1
    If ms > m_tickMaxMs Then m_tickMaxMs = ms
    If worked Then
        m_tickWorkCount = m_tickWorkCount + 1
        AppLog "-", "pump", "tick_ms=" & FmtF(ms) & " state=" & CStr(m_state)
    ElseIf (m_tickCount Mod 60) = 0 Then
        AppLog "-", "pump", "stats ticks=" & m_tickCount & " worked=" & m_tickWorkCount & _
            " max_tick_ms=" & FmtF(m_tickMaxMs)
    End If

Done:
    If Err.Number <> 0 Then
        AppLog "-", "error", "stage=pump msg=" & Err.Number & " " & Err.Description
    End If
    On Error Resume Next
    PumpSchedule False
    m_inTick = False
    On Error GoTo 0
End Sub

Private Function TickBody() As Boolean
    Dim worked As Boolean
    worked = False

    ' the BE is started from the first tick (see Rdv3AppStart)
    If Not m_spawnDone And m_state = ST_CHECKING Then
        m_spawnDone = True
        If Not SpawnBe() Then Exit Function
        TickBody = True
        Exit Function
    End If

    Select Case m_state
        Case ST_CHECKING, ST_APPLY_WAIT
            AnimTick
            If DispatchChannel() Then worked = True
            If Not worked Then CheckBeHealth
        Case ST_READY
            If DispatchChannel() Then worked = True
            If m_markPending Then CheckMarkPending
    End Select
    TickBody = worked
End Function

Private Sub AnimTick()
    m_animStep = (m_animStep + 1) Mod 4
    Select Case m_state
        Case ST_CHECKING
            Rdv3UiState "更新を確認中" & String$(m_animStep, "・")
        Case ST_APPLY_WAIT
            Rdv3UiState "台帳を更新中" & String$(m_animStep, "・")
    End Select
End Sub

' start the BE, on this thread, from the first pump tick. Everything after the
' bootstrap call runs in the BE's own process; spawn_ms is the FE occupancy and
' is logged as such.
Private Function SpawnBe() As Boolean
    Dim errMsg As String
    Dim spawnMs As Double
    If Not Rdv3HostSpawn(m_sid, m_dataDir, m_ledgerPath, m_beLogPath, errMsg, spawnMs) Then
        AppLog "R1", "error", "stage=spawn msg=" & errMsg & " spawn_ms=" & FmtF(spawnMs)
        EnterDead "更新確認を開始できませんでした: " & errMsg
        Exit Function
    End If
    m_phaseStart = Timer                  ' the phase clock starts with the BE
    AppLog "R1", "worker", "spawn ok be=" & Rdv3HostBeId() & " fe=" & Rdv3SelfId() & _
        " sid=" & m_sid & " spawn_ms=" & FmtF(spawnMs)
    SpawnBe = True
End Function

' Liveness without a process id: the BE holds its lease lock for its whole
' life, so an unlocked (or missing) lease means it is gone, however it went.
Private Sub CheckBeHealth()
    Dim el As Double

    If Not Rdv3HostStarted() Then Exit Sub
    el = Timer - m_phaseStart
    If el < 0 Then el = el + 86400

    If Not Rdv3HostBeAlive() Then
        AppLog "-", "error", "stage=worker msg=BE process died"
        EnterDead "worker プロセスが終了しました。再起動してください"
        Exit Sub
    End If
    If el > PHASE_TIMEOUT_S Then
        AppLog "-", "timeout", "state=" & CStr(m_state) & " after_s=" & FmtF(el)
        StopBeNow
        EnterDead "更新確認がタイムアウトしました。再起動してください"
    End If
End Sub

'------------------------------------------------------------------------------
' channel dispatch: one aggregate read, then per-kind version dedup
'------------------------------------------------------------------------------
Private Function DispatchChannel() As Boolean
    Dim recs As Collection
    Dim rec As Variant
    Dim i As Long
    Dim worked As Boolean

    Set recs = Rdv3ChReadAgg(m_sid)
    For i = 1 To recs.Count
        rec = recs(i)
        Select Case CStr(rec(0))
            Case RDV3_K_CHECK
                If CLng(rec(1)) > m_verCheck Then
                    m_verCheck = CLng(rec(1))
                    HandleCheck CStr(rec(2))
                    worked = True
                End If
            Case RDV3_K_APPLY
                If CLng(rec(1)) > m_verApply Then
                    m_verApply = CLng(rec(1))
                    HandleApply CStr(rec(2))
                    worked = True
                End If
            Case RDV3_K_READY
                If CLng(rec(1)) > m_verReady Then
                    m_verReady = CLng(rec(1))
                    HandleReady CStr(rec(2))
                    worked = True
                End If
            Case RDV3_K_RESULT
                If CLng(rec(1)) > m_verResult Then
                    m_verResult = CLng(rec(1))
                    HandleResult CStr(rec(2)), rec(3)
                    worked = True
                End If
            ' the save confirmation has its own slot and its own counter, so a
            ' search answered in the same second neither replaces it nor delays
            ' it: both are read in this one pass and both are rendered
            Case RDV3_K_MARK
                If CLng(rec(1)) > m_verMark Then
                    m_verMark = CLng(rec(1))
                    HandleMark CStr(rec(2))
                    worked = True
                End If
            Case RDV3_K_STATE
                If CLng(rec(1)) > m_verState Then
                    m_verState = CLng(rec(1))
                    HandleState CStr(rec(2))
                End If
            Case RDV3_K_BEERR
                If CLng(rec(1)) > m_verErr Then
                    m_verErr = CLng(rec(1))
                    HandleBeErr CStr(rec(2))
                    worked = True
                End If
        End Select
    Next i
    DispatchChannel = worked
End Function

Private Sub HandleCheck(ByVal metaString As String)
    Dim d As Object
    Dim res As String
    Dim i As Long
    Dim ans As VbMsgBoxResult
    Set d = Rdv3ChMeta(metaString)
    res = CStr(d("res"))

    For i = 0 To RDV3_STAGES - 1
        Dim sec As String
        If i < 3 Then
            sec = "read"
        ElseIf i < 6 Then
            sec = "index"
        Else
            sec = "join"
        End If
        AppLog "R1", sec, Rdv3StageKey(i) & " ms=" & CStr(d(Rdv3StageKey(i)))
    Next i
    m_lastMergeMs = Val(CStr(d("merge")))
    AppLog "R1", "merge", "rows=" & CStr(d("rows")) & " matchedAB=" & CStr(d("mab")) & _
        " matchedBC=" & CStr(d("mbc")) & " checksum=" & CStr(d("chk")) & _
        " ms=" & CStr(d("merge")) & " compose_ms=" & CStr(d("compose"))
    AppLog "R1", "verify", "oracle=" & CStr(d("oracle"))
    If CStr(d("oracle")) = "ng" Then Rdv3UiError "検算 NG: 統合結果が expected.txt と一致しません"
    AppLog "R1", "load", "saved rows=" & CStr(d("srows")) & " ms=" & CStr(d("load")) & _
        " src=" & CStr(d("lsrc"))
    AppLog "R1", "compare", "result=" & res & IIf(d.Exists("fdiff"), " first_diff_row=" & CStr(d("fdiff")), "")
    Rdv3UiMergeMs m_lastMergeMs

    Select Case res
        Case "same"
            AppLog "R1", "decision", "no difference"
            ' READY follows from the BE on its own
        Case "diff"
            If m_readOnly Then
                MsgBox "CSV に変更がありますが、読み取り専用で開かれているため台帳を更新できません。" & vbCrLf & _
                       "保存済みの台帳のまま起動します。", vbOKOnly + vbExclamation, "更新の確認"
                AppLog "R1", "decision", "difference found; skipped (workbook is read-only)"
                SendReq RDV3_RQ_DECISION, "adopt"
            Else
                ans = MsgBox("CSV に変更があります。統合台帳を更新しますか?" & vbCrLf & _
                             "(処理済みは変更のないレコードへ引き継がれます)", _
                             vbYesNo + vbQuestion, "更新の確認")
                AppLog "R1", "decision", "difference found; update " & IIf(ans = vbYes, "approved", "rejected")
                If ans = vbYes Then
                    SendReq RDV3_RQ_DECISION, "apply"
                    m_state = ST_APPLY_WAIT
                    m_phaseStart = Timer
                Else
                    SendReq RDV3_RQ_DECISION, "adopt"
                End If
            End If
        Case "ledgerbad"
            Rdv3UiError "台帳を読み込めませんでした: " & CStr(d("lederr"))
            If m_readOnly Then
                MsgBox "保存済みの統合台帳が読めず、読み取り専用のため作り直せません。", vbOKOnly + vbExclamation, "更新の確認"
                AppLog "R1", "decision", "ledger unreadable; blocked (read-only)"
                SendReq RDV3_RQ_DECISION, "block"
            Else
                ans = MsgBox("保存済みの統合台帳が読めません:" & vbCrLf & CStr(d("lederr")) & vbCrLf & _
                             "CSV から作り直しますか? (処理済み状態は失われます)", _
                             vbYesNo + vbQuestion, "更新の確認")
                AppLog "R1", "decision", "ledger unreadable; rebuild " & IIf(ans = vbYes, "approved", "declined")
                If ans = vbYes Then
                    SendReq RDV3_RQ_DECISION, "apply"
                    m_state = ST_APPLY_WAIT
                    m_phaseStart = Timer
                Else
                    SendReq RDV3_RQ_DECISION, "block"
                End If
            End If
    End Select
End Sub

' the BE has carried, written and SAVED the ledger workbook itself; this is
' purely a figures record for the FE log
Private Sub HandleApply(ByVal metaString As String)
    Dim d As Object
    Set d = Rdv3ChMeta(metaString)
    AppLog "R1", "carry", "carried=" & CStr(d("carried")) & " reset=" & CStr(d("reset")) & _
        " new=" & CStr(d("new")) & " dropped=" & CStr(d("drop"))
    AppLog "R1", "persist", "ledger written by BE rows=" & CStr(d("rows")) & _
        " write_ms=" & CStr(d("write")) & " save_ms=" & CStr(d("save"))
End Sub

Private Sub HandleReady(ByVal metaString As String)
    Dim d As Object
    Set d = Rdv3ChMeta(metaString)
    m_readyNote = CStr(d("note"))
    m_readyRows = CLng(Val(CStr(d("rows"))))
    AppLog "R1", "index", "table=LEDGER rows=" & CStr(d("rows")) & " ms=" & CStr(d("idx")) & " (BE)"

    If m_readyNote = "blocked" Then
        EnterBlocked
        Exit Sub
    End If
    EnterReadyUi
End Sub

Private Function ReadyNoteText() As String
    Select Case m_readyNote
        Case "nodiff": ReadyNoteText = "更新はありません (台帳は最新です)"
        Case "applied": ReadyNoteText = "台帳を更新しました"
        Case "adopted": ReadyNoteText = "更新を見送りました (保存済み台帳のまま)"
        Case "savedfail": ReadyNoteText = "保存済み台帳のまま (更新確認は失敗)"
        Case Else: ReadyNoteText = m_readyNote
    End Select
End Function

Private Sub EnterReadyUi()
    m_state = ST_READY
    Rdv3UiLedgerInfo Format$(m_readyRows, "#,##0") & " 件   " & ReadyNoteText()
    Rdv3UiState IIf(m_watchOn, "監視中", "停止中 (監視再開で戻ります)")
    AppLog "R1", "decision", "ready rows=" & m_readyRows & " note=" & m_readyNote
    If Not m_startupLogged Then
        ' boot (Rdv3AppStart) -> operable. The FE book open before it and the
        ' Excel process start are outside; the bench measures those separately.
        m_startupLogged = True
        AppLog "R1", "startup", "boot_to_ready_ms=" & FmtF(Rdv3MsSince(m_bootT0))
    End If
End Sub

' BLOCKED/DEAD do not cancel the armed tick (cancellation is unreliable near
' the due time); the tick fires once more, does nothing, and the chain ends
' because PumpSchedule refuses these states.
Private Sub EnterBlocked()
    m_state = ST_BLOCKED
    Rdv3UiState "台帳がありません"
    Rdv3UiLedgerInfo "なし -- 検索できません"
    AppLog "R1", "decision", "blocked (no ledger)"
End Sub

Private Sub EnterDead(ByVal msg As String)
    m_state = ST_DEAD
    Rdv3UiError msg
    Rdv3UiState "worker 停止"
    AppLog "-", "decision", "dead: " & msg
End Sub

Private Sub HandleState(ByVal metaString As String)
    Dim d As Object
    Dim st As String
    Set d = Rdv3ChMeta(metaString)
    st = CStr(d("st"))
    If Left$(st, 3) = "hb_" Then st = Mid$(st, 4)
    Select Case st
        Case "bound"
            m_watchOn = True
            Rdv3UiWatchButton True
            Rdv3UiNotepad CStr(d("title")) & "  (hwnd " & CStr(d("hwnd")) & ")"
            ' a heartbeat must not paint over "処理済みを保存中"
            If m_state = ST_READY And Not m_markPending Then Rdv3UiState "監視中"
            If m_lastWatchSt <> "bound" Then
                AppLog "-", "watch", "bound hwnd=" & CStr(d("hwnd")) & " title=" & CStr(d("title"))
            End If
        Case "waiting"
            m_watchOn = True
            Rdv3UiWatchButton True
            Rdv3UiNotepad "未接続 -- " & CStr(d("why"))
            If m_state = ST_READY And Not m_markPending Then Rdv3UiState "メモ帳を待機中"
            If m_lastWatchSt <> "waiting" Then
                AppLog "-", "watch", "waiting why=" & CStr(d("why"))
            End If
        Case "watch_off"
            m_watchOn = False
            Rdv3UiWatchButton False
            If m_state = ST_READY And Not m_markPending Then Rdv3UiState "停止中 (監視再開で戻ります)"
            If m_lastWatchSt <> "watch_off" Then AppLog "-", "watch", "off"
    End Select
    If st = "bound" Or st = "waiting" Or st = "watch_off" Then m_lastWatchSt = st
End Sub

Private Sub HandleBeErr(ByVal metaString As String)
    Dim d As Object
    Set d = Rdv3ChMeta(metaString)
    Rdv3UiError "worker エラー (" & CStr(d("stage")) & "): " & CStr(d("msg"))
    AppLog "-", "error", "stage=be-" & CStr(d("stage")) & " msg=" & CStr(d("msg"))
End Sub

'------------------------------------------------------------------------------
' RESULT rendering (the end of the search-time boundary)
'------------------------------------------------------------------------------
Private Sub HandleResult(ByVal metaString As String, ByVal values As Variant)
    Dim d As Object
    Dim res As String
    Dim sid As String
    Dim elapsed As Double
    Dim vA(0 To 9) As String, vB(0 To 9) As String, vC(0 To 9) As String
    Dim i As Long, c As Long
    Dim rows() As Variant
    Dim n As Long, show As Long

    Set d = Rdv3ChMeta(metaString)
    res = CStr(d("res"))

    If m_state <> ST_READY Then
        AppLog "-", "search", "result skipped (state=" & CStr(m_state) & ") res=" & res
        Exit Sub
    End If

    m_searchSeq = m_searchSeq + 1
    sid = "S" & CStr(m_searchSeq)

    Select Case res
        Case "single", "picked"
            For i = 0 To 9
                vA(i) = SafeCell(values, i + 1, 1)
                vB(i) = SafeCell(values, i + 1, 2)
                vC(i) = SafeCell(values, i + 1, 3)
            Next i
            m_shownRow = CLng(Val(CStr(d("row"))))
            m_shownKey2 = CStr(d("k2"))
            If res = "single" Then m_shownKey = CStr(d("key"))
            m_candCount = 0
            Dim verdict As String
            If res = "picked" Then
                verdict = "一致 1 件   番号2 = " & m_shownKey2 & "   (候補 " & CStr(d("total")) & _
                          " 件中 " & CStr(d("slot")) & " 件目)"
            Else
                verdict = "一致 1 件   番号2 = " & m_shownKey2
            End If
            Rdv3UiShowRecord m_shownKey, verdict, vA, vB, vC, _
                "処理済み: " & IIf(CStr(d("proc")) = "1", "TRUE", "FALSE")
            elapsed = Rdv3MsBetween(CCur(Val(CStr(d("t0")))), Rdv3Ticks())
            If res = "picked" Then
                AppLog sid, "display", "picked=" & CStr(d("slot")) & " key2=" & m_shownKey2 & _
                    " proc=" & CStr(d("proc")) & " ms=" & FmtF(elapsed)
            Else
                Rdv3UiSearchMs elapsed
                AppLog sid, "search", "key=" & m_shownKey & " source=" & CStr(d("src")) & _
                    " hits=1 ms=" & FmtF(elapsed)
            End If
        Case "none"
            m_shownRow = -1
            m_shownKey = CStr(d("key"))
            m_candCount = 0
            Rdv3UiShowEmpty m_shownKey, "該当なし"
            elapsed = Rdv3MsBetween(CCur(Val(CStr(d("t0")))), Rdv3Ticks())
            Rdv3UiSearchMs elapsed
            AppLog sid, "search", "key=" & m_shownKey & " source=" & CStr(d("src")) & _
                " hits=0 ms=" & FmtF(elapsed)
        Case "multi"
            n = CLng(Val(CStr(d("hits"))))
            show = CLng(Val(CStr(d("shown"))))
            m_shownRow = -1
            m_shownKey = CStr(d("key"))
            m_candCount = show
            m_candTotal = n
            ReDim rows(1 To show, 1 To 8)
            For i = 1 To show
                For c = 1 To 8
                    rows(i, c) = SafeCell(values, i, c)
                Next c
                m_candRows(i - 1) = CLng(Val(SafeCell(values, i, 9)))
            Next i
            Dim note As String
            If show < n Then
                note = "候補 " & n & " 件中 " & show & " 件を表示"
            Else
                note = ""
            End If
            Rdv3UiShowCandidates m_shownKey, "候補 " & n & " 件 -- 行頭の「選択」で表示", rows, show, note
            elapsed = Rdv3MsBetween(CCur(Val(CStr(d("t0")))), Rdv3Ticks())
            Rdv3UiSearchMs elapsed
            AppLog sid, "search", "key=" & m_shownKey & " source=" & CStr(d("src")) & _
                " hits=" & n & " ms=" & FmtF(elapsed)
            AppLog sid, "candidate", "count=" & n & " shown=" & show
    End Select
End Sub

Private Function SafeCell(ByVal values As Variant, ByVal r As Long, ByVal c As Long) As String
    On Error GoTo Empty_
    If Not IsArray(values) Then GoTo Empty_
    If IsEmpty(values(r, c)) Then GoTo Empty_
    SafeCell = CStr(values(r, c))
    Exit Function
Empty_:
    SafeCell = ""
End Function

' The "processed" confirmation from the BE, on its own channel slot: the BE has
' updated the ledger workbook, SAVED it and rewritten the sidecar -- or
' explicitly failed. It carries req=<version of the request it answers>, and
' only the request this FE is actually waiting on is accepted: anything else
' is a stale or foreign record and is logged and dropped, never acted on.
Private Sub HandleMark(ByVal metaString As String)
    Dim d As Object
    Dim res As String
    Dim reqVer As Long

    Set d = Rdv3ChMeta(metaString)
    res = CStr(d("res"))
    reqVer = CLng(Val(CStr(d("req"))))

    If Not m_markPending Then
        AppLog "-", "processed", "confirmation ignored (no save is pending) res=" & res & _
            " req=" & CStr(reqVer)
        Exit Sub
    End If
    If reqVer <> m_markReqVer Then
        AppLog "-", "processed", "confirmation ignored (answers request " & CStr(reqVer) & _
            ", waiting for " & CStr(m_markReqVer) & ") res=" & res
        Exit Sub
    End If
    HandleMarkResult res, d
End Sub

Private Sub HandleMarkResult(ByVal res As String, ByVal d As Object)
    Dim tag As String
    Dim e2e As Double
    m_procSeq = m_procSeq + 1
    tag = "P" & CStr(m_procSeq)
    If res = "marked" Then
        If CLng(Val(CStr(d("row")))) = m_shownRow Then
            ThisWorkbook.Worksheets(RDV3_SHEET_UI).Range(RDV3_C_PROCSTATE).Value2 = _
                "処理済み: " & IIf(CStr(d("proc")) = "1", "TRUE", "FALSE")
        End If
        e2e = Rdv3MsBetween(CCur(Val(CStr(d("t0")))), Rdv3Ticks())
        AppLog tag, "processed", "key2=" & CStr(d("k2")) & " row=" & (CLng(Val(CStr(d("row")))) + 1) & _
            " value=" & IIf(CStr(d("proc")) = "1", "TRUE", "FALSE") & _
            " save_ms=" & CStr(d("save_ms")) & " e2e_ms=" & FmtF(e2e)
        EndMarkSave tag, True, ""
    Else
        If CLng(Val(CStr(d("row")))) = m_shownRow Then
            ThisWorkbook.Worksheets(RDV3_SHEET_UI).Range(RDV3_C_PROCSTATE).Value2 = "処理済み: FALSE"
        End If
        Rdv3UiError "処理済みを保存できませんでした: " & CStr(d("msg"))
        AppLog tag, "error", "stage=processed msg=" & CStr(d("msg"))
        EndMarkSave tag, False, CStr(d("msg"))
    End If
End Sub

' The save is decided -- saved or failed, either is a decision -- so the next
' mark and the close are released again.
Private Sub EndMarkSave(ByVal tag As String, ByVal ok As Boolean, ByVal why As String)
    If Not m_markPending Then Exit Sub
    m_markPending = False
    m_markKey2 = ""
    If m_state = ST_READY Then
        Rdv3UiState IIf(m_watchOn, "監視中", "停止中 (監視再開で戻ります)")
    End If
    AppLog tag, "exit", "processed save decided (" & IIf(ok, "saved", "failed") & _
        IIf(Len(why) > 0, ": " & why, "") & "); exit released"
    If m_closeAskedWhileSaving Then
        m_closeAskedWhileSaving = False
        If ok Then
            Rdv3UiError "処理済みの保存が完了しました。終了できます"
        Else
            Rdv3UiError "処理済みの保存は失敗として確定しました。終了できます"
        End If
    End If
End Sub

' While a mark is undecided the FE refuses to close, so it must be able to
' decide the mark itself when the BE can no longer answer: a dead BE or a save
' that overran the ceiling is reported as a FAILED save (explicitly, on screen
' and in the log), never as a silent success and never as a permanent hold.
Private Sub CheckMarkPending()
    Dim el As Double
    If Not m_markPending Then Exit Sub
    el = Timer - m_markSince
    If el < 0 Then el = el + 86400
    If Rdv3HostStarted() Then
        If Not Rdv3HostBeAlive() Then
            ThisWorkbook.Worksheets(RDV3_SHEET_UI).Range(RDV3_C_PROCSTATE).Value2 = "処理済み: 不明 (保存未確定)"
            Rdv3UiError "処理済みの保存を確認できませんでした: worker プロセスが終了しました。台帳を開いて確認してください"
            AppLog "-", "error", "stage=processed msg=BE died while the save was undecided key2=" & m_markKey2
            EndMarkSave "-", False, "worker died"
            Exit Sub
        End If
    End If
    If el > MARK_TIMEOUT_S Then
        ThisWorkbook.Worksheets(RDV3_SHEET_UI).Range(RDV3_C_PROCSTATE).Value2 = "処理済み: 不明 (保存未確定)"
        Rdv3UiError "処理済みの保存が " & Format$(el, "0") & " 秒たっても確定しません。台帳を開いて確認してください"
        AppLog "-", "timeout", "stage=processed after_s=" & FmtF(el) & " key2=" & m_markKey2
        EndMarkSave "-", False, "timeout"
    End If
End Sub

' Called from Workbook_BeforeClose BEFORE anything is torn down. True = refuse
' this close; the operator has just been told why.
Public Function Rdv3AppCloseHeldBySave() As Boolean
    If Not m_markPending Then Exit Function
    m_closeAskedWhileSaving = True
    Rdv3UiError "処理済みを保存中です。確定するまで終了できません"
    AppLog "-", "exit", "close refused: a processed save is still in flight (key2=" & m_markKey2 & ")"
    Rdv3AppCloseHeldBySave = True
    MsgBox "処理済みの保存がまだ確定していません。" & vbCrLf & _
           "保存が成功または失敗として確定してから、もう一度閉じてください。", _
           vbOKOnly + vbInformation, "保存中のため終了できません"
End Function

'------------------------------------------------------------------------------
' click dispatch (short: validate, write a request file, arm fast ticks)
'------------------------------------------------------------------------------
Private Sub SendReq(ByVal kind As String, ByVal args As String)
    m_reqVer = m_reqVer + 1
    If Not Rdv3ChWriteReq(m_sid, kind, m_reqVer, args) Then
        AppLog "-", "error", "stage=request msg=request write failed kind=" & kind
        Rdv3UiError "要求を送れませんでした (" & kind & ")"
    End If
End Sub

Public Sub Rdv3HandleClick(ByVal addr As String)
    Dim row As Long
    On Error GoTo Done
    Select Case addr
        Case RDV3_BTN_SEARCH
            DoManualSearch
        Case RDV3_BTN_CLEAR
            DoClear
        Case RDV3_BTN_PROCESSED
            DoProcessed
        Case RDV3_BTN_WATCH
            DoWatchToggle
        Case Else
            If Left$(addr, 1) = "B" And IsNumeric(Mid$(addr, 2)) Then
                row = CLng(Mid$(addr, 2))
                If row >= RDV3_CAND_TOP And row < RDV3_CAND_TOP + RDV3_MAXSHOW Then
                    DoPick row - RDV3_CAND_TOP
                End If
            End If
    End Select
Done:
    Rdv3PumpEnsureArmed
End Sub

Private Sub DoManualSearch()
    Dim key As String
    Dim t0 As Double
    t0 = Rdv3Ticks()
    If m_state <> ST_READY Then
        Rdv3UiError "更新確認が終わるまで操作できません"
        AppLog "-", "search", "ignored reason=not-ready source=manual"
        Exit Sub
    End If
    key = Rdv3UiInputKey()
    If Not Rdv3IsKey(key) Then
        Rdv3UiError "番号1 は 8 桁の数字で入力してください"
        AppLog "-", "search", "ignored key=" & key & " reason=bad-key source=manual"
        Exit Sub
    End If
    Rdv3UiError ""
    SendReq RDV3_RQ_SEARCH, key & "|" & Trim$(Str$(CDbl(t0))) & "|manual"
    AppLog "-", "search", "dispatched key=" & key & " source=manual"
    m_fastFollow = 6
End Sub

Private Sub DoClear()
    If m_state <> ST_READY And m_state <> ST_BLOCKED Then Exit Sub
    Rdv3UiClearResult
    Rdv3UiError ""
    m_shownRow = -1
    m_shownKey = ""
    m_shownKey2 = ""
    m_candCount = 0
    AppLog "-", "clear", "input and result cleared"
End Sub

' "processed": the FE only sends the request. The BE updates and saves the
' ledger workbook and the sidecar, then confirms on the MARK slot (marked /
' markerr); until then the screen honestly says the save is in flight.
Private Sub DoProcessed()
    Dim t0 As Double
    If m_state <> ST_READY Then
        Rdv3UiError "更新確認が終わるまで操作できません"
        Exit Sub
    End If
    ' one save at a time: a second mark while the first is undecided would ask
    ' the BE to rewrite the same workbook from under its own save
    If m_markPending Then
        Rdv3UiError "処理済みを保存中です。確定するまで次の操作はできません"
        AppLog "-", "processed", "refused: a processed save is still in flight (key2=" & m_markKey2 & ")"
        Exit Sub
    End If
    If m_shownRow < 0 Then
        Rdv3UiError "処理済みにする統合レコードが表示されていません"
        Exit Sub
    End If
    If m_readOnly Then
        Rdv3UiError "読み取り専用で開かれているため処理済みは登録できません"
        AppLog "-", "processed", "refused (workbook is read-only)"
        Exit Sub
    End If
    If MsgBox("表示中の統合レコード (番号2 = " & m_shownKey2 & ") を処理済みにします。よろしいですか?", _
              vbYesNo + vbQuestion, "処理済みの確認") <> vbYes Then
        AppLog "-", "processed", "declined key2=" & m_shownKey2
        Exit Sub
    End If
    Rdv3UiError ""
    t0 = Rdv3Ticks()
    ThisWorkbook.Worksheets(RDV3_SHEET_UI).Range(RDV3_C_PROCSTATE).Value2 = "処理済み: TRUE (保存中...)"
    ' the save is now running in the BE and its outcome is undecided: hold the
    ' exit and the next mark until MARK marked / markerr says which it was
    m_markPending = True
    m_markSince = Timer
    m_markKey2 = m_shownKey2
    m_closeAskedWhileSaving = False
    Rdv3UiState "処理済みを保存中"
    SendReq RDV3_RQ_MARK, CStr(m_shownRow) & "|1|" & Trim$(Str$(CDbl(t0)))
    ' the version SendReq just used: the confirmation must name it
    m_markReqVer = m_reqVer
    m_fastFollow = 6
    AppLog "-", "processed", "dispatched key2=" & m_shownKey2 & " row=" & CStr(m_shownRow + 1) & _
        " req=" & CStr(m_markReqVer) & " (exit held until it is decided)"
End Sub

Private Sub DoWatchToggle()
    If m_state <> ST_READY Then
        Rdv3UiError "更新確認が終わるまで操作できません"
        Exit Sub
    End If
    If m_watchOn Then
        SendReq RDV3_RQ_WATCH, "0"
        AppLog "-", "watch", "stop requested"
    Else
        SendReq RDV3_RQ_WATCH, "1"
        AppLog "-", "watch", "restart requested"
    End If
    m_fastFollow = 3
End Sub

Private Sub DoPick(ByVal slot As Long)
    Dim t0 As Double
    If m_state <> ST_READY Then Exit Sub
    If slot < 0 Or slot >= m_candCount Then Exit Sub
    t0 = Rdv3Ticks()
    SendReq RDV3_RQ_PICK, CStr(m_candRows(slot)) & "|" & CStr(slot + 1) & "|" & _
        CStr(m_candTotal) & "|" & Trim$(Str$(CDbl(t0)))
    AppLog "-", "display", "pick dispatched slot=" & CStr(slot + 1)
    m_candCount = 0        ' the picked record replaces the list when it arrives
    m_fastFollow = 6
End Sub

'------------------------------------------------------------------------------
' shutdown
'------------------------------------------------------------------------------
Private Sub StopBeNow()
    Dim forced As Boolean
    Dim note As String
    Rdv3HostStop m_sid, forced, note
    AppLog "-", "worker", "exit " & note & " forced=" & CStr(forced)
End Sub

' Called from Workbook_BeforeClose. Returns True when the workbook may close
' now; False when a due pump tick is still queued -- the caller must set
' Cancel = True, the tick fires within about a second, and Rdv3FinishClose
' closes the workbook for real. Everything else (BE stop, lease, session
' files, log) is torn down on the FIRST call either way.
Public Function Rdv3AppPrepareClose() As Boolean
    On Error Resume Next
    m_state = ST_DEAD          ' a tick that still fires does nothing and ends the chain
    If Not m_closePrepared Then
        m_closePrepared = True
        If m_started Then StopBeNow
        Rdv3ChReleaseLease
        Rdv3ChDeleteSession m_sid
        AppLog "-", "exit", "closing"
    End If
    ' the FE persists nothing: never let Excel raise a save prompt for the
    ' UI cell writes
    ThisWorkbook.Saved = True
    ' Deterministic hand-off: OnTime cancellation is not trustworthy near the
    ' due time (it can report success while the tick still fires - measured as
    ' a reopen + security-notice zombie). So whenever a tick is armed the
    ' close is deferred one cycle and the live tick itself finishes the close.
    ' armed=False is reliable the other way: ticks clear it on entry, so a
    ' crashed tick can never leave a phantom schedule behind.
    If m_pumpArmed Then
        m_closePending = True
        AppLog "-", "exit", "close deferred until the armed pump tick fires"
        Rdv3AppPrepareClose = False
    Else
        Rdv3AppPrepareClose = True
    End If
End Function

Public Sub Rdv3FinishClose()
    On Error Resume Next
    AppLog "-", "exit", "deferred close"
    ThisWorkbook.Saved = True
    ThisWorkbook.Close False
End Sub
