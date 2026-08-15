Attribute VB_Name = "modRdv3Be"
'==============================================================================
' modRdv3Be -- the BE: everything heavy, in a dedicated invisible Excel this
' app started. Modeled on a proven resident worker from an earlier project: a
' self-pacing resident loop in its own process, publishing results to the file
' channel. THE BE NEVER CALLS INTO THE FE PROCESS -- not one COM call, not one
' cell write. A busy or cell-editing FE is therefore completely undisturbed.
'
' The BE owns THE WHOLE LEDGER LIFE CYCLE. The FE never touches the ledger:
'   - the ledger workbook (ReaderDataViewer-Ledger.xlsx, LEDGER + META sheets)
'     and its sidecar state file (<ledger>.state): read, compare, carry,
'     write, save -- all in this process
'   - a no-difference boot reads only the sidecar (a full TSV mirror of the
'     saved ledger written atomically at every save) and never opens the
'     workbook at all; the workbook is opened lazily for apply / mark, and as
'     the slow self-heal path when the sidecar is missing or out of step
'   - the merge (modRdv3Engine, Scripting.Dictionary indexes)
'   - the content comparison (full line-by-line equality, never time/size)
'     and the processed carry-over (identity = key2 + identical content)
'   - the key-1 search index and every search / pick / candidate answer
'   - the "processed" persistence: workbook cell + Save + sidecar, then a
'     confirmation RESULT; the FE only ever sent the request
'   - the Notepad watch (modRdv3Uia), polled at 40 ms inside the loop
'
' Publishes (atomic, version-tagged; see modRdv3Chan):
'   CHECK   merge stages + comparison outcome        -> FE decides via MsgBox
'   APPLY   carry stats + ledger write/save figures  -> FE logs them
'   READY   active rows + note                       -> FE enables operations
'   RESULT  search / pick render block                -> FE renders
'   MARK    "that one record is on disk" (or failed) -> FE releases the guard.
'           Its own slot, and it carries req=<version of the request it
'           answers>, so a search answered a moment later can neither replace
'           it nor be mistaken for it.
'   STATE   watch binding / heartbeat                -> FE status line
'   BEERR   a failure, explicitly                    -> FE error line
'
' Consumes (latest-wins request files): decision / search / pick / mark / watch.
' Liveness: stop flag, fe_gone tombstone, and the FE lease lock probe; any of
' them failing ends the loop and the BE quits itself (a crashed FE releases
' the lease lock instantly, so no PID guessing is ever needed).
'
' This module may reference modRdv3Spec, modRdv3Chan, modRdv3Engine and
' modRdv3Uia only: it must compile inside the worker book.
'==============================================================================
Option Explicit

' TIME-based cadences. v1 counted loop iterations, but one iteration is
' 40 ms + the UIA poll, which can stretch past 200 ms on a loaded machine --
' 25 iterations then took longer than the FE's stop-wait and the FE reached
' its last-resort kill (measured: forced=True). Seconds do not stretch.
Private Const REQ_POLL_S As Double = 0.12
Private Const LIVE_POLL_S As Double = 1#
Private Const HEARTBEAT_S As Double = 3#

Private g_sid As String
Private g_dataDir As String
Private g_ledgerPath As String
Private g_sidecarPath As String
Private g_logPath As String
Private g_active As Boolean
Private g_state As String                          ' CHECKING/AWAIT/READY/BLOCKED
Private g_ver As Long
Private g_exitReason As String

' active ledger (BE memory)
Private m_lines() As String
Private m_proc() As Boolean
Private m_rows As Long
Private m_dict As Object
Private m_savedOk As Boolean
Private m_savedErr As String
Private m_savedLoadMs As Double
Private m_stateSrc As String                       ' "sidecar" / "book" / ""
Private m_marks As Long                            ' persisted mark counter
Private m_lb As Object                             ' the ledger workbook, open lazily (RW)
Private m_saveMethod As Long                       ' how a mark reaches the file

' request bookkeeping (latest processed version per kind)
Private m_qDecision As Long
Private m_qSearch As Long
Private m_qPick As Long
Private m_qMark As Long
Private m_qWatch As Long

' watch state
Private w_on As Boolean
Private w_pending As String
Private w_pendSince As Double
Private w_pendPolls As Long
Private w_lastFired As String
Private w_sawEmpty As Boolean
Private w_bound As Boolean
Private w_saidWaiting As Boolean

Private m_expRows As Long
Private m_expChk As Double
Private m_expLoaded As Boolean

' build-time compile probe (also forces this module's dependencies to compile)
Public Function Rdv3BeIsActive() As Boolean
    Rdv3BeIsActive = g_active
End Function

' build-time initial ledger: merge the shipped CSVs and write the ledger
' workbook + sidecar through the EXACT code paths the product uses at apply
' time. Called by the builder inside its own Excel instance.
Public Function Rdv3BeBuildInitial(ByVal dataDir As String, ByVal ledgerPath As String) As String
    Dim i As Long
    Dim errMsg As String
    Dim saveMs As Double
    On Error GoTo Fail
    g_dataDir = dataDir
    g_ledgerPath = ledgerPath
    g_sidecarPath = Rdv3SidecarPath(ledgerPath)
    g_logPath = ""
    If Not Rdv3EngMerge(dataDir) Then
        Rdv3BeBuildInitial = "ERR merge: " & Rdv3EngErrText()
        Exit Function
    End If
    m_rows = Rdv3EngLineCount()
    If m_rows > 0 Then
        ReDim m_lines(0 To m_rows - 1)
        ReDim m_proc(0 To m_rows - 1)
        For i = 0 To m_rows - 1
            m_lines(i) = Rdv3EngLine(i)
        Next i
    End If
    m_marks = 0
    If Not WriteLedgerBookAll(errMsg, saveMs) Then
        Rdv3BeBuildInitial = "ERR book: " & errMsg
        Exit Function
    End If
    If Not SidecarWrite(errMsg) Then
        Rdv3BeBuildInitial = "ERR sidecar: " & errMsg
        Exit Function
    End If
    If Not m_lb Is Nothing Then
        m_lb.Close False
        Set m_lb = Nothing
    End If
    Rdv3BeBuildInitial = "ok rows=" & CStr(m_rows) & " checksum=" & Format$(Rdv3EngChecksum(), "0") & _
        " save_ms=" & Format$(saveMs, "0.0")
    Exit Function
Fail:
    Rdv3BeBuildInitial = "ERR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not m_lb Is Nothing Then m_lb.Close False
    Set m_lb = Nothing
End Function

'------------------------------------------------------------------------------
' bootstrap (called by the FE via Application.Run; returns immediately)
'------------------------------------------------------------------------------
Public Sub Rdv3BeBootstrap(ByVal sid As String, ByVal dataDir As String, _
                           ByVal ledgerPath As String, ByVal logPath As String)
    On Error GoTo Failed
    g_sid = sid
    g_dataDir = dataDir
    g_ledgerPath = ledgerPath
    g_sidecarPath = Rdv3SidecarPath(ledgerPath)
    g_logPath = logPath
    ' how a "processed" mark reaches the file. book unless a "<ledger>.savemethod"
    ' file next to the ledger names another one -- and no distribution ships
    ' that file, so what ships is unchanged (modRdv3Save).
    m_saveMethod = Rdv3SaveMethodFromFile(ledgerPath)
    g_active = True
    g_state = "CHECKING"
    g_ver = 0
    g_exitReason = ""
    Application.DisplayAlerts = False
    ' The FE drops its COM references as soon as this call returns, so this
    ' instance is a ref-count-zero automation server. Excel arms an automatic
    ' shutdown for those and fires it at the next message pump -- measured: a
    ' large Workbooks.Open mid-BootWork pumped, and the process tore itself
    ' down cleanly, no dialog, no crash event. UserControl = True marks the
    ' instance self-owned and disables that shutdown; this process ends only
    ' through its own SelfQuit. It is set HERE, inside the call the FE is still
    ' waiting on, so there is no unprotected gap.
    Application.UserControl = True

    ' the lease the FE probes for liveness: open before returning, so the FE
    ' can never read "no lease" and conclude the BE died before it started
    If Not Rdv3ChBeLeaseOpen(sid) Then
        Rdv3ChWriteFlag Rdv3ChBeDonePath(sid), "reason=be_lease_failed"
        Application.DisplayAlerts = False
        Application.Quit
        Exit Sub
    End If

    BeLog "bootstrap sid=" & sid & " be=" & Rdv3SelfId() & " data=" & dataDir & _
        " ledger=" & ledgerPath & " savemethod=" & Rdv3SaveMethodName(m_saveMethod)
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!Rdv3BeMain"
    Exit Sub
Failed:
    On Error Resume Next
    Rdv3ChWriteFlag Rdv3ChBeDonePath(sid), "reason=bootstrap_error_" & CStr(Err.Number)
    Application.DisplayAlerts = False
    Application.Quit
End Sub

'------------------------------------------------------------------------------
' the resident loop (each iteration is bounded)
'------------------------------------------------------------------------------
Public Sub Rdv3BeMain()
    Dim iter As Long
    On Error GoTo Failed

    ' the lease opened in the bootstrap is the FE's liveness signal; this is
    ' only the diagnostic note that the resident loop itself is now running
    BeLog "loop start be=" & Rdv3SelfId()

    ' This loop paces itself with the sub-second wait; if that primitive is not
    ' delivering waits, say so and stop -- never fall back to a spin. Tested
    ' here rather than in the bootstrap so the FE is not held for it.
    Dim waitWhy As String
    If Not Rdv3WaitReady(waitWhy) Then
        BeLog "FATAL wait primitive unavailable: " & waitWhy
        PublishErr "wait", "待機処理を初期化できません: " & waitWhy
        g_exitReason = "wait_unavailable"
        SelfQuit
        Exit Sub
    End If

    PublishState "boot", ""
    BootWork

    Dim lastReq As Double, lastLive As Double, lastHb As Double, nowT As Double
    lastReq = Timer: lastLive = Timer: lastHb = Timer
    Do While g_active
        If w_on And g_state = "READY" Then PollWatch
        nowT = Timer
        If nowT < lastReq Then lastReq = nowT     ' midnight wrap
        If nowT < lastLive Then lastLive = nowT
        If nowT < lastHb Then lastHb = nowT
        If nowT - lastReq >= REQ_POLL_S Then
            lastReq = nowT
            PollRequests
        End If
        If nowT - lastLive >= LIVE_POLL_S Then
            lastLive = nowT
            If Not Alive() Then Exit Do
        End If
        If nowT - lastHb >= HEARTBEAT_S Then
            lastHb = nowT
            Heartbeat
        End If
        If Not Rdv3WaitMs(RDV3_POLL_MS) Then
            g_exitReason = "wait_unavailable"
            PublishErr "wait", "待機処理が機能しなくなりました (空回りを避けるため停止します)"
            Exit Do
        End If
    Loop
    SelfQuit
    Exit Sub
Failed:
    Dim n As Long, d As String
    n = Err.Number
    d = Err.Description
    On Error Resume Next
    PublishErr "loop", "BE error " & n & ": " & d
    BeLog "FATAL " & n & ": " & d
    If Len(g_exitReason) = 0 Then g_exitReason = "be_error_" & CStr(n)
    SelfQuit
End Sub

Private Function Alive() As Boolean
    If Rdv3ChFlagExists(Rdv3ChStopPath(g_sid)) Then
        g_exitReason = "stop_requested"
        Exit Function
    End If
    If Rdv3ChFlagExists(Rdv3ChGonePath(g_sid)) Then
        g_exitReason = "fe_gone"
        Exit Function
    End If
    If Not Rdv3ChFeLeaseAlive(g_sid) Then
        g_exitReason = "fe_lease_lost"
        Exit Function
    End If
    Alive = True
End Function

Private Sub SelfQuit()
    On Error Resume Next
    If Not m_lb Is Nothing Then
        m_lb.Close False                     ' every change was saved when made
        Set m_lb = Nothing
    End If
    If Len(g_exitReason) = 0 Then g_exitReason = "stopped"
    BeLog "self-quit reason=" & g_exitReason
    Rdv3ChWriteFlag Rdv3ChBeDonePath(g_sid), "reason=" & g_exitReason & vbCrLf & _
        "closed_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Rdv3ChBeLeaseRelease                     ' the FE's liveness probe, closed cleanly
    ThisWorkbook.Saved = True
    Application.DisplayAlerts = False
    Application.Quit
End Sub

'------------------------------------------------------------------------------
' boot: merge, load the saved state (sidecar fast path), compare, publish
'------------------------------------------------------------------------------
Private Sub BootWork()
    Dim t As Double
    Dim same As Boolean
    Dim firstDiff As Long
    Dim meta As String
    Dim i As Long
    Dim oracle As String

    LoadExpected

    ' merge FIRST, on a clean heap: opening a large workbook beforehand leaves
    ' Excel's allocator so fragmented that the same pure-VBA merge measured
    ' ~5x slower (7.8 s vs 1.4 s on this data). The merge figure on screen
    ' must be the engine's cost, not the allocator's.
    PublishState "merging", ""
    If Not Rdv3EngMerge(g_dataDir) Then
        PublishErr "merge", Rdv3EngErrText()
        PublishState "loading", ""
        t = Rdv3Ticks()
        m_savedOk = LoadState(m_savedErr)
        m_savedLoadMs = Rdv3MsSince(t)
        If m_savedOk Then
            AdoptActive "savedfail"
        Else
            PublishReady 0, 0, "blocked"
            g_state = "BLOCKED"
        End If
        Exit Sub
    End If
    BeLog "merge ok rows=" & CStr(Rdv3EngRows()) & " ms=" & Format$(Rdv3EngMergeMs(), "0.0") & _
        " compose_ms=" & Format$(Rdv3EngComposeMs(), "0.0") & _
        " checksum=" & Format$(Rdv3EngChecksum(), "0")

    PublishState "loading", ""
    t = Rdv3Ticks()
    m_savedOk = LoadState(m_savedErr)
    m_savedLoadMs = Rdv3MsSince(t)
    BeLog "saved state loaded ok=" & CStr(m_savedOk) & " src=" & m_stateSrc & _
        " rows=" & CStr(m_rows) & " ms=" & Format$(m_savedLoadMs, "0.0") & _
        IIf(m_savedOk, "", " err=" & m_savedErr)

    oracle = "skip"
    If m_expLoaded Then
        If m_expRows = Rdv3EngRows() And m_expChk = Rdv3EngChecksum() Then
            oracle = "ok"
        Else
            oracle = "ng"
        End If
    End If

    firstDiff = -1
    If m_savedOk Then
        PublishState "comparing", ""
        t = Rdv3Ticks()
        same = CompareSavedToNew(firstDiff)
        BeLog "compare same=" & CStr(same) & " first_diff=" & CStr(firstDiff + 1) & _
            " ms=" & Format$(Rdv3MsSince(t), "0.0")
    Else
        same = False
    End If

    meta = "res=" & IIf(m_savedOk, IIf(same, "same", "diff"), "ledgerbad")
    For i = 0 To RDV3_STAGES - 1
        meta = meta & ";" & Rdv3StageKey(i) & "=" & Format$(Rdv3EngStageMs(i), "0.00")
    Next i
    meta = meta & ";merge=" & Format$(Rdv3EngMergeMs(), "0.00")
    meta = meta & ";compose=" & Format$(Rdv3EngComposeMs(), "0.00")
    meta = meta & ";rows=" & CStr(Rdv3EngRows())
    meta = meta & ";mab=" & CStr(Rdv3EngMatchedAB())
    meta = meta & ";mbc=" & CStr(Rdv3EngMatchedBC())
    meta = meta & ";chk=" & Format$(Rdv3EngChecksum(), "0")
    meta = meta & ";oracle=" & oracle
    meta = meta & ";load=" & Format$(m_savedLoadMs, "0.00")
    meta = meta & ";lsrc=" & m_stateSrc
    meta = meta & ";srows=" & CStr(m_rows)
    If Not m_savedOk Then meta = meta & ";lederr=" & Replace(m_savedErr, ";", ",")
    If firstDiff >= 0 Then meta = meta & ";fdiff=" & CStr(firstDiff + 1)

    Publish RDV3_K_CHECK, meta, Empty

    If m_savedOk And same Then
        ' no difference: the saved content is the active ledger; no decision
        AdoptActive "nodiff"
    Else
        ' diff / unreadable: the FE asks the operator and answers by request
        g_state = "AWAIT"
    End If
End Sub

'------------------------------------------------------------------------------
' saved state: the ledger workbook is the authority (it is what a person can
' open and read); the sidecar is its full TSV mirror written atomically at
' every save. The fast path trusts the sidecar ONLY while the workbook's
' mtime still matches the mtime the sidecar recorded at save time -- any
' mismatch (a crash between the two writes, a hand-replaced book) falls back
' to reading the workbook itself and re-writing the sidecar. The guard can
' only ever cause an extra verification, never a false "no difference".
'------------------------------------------------------------------------------
Private Function LoadState(ByRef errMsg As String) As Boolean
    Dim scErr As String
    errMsg = ""
    m_stateSrc = ""
    m_rows = 0
    m_marks = 0

    If Dir$(g_ledgerPath) = "" Then
        errMsg = "台帳ブックが見つかりません: " & g_ledgerPath
        Exit Function
    End If

    If LoadStateFromSidecar(scErr) Then
        m_stateSrc = "sidecar"
        LoadState = True
        Exit Function
    End If
    BeLog "sidecar unusable (" & scErr & "); reading the ledger workbook"
    If LoadStateFromBook(errMsg) Then
        m_stateSrc = "book"
        LoadState = True
    End If
End Function

Private Function FileTimeStr(ByVal path As String) As String
    On Error Resume Next
    FileTimeStr = Format$(FileDateTime(path), "yyyy-mm-dd hh:nn:ss")
    On Error GoTo 0
End Function

Private Function LoadStateFromSidecar(ByRef errMsg As String) As Boolean
    Dim txt As String
    Dim lines As Variant
    Dim head As Variant
    Dim rows As Long
    Dim i As Long
    Dim ln As String
    Dim d As Object

    errMsg = ""
    If Dir$(g_sidecarPath) = "" Then
        errMsg = "sidecar missing"
        Exit Function
    End If
    On Error GoTo Fail
    ReadStateFile g_sidecarPath, txt
    lines = Split(txt, vbCrLf)
    txt = ""                                   ' the lines array owns it now
    If UBound(lines) < 0 Then
        errMsg = "sidecar empty"
        Exit Function
    End If
    head = Split(CStr(lines(0)), vbTab)
    If UBound(head) < 1 Then
        errMsg = "sidecar header short"
        Exit Function
    End If
    If CStr(head(0)) <> "rdv3state" Or CStr(head(1)) <> "1" Then
        errMsg = "sidecar magic/version mismatch"
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(head)
        Dim eq As Long
        eq = InStr(CStr(head(i)), "=")
        If eq > 1 Then d(Left$(CStr(head(i)), eq - 1)) = Mid$(CStr(head(i)), eq + 1)
    Next i
    If Not d.Exists("rows") Or Not d.Exists("bookmtime") Then
        errMsg = "sidecar header incomplete"
        Exit Function
    End If
    If CStr(d("bookmtime")) <> FileTimeStr(g_ledgerPath) Then
        errMsg = "ledger workbook mtime differs from the sidecar record"
        Exit Function
    End If
    rows = CLng(Val(CStr(d("rows"))))
    If rows < 0 Or UBound(lines) < rows Then
        errMsg = "sidecar row count mismatch"
        Exit Function
    End If
    If d.Exists("marks") Then m_marks = CLng(Val(CStr(d("marks"))))

    m_rows = rows
    If rows > 0 Then
        ReDim m_lines(0 To rows - 1)
        ReDim m_proc(0 To rows - 1)
        For i = 0 To rows - 1
            ln = CStr(lines(1 + i))
            If Len(ln) < 2 Then
                errMsg = "sidecar row " & (i + 1) & " short"
                m_rows = 0
                Exit Function
            End If
            m_proc(i) = (Left$(ln, 1) = "1")
            m_lines(i) = Mid$(ln, 3)
        Next i
    Else
        Erase m_lines
        Erase m_proc
    End If
    LoadStateFromSidecar = True
    Exit Function
Fail:
    errMsg = "sidecar read error " & Err.Number & ": " & Err.Description
    m_rows = 0
End Function

' slow self-heal path: read the ledger workbook itself (read-only), then
' rewrite the sidecar so the next boot takes the fast path again
Private Function LoadStateFromBook(ByRef errMsg As String) As Boolean
    Dim wb As Object
    Dim ws As Object
    Dim mt As Object
    Dim v As Variant
    Dim rows As Long
    Dim c As Long
    Dim arr As Variant
    Dim r0 As Long, r1 As Long, i As Long, n As Long
    Dim parts(0 To RDV3_CONTENT_COLS - 1) As String
    Dim scErr As String

    errMsg = ""
    LoadStateFromBook = False
    m_rows = 0
    On Error GoTo Fail

    Set wb = Application.Workbooks.Open(g_ledgerPath, 0, True)

    Set ws = wb.Worksheets(RDV3_LB_SHEET)
    Set mt = wb.Worksheets(RDV3_LB_META)

    For c = 1 To RDV3_LED_COLS
        If CStr(ws.Cells(1, c).Value2) <> Rdv3LedgerHead(c - 1) Then
            errMsg = "台帳ヘッダーが一致しません (列 " & c & ")"
            GoTo CloseOut
        End If
    Next c
    v = mt.Range(RDV3_LM_ROWS).Value2
    If IsEmpty(v) Or Not IsNumeric(v) Then
        errMsg = "台帳 META の件数が読めません"
        GoTo CloseOut
    End If
    rows = CLng(v)
    If rows < 0 Then
        errMsg = "台帳 META の件数が不正です"
        GoTo CloseOut
    End If
    v = mt.Range(RDV3_LM_MARKS).Value2
    If IsNumeric(v) Then m_marks = CLng(v)

    m_rows = rows
    If rows > 0 Then
        ReDim m_lines(0 To rows - 1)
        ReDim m_proc(0 To rows - 1)
        r0 = 0
        Do While r0 < rows
            r1 = r0 + 16384
            If r1 > rows Then r1 = rows
            arr = ws.Range(ws.Cells(2 + r0, 1), ws.Cells(1 + r1, RDV3_LED_COLS)).Value2
            n = r1 - r0
            For i = 1 To n
                For c = 0 To RDV3_CONTENT_COLS - 1
                    If IsEmpty(arr(i, c + 2)) Or IsError(arr(i, c + 2)) Then
                        parts(c) = ""
                    Else
                        parts(c) = CStr(arr(i, c + 2))
                    End If
                Next c
                m_lines(r0 + i - 1) = Join(parts, vbTab)
                If VarType(arr(i, 1)) = vbBoolean Then
                    m_proc(r0 + i - 1) = arr(i, 1)
                Else
                    m_proc(r0 + i - 1) = (UCase$(CStr(arr(i, 1))) = "TRUE")
                End If
            Next i
            r0 = r1
        Loop
    Else
        Erase m_lines
        Erase m_proc
    End If
    LoadStateFromBook = True

CloseOut:
    On Error Resume Next
    wb.Close False
    Set wb = Nothing
    On Error GoTo 0
    If LoadStateFromBook Then
        If SidecarWrite(scErr) Then
            BeLog "sidecar healed from the workbook"
        Else
            BeLog "sidecar heal FAILED: " & scErr
        End If
    End If
    Exit Function
Fail:
    errMsg = "台帳読込エラー " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

'------------------------------------------------------------------------------
' sidecar IO. The sidecar is UTF-16LE, no BOM: the ledger content comes from
' UTF-8 CSVs, so a CP932 Print# could not round-trip every character, and the
' conversion has to be one bulk operation rather than a per-character loop.
'
' v2 did that bulk conversion with MultiByteToWideChar and kept the file in
' UTF-8. With Win32 gone the file itself carries VBA's own string encoding
' instead, and then no conversion exists at all: VB's Byte()<->String
' assignment copies the raw UTF-16LE bytes, and Get/Put move them straight to
' and from disk. Measured side by side, same 22 MB body, same process, warm:
'
'   UTF-16LE  Get + Byte()->String        23.4 ms   (this build)
'   UTF-8     Get + MultiByteToWideChar   26.0 ms   (what it replaces)
'
' So dropping Win32 costs nothing at the conversion. What it does cost is size:
' the file is twice as big (44.6 MB), and in the app that shows up as about
' 200 ms more in the state-load stage (386 -> 590 ms), because 22 MB more has
' to be faulted in right after the merge has filled the heap. Both COM routes
' were measured on the same body and are unusable: ADODB.Stream.ReadText(all)
' took 22.2 SECONDS, and mscorlib's System.Text.UTF8Encoding cannot even be
' created from VBA on this machine (CreateObject raises 0x80131500).
'------------------------------------------------------------------------------
' ByRef, not a String return: the body is ~45 MB and a returned String would be
' copied once more into the caller. The byte buffer is freed before the caller
' starts splitting, so the peak stays at two copies instead of three.
Private Sub ReadStateFile(ByVal path As String, ByRef s As String)
    Dim f As Integer
    Dim b() As Byte
    Dim n As Long
    s = ""
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n < 2 Then
        Close #f
        Exit Sub
    End If
    If (n Mod 2) <> 0 Then n = n - 1          ' truncated tail: drop the odd byte
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f
    s = b                                      ' raw UTF-16LE bytes -> String
    Erase b
End Sub

Private Function WriteStateFileAtomic(ByVal path As String, ByRef txt As String, _
                                      ByRef errMsg As String) As Boolean
    Dim f As Integer
    Dim b() As Byte
    Dim tmp As String
    errMsg = ""
    tmp = path & ".tmp"
    On Error GoTo Fail
    If Len(Dir$(tmp)) > 0 Then Kill tmp
    f = FreeFile
    Open tmp For Binary Access Write As #f
    If Len(txt) > 0 Then
        b = txt                                ' String -> raw UTF-16LE bytes
        Put #f, 1, b
    End If
    Close #f
    f = 0
    ' same replace rule as the channel: park the live sidecar by rename, never
    ' delete it before the new one is in place (modRdv3Chan.Rdv3ChReplaceFile)
    If Not Rdv3ChReplaceFile(tmp, path) Then
        errMsg = "sidecar replace failed (err " & CStr(Rdv3ChLastReplaceErr()) & ")"
        Exit Function
    End If
    WriteStateFileAtomic = True
    Exit Function
Fail:
    errMsg = "state write error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If f <> 0 Then Close #f
    If Len(Dir$(tmp)) > 0 Then Kill tmp
End Function

' full mirror: header, then "<0|1><TAB><28 content columns>" per row. Written
' AFTER the workbook save so bookmtime matches the file it describes.
Private Function SidecarWrite(ByRef errMsg As String) As Boolean
    Dim body() As String
    Dim i As Long
    Dim head As String
    If m_rows > 0 Then
        ReDim body(0 To m_rows - 1)
        For i = 0 To m_rows - 1
            body(i) = IIf(m_proc(i), "1", "0") & vbTab & m_lines(i)
        Next i
    End If
    head = "rdv3state" & vbTab & "1" & vbTab & "rows=" & CStr(m_rows) & _
           vbTab & "chk=" & Format$(Rdv3EngChecksum(), "0") & _
           vbTab & "marks=" & CStr(m_marks) & _
           vbTab & "saved=" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & _
           vbTab & "bookmtime=" & FileTimeStr(g_ledgerPath)
    If m_rows > 0 Then
        SidecarWrite = WriteStateFileAtomic(g_sidecarPath, head & vbCrLf & Join(body, vbCrLf) & vbCrLf, errMsg)
    Else
        SidecarWrite = WriteStateFileAtomic(g_sidecarPath, head & vbCrLf, errMsg)
    End If
End Function

'------------------------------------------------------------------------------
' ledger workbook IO (BE process only; the FE never opens this file)
'------------------------------------------------------------------------------
Private Function EnsureLedgerBook(ByRef errMsg As String, _
                                  Optional ByVal recreate As Boolean = False) As Boolean
    errMsg = ""
    If Not m_lb Is Nothing Then
        EnsureLedgerBook = True
        Exit Function
    End If
    On Error GoTo Fail
    If Dir$(g_ledgerPath) <> "" Then
        Err.Clear
        On Error Resume Next
        Set m_lb = Application.Workbooks.Open(g_ledgerPath, 0, False)
        If Err.Number <> 0 Or m_lb Is Nothing Then
            errMsg = "台帳ブックを開けません " & Err.Number & ": " & Err.Description
            Set m_lb = Nothing
            On Error GoTo Fail
            If Not recreate Then Exit Function
            ' an approved rebuild may replace an unreadable ledger book
            BeLog "ledger book unreadable; recreating (" & errMsg & ")"
            Kill g_ledgerPath
        Else
            On Error GoTo Fail
            If m_lb.ReadOnly Then
                errMsg = "台帳ブックが読み取り専用で開かれています (他で開いていませんか)"
                m_lb.Close False
                Set m_lb = Nothing
                Exit Function
            End If
            errMsg = ""
            EnsureLedgerBook = True
            Exit Function
        End If
    End If
    ' create from nothing (first rebuild, or the recreate branch above)
    Set m_lb = Application.Workbooks.Add(-4167)              ' one worksheet
    m_lb.Worksheets(1).Name = RDV3_LB_SHEET
    m_lb.Worksheets.Add(, m_lb.Worksheets(1)).Name = RDV3_LB_META
    m_lb.SaveAs g_ledgerPath, 51                             ' xlOpenXMLWorkbook
    errMsg = ""
    EnsureLedgerBook = True
    Exit Function
Fail:
    errMsg = "台帳ブックを開けません " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not m_lb Is Nothing Then m_lb.Close False
    Set m_lb = Nothing
End Function

' rewrite the whole LEDGER + META content from BE memory and save. Excel's
' own Save writes through a temp file, so a crash mid-save keeps the old book.
Private Function WriteLedgerBookAll(ByRef errMsg As String, ByRef saveMs As Double) As Boolean
    Dim ws As Object
    Dim mt As Object
    Dim c As Long
    Dim r0 As Long, r1 As Long, i As Long, n As Long
    Dim arr As Variant
    Dim fields As Variant
    Dim t As Double

    errMsg = ""
    saveMs = 0
    If Not EnsureLedgerBook(errMsg, True) Then Exit Function
    On Error GoTo Fail

    Set ws = m_lb.Worksheets(RDV3_LB_SHEET)
    Set mt = m_lb.Worksheets(RDV3_LB_META)

    ws.Cells.ClearContents
    ws.Cells.NumberFormat = "@"                  ' keep zero-padded keys as text
    For c = 1 To RDV3_LED_COLS
        ws.Cells(1, c).Value2 = Rdv3LedgerHead(c - 1)
    Next c

    r0 = 0
    Do While r0 < m_rows
        r1 = r0 + 16384
        If r1 > m_rows Then r1 = m_rows
        n = r1 - r0
        ReDim arr(1 To n, 1 To RDV3_LED_COLS)
        For i = 1 To n
            arr(i, 1) = m_proc(r0 + i - 1)
            fields = Split(m_lines(r0 + i - 1), vbTab)
            For c = 0 To RDV3_CONTENT_COLS - 1
                If c <= UBound(fields) Then arr(i, c + 2) = CStr(fields(c))
            Next c
        Next i
        ws.Range(ws.Cells(2 + r0, 1), ws.Cells(1 + r1, RDV3_LED_COLS)).Value2 = arr
        r0 = r1
    Loop

    mt.Cells.ClearContents
    mt.Range("A1").Value2 = "rows":     mt.Range(RDV3_LM_ROWS).Value2 = m_rows
    mt.Range("A2").Value2 = "checksum": mt.Range(RDV3_LM_CHECKSUM).Value2 = Format$(Rdv3EngChecksum(), "0")
    mt.Range("A3").Value2 = "marks":    mt.Range(RDV3_LM_MARKS).Value2 = m_marks
    mt.Range("A4").Value2 = "saved":    mt.Range(RDV3_LM_SAVED).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    mt.Range("A5").Value2 = "format":   mt.Range(RDV3_LM_FORMAT).Value2 = "rdv3-ledger v1"

    t = Rdv3Ticks()
    m_lb.Save
    saveMs = Rdv3MsSince(t)
    WriteLedgerBookAll = True
    Exit Function
Fail:
    errMsg = "台帳ブック書込エラー " & Err.Number & ": " & Err.Description
End Function

Private Sub LoadExpected()
    Dim f As Integer, ln As String, eq As Long, k As String, v As String
    Dim p As String
    m_expLoaded = False
    m_expRows = 0
    m_expChk = -1
    p = g_dataDir & "\expected.txt"
    If Dir$(p) = "" Then Exit Sub
    On Error GoTo Fail
    f = FreeFile
    Open p For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        eq = InStr(ln, "=")
        If eq > 1 Then
            k = Left$(ln, eq - 1)
            v = Mid$(ln, eq + 1)
            If k = "rows" Then m_expRows = CLng(v)
            If k = "joinchecksum" Then m_expChk = CDbl(v)
        End If
    Loop
    Close #f
    m_expLoaded = (m_expRows > 0)
    Exit Sub
Fail:
    On Error Resume Next
    Close #f
End Sub

' ordered content comparison, saved (m_lines) vs new merge (engine lines)
Private Function CompareSavedToNew(ByRef firstDiff As Long) As Boolean
    Dim i As Long
    firstDiff = -1
    If m_rows <> Rdv3EngLineCount() Then
        firstDiff = m_rows
        If Rdv3EngLineCount() < m_rows Then firstDiff = Rdv3EngLineCount()
        Exit Function
    End If
    For i = 0 To m_rows - 1
        If StrComp(m_lines(i), Rdv3EngLine(i), vbBinaryCompare) <> 0 Then
            firstDiff = i
            Exit Function
        End If
    Next i
    CompareSavedToNew = True
End Function

'------------------------------------------------------------------------------
' adopt / apply
'------------------------------------------------------------------------------
Private Sub AdoptActive(ByVal note As String)
    Dim t As Double
    t = Rdv3Ticks()
    BuildIndex
    BeLog "index rows=" & CStr(m_rows) & " ms=" & Format$(Rdv3MsSince(t), "0.0")
    g_state = "READY"
    w_on = True
    PublishReady m_rows, Rdv3MsSince(t), note
End Sub

Private Sub DoApply()
    Dim carried As Long, resetN As Long, newN As Long, dropped As Long
    Dim np() As Boolean
    Dim newCount As Long
    Dim oldByKey2 As Object
    Dim i As Long
    Dim k2 As String
    Dim used As Long
    Dim t As Double
    Dim parts As Long
    Dim at As Long
    Dim chunk As Long
    Dim newLines() As String

    PublishState "applying", ""
    newCount = Rdv3EngLineCount()

    ' the carry: identity = key2, TRUE only if the full content is identical
    t = Rdv3Ticks()
    If newCount > 0 Then ReDim np(0 To newCount - 1)
    If m_savedOk And m_rows > 0 Then
        Set oldByKey2 = CreateObject("Scripting.Dictionary")
        oldByKey2.CompareMode = 0
        For i = 0 To m_rows - 1
            k2 = LineField(m_lines(i), 1)
            If Not oldByKey2.Exists(k2) Then oldByKey2.Add k2, i
        Next i
        used = 0
        For i = 0 To newCount - 1
            k2 = LineField(Rdv3EngLine(i), 1)
            If oldByKey2.Exists(k2) Then
                used = used + 1
                If StrComp(m_lines(oldByKey2.Item(k2)), Rdv3EngLine(i), vbBinaryCompare) = 0 Then
                    np(i) = m_proc(oldByKey2.Item(k2))
                    If np(i) Then carried = carried + 1
                Else
                    np(i) = False
                    If m_proc(oldByKey2.Item(k2)) Then resetN = resetN + 1
                End If
            Else
                np(i) = False
                newN = newN + 1
            End If
        Next i
        dropped = m_rows - used
        If dropped < 0 Then dropped = 0
    Else
        newN = newCount
    End If
    BeLog "carry carried=" & carried & " reset=" & resetN & " new=" & newN & _
        " dropped=" & dropped & " ms=" & Format$(Rdv3MsSince(t), "0.0")

    ' adopt the new content as the BE's active ledger
    ReDim newLines(0 To newCount - 1)
    For i = 0 To newCount - 1
        newLines(i) = Rdv3EngLine(i)
    Next i
    m_rows = newCount
    m_lines = newLines
    m_proc = np
    m_savedOk = True
    m_savedErr = ""

    ' persist: the BE writes and saves the ledger workbook itself, then the
    ' sidecar mirror. The FE never sees a single ledger row. A failure here is
    ' loud and blocking: memory and disk must not diverge silently.
    Dim perr As String
    Dim saveMs As Double
    t = Rdv3Ticks()
    If Not WriteLedgerBookAll(perr, saveMs) Then
        PublishErr "apply", "台帳ブックの書込に失敗: " & perr
        BeLog "apply FAILED: " & perr
        g_state = "BLOCKED"
        PublishReady 0, 0, "blocked"
        Exit Sub
    End If
    If Not SidecarWrite(perr) Then
        ' the workbook (the authority) is saved; a failed sidecar only costs
        ' the next boot the slow self-heal path. Say so, loudly.
        BeLog "sidecar write FAILED after apply: " & perr
    End If
    BeLog "ledger written rows=" & newCount & " total_ms=" & Format$(Rdv3MsSince(t), "0.0") & _
        " save_ms=" & Format$(saveMs, "0.0")

    Publish RDV3_K_APPLY, "rows=" & newCount & ";carried=" & carried & _
        ";reset=" & resetN & ";new=" & newN & ";drop=" & dropped & _
        ";write=" & Format$(Rdv3MsSince(t), "0.00") & _
        ";save=" & Format$(saveMs, "0.00"), Empty
    AdoptActive "applied"
End Sub

Private Sub BuildIndex()
    Dim i As Long
    Dim k As String
    Dim col As Collection
    Set m_dict = CreateObject("Scripting.Dictionary")
    m_dict.CompareMode = 0
    For i = 0 To m_rows - 1
        k = LineField(m_lines(i), 0)
        If m_dict.Exists(k) Then
            m_dict.Item(k).Add i
        Else
            Set col = New Collection
            col.Add i
            m_dict.Add k, col
        End If
    Next i
End Sub

Private Function LineField(ByRef line As String, ByVal col As Long) As String
    Dim p As Long, q As Long, k As Long
    p = 1
    For k = 1 To col
        q = InStr(p, line, vbTab, vbBinaryCompare)
        If q = 0 Then
            LineField = ""
            Exit Function
        End If
        p = q + 1
    Next k
    q = InStr(p, line, vbTab, vbBinaryCompare)
    If q = 0 Then q = Len(line) + 1
    LineField = Mid$(line, p, q - p)
End Function

'------------------------------------------------------------------------------
' requests (latest-wins per kind)
'------------------------------------------------------------------------------
Private Sub PollRequests()
    Dim ver As Long
    Dim args As String
    Dim parts As Variant

    If Rdv3ChReadReq(g_sid, RDV3_RQ_DECISION, ver, args) Then
        If ver > m_qDecision Then
            m_qDecision = ver
            BeLog "req decision " & args
            If g_state = "AWAIT" Then
                Select Case args
                    Case "apply"
                        DoApply
                    Case "adopt"
                        If m_savedOk Then
                            AdoptActive "adopted"
                        Else
                            g_state = "BLOCKED"
                            PublishReady 0, 0, "blocked"
                        End If
                    Case "block"
                        g_state = "BLOCKED"
                        PublishReady 0, 0, "blocked"
                End Select
            End If
        End If
    End If

    If Rdv3ChReadReq(g_sid, RDV3_RQ_SEARCH, ver, args) Then
        If ver > m_qSearch Then
            m_qSearch = ver
            parts = Split(args, "|")
            If UBound(parts) >= 2 And g_state = "READY" Then
                DoSearch CStr(parts(0)), CStr(parts(1)), CStr(parts(2))
            End If
        End If
    End If

    If Rdv3ChReadReq(g_sid, RDV3_RQ_PICK, ver, args) Then
        If ver > m_qPick Then
            m_qPick = ver
            parts = Split(args, "|")
            If UBound(parts) >= 3 And g_state = "READY" Then
                DoPick CLng(Val(parts(0))), CLng(Val(parts(1))), CLng(Val(parts(2))), CStr(parts(3))
            End If
        End If
    End If

    If Rdv3ChReadReq(g_sid, RDV3_RQ_MARK, ver, args) Then
        If ver > m_qMark Then
            m_qMark = ver
            parts = Split(args, "|")
            If UBound(parts) >= 2 And g_state = "READY" Then
                ' the request's own version travels with the answer, so the FE
                ' can tell THIS confirmation from any other
                DoMark CLng(Val(parts(0))), (CStr(parts(1)) = "1"), CStr(parts(2)), ver
            End If
        End If
    End If

    If Rdv3ChReadReq(g_sid, RDV3_RQ_WATCH, ver, args) Then
        If ver > m_qWatch Then
            m_qWatch = ver
            w_on = (args = "1")
            BeLog "req watch " & args
            If Not w_on Then
                PublishState "watch_off", ""
            Else
                w_saidWaiting = False
                Rdv3UiaReset
            End If
        End If
    End If
End Sub

'------------------------------------------------------------------------------
' search / pick: answered from BE memory, rendered by the FE pump
'------------------------------------------------------------------------------
Private Sub DoSearch(ByVal key As String, ByVal t0s As String, ByVal source As String)
    Dim hits As Collection
    Dim n As Long, show As Long
    Dim i As Long, row As Long
    Dim meta As String
    Dim vals As Variant
    Dim t As Double

    t = Rdv3Ticks()
    Set hits = Nothing
    If Not m_dict Is Nothing Then
        If m_dict.Exists(key) Then Set hits = m_dict.Item(key)
    End If
    If hits Is Nothing Then
        n = 0
    Else
        n = hits.Count
    End If

    If n = 1 Then
        row = hits.Item(1)
        vals = RecordBlock(row)
        meta = "res=single;key=" & key & ";hits=1;src=" & source & ";t0=" & t0s & _
               ";k2=" & LineField(m_lines(row), 1) & ";proc=" & IIf(m_proc(row), "1", "0") & _
               ";row=" & CStr(row)
        Publish RDV3_K_RESULT, meta, vals
    ElseIf n = 0 Then
        meta = "res=none;key=" & key & ";hits=0;src=" & source & ";t0=" & t0s
        Publish RDV3_K_RESULT, meta, Empty
    Else
        show = n
        If show > RDV3_MAXSHOW Then show = RDV3_MAXSHOW
        ReDim vals(1 To show, 1 To 9)
        For i = 1 To show
            row = hits.Item(i)
            FillCandRow vals, i, row
        Next i
        meta = "res=multi;key=" & key & ";hits=" & CStr(n) & ";shown=" & CStr(show) & _
               ";src=" & source & ";t0=" & t0s
        Publish RDV3_K_RESULT, meta, vals
    End If
    BeLog "search key=" & key & " src=" & source & " hits=" & n & _
        " serve_ms=" & Format$(Rdv3MsSince(t), "0.0")
End Sub

Private Sub DoPick(ByVal row As Long, ByVal slot As Long, ByVal total As Long, ByVal t0s As String)
    Dim meta As String
    Dim vals As Variant
    If row < 0 Or row >= m_rows Then Exit Sub
    vals = RecordBlock(row)
    meta = "res=picked;hits=1;t0=" & t0s & ";k2=" & LineField(m_lines(row), 1) & _
           ";proc=" & IIf(m_proc(row), "1", "0") & ";row=" & CStr(row) & _
           ";slot=" & CStr(slot) & ";total=" & CStr(total)
    Publish RDV3_K_RESULT, meta, vals
    BeLog "pick row=" & row & " slot=" & slot
End Sub

' "processed": the BE persists it -- workbook cell, workbook save, sidecar --
' and only then confirms. On any failure the memory flag is reverted and the
' FE gets an explicit markerr: state never degrades silently.
'
' The confirmation goes out on RDV3_K_MARK, its own channel slot, and carries
' req=<the version of the request it answers>: a search answered a moment later
' can no longer take its place, and the FE can tell which request it belongs to.
Private Sub DoMark(ByVal row As Long, ByVal newVal As Boolean, ByVal t0s As String, _
                   ByVal reqVer As Long)
    Dim errMsg As String
    Dim scErr As String
    Dim oldVal As Boolean
    Dim ws As Object
    Dim mt As Object
    Dim t As Double
    Dim saveMs As Double
    Dim k2 As String
    Dim oldMarks As Long

    If row < 0 Or row >= m_rows Then
        Publish RDV3_K_MARK, "res=markerr;row=" & CStr(row) & ";req=" & CStr(reqVer) & _
            ";t0=" & t0s & ";msg=" & Replace("行が不正です", ";", ","), Empty
        Exit Sub
    End If
    k2 = LineField(m_lines(row), 1)
    oldVal = m_proc(row)
    oldMarks = m_marks
    m_proc(row) = newVal

    On Error GoTo Fail
    If m_saveMethod = RDV3_SAVE_BOOK Then
        If Not EnsureLedgerBook(errMsg) Then GoTo FailMsg
        Set ws = m_lb.Worksheets(RDV3_LB_SHEET)
        Set mt = m_lb.Worksheets(RDV3_LB_META)
        ws.Cells(2 + row, 1).Value2 = newVal
        m_marks = m_marks + 1
        mt.Range(RDV3_LM_MARKS).Value2 = m_marks
        mt.Range(RDV3_LM_SAVED).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
        t = Rdv3Ticks()
        m_lb.Save
        saveMs = Rdv3MsSince(t)
    Else
        ' the comparison methods work on the CLOSED file, so the workbook this
        ' process may be holding open has to be let go first
        If Not m_lb Is Nothing Then
            m_lb.Close False
            Set m_lb = Nothing
            Rdv3SaveReset
        End If
        t = Rdv3Ticks()
        If m_saveMethod = RDV3_SAVE_ADO Then
            If Not Rdv3SaveOneAdo(g_ledgerPath, 2 + row, newVal, errMsg) Then GoTo FailMsg
        Else
            If Not Rdv3SaveOneZip(g_ledgerPath, 2 + row, m_rows, newVal, errMsg) Then GoTo FailMsg
        End If
        saveMs = Rdv3MsSince(t)
        ' the counter still moves and still reaches the sidecar, which is where
        ' the next boot reads it; the ledger's META copy is only refreshed by
        ' the book method, so on these two it can lag until the next full write
        m_marks = m_marks + 1
        BeLog "mark via " & Rdv3SaveMethodName(m_saveMethod) & " " & Rdv3SaveLastNote()
    End If

    If Not SidecarWrite(scErr) Then
        ' the saved workbook is the authority; a stale sidecar only costs the
        ' next boot the slow self-heal path
        BeLog "sidecar write FAILED after mark: " & scErr
    End If

    Publish RDV3_K_MARK, "res=marked;row=" & CStr(row) & ";k2=" & k2 & _
        ";proc=" & IIf(newVal, "1", "0") & ";save_ms=" & Format$(saveMs, "0.00") & _
        ";req=" & CStr(reqVer) & ";t0=" & t0s, Empty
    BeLog "mark row=" & row & " k2=" & k2 & " value=" & IIf(newVal, "TRUE", "FALSE") & _
        " save_ms=" & Format$(saveMs, "0.0")
    Exit Sub
Fail:
    errMsg = "処理済みの保存に失敗 " & Err.Number & ": " & Err.Description
FailMsg:
    On Error Resume Next
    m_proc(row) = oldVal
    m_marks = oldMarks
    Publish RDV3_K_MARK, "res=markerr;row=" & CStr(row) & ";k2=" & k2 & _
        ";req=" & CStr(reqVer) & ";t0=" & t0s & ";msg=" & Replace(errMsg, ";", ","), Empty
    BeLog "mark FAILED row=" & row & ": " & errMsg
End Sub

' the 30-field record block: 10 rows x 3 columns (A value, B value, C value)
Private Function RecordBlock(ByVal row As Long) As Variant
    Dim f(0 To RDV3_CONTENT_COLS - 1) As String
    Dim vals(1 To 10, 1 To 3) As Variant
    Dim i As Long
    Dim hasA As Boolean, hasC As Boolean
    SplitLine m_lines(row), f
    hasA = False
    For i = 2 To 10
        If Len(f(i)) > 0 Then
            hasA = True
            Exit For
        End If
    Next i
    hasC = False
    For i = 19 To 27
        If Len(f(i)) > 0 Then
            hasC = True
            Exit For
        End If
    Next i
    If hasA Then vals(1, 1) = f(0) Else vals(1, 1) = ""
    For i = 1 To 9
        vals(1 + i, 1) = f(1 + i)
    Next i
    vals(1, 2) = f(0)
    vals(2, 2) = f(1)
    For i = 2 To 9
        vals(1 + i, 2) = f(9 + i)
    Next i
    If hasC Then vals(1, 3) = f(1) Else vals(1, 3) = ""
    For i = 1 To 9
        vals(1 + i, 3) = f(18 + i)
    Next i
    RecordBlock = vals
End Function

' candidate slot: the eight identification columns + the ledger row (col 9)
Private Sub FillCandRow(ByRef vals As Variant, ByVal i As Long, ByVal row As Long)
    Dim f(0 To RDV3_CONTENT_COLS - 1) As String
    SplitLine m_lines(row), f
    vals(i, 1) = f(1)      ' key2
    vals(i, 2) = f(18)     ' b_line
    vals(i, 3) = f(11)     ' b_slip
    vals(i, 4) = f(12)     ' b_date
    vals(i, 5) = f(13)     ' b_qty
    vals(i, 6) = f(16)     ' b_status
    vals(i, 7) = f(19)     ' c_item
    vals(i, 8) = f(20)     ' c_maker
    vals(i, 9) = row
End Sub

Private Sub SplitLine(ByRef line As String, ByRef out() As String)
    Dim parts As Variant
    Dim i As Long
    parts = Split(line, vbTab)
    For i = 0 To RDV3_CONTENT_COLS - 1
        If i <= UBound(parts) Then
            out(i) = parts(i)
        Else
            out(i) = ""
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' the Notepad watch (same rules as every build before this one), 40 ms cadence
'------------------------------------------------------------------------------
Private Sub PollWatch()
    Dim raw As Variant
    Dim cand As String
    Dim held As Double
    Dim t0 As Double

    If Not Rdv3UiaBound() Then
        If Rdv3UiaBind() Then
            w_bound = True
            w_saidWaiting = False
            w_pending = ""
            w_pendPolls = 0
            PublishState "bound", "hwnd=" & CStr(Rdv3UiaHwnd()) & ";title=" & Replace(Rdv3UiaTitle(), ";", ",")
            BeLog "watch bound hwnd=" & CStr(Rdv3UiaHwnd())
        Else
            If Not w_saidWaiting Then
                w_saidWaiting = True
                w_bound = False
                PublishState "waiting", "why=" & Replace(Rdv3UiaWhy(), ";", ",")
            End If
            Rdv3WaitMs 360     ' bind retry backoff on top of the loop's 40 ms
            Exit Sub
        End If
    End If

    raw = Rdv3UiaRead()
    If IsNull(raw) Then
        Rdv3UiaReset
        If w_bound Then
            w_bound = False
            w_saidWaiting = False
        End If
        Exit Sub
    End If

    cand = Rdv3Candidate(CStr(raw))
    If Len(cand) = 0 Then
        w_sawEmpty = True
        w_pending = ""
        w_pendPolls = 0
    ElseIf Not Rdv3IsKey(cand) Then
        w_pending = ""
        w_pendPolls = 0
    ElseIf cand = w_lastFired And Not w_sawEmpty Then
        w_pending = ""
        w_pendPolls = 0
    ElseIf cand <> w_pending Then
        w_pending = cand
        w_pendSince = Rdv3Ticks()
        w_pendPolls = 1
    Else
        w_pendPolls = w_pendPolls + 1
        held = Rdv3MsSince(w_pendSince)
        If held >= RDV3_STABLE_MS Then
            t0 = Rdv3Ticks()
            w_lastFired = cand
            w_sawEmpty = False
            w_pending = ""
            BeLog "detect key=" & cand & " latency_ms=" & Format$(held, "0.0") & _
                " polls=" & CStr(w_pendPolls)
            DoSearch cand, Trim$(Str$(CDbl(t0))), "detect"
            w_pendPolls = 0
        End If
    End If
End Sub

Private Sub Heartbeat()
    Dim st As String
    If g_state = "READY" Then
        If Not w_on Then
            st = "watch_off"
        ElseIf w_bound Then
            st = "bound"
        Else
            st = "waiting"
        End If
    Else
        st = LCase$(g_state)
    End If
    PublishState "hb_" & st, IIf(w_bound, "hwnd=" & CStr(Rdv3UiaHwnd()) & ";title=" & Replace(Rdv3UiaTitle(), ";", ","), "")
End Sub

'------------------------------------------------------------------------------
' publish helpers
'------------------------------------------------------------------------------
Private Sub Publish(ByVal kind As String, ByVal meta As String, ByVal values As Variant)
    Dim tries As Long
    g_ver = g_ver + 1
    For tries = 1 To 6
        If Rdv3ChPublish(g_sid, kind, g_ver, meta, values) Then Exit Sub
        Rdv3WaitMs 40
    Next tries
    ' nothing was lost -- the live aggregate and the pending tmp are both still
    ' there (Rdv3ChReplaceFile) -- but this record did not reach the FE, so say
    ' so with the reason rather than dropping it silently
    BeLog "publish skipped kind=" & kind & " after 6 tries, last replace err=" & _
        CStr(Rdv3ChLastReplaceErr())
End Sub

Private Sub PublishState(ByVal st As String, ByVal extra As String)
    Publish RDV3_K_STATE, "st=" & st & IIf(Len(extra) > 0, ";" & extra, ""), Empty
End Sub

Private Sub PublishReady(ByVal rows As Long, ByVal idxMs As Double, ByVal note As String)
    Publish RDV3_K_READY, "rows=" & CStr(rows) & ";idx=" & Format$(idxMs, "0.00") & ";note=" & note, Empty
End Sub

Private Sub PublishErr(ByVal stage As String, ByVal msg As String)
    Publish RDV3_K_BEERR, "stage=" & stage & ";msg=" & Replace(Replace(msg, ";", ","), "=", ":"), Empty
    BeLog "ERROR stage=" & stage & " msg=" & msg
End Sub

Private Sub BeLog(ByVal msg As String)
    Dim f As Integer
    If Len(g_logPath) = 0 Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Open g_logPath For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " BE " & msg
    Close #f
    On Error GoTo 0
End Sub
