Attribute VB_Name = "modRdv2App"
'==============================================================================
' modRdv2App -- the loop, the screen, the candidate dialog, and where the clock
' stops.
'
'   t0             a number was confirmed (start of the merge-select)
'   stages 1..9    modRdv2Engine: read three, index three, join twice, search
'   stage 10       one candidate -> the record on the sheet
'                  several       -> the dialog is built and its list filled
'   stop clock     total = now - t0. This never contains a person.
'   (dialog)       the operator reads and picks. Not measured at all.
'   pick           the chosen candidate's record, timed on its own.
'
' The timings are written only after the clock has stopped, so the screen is
' never timing the drawing of its own stopwatch.
'
' Detection uses the same two rules as the C# builds: exactly 8 digits, and the
' value must survive RDV2_STABLE_MS of polling. The same number does not fire
' twice in a row unless the field went empty in between.
'==============================================================================
Option Explicit

Private m_running As Boolean
Private m_seq As Long
Private m_pending As String
Private m_pendSince As Currency
Private m_pendPolls As Long
Private m_lastFired As String
Private m_sawEmpty As Boolean
Private m_detectMs As Double
Private m_polls As Long
Private m_totalMs As Double
Private m_pickMs As Double
Private m_dataDir As String
Private m_logPath As String
Private m_autoPick As Long
Private m_expRows As Long
Private m_expChk As Double
Private m_expLoaded As Boolean
Private m_oracleOk As Boolean
Private m_oracleNote As String

Private Function Vw() As Object
    Set Vw = ThisWorkbook.Worksheets(RDV2_SHEET)
End Function

Public Function Rdv2DataDir() As String
    Dim ws As Object
    Dim c As String
    Dim p As String
    Set ws = Vw()
    c = Trim$(CStr(ws.Range(RDV2_CELL_DATADIR).Value))
    If Len(c) > 0 Then
        If Dir$(c & "\tableA.csv") <> "" Then
            Rdv2DataDir = c
            Exit Function
        End If
    End If
    p = ThisWorkbook.path
    If Dir$(p & "\..\data-100k\tableA.csv") <> "" Then
        Rdv2DataDir = p & "\..\data-100k"
        Exit Function
    End If
    If Dir$(p & "\data-100k\tableA.csv") <> "" Then
        Rdv2DataDir = p & "\data-100k"
        Exit Function
    End If
    Rdv2DataDir = ""
End Function

Private Sub LoadExpected()
    Dim f As Integer, line As String, eq As Long, k As String, v As String
    Dim p As String
    m_expLoaded = False
    m_expRows = 0
    m_expChk = -1
    p = m_dataDir & "\expected.txt"
    If Dir$(p) = "" Then Exit Sub
    f = FreeFile
    Open p For Input As #f
    Do While Not EOF(f)
        Line Input #f, line
        eq = InStr(line, "=")
        If eq > 1 Then
            k = Left$(line, eq - 1)
            v = Mid$(line, eq + 1)
            If k = "rows" Then m_expRows = CLng(v)
            If k = "joinchecksum" Then m_expChk = CDbl(v)
        End If
    Loop
    Close #f
    m_expLoaded = (m_expRows > 0)
End Sub

'------------------------------------------------------------------------------
' screen
'------------------------------------------------------------------------------
Public Sub Rdv2PaintFrame()
    Dim ws As Object
    Set ws = Vw()
    ws.Range(RDV2_RG_KEY).ClearContents
    ws.Range(RDV2_RG_RECORD).ClearContents
    ws.Range(RDV2_RG_STAGE).ClearContents
    ws.Range(RDV2_COL_FIRST & RDV2_HIST_TOP & ":" & RDV2_COL_LAST & (RDV2_HIST_TOP + RDV2_HIST_ROWS - 1)).ClearContents
    ws.Range(RDV2_COL_FIRST & RDV2_CAND_TOP & ":" & RDV2_COL_LAST & (RDV2_CAND_TOP + RDV2_CAND_ROWS - 1)).ClearContents
    ws.Range(RDV2_CELL_ERROR).Value = ""
    m_seq = 0
End Sub

Private Sub SetState(ByVal state As String, ByVal detail As String)
    Dim ws As Object
    On Error Resume Next
    Set ws = Vw()
    ws.Range(RDV2_CELL_STATE).Value = state
    If Len(detail) > 0 Then
        ws.Range(RDV2_CELL_NOTEPAD).Value = detail
        ws.Range(RDV2_CELL_DATA).Value = m_dataDir
        ws.Range(RDV2_CELL_INDEX).Value = Rdv2IndexKind()
    End If
End Sub

' the record of one candidate. This is the end of the measured region when
' there was exactly one, and the "選択後の表示" figure when the operator picked.
Private Sub ShowRecord(ByVal key As String)
    Dim ws As Object
    Dim rec(1 To 10, 1 To 7) As Variant
    Dim i As Long
    Dim verdict As String

    Set ws = Vw()
    If Len(Rdv2ErrText()) > 0 Then
        verdict = "エラー"
    ElseIf Rdv2CandCount() = 0 Then
        verdict = "該当なし"
    ElseIf Rdv2Chosen() >= 0 Then
        verdict = "一致 1 件   key2 = " & Rdv2CandCell(Rdv2Chosen(), 0)
        If Rdv2CandCount() > 1 Then
            verdict = verdict & "   (候補 " & Rdv2CandCount() & " 件中 " & (Rdv2Chosen() + 1) & " 件目)"
        End If
    Else
        verdict = "候補 " & Rdv2CandCount() & " 件"
    End If

    For i = 1 To RDV2_FIELDS
        rec(i, 1) = i
        rec(i, 2) = Rdv2NameOf(0, i - 1)
        rec(i, 3) = Rdv2ValueOf(0, i - 1)
        rec(i, 4) = Rdv2NameOf(1, i - 1)
        rec(i, 5) = Rdv2ValueOf(1, i - 1)
        rec(i, 6) = Rdv2NameOf(2, i - 1)
        rec(i, 7) = Rdv2ValueOf(2, i - 1)
    Next i
    ws.Range(RDV2_CELL_KEY).Value = key
    ws.Range(RDV2_CELL_VERDICT).Value = verdict
    ws.Range(RDV2_RG_RECORD).Value = rec
End Sub

' part of stage 10 when there are several candidates: build the dialog and fill
' its list. Showing it comes afterwards, once the clock has stopped.
Private Function BuildPick(ByVal key As String) As Object
    Dim f As frmRdv2Pick
    Set f = New frmRdv2Pick
    f.Fill key, Rdv2CandCount()
    Set BuildPick = f
End Function

' an echo of the candidate list on the sheet, written after the clock has
' stopped so the operator can still see what the other candidates were
Private Sub ShowCandidateSheet()
    Dim ws As Object
    Dim n As Long, i As Long, c As Long
    Dim tbl() As Variant
    Set ws = Vw()
    ws.Range(RDV2_COL_FIRST & RDV2_CAND_TOP & ":" & RDV2_COL_LAST & (RDV2_CAND_TOP + RDV2_CAND_ROWS - 1)).ClearContents
    n = Rdv2CandCount()
    If n <= 1 Then Exit Sub
    If n > RDV2_CAND_ROWS Then n = RDV2_CAND_ROWS
    ReDim tbl(1 To n, 1 To 8)
    For i = 1 To n
        For c = 1 To 8
            tbl(i, c) = Rdv2CandCell(i - 1, c - 1)
        Next c
    Next i
    ws.Range(RDV2_COL_FIRST & RDV2_CAND_TOP & ":" & RDV2_COL_LAST & (RDV2_CAND_TOP + n - 1)).Value = tbl
End Sub

Private Sub ShowStats(ByVal key As String)
    Dim ws As Object
    Dim tbl(1 To 17, 1 To 3) As Variant
    Dim i As Long, r As Long
    Dim total As Double

    Set ws = Vw()
    total = m_totalMs
    For i = 0 To RDV2_STAGES - 1
        tbl(i + 1, 1) = Rdv2StageName(i)
        tbl(i + 1, 2) = Rdv2StageMs(i)
        If total > 0 Then tbl(i + 1, 3) = 100# * Rdv2StageMs(i) / total
    Next i
    r = RDV2_STAGES + 1
    tbl(r, 1) = "その他 (差分)"
    tbl(r, 2) = total - Rdv2StageSum()
    If total > 0 Then tbl(r, 3) = 100# * (total - Rdv2StageSum()) / total
    r = r + 1
    tbl(r, 1) = "合計 (選択待ちを除く)"
    tbl(r, 2) = total
    tbl(r, 3) = 100#
    r = r + 1
    tbl(r, 1) = "選択後の表示 (計測外)"
    If m_pickMs >= 0 Then tbl(r, 2) = m_pickMs
    r = r + 1
    tbl(r, 1) = "検知遅延 (参考・計測外)"
    tbl(r, 2) = m_detectMs
    r = r + 1
    tbl(r, 1) = "候補数"
    tbl(r, 2) = Rdv2CandCount()
    r = r + 1
    tbl(r, 1) = "データ行数"
    tbl(r, 2) = Rdv2Rows()
    r = r + 1
    tbl(r, 1) = "索引走査回数"
    tbl(r, 2) = Rdv2Probes()
    ws.Range(RDV2_RG_STAGE).Value = tbl
    ws.Range(RDV2_CELL_ERROR).Value = Rdv2ErrText()
End Sub

Private Sub PushHistory(ByVal key As String)
    Dim ws As Object
    Dim row(1 To 1, 1 To 8) As Variant
    Dim top As Long, last As Long

    Set ws = Vw()
    top = RDV2_HIST_TOP
    last = RDV2_HIST_TOP + RDV2_HIST_ROWS - 1
    ws.Range(RDV2_COL_FIRST & top & ":" & RDV2_COL_LAST & (last - 1)).Value = _
        ws.Range(RDV2_COL_FIRST & (top + 1) & ":" & RDV2_COL_LAST & last).Value
    row(1, 1) = m_seq
    row(1, 2) = Format$(Now, "hh:nn:ss")
    row(1, 3) = key
    row(1, 4) = Rdv2CandCount()
    row(1, 5) = m_totalMs
    If m_pickMs >= 0 Then
        row(1, 6) = m_pickMs
    Else
        row(1, 6) = ""
    End If
    row(1, 7) = m_detectMs
    If Len(Rdv2ErrText()) > 0 Then
        row(1, 8) = "エラー"
    ElseIf m_oracleOk Then
        row(1, 8) = "検算 OK"
    Else
        row(1, 8) = "検算 NG " & m_oracleNote
    End If
    ws.Range(RDV2_COL_FIRST & last & ":" & RDV2_COL_LAST & last).Value = row
End Sub

Private Sub CheckOracle()
    m_oracleOk = True
    m_oracleNote = ""
    If Not m_expLoaded Then
        m_oracleNote = "expected.txt なし"
        Exit Sub
    End If
    If m_expRows <> Rdv2Rows() Then
        m_oracleOk = False
        m_oracleNote = "rows"
    ElseIf m_expChk <> Rdv2Checksum() Then
        m_oracleOk = False
        m_oracleNote = "checksum"
    End If
End Sub

' opt-in only: nothing is written unless a log path is in the settings cell
Private Sub AppendLog(ByVal key As String)
    Dim f As Integer, i As Long
    Dim s As String
    Dim isNew As Boolean
    If Len(m_logPath) = 0 Then Exit Sub
    On Error Resume Next
    isNew = (Dir$(m_logPath) = "")
    f = FreeFile
    Open m_logPath For Append As #f
    If isNew Then
        s = "seq" & vbTab & "time" & vbTab & "key" & vbTab & "index"
        For i = 0 To RDV2_STAGES - 1
            s = s & vbTab & Rdv2StageKey(i)
        Next i
        s = s & vbTab & "other" & vbTab & "total" & vbTab & "pick" & vbTab & "detect" & vbTab & _
            "polls" & vbTab & "cands" & vbTab & "rows" & vbTab & "probes" & vbTab & "checksum" & _
            vbTab & "matchedAB" & vbTab & "matchedBC" & vbTab & "oracle" & vbTab & "error"
        Print #f, s
    End If
    s = m_seq & vbTab & Format$(Now, "hh:nn:ss") & vbTab & key & vbTab & Rdv2IndexKind()
    For i = 0 To RDV2_STAGES - 1
        s = s & vbTab & Format$(Rdv2StageMs(i), "0.00")
    Next i
    s = s & vbTab & Format$(m_totalMs - Rdv2StageSum(), "0.00")
    s = s & vbTab & Format$(m_totalMs, "0.00")
    s = s & vbTab & Format$(m_pickMs, "0.00")
    s = s & vbTab & Format$(m_detectMs, "0.00")
    s = s & vbTab & m_polls
    s = s & vbTab & Rdv2CandCount()
    s = s & vbTab & Rdv2Rows()
    s = s & vbTab & Format$(Rdv2Probes(), "0")
    s = s & vbTab & Format$(Rdv2Checksum(), "0")
    s = s & vbTab & Rdv2MatchedAB()
    s = s & vbTab & Rdv2MatchedBC()
    If m_oracleOk Then
        s = s & vbTab & "ok"
    Else
        s = s & vbTab & "NG:" & m_oracleNote
    End If
    s = s & vbTab & Replace(Replace(Rdv2ErrText(), vbTab, " "), vbCrLf, " ")
    Print #f, s
    Close #f
End Sub

'------------------------------------------------------------------------------
' one merge-select
'------------------------------------------------------------------------------
Private Sub RunOne(ByVal key As String, ByVal t0 As Currency, ByVal detectMs As Double, ByVal polls As Long)
    Dim t As Currency
    Dim p0 As Currency
    Dim ok As Boolean
    Dim n As Long
    Dim picked As Long
    Dim frm As Object

    m_seq = m_seq + 1
    m_detectMs = detectMs
    m_polls = polls
    m_pickMs = -1
    SetState "処理中", ""

    ok = Rdv2RunMergeSelect(m_dataDir, key)

    t = Rdv2Ticks()
    n = Rdv2CandCount()
    If Len(Rdv2ErrText()) > 0 Or n = 0 Then
        ShowRecord key
    ElseIf n = 1 Then
        Rdv2FillCandidates
        Rdv2FillRecord 0
        ShowRecord key
    Else
        Rdv2FillCandidates
        Set frm = BuildPick(key)
    End If
    Rdv2SetStageMs RDV2_ST_SHOW, Rdv2MsSince(t)
    m_totalMs = Rdv2MsSince(t0)

    CheckOracle
    ShowStats key

    If n > 1 And Len(Rdv2ErrText()) = 0 Then
        SetState "候補を選択してください", ""
        frm.AutoPick = m_autoPick
        frm.Show 1
        picked = frm.Chosen
        Unload frm
        Set frm = Nothing
        If picked >= 0 Then
            p0 = Rdv2Ticks()
            Rdv2FillRecord picked
            ShowRecord key
            m_pickMs = Rdv2MsSince(p0)
            ShowStats key
        End If
        ShowCandidateSheet
    End If

    PushHistory key
    AppendLog key
    SetState "監視中", ""
End Sub

'------------------------------------------------------------------------------
' the loop
'------------------------------------------------------------------------------
Public Sub RDV2_StartMonitor()
    Dim ws As Object
    Dim raw As Variant
    Dim cand As String
    Dim held As Double
    Dim t0 As Currency
    Dim stopped As Boolean
    Dim pick As String

    If m_running Then Exit Sub
    On Error GoTo Fail
    Application.EnableCancelKey = xlErrorHandler

    Set ws = Vw()
    m_dataDir = Rdv2DataDir()
    m_logPath = Trim$(CStr(ws.Range(RDV2_CELL_LOG).Value))
    pick = Trim$(CStr(ws.Range(RDV2_CELL_AUTOPICK).Value))
    If Len(pick) > 0 And IsNumeric(pick) Then
        m_autoPick = CLng(pick)
    Else
        m_autoPick = -1
    End If
    If Len(m_dataDir) = 0 Then
        ws.Range(RDV2_CELL_ERROR).Value = "データが見つかりません。build\gen_data2.ps1 を先に実行してください。"
        Exit Sub
    End If
    LoadExpected

    m_running = True
    m_pending = ""
    m_pendPolls = 0
    m_lastFired = ""
    m_sawEmpty = True
    SetState "監視中", ""

    Do While m_running
        If UCase$(Trim$(CStr(ws.Range(RDV2_CELL_STOP).Value))) = "STOP" Then
            stopped = True
            Exit Do
        End If

        If Not Rdv2UiaBound() Then
            If Rdv2UiaBind() Then
                SetState "監視中", Rdv2UiaTitle() & "  (hwnd " & CStr(Rdv2UiaHwnd()) & ")"
                m_pending = ""
                m_pendPolls = 0
            Else
                SetState "メモ帳を待機中", "メモ帳: 未接続 -- " & Rdv2UiaWhy()
                DoEvents
                Rdv2Sleep 400
            End If
        End If

        If Rdv2UiaBound() Then
            raw = Rdv2UiaRead()
            If Not IsNull(raw) Then
                cand = Rdv2Candidate(CStr(raw))
                If Len(cand) = 0 Then
                    m_sawEmpty = True
                    m_pending = ""
                    m_pendPolls = 0
                ElseIf Not Rdv2IsKey(cand) Then
                    m_pending = ""
                    m_pendPolls = 0
                ElseIf cand = m_lastFired And Not m_sawEmpty Then
                    m_pending = ""
                    m_pendPolls = 0
                ElseIf cand <> m_pending Then
                    m_pending = cand
                    m_pendSince = Rdv2Ticks()
                    m_pendPolls = 1
                Else
                    m_pendPolls = m_pendPolls + 1
                    held = Rdv2MsSince(m_pendSince)
                    If held >= RDV2_STABLE_MS Then
                        t0 = Rdv2Ticks()
                        m_lastFired = cand
                        m_sawEmpty = False
                        m_pending = ""
                        RunOne cand, t0, held, m_pendPolls
                        m_pendPolls = 0
                    End If
                End If
            Else
                Rdv2UiaReset
            End If
        End If

        DoEvents
        Rdv2Sleep RDV2_POLL_MS
    Loop

    m_running = False
    If stopped Then
        SetState "停止", ""
    Else
        SetState "待機", ""
    End If
    Exit Sub
Fail:
    m_running = False
    On Error Resume Next
    Vw().Range(RDV2_CELL_ERROR).Value = "監視エラー " & Err.Number & ": " & Err.Description
    SetState "停止", ""
End Sub

' RDV2_StartMonitor does not return until the monitor stops, so a caller that
' drives this workbook over COM has to start it through OnTime or it deadlocks
' waiting for its own Application.Run to come back.
Public Sub RDV2_StartMonitorAsync()
    Vw().Range(RDV2_CELL_STOP).Value = ""
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!RDV2_StartMonitor"
End Sub

Public Sub RDV2_StopMonitor()
    m_running = False
    On Error Resume Next
    Vw().Range(RDV2_CELL_STOP).Value = "STOP"
End Sub

Public Sub RDV2_ManualRun()
    Dim ws As Object
    Dim key As String
    Dim t0 As Currency
    Set ws = Vw()
    key = Trim$(CStr(ws.Range(RDV2_CELL_MANUAL).Value))
    If Not Rdv2IsKey(key) Then
        ws.Range(RDV2_CELL_ERROR).Value = "手動実行には " & RDV2_KEYLEN & " 桁の数字を入れてください。"
        Exit Sub
    End If
    If Len(m_dataDir) = 0 Then
        m_dataDir = Rdv2DataDir()
        m_logPath = Trim$(CStr(ws.Range(RDV2_CELL_LOG).Value))
        LoadExpected
    End If
    t0 = Rdv2Ticks()
    RunOne key, t0, 0, 0
End Sub

Public Sub RDV2_Reset()
    Rdv2PaintFrame
    Vw().Range(RDV2_CELL_STOP).Value = ""
End Sub
