Attribute VB_Name = "modRdvApp"
'==============================================================================
' modRdvApp -- the loop, the screen, and the measurement boundary.
'
'   t0            a number was confirmed (start of the merge-select)
'   stages 1..6   modRdvEngine: read three CSVs, join A-B, join B-C, look up
'   stage 7       write the record to the sheet
'   stop clock    total = now - t0
'   afterwards    write the timings, the history line and the log
'
' The timings are written only after the clock has stopped, so the screen is
' never timing the drawing of its own stopwatch. Nothing touches the status bar
' while a run is in flight.
'
' Detection uses the same two rules as the C# builds: exactly 8 digits, and the
' value must survive RDV_STABLE_MS of polling. The same number does not fire
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
Private m_dataDir As String
Private m_logPath As String
Private m_expRows As Long
Private m_expChk As Double
Private m_expLoaded As Boolean
Private m_oracleOk As Boolean
Private m_oracleNote As String

Private Function Vw() As Object
    Set Vw = ThisWorkbook.Worksheets(RDV_SHEET)
End Function

Public Function RdvDataDir() As String
    Dim ws As Object
    Dim c As String
    Dim p As String
    Set ws = Vw()
    c = Trim$(CStr(ws.Range(RDV_CELL_DATA).Value))
    If Len(c) > 0 Then
        If Dir$(c & "\tableA.csv") <> "" Then
            RdvDataDir = c
            Exit Function
        End If
    End If
    p = ThisWorkbook.path
    If Dir$(p & "\..\data\tableA.csv") <> "" Then
        RdvDataDir = p & "\..\data"
        Exit Function
    End If
    If Dir$(p & "\data\tableA.csv") <> "" Then
        RdvDataDir = p & "\data"
        Exit Function
    End If
    RdvDataDir = ""
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
Public Sub RdvPaintFrame()
    Dim ws As Object
    Set ws = Vw()
    ws.Range(RDV_RG_KEY).ClearContents
    ws.Range(RDV_RG_RECORD).ClearContents
    ws.Range(RDV_RG_STAGE).ClearContents
    ws.Range(RDV_HIST_COL & RDV_HIST_TOP & ":" & RDV_HIST_COLEND & (RDV_HIST_TOP + RDV_HIST_ROWS - 1)).ClearContents
    ws.Range(RDV_CELL_ERROR).Value = ""
    m_seq = 0
End Sub

Private Sub SetState(ByVal state As String, ByVal detail As String)
    Dim ws As Object
    On Error Resume Next
    Set ws = Vw()
    ws.Range("C2").Value = state
    If Len(detail) > 0 Then
        ws.Range("C3").Value = detail
        ws.Range("C4").Value = m_dataDir
    End If
End Sub

' stage 7: the record, and nothing else
Private Sub ShowRecord(ByVal key As String)
    Dim ws As Object
    Dim rec(1 To 10, 1 To 7) As Variant
    Dim top(1 To 1, 1 To 3) As Variant
    Dim i As Long
    Dim dummy As Variant

    Set ws = Vw()
    top(1, 1) = key
    top(1, 2) = ""
    If Len(RdvError()) > 0 Then
        top(1, 3) = "エラー"
    ElseIf RdvFound() Then
        top(1, 3) = "一致 1 件   key2 = " & RdvKey2()
    Else
        top(1, 3) = "該当なし"
    End If
    For i = 0 To RDV_FIELDS - 1
        rec(i + 1, 1) = i + 1
        rec(i + 1, 2) = RdvNameOf(0, i)
        rec(i + 1, 3) = RdvValueOf(0, i)
        rec(i + 1, 4) = RdvNameOf(1, i)
        rec(i + 1, 5) = RdvValueOf(1, i)
        rec(i + 1, 6) = RdvNameOf(2, i)
        rec(i + 1, 7) = RdvValueOf(2, i)
    Next i
    ws.Range(RDV_RG_KEY).Value2 = top
    ws.Range(RDV_RG_RECORD).Value2 = rec
    dummy = ws.Range(RDV_CELL_COMMIT).Value2
End Sub

' after the clock has stopped
Private Sub ShowStats(ByVal key As String)
    Dim ws As Object
    Dim st(1 To 12, 1 To 3) As Variant
    Dim hist(1 To 1, 1 To 6) As Variant
    Dim i As Long, r As Long
    Dim total As Double, other As Double

    Set ws = Vw()
    total = m_totalMs
    For i = 0 To RDV_STAGES - 1
        st(i + 1, 1) = RdvStageName(i)
        st(i + 1, 2) = RdvStageMs(i)
        If total > 0 Then st(i + 1, 3) = 100# * RdvStageMs(i) / total Else st(i + 1, 3) = 0
    Next i
    other = total - RdvStageSum()
    st(8, 1) = "その他 (差分)"
    st(8, 2) = other
    If total > 0 Then st(8, 3) = 100# * other / total Else st(8, 3) = 0
    st(9, 1) = "合計 (merge-select)"
    st(9, 2) = total
    st(9, 3) = 100
    st(10, 1) = "検知遅延 (参考・計測外)"
    st(10, 2) = m_detectMs
    st(10, 3) = ""
    st(11, 1) = "データ行数"
    st(11, 2) = RdvRows()
    st(11, 3) = ""
    st(12, 1) = "プローブ数"
    st(12, 2) = RdvProbes()
    st(12, 3) = ""
    ws.Range(RDV_RG_STAGE).Value2 = st

    hist(1, 1) = m_seq
    hist(1, 2) = Format$(Now, "hh:nn:ss")
    hist(1, 3) = key
    hist(1, 4) = total
    hist(1, 5) = m_detectMs
    If Len(RdvError()) > 0 Then
        hist(1, 6) = "エラー"
    ElseIf m_oracleOk Then
        hist(1, 6) = "検算 OK"
    Else
        hist(1, 6) = "検算 NG " & m_oracleNote
    End If
    r = ((m_seq - 1) Mod RDV_HIST_ROWS) + RDV_HIST_TOP
    ws.Range(RDV_HIST_COL & r & ":" & RDV_HIST_COLEND & r).Value2 = hist
    ws.Range(RDV_CELL_ERROR).Value = RdvError()
End Sub

Private Sub CheckOracle()
    m_oracleOk = True
    m_oracleNote = ""
    If Not m_expLoaded Then
        m_oracleNote = "expected.txt なし"
        Exit Sub
    End If
    If m_expRows <> RdvRows() Then
        m_oracleOk = False
        m_oracleNote = "rows"
    ElseIf m_expChk <> RdvChecksum() Then
        m_oracleOk = False
        m_oracleNote = "checksum"
    ElseIf RdvMatchedAB() <> RdvRows() Or RdvMatchedBC() <> RdvRows() Then
        m_oracleOk = False
        m_oracleNote = "matched"
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
        s = "seq" & vbTab & "time" & vbTab & "key"
        For i = 0 To RDV_STAGES - 1
            s = s & vbTab & RdvStageKey(i)
        Next i
        s = s & vbTab & "other" & vbTab & "total" & vbTab & "detect" & vbTab & "polls" & vbTab & _
            "rows" & vbTab & "probes" & vbTab & "checksum" & vbTab & "matchedAB" & vbTab & _
            "matchedBC" & vbTab & "oracle" & vbTab & "error"
        Print #f, s
    End If
    s = m_seq & vbTab & Format$(Now, "hh:nn:ss") & vbTab & key
    For i = 0 To RDV_STAGES - 1
        s = s & vbTab & Format$(RdvStageMs(i), "0.00")
    Next i
    s = s & vbTab & Format$(m_totalMs - RdvStageSum(), "0.00")
    s = s & vbTab & Format$(m_totalMs, "0.00")
    s = s & vbTab & Format$(m_detectMs, "0.00")
    s = s & vbTab & m_polls
    s = s & vbTab & RdvRows()
    s = s & vbTab & Format$(RdvProbes(), "0")
    s = s & vbTab & Format$(RdvChecksum(), "0")
    s = s & vbTab & RdvMatchedAB()
    s = s & vbTab & RdvMatchedBC()
    If m_oracleOk Then
        s = s & vbTab & "ok"
    Else
        s = s & vbTab & "NG:" & m_oracleNote
    End If
    s = s & vbTab & Replace(Replace(RdvError(), vbTab, " "), vbCrLf, " ")
    Print #f, s
    Close #f
End Sub

'------------------------------------------------------------------------------
' one merge-select
'------------------------------------------------------------------------------
Private Sub RunOne(ByVal key As String, ByVal t0 As Currency, ByVal detectMs As Double, ByVal polls As Long)
    Dim t As Currency
    Dim ok As Boolean

    m_seq = m_seq + 1
    m_detectMs = detectMs
    m_polls = polls
    SetState "処理中", ""

    ok = RdvRunMergeSelect(m_dataDir, key)

    t = RdvTicks()
    ShowRecord key
    RdvSetStageMs RDV_ST_SHOW, RdvMsSince(t)
    m_totalMs = RdvMsSince(t0)

    CheckOracle
    ShowStats key
    AppendLog key
    SetState "監視中", ""
End Sub

'------------------------------------------------------------------------------
' the loop
'------------------------------------------------------------------------------
Public Sub RDV_StartMonitor()
    Dim ws As Object
    Dim raw As Variant
    Dim cand As String
    Dim held As Double
    Dim t0 As Currency
    Dim stopped As Boolean

    If m_running Then Exit Sub
    On Error GoTo Fail
    Application.EnableCancelKey = xlErrorHandler

    Set ws = Vw()
    m_dataDir = RdvDataDir()
    m_logPath = Trim$(CStr(ws.Range(RDV_CELL_LOG).Value))
    If Len(m_dataDir) = 0 Then
        ws.Range(RDV_CELL_ERROR).Value = "データが見つかりません。build\gen_data.ps1 を先に実行してください。"
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
        If UCase$(Trim$(CStr(ws.Range(RDV_CELL_STOP).Value))) = "STOP" Then
            stopped = True
            Exit Do
        End If

        If Not RdvUiaBound() Then
            If RdvUiaBind() Then
                SetState "監視中", RdvUiaTitle() & "  (hwnd " & CStr(RdvUiaHwnd()) & ")"
                m_pending = ""
                m_pendPolls = 0
            Else
                SetState "メモ帳を待機中", "メモ帳: 未接続 -- " & RdvUiaWhy()
                DoEvents
                RdvSleep 400
            End If
        End If

        If RdvUiaBound() Then
            raw = RdvUiaRead()
            If Not IsNull(raw) Then
                cand = RdvCandidate(CStr(raw))
                If Len(cand) = 0 Then
                    m_sawEmpty = True
                    m_pending = ""
                    m_pendPolls = 0
                ElseIf cand <> m_pending Then
                    m_pending = cand
                    m_pendSince = RdvTicks()
                    m_pendPolls = 1
                Else
                    m_pendPolls = m_pendPolls + 1
                    held = RdvMsSince(m_pendSince)
                    If held >= RDV_STABLE_MS Then
                        If RdvIsKey(cand) Then
                            If cand <> m_lastFired Or m_sawEmpty Then
                                m_lastFired = cand
                                m_sawEmpty = False
                                t0 = RdvTicks()
                                RunOne cand, t0, RdvMsBetween(m_pendSince, t0), m_pendPolls
                            End If
                        End If
                    End If
                End If
            End If
            DoEvents
            RdvSleep RDV_POLL_MS
        End If
    Loop

    m_running = False
    SetState "停止", ""
    Exit Sub
Fail:
    m_running = False
    If Err.Number = 18 Then
        SetState "停止 (Esc)", ""
    Else
        SetState "エラー", ""
        Vw().Range(RDV_CELL_ERROR).Value = "監視エラー " & Err.Number & ": " & Err.Description
    End If
End Sub

Public Sub RDV_StartMonitorAsync()
    Vw().Range(RDV_CELL_STOP).Value = ""
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!RDV_StartMonitor"
End Sub

Public Sub RDV_StopMonitor()
    m_running = False
    Vw().Range(RDV_CELL_STOP).Value = "STOP"
End Sub

Public Sub RDV_ManualRun()
    Dim key As String
    Dim t0 As Currency
    key = Trim$(CStr(Vw().Range(RDV_CELL_MANUAL).Value))
    If Not RdvIsKey(key) Then
        Vw().Range(RDV_CELL_ERROR).Value = RDV_CELL_MANUAL & " に 8 桁の番号を入れてください。"
        Exit Sub
    End If
    If Len(m_dataDir) = 0 Then
        m_dataDir = RdvDataDir()
        m_logPath = Trim$(CStr(Vw().Range(RDV_CELL_LOG).Value))
        LoadExpected
    End If
    t0 = RdvTicks()
    RunOne key, t0, 0, 0
End Sub

Public Sub RDV_Reset()
    RdvReset
    RdvUiaReset
    RdvPaintFrame
    SetState "停止", "メモ帳: 未接続"
End Sub
