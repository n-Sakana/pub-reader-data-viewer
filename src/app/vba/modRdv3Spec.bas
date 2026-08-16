Attribute VB_Name = "modRdv3Spec"
'==============================================================================
' modRdv3Spec -- constants, clock, sheet addresses and helpers for the
' practical VBA build.
'
' The practical build is NOT the benchmark. src\vba and src\v2\vba are frozen;
' this directory copies from them where the job is the same (clock, detection
' rules) and never changes them.
'
' One index method only: Scripting.Dictionary, late bound. There is no
' hand-built hash here and no fallback to one.
'
' NO WIN32 AND NO SHELL ANYWHERE IN THIS BUILD. The practical VBA build runs on
' VBA + COM only; the frozen comparison builds (src\vba, src\v2\vba) keep their
' Declare blocks and are not touched. The two services Win32 used to provide
' here are replaced in kind, both measured:
'
'   clock  QueryPerformanceCounter -> VBA Timer. Timer is a Single, so it
'          advances in steps of 1/256 s (measured on this machine: 3.906 ms);
'          differences are taken in Double and corrected for the midnight wrap.
'          Every figure this app reports -- merge ~1.3 s, compose ~1.8 s, state
'          load ~0.4 s, search 0.05-1 s, detect latency ~0.14 s -- is two to
'          three orders above that step. Sub-4 ms figures in the BE log are
'          therefore quantised, and that is the whole cost.
'   sleep  kernel32 Sleep -> a WMI event source waited on with a timeout.
'          Application.Wait cannot do it (measured: one-second granularity,
'          aligned to the second boundary -- a +40 ms target returns at once,
'          a +400 ms target waited 861 ms), and Application.OnTime is worse
'          (a sub-second target fires in ~1 ms, which is a spin). NextEvent
'          blocks for its timeout and then raises wbemErrTimedout: a real
'          wait with the thread parked. 40 ms requested measured 45-54 ms.
'==============================================================================
Option Explicit

Public Const RDV3_KEYLEN As Long = 8
Public Const RDV3_FIELDS As Long = 10
Public Const RDV3_CONTENT_COLS As Long = 28
Public Const RDV3_LED_COLS As Long = 29          ' processed + 28 content
Public Const RDV3_MAXSHOW As Long = 10           ' candidate rows on screen
Public Const RDV3_MODV As Double = 1000000007#
Public Const RDV3_COMMA As Byte = 44

' detection: identical to the frozen builds
Public Const RDV3_POLL_MS As Long = 40
Public Const RDV3_STABLE_MS As Long = 120

' the eight merge stages whose sum is the merge time on screen
Public Const RDV3_STAGES As Long = 8
Public Const RDV3_ST_READA As Long = 0
Public Const RDV3_ST_READB As Long = 1
Public Const RDV3_ST_READC As Long = 2
Public Const RDV3_ST_IDXA As Long = 3
Public Const RDV3_ST_IDXB As Long = 4
Public Const RDV3_ST_IDXC As Long = 5
Public Const RDV3_ST_JOINAB As Long = 6
Public Const RDV3_ST_JOINBC As Long = 7

Public Const RDV3_WORKER_TIMEOUT_S As Double = 180#

' FE book sheets (the FE holds NO ledger; it is a small UI-only book)
Public Const RDV3_SHEET_UI As String = "UI"
Public Const RDV3_SHEET_META As String = "META"

' ledger workbook (BE-owned; the FE never opens it): LEDGER sheet name matches
' the C# build's viewing ledger, plus a small META sheet
Public Const RDV3_LB_SHEET As String = "LEDGER"
Public Const RDV3_LB_META As String = "META"
Public Const RDV3_LM_ROWS As String = "B1"
Public Const RDV3_LM_CHECKSUM As String = "B2"
Public Const RDV3_LM_MARKS As String = "B3"
Public Const RDV3_LM_SAVED As String = "B4"
Public Const RDV3_LM_FORMAT As String = "B5"

' UI sheet addresses (the builder paints these; the code writes them)
Public Const RDV3_C_STATE As String = "C3"
Public Const RDV3_C_NOTEPAD As String = "C4"
Public Const RDV3_C_LEDINFO As String = "C5"
Public Const RDV3_C_MERGEMS As String = "C6"
Public Const RDV3_C_SEARCHMS As String = "E6"
Public Const RDV3_C_ERROR As String = "C7"
Public Const RDV3_C_INPUT As String = "C9"
Public Const RDV3_BTN_SEARCH As String = "E9"
Public Const RDV3_BTN_CLEAR As String = "F9"
Public Const RDV3_BTN_PROCESSED As String = "G9"
Public Const RDV3_BTN_WATCH As String = "I9"
Public Const RDV3_C_KEY As String = "C11"
Public Const RDV3_C_VERDICT As String = "E11"
Public Const RDV3_C_PROCSTATE As String = "I11"
Public Const RDV3_REC_TOP As Long = 14           ' record rows 14..23, cols B..H
Public Const RDV3_C_CANDNOTE As String = "B25"
Public Const RDV3_CAND_TOP As Long = 27          ' candidate rows 27..36, cols B..J
Public Const RDV3_CAND_LABEL As String = "候補一覧 (行頭の「選択」で表示)"

' META sheet addresses
Public Const RDV3_M_ROWS As String = "C1"
Public Const RDV3_M_UPDATED As String = "C2"
Public Const RDV3_M_FORMAT As String = "C3"
Public Const RDV3_M_DATADIR As String = "C4"
Public Const RDV3_M_LOGPATH As String = "C5"
Public Const RDV3_M_WORKER_TOP As Long = 8       ' worker.xlsm base64 in E8 down

' wait primitive state (module level: declarations precede procedures)
Private m_evSrc As Object
Private m_evDead As Boolean
Private m_evShort As Long
Private m_evRebuilt As Boolean

' ---- the rule this SESSION is running on ------------------------------------
' The constants above are the built-in defaults -- the values the shipped
' ReaderDataViewer.json carries. What is actually in force comes out of that file
' at start-up (modRdv3Cfg) and is settled ONCE, here, in both processes.
'
' The key length in particular is one length for the whole program: the width of
' the key column in the CSVs, the width the index is built on, and the width the
' screen accepts. They cannot differ, or the operator types a number the index
' can never answer -- which is why it is settled before the CSVs are read and
' never changed under a running session. key.digitsOnly is not like that: it only
' says which CHARACTERS count, nothing built at start-up depends on it, and the
' settings screen applies it at once.
Private m_keyLen As Long
Private m_keyDigits As Boolean
Private m_maxShow As Long
Private m_pollMs As Long
Private m_stableMs As Long
Private m_rebindMs As Long
Private m_specSet As Boolean

'------------------------------------------------------------------------------
' clock: VBA Timer, in Double, wrap-corrected. See the module header for what
' this costs against the QueryPerformanceCounter it replaces.
'------------------------------------------------------------------------------
Public Function Rdv3Ticks() As Double
    Rdv3Ticks = Timer
End Function

Public Function Rdv3MsBetween(ByVal t0 As Double, ByVal t1 As Double) As Double
    Dim d As Double
    d = t1 - t0
    If d < 0 Then d = d + 86400#                 ' the clock passed midnight
    Rdv3MsBetween = d * 1000#
End Function

Public Function Rdv3MsSince(ByVal t0 As Double) As Double
    Rdv3MsSince = Rdv3MsBetween(t0, Timer)
End Function

'------------------------------------------------------------------------------
' sub-second blocking wait, without Win32 and without a spin.
'
' The source is an extrinsic WMI event class this machine does not raise on its
' own: nothing is polled on its behalf, so the source costs nothing while we sit
' in NextEvent. NextEvent(ms) parks the thread for the timeout and then raises
' wbemErrTimedout, which is the wait.
'
' Returns False when the primitive is NOT delivering waits. That is a hard
' failure, never a silent fall back to spinning: the BE tests it once at the
' start of its resident loop (Rdv3WaitReady) and then treats every later False
' as fatal, reporting it and stopping. m_evShort counts consecutive waits that
' came back far too early, which is how a WMI service that died under us is
' caught.
'------------------------------------------------------------------------------
Public Function Rdv3WaitMs(ByVal ms As Long) As Boolean
    Dim svc As Object
    Dim t0 As Double
    Dim el As Double

    If m_evDead Then Exit Function
    If ms <= 0 Then
        Rdv3WaitMs = True
        Exit Function
    End If
    If m_evSrc Is Nothing Then
        On Error GoTo Fail
        Set svc = GetObject("winmgmts:\\.\root\cimv2")
        Set m_evSrc = svc.ExecNotificationQuery("SELECT * FROM Win32_PowerManagementEvent")
        On Error GoTo 0
    End If

    t0 = Timer
    On Error Resume Next
    Err.Clear
    m_evSrc.NextEvent ms                          ' returns by timing out
    Err.Clear
    On Error GoTo 0
    el = Rdv3MsBetween(t0, Timer)

    ' A power event (or a broken source) can return early. One early return is
    ' legitimate; twenty in a row means we are not waiting at all any more.
    ' A WMI service that restarted under us is worth one rebuild of the source
    ' before giving up -- but only one, so a permanently broken WMI cannot turn
    ' this into a spin.
    If el < CDbl(ms) * 0.25 Then
        m_evShort = m_evShort + 1
        If m_evShort >= 20 Then
            m_evShort = 0
            Set m_evSrc = Nothing
            If m_evRebuilt Then
                m_evDead = True
                Exit Function
            End If
            m_evRebuilt = True
        End If
    Else
        m_evShort = 0
    End If
    Rdv3WaitMs = True
    Exit Function
Fail:
    m_evDead = True
End Function

' One-off self test: ask for 40 ms and insist on getting at least half of it.
' The first call builds the event source, so it is not the one measured --
' otherwise the setup cost alone could pass the test for a wait that returns
' instantly.
Public Function Rdv3WaitReady(ByRef why As String) As Boolean
    Dim t0 As Double
    Dim el As Double
    why = ""
    If Not Rdv3WaitMs(40) Then
        why = "WMI の待機イベントソースを作成できません"
        Exit Function
    End If
    t0 = Timer
    If Not Rdv3WaitMs(40) Then
        why = "WMI の待機が失敗しました"
        Exit Function
    End If
    el = Rdv3MsBetween(t0, Timer)
    If el < 20# Then
        why = "待機が効いていません (40 ms 要求で " & Format$(el, "0.0") & " ms)"
        Exit Function
    End If
    Rdv3WaitReady = True
End Function

'------------------------------------------------------------------------------
' the session rule
'------------------------------------------------------------------------------
Private Sub SpecDefaults()
    m_keyLen = RDV3_KEYLEN
    m_keyDigits = True
    m_maxShow = RDV3_MAXSHOW
    m_pollMs = RDV3_POLL_MS
    m_stableMs = RDV3_STABLE_MS
    m_rebindMs = 400
    m_specSet = True
End Sub

' Called once per process, right after the settings file is read, before the
' CSVs are touched. candidateRowsShown is capped at the number of candidate rows
' the sheet actually has: the setting may ask for more, and painting rows that do
' not exist would be a silent lie about how many were shown.
Public Sub Rdv3SpecApply(ByVal keyLen As Long, ByVal digitsOnly As Boolean, _
                         ByVal candidateRows As Long, ByVal pollMs As Long, _
                         ByVal stableMs As Long, ByVal rebindMs As Long)
    SpecDefaults
    If keyLen >= 1 And keyLen <= 64 Then m_keyLen = keyLen
    m_keyDigits = digitsOnly
    If candidateRows >= 1 Then
        If candidateRows > RDV3_MAXSHOW Then
            m_maxShow = RDV3_MAXSHOW
        Else
            m_maxShow = candidateRows
        End If
    End If
    If pollMs >= 5 And pollMs <= 5000 Then m_pollMs = pollMs
    If stableMs >= 0 And stableMs <= 60000 Then m_stableMs = stableMs
    If rebindMs >= 50 And rebindMs <= 60000 Then m_rebindMs = rebindMs
End Sub

' What a running session may adopt from a re-read settings file. The key LENGTH
' is deliberately absent: the CSVs were read and the index was built with the one
' this session started on, and a shorter key would leave the screen asking for
' numbers that index can never answer.
Public Sub Rdv3SpecApplyLive(ByVal digitsOnly As Boolean, ByVal candidateRows As Long, _
                             ByVal pollMs As Long, ByVal stableMs As Long, _
                             ByVal rebindMs As Long)
    If Not m_specSet Then SpecDefaults
    m_keyDigits = digitsOnly
    If candidateRows >= 1 Then
        If candidateRows > RDV3_MAXSHOW Then
            m_maxShow = RDV3_MAXSHOW
        Else
            m_maxShow = candidateRows
        End If
    End If
    If pollMs >= 5 And pollMs <= 5000 Then m_pollMs = pollMs
    If stableMs >= 0 And stableMs <= 60000 Then m_stableMs = stableMs
    If rebindMs >= 50 And rebindMs <= 60000 Then m_rebindMs = rebindMs
End Sub

' digitsOnly on its own, so the settings screen can apply it without a restart
Public Sub Rdv3SpecSetDigitsOnly(ByVal digitsOnly As Boolean)
    If Not m_specSet Then SpecDefaults
    m_keyDigits = digitsOnly
End Sub

Public Function Rdv3KeyLen() As Long
    If Not m_specSet Then SpecDefaults
    Rdv3KeyLen = m_keyLen
End Function

Public Function Rdv3KeyDigits() As Boolean
    If Not m_specSet Then SpecDefaults
    Rdv3KeyDigits = m_keyDigits
End Function

Public Function Rdv3MaxShow() As Long
    If Not m_specSet Then SpecDefaults
    Rdv3MaxShow = m_maxShow
End Function

Public Function Rdv3PollMs() As Long
    If Not m_specSet Then SpecDefaults
    Rdv3PollMs = m_pollMs
End Function

Public Function Rdv3StableMs() As Long
    If Not m_specSet Then SpecDefaults
    Rdv3StableMs = m_stableMs
End Function

' how often an UNBOUND target is looked for again. Binding walks the focused
' element's ancestors and may ask WMI for a process name, so doing it at the poll
' cadence would be 25 of those a second for as long as one target is unbound.
Public Function Rdv3RebindMs() As Long
    If Not m_specSet Then SpecDefaults
    Rdv3RebindMs = m_rebindMs
End Function

Public Function Rdv3IsKey(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    Dim n As Long
    n = Rdv3KeyLen()
    If Len(s) <> n Then
        Rdv3IsKey = False
        Exit Function
    End If
    If Not Rdv3KeyDigits() Then
        Rdv3IsKey = True
        Exit Function
    End If
    For i = 1 To n
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then
            Rdv3IsKey = False
            Exit Function
        End If
    Next i
    Rdv3IsKey = True
End Function

' the last non-empty line of whatever the notepad field holds -- same rule as
' the frozen builds
Public Function Rdv3Candidate(ByVal s As String) As String
    Dim e As Long, st As Long, ch As String
    If Len(s) = 0 Then
        Rdv3Candidate = ""
        Exit Function
    End If
    e = Len(s)
    Do While e > 0
        ch = Mid$(s, e, 1)
        If ch <> vbLf And ch <> vbCr And ch <> " " And ch <> vbTab Then Exit Do
        e = e - 1
    Loop
    If e = 0 Then
        Rdv3Candidate = ""
        Exit Function
    End If
    st = e
    Do While st > 1
        ch = Mid$(s, st - 1, 1)
        If ch = vbLf Or ch = vbCr Then Exit Do
        st = st - 1
    Loop
    Rdv3Candidate = Trim$(Mid$(s, st, e - st + 1))
End Function

Public Function Rdv3StageKey(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv3StageKey = "readA"
        Case 1: Rdv3StageKey = "readB"
        Case 2: Rdv3StageKey = "readC"
        Case 3: Rdv3StageKey = "idxA"
        Case 4: Rdv3StageKey = "idxB"
        Case 5: Rdv3StageKey = "idxC"
        Case 6: Rdv3StageKey = "joinAB"
        Case 7: Rdv3StageKey = "joinBC"
        Case Else: Rdv3StageKey = ""
    End Select
End Function

' ledger column names: "processed" head + 28 content columns. Must stay
' identical to Rdv3Ledger.ContentHead in the C# build and to the LEDGER sheet
' header row the builder writes.
Public Function Rdv3LedgerHead(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv3LedgerHead = "処理済み"
        Case 1: Rdv3LedgerHead = "key1"
        Case 2: Rdv3LedgerHead = "key2"
        Case 3: Rdv3LedgerHead = "a_code"
        Case 4: Rdv3LedgerHead = "a_name"
        Case 5: Rdv3LedgerHead = "a_grade"
        Case 6: Rdv3LedgerHead = "a_date"
        Case 7: Rdv3LedgerHead = "a_amount"
        Case 8: Rdv3LedgerHead = "a_rate"
        Case 9: Rdv3LedgerHead = "a_flag"
        Case 10: Rdv3LedgerHead = "a_dept"
        Case 11: Rdv3LedgerHead = "a_note"
        Case 12: Rdv3LedgerHead = "b_slip"
        Case 13: Rdv3LedgerHead = "b_date"
        Case 14: Rdv3LedgerHead = "b_qty"
        Case 15: Rdv3LedgerHead = "b_unit"
        Case 16: Rdv3LedgerHead = "b_total"
        Case 17: Rdv3LedgerHead = "b_status"
        Case 18: Rdv3LedgerHead = "b_line"
        Case 19: Rdv3LedgerHead = "b_memo"
        Case 20: Rdv3LedgerHead = "c_item"
        Case 21: Rdv3LedgerHead = "c_maker"
        Case 22: Rdv3LedgerHead = "c_cat"
        Case 23: Rdv3LedgerHead = "c_price"
        Case 24: Rdv3LedgerHead = "c_stock"
        Case 25: Rdv3LedgerHead = "c_loc"
        Case 26: Rdv3LedgerHead = "c_lot"
        Case 27: Rdv3LedgerHead = "c_exp"
        Case 28: Rdv3LedgerHead = "c_remark"
        Case Else: Rdv3LedgerHead = ""
    End Select
End Function

' the 10 field names of each table's record view (key1/key2 re-shown in the
' tables they belong to; same unfolding as the C# build)
Public Function Rdv3NameOfA(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv3NameOfA = "key1"
        Case Else: Rdv3NameOfA = Rdv3LedgerHead(2 + i)
    End Select
End Function

Public Function Rdv3NameOfB(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv3NameOfB = "key1"
        Case 1: Rdv3NameOfB = "key2"
        Case Else: Rdv3NameOfB = Rdv3LedgerHead(10 + i)
    End Select
End Function

Public Function Rdv3NameOfC(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv3NameOfC = "key2"
        Case Else: Rdv3NameOfC = Rdv3LedgerHead(19 + i)
    End Select
End Function

' Process identity for the log. GetCurrentProcessId is Win32; Excel's own main
' window handle is a COM property and identifies the instance just as well
' (the FE logs its own, and its BE's, so a run can still be traced end to end).
Public Function Rdv3SelfId() As String
    On Error Resume Next
    Rdv3SelfId = "hwnd:" & CStr(Application.hwnd)
    On Error GoTo 0
End Function

Public Function Rdv3FmtMs(ByVal ms As Double) As String
    Rdv3FmtMs = Format$(ms, "#,##0.0") & " ms"
End Function

' the sidecar state file next to the ledger workbook: header line, then one
' line per ledger row ("0/1" processed TAB the 28 content columns). It is the
' ledger-side canonical metadata written atomically together with every ledger
' save, so a no-difference boot never has to open the 14 MB workbook at all.
Public Function Rdv3SidecarPath(ByVal ledgerPath As String) As String
    Dim dot As Long
    dot = InStrRev(ledgerPath, ".")
    If dot > 0 Then
        Rdv3SidecarPath = Left$(ledgerPath, dot - 1) & ".state"
    Else
        Rdv3SidecarPath = ledgerPath & ".state"
    End If
End Function
