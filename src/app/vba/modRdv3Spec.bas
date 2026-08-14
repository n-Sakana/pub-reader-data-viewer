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
' The clock is QueryPerformanceCounter received into Currency, same as the
' frozen builds: VBA's Timer is a Single and cannot time a 30 ms stage.
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
Public Const RDV3_SPAWN_TIMEOUT_S As Double = 30#

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

' (Rdv3SidecarPath lives at the end of this module: procedures may not appear
' before the Declare block.)

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

#If VBA7 Then
    Public Declare PtrSafe Function Rdv3Qpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef c As Currency) As Long
    Public Declare PtrSafe Function Rdv3Qpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef f As Currency) As Long
    Public Declare PtrSafe Sub Rdv3Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Public Declare PtrSafe Function Rdv3GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
    Public Declare PtrSafe Function Rdv3GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" ( _
        ByVal hwnd As LongPtr, ByRef pid As Long) As Long
    Public Declare PtrSafe Function Rdv3OpenProcess Lib "kernel32" Alias "OpenProcess" ( _
        ByVal access As Long, ByVal inherit As Long, ByVal pid As Long) As LongPtr
    Public Declare PtrSafe Function Rdv3GetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" ( _
        ByVal hProc As LongPtr, ByRef code As Long) As Long
    Public Declare PtrSafe Function Rdv3TerminateProcess Lib "kernel32" Alias "TerminateProcess" ( _
        ByVal hProc As LongPtr, ByVal code As Long) As Long
    Public Declare PtrSafe Function Rdv3CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal h As LongPtr) As Long
    Public Declare PtrSafe Function Rdv3ForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As LongPtr
    Public Declare PtrSafe Function Rdv3FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Public Declare PtrSafe Function Rdv3IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function Rdv3GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
        ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
    Public Declare PtrSafe Function Rdv3MB2WC Lib "kernel32" Alias "MultiByteToWideChar" ( _
        ByVal cp As Long, ByVal flags As Long, ByVal mb As LongPtr, ByVal cbMb As Long, _
        ByVal wc As LongPtr, ByVal ccWc As Long) As Long
    Public Declare PtrSafe Function Rdv3WC2MB Lib "kernel32" Alias "WideCharToMultiByte" ( _
        ByVal cp As Long, ByVal flags As Long, ByVal wc As LongPtr, ByVal ccWc As Long, _
        ByVal mb As LongPtr, ByVal cbMb As Long, ByVal defChar As LongPtr, ByVal usedDef As LongPtr) As Long
#Else
    Public Declare Function Rdv3Qpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef c As Currency) As Long
    Public Declare Function Rdv3Qpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef f As Currency) As Long
    Public Declare Sub Rdv3Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Public Declare Function Rdv3GetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
    Public Declare Function Rdv3GetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" ( _
        ByVal hwnd As Long, ByRef pid As Long) As Long
    Public Declare Function Rdv3OpenProcess Lib "kernel32" Alias "OpenProcess" ( _
        ByVal access As Long, ByVal inherit As Long, ByVal pid As Long) As Long
    Public Declare Function Rdv3GetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" ( _
        ByVal hProc As Long, ByRef code As Long) As Long
    Public Declare Function Rdv3TerminateProcess Lib "kernel32" Alias "TerminateProcess" ( _
        ByVal hProc As Long, ByVal code As Long) As Long
    Public Declare Function Rdv3CloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal h As Long) As Long
    Public Declare Function Rdv3ForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
    Public Declare Function Rdv3FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hwndParent As Long, ByVal hwndChildAfter As Long, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Public Declare Function Rdv3IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
    Public Declare Function Rdv3GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
        ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
    Public Declare Function Rdv3MB2WC Lib "kernel32" Alias "MultiByteToWideChar" ( _
        ByVal cp As Long, ByVal flags As Long, ByVal mb As Long, ByVal cbMb As Long, _
        ByVal wc As Long, ByVal ccWc As Long) As Long
    Public Declare Function Rdv3WC2MB Lib "kernel32" Alias "WideCharToMultiByte" ( _
        ByVal cp As Long, ByVal flags As Long, ByVal wc As Long, ByVal ccWc As Long, _
        ByVal mb As Long, ByVal cbMb As Long, ByVal defChar As Long, ByVal usedDef As Long) As Long
#End If

Private m_Freq As Currency

Public Function Rdv3Ticks() As Currency
    Dim c As Currency
    Rdv3Qpc c
    Rdv3Ticks = c
End Function

Public Function Rdv3MsBetween(ByVal t0 As Currency, ByVal t1 As Currency) As Double
    If m_Freq = 0 Then Rdv3Qpf m_Freq
    Rdv3MsBetween = (CDbl(t1) - CDbl(t0)) * 1000# / CDbl(m_Freq)
End Function

Public Function Rdv3MsSince(ByVal t0 As Currency) As Double
    Rdv3MsSince = Rdv3MsBetween(t0, Rdv3Ticks())
End Function

Public Function Rdv3IsKey(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) <> RDV3_KEYLEN Then
        Rdv3IsKey = False
        Exit Function
    End If
    For i = 1 To RDV3_KEYLEN
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

#If VBA7 Then
Public Function Rdv3WindowTitle(ByVal h As LongPtr) As String
#Else
Public Function Rdv3WindowTitle(ByVal h As Long) As String
#End If
    Dim buf As String
    Dim n As Long
    buf = String$(320, vbNullChar)
    n = Rdv3GetWindowText(h, StrPtr(buf), 320)
    If n > 0 Then
        Rdv3WindowTitle = Left$(buf, n)
    Else
        Rdv3WindowTitle = ""
    End If
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

' is the process I started still alive?
Public Function Rdv3PidAlive(ByVal pid As Long) As Boolean
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Dim code As Long
    Rdv3PidAlive = False
    If pid = 0 Then Exit Function
    h = Rdv3OpenProcess(&H1000&, 0, pid)         ' PROCESS_QUERY_LIMITED_INFORMATION
    If h = 0 Then Exit Function
    If Rdv3GetExitCodeProcess(h, code) <> 0 Then
        If code = 259 Then Rdv3PidAlive = True   ' STILL_ACTIVE
    End If
    Rdv3CloseHandle h
End Function

Public Function Rdv3KillPid(ByVal pid As Long) As Boolean
    #If VBA7 Then
        Dim h As LongPtr
    #Else
        Dim h As Long
    #End If
    Rdv3KillPid = False
    If pid = 0 Then Exit Function
    h = Rdv3OpenProcess(&H1&, 0, pid)            ' PROCESS_TERMINATE
    If h = 0 Then Exit Function
    If Rdv3TerminateProcess(h, 1) <> 0 Then Rdv3KillPid = True
    Rdv3CloseHandle h
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
