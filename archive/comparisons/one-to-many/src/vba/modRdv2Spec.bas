Attribute VB_Name = "modRdv2Spec"
'==============================================================================
' modRdv2Spec -- constants, clock, hash and sheet layout for the one-to-many
' build.
'
' Twin of src\v2\csharp\Rdv2Core.cs, and it has to stay a twin: the four
' methods only mean anything if VBA and C# do the same work in the same order.
' src\vba is the frozen 1:1 build and is not touched by any of this.
'
' Ten stages here, seven there. Reading and indexing are timed separately in
' this build, because what the index costs is the whole question.
'
' The clock is QueryPerformanceCounter received into Currency. VBA's Timer is a
' Single, so after midday one ULP is about 3.9 ms and a 30 ms stage cannot be
' measured at all.
'==============================================================================
Option Explicit

Public Const RDV2_KEYLEN As Long = 8
Public Const RDV2_FIELDS As Long = 10
Public Const RDV2_HASHMASK As Long = &H1FFFFF     ' 21 bits
Public Const RDV2_MODV As Double = 1000000007#
Public Const RDV2_COMMA As Byte = 44
Public Const RDV2_MAXCAND As Long = 256

Public Const RDV2_STAGES As Long = 10
Public Const RDV2_ST_READA As Long = 0
Public Const RDV2_ST_READB As Long = 1
Public Const RDV2_ST_READC As Long = 2
Public Const RDV2_ST_IDXA As Long = 3
Public Const RDV2_ST_IDXB As Long = 4
Public Const RDV2_ST_IDXC As Long = 5
Public Const RDV2_ST_JOINAB As Long = 6
Public Const RDV2_ST_JOINBC As Long = 7
Public Const RDV2_ST_SEARCH As Long = 8
Public Const RDV2_ST_SHOW As Long = 9

' detection: identical to Rdv2Watch.cs
Public Const RDV2_POLL_MS As Long = 40
Public Const RDV2_STABLE_MS As Long = 120

' sheet layout
Public Const RDV2_SHEET As String = "VIEW"
Public Const RDV2_CELL_STATE As String = "C2"
Public Const RDV2_CELL_NOTEPAD As String = "C3"
Public Const RDV2_CELL_DATA As String = "C4"
Public Const RDV2_CELL_INDEX As String = "C5"
Public Const RDV2_CELL_KEY As String = "C7"
Public Const RDV2_CELL_VERDICT As String = "E7"
Public Const RDV2_RG_KEY As String = "C7:H7"
Public Const RDV2_RG_RECORD As String = "B10:H19"
Public Const RDV2_RG_STAGE As String = "J10:L26"
Public Const RDV2_HIST_TOP As Long = 23
Public Const RDV2_HIST_ROWS As Long = 20
Public Const RDV2_CAND_TOP As Long = 45
Public Const RDV2_CAND_ROWS As Long = 10
Public Const RDV2_COL_FIRST As String = "B"
Public Const RDV2_COL_LAST As String = "I"
Public Const RDV2_CELL_ERROR As String = "C57"
Public Const RDV2_CELL_STOP As String = "C58"
Public Const RDV2_CELL_LOG As String = "C59"
Public Const RDV2_CELL_DATADIR As String = "C60"
Public Const RDV2_CELL_AUTOPICK As String = "C61"
Public Const RDV2_CELL_MANUAL As String = "F2"

#If VBA7 Then
    Public Declare PtrSafe Function Rdv2Qpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef c As Currency) As Long
    Public Declare PtrSafe Function Rdv2Qpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef f As Currency) As Long
    Public Declare PtrSafe Sub Rdv2Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Public Declare PtrSafe Function Rdv2ForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As LongPtr
    Public Declare PtrSafe Function Rdv2FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Public Declare PtrSafe Function Rdv2IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function Rdv2GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
        ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
#Else
    Public Declare Function Rdv2Qpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef c As Currency) As Long
    Public Declare Function Rdv2Qpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef f As Currency) As Long
    Public Declare Sub Rdv2Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Public Declare Function Rdv2ForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
    Public Declare Function Rdv2FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hwndParent As Long, ByVal hwndChildAfter As Long, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Public Declare Function Rdv2IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
    Public Declare Function Rdv2GetWindowText Lib "user32" Alias "GetWindowTextW" ( _
        ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
#End If

Private m_Freq As Currency

Public Function Rdv2Ticks() As Currency
    Dim c As Currency
    Rdv2Qpc c
    Rdv2Ticks = c
End Function

Public Function Rdv2MsBetween(ByVal t0 As Currency, ByVal t1 As Currency) As Double
    If m_Freq = 0 Then Rdv2Qpf m_Freq
    Rdv2MsBetween = (CDbl(t1) - CDbl(t0)) * 1000# / CDbl(m_Freq)
End Function

Public Function Rdv2MsSince(ByVal t0 As Currency) As Double
    Rdv2MsSince = Rdv2MsBetween(t0, Rdv2Ticks())
End Function

' smallest power of two at least twice the row count, capped at what the 21-bit
' hash can address. Rdv2Spec.Capacity() in C# computes the same number.
Public Function Rdv2Capacity(ByVal rows As Long) As Long
    Dim cap As Long
    cap = 1024
    Do While cap < rows * 2 And cap < 2097152
        cap = cap * 2
    Loop
    Rdv2Capacity = cap
End Function

' The canonical hash, same polynomial as Rdv2Spec.Hash in C#. clsRdv2Idx
' (hash flavour) inlines these nine lines in its two hot loops because a VBA
' call costs more than the hash itself; build\build_workbooks2.ps1 checks that
' every copy is still identical to this one, character for character.
Public Function Rdv2HashBytes(ByRef b() As Byte, ByVal p As Long) As Long
    Dim h As Long
    h = b(p)
    h = ((h * 131) + b(p + 1)) And RDV2_HASHMASK
    h = ((h * 131) + b(p + 2)) And RDV2_HASHMASK
    h = ((h * 131) + b(p + 3)) And RDV2_HASHMASK
    h = ((h * 131) + b(p + 4)) And RDV2_HASHMASK
    h = ((h * 131) + b(p + 5)) And RDV2_HASHMASK
    h = ((h * 131) + b(p + 6)) And RDV2_HASHMASK
    h = ((h * 131) + b(p + 7)) And RDV2_HASHMASK
    h = (h Xor (h \ 1024)) And RDV2_HASHMASK
    Rdv2HashBytes = h
End Function

#If VBA7 Then
Public Function Rdv2WindowTitle(ByVal h As LongPtr) As String
#Else
Public Function Rdv2WindowTitle(ByVal h As Long) As String
#End If
    Dim buf As String
    Dim n As Long
    buf = String$(320, vbNullChar)
    n = Rdv2GetWindowText(h, StrPtr(buf), 320)
    If n > 0 Then
        Rdv2WindowTitle = Left$(buf, n)
    Else
        Rdv2WindowTitle = ""
    End If
End Function

Public Function Rdv2IsKey(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) <> RDV2_KEYLEN Then
        Rdv2IsKey = False
        Exit Function
    End If
    For i = 1 To RDV2_KEYLEN
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then
            Rdv2IsKey = False
            Exit Function
        End If
    Next i
    Rdv2IsKey = True
End Function

' the last non-empty line of whatever the field holds
Public Function Rdv2Candidate(ByVal s As String) As String
    Dim e As Long, st As Long, ch As String
    If Len(s) = 0 Then
        Rdv2Candidate = ""
        Exit Function
    End If
    e = Len(s)
    Do While e > 0
        ch = Mid$(s, e, 1)
        If ch <> vbLf And ch <> vbCr And ch <> " " And ch <> vbTab Then Exit Do
        e = e - 1
    Loop
    If e = 0 Then
        Rdv2Candidate = ""
        Exit Function
    End If
    st = e
    Do While st > 1
        ch = Mid$(s, st - 1, 1)
        If ch = vbLf Or ch = vbCr Then Exit Do
        st = st - 1
    Loop
    Rdv2Candidate = Trim$(Mid$(s, st, e - st + 1))
End Function

Public Function Rdv2StageName(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv2StageName = "1 表A 読込"
        Case 1: Rdv2StageName = "2 表B 読込"
        Case 2: Rdv2StageName = "3 表C 読込"
        Case 3: Rdv2StageName = "4 表A 索引作成"
        Case 4: Rdv2StageName = "5 表B 索引作成"
        Case 5: Rdv2StageName = "6 表C 索引作成"
        Case 6: Rdv2StageName = "7 A-B 結合 (番号1)"
        Case 7: Rdv2StageName = "8 B-C 結合 (番号2)"
        Case 8: Rdv2StageName = "9 番号1 検索 (候補抽出)"
        Case 9: Rdv2StageName = "10 表示準備"
        Case Else: Rdv2StageName = ""
    End Select
End Function

Public Function Rdv2StageKey(ByVal i As Long) As String
    Select Case i
        Case 0: Rdv2StageKey = "readA"
        Case 1: Rdv2StageKey = "readB"
        Case 2: Rdv2StageKey = "readC"
        Case 3: Rdv2StageKey = "idxA"
        Case 4: Rdv2StageKey = "idxB"
        Case 5: Rdv2StageKey = "idxC"
        Case 6: Rdv2StageKey = "joinAB"
        Case 7: Rdv2StageKey = "joinBC"
        Case 8: Rdv2StageKey = "search"
        Case 9: Rdv2StageKey = "present"
        Case Else: Rdv2StageKey = ""
    End Select
End Function
