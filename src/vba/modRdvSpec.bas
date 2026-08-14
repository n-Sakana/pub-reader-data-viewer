Attribute VB_Name = "modRdvSpec"
'==============================================================================
' modRdvSpec -- the constants, the clock, and the sheet layout.
'
' Everything here has a twin in src\csharp\RdvCore.cs and it has to stay a twin:
' the whole point of the comparison is that VBA and C# do the SAME work, so a
' difference in the numbers is a difference between the methods.
'
' The clock is QueryPerformanceCounter received into Currency. VBA's Timer is a
' Single, so after midday one ULP is about 3.9 ms and a 30 ms stage cannot be
' measured at all.
'==============================================================================
Option Explicit

Public Const RDV_KEYLEN As Long = 8
Public Const RDV_FIELDS As Long = 10
Public Const RDV_HASHMASK As Long = &H1FFFFF     ' 21 bits
Public Const RDV_HASHBITS As Long = 21
Public Const RDV_MODV As Double = 1000000007#
Public Const RDV_COMMA As Byte = 44

Public Const RDV_STAGES As Long = 7
Public Const RDV_ST_LOADA As Long = 0
Public Const RDV_ST_LOADB As Long = 1
Public Const RDV_ST_LOADC As Long = 2
Public Const RDV_ST_JOINAB As Long = 3
Public Const RDV_ST_JOINBC As Long = 4
Public Const RDV_ST_SELECT As Long = 5
Public Const RDV_ST_SHOW As Long = 6

' detection: identical to RdvWatch.cs
Public Const RDV_POLL_MS As Long = 40
Public Const RDV_STABLE_MS As Long = 120

' sheet layout: identical to RdvSheet in src\csharp\RdvUiExcel.cs
Public Const RDV_SHEET As String = "VIEW"
Public Const RDV_RG_HEADER As String = "C2:C4"
Public Const RDV_RG_KEY As String = "C6:E6"
Public Const RDV_RG_RECORD As String = "B9:H18"
Public Const RDV_CELL_COMMIT As String = "B9"
Public Const RDV_RG_STAGE As String = "J9:L20"
Public Const RDV_HIST_COL As String = "B"
Public Const RDV_HIST_COLEND As String = "G"
Public Const RDV_HIST_TOP As Long = 22
Public Const RDV_HIST_ROWS As Long = 20
Public Const RDV_CELL_ERROR As String = "C44"
Public Const RDV_CELL_STOP As String = "C45"
Public Const RDV_CELL_LOG As String = "C46"
Public Const RDV_CELL_DATA As String = "C47"
Public Const RDV_CELL_MANUAL As String = "F2"

#If VBA7 Then
    Public Declare PtrSafe Function RdvQpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef c As Currency) As Long
    Public Declare PtrSafe Function RdvQpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef f As Currency) As Long
    Public Declare PtrSafe Sub RdvSleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Public Declare PtrSafe Function RdvForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As LongPtr
    Public Declare PtrSafe Function RdvFindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
    Public Declare PtrSafe Function RdvIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As LongPtr) As Long
    Public Declare PtrSafe Function RdvGetWindowText Lib "user32" Alias "GetWindowTextW" ( _
        ByVal hwnd As LongPtr, ByVal lpString As LongPtr, ByVal cch As Long) As Long
#Else
    Public Declare Function RdvQpc Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef c As Currency) As Long
    Public Declare Function RdvQpf Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef f As Currency) As Long
    Public Declare Sub RdvSleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
    Public Declare Function RdvForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
    Public Declare Function RdvFindWindowEx Lib "user32" Alias "FindWindowExA" ( _
        ByVal hwndParent As Long, ByVal hwndChildAfter As Long, _
        ByVal lpszClass As String, ByVal lpszWindow As String) As Long
    Public Declare Function RdvIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
    Public Declare Function RdvGetWindowText Lib "user32" Alias "GetWindowTextW" ( _
        ByVal hwnd As Long, ByVal lpString As Long, ByVal cch As Long) As Long
#End If

Private m_Freq As Currency

Public Function RdvTicks() As Currency
    Dim c As Currency
    RdvQpc c
    RdvTicks = c
End Function

Public Function RdvMsBetween(ByVal t0 As Currency, ByVal t1 As Currency) As Double
    If m_Freq = 0 Then RdvQpf m_Freq
    RdvMsBetween = (CDbl(t1) - CDbl(t0)) * 1000# / CDbl(m_Freq)
End Function

Public Function RdvMsSince(ByVal t0 As Currency) As Double
    RdvMsSince = RdvMsBetween(t0, RdvTicks())
End Function

' smallest power of two at least twice the row count, capped at what the 21-bit
' hash can address. RdvSpec.Capacity() in C# computes the same number, and the
' probe counts the two methods report only match if it does.
Public Function RdvCapacity(ByVal rows As Long) As Long
    Dim cap As Long
    cap = 1024
    Do While cap < rows * 2 And cap < 2097152
        cap = cap * 2
    Loop
    RdvCapacity = cap
End Function

' The canonical hash. The three hot loops in modRdvEngine inline these same
' nine lines because a VBA procedure call costs more than the hash itself;
' build\build_workbooks.ps1 checks that every copy is still identical to this
' one, character for character.
Public Function RdvHashBytes(ByRef b() As Byte, ByVal p As Long) As Long
    Dim h As Long
    h = b(p)
    h = ((h * 131) + b(p + 1)) And RDV_HASHMASK
    h = ((h * 131) + b(p + 2)) And RDV_HASHMASK
    h = ((h * 131) + b(p + 3)) And RDV_HASHMASK
    h = ((h * 131) + b(p + 4)) And RDV_HASHMASK
    h = ((h * 131) + b(p + 5)) And RDV_HASHMASK
    h = ((h * 131) + b(p + 6)) And RDV_HASHMASK
    h = ((h * 131) + b(p + 7)) And RDV_HASHMASK
    h = (h Xor (h \ 1024)) And RDV_HASHMASK
    RdvHashBytes = h
End Function

#If VBA7 Then
Public Function RdvWindowTitle(ByVal h As LongPtr) As String
#Else
Public Function RdvWindowTitle(ByVal h As Long) As String
#End If
    Dim buf As String
    Dim n As Long
    buf = String$(320, vbNullChar)
    n = RdvGetWindowText(h, StrPtr(buf), 320)
    If n > 0 Then
        RdvWindowTitle = Left$(buf, n)
    Else
        RdvWindowTitle = ""
    End If
End Function

Public Function RdvIsKey(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) <> RDV_KEYLEN Then
        RdvIsKey = False
        Exit Function
    End If
    For i = 1 To RDV_KEYLEN
        c = AscW(Mid$(s, i, 1))
        If c < 48 Or c > 57 Then
            RdvIsKey = False
            Exit Function
        End If
    Next i
    RdvIsKey = True
End Function

' the last non-empty line of whatever the field holds
Public Function RdvCandidate(ByVal s As String) As String
    Dim e As Long, st As Long, ch As String
    If Len(s) = 0 Then
        RdvCandidate = ""
        Exit Function
    End If
    e = Len(s)
    Do While e > 0
        ch = Mid$(s, e, 1)
        If ch <> vbLf And ch <> vbCr And ch <> " " And ch <> vbTab Then Exit Do
        e = e - 1
    Loop
    If e = 0 Then
        RdvCandidate = ""
        Exit Function
    End If
    st = e
    Do While st > 1
        ch = Mid$(s, st - 1, 1)
        If ch = vbLf Or ch = vbCr Then Exit Do
        st = st - 1
    Loop
    RdvCandidate = Trim$(Mid$(s, st, e - st + 1))
End Function

Public Function RdvStageName(ByVal i As Long) As String
    Select Case i
        Case 0: RdvStageName = "1 表A 読込・索引"
        Case 1: RdvStageName = "2 表B 読込・索引"
        Case 2: RdvStageName = "3 表C 読込・索引"
        Case 3: RdvStageName = "4 A-B 結合 (番号1)"
        Case 4: RdvStageName = "5 B-C 結合 (番号2)"
        Case 5: RdvStageName = "6 番号1 検索"
        Case 6: RdvStageName = "7 画面表示"
        Case Else: RdvStageName = ""
    End Select
End Function

Public Function RdvStageKey(ByVal i As Long) As String
    Select Case i
        Case 0: RdvStageKey = "loadA"
        Case 1: RdvStageKey = "loadB"
        Case 2: RdvStageKey = "loadC"
        Case 3: RdvStageKey = "joinAB"
        Case 4: RdvStageKey = "joinBC"
        Case 5: RdvStageKey = "select"
        Case 6: RdvStageKey = "display"
        Case Else: RdvStageKey = ""
    End Select
End Function
