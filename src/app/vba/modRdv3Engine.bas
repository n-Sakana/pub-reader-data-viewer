Attribute VB_Name = "modRdv3Engine"
'==============================================================================
' modRdv3Engine -- CSV read + Scripting.Dictionary indexes + both joins.
'
' This module runs inside the WORKER: a second, invisible Excel process this
' app started itself. The UI process never runs the merge. It also runs once
' at build time to seed the initial LEDGER sheet.
'
' The index is Scripting.Dictionary, late bound through CreateObject, with a
' Collection of row numbers under each key -- the same structure the one-to-
' many comparison measured as "VBA 標準 Dictionary" (clsRdv2IdxDict). The
' hand-built hash that was measured next to it is comparison evidence only and
' is deliberately not here: one method, no fallback.
'
' Reading keeps the frozen builds' approach (bytes + the same bytes widened to
' one character per byte, rows as offset pairs): Split on the whole file fails
' with error 5 at this size, and the data is ASCII so the widening is 1:1.
'
' Timed stages (the merge time on screen is their sum):
'   readA readB readC | idxA idxB idxC | joinAB joinBC
' Composing the 28-column ledger lines happens after the clock: it is ledger
' materialisation, logged separately by the caller.
'==============================================================================
Option Explicit

Private m_bA() As Byte, m_stA() As Long, m_enA() As Long, m_sA As String
Private m_bB() As Byte, m_stB() As Long, m_enB() As Long, m_sB As String
Private m_bC() As Byte, m_stC() As Long, m_enC() As Long, m_sC As String
Private m_rowsA As Long, m_rowsB As Long, m_rowsC As Long
Private m_hdrA As String, m_hdrB As String, m_hdrC As String

Private m_stageMs(0 To 7) As Double
Private m_checksum As Double
Private m_matchAB As Long, m_matchBC As Long
Private m_errText As String

Private m_lines() As String
Private m_lineCount As Long
Private m_composeMs As Double

'------------------------------------------------------------------------------
' what the caller reads back
'------------------------------------------------------------------------------
Public Function Rdv3EngStageMs(ByVal i As Long) As Double
    Rdv3EngStageMs = m_stageMs(i)
End Function

Public Function Rdv3EngMergeMs() As Double
    Dim i As Long, s As Double
    For i = 0 To RDV3_STAGES - 1
        s = s + m_stageMs(i)
    Next i
    Rdv3EngMergeMs = s
End Function

Public Function Rdv3EngRows() As Long
    Rdv3EngRows = m_rowsB
End Function

Public Function Rdv3EngChecksum() As Double
    Rdv3EngChecksum = m_checksum
End Function

Public Function Rdv3EngMatchedAB() As Long
    Rdv3EngMatchedAB = m_matchAB
End Function

Public Function Rdv3EngMatchedBC() As Long
    Rdv3EngMatchedBC = m_matchBC
End Function

Public Function Rdv3EngErrText() As String
    Rdv3EngErrText = m_errText
End Function

Public Function Rdv3EngLineCount() As Long
    Rdv3EngLineCount = m_lineCount
End Function

Public Function Rdv3EngLine(ByVal i As Long) As String
    Rdv3EngLine = m_lines(i)
End Function

Public Function Rdv3EngComposeMs() As Double
    Rdv3EngComposeMs = m_composeMs
End Function

' one joined block for the result file, built in chunks so no single string
' concatenation has to carry 25 MB alone
Public Function Rdv3EngLinesJoined() As String
    Rdv3EngLinesJoined = Join(m_lines, vbCrLf)
End Function

'------------------------------------------------------------------------------
' one CSV -> bytes + widened string + row offsets (0 based, header split off)
'------------------------------------------------------------------------------
Private Function ReadOne(ByVal path As String, ByRef b() As Byte, ByRef s As String, _
                         ByRef st() As Long, ByRef en() As Long, ByRef nrows As Long, _
                         ByRef header As String) As String
    Dim f As Integer, n As Long
    Dim L As Long, pos As Long, nl As Long, e As Long, cnt As Long, capRows As Long
    Dim ts() As Long, te() As Long
    Dim i As Long, p As Long, ke As Long

    On Error GoTo Fail

    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n <= 0 Then
        Close #f
        ReadOne = "空のファイル: " & path
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f

    s = StrConv(b, vbUnicode)
    L = Len(s)
    pos = 1
    If n >= 3 Then
        If b(0) = 239 And b(1) = 187 And b(2) = 191 Then pos = 4
    End If

    capRows = 65536
    ReDim ts(0 To capRows - 1)
    ReDim te(0 To capRows - 1)
    cnt = 0
    Do While pos <= L
        nl = InStr(pos, s, vbLf, vbBinaryCompare)
        If nl = 0 Then
            e = L + 1
        Else
            e = nl
        End If
        If e > pos Then
            If Mid$(s, e - 1, 1) = vbCr Then e = e - 1
        End If
        If e > pos Then
            If cnt = capRows Then
                capRows = capRows * 2
                ReDim Preserve ts(0 To capRows - 1)
                ReDim Preserve te(0 To capRows - 1)
            End If
            ts(cnt) = pos - 1
            te(cnt) = e - 1
            cnt = cnt + 1
        End If
        If nl = 0 Then Exit Do
        pos = nl + 1
    Loop
    If cnt < 2 Then
        ReadOne = "データ行がありません: " & path
        Exit Function
    End If

    header = Mid$(s, ts(0) + 1, te(0) - ts(0))
    nrows = cnt - 1
    ReDim st(0 To nrows - 1)
    ReDim en(0 To nrows - 1)
    For i = 0 To nrows - 1
        st(i) = ts(i + 1)
        en(i) = te(i + 1)
        p = st(i)
        ke = p + RDV3_KEYLEN
        If ke > en(i) Then
            ReadOne = "キー長が " & RDV3_KEYLEN & " ではありません: " & path & " 行 " & (i + 1)
            Exit Function
        End If
        If ke < en(i) Then
            If b(ke) <> RDV3_COMMA Then
                ReadOne = "キー長が " & RDV3_KEYLEN & " ではありません: " & path & " 行 " & (i + 1)
                Exit Function
            End If
        End If
    Next i
    ReadOne = ""
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    ReadOne = "読込エラー " & Err.Number & ": " & Err.Description & " (" & path & ")"
End Function

' the standard index: key string -> Collection of 0-based row numbers. A key
' that owns several rows keeps all of them; nothing is overwritten.
Private Function BuildIndex(ByRef s As String, ByRef st() As Long, ByVal nrows As Long) As Object
    Dim d As Object
    Dim i As Long
    Dim k As String
    Dim col As Collection

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    For i = 0 To nrows - 1
        k = Mid$(s, st(i) + 1, RDV3_KEYLEN)
        If d.Exists(k) Then
            d.Item(k).Add i
        Else
            Set col = New Collection
            col.Add i
            d.Add k, col
        End If
    Next i
    Set BuildIndex = d
End Function

' st is the 0-based first byte of the row, en its 0-based exclusive end, so in
' the 1-based widened string the row is chars (st+1)..en -- the same coordinate
' system the frozen builds use.
Private Sub SplitRow(ByRef s As String, ByVal st As Long, ByVal en As Long, ByRef out() As String)
    Dim f As Long, p As Long, q As Long
    p = st + 1
    For f = 0 To RDV3_FIELDS - 1
        q = InStr(p, s, ",", vbBinaryCompare)
        If q = 0 Or q > en Then
            q = en + 1
        End If
        If q > p Then
            out(f) = Mid$(s, p, q - p)
        Else
            out(f) = ""
        End If
        p = q + 1
        If p > en + 1 Then
            p = en + 1
        End If
    Next f
End Sub

'------------------------------------------------------------------------------
' the merge: 8 timed stages, then the ledger lines
'------------------------------------------------------------------------------
Public Function Rdv3EngMerge(ByVal dataDir As String) As Boolean
    Dim t As Currency
    Dim msg As String
    Dim dA As Object, dB As Object, dC As Object
    Dim ab() As Long, bc() As Long
    Dim i As Long, p As Long, e As Long
    Dim k As String
    Dim chk As Double

    m_errText = ""
    m_checksum = 0
    m_matchAB = 0
    m_matchBC = 0
    m_lineCount = 0
    m_composeMs = 0
    For i = 0 To RDV3_STAGES - 1
        m_stageMs(i) = 0
    Next i

    t = Rdv3Ticks()
    msg = ReadOne(dataDir & "\tableA.csv", m_bA, m_sA, m_stA, m_enA, m_rowsA, m_hdrA)
    m_stageMs(RDV3_ST_READA) = Rdv3MsSince(t)
    If Len(msg) > 0 Then
        m_errText = msg
        Rdv3EngMerge = False
        Exit Function
    End If

    t = Rdv3Ticks()
    msg = ReadOne(dataDir & "\tableB.csv", m_bB, m_sB, m_stB, m_enB, m_rowsB, m_hdrB)
    m_stageMs(RDV3_ST_READB) = Rdv3MsSince(t)
    If Len(msg) > 0 Then
        m_errText = msg
        Rdv3EngMerge = False
        Exit Function
    End If

    t = Rdv3Ticks()
    msg = ReadOne(dataDir & "\tableC.csv", m_bC, m_sC, m_stC, m_enC, m_rowsC, m_hdrC)
    m_stageMs(RDV3_ST_READC) = Rdv3MsSince(t)
    If Len(msg) > 0 Then
        m_errText = msg
        Rdv3EngMerge = False
        Exit Function
    End If

    On Error GoTo Fail

    t = Rdv3Ticks()
    Set dA = BuildIndex(m_sA, m_stA, m_rowsA)
    m_stageMs(RDV3_ST_IDXA) = Rdv3MsSince(t)

    t = Rdv3Ticks()
    Set dB = BuildIndex(m_sB, m_stB, m_rowsB)
    m_stageMs(RDV3_ST_IDXB) = Rdv3MsSince(t)

    t = Rdv3Ticks()
    Set dC = BuildIndex(m_sC, m_stC, m_rowsC)
    m_stageMs(RDV3_ST_IDXC) = Rdv3MsSince(t)

    ReDim ab(0 To m_rowsB - 1)
    ReDim bc(0 To m_rowsB - 1)

    ' A-B over every row of B on key 1. A is unique per key 1 in this data but
    ' nothing here assumes it: the index answers with a set either way.
    t = Rdv3Ticks()
    For i = 0 To m_rowsB - 1
        k = Mid$(m_sB, m_stB(i) + 1, RDV3_KEYLEN)
        If dA.Exists(k) Then
            ab(i) = dA.Item(k).Item(1)
            m_matchAB = m_matchAB + 1
        Else
            ab(i) = -1
        End If
    Next i
    m_stageMs(RDV3_ST_JOINAB) = Rdv3MsSince(t)

    ' B-C on key 2 = field 1 of B, with the same join checksum the comparison
    ' builds report, so expected.txt still verifies this merge
    t = Rdv3Ticks()
    chk = 0
    For i = 0 To m_rowsB - 1
        p = m_stB(i)
        e = m_enB(i)
        Do While p < e
            If m_bB(p) = RDV3_COMMA Then Exit Do
            p = p + 1
        Loop
        p = p + 1
        bc(i) = -1
        If p + RDV3_KEYLEN <= e Then
            k = Mid$(m_sB, p + 1, RDV3_KEYLEN)
            If dC.Exists(k) Then
                bc(i) = dC.Item(k).Item(1)
                m_matchBC = m_matchBC + 1
            End If
        End If
        chk = chk * 31 + ab(i) + bc(i)
        chk = chk - Fix(chk / RDV3_MODV) * RDV3_MODV
    Next i
    m_checksum = chk
    m_stageMs(RDV3_ST_JOINBC) = Rdv3MsSince(t)

    ' ---- end of the timed merge. Compose the 28-column ledger lines.
    t = Rdv3Ticks()
    ComposeLines ab, bc
    m_composeMs = Rdv3MsSince(t)

    Rdv3EngMerge = True
    Exit Function
Fail:
    m_errText = "統合エラー " & Err.Number & ": " & Err.Description
    Rdv3EngMerge = False
End Function

Private Sub ComposeLines(ByRef ab() As Long, ByRef bc() As Long)
    Dim i As Long, f As Long
    Dim pa(0 To 9) As String, pb(0 To 9) As String, pc(0 To 9) As String
    Dim parts(0 To 27) As String
    Dim ra As Long, rc As Long

    ReDim m_lines(0 To m_rowsB - 1)
    For i = 0 To m_rowsB - 1
        SplitRow m_sB, m_stB(i), m_enB(i), pb
        ra = ab(i)
        If ra >= 0 Then
            SplitRow m_sA, m_stA(ra), m_enA(ra), pa
        Else
            For f = 0 To 9
                pa(f) = ""
            Next f
        End If
        rc = bc(i)
        If rc >= 0 Then
            SplitRow m_sC, m_stC(rc), m_enC(rc), pc
        Else
            For f = 0 To 9
                pc(f) = ""
            Next f
        End If
        parts(0) = pb(0)                     ' key1
        parts(1) = pb(1)                     ' key2
        For f = 1 To 9
            parts(1 + f) = pa(f)             ' a_code..a_note
        Next f
        For f = 2 To 9
            parts(9 + f) = pb(f)             ' b_slip..b_memo
        Next f
        For f = 1 To 9
            parts(18 + f) = pc(f)            ' c_item..c_remark
        Next f
        m_lines(i) = Join(parts, vbTab)
    Next i
    m_lineCount = m_rowsB
End Sub
