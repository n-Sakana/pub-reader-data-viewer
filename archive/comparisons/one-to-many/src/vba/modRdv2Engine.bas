Attribute VB_Name = "modRdv2Engine"
'==============================================================================
' modRdv2Engine -- merge-select over a one-to-many key, in Excel's own process.
'
' Same ten stages as src\v2\csharp\Rdv2Core.cs, in the same order: read three
' CSVs, build three indexes, join A-B on key 1, join B-C on key 2, search key 1
' for the set of candidates, present. Reading and indexing are timed apart
' because what the index costs is the question this build exists to answer.
'
' The index is not in this file. modRdv2Engine says New clsRdv2Idx and the
' workbook it was built into decides whether that is the hand-built hash or
' Scripting.Dictionary. Same call sites for both.
'
' Two things VBA forces, both measured rather than assumed:
'
'   Split(text, vbLf) on a large string fails outright with error 5, so lines
'   are never materialised. A row is a pair of byte offsets.
'
'   The file is kept in two forms: the bytes, and the same bytes widened one
'   byte to one character. The widened copy is what makes Mid$ able to cut a
'   key or a field out, and both index flavours are handed both forms, so
'   neither is charged for something the other got for free. The data is ASCII,
'   which is what makes the widening one to one.
'==============================================================================
Option Explicit

Private m_bA() As Byte, m_stA() As Long, m_enA() As Long, m_sA As String
Private m_bB() As Byte, m_stB() As Long, m_enB() As Long, m_sB As String
Private m_bC() As Byte, m_stC() As Long, m_enC() As Long, m_sC As String
Private m_rowsA As Long, m_rowsB As Long, m_rowsC As Long
Private m_hdrA As String, m_hdrB As String, m_hdrC As String
Private m_idxA As clsRdv2Idx, m_idxB As clsRdv2Idx, m_idxC As clsRdv2Idx
Private m_abIdx() As Long, m_bcIdx() As Long

Private m_stageMs(0 To 9) As Double
Private m_checksum As Double
Private m_matchAB As Long, m_matchBC As Long
Private m_errText As String

Private m_candN As Long
Private m_candB() As Long, m_candA() As Long, m_candC() As Long
Private m_candCol() As String
Private m_chosen As Long

Private m_nameA(0 To 9) As String, m_nameB(0 To 9) As String, m_nameC(0 To 9) As String
Private m_valA(0 To 9) As String, m_valB(0 To 9) As String, m_valC(0 To 9) As String

'------------------------------------------------------------------------------
' what the screen and the log read back
'------------------------------------------------------------------------------
Public Function Rdv2StageMs(ByVal i As Long) As Double
    Rdv2StageMs = m_stageMs(i)
End Function

Public Sub Rdv2SetStageMs(ByVal i As Long, ByVal v As Double)
    m_stageMs(i) = v
End Sub

Public Function Rdv2StageSum() As Double
    Dim i As Long, s As Double
    For i = 0 To RDV2_STAGES - 1
        s = s + m_stageMs(i)
    Next i
    Rdv2StageSum = s
End Function

Public Function Rdv2Checksum() As Double
    Rdv2Checksum = m_checksum
End Function

Public Function Rdv2Probes() As Double
    Dim p As Double
    If Not m_idxA Is Nothing Then
        p = p + m_idxA.Probes
    End If
    If Not m_idxB Is Nothing Then
        p = p + m_idxB.Probes
    End If
    If Not m_idxC Is Nothing Then
        p = p + m_idxC.Probes
    End If
    Rdv2Probes = p
End Function

Public Function Rdv2IndexKind() As String
    Dim o As clsRdv2Idx
    If m_idxB Is Nothing Then
        Set o = New clsRdv2Idx
        Rdv2IndexKind = o.Kind
    Else
        Rdv2IndexKind = m_idxB.Kind
    End If
End Function

Public Function Rdv2Rows() As Long
    Rdv2Rows = m_rowsB
End Function

Public Function Rdv2MatchedAB() As Long
    Rdv2MatchedAB = m_matchAB
End Function

Public Function Rdv2MatchedBC() As Long
    Rdv2MatchedBC = m_matchBC
End Function

Public Function Rdv2ErrText() As String
    Rdv2ErrText = m_errText
End Function

Public Function Rdv2CandCount() As Long
    Rdv2CandCount = m_candN
End Function

Public Function Rdv2Chosen() As Long
    Rdv2Chosen = m_chosen
End Function

' the eight identification columns of candidate i, 0 based:
' key2, line, slip, date, qty, status, item, maker
Public Function Rdv2CandCell(ByVal i As Long, ByVal c As Long) As String
    If i < 0 Or i >= m_candN Then
        Rdv2CandCell = ""
        Exit Function
    End If
    Rdv2CandCell = m_candCol(i, c)
End Function

Public Function Rdv2NameOf(ByVal tbl As Long, ByVal i As Long) As String
    Select Case tbl
        Case 0: Rdv2NameOf = m_nameA(i)
        Case 1: Rdv2NameOf = m_nameB(i)
        Case Else: Rdv2NameOf = m_nameC(i)
    End Select
End Function

Public Function Rdv2ValueOf(ByVal tbl As Long, ByVal i As Long) As String
    Select Case tbl
        Case 0: Rdv2ValueOf = m_valA(i)
        Case 1: Rdv2ValueOf = m_valB(i)
        Case Else: Rdv2ValueOf = m_valC(i)
    End Select
End Function

'------------------------------------------------------------------------------
' one field out of one row, from the widened copy
'------------------------------------------------------------------------------
Private Function FieldStr(ByRef s As String, ByVal st As Long, ByVal en As Long, ByVal f As Long) As String
    Dim p As Long, q As Long, k As Long
    p = st + 1
    For k = 1 To f
        q = InStr(p, s, ",", vbBinaryCompare)
        If q = 0 Or q > en Then
            FieldStr = ""
            Exit Function
        End If
        p = q + 1
    Next k
    q = InStr(p, s, ",", vbBinaryCompare)
    If q = 0 Or q > en Then
        q = en + 1
    End If
    FieldStr = Mid$(s, p, q - p)
End Function

Private Sub SplitRow(ByRef s As String, ByVal st As Long, ByVal en As Long, ByRef out() As String)
    Dim f As Long, p As Long, q As Long
    p = st + 1
    For f = 0 To RDV2_FIELDS - 1
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

Private Sub SplitHead(ByVal header As String, ByRef out() As String)
    Dim parts() As String
    Dim i As Long
    parts = Split(header, ",")
    For i = 0 To RDV2_FIELDS - 1
        If i <= UBound(parts) Then
            out(i) = parts(i)
        Else
            out(i) = ""
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' stages 1..3: read one CSV. Indexing is stage 4..6 and is timed on its own.
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
            ts(cnt) = pos - 1              ' 0 based byte offsets, as in C#
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
        ke = p + RDV2_KEYLEN
        If ke > en(i) Then
            ReadOne = "キー長が " & RDV2_KEYLEN & " ではありません: " & path & " 行 " & (i + 1)
            Exit Function
        End If
        If ke < en(i) Then
            If b(ke) <> RDV2_COMMA Then
                ReadOne = "キー長が " & RDV2_KEYLEN & " ではありません: " & path & " 行 " & (i + 1)
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

'------------------------------------------------------------------------------
' stages 1..9. The caller stamps stage 10 once the screen or the dialog is up.
'------------------------------------------------------------------------------
Public Function Rdv2RunMergeSelect(ByVal dataDir As String, ByVal key As String) As Boolean
    Dim t As Currency
    Dim msg As String
    Dim i As Long, j As Long, n As Long, p As Long, e As Long, c As Long
    Dim chk As Double
    Dim found() As Long
    Dim kb() As Byte
    Dim ks As String
    Dim tmp As Long

    m_errText = ""
    m_candN = 0
    m_chosen = -1
    m_checksum = 0
    m_matchAB = 0
    m_matchBC = 0
    For i = 0 To RDV2_FIELDS - 1
        m_nameA(i) = "": m_valA(i) = ""
        m_nameB(i) = "": m_valB(i) = ""
        m_nameC(i) = "": m_valC(i) = ""
    Next i

    t = Rdv2Ticks()
    msg = ReadOne(dataDir & "\tableA.csv", m_bA, m_sA, m_stA, m_enA, m_rowsA, m_hdrA)
    m_stageMs(RDV2_ST_READA) = Rdv2MsSince(t)
    If Len(msg) > 0 Then
        m_errText = msg
        Rdv2RunMergeSelect = False
        Exit Function
    End If

    t = Rdv2Ticks()
    msg = ReadOne(dataDir & "\tableB.csv", m_bB, m_sB, m_stB, m_enB, m_rowsB, m_hdrB)
    m_stageMs(RDV2_ST_READB) = Rdv2MsSince(t)
    If Len(msg) > 0 Then
        m_errText = msg
        Rdv2RunMergeSelect = False
        Exit Function
    End If

    t = Rdv2Ticks()
    msg = ReadOne(dataDir & "\tableC.csv", m_bC, m_sC, m_stC, m_enC, m_rowsC, m_hdrC)
    m_stageMs(RDV2_ST_READC) = Rdv2MsSince(t)
    If Len(msg) > 0 Then
        m_errText = msg
        Rdv2RunMergeSelect = False
        Exit Function
    End If

    t = Rdv2Ticks()
    Set m_idxA = New clsRdv2Idx
    m_idxA.Build m_bA, m_sA, m_stA, m_rowsA
    m_stageMs(RDV2_ST_IDXA) = Rdv2MsSince(t)

    t = Rdv2Ticks()
    Set m_idxB = New clsRdv2Idx
    m_idxB.Build m_bB, m_sB, m_stB, m_rowsB
    m_stageMs(RDV2_ST_IDXB) = Rdv2MsSince(t)

    t = Rdv2Ticks()
    Set m_idxC = New clsRdv2Idx
    m_idxC.Build m_bC, m_sC, m_stC, m_rowsC
    m_stageMs(RDV2_ST_IDXC) = Rdv2MsSince(t)

    ReDim found(0 To RDV2_MAXCAND - 1)
    ReDim m_abIdx(0 To m_rowsB - 1)
    ReDim m_bcIdx(0 To m_rowsB - 1)

    ' stage 7: A-B over every row of B. Table A holds one row per key 1, but
    ' nothing here assumes it -- the index answers with a set either way.
    t = Rdv2Ticks()
    For i = 0 To m_rowsB - 1
        n = m_idxA.Find(m_bB, m_sB, m_stB(i), found)
        If n > 0 Then
            m_abIdx(i) = found(0)
            m_matchAB = m_matchAB + 1
        Else
            m_abIdx(i) = -1
        End If
    Next i
    m_stageMs(RDV2_ST_JOINAB) = Rdv2MsSince(t)

    ' stage 8: B-C on key 2, which is field 1 of B
    t = Rdv2Ticks()
    chk = 0
    For i = 0 To m_rowsB - 1
        p = m_stB(i)
        e = m_enB(i)
        Do While p < e
            If m_bB(p) = RDV2_COMMA Then Exit Do
            p = p + 1
        Loop
        p = p + 1
        m_bcIdx(i) = -1
        If p + RDV2_KEYLEN <= e Then
            n = m_idxC.Find(m_bB, m_sB, p, found)
            If n > 0 Then
                m_bcIdx(i) = found(0)
                m_matchBC = m_matchBC + 1
            End If
        End If
        chk = chk * 31 + m_abIdx(i) + m_bcIdx(i)
        chk = chk - Fix(chk / RDV2_MODV) * RDV2_MODV
    Next i
    m_checksum = chk
    m_stageMs(RDV2_ST_JOINBC) = Rdv2MsSince(t)

    ' stage 9: key 1 -> the set of B rows. This is where one-to-many shows up:
    ' the answer is a set, and every member of it is kept.
    t = Rdv2Ticks()
    n = 0
    If Rdv2IsKey(key) Then
        kb = StrConv(key, vbFromUnicode)
        ks = key
        n = m_idxB.Find(kb, ks, 0, found)
    End If
    m_candN = n
    If n > 0 Then
        ReDim m_candB(0 To n - 1)
        ReDim m_candA(0 To n - 1)
        ReDim m_candC(0 To n - 1)
        For i = 0 To n - 1
            m_candB(i) = found(i)
        Next i
        ' file order, so the list the operator sees is stable
        For i = 1 To n - 1
            tmp = m_candB(i)
            j = i - 1
            Do While j >= 0
                If m_candB(j) <= tmp Then Exit Do
                m_candB(j + 1) = m_candB(j)
                j = j - 1
            Loop
            m_candB(j + 1) = tmp
        Next i
        For i = 0 To n - 1
            c = m_candB(i)
            m_candA(i) = m_abIdx(c)
            m_candC(i) = m_bcIdx(c)
        Next i
    End If
    m_stageMs(RDV2_ST_SEARCH) = Rdv2MsSince(t)

    Rdv2RunMergeSelect = True
End Function

'------------------------------------------------------------------------------
' part of stage 10: the columns the operator picks by. Only the candidates are
' materialised, never all 100,000 rows.
'------------------------------------------------------------------------------
Public Sub Rdv2FillCandidates()
    Dim i As Long, rb As Long, rc As Long
    If m_candN <= 0 Then Exit Sub
    ReDim m_candCol(0 To m_candN - 1, 0 To 7)
    For i = 0 To m_candN - 1
        rb = m_candB(i)
        m_candCol(i, 0) = FieldStr(m_sB, m_stB(rb), m_enB(rb), 1)     ' key2
        m_candCol(i, 1) = FieldStr(m_sB, m_stB(rb), m_enB(rb), 8)     ' b_line
        m_candCol(i, 2) = FieldStr(m_sB, m_stB(rb), m_enB(rb), 2)     ' b_slip
        m_candCol(i, 3) = FieldStr(m_sB, m_stB(rb), m_enB(rb), 3)     ' b_date
        m_candCol(i, 4) = FieldStr(m_sB, m_stB(rb), m_enB(rb), 4)     ' b_qty
        m_candCol(i, 5) = FieldStr(m_sB, m_stB(rb), m_enB(rb), 7)     ' b_status
        rc = m_candC(i)
        If rc >= 0 Then
            m_candCol(i, 6) = FieldStr(m_sC, m_stC(rc), m_enC(rc), 1) ' c_item
            m_candCol(i, 7) = FieldStr(m_sC, m_stC(rc), m_enC(rc), 2) ' c_maker
        Else
            m_candCol(i, 6) = ""
            m_candCol(i, 7) = ""
        End If
    Next i
End Sub

' the 30 fields of one chosen candidate
Public Sub Rdv2FillRecord(ByVal which As Long)
    Dim ra As Long, rb As Long, rc As Long
    Dim i As Long
    If which < 0 Or which >= m_candN Then Exit Sub
    m_chosen = which
    rb = m_candB(which)
    ra = m_candA(which)
    rc = m_candC(which)

    SplitHead m_hdrA, m_nameA
    SplitHead m_hdrB, m_nameB
    SplitHead m_hdrC, m_nameC
    SplitRow m_sB, m_stB(rb), m_enB(rb), m_valB
    If ra >= 0 Then
        SplitRow m_sA, m_stA(ra), m_enA(ra), m_valA
    Else
        For i = 0 To RDV2_FIELDS - 1
            m_valA(i) = ""
        Next i
    End If
    If rc >= 0 Then
        SplitRow m_sC, m_stC(rc), m_enC(rc), m_valC
    Else
        For i = 0 To RDV2_FIELDS - 1
            m_valC(i) = ""
        Next i
    End If
End Sub
