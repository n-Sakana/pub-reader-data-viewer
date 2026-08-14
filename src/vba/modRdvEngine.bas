Attribute VB_Name = "modRdvEngine"
'==============================================================================
' modRdvEngine -- the merge-select engine, in Excel's own process.
'
' Same algorithm as src\csharp\RdvCore.cs, stage for stage:
'   read the file into bytes, scan the line offsets, index the key with an
'   open-addressing table, join A-B on key 1, join B-C on key 2, look one row
'   up, hand the fields to the screen.
'
' Two things VBA forces that C# does not, both measured rather than assumed:
'
'   Scripting.Dictionary is not usable at this size. Measured on this data:
'   1,000,000 inserts 28.2 s and 1,000,000 lookups 29.6 s, about 30 us per
'   operation. The table below does the same work in 0.19 s and 0.44 s. So the
'   hash table is hand-built -- not for cleverness, because there is no stock
'   one that finishes.
'
'   Split(text, vbLf) on an 80 MB string fails outright with error 5, so lines
'   are never materialised. A row is a pair of byte offsets and a key is eight
'   bytes at a known offset, exactly as in the C# build.
'
' The string form of the file exists only while the line offsets are being
' scanned, because InStr on a String is the only fast way VBA has to find a
' line break. It is released immediately afterwards; every later stage reads
' the byte array.
'==============================================================================
Option Explicit

Private m_bA() As Byte, m_stA() As Long, m_enA() As Long, m_slA() As Long
Private m_bB() As Byte, m_stB() As Long, m_enB() As Long, m_slB() As Long
Private m_bC() As Byte, m_stC() As Long, m_enC() As Long, m_slC() As Long
Private m_rowsA As Long, m_rowsB As Long, m_rowsC As Long
Private m_maskA As Long, m_maskB As Long, m_maskC As Long
Private m_hdrA As String, m_hdrB As String, m_hdrC As String
Private m_probA As Double, m_probB As Double, m_probC As Double
Private m_abIdx() As Long, m_bcIdx() As Long

Private m_stageMs(0 To 6) As Double
Private m_checksum As Double
Private m_matchAB As Long, m_matchBC As Long
Private m_found As Boolean
Private m_key1 As String, m_key2 As String
Private m_nameA(0 To 9) As String, m_nameB(0 To 9) As String, m_nameC(0 To 9) As String
Private m_valA(0 To 9) As String, m_valB(0 To 9) As String, m_valC(0 To 9) As String
Private m_error As String

Public Function RdvStageMs(ByVal i As Long) As Double
    RdvStageMs = m_stageMs(i)
End Function

Public Sub RdvSetStageMs(ByVal i As Long, ByVal v As Double)
    m_stageMs(i) = v
End Sub

Public Function RdvStageSum() As Double
    Dim i As Long, s As Double
    For i = 0 To RDV_STAGES - 1
        s = s + m_stageMs(i)
    Next i
    RdvStageSum = s
End Function

Public Function RdvChecksum() As Double
    RdvChecksum = m_checksum
End Function

Public Function RdvProbes() As Double
    RdvProbes = m_probA + m_probB + m_probC
End Function

Public Function RdvRows() As Long
    RdvRows = m_rowsB
End Function

Public Function RdvMatchedAB() As Long
    RdvMatchedAB = m_matchAB
End Function

Public Function RdvMatchedBC() As Long
    RdvMatchedBC = m_matchBC
End Function

Public Function RdvFound() As Boolean
    RdvFound = m_found
End Function

Public Function RdvKey2() As String
    RdvKey2 = m_key2
End Function

Public Function RdvError() As String
    RdvError = m_error
End Function

Public Function RdvNameOf(ByVal tbl As Long, ByVal i As Long) As String
    Select Case tbl
        Case 0: RdvNameOf = m_nameA(i)
        Case 1: RdvNameOf = m_nameB(i)
        Case Else: RdvNameOf = m_nameC(i)
    End Select
End Function

Public Function RdvValueOf(ByVal tbl As Long, ByVal i As Long) As String
    Select Case tbl
        Case 0: RdvValueOf = m_valA(i)
        Case 1: RdvValueOf = m_valB(i)
        Case Else: RdvValueOf = m_valC(i)
    End Select
End Function

'------------------------------------------------------------------------------
' stage 1..3: read one CSV and index it
'------------------------------------------------------------------------------
Private Function LoadOne(ByVal path As String, ByRef b() As Byte, ByRef st() As Long, _
                         ByRef en() As Long, ByRef sl() As Long, ByRef nrows As Long, _
                         ByRef mask As Long, ByRef probes As Double, ByRef header As String) As String
    Dim f As Integer, n As Long
    Dim s As String
    Dim L As Long, pos As Long, nl As Long, e As Long, cnt As Long, capRows As Long
    Dim cap As Long, i As Long, h As Long, idx As Long, p As Long, ke As Long

    On Error GoTo Fail

    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n <= 0 Then
        Close #f
        LoadOne = "空のファイル: " & path
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f

    ' the widened copy exists only for the line scan below
    s = StrConv(b, vbUnicode)
    L = Len(s)
    pos = 1
    If n >= 3 Then
        If b(0) = 239 And b(1) = 187 And b(2) = 191 Then pos = 4
    End If

    capRows = 65536
    ReDim st(0 To capRows - 1)
    ReDim en(0 To capRows - 1)
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
                ReDim Preserve st(0 To capRows - 1)
                ReDim Preserve en(0 To capRows - 1)
            End If
            st(cnt) = pos - 1          ' 0-based byte offsets, as in the C# build
            en(cnt) = e - 1
            cnt = cnt + 1
        End If
        If nl = 0 Then Exit Do
        pos = nl + 1
    Loop
    If cnt < 2 Then
        LoadOne = "データ行がありません: " & path
        Exit Function
    End If
    header = Mid$(s, st(0) + 1, en(0) - st(0))
    s = ""                              ' 158 MB released before the joins start
    nrows = cnt - 1

    cap = RdvCapacity(nrows)
    mask = cap - 1
    ReDim sl(0 To cap - 1)
    probes = 0
    For i = 0 To nrows - 1
        p = st(i + 1)
        ke = p + RDV_KEYLEN
        If ke > en(i + 1) Then
            LoadOne = "キー長が " & RDV_KEYLEN & " ではありません: " & path & " 行 " & (i + 1)
            Exit Function
        End If
        If ke < en(i + 1) Then
            If b(ke) <> RDV_COMMA Then
                LoadOne = "キー長が " & RDV_KEYLEN & " ではありません: " & path & " 行 " & (i + 1)
                Exit Function
            End If
        End If
        ' --- RdvHashBytes, inlined ---
        h = b(p)
        h = ((h * 131) + b(p + 1)) And RDV_HASHMASK
        h = ((h * 131) + b(p + 2)) And RDV_HASHMASK
        h = ((h * 131) + b(p + 3)) And RDV_HASHMASK
        h = ((h * 131) + b(p + 4)) And RDV_HASHMASK
        h = ((h * 131) + b(p + 5)) And RDV_HASHMASK
        h = ((h * 131) + b(p + 6)) And RDV_HASHMASK
        h = ((h * 131) + b(p + 7)) And RDV_HASHMASK
        h = (h Xor (h \ 1024)) And RDV_HASHMASK
        ' --- end inline ---
        idx = h And mask
        Do While sl(idx) <> 0
            probes = probes + 1
            idx = (idx + 1) And mask
        Loop
        sl(idx) = i + 1
    Next i
    LoadOne = ""
    Exit Function
Fail:
    On Error Resume Next
    Close #f
    LoadOne = "読込エラー " & Err.Number & ": " & Err.Description & " (" & path & ")"
End Function

'------------------------------------------------------------------------------
' stage 4: A-B on key 1, one to one
'------------------------------------------------------------------------------
Private Sub JoinAB(ByRef pb() As Byte, ByRef pst() As Long, ByVal prows As Long, _
                   ByRef tb() As Byte, ByRef tst() As Long, ByRef tsl() As Long, _
                   ByVal tmask As Long, ByRef out() As Long, ByRef matched As Long)
    Dim i As Long, h As Long, idx As Long, p As Long, q As Long, r As Long
    matched = 0
    For i = 0 To prows - 1
        p = pst(i + 1)
        ' --- RdvHashBytes, inlined ---
        h = pb(p)
        h = ((h * 131) + pb(p + 1)) And RDV_HASHMASK
        h = ((h * 131) + pb(p + 2)) And RDV_HASHMASK
        h = ((h * 131) + pb(p + 3)) And RDV_HASHMASK
        h = ((h * 131) + pb(p + 4)) And RDV_HASHMASK
        h = ((h * 131) + pb(p + 5)) And RDV_HASHMASK
        h = ((h * 131) + pb(p + 6)) And RDV_HASHMASK
        h = ((h * 131) + pb(p + 7)) And RDV_HASHMASK
        h = (h Xor (h \ 1024)) And RDV_HASHMASK
        ' --- end inline ---
        idx = h And tmask
        out(i) = -1
        Do
            r = tsl(idx)
            If r = 0 Then Exit Do
            q = tst(r)
            If tb(q) = pb(p) And tb(q + 1) = pb(p + 1) And tb(q + 2) = pb(p + 2) _
               And tb(q + 3) = pb(p + 3) And tb(q + 4) = pb(p + 4) And tb(q + 5) = pb(p + 5) _
               And tb(q + 6) = pb(p + 6) And tb(q + 7) = pb(p + 7) Then
                out(i) = r - 1
                matched = matched + 1
                Exit Do
            End If
            idx = (idx + 1) And tmask
        Loop
    Next i
End Sub

'------------------------------------------------------------------------------
' stage 5: B-C on key 2 (field 1 of B), one to one, and fold the checksum
'------------------------------------------------------------------------------
Private Sub JoinBC(ByRef pb() As Byte, ByRef pst() As Long, ByRef pen() As Long, ByVal prows As Long, _
                   ByRef tb() As Byte, ByRef tst() As Long, ByRef tsl() As Long, _
                   ByVal tmask As Long, ByRef ab() As Long, ByRef out() As Long, _
                   ByRef matched As Long, ByRef checksum As Double)
    Dim i As Long, h As Long, idx As Long, p As Long, q As Long, r As Long, e As Long
    Dim chk As Double
    matched = 0
    chk = 0
    For i = 0 To prows - 1
        p = pst(i + 1)
        e = pen(i + 1)
        Do While p < e
            If pb(p) = RDV_COMMA Then Exit Do
            p = p + 1
        Loop
        p = p + 1
        r = 0
        out(i) = -1
        If p + RDV_KEYLEN <= e Then
            ' --- RdvHashBytes, inlined ---
            h = pb(p)
            h = ((h * 131) + pb(p + 1)) And RDV_HASHMASK
            h = ((h * 131) + pb(p + 2)) And RDV_HASHMASK
            h = ((h * 131) + pb(p + 3)) And RDV_HASHMASK
            h = ((h * 131) + pb(p + 4)) And RDV_HASHMASK
            h = ((h * 131) + pb(p + 5)) And RDV_HASHMASK
            h = ((h * 131) + pb(p + 6)) And RDV_HASHMASK
            h = ((h * 131) + pb(p + 7)) And RDV_HASHMASK
            h = (h Xor (h \ 1024)) And RDV_HASHMASK
            ' --- end inline ---
            idx = h And tmask
            Do
                r = tsl(idx)
                If r = 0 Then Exit Do
                q = tst(r)
                If tb(q) = pb(p) And tb(q + 1) = pb(p + 1) And tb(q + 2) = pb(p + 2) _
                   And tb(q + 3) = pb(p + 3) And tb(q + 4) = pb(p + 4) And tb(q + 5) = pb(p + 5) _
                   And tb(q + 6) = pb(p + 6) And tb(q + 7) = pb(p + 7) Then
                    out(i) = r - 1
                    matched = matched + 1
                    Exit Do
                End If
                idx = (idx + 1) And tmask
            Loop
        End If
        chk = chk * 31 + ab(i) + out(i)
        chk = chk - Fix(chk / RDV_MODV) * RDV_MODV
    Next i
    checksum = chk
End Sub

'------------------------------------------------------------------------------
' stages 1..6. The caller stamps stage 7 after the screen is up.
'------------------------------------------------------------------------------
Public Function RdvRunMergeSelect(ByVal dataDir As String, ByVal key As String) As Boolean
    Dim t As Currency
    Dim msg As String
    Dim i As Long

    m_error = ""
    m_found = False
    m_key1 = key
    m_key2 = ""
    For i = 0 To RDV_FIELDS - 1
        m_nameA(i) = "": m_valA(i) = ""
        m_nameB(i) = "": m_valB(i) = ""
        m_nameC(i) = "": m_valC(i) = ""
    Next i

    t = RdvTicks()
    msg = LoadOne(dataDir & "\tableA.csv", m_bA, m_stA, m_enA, m_slA, m_rowsA, m_maskA, m_probA, m_hdrA)
    m_stageMs(RDV_ST_LOADA) = RdvMsSince(t)
    If Len(msg) > 0 Then
        m_error = msg
        RdvRunMergeSelect = False
        Exit Function
    End If

    t = RdvTicks()
    msg = LoadOne(dataDir & "\tableB.csv", m_bB, m_stB, m_enB, m_slB, m_rowsB, m_maskB, m_probB, m_hdrB)
    m_stageMs(RDV_ST_LOADB) = RdvMsSince(t)
    If Len(msg) > 0 Then
        m_error = msg
        RdvRunMergeSelect = False
        Exit Function
    End If

    t = RdvTicks()
    msg = LoadOne(dataDir & "\tableC.csv", m_bC, m_stC, m_enC, m_slC, m_rowsC, m_maskC, m_probC, m_hdrC)
    m_stageMs(RDV_ST_LOADC) = RdvMsSince(t)
    If Len(msg) > 0 Then
        m_error = msg
        RdvRunMergeSelect = False
        Exit Function
    End If

    ReDim m_abIdx(0 To m_rowsB - 1)
    ReDim m_bcIdx(0 To m_rowsB - 1)

    t = RdvTicks()
    JoinAB m_bB, m_stB, m_rowsB, m_bA, m_stA, m_slA, m_maskA, m_abIdx, m_matchAB
    m_stageMs(RDV_ST_JOINAB) = RdvMsSince(t)

    t = RdvTicks()
    JoinBC m_bB, m_stB, m_enB, m_rowsB, m_bC, m_stC, m_slC, m_maskC, m_abIdx, m_bcIdx, m_matchBC, m_checksum
    m_stageMs(RDV_ST_JOINBC) = RdvMsSince(t)

    t = RdvTicks()
    SelectOne key
    m_stageMs(RDV_ST_SELECT) = RdvMsSince(t)

    RdvRunMergeSelect = True
End Function

'------------------------------------------------------------------------------
' stage 6: find key 1 in the joined table and materialise the 30 fields
'------------------------------------------------------------------------------
Private Sub SelectOne(ByVal key As String)
    Dim kb() As Byte
    Dim i As Long, h As Long, idx As Long, r As Long, q As Long, bi As Long
    Dim ai As Long, ci As Long
    Dim parts() As String

    m_found = False
    If Not RdvIsKey(key) Then Exit Sub

    ReDim kb(0 To RDV_KEYLEN - 1)
    For i = 0 To RDV_KEYLEN - 1
        kb(i) = CByte(AscW(Mid$(key, i + 1, 1)))
    Next i

    h = RdvHashBytes(kb, 0)
    idx = h And m_maskB
    bi = -1
    Do
        r = m_slB(idx)
        If r = 0 Then Exit Do
        q = m_stB(r)
        If m_bB(q) = kb(0) And m_bB(q + 1) = kb(1) And m_bB(q + 2) = kb(2) _
           And m_bB(q + 3) = kb(3) And m_bB(q + 4) = kb(4) And m_bB(q + 5) = kb(5) _
           And m_bB(q + 6) = kb(6) And m_bB(q + 7) = kb(7) Then
            bi = r - 1
            Exit Do
        End If
        idx = (idx + 1) And m_maskB
    Loop
    If bi < 0 Then Exit Sub

    m_found = True
    ai = m_abIdx(bi)
    ci = m_bcIdx(bi)

    parts = Split(m_hdrB, ",")
    For i = 0 To RDV_FIELDS - 1
        If i <= UBound(parts) Then m_nameB(i) = parts(i)
    Next i
    SplitRow m_bB, m_stB(bi + 1), m_enB(bi + 1), m_valB
    m_key2 = m_valB(1)

    If ai >= 0 Then
        parts = Split(m_hdrA, ",")
        For i = 0 To RDV_FIELDS - 1
            If i <= UBound(parts) Then m_nameA(i) = parts(i)
        Next i
        SplitRow m_bA, m_stA(ai + 1), m_enA(ai + 1), m_valA
    End If
    If ci >= 0 Then
        parts = Split(m_hdrC, ",")
        For i = 0 To RDV_FIELDS - 1
            If i <= UBound(parts) Then m_nameC(i) = parts(i)
        Next i
        SplitRow m_bC, m_stC(ci + 1), m_enC(ci + 1), m_valC
    End If
End Sub

Private Sub SplitRow(ByRef b() As Byte, ByVal s As Long, ByVal e As Long, ByRef v() As String)
    Dim f As Long, p As Long, q As Long, k As Long
    Dim tmp() As Byte
    p = s
    For f = 0 To RDV_FIELDS - 1
        q = p
        Do While q < e
            If b(q) = RDV_COMMA Then Exit Do
            q = q + 1
        Loop
        If q > p Then
            ReDim tmp(0 To q - p - 1)
            For k = 0 To q - p - 1
                tmp(k) = b(p + k)
            Next k
            v(f) = StrConv(tmp, vbUnicode)
        Else
            v(f) = ""
        End If
        p = q + 1
        If p > e Then p = e
    Next f
End Sub

' free everything between runs so the next merge-select starts from cold arrays
Public Sub RdvReset()
    Erase m_bA
    Erase m_stA
    Erase m_enA
    Erase m_slA
    Erase m_bB
    Erase m_stB
    Erase m_enB
    Erase m_slB
    Erase m_bC
    Erase m_stC
    Erase m_enC
    Erase m_slC
    Erase m_abIdx
    Erase m_bcIdx
    m_rowsA = 0
    m_rowsB = 0
    m_rowsC = 0
End Sub
