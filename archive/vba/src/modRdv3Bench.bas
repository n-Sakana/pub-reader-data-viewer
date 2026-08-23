Attribute VB_Name = "modRdv3Bench"
'==============================================================================
' modRdv3Bench -- the measuring harness for the three save methods. It is NOT
' part of any distribution: build_workbook_app.ps1 does not import it, and
' bench_save.ps1 loads it into an Excel instance it started itself.
'
' What one measured item is, for all three methods alike:
'
'     set ONE processed flag in the 100,000-row ledger workbook and commit it
'     to disk, with nothing held back for later
'
' so the numbers are comparable by construction. The ledger copy, the target
' rows, the row count and the machine are the same for every method; only the
' way the bytes reach the disk differs. The sidecar mirror the product also
' writes is NOT part of the measured item -- it is identical work in all three
' methods and would only add the same constant to each -- and it is measured
' once on its own so the product-level cost can still be read off.
'
' Verification is a separate pass over the file that was actually produced:
' every target row must have kept its flag and every other row must be
' unchanged, checked against the baseline mirror of the copy before the run.
'==============================================================================
Option Explicit

Private m_lines() As String
Private m_proc() As Boolean
Private m_rows As Long
Private m_tsv As String

' one open run, driven one item per call: a single COM call that blocked for a
' whole 100-item run would be minutes long, and the harness could not report
' progress or survive a stall. Open, step, step, ..., close.
Private m_runMethod As Long
Private m_runPath As String
Private m_runWb As Object
Private m_runWs As Object
Private m_runMt As Object
Private m_runMarks As Long
Private m_runOpen As Boolean
Private m_writeMeta As Boolean         ' the shipped mark also stamps META

'------------------------------------------------------------------------------
' open / step / close
'------------------------------------------------------------------------------
Public Function Rdv3BenchOpen(ByVal methodName As String, ByVal ledgerPath As String, _
                              ByVal tsvPath As String) As String
    Dim errMsg As String
    Dim t As Double
    Dim openMs As Double
    Dim note As String

    If m_runOpen Then
        Rdv3BenchOpen = "ERR a run is already open"
        Exit Function
    End If
    m_tsv = tsvPath
    ' "book-flagonly" is the same method with the META stamp left out, so the
    ' package diff shows what EXCEL changes rather than what the harness does
    m_writeMeta = (LCase$(methodName) <> "book-flagonly")
    m_runMethod = Rdv3SaveMethodId(Replace(LCase$(methodName), "-flagonly", ""))
    m_runPath = ledgerPath
    If Not LoadSidecar(Rdv3SidecarPath(ledgerPath), errMsg) Then
        Rdv3BenchOpen = "ERR sidecar: " & errMsg
        Exit Function
    End If
    Rdv3SaveReset

    On Error GoTo Fail
    ' Nothing method-specific happens here. Opening the ledger workbook,
    ' settling the ADO provider, walking the deflate stream -- all of that is
    ' what the first save costs in normal use, so it belongs to the first
    ' record, not to a setup step outside the measurement.
    m_runOpen = True
    Rdv3BenchOpen = "ok rows=" & CStr(m_rows)
    Exit Function
Fail:
    Rdv3BenchOpen = "ERR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not m_runWb Is Nothing Then m_runWb.Close False
    Set m_runWb = Nothing
End Function

' ONE record, committed to disk before this returns
Public Function Rdv3BenchStep(ByVal idx As Long, ByVal row As Long) As String
    Dim t As Double
    Dim ms As Double
    Dim ok As Boolean
    Dim errMsg As String
    Dim note As String

    If Not m_runOpen Then
        Rdv3BenchStep = "ERR no run is open"
        Exit Function
    End If
    If row < 0 Or row >= m_rows Then
        Rdv3BenchStep = "ERR row out of range: " & CStr(row)
        Exit Function
    End If
    On Error GoTo Fail
    t = Rdv3Ticks()
    Select Case m_runMethod
        Case RDV3_SAVE_BOOK
            If m_runWb Is Nothing Then
                ' the open is part of the first save, the way it is in the app
                Set m_runWb = Application.Workbooks.Open(m_runPath, 0, False)
                If m_runWb.ReadOnly Then
                    m_runWb.Close False
                    Set m_runWb = Nothing
                    Rdv3BenchStep = "ERR ledger opened read-only"
                    Exit Function
                End If
                Set m_runWs = m_runWb.Worksheets(RDV3_LB_SHEET)
                Set m_runMt = m_runWb.Worksheets(RDV3_LB_META)
                If IsNumeric(m_runMt.Range(RDV3_LM_MARKS).Value2) Then
                    m_runMarks = CLng(m_runMt.Range(RDV3_LM_MARKS).Value2)
                End If
                note = "opened the ledger workbook"
            End If
            m_runWs.Cells(2 + row, 1).Value2 = True
            If m_writeMeta Then
                ' what the shipped mark does: it also stamps the META sheet,
                ' which is a second part changed on top of the flag
                m_runMarks = m_runMarks + 1
                m_runMt.Range(RDV3_LM_MARKS).Value2 = m_runMarks
                m_runMt.Range(RDV3_LM_SAVED).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
            End If
            m_runWb.Save
            ok = True
        Case RDV3_SAVE_ADO
            ok = Rdv3SaveOneAdo(m_runPath, 2 + row, True, errMsg)
        Case RDV3_SAVE_ZIP
            ok = Rdv3SaveOneZip(m_runPath, 2 + row, m_rows, True, errMsg)
            note = Rdv3SaveLastNote()
    End Select
    ms = Rdv3MsSince(t)
    If Not ok Then
        Tsv Rdv3SaveMethodName(m_runMethod), idx, row, ms, "FAIL " & errMsg
        Rdv3BenchStep = "ERR " & errMsg
        Exit Function
    End If
    m_proc(row) = True
    Tsv Rdv3SaveMethodName(m_runMethod), idx, row, ms, note
    Rdv3BenchStep = "ok ms=" & Format$(ms, "0.0") & " file=" & CStr(FileLen(m_runPath)) & _
        IIf(Len(note) > 0, " " & note, "")
    Exit Function
Fail:
    Rdv3BenchStep = "ERR " & Err.Number & ": " & Err.Description
End Function

Public Function Rdv3BenchClose() As String
    On Error Resume Next
    If Not m_runWb Is Nothing Then
        m_runWb.Close False
        Set m_runWb = Nothing
    End If
    Set m_runWs = Nothing
    Set m_runMt = Nothing
    m_runOpen = False
    Rdv3BenchClose = "ok file=" & CStr(FileLen(m_runPath))
End Function

' the sidecar write the product also does on every mark, measured on its own so
' the product-level per-record cost can be read off the method figures
Public Function Rdv3BenchSidecar(ByVal ledgerPath As String) As String
    Dim errMsg As String
    Dim t As Double
    Dim body() As String
    Dim i As Long
    Dim head As String
    Dim txt As String
    Dim f As Integer
    Dim tmp As String
    Dim b() As Byte

    If Not LoadSidecar(Rdv3SidecarPath(ledgerPath), errMsg) Then
        Rdv3BenchSidecar = "ERR " & errMsg
        Exit Function
    End If
    On Error GoTo Fail
    t = Rdv3Ticks()
    ReDim body(0 To m_rows - 1)
    For i = 0 To m_rows - 1
        body(i) = IIf(m_proc(i), "1", "0") & vbTab & m_lines(i)
    Next i
    head = "rdv3state" & vbTab & "1" & vbTab & "rows=" & CStr(m_rows)
    txt = head & vbCrLf & Join(body, vbCrLf) & vbCrLf
    tmp = ledgerPath & ".sidecarbench"
    If Len(Dir$(tmp)) > 0 Then Kill tmp
    f = FreeFile
    Open tmp For Binary Access Write As #f
    b = txt
    Put #f, 1, b
    Close #f
    Rdv3BenchSidecar = "ok sidecar_ms=" & Format$(Rdv3MsSince(t), "0.0") & _
        " bytes=" & CStr(FileLen(tmp))
    Kill tmp
    Exit Function
Fail:
    Rdv3BenchSidecar = "ERR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #f
End Function

'------------------------------------------------------------------------------
' verification: reopen what was produced and check it row by row
'------------------------------------------------------------------------------
Public Function Rdv3BenchVerify(ByVal ledgerPath As String, ByVal baselineSidecar As String, _
                                ByVal rowsCsv As String) As String
    Dim wb As Object
    Dim ws As Object
    Dim arr As Variant
    Dim r0 As Long, r1 As Long, i As Long, c As Long, nn As Long
    Dim parts(0 To RDV3_CONTENT_COLS - 1) As String
    Dim line As String
    Dim errMsg As String
    Dim want As Object
    Dim rowsArr As Variant
    Dim k As Long
    Dim procNow As Boolean
    Dim badContent As Long
    Dim badProc As Long
    Dim targetOk As Long
    Dim firstBad As String
    Dim rows As Long
    Dim t As Double

    If Not LoadSidecar(baselineSidecar, errMsg) Then
        Rdv3BenchVerify = "ERR baseline: " & errMsg
        Exit Function
    End If
    Set want = CreateObject("Scripting.Dictionary")
    rowsArr = Split(rowsCsv, ",")
    For k = LBound(rowsArr) To UBound(rowsArr)
        want(CLng(Val(Trim$(CStr(rowsArr(k)))))) = True
    Next k

    On Error GoTo Fail
    t = Rdv3Ticks()
    Set wb = Application.Workbooks.Open(ledgerPath, 0, True)
    ' the whole workbook has to survive a save method, not just the one sheet
    ' it wrote: a method that drops the META sheet would still pass a row check
    Set ws = Nothing
    For c = 1 To wb.Worksheets.Count
        If wb.Worksheets(c).Name = RDV3_LB_META Then Set ws = wb.Worksheets(c)
    Next c
    If ws Is Nothing Then
        Rdv3BenchVerify = "ERR the META sheet is gone from the workbook"
        GoTo CloseOut
    End If
    If Not IsNumeric(ws.Range(RDV3_LM_ROWS).Value2) Then
        Rdv3BenchVerify = "ERR META rows is not a number after the save"
        GoTo CloseOut
    End If
    Set ws = wb.Worksheets(RDV3_LB_SHEET)
    For c = 1 To RDV3_LED_COLS
        If CStr(ws.Cells(1, c).Value2) <> Rdv3LedgerHead(c - 1) Then
            Rdv3BenchVerify = "ERR header column " & CStr(c) & " = [" & _
                CStr(ws.Cells(1, c).Value2) & "]"
            GoTo CloseOut
        End If
    Next c

    rows = m_rows
    r0 = 0
    Do While r0 < rows
        r1 = r0 + 16384
        If r1 > rows Then r1 = rows
        arr = ws.Range(ws.Cells(2 + r0, 1), ws.Cells(1 + r1, RDV3_LED_COLS)).Value2
        nn = r1 - r0
        For i = 1 To nn
            For c = 0 To RDV3_CONTENT_COLS - 1
                If IsEmpty(arr(i, c + 2)) Or IsError(arr(i, c + 2)) Then
                    parts(c) = ""
                Else
                    parts(c) = CStr(arr(i, c + 2))
                End If
            Next c
            line = Join(parts, vbTab)
            If VarType(arr(i, 1)) = vbBoolean Then
                procNow = arr(i, 1)
            Else
                procNow = (UCase$(CStr(arr(i, 1))) = "TRUE")
            End If
            k = r0 + i - 1
            If StrComp(line, m_lines(k), vbBinaryCompare) <> 0 Then
                badContent = badContent + 1
                If Len(firstBad) = 0 Then
                    firstBad = "content row " & CStr(k + 1) & " got[" & Left$(line, 60) & _
                        "] want[" & Left$(m_lines(k), 60) & "]"
                End If
            End If
            If want.Exists(k) Then
                If procNow Then
                    targetOk = targetOk + 1
                Else
                    badProc = badProc + 1
                    If Len(firstBad) = 0 Then firstBad = "target row " & CStr(k + 1) & " is FALSE"
                End If
            Else
                If procNow <> m_proc(k) Then
                    badProc = badProc + 1
                    If Len(firstBad) = 0 Then firstBad = "non-target row " & CStr(k + 1) & " changed"
                End If
            End If
        Next i
        r0 = r1
    Loop

    Rdv3BenchVerify = "ok rows=" & CStr(rows) & " targets=" & CStr(want.Count) & _
        " target_true=" & CStr(targetOk) & " content_mismatch=" & CStr(badContent) & _
        " proc_mismatch=" & CStr(badProc) & " read_ms=" & Format$(Rdv3MsSince(t), "0.0") & _
        IIf(Len(firstBad) > 0, " first=" & firstBad, "")
CloseOut:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    Set wb = Nothing
    Exit Function
Fail:
    Rdv3BenchVerify = "ERR " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
End Function

' the key2 of a ledger row, so the harness can name its targets by key
Public Function Rdv3BenchRowOfKey2(ByVal sidecarPath As String, ByVal keysCsv As String) As String
    Dim errMsg As String
    Dim d As Object
    Dim i As Long
    Dim p As Long
    Dim q As Long
    Dim k2 As String
    Dim keys As Variant
    Dim out As String
    Dim j As Long

    If Not LoadSidecar(sidecarPath, errMsg) Then
        Rdv3BenchRowOfKey2 = "ERR " & errMsg
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    For i = 0 To m_rows - 1
        p = InStr(1, m_lines(i), vbTab, vbBinaryCompare)
        q = InStr(p + 1, m_lines(i), vbTab, vbBinaryCompare)
        If q = 0 Then q = Len(m_lines(i)) + 1
        k2 = Mid$(m_lines(i), p + 1, q - p - 1)
        If Not d.Exists(k2) Then d.Add k2, i
    Next i
    keys = Split(keysCsv, ",")
    For j = LBound(keys) To UBound(keys)
        k2 = Trim$(CStr(keys(j)))
        If Not d.Exists(k2) Then
            Rdv3BenchRowOfKey2 = "ERR key2 not found: " & k2
            Exit Function
        End If
        If Len(out) > 0 Then out = out & ","
        out = out & CStr(d.Item(k2))
    Next j
    Rdv3BenchRowOfKey2 = out
End Function

' every key2 in ledger-row order, so the harness can pick spread-out targets
Public Function Rdv3BenchKey2At(ByVal sidecarPath As String, ByVal rowsCsv As String) As String
    Dim errMsg As String
    Dim rowsArr As Variant
    Dim j As Long
    Dim i As Long
    Dim p As Long
    Dim q As Long
    Dim out As String

    If Not LoadSidecar(sidecarPath, errMsg) Then
        Rdv3BenchKey2At = "ERR " & errMsg
        Exit Function
    End If
    rowsArr = Split(rowsCsv, ",")
    For j = LBound(rowsArr) To UBound(rowsArr)
        i = CLng(Val(Trim$(CStr(rowsArr(j)))))
        If i < 0 Or i >= m_rows Then
            Rdv3BenchKey2At = "ERR row out of range: " & CStr(i)
            Exit Function
        End If
        p = InStr(1, m_lines(i), vbTab, vbBinaryCompare)
        q = InStr(p + 1, m_lines(i), vbTab, vbBinaryCompare)
        If q = 0 Then q = Len(m_lines(i)) + 1
        If Len(out) > 0 Then out = out & ","
        out = out & Mid$(m_lines(i), p + 1, q - p - 1)
    Next j
    Rdv3BenchKey2At = out
End Function

Public Function Rdv3BenchRows(ByVal sidecarPath As String) As String
    Dim errMsg As String
    If Not LoadSidecar(sidecarPath, errMsg) Then
        Rdv3BenchRows = "ERR " & errMsg
        Exit Function
    End If
    Rdv3BenchRows = CStr(m_rows)
End Function

' the throughput of a plain VBA byte loop on this machine, which is the ceiling
' any hand-written scan (checksum, compressor, parser) has to live under
Public Function Rdv3BenchByteLoop(ByVal mb As Long) As String
    Dim b() As Byte
    Dim i As Long
    Dim n As Long
    Dim s As Long
    Dim t As Double
    Dim ms As Double
    n = mb * 1048576
    ReDim b(0 To n - 1)
    For i = 0 To 255
        b(i) = i
    Next i
    t = Rdv3Ticks()
    For i = 0 To n - 1
        s = (s Xor CLng(b(i))) And &HFFFF&
    Next i
    ms = Rdv3MsSince(t)
    Rdv3BenchByteLoop = "bytes=" & CStr(n) & " ms=" & Format$(ms, "0.0") & _
        " mb_per_s=" & Format$(CDbl(mb) / (ms / 1000#), "0.00") & " sink=" & CStr(s)
End Function

'------------------------------------------------------------------------------
' the sidecar reader (the same UTF-16LE mirror the BE writes)
'------------------------------------------------------------------------------
Private Function LoadSidecar(ByVal path As String, ByRef errMsg As String) As Boolean
    Dim f As Integer
    Dim b() As Byte
    Dim n As Long
    Dim txt As String
    Dim lines As Variant
    Dim head As Variant
    Dim i As Long
    Dim eq As Long
    Dim rows As Long
    Dim ln As String

    errMsg = ""
    m_rows = 0
    If Dir$(path) = "" Then
        errMsg = "sidecar missing: " & path
        Exit Function
    End If
    On Error GoTo Fail
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If (n Mod 2) <> 0 Then n = n - 1
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f
    txt = b
    Erase b
    lines = Split(txt, vbCrLf)
    txt = ""
    head = Split(CStr(lines(0)), vbTab)
    If CStr(head(0)) <> "rdv3state" Then
        errMsg = "sidecar magic mismatch"
        Exit Function
    End If
    rows = -1
    For i = 2 To UBound(head)
        eq = InStr(CStr(head(i)), "=")
        If eq > 1 Then
            If Left$(CStr(head(i)), eq - 1) = "rows" Then rows = CLng(Val(Mid$(CStr(head(i)), eq + 1)))
        End If
    Next i
    If rows < 0 Or UBound(lines) < rows Then
        errMsg = "sidecar row count mismatch"
        Exit Function
    End If
    ReDim m_lines(0 To rows - 1)
    ReDim m_proc(0 To rows - 1)
    For i = 0 To rows - 1
        ln = CStr(lines(1 + i))
        m_proc(i) = (Left$(ln, 1) = "1")
        m_lines(i) = Mid$(ln, 3)
    Next i
    m_rows = rows
    LoadSidecar = True
    Exit Function
Fail:
    errMsg = "sidecar read error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    Close #f
End Function

'------------------------------------------------------------------------------
Private Sub Tsv(ByVal method As String, ByVal idx As Long, ByVal row As Long, _
                ByVal ms As Double, ByVal extra As String)
    Dim f As Integer
    If Len(m_tsv) = 0 Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Open m_tsv For Append As #f
    Print #f, method & vbTab & CStr(idx) & vbTab & CStr(row) & vbTab & _
        Format$(ms, "0.0") & vbTab & extra
    Close #f
    On Error GoTo 0
End Sub

Public Function Rdv3BenchTouch() As String
    Rdv3BenchTouch = "bench ok " & Rdv3SaveSelfTest()
End Function
