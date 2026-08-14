Attribute VB_Name = "modRdv3Chan"
'==============================================================================
' modRdv3Chan -- the file channel between FE and BE. Modeled directly on a
' proven file channel from an earlier project: the worker never calls into
' the FE process, results are published to files atomically (write "<path>.tmp"
' then Kill the target and Name the tmp over it), and the FE pulls them from
' its own OnTime pump. A reader that finds a file missing or unreadable skips
' that cycle and retries on the next.
'
' Why files and not COM: a cross-process COM write parks against an FE whose
' user is editing a cell, the OLE channel retries flood the message queue, and
' the FE's keyboard input starves -- a permanent freeze that worker-side
' mitigation cannot fix (measured in that earlier project).
'
' Files (all under %TEMP%\rdv3\, session-scoped by an FE-generated id):
'   rdv3_<sid>_agg.dat        BE -> FE: latest record per kind, version-tagged
'   rdv3_<sid>_req_<k>.dat    FE -> BE: latest request of kind k, version-tagged
'   rdv3_<sid>_stop.flag      FE -> BE: stop now
'   rdv3_<sid>_fe_gone.flag   FE tombstone written at clean shutdown
'   rdv3_<sid>_fe_lease.lock  held open (deny-all) by the FE for its lifetime;
'                             the BE probes the lock: locked = FE alive (even
'                             mid cell-edit), unlocked/missing = FE gone
'   rdv3_<sid>_be_pid.flag    BE -> FE: the BE's own process id, written first
'                             thing in bootstrap (the FE spawns the BE through
'                             a detached cscript and never holds a COM ref)
'   rdv3_<sid>_spawn_err.flag spawn script -> FE: the spawn failed, with why
'   rdv3_<sid>_spawn.vbs      the spawn script the FE shells out to
'   rdv3_<sid>_be_done.flag   BE exit note (diagnostic)
'   rdv3_<sid>_worker.xlsm    the extracted worker book copy
'
' The 100k-row ledger itself never crosses this channel: it lives in its own
' ledger workbook owned entirely by the BE (read, compare, carry, write, save).
' The FE only ever receives display-sized RESULT records.
'
' Aggregate record, one line per kind (tab separated):
'   r <TAB> kind <TAB> version <TAB> meta(k=v;...) <TAB> rows <TAB> cols <TAB> cells...
' Cells are encoded S<text> / N<number> / E (empty), tabs stripped from text.
' The session id comes from a timestamp, never from Application.hWnd (which
' can change between startup and later calls).
'
' This module may reference modRdv3Spec only: it compiles in both books.
'==============================================================================
Option Explicit

Public Const RDV3_K_CHECK As String = "CHECK"
Public Const RDV3_K_READY As String = "READY"
Public Const RDV3_K_RESULT As String = "RESULT"
Public Const RDV3_K_STATE As String = "STATE"
Public Const RDV3_K_APPLY As String = "APPLY"
Public Const RDV3_K_BEERR As String = "BEERR"

Public Const RDV3_RQ_DECISION As String = "decision"
Public Const RDV3_RQ_SEARCH As String = "search"
Public Const RDV3_RQ_PICK As String = "pick"
Public Const RDV3_RQ_MARK As String = "mark"
Public Const RDV3_RQ_WATCH As String = "watch"

' FE lease handle (held open for the whole session; see Rdv3ChEnsureLease)
Private m_leaseNo As Integer
Private m_leaseSid As String

'------------------------------------------------------------------------------
' paths
'------------------------------------------------------------------------------
Public Function Rdv3ChRoot() As String
    Dim root As String
    root = Environ$("TEMP") & "\rdv3"
    If Dir$(root, vbDirectory) = "" Then MkDir root
    Rdv3ChRoot = root
End Function

Public Function Rdv3ChBase(ByVal sid As String) As String
    Rdv3ChBase = Rdv3ChRoot() & "\rdv3_" & sid
End Function

Public Function Rdv3ChAggPath(ByVal sid As String) As String
    Rdv3ChAggPath = Rdv3ChBase(sid) & "_agg.dat"
End Function

Public Function Rdv3ChReqPath(ByVal sid As String, ByVal kind As String) As String
    Rdv3ChReqPath = Rdv3ChBase(sid) & "_req_" & kind & ".dat"
End Function

Public Function Rdv3ChStopPath(ByVal sid As String) As String
    Rdv3ChStopPath = Rdv3ChBase(sid) & "_stop.flag"
End Function

Public Function Rdv3ChGonePath(ByVal sid As String) As String
    Rdv3ChGonePath = Rdv3ChBase(sid) & "_fe_gone.flag"
End Function

Public Function Rdv3ChLeasePath(ByVal sid As String) As String
    Rdv3ChLeasePath = Rdv3ChBase(sid) & "_fe_lease.lock"
End Function

Public Function Rdv3ChBePidPath(ByVal sid As String) As String
    Rdv3ChBePidPath = Rdv3ChBase(sid) & "_be_pid.flag"
End Function

Public Function Rdv3ChSpawnErrPath(ByVal sid As String) As String
    Rdv3ChSpawnErrPath = Rdv3ChBase(sid) & "_spawn_err.flag"
End Function

Public Function Rdv3ChSpawnVbsPath(ByVal sid As String) As String
    Rdv3ChSpawnVbsPath = Rdv3ChBase(sid) & "_spawn.vbs"
End Function

Public Function Rdv3ChBeDonePath(ByVal sid As String) As String
    Rdv3ChBeDonePath = Rdv3ChBase(sid) & "_be_done.flag"
End Function

Public Function Rdv3ChWorkerCopyPath(ByVal sid As String) As String
    Rdv3ChWorkerCopyPath = Rdv3ChBase(sid) & "_worker.xlsm"
End Function

'------------------------------------------------------------------------------
' small flags
'------------------------------------------------------------------------------
Public Function Rdv3ChFlagExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    Rdv3ChFlagExists = (Len(Dir$(path)) > 0)
    On Error GoTo 0
End Function

Public Sub Rdv3ChWriteFlag(ByVal path As String, ByVal text As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
    On Error GoTo 0
End Sub

Public Sub Rdv3ChDeleteQuiet(ByVal path As String)
    If Len(path) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' atomic write: tmp -> Kill dest -> Name tmp As dest
'------------------------------------------------------------------------------
Private Function AtomicReplace(ByVal tmpPath As String, ByVal path As String) As Boolean
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    Err.Clear
    Name tmpPath As path
    If Err.Number <> 0 Then
        Kill tmpPath
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    AtomicReplace = True
End Function

'------------------------------------------------------------------------------
' cell encoding (canonical N/S/E)
'------------------------------------------------------------------------------
Public Function Rdv3ChEncodeCell(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbEmpty, vbNull
            Rdv3ChEncodeCell = "E"
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte, vbDecimal, vbDate, vbBoolean
            Rdv3ChEncodeCell = "N" & Trim$(Str$(CDbl(v)))
        Case Else
            Rdv3ChEncodeCell = "S" & Replace(Replace(Replace(CStr(v), vbTab, " "), vbCr, " "), vbLf, " ")
    End Select
End Function

Public Function Rdv3ChDecodeCell(ByVal token As String) As Variant
    If Len(token) = 0 Then
        Rdv3ChDecodeCell = Empty
    ElseIf Left$(token, 1) = "N" Then
        Rdv3ChDecodeCell = Val(Mid$(token, 2))
    ElseIf Left$(token, 1) = "S" Then
        Rdv3ChDecodeCell = Mid$(token, 2)
    Else
        Rdv3ChDecodeCell = Empty
    End If
End Function

Private Function CleanMeta(ByVal s As String) As String
    CleanMeta = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

'------------------------------------------------------------------------------
' aggregate: BE publishes, FE reads once per pump tick. One writer (the single
' BE) so no cross-writer lock is needed; atomic replace still protects readers.
'------------------------------------------------------------------------------
Public Function Rdv3ChPublish(ByVal sid As String, ByVal kind As String, ByVal version As Long, _
                              ByVal meta As String, ByVal values As Variant) As Boolean
    Dim path As String
    Dim tmp As String
    Dim inNo As Integer, outNo As Integer
    Dim lines As Collection
    Dim ln As String
    Dim parts As Variant
    Dim i As Long
    Dim rec As String

    path = Rdv3ChAggPath(sid)
    tmp = path & ".tmp"
    Set lines = New Collection
    On Error GoTo Failed

    If Len(Dir$(path)) > 0 Then
        inNo = FreeFile
        Open path For Input Shared As #inNo
        Do Until EOF(inNo)
            Line Input #inNo, ln
            parts = Split(ln, vbTab)
            If UBound(parts) < 1 Then
                lines.Add ln
            ElseIf CStr(parts(0)) <> "r" Or CStr(parts(1)) <> kind Then
                lines.Add ln
            End If
        Loop
        Close #inNo
        inNo = 0
    End If

    rec = EncodeRecord(kind, version, meta, values)

    outNo = FreeFile
    Open tmp For Output As #outNo
    For i = 1 To lines.Count
        Print #outNo, CStr(lines(i))
    Next i
    Print #outNo, rec
    Close #outNo
    outNo = 0

    Rdv3ChPublish = AtomicReplace(tmp, path)
    Exit Function
Failed:
    On Error Resume Next
    If inNo <> 0 Then Close #inNo
    If outNo <> 0 Then Close #outNo
    If Len(Dir$(tmp)) > 0 Then Kill tmp
    On Error GoTo 0
End Function

Private Function EncodeRecord(ByVal kind As String, ByVal version As Long, _
                              ByVal meta As String, ByVal values As Variant) As String
    Dim rows As Long, cols As Long
    Dim r As Long, c As Long
    Dim s As String
    If IsArray(values) Then
        rows = UBound(values, 1) - LBound(values, 1) + 1
        cols = UBound(values, 2) - LBound(values, 2) + 1
    Else
        rows = 0
        cols = 0
    End If
    s = "r" & vbTab & kind & vbTab & CStr(version) & vbTab & CleanMeta(meta) & vbTab & _
        CStr(rows) & vbTab & CStr(cols)
    If rows > 0 Then
        For r = LBound(values, 1) To UBound(values, 1)
            For c = LBound(values, 2) To UBound(values, 2)
                s = s & vbTab & Rdv3ChEncodeCell(values(r, c))
            Next c
        Next r
    End If
    EncodeRecord = s
End Function

' Reads every kind's latest record. Callers dedup by version per kind.
' Returns a Collection of Variant arrays: (kind, version, metaString, values2D-or-Empty)
Public Function Rdv3ChReadAgg(ByVal sid As String) As Collection
    Dim path As String
    Dim f As Integer
    Dim ln As String
    Dim recs As Collection
    Dim one As Variant

    Set recs = New Collection
    path = Rdv3ChAggPath(sid)
    If Not Rdv3ChFlagExists(path) Then
        Set Rdv3ChReadAgg = recs
        Exit Function
    End If
    On Error GoTo Failed
    f = FreeFile
    Open path For Input Shared As #f
    Do Until EOF(f)
        Line Input #f, ln
        one = DecodeRecord(ln)
        If IsArray(one) Then recs.Add one
    Loop
    Close #f
    f = 0
Failed:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
    Set Rdv3ChReadAgg = recs
End Function

Private Function DecodeRecord(ByVal ln As String) As Variant
    Dim parts As Variant
    Dim rows As Long, cols As Long
    Dim vals As Variant
    Dim r As Long, c As Long, at As Long
    On Error GoTo Failed
    parts = Split(ln, vbTab)
    If UBound(parts) < 5 Then Exit Function
    If CStr(parts(0)) <> "r" Then Exit Function
    rows = CLng(Val(parts(4)))
    cols = CLng(Val(parts(5)))
    If rows < 0 Or cols < 0 Or rows > 10000 Or cols > 1000 Then Exit Function
    If rows > 0 Then
        If UBound(parts) < 5 + rows * cols Then Exit Function
        ReDim vals(1 To rows, 1 To cols)
        at = 6
        For r = 1 To rows
            For c = 1 To cols
                vals(r, c) = Rdv3ChDecodeCell(CStr(parts(at)))
                at = at + 1
            Next c
        Next r
    Else
        vals = Empty
    End If
    DecodeRecord = Array(CStr(parts(1)), CLng(Val(parts(2))), CStr(parts(3)), vals)
    Exit Function
Failed:
End Function

' meta string "k=v;k=v" -> Scripting.Dictionary
Public Function Rdv3ChMeta(ByVal metaString As String) As Object
    Dim d As Object
    Dim parts As Variant
    Dim i As Long, eq As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    parts = Split(metaString, ";")
    For i = LBound(parts) To UBound(parts)
        eq = InStr(parts(i), "=")
        If eq > 1 Then d(Left$(parts(i), eq - 1)) = Mid$(parts(i), eq + 1)
    Next i
    Set Rdv3ChMeta = d
End Function

'------------------------------------------------------------------------------
' requests: FE writes the latest request of a kind, BE consumes by version
'------------------------------------------------------------------------------
Public Function Rdv3ChWriteReq(ByVal sid As String, ByVal kind As String, _
                               ByVal version As Long, ByVal args As String) As Boolean
    Dim path As String
    Dim tmp As String
    Dim f As Integer
    path = Rdv3ChReqPath(sid, kind)
    tmp = path & ".tmp"
    On Error GoTo Failed
    f = FreeFile
    Open tmp For Output As #f
    Print #f, "q" & vbTab & CStr(version) & vbTab & CleanMeta(args)
    Close #f
    f = 0
    Rdv3ChWriteReq = AtomicReplace(tmp, path)
    Exit Function
Failed:
    On Error Resume Next
    If f <> 0 Then Close #f
    If Len(Dir$(tmp)) > 0 Then Kill tmp
    On Error GoTo 0
End Function

Public Function Rdv3ChReadReq(ByVal sid As String, ByVal kind As String, _
                              ByRef version As Long, ByRef args As String) As Boolean
    Dim path As String
    Dim f As Integer
    Dim ln As String
    Dim parts As Variant
    path = Rdv3ChReqPath(sid, kind)
    If Not Rdv3ChFlagExists(path) Then Exit Function
    On Error GoTo Failed
    f = FreeFile
    Open path For Input Shared As #f
    Line Input #f, ln
    Close #f
    f = 0
    parts = Split(ln, vbTab)
    If UBound(parts) < 2 Then Exit Function
    If CStr(parts(0)) <> "q" Then Exit Function
    version = CLng(Val(parts(1)))
    args = CStr(parts(2))
    Rdv3ChReadReq = True
    Exit Function
Failed:
    On Error Resume Next
    If f <> 0 Then Close #f
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' FE lease: the FE keeps the file open with a
' deny-all lock; the BE probes by trying to lock it. Locked = FE alive (even
' mid cell-edit). A crash releases the lock the instant the process dies.
'------------------------------------------------------------------------------
Public Function Rdv3ChEnsureLease(ByVal sid As String) As Boolean
    Dim f As Integer
    If m_leaseNo <> 0 Then
        Rdv3ChEnsureLease = True
        Exit Function
    End If
    On Error Resume Next
    Rdv3ChDeleteQuiet Rdv3ChLeasePath(sid)
    Err.Clear
    f = FreeFile
    Open Rdv3ChLeasePath(sid) For Output Lock Read Write As #f
    If Err.Number = 0 Then
        Print #f, "fe_pid=" & CStr(Rdv3GetCurrentProcessId())
        m_leaseNo = f
        m_leaseSid = sid
        Rdv3ChEnsureLease = True
    Else
        Close #f
    End If
    On Error GoTo 0
End Function

Public Sub Rdv3ChReleaseLease()
    If m_leaseNo = 0 Then Exit Sub
    On Error Resume Next
    Close #m_leaseNo
    m_leaseNo = 0
    If Len(m_leaseSid) > 0 Then Rdv3ChDeleteQuiet Rdv3ChLeasePath(m_leaseSid)
    m_leaseSid = ""
    On Error GoTo 0
End Sub

' BE side: True while the FE holds the lease lock
Public Function Rdv3ChFeLeaseAlive(ByVal sid As String) As Boolean
    Dim f As Integer
    Dim errNo As Long
    Dim path As String
    path = Rdv3ChLeasePath(sid)
    If Not Rdv3ChFlagExists(path) Then Exit Function
    On Error Resume Next
    Err.Clear
    f = FreeFile
    Open path For Binary Access Write Lock Read Write As #f
    errNo = Err.Number
    If errNo = 0 Then Close #f
    Err.Clear
    On Error GoTo 0
    Rdv3ChFeLeaseAlive = (errNo <> 0)
End Function

' delete THIS session's channel files (FE shutdown, after the BE exited).
' The lease must already be released; the fe_gone tombstone is kept so a BE
' that somehow lingers still learns the FE left, and the daily sweep takes it.
Public Sub Rdv3ChDeleteSession(ByVal sid As String)
    Dim base As String
    Dim nm As String
    Dim victims As Collection
    Dim v As Variant
    base = Rdv3ChRoot() & "\"
    Set victims = New Collection
    On Error Resume Next
    nm = Dir$(base & "rdv3_" & sid & "_*")
    Do While Len(nm) > 0
        If InStr(nm, "_fe_gone") = 0 Then victims.Add base & nm
        nm = Dir$()
    Loop
    For Each v In victims
        Kill CStr(v)
    Next v
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' session sweep: files of sessions whose FE is gone (lease unlocked) and that
' are older than a day -- best-effort cleanup at boot
'------------------------------------------------------------------------------
Public Function Rdv3ChSweepStale() As Long
    Dim root As String
    Dim nm As String
    Dim cnt As Long
    Dim full As String
    root = Rdv3ChRoot() & "\"
    On Error Resume Next
    nm = Dir$(root & "rdv3_*")
    Do While Len(nm) > 0
        full = root & nm
        If FileDateTime(full) < DateAdd("d", -1, Now) Then
            Kill full             ' a live session's lease is locked, so Kill skips it
            If Err.Number = 0 Then cnt = cnt + 1
            Err.Clear
        End If
        nm = Dir$()
    Loop
    On Error GoTo 0
    Rdv3ChSweepStale = cnt
End Function
