Attribute VB_Name = "modRdv3Chan"
'==============================================================================
' modRdv3Chan -- the file channel between FE and BE. Modeled directly on a
' proven file channel from an earlier project: the worker never calls into
' the FE process, results are published to files atomically (write "<path>.tmp",
' park the live file under "<path>.bak" by rename, rename the tmp into place,
' then drop the parked copy -- see Rdv3ChReplaceFile for why it is that order),
' and the FE pulls them from its own OnTime pump. A reader that finds a file
' missing or unreadable skips that cycle and retries on the next.
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
'   rdv3_<sid>_be_lease.lock  the same thing the other way round: held open by
'                             the BE for its lifetime, probed by the FE. This
'                             is how a BE that died without saying so is
'                             noticed now that no process id is taken (the
'                             OS drops the lock the instant the process ends)
'   rdv3_<sid>_be_done.flag   BE exit note (diagnostic)
'   rdv3_<sid>_worker.xlsm    the extracted worker book copy
'   <any of the above>.tmp   the half-written next version (writer only)
'   <any of the above>.bak   the live version parked during a replace; exists
'                             only inside that window, or after a crash in it
'
' The 100k-row ledger itself never crosses this channel: it lives in its own
' ledger workbook owned entirely by the BE (read, compare, carry, write, save).
' The FE only ever receives display-sized RESULT records.
'
' Aggregate record, one line per kind (tab separated):
'   r <TAB> kind <TAB> version <TAB> meta(k=v;...) <TAB> rows <TAB> cols <TAB> cells...
'
' One line PER KIND means the kind is the delivery slot, and kinds fall into
' two classes:
'   latest-wins   STATE / CHECK / READY / APPLY / BEERR / RESULT -- only the
'                 newest record carries meaning (a heartbeat, the newest search
'                 answer). Losing an older one costs nothing.
'   must-deliver  MARK -- the confirmation that ONE processed record reached
'                 the disk. It gets its own slot precisely so a search answer
'                 cannot take its place, and it is rewritten only by the next
'                 save request, which the exit guard will not allow until this
'                 one has been consumed.
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
' A "processed" save is confirmed in its OWN slot, never in RESULT. The
' aggregate keeps one record per kind, so two records of the SAME kind
' published between two FE ticks lose the older one -- which is right for a
' search (only the newest answer is wanted) and wrong for a save confirmation
' (measured: a search issued 200 ms into a save replaced res=marked 441 ms
' later, the screen never saw the save, and the exit guard then held the book
' for its full 180 s ceiling and called a SUCCESSFUL save undecided;
' work\race-evidence-before\). MARK is only ever rewritten by the NEXT save
' request, and the exit guard refuses one until the FE has taken this one, so
' an unread confirmation cannot be overwritten.
Public Const RDV3_K_MARK As String = "MARK"

' The element the operator pointed the settings screen at. Its own kind for the
' same reason MARK has one: it MUST arrive. A heartbeat or a search answer
' published in the same second would otherwise take its place and the settings
' screen would sit there having asked for something that already happened.
Public Const RDV3_K_INSPECT As String = "INSPECT"

' The BE's answer to a settings save. Its own kind for the same reason: the
' settings screen has to be told whether the running session actually took
' the new values, and a heartbeat must not be able to take its place.
Public Const RDV3_K_CONFIG As String = "CONFIG"

Public Const RDV3_RQ_DECISION As String = "decision"
Public Const RDV3_RQ_SEARCH As String = "search"
Public Const RDV3_RQ_PICK As String = "pick"
' the element picker, which is not the candidate pick above
Public Const RDV3_RQ_INSPECT As String = "inspect"
' re-read the settings file and adopt what may be adopted without a restart
Public Const RDV3_RQ_CONFIG As String = "config"
Public Const RDV3_RQ_MARK As String = "mark"
Public Const RDV3_RQ_WATCH As String = "watch"

' FE lease handle (held open for the whole session; see Rdv3ChEnsureLease)
Private m_leaseNo As Integer
Private m_leaseSid As String

' BE lease handle (the mirror image, held by the BE process)
Private m_beLeaseNo As Integer
Private m_beLeaseSid As String
' the ledger lock (held open for the session; see Rdv3ChLedgerLockTake)
Private m_ledgerLockNo As Integer

' last failure code from Rdv3ChReplaceFile, for the caller's log
Private m_lastReplaceErr As Long

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

Public Function Rdv3ChBeLeasePath(ByVal sid As String) As String
    Rdv3ChBeLeasePath = Rdv3ChBase(sid) & "_be_lease.lock"
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
' atomic replace (was AtomicReplace: tmp -> Kill dest -> Name tmp As dest).
'
' The old order could lose data two ways, and did: a Name that failed AFTER the
' Kill destroyed the live file -- and with it every OTHER kind's record in the
' aggregate -- and the error path then killed the tmp as well, so the new record
' went with it. Neither loss was reported: the function just returned False and
' the next publish rebuilt an aggregate that no longer held the missing record.
' Measured symptom: searches the BE served that the FE never rendered (2 in 39
' on this machine, present before this change too).
'
' The live file is now PARKED by rename, never deleted before the replacement is
' in place:
'
'   path -> path.bak      (park; on failure nothing has moved)
'   tmp  -> path          (on failure the parked file goes straight back)
'   Kill path.bak         (only once the new file is live)
'
' So every failure leaves BOTH the live file and the tmp on disk and returns
' False: the caller can simply try again. A process death between the two
' renames leaves the live file parked, and the next writer-side call puts it
' back (HealParked). A .bak that cannot be deleted afterwards is harmless -- the
' replacement already happened -- and the next call removes it.
'
' The last failure code is kept for the caller's log: nothing here is silent.
'------------------------------------------------------------------------------
Public Function Rdv3ChReplaceFile(ByVal tmpPath As String, ByVal path As String) As Boolean
    Dim bak As String
    Dim e As Long

    m_lastReplaceErr = 0
    bak = path & ".bak"
    On Error Resume Next
    HealParked path

    If Len(Dir$(tmpPath)) = 0 Then
        m_lastReplaceErr = 53                    ' file not found: nothing to put in place
        On Error GoTo 0
        Exit Function
    End If

    If Len(Dir$(path)) = 0 Then
        Err.Clear
        Name tmpPath As path                     ' nothing to displace
        e = Err.Number
        Err.Clear
        m_lastReplaceErr = e
        On Error GoTo 0
        Rdv3ChReplaceFile = (e = 0)
        Exit Function
    End If

    If Len(Dir$(bak)) > 0 Then
        Err.Clear
        Kill bak                                 ' stale park from an earlier failure
        Err.Clear
    End If

    Name path As bak
    e = Err.Number
    Err.Clear
    If e <> 0 Then
        m_lastReplaceErr = e                     ' live file untouched, tmp kept
        On Error GoTo 0
        Exit Function
    End If

    Name tmpPath As path
    e = Err.Number
    Err.Clear
    If e <> 0 Then
        Name bak As path                         ' put the live file back
        Err.Clear
        m_lastReplaceErr = e                     ' tmp kept: the retry reuses it
        On Error GoTo 0
        Exit Function
    End If

    Kill bak
    Err.Clear
    On Error GoTo 0
    Rdv3ChReplaceFile = True
End Function

' the error number of the last failed replace, for the caller's log
Public Function Rdv3ChLastReplaceErr() As Long
    Rdv3ChLastReplaceErr = m_lastReplaceErr
End Function

' A crash between the two renames leaves the live file parked. Only the writer
' of a given file calls this (the BE for the aggregate, the FE for requests), so
' the two processes never rename the same file.
Private Sub HealParked(ByVal path As String)
    Dim bak As String
    bak = path & ".bak"
    On Error Resume Next
    If Len(Dir$(path)) = 0 Then
        If Len(Dir$(bak)) > 0 Then
            Err.Clear
            Name bak As path
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub

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

    ' the other kinds' records are carried over from the live file, so a live
    ' file still parked from an interrupted replace has to be put back BEFORE
    ' that read -- otherwise this publish would quietly write an aggregate that
    ' holds nothing but its own record
    HealParked path

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

    Rdv3ChPublish = Rdv3ChReplaceFile(tmp, path)
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
    Rdv3ChWriteReq = Rdv3ChReplaceFile(tmp, path)
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
'------------------------------------------------------------------------------
' The LEDGER lock: one writer per ledger FILE.
'
' The settings file can point this build at any ledger, so two copies of the FE
' can now be aimed at the same xlsx -- and two processes writing the same
' workbook is how a saved record gets lost. The guard is a lock file beside the
' ledger, opened with Lock Read Write and held for the whole session; the OS
' releases it however the process ends.
'
' It guards the FILE, not the spelling. "x\led.xlsx" and "x\.\led.xlsx" name the
' same lock file on disk, so a second instance is refused however its path was
' written -- no normalising, and nothing to get wrong.
'------------------------------------------------------------------------------
Public Function Rdv3ChLedgerLockPath(ByVal ledgerPath As String) As String
    Rdv3ChLedgerLockPath = ledgerPath & ".lock"
End Function

Public Function Rdv3ChLedgerLockTake(ByVal ledgerPath As String) As Boolean
    Dim f As Integer
    If m_ledgerLockNo <> 0 Then
        Rdv3ChLedgerLockTake = True
        Exit Function
    End If
    On Error Resume Next
    f = FreeFile
    Err.Clear
    Open Rdv3ChLedgerLockPath(ledgerPath) For Output Lock Read Write As #f
    If Err.Number = 0 Then
        Print #f, "owner=" & Rdv3SelfId()
        m_ledgerLockNo = f
        Rdv3ChLedgerLockTake = True
    Else
        Close #f
    End If
    On Error GoTo 0
End Function

Public Sub Rdv3ChLedgerLockRelease()
    On Error Resume Next
    If m_ledgerLockNo <> 0 Then Close #m_ledgerLockNo
    m_ledgerLockNo = 0
    On Error GoTo 0
End Sub

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
        Print #f, "fe=" & Rdv3SelfId()
        m_leaseNo = f
        m_leaseSid = sid
        Rdv3ChEnsureLease = True
    Else
        Close #f
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' BE lease: the mirror of the FE lease. The BE holds it open for its whole life
' and the FE probes it, so "the BE is gone" is answered by the OS releasing a
' lock rather than by asking Win32 about a process id -- and a COM call into a
' possibly busy BE (which would park the FE behind Excel's server-busy dialog)
' is never needed for liveness.
'------------------------------------------------------------------------------
Public Function Rdv3ChBeLeaseOpen(ByVal sid As String) As Boolean
    Dim f As Integer
    If m_beLeaseNo <> 0 Then
        Rdv3ChBeLeaseOpen = True
        Exit Function
    End If
    On Error Resume Next
    Rdv3ChDeleteQuiet Rdv3ChBeLeasePath(sid)
    Err.Clear
    f = FreeFile
    Open Rdv3ChBeLeasePath(sid) For Output Lock Read Write As #f
    If Err.Number = 0 Then
        Print #f, "be=" & Rdv3SelfId()
        m_beLeaseNo = f
        m_beLeaseSid = sid
        Rdv3ChBeLeaseOpen = True
    Else
        Close #f
    End If
    On Error GoTo 0
End Function

Public Sub Rdv3ChBeLeaseRelease()
    If m_beLeaseNo = 0 Then Exit Sub
    On Error Resume Next
    Close #m_beLeaseNo
    m_beLeaseNo = 0
    If Len(m_beLeaseSid) > 0 Then Rdv3ChDeleteQuiet Rdv3ChBeLeasePath(m_beLeaseSid)
    m_beLeaseSid = ""
    On Error GoTo 0
End Sub

' FE side: True while the BE holds its lease lock
Public Function Rdv3ChBeLeaseAlive(ByVal sid As String) As Boolean
    Dim f As Integer
    Dim errNo As Long
    Dim path As String
    path = Rdv3ChBeLeasePath(sid)
    If Not Rdv3ChFlagExists(path) Then Exit Function
    On Error Resume Next
    Err.Clear
    f = FreeFile
    Open path For Binary Access Write Lock Read Write As #f
    errNo = Err.Number
    If errNo = 0 Then Close #f
    Err.Clear
    On Error GoTo 0
    Rdv3ChBeLeaseAlive = (errNo <> 0)
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
