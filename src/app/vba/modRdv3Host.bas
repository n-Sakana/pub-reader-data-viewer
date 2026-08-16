Attribute VB_Name = "modRdv3Host"
'==============================================================================
' modRdv3Host -- the FE side of owning the BE process: start it, know whether
' it is alive, stop it. No Shell, no cscript, no Win32.
'
' HOW THE BE IS STARTED
' Exactly the shape the proven FE/BE project uses (xltoolrack, JobHost.StartJob):
'
'   CreateObject("Excel.Application")     a NEW invisible Excel, ours alone
'   .Workbooks.Open(worker copy, ...)     under AutomationSecurity = low, so
'                                         the unsigned worker macros load
'   .Run "'<book>'!Rdv3BeBootstrap", ...  which ONLY arms Application.OnTime
'                                         and returns; the resident loop starts
'                                         from the BE's own message pump
'
' The bootstrap returning immediately is what keeps this asynchronous: the FE
' is occupied for the process start and the worker open (measured and logged as
' spawn_ms, ~1 s), never for the work itself. v2 moved that occupancy off the
' FE thread by shelling a detached cscript; with Shell gone the spawn is done
' from the FIRST PUMP TICK instead of from Workbook_Open, so the UI is painted
' and the window is up before the FE goes quiet for it.
'
' THE FE STILL HOLDS NO COM REFERENCE TO THE BE. The references are dropped as
' soon as the bootstrap returns, because a COM call into a BE that is busy in
' VBA parks the FE behind Excel's server-busy dialog -- the exact class of
' freeze this architecture exists to avoid. A zero-reference invisible Excel
' would normally shut itself down at the next message pump; the bootstrap has
' already declared Application.UserControl = True (measured fix, see
' docs\app.md), so this instance ends only through its own SelfQuit.
'
' LIVENESS is therefore never a COM call and never a process id: the BE holds
' rdv3_<sid>_be_lease.lock open for its whole life and the FE probes the lock.
' The OS releases it the instant the process ends, however it ended.
'
' STOP is the file protocol: stop flag + fe_gone tombstone, then wait for the
' BE's own quit (its loop checks both about once a second). There is no
' last-resort TerminateProcess any more -- that was the only Win32 left here --
' so a BE that hangs inside a long operation is reported, not killed.
'==============================================================================
Option Explicit

Private m_active As Boolean
Private m_sid As String
Private m_beId As String

' identity of the BE instance, for the log (its own Excel window handle)
Public Function Rdv3HostBeId() As String
    Rdv3HostBeId = m_beId
End Function

Public Function Rdv3HostStarted() As Boolean
    Rdv3HostStarted = m_active
End Function

' True while the BE process holds its lease lock
Public Function Rdv3HostBeAlive() As Boolean
    If Not m_active Then Exit Function
    If Len(m_sid) = 0 Then Exit Function
    Rdv3HostBeAlive = Rdv3ChBeLeaseAlive(m_sid)
End Function

' Start the BE. Synchronous, and short: everything after the bootstrap call
' happens in the BE's own process. spawnMs reports the FE occupancy.
Public Function Rdv3HostSpawn(ByVal sid As String, ByVal dataDir As String, _
                              ByVal ledgerPath As String, ByVal beLogPath As String, _
                              ByVal configPath As String, _
                              ByRef errMsg As String, ByRef spawnMs As Double) As Boolean
    Dim copyPath As String
    Dim beApp As Object
    Dim beBook As Object
    Dim prevSec As Long
    Dim secChanged As Boolean
    Dim t0 As Double

    errMsg = ""
    spawnMs = 0
    Rdv3HostSpawn = False
    m_active = False
    m_beId = ""
    m_sid = sid
    t0 = Rdv3Ticks()
    On Error GoTo Fail

    ' the FE lease must exist before the BE first probes it
    If Not Rdv3ChEnsureLease(sid) Then
        errMsg = "lease file could not be created"
        Exit Function
    End If
    Rdv3ChDeleteQuiet Rdv3ChStopPath(sid)
    Rdv3ChDeleteQuiet Rdv3ChGonePath(sid)
    Rdv3ChDeleteQuiet Rdv3ChAggPath(sid)
    Rdv3ChDeleteQuiet Rdv3ChBeDonePath(sid)
    Rdv3ChDeleteQuiet Rdv3ChBeLeasePath(sid)

    copyPath = Rdv3ChWorkerCopyPath(sid)
    WriteWorkerBook copyPath

    Set beApp = CreateObject("Excel.Application")
    beApp.Visible = False
    beApp.DisplayAlerts = False
    beApp.ScreenUpdating = False
    beApp.EnableEvents = False
    prevSec = beApp.AutomationSecurity
    beApp.AutomationSecurity = 1                  ' msoAutomationSecurityLow
    secChanged = True
    Set beBook = beApp.Workbooks.Open(copyPath, 0, True)
    beApp.AutomationSecurity = prevSec
    secChanged = False
    m_beId = "hwnd:" & CStr(beApp.hwnd)

    ' arms the BE's OnTime and returns; the loop runs in that process
    beApp.Run "'" & Replace(beBook.Name, "'", "''") & "'!Rdv3BeBootstrap", _
        sid, dataDir, ledgerPath, beLogPath, configPath

    ' hand the instance over to itself (UserControl was set by the bootstrap)
    Set beBook = Nothing
    Set beApp = Nothing

    m_active = True
    spawnMs = Rdv3MsSince(t0)
    Rdv3HostSpawn = True
    Exit Function
Fail:
    errMsg = "BE spawn failed " & Err.Number & ": " & Err.Description
    spawnMs = Rdv3MsSince(t0)
    On Error Resume Next
    If Not beApp Is Nothing Then
        If secChanged Then beApp.AutomationSecurity = prevSec
        ' nothing of ours is running in there yet, so Quit is safe here and
        ' only here (before the bootstrap, the instance has no VBA of ours)
        If Not beBook Is Nothing Then beBook.Close False
        beApp.Quit
    End If
    Set beBook = Nothing
    Set beApp = Nothing
    m_active = False
End Function

' Stop: flag + tombstone, then wait for the BE's own quit. forced stays False --
' there is no kill path left; note carries what actually happened.
Public Sub Rdv3HostStop(ByVal sid As String, ByRef forced As Boolean, ByRef note As String)
    Dim i As Long
    forced = False
    note = "be=" & IIf(Len(m_beId) > 0, m_beId, "?")
    Rdv3ChWriteFlag Rdv3ChStopPath(sid), "stop"
    Rdv3ChWriteFlag Rdv3ChGonePath(sid), "closed_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If Not m_active Then
        note = "no active BE"
        Exit Sub
    End If

    ' The BE checks the stop flag about once a second; 12 s is the same grace
    ' the pid-based version allowed before it reached for TerminateProcess.
    ' Application.Wait is the right primitive HERE (and only here): its
    ' one-second granularity is exactly this cadence, and it keeps the FE's
    ' shutdown path independent of the BE's sub-second wait primitive.
    For i = 1 To 12
        If Not Rdv3ChBeLeaseAlive(sid) Then Exit For
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
    If Rdv3ChBeLeaseAlive(sid) Then
        note = note & " still_running (no kill path: reported, not killed)"
    Else
        note = note & " exited"
    End If
    m_active = False
End Sub

' the embedded worker book: base64 in META!E<RDV3_M_WORKER_TOP> downward
Private Sub WriteWorkerBook(ByVal path As String)
    Dim ws As Object
    Dim r As Long
    Dim s As String
    Dim chunk As String
    Dim b() As Byte
    Dim f As Integer

    Set ws = ThisWorkbook.Worksheets(RDV3_SHEET_META)
    r = RDV3_M_WORKER_TOP
    s = ""
    Do
        chunk = CStr(ws.Cells(r, 5).Value)
        If Len(chunk) = 0 Then Exit Do
        s = s & chunk
        r = r + 1
    Loop
    If Len(s) = 0 Then
        Err.Raise vbObjectError + 1001, , "embedded worker book is missing from META"
    End If
    b = B64Decode(s)
    If Dir$(path) <> "" Then Kill path
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f
End Sub

Private Function B64Decode(ByVal s As String) As Byte()
    Dim el As Object
    Set el = CreateObject("MSXML2.DOMDocument.6.0").createElement("b64")
    el.DataType = "bin.base64"
    el.Text = s
    B64Decode = el.nodeTypedValue
End Function
