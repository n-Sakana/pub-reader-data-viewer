Attribute VB_Name = "modRdv3Host"
'==============================================================================
' modRdv3Host -- spawn and stop the BE process, WITHOUT ever blocking the FE.
'
' v1 spawned the BE with a synchronous CreateObject + Workbooks.Open on the
' FE's own thread: ~2 s of measured FE occupancy at boot. v2 hands the whole
' job to a detached cscript (spawn.vbs, generated per session): Shell returns
' in milliseconds, the script creates the invisible Excel, opens the worker
' copy and runs the bootstrap. The FE learns the outcome purely from files:
'   rdv3_<sid>_be_pid.flag     written by the BE loop itself (its real pid)
'   rdv3_<sid>_spawn_err.flag  written by the script when any step fails
' and the pump raises a loud timeout if neither appears in RDV3_SPAWN_TIMEOUT_S.
'
' The FE therefore holds NO COM reference to the BE at any point. The script
' holds its refs only until the BE loop has started (an idle ref-free
' invisible Excel would shut itself down; a running loop keeps it alive), then
' polls the pid flag and exits.
'
' Stop stays file-based: stop flag + fe_gone tombstone, wait for the BE's own
' self-quit by pid, and only as a last resort terminate OUR OWN child pid.
'==============================================================================
Option Explicit

Private m_bePid As Long
Private m_active As Boolean

Public Function Rdv3HostBePid() As Long
    Rdv3HostBePid = m_bePid
End Function

Public Function Rdv3HostBeAlive() As Boolean
    Rdv3HostBeAlive = False
    If Not m_active Then Exit Function
    If m_bePid = 0 Then Exit Function
    Rdv3HostBeAlive = Rdv3PidAlive(m_bePid)
End Function

' the FE pump calls this when the be_pid flag appears
Public Sub Rdv3HostNotePid(ByVal pid As Long)
    m_bePid = pid
    m_active = True
End Sub

' asynchronous spawn: prepare the session files, then Shell a detached script.
' Returns as soon as the script is launched (milliseconds); success/failure
' arrives through the channel flags.
Public Function Rdv3HostSpawnAsync(ByVal sid As String, ByVal dataDir As String, _
                                   ByVal ledgerPath As String, ByVal beLogPath As String, _
                                   ByRef errMsg As String) As Boolean
    Dim copyPath As String
    Dim vbsPath As String
    Dim taskId As Double
    errMsg = ""
    Rdv3HostSpawnAsync = False
    On Error GoTo Fail

    m_bePid = 0
    m_active = False

    ' the lease must exist before the BE first probes it
    If Not Rdv3ChEnsureLease(sid) Then
        errMsg = "lease file could not be created"
        Exit Function
    End If
    Rdv3ChDeleteQuiet Rdv3ChStopPath(sid)
    Rdv3ChDeleteQuiet Rdv3ChGonePath(sid)
    Rdv3ChDeleteQuiet Rdv3ChAggPath(sid)
    Rdv3ChDeleteQuiet Rdv3ChBeDonePath(sid)
    Rdv3ChDeleteQuiet Rdv3ChBePidPath(sid)
    Rdv3ChDeleteQuiet Rdv3ChSpawnErrPath(sid)

    copyPath = Rdv3ChWorkerCopyPath(sid)
    WriteWorkerBook copyPath

    vbsPath = Rdv3ChSpawnVbsPath(sid)
    WriteSpawnScript vbsPath, copyPath, sid, dataDir, ledgerPath, beLogPath

    taskId = Shell("cscript.exe //nologo """ & vbsPath & """", vbHide)
    If taskId = 0 Then
        errMsg = "cscript did not start"
        Exit Function
    End If
    Rdv3HostSpawnAsync = True
    Exit Function
Fail:
    errMsg = "BE spawn failed " & Err.Number & ": " & Err.Description
End Function

' Normal stop: flag + tombstone, then wait for the BE's own Quit. The PID
' terminate is a last resort, on OUR child only (the pid the BE reported for
' itself), and a forced end is reported to the caller for the log.
Public Sub Rdv3HostStop(ByVal sid As String, ByRef forced As Boolean, ByRef note As String)
    Dim i As Long
    forced = False
    note = "pid=" & m_bePid
    If Not m_active Or m_bePid = 0 Then
        note = "no active BE"
        Rdv3ChWriteFlag Rdv3ChStopPath(sid), "stop"
        Rdv3ChWriteFlag Rdv3ChGonePath(sid), "closed_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")
        m_active = False
        Exit Sub
    End If
    Rdv3ChWriteFlag Rdv3ChStopPath(sid), "stop"
    Rdv3ChWriteFlag Rdv3ChGonePath(sid), "closed_at=" & Format$(Now, "yyyy-mm-dd hh:nn:ss")

    For i = 1 To 30                        ' the BE polls liveness about once a second
        If Not Rdv3PidAlive(m_bePid) Then Exit For
        Rdv3Sleep 400
    Next i
    If Rdv3PidAlive(m_bePid) Then
        forced = Rdv3KillPid(m_bePid)
        note = note & " terminated=" & CStr(forced)
    Else
        note = note & " exited"
    End If
    m_active = False
End Sub

' the spawn script, generated with literal values (no argument quoting games).
' Any failure writes the spawn_err flag and tries to quit the half-made Excel.
' After Run returns the script waits for the BE loop's own pid flag before
' releasing its references -- see the module header for why.
Private Sub WriteSpawnScript(ByVal vbsPath As String, ByVal copyPath As String, _
                             ByVal sid As String, ByVal dataDir As String, _
                             ByVal ledgerPath As String, ByVal beLogPath As String)
    Dim f As Integer
    Dim q As String
    q = """"
    If Dir$(vbsPath) <> "" Then Kill vbsPath
    f = FreeFile
    Open vbsPath For Output As #f
    Print #f, "On Error Resume Next"
    Print #f, "Dim xl, wb, fso, i, ok"
    Print #f, "Set fso = CreateObject(" & q & "Scripting.FileSystemObject" & q & ")"
    Print #f, "errPath = " & VbsStr(Rdv3ChSpawnErrPath(sid))
    Print #f, "pidPath = " & VbsStr(Rdv3ChBePidPath(sid))
    Print #f, "Sub Fail(msg)"
    Print #f, "  On Error Resume Next"
    Print #f, "  Dim ts"
    Print #f, "  Set ts = fso.CreateTextFile(errPath, True)"
    Print #f, "  ts.WriteLine msg"
    Print #f, "  ts.Close"
    Print #f, "  If IsObject(xl) Then xl.Quit"
    Print #f, "  WScript.Quit 1"
    Print #f, "End Sub"
    Print #f, "Set xl = CreateObject(" & q & "Excel.Application" & q & ")"
    Print #f, "If Not IsObject(xl) Then Fail " & q & "CreateObject failed: " & q & " & Err.Description"
    Print #f, "xl.Visible = False"
    Print #f, "xl.DisplayAlerts = False"
    Print #f, "xl.ScreenUpdating = False"
    Print #f, "Err.Clear"
    Print #f, "xl.AutomationSecurity = 1"       ' this child, this open only
    Print #f, "Err.Clear"
    Print #f, "xl.EnableEvents = False"
    Print #f, "Set wb = xl.Workbooks.Open(" & VbsStr(copyPath) & ", 0, True)"
    Print #f, "If Not IsObject(wb) Then Fail " & q & "worker open failed: " & q & " & Err.Description"
    Print #f, "Err.Clear"
    Print #f, "xl.Run " & q & "'" & q & " & wb.Name & " & q & "'!Rdv3BeBootstrap" & q & ", " & _
              VbsStr(sid) & ", " & VbsStr(dataDir) & ", " & VbsStr(ledgerPath) & ", " & VbsStr(beLogPath)
    Print #f, "If Err.Number <> 0 Then Fail " & q & "bootstrap failed: " & q & " & Err.Description"
    ' hold our refs until the BE loop reports its pid (so the idle Excel
    ' cannot shut itself down in the gap), then let go
    Print #f, "ok = False"
    Print #f, "For i = 1 To 100"
    Print #f, "  If fso.FileExists(pidPath) Then ok = True : Exit For"
    Print #f, "  WScript.Sleep 200"
    Print #f, "Next"
    Print #f, "If Not ok Then Fail " & q & "BE loop did not start (no pid flag)" & q
    Print #f, "Set wb = Nothing"
    Print #f, "Set xl = Nothing"
    Close #f
End Sub

' a VBS string literal ("" doubling)
Private Function VbsStr(ByVal s As String) As String
    VbsStr = """" & Replace(s, """", """""") & """"
End Function

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
