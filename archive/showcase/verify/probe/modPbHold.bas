Attribute VB_Name = "modPbHold"
'==============================================================================
' modPbHold -- VERIFICATION ONLY, opened by the SHELL so the add-ins are there.
'
' The previous probe showed the cost is not the cells, it is the REPAINT that
' each separate paint triggers: 30 fills cost 16 s live and 8 ms when drawing
' is held off and released once. This asks the question that decides whether
' "1 cell = 1 px" can stay: at 812 x 812, with drawing held, what does a whole
' screen build cost, and what does one tick's worth of painting cost?
'==============================================================================
Option Explicit

Private m_res As String

Private Sub Say(ByVal s As String)
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
    For i = 1 To 60
        Err.Clear
        f = FreeFile
        Open m_res For Append As #f
        If Err.Number = 0 Then
            Print #f, s
            Close #f
            Exit For
        End If
        DoEvents
    Next i
End Sub

Private Sub T(ByVal tag As String, ByVal t0 As Double)
    Dim ms As Double
    ms = (Timer - t0) * 1000
    If ms < 0 Then ms = ms + 86400000#
    Say Left$(tag & String$(32, " "), 32) & Format$(ms, "0") & " ms"
End Sub

Public Sub Auto_Open()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "AUTO_OPEN addins=" & Application.AddIns.Count & " com=" & Application.COMAddIns.Count
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PbHoRun"
End Sub

Public Sub PbHoPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbHoRun()
    Dim ws As Worksheet
    Dim home As Worksheet
    Dim t0 As Double
    Dim i As Long
    Dim n As Long

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.DisplayAlerts = False
    Say "=== hold probe start ==="
    n = 812

    Set home = ThisWorkbook.Worksheets(1)
    home.Activate
    On Error Resume Next
    ThisWorkbook.Worksheets("hold").Delete
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "hold"

    Application.ScreenUpdating = False
    t0 = Timer
    ws.Range(ws.Rows(1), ws.Rows(n)).RowHeight = 1
    T "rows behind held", t0
    t0 = Timer
    ws.Range(ws.Columns(1), ws.Columns(n)).ColumnWidth = 0.08
    T "cols behind held", t0
    Application.ScreenUpdating = True

    ' a whole screen build: canvas plus ~500 fills, all with drawing held
    t0 = Timer
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(1, 1), ws.Cells(n, n)).Interior.Color = &HF7F7F7
    For i = 1 To 500
        ws.Range(ws.Cells(20 + (i Mod 700), 20), ws.Cells(20 + (i Mod 700) + 8, 400)).Interior.Color = &HE8ECFB
    Next i
    Application.ScreenUpdating = True
    T "build 500 fills held (behind)", t0

    ' now show it -- this is where Excel died before
    t0 = Timer
    ws.Activate
    T "ACTIVATE 812x812", t0

    ' one tick's worth of painting, held
    t0 = Timer
    Application.ScreenUpdating = False
    For i = 1 To 30
        ws.Range(ws.Cells(300 + i, 100), ws.Cells(340 + i, 300)).Interior.Color = &HC8FFC8
    Next i
    Application.ScreenUpdating = True
    T "tick 30 fills held (front)", t0

    ' a second tick, to see whether it stays cheap
    t0 = Timer
    Application.ScreenUpdating = False
    For i = 1 To 30
        ws.Range(ws.Cells(400 + i, 100), ws.Cells(440 + i, 300)).Interior.Color = &HFFC8C8
    Next i
    Application.ScreenUpdating = True
    T "tick again 30 fills held", t0

    ' a full canvas repaint with drawing held, in front (the theme switch)
    t0 = Timer
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(1, 1), ws.Cells(n, n)).Interior.Color = &HEFEFEF
    Application.ScreenUpdating = True
    T "full canvas held (front)", t0

    ' merged labels, held
    t0 = Timer
    Application.ScreenUpdating = False
    For i = 1 To 40
        ws.Range(ws.Cells(600 + i * 4, 10), ws.Cells(600 + i * 4 + 3, 200)).Merge
    Next i
    Application.ScreenUpdating = True
    T "40 merges held (front)", t0

    home.Activate
    Say "=== hold probe done ==="
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    Say "ERR " & Err.Number & " " & Err.Description
    Say "=== hold probe done ==="
End Sub
