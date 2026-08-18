Attribute VB_Name = "modPbCols2"
'==============================================================================
' modPbCols2 -- VERIFICATION ONLY, and it must run from a SHELL-started Excel.
'
' Measured: with the user's add-ins loaded (5 add-ins, 3 COM add-ins), setting
' 812 column widths costs 26.7 s and selecting one cell costs 2.4 s. The same
' calls cost 0 ms in an automation Excel, which loads no add-ins. So the screen
' build has to stop doing those things -- this measures the ways around them.
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
    Say Left$(tag & String$(34, " "), 34) & Format$(ms, "0") & " ms"
End Sub

Private Function Qual(ByVal p As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & p
End Function

Public Sub Auto_Open()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "AUTO_OPEN addins=" & Application.AddIns.Count & " com=" & Application.COMAddIns.Count
    ListAddIns
    Application.OnTime Now, Qual("PbC2Run")
End Sub

Public Sub PbC2Ping()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Private Sub ListAddIns()
    Dim i As Long
    On Error Resume Next
    For i = 1 To Application.AddIns.Count
        If Application.AddIns(i).Installed Then
            Say "  addin: " & Application.AddIns(i).Name
        End If
    Next i
    For i = 1 To Application.COMAddIns.Count
        If Application.COMAddIns(i).Connect Then
            Say "  comaddin: " & Application.COMAddIns(i).progID & " | " & Application.COMAddIns(i).Description
        End If
    Next i
End Sub

Private Function NewSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = nm
    Set NewSheet = ws
End Function

Public Sub PbC2Run()
    Dim ws As Worksheet
    Dim front As Worksheet
    Dim t0 As Double
    Dim i As Long

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.DisplayAlerts = False
    Set front = ThisWorkbook.Worksheets(1)
    Say "=== cols2 probe start (state=" & Application.WindowState & ") ==="

    ' A: StandardWidth instead of a column range, sheet in front, maximized
    Application.WindowState = xlMaximized
    Set ws = NewSheet("s1")
    ws.Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.StandardWidth = 0.08: T "A standardwidth front", t0
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "A rows front", t0
    Application.ScreenUpdating = True

    ' B: everything while ANOTHER sheet is in front, then activate once
    Set ws = NewSheet("s2")
    front.Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08: T "B cols behind", t0
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "B rows behind", t0
    t0 = Timer: ws.Range(ws.Cells(1, 1), ws.Cells(812, 812)).Interior.Color = &HF0F0F0: T "B canvas behind", t0
    Application.ScreenUpdating = True
    t0 = Timer: ws.Activate: T "B activate after", t0

    ' C: painting and selecting once the sheet IS in front (the tick cost)
    t0 = Timer: ws.Range(ws.Cells(100, 100), ws.Cells(251, 467)).Interior.Color = &HD0D0FF: T "C region 368x152", t0
    t0 = Timer
    For i = 1 To 30
        ws.Range(ws.Cells(300 + i, 100), ws.Cells(340 + i, 300)).Interior.Color = &HC0FFC0
    Next i
    T "C 30 fills", t0
    t0 = Timer: ws.Cells(812, 812).Select: T "C select corner", t0
    t0 = Timer: ws.Cells(1, 1).Select: T "C select A1", t0
    t0 = Timer: ws.Range(ws.Cells(400, 20), ws.Cells(420, 200)).Merge: T "C merge", t0
    t0 = Timer: ws.Cells(400, 20).Value = "hello": T "C set value", t0
    t0 = Timer: Say "C read selection " & Selection.Address: T "C read selection", t0

    ' D: how much of it is the WINDOW being huge
    Application.WindowState = xlNormal
    Application.Width = 700
    Application.Height = 600
    t0 = Timer: ws.Cells(812, 812).Select: T "D select in small window", t0
    t0 = Timer: ws.Range(ws.Cells(500, 100), ws.Cells(651, 467)).Interior.Color = &HFFD0D0: T "D region small window", t0

    Say "=== cols2 probe done ==="
End Sub
