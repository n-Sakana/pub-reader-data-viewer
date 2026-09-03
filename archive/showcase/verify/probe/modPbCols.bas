Attribute VB_Name = "modPbCols"
'==============================================================================
' modPbCols -- VERIFICATION ONLY. Setting 812 column widths took 27 seconds in
' the real screen build and 0 ms in the first probe. This finds the condition
' that makes the difference instead of guessing at it: window state, screen
' updating, and whether the sheet is the one in front.
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

' シェルから開かれたとき（＝アドインが載っている、ふつうの起動）に走らせる口
Public Sub Auto_Open()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "AUTO_OPEN addins=" & Application.AddIns.Count & _
        " comaddins=" & Application.COMAddIns.Count & _
        " books=" & Application.Workbooks.Count
    PbCoArm
End Sub

Public Sub PbCoPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbCoArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, Qual("PbCoRun")
End Sub

Private Function NewSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Worksheets(nm).Delete
    On Error GoTo 0
    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = nm
    Set NewSheet = ws
End Function

Public Sub PbCoRun()
    Dim ws As Worksheet
    Dim t0 As Double

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.DisplayAlerts = False
    Say "=== cols probe start ==="
    Say "window state at entry: " & Application.WindowState & _
        "  appW=" & Application.Width & " appH=" & Application.Height

    '--- 1: normal window, screen updating off (what the first probe did)
    Application.WindowState = xlNormal
    Set ws = NewSheet("c1")
    ws.Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "normal_upd_off rows", t0
    t0 = Timer: ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08: T "normal_upd_off cols", t0
    Application.ScreenUpdating = True

    '--- 2: maximized window, screen updating off (what the app did)
    Application.WindowState = xlMaximized
    Set ws = NewSheet("c2")
    ws.Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "max_upd_off rows", t0
    t0 = Timer: ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08: T "max_upd_off cols", t0
    t0 = Timer: ws.Columns(1).ColumnWidth = 0.09: T "max_upd_off one col", t0
    t0 = Timer: ws.Cells(812, 812).Select: T "max_upd_off select corner", t0
    Application.ScreenUpdating = True

    '--- 3: maximized, screen updating ON
    Set ws = NewSheet("c3")
    ws.Activate
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "max_upd_on rows", t0
    t0 = Timer: ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08: T "max_upd_on cols", t0

    '--- 4: maximized, sheet NOT in front
    Set ws = NewSheet("c4")
    ThisWorkbook.Worksheets("c3").Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "max_back rows", t0
    t0 = Timer: ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08: T "max_back cols", t0
    t0 = Timer: ws.Activate: T "max_back then activate", t0
    Application.ScreenUpdating = True

    '--- 5: does the ORDER matter (columns before rows)
    Set ws = NewSheet("c5")
    ws.Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08: T "cols_first cols", t0
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "cols_first rows", t0
    Application.ScreenUpdating = True

    '--- 6: StandardWidth for the whole sheet instead of a column range
    Set ws = NewSheet("c6")
    ws.Activate
    Application.ScreenUpdating = False
    t0 = Timer: ws.StandardWidth = 0.08: T "standardwidth", t0
    t0 = Timer: ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1: T "standardwidth rows", t0
    Application.ScreenUpdating = True

    Say "=== cols probe done ==="
End Sub
