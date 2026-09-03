Attribute VB_Name = "modPbRender"
'==============================================================================
' modPbRender -- VERIFICATION ONLY, and it has to run in a VISIBLE Excel.
'
' The first paint probe ran invisible and reported 0 ms for everything, which
' was useless: the real screen is 812x812 one-pixel cells in a window that is
' actually being drawn. This measures the same operations where they are
' expensive -- on a sheet that IS on screen -- and measures the way out:
' doing the same work while another sheet is in front.
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
    Say Left$(tag & String$(30, " "), 30) & Format$(ms, "0") & " ms"
End Sub

Private Function Qual(ByVal p As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & p
End Function

Public Sub PbRnPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbRnArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, Qual("PbRnRun")
End Sub

Public Sub PbRnRun()
    Dim wsA As Worksheet
    Dim wsB As Worksheet
    Dim t0 As Double
    Dim i As Long

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "=== render probe start (visible=" & Application.Visible & ") ==="

    Set wsA = ThisWorkbook.Worksheets(1)
    On Error Resume Next
    Set wsB = ThisWorkbook.Worksheets("scratch")
    On Error GoTo 0
    If wsB Is Nothing Then
        Set wsB = ThisWorkbook.Worksheets.Add(After:=wsA)
        wsB.Name = "scratch"
    End If

    '--- A: the sheet IS in front (this is what the app did)
    wsA.Activate
    Application.ScreenUpdating = False
    t0 = Timer
    wsA.Range(wsA.Rows(1), wsA.Rows(812)).RowHeight = 1
    T "front_812_rowheight", t0
    t0 = Timer
    wsA.Range(wsA.Columns(1), wsA.Columns(812)).ColumnWidth = 0.08
    T "front_812_colwidth", t0
    t0 = Timer
    wsA.Range(wsA.Cells(1, 1), wsA.Cells(812, 812)).Interior.Color = &HF0F0F0
    T "front_canvas_fill_noupd", t0
    t0 = Timer
    Application.ScreenUpdating = True
    T "front_screenupdating_on", t0
    t0 = Timer
    wsA.Range(wsA.Cells(100, 100), wsA.Cells(251, 467)).Interior.Color = &HD0D0FF
    T "front_region_368x152", t0
    t0 = Timer
    For i = 1 To 20
        wsA.Range(wsA.Cells(400 + i, 100), wsA.Cells(400 + i, 300)).Interior.Color = &HFFD0D0
    Next i
    T "front_20_thin_fills", t0
    t0 = Timer
    wsA.Cells(812, 812).Select
    T "front_select_corner", t0

    '--- B: the same work while ANOTHER sheet is in front
    wsB.Activate
    t0 = Timer
    wsA.Range(wsA.Rows(1), wsA.Rows(812)).RowHeight = 2
    T "back_812_rowheight", t0
    t0 = Timer
    wsA.Range(wsA.Rows(1), wsA.Rows(812)).RowHeight = 1
    T "back_812_rowheight_2", t0
    t0 = Timer
    wsA.Range(wsA.Cells(1, 1), wsA.Cells(812, 812)).Interior.Color = &HE8E8E8
    T "back_canvas_fill", t0
    t0 = Timer
    wsA.Activate
    T "back_then_activate", t0

    '--- C: what a tick would cost, with the sheet in front
    t0 = Timer
    For i = 1 To 30
        wsA.Range(wsA.Cells(100 + i, 100), wsA.Cells(140 + i, 300)).Interior.Color = &HC0FFC0
    Next i
    T "tick_30_fills_200x40", t0
    t0 = Timer
    wsA.Range("A1").Value = "x" & Timer
    T "tick_one_cell_value", t0

    '--- D: does zooming out change anything (fewer device pixels to draw)
    t0 = Timer
    ActiveWindow.Zoom = 50
    T "zoom_50", t0
    t0 = Timer
    wsA.Range(wsA.Cells(300, 100), wsA.Cells(451, 467)).Interior.Color = &HD0FFFF
    T "zoom50_region_368x152", t0
    ActiveWindow.Zoom = 100

    Say "=== render probe done ==="
End Sub
