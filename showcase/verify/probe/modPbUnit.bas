Attribute VB_Name = "modPbUnit"
'==============================================================================
' modPbUnit -- VERIFICATION ONLY, and it must be opened by the SHELL so the
' user's add-ins are loaded, because that is the condition where this hurts.
'
' The question it answers is the one that decides the design: is the cost of a
' pseudo-pixel canvas driven by the NUMBER OF CELLS on screen? The screen is
' 812 x 812 design pixels either way; what changes is how many cells that is.
'   1 cell = 1 px  ->  812 x 812  = 659,344 cells
'   1 cell = 2 px  ->  406 x 406  = 164,836 cells
'   1 cell = 4 px  ->  203 x 203  =  41,209 cells
' Each sheet is built while it is NOT in front, then shown, and the showing is
' what gets timed -- that is the step that hung and then crashed Excel.
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

Public Sub Auto_Open()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "AUTO_OPEN addins=" & Application.AddIns.Count & " com=" & Application.COMAddIns.Count
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PbUnRun"
End Sub

Public Sub PbUnPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbUnRun()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.DisplayAlerts = False
    Say "=== unit probe start ==="
    OneUnit 2, 406, 2#
    OneUnit 4, 203, 4#
    Say "=== unit probe done ==="
End Sub

' unitPx: how many design pixels one cell is
' n:      how many cells that makes for a 812 px side
' ptPer:  the height of one cell in points (1 design px = 1 pt here, measured
'         on this display: one device pixel is 0.5 pt and the chosen pixel is 1 pt)
Private Sub OneUnit(ByVal unitPx As Long, ByVal n As Long, ByVal ptPer As Double)
    Dim ws As Worksheet
    Dim home As Worksheet
    Dim t0 As Double
    Dim i As Long

    On Error GoTo Failed
    Set home = ThisWorkbook.Worksheets(1)
    home.Activate
    On Error Resume Next
    ThisWorkbook.Worksheets("u" & unitPx).Delete
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = "u" & unitPx

    ' build it while it is NOT in front
    t0 = Timer
    ws.Range(ws.Rows(1), ws.Rows(n)).RowHeight = ptPer
    T "unit" & unitPx & " rows behind", t0
    t0 = Timer
    ws.Range(ws.Columns(1), ws.Columns(n)).ColumnWidth = ColWidthFor(ws, ptPer)
    T "unit" & unitPx & " cols behind", t0
    t0 = Timer
    ws.Range(ws.Cells(1, 1), ws.Cells(n, n)).Interior.Color = &HF7F7F7
    T "unit" & unitPx & " canvas behind", t0

    ' THE STEP THAT MATTERS: put it in front
    t0 = Timer
    ws.Activate
    T "unit" & unitPx & " ACTIVATE", t0

    ' and what a repaint costs once it is in front
    t0 = Timer
    ws.Range(ws.Cells(1, 1), ws.Cells(n, n)).Interior.Color = &HEFEFEF
    T "unit" & unitPx & " repaint in front", t0
    t0 = Timer
    For i = 1 To 30
        ws.Range(ws.Cells(10 + i, 10), ws.Cells(30 + i, 120)).Interior.Color = &HD8E0FF
    Next i
    T "unit" & unitPx & " 30 fills live", t0

    ' the same 30 fills with drawing held off and released once -- this is what
    ' a tick would actually do
    t0 = Timer
    Application.ScreenUpdating = False
    For i = 1 To 30
        ws.Range(ws.Cells(60 + i, 10), ws.Cells(80 + i, 120)).Interior.Color = &HFFE0D8
    Next i
    Application.ScreenUpdating = True
    T "unit" & unitPx & " 30 fills held", t0
    home.Activate
    Exit Sub
Failed:
    Say "ERR unit" & unitPx & " " & Err.Number & " " & Err.Description
End Sub

Private Function ColWidthFor(ByVal ws As Worksheet, ByVal wantPt As Double) As Double
    Dim cw As Double
    ColWidthFor = 0.08
    cw = 0.01
    Do While cw <= 2#
        ws.Columns(1).ColumnWidth = cw
        If ws.Columns(1).Width >= wantPt - 0.001 Then
            ColWidthFor = cw
            Exit Function
        End If
        cw = cw + 0.01
    Loop
End Function
