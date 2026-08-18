Attribute VB_Name = "modPbPaint"
'==============================================================================
' modPbPaint -- VERIFICATION ONLY. What does it actually cost to paint a
' 812x812 pseudo-pixel canvas out of cells? The screen build turned out to be
' far slower than expected, so every candidate is timed on its own instead of
' being guessed at.
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

Private Function Qual(ByVal p As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & p
End Function

Public Sub PbPtPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbPtArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, Qual("PbPtRun")
End Sub

Public Sub PbPtRun()
    Dim ws As Worksheet
    Dim t As Double
    Dim i As Long
    Dim r As Long, c As Long

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Set ws = ThisWorkbook.Worksheets(1)
    Say "=== paint probe start ==="
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    t = Timer
    ws.Cells.ClearFormats
    Say "clearformats_all      " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 1
    ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08
    Say "set_812_row_col       " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    ws.Cells.Interior.Color = &HF7F7F7
    Say "sheetwide_interior    " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    ws.Range(ws.Cells(1, 1), ws.Cells(812, 812)).Interior.Color = &HF0F0F0
    Say "canvas_812x812        " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    ws.Range(ws.Cells(1, 1), ws.Cells(248, 796)).Interior.Color = &HFFFFFF
    Say "card_796x248          " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    For i = 1 To 300
        ws.Range(ws.Cells(300 + (i Mod 40), 10), ws.Cells(300 + (i Mod 40) + 18, 210)).Interior.Color = &HE0E0E0
    Next i
    Say "300_fills_200x18      " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    For i = 1 To 20
        ws.Range(ws.Cells(500 + i * 3, 10), ws.Cells(500 + i * 3 + 2, 370)).Merge
    Next i
    Say "20_merges_360x3       " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    For i = 1 To 20
        ws.Range(ws.Cells(600 + i * 20, 10), ws.Cells(600 + i * 20 + 16, 370)).Merge
    Next i
    Say "20_merges_360x17      " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    For i = 1 To 5
        ws.Range(ws.Cells(60 + i * 60, 400), ws.Cells(60 + i * 60 + 54, 760)).Merge
    Next i
    Say "5_merges_360x55       " & Format$((Timer - t) * 1000, "0") & " ms"

    Say "about to try Names.Add with a Range object"
    On Error Resume Next
    For i = 1 To 5
        Err.Clear
        t = Timer
        ws.Names.Add Name:="pbr" & i, RefersTo:=ws.Range(ws.Cells(i, 1), ws.Cells(i, 4))
        Say "name_range_obj  #" & i & "   " & Format$((Timer - t) * 1000, "0") & " ms err=" & _
            Err.Number & " " & Err.Description
    Next i
    Say "about to try Names.Add with a RefersTo string"
    For i = 1 To 5
        Err.Clear
        t = Timer
        ws.Names.Add Name:="pbs" & i, RefersTo:="='" & ws.Name & "'!" & _
            ws.Range(ws.Cells(i, 6), ws.Cells(i, 9)).Address
        Say "name_refersto_str #" & i & " " & Format$((Timer - t) * 1000, "0") & " ms err=" & _
            Err.Number & " " & Err.Description
    Next i
    On Error GoTo 0
    t = Timer
    Say "lookup " & ws.Range("pbr1").Address & " " & ws.Range("pbs1").Address & _
        "  " & Format$((Timer - t) * 1000, "0") & " ms"

    t = Timer
    For r = 0 To 9
        For c = 0 To 9
            ws.Range(ws.Cells(700 + r * 10, 400 + c * 10), ws.Cells(700 + r * 10 + 8, 400 + c * 10 + 8)).Interior.Color = &H8080FF
        Next c
    Next r
    Say "matrix_100_marks      " & Format$((Timer - t) * 1000, "0") & " ms"

    Application.ScreenUpdating = True
    Say "=== paint probe done ==="
End Sub
