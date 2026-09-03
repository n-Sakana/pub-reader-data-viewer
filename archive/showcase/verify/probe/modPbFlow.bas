Attribute VB_Name = "modPbFlow"
'==============================================================================
' modPbFlow -- VERIFICATION ONLY. On a 812x812 board of 1px cells, what does a
' small per-tick update really cost, and does the answer change between
'   A: ScreenUpdating False -> write -> True   (the shipped Hold/Release shape)
'   B: write with ScreenUpdating left True     (no toggle at all)
' The 1px build on this machine ticks every ~3.4 s although the VBA body takes
' ~30 ms; the suspicion is that the TOGGLE invalidates the whole viewport
' (659,344 visible cells) while a plain small write would only repaint its own
' rectangle. Measured as OnTime cadence, the same way the real pump runs.
'==============================================================================
Option Explicit

Private m_res As String
Private m_ws As Worksheet
Private m_phase As Long
Private m_i As Long
Private m_last As Double
Private m_uia As Object                        ' UIAutomationClient.IUIAutomation
Private m_root As Object
Private m_np As Object                         ' element array from FindAll
Private m_self As Object                       ' UIA element of THIS Excel window

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

Public Sub PbFwPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbFwArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, Qual("PbFwRun")
End Sub

' Build the 1px board and start the cadence phases.
Public Sub PbFwRun()
    Dim t As Double

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    On Error GoTo Failed
    Application.DisplayAlerts = False
    Application.WindowState = xlMaximized

    Set m_ws = ThisWorkbook.Worksheets.Add()
    m_ws.Activate
    ' display settings BEFORE the columns are narrowed (the 276 s trap)
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    Application.ScreenUpdating = False
    t = Timer
    m_ws.Range(m_ws.Rows(1), m_ws.Rows(812)).RowHeight = 0.6
    m_ws.Range(m_ws.Columns(1), m_ws.Columns(812)).ColumnWidth = 0.03
    Say "grid_ms " & Format$((Timer - t) * 1000, "0") & " rowpt=" & m_ws.Rows(1).Height & " colpt=" & m_ws.Columns(1).Width
    t = Timer
    ' 8 colour bands so the board is not one uniform format run
    Dim b As Long
    For b = 0 To 7
        m_ws.Range(m_ws.Rows(b * 101 + 1), m_ws.Rows(IIf(b = 7, 812, b * 101 + 101))).Interior.Color = _
            RGB(230 - b * 10, 235 - b * 8, 245 - b * 5)
    Next b
    Say "fill_ms " & Format$((Timer - t) * 1000, "0")
    Application.ScreenUpdating = True
    DoEvents
    Say "board up, visible rows " & ActiveWindow.VisibleRange.Rows.Count & _
        " cols " & ActiveWindow.VisibleRange.Columns.Count
    m_phase = 1
    m_i = 0
    m_last = Timer
    Application.OnTime Now + TimeSerial(0, 0, 1), Qual("PbFwTick")
    Exit Sub
Failed:
    Say "RUN FAILED " & Err.Number & " " & Err.Description
    Say "DONE"
End Sub

' One pump tick.
'   1 = Hold/Release, plain cell        2 = direct write, plain cell
'   3 = direct write + dot fill        (then 60 merged label cells are built)
'   4 = Hold/Release, write into a merged cell
'   5 = direct write into a merged cell (then 6 hyperlinks are added)
'   6 = Hold/Release, merged cell, hyperlinks present
Public Sub PbFwTick()
    Dim gap As Double
    Dim t0 As Double
    Dim op As Double

    On Error GoTo Failed
    gap = Timer - m_last
    If gap < 0 Then gap = gap + 86400
    m_last = Timer

    m_i = m_i + 1
    t0 = Timer
    Select Case m_phase
        Case 1
            Application.ScreenUpdating = False
            m_ws.Cells(6, 400).Value = "t" & m_i
            Application.ScreenUpdating = True
        Case 2
            m_ws.Cells(6, 400).Value = "u" & m_i
        Case 3
            m_ws.Cells(6, 400).Value = "v" & m_i
            m_ws.Range(m_ws.Cells(700, 200 + (m_i Mod 8) * 29), _
                       m_ws.Cells(707, 226 + (m_i Mod 8) * 29)).Interior.Color = _
                IIf(m_i Mod 2 = 0, RGB(25, 122, 75), RGB(230, 230, 230))
        Case 4, 6
            Application.ScreenUpdating = False
            m_ws.Range("fw_clock").Cells(1, 1).Value = "w" & m_i
            Application.ScreenUpdating = True
        Case 5
            m_ws.Range("fw_clock").Cells(1, 1).Value = "x" & m_i
        Case 7
            ' complex board, direct small write, no toggle
            m_ws.Range("fw_clock").Cells(1, 1).Value = "y" & m_i
        Case 8
            ' complex board, Hold/Release around the same small write
            Application.ScreenUpdating = False
            m_ws.Range("fw_clock").Cells(1, 1).Value = "z" & m_i
            Application.ScreenUpdating = True
        Case 9
            ' complex board, direct dot fill (the animation op)
            m_ws.Range(m_ws.Cells(700, 200 + (m_i Mod 8) * 29), _
                       m_ws.Cells(707, 226 + (m_i Mod 8) * 29)).Interior.Color = _
                IIf(m_i Mod 2 = 0, RGB(25, 122, 75), RGB(230, 230, 230))
        Case 10
            ' UIA client alive in-process, direct small write
            m_ws.Range("fw_clock").Cells(1, 1).Value = "p" & m_i
        Case 11
            ' UIA client alive, Hold/Release around the same write
            Application.ScreenUpdating = False
            m_ws.Range("fw_clock").Cells(1, 1).Value = "q" & m_i
            Application.ScreenUpdating = True
        Case 12
            ' UIA client alive, a realistic tick: clock + status + 2 dot fills
            m_ws.Range("fw_clock").Cells(1, 1).Value = "r" & m_i
            m_ws.Range("fw_clock").Offset(40, 0).Cells(1, 1).Value = "poll " & m_i
            m_ws.Range(m_ws.Cells(700, 200 + (m_i Mod 8) * 29), _
                       m_ws.Cells(707, 226 + (m_i Mod 8) * 29)).Interior.Color = RGB(25, 122, 75)
            m_ws.Range(m_ws.Cells(700, 200 + ((m_i + 4) Mod 8) * 29), _
                       m_ws.Cells(707, 226 + ((m_i + 4) Mod 8) * 29)).Interior.Color = RGB(230, 230, 230)
        Case 13
            ' window in NORMAL state sized to the board (like the real app),
            ' same direct small write as phase 10
            m_ws.Range("fw_clock").Cells(1, 1).Value = "s" & m_i
        Case 14
            ' normal-state window, realistic 4-write tick
            m_ws.Range("fw_clock").Cells(1, 1).Value = "n" & m_i
            m_ws.Range("fw_clock").Offset(40, 0).Cells(1, 1).Value = "poll " & m_i
            m_ws.Range(m_ws.Cells(700, 200 + (m_i Mod 8) * 29), _
                       m_ws.Cells(707, 226 + (m_i Mod 8) * 29)).Interior.Color = RGB(25, 122, 75)
            m_ws.Range(m_ws.Cells(700, 200 + ((m_i + 4) Mod 8) * 29), _
                       m_ws.Cells(707, 226 + ((m_i + 4) Mod 8) * 29)).Interior.Color = RGB(230, 230, 230)
        Case 15
            ' horizontal scrollbar and sheet tabs HIDDEN, direct small write
            m_ws.Range("fw_clock").Cells(1, 1).Value = "h" & m_i
        Case 16
            ' strips shown again, same write (recovery check)
            m_ws.Range("fw_clock").Cells(1, 1).Value = "g" & m_i
        Case 17
            ' UIA element of THIS Excel window held (like the real app), write
            m_ws.Range("fw_clock").Cells(1, 1).Value = "e" & m_i
        Case 18
            ' self element released again, write (recovery check)
            m_ws.Range("fw_clock").Cells(1, 1).Value = "f" & m_i
    End Select
    op = (Timer - t0) * 1000

    Say "phase " & m_phase & " tick " & m_i & " gap_ms " & Format$(gap * 1000, "0") & _
        " op_ms " & Format$(op, "0")

    If m_i >= 5 Then
        m_phase = m_phase + 1
        m_i = 0
        If m_phase = 4 Then BuildMerges
        If m_phase = 6 Then BuildLinks
        If m_phase = 7 Then BuildComplex
        If m_phase = 10 Then BuildUiaClient
        If m_phase = 13 Then GoNormalWindow
        If m_phase = 15 Then StripBars True
        If m_phase = 16 Then StripBars False
        If m_phase = 17 Then SelfElement True
        If m_phase = 18 Then SelfElement False
        If m_phase > 18 Then
            Say "DONE"
            Exit Sub
        End If
    End If
    Application.OnTime Now + TimeSerial(0, 0, 1), Qual("PbFwTick")
    Exit Sub
Failed:
    Say "TICK FAILED " & Err.Number & " " & Err.Description
    Say "DONE"
End Sub

' 60 merged label cells the way the real screen makes them: merged range,
' Noto Sans JP, small font, one of them named fw_clock (the write target).
Private Sub BuildMerges()
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim rg As Range
    Dim t As Double

    On Error Resume Next
    t = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For i = 0 To 59
        r = 20 + (i \ 6) * 40
        c = 20 + (i Mod 6) * 130
        Set rg = m_ws.Range(m_ws.Cells(r, c), m_ws.Cells(r + 13, c + 99))
        rg.Merge
        rg.Value = "LABEL " & i
        rg.Font.Name = "Noto Sans JP"
        rg.Font.Size = 8
        rg.VerticalAlignment = xlCenter
    Next i
    m_ws.Names.Add "fw_clock", m_ws.Range(m_ws.Cells(20, 20), m_ws.Cells(33, 119))
    Application.ScreenUpdating = True
    DoEvents
    Say "merges built in " & Format$((Timer - t) * 1000, "0") & " ms"
End Sub

' Make the board's format-run count comparable to the real screen: alternate
' the colour of every other column over the whole 812x812 board (about 330k
' runs), the way the real screen's rounded corners and 1px strips shred rows
' into many short runs.
Private Sub BuildComplex()
    Dim c As Long
    Dim t As Double

    On Error Resume Next
    t = Timer
    Application.ScreenUpdating = False
    For c = 1 To 812 Step 2
        m_ws.Range(m_ws.Cells(1, c), m_ws.Cells(812, c)).Interior.Color = RGB(247, 247, 247)
    Next c
    Application.ScreenUpdating = True
    DoEvents
    Say "complex built in " & Format$((Timer - t) * 1000, "0") & " ms"
End Sub

' The real app is itself a UI Automation CLIENT (CUIAutomation instance, root
' element, FindAll, live element references). Loading that machinery into the
' process is the one structural thing this probe board lacked. If per-write
' cost jumps once it is present, the tax belongs to UIA eventing, not to the
' board.
Private Sub BuildUiaClient()
    Dim cond As Object
    Dim t As Double

    On Error Resume Next
    t = Timer
    Set m_uia = New UIAutomationClient.CUIAutomation
    Set m_root = m_uia.GetRootElement
    Set cond = m_uia.CreatePropertyCondition(30012, "Notepad")   ' ClassName
    Set m_np = m_root.FindAll(2, cond)                           ' TreeScope_Children
    Say "uia client up in " & Format$((Timer - t) * 1000, "0") & " ms, notepads " & m_np.Length
End Sub

' The real app runs in a NORMAL-state window sized to the board; every probe
' phase so far ran maximized. If that alone changes the per-write cost, the
' tax is in the window presentation, not in the board or the client objects.
Private Sub GoNormalWindow()
    On Error Resume Next
    Application.WindowState = xlNormal
    Application.Left = 0
    Application.Top = 0
    Application.Width = 530
    Application.Height = 580
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    DoEvents
    Say "window normal " & Application.Width & "x" & Application.Height
End Sub

' The app hides the horizontal scrollbar and the sheet tabs to win back 33
' device px of height. Stripping every board feature off the live board did
' not remove the per-write tax, so the last standing difference to this probe
' is exactly these two window settings.
Private Sub StripBars(ByVal hide As Boolean)
    On Error Resume Next
    ActiveWindow.DisplayHorizontalScrollBar = Not hide
    ActiveWindow.DisplayWorkbookTabs = Not hide
    DoEvents
    Say "bars hidden=" & hide
End Sub

' The real app materialises the UIA element of ITS OWN window
' (ElementFromHandle(Application.hwnd) in MeasurePxPerPt and ScanWindows).
' That is the one UIA act this probe has not copied: it wakes Excel's own
' UIA provider inside the process being measured.
Private Sub SelfElement(ByVal grab As Boolean)
    Dim t As Double
    Dim rc As Variant
    On Error Resume Next
    t = Timer
    If grab Then
        Set m_self = m_uia.ElementFromHandle(ByVal Application.hwnd)
        Say "self element up in " & Format$((Timer - t) * 1000, "0") & " ms nothing=" & (m_self Is Nothing)
    Else
        Set m_self = Nothing
        Say "self element released"
    End If
End Sub

Private Sub BuildLinks()
    Dim i As Long
    Dim rg As Range
    Dim t As Double

    On Error Resume Next
    t = Timer
    Application.ScreenUpdating = False
    For i = 0 To 5
        Set rg = m_ws.Range("fw_clock").Offset(40 * i, 260)
        m_ws.Hyperlinks.Add Anchor:=rg, Address:="", _
            SubAddress:="'" & m_ws.Name & "'!" & rg.Cells(1, 1).Address, ScreenTip:="tip " & i
    Next i
    Application.ScreenUpdating = True
    DoEvents
    Say "links built in " & Format$((Timer - t) * 1000, "0") & " ms"
End Sub
