Attribute VB_Name = "modPbProbe"
'==============================================================================
' modPbProbe -- VERIFICATION ONLY. Not part of the deliverable.
'
' Measures the platform facts the requirement document (v2, section 10) says
' have to be measured before the implementation can choose:
'   - which UI Automation calls are safe from Excel's own UI thread
'     (this repo has a measured deadlock for "walk the desktop")
'   - the real cell pitch for a 1px pseudo-pixel grid on THIS display
'   - whether a Shape can carry OnAction and a hyperlink ScreenTip at once
'   - whether an ActiveX text box can be created at all
'   - what Application.OnTime actually does with a sub-second interval
'
' Every step writes its own START line BEFORE it runs and its result line
' after, so a call that never returns shows up in the log as a START with no
' end. The harness outside kills the process it owns on a deadline and can
' re-run with that step in the skip list (sheet1!A2).
'==============================================================================
Option Explicit

Private Const TS_CHILDREN As Long = 2
Private Const TS_DESCENDANTS As Long = 4
Private Const UIA_ProcessIdPropertyId As Long = 30002
Private Const UIA_NamePropertyId As Long = 30005
Private Const UIA_ClassNamePropertyId As Long = 30012
Private Const UIA_NativeWindowHandlePropertyId As Long = 30020
Private Const UIA_ValuePatternId As Long = 10002
Private Const UIA_TransformPatternId As Long = 10003
Private Const UIA_WindowPatternId As Long = 10009

Private m_res As String
Private m_skip As String
Private m_chain As Long
Private m_chainT0 As Double
Private m_chainLast As Double
Private m_chainStep As Double
Private m_chainTag As String
Private m_uia As UIAutomationClient.IUIAutomation
Private m_root As UIAutomationClient.IUIAutomationElement
Private m_top As UIAutomationClient.IUIAutomationElement

'------------------------------------------------------------------ log helpers
' The watcher outside reads this file while VBA appends to it. A collision is
' a permission error, not a lost line: retry rather than fail.
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

Private Function Beginning(ByVal tag As String) As Boolean
    If InStr(1, "," & m_skip & ",", "," & tag & ",", vbTextCompare) > 0 Then
        Say "SKIP  " & tag
        Beginning = False
        Exit Function
    End If
    Say "START " & tag
    Beginning = True
End Function

Private Sub Done1(ByVal tag As String, ByVal val As String)
    Say "OK    " & tag & " = " & val
End Sub

Private Sub Fail1(ByVal tag As String)
    Dim n As Long
    Dim d As String
    n = Err.Number
    d = Err.Description
    Say "ERR   " & tag & " = " & n & " " & d
End Sub

Private Function Qual(ByVal procName As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

'------------------------------------------------------------------ entry point
' The harness calls this FIRST, in an INVISIBLE Excel. A VBA COMPILE ERROR is a
' modal, and a modal nobody can see is a hang: if this does not answer, the
' harness kills the process it owns and nothing is ever put on the operator's
' screen.
Public Sub PbProbePing()
    ReadArgs
    Say "PONG  project compiles"
End Sub

' Called through Application.Run: only schedules, so the caller's COM call
' returns before anything risky happens.
Public Sub PbProbeArm()
    ReadArgs
    Application.OnTime Now, Qual("PbProbeRun")
End Sub

Private Sub ReadArgs()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    m_skip = CStr(ThisWorkbook.Worksheets(1).Range("A2").Value)
End Sub

Public Sub PbProbeRun()
    ReadArgs
    Say "=== probe start " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " skip=[" & m_skip & "] ==="
    ProbeEnv
    ProbeUia
    ProbeGrid
    ProbeShapeTip
    ProbeActiveX
    StartChain "ontime_half", 0.5
End Sub

'------------------------------------------------------------------ environment
Private Sub ProbeEnv()
    On Error GoTo Failed
    If Not Beginning("env") Then Exit Sub
    Done1 "env", "ver=" & Application.Version & _
        " build=" & Application.Build & _
        " hwnd=" & CStr(Application.hwnd) & _
        " left=" & Application.Left & " top=" & Application.Top & _
        " width=" & Application.Width & " height=" & Application.Height & _
        " usableW=" & Application.UsableWidth & " usableH=" & Application.UsableHeight
    Exit Sub
Failed:
    Fail1 "env"
End Sub

'--------------------------------------------------------------- UI Automation
' Ordered by rising risk, so a hang still leaves the cheaper answers on disk.
Private Sub ProbeUia()
    UiaNew
    UiaRoot
    UiaFocused
    UiaFocusedTop
    UiaTopPatterns
    UiaFromHandleSelf
    UiaDesktopFindFirstNotepad
    UiaDesktopChildren
    UiaDesktopChildrenProps
    UiaHitTest
End Sub

Private Sub UiaNew()
    On Error GoTo Failed
    If Not Beginning("uia_new") Then Exit Sub
    Set m_uia = New UIAutomationClient.CUIAutomation
    Done1 "uia_new", "ok"
    Exit Sub
Failed:
    Fail1 "uia_new"
End Sub

Private Sub UiaRoot()
    Dim rc As UIAutomationClient.tagRECT
    On Error GoTo Failed
    If Not Beginning("uia_root") Then Exit Sub
    Set m_root = m_uia.GetRootElement
    rc = m_root.CurrentBoundingRectangle
    Done1 "uia_root", "rect=" & rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom & _
        " cls=" & m_root.CurrentClassName
    Exit Sub
Failed:
    Fail1 "uia_root"
End Sub

Private Sub UiaFocused()
    Dim el As UIAutomationClient.IUIAutomationElement
    On Error GoTo Failed
    If Not Beginning("uia_focused") Then Exit Sub
    Set el = m_uia.GetFocusedElement
    If el Is Nothing Then
        Done1 "uia_focused", "nothing"
    Else
        Done1 "uia_focused", "cls=" & el.CurrentClassName & " pid=" & el.CurrentProcessId & _
            " type=" & el.CurrentControlType
    End If
    Exit Sub
Failed:
    Fail1 "uia_focused"
End Sub

' climb to the top-level window and read properties OFF IT. The repo's v2 note
' says exactly this hangs when the window is Excel's own.
Private Sub UiaFocusedTop()
    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim par As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    Dim depth As Long

    On Error GoTo Failed
    If Not Beginning("uia_focused_top") Then Exit Sub
    Set walker = m_uia.ControlViewWalker
    Set par = m_uia.GetFocusedElement
    For depth = 1 To 12
        Set el = walker.GetParentElement(par)
        If el Is Nothing Then Exit For
        If m_uia.CompareElements(el, m_root) <> 0 Then Exit For
        Set par = el
    Next depth
    Set m_top = par
    rc = par.CurrentBoundingRectangle
    Done1 "uia_focused_top", "cls=" & par.CurrentClassName & " name=[" & par.CurrentName & "]" & _
        " pid=" & par.CurrentProcessId & _
        " hwnd=" & CStr(par.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)) & _
        " rect=" & rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom
    Exit Sub
Failed:
    Fail1 "uia_focused_top"
End Sub

Private Sub UiaTopPatterns()
    Dim s As String
    On Error GoTo Failed
    If Not Beginning("uia_top_patterns") Then Exit Sub
    If m_top Is Nothing Then
        Done1 "uia_top_patterns", "no top element"
        Exit Sub
    End If
    s = "transform=" & CStr(Not m_top.GetCurrentPattern(UIA_TransformPatternId) Is Nothing)
    s = s & " window=" & CStr(Not m_top.GetCurrentPattern(UIA_WindowPatternId) Is Nothing)
    Done1 "uia_top_patterns", s
    Exit Sub
Failed:
    Fail1 "uia_top_patterns"
End Sub

Private Sub UiaFromHandleSelf()
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    On Error GoTo Failed
    If Not Beginning("uia_from_handle_self") Then Exit Sub
    Set el = m_uia.ElementFromHandle(ByVal Application.hwnd)
    If el Is Nothing Then
        Done1 "uia_from_handle_self", "nothing"
    Else
        rc = el.CurrentBoundingRectangle
        Done1 "uia_from_handle_self", "cls=" & el.CurrentClassName & _
            " name=[" & el.CurrentName & "]" & _
            " rect=" & rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom
    End If
    Exit Sub
Failed:
    Fail1 "uia_from_handle_self"
End Sub

' IUIAutomation.ElementFromPoint takes a POINT BY VALUE and VBA cannot pass a
' user-defined type ByVal -- it is a COMPILE error, not a runtime one, so it
' cannot be caught. (This is the "tagPOINT の扱い" the requirement document
' leaves to implementation time, section 10.) The same answer is built here out
' of calls VBA CAN make: hit-test the point against the top-level windows'
' bounding rectangles, nearest to the front first.
Private Sub UiaHitTest()
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    Dim root As UIAutomationClient.tagRECT
    Dim px As Long
    Dim py As Long
    Dim i As Long

    On Error GoTo Failed
    If Not Beginning("uia_hittest") Then Exit Sub
    If m_top Is Nothing Then
        Done1 "uia_hittest", "no top element"
        Exit Sub
    End If
    root = m_top.CurrentBoundingRectangle
    px = root.Left + 60
    py = root.Top + 60
    Set arr = m_root.FindAll(TS_CHILDREN, m_uia.CreateTrueCondition)
    For i = 0 To arr.Length - 1
        Set el = arr.GetElement(i)
        rc = el.CurrentBoundingRectangle
        If px >= rc.Left And px < rc.Right And py >= rc.Top And py < rc.Bottom Then
            Done1 "uia_hittest", "pt=" & px & "," & py & " zorder=" & i & _
                " cls=" & el.CurrentClassName & " pid=" & el.CurrentProcessId & _
                " name=[" & Left$(el.CurrentName, 40) & "]" & _
                " rect=" & rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom
            Exit Sub
        End If
    Next i
    Done1 "uia_hittest", "pt=" & px & "," & py & " no window contains it"
    Exit Sub
Failed:
    Fail1 "uia_hittest"
End Sub

Private Sub UiaDesktopFindFirstNotepad()
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim el As UIAutomationClient.IUIAutomationElement
    On Error GoTo Failed
    If Not Beginning("uia_findfirst_notepad") Then Exit Sub
    Set cond = m_uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Notepad")
    Set el = m_root.FindFirst(TS_CHILDREN, cond)
    If el Is Nothing Then
        Done1 "uia_findfirst_notepad", "none (no notepad window right now)"
    Else
        Done1 "uia_findfirst_notepad", "name=[" & el.CurrentName & "] pid=" & el.CurrentProcessId
    End If
    Exit Sub
Failed:
    Fail1 "uia_findfirst_notepad"
End Sub

Private Sub UiaDesktopChildren()
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    On Error GoTo Failed
    If Not Beginning("uia_desktop_children") Then Exit Sub
    Set arr = m_root.FindAll(TS_CHILDREN, m_uia.CreateTrueCondition)
    Done1 "uia_desktop_children", "count=" & arr.Length
    Exit Sub
Failed:
    Fail1 "uia_desktop_children"
End Sub

Private Sub UiaDesktopChildrenProps()
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    Dim s As String
    Dim i As Long
    On Error GoTo Failed
    If Not Beginning("uia_desktop_children_props") Then Exit Sub
    Set arr = m_root.FindAll(TS_CHILDREN, m_uia.CreateTrueCondition)
    s = ""
    For i = 0 To arr.Length - 1
        Set el = arr.GetElement(i)
        rc = el.CurrentBoundingRectangle
        s = s & "|" & el.CurrentClassName & "~" & Left$(el.CurrentName, 20) & _
            "~" & rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom
    Next i
    Done1 "uia_desktop_children_props", s
    Exit Sub
Failed:
    Fail1 "uia_desktop_children_props"
End Sub

'------------------------------------------------------- the pseudo-pixel grid
Private Sub ProbeGrid()
    Dim ws As Worksheet
    Dim cands As Variant
    Dim i As Long
    Dim s As String
    Dim cw As Double

    On Error GoTo Failed
    If Not Beginning("grid") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(1)
    ' what row heights the display actually accepts (the device pixel)
    cands = Array(0.1, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.75, 0.8, 1#, 1.2, 1.5)
    s = ""
    For i = LBound(cands) To UBound(cands)
        ws.Rows(1).RowHeight = cands(i)
        s = s & "|rh " & cands(i) & "->" & ws.Rows(1).Height
    Next i
    Done1 "grid_rowheights", s

    ' what column widths the display actually accepts
    s = ""
    cw = 0.01
    Do While cw <= 0.2
        ws.Columns(1).ColumnWidth = cw
        s = s & "|cw " & Format$(cw, "0.00") & "->" & ws.Columns(1).Width
        cw = cw + 0.01
    Loop
    Done1 "grid_colwidths", s

    If Not Beginning("grid812") Then Exit Sub
    Application.ScreenUpdating = False
    ws.Range(ws.Columns(1), ws.Columns(812)).ColumnWidth = 0.08
    ws.Range(ws.Rows(1), ws.Rows(812)).RowHeight = 0.75
    Application.ScreenUpdating = True
    Done1 "grid812", "colW=" & ws.Columns(1).Width & " rowH=" & ws.Rows(1).Height & _
        " left813=" & ws.Cells(1, 813).Left & " top813=" & ws.Cells(813, 1).Top & _
        " zoom=" & ActiveWindow.Zoom
    Exit Sub
Failed:
    Fail1 "grid"
    Application.ScreenUpdating = True
End Sub

'------------------------------------- Shape.OnAction together with a ScreenTip
Private Sub ProbeShapeTip()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim s As String

    On Error GoTo Failed
    If Not Beginning("shapetip") Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(1)
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 10, 400, 90, 20)
    shp.Name = "probeBtn"
    shp.TextFrame2.TextRange.Text = "probe"
    shp.OnAction = "PbProbeClicked"
    s = "after_set=[" & shp.OnAction & "]"
    ws.Hyperlinks.Add Anchor:=shp, Address:="", SubAddress:="'" & ws.Name & "'!A1", _
                      ScreenTip:="probe tooltip"
    s = s & " after_link=[" & shp.OnAction & "]"
    s = s & " tip=[" & shp.Hyperlink.ScreenTip & "]"
    shp.OnAction = "PbProbeClicked"
    s = s & " reset=[" & shp.OnAction & "]"
    Done1 "shapetip", s
    Exit Sub
Failed:
    Fail1 "shapetip"
End Sub

Public Sub PbProbeClicked()
    Say "EVENT shape_clicked t=" & Format$(Timer, "0.000")
End Sub

'---------------------------------------------------------- ActiveX text box
Private Sub ProbeActiveX()
    Dim ws As Worksheet
    Dim ole As OLEObject
    Dim shp As Shape
    Dim s As String

    Set ws = ThisWorkbook.Worksheets(1)

    On Error GoTo Failed1
    If Not Beginning("activex_oleobjects_add") Then GoTo Second
    ws.Activate
    Set ole = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", _
        Left:=10, Top:=430, Width:=200, Height:=60)
    ole.Name = "probeTb"
    ole.Object.MultiLine = True
    ole.Object.Text = "hello" & vbCrLf & "world"
    s = "text=[" & Replace(Replace(ole.Object.Text, vbCr, "\r"), vbLf, "\n") & "]"
    ole.Object.Text = ""
    Done1 "activex_oleobjects_add", s & " cleared=[" & ole.Object.Text & "]"
    GoTo Second
Failed1:
    Fail1 "activex_oleobjects_add"

Second:
    On Error GoTo Failed2
    If Not Beginning("activex_addoleobject") Then GoTo Third
    Set shp = ws.Shapes.AddOLEObject(ClassType:="Forms.TextBox.1", _
        Left:=230, Top:=430, Width:=200, Height:=60)
    Done1 "activex_addoleobject", "name=" & shp.Name
    GoTo Third
Failed2:
    Fail1 "activex_addoleobject"

Third:
    On Error GoTo Failed3
    If Not Beginning("activex_createobject") Then Exit Sub
    Done1 "activex_createobject", TypeName(CreateObject("Forms.TextBox.1"))
    Exit Sub
Failed3:
    Fail1 "activex_createobject"
End Sub

'--------------------------------------------------------- OnTime timing chain
Private Sub StartChain(ByVal tag As String, ByVal stepSec As Double)
    m_chainTag = tag
    m_chainStep = stepSec
    m_chain = 0
    m_chainT0 = Timer
    m_chainLast = m_chainT0
    If Not Beginning(tag) Then
        Say "=== probe done ==="
        Exit Sub
    End If
    Application.OnTime Now + m_chainStep / 86400#, Qual("PbProbeChain")
End Sub

Public Sub PbProbeChain()
    Dim t As Double
    t = Timer
    m_chain = m_chain + 1
    Say "TICK  " & m_chainTag & " n=" & m_chain & _
        " gap=" & Format$(t - m_chainLast, "0.000") & _
        " total=" & Format$(t - m_chainT0, "0.000")
    m_chainLast = t
    If m_chain < 8 Then
        Application.OnTime Now + m_chainStep / 86400#, Qual("PbProbeChain")
        Exit Sub
    End If
    Done1 m_chainTag, "8 ticks in " & Format$(t - m_chainT0, "0.000") & "s"
    If m_chainStep < 0.9 Then
        StartChain "ontime_one", 1#
        Exit Sub
    End If
    Say "=== probe done ==="
End Sub
