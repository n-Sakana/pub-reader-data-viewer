Attribute VB_Name = "modPbTx"
'==============================================================================
' modPbTx -- VERIFICATION ONLY. On this machine the app's MoveWin gets
' Nothing from GetCurrentPattern(TransformPattern) for a Notepad top window,
' while the managed .NET UIA client reports CanMove=True on the same window.
' Compare the two native clients (CUIAutomation vs CUIAutomation8) and the
' property route (IsTransformPatternAvailable) to find a working move path.
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

Public Sub PbTxPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbTxArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PbTxRun"
End Sub

Public Sub PbTxRun()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    On Error Resume Next
    TryClient "CUIAutomation", New UIAutomationClient.CUIAutomation
    TryClient "CUIAutomation8", New UIAutomationClient.CUIAutomation8
    ScanDescendants
    Say "DONE"
End Sub

' does ANY element under the notepad window expose TransformPattern?
Private Sub ScanDescendants()
    Dim uia As UIAutomationClient.IUIAutomation
    Dim root As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim win As UIAutomationClient.IUIAutomationElement
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim i As Long
    Dim n As Long

    On Error GoTo Failed
    Set uia = New UIAutomationClient.CUIAutomation
    Set root = uia.GetRootElement
    Set cond = uia.CreatePropertyCondition(30012, "Notepad")
    Set win = root.FindFirst(2, cond)
    If win Is Nothing Then
        Say "scan: no notepad"
        Exit Sub
    End If
    ' 30025 = IsTransformPatternAvailable
    Set cond = uia.CreatePropertyCondition(30025, True)
    Set arr = win.FindAll(4, cond)          ' TreeScope_Descendants
    Say "scan: descendants with Transform = " & arr.Length
    n = arr.Length
    If n > 6 Then n = 6
    For i = 0 To n - 1
        Set el = arr.GetElement(i)
        Say "  [" & i & "] type=" & el.CurrentControlType & " name=" & Left$(el.CurrentName, 40) & _
            " class=" & el.CurrentClassName
    Next i
    ' and the FE Excel window itself, for contrast
    Set cond = uia.CreatePropertyCondition(30012, "XLMAIN")
    Set win = root.FindFirst(2, cond)
    If Not win Is Nothing Then
        Say "scan: XLMAIN IsTransformAvailable=" & CStr(win.GetCurrentPropertyValue(30025))
    End If
    Exit Sub
Failed:
    Say "scan: FAILED " & Err.Number & " " & Err.Description
End Sub

Private Sub TryClient(ByVal label As String, ByVal uia As UIAutomationClient.IUIAutomation)
    Dim root As UIAutomationClient.IUIAutomationElement
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim tp As UIAutomationClient.IUIAutomationTransformPattern
    Dim wp As UIAutomationClient.IUIAutomationWindowPattern
    Dim avail As Variant
    Dim rc As UIAutomationClient.tagRECT

    On Error GoTo Failed
    If uia Is Nothing Then
        Say label & ": client is Nothing"
        Exit Sub
    End If
    Set root = uia.GetRootElement
    Set cond = uia.CreatePropertyCondition(30012, "Notepad")
    Set arr = root.FindAll(2, cond)
    If arr.Length = 0 Then
        Say label & ": no notepad window"
        Exit Sub
    End If
    Set el = arr.GetElement(0)
    avail = el.GetCurrentPropertyValue(30025)          ' IsTransformPatternAvailable
    Say label & ": IsTransformAvailable=" & CStr(avail)
    Err.Clear
    Set tp = el.GetCurrentPattern(10003)
    Say label & ": GetCurrentPattern(Transform) nothing=" & (tp Is Nothing) & " err=" & Err.Number
    If Not tp Is Nothing Then
        Say label & ": CanMove=" & tp.CurrentCanMove
        rc = el.CurrentBoundingRectangle
        tp.Move CDbl(rc.Left + 8), CDbl(rc.Top)
        Say label & ": moved err=" & Err.Number
        tp.Move CDbl(rc.Left), CDbl(rc.Top)
    End If
    Err.Clear
    Set wp = el.GetCurrentPattern(10009)
    Say label & ": GetCurrentPattern(Window) nothing=" & (wp Is Nothing) & " err=" & Err.Number
    Exit Sub
Failed:
    Say label & ": FAILED " & Err.Number & " " & Err.Description
End Sub
