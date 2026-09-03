Attribute VB_Name = "modPbAx"
'==============================================================================
' modPbAx -- VERIFICATION ONLY. Narrow probe: can this Excel create the ActiveX
' text box the requirement document's sync panel needs, and if not, what does
' work instead. Every variant is tried in its own guarded block so one failure
' does not hide the next answer.
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

Private Sub Note(ByVal tag As String, ByVal val As String)
    Say "OK    " & tag & " = " & val
End Sub

Private Sub Oops(ByVal tag As String)
    Dim n As Long
    Dim d As String
    n = Err.Number
    d = Err.Description
    Say "ERR   " & tag & " = " & n & " " & d
End Sub

Private Function Qual(ByVal procName As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

Public Sub PbAxPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG  project compiles"
End Sub

Public Sub PbAxArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, Qual("PbAxRun")
End Sub

Public Sub PbAxRun()
    Dim ws As Worksheet
    Dim ole As OLEObject
    Dim shp As Shape
    Dim o As Object

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Set ws = ThisWorkbook.Worksheets(1)
    Say "=== ax probe start ==="
    Say "START ctx"
    Note "ctx", "visible=" & Application.Visible & " interactive=" & Application.Interactive & _
        " books=" & Application.Workbooks.Count & " sheetProtected=" & ws.ProtectContents

    On Error GoTo E1
    Say "START ole_add"
    ws.Activate
    Set ole = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", _
        Left:=10, Top:=10, Width:=200, Height:=60)
    Note "ole_add", "name=" & ole.Name & " progid=" & ole.progID
    On Error Resume Next
    ole.Object.MultiLine = True
    ole.Object.Text = "abc"
    Note "ole_text", "[" & ole.Object.Text & "]"
    On Error GoTo E2
    GoTo S2
E1:
    Oops "ole_add"
    On Error GoTo E2

S2:
    Say "START ole_add_progid"
    Set ole = ws.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, _
        DisplayAsIcon:=False, Left:=10, Top:=90, Width:=200, Height:=60)
    Note "ole_add_progid", "name=" & ole.Name
    GoTo S3
E2:
    Oops "ole_add_progid"

S3:
    On Error GoTo E3
    Say "START addoleobject"
    Set shp = ws.Shapes.AddOLEObject(ClassType:="Forms.TextBox.1", _
        Left:=230, Top:=10, Width:=200, Height:=60)
    Note "addoleobject", "name=" & shp.Name
    GoTo S4
E3:
    Oops "addoleobject"

S4:
    On Error GoTo E4
    Say "START ole_add_label"
    Set ole = ws.OLEObjects.Add(ClassType:="Forms.Label.1", _
        Left:=460, Top:=10, Width:=120, Height:=30)
    Note "ole_add_label", "name=" & ole.Name
    GoTo S5
E4:
    Oops "ole_add_label"

S5:
    On Error GoTo E5
    Say "START createobject_textbox"
    Set o = CreateObject("Forms.TextBox.1")
    Note "createobject_textbox", TypeName(o)
    GoTo S6
E5:
    Oops "createobject_textbox"

S6:
    On Error GoTo E6
    Say "START ole_add_msforms2"
    Set ole = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
        Left:=460, Top:=60, Width:=120, Height:=30)
    Note "ole_add_msforms2", "name=" & ole.Name
    GoTo S7
E6:
    Oops "ole_add_msforms2"

S7:
    ' the fallback the implementation would have to use instead: a merged range
    ' the operator types into. Timers are suspended while a cell is in edit
    ' mode, so this is measured, not assumed.
    On Error GoTo E7
    Say "START merged_cell_input"
    ws.Range("H10:M20").Merge
    ws.Range("H10").Value = "typed here"
    Note "merged_cell_input", "[" & ws.Range("H10").Value & "]"
    GoTo Fin
E7:
    Oops "merged_cell_input"

Fin:
    Say "=== ax probe done ==="
End Sub
