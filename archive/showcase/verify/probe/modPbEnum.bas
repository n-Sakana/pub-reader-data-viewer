Attribute VB_Name = "modPbEnum"
'==============================================================================
' modPbEnum -- VERIFICATION ONLY, and it must be opened by the SHELL.
'
' 問い：デスクトップ列挙は本当に使えないのか。使えないとして、
'       「メモ帳だけ」の条件付き列挙なら安全なのか。描画と同時が悪いのか。
'
' このリポジトリの測定（src\v2\vba\modRdv2Uia.bas）はこう書いている：
'   全列挙は UIA が「すべての最上位ウィンドウ」にプロパティを尋ねるので、その
'   中の Excel 自身のプロバイダと、すでに止まっているこのスレッドが待ち合う。
'   C# ビルドが同じ列挙を安全にできるのは別の MTA スレッドから呼ぶから。
'
' そこで 3 通りを同じ条件で回す。1 周ごとに前後を記録するので、返ってこなく
' なった周回がそのまま残る。
'   A: ClassName="Notepad" の条件付き列挙（+ 各窓の rect / name / hwnd）
'   B: 条件付き列挙 + 毎周セルを塗る（描画と同時が悪いのかを見る）
'   C: 全列挙（TrueCondition）+ 各窓のプロパティ読み
'==============================================================================
Option Explicit

Private Const TS_CHILDREN As Long = 2
Private Const UIA_NamePropertyId As Long = 30005
Private Const UIA_ClassNamePropertyId As Long = 30012
Private Const UIA_NativeWindowHandlePropertyId As Long = 30020
Private Const LOOPS As Long = 100

Private m_res As String
Private m_uia As UIAutomationClient.IUIAutomation
Private m_root As UIAutomationClient.IUIAutomationElement

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

Public Sub Auto_Open()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "AUTO_OPEN addins=" & Application.AddIns.Count & " com=" & Application.COMAddIns.Count
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PbEnRun"
End Sub

Public Sub PbEnPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbEnRun()
    Dim ws As Worksheet
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.DisplayAlerts = False
    Say "=== enum probe start ==="

    On Error GoTo Failed
    Set m_uia = New UIAutomationClient.CUIAutomation
    Set m_root = m_uia.GetRootElement

    ' 疑似ピクセルの盤面をこの端末の本番と同じ粒度で用意する（2px = 406x406）
    Set ws = ThisWorkbook.Worksheets(1)
    Application.ScreenUpdating = False
    ws.Range(ws.Rows(1), ws.Rows(406)).RowHeight = 1.5
    ws.Range(ws.Columns(1), ws.Columns(406)).ColumnWidth = 0.14
    ws.Range(ws.Cells(1, 1), ws.Cells(406, 406)).Interior.Color = &HF7F7F7
    Application.ScreenUpdating = True
    Say "board ready 406x406 rowH=" & ws.Rows(1).Height & " colW=" & ws.Columns(1).Width

    RunPhase "A_filtered_only", True, False, ws
    RunPhase "B_filtered_plus_paint", True, True, ws
    RunPhase "C_full_enum", False, False, ws

    Say "=== enum probe done ==="
    Exit Sub
Failed:
    Say "ERR setup " & Err.Number & " " & Err.Description
    Say "=== enum probe done ==="
End Sub

Private Sub RunPhase(ByVal tag As String, ByVal filtered As Boolean, _
                     ByVal paint As Boolean, ByVal ws As Worksheet)
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    Dim i As Long
    Dim k As Long
    Dim t0 As Double
    Dim ms As Double
    Dim worst As Double
    Dim total As Double
    Dim seen As Long
    Dim h As Variant

    On Error GoTo Failed
    Say "START " & tag
    If filtered Then
        Set cond = m_uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Notepad")
    Else
        Set cond = m_uia.CreateTrueCondition
    End If

    For i = 1 To LOOPS
        Say "  ." & tag & " " & i          ' 返らなくなった周回がここに残る
        t0 = Timer
        Set arr = m_root.FindAll(TS_CHILDREN, cond)
        seen = arr.Length
        For k = 0 To arr.Length - 1
            Set el = arr.GetElement(k)
            rc = el.CurrentBoundingRectangle
            h = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
            If Len(el.CurrentName) > 0 Then h = h
        Next k
        If paint Then
            Application.ScreenUpdating = False
            ws.Range(ws.Cells(20 + (i Mod 60), 20), ws.Cells(40 + (i Mod 60), 200)).Interior.Color = _
                IIf(i Mod 2 = 0, &HE8ECFB, &HFFFFFF)
            Application.ScreenUpdating = True
        End If
        ms = (Timer - t0) * 1000
        If ms < 0 Then ms = ms + 86400000#
        total = total + ms
        If ms > worst Then worst = ms
    Next i
    Say "OK    " & tag & " windows=" & seen & " avg=" & Format$(total / LOOPS, "0.0") & _
        "ms worst=" & Format$(worst, "0") & "ms"
    Exit Sub
Failed:
    Say "ERR   " & tag & " at loop " & i & " : " & Err.Number & " " & Err.Description
End Sub
