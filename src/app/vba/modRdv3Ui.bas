Attribute VB_Name = "modRdv3Ui"
'==============================================================================
' modRdv3Ui -- everything the UI sheet shows. Rendering only; no decisions.
'
' The screen is cells. build\ui_grid_app.ps1 gives the sheet cells small enough
' to be used as PIXELS and merges each of the reference's rectangles
' (Rdv3Geom.cs / docs\ui-ref-geom.json) into one of them, then names it. So this
' module never knows an address: it writes through the names the builder made
' (rdvKeyValue, rdvVerdict, rdvCand_r_c, ...). A rectangle that moves in the
' reference moves in the builder, and nothing here changes.
'
' No UserForm, no ActiveX, no Forms control, no Shape. The buttons are hyperlink
' cells and the clicks arrive in modRdv3App through Worksheet_FollowHyperlink.
'
' What the run time is allowed to touch is deliberately small: Value2 on named
' ranges, in bulk where a block is written at once, plus the fill of the row that
' just gained or lost the selection. Colour, border and font are settled at build
' time. Re-formatting the sheet on a hot path is what makes an Excel screen feel
' slow, and the ledger figures are measured through this code.
'
' The two performance figures the spec allows on screen -- merge time and search
' time -- are written here and nowhere else. Every finer figure goes to the log.
'==============================================================================
Option Explicit

' the candidate table's shape, as the builder painted it
Public Const RDV3_UI_CAND_ROWS As Long = 10

' the same idle strings the C# screen uses (Rdv3Text.SubIdle / CandCount /
' PickToSee), so both builds say the same thing when they have nothing to show
Public Const RDV3_UI_IDLE_SUB As String = "未検索"
Public Const RDV3_UI_CAND_FMT As String = "候補 {n} 件"
Public Const RDV3_UI_PICK_TO_SEE As String = "候補一覧から行を選択すると表示されます。"
Public Const RDV3_UI_CAND_COLS As Long = 10
' the whole card, so the window can be fitted to it without guessing
Public Const RDV3_UI_CARD As String = "rdvCard"

' META!H3/H4: what this workbook changed about the Excel it runs in, written
' down where a VBA project reset cannot lose it (see Rdv3UiRestoreShell)
Private Const SHELL_ROW As Long = 3          ' META!H3/H4: what we changed
Private Const SHELL_COL As Long = 8


Private m_animBase As String
Private m_animStep As Long
Private m_animLast As Double
Private m_animOn As Boolean
Private m_selRow As Long             ' the candidate row drawn as selected, or -1
Private m_names As Object            ' name -> address, resolved once

Private Function UiWs() As Object
    Set UiWs = ThisWorkbook.Worksheets(RDV3_SHEET_UI)
End Function

' The builder's names are the contract between the sheet and this code. If one is
' missing the screen is not the screen this build was made for, and writing to
' the wrong cell would be worse than saying so.
Private Function NamedRange(ByVal nm As String) As Object
    Dim addr As String
    On Error GoTo Missing
    If m_names Is Nothing Then
        Set m_names = CreateObject("Scripting.Dictionary")
        m_names.CompareMode = 1
    End If
    If m_names.Exists(nm) Then
        Set NamedRange = UiWs().Range(CStr(m_names(nm)))
        Exit Function
    End If
    addr = ThisWorkbook.Names(nm).RefersToRange.Address
    m_names(nm) = addr
    Set NamedRange = UiWs().Range(addr)
    Exit Function
Missing:
End Function

Private Sub PutName(ByVal nm As String, ByVal v As Variant)
    Dim r As Object
    On Error Resume Next
    Set r = NamedRange(nm)
    If r Is Nothing Then Exit Sub
    r.Value2 = v
End Sub

' MAKE THE WORKBOOK LOOK AND BEHAVE LIKE THE APPLICATION IT IS.
'
' The C# build owns its window: its card is the whole window, so the screen, the
' record and the status bar are all in front of the operator at once. A workbook
' opens inside Excel, and Excel's own chrome (the ribbon, the formula bar, the
' name box) takes about a third of the height -- so the same card ran off the
' bottom and the record and the status bar were a page-down away. That is not
' the C# UI/UX.
'
' So the app hides Excel's chrome while it owns the window, and zooms the sheet
' until the whole card fits, exactly as the C# build scales itself to the DPI.
' Everything here is Excel's own API: no Win32, no Shell, no external helper.
' The ribbon and the formula bar are borrowed only while this workbook is alone
' in the operator's Excel, and handed back from what is recorded in META.
Public Sub Rdv3UiFitWindow()
    Dim ws As Object
    Dim w As Object
    Dim cardW As Double
    Dim cardH As Double
    Dim zw As Double
    Dim zh As Double
    Dim z As Long
    Dim pass As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RDV3_SHEET_UI)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    HideExcelShell
    Set w = ActiveWindow
    If w Is Nothing Then Exit Sub
    w.DisplayGridlines = False
    w.DisplayHeadings = False
    w.DisplayWorkbookTabs = False
    w.WindowState = xlMaximized
    w.ScrollRow = 1
    w.ScrollColumn = 1

    ' the card's own size, in points, taken from the sheet rather than assumed
    cardW = ws.Range(RDV3_UI_CARD).Width
    cardH = ws.Range(RDV3_UI_CARD).Height
    If cardW <= 0 Or cardH <= 0 Then Exit Sub

    ' A first estimate from the window's own measurements ...
    zw = (w.UsableWidth * 0.99) / cardW
    zh = (w.UsableHeight * 0.99) / cardH
    If zh < zw Then zw = zh
    ' x100, NOT x the current zoom: UsableWidth/Height are reported in points of
    ' the WINDOW and do not change with zoom, so folding the current zoom in
    ' would apply the ratio twice
    z = CLng(zw * 100#)
    If z > 100 Then z = 100
    If z < 40 Then z = 40
    w.Zoom = z

    ' ... then ASK. UsableHeight is close but not exact (the scroll bars and the
    ' chrome settle after the zoom), and one row short means the status bar --
    ' the state, the counts and the timings -- is off the screen, which is the
    ' whole point of fitting. VisibleRange is what Excel actually shows, so step
    ' down until the card's last row is really on it.
    For pass = 1 To 20
        If CardIsVisible(ws, w) Then Exit For
        If w.Zoom <= 40 Then Exit For
        w.Zoom = w.Zoom - 2
        w.ScrollRow = 1
        w.ScrollColumn = 1
    Next pass
    w.ScrollRow = 1
    w.ScrollColumn = 1
End Sub

' is every row and column of the card on the screen right now?
Private Function CardIsVisible(ByVal ws As Object, ByVal w As Object) As Boolean
    Dim card As Object
    Dim vis As Object
    On Error Resume Next
    Set card = ws.Range(RDV3_UI_CARD)
    Set vis = w.VisibleRange
    If card Is Nothing Or vis Is Nothing Then Exit Function
    CardIsVisible = (vis.Row + vis.Rows.Count - 1 >= card.Row + card.Rows.Count - 1) And _
                    (vis.Column + vis.Columns.Count - 1 >= card.Column + card.Columns.Count - 1)
End Function

' THE FORMULA BAR AND THE RIBBON BELONG TO THE WHOLE EXCEL, and this workbook
' runs in the OPERATOR'S Excel -- docs/app.md says so, and a plain double-click
' is absorbed into an Excel that is already running (measured). So there is no
' instance this build may call its own, and the chrome is borrowed under strict
' terms:
'
'   - only while this workbook is the ONLY one open, which is the case where
'     hiding the chrome shows the operator nothing but this app anyway;
'   - the previous state of BOTH is written into META (H3/H4) before anything
'     changes, so what goes back is what was there -- including a ribbon the
'     operator had already collapsed, which must not be forced open;
'   - and the moment another workbook joins this Excel, the chrome goes back
'     immediately (Rdv3UiShellRecheck, called from the pump), because from then
'     on it is not only this app's screen.
'
' The window's own gridlines, headings, sheet tabs, zoom and window state are
' properties of THIS workbook's window and go when it closes, so they are not
' restored here.
Private Function ShellCell(ByVal offset As Long) As Object
    On Error Resume Next
    Set ShellCell = ThisWorkbook.Worksheets(RDV3_SHEET_META).Cells(SHELL_ROW + offset, SHELL_COL)
End Function

' the ribbon is "up" when its bar is tall; a collapsed one is a strip
Private Function RibbonShown() As Boolean
    On Error Resume Next
    RibbonShown = (Application.CommandBars("Ribbon").Height > 100)
End Function

Private Sub HideExcelShell()
    Dim c As Object
    On Error Resume Next
    If Application.Workbooks.Count <> 1 Then Exit Sub
    Set c = ShellCell(0)
    If c Is Nothing Then Exit Sub
    If Len(CStr(c.Value2)) > 0 Then Exit Sub               ' already recorded: do not stack
    c.Value2 = IIf(Application.DisplayFormulaBar, "1", "0")
    ShellCell(1).Value2 = IIf(RibbonShown(), "1", "0")     ' what it WAS, not what we did
    Application.DisplayFormulaBar = False
    If CStr(ShellCell(1).Value2) = "1" Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    End If
    ThisWorkbook.Saved = True
End Sub

' Another workbook in this Excel means the chrome is no longer only this app's
' business. Called from the pump, so it takes effect within a tick.
Public Sub Rdv3UiShellRecheck()
    On Error Resume Next
    If Application.Workbooks.Count > 1 Then Rdv3UiRestoreShell
End Sub

' Put back exactly what was recorded, and nothing else.
Public Sub Rdv3UiRestoreShell()
    Dim c As Object
    On Error Resume Next
    Set c = ShellCell(0)
    If c Is Nothing Then Exit Sub
    If Len(CStr(c.Value2)) = 0 Then Exit Sub               ' we changed nothing
    Application.DisplayFormulaBar = (CStr(c.Value2) = "1")
    If CStr(ShellCell(1).Value2) = "1" Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    End If
    c.Value2 = ""
    ShellCell(1).Value2 = ""
    ThisWorkbook.Saved = True
End Sub

Public Function Rdv3UiRangeOf(ByVal nm As String) As Object
    Set Rdv3UiRangeOf = NamedRange(nm)
End Function

'------------------------------------------------------------------------------
' status bar and summary
'------------------------------------------------------------------------------
Public Sub Rdv3UiState(ByVal text As String)
    PutName "rdvStState", text
End Sub

Public Sub Rdv3UiNotepad(ByVal text As String)
    PutName "rdvStWatch", text
End Sub

Public Sub Rdv3UiLedgerInfo(ByVal text As String)
    PutName "rdvStLedger", text
End Sub

Public Sub Rdv3UiMergeMs(ByVal ms As Double)
    ' ui-spec 9: a figure that has not been measured yet gets no segment at all,
    ' rather than an empty frame
    If ms >= 0 Then
        PutName "rdvStMerge", "マージ " & Rdv3FmtMs(ms)
    Else
        PutName "rdvStMerge", ""
    End If
End Sub

Public Sub Rdv3UiSearchMs(ByVal ms As Double)
    If ms >= 0 Then
        PutName "rdvStSearch", "検索 " & Rdv3FmtMs(ms)
    Else
        PutName "rdvStSearch", ""
    End If
End Sub

Public Sub Rdv3UiIdentity(ByVal userName As String, ByVal host As String, _
                          ByVal role As String, ByVal selfId As String, _
                          ByVal logName As String)
    PutName "rdvIdentName", userName
    PutName "rdvIdentSub", host
    PutName "rdvIdentTag", role
    PutName "rdvStId", selfId & "   " & logName
End Sub

Public Sub Rdv3UiRows(ByVal rows As Long, ByVal savedAt As String)
    If rows > 0 Then
        PutName "rdvRowsValue", Format$(rows, "#,##0") & " 件"
    Else
        PutName "rdvRowsValue", ""
    End If
    If Len(savedAt) > 0 Then
        PutName "rdvRowsSub", "最終更新 " & savedAt
    Else
        PutName "rdvRowsSub", ""
    End If
End Sub

' The system error row. ui-spec 11 keeps this for the app's own failures; a
' number the operator mistyped is answered under the input box instead.
Public Sub Rdv3UiError(ByVal text As String)
    If Len(text) > 0 Then
        PutName "rdvErrTag", "エラー"
        PutName "rdvErrText", text
    Else
        PutName "rdvErrTag", ""
        PutName "rdvErrText", ""
    End If
End Sub

Public Sub Rdv3UiInputError(ByVal text As String)
    PutName "rdvInputErr", text
End Sub

' Cells(1, 1), not the range: every element of this screen is a MERGE, and
' Value2 on a multi-cell range hands back a 2-D array. CStr of an array is a type
' mismatch, and an unhandled error inside an Application.Run in an invisible
' Excel is a modal nobody can see -- the call never returns (AGENTS.md). Reading
' the screen must not be able to do that, so it takes the top-left cell and
' cannot raise.
Public Function Rdv3UiInputKey() As String
    Dim r As Object
    On Error Resume Next
    Set r = NamedRange("rdvInput")
    If r Is Nothing Then Exit Function
    Rdv3UiInputKey = Trim$(CStr(r.Cells(1, 1).Value2))
End Function

Public Sub Rdv3UiSetInput(ByVal text As String)
    PutName "rdvInput", text
End Sub

'------------------------------------------------------------------------------
' the checking animation
'------------------------------------------------------------------------------
Public Sub Rdv3UiAnimStart(ByVal baseText As String)
    m_animBase = baseText
    m_animStep = 0
    m_animLast = 0
    m_animOn = True
    Rdv3UiState baseText
End Sub

Public Sub Rdv3UiAnimStop()
    m_animOn = False
End Sub

' cheap enough to call from every chunk: repaints at most ~3 times a second
Public Sub Rdv3UiAnimPump()
    If Not m_animOn Then Exit Sub
    If Timer - m_animLast < 0.3 And Timer >= m_animLast Then Exit Sub
    m_animLast = Timer
    m_animStep = (m_animStep + 1) Mod 4
    Rdv3UiState m_animBase & String$(m_animStep, "・")
End Sub

'------------------------------------------------------------------------------
' the summary's first two blocks: the number, and what it resolved to
'------------------------------------------------------------------------------
Public Sub Rdv3UiKey(ByVal key As String, ByVal sub1 As String)
    PutName "rdvKeyValue", key
    PutName "rdvKeySub", sub1
End Sub

' ui-spec 4: HOLD and VOID are NG, anything else is OK, nothing selected is a
' dash. The rule lives here and nowhere else.
Public Function Rdv3UiVerdictOf(ByVal bStatus As String) As String
    Select Case UCase$(Trim$(bStatus))
        Case "HOLD", "VOID"
            Rdv3UiVerdictOf = "NG"
        Case ""
            Rdv3UiVerdictOf = ""
        Case Else
            Rdv3UiVerdictOf = "OK"
    End Select
End Function

Public Sub Rdv3UiStatusBlock(ByVal value As String, ByVal sub1 As String)
    PutName "rdvStatusValue", value
    PutName "rdvStatusSub", sub1
End Sub

'------------------------------------------------------------------------------
' the candidate table
'
' rows(1..n, 1..RDV3_UI_CAND_COLS) is written as ONE block per row, which is one
' COM call per row rather than one per cell. The table is small (the shown count
' is capped by the sheet) so this stays well inside the pump's tick.
'------------------------------------------------------------------------------
Public Sub Rdv3UiShowCandidates(ByRef rows() As Variant, ByVal n As Long)
    Dim r As Long
    Dim c As Long
    Dim cell As Object

    Rdv3UiClearCandidates
    If n <= 0 Then Exit Sub
    If n > RDV3_UI_CAND_ROWS Then n = RDV3_UI_CAND_ROWS
    For r = 1 To n
        For c = 1 To RDV3_UI_CAND_COLS
            Set cell = NamedRange("rdvCand_" & CStr(r - 1) & "_" & CStr(c - 1))
            If Not cell Is Nothing Then cell.Value2 = rows(r, c)
        Next c
    Next r
End Sub

Public Sub Rdv3UiClearCandidates()
    Dim r As Long
    Dim c As Long
    Dim cell As Object
    Rdv3UiSelectRow -1
    For r = 0 To RDV3_UI_CAND_ROWS - 1
        For c = 0 To RDV3_UI_CAND_COLS - 1
            Set cell = NamedRange("rdvCand_" & CStr(r) & "_" & CStr(c))
            If Not cell Is Nothing Then cell.Value2 = ""
        Next c
    Next r
End Sub

' The selected row: accent-100 behind it, the way the reference marks it. Only
' the row that gains it and the row that loses it are touched -- the rest of the
' sheet is never re-formatted.
Public Sub Rdv3UiSelectRow(ByVal slot As Long)
    Dim c As Long
    Dim cell As Object
    If m_selRow = slot Then Exit Sub
    If m_selRow >= 0 Then
        For c = 0 To RDV3_UI_CAND_COLS - 1
            Set cell = NamedRange("rdvCand_" & CStr(m_selRow) & "_" & CStr(c))
            If Not cell Is Nothing Then cell.Interior.Color = 16777215
        Next c
    End If
    m_selRow = slot
    If slot >= 0 Then
        For c = 0 To RDV3_UI_CAND_COLS - 1
            Set cell = NamedRange("rdvCand_" & CStr(slot) & "_" & CStr(c))
            If Not cell Is Nothing Then cell.Interior.Color = 16774894   ' accent-100
        Next c
    End If
End Sub

' the 処理済み column of one candidate row. Written when a mark is CONFIRMED,
' so the row the operator picked from says the same thing the record does.
Public Sub Rdv3UiSetCandProcessed(ByVal slot As Long, ByVal text As String)
    Dim cell As Object
    If slot < 0 Or slot >= RDV3_UI_CAND_ROWS Then Exit Sub
    Set cell = NamedRange("rdvCand_" & CStr(slot) & "_" & CStr(RDV3_UI_CAND_COLS - 1))
    If cell Is Nothing Then Exit Sub
    cell.Value2 = text
End Sub

Public Sub Rdv3UiVerdict(ByVal text As String, ByVal tag As String)
    PutName "rdvVerdict", text
    PutName "rdvListTag", tag
End Sub

'------------------------------------------------------------------------------
' the merged record
'
' ui-spec 4 and 4.2: the left column carries the seven values the candidate table
' does NOT already show, and the two boxes carry b_memo and c_remark. Everything
' else stays in the ledger and is simply not on this screen.
'------------------------------------------------------------------------------
Public Sub Rdv3UiShowRecord(ByRef f() As String, ByVal procText As String, _
                            ByVal key2 As String)
    ' the same seven columns the C# build shows, in the same order:
    ' Rdv3Ui.KvCols = 3, 2, 4, 9, 5, 6, 7 (a_name a_code a_grade a_dept a_date
    ' a_amount, and a_rate with a_flag)
    PutName "rdvKv0", Blank(f(3))
    PutName "rdvKv1", Blank(f(2))
    PutName "rdvKv2", Blank(f(4))
    PutName "rdvKv3", Blank(f(9))
    PutName "rdvKv4", Blank(f(5))
    PutName "rdvKv5", Blank(f(6))
    PutName "rdvKv6", Blank(f(7)) & " ・ " & Blank(f(8))
    PutName "rdvMemo", f(18)
    PutName "rdvRemark", f(27)
    PutName "rdvRecTagStatus", f(16)
    PutName "rdvRecTagProc", procText
    PutName "rdvRecTagKey2", "番号2 = " & key2
End Sub

' ui-spec 9: nothing selected leaves the cell empty; selected but empty in the
' ledger says so, because "shown, and there is no value" is not the same answer
' as "nothing is shown".
Private Function Blank(ByVal s As String) As String
    If Len(s) = 0 Then
        Blank = "N/A"
    Else
        Blank = s
    End If
End Function

Public Sub Rdv3UiClearRecord()
    PutName "rdvKv0", ""
    PutName "rdvKv1", ""
    PutName "rdvKv2", ""
    PutName "rdvKv3", ""
    PutName "rdvKv4", ""
    PutName "rdvKv5", ""
    PutName "rdvKv6", ""
    PutName "rdvMemo", RDV3_UI_PICK_TO_SEE
    PutName "rdvRemark", RDV3_UI_PICK_TO_SEE
    PutName "rdvRecTagStatus", ""
    PutName "rdvRecTagProc", ""
    PutName "rdvRecTagKey2", ""
End Sub

Public Sub Rdv3UiProcState(ByVal text As String)
    PutName "rdvRecTagProc", text
End Sub

' クリア: the input and the current result only. The ledger, the processed flags
' and the two timing figures stay.
Public Sub Rdv3UiClearResult()
    Rdv3UiSetInput ""
    Rdv3UiInputError ""
    Rdv3UiKey "", RDV3_UI_IDLE_SUB
    Rdv3UiStatusBlock "", ""
    Rdv3UiVerdict "", Replace(RDV3_UI_CAND_FMT, "{n}", "0")
    Rdv3UiClearRecord
    Rdv3UiClearCandidates
End Sub

'------------------------------------------------------------------------------
' operable / not operable
'
' The C# screen greys the controls it will refuse (Rdv3Ui.EnableOps). The sheet
' does the same with the one thing a cell has: the button's ink. The handlers
' still refuse on their own -- this says so before the click, it does not replace
' the check.
'------------------------------------------------------------------------------
Public Sub Rdv3UiEnableOps(ByVal enabled As Boolean)
    PaintBtn "rdvBtnSearch", enabled, True
    PaintBtn "rdvBtnClear", enabled, False
    PaintBtn "rdvBtnProcessed", enabled, False
End Sub

Public Sub Rdv3UiEnableProcessed(ByVal enabled As Boolean)
    PaintBtn "rdvBtnProcessed", enabled, False
End Sub

Private Sub PaintBtn(ByVal nm As String, ByVal enabled As Boolean, ByVal primary As Boolean)
    Dim r As Object
    Set r = NamedRange(nm)
    If r Is Nothing Then Exit Sub
    If primary Then
        If enabled Then
            r.Interior.Color = 10911833          ' accent
            r.Font.Color = 15987442              ' bg
        Else
            r.Interior.Color = 13553101          ' divider grey
            r.Font.Color = 15987442
        End If
    Else
        If enabled Then
            r.Font.Color = 2105117               ' ink
        Else
            r.Font.Color = 13553101              ' divider grey: visibly not available
        End If
    End If
End Sub
