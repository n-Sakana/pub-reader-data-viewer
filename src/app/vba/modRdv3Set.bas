Attribute VB_Name = "modRdv3Set"
'==============================================================================
' modRdv3Set -- the settings screen.
'
' The C# build opens a modal window for this. A workbook has no second window
' worth having, so the settings are a SHEET, painted by the same pseudo-pixel
' grid as the main screen (build\ui_grid_app.ps1) and carrying the same three
' groups the C# dialog has: 監視対象 / 動作 / ファイル.
'
' What it edits is ReaderDataViewer.json -- the SAME file the C# build reads and
' writes (modRdv3Cfg). Cancel re-reads the file, so nothing is committed until
' 保存 is pressed.
'
' The live/next-start boundary is the one docs\settings.md states and the C#
' build keeps: paths and key.length take effect at the next start; digitsOnly,
' the watch targets, the timings and the candidate row count take effect as soon
' as they are saved. The screen says which is which rather than pretending.
'
' The FE never saves ITSELF. Editing these cells dirties the workbook, and a
' dirty FE would ask to be saved on close -- so the sheet is marked clean again
' after every write, exactly as the rest of this build does.
'==============================================================================
Option Explicit

Public Const RDV3_SHEET_SET As String = "SETTINGS"
Public Const RDV3_SET_ROWS As Long = 6          ' target rows the sheet carries

Private m_open As Boolean
' what the file said when the sheet was filled: the save edits THIS, so members
' the sheet has no column for survive a round trip
Private m_base As Collection
' the picked element, held for preview only until 採用 is pressed
Private m_pv As Object
Private m_expectReq As String
Private m_picking As Boolean

Private Function SetWs() As Object
    On Error Resume Next
    Set SetWs = ThisWorkbook.Worksheets(RDV3_SHEET_SET)
End Function

Public Function Rdv3SetIsOpen() As Boolean
    Rdv3SetIsOpen = m_open
End Function

' Show the settings sheet, filled from what is in memory (which is what the file
' said, plus anything a previous save changed).
Public Function Rdv3SetOpen() As String
    Dim ws As Object
    Set ws = SetWs()
    If ws Is Nothing Then
        Rdv3SetOpen = "設定シートがありません (このブックは設定画面なしで作られています)"
        Exit Function
    End If
    On Error GoTo Failed
    Rdv3SetLoad
    ws.Visible = -1                              ' xlSheetVisible
    ws.Activate
    m_open = True
    Exit Function
Failed:
    Rdv3SetOpen = "設定画面を開けませんでした " & Err.Number & ": " & Err.Description
End Function

Public Function Rdv3SetClose() As String
    Dim ws As Object
    Set ws = SetWs()
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    ' leaving the screen ends the pick. Otherwise the BE keeps sampling for the
    ' rest of its window and goes on writing previews into a hidden sheet, so
    ' the next time the operator opens it they are looking at an answer to a
    ' question they walked away from.
    If m_picking Then Rdv3SetCancelPreview
    m_open = False
    ThisWorkbook.Worksheets(RDV3_SHEET_UI).Activate
    ws.Visible = 0                               ' xlSheetHidden
End Function

' The settings sheet's three buttons. Same dispatch idea as the main screen: the
' builder's names say which cell is which, so this module holds no addresses.
Public Sub Rdv3SetClick(ByVal addr As String)
    On Error Resume Next
    If HitsSet(addr, "setBtnSave") Then
        Rdv3SetSave
    ElseIf HitsSet(addr, "setBtnCancel") Then
        ' the file is the truth; re-reading it is what "cancel" means -- and an
        ' answer still on its way would be describing the edits just discarded
        If m_picking Then Rdv3SetCancelPreview
        Rdv3CfgLoad Rdv3CfgSourcePath()
        Rdv3SetLoad
        PutSet "setNote", "編集を捨てて、設定ファイルを読み直しました"
        MarkClean
    ElseIf HitsSet(addr, "setBtnPick") Then
        ' no duration here: the BE owns the deadline and says it in its own
        ' answer. What THIS screen owns is that picking changes nothing.
        PutSet "setNote", "画面から選ぶ: 選んでいる間は設定を変えません。「この要素を使う」で反映します"
        MarkClean
        Rdv3AppRequestInspect
    ElseIf HitsSet(addr, "setBtnAdopt") Then
        Rdv3SetAdoptPreview
    ElseIf HitsSet(addr, "setBtnPickCancel") Then
        Rdv3SetCancelPreview
    ElseIf HitsSet(addr, "setBtnBack") Then
        Rdv3SetClose
    End If
End Sub

' THE PICKED ELEMENT IS A PREVIEW, NOT A CHANGE.
'
' Clicking a field in the other application tells this screen what that field is.
' It does NOT alter the settings: the description lands in the preview block,
' beside what the chosen row holds now, and nothing moves until 採用 is pressed.
' Choosing a different row, or picking again, only changes what is being
' previewed. 取消 throws the preview away.
'
' Applying on the click would mean a stray click while looking for the right box
' had already rewritten a target -- and the operator would have no way back.
' The BE's answer to a save: what the RUNNING session actually took. The screen
' reports the SESSION, not the file -- the file is what the next start will use,
' and saying "applied" for something that waits until then would be a lie the
' operator only finds out about later.
Public Sub Rdv3SetConfigResult(ByVal metaString As String)
    Dim d As Object
    On Error Resume Next
    Set d = Rdv3ChMeta(metaString)
    If CStr(d("res")) <> "ok" Then
        PutSet "setNote", "保存はしましたが、動作中のセッションへ反映できませんでした: " & CStr(d("msg"))
        MarkClean
        Exit Sub
    End If
    PutSet "setNote", "保存しました。監視 " & CStr(d("watched")) & "/" & CStr(d("targets")) & _
        " 件・監視間隔 " & CStr(d("poll")) & "ms・確定待ち " & CStr(d("stable")) & "ms・再接続 " & _
        CStr(d("rebind")) & "ms・候補 " & CStr(d("cand")) & " 行・数字のみ " & _
        IIf(CStr(d("digits")) = "1", "はい", "いいえ") & _
        " を今のセッションに反映しました。桁数 (" & CStr(d("keylen")) & ") とファイルは次回の起動からです"
    MarkClean
End Sub

Public Sub Rdv3SetInspectResult(ByVal metaString As String)
    Dim d As Object
    Dim res As String

    On Error Resume Next
    Set d = Rdv3ChMeta(metaString)
    ' an answer to a press the operator has already moved on from is not this
    ' screen's answer, and must not appear as though it were
    If CStr(d("req")) <> m_expectReq Then
        Exit Sub
    End If
    res = CStr(d("res"))
    If res <> "preview" Then
        ' waiting / timeout / closed: a message, never a change
        Set m_pv = Nothing
        ClearPreview
        PutSet "setPvNote", CStr(d("msg"))
        m_picking = (res = "waiting")
        ' timeout or closed: the BE has stopped, so the shortcut goes with it
        If Not m_picking Then BindEscape False
        MarkClean
        Exit Sub
    End If

    ' a preview and only a preview: the row is untouched until 「この要素を使う」
    Set m_pv = d
    m_picking = True
    ShowPreview
    MarkClean
End Sub

' Esc stops the pick, for as long as Excel is the window with the keyboard. It
' cannot reach here while the operator is inside the target application -- that
' is the same boundary that rules out Ctrl+Shift -- so the panel's 「閉じる」
' stays the way that always works, and this is the shortcut for when the
' settings screen is in front.
Private Sub BindEscape(ByVal on_ As Boolean)
    On Error Resume Next
    If on_ Then
        Application.OnKey "{ESC}", "Rdv3SetEscape"
    Else
        Application.OnKey "{ESC}"
    End If
End Sub

Public Sub Rdv3SetEscape()
    If m_picking Then Rdv3SetCancelPreview
End Sub

Public Sub Rdv3SetInspectExpect(ByVal reqId As String)
    m_expectReq = reqId
    ' PICKING STARTS HERE, not when the BE's first answer arrives. The request is
    ' already on its way, so a 戻る / 取消 / Esc in that gap has to send the stop
    ' -- otherwise the BE samples out its whole window and publishes previews
    ' into a sheet nobody is looking at.
    m_picking = True
    BindEscape True
    Set m_pv = Nothing
    ClearPreview
    ' a placeholder for the moment between the press and the BE's answer. The
    ' BE's message replaces it and carries the deadline; stating the deadline
    ' twice is what let the two drift apart.
    PutSet "setPvNote", "対象アプリの欄をクリックすると、そこをプレビューします"
    MarkClean
End Sub

' what was picked, next to what the chosen row holds now, so the operator can
' see exactly what 採用 would change before deciding
Private Sub ShowPreview()
    Dim r As Long
    r = TargetRow()
    PutSet "setPvClass", CStr(m_pv("cls"))
    PutSet "setPvProc", CStr(m_pv("proc"))
    PutSet "setPvWName", CStr(m_pv("wname"))
    PutSet "setPvFid", CStr(m_pv("fid"))
    PutSet "setPvFType", CStr(m_pv("ftype"))
    PutSet "setPvRead", CStr(m_pv("read"))
    PutSet "setPvNowClass", GetSet("setT" & CStr(r) & "_class")
    PutSet "setPvNowProc", GetSet("setT" & CStr(r) & "_proc")
    PutSet "setPvNowFid", GetSet("setT" & CStr(r) & "_fid")
    PutSet "setPvNowFType", GetSet("setT" & CStr(r) & "_ftype")
    PutSet "setPvNote", (r + 1) & " 行目へ「この要素を使う」で反映します。今は何も変えていません" & _
        IIf(Len(CStr(m_pv("fid"))) = 0, " (この欄には AutomationId が無いので、種類と位置で特定します)", "")
End Sub

Private Sub ClearPreview()
    Dim k As Variant
    For Each k In Array("setPvClass", "setPvProc", "setPvWName", "setPvFid", "setPvFType", _
                        "setPvRead", "setPvNowClass", "setPvNowProc", "setPvNowFid", "setPvNowFType")
        PutSet CStr(k), ""
    Next k
End Sub

Private Function TargetRow() As Long
    Dim r As Long
    r = CLng(Val(GetSet("setPickRow"))) - 1
    If r < 0 Or r >= RDV3_SET_ROWS Then r = 0
    TargetRow = r
End Function

' The explicit step. Only this writes anything, and it writes to the CELLS --
' the file is still only touched by 保存.
'
' The members this sheet has no column for (the field's class and name, its
' position among identical fields, the intermediate path) go into the row's
' underlying target, which the save carries through. Without them an Edit with no
' AutomationId could not be told from the one beside it.
Public Sub Rdv3SetAdoptPreview()
    Dim r As Long
    Dim t As Object
    Dim ids As Variant
    Dim i As Long

    If m_pv Is Nothing Then
        PutSet "setPvNote", "取り込んだ内容がありません。先に「画面から選ぶ」を押してください"
        MarkClean
        Exit Sub
    End If
    r = TargetRow()
    PutSet "setT" & CStr(r) & "_class", CStr(m_pv("cls"))
    PutSet "setT" & CStr(r) & "_proc", CStr(m_pv("proc"))
    PutSet "setT" & CStr(r) & "_fid", CStr(m_pv("fid"))
    PutSet "setT" & CStr(r) & "_ftype", CStr(m_pv("ftype"))
    PutSet "setT" & CStr(r) & "_read", CStr(m_pv("read"))
    ' the old row's title pattern described the old window. The save writes this
    ' cell into the target whatever the adopt did, so clearing the descriptor
    ' without clearing the cell would put the stale condition straight back.
    PutSet "setT" & CStr(r) & "_like", ""
    If Len(GetSet("setT" & CStr(r) & "_name")) = 0 Then
        PutSet "setT" & CStr(r) & "_name", CStr(m_pv("wname"))
    End If
    If Len(GetSet("setT" & CStr(r) & "_on")) = 0 Then
        PutSet "setT" & CStr(r) & "_on", "はい"
    End If

    ' A DIFFERENT ELEMENT MEANS A DIFFERENT DESCRIPTOR, so every matcher is
    ' rebuilt from the preview and none of the old one survives. Carrying the
    ' previous row's hidden members over left conditions that describe the
    ' element that is no longer there -- the built-in default's
    ' requireValuePattern=true, adopted onto a TextPattern-only field, is
    ' rejected by the matcher and the target can never bind. C# builds a whole
    ' new Rdv3Target here for exactly this reason; only what the OPERATOR typed
    ' (the display name, and 有効) is theirs to keep.
    If m_base Is Nothing Then Set m_base = Rdv3CfgCloneTargets()
    Do While m_base.Count < r + 1
        m_base.Add Rdv3CfgNewTarget()
    Loop
    Set t = Rdv3CfgNewTarget()
    t("enabled") = True
    t("name") = GetSet("setT" & CStr(r) & "_name")
    t("window")("automationId") = CStr(m_pv("wid"))
    t("window")("className") = CStr(m_pv("cls"))
    t("window")("processName") = CStr(m_pv("proc"))
    t("window")("descendants") = False
    t("field")("automationId") = CStr(m_pv("fid"))
    t("field")("className") = IIf(Len(CStr(m_pv("fid"))) = 0, CStr(m_pv("fcls")), "")
    t("field")("descendants") = True
    t("field")("index") = IIf(Len(CStr(m_pv("fid"))) = 0, CLng(Val(CStr(m_pv("fidx")))), 0)
    t("field")("requireValuePattern") = (CStr(m_pv("read")) = "value")
    t("read") = CStr(m_pv("read"))
    If Len(CStr(m_pv("path"))) > 0 Then
        ids = Split(CStr(m_pv("path")), "/")
        For i = LBound(ids) To UBound(ids)
            If Len(Trim$(CStr(ids(i)))) > 0 Then
                t("path").Add StepWithId(Trim$(CStr(ids(i))))
            End If
        Next i
    End If
    ReplaceAt m_base, r + 1, t

    ' the picker is finished with: stop it sampling, and say plainly that the
    ' FILE has not changed -- that is still only 保存
    Set m_pv = Nothing
    m_picking = False
    BindEscape False
    Rdv3AppRequestInspectStop
    ClearPreview
    PutSet "setPvNote", (r + 1) & " 行目に反映しました。設定ファイルはまだ変わっていません。「保存」を押すと書き込みます"
    MarkClean
End Sub

Private Function StepWithId(ByVal id As String) As Object
    Dim m As Object
    Set m = Rdv3CfgNewMatch()
    m("automationId") = id
    Set StepWithId = m
End Function

' a Collection item cannot be assigned, so replacing one is a remove and an
' insert at the same position
Private Sub ReplaceAt(ByVal c As Collection, ByVal idx As Long, ByVal item As Object)
    If idx < 1 Or idx > c.Count Then Exit Sub
    If idx = c.Count Then
        c.Remove idx
        c.Add item
    Else
        c.Remove idx
        c.Add item, , idx
    End If
End Sub

' Esc / 閉じる: the picker stops and NOTHING was adopted. The working cells and
' the file are both exactly as they were.
Public Sub Rdv3SetCancelPreview()
    Set m_pv = Nothing
    m_expectReq = ""
    m_picking = False
    BindEscape False
    Rdv3AppRequestInspectStop
    ClearPreview
    PutSet "setPvNote", "ピッカーを閉じました。設定は何も変わっていません"
    MarkClean
End Sub

Private Function HitsSet(ByVal addr As String, ByVal nm As String) As Boolean
    Dim r As Object
    On Error Resume Next
    Set r = SetRange(nm)
    If r Is Nothing Then Exit Function
    HitsSet = (r.Cells(1, 1).Address(0, 0) = addr)
End Function

'------------------------------------------------------------------------------
' file -> sheet
'------------------------------------------------------------------------------
Public Sub Rdv3SetLoad()
    Dim t As Object
    Dim i As Long

    Set m_base = Rdv3CfgCloneTargets()

    PutSet "setKeyLen", CStr(Rdv3CfgKeyLength())
    PutSet "setDigits", Yn(Rdv3CfgKeyDigitsOnly())
    PutSet "setPoll", CStr(Rdv3CfgPollMs())
    PutSet "setStable", CStr(Rdv3CfgStableMs())
    PutSet "setRebind", CStr(Rdv3CfgRebindMs())
    PutSet "setCand", CStr(Rdv3CfgCandidateRows())
    PutSet "setFocus", Yn(Rdv3CfgPreferFocused())
    PutSet "setDataDir", Rdv3CfgDataDir()
    PutSet "setLedger", Rdv3CfgLedger()
    PutSet "setLog", Rdv3CfgLog()
    PutSet "setFile", Rdv3CfgSourcePath()

    For i = 0 To RDV3_SET_ROWS - 1
        Set t = Nothing
        If i < Rdv3CfgTargetCount() Then Set t = Rdv3CfgTarget(i + 1)
        If t Is Nothing Then
            PutSet "setT" & CStr(i) & "_on", ""
            PutSet "setT" & CStr(i) & "_name", ""
            PutSet "setT" & CStr(i) & "_class", ""
            PutSet "setT" & CStr(i) & "_proc", ""
            PutSet "setT" & CStr(i) & "_like", ""
            PutSet "setT" & CStr(i) & "_fid", ""
            PutSet "setT" & CStr(i) & "_ftype", ""
            PutSet "setT" & CStr(i) & "_read", ""
            PutSet "setT" & CStr(i) & "_why", ""
        Else
            PutSet "setT" & CStr(i) & "_on", Yn(CBool(t("enabled")))
            PutSet "setT" & CStr(i) & "_name", CStr(t("name"))
            PutSet "setT" & CStr(i) & "_class", CStr(t("window")("className"))
            PutSet "setT" & CStr(i) & "_proc", CStr(t("window")("processName"))
            PutSet "setT" & CStr(i) & "_like", CStr(t("window")("nameLike"))
            PutSet "setT" & CStr(i) & "_fid", CStr(t("field")("automationId"))
            PutSet "setT" & CStr(i) & "_ftype", JoinTypes(t("field")("controlTypes"))
            PutSet "setT" & CStr(i) & "_read", CStr(t("read"))
            PutSet "setT" & CStr(i) & "_why", Rdv3CfgWhyNotWatchable(t)
        End If
    Next i
    PutSet "setNote", ""
    MarkClean
End Sub

'------------------------------------------------------------------------------
' sheet -> file
'------------------------------------------------------------------------------
' THE SHEET SHOWS PART OF A TARGET, SO A SAVE MUST NOT WRITE ONLY THAT PART.
'
' ReaderDataViewer.json is the C# build's file too, and it carries members this
' sheet has no column for: window automationId / name / index / scope, the
' intermediate path, the field's className / name / nameLike / processName /
' requireValuePattern / index / scope -- and any target past the sixth. Building
' the list from the cells alone would delete every one of them, so opening the
' settings screen and pressing 保存 without typing anything would quietly cut a
' C#-written configuration down to what this sheet happens to display.
'
' So the save EDITS: each row starts as a deep copy of the target it was loaded
' from, the shown members are overwritten, and everything else is carried
' through untouched. Targets past the sixth are carried through whole.
'
' Deleting is explicit -- 有効 = 削除 -- because a blank row cannot mean
' "remove this", not when the thing being removed may be a matcher this sheet
' never showed in the first place.
Public Function Rdv3SetSave() As String
    Dim list As Collection
    Dim t As Object
    Dim orig As Object
    Dim i As Long
    Dim nm As String
    Dim why As String
    Dim err1 As String
    Dim dropped As Long
    Dim removed As Long
    Dim kept As Long

    On Error GoTo Failed
    If m_base Is Nothing Then Set m_base = Rdv3CfgCloneTargets()
    Set list = New Collection

    For i = 0 To RDV3_SET_ROWS - 1
        Set orig = Nothing
        If i < m_base.Count Then Set orig = m_base.Item(i + 1)
        nm = Trim$(GetSet("setT" & CStr(i) & "_on"))
        If nm = "削除" Then
            removed = removed + 1
        ElseIf RowIsBlank(i) And orig Is Nothing Then
            ' an unused row, not an instruction
        Else
            If orig Is Nothing Then
                Set t = Rdv3CfgNewTarget()
            Else
                Set t = Rdv3CfgCloneTarget(orig)
                kept = kept + 1
            End If
            t("enabled") = IsYes(GetSet("setT" & CStr(i) & "_on"))
            t("name") = Trim$(GetSet("setT" & CStr(i) & "_name"))
            t("window")("className") = Trim$(GetSet("setT" & CStr(i) & "_class"))
            t("window")("processName") = Trim$(GetSet("setT" & CStr(i) & "_proc"))
            t("window")("nameLike") = Trim$(GetSet("setT" & CStr(i) & "_like"))
            t("field")("automationId") = Trim$(GetSet("setT" & CStr(i) & "_fid"))
            SplitTypes t("field")("controlTypes"), GetSet("setT" & CStr(i) & "_ftype")
            t("read") = Rdv3CfgReadModeName(Rdv3CfgReadModeId(GetSet("setT" & CStr(i) & "_read")))
            ' a brand new row with nothing but a control type still needs
            ' something to hold on to; an edited one keeps whatever it had
            If orig Is Nothing Then
                If Rdv3CfgMatchIsEmpty(t("field")) Then t("field")("requireValuePattern") = True
            End If
            why = Rdv3CfgWhyNotWatchable(t)
            PutSet "setT" & CStr(i) & "_why", why
            If Len(why) > 0 Then
                ' it would match any window at all: shown back with the reason
                ' rather than saved as a target that reads a stranger's screen
                dropped = dropped + 1
            Else
                list.Add t
            End If
        End If
    Next i

    ' anything past the last row this sheet has, carried through as it was
    For i = RDV3_SET_ROWS + 1 To m_base.Count
        list.Add m_base.Item(i)
    Next i

    Rdv3CfgSetKey CLng(Val(GetSet("setKeyLen"))), IsYes(GetSet("setDigits"))
    Rdv3CfgSetWatch CLng(Val(GetSet("setPoll"))), CLng(Val(GetSet("setStable"))), _
                    CLng(Val(GetSet("setRebind"))), IsYes(GetSet("setFocus"))
    Rdv3CfgSetCandidateRows CLng(Val(GetSet("setCand")))
    Rdv3CfgSetPaths Trim$(GetSet("setDataDir")), Trim$(GetSet("setLedger")), Trim$(GetSet("setLog"))
    Rdv3CfgSetTargets list

    err1 = Rdv3CfgSave(Rdv3CfgSourcePath())
    If Len(err1) > 0 Then
        Rdv3SetSave = err1
        PutSet "setNote", "保存できませんでした: " & err1
        Exit Function
    End If

    ' What may be adopted without a restart, and only that. This is the FE's own
    ' copy; the BE is a SEPARATE PROCESS and saw none of it, so it is asked to
    ' re-read the same file and answer with what it actually took. Until that
    ' answer arrives the screen says so rather than claiming the session changed.
    Rdv3SpecSetDigitsOnly Rdv3CfgKeyDigitsOnly()
    Rdv3SetLoad
    Rdv3AppRequestConfig
    PutSet "setNote", "保存しました" & _
        IIf(dropped > 0, " (" & CStr(dropped) & " 件は条件が空のため対象外)", "") & _
        IIf(removed > 0, " (" & CStr(removed) & " 件を削除)", "") & _
        "。反映を確認中..."
    MarkClean
    Exit Function
Failed:
    Rdv3SetSave = "設定を保存できませんでした " & Err.Number & ": " & Err.Description
    PutSet "setNote", Rdv3SetSave
End Function

' every cell this sheet lets a person type into, for one row
Private Function RowIsBlank(ByVal i As Long) As Boolean
    Dim k As Variant
    For Each k In Array("on", "name", "class", "proc", "like", "fid", "ftype", "read")
        If Len(Trim$(GetSet("setT" & CStr(i) & "_" & CStr(k)))) > 0 Then Exit Function
    Next k
    RowIsBlank = True
End Function

'------------------------------------------------------------------------------
Private Function JoinTypes(ByVal c As Collection) As String
    Dim i As Long
    Dim s As String
    For i = 1 To c.Count
        If i > 1 Then s = s & ", "
        s = s & CStr(c.Item(i))
    Next i
    JoinTypes = s
End Function

Private Sub SplitTypes(ByVal c As Collection, ByVal s As String)
    Dim parts As Variant
    Dim i As Long
    Dim v As String
    Do While c.Count > 0
        c.Remove 1
    Loop
    If Len(Trim$(s)) = 0 Then Exit Sub
    parts = Split(Replace(Replace(s, "/", ","), " ", ","), ",")
    For i = LBound(parts) To UBound(parts)
        v = Trim$(CStr(parts(i)))
        If Len(v) > 0 Then c.Add v
    Next i
End Sub

Private Function Yn(ByVal b As Boolean) As String
    If b Then
        Yn = "はい"
    Else
        Yn = "いいえ"
    End If
End Function

Private Function IsYes(ByVal s As String) As Boolean
    Select Case LCase$(Trim$(s))
        Case "はい", "yes", "y", "true", "1", "on"
            IsYes = True
    End Select
End Function

Private Function SetRange(ByVal nm As String) As Object
    On Error Resume Next
    Set SetRange = ThisWorkbook.Names(nm).RefersToRange
End Function

Private Sub PutSet(ByVal nm As String, ByVal v As Variant)
    Dim r As Object
    On Error Resume Next
    Set r = SetRange(nm)
    If r Is Nothing Then Exit Sub
    r.Value2 = v
End Sub

' Cells(1, 1) for the same reason as Rdv3UiInputKey: these are merged blocks and
' Value2 on the whole block is an array.
Private Function GetSet(ByVal nm As String) As String
    Dim r As Object
    On Error Resume Next
    Set r = SetRange(nm)
    If r Is Nothing Then Exit Function
    GetSet = Trim$(CStr(r.Cells(1, 1).Value2))
End Function

' The FE has nothing to save and must never ask to be saved: writing into these
' cells dirties the book, so it is marked clean again straight away.
Private Sub MarkClean()
    On Error Resume Next
    ThisWorkbook.Saved = True
End Sub

' THE round trip this sheet has to survive: a configuration written by the C#
' build, loaded into these cells, saved again with nothing typed, and still
' carrying every member this sheet has no column for. Needs the SETTINGS sheet,
' so it is run by the builder against the painted book.
Public Function Rdv3SetSelfTest() As String
    Dim p As String
    Dim s As String
    Dim t As Object
    Dim bad As String
    Dim err1 As String

    On Error GoTo Failed
    p = Environ$("TEMP") & "\rdv3set-selftest.json"
    s = "{ ""watch"": { ""targets"": [" & _
        "{ ""enabled"": true, ""name"": ""deep""," & _
        "  ""window"": { ""automationId"": ""winId"", ""className"": ""LobApp"", ""name"": ""exact""," & _
        "                ""index"": 2, ""scope"": ""children"" }," & _
        "  ""path"": [ { ""automationId"": ""tabMain"" }, { ""automationId"": ""pnl"" } ]," & _
        "  ""field"": { ""automationId"": ""txt"", ""className"": ""EditCls"", ""nameLike"": ""bar*""," & _
        "               ""processName"": ""LobApp"", ""controlTypes"": [""Edit""], ""index"": 3 }," & _
        "  ""read"": ""text"" }," & _
        "{ ""enabled"": false, ""name"": ""t2"", ""window"": { ""className"": ""C2"" }, ""field"": { ""controlTypes"": [""Edit""] } }," & _
        "{ ""enabled"": true, ""name"": ""t3"", ""window"": { ""className"": ""C3"" }, ""field"": { ""controlTypes"": [""Edit""] } }," & _
        "{ ""enabled"": true, ""name"": ""t4"", ""window"": { ""className"": ""C4"" }, ""field"": { ""controlTypes"": [""Edit""] } }," & _
        "{ ""enabled"": true, ""name"": ""t5"", ""window"": { ""className"": ""C5"" }, ""field"": { ""controlTypes"": [""Edit""] } }," & _
        "{ ""enabled"": true, ""name"": ""t6"", ""window"": { ""className"": ""C6"" }, ""field"": { ""controlTypes"": [""Edit""] } }," & _
        "{ ""enabled"": true, ""name"": ""t7"", ""window"": { ""className"": ""C7"" }, ""field"": { ""controlTypes"": [""Edit""] } }" & _
        "] } }"
    If Not Rdv3CfgWriteRawForTest(p, s) Then
        Rdv3SetSelfTest = "FAIL cannot write the test file"
        Exit Function
    End If
    If Not Rdv3CfgLoad(p) Then
        Rdv3SetSelfTest = "FAIL parse: " & Rdv3CfgError()
        Exit Function
    End If

    ' fill the sheet, then save it back WITHOUT typing anything
    Rdv3SetLoad
    err1 = Rdv3SetSave()
    If Len(err1) > 0 Then
        Rdv3SetSelfTest = "FAIL save: " & err1
        Exit Function
    End If

    Rdv3CfgDefaults
    Rdv3CfgLoad p
    bad = bad & ExpectSet("targets", CStr(Rdv3CfgTargetCount()), "7")
    If Rdv3CfgTargetCount() >= 7 Then
        Set t = Rdv3CfgTarget(1)
        bad = bad & ExpectSet("win autoId", CStr(t("window")("automationId")), "winId")
        bad = bad & ExpectSet("win name", CStr(t("window")("name")), "exact")
        bad = bad & ExpectSet("win index", CStr(CLng(t("window")("index"))), "2")
        bad = bad & ExpectSet("path", CStr(t("path").Count), "2")
        bad = bad & ExpectSet("fld class", CStr(t("field")("className")), "EditCls")
        bad = bad & ExpectSet("fld like", CStr(t("field")("nameLike")), "bar*")
        bad = bad & ExpectSet("fld index", CStr(CLng(t("field")("index"))), "3")
        bad = bad & ExpectSet("read", CStr(t("read")), "text")
        Set t = Rdv3CfgTarget(2)
        bad = bad & ExpectSet("t2 off", CStr(CBool(t("enabled"))), CStr(False))
        Set t = Rdv3CfgTarget(7)
        bad = bad & ExpectSet("7th kept", CStr(t("name")), "t7")
    End If

    On Error Resume Next
    Kill p
    On Error GoTo 0
    Rdv3CfgDefaults
    If Len(bad) > 0 Then
        Rdv3SetSelfTest = "FAIL" & bad
    Else
        Rdv3SetSelfTest = "ok sheet round trip keeps unshown members and the 7th target"
    End If
    Exit Function
Failed:
    Rdv3SetSelfTest = "FAIL " & Err.Number & ": " & Err.Description
End Function

' THE OTHER round trip: 「この要素を使う」 onto a row that already describes a
' DIFFERENT element. Every matcher must come from the picked element and none of
' the old one may survive -- a leftover requireValuePattern or nameLike describes
' something that is no longer there, and the target then never binds. The rows
' that were NOT re-picked must still keep everything, so both properties are
' asserted in the same pass.
Public Function Rdv3SetAdoptSelfTest() As String
    Dim p As String
    Dim s As String
    Dim t As Object
    Dim bad As String
    Dim err1 As String

    On Error GoTo Failed
    p = Environ$("TEMP") & "\rdv3adopt-selftest.json"
    s = "{ ""watch"": { ""targets"": [" & _
        "{ ""enabled"": true, ""name"": ""old""," & _
        "  ""window"": { ""automationId"": ""oldWin"", ""className"": ""OldCls"", ""name"": ""oldExact""," & _
        "                ""nameLike"": ""old*"", ""index"": 2, ""scope"": ""children"" }," & _
        "  ""path"": [ { ""automationId"": ""oldTab"" } ]," & _
        "  ""field"": { ""automationId"": ""oldTxt"", ""className"": ""OldEdit"", ""nameLike"": ""oldbar*""," & _
        "               ""processName"": ""OldApp"", ""requireValuePattern"": true," & _
        "               ""controlTypes"": [""Edit""], ""index"": 5 }," & _
        "  ""read"": ""value"" }," & _
        "{ ""enabled"": true, ""name"": ""keep"", ""window"": { ""className"": ""KeepCls"", ""name"": ""keepExact"" }," & _
        "  ""field"": { ""className"": ""KeepEdit"", ""nameLike"": ""keep*"", ""controlTypes"": [""Edit""] } }" & _
        "] } }"
    If Not Rdv3CfgWriteRawForTest(p, s) Then
        Rdv3SetAdoptSelfTest = "FAIL cannot write the test file"
        Exit Function
    End If
    If Not Rdv3CfgLoad(p) Then
        Rdv3SetAdoptSelfTest = "FAIL parse: " & Rdv3CfgError()
        Exit Function
    End If

    Rdv3SetLoad
    PutSet "setPickRow", "1"
    ' a picked element that shares nothing with the row's current contents
    Rdv3SetInspectExpect "SELFTEST"
    Rdv3SetInspectResult "res=preview;req=SELFTEST;cls=NewCls;proc=NewApp;wname=NewWindow;" & _
        "wid=newWin;fid=;fcls=NewEdit;fname=;ftype=Edit;fidx=3;path=pnlA/pnlB;read=text"
    If m_pv Is Nothing Then
        Rdv3SetAdoptSelfTest = "FAIL the preview was not accepted"
        Exit Function
    End If
    Rdv3SetAdoptPreview
    err1 = Rdv3SetSave()
    If Len(err1) > 0 Then
        Rdv3SetAdoptSelfTest = "FAIL save: " & err1
        Exit Function
    End If

    Rdv3CfgDefaults
    Rdv3CfgLoad p
    bad = bad & ExpectSet("targets", CStr(Rdv3CfgTargetCount()), "2")
    If Rdv3CfgTargetCount() >= 2 Then
        Set t = Rdv3CfgTarget(1)
        ' what the picked element says
        bad = bad & ExpectSet("win id", CStr(t("window")("automationId")), "newWin")
        bad = bad & ExpectSet("win cls", CStr(t("window")("className")), "NewCls")
        bad = bad & ExpectSet("win proc", CStr(t("window")("processName")), "NewApp")
        bad = bad & ExpectSet("fld cls", CStr(t("field")("className")), "NewEdit")
        bad = bad & ExpectSet("fld idx", CStr(CLng(t("field")("index"))), "3")
        bad = bad & ExpectSet("read", CStr(t("read")), "text")
        bad = bad & ExpectSet("path", CStr(t("path").Count), "2")
        ' and NONE of what the old element said
        bad = bad & ExpectSet("win name gone", CStr(t("window")("name")), "")
        bad = bad & ExpectSet("win like gone", CStr(t("window")("nameLike")), "")
        bad = bad & ExpectSet("win index gone", CStr(CLng(t("window")("index"))), "0")
        bad = bad & ExpectSet("fld id gone", CStr(t("field")("automationId")), "")
        bad = bad & ExpectSet("fld like gone", CStr(t("field")("nameLike")), "")
        bad = bad & ExpectSet("fld proc gone", CStr(t("field")("processName")), "")
        bad = bad & ExpectSet("value req gone", CStr(CBool(t("field")("requireValuePattern"))), CStr(False))
        ' the row that was NOT re-picked still carries everything it had
        Set t = Rdv3CfgTarget(2)
        bad = bad & ExpectSet("keep name", CStr(t("window")("name")), "keepExact")
        bad = bad & ExpectSet("keep like", CStr(t("field")("nameLike")), "keep*")
    End If

    On Error Resume Next
    Kill p
    On Error GoTo 0
    Rdv3CfgDefaults
    Rdv3SetLoad
    PutSet "setNote", ""
    PutSet "setPvNote", ""
    If Len(bad) > 0 Then
        Rdv3SetAdoptSelfTest = "FAIL" & bad
    Else
        Rdv3SetAdoptSelfTest = "ok adopt replaces every matcher and leaves the other rows whole"
    End If
    Exit Function
Failed:
    Rdv3SetAdoptSelfTest = "FAIL " & Err.Number & ": " & Err.Description
End Function

Private Function ExpectSet(ByVal what As String, ByVal got As String, ByVal want As String) As String
    If got <> want Then ExpectSet = " [" & what & "=" & got & " want " & want & "]"
End Function
