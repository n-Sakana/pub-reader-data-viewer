Attribute VB_Name = "modRdv3Cfg"
'==============================================================================
' modRdv3Cfg -- ReaderDataViewer.json, read and written by the VBA build.
'
' THE SAME FILE THE C# BUILD READS, not a second format of its own. There is one
' settings document (docs\settings.md), one set of member names and one set of
' rules; a build that invented its own file would leave this product with two
' settings formats forever, and "the VBA build can be configured like the C# one"
' would stop being true the first time either side gained a member.
'
' So this module is the VBA half of Rdv3Json.cs + Rdv3Config.cs, and it has to
' agree with them on every answer:
'
'   - JSONC: // to the end of the line and /* ... */ anywhere outside a string
'     are comments (same as Rdv3Json.SkipWhite)
'   - a missing file is not an error: the built-in defaults ARE the shipped
'     ReaderDataViewer.json, and the log says which was used
'   - a file that cannot be parsed is reported and the defaults are used, because
'     a settings typo must not stop the operator working
'   - a single value out of range falls back on ITS OWN default, not on the
'     nearest limit, and says so in the notes
'   - watch.targets is watched as written, down to "watch nothing"; a target that
'     is switched off is kept, and only one that could match ANY window is
'     dropped
'
' NO WIN32 AND NO SHELL, and no new COM dependency either: UTF-8 is encoded and
' decoded here, by hand. ADODB.Stream would do it, but the one measurement this
' repository has of it is 22.2 seconds for the 45 MB sidecar (modRdv3Be), and
' rather than reason about where that curve starts to matter for a 4 KB file the
' conversion is simply written out -- it is thirty lines and it cannot surprise
' anyone.
'
' Shapes used here, so the rest of the build can read them without a class:
'   matcher  Dictionary: automationId, className, name, nameLike, processName,
'            controlTypes (Collection of String), requireValuePattern, index,
'            descendants
'   target   Dictionary: enabled, name, window (matcher), path (Collection of
'            matcher), field (matcher), read ("value"/"text"/"name")
'==============================================================================
Option Explicit

Public Const RDV3_READ_VALUE As Long = 0
Public Const RDV3_READ_TEXT As Long = 1
Public Const RDV3_READ_NAME As Long = 2
Public Const RDV3_CFG_SCHEMA As Long = 1

' ---- what the file says (defaults are the shipped file's values) ------------
Private m_dataDir As String
Private m_ledger As String
Private m_logPath As String
Private m_keyLength As Long
Private m_keyDigitsOnly As Boolean
Private m_pollMs As Long
Private m_stableMs As Long
Private m_rebindMs As Long
Private m_preferFocused As Boolean
Private m_candidateRows As Long
Private m_checkTimeoutMs As Long
Private m_searchTimeoutMs As Long
Private m_saveTimeoutMs As Long
Private m_markOverdueMs As Long
Private m_pumpMs As Long
Private m_targets As Collection

Private m_sourcePath As String
Private m_loaded As Boolean
Private m_error As String
Private m_notes As Collection
Private m_ready As Boolean

' ---- parser state ----------------------------------------------------------
Private m_src As String
Private m_pos As Long
Private m_line As Long
' where the reader had got to. A settings file is read before there is anywhere
' to log to, so when it fails the message is the only evidence there will be:
' this makes it name the member it was on instead of just "type mismatch".
Private m_stage As String

' ---- control type names UI Automation knows --------------------------------
Private m_typeIds As Object
Private m_typeNames As Object

'==============================================================================
' the values
'==============================================================================
Public Function Rdv3CfgDataDir() As String
    EnsureReady
    Rdv3CfgDataDir = m_dataDir
End Function

Public Function Rdv3CfgLedger() As String
    EnsureReady
    Rdv3CfgLedger = m_ledger
End Function

Public Function Rdv3CfgLog() As String
    EnsureReady
    Rdv3CfgLog = m_logPath
End Function

Public Function Rdv3CfgKeyLength() As Long
    EnsureReady
    Rdv3CfgKeyLength = m_keyLength
End Function

Public Function Rdv3CfgKeyDigitsOnly() As Boolean
    EnsureReady
    Rdv3CfgKeyDigitsOnly = m_keyDigitsOnly
End Function

Public Function Rdv3CfgPollMs() As Long
    EnsureReady
    Rdv3CfgPollMs = m_pollMs
End Function

Public Function Rdv3CfgStableMs() As Long
    EnsureReady
    Rdv3CfgStableMs = m_stableMs
End Function

Public Function Rdv3CfgRebindMs() As Long
    EnsureReady
    Rdv3CfgRebindMs = m_rebindMs
End Function

Public Function Rdv3CfgPreferFocused() As Boolean
    EnsureReady
    Rdv3CfgPreferFocused = m_preferFocused
End Function

Public Function Rdv3CfgCandidateRows() As Long
    EnsureReady
    Rdv3CfgCandidateRows = m_candidateRows
End Function

Public Function Rdv3CfgCheckTimeoutMs() As Long
    EnsureReady
    Rdv3CfgCheckTimeoutMs = m_checkTimeoutMs
End Function

Public Function Rdv3CfgSearchTimeoutMs() As Long
    EnsureReady
    Rdv3CfgSearchTimeoutMs = m_searchTimeoutMs
End Function

Public Function Rdv3CfgSaveTimeoutMs() As Long
    EnsureReady
    Rdv3CfgSaveTimeoutMs = m_saveTimeoutMs
End Function

Public Function Rdv3CfgMarkOverdueMs() As Long
    EnsureReady
    Rdv3CfgMarkOverdueMs = m_markOverdueMs
End Function

Public Function Rdv3CfgPumpMs() As Long
    EnsureReady
    Rdv3CfgPumpMs = m_pumpMs
End Function

Public Function Rdv3CfgSourcePath() As String
    Rdv3CfgSourcePath = m_sourcePath
End Function

Public Function Rdv3CfgLoaded() As Boolean
    Rdv3CfgLoaded = m_loaded
End Function

Public Function Rdv3CfgError() As String
    Rdv3CfgError = m_error
End Function

Public Function Rdv3CfgNoteCount() As Long
    If m_notes Is Nothing Then Exit Function
    Rdv3CfgNoteCount = m_notes.Count
End Function

Public Function Rdv3CfgNote(ByVal i As Long) As String
    If m_notes Is Nothing Then Exit Function
    If i < 1 Or i > m_notes.Count Then Exit Function
    Rdv3CfgNote = CStr(m_notes.Item(i))
End Function

Public Function Rdv3CfgTargetCount() As Long
    EnsureReady
    Rdv3CfgTargetCount = m_targets.Count
End Function

Public Function Rdv3CfgTarget(ByVal i As Long) As Object
    EnsureReady
    If i < 1 Or i > m_targets.Count Then Exit Function
    Set Rdv3CfgTarget = m_targets.Item(i)
End Function

'==============================================================================
' setters, for the settings sheet. They change what is IN MEMORY; the file is
' only touched by Rdv3CfgSave.
'==============================================================================
Public Sub Rdv3CfgSetPaths(ByVal dataDir As String, ByVal ledger As String, ByVal logPath As String)
    EnsureReady
    m_dataDir = dataDir
    m_ledger = ledger
    m_logPath = logPath
End Sub

Public Sub Rdv3CfgSetKey(ByVal keyLen As Long, ByVal digitsOnly As Boolean)
    EnsureReady
    m_keyLength = Clip(keyLen, 1, 64)
    m_keyDigitsOnly = digitsOnly
End Sub

Public Sub Rdv3CfgSetWatch(ByVal pollMs As Long, ByVal stableMs As Long, _
                           ByVal rebindMs As Long, ByVal preferFocused As Boolean)
    EnsureReady
    m_pollMs = Clip(pollMs, 5, 5000)
    m_stableMs = Clip(stableMs, 0, 60000)
    m_rebindMs = Clip(rebindMs, 50, 60000)
    m_preferFocused = preferFocused
End Sub

Public Sub Rdv3CfgSetCandidateRows(ByVal n As Long)
    EnsureReady
    m_candidateRows = Clip(n, 1, 1000)
End Sub

Public Sub Rdv3CfgSetTargets(ByVal list As Collection)
    EnsureReady
    If list Is Nothing Then Exit Sub
    Set m_targets = list
End Sub

' the settings sheet edits a copy, so a cancel costs nothing
Public Function Rdv3CfgCloneTargets() As Collection
    Dim out As Collection
    Dim i As Long
    EnsureReady
    Set out = New Collection
    For i = 1 To m_targets.Count
        out.Add CloneTarget(m_targets.Item(i))
    Next i
    Set Rdv3CfgCloneTargets = out
End Function

' a deep copy of ONE target, for an editor that must not lose what it does not
' show
Public Function Rdv3CfgCloneTarget(ByVal t As Object) As Object
    If t Is Nothing Then Exit Function
    Set Rdv3CfgCloneTarget = CloneTarget(t)
End Function

Private Function CloneTarget(ByVal t As Object) As Object
    Dim c As Object
    Dim steps As Collection
    Dim src As Collection
    Dim i As Long
    Set c = Rdv3CfgNewTarget()
    c("enabled") = t("enabled")
    c("name") = t("name")
    Set c("window") = CloneMatch(t("window"))
    Set c("field") = CloneMatch(t("field"))
    c("read") = t("read")
    Set steps = New Collection
    Set src = t("path")
    For i = 1 To src.Count
        steps.Add CloneMatch(src.Item(i))
    Next i
    Set c("path") = steps
    Set CloneTarget = c
End Function

Private Function CloneMatch(ByVal m As Object) As Object
    Dim c As Object
    Dim types As Collection
    Dim src As Collection
    Dim i As Long
    Set c = NewMatch()
    c("automationId") = m("automationId")
    c("className") = m("className")
    c("name") = m("name")
    c("nameLike") = m("nameLike")
    c("processName") = m("processName")
    c("requireValuePattern") = m("requireValuePattern")
    c("index") = m("index")
    c("descendants") = m("descendants")
    Set types = New Collection
    Set src = m("controlTypes")
    For i = 1 To src.Count
        types.Add CStr(src.Item(i))
    Next i
    Set c("controlTypes") = types
    Set CloneMatch = c
End Function

'==============================================================================
' the shapes
'==============================================================================
Public Function Rdv3CfgNewMatch() As Object
    Set Rdv3CfgNewMatch = NewMatch()
End Function

Public Function Rdv3CfgNewTarget() As Object
    Dim t As Object
    Set t = NewDict()
    t("enabled") = True
    t("name") = ""
    Set t("window") = NewMatch()
    Set t("field") = NewMatch()
    Set t("path") = New Collection
    t("read") = "value"
    t("window")("descendants") = False
    Set Rdv3CfgNewTarget = t
End Function

Private Function NewMatch() As Object
    Dim m As Object
    Set m = NewDict()
    m("automationId") = ""
    m("className") = ""
    m("name") = ""
    m("nameLike") = ""
    m("processName") = ""
    Set m("controlTypes") = New Collection
    m("requireValuePattern") = False
    m("index") = 0
    m("descendants") = True
    Set NewMatch = m
End Function

Private Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                      ' text: the C# reader is case-insensitive
    Set NewDict = d
End Function

'==============================================================================
' validity -- the same two questions Rdv3Match.IsEmpty and
' Rdv3Target.WhyNotWatchable answer in the C# build
'==============================================================================

' How many of the listed control types UI Automation actually knows. One it does
' not know is dropped from the search, so counting it as a constraint would let a
' typo -- "Edti" for "Edit" -- pass for a narrowed matcher and leave one that in
' fact accepts every window there is.
Public Function Rdv3CfgKnownTypeCount(ByVal m As Object) As Long
    Dim types As Collection
    Dim i As Long
    Dim n As Long
    If m Is Nothing Then Exit Function
    Set types = m("controlTypes")
    For i = 1 To types.Count
        If Rdv3CfgControlTypeId(CStr(types.Item(i))) <> 0 Then n = n + 1
    Next i
    Rdv3CfgKnownTypeCount = n
End Function

Public Function Rdv3CfgMatchIsEmpty(ByVal m As Object) As Boolean
    If m Is Nothing Then
        Rdv3CfgMatchIsEmpty = True
        Exit Function
    End If
    If Len(CStr(m("automationId"))) > 0 Then Exit Function
    If Len(CStr(m("className"))) > 0 Then Exit Function
    If Len(CStr(m("name"))) > 0 Then Exit Function
    If Len(CStr(m("nameLike"))) > 0 Then Exit Function
    If Len(CStr(m("processName"))) > 0 Then Exit Function
    If CBool(m("requireValuePattern")) Then Exit Function
    If Rdv3CfgKnownTypeCount(m) > 0 Then Exit Function
    Rdv3CfgMatchIsEmpty = True
End Function

Public Function Rdv3CfgWhyNotWatchable(ByVal t As Object) As String
    If t Is Nothing Then
        Rdv3CfgWhyNotWatchable = "対象がありません"
        Exit Function
    End If
    If Rdv3CfgMatchIsEmpty(t("window")) Then
        Rdv3CfgWhyNotWatchable = "ウィンドウの条件が空で、どの画面にも一致します"
        Exit Function
    End If
    If Rdv3CfgMatchIsEmpty(t("field")) Then
        Rdv3CfgWhyNotWatchable = "欄の条件が空で、どの欄にも一致します"
        Exit Function
    End If
End Function

Public Function Rdv3CfgWatchable(ByVal t As Object) As Boolean
    If t Is Nothing Then Exit Function
    If Not CBool(t("enabled")) Then Exit Function
    Rdv3CfgWatchable = (Len(Rdv3CfgWhyNotWatchable(t)) = 0)
End Function

Public Function Rdv3CfgWatchableCount() As Long
    Dim i As Long
    Dim n As Long
    EnsureReady
    For i = 1 To m_targets.Count
        If Rdv3CfgWatchable(m_targets.Item(i)) Then n = n + 1
    Next i
    Rdv3CfgWatchableCount = n
End Function

Public Function Rdv3CfgReadModeId(ByVal s As String) As Long
    Select Case LCase$(Trim$(s))
        Case "text"
            Rdv3CfgReadModeId = RDV3_READ_TEXT
        Case "name"
            Rdv3CfgReadModeId = RDV3_READ_NAME
        Case Else
            Rdv3CfgReadModeId = RDV3_READ_VALUE
    End Select
End Function

Public Function Rdv3CfgReadModeName(ByVal id As Long) As String
    Select Case id
        Case RDV3_READ_TEXT
            Rdv3CfgReadModeName = "text"
        Case RDV3_READ_NAME
            Rdv3CfgReadModeName = "name"
        Case Else
            Rdv3CfgReadModeName = "value"
    End Select
End Function

'==============================================================================
' control type names -> UI Automation control type ids. The same 39 names the
' C# build accepts (Rdv3Uia.Types), so the settings file means the same thing
' on both sides. 0 = a name UI Automation does not know.
'==============================================================================
Public Function Rdv3CfgControlTypeId(ByVal nm As String) As Long
    Dim k As String
    If m_typeIds Is Nothing Then BuildTypes
    k = LCase$(Trim$(nm))
    If Len(k) = 0 Then Exit Function
    If m_typeIds.Exists(k) Then Rdv3CfgControlTypeId = CLng(m_typeIds(k))
End Function

Public Function Rdv3CfgControlTypeName(ByVal id As Long) As String
    If m_typeNames Is Nothing Then BuildTypes
    If m_typeNames.Exists(CStr(id)) Then Rdv3CfgControlTypeName = CStr(m_typeNames(CStr(id)))
End Function

Private Sub BuildTypes()
    Set m_typeIds = CreateObject("Scripting.Dictionary")
    Set m_typeNames = CreateObject("Scripting.Dictionary")
    AddType "Button", 50000
    AddType "Calendar", 50001
    AddType "CheckBox", 50002
    AddType "ComboBox", 50003
    AddType "Edit", 50004
    AddType "Hyperlink", 50005
    AddType "Image", 50006
    AddType "ListItem", 50007
    AddType "List", 50008
    AddType "Menu", 50009
    AddType "MenuBar", 50010
    AddType "MenuItem", 50011
    AddType "ProgressBar", 50012
    AddType "RadioButton", 50013
    AddType "ScrollBar", 50014
    AddType "Slider", 50015
    AddType "Spinner", 50016
    AddType "StatusBar", 50017
    AddType "Tab", 50018
    AddType "TabItem", 50019
    AddType "Text", 50020
    AddType "ToolBar", 50021
    AddType "ToolTip", 50022
    AddType "Tree", 50023
    AddType "TreeItem", 50024
    AddType "Custom", 50025
    AddType "Group", 50026
    AddType "Thumb", 50027
    AddType "DataGrid", 50028
    AddType "DataItem", 50029
    AddType "Document", 50030
    AddType "SplitButton", 50031
    AddType "Window", 50032
    AddType "Pane", 50033
    AddType "Header", 50034
    AddType "HeaderItem", 50035
    AddType "Table", 50036
    AddType "TitleBar", 50037
    AddType "Separator", 50038
End Sub

Private Sub AddType(ByVal nm As String, ByVal id As Long)
    m_typeIds(LCase$(nm)) = id
    m_typeNames(CStr(id)) = nm
End Sub

'==============================================================================
' defaults -- exactly what the shipped ReaderDataViewer.json says
'==============================================================================
Public Sub Rdv3CfgDefaults()
    Dim t As Object
    Dim types As Collection
    m_dataDir = "data"
    m_ledger = "ReaderDataViewer-Ledger.xlsx"
    m_logPath = "ReaderDataViewer.log"
    m_keyLength = 8
    m_keyDigitsOnly = True
    m_pollMs = 40
    m_stableMs = 120
    m_rebindMs = 400
    m_preferFocused = True
    m_candidateRows = 10
    m_checkTimeoutMs = 180000
    m_searchTimeoutMs = 30000
    m_saveTimeoutMs = 60000
    m_markOverdueMs = 180000
    m_pumpMs = 1000
    m_sourcePath = ""
    m_loaded = False
    m_error = ""
    Set m_notes = New Collection
    Set m_targets = New Collection
    m_ready = True
    Set t = Rdv3CfgNewTarget()
    t("name") = "メモ帳"
    t("window")("className") = "Notepad"
    Set types = t("field")("controlTypes")
    types.Add "Document"
    types.Add "Edit"
    t("field")("requireValuePattern") = True
    m_targets.Add t
End Sub

Private Sub EnsureReady()
    If Not m_ready Then Rdv3CfgDefaults
End Sub

Private Sub Note(ByVal s As String)
    If m_notes Is Nothing Then Set m_notes = New Collection
    m_notes.Add s
End Sub

Private Function Clip(ByVal v As Long, ByVal lo As Long, ByVal hi As Long) As Long
    If v < lo Then
        Clip = lo
    ElseIf v > hi Then
        Clip = hi
    Else
        Clip = v
    End If
End Function

' Out of range means THIS member falls back on its own default, which is what
' the shipped file and docs\settings.md promise. The nearest limit would answer a
' different question: a typed "0" for jobs.checkTimeoutMs is not a request for
' the shortest legal timeout.
Private Function Ranged(ByVal v As Long, ByVal lo As Long, ByVal hi As Long, _
                        ByVal def As Long, ByVal what As String) As Long
    If v < lo Or v > hi Then
        Note what & " " & CStr(v) & " は範囲 [" & CStr(lo) & ".." & CStr(hi) & _
             "] の外です。既定値 " & CStr(def) & " を使います"
        Ranged = def
    Else
        Ranged = v
    End If
End Function

'==============================================================================
' load
'==============================================================================
Public Function Rdv3CfgLoad(ByVal path As String) As Boolean
    Dim root As Object
    Dim raw As String

    Rdv3CfgDefaults
    m_sourcePath = path
    If Len(path) = 0 Then Exit Function
    If Len(Dir$(path)) = 0 Then Exit Function

    On Error GoTo Failed
    m_stage = "read"
    raw = ReadUtf8(path)
    m_stage = "parse"
    Set root = ParseJsonc(raw)
    On Error GoTo 0
    If root Is Nothing Then
        m_error = "最上位がオブジェクトではありません"
        Exit Function
    End If

    On Error GoTo Failed
    ApplyRoot root
    On Error GoTo 0
    m_loaded = True
    Rdv3CfgLoad = True
    Exit Function
Failed:
    m_error = Err.Description & " [" & m_stage & Where() & "]"
End Function

' where the reader had got to, for the message. A settings file is read before
' there is anywhere to log to, so this message is the only evidence there will be
' -- "type mismatch" on its own names nothing an operator could act on.
Private Function Where() As String
    Dim from As Long
    On Error Resume Next
    If m_stage <> "parse" Then Exit Function
    from = m_pos - 20
    If from < 1 Then from = 1
    Where = " line " & CStr(m_line) & " at " & _
            Replace(Replace(Mid$(m_src, from, 40), vbCr, " "), vbLf, " ")
    On Error GoTo 0
End Function

Private Sub ApplyRoot(ByVal root As Object)
    Dim o As Object
    Dim schema As Long

    schema = GetLong(root, "schema", RDV3_CFG_SCHEMA)
    If schema <> RDV3_CFG_SCHEMA Then
        Note "schema " & CStr(schema) & " はこのビルドが知る " & CStr(RDV3_CFG_SCHEMA) & _
             " ではありません (そのまま読みます)"
    End If

    m_stage = "paths"
    Set o = GetObj(root, "paths")
    If Not o Is Nothing Then
        m_dataDir = GetStr(o, "dataDir", m_dataDir)
        m_ledger = GetStr(o, "ledger", m_ledger)
        m_logPath = GetStr(o, "log", m_logPath)
    End If

    m_stage = "key"
    Set o = GetObj(root, "key")
    If Not o Is Nothing Then
        m_keyLength = Ranged(GetLong(o, "length", m_keyLength), 1, 64, 8, "key.length")
        m_keyDigitsOnly = GetBool(o, "digitsOnly", m_keyDigitsOnly)
    End If

    m_stage = "search"
    Set o = GetObj(root, "search")
    If Not o Is Nothing Then
        m_candidateRows = Ranged(GetLong(o, "candidateRowsShown", m_candidateRows), _
                                 1, 1000, 10, "search.candidateRowsShown")
    End If

    m_stage = "jobs"
    Set o = GetObj(root, "jobs")
    If Not o Is Nothing Then
        m_checkTimeoutMs = Ranged(GetLong(o, "checkTimeoutMs", m_checkTimeoutMs), 1000, 3600000, 180000, "jobs.checkTimeoutMs")
        m_searchTimeoutMs = Ranged(GetLong(o, "searchTimeoutMs", m_searchTimeoutMs), 1000, 3600000, 30000, "jobs.searchTimeoutMs")
        m_saveTimeoutMs = Ranged(GetLong(o, "saveTimeoutMs", m_saveTimeoutMs), 1000, 3600000, 60000, "jobs.saveTimeoutMs")
        m_markOverdueMs = Ranged(GetLong(o, "markOverdueMs", m_markOverdueMs), 1000, 3600000, 180000, "jobs.markOverdueMs")
        m_pumpMs = Ranged(GetLong(o, "pumpMs", m_pumpMs), 100, 60000, 1000, "jobs.pumpMs")
    End If

    m_stage = "watch"
    Set o = GetObj(root, "watch")
    If Not o Is Nothing Then ApplyWatch o
End Sub

Private Sub ApplyWatch(ByVal w As Object)
    Dim arr As Collection
    Dim list As Collection
    Dim t As Object
    Dim why As String
    Dim i As Long

    m_pollMs = Ranged(GetLong(w, "pollMs", m_pollMs), 5, 5000, 40, "watch.pollMs")
    m_stableMs = Ranged(GetLong(w, "stableMs", m_stableMs), 0, 60000, 120, "watch.stableMs")
    m_rebindMs = Ranged(GetLong(w, "rebindMs", m_rebindMs), 50, 60000, 400, "watch.rebindMs")
    m_preferFocused = GetBool(w, "preferFocusedWindow", m_preferFocused)

    m_stage = "watch.targets"
    Set arr = GetArr(w, "targets")
    If arr Is Nothing Then Exit Sub

    ' The file answered the question, so the file decides -- down to "watch
    ' nothing". A target the operator turned OFF is kept: it is still theirs, the
    ' settings sheet still lists it and the next save still writes it. Only one
    ' that cannot work at all is dropped.
    Set list = New Collection
    For i = 1 To arr.Count
        m_stage = "targets[" & CStr(i) & "]"
        Set t = ReadTarget(arr.Item(i), i)
        why = Rdv3CfgWhyNotWatchable(t)
        If Len(why) > 0 Then
            Note "watch.targets[" & CStr(i - 1) & "]: " & why & "。この対象は監視しません"
        Else
            list.Add t
        End If
    Next i
    If list.Count = 0 Then
        Note "watch.targets に監視できる対象がありません。何も監視しません"
    End If
    Set m_targets = list
End Sub

Private Function ReadTarget(ByRef v As Variant, ByVal at As Long) As Object
    Dim t As Object
    Dim o As Object
    Dim arr As Collection
    Dim stp As Object
    Dim where As String
    Dim i As Long

    Set t = Rdv3CfgNewTarget()
    where = "watch.targets[" & CStr(at - 1) & "]"
    If Not IsObject(v) Then
        Note where & " はオブジェクトではありません"
        Set ReadTarget = t
        Exit Function
    End If
    If v Is Nothing Then
        Set ReadTarget = t
        Exit Function
    End If
    If TypeName(v) <> "Dictionary" Then
        Note where & " はオブジェクトではありません"
        Set ReadTarget = t
        Exit Function
    End If
    Set o = v
    t("enabled") = GetBool(o, "enabled", True)
    t("name") = GetStr(o, "name", "")
    m_stage = m_stage & ".window"
    Set t("window") = ReadMatch(GetObj(o, "window"), where & ".window")
    t("window")("descendants") = False           ' top level windows sit under the desktop
    m_stage = m_stage & "/field"
    Set t("field") = ReadMatch(GetObj(o, "field"), where & ".field")
    t("read") = Rdv3CfgReadModeName(Rdv3CfgReadModeId(GetStr(o, "read", "value")))

    Set arr = GetArr(o, "path")
    If Not arr Is Nothing Then
        For i = 1 To arr.Count
            Set stp = ReadMatchItem(arr, i, where & ".path[" & CStr(i - 1) & "]")
            If Rdv3CfgMatchIsEmpty(stp) Then
                Note where & ".path[" & CStr(i - 1) & "] は何も絞っていません。この段は飛ばします"
            Else
                t("path").Add stp
            End If
        Next i
    End If

    If Len(CStr(t("name"))) = 0 Then
        If Len(CStr(t("window")("processName"))) > 0 Then
            t("name") = t("window")("processName")
        Else
            t("name") = t("window")("className")
        End If
    End If
    Set ReadTarget = t
End Function

Private Function ReadMatchItem(ByVal arr As Collection, ByVal i As Long, _
                               ByVal where As String) As Object
    Dim o As Object
    If IsObject(arr.Item(i)) Then
        If TypeName(arr.Item(i)) = "Dictionary" Then Set o = arr.Item(i)
    End If
    Set ReadMatchItem = ReadMatch(o, where)
End Function

Private Function ReadMatch(ByVal o As Object, ByVal where As String) As Object
    Dim m As Object
    Dim arr As Collection
    Dim scope As String
    Dim nm As String
    Dim i As Long

    Set m = NewMatch()
    If o Is Nothing Then
        Set ReadMatch = m
        Exit Function
    End If
    m("automationId") = GetStr(o, "automationId", "")
    m("className") = GetStr(o, "className", "")
    m("name") = GetStr(o, "name", "")
    m("nameLike") = GetStr(o, "nameLike", "")
    m("processName") = GetStr(o, "processName", "")
    m("requireValuePattern") = GetBool(o, "requireValuePattern", False)
    m("index") = Clip(GetLong(o, "index", 0), 0, 100000)

    scope = LCase$(Trim$(GetStr(o, "scope", "")))
    If scope = "children" Then
        m("descendants") = False
    ElseIf scope = "descendants" Or Len(scope) = 0 Then
        m("descendants") = True
    Else
        Note where & ": scope " & scope & " は children でも descendants でもありません"
    End If

    Set arr = GetArr(o, "controlTypes")
    If Not arr Is Nothing Then
        For i = 1 To arr.Count
            If Not IsObject(arr.Item(i)) Then
                nm = CStr(arr.Item(i))
                m("controlTypes").Add nm
                If Rdv3CfgControlTypeId(nm) = 0 Then
                    Note where & ": controlType " & nm & " は UI Automation が知らない名前です"
                End If
            End If
        Next i
    End If
    Set ReadMatch = m
End Function

'==============================================================================
' save -- the same layout Rdv3Config.Save writes, so the file stays one format.
' Comments a person added are not kept (the reader accepts them; the writer does
' not invent them), so the header says who wrote it and when.
'==============================================================================
Public Function Rdv3CfgSave(ByVal path As String) As String
    Dim sb As String
    Dim t As Object
    Dim steps As Collection
    Dim i As Long
    Dim k As Long

    EnsureReady
    On Error GoTo Failed
    sb = "// Reader Data Viewer settings" & vbCrLf
    sb = sb & "// written by the app on " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    sb = sb & "// what each member means: docs/settings.md" & vbCrLf
    sb = sb & "{" & vbCrLf
    sb = sb & "  ""schema"": " & CStr(RDV3_CFG_SCHEMA) & "," & vbCrLf
    sb = sb & "  ""paths"": {" & vbCrLf
    sb = sb & "    ""dataDir"": " & JsonStr(m_dataDir) & "," & vbCrLf
    sb = sb & "    ""ledger"": " & JsonStr(m_ledger) & "," & vbCrLf
    sb = sb & "    ""log"": " & JsonStr(m_logPath) & vbCrLf
    sb = sb & "  }," & vbCrLf
    sb = sb & "  ""key"": { ""length"": " & CStr(m_keyLength) & ", ""digitsOnly"": " & _
         JsonBool(m_keyDigitsOnly) & " }," & vbCrLf
    sb = sb & "  ""search"": { ""candidateRowsShown"": " & CStr(m_candidateRows) & " }," & vbCrLf
    sb = sb & "  ""jobs"": {" & vbCrLf
    sb = sb & "    ""checkTimeoutMs"": " & CStr(m_checkTimeoutMs) & "," & vbCrLf
    sb = sb & "    ""searchTimeoutMs"": " & CStr(m_searchTimeoutMs) & "," & vbCrLf
    sb = sb & "    ""saveTimeoutMs"": " & CStr(m_saveTimeoutMs) & "," & vbCrLf
    sb = sb & "    ""markOverdueMs"": " & CStr(m_markOverdueMs) & "," & vbCrLf
    sb = sb & "    ""pumpMs"": " & CStr(m_pumpMs) & vbCrLf
    sb = sb & "  }," & vbCrLf
    sb = sb & "  ""watch"": {" & vbCrLf
    sb = sb & "    ""pollMs"": " & CStr(m_pollMs) & ", ""stableMs"": " & CStr(m_stableMs)
    sb = sb & ", ""rebindMs"": " & CStr(m_rebindMs) & ", ""preferFocusedWindow"": " & _
         JsonBool(m_preferFocused) & "," & vbCrLf
    sb = sb & "    ""targets"": [" & vbCrLf
    For i = 1 To m_targets.Count
        Set t = m_targets.Item(i)
        sb = sb & "      {" & vbCrLf
        sb = sb & "        ""enabled"": " & JsonBool(CBool(t("enabled"))) & "," & vbCrLf
        sb = sb & "        ""name"": " & JsonStr(CStr(t("name"))) & "," & vbCrLf
        sb = sb & "        ""window"": " & MatchJson(t("window")) & "," & vbCrLf
        sb = sb & "        ""path"": ["
        Set steps = t("path")
        For k = 1 To steps.Count
            If k > 1 Then sb = sb & ","
            sb = sb & vbCrLf & "          " & MatchJson(steps.Item(k))
        Next k
        If steps.Count > 0 Then
            sb = sb & vbCrLf & "        ]"
        Else
            sb = sb & "]"
        End If
        sb = sb & "," & vbCrLf
        sb = sb & "        ""field"": " & MatchJson(t("field")) & "," & vbCrLf
        sb = sb & "        ""read"": " & JsonStr(CStr(t("read"))) & vbCrLf
        If i < m_targets.Count Then
            sb = sb & "      }," & vbCrLf
        Else
            sb = sb & "      }" & vbCrLf
        End If
    Next i
    sb = sb & "    ]" & vbCrLf
    sb = sb & "  }" & vbCrLf
    sb = sb & "}" & vbCrLf

    Rdv3CfgSave = WriteUtf8Atomic(path, sb)
    If Len(Rdv3CfgSave) = 0 Then
        m_sourcePath = path
        m_loaded = True
        m_error = ""
    End If
    Exit Function
Failed:
    Rdv3CfgSave = "設定の書き出しに失敗 " & Err.Number & ": " & Err.Description
End Function

Private Function MatchJson(ByVal m As Object) As String
    Dim s As String
    Dim types As Collection
    Dim i As Long
    s = "{ "
    s = s & Pair(s, "automationId", CStr(m("automationId")))
    s = s & Pair(s, "className", CStr(m("className")))
    s = s & Pair(s, "name", CStr(m("name")))
    s = s & Pair(s, "nameLike", CStr(m("nameLike")))
    s = s & Pair(s, "processName", CStr(m("processName")))
    Set types = m("controlTypes")
    If types.Count > 0 Then
        If Len(s) > 2 Then s = s & ", "
        s = s & """controlTypes"": ["
        For i = 1 To types.Count
            If i > 1 Then s = s & ", "
            s = s & JsonStr(CStr(types.Item(i)))
        Next i
        s = s & "]"
    End If
    If CBool(m("requireValuePattern")) Then
        If Len(s) > 2 Then s = s & ", "
        s = s & """requireValuePattern"": true"
    End If
    If CLng(m("index")) > 0 Then
        If Len(s) > 2 Then s = s & ", "
        s = s & """index"": " & CStr(CLng(m("index")))
    End If
    If Len(s) > 2 Then s = s & ", "
    If CBool(m("descendants")) Then
        s = s & """scope"": ""descendants"""
    Else
        s = s & """scope"": ""children"""
    End If
    MatchJson = s & " }"
End Function

Private Function Pair(ByVal soFar As String, ByVal k As String, ByVal v As String) As String
    If Len(v) = 0 Then Exit Function
    If Len(soFar) > 2 Then Pair = ", "
    Pair = Pair & JsonStr(k) & ": " & JsonStr(v)
End Function

Private Function JsonBool(ByVal b As Boolean) As String
    If b Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function

Private Function JsonStr(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim c As Long
    out = """"
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        Select Case c
            Case 34
                out = out & "\"""
            Case 92
                out = out & "\\"
            Case 10
                out = out & "\n"
            Case 13
                out = out & "\r"
            Case 9
                out = out & "\t"
            Case Else
                If c < 32 Then
                    out = out & "\u" & Right$("000" & Hex$(c), 4)
                Else
                    out = out & ChrW$(c)
                End If
        End Select
    Next i
    JsonStr = out & """"
End Function

'==============================================================================
' member access on a parsed object
'==============================================================================
Private Function GetObj(ByVal o As Object, ByVal k As String) As Object
    If o Is Nothing Then Exit Function
    If Not o.Exists(k) Then Exit Function
    If Not IsObject(o(k)) Then Exit Function
    If TypeName(o(k)) <> "Dictionary" Then Exit Function
    Set GetObj = o(k)
End Function

Private Function GetArr(ByVal o As Object, ByVal k As String) As Collection
    If o Is Nothing Then Exit Function
    If Not o.Exists(k) Then Exit Function
    If Not IsObject(o(k)) Then Exit Function
    If TypeName(o(k)) <> "Collection" Then Exit Function
    Set GetArr = o(k)
End Function

Private Function GetStr(ByVal o As Object, ByVal k As String, ByVal def As String) As String
    GetStr = def
    If o Is Nothing Then Exit Function
    If Not o.Exists(k) Then Exit Function
    If IsObject(o(k)) Then Exit Function
    If IsNull(o(k)) Then Exit Function
    GetStr = CStr(o(k))
End Function

Private Function GetLong(ByVal o As Object, ByVal k As String, ByVal def As Long) As Long
    GetLong = def
    If o Is Nothing Then Exit Function
    If Not o.Exists(k) Then Exit Function
    If IsObject(o(k)) Then Exit Function
    If IsNull(o(k)) Then Exit Function
    If Not IsNumeric(o(k)) Then Exit Function
    GetLong = CLng(Int(CDbl(o(k))))
End Function

Private Function GetBool(ByVal o As Object, ByVal k As String, ByVal def As Boolean) As Boolean
    GetBool = def
    If o Is Nothing Then Exit Function
    If Not o.Exists(k) Then Exit Function
    If IsObject(o(k)) Then Exit Function
    If IsNull(o(k)) Then Exit Function
    If VarType(o(k)) = vbBoolean Then GetBool = CBool(o(k))
End Function

'==============================================================================
' JSONC -- the same syntax Rdv3Json.cs accepts
'==============================================================================
Private Function ParseJsonc(ByVal text As String) As Object
    Dim vo As Object
    Dim vv As Variant
    Dim isObj As Boolean
    m_src = text
    m_pos = 1
    m_line = 1
    If Len(m_src) > 0 Then
        If Mid$(m_src, 1, 1) = ChrW$(65279) Then m_pos = 2
    End If
    ReadValueInto vo, vv, isObj
    SkipWhite
    If m_pos <= Len(m_src) Then JsonFail "最上位の値の後ろに余分な文字があります"
    If Not isObj Then Exit Function
    If TypeName(vo) <> "Dictionary" Then Exit Function
    Set ParseJsonc = vo
End Function

Private Sub JsonFail(ByVal msg As String)
    Err.Raise vbObjectError + 513, "modRdv3Cfg", msg & " (" & CStr(m_line) & " 行目)"
End Sub

Private Sub SkipWhite()
    Dim c As String
    Dim n As Long
    n = Len(m_src)
    Do While m_pos <= n
        c = Mid$(m_src, m_pos, 1)
        If c = vbLf Then
            m_line = m_line + 1
            m_pos = m_pos + 1
        ElseIf c = " " Or c = vbTab Or c = vbCr Then
            m_pos = m_pos + 1
        ElseIf c = "/" And m_pos < n And Mid$(m_src, m_pos + 1, 1) = "/" Then
            Do While m_pos <= n
                If Mid$(m_src, m_pos, 1) = vbLf Then Exit Do
                m_pos = m_pos + 1
            Loop
        ElseIf c = "/" And m_pos < n And Mid$(m_src, m_pos + 1, 1) = "*" Then
            m_pos = m_pos + 2
            Do While m_pos < n
                If Mid$(m_src, m_pos, 1) = "*" And Mid$(m_src, m_pos + 1, 1) = "/" Then Exit Do
                If Mid$(m_src, m_pos, 1) = vbLf Then m_line = m_line + 1
                m_pos = m_pos + 1
            Loop
            m_pos = m_pos + 2
            If m_pos > n + 1 Then m_pos = n + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function PeekChar() As String
    SkipWhite
    If m_pos > Len(m_src) Then JsonFail "ファイルが途中で終わっています"
    PeekChar = Mid$(m_src, m_pos, 1)
End Function

' Three outputs, not one Variant, and not a return value. A JSON value is either
' an object or a scalar, and VBA needs Set for the one and Let for the other --
' but a Variant that is STILL HOLDING an object from the previous member does not
' simply take a Let: the assignment goes through that object's default member
' instead (for a Dictionary that is Item(key), which wants an argument, so the
' parser died with "type mismatch" on the second member of every object). Keeping
' the object in an As Object variable and the scalar in a Variant that never holds
' one means neither assignment can ever be routed through anything.
Private Sub ReadValueInto(ByRef outObj As Object, ByRef outVal As Variant, ByRef isObj As Boolean)
    Dim c As String
    Set outObj = Nothing
    outVal = Empty
    isObj = False
    c = PeekChar()
    If c = "{" Then
        Set outObj = ReadObjectAt()
        isObj = True
    ElseIf c = "[" Then
        Set outObj = ReadArrayAt()
        isObj = True
    ElseIf c = """" Then
        outVal = ReadStringAt()
    ElseIf c = "-" Or (c >= "0" And c <= "9") Then
        outVal = ReadNumberAt()
    ElseIf WordAt("true") Then
        outVal = True
    ElseIf WordAt("false") Then
        outVal = False
    ElseIf WordAt("null") Then
        outVal = Null
    Else
        JsonFail "値が来るはずの場所です"
    End If
End Sub

Private Function WordAt(ByVal w As String) As Boolean
    If m_pos + Len(w) - 1 > Len(m_src) Then Exit Function
    If Mid$(m_src, m_pos, Len(w)) <> w Then Exit Function
    m_pos = m_pos + Len(w)
    WordAt = True
End Function

Private Function ReadObjectAt() As Object
    Dim o As Object
    Dim vo As Object
    Dim vv As Variant
    Dim isObj As Boolean
    Dim k As String
    Dim c As String
    Set o = NewDict()
    m_pos = m_pos + 1
    Do
        c = PeekChar()
        If c = "}" Then
            m_pos = m_pos + 1
            Set ReadObjectAt = o
            Exit Function
        End If
        If c <> """" Then JsonFail "メンバー名は引用符で囲みます"
        k = ReadStringAt()
        If PeekChar() <> ":" Then JsonFail "メンバー名の後には : が要ります"
        m_pos = m_pos + 1
        ReadValueInto vo, vv, isObj
        If isObj Then
            Set o(k) = vo
        Else
            o(k) = vv
        End If
        c = PeekChar()
        If c = "," Then
            m_pos = m_pos + 1
        ElseIf c = "}" Then
            m_pos = m_pos + 1
            Set ReadObjectAt = o
            Exit Function
        Else
            JsonFail "オブジェクトの中は , か } です"
        End If
    Loop
End Function

Private Function ReadArrayAt() As Collection
    Dim a As Collection
    Dim vo As Object
    Dim vv As Variant
    Dim isObj As Boolean
    Dim c As String
    Set a = New Collection
    m_pos = m_pos + 1
    Do
        c = PeekChar()
        If c = "]" Then
            m_pos = m_pos + 1
            Set ReadArrayAt = a
            Exit Function
        End If
        ReadValueInto vo, vv, isObj
        If isObj Then
            a.Add vo
        Else
            a.Add vv
        End If
        c = PeekChar()
        If c = "," Then
            m_pos = m_pos + 1
        ElseIf c = "]" Then
            m_pos = m_pos + 1
            Set ReadArrayAt = a
            Exit Function
        Else
            JsonFail "配列の中は , か ] です"
        End If
    Loop
End Function

Private Function ReadStringAt() As String
    Dim out As String
    Dim c As String
    Dim e As String
    Dim n As Long
    n = Len(m_src)
    m_pos = m_pos + 1
    Do
        If m_pos > n Then JsonFail "文字列が閉じていません"
        c = Mid$(m_src, m_pos, 1)
        m_pos = m_pos + 1
        If c = """" Then
            ReadStringAt = out
            Exit Function
        End If
        If c = vbLf Then JsonFail "文字列の中で改行しています"
        If c <> "\" Then
            out = out & c
        Else
            If m_pos > n Then JsonFail "文字列が閉じていません"
            e = Mid$(m_src, m_pos, 1)
            m_pos = m_pos + 1
            Select Case e
                Case "n"
                    out = out & vbLf
                Case "t"
                    out = out & vbTab
                Case "r"
                    out = out & vbCr
                Case "b"
                    out = out & Chr$(8)
                Case "f"
                    out = out & Chr$(12)
                Case "/", "\", """"
                    out = out & e
                Case "u"
                    If m_pos + 3 > n Then JsonFail "\u の後が足りません"
                    out = out & ChrW$(HexQuad(Mid$(m_src, m_pos, 4)))
                    m_pos = m_pos + 4
                Case Else
                    JsonFail "知らないエスケープです"
            End Select
        End If
    Loop
End Function

' The four hex digits of a \u escape, read by hand. Neither CLng nor Val will do
' it: both read "&HFFFF" as a SIGNED 16 bit literal and answer -1, and the "&"
' suffix that forces a Long in source code is not something either accepts inside
' a string -- CLng("&H30e1&") is a type mismatch.
Private Function HexQuad(ByVal h As String) As Long
    Dim i As Long
    Dim c As Long
    Dim v As Long
    For i = 1 To Len(h)
        c = AscW(Mid$(h, i, 1))
        If c >= 48 And c <= 57 Then
            v = v * 16 + (c - 48)
        ElseIf c >= 65 And c <= 70 Then
            v = v * 16 + (c - 55)
        ElseIf c >= 97 And c <= 102 Then
            v = v * 16 + (c - 87)
        Else
            JsonFail "\u の後は 16 進 4 桁です"
        End If
    Next i
    HexQuad = v
End Function

Private Function ReadNumberAt() As Double
    Dim st As Long
    Dim c As String
    Dim n As Long
    n = Len(m_src)
    st = m_pos
    If m_pos <= n Then
        If Mid$(m_src, m_pos, 1) = "-" Then m_pos = m_pos + 1
    End If
    Do While m_pos <= n
        c = Mid$(m_src, m_pos, 1)
        If (c >= "0" And c <= "9") Or c = "." Or c = "e" Or c = "E" Or c = "+" Or c = "-" Then
            m_pos = m_pos + 1
        Else
            Exit Do
        End If
    Loop
    If m_pos = st Then JsonFail "数値が読めません"
    ReadNumberAt = Val(Mid$(m_src, st, m_pos - st))
End Function

'==============================================================================
' UTF-8, by hand. No ADODB and no Win32.
'==============================================================================
Private Function ReadUtf8(ByVal path As String) As String
    Dim f As Integer
    Dim b() As Byte
    Dim n As Long
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n = 0 Then
        Close #f
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f
    ReadUtf8 = Utf8ToText(b)
End Function

Private Function Utf8ToText(ByRef b() As Byte) As String
    Dim i As Long
    Dim n As Long
    Dim c As Long
    Dim cp As Long
    Dim extra As Long
    Dim sb() As String
    Dim used As Long

    n = UBound(b) - LBound(b) + 1
    If n <= 0 Then Exit Function
    ReDim sb(0 To n)
    i = LBound(b)
    ' a UTF-8 BOM is not part of the text
    If n >= 3 Then
        If b(i) = 239 And b(i + 1) = 187 And b(i + 2) = 191 Then i = i + 3
    End If
    Do While i <= UBound(b)
        c = b(i)
        If c < 128 Then
            cp = c
            extra = 0
        ElseIf c >= 192 And c < 224 Then
            cp = c And 31
            extra = 1
        ElseIf c >= 224 And c < 240 Then
            cp = c And 15
            extra = 2
        ElseIf c >= 240 And c < 248 Then
            cp = c And 7
            extra = 3
        Else
            ' a stray continuation byte: the file is not the UTF-8 it claims to
            ' be. Say so rather than invent a character.
            Err.Raise vbObjectError + 514, "modRdv3Cfg", "UTF-8 として読めないバイトがあります"
        End If
        i = i + 1
        Do While extra > 0
            If i > UBound(b) Then
                Err.Raise vbObjectError + 514, "modRdv3Cfg", "UTF-8 の途中でファイルが終わっています"
            End If
            If (b(i) And 192) <> 128 Then
                Err.Raise vbObjectError + 514, "modRdv3Cfg", "UTF-8 として読めないバイトがあります"
            End If
            cp = (cp * 64) + (b(i) And 63)
            i = i + 1
            extra = extra - 1
        Loop
        If cp < 65536 Then
            sb(used) = ChrW$(cp)
        Else
            cp = cp - 65536
            sb(used) = ChrW$(55296 + (cp \ 1024)) & ChrW$(56320 + (cp Mod 1024))
        End If
        used = used + 1
    Loop
    If used = 0 Then Exit Function
    ReDim Preserve sb(0 To used - 1)
    Utf8ToText = Join(sb, "")
End Function

Private Function TextToUtf8(ByVal s As String) As Byte()
    Dim i As Long
    Dim c As Long
    Dim lo As Long
    Dim b() As Byte
    Dim used As Long
    Dim n As Long

    n = Len(s)
    If n = 0 Then
        ReDim b(0 To 0)
        TextToUtf8 = b
        Exit Function
    End If
    ReDim b(0 To n * 4 - 1)
    i = 1
    Do While i <= n
        c = AscW(Mid$(s, i, 1))
        If c < 0 Then c = c + 65536
        ' a surrogate pair is one character
        If c >= 55296 And c <= 56319 And i < n Then
            lo = AscW(Mid$(s, i + 1, 1))
            If lo < 0 Then lo = lo + 65536
            If lo >= 56320 And lo <= 57343 Then
                c = 65536 + (c - 55296) * 1024 + (lo - 56320)
                i = i + 1
            End If
        End If
        If c < 128 Then
            b(used) = c
            used = used + 1
        ElseIf c < 2048 Then
            b(used) = 192 + (c \ 64)
            b(used + 1) = 128 + (c Mod 64)
            used = used + 2
        ElseIf c < 65536 Then
            b(used) = 224 + (c \ 4096)
            b(used + 1) = 128 + ((c \ 64) Mod 64)
            b(used + 2) = 128 + (c Mod 64)
            used = used + 3
        Else
            b(used) = 240 + (c \ 262144)
            b(used + 1) = 128 + ((c \ 4096) Mod 64)
            b(used + 2) = 128 + ((c \ 64) Mod 64)
            b(used + 3) = 128 + (c Mod 64)
            used = used + 4
        End If
        i = i + 1
    Loop
    ReDim Preserve b(0 To used - 1)
    TextToUtf8 = b
End Function

' UTF-8 without a BOM, written beside the live file and renamed over it -- the
' same replace rule the channel and the sidecar use, so a settings file is never
' deleted before its replacement exists.
Private Function WriteUtf8Atomic(ByVal path As String, ByVal text As String) As String
    Dim f As Integer
    Dim b() As Byte
    Dim tmp As String
    tmp = path & ".tmp"
    On Error GoTo Failed
    If Len(Dir$(tmp)) > 0 Then Kill tmp
    b = TextToUtf8(text)
    f = FreeFile
    Open tmp For Binary Access Write As #f
    If Len(text) > 0 Then Put #f, 1, b
    Close #f
    f = 0
    If Not Rdv3ChReplaceFile(tmp, path) Then
        WriteUtf8Atomic = "設定ファイルの置き換えに失敗 (err " & CStr(Rdv3ChLastReplaceErr()) & ")"
        Exit Function
    End If
    Exit Function
Failed:
    WriteUtf8Atomic = "設定の書き込みに失敗 " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If f <> 0 Then Close #f
End Function

'==============================================================================
' paths, resolved against the folder holding the workbook (same rule as the C#
' build: a relative path is relative to the folder the program sits in)
'==============================================================================
Public Function Rdv3CfgResolve(ByVal raw As String, ByVal baseDir As String, _
                               ByVal fallback As String) As String
    Dim s As String
    s = Trim$(raw)
    If Len(s) = 0 Then
        Rdv3CfgResolve = fallback
        Exit Function
    End If
    If Mid$(s, 2, 1) = ":" Or Left$(s, 2) = "\\" Then
        Rdv3CfgResolve = s
    Else
        Rdv3CfgResolve = baseDir & "\" & s
    End If
End Function

' one line for the log, so a support question starts from what was read
Public Function Rdv3CfgDescribe() As String
    Dim s As String
    Dim t As Object
    Dim steps As Collection
    Dim i As Long
    Dim k As Long
    EnsureReady
    If m_loaded Then
        s = "loaded "
    Else
        s = "defaults "
    End If
    s = s & m_sourcePath
    If Len(m_error) > 0 Then s = s & " ERROR=" & m_error
    s = s & " key=" & CStr(m_keyLength)
    If m_keyDigitsOnly Then
        s = s & "digits"
    Else
        s = s & "any"
    End If
    s = s & " poll=" & CStr(m_pollMs) & "/" & CStr(m_stableMs)
    s = s & " targets=" & CStr(m_targets.Count) & " watched=" & CStr(Rdv3CfgWatchableCount())
    For i = 1 To m_targets.Count
        Set t = m_targets.Item(i)
        s = s & " [" & CStr(t("name"))
        If Not CBool(t("enabled")) Then
            s = s & " OFF"
        ElseIf Len(Rdv3CfgWhyNotWatchable(t)) > 0 Then
            s = s & " UNUSABLE"
        End If
        s = s & ": win " & MatchDescribe(t("window"))
        Set steps = t("path")
        For k = 1 To steps.Count
            s = s & " > " & MatchDescribe(steps.Item(k))
        Next k
        s = s & " > field " & MatchDescribe(t("field"))
        s = s & " read=" & CStr(t("read")) & "]"
    Next i
    Rdv3CfgDescribe = s
End Function

Private Function MatchDescribe(ByVal m As Object) As String
    Dim s As String
    Dim types As Collection
    Dim joined As String
    Dim i As Long
    s = Seg(s, "automationId", CStr(m("automationId")))
    s = s & Seg(s, "class", CStr(m("className")))
    s = s & Seg(s, "name", CStr(m("name")))
    s = s & Seg(s, "nameLike", CStr(m("nameLike")))
    s = s & Seg(s, "process", CStr(m("processName")))
    Set types = m("controlTypes")
    If types.Count > 0 Then
        For i = 1 To types.Count
            If i > 1 Then joined = joined & "/"
            joined = joined & CStr(types.Item(i))
        Next i
        s = s & Seg(s, "type", joined)
    End If
    If CBool(m("requireValuePattern")) Then s = s & Seg(s, "value", "yes")
    If CLng(m("index")) > 0 Then s = s & Seg(s, "index", CStr(CLng(m("index"))))
    If CBool(m("descendants")) Then
        s = s & Seg(s, "scope", "descendants")
    Else
        s = s & Seg(s, "scope", "children")
    End If
    If Len(s) = 0 Then
        MatchDescribe = "(anything)"
    Else
        MatchDescribe = s
    End If
End Function

Private Function Seg(ByVal soFar As String, ByVal k As String, ByVal v As String) As String
    If Len(v) = 0 Then Exit Function
    If Len(soFar) > 0 Then Seg = " "
    Seg = Seg & k & "=" & v
End Function

'==============================================================================
' self test -- run by build\build_workbook_app.ps1 before the workbook ships.
'
' The C# build has test_settings_contract.ps1 for exactly these promises; this
' is the VBA half of it, and it runs inside the Excel that will host the code
' rather than against a copy of it. It writes only under TEMP.
'==============================================================================
Public Function Rdv3CfgSelfTest() As String
    Dim dir As String
    Dim p As String
    Dim s As String
    Dim t As Object
    Dim bad As String
    Dim n As Long

    On Error GoTo Failed
    dir = Environ$("TEMP")
    If Len(dir) = 0 Then
        Rdv3CfgSelfTest = "FAIL no TEMP"
        Exit Function
    End If
    p = dir & "\rdv3cfg-selftest.json"

    ' ---- 1. every comment form the C# reader accepts, and strings that only
    ' look like comments
    s = "// whole line" & vbCrLf & _
        "{" & vbCrLf & _
        "  /* block" & vbCrLf & "     over two lines */" & vbCrLf & _
        "  ""schema"": 1,   // trailing" & vbCrLf & _
        "  ""paths"": { ""dataDir"": ""D:\\a//b"", ""ledger"": ""L.xlsx"", ""log"": ""r.log"" }," & vbCrLf & _
        "  ""key"": { ""length"": 6, ""digitsOnly"": false }," & vbCrLf & _
        "  ""jobs"": { ""checkTimeoutMs"": 0 }," & vbCrLf & _
        "  ""watch"": { ""pollMs"": 55, ""targets"": [" & vbCrLf & _
        "    { ""enabled"": true, ""name"": ""\u30e1\u30e2"", ""window"": { ""className"": ""Notepad"" }," & vbCrLf & _
        "      ""field"": { ""controlTypes"": [""Edit""] }, ""read"": ""value"" }," & vbCrLf & _
        "    { ""enabled"": false, ""name"": ""off"", ""window"": { ""processName"": ""Lob"" }," & vbCrLf & _
        "      ""field"": { ""controlTypes"": [""Edit""] } }," & vbCrLf & _
        "    { ""enabled"": true, ""name"": ""typo"", ""window"": { ""className"": ""X"" }," & vbCrLf & _
        "      ""field"": { ""controlTypes"": [""Edti""] } }" & vbCrLf & _
        "  ] }" & vbCrLf & "}" & vbCrLf
    If Not WriteRaw(p, s) Then
        Rdv3CfgSelfTest = "FAIL cannot write the test file"
        Exit Function
    End If
    If Not Rdv3CfgLoad(p) Then
        Rdv3CfgSelfTest = "FAIL parse: " & Rdv3CfgError()
        Exit Function
    End If
    bad = bad & Expect("dataDir", Rdv3CfgDataDir(), "D:\a//b")
    bad = bad & Expect("key.length", CStr(Rdv3CfgKeyLength()), "6")
    bad = bad & Expect("digitsOnly", CStr(Rdv3CfgKeyDigitsOnly()), CStr(False))
    bad = bad & Expect("pollMs", CStr(Rdv3CfgPollMs()), "55")
    ' out of range falls back on its OWN default, not on the nearest limit
    bad = bad & Expect("checkTimeoutMs", CStr(Rdv3CfgCheckTimeoutMs()), "180000")
    ' the disabled one is kept, the misspelt controlType one is dropped
    bad = bad & Expect("targets", CStr(Rdv3CfgTargetCount()), "2")
    bad = bad & Expect("watched", CStr(Rdv3CfgWatchableCount()), "1")
    If Rdv3CfgTargetCount() = 2 Then
        Set t = Rdv3CfgTarget(1)
        bad = bad & Expect("u-escape", CStr(t("name")), ChrW$(&H30E1) & ChrW$(&H30E2))
        Set t = Rdv3CfgTarget(2)
        bad = bad & Expect("kept off", CStr(t("enabled")), CStr(False))
    End If

    ' ---- 2. an explicit empty list means nothing is watched
    If Not WriteRaw(p, "{ ""watch"": { ""targets"": [] } }") Then
        Rdv3CfgSelfTest = "FAIL cannot write the test file"
        Exit Function
    End If
    Rdv3CfgLoad p
    bad = bad & Expect("empty list", CStr(Rdv3CfgTargetCount()), "0")

    ' ---- 3. no watch member at all keeps the built-in target
    If Not WriteRaw(p, "{ ""key"": { ""length"": 8 } }") Then
        Rdv3CfgSelfTest = "FAIL cannot write the test file"
        Exit Function
    End If
    Rdv3CfgLoad p
    bad = bad & Expect("built-in", CStr(Rdv3CfgTargetCount()), "1")

    ' ---- 4. write what we hold, read it back, and get the same answers. This is
    ' the round trip the settings sheet depends on, UTF-8 and all.
    Rdv3CfgDefaults
    Rdv3CfgSetPaths ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF), "L.xlsx", "r.log"
    Rdv3CfgSetKey 6, False
    Rdv3CfgSetWatch 55, 130, 410, False
    Rdv3CfgSetCandidateRows 7
    Set t = Rdv3CfgTarget(1)
    t("name") = ChrW$(&H30E1) & ChrW$(&H30E2) & ChrW$(&H5E33)
    s = Rdv3CfgSave(p)
    If Len(s) > 0 Then
        Rdv3CfgSelfTest = "FAIL save: " & s
        Exit Function
    End If
    Rdv3CfgDefaults
    If Not Rdv3CfgLoad(p) Then
        Rdv3CfgSelfTest = "FAIL reload: " & Rdv3CfgError()
        Exit Function
    End If
    bad = bad & Expect("rt dataDir", Rdv3CfgDataDir(), ChrW$(&H30C7) & ChrW$(&H30FC) & ChrW$(&H30BF))
    bad = bad & Expect("rt key", CStr(Rdv3CfgKeyLength()), "6")
    bad = bad & Expect("rt digits", CStr(Rdv3CfgKeyDigitsOnly()), CStr(False))
    bad = bad & Expect("rt poll", CStr(Rdv3CfgPollMs()), "55")
    bad = bad & Expect("rt stable", CStr(Rdv3CfgStableMs()), "130")
    bad = bad & Expect("rt rebind", CStr(Rdv3CfgRebindMs()), "410")
    bad = bad & Expect("rt focus", CStr(Rdv3CfgPreferFocused()), CStr(False))
    bad = bad & Expect("rt cand", CStr(Rdv3CfgCandidateRows()), "7")
    bad = bad & Expect("rt targets", CStr(Rdv3CfgTargetCount()), "1")
    If Rdv3CfgTargetCount() = 1 Then
        Set t = Rdv3CfgTarget(1)
        bad = bad & Expect("rt name", CStr(t("name")), ChrW$(&H30E1) & ChrW$(&H30E2) & ChrW$(&H5E33))
        bad = bad & Expect("rt class", CStr(t("window")("className")), "Notepad")
    End If

    ' ---- 5. UTF-8 by hand, over the whole BMP plus a surrogate pair
    n = Utf8RoundTrip()
    If n <> 0 Then
        Rdv3CfgSelfTest = "FAIL utf8 round trip at code point " & CStr(n)
        Exit Function
    End If

    On Error Resume Next
    Kill p
    On Error GoTo 0
    Rdv3CfgDefaults

    If Len(bad) > 0 Then
        Rdv3CfgSelfTest = "FAIL" & bad
    Else
        Rdv3CfgSelfTest = "ok jsonc=3forms range=default targets=keep/drop roundtrip=utf8"
    End If
    Exit Function
Failed:
    Rdv3CfgSelfTest = "FAIL " & Err.Number & ": " & Err.Description
End Function

Private Function Expect(ByVal what As String, ByVal got As String, ByVal want As String) As String
    If got <> want Then Expect = " [" & what & "=" & got & " want " & want & "]"
End Function

' the raw writer, for the settings sheet's own round-trip test
Public Function Rdv3CfgWriteRawForTest(ByVal path As String, ByVal text As String) As Boolean
    Rdv3CfgWriteRawForTest = WriteRaw(path, text)
End Function

Private Function WriteRaw(ByVal path As String, ByVal text As String) As Boolean
    Dim f As Integer
    Dim b() As Byte
    On Error GoTo Failed
    If Len(Dir$(path)) > 0 Then Kill path
    b = TextToUtf8(text)
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    WriteRaw = True
    Exit Function
Failed:
    On Error Resume Next
    If f <> 0 Then Close #f
End Function

' Encoded and decoded by the two functions above. Returns 0, or the first code
' point that did not survive.
'
' Every code point would be 65,533 round trips, and inside an automation Excel
' that is not a moment -- it turned this build step into a wait long enough to
' look like the invisible-modal hang it is not. So: the boundary of every encoded
' length (where an off-by-one lives), all of the one byte range, and a stride
' through the two and three byte ranges. Both continuation shapes and the
' surrogate pair are exercised; the build stays quick.
Private Function Utf8RoundTrip() As Long
    Dim cp As Long
    Dim i As Long
    Dim edges As Variant

    ' 127/128 = 1 to 2 bytes, 2047/2048 = 2 to 3, 55295/57344 = around the
    ' surrogate block, 65535 = the one AscW returns as -1
    edges = Array(1, 127, 128, 2047, 2048, 55295, 57344, 65533, 65534, 65535)
    For i = LBound(edges) To UBound(edges)
        If Not TripOk(CLng(edges(i))) Then
            Utf8RoundTrip = CLng(edges(i))
            Exit Function
        End If
    Next i
    For cp = 1 To 127
        If Not TripOk(cp) Then
            Utf8RoundTrip = cp
            Exit Function
        End If
    Next cp
    For cp = 128 To 2047 Step 37
        If Not TripOk(cp) Then
            Utf8RoundTrip = cp
            Exit Function
        End If
    Next cp
    For cp = 2048 To 65533 Step 337
        If Not TripOk(cp) Then
            Utf8RoundTrip = cp
            Exit Function
        End If
    Next cp
    If Not PairOk() Then Utf8RoundTrip = 131137
End Function

' Lone surrogates are not characters and cannot round trip. U+FEFF on its own
' cannot either, and should not: at the START of a file those three bytes are a
' byte order mark, and the reader is right to drop them.
Private Function TripOk(ByVal cp As Long) As Boolean
    Dim s As String
    Dim b() As Byte
    If cp >= 55296 And cp <= 57343 Then
        TripOk = True
        Exit Function
    End If
    If cp = 65279 Then
        TripOk = True
        Exit Function
    End If
    s = ChrW$(cp)
    b = TextToUtf8(s)
    TripOk = (Utf8ToText(b) = s)
End Function

' one astral character, which is two UTF-16 units and four UTF-8 bytes
Private Function PairOk() As Boolean
    Dim s As String
    Dim b() As Byte
    s = ChrW$(55362) & ChrW$(56321)
    b = TextToUtf8(s)
    PairOk = (Utf8ToText(b) = s)
End Function

' A settings round trip that keeps what the screen does not show. The VBA
' settings sheet has no column for the intermediate path, the window's
' automationId, the field's className and so on -- and ReaderDataViewer.json is
' the C# build's file too, so losing them would mean "either build can write it"
' stopped being true. Run by the builder.
Public Function Rdv3CfgRoundTripTest() As String
    Dim dir As String
    Dim p As String
    Dim s As String
    Dim t As Object
    Dim bad As String

    On Error GoTo Failed
    dir = Environ$("TEMP")
    p = dir & "\rdv3cfg-roundtrip.json"

    ' a target using members the VBA sheet has no column for
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
    If Not WriteRaw(p, s) Then
        Rdv3CfgRoundTripTest = "FAIL cannot write"
        Exit Function
    End If
    If Not Rdv3CfgLoad(p) Then
        Rdv3CfgRoundTripTest = "FAIL parse: " & Rdv3CfgError()
        Exit Function
    End If
    bad = bad & Expect("targets in", CStr(Rdv3CfgTargetCount()), "7")

    ' write it straight back out and read it again: nothing may change
    s = Rdv3CfgSave(p)
    If Len(s) > 0 Then
        Rdv3CfgRoundTripTest = "FAIL save: " & s
        Exit Function
    End If
    Rdv3CfgDefaults
    Rdv3CfgLoad p
    bad = bad & Expect("targets out", CStr(Rdv3CfgTargetCount()), "7")
    If Rdv3CfgTargetCount() >= 7 Then
        Set t = Rdv3CfgTarget(1)
        bad = bad & Expect("win autoId", CStr(t("window")("automationId")), "winId")
        bad = bad & Expect("win name", CStr(t("window")("name")), "exact")
        bad = bad & Expect("win index", CStr(CLng(t("window")("index"))), "2")
        bad = bad & Expect("path steps", CStr(t("path").Count), "2")
        bad = bad & Expect("fld class", CStr(t("field")("className")), "EditCls")
        bad = bad & Expect("fld like", CStr(t("field")("nameLike")), "bar*")
        bad = bad & Expect("fld proc", CStr(t("field")("processName")), "LobApp")
        bad = bad & Expect("fld index", CStr(CLng(t("field")("index"))), "3")
        bad = bad & Expect("read", CStr(t("read")), "text")
        Set t = Rdv3CfgTarget(7)
        bad = bad & Expect("7th kept", CStr(t("name")), "t7")
    End If

    On Error Resume Next
    Kill p
    On Error GoTo 0
    Rdv3CfgDefaults
    If Len(bad) > 0 Then
        Rdv3CfgRoundTripTest = "FAIL" & bad
    Else
        Rdv3CfgRoundTripTest = "ok deep-target round trip keeps every member"
    End If
    Exit Function
Failed:
    Rdv3CfgRoundTripTest = "FAIL " & Err.Number & ": " & Err.Description
End Function
