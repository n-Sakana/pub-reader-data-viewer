Attribute VB_Name = "modRdv3Uia"
'==============================================================================
' modRdv3Uia -- Excel VBA as a UI Automation CLIENT.
'
' Every UIA client interface derives from IUnknown and has no IDispatch, and
' VBA's Object IS IDispatch: CreateObject would assign but never call. The
' reference is added by build\build_workbook_app.ps1 AFTER the imports.
'
' HOW A WINDOW IS FOUND, AND WHY THIS WAY
' The frozen builds enumerated top-level windows with FindWindowEx. Walking the
' DESKTOP with UIA from Excel's own UI thread deadlocks on Excel's own provider
' (measured; see modRdv2Uia), and Win32 is not available here, so THE DESKTOP IS
' NEVER WALKED. The entry point is IUIAutomation.GetFocusedElement: it resolves
' exactly one element -- the one with keyboard focus -- and from there
' ControlViewWalker climbs to the top-level window.
'
' THE RESIDUAL THIS LEAVES, STATED PLAINLY. The C# build finds a window whether
' or not it has ever had the focus; this one adopts a window the first time the
' operator works in it. Once adopted the binding is kept and polled, so several
' windows can be watched at the same time -- but each has to be visited once.
' The status line says so rather than looking broken. Closing that gap would
' mean walking the desktop, which is the measured deadlock, or Win32, which this
' build does not have.
'
' WHAT IS WATCHED comes from ReaderDataViewer.json (modRdv3Cfg): any number of
' targets, each a window matcher, optional intermediate steps, a field matcher
' and a way to read it (ValuePattern / TextPattern / the Name property). Each
' target keeps its own pending / last-fired state, so two applications showing
' the same number are two readings and one going away does not disturb the rest.
'
' TWO UIA MEMBERS ARE UNUSABLE FROM VBA and both have replacements here:
'   IUIAutomationElement.CurrentNativeWindowHandle returns UIA_HWND (void*),
'   which VBA rejects at COMPILE time ("automation type not supported"). In an
'   automation Excel that compile error is an INVISIBLE modal and the process
'   hangs forever (measured). GetCurrentPropertyValue returns a VARIANT and is
'   used instead. CurrentName gives the window title, so GetWindowTextW goes
'   the same way.
'
' This module only ever reads. It never starts, closes or kills anything.
'==============================================================================
Option Explicit

Private Const UIA_ProcessIdPropertyId As Long = 30002
Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_NamePropertyId As Long = 30005
Private Const UIA_AutomationIdPropertyId As Long = 30011
Private Const UIA_ClassNamePropertyId As Long = 30012
Private Const UIA_NativeWindowHandlePropertyId As Long = 30020
Private Const UIA_IsValuePatternAvailablePropertyId As Long = 30043
Private Const UIA_ValuePatternId As Long = 10002
Private Const UIA_TextPatternId As Long = 10014
Private Const WALK_MAX As Long = 12
Private Const MAX_SLOTS As Long = 16

Private m_Uia As UIAutomationClient.IUIAutomation
Private m_targets As Collection                 ' the configured targets
Private m_why As String

' one bound target per slot
Private s_used(0 To MAX_SLOTS - 1) As Boolean
Private s_idx(0 To MAX_SLOTS - 1) As Long        ' which configured target
Private s_field(0 To MAX_SLOTS - 1) As UIAutomationClient.IUIAutomationElement
Private s_val(0 To MAX_SLOTS - 1) As UIAutomationClient.IUIAutomationValuePattern
Private s_read(0 To MAX_SLOTS - 1) As Long
Private s_hwnd(0 To MAX_SLOTS - 1) As Variant
Private s_title(0 To MAX_SLOTS - 1) As String
Private s_name(0 To MAX_SLOTS - 1) As String
Private s_pending(0 To MAX_SLOTS - 1) As String
Private s_since(0 To MAX_SLOTS - 1) As Double
Private s_polls(0 To MAX_SLOTS - 1) As Long
Private s_fired(0 To MAX_SLOTS - 1) As String
Private s_empty(0 To MAX_SLOTS - 1) As Boolean


'==============================================================================
' configuration
'==============================================================================
Public Sub Rdv3UiaConfigure(ByVal targets As Collection)
    Set m_targets = targets
    Rdv3UiaResetAll
End Sub

Public Sub Rdv3UiaResetAll()
    Dim i As Long
    For i = 0 To MAX_SLOTS - 1
        s_used(i) = False
        Set s_field(i) = Nothing
        Set s_val(i) = Nothing
        s_hwnd(i) = 0
        s_title(i) = ""
        s_name(i) = ""
        s_pending(i) = ""
        s_polls(i) = 0
        s_fired(i) = ""
        s_empty(i) = True
    Next i
End Sub

Public Function Rdv3UiaWhy() As String
    Rdv3UiaWhy = m_why
End Function

Public Function Rdv3UiaBoundCount() As Long
    Dim i As Long
    Dim n As Long
    For i = 0 To MAX_SLOTS - 1
        If s_used(i) Then n = n + 1
    Next i
    Rdv3UiaBoundCount = n
End Function

Public Function Rdv3UiaWantedCount() As Long
    Dim i As Long
    Dim n As Long
    If m_targets Is Nothing Then Exit Function
    For i = 1 To m_targets.Count
        If Rdv3CfgWatchable(m_targets.Item(i)) Then n = n + 1
    Next i
    Rdv3UiaWantedCount = n
End Function

' the first bound target, for the status line
Public Function Rdv3UiaHwnd() As Variant
    Dim i As Long
    Rdv3UiaHwnd = 0
    For i = 0 To MAX_SLOTS - 1
        If s_used(i) Then
            Rdv3UiaHwnd = s_hwnd(i)
            Exit Function
        End If
    Next i
End Function

Public Function Rdv3UiaTitle() As String
    Dim i As Long
    For i = 0 To MAX_SLOTS - 1
        If s_used(i) Then
            Rdv3UiaTitle = s_title(i)
            Exit Function
        End If
    Next i
End Function

Public Function Rdv3UiaBound() As Boolean
    Rdv3UiaBound = (Rdv3UiaBoundCount() > 0)
End Function

'==============================================================================
' binding: only ever from the window that HAS THE FOCUS
'==============================================================================
Public Function Rdv3UiaBind() As Boolean
    Dim win As UIAutomationClient.IUIAutomationElement
    Dim t As Object
    Dim i As Long
    Dim slot As Long

    On Error GoTo Failed
    Rdv3UiaBind = False
    If m_targets Is Nothing Then
        m_why = "監視対象が設定されていません"
        Exit Function
    End If
    If m_Uia Is Nothing Then Set m_Uia = New UIAutomationClient.CUIAutomation

    Set win = FocusedTopLevel()
    If win Is Nothing Then
        m_why = TellWhat()
        Exit Function
    End If

    For i = 1 To m_targets.Count
        Set t = m_targets.Item(i)
        If Rdv3CfgWatchable(t) Then
            If Not AlreadyBound(i) Then
                If MatchElement(win, t("window")) Then
                    slot = FreeSlot()
                    If slot >= 0 Then
                        If BindOne(slot, i, t, win) Then Rdv3UiaBind = True
                    End If
                End If
            End If
        End If
    Next i
    If Not Rdv3UiaBind Then m_why = TellWhat()
    Exit Function
Failed:
    m_why = "接続に失敗 " & Err.Number & ": " & Err.Description
End Function

' what the status line says while nothing (or not everything) is bound
Private Function TellWhat() As String
    Dim n As Long
    Dim w As Long
    n = Rdv3UiaBoundCount()
    w = Rdv3UiaWantedCount()
    If w = 0 Then
        TellWhat = "監視対象がありません (設定で追加してください)"
    ElseIf n = 0 Then
        TellWhat = "対象の欄を一度クリックすると接続します"
    ElseIf n < w Then
        TellWhat = CStr(n) & "/" & CStr(w) & " 接続。残りはその欄を一度クリックすると接続します"
    Else
        TellWhat = ""
    End If
End Function

Private Function AlreadyBound(ByVal cfgIndex As Long) As Boolean
    Dim i As Long
    For i = 0 To MAX_SLOTS - 1
        If s_used(i) Then
            If s_idx(i) = cfgIndex Then
                AlreadyBound = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FreeSlot() As Long
    Dim i As Long
    For i = 0 To MAX_SLOTS - 1
        If Not s_used(i) Then
            FreeSlot = i
            Exit Function
        End If
    Next i
    FreeSlot = -1
End Function

Private Function BindOne(ByVal slot As Long, ByVal cfgIndex As Long, _
                         ByVal t As Object, _
                         ByVal win As UIAutomationClient.IUIAutomationElement) As Boolean
    Dim at As UIAutomationClient.IUIAutomationElement
    Dim fld As UIAutomationClient.IUIAutomationElement
    Dim steps As Collection
    Dim k As Long

    On Error GoTo Failed
    Set at = win
    Set steps = t("path")
    For k = 1 To steps.Count
        Set at = FindOne(at, steps.Item(k))
        If at Is Nothing Then Exit Function
    Next k
    Set fld = FindOne(at, t("field"))
    If fld Is Nothing Then Exit Function

    s_read(slot) = Rdv3CfgReadModeId(CStr(t("read")))
    If s_read(slot) = RDV3_READ_VALUE Then
        Set s_val(slot) = fld.GetCurrentPattern(UIA_ValuePatternId)
        If s_val(slot) Is Nothing Then Exit Function
    End If
    Set s_field(slot) = fld
    s_idx(slot) = cfgIndex
    s_hwnd(slot) = win.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    s_title(slot) = win.CurrentName
    s_name(slot) = CStr(t("name"))
    s_pending(slot) = ""
    s_polls(slot) = 0
    s_fired(slot) = ""
    s_empty(slot) = True
    s_used(slot) = True
    BindOne = True
    Exit Function
Failed:
    Set s_val(slot) = Nothing
    Set s_field(slot) = Nothing
End Function

' focused element -> up the control view -> its top level window. Bounded: a
' focused element is a handful of levels below its window, and this must not
' become a walk.
Private Function FocusedTopLevel() As UIAutomationClient.IUIAutomationElement
    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim parent As UIAutomationClient.IUIAutomationElement
    Dim root As UIAutomationClient.IUIAutomationElement
    Dim depth As Long

    On Error GoTo Failed
    Set el = m_Uia.GetFocusedElement
    If el Is Nothing Then Exit Function
    Set root = m_Uia.GetRootElement
    Set walker = m_Uia.ControlViewWalker
    For depth = 1 To WALK_MAX
        Set parent = walker.GetParentElement(el)
        If parent Is Nothing Then Exit For
        If m_Uia.CompareElements(parent, root) <> 0 Then Exit For
        Set el = parent
    Next depth
    Set FocusedTopLevel = el
    Exit Function
Failed:
    ' a provider that is busy or gone answers with an error; that is a "not
    ' now", not a failure of the app, and the next poll tries again
    Set FocusedTopLevel = Nothing
End Function

'==============================================================================
' matching, against the same members the C# build matches on
'==============================================================================
Private Function MatchElement(ByVal el As UIAutomationClient.IUIAutomationElement, _
                              ByVal m As Object) As Boolean
    Dim types As Collection
    Dim i As Long
    Dim want As Long
    Dim okType As Boolean

    On Error GoTo Failed
    If Len(CStr(m("automationId"))) > 0 Then
        If StrComp(el.CurrentAutomationId, CStr(m("automationId")), vbBinaryCompare) <> 0 Then Exit Function
    End If
    If Len(CStr(m("className"))) > 0 Then
        If StrComp(el.CurrentClassName, CStr(m("className")), vbBinaryCompare) <> 0 Then Exit Function
    End If
    If Len(CStr(m("name"))) > 0 Then
        If StrComp(el.CurrentName, CStr(m("name")), vbBinaryCompare) <> 0 Then Exit Function
    End If
    If Len(CStr(m("nameLike"))) > 0 Then
        If Not Rdv3UiaLike(el.CurrentName, CStr(m("nameLike"))) Then Exit Function
    End If
    If Len(CStr(m("processName"))) > 0 Then
        ' processName cannot be resolved any more: the only way to turn a
        ' process id into a name here was WMI, which this build may no longer
        ' use. Rather than ignore the constraint -- which would let the matcher
        ' bind to a window the operator never named -- a target that asks for a
        ' process name matches nothing. Match on className / nameLike instead.
        Exit Function
    End If
    If CBool(m("requireValuePattern")) Then
        If CBool(el.GetCurrentPropertyValue(UIA_IsValuePatternAvailablePropertyId)) = False Then Exit Function
    End If
    Set types = m("controlTypes")
    If Rdv3CfgKnownTypeCount(m) > 0 Then
        okType = False
        For i = 1 To types.Count
            want = Rdv3CfgControlTypeId(CStr(types.Item(i)))
            If want <> 0 Then
                If CLng(el.CurrentControlType) = want Then okType = True
            End If
        Next i
        If Not okType Then Exit Function
    End If
    MatchElement = True
    Exit Function
Failed:
    MatchElement = False
End Function

' the index-th descendant (or child) that matches, or Nothing. "which one of the
' several that match" -- never the nearest one available, which would read a
' number off the wrong part of the screen.
Private Function FindOne(ByVal from As UIAutomationClient.IUIAutomationElement, _
                         ByVal m As Object) As UIAutomationClient.IUIAutomationElement
    Dim found As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim scope As Long
    Dim i As Long
    Dim hit As Long
    Dim want As Long

    On Error GoTo Failed
    If from Is Nothing Then Exit Function
    If CBool(m("descendants")) Then
        scope = TreeScope_Descendants
    Else
        scope = TreeScope_Children
    End If
    want = CLng(m("index"))
    Set found = from.FindAll(scope, m_Uia.CreateTrueCondition)
    If found Is Nothing Then Exit Function
    For i = 0 To found.Length - 1
        Set el = found.GetElement(i)
        If MatchElement(el, m) Then
            If hit = want Then
                Set FindOne = el
                Exit Function
            End If
            hit = hit + 1
        End If
    Next i
    Exit Function
Failed:
    Set FindOne = Nothing
End Function

' * and ? and NOTHING ELSE, matched by hand.
'
' VBA's own Like operator also treats #, [...] and [!...] as pattern syntax; the
' C# build's matcher (Rdv3Uia.Like) treats them as ordinary characters. Since
' ReaderDataViewer.json is the same file for both builds, "注文 [A] #1" as a
' nameLike has to mean the same window on both sides -- so this is written out
' rather than handed to Like.
Public Function Rdv3UiaLike(ByVal text As String, ByVal pattern As String) As Boolean
    If Len(pattern) = 0 Then
        Rdv3UiaLike = True
        Exit Function
    End If
    Rdv3UiaLike = LikeAt(LCase$(text), 1, LCase$(pattern), 1)
End Function

Private Function LikeAt(ByVal s As String, ByVal si As Long, _
                        ByVal p As String, ByVal pi As Long) As Boolean
    Dim c As String
    Dim k As Long
    Do While pi <= Len(p)
        c = Mid$(p, pi, 1)
        If c = "*" Then
            Do While pi < Len(p)
                If Mid$(p, pi + 1, 1) <> "*" Then Exit Do
                pi = pi + 1
            Loop
            If pi = Len(p) Then
                LikeAt = True
                Exit Function
            End If
            For k = si To Len(s) + 1
                If LikeAt(s, k, p, pi + 1) Then
                    LikeAt = True
                    Exit Function
                End If
            Next k
            Exit Function
        End If
        If si > Len(s) Then Exit Function
        If c <> "?" Then
            If c <> Mid$(s, si, 1) Then Exit Function
        End If
        si = si + 1
        pi = pi + 1
    Loop
    LikeAt = (si = Len(s) + 1)
End Function


'==============================================================================
' polling
'==============================================================================
' One poll of every bound target. Returns True when one of them has held a valid
' key for stableMs; key and fromName then say what and where from.
Public Function Rdv3UiaTick(ByRef key As String, ByRef fromName As String, _
                            ByRef detectMs As Double, ByRef polls As Long) As Boolean
    Dim i As Long
    Dim k As Long
    Dim n As Long
    Dim raw As Variant
    Dim cand As String
    Dim held As Double
    Dim order() As Long

    order = PollOrder(n)
    For i = 0 To n - 1
        k = order(i)
        If s_used(k) Then
            raw = ReadSlot(k)
            If IsNull(raw) Then
                Drop k
            Else
                cand = Rdv3Candidate(CStr(raw))
                If Len(cand) = 0 Then
                    s_empty(k) = True
                    s_pending(k) = ""
                    s_polls(k) = 0
                ElseIf cand <> s_pending(k) Then
                    s_pending(k) = cand
                    s_since(k) = Rdv3Ticks()
                    s_polls(k) = 1
                Else
                    s_polls(k) = s_polls(k) + 1
                    held = Rdv3MsSince(s_since(k))
                    If held >= Rdv3StableMs() And Rdv3IsKey(cand) Then
                        If cand <> s_fired(k) Or s_empty(k) Then
                            s_fired(k) = cand
                            s_empty(k) = False
                            key = cand
                            fromName = s_name(k)
                            detectMs = held
                            polls = s_polls(k)
                            Rdv3UiaTick = True
                            Exit Function          ' one reading per tick
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Function

' WHICH SLOT IS ASKED FIRST. watch.preferFocusedWindow says that when two
' applications both hold a valid number, the one the operator is actually using
' wins -- so the slot whose window has the keyboard is moved to the front and
' the rest keep the order they were configured in. Off, or with fewer than two
' bound, the configured order stands and nothing is asked of UIA.
' (C# does the same swap in Rdv3Watch before its read loop.)
Private Function PollOrder(ByRef n As Long) As Long()
    ' dynamic, not fixed: a fixed-size array is not assignable, and in an
    ' invisible automation Excel that compile error is a modal nobody can see
    Dim ord() As Long
    Dim i As Long
    Dim h As Variant
    Dim tmp As Long

    ReDim ord(0 To MAX_SLOTS - 1)
    For i = 0 To MAX_SLOTS - 1
        ord(i) = i
    Next i
    n = MAX_SLOTS
    If Not Rdv3CfgPreferFocused() Then
        PollOrder = ord
        Exit Function
    End If
    If Rdv3UiaBoundCount() < 2 Then
        PollOrder = ord
        Exit Function
    End If
    h = FocusedTopHwnd()
    If IsEmpty(h) Then
        PollOrder = ord
        Exit Function
    End If
    For i = 0 To MAX_SLOTS - 1
        If s_used(ord(i)) Then
            If CStr(s_hwnd(ord(i))) = CStr(h) Then
                tmp = ord(0)
                ord(0) = ord(i)
                ord(i) = tmp
                Exit For
            End If
        End If
    Next i
    PollOrder = ord
End Function

' the window handle of the top level window that has the keyboard, or Empty.
' The same walk the binder uses -- there is one route to the focused window here
' and this asks it for its handle.
Private Function FocusedTopHwnd() As Variant
    Dim el As UIAutomationClient.IUIAutomationElement
    On Error GoTo Failed
    If m_Uia Is Nothing Then Exit Function
    Set el = FocusedTopLevel()
    If el Is Nothing Then Exit Function
    FocusedTopHwnd = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    Exit Function
Failed:
    FocusedTopHwnd = Empty
End Function

' the raw text of one bound field. Null when the binding has gone stale, so the
' caller can tell "the field is empty" from "the window went away".
Private Function ReadSlot(ByVal i As Long) As Variant
    Dim tp As Object
    Dim rng As Object
    On Error GoTo Failed
    Select Case s_read(i)
        Case RDV3_READ_NAME
            ReadSlot = s_field(i).CurrentName
        Case RDV3_READ_TEXT
            Set tp = s_field(i).GetCurrentPattern(UIA_TextPatternId)
            If tp Is Nothing Then
                ReadSlot = Null
                Exit Function
            End If
            Set rng = tp.DocumentRange
            ReadSlot = rng.GetText(-1)
        Case Else
            If s_val(i) Is Nothing Then
                ReadSlot = Null
                Exit Function
            End If
            ReadSlot = s_val(i).CurrentValue
    End Select
    Exit Function
Failed:
    ReadSlot = Null
End Function

Private Sub Drop(ByVal i As Long)
    s_used(i) = False
    Set s_field(i) = Nothing
    Set s_val(i) = Nothing
    s_hwnd(i) = 0
    s_title(i) = ""
End Sub

'==============================================================================
' the element picker: what the operator is focused on RIGHT NOW, described the
' way the settings file describes it.
'
' The C# build points at an element with the mouse (AutomationElement.FromPoint
' plus the cursor position). Neither is reachable here without Win32, so the
' operator FOCUSES the field instead -- clicks into it -- and this reads the
' focused element. Same answer, one different gesture, and it needs nothing this
' build is not allowed to use.
'==============================================================================
Public Function Rdv3UiaPickFocused(ByVal reqId As String) As String
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim win As UIAutomationClient.IUIAutomationElement
    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    Dim parent As UIAutomationClient.IUIAutomationElement
    Dim root As UIAutomationClient.IUIAutomationElement
    Dim chain As Collection
    Dim depth As Long
    Dim s As String
    Dim readMode As String

    On Error GoTo Failed
    If m_Uia Is Nothing Then Set m_Uia = New UIAutomationClient.CUIAutomation
    Set el = m_Uia.GetFocusedElement
    If el Is Nothing Then
        Rdv3UiaPickFocused = "err=フォーカスされている要素がありません"
        Exit Function
    End If
    ' the ancestors, nearest first, up to the top level window
    Set chain = New Collection
    Set root = m_Uia.GetRootElement
    Set walker = m_Uia.ControlViewWalker
    Set win = el
    For depth = 1 To WALK_MAX
        Set parent = walker.GetParentElement(win)
        If parent Is Nothing Then Exit For
        If m_Uia.CompareElements(parent, root) <> 0 Then Exit For
        Set win = parent
        chain.Add win
    Next depth

    ' Excel's own window is not a target: picking it would be picking this app.
    ' Asked on the TOP LEVEL window through UIA (class XLMAIN) because the
    ' process name behind a process id needed WMI, which this build may not use.
    If StrComp(win.CurrentClassName, "XLMAIN", vbTextCompare) = 0 Then
        Rdv3UiaPickFocused = "err=対象アプリの欄をクリックしてから取得してください"
        Exit Function
    End If

    s = "req=" & reqId
    s = s & ";cls=" & Clean(win.CurrentClassName)
    s = s & ";wname=" & Clean(win.CurrentName)
    ' the window's OWN automationId. It is not a path step (see PathIds), but it
    ' is the strongest thing the window matcher can hold, so 採用 gets to use it.
    s = s & ";wid=" & Clean(win.CurrentAutomationId)
    s = s & ";fid=" & Clean(el.CurrentAutomationId)
    s = s & ";fcls=" & Clean(el.CurrentClassName)
    s = s & ";fname=" & Clean(el.CurrentName)
    s = s & ";ftype=" & Clean(Rdv3CfgControlTypeName(CLng(el.CurrentControlType)))
    ' An automationId is enough on its own. Without one, the same window may hold
    ' several identical Edits, and "the first that matches" would read the wrong
    ' box -- so the position among the ones that DO match is carried too. It is
    ' measured the way the RESOLVER will read it back: from the anchor the saved
    ' path ends at, with the matcher 採用 is going to save.
    If CBool(el.GetCurrentPropertyValue(UIA_IsValuePatternAvailablePropertyId)) Then
        readMode = "value"
    ElseIf Not el.GetCurrentPattern(UIA_TextPatternId) Is Nothing Then
        readMode = "text"
    Else
        readMode = "name"
    End If
    s = s & ";fidx=" & CStr(FieldIndexOf(chain, win, el, readMode))
    s = s & ";path=" & Clean(PathIds(chain))
    s = s & ";read=" & readMode
    Rdv3UiaPickFocused = s
    Exit Function
Failed:
    Rdv3UiaPickFocused = "err=取得に失敗 " & Err.Number & ": " & Err.Description
End Function

' WHICH of the fields that match, counted the way the resolver will read it back.
'
' The first version counted over the whole top-level window and matched on
' controlType + automationId only. The resolver does neither: it walks the saved
' path first and searches the LAST step's element, with the field matcher the
' adopt actually saves (className when there is no automationId, the control
' type, and requireValuePattern for a value-read). Two panels holding one id-less
' Edit each then gave the picker 1 and the resolver 0, and the target never bound.
'
' So the ordinal is measured from the same anchor, with the same matcher.
Private Function FieldIndexOf(ByVal chain As Collection, _
                              ByVal win As UIAutomationClient.IUIAutomationElement, _
                              ByVal el As UIAutomationClient.IUIAutomationElement, _
                              ByVal readMode As String) As Long
    Dim anchor As UIAutomationClient.IUIAutomationElement
    Dim m As Object
    Dim found As UIAutomationClient.IUIAutomationElementArray
    Dim other As UIAutomationClient.IUIAutomationElement
    Dim i As Long
    Dim n As Long

    On Error GoTo Failed
    Set anchor = PathAnchor(chain, win)
    Set m = Rdv3CfgNewMatch()
    m("automationId") = el.CurrentAutomationId
    If Len(CStr(m("automationId"))) = 0 Then m("className") = el.CurrentClassName
    m("controlTypes").Add Rdv3CfgControlTypeName(CLng(el.CurrentControlType))
    m("requireValuePattern") = (readMode = "value")
    m("descendants") = True
    Set found = anchor.FindAll(TreeScope_Descendants, m_Uia.CreateTrueCondition)
    If found Is Nothing Then Exit Function
    For i = 0 To found.Length - 1
        Set other = found.GetElement(i)
        If MatchElement(other, m) Then
            If m_Uia.CompareElements(other, el) <> 0 Then
                FieldIndexOf = n
                Exit Function
            End If
            n = n + 1
        End If
    Next i
    Exit Function
Failed:
    FieldIndexOf = 0
End Function

' the element the saved path ends at: the innermost ancestor PathIds keeps, or
' the window when it keeps none. Binding searches for the field from exactly here.
Private Function PathAnchor(ByVal chain As Collection, _
                            ByVal win As UIAutomationClient.IUIAutomationElement) _
                            As UIAutomationClient.IUIAutomationElement
    Dim i As Long
    On Error Resume Next
    Set PathAnchor = win
    For i = 1 To chain.Count - 1
        If Len(chain.Item(i).CurrentAutomationId) > 0 Then
            Set PathAnchor = chain.Item(i)
            Exit Function
        End If
    Next i
End Function

' the automationIds of the ancestors that have one, outermost first, as the
' settings file's intermediate path. Ancestors without an id are left out: they
' would not narrow anything and would only break when the app re-lays them out.
'
' The LAST entry of the chain is the top level window itself, and it is NOT a
' path step: binding matches the window first and then searches its DESCENDANTS
' for each step, and a window is not its own descendant. Including it made every
' target adopted from an app whose window carries an automationId impossible to
' bind. (C# builds its steps from chain[1..] for the same reason.)
Private Function PathIds(ByVal chain As Collection) As String
    Dim i As Long
    Dim id As String
    Dim s As String
    On Error Resume Next
    For i = chain.Count - 1 To 1 Step -1
        id = chain.Item(i).CurrentAutomationId
        If Len(id) > 0 Then
            If Len(s) > 0 Then s = s & "/"
            s = s & id
        End If
    Next i
    PathIds = s
End Function

Private Function Clean(ByVal s As String) As String
    Clean = Replace(Replace(Replace(s, ";", ","), vbCr, " "), vbLf, " ")
End Function
