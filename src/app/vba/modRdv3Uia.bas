Attribute VB_Name = "modRdv3Uia"
'==============================================================================
' modRdv3Uia -- Excel VBA as a UI Automation CLIENT, watching Notepad.
' Derived from the frozen src\v2\vba\modRdv2Uia.bas (kept separate so the
' frozen comparison sources stay untouched), with ONE difference: the window
' is found through UI Automation itself, not through Win32.
'
' Required reference (late binding is impossible here, not merely inconvenient)
'     "UIAutomationClient" -- TypeLib {944DE083-8FB8-45CF-BCB7-C477ACB2F897}
' Every UIA client interface derives from IUnknown and has no IDispatch, and
' VBA's Object IS IDispatch: CreateObject would assign but never call. The
' reference is added by build\build_workbook_app.ps1 AFTER the imports.
'
' HOW THE WINDOW IS FOUND, AND WHY THIS WAY
' The frozen builds enumerated top-level windows with FindWindowEx and picked
' the foreground Notepad, because walking the DESKTOP with UIA from Excel's own
' UI thread deadlocks on Excel's own provider (measured; see modRdv2Uia). With
' Win32 gone, the entry point is IUIAutomation.GetFocusedElement: it resolves
' exactly one element -- the one with keyboard focus -- and never walks the
' desktop, so the deadlock has no way in. From there ControlViewWalker climbs
' to the top-level window and the class name decides whether it is Notepad.
' The operator therefore chooses the window the same way as before: by working
' in it. The one behaviour change is that a Notepad window that has never had
' the focus is not adopted; the status line says so.
'
' TWO UIA MEMBERS ARE UNUSABLE FROM VBA and both have replacements here:
'   IUIAutomationElement.CurrentNativeWindowHandle returns UIA_HWND (void*),
'   which VBA rejects at COMPILE time ("automation type not supported"). In an
'   automation Excel that compile error is an INVISIBLE modal and the process
'   hangs forever (measured). GetCurrentPropertyValue returns a VARIANT and is
'   used instead. CurrentName gives the window title, so GetWindowTextW goes
'   the same way.
'
' This module only ever reads. It never starts, closes or kills Notepad.
'==============================================================================
Option Explicit

Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_NativeWindowHandlePropertyId As Long = 30020
Private Const UIA_ValuePatternId As Long = 10002
Private Const UIA_DocumentControlTypeId As Long = 50030
Private Const UIA_EditControlTypeId As Long = 50004
Private Const NOTEPAD_CLASS As String = "Notepad"
Private Const WALK_MAX As Long = 8

Private m_Uia As UIAutomationClient.IUIAutomation
Private m_Doc As UIAutomationClient.IUIAutomationElement
Private m_Val As UIAutomationClient.IUIAutomationValuePattern
Private m_Hwnd As Variant
Private m_Title As String
Private m_Why As String

Public Function Rdv3UiaWhy() As String
    Rdv3UiaWhy = m_Why
End Function

Public Function Rdv3UiaBound() As Boolean
    Rdv3UiaBound = Not (m_Val Is Nothing)
End Function

Public Function Rdv3UiaHwnd() As Variant
    Rdv3UiaHwnd = m_Hwnd
End Function

Public Function Rdv3UiaTitle() As String
    Rdv3UiaTitle = m_Title
End Function

Public Sub Rdv3UiaReset()
    Set m_Val = Nothing
    Set m_Doc = Nothing
    m_Hwnd = 0
    m_Title = ""
End Sub

' The Notepad window the operator is working in wins -- the same rule as every
' build before this one, resolved through focus instead of through Win32.
Public Function Rdv3UiaBind() As Boolean
    Dim c1 As UIAutomationClient.IUIAutomationCondition
    Dim c2 As UIAutomationClient.IUIAutomationCondition
    Dim cOr As UIAutomationClient.IUIAutomationCondition
    Dim d As UIAutomationClient.IUIAutomationElement
    Dim pick As UIAutomationClient.IUIAutomationElement

    On Error GoTo Fail
    Rdv3UiaBind = False
    m_Why = "uia"
    If m_Uia Is Nothing Then Set m_Uia = New UIAutomationClient.CUIAutomation

    m_Why = "find"
    Set pick = NotepadFromFocus()
    If pick Is Nothing Then
        m_Why = "メモ帳の入力欄をクリックすると接続します"
        Exit Function
    End If

    m_Why = "doc"
    Set c1 = m_Uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_DocumentControlTypeId)
    Set c2 = m_Uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId)
    Set cOr = m_Uia.CreateOrCondition(c1, c2)
    Set d = pick.FindFirst(TreeScope_Descendants, cOr)
    If d Is Nothing Then
        m_Why = "メモ帳の入力欄が見つかりません"
        Exit Function
    End If

    m_Why = "pattern"
    Set m_Val = d.GetCurrentPattern(UIA_ValuePatternId)
    If m_Val Is Nothing Then
        m_Why = "入力欄が ValuePattern を持っていません"
        Exit Function
    End If

    Set m_Doc = d
    m_Hwnd = pick.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    m_Title = pick.CurrentName
    m_Why = ""
    Rdv3UiaBind = True
    Exit Function
Fail:
    m_Why = m_Why & " で失敗: " & Err.Number & " " & Err.Description
    Rdv3UiaReset
    Rdv3UiaBind = False
End Function

' focused element -> up the control view -> the Notepad top-level window, or
' Nothing when the focus is somewhere else entirely. Bounded: a focused element
' is a handful of levels below its window, and this must not become a walk.
Private Function NotepadFromFocus() As UIAutomationClient.IUIAutomationElement
    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim depth As Long

    On Error GoTo Fail
    Set el = m_Uia.GetFocusedElement
    If el Is Nothing Then Exit Function
    If el.CurrentClassName = NOTEPAD_CLASS Then
        Set NotepadFromFocus = el
        Exit Function
    End If
    Set walker = m_Uia.ControlViewWalker
    For depth = 1 To WALK_MAX
        Set el = walker.GetParentElement(el)
        If el Is Nothing Then Exit Function
        If el.CurrentClassName = NOTEPAD_CLASS Then
            Set NotepadFromFocus = el
            Exit Function
        End If
    Next depth
    Exit Function
Fail:
    ' a provider that is busy or gone answers with an error; that is a "not
    ' now", not a failure of the app, and the next poll tries again
    Set NotepadFromFocus = Nothing
End Function

' one poll. Null when the binding has gone stale, so the caller can tell "the
' field is empty" from "the window went away".
Public Function Rdv3UiaRead() As Variant
    On Error GoTo Fail
    If m_Val Is Nothing Then
        Rdv3UiaRead = Null
        Exit Function
    End If
    Rdv3UiaRead = m_Val.CurrentValue
    Exit Function
Fail:
    Rdv3UiaReset
    Rdv3UiaRead = Null
End Function
