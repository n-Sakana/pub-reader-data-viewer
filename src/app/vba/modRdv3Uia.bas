Attribute VB_Name = "modRdv3Uia"
'==============================================================================
' modRdv3Uia -- Excel VBA as a UI Automation CLIENT, watching Notepad.
' A deliberate copy of the frozen src\v2\vba\modRdv2Uia.bas with renamed
' entry points, kept separate so the frozen comparison sources stay untouched
' (the same convention modRdv2Uia itself followed toward src\vba).
'
' Required reference (late binding is impossible here, not merely inconvenient)
'     "UIAutomationClient" -- TypeLib {944DE083-8FB8-45CF-BCB7-C477ACB2F897}
' Every UIA client interface derives from IUnknown and has no IDispatch, and
' VBA's Object IS IDispatch: CreateObject would assign but never call. The
' reference is added by build\build_workbook_app.ps1 AFTER the imports.
'
' Windows are enumerated with Win32, not with UIA, and that is not a style
' choice: walking the desktop with UIA from Excel's own UI thread deadlocks on
' Excel's own provider (measured; see modRdv2Uia). FindWindowEx never enters
' UIA, and ElementFromHandle attaches to one window without walking the tree.
'
' This module only ever reads. It never starts, closes or kills Notepad.
'==============================================================================
Option Explicit

Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_ValuePatternId As Long = 10002
Private Const UIA_DocumentControlTypeId As Long = 50030
Private Const UIA_EditControlTypeId As Long = 50004

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

' Notepad windows only; the one in the foreground wins so the operator can
' pick a window just by clicking it. Same rule as every build before this one.
Public Function Rdv3UiaBind() As Boolean
    Dim c1 As UIAutomationClient.IUIAutomationCondition
    Dim c2 As UIAutomationClient.IUIAutomationCondition
    Dim cOr As UIAutomationClient.IUIAutomationCondition
    Dim d As UIAutomationClient.IUIAutomationElement
    Dim pick As UIAutomationClient.IUIAutomationElement
    #If VBA7 Then
        Dim h As LongPtr, hPick As LongPtr, hLast As LongPtr, fg As LongPtr
    #Else
        Dim h As Long, hPick As Long, hLast As Long, fg As Long
    #End If

    On Error GoTo Fail
    Rdv3UiaBind = False
    m_Why = "uia"
    If m_Uia Is Nothing Then Set m_Uia = New UIAutomationClient.CUIAutomation

    m_Why = "find"
    fg = Rdv3ForegroundWindow()
    h = 0
    Do
        h = Rdv3FindWindowEx(0, h, "Notepad", vbNullString)
        If h = 0 Then Exit Do
        If Rdv3IsWindowVisible(h) <> 0 Then
            hLast = h
            If h = fg Then hPick = h
        End If
    Loop
    If hPick = 0 Then hPick = hLast
    If hPick = 0 Then
        m_Why = "メモ帳のウィンドウがありません"
        Exit Function
    End If

    m_Why = "attach"
    Set pick = m_Uia.ElementFromHandle(ByVal hPick)
    If pick Is Nothing Then
        m_Why = "メモ帳のウィンドウに接続できません"
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

    ' handle and title from Win32: reading a UIA property off a TOP-LEVEL
    ' window from Excel's own UI thread never returns (measured; modRdv2Uia)
    Set m_Doc = d
    m_Hwnd = hPick
    m_Title = Rdv3WindowTitle(hPick)
    m_Why = ""
    Rdv3UiaBind = True
    Exit Function
Fail:
    m_Why = m_Why & " で失敗: " & Err.Number & " " & Err.Description
    Rdv3UiaReset
    Rdv3UiaBind = False
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
