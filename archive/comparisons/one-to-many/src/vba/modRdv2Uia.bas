Attribute VB_Name = "modRdv2Uia"
'==============================================================================
' modRdv2Uia -- Excel VBA as a UI Automation CLIENT, watching Notepad.
'
' Required reference (late binding is impossible here, not merely inconvenient)
'     Tools > References > "UIAutomationClient"
'     TypeLib {944DE083-8FB8-45CF-BCB7-C477ACB2F897} v1.0
'     -> C:\Windows\System32\UIAutomationCore.dll
'
' Every UIA client interface derives from IUnknown and has no IDispatch, and
' VBA's Object IS IDispatch: CreateObject would assign but never call. So the
' reference is added by build\build_workbooks.ps1, after the modules are
' imported (adding it before makes Application.Run stop finding the macros).
'
' This module only ever reads. It never starts, closes or kills Notepad.
'==============================================================================
' Copied from the module of the same job in src\vba, with the same rules and the
' same measured
' reason for enumerating with Win32 instead of UIA. Kept separate on purpose:
' the 1:1 sources are frozen so the existing workbooks stay reproducible.
Option Explicit

Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_NamePropertyId As Long = 30005
Private Const UIA_ClassNamePropertyId As Long = 30012
Private Const UIA_ValuePatternId As Long = 10002
Private Const UIA_DocumentControlTypeId As Long = 50030
Private Const UIA_EditControlTypeId As Long = 50004

Private m_Uia As UIAutomationClient.IUIAutomation
Private m_Doc As UIAutomationClient.IUIAutomationElement
Private m_Val As UIAutomationClient.IUIAutomationValuePattern
Private m_Hwnd As Variant
Private m_Title As String
Private m_Why As String

' why the last bind attempt did not produce a text field. Silently retrying
' forever is the one thing a watcher must not do: if it cannot attach, the
' screen has to say what stopped it.
Public Function Rdv2UiaWhy() As String
    Rdv2UiaWhy = m_Why
End Function

Public Function Rdv2UiaBound() As Boolean
    Rdv2UiaBound = Not (m_Val Is Nothing)
End Function

Public Function Rdv2UiaHwnd() As Variant
    Rdv2UiaHwnd = m_Hwnd
End Function

Public Function Rdv2UiaTitle() As String
    Rdv2UiaTitle = m_Title
End Function

Public Sub Rdv2UiaReset()
    Set m_Val = Nothing
    Set m_Doc = Nothing
    m_Hwnd = 0
    m_Title = ""
End Sub

' Same rule as Rdv2Watch.Bind() in the C# builds -- Notepad windows only, the
' one in the foreground wins so the operator can pick by clicking -- but the
' windows are enumerated with Win32, not with UIA, and that part is not a
' style choice.
'
' Measured here: m_Uia.GetRootElement.FindAll(TreeScope_Children, ...) never
' returns when it is called on Excel's own UI thread. Enumerating the desktop's
' children makes UIA ask every top-level window for its properties, and one of
' those windows is Excel's. Excel is a UIA provider, its provider lives on the
' thread that is already blocked inside the client call, and the two wait for
' each other. The C# builds do the same enumeration safely because they poll
' from a separate MTA thread; VBA has no second thread to move it to.
'
' FindWindowEx never enters UIA, and ElementFromHandle attaches to one window
' without walking the tree, so nothing ever asks Excel about itself.
Public Function Rdv2UiaBind() As Boolean
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
    Rdv2UiaBind = False
    m_Why = "uia"
    If m_Uia Is Nothing Then Set m_Uia = New UIAutomationClient.CUIAutomation

    m_Why = "find"
    fg = Rdv2ForegroundWindow()
    h = 0
    Do
        h = Rdv2FindWindowEx(0, h, "Notepad", vbNullString)
        If h = 0 Then Exit Do
        If Rdv2IsWindowVisible(h) <> 0 Then
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

    ' The window handle and the title come from Win32, not from
    ' pick.CurrentName / pick.CurrentNativeWindowHandle. Measured: reading a
    ' UIA property off a TOP-LEVEL window from Excel's own UI thread never
    ' returns, the same deadlock as walking the desktop. Reading from the
    ' Document element below it is fine, which is why polling works at all.
    Set m_Doc = d
    m_Hwnd = hPick
    m_Title = Rdv2WindowTitle(hPick)
    m_Why = ""
    Rdv2UiaBind = True
    Exit Function
Fail:
    m_Why = m_Why & " で失敗: " & Err.Number & " " & Err.Description
    Rdv2UiaReset
    Rdv2UiaBind = False
End Function

' one poll. Returns Null when the binding has gone stale, so the caller can
' tell "the field is empty" from "the window went away".
Public Function Rdv2UiaRead() As Variant
    On Error GoTo Fail
    If m_Val Is Nothing Then
        Rdv2UiaRead = Null
        Exit Function
    End If
    Rdv2UiaRead = m_Val.CurrentValue
    Exit Function
Fail:
    Rdv2UiaReset
    Rdv2UiaRead = Null
End Function
