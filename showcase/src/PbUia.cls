VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbUia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbUia - UI Automation のクライアントと、画面そのもの
'
' 責務：UIA クライアントを 1 つだけ作って持ち、画面の矩形と「メモ帳の窓」の
' 列挙を提供する。掴んだ相手の面倒は PbNotepad が見る。
'
' 二つの制約がこのクラスの形を決めている。
'
' 1. **自分の窓の UIA 要素を作らない。** 作ると Excel 自身の UIA プロバイダが
'    このプロセスの中で動き出し、以後セルへの書き込み 1 回ごとに 130ms 級の
'    同期コストが乗る（素の盤なら 5ms）。1 秒のティックはそれで壊れる。
'    だから自分の窓の矩形は Application の座標から作り、フォーカスの判定も
'    相手（メモ帳）側の要素に尋ねる。唯一の例外が起動時のプロセス番号取得で、
'    そこは 1 回だけ作ってすぐ手放す。
'
' 2. **デスクトップ直下の列挙は詰まり得る。** UIA はすべてのトップレベル窓の
'    プロバイダに尋ねるので、応答しないプロセスが 1 つあるとそこで止まる。
'    ClassName で絞っても速さはほとんど変わらない（118.1ms 対 118.8ms）が、
'    触りにいく相手の数は 22 から 2 に減る。呼ぶ回数自体も抑える。
'==============================================================================
Option Explicit

Private m_uia As UIAutomationClient.IUIAutomation
Private m_root As UIAutomationClient.IUIAutomationElement
Private m_log As PbLog
Private m_note As String
Private m_ownPid As String
Private m_scrL As Long, m_scrT As Long, m_scrR As Long, m_scrB As Long

Public Sub Init(ByVal logger As PbLog)
    Set m_log = logger
End Sub

Public Function Ensure() As Boolean
    On Error GoTo Failed
    If m_uia Is Nothing Then Set m_uia = New UIAutomationClient.CUIAutomation
    If m_root Is Nothing Then Set m_root = m_uia.GetRootElement
    Ensure = (Not m_uia Is Nothing) And (Not m_root Is Nothing)
    Exit Function
Failed:
    m_note = "UIA 接続不可 " & Err.Number
    Ensure = False
End Function

Public Property Get Client() As UIAutomationClient.IUIAutomation
    Set Client = m_uia
End Property

Public Property Get Note() As String
    Note = m_note
End Property

Public Property Let Note(ByVal s As String)
    m_note = s
End Property

'--- 画面 ---------------------------------------------------------------------
Public Property Get ScreenLeft() As Long
    ScreenLeft = m_scrL
End Property

Public Property Get ScreenTop() As Long
    ScreenTop = m_scrT
End Property

Public Property Get ScreenRight() As Long
    ScreenRight = m_scrR
End Property

Public Property Get ScreenBottom() As Long
    ScreenBottom = m_scrB
End Property

Public Property Get ScreenWidth() As Long
    ScreenWidth = m_scrR - m_scrL
End Property

Public Property Get ScreenHeight() As Long
    ScreenHeight = m_scrB - m_scrT
End Property

Public Property Get HasScreen() As Boolean
    HasScreen = (m_scrR > m_scrL) And (m_scrB > m_scrT)
End Property

' 画面の大きさ（デバイス px）。UIA の root の矩形がそのまま画面全体になる。
' 取れない端末では、最大化した窓の大きさで代用する。pxPerPt はそのときだけ使う。
Public Sub ReadScreen(ByVal pxPerPt As Double)
    Dim rc As UIAutomationClient.tagRECT
    On Error GoTo Fallback
    If Not Ensure() Then GoTo Fallback
    rc = m_root.CurrentBoundingRectangle
    If rc.Right - rc.Left < 320 Or rc.Bottom - rc.Top < 240 Then GoTo Fallback
    m_scrL = rc.Left: m_scrT = rc.Top: m_scrR = rc.Right: m_scrB = rc.Bottom
    Exit Sub
Fallback:
    On Error Resume Next
    Application.WindowState = xlMaximized
    m_scrL = 0: m_scrT = 0
    m_scrR = CLng(Application.Width * pxPerPt)
    m_scrB = CLng(Application.Height * pxPerPt)
    If m_scrR < 320 Then m_scrR = 1280
    If m_scrB < 240 Then m_scrB = 720
End Sub

'--- 窓の列挙 -----------------------------------------------------------------
Public Function FindNotepads() As UIAutomationClient.IUIAutomationElementArray
    Dim cond As UIAutomationClient.IUIAutomationCondition
    On Error Resume Next
    If Not Ensure() Then Exit Function
    Set cond = m_uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Notepad")
    Set FindNotepads = m_root.FindAll(TS_CHILDREN, cond)
End Function

' メモ帳の文書要素（Document か Edit）。ValuePattern はここに付いている。
Public Function FindTextElement(ByVal win As UIAutomationClient.IUIAutomationElement) _
        As UIAutomationClient.IUIAutomationElement
    Dim c1 As UIAutomationClient.IUIAutomationCondition
    Dim c2 As UIAutomationClient.IUIAutomationCondition
    On Error Resume Next
    If m_uia Is Nothing Or win Is Nothing Then Exit Function
    Set c1 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_DocumentControlTypeId)
    Set c2 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId)
    Set FindTextElement = win.FindFirst(TS_DESCENDANTS, m_uia.CreateOrCondition(c1, c2))
End Function

'--- 自分のプロセス番号 -------------------------------------------------------
' Win32 なしで取れる唯一の口が UIA なので、起動時に 1 回だけ自分の窓の要素を
' 作って読み、すぐ手放す。持ち続けたり毎ティック作り直したりすると、その間
' ずっと書き込みが 130ms 級になる。
Public Sub ReadOwnPid()
    Dim el As UIAutomationClient.IUIAutomationElement
    On Error Resume Next
    m_ownPid = ""
    If Not Ensure() Then Exit Sub
    Set el = m_uia.ElementFromHandle(ByVal Application.hwnd)
    If el Is Nothing Then Exit Sub
    m_ownPid = CStr(el.CurrentProcessId)
    Set el = Nothing
End Sub

Public Property Get OwnPidText() As String
    If Len(m_ownPid) = 0 Then
        OwnPidText = "―"
    Else
        OwnPidText = m_ownPid
    End If
End Property

'--- 解放 ---------------------------------------------------------------------
' プロセスをまたぐ COM 参照なので、握ったままだと Excel は終われない
' （後始末は 2 秒で終わっているのにプロセスだけが 45 秒残った）。
' 呼ぶ順序は「掴んだ要素 → ここ」。逆にしない。
Public Sub ReleaseAll()
    On Error Resume Next
    Set m_root = Nothing
    Set m_uia = Nothing
End Sub
