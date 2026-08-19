VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbNotepad - つないでいるメモ帳 1 枚
'
' 責務：窓を 1 つだけ掴み、その本文を読み書きし、位置と大きさを変える。
' アプリはメモ帳を起動も終了もしない。既に開いている窓につなぐだけ。
'
' 二つの不変条件がある。壊すと利用者の文章が消える。
'
' 1. **いま掴んでいる窓が残っているかぎり、持ち替えない。** FindAll が返す
'    並びは Z 順で、利用者が窓をクリックしただけで入れ替わる。並び順で選ぶと、
'    入れ替わった先へ前の窓の本文を書き戻して文章を消す。掴んでいる窓が
'    居なくなったときだけ選び直す。
'
' 2. **結び直した直後の 1 回は必ずメモ帳 → セルの向きで揃える。** そこで
'    書き戻しの向きに入ると、画面側に残っていた古い文字で相手の本文を潰す。
'
' 最小化されても要素は手放さない。矩形は (-32000,-32000) へ飛ぶがハンドルは
' 生きているので、掴んだままにしておけば最大化でも範囲選択でも戻せる。
' 手放すと二度と戻せない。
'==============================================================================
Option Explicit

Private m_uia As PbUia
Private m_log As PbLog

Private m_win As UIAutomationClient.IUIAutomationElement
Private m_doc As UIAutomationClient.IUIAutomationElement
Private m_val As UIAutomationClient.IUIAutomationValuePattern
Private m_hwnd As Variant
Private m_title As String
Private m_pid As String
Private m_bound As Boolean
Private m_min As Boolean
Private m_fresh As Boolean
Private m_lastText As String
Private m_lastCell As String
Private m_l As Long, m_t As Long, m_r As Long, m_b As Long

Public Sub Init(ByVal uia As PbUia, ByVal logger As PbLog)
    Set m_uia = uia
    Set m_log = logger
End Sub

Public Property Get IsBound() As Boolean
    IsBound = m_bound
End Property

Public Property Get IsMinimized() As Boolean
    IsMinimized = m_min
End Property

Public Property Get Title() As String
    Title = m_title
End Property

Public Property Get Pid() As String
    Pid = m_pid
End Property

Public Property Get RectLeft() As Long
    RectLeft = m_l
End Property

Public Property Get RectTop() As Long
    RectTop = m_t
End Property

Public Property Get RectRight() As Long
    RectRight = m_r
End Property

Public Property Get RectBottom() As Long
    RectBottom = m_b
End Property

Public Property Get LastText() As String
    LastText = m_lastText
End Property

Public Property Get LastCell() As String
    LastCell = m_lastCell
End Property

Public Property Get IsFresh() As Boolean
    IsFresh = m_fresh
End Property

Public Property Get Element() As UIAutomationClient.IUIAutomationElement
    Set Element = m_win
End Property

'--- 掴み直し -----------------------------------------------------------------
' 見えている窓を数え、掴む相手を決め、矩形と最小化の別を取り込む。
' 戻り値は「見えたメモ帳の窓の数」。
Public Function Rescan() As Long
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim keepEl As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    Dim i As Long
    Dim n As Long
    Dim h As Variant
    Dim keepH As Variant

    On Error GoTo Failed
    Set arr = m_uia.FindNotepads()
    If arr Is Nothing Then Exit Function
    n = arr.Length

    ' 1 巡目：いま掴んでいる窓がまだ居るなら、それを使い続ける
    If m_bound Then
        For i = 0 To n - 1
            Set el = arr.GetElement(i)
            h = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
            If CStr(m_hwnd) = CStr(h) Then
                Set keepEl = el
                keepH = h
                Exit For
            End If
        Next i
    End If
    ' 2 巡目：居なければ選び直す。見えている窓を優先し、無ければ先頭
    If keepEl Is Nothing Then
        For i = 0 To n - 1
            Set el = arr.GetElement(i)
            rc = el.CurrentBoundingRectangle
            If rc.Right > rc.Left And rc.Bottom > rc.Top _
               And rc.Left > -10000 And Not CBool(el.CurrentIsOffscreen) Then
                Set keepEl = el
                keepH = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
                Exit For
            End If
        Next i
    End If
    If keepEl Is Nothing And n > 0 Then
        Set keepEl = arr.GetElement(0)
        keepH = keepEl.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    End If

    If keepEl Is Nothing Then
        Unbind
        Rescan = 0
        Exit Function
    End If

    Bind keepEl, keepH
    rc = keepEl.CurrentBoundingRectangle
    m_l = rc.Left: m_t = rc.Top: m_r = rc.Right: m_b = rc.Bottom
    m_min = (rc.Right <= rc.Left) Or (rc.Bottom <= rc.Top) Or (rc.Left < -10000) _
            Or CBool(keepEl.CurrentIsOffscreen)
    Rescan = n
    Exit Function
Failed:
    Rescan = 0
End Function

Private Sub Bind(ByVal win As UIAutomationClient.IUIAutomationElement, ByVal h As Variant)
    Dim d As UIAutomationClient.IUIAutomationElement
    On Error GoTo Failed
    Set m_win = win
    m_title = win.CurrentName
    m_pid = CStr(win.CurrentProcessId)
    If m_bound Then
        If CStr(m_hwnd) = CStr(h) Then Exit Sub      ' 同じ窓なら結び直さない
    End If
    Set d = m_uia.FindTextElement(win)
    If d Is Nothing Then
        m_bound = False
        Exit Sub
    End If
    ' パターンは必ず型どおりに受ける。UIA のパターンは IUnknown 系で IDispatch を
    ' 持たないため、As Object（遅延バインディング）で受けると Set の時点で失敗し、
    ' 「何をしても動かない」になる。
    Set m_val = d.GetCurrentPattern(UIA_ValuePatternId)
    If m_val Is Nothing Then
        m_bound = False
        Exit Sub
    End If
    Set m_doc = d
    m_hwnd = h
    m_bound = True
    m_fresh = True
    m_lastText = ""
    m_lastCell = ""
    Exit Sub
Failed:
    m_bound = False
End Sub

Private Sub Unbind()
    m_bound = False
    m_min = False
    m_title = ""
    Set m_win = Nothing
    Set m_doc = Nothing
    Set m_val = Nothing
End Sub

'--- 本文 ---------------------------------------------------------------------
Public Function ReadText() As String
    On Error Resume Next
    If Not m_bound Then Exit Function
    ReadText = ToLf(CStr(m_val.CurrentValue))
End Function

Public Function WriteText(ByVal s As String) As Boolean
    On Error GoTo Failed
    If Not m_bound Then Exit Function
    m_val.SetValue ToCrLf(s)
    m_lastCell = s
    m_lastText = ToLf(s)
    WriteText = True
    Exit Function
Failed:
    m_bound = False
    WriteText = False
End Function

' セルへ写したあとに控えを揃える。控えは 2 つの正規形で持つ。
'   メモ帳側 … LF 形（ValuePattern が LF で返すため）
'   セル側   … 「書いた値」ではなく「書いた結果セルが返す値」。Excel は代入で
'              先頭の ' を落とすなど値を黙って直すことがあり、意図した値で
'              控えると次のティックが「セルが編集された」と誤読して、直った
'              あとの文字列を相手へ書き戻してしまう。
Public Sub NoteMirrored(ByVal npText As String, ByVal cellReadBack As String)
    m_lastText = ToLf(npText)
    m_lastCell = cellReadBack
    m_fresh = False
End Sub

Public Sub MarkDisconnected()
    m_bound = False
End Sub

Public Function HasFocus() As Boolean
    On Error Resume Next
    If Not m_bound Then Exit Function
    HasFocus = CBool(m_doc.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId))
End Function

'--- 窓の操作 -----------------------------------------------------------------
' 最大化されていると CanMove が False になるので、先に WindowPattern で
' 通常表示へ戻す。最小化から戻すのも同じ経路。
Public Function MoveResize(ByVal x As Long, ByVal y As Long, _
                           ByVal w As Long, ByVal h As Long) As Boolean
    Dim wp As UIAutomationClient.IUIAutomationWindowPattern
    Dim tp As UIAutomationClient.IUIAutomationTransformPattern

    On Error GoTo Failed
    If m_win Is Nothing Then
        m_log.Note "  movewin: el nothing"
        Exit Function
    End If
    Set wp = m_win.GetCurrentPattern(UIA_WindowPatternId)
    If Not wp Is Nothing Then
        On Error Resume Next
        wp.SetWindowVisualState WindowVisualState_Normal
        On Error GoTo Failed
    End If
    Set tp = m_win.GetCurrentPattern(UIA_TransformPatternId)
    If tp Is Nothing Then
        ' 取れなかったときは「相手が公開していない」のか「こちらの訊き方が
        ' 悪い」のかをログで分ける。IsTransformPatternAvailable が True なのに
        ' パターンが Nothing なら、それは訊き方（ID・受け型）の側の問題。
        m_log.Note "  movewin: no transform (avail=" & _
            CStr(m_win.GetCurrentPropertyValue(UIA_IsTransformPatternAvailablePropertyId)) & ")"
        Exit Function
    End If
    If Not CBool(tp.CurrentCanMove) Then
        m_log.Note "  movewin: canmove false"
        Exit Function
    End If
    tp.Move CDbl(x), CDbl(y)
    If CBool(tp.CurrentCanResize) Then tp.Resize CDbl(w), CDbl(h)
    m_min = False
    m_log.Note "  movewin: ok " & x & "," & y & " " & w & "x" & h
    MoveResize = True
    Exit Function
Failed:
    m_log.Note "  movewin: err " & Err.Number & " " & Err.Description
    MoveResize = False
End Function

Public Function Minimize() As Boolean
    Dim wp As UIAutomationClient.IUIAutomationWindowPattern
    On Error GoTo Failed
    If m_win Is Nothing Then Exit Function
    Set wp = m_win.GetCurrentPattern(UIA_WindowPatternId)
    If wp Is Nothing Then Exit Function
    wp.SetWindowVisualState WindowVisualState_Minimized
    m_min = True
    Minimize = True
    Exit Function
Failed:
    Minimize = False
End Function

' 前面へ。相手のプロセスの窓に対して使う（自分の窓の要素は作らない、という
' 原則は守っている）。
Public Sub Focus()
    On Error Resume Next
    If Not m_win Is Nothing Then m_win.SetFocus
End Sub

'--- 解放 ---------------------------------------------------------------------
' UIA クライアントを捨てる前に、掴んだ要素をここで先に手放す。
Public Sub ReleaseAll()
    On Error Resume Next
    m_bound = False
    Set m_val = Nothing
    Set m_doc = Nothing
    Set m_win = Nothing
End Sub
