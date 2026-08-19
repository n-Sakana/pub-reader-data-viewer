VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbWorkspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbWorkspace - 画面の外側（デスクトップ）でこのアプリが扱うものすべて
'
' 責務：UI Automation のクライアントを 1 つだけ持ち、画面の矩形を測り、
' つないでいるメモ帳 1 枚を掴んで読み書き・移動し、ミニマップに映す窓の
' 一覧を作る。どれも「外の窓」を相手にする仕事で、寿命も資源の境界も同じ
' 一つ（UIA クライアントとその要素）なので、まとめて持つ。
'
' 四つの制約がこのクラスの形を決めている。
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
'
' 3. **いま掴んでいる窓が残っているかぎり、持ち替えない。** FindAll が返す
'    並びは Z 順で、利用者が窓をクリックしただけで入れ替わる。並び順で選ぶと、
'    入れ替わった先へ前の窓の本文を書き戻して文章を消す。掴んでいる窓が
'    居なくなったときだけ選び直す。結び直した直後の 1 回は必ずメモ帳 → セルの
'    向きで揃える。逆向きに入ると、画面側に残っていた古い文字で相手の本文を潰す。
'
' 4. **最小化されても要素は手放さない。** 矩形は (-32000,-32000) へ飛ぶが
'    ハンドルは生きているので、掴んだままにしておけば最大化でも範囲選択でも
'    戻せる。手放すと二度と戻せない。
'
' アプリはメモ帳を起動も終了もしない。既に開いている窓につなぐだけ。
'==============================================================================
Option Explicit

Private m_log As PbLog
Private m_canvas As PbCanvas

'--- UIA クライアントと画面
Private m_uia As UIAutomationClient.IUIAutomation
Private m_root As UIAutomationClient.IUIAutomationElement
Private m_ownPid As String
Private m_scrL As Long, m_scrT As Long, m_scrR As Long, m_scrB As Long

'--- 掴んでいるメモ帳 1 枚
Private m_npWin As UIAutomationClient.IUIAutomationElement
Private m_npDoc As UIAutomationClient.IUIAutomationElement
Private m_npVal As UIAutomationClient.IUIAutomationValuePattern
Private m_npHwnd As Variant
Private m_npTitle As String
Private m_npPid As String
Private m_npBound As Boolean
Private m_npMin As Boolean
Private m_npFresh As Boolean
Private m_npText As String
Private m_npCell As String
Private m_npL As Long, m_npT As Long, m_npR As Long, m_npB As Long

'--- ミニマップに映す窓の一覧
Private m_winN As Long
Private m_winL(0 To PB_MAXWIN - 1) As Long
Private m_winT(0 To PB_MAXWIN - 1) As Long
Private m_winR(0 To PB_MAXWIN - 1) As Long
Private m_winB(0 To PB_MAXWIN - 1) As Long
Private m_winLabel(0 To PB_MAXWIN - 1) As String
Private m_winPid(0 To PB_MAXWIN - 1) As String
Private m_winIsNp(0 To PB_MAXWIN - 1) As Boolean

Public Sub Init(ByVal logger As PbLog, ByVal canvas As PbCanvas)
    Set m_log = logger
    Set m_canvas = canvas
End Sub

'==============================================================================
' UIA クライアントと画面
'==============================================================================
Private Function Ensure() As Boolean
    On Error GoTo Failed
    If m_uia Is Nothing Then Set m_uia = New UIAutomationClient.CUIAutomation
    If m_root Is Nothing Then Set m_root = m_uia.GetRootElement
    Ensure = (Not m_uia Is Nothing) And (Not m_root Is Nothing)
    Exit Function
Failed:
    Ensure = False
End Function

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

' Win32 なしで自分のプロセス番号を取れる唯一の口が UIA なので、起動時に 1 回
' だけ自分の窓の要素を作って読み、すぐ手放す。持ち続けたり毎ティック作り直したり
' すると、その間ずっと書き込みが 130ms 級になる。
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

Private Function FindNotepads() As UIAutomationClient.IUIAutomationElementArray
    Dim cond As UIAutomationClient.IUIAutomationCondition
    On Error Resume Next
    If Not Ensure() Then Exit Function
    Set cond = m_uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Notepad")
    Set FindNotepads = m_root.FindAll(TS_CHILDREN, cond)
End Function

' メモ帳の文書要素（Document か Edit）。ValuePattern はここに付いている。
Private Function FindTextElement(ByVal win As UIAutomationClient.IUIAutomationElement) _
        As UIAutomationClient.IUIAutomationElement
    Dim c1 As UIAutomationClient.IUIAutomationCondition
    Dim c2 As UIAutomationClient.IUIAutomationCondition
    On Error Resume Next
    If m_uia Is Nothing Or win Is Nothing Then Exit Function
    Set c1 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_DocumentControlTypeId)
    Set c2 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId)
    Set FindTextElement = win.FindFirst(TS_DESCENDANTS, m_uia.CreateOrCondition(c1, c2))
End Function

'==============================================================================
' つないでいるメモ帳 1 枚
'==============================================================================
Public Property Get IsBound() As Boolean
    IsBound = m_npBound
End Property

Public Property Get IsMinimized() As Boolean
    IsMinimized = m_npMin
End Property

Public Property Get Title() As String
    Title = m_npTitle
End Property

Public Property Get Pid() As String
    Pid = m_npPid
End Property

Public Property Get LastText() As String
    LastText = m_npText
End Property

Public Property Get LastCell() As String
    LastCell = m_npCell
End Property

Public Property Get IsFresh() As Boolean
    IsFresh = m_npFresh
End Property

Public Property Get Element() As UIAutomationClient.IUIAutomationElement
    Set Element = m_npWin
End Property

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
    Set arr = FindNotepads()
    If arr Is Nothing Then Exit Function
    n = arr.Length

    ' 1 巡目：いま掴んでいる窓がまだ居るなら、それを使い続ける
    If m_npBound Then
        For i = 0 To n - 1
            Set el = arr.GetElement(i)
            h = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
            If CStr(m_npHwnd) = CStr(h) Then
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
    m_npL = rc.Left: m_npT = rc.Top: m_npR = rc.Right: m_npB = rc.Bottom
    m_npMin = (rc.Right <= rc.Left) Or (rc.Bottom <= rc.Top) Or (rc.Left < -10000) _
              Or CBool(keepEl.CurrentIsOffscreen)
    Rescan = n
    Exit Function
Failed:
    Rescan = 0
End Function

Private Sub Bind(ByVal win As UIAutomationClient.IUIAutomationElement, ByVal h As Variant)
    Dim d As UIAutomationClient.IUIAutomationElement
    On Error GoTo Failed
    Set m_npWin = win
    m_npTitle = win.CurrentName
    m_npPid = CStr(win.CurrentProcessId)
    If m_npBound Then
        If CStr(m_npHwnd) = CStr(h) Then Exit Sub    ' 同じ窓なら結び直さない
    End If
    Set d = FindTextElement(win)
    If d Is Nothing Then
        m_npBound = False
        Exit Sub
    End If
    ' パターンは必ず型どおりに受ける。UIA のパターンは IUnknown 系で IDispatch を
    ' 持たないため、As Object（遅延バインディング）で受けると Set の時点で失敗し、
    ' 「何をしても動かない」になる。
    Set m_npVal = d.GetCurrentPattern(UIA_ValuePatternId)
    If m_npVal Is Nothing Then
        m_npBound = False
        Exit Sub
    End If
    Set m_npDoc = d
    m_npHwnd = h
    m_npBound = True
    m_npFresh = True
    m_npText = ""
    m_npCell = ""
    Exit Sub
Failed:
    m_npBound = False
End Sub

Private Sub Unbind()
    m_npBound = False
    m_npMin = False
    m_npTitle = ""
    Set m_npWin = Nothing
    Set m_npDoc = Nothing
    Set m_npVal = Nothing
End Sub

Public Function ReadText() As String
    On Error Resume Next
    If Not m_npBound Then Exit Function
    ReadText = ToLf(CStr(m_npVal.CurrentValue))
End Function

Public Function WriteText(ByVal s As String) As Boolean
    On Error GoTo Failed
    If Not m_npBound Then Exit Function
    m_npVal.SetValue ToCrLf(s)
    m_npCell = s
    m_npText = ToLf(s)
    WriteText = True
    Exit Function
Failed:
    m_npBound = False
    WriteText = False
End Function

' セルへ写したあとに控えを揃える。控えは 2 つの正規形で持つ。
'   メモ帳側 … LF 形（ValuePattern が LF で返すため）
'   セル側   … 「書いた値」ではなく「書いた結果セルが返す値」。Excel は代入で
'              先頭の ' を落とすなど値を黙って直すことがあり、意図した値で
'              控えると次のティックが「セルが編集された」と誤読して、直った
'              あとの文字列を相手へ書き戻してしまう。
Public Sub NoteMirrored(ByVal npText As String, ByVal cellReadBack As String)
    m_npText = ToLf(npText)
    m_npCell = cellReadBack
    m_npFresh = False
End Sub

Public Sub MarkDisconnected()
    m_npBound = False
End Sub

Public Function HasFocus() As Boolean
    On Error Resume Next
    If Not m_npBound Then Exit Function
    HasFocus = CBool(m_npDoc.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId))
End Function

' 最大化されていると CanMove が False になるので、先に WindowPattern で
' 通常表示へ戻す。最小化から戻すのも同じ経路。
Public Function MoveResize(ByVal x As Long, ByVal y As Long, _
                           ByVal w As Long, ByVal h As Long) As Boolean
    Dim wp As UIAutomationClient.IUIAutomationWindowPattern
    Dim tp As UIAutomationClient.IUIAutomationTransformPattern

    On Error GoTo Failed
    If m_npWin Is Nothing Then
        m_log.Note "  movewin: el nothing"
        Exit Function
    End If
    Set wp = m_npWin.GetCurrentPattern(UIA_WindowPatternId)
    If Not wp Is Nothing Then
        On Error Resume Next
        wp.SetWindowVisualState WindowVisualState_Normal
        On Error GoTo Failed
    End If
    Set tp = m_npWin.GetCurrentPattern(UIA_TransformPatternId)
    If tp Is Nothing Then
        ' 取れなかったときは「相手が公開していない」のか「こちらの訊き方が
        ' 悪い」のかをログで分ける。IsTransformPatternAvailable が True なのに
        ' パターンが Nothing なら、それは訊き方（ID・受け型）の側の問題。
        m_log.Note "  movewin: no transform (avail=" & _
            CStr(m_npWin.GetCurrentPropertyValue(UIA_IsTransformPatternAvailablePropertyId)) & ")"
        Exit Function
    End If
    If Not CBool(tp.CurrentCanMove) Then
        m_log.Note "  movewin: canmove false"
        Exit Function
    End If
    tp.Move CDbl(x), CDbl(y)
    If CBool(tp.CurrentCanResize) Then tp.Resize CDbl(w), CDbl(h)
    m_npMin = False
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
    If m_npWin Is Nothing Then Exit Function
    Set wp = m_npWin.GetCurrentPattern(UIA_WindowPatternId)
    If wp Is Nothing Then Exit Function
    wp.SetWindowVisualState WindowVisualState_Minimized
    m_npMin = True
    Minimize = True
    Exit Function
Failed:
    Minimize = False
End Function

' 前面へ。相手のプロセスの窓に対して使う（自分の窓の要素は作らない、という
' 原則は守っている）。
Public Sub Focus()
    On Error Resume Next
    If Not m_npWin Is Nothing Then m_npWin.SetFocus
End Sub

'==============================================================================
' ミニマップに映す窓の一覧
'
' 中身はこの Excel と、つないでいるメモ帳 1 枚だけ。自分の窓の矩形は
' Application の座標（ポイント）から作る（上の制約 1）。ベンチの表示用 Excel は
' ここへ入れない。あれは表示専用で、操作の対象に混ぜない。
'==============================================================================
Public Sub RebuildWindows()
    Dim ppp As Double
    m_winN = 0
    ppp = m_canvas.PxPerPt
    If ppp > 0 Then
        AddWindow CLng(Application.Left * ppp), CLng(Application.Top * ppp), _
                  CLng((Application.Left + Application.Width) * ppp), _
                  CLng((Application.Top + Application.Height) * ppp), _
                  "この Excel", OwnPidText, False
    End If
    If m_npBound And Not m_npMin Then
        AddWindow m_npL, m_npT, m_npR, m_npB, m_npTitle, m_npPid, True
    End If
End Sub

Private Sub AddWindow(ByVal l1 As Long, ByVal t1 As Long, ByVal r1 As Long, ByVal b1 As Long, _
                      ByVal lbl As String, ByVal pidText As String, ByVal isNp As Boolean)
    If m_winN >= PB_MAXWIN Then Exit Sub
    m_winL(m_winN) = l1: m_winT(m_winN) = t1: m_winR(m_winN) = r1: m_winB(m_winN) = b1
    m_winLabel(m_winN) = lbl
    m_winPid(m_winN) = pidText
    m_winIsNp(m_winN) = isNp
    m_winN = m_winN + 1
End Sub

Public Property Get Count() As Long
    Count = m_winN
End Property

Public Property Get L(ByVal i As Long) As Long
    L = m_winL(i)
End Property

Public Property Get T(ByVal i As Long) As Long
    T = m_winT(i)
End Property

Public Property Get R(ByVal i As Long) As Long
    R = m_winR(i)
End Property

Public Property Get B(ByVal i As Long) As Long
    B = m_winB(i)
End Property

Public Property Get Label(ByVal i As Long) As String
    Label = m_winLabel(i)
End Property

Public Property Get PidOf(ByVal i As Long) As String
    PidOf = m_winPid(i)
End Property

Public Property Get IsNotepad(ByVal i As Long) As Boolean
    IsNotepad = m_winIsNp(i)
End Property

' 並びが前と同じなら塗り替えない、の判定に使う。
Public Property Get Signature() As String
    Dim s As String
    Dim i As Long
    For i = 0 To m_winN - 1
        s = s & m_winL(i) & "," & m_winT(i) & "," & m_winR(i) & "," & m_winB(i) & ";"
    Next i
    Signature = s & "|" & m_winN
End Property

' 画面上の 1 点がどの窓の中か。無ければ -1。
Public Function HitTest(ByVal px As Long, ByVal py As Long) As Long
    Dim i As Long
    HitTest = -1
    For i = 0 To m_winN - 1
        If px >= m_winL(i) And px < m_winR(i) And py >= m_winT(i) And py < m_winB(i) Then
            HitTest = i
            Exit Function
        End If
    Next i
End Function

'==============================================================================
' 解放
'
' プロセスをまたぐ COM 参照なので、握ったままだと Excel は終われない
' （後始末は 2 秒で終わっているのにプロセスだけが 45 秒残った）。順序は
' 「掴んだ要素 → クライアント」。逆にしない。1 つのクラスにまとめてあるのは、
' この順序を呼ぶ側の書き方に委ねないため。
'==============================================================================
Public Sub ReleaseAll()
    On Error Resume Next
    m_npBound = False
    Set m_npVal = Nothing
    Set m_npDoc = Nothing
    Set m_npWin = Nothing
    Set m_root = Nothing
    Set m_uia = Nothing
End Sub
