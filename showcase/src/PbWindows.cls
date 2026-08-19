VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbWindows - ミニマップに映す窓の一覧
'
' 責務：「いま画面のどこに何があるか」を毎スキャンで作り直す。中身は
' この Excel と、つないでいるメモ帳 1 枚だけ。
'
' 自分の窓の矩形は **Application の座標（ポイント）から作る。** UIA の
' ElementFromHandle(自分の hwnd) は使わない。自分の窓の UIA 要素を作ると
' Excel 自身のプロバイダがこのプロセスで動き出し、以後のセル書き込みが
' 1 回 130ms 級になる。
'
' ベンチの表示用 Excel はここへ入れない。あれは表示専用で、操作の対象に
' 混ぜない。
'==============================================================================
Option Explicit

Private m_uia As PbUia
Private m_np As PbNotepad
Private m_canvas As PbCanvas

Private m_n As Long
Private m_l(0 To PB_MAXWIN - 1) As Long
Private m_t(0 To PB_MAXWIN - 1) As Long
Private m_r(0 To PB_MAXWIN - 1) As Long
Private m_b(0 To PB_MAXWIN - 1) As Long
Private m_label(0 To PB_MAXWIN - 1) As String
Private m_pid(0 To PB_MAXWIN - 1) As String
Private m_isNp(0 To PB_MAXWIN - 1) As Boolean

Public Sub Init(ByVal uia As PbUia, ByVal np As PbNotepad, ByVal canvas As PbCanvas)
    Set m_uia = uia
    Set m_np = np
    Set m_canvas = canvas
End Sub

Public Property Get Count() As Long
    Count = m_n
End Property

Public Property Get L(ByVal i As Long) As Long
    L = m_l(i)
End Property

Public Property Get T(ByVal i As Long) As Long
    T = m_t(i)
End Property

Public Property Get R(ByVal i As Long) As Long
    R = m_r(i)
End Property

Public Property Get B(ByVal i As Long) As Long
    B = m_b(i)
End Property

Public Property Get Label(ByVal i As Long) As String
    Label = m_label(i)
End Property

Public Property Get PidOf(ByVal i As Long) As String
    PidOf = m_pid(i)
End Property

Public Property Get IsNotepad(ByVal i As Long) As Boolean
    IsNotepad = m_isNp(i)
End Property

Public Sub Rebuild()
    Dim ppp As Double
    m_n = 0
    ppp = m_canvas.PxPerPt
    If ppp > 0 Then
        Add CLng(Application.Left * ppp), CLng(Application.Top * ppp), _
            CLng((Application.Left + Application.Width) * ppp), _
            CLng((Application.Top + Application.Height) * ppp), _
            "この Excel", m_uia.OwnPidText, False
    End If
    If m_np.IsBound And Not m_np.IsMinimized Then
        Add m_np.RectLeft, m_np.RectTop, m_np.RectRight, m_np.RectBottom, _
            m_np.Title, m_np.Pid, True
    End If
End Sub

Private Sub Add(ByVal l1 As Long, ByVal t1 As Long, ByVal r1 As Long, ByVal b1 As Long, _
                ByVal lbl As String, ByVal pidText As String, ByVal isNp As Boolean)
    If m_n >= PB_MAXWIN Then Exit Sub
    m_l(m_n) = l1: m_t(m_n) = t1: m_r(m_n) = r1: m_b(m_n) = b1
    m_label(m_n) = lbl
    m_pid(m_n) = pidText
    m_isNp(m_n) = isNp
    m_n = m_n + 1
End Sub

' 並びが前と同じなら塗り替えない、の判定に使う。
Public Property Get Signature() As String
    Dim s As String
    Dim i As Long
    For i = 0 To m_n - 1
        s = s & m_l(i) & "," & m_t(i) & "," & m_r(i) & "," & m_b(i) & ";"
    Next i
    Signature = s & "|" & m_n
End Property

' 画面上の 1 点がどの窓の中か。無ければ -1。
Public Function HitTest(ByVal px As Long, ByVal py As Long) As Long
    Dim i As Long
    HitTest = -1
    For i = 0 To m_n - 1
        If px >= m_l(i) And px < m_r(i) And py >= m_t(i) And py < m_b(i) Then
            HitTest = i
            Exit Function
        End If
    Next i
End Function
