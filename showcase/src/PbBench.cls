VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbBench"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbBench - 円周率ベンチ（15 秒）
'
' 責務：配置と準備、15 秒の計測、桁の受け渡し、そして後始末。計算そのものは
' BE 側（modPbBackend）で、このクラスは 1 桁も計算しない。
'
' 見せたいことは 1 つだけ。**非表示の別プロセス Excel が重い計算をしている
' 最中も、FE の 1 秒アニメ・メモ帳との双方向同期・窓の移動 / リサイズが
' 止まらない。**
'
' 登場人物
'   FE Excel     … 疑似ピクセルアプリ。1 秒ポンプで全部を回す
'   BE Excel     … 別プロセス・非表示。15 秒だけ円周率を実際に計算する
'   表示用 Excel … BE とは別の、保存しない一時ブック。書くのは FE
'   メモ帳       … つないでいる 1 枚
'
' 数え方の約束
'   「計算できた桁」と「表示できた桁」は別に数える。**正式な結果は後者。**
'   内部で何桁進んでいても、表示用 Excel まで届いていない桁は数えない。
'   そして 15 秒を越えたティックでは表示を足さない。終わりを見張るのは 1 秒
'   ポンプなので、最後に数えられるのは 14 秒台のティックまでになる。
'==============================================================================
Option Explicit

Private m_screen As PbScreen
Private m_ch As PbChannel
Private m_be As PbBeSession
Private m_disp As PbDisplayBook
Private m_np As PbNotepad
Private m_uia As PbUia
Private m_canvas As PbCanvas
Private m_log As PbLog

Private m_running As Boolean
Private m_t0 As Double
Private m_calc As Long           ' BE が計算した桁数（先頭の 3 を含む）
Private m_shown As Long          ' 表示用 Excel へ実際に書いた小数以下の桁数
Private m_lastProg As String

Public Sub Init(ByVal screen As PbScreen, ByVal channel As PbChannel, _
                ByVal be As PbBeSession, ByVal disp As PbDisplayBook, _
                ByVal np As PbNotepad, ByVal uia As PbUia, _
                ByVal canvas As PbCanvas, ByVal logger As PbLog)
    Set m_screen = screen
    Set m_ch = channel
    Set m_be = be
    Set m_disp = disp
    Set m_np = np
    Set m_uia = uia
    Set m_canvas = canvas
    Set m_log = logger
End Sub

Public Property Get IsRunning() As Boolean
    IsRunning = m_running
End Property

'--- 開始 ---------------------------------------------------------------------
Public Sub Start()
    If m_running Then Exit Sub
    On Error GoTo Failed
    m_screen.SetText "pb_jobstat", "用意しています ･ 表示用 Excel を作成中"
    m_screen.ClearLamps
    m_calc = 0
    m_shown = 0
    m_lastProg = ""

    If Not m_be.IsStarted Then m_be.Start
    If Not m_be.IsStarted Then
        m_screen.SetText "pb_jobstat", "BE を起こせませんでした ･ " & m_be.StatusText
        Exit Sub
    End If
    m_ch.KillQuiet m_ch.PiPath
    m_ch.KillQuiet m_ch.ProgPath
    m_be.ForgetProgress

    If Not m_disp.Ensure() Then
        m_screen.SetText "pb_jobstat", "表示用 Excel を作れませんでした"
        m_screen.SetText "pb_disp_state", "失敗"
        Exit Sub
    End If
    m_screen.SetText "pb_disp_state", "作成"
    Layout

    ' ここまでが「配置と準備」。15 秒はこの下から測る。
    m_t0 = Timer
    m_running = True
    m_ch.SendCommand "pi", """secs"":" & CLng(PB_BENCHSECS)
    m_screen.Flash "pb_lane_out"
    m_log.Note "bench: started"
    m_be.StatusText = "計算中"
    m_screen.SetText "pb_be_state", "計算中"
    m_screen.SetText "pb_jobstat", "計算中 ･ 経過 0.0 秒"
    Exit Sub
Failed:
    m_running = False
    m_log.Note "bench: start failed " & Err.Number & " " & Err.Description
    m_screen.SetText "pb_jobstat", "開始に失敗 " & Err.Number
    m_disp.CloseQuiet
    m_screen.SetText "pb_disp_state", "終了"
End Sub

' 開始時の配置。FE は左半分のまま動かさない（起動時に置いたきり）。
' 表示用 Excel を右上 4 分の 1、メモ帳を右下 4 分の 1 へ。三つが重ならずに
' 画面を埋めるので、どれも隠れない。
Private Sub Layout()
    Dim hx As Long, hy As Long, hw As Long, hh As Long

    On Error Resume Next
    hx = m_uia.ScreenLeft + m_uia.ScreenWidth \ 2
    hw = m_uia.ScreenRight - hx
    hy = m_uia.ScreenTop
    hh = m_uia.ScreenHeight \ 2

    m_disp.PlaceAt hx, hy, hw, hh, m_canvas.PxPerPt
    If m_np.IsBound Then
        m_np.MoveResize hx, hy + hh, hw, m_uia.ScreenHeight - hh
        m_np.Focus
    End If
    m_screen.InvalidateMap
    m_log.Note "bench: layout disp=" & hx & "," & hy & " " & hw & "x" & hh
End Sub

'--- 毎ティック ---------------------------------------------------------------
Public Sub Pump()
    Dim el As Double
    Dim js As String

    If Not m_running Then Exit Sub
    On Error GoTo Failed
    el = ElapsedSince(m_t0)
    ' 15 秒を過ぎたティックでは、もう表示しない（上の「数え方の約束」を参照）。
    If el >= PB_BENCHSECS Then
        Finish
        Exit Sub
    End If

    js = m_ch.ReadProgress()
    If Len(js) > 0 And js <> m_lastProg Then
        m_lastProg = js
        m_screen.Flash "pb_lane_in"
        m_screen.Unflash "pb_lane_out"
        m_calc = CLng(Val(JVal(js, "n")))
        If JVal(js, "state") = "error" Then
            m_screen.SetText "pb_jobstat", "BE エラー " & JVal(js, "msg")
            Finish
            Exit Sub
        End If
    Else
        m_screen.Unflash "pb_lane_in"
    End If

    ShowDigits m_ch.ReadDigits()

    m_screen.SetText "pb_jobstat", "計算中 ･ 経過 " & Format$(el, "0.0") & " 秒 ･ 計算 " & _
        Decimals(m_calc) & " 桁 ･ 表示 " & m_shown & " 桁"
    Exit Sub
Failed:
    m_log.Note "bench: pump err " & Err.Number & " " & Err.Description
    Finish
End Sub

' まだ表示していない桁を、10 桁ずつ表示用 Excel のセルへ流す。表示できた桁数
' だけが結果になるので、ここは「実際に書けた分」しか進めない。
Private Sub ShowDigits(ByVal digits As String)
    Dim dec As String
    Dim g As Long
    Dim groups As Long

    On Error GoTo Failed
    If Len(digits) < 2 Then Exit Sub
    If Not m_disp.IsOpen Then Exit Sub
    dec = Mid$(digits, 2)                       ' 先頭の 1 桁は整数部の 3
    groups = Len(dec) \ PI_PERCELL              ' 埋まりきった 10 桁だけ出す
    g = m_shown \ PI_PERCELL
    Do While g < groups
        m_disp.WriteGroup 4 + g \ PI_PERROW, 1 + (g Mod PI_PERROW), _
                          Mid$(dec, g * PI_PERCELL + 1, PI_PERCELL)
        g = g + 1
        m_shown = g * PI_PERCELL
    Loop
    m_screen.PaintDigits dec, m_shown
    Exit Sub
Failed:
    m_log.Note "bench: show err " & Err.Number & " " & Err.Description
End Sub

' 先頭の 1 桁は整数部の 3。数えるのは小数点以下だけ。
Private Function Decimals(ByVal n As Long) As Long
    If n > 0 Then Decimals = n - 1
End Function

'--- 終了 ---------------------------------------------------------------------
' 15 秒で自動終了。BE を止め、表示用 Excel を保存せずに閉じる。Temp の掃除は
' BE の bye を見てから（PbBeSession.PumpIdle）。例外や途中終了、閉じる操作から
' も同じ道を通す。
Public Sub Finish()
    Dim el As Double

    On Error Resume Next
    If Not m_running Then Exit Sub
    m_running = False
    el = ElapsedSince(m_t0)

    m_be.RequestQuit
    m_disp.CloseQuiet
    m_screen.SetText "pb_disp_state", "終了"
    m_ch.KillQuiet m_ch.PiPath
    m_screen.Unflash "pb_lane_in"
    m_screen.Unflash "pb_lane_out"
    m_screen.SetText "pb_be_state", "終了中"
    m_screen.SetText "pb_jobstat", "完了 " & Format$(el, "0.00") & " 秒 ･ 計算 " & _
        Decimals(m_calc) & " 桁 ･ 表示 " & m_shown & " 桁（結果）"
    m_log.Note "bench: finished " & Format$(el, "0.00") & "s calc=" & m_calc & _
        " shown=" & m_shown
End Sub
