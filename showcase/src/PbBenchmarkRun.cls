VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbBenchmarkRun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbBenchmarkRun - 円周率ベンチ 1 回ぶん（15 秒）
'
' 責務：配置と準備、15 秒の計測、桁の受け渡し、そして後始末。指示を送る口
' （Temp の JSON）、桁を映す表示用 Excel、数え方のどれもが「1 回の run」と
' 同じ寿命・同じ状態を共有するので、まとめて持つ。計算そのものは BE 側
' （modPbBackend）で、このクラスは 1 桁も計算しない。
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
'
' このクラスは BE のプロセスの中でも作られる。BE が使うのは Temp のやりとりの
' 部分だけなので、Init は記録係だけでも呼べるようにしてある。
'==============================================================================
Option Explicit

Private m_log As PbLog
Private m_screen As PbScreen
Private m_be As PbBeSession
Private m_ws As PbWorkspace
Private m_canvas As PbCanvas

'--- Temp のやりとり
Private m_seq As Long

'--- 表示用 Excel
Private m_dispApp As Object
Private m_dispBook As Object
Private m_dispSheet As Object

'--- 走っている 1 回
Private m_running As Boolean
Private m_t0 As Double
Private m_calc As Long           ' BE が計算した桁数（先頭の 3 を含む）
Private m_shown As Long          ' 表示用 Excel へ実際に書いた小数以下の桁数
Private m_lastProg As String

Public Sub Init(ByVal logger As PbLog, _
                Optional ByVal screen As PbScreen, _
                Optional ByVal be As PbBeSession, _
                Optional ByVal ws As PbWorkspace, _
                Optional ByVal canvas As PbCanvas)
    Set m_log = logger
    Set m_screen = screen
    Set m_be = be
    Set m_ws = ws
    Set m_canvas = canvas
End Sub

'==============================================================================
' FE と BE のあいだ ─ Temp のファイルだけで往復する
'
' %TEMP%\pixelbridge\ の下だけを触る。置き場所を知っているのはここだけで、
' 他はパスを組み立てない。
'
'   command.json  FE → BE の指示（cmd と seq）
'   progress.json BE → FE の状態
'   fe.json       FE が生きている印。BE はこれが更新されなくなったら自分で終わる
'   pi.txt        計算できた桁そのもの
'   pixelbridge_be.xlsm  BE 用のブックのコピー
'
' 守っている約束
'   - **本文のような自由な文字列を JSON へ入れない。** 引用符も改行もそのまま
'     置ける別ファイルにする。JSON に入れていた頃、日本語の本文が読めずに
'     BE が毎回失敗した。
'   - **ファイルはバイトで読む。** Input$(LOF(f), f) は「LOF バイト分の文字」を
'     要求する読み方で、ASCII のうちは動くが CP932 の日本語では必ず末尾を越えて
'     読み取り全体が落ちる。
'   - **配布物のフォルダには何も書かない。** BE 用のコピーも Temp に置く。
'     置き場所を Temp にできるのは、開くときに AutomationSecurity を落として
'     いるから（信頼された場所かどうかとは別の口）。
'==============================================================================
Public Property Get Folder() As String
    Dim p As String
    p = Environ$("TEMP")
    If Right$(p, 1) <> "\" Then p = p & "\"
    Folder = p & "pixelbridge\"
End Property

Public Property Get CmdPath() As String
    CmdPath = Folder & "command.json"
End Property

Public Property Get ProgPath() As String
    ProgPath = Folder & "progress.json"
End Property

Public Property Get FePath() As String
    FePath = Folder & "fe.json"
End Property

Public Property Get PiPath() As String
    PiPath = Folder & "pi.txt"
End Property

Public Property Get BeBookPath() As String
    BeBookPath = Folder & "pixelbridge_be.xlsm"
End Property

Public Sub EnsureFolder()
    On Error Resume Next
    If Dir$(Folder, vbDirectory) = "" Then MkDir Left$(Folder, Len(Folder) - 1)
End Sub

Public Sub WriteAll(ByVal path As String, ByVal s As String)
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
    ' 相手が同じ瞬間に読んでいると開けないことがある。少し粘る。
    For i = 1 To 20
        Err.Clear
        f = FreeFile
        Open path For Output As #f
        If Err.Number = 0 Then
            Print #f, s
            Close #f
            Exit For
        End If
    Next i
End Sub

Public Function ReadAll(ByVal path As String) As String
    Dim f As Integer
    Dim b() As Byte
    Dim n As Long
    On Error GoTo Failed
    If Dir$(path) = "" Then Exit Function
    f = FreeFile
    Open path For Binary Access Read As #f
    n = LOF(f)
    If n <= 0 Then
        Close #f
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #f, 1, b
    Close #f
    ReadAll = StrConv(b, vbUnicode)
    Exit Function
Failed:
    On Error Resume Next
    Close #f
    ReadAll = ""
End Function

Public Sub KillQuiet(ByVal path As String)
    On Error Resume Next
    If Dir$(path) <> "" Then Kill path
End Sub

' FE が生きている印。BE はこの中身が変わらなくなったら置き去りと判断して
' 自分から終わる。「ファイルが有るか」では足りない（強制終了で残るため）。
Public Sub Heartbeat()
    WriteAll FePath(), "{""fe"":""alive"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
End Sub

' 指示を 1 つ送る。seq は単調増加で、BE は「前より大きい seq」だけを実行する。
' 走行中の BE も seq を見て、新しい指示が来たことに気づいて畳む。
Public Function SendCommand(ByVal cmd As String, ByVal extraJson As String) As Long
    Dim js As String
    m_seq = m_seq + 1
    js = "{""cmd"":""" & cmd & """,""seq"":" & m_seq
    If Len(extraJson) > 0 Then js = js & "," & extraJson
    js = js & ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    WriteAll CmdPath(), js
    SendCommand = m_seq
End Function

Public Function ReadProgress() As String
    ReadProgress = ReadAll(ProgPath())
End Function

' BE が置いた桁の列。WriteAll の Print # が末尾へ改行を足すので落とす。
Public Function ReadDigits() As String
    Dim s As String
    s = ReadAll(PiPath())
    Do While Len(s) > 0 And (Right$(s, 1) = vbLf Or Right$(s, 1) = vbCr)
        s = Left$(s, Len(s) - 1)
    Loop
    ReadDigits = s
End Function

' やりとりに使ったものを全部消し、フォルダも畳む。BE 用のコピーもこの中に
' あるので、先に消さないとフォルダが空にならず RmDir が失敗して残る。
Public Sub CleanAll()
    On Error Resume Next
    KillQuiet PiPath()
    KillQuiet ProgPath()
    KillQuiet CmdPath()
    KillQuiet FePath()
    KillQuiet BeBookPath()
    RmDir Left$(Folder, Len(Folder) - 1)
    If Not m_log Is Nothing Then m_log.Note "channel: temp cleaned"
End Sub

'==============================================================================
' 表示用 Excel（桁を映すだけの一時ブック）
'
' 別プロセスの Excel を 1 つ持ち、保存しないブックへ桁を書く。BE とは別物で、
' 計算はしない。書くのは FE で、BE はこのブックを知らない。
'
' BE との違いが 2 つある。
'
' 1. **AutomationSecurity を触らない。** 開くファイルが無く、マクロを持たない
'    新規ブックだから。
'
' 2. **COM 参照を持ち続けてよい。** BE の参照を手放すのは「忙しい相手を叩くと
'    こちらが固まる」からで、表示用は FE が叩く側。持っていないと毎ティック
'    掴み直すことになり、そちらのほうが高い。
'
' 閉じるときの順序は「ブックを閉じる → Quit → 参照を捨てる」。逆にすると、
' ブックを持たない Excel だけが残ることがある。
'==============================================================================
Public Property Get DisplayIsOpen() As Boolean
    DisplayIsOpen = Not (m_dispSheet Is Nothing)
End Property

Public Function EnsureDisplay() As Boolean
    On Error GoTo Failed
    If Not m_dispApp Is Nothing Then
        ' 生きているか確かめる。落ちていれば作り直す。
        If m_dispApp.Workbooks.Count > 0 Then
            PrepareDisplaySheet
            EnsureDisplay = True
            Exit Function
        End If
    End If
    Set m_dispApp = CreateObject("Excel.Application")
    m_dispApp.DisplayAlerts = False
    m_dispApp.EnableEvents = False
    m_dispApp.UserControl = False
    Set m_dispBook = m_dispApp.Workbooks.Add
    Set m_dispSheet = m_dispBook.Worksheets(1)
    m_dispApp.Visible = True
    PrepareDisplaySheet
    EnsureDisplay = True
    Exit Function
Failed:
    m_log.Note "disp: FAILED " & Err.Number & " " & Err.Description
    CloseDisplay
    EnsureDisplay = False
End Function

Private Sub PrepareDisplaySheet()
    On Error Resume Next
    With m_dispSheet
        .Cells.Clear
        .Cells.Font.Name = m_canvas.FontMono
        .Cells.Font.Size = 11
        .Range("A1:J1").Merge
        .Range("A1").Value = "pi = 3.   BE が計算した桁を FE がここへ流しています"
        .Range("A1").Font.Bold = True
        .Range("A1").Font.Size = 14
        .Range("A2").Value = "小数点以下 " & PI_PERCELL & " 桁ずつ ･ 1 行 " & _
            (PI_PERCELL * PI_PERROW) & " 桁"
        .Range("A2").Font.Size = 10
        .Columns("A:J").ColumnWidth = 12
        .Rows(1).RowHeight = 24
    End With
    ' 桁は文字列で書くので「数値が文字列として保存されています」の緑三角が
    ' 全セルに出る。使い捨ての別プロセスなので、その Excel のチェックだけ切る
    ' （利用者の Excel の設定には触っていない）。
    m_dispApp.ErrorCheckingOptions.NumberAsText = False
    m_dispApp.ActiveWindow.DisplayGridlines = False
End Sub

' 画面の指定の場所へ置く。座標はデバイス px、Excel はポイントで受けるので変換する。
Public Sub PlaceDisplayAt(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
                          ByVal pxPerPt As Double)
    On Error Resume Next
    If m_dispApp Is Nothing Or pxPerPt <= 0 Then Exit Sub
    m_dispApp.WindowState = xlNormal
    m_dispApp.Left = x / pxPerPt
    m_dispApp.Top = y / pxPerPt
    m_dispApp.Width = w / pxPerPt
    m_dispApp.Height = h / pxPerPt
    m_dispApp.Visible = True
End Sub

' 10 桁ぶんを 1 セルへ。先頭に ' を付けて、数値へ化けるのを止める。
Public Sub WriteDisplayGroup(ByVal r As Long, ByVal c As Long, ByVal piece As String)
    On Error Resume Next
    If m_dispSheet Is Nothing Then Exit Sub
    m_dispSheet.Cells(r, c).Value = "'" & piece
End Sub

' 保存せずに閉じる。順序を変えない（上のコメントを参照）。
Public Sub CloseDisplay()
    On Error Resume Next
    If Not m_dispBook Is Nothing Then
        m_dispBook.Saved = True
        m_dispBook.Close SaveChanges:=False
    End If
    If Not m_dispApp Is Nothing Then
        m_dispApp.DisplayAlerts = False
        m_dispApp.Quit
    End If
    Set m_dispSheet = Nothing
    Set m_dispBook = Nothing
    Set m_dispApp = Nothing
End Sub

'==============================================================================
' 15 秒の計測
'==============================================================================
Public Property Get IsRunning() As Boolean
    IsRunning = m_running
End Property

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
    KillQuiet PiPath()
    KillQuiet ProgPath()
    m_be.ForgetProgress

    If Not EnsureDisplay() Then
        m_screen.SetText "pb_jobstat", "表示用 Excel を作れませんでした"
        m_screen.SetText "pb_disp_state", "失敗"
        Exit Sub
    End If
    m_screen.SetText "pb_disp_state", "作成"
    Layout

    ' ここまでが「配置と準備」。15 秒はこの下から測る。
    m_t0 = Timer
    m_running = True
    SendCommand "pi", """secs"":" & CLng(PB_BENCHSECS)
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
    CloseDisplay
    m_screen.SetText "pb_disp_state", "終了"
End Sub

' 開始時の配置。FE は左半分のまま動かさない（起動時に置いたきり）。
' 表示用 Excel を右上 4 分の 1、メモ帳を右下 4 分の 1 へ。三つが重ならずに
' 画面を埋めるので、どれも隠れない。
Private Sub Layout()
    Dim hx As Long, hy As Long, hw As Long, hh As Long

    On Error Resume Next
    hx = m_ws.ScreenLeft + m_ws.ScreenWidth \ 2
    hw = m_ws.ScreenRight - hx
    hy = m_ws.ScreenTop
    hh = m_ws.ScreenHeight \ 2

    PlaceDisplayAt hx, hy, hw, hh, m_canvas.PxPerPt
    If m_ws.IsBound Then
        m_ws.MoveResize hx, hy + hh, hw, m_ws.ScreenHeight - hh
        m_ws.Focus
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

    js = ReadProgress()
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

    ShowDigits ReadDigits()

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
    If Not DisplayIsOpen Then Exit Sub
    dec = Mid$(digits, 2)                       ' 先頭の 1 桁は整数部の 3
    groups = Len(dec) \ PI_PERCELL              ' 埋まりきった 10 桁だけ出す
    g = m_shown \ PI_PERCELL
    Do While g < groups
        WriteDisplayGroup 4 + g \ PI_PERROW, 1 + (g Mod PI_PERROW), _
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
    CloseDisplay
    m_screen.SetText "pb_disp_state", "終了"
    KillQuiet PiPath()
    m_screen.Unflash "pb_lane_in"
    m_screen.Unflash "pb_lane_out"
    m_screen.SetText "pb_be_state", "終了中"
    m_screen.SetText "pb_jobstat", "完了 " & Format$(el, "0.00") & " 秒 ･ 計算 " & _
        Decimals(m_calc) & " 桁 ･ 表示 " & m_shown & " 桁（結果）"
    m_log.Note "bench: finished " & Format$(el, "0.00") & "s calc=" & m_calc & _
        " shown=" & m_shown
End Sub
