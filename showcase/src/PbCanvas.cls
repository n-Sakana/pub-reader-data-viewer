VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbCanvas - 疑似ピクセルの画布（シートそのものと、描く道具）
'
' 責務：セルを画素として扱えるようにし、矩形・角丸・文字・ボタンを置く。
' 「何を描くか」は知らない（それは PbScreen）。ここが持つのは「どう描くか」。
'
' 図形もフォームコントロールも ActiveX も使わない。画面はぜんぶセル。
'
' この設計を決めた実測
'
' - **ScreenUpdating の False → True は、変えた量に関係なくビューポート全体を
'   無効化する。** その再描画の値段は見えているセルの色の切り替わり数で決まり、
'   1px の実盤面では 1 回 2.4 秒かかった。逆に True のまま少しだけ書けば汚した
'   矩形しか描き直さない（同じ盤で直書き 1016ms、トグル 1711ms）。
'   だから **毎秒のティックはトグルせず直書きし、Hold / Unhold は初期構築と
'   後始末だけに使う。** ここが軽さの本体。
'
' - **表示設定の切り替えは、列幅を詰める前にやる。** グリッド線と見出しの
'   切り替えはシートの行と列を全部並べ直す。列幅が既定のうちは一瞬だが、
'   812 列を 0.08 まで詰めたあとにやると 276 秒かかった。
'
' - **盤面を持ったまま保存しない。** 66 万セルぶんの寸法を持ったシートは、
'   読み込んだあと前に出す一手だけで Excel を応答不能にする。開くたびに
'   作り直すほうが速い（行高 2ms・列幅 0ms）。
'==============================================================================
Option Explicit

Private m_ws As Worksheet
Private m_log As PbLog
Private m_pitch As Double          ' 1 セルの実寸（ポイント）
Private m_colW As Double
Private m_devPt As Double          ' 1 デバイスピクセルのポイント数（実測）
Private m_pxPerPt As Double        ' 画面ピクセル / ポイント（= 1 / m_devPt）
Private m_held As Boolean
Private m_dirty As Boolean
Private m_repaints As Long
Private m_fontJp As String
Private m_fontMono As String

Public Sub Init(ByVal logger As PbLog)
    Set m_log = logger
    m_fontJp = PickFontJp()
    m_fontMono = PickFontMono()
End Sub

'--- 参照 ---------------------------------------------------------------------
Public Property Get Sheet() As Worksheet
    Set Sheet = m_ws
End Property

Public Property Get Pitch() As Double
    Pitch = m_pitch
End Property

Public Property Get PxPerPt() As Double
    PxPerPt = m_pxPerPt
End Property

' 1 設計 px が何ポイントか。
Public Property Get PxPt() As Double
    PxPt = m_pitch / PB_UNIT
End Property

Public Property Get FontMono() As String
    FontMono = m_fontMono
End Property

Public Property Get Dirty() As Boolean
    Dirty = m_dirty
End Property

Public Property Let Dirty(ByVal v As Boolean)
    m_dirty = v
End Property

Public Property Get Repaints() As Long
    Repaints = m_repaints
End Property

Public Sub CountRepaint()
    m_repaints = m_repaints + 1
End Sub

'--- 描画の抑止 ---------------------------------------------------------------
Public Sub Hold()
    If m_held Then Exit Sub
    Application.ScreenUpdating = False
    m_held = True
End Sub

' Unhold 1 回 =「まとめ描き」1 回。ティック側の直書きは Dirty で数え、
' どちらも Repaints に足す。演出ではなく実測値として画面に出す。
Public Sub Unhold()
    If Not m_held Then Exit Sub
    Application.ScreenUpdating = True
    m_held = False
    m_repaints = m_repaints + 1
    m_dirty = False
End Sub

'--- シートの用意 -------------------------------------------------------------
Public Sub Build()
    Set m_ws = EnsureSheet()
    ' 幅を詰める前に表示設定を済ませる（上のコメントを参照）
    ApplyChrome
    ResetSheet
    MeasureDevPt
    m_pxPerPt = 0
    If m_devPt > 0 Then m_pxPerPt = 1# / m_devPt
End Sub

Private Function EnsureSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PB_SHEET)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If ActiveSheet Is ws Then
            On Error Resume Next
            OtherSheet().Activate
            On Error GoTo 0
        End If
        Application.DisplayAlerts = False
        ws.Delete
        Set ws = Nothing
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = PB_SHEET
    On Error Resume Next
    ThisWorkbook.Worksheets(1).Activate
    On Error GoTo 0
    Set EnsureSheet = ws
End Function

Public Function OtherSheet() As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> PB_SHEET And ws.Visible = xlSheetVisible Then
            Set OtherSheet = ws
            Exit Function
        End If
    Next ws
End Function

Private Sub ResetSheet()
    Dim i As Long
    On Error Resume Next
    Cel(0, 0, PB_W, PB_H).UnMerge
    Cel(0, 0, PB_W, PB_H).ClearContents
    m_ws.Hyperlinks.Delete
    For i = m_ws.Names.Count To 1 Step -1
        m_ws.Names(i).Delete
    Next i
End Sub

' 横スクロールバーとシートタブも消す。1080p の縦は薄く、この 1 段（約 33
' デバイス px）が「812 行が入るか入らないか」を分ける。リボン・数式バー・
' ステータスバーは残す。どれも終了時に戻す。
Public Sub ApplyChrome()
    On Error Resume Next
    m_ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
End Sub

Public Sub RestoreChrome()
    On Error Resume Next
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
End Sub

Public Sub ScrollHome()
    On Error Resume Next
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Public Sub DeleteSheet()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(PB_SHEET).Delete
End Sub

'--- 寸法 ---------------------------------------------------------------------
' 1 デバイスピクセルが何ポイントか。行高は表示 DPI の 1 デバイスピクセルへ
' 丸められるので、丸められずに残る最小の高さがそのまま 1 デバイス px になる。
Private Sub MeasureDevPt()
    Dim t As Double
    m_devPt = 0
    t = 0.05
    Do While t <= 2#
        m_ws.Rows(1).RowHeight = t
        If m_ws.Rows(1).Height > 0 Then
            m_devPt = m_ws.Rows(1).Height
            Exit Do
        End If
        t = t + 0.05
    Loop
    If m_devPt <= 0 Then m_devPt = 0.75
End Sub

' この Excel を画面の左半分へ合わせる。**起動時に一度だけ呼ぶ。以後この窓は
' 位置も大きさも変えない。** メモ帳の最大化・最小化・範囲リサイズでも動かさない。
Public Sub PlaceLeftHalf(ByVal scrL As Long, ByVal scrT As Long, _
                         ByVal scrR As Long, ByVal scrB As Long)
    On Error Resume Next
    If m_pxPerPt <= 0 Then Exit Sub
    If Application.WindowState <> xlNormal Then Application.WindowState = xlNormal
    Application.Left = scrL / m_pxPerPt
    Application.Top = scrT / m_pxPerPt
    Application.Width = ((scrR - scrL) \ 2) / m_pxPerPt
    Application.Height = (scrB - scrT) / m_pxPerPt
    m_log.Note "  place: screen " & (scrR - scrL) & "x" & (scrB - scrT) & _
        " app " & Application.Left & "," & Application.Top & " " & _
        Application.Width & "x" & Application.Height
End Sub

' 1 セルを何デバイス px にするかを決める。
'
' ねらいは「1 設計 px = 0.75pt」。DADS の 1px はこの寸法で、96dpi ならちょうど
' 1 デバイスピクセルになる。ただし **その寸法で盤面がこの窓に収まらないなら
' 意味がない。** 0.75pt どおりだと 1080p / 125% では盤面が 1015～1218 デバイス px
' になり、下 4 割が最初から画面の外だった。だから「窓に入る最大の整数デバイス
' px」を上限にする。
'
' 窓はもう最終の場所（左半分）に置いてあり、見出しもタブも横スクロールバーも
' 消した後なので、UsableHeight / UsableWidth がそのままシートの見える広さになる。
' 端の列が半分だけ見えるのを避けるため「必要数 + 1」まで入る大きさを選び、
' 幅からは縦スクロールバーぶん（UsableWidth はこれを含む）の 24 デバイス px を
' 引いておく。
Public Sub Calibrate()
    Dim k As Long
    Dim best As Long
    Dim bestErr As Double
    Dim e As Double
    Dim cells As Long
    Dim kFit As Long
    Dim kf As Long

    If GridAlreadyOk() Then Exit Sub
    cells = PB_W \ PB_UNIT

    best = 1
    bestErr = 1E+18
    For k = 1 To 16
        e = Abs(k * m_devPt - 0.75 * PB_UNIT)
        If e <= bestErr Then
            bestErr = e
            best = k
        End If
    Next k

    kFit = Int((ActiveWindow.UsableHeight / m_devPt) / (cells + 1))
    kf = Int((ActiveWindow.UsableWidth / m_devPt - 24) / (cells + 1))
    If kf < kFit Then kFit = kf
    If kFit < 1 Then kFit = 1
    If best > kFit Then best = kFit

    m_pitch = best * m_devPt
    m_colW = FindColWidth(m_pitch)
    m_ws.Range(m_ws.Columns(1), m_ws.Columns(cells)).ColumnWidth = m_colW
    m_ws.Range(m_ws.Rows(1), m_ws.Rows(cells)).RowHeight = m_pitch
    m_log.Note "  cal: devPt=" & m_devPt & " pitch=" & m_pitch & " colw=" & m_colW & _
        " kFit=" & kFit & " usable=" & ActiveWindow.UsableWidth & "x" & ActiveWindow.UsableHeight
End Sub

Private Function GridAlreadyOk() As Boolean
    Dim h As Double
    Dim w As Double
    Dim cells As Long
    On Error GoTo Failed
    cells = PB_W \ PB_UNIT
    h = m_ws.Rows(1).Height
    w = m_ws.Columns(1).Width
    If h <= 0 Or w <= 0 Then Exit Function
    If Abs(h - w) > 0.001 Then Exit Function
    If h > 2.5 * PB_UNIT Then Exit Function
    If Abs(m_ws.Cells(cells + 1, 1).Top - cells * h) > 0.5 Then Exit Function
    m_pitch = h
    m_colW = m_ws.Columns(1).ColumnWidth
    GridAlreadyOk = True
    Exit Function
Failed:
    GridAlreadyOk = False
End Function

' 列幅はポイントで指定できない（文字数指定）。ほしい実寸になる値を探す。
Private Function FindColWidth(ByVal wantPt As Double) As Double
    Dim cw As Double
    Dim got As Double
    FindColWidth = 0.08
    cw = 0.01
    Do While cw <= 3#
        m_ws.Columns(1).ColumnWidth = cw
        got = m_ws.Columns(1).Width
        If got >= wantPt - 0.001 Then
            If got <= wantPt + 0.001 Then FindColWidth = cw
            Exit Function
        End If
        cw = cw + 0.01
    Loop
End Function

' 窓は動かさない。入っているかを見て記録し、余白を地の色で塗るだけ。
'
' 見るのは UsableWidth / UsableHeight ではなく VisibleRange。UsableWidth は
' 縦スクロールバーを含み、UsableHeight はシート見出しと横スクロールバーを
' 含むので、そのまま信じると設計の右端と下端が窓の外へ出る。
'
' 余白を塗るのは、1px ビルドだけ盤外に細い黒線が数本出るため。行が 0.6pt まで
' 潰れた盤で、結合セルのある行の帯が塗っていない余白側にはみ出して描かれる。
' セルの値・書式・結合・図形のどれを問い合わせても空のまま返る（この線は
' 単一モジュール時代の版にも同じ位置に出ていた）。塗りのあるセルの上には
' 出ないので、見えている余白を地の色で塗って隠す。盤の外なので設計には触れない。
Public Sub CheckFit()
    Dim needC As Long
    Dim needR As Long
    Dim vis As Range
    Dim lastC As Long
    Dim lastR As Long

    On Error Resume Next
    needC = PB_W \ PB_UNIT
    needR = PB_H \ PB_UNIT
    Set vis = ActiveWindow.VisibleRange
    If vis Is Nothing Then Exit Sub
    m_log.Note "  fit: visible " & vis.Columns.Count & "x" & vis.Rows.Count & _
        " need " & needC & "x" & needR
    lastC = vis.Column + vis.Columns.Count - 1
    lastR = vis.Row + vis.Rows.Count - 1
    If lastC > needC Then
        m_ws.Range(m_ws.Cells(1, needC + 1), m_ws.Cells(lastR, lastC)).Interior.Color = C_BODY
    End If
    If lastR > needR Then
        m_ws.Range(m_ws.Cells(needR + 1, 1), m_ws.Cells(lastR, lastC)).Interior.Color = C_BODY
    End If
End Sub

'--- 描く道具 -----------------------------------------------------------------
' 設計 px の矩形 → セル範囲。
Public Function Cel(ByVal x As Long, ByVal y As Long, _
                    ByVal w As Long, ByVal h As Long) As Range
    Dim c1 As Long
    Dim r1 As Long
    Dim c2 As Long
    Dim r2 As Long
    c1 = x \ PB_UNIT + 1
    r1 = y \ PB_UNIT + 1
    c2 = (x + w - 1) \ PB_UNIT + 1
    r2 = (y + h - 1) \ PB_UNIT + 1
    If c2 < c1 Then c2 = c1
    If r2 < r1 Then r2 = r1
    Set Cel = m_ws.Range(m_ws.Cells(r1, c1), m_ws.Cells(r2, c2))
End Function

Public Sub Fill(ByVal x As Long, ByVal y As Long, ByVal w As Long, _
                ByVal h As Long, ByVal c As Long)
    Dim t0 As Double
    If w <= 0 Or h <= 0 Then Exit Sub
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If x + w > PB_W Then w = PB_W - x
    If y + h > PB_H Then h = PB_H - y
    If w <= 0 Or h <= 0 Then Exit Sub
    m_dirty = True
    t0 = Timer
    Cel(x, y, w, h).Interior.Color = c
    m_log.Span "fill " & x & "," & y & " " & w & "x" & h, t0
End Sub

Public Sub RoundRect(ByVal x As Long, ByVal y As Long, ByVal w As Long, _
                     ByVal h As Long, ByVal r As Long, ByVal c As Long)
    Dim dy As Long
    Dim ins As Long
    r = RadFor(r)
    If r * 2 > h Then r = h \ 2
    If r * 2 > w Then r = w \ 2
    If r <= 0 Then
        Fill x, y, w, h, c
        Exit Sub
    End If
    For dy = 0 To r - 1
        ins = RadInset(r, dy)
        Fill x + ins, y + dy, w - 2 * ins, 1, c
        Fill x + ins, y + h - 1 - dy, w - 2 * ins, 1, c
    Next dy
    Fill x, y + r, w, h - 2 * r, c
End Sub

Public Sub CardBox(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    RoundRect x, y, w, h, 8, C_LINE
    RoundRect x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, h - PB_UNIT * 2, 8 - PB_UNIT, C_BODY
End Sub

' 文字は結合セルで置く。結合セルは溢れた分を隣へ流さずそこで切るので、
' 枠は中身の実寸に合わせる（狭いと末尾が消える）。
Public Function Txt(ByVal nm As String, ByVal x As Long, ByVal y As Long, _
                    ByVal w As Long, ByVal h As Long, ByVal s As String, _
                    ByVal sizePx As Double, ByVal colr As Long, _
                    ByVal bold As Boolean, ByVal alignH As Long, _
                    ByVal mono As Boolean) As Range
    Dim rg As Range
    m_dirty = True
    Set rg = Cel(x, y, w, h)
    rg.Merge
    With rg
        ' 文字列書式を先に敷く。既定のままだと "01:07" が時刻として取り込まれ、
        ' 表示が "1:07" に化ける。しかも読み返した値が書いた値と一致しなくなり、
        ' SetTxt が毎ティック書き直して無駄に描画する。
        .NumberFormat = "@"
        .Value = s
        .Font.Name = IIf(mono, m_fontMono, m_fontJp)
        .Font.Size = RoundFont(sizePx * PxPt)
        .Font.Color = colr
        .Font.bold = bold
        .HorizontalAlignment = alignH
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    ' 文字列書式にした副作用で、"0" のような札に「数値が文字列として保存されて
    ' います」の緑三角が出る。そのセルに限ってチェックを黙らせる（利用者の
    ' Excel 全体の設定には触らない）。Range.Errors は 1 セルにしか無いので、
    ' 結合範囲のままではなく左上のセルへ。
    On Error Resume Next
    rg.Cells(1, 1).Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
    If Len(nm) > 0 Then NameIt nm, rg
    Set Txt = rg
End Function

Public Sub NameIt(ByVal nm As String, ByVal rg As Range)
    On Error Resume Next
    m_ws.Names.Add Name:=nm, RefersTo:=rg
End Sub

' 名前で指す欄の文字を書き換える。中身が同じなら何もしない（無駄な再描画を
' 出さないため）。
Public Sub SetTxt(ByVal nm As String, ByVal s As String)
    Dim rg As Range
    Dim t0 As Double
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If CStr(rg.Cells(1, 1).Value) = s Then Exit Sub
    m_dirty = True
    t0 = Timer
    rg.Cells(1, 1).Value = s
    m_log.Span "settxt " & nm, t0
End Sub

Public Function TextOf(ByVal nm As String) As String
    On Error Resume Next
    TextOf = CStr(m_ws.Range(nm).Cells(1, 1).Value)
End Function

Public Function RangeOf(ByVal nm As String) As Range
    On Error Resume Next
    Set RangeOf = m_ws.Range(nm)
End Function

' 名前の付いた矩形の背景だけ塗り替える。
Public Sub SetFillOf(ByVal nm As String, ByVal c As Long)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If rg.Interior.Color = c Then Exit Sub
    m_dirty = True
    rg.Interior.Color = c
End Sub

' ボタン：セルで描いた矩形と、ホバー用の ScreenTip を持つハイパーリンク。
' 押されたことは「そのセルが選択された」で知る。選択を戻すことはしない
' （1 回 2.4 秒かかるうえ、利用者のカーソルを勝手に動かすことになる）。
' 戻り値は札のセル範囲。押下の判定に使う住所は呼び出し側が控える。
Public Function Button(ByVal key As String, ByVal x As Long, ByVal y As Long, _
                       ByVal w As Long, ByVal h As Long, ByVal caption As String, _
                       ByVal tip As String, ByVal style As Long, _
                       Optional ByVal fs As Double = 11) As Range
    Dim rg As Range
    Dim faceC As Long
    Dim textC As Long
    Dim edgeC As Long

    Select Case style
        Case 1: faceC = C_KEY: textC = C_ONBAND: edgeC = C_KEY        ' 主
        Case 2: faceC = C_ERRBG: textC = C_ERR: edgeC = C_ERR         ' 停止
        Case Else: faceC = C_BODY: textC = C_KEY: edgeC = C_KEY       ' 副
    End Select
    RoundRect x, y, w, h, 6, edgeC
    RoundRect x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, h - PB_UNIT * 2, 6 - PB_UNIT, faceC
    Set rg = Txt("pb_btn_" & key, x + PB_UNIT * 2, y + PB_UNIT, _
                 w - PB_UNIT * 4, h - PB_UNIT * 2, caption, fs, textC, True, xlCenter, False)
    ' 小さいピッチでは名目のポイント数より実描画が広くなり、中央揃えの札は
    ' 両端から欠ける。幅に収まるまで Excel に縮めさせる（ただし結合セルでは
    ' 効かないので、枠の実寸を合わせるほうが本命）。
    rg.ShrinkToFit = True
    rg.Interior.Color = faceC
    On Error Resume Next
    m_ws.Hyperlinks.Add Anchor:=rg, Address:="", _
        SubAddress:="'" & m_ws.Name & "'!" & rg.Cells(1, 1).Address, ScreenTip:=tip
    rg.Font.Color = textC
    rg.Font.Underline = xlUnderlineStyleNone
    rg.Font.bold = True
    On Error GoTo 0
    Set Button = rg
End Function

' 札の中の文字は枠の内寸いっぱいの高さで置く。上下に PB_UNIT*2 ずつ空けると、
' 4px ビルドでは文字の入る高さが 8 デバイス px しか残らず、10px の和文が縦に
' 切れて別の字に見える。
Public Sub Chip(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal s As String)
    Fill x, y, w, 24, C_PANEL
    Fill x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, 24 - PB_UNIT * 2, C_BODY
    Txt "", x + PB_UNIT * 2, y + PB_UNIT, w - PB_UNIT * 4, 24 - PB_UNIT * 2, _
        s, 10, C_TEXT, False, xlCenter, True
End Sub

Public Sub ProcBox(ByVal side As String, ByVal x As Long, ByVal y As Long, _
                   ByVal w As Long, ByVal h As Long, ByVal caption As String)
    RoundRect x, y, w, h, 6, C_LINE
    RoundRect x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, h - PB_UNIT * 2, 6 - PB_UNIT, C_BODY
    Txt "", x + 10, y + 10, 40, 13, caption, 10, C_KEY, True, xlLeft, True
    Txt "pb_" & side & "_state", x + 52, y + 9, w - 62, 14, "待機", 10, C_SUB, False, xlRight, True
End Sub
