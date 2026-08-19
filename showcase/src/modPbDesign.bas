Attribute VB_Name = "modPbDesign"
'==============================================================================
' modPbDesign - 設計の定数と、状態を持たない座標・色の計算
'
' 責務：この 1 枚だけが「画面はどういう寸法か」を知っている。ここに状態は無く、
' 副作用も無い。値を返すだけなので、どのクラスからでも安全に呼べる。
'
' 画面は 812 x 812 の設計ピクセルで描く。1 セルが何設計ピクセルを受け持つかは
' PB_UNIT で決まり、ビルドが 1 / 2 / 4 のどれかへ書き換えて別々の成果物を作る。
' PB_UNIT 以外の寸法は 3 種で共通なので、レイアウトを直すときはここと
' PbScreen だけを見ればよい。
'==============================================================================
Option Explicit

'--- 1 セルが受け持つ設計 px。ビルドがこの 1 行を 1 / 2 / 4 に書き換える。
Public Const PB_UNIT As Long = 4              ' 1 セルが何設計 px 分か（ビルド時に選ぶ: 1 / 2 / 4）

' 設計の格子。1 セルはどのビルドでもこの値以下なので、左端と幅をこの格子へ
' 載せた矩形は 1 / 2 / 4px のどれでも必ずセル境界に一致する。格子から外れた
' 矩形は隣の要素と同じセルを共有し、あとから塗ったほうが相手を潰す。
Public Const PB_GRID As Long = 4

'--- 名乗り
Public Const PB_APP As String = "PIXEL BRIDGE"
Public Const PB_SUB As String = "notepad.exe / UI Automation ･ FE ⇔ BE ･ pi 30s"
Public Const PB_SHEET As String = "PIXELBRIDGE"
Public Const PB_BE_MARK As String = "_be"

'--- 盤面
Public Const PB_W As Long = 812
Public Const PB_H As Long = 812
Public Const PB_DOTS As Long = 8
' 追う窓は 2 つ：この Excel と、つなぐメモ帳 1 枚。
Public Const PB_MAXWIN As Long = 2
Public Const PB_MAXBTN As Long = 8
' 窓の探し直しは何ティックに 1 度か。デスクトップ直下の列挙は他人のプロバイダに
' 触りにいくので、回数そのものを減らしてある。
Public Const PB_SCANEVERY As Long = 4

' ベンチの長さ。秒数と、計器の目盛の本数は同じ 1 つの値から出す（30 秒 = 30 本）。
' 測りはじめるのは、窓の配置と表示用 Excel の用意が終わってから。
Public Const PB_BENCHCOLS As Long = 30
Public Const PB_BENCHSECS As Double = PB_BENCHCOLS

' 30 秒の計器。ベンチのカードでいちばん大きい面をこれが占める。
'
' 見せたいことは 3 つで、それぞれに 1 つずつ図がある。
'
'   帯 TL_COLS 列    1 列が 1 秒。BE が計算に潜っているあいだ FE の 1 秒ポンプが
'                    何回まわったかが、そのまま左から右へ積み上がる。高さはその
'                    1 秒に表示用 Excel まで届いた桁で、1 桁も届かなかった秒でも
'                    根（TL_STUB）は必ず立つ。根が並んで途切れない床になること
'                    自体が「どの秒も FE はティックした」という意味になる。
'   走査点           いま塗っている列の上に出て、1 秒ごとに右へ 1 列動く。
'                    動いていること自体が FE が生きている証拠。
'   下線             経過 / 全体。待機は空、実行で伸び、完了で満ちる。
'
' **列と列のあいだに隙間を置かない。** 隙間を空けて細い棒を 30 本並べた版は、
' 稜線ではなく「粒の集まり」に見えてしまった。隣どうしを接して並べれば、
' 高さの違いは 1 本の帯の稜線として読める。
'
' 列の左端は PB_GRID の格子へ丸める（1 / 2 / 4px のどれでもセル境界に載せる
' ため）。352 / 30 は割り切れないので、幅は 12 が 28 列・8 が 2 列になる。
' 接して並べているので、この差は稜線の中に埋もれて見えない。
Public Const TL_COLS As Long = PB_BENCHCOLS
Public Const TL_X As Long = 432
Public Const TL_Y As Long = 624
Public Const TL_W As Long = 352
Public Const TL_H As Long = 64
Public Const TL_HEADY As Long = 624           ' 走査点の帯
Public Const TL_HEADH As Long = 8
Public Const TL_TOP As Long = 632             ' 帯の上端
Public Const TL_TALL As Long = 36             ' 列の高さ（満杯）
Public Const TL_STUB As Long = 4              ' ティックした証拠の根
Public Const TL_BASEY As Long = 672           ' 下線
Public Const TL_BASEH As Long = 8

' 表示用 Excel の並べ方。1 セルに 10 桁、1 行に 10 セル＝ 100 桁。
' 桁は PI_ROW0 行目から下へ伸びる。上の 3 行は見出しなので凍らせて残す。
Public Const PI_PERCELL As Long = 10
Public Const PI_PERROW As Long = 10
Public Const PI_ROW0 As Long = 4
Public Const PI_FREEZE As Long = 3

'--- 配色（デジタル庁デザインシステム系。テーマは 1 つだけ）
' Excel の Interior.Color は BGR。#0017C1 は &HC11700 と書く。
Public Const C_BAND As Long = &HC11700
Public Const C_KEY As Long = &HC11700
Public Const C_ONBAND As Long = &HFFFFFF
Public Const C_BODY As Long = &HFFFFFF
Public Const C_PANEL As Long = &HF7F7F7
Public Const C_LINE As Long = &HE6E6E6
Public Const C_LINE2 As Long = &HCCCCCC
Public Const C_TEXT As Long = &H1A1A1A
Public Const C_SUB As Long = &H767676
Public Const C_SOFT As Long = &HFEF1E8
Public Const C_OK As Long = &H4B7A19
Public Const C_LAMP As Long = &HA0E069
Public Const C_OFFLAMP As Long = &HB3B3B3
Public Const C_ERR As Long = &H2F2FD3
Public Const C_ERRBG As Long = &HECECFD
Public Const C_WINOTHER As Long = &H999999

' 計器の 2 色。**緑 = FE の時間（1 秒ごとの拍と経過）、青 = BE の成果（届いた
' 桁）。** ぜんぶ同じ色だと、動いている 3 つの図が 1 つの塊に見えて、どれが
' 何を言っているのか読めない。色で担当を分けると、「BE が計算しているあいだも
' FE の緑は途切れず進んでいる」というこのショーケースの主張が、形ではなく色で
' 読めるようになる。
Public Const C_TICK As Long = C_OK
Public Const C_DIGIT As Long = C_KEY

'--- セル番号 → 設計 px
Public Function PxOfCol(ByVal colIdx As Long) As Long
    PxOfCol = (colIdx - 1) * PB_UNIT
End Function

Public Function PxOfRow(ByVal rowIdx As Long) As Long
    PxOfRow = (rowIdx - 1) * PB_UNIT
End Function

' 2 色を混ぜる。**セルは 1 つの色しか持てないので、角丸のアンチエイリアスは
' これで作る。** a は前景の被覆率（0 = 下地のまま、1 = 前景のまま）。
' Interior.Color は BGR の 3 バイトなので、バイトごとに線形に混ぜる。
'
' 角のセルを「塗る / 塗らない」の 2 択にしていた頃は、1px では 1 デバイス px の
' 硬い階段になり、2px では 1 設計 px ずつ描いた 2 本が同じセル行へ写って、本来の
' 引っ込み [5,3,2,1,0] が [1,0,0,0] まで潰れていた。被覆率で混ぜれば、落とせる
' 画素が残らない粒度でも角は丸く見える。
Public Function Blend(ByVal bg As Long, ByVal fg As Long, ByVal a As Double) As Long
    Dim i As Long
    Dim m As Long
    Dim v As Long
    Dim out As Long
    If a <= 0 Then
        Blend = bg
        Exit Function
    End If
    If a >= 1 Then
        Blend = fg
        Exit Function
    End If
    m = 1
    For i = 0 To 2
        v = CLng(((bg \ m) And 255) * (1 - a) + ((fg \ m) And 255) * a)
        If v < 0 Then v = 0
        If v > 255 Then v = 255
        out = out + v * m
        m = m * 256
    Next i
    Blend = out
End Function

' フォントサイズは 0.5pt 刻みへ丸める。Excel はそれ以下を勝手に丸めるので、
' 丸めた値を自分で持っておかないと「設定した値」と「実際の値」がずれる。
Public Function RoundFont(ByVal pt As Double) As Double
    Dim v As Double
    v = Int(pt * 2 + 0.5) / 2
    If v < 1 Then v = 1
    RoundFont = v
End Function

' i 列目の左端（設計 px）。0 から TL_COLS まで（TL_COLS を渡すと右端）。
' PB_GRID の格子へ丸めるので、隣の列と同じセルを取り合わない。
Public Function TlSlotX(ByVal i As Long) As Long
    Dim d As Long
    If i <= 0 Then
        d = 0
    ElseIf i >= TL_COLS Then
        d = TL_W
    Else
        d = ((i * TL_W) \ TL_COLS \ PB_GRID) * PB_GRID
    End If
    TlSlotX = TL_X + d
End Function

' i 列目の幅。隣の列の左端との差なので、列どうしのあいだに隙間はできない。
Public Function TlSlotW(ByVal i As Long) As Long
    TlSlotW = TlSlotX(i + 1) - TlSlotX(i)
End Function

' 長さを PB_GRID の格子へ丸める（近いほう）。
Public Function SnapGrid(ByVal v As Long) As Long
    SnapGrid = ((v + PB_GRID \ 2) \ PB_GRID) * PB_GRID
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

' 書体は「入っていれば使う、無ければ既定へ落とす」。フォント名を Excel へ
' 渡すだけだと、無い書体は黙って別の書体で描かれ、字幅が変わって札が欠ける。
' 入っているかはフォントフォルダを見て決める。
Public Function PickFontJp() As String
    PickFontJp = PickFont("Noto Sans JP", "NotoSansJP*", "Yu Gothic UI")
End Function

Public Function PickFontMono() As String
    PickFontMono = PickFont("Noto Sans Mono", "NotoSansMono*", "Consolas")
End Function

Private Function PickFont(ByVal want As String, ByVal filePat As String, _
                          ByVal fallback As String) As String
    On Error Resume Next
    If Len(Dir$(Environ$("WINDIR") & "\Fonts\" & filePat)) > 0 Then
        PickFont = want
        Exit Function
    End If
    If Len(Dir$(Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\" & filePat)) > 0 Then
        PickFont = want
        Exit Function
    End If
    PickFont = fallback
End Function
