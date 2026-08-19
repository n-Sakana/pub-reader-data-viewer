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

'--- 名乗り
Public Const PB_APP As String = "PIXEL BRIDGE"
Public Const PB_SUB As String = "notepad.exe / UI Automation ･ FE ⇔ BE ･ pi 15s"
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

' 桁の盤。67 x 12 = 804 ランプ。円周率の桁を左上から右下へ 1 桁 1 ランプで置き、
' 桁の値（0～9）がそのままランプの色になる。ランプ 1 つの辺は PB_UNIT の倍数
' なので、1px でも 2px でも 4px でもセルに割り切れる。
Public Const LAMP_COLS As Long = 67
Public Const LAMP_ROWS As Long = 12
Public Const LAMP_PX As Long = 4
Public Const LAMP_N As Long = LAMP_COLS * LAMP_ROWS

' ベンチの長さ（秒）。窓の配置と表示用 Excel の用意が終わってから測りはじめる。
Public Const PB_BENCHSECS As Double = 15#

' 表示用 Excel の並べ方。1 セルに 10 桁、1 行に 10 セル＝ 100 桁。
Public Const PI_PERCELL As Long = 10
Public Const PI_PERROW As Long = 10

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

'--- セル番号 → 設計 px
Public Function PxOfCol(ByVal colIdx As Long) As Long
    PxOfCol = (colIdx - 1) * PB_UNIT
End Function

Public Function PxOfRow(ByVal rowIdx As Long) As Long
    PxOfRow = (rowIdx - 1) * PB_UNIT
End Function

' 角丸はコーナーの画素を落として作る。1 セルが 4 設計 px を受け持つ粒度では、
' 角丸に使える画素が残らないのでやめる（モック準拠）。
Public Function RadFor(ByVal r As Long) As Long
    If PB_UNIT >= 4 Then
        RadFor = 0
    Else
        RadFor = r
    End If
End Function

Public Function RadInset(ByVal r As Long, ByVal dy As Long) As Long
    Dim d As Double
    d = r - 0.5 - dy
    RadInset = r - CLng(Sqr(r * r - d * d) + 0.5)
    If RadInset < 0 Then RadInset = 0
    If RadInset > r Then RadInset = r
End Function

' フォントサイズは 0.5pt 刻みへ丸める。Excel はそれ以下を勝手に丸めるので、
' 丸めた値を自分で持っておかないと「設定した値」と「実際の値」がずれる。
Public Function RoundFont(ByVal pt As Double) As Double
    Dim v As Double
    v = Int(pt * 2 + 0.5) / 2
    If v < 1 Then v = 1
    RoundFont = v
End Function

' 桁（0-9）を青の 10 段階へ。モックの色並びをそのまま使う。
Public Function LampColor(ByVal d As Long) As Long
    Select Case d
        Case 0: LampColor = &HFEF3ED
        Case 1: LampColor = &HFCE1D3
        Case 2: LampColor = &HFACDB5
        Case 3: LampColor = &HF8B493
        Case 4: LampColor = &HF6976F
        Case 5: LampColor = &HF3784A
        Case 6: LampColor = &HEC572B
        Case 7: LampColor = &HD43D1B
        Case 8: LampColor = &HA8280F
        Case Else: LampColor = &H731206
    End Select
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
