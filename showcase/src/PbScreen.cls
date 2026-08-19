VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbScreen - 812 x 812 の盤面（何をどこへ描くか）
'
' 責務：モックどおりの配置でカードを組み、毎ティックの差分だけを塗り替える。
' 描く道具は PbCanvas が持つ。ここが持つのは座標と文言、そしてボタンの台帳。
'
' 覚えておくこと
'
' - **隣り合う要素が同じセルを共有してはいけない。** 結合セルは 1 セル塗られた
'   だけで全体が塗り潰される。かつて札と枠が同じ行に落ちていて、枠の帯を塗った
'   瞬間に札が灰色（フォーカス時は青）のベタ塗りになり、下の文字ごと消えた。
'   だから座標は 4 設計 px の格子に載せる。
'
' - **ミニマップの窓の矩形は画面の矩形で切り取る。** 窓は画面の外へ出ている
'   ことがあり（端に置いた窓、タスクバーの下へ沈んだ窓）、そのまま描くと
'   ミニマップの枠を越えて隣のカードの上まで塗る。
'
' - **ミニマップの中には結合セルを置かない。** 位置が変わるたびに結合し直す
'   ことになり、その塗り替えの再描画だけでポンプが止まった。
'==============================================================================
Option Explicit

Private m_cv As PbCanvas
Private m_log As PbLog

Private m_btnN As Long
Private m_btnKey(0 To PB_MAXBTN - 1) As String
Private m_btnCell(0 To PB_MAXBTN - 1) As String

Private m_mapSig As String
Private m_dotAt As Long
Private m_dotShown As Long
Private m_frameOn As Long

'--- 30 秒の計器の途中経過
Private m_tlDone As Long          ' 確定させた目盛の番号（-1 = まだ無い）
Private m_tlHead As Long          ' 走査点がいまいる目盛
Private m_tlStub As Long          ' 根だけ立てた目盛
Private m_tlShown As Long         ' 直前の目盛の切れ目で数えた表示済み桁
Private m_tlBase As Long          ' 下線の長さ（設計 px）

Public Sub Init(ByVal canvas As PbCanvas, ByVal logger As PbLog)
    Set m_cv = canvas
    Set m_log = logger
End Sub

'==============================================================================
' 組み立て
'==============================================================================
Public Sub BuildAll()
    Dim prevEv As Boolean
    Dim prevAl As Boolean

    prevEv = Application.EnableEvents
    prevAl = Application.DisplayAlerts
    Application.EnableEvents = False
    ' 文字は結合セルで描くので「左上の値だけ残ります」を何十回も聞かれる。黙らせる。
    Application.DisplayAlerts = False
    On Error GoTo Failed
    m_cv.Hold

    m_btnN = 0
    m_cv.Fill 0, 0, PB_W, PB_H, C_LINE
    m_cv.Fill PB_UNIT, PB_UNIT, PB_W - PB_UNIT * 2, PB_H - PB_UNIT * 2, C_BODY
    BuildHeader
    BuildMiniMapCard
    BuildValueCard
    BuildBenchCard
    BuildFooter
    m_log.Note "build: done"

    Application.DisplayAlerts = prevAl
    Application.EnableEvents = prevEv
    Exit Sub
Failed:
    Application.DisplayAlerts = prevAl
    Application.EnableEvents = prevEv
    m_cv.Unhold
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub BuildHeader()
    m_cv.Fill 2, 2, 808, 52, C_BAND
    m_cv.Txt "", 14, 16, 152, 25, PB_APP, 19, C_ONBAND, True, xlLeft, False
    m_cv.Txt "", 172, 21, 300, 13, PB_SUB, 10, C_ONBAND, False, xlLeft, True
    ' アニメーションは常時 1 秒。速さの切替は置かない。
    m_cv.Txt "pb_uptime", 596, 18, 56, 21, "00:00", 15, C_ONBAND, True, xlLeft, True
    m_cv.Fill 660, 24, 10, 10, C_LAMP
    m_cv.NameIt "pb_lamp", m_cv.Cel(660, 24, 10, 10)
    m_cv.Txt "pb_status", 676, 18, 122, 21, "同期中 poll 0", 11, C_ONBAND, False, xlLeft, True
End Sub

Private Sub BuildMiniMapCard()
    m_cv.CardBox 14, 66, 784, 450
    m_cv.Txt "", 28, 82, 60, 15, "MINIMAP", 11, C_TEXT, True, xlLeft, True
    m_cv.Txt "pb_res", 95, 78, 150, 22, "―", 11, C_TEXT, False, xlLeft, True
    m_cv.Txt "pb_scale", 254, 83, 200, 13, "―", 10, C_SUB, False, xlLeft, True
    ' この 2 つが動かすのは、つないでいるメモ帳 1 枚だけ。Excel は起動時に
    ' 画面の左半分へ置いたきり動かさない。
    RegisterButton "npmax", m_cv.Button("npmax", 560, 76, 108, 28, "最大化", _
        "メモ帳を画面の右半分いっぱいへ移動・リサイズします（OS の最大化ではありません）", 0, 10)
    RegisterButton "npmin", m_cv.Button("npmin", 676, 76, 108, 28, "最小化", _
        "メモ帳をタスクバーへ格納します。最大化か、ミニマップの範囲選択で戻せます", 0, 10)

    ' 要素情報は 2 行のストリップに畳む
    m_cv.Fill 16, 116, 780, 40, C_PANEL
    m_cv.Txt "", 28, 120, 102, 14, "ElementFromPoint", 10, C_KEY, False, xlLeft, True
    m_cv.Txt "pb_pt_win", 138, 120, 540, 14, "―", 10, C_TEXT, False, xlLeft, True
    m_cv.Txt "pb_pt_xy", 680, 120, 104, 14, "PT ―", 10, C_SUB, False, xlRight, True
    m_cv.Txt "", 28, 136, 100, 14, "Rect ･ CanMove", 10, C_SUB, False, xlLeft, True
    m_cv.Txt "pb_pt_rect", 136, 136, 396, 14, "―", 10, C_TEXT, False, xlLeft, True
    m_cv.Txt "pb_drag", 540, 136, 244, 14, "ミニマップの範囲でメモ帳を置く", 10, C_SUB, False, xlRight, False

    m_cv.Fill 16, 156, 780, 358, C_PANEL
    m_cv.NameIt "pb_map", m_cv.Cel(16, 156, 780, 358)
    m_mapSig = ""
End Sub

' つなぐメモ帳は 1 枚だけ。入力面もこの 1 つで、ここへ直接書いた内容がその窓へ
' 流れ、その窓で打った文字がここへ返る。
Private Sub BuildValueCard()
    Dim rg As Range
    m_cv.CardBox 14, 528, 392, 224
    m_cv.Txt "", 28, 544, 92, 16, "ValuePattern", 11, C_KEY, True, xlLeft, True
    m_cv.Txt "pb_focus", 124, 546, 176, 14, "focus ―", 10, C_SUB, False, xlLeft, True
    RegisterButton "sync", m_cv.Button("sync", 304, 540, 88, 28, "同期を停止", _
        "1 秒ごとに本文を取得し、差分があれば SetValue で書き戻します（Ctrl+Shift+M）", 2)

    m_cv.Txt "pb_np_title", 28, 572, 268, 16, "―", 10, C_SUB, False, xlLeft, False
    m_cv.Txt "pb_np_state", 300, 572, 92, 16, "", 10, C_KEY, False, xlRight, False
    ' 入力面の内寸は PB_UNIT に依らない固定値にする。PB_UNIT 倍のインセットだと
    ' 同じ設計でも 1px / 2px だけ最終行が半分描かれて切れる。
    m_cv.Fill 28, 592, 364, 124, C_LINE2
    m_cv.NameIt "pb_np_frame", m_cv.Cel(28, 592, 364, 124)
    m_cv.Fill 28 + PB_UNIT, 592 + PB_UNIT, 364 - PB_UNIT * 2, 124 - PB_UNIT * 2, C_BODY
    Set rg = m_cv.Txt("pb_np_text", 40, 600, 340, 108, "", 11, C_TEXT, False, xlLeft, False)
    With rg
        .ShrinkToFit = False
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = C_BODY
        ' 既定の書式のままだと「1-2」が日付、「007」が 7 に化け、次のティックが
        ' 「セルが編集された」と誤読して、化けた文字列を相手へ書き戻す。
        .NumberFormat = "@"
    End With
    m_frameOn = C_LINE2

    m_cv.Txt "pb_get", 28, 724, 132, 16, "← GetValue --:--:--", 10, C_SUB, False, xlLeft, True
    BuildDots
End Sub

Private Sub BuildDots()
    Dim i As Long
    For i = 0 To PB_DOTS - 1
        m_cv.Fill 166 + i * 29, 726, 27, 8, C_LINE
    Next i
    m_cv.NameIt "pb_dots", m_cv.Cel(166, 726, PB_DOTS * 29 - 2, 8)
    m_dotShown = -1
End Sub

' 円周率ベンチのカード。FE / BE / 表示用 Excel の状態、Temp のレーン、桁の盤、
' 経過と桁数、実行ボタン 1 つ。
Private Sub BuildBenchCard()
    m_cv.CardBox 418, 528, 380, 224
    m_cv.Txt "", 432, 540, 120, 20, "円周率ベンチ", 11, C_KEY, True, xlLeft, True
    m_cv.Txt "", 676, 541, 108, 13, "%TEMP%\pixelbridge\", 9, C_SUB, False, xlRight, True

    m_cv.ProcBox "fe", 432, 564, 116, 32, "FE"
    m_cv.ProcBox "be", 552, 564, 116, 32, "BE"
    m_cv.ProcBox "disp", 672, 564, 112, 32, "表示"

    m_cv.Txt "", 432, 604, 80, 13, "command.json", 9, C_SUB, False, xlLeft, True
    m_cv.Fill 512, 610, 74, 2, C_LINE
    m_cv.NameIt "pb_lane_out", m_cv.Cel(512, 610, 74, 2)
    m_cv.Txt "", 588, 604, 12, 13, "→", 9, C_SUB, False, xlLeft, True
    m_cv.Txt "", 612, 604, 12, 13, "←", 9, C_SUB, False, xlLeft, True
    m_cv.Fill 626, 610, 74, 2, C_LINE
    m_cv.NameIt "pb_lane_in", m_cv.Cel(626, 610, 74, 2)
    m_cv.Txt "", 704, 604, 80, 13, "progress.json", 9, C_SUB, False, xlRight, True

    ' 30 秒の計器。枠と札は行を分ける（結合セルは 1 セル塗られただけで全体が潰れる）。
    BuildTimeline
    m_cv.Txt "pb_jobstat", 432, 692, 352, 16, _
        "待機 ･ 円周率計算を押すと " & CLng(PB_BENCHSECS) & " 秒", 10, C_SUB, False, xlLeft, True

    RegisterButton "runpi", m_cv.Button("runpi", 432, 712, 352, 28, "円周率計算", _
        "非表示の別プロセス Excel が 30 秒だけ円周率を計算し、その桁を表示用 Excel へ流します（Ctrl+Shift+P）", 1, 11)
End Sub

Private Sub BuildFooter()
    m_cv.Fill 2, 764, 808, 46, C_BODY
    m_cv.Fill 2, 764, 808, PB_UNIT, C_LINE
    m_cv.Txt "", 14, 780, 40, 16, "ONKEY", 10, C_SUB, False, xlLeft, True
    ' 札の幅は中身の実寸から取る。狭いと末尾の字が枠の外で切れる。
    m_cv.Chip 56, 776, 124, "^+M 同期 入 / 切"
    m_cv.Chip 184, 776, 116, "^+P 円周率計算"
    m_cv.Txt "pb_meta1", 400, 780, 236, 16, "", 10, C_SUB, False, xlRight, True
    m_cv.Txt "pb_meta2", 644, 780, 154, 16, "", 10, C_SUB, False, xlRight, True
End Sub

'==============================================================================
' ボタンの台帳
'==============================================================================
Private Sub RegisterButton(ByVal key As String, ByVal rg As Range)
    If m_btnN >= PB_MAXBTN Then Exit Sub
    m_btnKey(m_btnN) = key
    m_btnCell(m_btnN) = rg.Cells(1, 1).Address
    m_btnN = m_btnN + 1
End Sub

' 選択されたセルがボタンなら、その識別子。違えば空。
Public Function ButtonKeyAt(ByVal addr As String) As String
    Dim i As Long
    For i = 0 To m_btnN - 1
        If m_btnCell(i) = addr Then
            ButtonKeyAt = m_btnKey(i)
            Exit Function
        End If
    Next i
End Function

Public Function InMap(ByVal c As Range) As Boolean
    Dim mp As Range
    On Error Resume Next
    Set mp = m_cv.RangeOf("pb_map")
    If mp Is Nothing Then Exit Function
    If c.Row >= mp.Row And c.Row < mp.Row + mp.Rows.Count Then
        If c.Column >= mp.Column And c.Column < mp.Column + mp.Columns.Count Then
            InMap = True
        End If
    End If
End Function

'==============================================================================
' 毎ティックの塗り替え
'==============================================================================
Public Sub PaintHeader(ByVal syncOn As Boolean, ByVal startAt As Double, ByVal polls As Long)
    Dim s As Double
    Dim mm As Long

    If syncOn Then
        m_cv.SetFillOf "pb_lamp", C_LAMP
    Else
        m_cv.SetFillOf "pb_lamp", C_OFFLAMP
    End If
    s = ElapsedSince(startAt)
    mm = Int(s) \ 60
    m_cv.SetTxt "pb_uptime", Format$(mm, "00") & ":" & Format$(Int(s) Mod 60, "00")
    If syncOn Then
        m_cv.SetTxt "pb_status", "同期中 poll " & polls
    Else
        m_cv.SetTxt "pb_status", "停止中"
    End If
End Sub

Public Sub PaintFooter()
    m_cv.SetTxt "pb_meta1", "1 セル = " & PB_UNIT & "px ･ 角丸 8px = " & _
        (8 \ PB_UNIT) & " セル（混色）"
    m_cv.SetTxt "pb_meta2", "計器 = " & TL_COLS & " 列 ･ 高さ " & _
        (TL_TALL \ PB_UNIT) & " セル"
End Sub

' ドット列は「同期が生きている」表示。毎ティック 1 つ進む。
Public Sub PaintDots(ByVal syncOn As Boolean)
    Dim rg As Range
    Dim i As Long
    Dim x As Long
    Dim y As Long

    If Not syncOn Then Exit Sub
    m_dotAt = (m_dotAt + 1) Mod PB_DOTS
    If m_dotAt = m_dotShown Then Exit Sub
    On Error Resume Next
    Set rg = m_cv.RangeOf("pb_dots")
    If rg Is Nothing Then Exit Sub
    x = PxOfCol(rg.Column)
    y = PxOfRow(rg.Row)
    For i = 0 To PB_DOTS - 1
        m_cv.Fill x + i * 29, y, 27, 8, IIf(i <= m_dotAt, C_OK, C_LINE)
    Next i
    m_dotShown = m_dotAt
End Sub

' 入力面の枠。フォーカスが相手にあるときだけ色を変える。帯だけ塗り替え、
' 変わっていなければ何もしない。
Public Sub SetFrameFocused(ByVal focused As Boolean)
    Dim c As Long
    Dim rg As Range
    Dim x As Long, y As Long, w As Long, h As Long
    c = IIf(focused, C_KEY, C_LINE2)
    If m_frameOn = c Then Exit Sub
    m_frameOn = c
    On Error Resume Next
    Set rg = m_cv.RangeOf("pb_np_frame")
    If rg Is Nothing Then Exit Sub
    x = PxOfCol(rg.Column): y = PxOfRow(rg.Row)
    w = rg.Columns.Count * PB_UNIT: h = rg.Rows.Count * PB_UNIT
    m_cv.Fill x, y, w, PB_UNIT, c
    m_cv.Fill x, y + h - PB_UNIT, w, PB_UNIT, c
    m_cv.Fill x, y, PB_UNIT, h, c
    m_cv.Fill x + w - PB_UNIT, y, PB_UNIT, h, c
End Sub

Public Sub SetText(ByVal nm As String, ByVal s As String)
    m_cv.SetTxt nm, s
End Sub

Public Sub Note(ByVal s As String)
    m_cv.SetTxt "pb_drag", s
End Sub

Public Sub Flash(ByVal nm As String)
    m_cv.SetFillOf nm, C_KEY
End Sub

Public Sub Unflash(ByVal nm As String)
    m_cv.SetFillOf nm, C_LINE
End Sub

Public Sub InvalidateMap()
    m_mapSig = ""
End Sub

'==============================================================================
' ミニマップ
'==============================================================================
' 盤の中のミニマップの矩形と、画面 → ミニマップの縮尺。
Public Sub MapGeom(ByRef mx As Long, ByRef my As Long, ByRef mw As Long, ByRef mh As Long, _
                   ByRef sc As Double, ByRef ox As Long, ByRef oy As Long, _
                   ByVal scrW As Long, ByVal scrH As Long)
    Dim mp As Range
    sc = 0
    On Error Resume Next
    Set mp = m_cv.RangeOf("pb_map")
    If mp Is Nothing Then Exit Sub
    mx = PxOfCol(mp.Column): my = PxOfRow(mp.Row)
    mw = mp.Columns.Count * PB_UNIT: mh = mp.Rows.Count * PB_UNIT
    If scrW <= 0 Or scrH <= 0 Then Exit Sub
    sc = mw / scrW
    If mh / scrH < sc Then sc = mh / scrH
    ox = mx + (mw - CLng(scrW * sc)) \ 2
    oy = my + (mh - CLng(scrH * sc)) \ 2
End Sub

Public Sub PaintMiniMap(ByVal ws As PbWorkspace)
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    Dim sig As String
    Dim i As Long
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim edge As Long, face As Long, barC As Long
    Dim cx1 As Long, cy1 As Long, cx2 As Long, cy2 As Long

    On Error Resume Next
    If ws.Count = 0 Then Exit Sub
    sig = ws.Signature
    If sig = m_mapSig Then Exit Sub
    m_mapSig = sig

    MapGeom mx, my, mw, mh, sc, ox, oy, ws.ScreenWidth, ws.ScreenHeight
    If sc <= 0 Then Exit Sub

    m_cv.SetTxt "pb_res", ws.ScreenWidth & "×" & ws.ScreenHeight & " 自動判定"
    m_cv.SetTxt "pb_scale", "×" & Format$(sc, "0.00") & " ･ " & _
        CLng(ws.ScreenWidth * sc) & "×" & CLng(ws.ScreenHeight * sc)

    m_cv.Fill mx, my, mw, mh, C_PANEL
    cx1 = ox: cy1 = oy
    cx2 = ox + CLng(ws.ScreenWidth * sc)
    cy2 = oy + CLng(ws.ScreenHeight * sc)
    m_cv.Fill cx1, cy1, cx2 - cx1, cy2 - cy1, C_BODY

    For i = ws.Count - 1 To 0 Step -1
        x1 = ox + CLng((ws.L(i) - ws.ScreenLeft) * sc)
        y1 = oy + CLng((ws.T(i) - ws.ScreenTop) * sc)
        x2 = ox + CLng((ws.R(i) - ws.ScreenLeft) * sc)
        y2 = oy + CLng((ws.B(i) - ws.ScreenTop) * sc)
        ' 画面の外へ出ている分は切り取る（切らないと隣のカードまで塗る）
        If x1 < cx1 Then x1 = cx1
        If y1 < cy1 Then y1 = cy1
        If x2 > cx2 Then x2 = cx2
        If y2 > cy2 Then y2 = cy2
        If x2 - x1 > 20 And y2 - y1 > 20 Then
            ' 操作できる窓（メモ帳）を強調する。Excel は固定なので添え物。
            If ws.IsNotepad(i) Then
                edge = C_KEY: face = C_SOFT: barC = C_KEY
            Else
                edge = C_WINOTHER: face = C_BODY: barC = C_SUB
            End If
            m_cv.Fill x1, y1, x2 - x1, y2 - y1, edge
            m_cv.Fill x1 + PB_UNIT, y1 + PB_UNIT, x2 - x1 - PB_UNIT * 2, y2 - y1 - PB_UNIT * 2, face
            m_cv.Fill x1 + PB_UNIT, y1 + PB_UNIT, x2 - x1 - PB_UNIT * 2, 14, barC
        End If
    Next i
End Sub

' ミニマップのセル → 画面の点。
Public Sub CellToScreen(ByVal c As Range, ByRef px As Long, ByRef py As Long, ByVal ws As PbWorkspace)
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    MapGeom mx, my, mw, mh, sc, ox, oy, ws.ScreenWidth, ws.ScreenHeight
    If sc <= 0 Then Exit Sub
    px = ws.ScreenLeft + CLng((PxOfCol(c.Column) - ox) / sc)
    py = ws.ScreenTop + CLng((PxOfRow(c.Row) - oy) / sc)
End Sub

' ミニマップの選択範囲 → 画面の矩形。
Public Function RangeToScreen(ByVal rg As Range, ByRef x1 As Long, ByRef y1 As Long, _
                              ByRef x2 As Long, ByRef y2 As Long, ByVal ws As PbWorkspace) As Boolean
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    MapGeom mx, my, mw, mh, sc, ox, oy, ws.ScreenWidth, ws.ScreenHeight
    If sc <= 0 Then Exit Function
    x1 = ws.ScreenLeft + CLng((PxOfCol(rg.Column) - ox) / sc)
    y1 = ws.ScreenTop + CLng((PxOfRow(rg.Row) - oy) / sc)
    x2 = ws.ScreenLeft + CLng((PxOfCol(rg.Column + rg.Columns.Count) - ox) / sc)
    y2 = ws.ScreenTop + CLng((PxOfRow(rg.Row + rg.Rows.Count) - oy) / sc)
    RangeToScreen = True
End Function

' ElementFromPoint の実演。その点にある窓の素性を出すだけで、動かす相手を
' 選ぶ操作ではない（相手は常につないだメモ帳 1 枚）。
Public Sub ShowPointInfo(ByVal idx As Long, ByVal ws As PbWorkspace)
    Dim tp As UIAutomationClient.IUIAutomationTransformPattern
    Dim canMove As String
    On Error GoTo Failed
    If idx < 0 Then
        m_cv.SetTxt "pb_pt_win", "その点に窓はありません"
        m_cv.SetTxt "pb_pt_rect", "―"
        Exit Sub
    End If
    m_cv.SetTxt "pb_pt_win", Left$(ws.Label(idx), 30) & " ･ " & _
        IIf(ws.IsNotepad(idx), "Notepad", "XLMAIN") & " ･ pid " & ws.PidOf(idx)
    canMove = "―"
    If Not ws.IsNotepad(idx) Then
        canMove = "固定（起動時に画面の左半分へ配置）"
    ElseIf Not ws.Element Is Nothing Then
        Set tp = ws.Element.GetCurrentPattern(UIA_TransformPatternId)
        If tp Is Nothing Then
            canMove = "False（TransformPattern なし）"
        Else
            canMove = "Move " & CStr(CBool(tp.CurrentCanMove)) & _
                      " / Resize " & CStr(CBool(tp.CurrentCanResize))
        End If
    End If
    m_cv.SetTxt "pb_pt_rect", ws.L(idx) & ", " & ws.T(idx) & " ･ " & _
        (ws.R(idx) - ws.L(idx)) & "×" & (ws.B(idx) - ws.T(idx)) & " ･ " & canMove
    Exit Sub
Failed:
    m_cv.SetTxt "pb_pt_rect", "―"
End Sub

Public Sub ShowPointXY(ByVal px As Long, ByVal py As Long)
    m_cv.SetTxt "pb_pt_xy", "PT " & px & ", " & py
End Sub

'==============================================================================
' 30 秒の計器
'
' ここがこのショーケースの主張そのものを描く。**BE が計算に潜っているあいだも
' FE は 1 秒ごとに目を覚まし、桁を運び、画面を描き替えている。** それを 3 つの
' 図で見せる。
'
'   帯 30 列    1 列が 1 秒。**床（緑）はその秒に FE がティックした証拠で、
'               その上に積む山（青）がその秒に表示用 Excel まで届いた桁。**
'               1 桁も届かなかった秒でも床は立つので、緑は途切れない。列は
'               隙間なく接して並べる（離すと稜線ではなく粒の集まりに見える）。
'   走査点      いま塗っている列の真上に出て、1 秒ごとに右へ 1 列動く（緑）。
'   下線        経過 / 全体（緑）。待機は空、実行で伸び、完了で満ちる。
'
' 色は 2 つだけ。**緑 = FE の時間、青 = BE の成果。** BE が計算に潜っている
' あいだも緑が 1 秒ごとに進み続けることが、このショーケースの主張そのもの。
'
' 前の版はここに円周率の桁の値を色にした 67×12 のランプ盤を置いていた。実測で
' 5,000 桁級が届くので **先頭 804 桁で埋まりきり、開始 2～3 秒で静止していた**。
' 残りの時間はただのモザイクで、動いていないので「FE が生きている」ことも
' 伝わらない。数える対象（届いた桁）は同じまま、見せ方だけを入れ替えてある。
'
' 毎ティックで塗り替えるのは、確定した目盛 1 本・走査点・下線の先端だけ。
' 盤ぜんぶを塗り直さない（1 秒に収めるための約束）。
'==============================================================================
Private Sub BuildTimeline()
    m_cv.Fill TL_X, TL_Y, TL_W, TL_H, C_PANEL
    m_cv.NameIt "pb_tl", m_cv.Cel(TL_X, TL_Y, TL_W, TL_H)
    ResetTimeline
End Sub

' 待機の姿へ戻す。帯は空のまま、走査点も下線も消す。
Public Sub ResetTimeline()
    On Error Resume Next
    m_cv.Fill TL_X, TL_HEADY, TL_W, TL_HEADH, C_PANEL
    m_cv.Fill TL_X, TL_TOP, TL_W, TL_TALL, C_LINE
    m_cv.Fill TL_X, TL_BASEY, TL_W, TL_BASEH, C_LINE2
    m_tlDone = -1
    m_tlHead = -1
    m_tlStub = -1
    m_tlShown = 0
    m_tlBase = 0
End Sub

Public Sub PaintTimeline(ByVal el As Double, ByVal shown As Long, ByVal target As Long)
    Advance el, shown, target, False
End Sub

' 走り終わり。いまの目盛まで確定させ、走査点を消し、下線を結果の色にする。
Public Sub FinishTimeline(ByVal el As Double, ByVal shown As Long, _
                          ByVal target As Long, ByVal ok As Boolean)
    On Error Resume Next
    Advance el, shown, target, True
    m_cv.Fill TL_X, TL_HEADY, TL_W, TL_HEADH, C_PANEL
    m_tlHead = -1
    m_tlBase = 0
    PaintBase el, IIf(ok, C_TICK, C_ERR)
End Sub

Private Sub Advance(ByVal el As Double, ByVal shown As Long, ByVal target As Long, _
                    ByVal closing As Boolean)
    Dim cur As Long
    Dim last As Long
    Dim i As Long
    Dim d As Long
    Dim n As Long

    On Error Resume Next
    cur = Int(el)
    If cur < 0 Then cur = 0
    If cur > TL_COLS - 1 Then cur = TL_COLS - 1
    ' 確定させるのは「終わった秒」まで。閉じるときだけ、いまの秒も確定させる。
    last = cur - 1
    If closing Then last = cur
    n = last - m_tlDone
    If n > 0 Then
        d = shown - m_tlShown
        If d < 0 Then d = 0
        ' ティックが 1 秒より遅れて 2 目盛ぶん進むことがある。その間に届いた
        ' 桁は目盛でならす（遅れた側だけ高くは見せない）。
        For i = m_tlDone + 1 To last
            DrawSlot i, d \ n, target
        Next i
        m_tlDone = last
        m_tlShown = shown
    End If
    If Not closing Then
        If cur > m_tlDone And cur <> m_tlStub Then
            DrawSlot cur, 0, target        ' いまの秒は、まず根だけ
            m_tlStub = cur
        End If
        MoveHead cur
        PaintBase el, C_TICK
    End If
End Sub

' 1 列ぶん。床（緑）は必ず置き、その上に「その 1 秒に届いた桁」の山（青）を
' 積む。満杯の目安はベンチ全体の目標桁数を秒で割った値で、超えたぶんは頭打ち。
' 隣の列と接しているので、床はつながって途切れない 1 本の線になる。
Private Sub DrawSlot(ByVal i As Long, ByVal digits As Long, ByVal target As Long)
    Dim h As Long
    Dim w As Long
    Dim nom As Double

    If i < 0 Or i > TL_COLS - 1 Then Exit Sub
    w = TlSlotW(i)
    m_cv.Fill TlSlotX(i), TL_TOP + TL_TALL - TL_STUB, w, TL_STUB, C_TICK
    If target <= 0 Then Exit Sub
    nom = target / TL_COLS
    If nom <= 0 Then Exit Sub
    h = SnapGrid(CLng((TL_TALL - TL_STUB) * MinD(1, digits / nom)))
    If h > TL_TALL - TL_STUB Then h = TL_TALL - TL_STUB
    If h > 0 Then m_cv.Fill TlSlotX(i), TL_TOP + TL_TALL - TL_STUB - h, w, h, C_DIGIT
End Sub

' 走査点。いま塗っている列の真上に、その列と同じ幅で出る。1 秒ごとに右へ
' 1 列動くので、動いていること自体が FE が生きている証拠になる。
Private Sub MoveHead(ByVal i As Long)
    If i = m_tlHead Then Exit Sub
    If m_tlHead >= 0 Then
        m_cv.Fill TlSlotX(m_tlHead), TL_HEADY, TlSlotW(m_tlHead), TL_HEADH, C_PANEL
    End If
    m_cv.Fill TlSlotX(i), TL_HEADY, TlSlotW(i), TL_HEADH, C_TICK
    m_tlHead = i
End Sub

' 下線は経過そのもの。待機は空、実行で伸び、完了で満ちる。
Private Sub PaintBase(ByVal el As Double, ByVal c As Long)
    Dim w As Long
    w = SnapGrid(CLng(TL_W * MinD(1, el / PB_BENCHSECS)))
    If w > TL_W Then w = TL_W
    If w <= m_tlBase Then Exit Sub
    m_cv.Fill TL_X, TL_BASEY, w, TL_BASEH, c
    m_tlBase = w
End Sub
