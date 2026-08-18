Attribute VB_Name = "modPixelBridge"
'==============================================================================
' VBA Pixel Bridge - Excel分離アーキテクチャ 技術ショーケース
' UI Automation 版（モック「PIXEL BRIDGE モック UI Automation版_ライトテーマ」準拠）
'
' 標準モジュール 1 枚だけ。参照設定は UIAutomationClient のみ。
' Win32 API / Declare / Shell / WScript.Shell / 外部 helper は一切なし。
' 図形（Shape）もフォームコントロールも ActiveX も使わない。画面はぜんぶセル。
'
' この版は UI Automation の実演に専念する。FE/BE 分離は画面から外し、BE の
' 起動もしない（何もしない BE を起こす意味がないため）。
'
' 画面は 812 x 812 の設計ピクセル。1 セルが何設計ピクセルぶんかは PB_UNIT で
' 決める。本番はこの定数を変えて別ファイルを作る。
'
' 実測して決めたこと（この端末での測定。showcase\README.md に記録）
'   - Application.OnTime は秒単位に量子化される。0.5 秒は出ないので 1 秒。
'   - 塗るたびに再描画が走り、その 1 回の値段は「見えているセルの数」で決まる。
'     だから ScreenUpdating は「実際に何か変えるとき」だけ落として戻す。
'     何も変わらないティックでは 1 回も描かない。ここが軽さの本体。
'   - IUIAutomation.ElementFromPoint は POINT を ByVal で取るため VBA から
'     呼べない（コンパイルエラー）。点 → 窓は矩形ヒットテストで出す。
'   - デスクトップ直下の列挙は UIA が Excel 自身にも尋ねるので詰まり得る。
'     100 周 x 3 通りでは詰まらなかったが（平均 118ms）安全の保証にはならない。
'     だから 4 ティックに 1 度だけ、しかも ClassName="Notepad" に絞って呼ぶ。
'==============================================================================
Option Explicit

'------------------------------------------------------------------ identity
Public Const PB_APP As String = "PIXEL BRIDGE"
Public Const PB_SUB As String = "notepad.exe / UI Automation ･ FE ⇔ BE"
Private Const PB_SHEET As String = "PIXELBRIDGE"
Private Const PB_BE_MARK As String = "_be"

'------------------------------------------------------------------ geometry
Private Const PB_UNIT As Long = 4              ' ビルドが 1 / 2 / 4 で上書きする。単体で取り込むとこの値
Private Const PB_W As Long = 812
Private Const PB_H As Long = 812
Private Const PB_DOTS As Long = 8
Private Const PB_MAXNP As Long = 2
Private Const PB_MAXWIN As Long = 3
Private Const PB_SCANEVERY As Long = 4         ' 窓の探し直しは何ティックに 1 度か
Private Const PB_MAXBTN As Long = 8

' 進捗の盤。モックと同じ 67 x 12 = 804 ランプ。原稿を左上から右下へ割り当て、
' 清書できたところまで点いていく。ランプ 1 つは PB_UNIT の倍数にしてあるので、
' 1px でも 2px でも 4px でもセルに割り切れる。
Private Const LAMP_COLS As Long = 67
Private Const LAMP_ROWS As Long = 12
Private Const LAMP_PX As Long = 4              ' ランプ 1 つの設計 px
Private Const LAMP_N As Long = LAMP_COLS * LAMP_ROWS

' 校正の総量（文字の読み直し回数）。原稿が短いと一瞬で終わって FE と BE の
' 差が見えないので、「原稿を何周読み直すか」でここへ揃える。この端末は
' 1 秒あたり約 2000 万回（実測）なので、300,000,000 でおよそ 15 秒になる。
Private Const PB_OPS As Double = 300000000#

'------------------------------------------------------------------ UIA ids
Private Const TS_CHILDREN As Long = 2
Private Const TS_DESCENDANTS As Long = 4
Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_ClassNamePropertyId As Long = 30012
Private Const UIA_NativeWindowHandlePropertyId As Long = 30020
Private Const UIA_IsValuePatternAvailablePropertyId As Long = 30043
Private Const UIA_ValuePatternId As Long = 10002
Private Const UIA_TransformPatternId As Long = 10003
Private Const UIA_WindowPatternId As Long = 10009
Private Const UIA_DocumentControlTypeId As Long = 50030
Private Const UIA_EditControlTypeId As Long = 50004
Private Const WindowVisualState_Normal As Long = 0

'------------------------------------------------------------------ palette
' デジタル庁デザインシステム（DADS β v2 系）。テーマは 1 つだけ。
Private Const C_BAND As Long = &HC11700        ' #0017C1 Blue-900
Private Const C_KEY As Long = &HC11700
Private Const C_ONBAND As Long = &HFFFFFF
Private Const C_BODY As Long = &HFFFFFF
Private Const C_PANEL As Long = &HF7F7F7
Private Const C_LINE As Long = &HE6E6E6
Private Const C_LINE2 As Long = &HCCCCCC
Private Const C_TEXT As Long = &H1A1A1A
Private Const C_SUB As Long = &H767676
Private Const C_SOFT As Long = &HFEF1E8        ' #E8F1FE
Private Const C_OK As Long = &H4B7A19          ' #197A4B
Private Const C_LAMP As Long = &HA0E069        ' #69E0A0
Private Const C_OFFLAMP As Long = &HB3B3B3
Private Const C_ERR As Long = &H2F2FD3         ' #D32F2F
Private Const C_ERRBG As Long = &HECECFD       ' #FDECEC
Private Const C_WINOTHER As Long = &H999999

'------------------------------------------------------------------ state
Private m_ws As Worksheet
Private m_pitch As Double                      ' 1 セルの実寸（ポイント）
Private m_colW As Double
Private m_pxPerPt As Double                    ' 画面ピクセル / ポイント（実測）
Private m_running As Boolean
Private m_held As Boolean                      ' いま描画を止めているか
Private m_startAt As Double
Private m_ticks As Long
Private m_armed As Boolean                     ' 予約が 1 本生きているか
Private m_nextTick As Date                     ' その予約の時刻（取り消しに要る）
Private m_inTick As Boolean                    ' ティックの再入防止
Private m_polls As Long
Private m_syncOn As Boolean
Private m_animRate As Long                     ' 0=なし 1=1秒 5=5秒
Private m_dotAt As Long
Private m_dotShown As Long
Private m_fontJp As String
Private m_fontMono As String
Private m_lastSel As String

Private m_uia As UIAutomationClient.IUIAutomation
Private m_root As UIAutomationClient.IUIAutomationElement
Private m_uiaNote As String
Private m_scrL As Long, m_scrT As Long, m_scrR As Long, m_scrB As Long

' 追える窓：0 = この Excel、1.. = メモ帳
Private m_winN As Long
Private m_winL(0 To PB_MAXWIN - 1) As Long
Private m_winT(0 To PB_MAXWIN - 1) As Long
Private m_winR(0 To PB_MAXWIN - 1) As Long
Private m_winB(0 To PB_MAXWIN - 1) As Long
Private m_winLabel(0 To PB_MAXWIN - 1) As String
Private m_winApp(0 To PB_MAXWIN - 1) As String
Private m_winPid(0 To PB_MAXWIN - 1) As String
Private m_winNp(0 To PB_MAXWIN - 1) As Long    ' -1 = Excel、0.. = メモ帳スロット
Private m_mapSig As String
Private m_picked As Long

' メモ帳スロット
Private m_npWin(0 To PB_MAXNP - 1) As UIAutomationClient.IUIAutomationElement
Private m_npVal(0 To PB_MAXNP - 1) As UIAutomationClient.IUIAutomationValuePattern
Private m_npTitle(0 To PB_MAXNP - 1) As String
Private m_npHwnd(0 To PB_MAXNP - 1) As Variant
Private m_npPid(0 To PB_MAXNP - 1) As String
Private m_npLastText(0 To PB_MAXNP - 1) As String
Private m_npLastCell(0 To PB_MAXNP - 1) As String
Private m_npBound(0 To PB_MAXNP - 1) As Boolean
Private m_frameOn(0 To PB_MAXNP - 1) As Long
Private m_focusHwnd As Variant

' ボタン（セルで描いてあり、押下は選択セルで拾う）
Private m_btnN As Long
Private m_btnKey(0 To PB_MAXBTN - 1) As String
Private m_btnCell(0 To PB_MAXBTN - 1) As String

' FE / BE
Private m_beStarted As Boolean
Private m_beNote As String
Private m_reqSeq As Long
Private m_job As String                        ' "" / "FE" / "BE"
Private m_jobT0 As Double
Private m_lampShown As Long                    ' もう塗ったランプの数
Private m_jobSrc As String                     ' 原稿（メモ帳 A から取った本文）
Private m_jobRounds As Long                    ' 校正の周回数
Private m_jobDone As Long                      ' 清書できた文字数
Private m_cleanShown As String                 ' メモ帳 B へ最後に書いた内容
Private m_srcSlot As Long                      ' 原稿にするスロット（中身で決める）
Private m_dstSlot As Long                      ' 清書にするスロット
Private m_repaints As Long                     ' 実際に描画を戻した回数
Private m_lastProg As String
Private m_beMissing As Long
Private m_beFeSeen As String                   ' 前に見た fe.json の中身
Private m_beAlive As Boolean
Private m_beLastSeq As Long
Private m_beSaidIdle As Boolean
Private m_beQuit As Boolean

'==============================================================================
' 入口
'==============================================================================
' Auto_Open は標準モジュールで動く唯一の起動フック。ユーザーがブックを開けば
' FE が立ち上がる。FE が Workbooks.Open で開いた BE 用コピーでは走らないので、
' FE 側から wb.RunAutoMacros xlAutoOpen で明示的に起動する（要件 v2 §3.1）。
' その BE では同じ Auto_Open が「自分は BE だ」と気づき、OnTime を仕込んで
' すぐ返る。だから FE→BE の呼び出しは即 return になる。
Public Sub Auto_Open()
    If IsBeBook() Then
        PbBeBootstrap
    Else
        PbShow
    End If
End Sub

Public Sub Auto_Close()
    If IsBeBook() Then Exit Sub
    PbShutdown
End Sub

Private Function IsBeBook() As Boolean
    IsBeBook = (InStr(1, ThisWorkbook.Name, PB_BE_MARK & ".", vbTextCompare) > 0)
End Function


Public Function PbPing() As String
    PbPing = PB_APP & " ok " & Application.Version
End Function

' 画面のどこかの文字を名前で読む。外からの検証がこれで実画面の中身を確かめる。
Public Function PbGet(ByVal nm As String) As String
    On Error Resume Next
    PbGet = CStr(m_ws.Range(nm).Cells(1, 1).Value)
End Function

Public Sub PbShow()
    On Error GoTo Failed
    m_startAt = Timer
    m_ticks = 0
    m_polls = 0
    m_syncOn = True
    ' 既定はアニメ 1 秒。ヘッダのボタンもその文字で組み立てるので、ここを
    ' 0（なし）にすると、ボタンは「1秒」と言っているのに時計が --:-- のまま
    ' 止まる、という食い違いになる（実測）。
    m_animRate = 1
    m_picked = 0
    PbLog "show: begin"
    Hold
    PickFonts
    ' BE は画面を作る前に起こす。作ったあとだと SaveCopyAs が 16 万セルの盤面
    ' ごと複製してしまい（94KB が 557KB）、その読み込みで FE が詰まる（実測）。
    PbEnsureBe
    PbMakeSheet
    PbLog "show: sheet pitch=" & m_pitch
    ' 盤面を前に出すのは、組み終わってから一度だけ。組む前に出すと 1px では
    ' 表示設定の変更で再配置が走り続けて戻ってこない（実測：CPU 478 秒）。
    PbFitCoarse
    PbBuildScreen
    PbLog "show: built"
    PbScrollHome
    PbLog "show: sheet shown"
    PbFitRefine
    PbLog "show: fitted"
    Release
    PbBindKeys
    m_running = True
    PbLog "show: shown"
    PbArm
    Exit Sub
Failed:
    Release
    MsgBox "起動に失敗しました。" & vbCrLf & Err.Number & " " & Err.Description, _
           vbExclamation, PB_APP
End Sub

' %TEMP%\pixelbridge.debug があるときだけ、どこまで進んだかを書く。
Private Sub PbLog(ByVal s As String)
    Static mode As Long
    Dim f As Integer
    If mode = 0 Then
        If Len(Dir$(Environ$("TEMP") & "\pixelbridge.debug")) > 0 Then mode = 1 Else mode = 2
    End If
    If mode <> 1 Then Exit Sub
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\pixelbridge-build.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & " " & Format$(Timer, "0.000") & " " & s
    Close #f
End Sub

'==============================================================================
' 描画を止める / 戻す
'
' ここがこの実装のいちばん効いている部分。塗るたびに再描画が走り、その 1 回の
' 値段は見えているセルの数で決まる（812x812 の 1px セルなら 1 回 3.3 秒）。
' だから「実際に何か変えるとき」だけ止めて、そのティックの終わりに戻す。
' 何も変わらないティックでは Hold が一度も呼ばれず、再描画も起きない。
'==============================================================================
Private Sub Hold()
    If m_held Then Exit Sub
    Application.ScreenUpdating = False
    m_held = True
End Sub

' 戻した回数がそのまま「描画回数」。演出ではなく実測値として画面に出す。
Private Sub Release()
    If Not m_held Then Exit Sub
    Application.ScreenUpdating = True
    m_held = False
    m_repaints = m_repaints + 1
End Sub

'==============================================================================
' 疑似ピクセルの寸法
'==============================================================================
Private Sub PbMakeSheet()
    On Error Resume Next
    Application.WindowState = xlMaximized
    On Error GoTo 0
    Set m_ws = EnsureSheet()
    ' 幅を詰める前に表示設定を済ませる（PbSheetChrome の説明を参照）
    PbSheetChrome
    ResetSheet
    PbCalibrate
    MeasurePxPerPt
End Sub

' 疑似ピクセルのシートは開くたびに新しく作る。寸法を持ったまま保存したものを
' 読み込むと、前に出す一手だけで Excel が応答しなくなる（実測）。
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

Private Function OtherSheet() As Worksheet
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

' シートは裏にあり描画も止まっている。その条件なら行高 812 行が 2ms、
' 列幅が 0ms で済む。前に出したままだと 12.1 秒 / 26.7 秒（実測）。
Private Sub PbCalibrate()
    Dim devPt As Double
    Dim t As Double
    Dim k As Long
    Dim best As Long
    Dim bestErr As Double
    Dim e As Double
    Dim cells As Long

    If GridAlreadyOk() Then Exit Sub
    cells = PB_W \ PB_UNIT

    devPt = 0
    t = 0.05
    Do While t <= 2#
        m_ws.Rows(1).RowHeight = t
        If m_ws.Rows(1).Height > 0 Then
            devPt = m_ws.Rows(1).Height
            Exit Do
        End If
        t = t + 0.05
    Loop
    If devPt <= 0 Then devPt = 0.75

    ' ねらいは「1 設計 px = 0.75pt」。DADS の 1px はこの寸法で、96dpi なら
    ' ちょうど 1 デバイスピクセルになる。1 セルは PB_UNIT px 分。
    best = 1
    bestErr = 1E+18
    For k = 1 To 16
        e = Abs(k * devPt - 0.75 * PB_UNIT)
        If e <= bestErr Then
            bestErr = e
            best = k
        End If
    Next k
    m_pitch = best * devPt
    m_colW = FindColWidth(m_pitch)
    m_ws.Range(m_ws.Columns(1), m_ws.Columns(cells)).ColumnWidth = m_colW
    m_ws.Range(m_ws.Rows(1), m_ws.Rows(cells)).RowHeight = m_pitch
    PbLog "  cal: devPt=" & devPt & " pitch=" & m_pitch & " colw=" & m_colW
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

' Excel の座標（ポイント）と UIA の座標（画面ピクセル）の比を実測する。
' 要件 v2 §7 の「単位系を UIA に統一（DPI 問題の回避）」はこれで満たす。
Private Sub MeasurePxPerPt()
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim rc As UIAutomationClient.tagRECT
    m_pxPerPt = 0
    On Error Resume Next
    If Not EnsureUia() Then Exit Sub
    Set el = m_uia.ElementFromHandle(ByVal Application.hwnd)
    If el Is Nothing Then Exit Sub
    rc = el.CurrentBoundingRectangle
    If Application.Width > 0 Then m_pxPerPt = (rc.Right - rc.Left) / Application.Width
End Sub

' 窓合わせは 2 段。
'
' 粗合わせ（PbFitCoarse）は盤面を作った直後、まだ前に出す前にやる。設計の
' 実寸から一息で寄せるだけなので、シートには触らない。
' 仕上げ（PbFitRefine）は盤面を組んで前に出したあと。実際に見えているセルの
' 数を見て、足りなければ広げる。
'
' 分けている理由：仕上げには ActiveWindow が要るが、そのために盤面を先に前へ
' 出すと、1px（812x812 = 659,344 セル）では表示設定の変更で再配置が何度も
' 走って戻ってこない。実測：1px で CPU 478 秒、応答なしのまま。4px では
' 起きないので気づきにくい。盤面を前に出すのは、組み終わってから一度だけ。
Private Sub PbFitCoarse()
    On Error Resume Next
    If Application.WindowState <> xlNormal Then Application.WindowState = xlNormal
    Application.Left = 0
    Application.Top = 0
    Application.Width = PB_W * PxPt() + 40
    Application.Height = PB_H * PxPt() + 220
    PbLog "  fit coarse: app " & Application.Width & "x" & Application.Height
End Sub

Private Sub PbFitRefine()
    Dim i As Long
    Dim needC As Long
    Dim needR As Long
    Dim cellPt As Double
    Dim vis As Range
    Dim dc As Long
    Dim dr As Long

    On Error Resume Next
    needC = PB_W \ PB_UNIT
    needR = PB_H \ PB_UNIT
    cellPt = PxPt() * PB_UNIT

    ' UsableWidth / UsableHeight には合わせない。UsableWidth は縦スクロール
    ' バーの幅を含み、UsableHeight はシート見出しと横スクロールバーの高さを
    ' 含む。そこへ合わせると、その分だけ設計の右端と下端が窓の外へ出る。
    ' 実測では右下の「BE で清書」ボタンが切れ、横スクロールバーが出た。
    ' VisibleRange なら、そのとき本当に見えている範囲そのもの。端の列が半分
    ' だけ見えている状態を数えないよう、必要数より 1 つ多く見えるまで広げる
    ' （modRdv3Ui の CardIsVisible と同じ考え方）。
    For i = 1 To 8
        Set vis = ActiveWindow.VisibleRange
        If vis Is Nothing Then Exit For
        dc = needC + 1 - vis.Columns.Count
        dr = needR + 1 - vis.Rows.Count
        If dc <= 0 And dr <= 0 Then Exit For
        If dc > 0 Then Application.Width = Application.Width + dc * cellPt
        If dr > 0 Then Application.Height = Application.Height + dr * cellPt
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    Next i
    ' 内側のブックウィンドウには触らない。現行の Excel は SDI で、それはアプリ
    ' の窓そのものだから。幅と高さを直接与えると枠が増えて UsableHeight が
    ' 609 から 420.5 に縮み、最大化すると画面全体に広がる（どちらも実測）。
    Set vis = ActiveWindow.VisibleRange
    If Not vis Is Nothing Then
        PbLog "  fit: visible " & vis.Columns.Count & "x" & vis.Rows.Count & _
            " need " & needC & "x" & needR & " passes=" & i
    End If
End Sub

' 表示設定は、列幅を詰める前の軽いシートに対して変える。
'
' グリッド線と見出しの切り替えは、そのシートの列と行を全部並べ直す。列幅が
' 既定のうちは一瞬だが、812 列を 0.08 まで詰めたあと（1px、659,344 セル）に
' やると戻ってこない。実測：盤面を組み終わってからこの 3 行を実行して 276 秒。
' 4px では 41,209 セルなのでどちらでも気づかない。1px で初めて牙をむく。
Private Sub PbSheetChrome()
    On Error Resume Next
    m_ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
End Sub

' 盤面を組んだあとは左上へ戻すだけ。ここは重くない。
Private Sub PbScrollHome()
    On Error Resume Next
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

'==============================================================================
' 描画の材料
'==============================================================================
' 設計は 812x812 の設計 px。実際のセルは PB_UNIT px 分を 1 セルとして受け持つ。
Private Function Cel(ByVal x As Long, ByVal y As Long, _
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

Private Function PxOfCol(ByVal colIdx As Long) As Long
    PxOfCol = (colIdx - 1) * PB_UNIT
End Function

Private Function PxOfRow(ByVal rowIdx As Long) As Long
    PxOfRow = (rowIdx - 1) * PB_UNIT
End Function

Private Function PxPt() As Double
    PxPt = m_pitch / PB_UNIT
End Function

Private Sub Fill(ByVal x As Long, ByVal y As Long, ByVal w As Long, _
                 ByVal h As Long, ByVal c As Long)
    If w <= 0 Or h <= 0 Then Exit Sub
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    If x + w > PB_W Then w = PB_W - x
    If y + h > PB_H Then h = PB_H - y
    If w <= 0 Or h <= 0 Then Exit Sub
    Hold
    Cel(x, y, w, h).Interior.Color = c
End Sub

' 角丸はコーナーの画素を落として作る。4px 粒度では角丸をやめる（モック準拠）。
Private Function RadFor(ByVal r As Long) As Long
    If PB_UNIT >= 4 Then
        RadFor = 0
    Else
        RadFor = r
    End If
End Function

Private Function RadInset(ByVal r As Long, ByVal dy As Long) As Long
    Dim d As Double
    d = r - 0.5 - dy
    RadInset = r - CLng(Sqr(r * r - d * d) + 0.5)
    If RadInset < 0 Then RadInset = 0
    If RadInset > r Then RadInset = r
End Function

Private Sub RoundRect(ByVal x As Long, ByVal y As Long, ByVal w As Long, _
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

Private Sub CardBox(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    RoundRect x, y, w, h, 8, C_LINE
    RoundRect x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, h - PB_UNIT * 2, 8 - PB_UNIT, C_BODY
End Sub

Private Function Txt(ByVal nm As String, ByVal x As Long, ByVal y As Long, _
                     ByVal w As Long, ByVal h As Long, ByVal s As String, _
                     ByVal sizePx As Double, ByVal colr As Long, _
                     ByVal bold As Boolean, ByVal alignH As Long, _
                     ByVal mono As Boolean) As Range
    Dim rg As Range
    Hold
    Set rg = Cel(x, y, w, h)
    rg.Merge
    With rg
        .Value = s
        .Font.Name = IIf(mono, m_fontMono, m_fontJp)
        .Font.Size = RoundFont(sizePx * PxPt())
        .Font.Color = colr
        .Font.bold = bold
        .HorizontalAlignment = alignH
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    If Len(nm) > 0 Then NameIt nm, rg
    Set Txt = rg
End Function

Private Function RoundFont(ByVal pt As Double) As Double
    Dim v As Double
    v = Int(pt * 2 + 0.5) / 2
    If v < 1 Then v = 1
    RoundFont = v
End Function

Private Sub NameIt(ByVal nm As String, ByVal rg As Range)
    On Error Resume Next
    m_ws.Names.Add Name:=nm, RefersTo:=rg
End Sub

Private Sub SetTxt(ByVal nm As String, ByVal s As String)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If CStr(rg.Cells(1, 1).Value) = s Then Exit Sub
    Hold
    rg.Cells(1, 1).Value = s
End Sub

' ボタン：セルで描いた矩形と、ホバー用の ScreenTip を持つハイパーリンク。
' 図形もコントロールも使わないので、押されたことは「そのセルが選択された」で
' 知る。選択を戻すことはしない（この端末では 1 回 2.4 秒かかるうえ、利用者の
' カーソルを勝手に動かすことになる）。
Private Sub Button(ByVal key As String, ByVal x As Long, ByVal y As Long, _
                   ByVal w As Long, ByVal h As Long, ByVal caption As String, _
                   ByVal tip As String, ByVal style As Long, _
                   Optional ByVal fs As Double = 11)
    Dim rg As Range
    Dim faceC As Long
    Dim textC As Long
    Dim edgeC As Long

    Select Case style
        Case 1: faceC = C_KEY: textC = C_ONBAND: edgeC = C_KEY        ' 主
        Case 2: faceC = C_ERRBG: textC = C_ERR: edgeC = C_ERR          ' 停止
        Case Else: faceC = C_BODY: textC = C_KEY: edgeC = C_KEY        ' 副
    End Select
    RoundRect x, y, w, h, 6, edgeC
    RoundRect x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, h - PB_UNIT * 2, 6 - PB_UNIT, faceC
    Set rg = Txt("pb_btn_" & key, x + PB_UNIT * 2, y + PB_UNIT, _
                 w - PB_UNIT * 4, h - PB_UNIT * 2, caption, fs, textC, True, xlCenter, False)
    rg.Interior.Color = faceC
    On Error Resume Next
    m_ws.Hyperlinks.Add Anchor:=rg, Address:="", _
        SubAddress:="'" & m_ws.Name & "'!" & rg.Cells(1, 1).Address, ScreenTip:=tip
    rg.Font.Color = textC
    rg.Font.Underline = xlUnderlineStyleNone
    rg.Font.bold = True
    On Error GoTo 0
    If m_btnN < PB_MAXBTN Then
        m_btnKey(m_btnN) = key
        m_btnCell(m_btnN) = rg.Cells(1, 1).Address
        m_btnN = m_btnN + 1
    End If
End Sub

'==============================================================================
' 画面を組む（モックの実寸そのまま。812x812）
'==============================================================================
Public Sub PbBuildScreen()
    Dim prevEv As Boolean
    Dim prevAl As Boolean

    prevEv = Application.EnableEvents
    prevAl = Application.DisplayAlerts
    Application.EnableEvents = False
    ' 文字は結合セルで描くので「左上の値だけ残ります」を何十回も聞かれる。黙らせる。
    Application.DisplayAlerts = False
    On Error GoTo Failed
    Hold

    m_btnN = 0
    Fill 0, 0, PB_W, PB_H, C_LINE
    Fill PB_UNIT, PB_UNIT, PB_W - PB_UNIT * 2, PB_H - PB_UNIT * 2, C_BODY
    BuildHeader
    BuildMiniMapCard
    BuildValueCard
    BuildInfoCard
    BuildFooter
    PbLog "build: done"

    Application.DisplayAlerts = prevAl
    Application.EnableEvents = prevEv
    Exit Sub
Failed:
    Application.DisplayAlerts = prevAl
    Application.EnableEvents = prevEv
    Release
    Err.Raise Err.Number, , Err.Description
End Sub

'------------------------------------------------------------------ ヘッダー
Private Sub BuildHeader()
    Fill 2, 2, 808, 52, C_BAND
    Txt "", 14, 16, 130, 25, PB_APP, 19, C_ONBAND, True, xlLeft, False
    Txt "", 149, 21, 170, 13, PB_SUB, 10, C_ONBAND, False, xlLeft, True

    ' モックの粒度スイッチの位置に、アニメーション頻度の切替を置く。
    ' 本番で粒度はビルドで決まるので、ここに要るのはこちら。
    Button "anim", 483, 12, 124, 32, "アニメ 1秒", _
        "画面が動く速さを 1秒 → 5秒 → なし の順に切り替えます（Ctrl+Shift+A）。" & _
        "「なし」では、変化がないティックで一度も描きません", 0

    Txt "pb_uptime", 619, 18, 56, 21, "00:00", 15, C_ONBAND, True, xlLeft, True
    Fill 686, 24, 10, 10, C_LAMP
    NameIt "pb_lamp", Cel(686, 24, 10, 10)
    Txt "pb_status", 702, 18, 96, 21, "同期中 poll 0", 11, C_ONBAND, False, xlLeft, True
End Sub

'------------------------------------------------------------------ ミニマップ
Private Sub BuildMiniMapCard()
    CardBox 14, 66, 784, 450
    Txt "", 28, 82, 60, 15, "MINIMAP", 11, C_TEXT, True, xlLeft, True
    ' 解像度は自動判定。モックのセレクタは表示だけにする。
    Txt "pb_res", 95, 78, 150, 22, "―", 11, C_TEXT, False, xlLeft, True
    Txt "pb_scale", 254, 83, 200, 13, "―", 10, C_SUB, False, xlLeft, True
    Button "arrL", 592, 76, 91, 28, "左右 2 分割", _
        "TransformPattern.Move / Resize で左右 2 分割にします", 0
    Button "arrT", 693, 76, 91, 28, "上下 2 分割", _
        "上下 2 分割。最大化中は CanMove=False のため先に通常表示へ戻します", 0

    ' 要素情報はカードをやめ、ミニマップ直下の 2 行ストリップに畳む（v3）
    Fill 16, 116, 780, 40, C_PANEL
    Txt "", 28, 120, 102, 14, "ElementFromPoint", 10, C_KEY, False, xlLeft, True
    Txt "pb_pt_win", 138, 120, 540, 14, "―", 10, C_TEXT, False, xlLeft, True
    Txt "pb_pt_xy", 680, 120, 104, 14, "PT ―", 10, C_SUB, False, xlRight, True
    Txt "", 28, 136, 100, 14, "Rect ･ CanMove", 10, C_SUB, False, xlLeft, True
    Txt "pb_pt_rect", 136, 136, 520, 14, "―", 10, C_TEXT, False, xlLeft, True
    Txt "pb_drag", 664, 136, 120, 14, "ドラッグで移動先を指定", 10, C_SUB, False, xlRight, False

    Fill 16, 156, 780, 358, C_PANEL
    NameIt "pb_map", Cel(16, 156, 780, 358)
    m_mapSig = ""
End Sub

'------------------------------------------------------------------ 同期カード
Private Sub BuildValueCard()
    CardBox 14, 528, 392, 224
    Txt "", 28, 546, 90, 15, "ValuePattern", 11, C_KEY, True, xlLeft, True
    Txt "pb_focus", 120, 547, 177, 13, "focus ―", 10, C_SUB, False, xlLeft, True
    Button "sync", 305, 540, 87, 28, "同期を停止", _
        "1 秒ごとに本文を取得し、差分があれば SetValue で書き戻します（Ctrl+Shift+M）", 2

    BuildSlot 0, 577, 594
    BuildSlot 1, 633, 650

    Txt "pb_get", 28, 723, 130, 13, "← GetValue --:--:--", 10, C_SUB, False, xlLeft, True
    BuildDots
End Sub

Private Sub BuildSlot(ByVal idx As Long, ByVal labelY As Long, ByVal boxY As Long)
    Txt "pb_np" & idx & "_title", 28, labelY, 250, 13, "―", 10, C_SUB, False, xlLeft, False
    Txt "pb_np" & idx & "_state", 280, labelY, 112, 13, "", 10, C_KEY, False, xlRight, False
    ' 入力面は結合セルそのもの。ここに直接書いた内容がメモ帳へ流れる。
    Fill 28, boxY, 364, 46, C_LINE2
    NameIt "pb_np" & idx & "_frame", Cel(28, boxY, 364, 46)
    Fill 28 + PB_UNIT, boxY + PB_UNIT, 364 - PB_UNIT * 2, 46 - PB_UNIT * 2, C_BODY
    With Txt("pb_np" & idx & "_text", 28 + PB_UNIT * 3, boxY + PB_UNIT * 2, _
             364 - PB_UNIT * 6, 46 - PB_UNIT * 4, "", 11, C_TEXT, False, xlLeft, False)
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = C_BODY
    End With
    m_frameOn(idx) = C_LINE2
End Sub

Private Sub BuildDots()
    Dim i As Long
    For i = 0 To PB_DOTS - 1
        Fill 166 + i * 29, 726, 27, 8, C_LINE
    Next i
    NameIt "pb_dots", Cel(166, 726, PB_DOTS * 29 - 2, 8)
    m_dotShown = -1
End Sub

'------------------------------------------------------------------ 要素情報
' 校正・清書カード。FE / BE の箱、Temp のレーン、進捗の盤、実行ボタン 2 つ。
Private Sub BuildInfoCard()
    Dim i As Long
    Dim lx As Long
    Dim ly As Long

    CardBox 418, 528, 380, 224
    Txt "", 432, 540, 92, 16, "校正・清書", 11, C_KEY, True, xlLeft, True
    ' ランプの色見本。清書した文字そのものから色が決まる。
    Txt "", 584, 541, 10, 13, "0", 9, C_SUB, False, xlLeft, True
    For i = 0 To 9
        Fill 593 + i * 6, 544, 6, 8, LampColor(i)
    Next i
    Txt "", 655, 541, 10, 13, "9", 9, C_SUB, False, xlLeft, True
    Txt "", 670, 541, 114, 13, "%TEMP%\pixelbridge\", 9, C_SUB, False, xlRight, True

    ProcBox "fe", 432, 564, 172, 32, "FE"
    ProcBox "be", 612, 564, 172, 32, "BE"

    Txt "", 432, 604, 80, 13, "command.json", 9, C_SUB, False, xlLeft, True
    Fill 512, 610, 74, 2, C_LINE
    NameIt "pb_lane_out", Cel(512, 610, 74, 2)
    Txt "", 588, 604, 12, 13, "→", 9, C_SUB, False, xlLeft, True
    Txt "", 612, 604, 12, 13, "←", 9, C_SUB, False, xlLeft, True
    Fill 626, 610, 74, 2, C_LINE
    NameIt "pb_lane_in", Cel(626, 610, 74, 2)
    Txt "", 704, 604, 80, 13, "progress.json", 9, C_SUB, False, xlRight, True

    ' 進捗の盤。67 x 12 のランプを 432,625 352x77 の枠の中に置く。原稿を
    ' 左上から右下へ割り当て、清書できたところまで点いていく。
    Fill 432, 625, 352, 77, C_PANEL
    lx = 432 + (352 - LAMP_COLS * LAMP_PX) \ 2
    ly = 625 + (77 - LAMP_ROWS * LAMP_PX) \ 2
    Fill lx, ly, LAMP_COLS * LAMP_PX, LAMP_ROWS * LAMP_PX, C_BODY
    NameIt "pb_lamps", Cel(lx, ly, LAMP_COLS * LAMP_PX, LAMP_ROWS * LAMP_PX)
    Txt "pb_jobstat", 432, 686, 352, 14, "待機", 10, C_SUB, False, xlLeft, True

    ' 2 つの札の長さが違うので、幅も 172 / 172 の等分ではなく中身に合わせる。
    ' 等分だと「BE で実行（固まらない）」が内寸（172 - 16 = 156px ≒ 117pt）に
    ' 収まらず、実機で先頭の B が欠けた。この 2 つだけ 10pt にして幅も分ける。
    Button "runfe", 432, 710, 160, 28, "FE で清書（固まる）", _
        "この Excel の中で校正します。終わるまで応答しなくなり、清書は最後にまとめて出ます", 0, 10
    Button "runbe", 600, 710, 184, 28, "BE で清書（固まらない）", _
        "非表示の別プロセス Excel に投げます。FE は動き続け、清書が少しずつ伸びます", 1, 10
End Sub

' 文字を青の 10 段階へ。モックの色並びをそのまま使う。
Private Function LampColor(ByVal d As Long) As Long
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

Private Sub ProcBox(ByVal side As String, ByVal x As Long, ByVal y As Long, _
                    ByVal w As Long, ByVal h As Long, ByVal caption As String)
    RoundRect x, y, w, h, 6, C_LINE
    RoundRect x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, h - PB_UNIT * 2, 6 - PB_UNIT, C_BODY
    Txt "", x + 10, y + 10, 40, 13, caption, 10, C_KEY, True, xlLeft, True
    Txt "pb_" & side & "_state", x + 52, y + 9, w - 62, 14, "待機", 10, C_SUB, False, xlRight, True
End Sub



'------------------------------------------------------------------ フッター
Private Sub BuildFooter()
    Fill 2, 764, 808, 46, C_BODY
    Fill 2, 764, 808, PB_UNIT, C_LINE
    Txt "", 14, 781, 40, 13, "ONKEY", 10, C_SUB, False, xlLeft, True
    Chip 56, 777, 104, "^+M 同期 入 / 切"
    Chip 166, 777, 80, "^+A アニメ"
    Chip 252, 777, 92, "^+F FE で清書"
    Chip 350, 777, 92, "^+B BE で清書"
    Txt "pb_meta1", 420, 781, 240, 13, "", 10, C_SUB, False, xlRight, True
    Txt "pb_meta2", 668, 781, 130, 13, "", 10, C_SUB, False, xlRight, True
End Sub

Private Sub Chip(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal s As String)
    Fill x, y, w, 22, C_PANEL
    Fill x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, 22 - PB_UNIT * 2, C_BODY
    Txt "", x + PB_UNIT * 2, y + PB_UNIT * 2, w - PB_UNIT * 4, 22 - PB_UNIT * 4, _
        s, 10, C_TEXT, False, xlCenter, True
End Sub

'==============================================================================
' ポンプ
'
' Application.OnTime は秒単位に量子化される（実測）。0.5 秒は出ないので 1 秒。
'==============================================================================
Private Function Qual(ByVal procName As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

' ポンプ ── XLToolRack の JobPump と同じ形にしてある。
'
' 守るのは 1 つだけ：「予約は常に 1 本」。m_armed が真なら二重に張らない。
' ティックは入口で m_armed を False に戻し（この予約は消費された）、本体の
' 最後にもう 1 本だけ張る。終わらせたいときは張らないだけでいい。取り消しは
' 保険であって、頼りにする仕組みではない。
'
' 一度これを破って「ダイアログを閉じた直後にもう 1 本張る」を足したところ、
' 予約が同時に 2 本生きる状態ができ、閉じるときに取り消しきれなくなった。
' 残った予約を走らせるため Excel はブックを開いたまま保ち、後始末は終わって
' いるのに窓だけが消えない（実測：150 秒待っても消えず）。増築を撤回して
' 実績のある形へ戻した。
Private Sub PbArm()
    On Error Resume Next
    If m_armed Then Exit Sub
    m_nextTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime m_nextTick, Qual("PbTick")
    m_armed = True
End Sub

Private Sub PbDisarm()
    On Error Resume Next
    If m_armed Then
        Application.OnTime m_nextTick, Qual("PbTick"), , False
    End If
    m_armed = False
End Sub

' 見張り。ブック自身のイベント（選択の変化・シート切り替え・窓の復帰）から
' 呼ばれる。鎖が切れていたら張り直す。切れたかどうかは「予約の時刻を過ぎても
' ティックが来ていないか」で見る。XLToolRack と同じ 5 秒の猶予を置くのは、
' Excel がセル編集中やダイアログ表示中は OnTime を後回しにするからで、その
' 遅れを故障と取り違えないため。
Public Sub PbEnsureArmed()
    On Error Resume Next
    If Not m_running Then Exit Sub
    If m_inTick Then Exit Sub
    If m_armed Then
        If Now <= DateAdd("s", 5, m_nextTick) Then Exit Sub
        PbLog "watchdog: schedule overdue, re-arming"
    Else
        PbLog "watchdog: chain lost, re-arming"
    End If
    m_armed = False
    PbArm
End Sub

Public Sub PbTick()
    On Error GoTo Failed
    If m_inTick Then Exit Sub
    If Not m_running Then Exit Sub
    ' この予約は消費された。次を張るのは本体の最後（と転んだときのハンドラ）。
    m_armed = False
    m_inTick = True
    m_ticks = m_ticks + 1
    PbLog "tick " & m_ticks & ": begin"

    ' 次を仕掛けるのはこの本体の最後（と、転んだときのハンドラ）。先頭で
    ' 仕掛けると「予約が常に 1 つ残っている」状態になり、閉じようとしても
    ' Excel がその予約を待ってしまう。実測：閉じる要求から Auto_Close が
    ' 走り出すまで 35 秒かかった（予約が発火 → まだ m_running なので次を
    ' 仕掛ける、の繰り返し）。最後に仕掛けておけば、終了時に立てた
    ' m_running = False へ最後の 1 発が落ちて、そこで鎖が終わる。
    ' 本体が途中で切れてポンプが死ぬ場合は、PbEnsureArmed（ThisWorkbook の
    ' イベントから呼ばれる見張り）が拾う。

    HandleSelection
    PbLog "  t: sel"
    If m_syncOn Then
        m_polls = m_polls + 1
        If m_winN = 0 Or (m_ticks Mod PB_SCANEVERY) = 1 Then ScanWindows
        PbLog "  t: scan n=" & m_winN
        SyncSlots
        PbLog "  t: sync"
        PaintFocused
        PbLog "  t: focus"
    End If
    PaintMiniMap
    PbLog "  t: map"
    PumpChannel
    PaintHeader
    PaintDots
    PaintFooter
    PbLog "  t: painted held=" & m_held

    Release
    PbLog "tick " & m_ticks & ": end"
    m_inTick = False
    PbArm
    Exit Sub
Failed:
    ' 1 ティックの失敗でポンプを止めない
    Release
    PbLog "tick " & m_ticks & ": err " & Err.Number & " " & Err.Description
    m_inTick = False
    PbArm
End Sub

' 時計もポーリング回数も、毎ティック値が変わる。変われば必ず 1 回再描画が
' 入るので、ここを放っておくと「何も変わらないティック」が一つもなくなる。
' だから数字の更新はアニメの設定に従わせる。「なし」なら数字は動かない。
Private Sub PaintHeader()
    Dim s As Double
    Dim mm As Long

    If m_syncOn Then
        SetLamp C_LAMP
    Else
        SetLamp C_OFFLAMP
    End If
    If m_animRate = 0 Then
        SetTxt "pb_uptime", "--:--"
        SetTxt "pb_status", IIf(m_syncOn, "同期中", "停止中")
        Exit Sub
    End If
    If (m_ticks Mod m_animRate) <> 0 Then Exit Sub
    s = Timer - m_startAt
    If s < 0 Then s = s + 86400
    mm = Int(s) \ 60
    SetTxt "pb_uptime", Format$(mm, "00") & ":" & Format$(Int(s) Mod 60, "00")
    If m_syncOn Then
        SetTxt "pb_status", "同期中 poll " & m_polls
    Else
        SetTxt "pb_status", "停止中"
    End If
End Sub

Private Sub SetLamp(ByVal c As Long)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range("pb_lamp")
    If rg Is Nothing Then Exit Sub
    If rg.Interior.Color = c Then Exit Sub
    Hold
    rg.Interior.Color = c
End Sub

Private Sub PaintFooter()
    SetTxt "pb_meta1", "1 セル = " & PB_UNIT & "px ･ " & _
        IIf(PB_UNIT >= 4, "角丸なし", "角丸 " & (8 \ PB_UNIT) & " セル")
    SetTxt "pb_meta2", "進捗盤 = " & (LAMP_COLS * LAMP_PX \ PB_UNIT) & "×" & _
        (LAMP_ROWS * LAMP_PX \ PB_UNIT) & " セル"
End Sub

' ドット列は「同期が生きている」表示。アニメの設定がそのままここの速さになる。
' なしのときは動かさない＝塗らない＝再描画も起きない。
Private Sub PaintDots()
    Dim rg As Range
    Dim i As Long
    Dim x As Long
    Dim y As Long

    If m_animRate = 0 Then Exit Sub
    If Not m_syncOn Then Exit Sub
    If (m_ticks Mod m_animRate) <> 0 Then Exit Sub
    m_dotAt = (m_dotAt + 1) Mod PB_DOTS
    If m_dotAt = m_dotShown Then Exit Sub
    On Error Resume Next
    Set rg = m_ws.Range("pb_dots")
    If rg Is Nothing Then Exit Sub
    x = PxOfCol(rg.Column)
    y = PxOfRow(rg.Row)
    For i = 0 To PB_DOTS - 1
        Fill x + i * 29, y, 27, 8, IIf(i <= m_dotAt, C_OK, C_LINE)
    Next i
    m_dotShown = m_dotAt
End Sub

'==============================================================================
' 押されたことは「選択セルが変わった」で知る
'
' 1 セルなら「そこを選んだ」、複数セルなら「その矩形へ置け」。ドラッグ選択は
' 離した瞬間に矩形になるので、これがそのままドラッグ配置の入口になる。
' 連続追従はできない（マウスイベントが取れないため）。
'==============================================================================
Private Sub HandleSelection()
    Dim rg As Range
    Dim a As String
    Dim i As Long

    On Error Resume Next
    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> PB_SHEET Then Exit Sub
    Set rg = Selection
    If rg Is Nothing Then Exit Sub
    a = rg.Address
    If a = m_lastSel Then Exit Sub
    m_lastSel = a

    ' ボタンは結合セルなので、押すと選択は「複数セル」になる。だから
    ' セル数で分ける前に、まずボタンかどうかを見る。
    For i = 0 To m_btnN - 1
        If m_btnCell(i) = rg.Cells(1, 1).Address Then
            DoAction m_btnKey(i)
            Exit Sub
        End If
    Next i

    If Not InMap(rg.Cells(1, 1)) Then Exit Sub
    ' ミニマップの中：1 セルなら「そこを選んだ」、範囲なら「そこへ置け」。
    ' ドラッグ選択は離した瞬間に矩形になるので、これが配置の入口になる。
    If rg.Cells.Count = 1 Then
        PickAt rg.Cells(1, 1)
    Else
        PlaceAt rg
    End If
End Sub

Private Function InMap(ByVal c As Range) As Boolean
    Dim mp As Range
    On Error Resume Next
    Set mp = m_ws.Range("pb_map")
    If mp Is Nothing Then Exit Function
    If c.Row >= mp.Row And c.Row < mp.Row + mp.Rows.Count Then
        If c.Column >= mp.Column And c.Column < mp.Column + mp.Columns.Count Then
            InMap = True
        End If
    End If
End Function

Public Sub DoAction(ByVal key As String)
    Select Case key
        Case "sync": ToggleSync
        Case "anim": CycleAnim
        Case "arrL": ArrangeWindows True
        Case "arrT": ArrangeWindows False
        Case "runfe": RunJob "FE"
        Case "runbe": RunJob "BE"
    End Select
End Sub

Private Sub ToggleSync()
    m_syncOn = Not m_syncOn
    SetTxt "pb_btn_sync", IIf(m_syncOn, "同期を停止", "同期を開始")
    PaintHeader
End Sub

Private Sub CycleAnim()
    Select Case m_animRate
        Case 1: m_animRate = 5
        Case 5: m_animRate = 0
        Case Else: m_animRate = 1
    End Select
    SetTxt "pb_btn_anim", "アニメ " & IIf(m_animRate = 0, "なし", CStr(m_animRate) & "秒")
    PaintHeader
End Sub

'==============================================================================
' UI Automation
'==============================================================================
Private Function EnsureUia() As Boolean
    On Error GoTo Failed
    If m_uia Is Nothing Then Set m_uia = New UIAutomationClient.CUIAutomation
    If m_root Is Nothing Then Set m_root = m_uia.GetRootElement
    EnsureUia = (Not m_uia Is Nothing) And (Not m_root Is Nothing)
    Exit Function
Failed:
    m_uiaNote = "UIA 接続不可 " & Err.Number
    EnsureUia = False
End Function

' 追う窓：この Excel と、メモ帳。
'
' メモ帳の探索はデスクトップ直下の列挙だが、ClassName で絞り、しかも
' PB_SCANEVERY ティックに 1 度しか呼ばない。全列挙と条件付き列挙は速さがほぼ
' 同じ（118.1ms 対 118.8ms、実測）なので、得なのは速度ではなく「他人のプロバイダ
' に触る数が 22 から 2 に減る」こと。UIA が Excel 自身にも尋ねる以上、詰まる
' 危険は消えない（このリポジトリの測定記録）ので、回数そのものを減らしてある。
Private Sub ScanWindows()
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim rc As UIAutomationClient.tagRECT
    Dim i As Long
    Dim slot As Long
    Dim h As Variant

    On Error GoTo Failed
    If Not EnsureUia() Then Exit Sub
    rc = m_root.CurrentBoundingRectangle
    m_scrL = rc.Left: m_scrT = rc.Top: m_scrR = rc.Right: m_scrB = rc.Bottom

    m_winN = 0
    Set el = m_uia.ElementFromHandle(ByVal Application.hwnd)
    If Not el Is Nothing Then
        rc = el.CurrentBoundingRectangle
        AddWin rc, "この Excel", "EXCEL.EXE", CStr(el.CurrentProcessId), -1
    End If

    Set cond = m_uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Notepad")
    Set arr = m_root.FindAll(TS_CHILDREN, cond)
    slot = 0
    For i = 0 To arr.Length - 1
        If slot >= PB_MAXNP Then Exit For
        Set el = arr.GetElement(i)
        rc = el.CurrentBoundingRectangle
        ' 最小化した窓は矩形が (-32000,-32000) になるだけで、幅も高さも残る。
        ' それを数えると、画面に出ていない窓が同期スロットを占める。見えて
        ' いない窓は、このデモでは対象にしない。
        If rc.Right > rc.Left And rc.Bottom > rc.Top _
           And Not CBool(el.CurrentIsOffscreen) Then
            h = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
            BindNotepad slot, el, h
            AddWin rc, m_npTitle(slot), "notepad.exe", m_npPid(slot), slot
            slot = slot + 1
        End If
    Next i
    For i = slot To PB_MAXNP - 1
        m_npBound(i) = False
        Set m_npWin(i) = Nothing
        Set m_npVal(i) = Nothing
        m_npTitle(i) = ""
    Next i
    m_uiaNote = "メモ帳 " & slot & " 窓"
    Exit Sub
Failed:
    m_uiaNote = "UIA " & Err.Number
End Sub

Private Sub AddWin(ByRef rc As UIAutomationClient.tagRECT, ByVal label As String, _
                   ByVal app As String, ByVal pid As String, ByVal npSlot As Long)
    If m_winN >= PB_MAXWIN Then Exit Sub
    m_winL(m_winN) = rc.Left
    m_winT(m_winN) = rc.Top
    m_winR(m_winN) = rc.Right
    m_winB(m_winN) = rc.Bottom
    m_winLabel(m_winN) = label
    m_winApp(m_winN) = app
    m_winPid(m_winN) = pid
    m_winNp(m_winN) = npSlot
    m_winN = m_winN + 1
End Sub

Private Sub BindNotepad(ByVal slot As Long, _
                        ByVal win As UIAutomationClient.IUIAutomationElement, _
                        ByVal h As Variant)
    Dim d As UIAutomationClient.IUIAutomationElement
    Dim c1 As UIAutomationClient.IUIAutomationCondition
    Dim c2 As UIAutomationClient.IUIAutomationCondition

    On Error GoTo Failed
    Set m_npWin(slot) = win
    m_npTitle(slot) = win.CurrentName
    m_npPid(slot) = CStr(win.CurrentProcessId)
    If m_npBound(slot) Then
        If CStr(m_npHwnd(slot)) = CStr(h) Then Exit Sub      ' 同じ窓なら結び直さない
    End If
    Set c1 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_DocumentControlTypeId)
    Set c2 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId)
    Set d = win.FindFirst(TS_DESCENDANTS, m_uia.CreateOrCondition(c1, c2))
    If d Is Nothing Then
        m_npBound(slot) = False
        Exit Sub
    End If
    Set m_npVal(slot) = d.GetCurrentPattern(UIA_ValuePatternId)
    If m_npVal(slot) Is Nothing Then
        m_npBound(slot) = False
        Exit Sub
    End If
    m_npHwnd(slot) = h
    m_npBound(slot) = True
    m_npLastText(slot) = ""
    m_npLastCell(slot) = ""
    Exit Sub
Failed:
    m_npBound(slot) = False
End Sub

'==============================================================================
' メモ帳との双方向フリー同期
'
' 入力面が変わっていれば SetValue で書き戻し、変わっていなければ GetValue だけ。
' 窓ごとに独立。判定も 1 文字制限も送信ボタンもない。
'==============================================================================
Private Sub SyncSlots()
    Dim i As Long
    m_focusHwnd = FocusedTopHwnd()
    For i = 0 To PB_MAXNP - 1
        SyncOne i
        If m_npBound(i) And Not IsEmpty(m_focusHwnd) Then
            If CStr(m_npHwnd(i)) = CStr(m_focusHwnd) Then
                FrameColor i, C_KEY
            Else
                FrameColor i, C_LINE2
            End If
        Else
            FrameColor i, C_LINE2
        End If
    Next i
End Sub

Private Sub SyncOne(ByVal slot As Long)
    Dim cur As String
    Dim np As String
    Dim rg As Range

    On Error GoTo Failed
    If Not m_npBound(slot) Then
        SetTxt "pb_np" & slot & "_title", "―"
        SetTxt "pb_np" & slot & "_state", ""
        Exit Sub
    End If
    SetTxt "pb_np" & slot & "_title", Left$(m_npTitle(slot), 34)

    Set rg = m_ws.Range("pb_np" & slot & "_text")
    If rg Is Nothing Then Exit Sub
    cur = CStr(rg.Cells(1, 1).Value)

    If cur <> m_npLastCell(slot) Then
        ' Excel 側が編集された → メモ帳へ書き戻す
        SetTxt "pb_np" & slot & "_state", "SetValue 待ち"
        m_npVal(slot).SetValue ToCrLf(cur)
        m_npLastCell(slot) = cur
        m_npLastText(slot) = ToCrLf(cur)
        SetTxt "pb_np" & slot & "_state", ""
        SetTxt "pb_get", "← SetValue " & Format$(Now, "hh:nn:ss")
        Exit Sub
    End If

    np = CStr(m_npVal(slot).CurrentValue)
    If np <> m_npLastText(slot) Then
        m_npLastText(slot) = np
        m_npLastCell(slot) = ToLf(np)
        Hold
        rg.Cells(1, 1).Value = ToLf(np)
        SetTxt "pb_get", "← GetValue " & Format$(Now, "hh:nn:ss")
    End If
    Exit Sub
Failed:
    m_npBound(slot) = False
    SetTxt "pb_np" & slot & "_state", "切断 " & Err.Number
End Sub

Private Function ToLf(ByVal s As String) As String
    ToLf = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function ToCrLf(ByVal s As String) As String
    ToCrLf = Replace(ToLf(s), vbLf, vbCrLf)
End Function

' 枠の帯だけ塗り替える。変わっていなければ何もしない。
Private Sub FrameColor(ByVal slot As Long, ByVal c As Long)
    Dim rg As Range
    Dim x As Long, y As Long, w As Long, h As Long
    If m_frameOn(slot) = c Then Exit Sub
    m_frameOn(slot) = c
    On Error Resume Next
    Set rg = m_ws.Range("pb_np" & slot & "_frame")
    If rg Is Nothing Then Exit Sub
    x = PxOfCol(rg.Column): y = PxOfRow(rg.Row)
    w = rg.Columns.Count * PB_UNIT: h = rg.Rows.Count * PB_UNIT
    Fill x, y, w, PB_UNIT, c
    Fill x, y + h - PB_UNIT, w, PB_UNIT, c
    Fill x, y, PB_UNIT, h, c
    Fill x + w - PB_UNIT, y, PB_UNIT, h, c
End Sub

' フォーカスを持っている最上位ウィンドウのハンドル。デスクトップは歩かない。
Private Function FocusedTopHwnd() As Variant
    Dim walker As UIAutomationClient.IUIAutomationTreeWalker
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim par As UIAutomationClient.IUIAutomationElement
    Dim depth As Long
    On Error GoTo Failed
    If Not EnsureUia() Then Exit Function
    Set par = m_uia.GetFocusedElement
    If par Is Nothing Then Exit Function
    Set walker = m_uia.ControlViewWalker
    For depth = 1 To 12
        Set el = walker.GetParentElement(par)
        If el Is Nothing Then Exit For
        If m_uia.CompareElements(el, m_root) <> 0 Then Exit For
        Set par = el
    Next depth
    FocusedTopHwnd = par.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
    Exit Function
Failed:
    FocusedTopHwnd = Empty
End Function

' フォーカス中の要素は、ValuePattern カードの見出しに 1 行で畳む（v3）。
Private Sub PaintFocused()
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim vp As UIAutomationClient.IUIAutomationValuePattern
    Dim s As String

    On Error GoTo Failed
    If Not EnsureUia() Then Exit Sub
    Set el = m_uia.GetFocusedElement
    If el Is Nothing Then Exit Sub
    s = "focus pid " & el.CurrentProcessId
    If CBool(el.GetCurrentPropertyValue(UIA_IsValuePatternAvailablePropertyId)) Then
        Set vp = el.GetCurrentPattern(UIA_ValuePatternId)
        If Not vp Is Nothing Then s = s & " ･ Len " & Len(CStr(vp.CurrentValue))
    Else
        s = s & " ･ " & CtrlTypeName(CLng(el.CurrentControlType))
    End If
    SetTxt "pb_focus", s
    Exit Sub
Failed:
    SetTxt "pb_focus", "focus ―"
End Sub

Private Function CtrlTypeName(ByVal id As Long) As String
    Select Case id
        Case 50000: CtrlTypeName = "Button (50000)"
        Case 50004: CtrlTypeName = "Edit (50004)"
        Case 50030: CtrlTypeName = "Document (50030)"
        Case 50032: CtrlTypeName = "Window (50032)"
        Case 50033: CtrlTypeName = "Pane (50033)"
        Case 50020: CtrlTypeName = "Text (50020)"
        Case 50011: CtrlTypeName = "Tab (50011)"
        Case 50019: CtrlTypeName = "TabItem (50019)"
        Case 50026: CtrlTypeName = "Group (50026)"
        Case Else: CtrlTypeName = "id " & id
    End Select
End Function

'==============================================================================
' ミニマップ
'==============================================================================
Private Sub MapGeom(ByRef mx As Long, ByRef my As Long, ByRef mw As Long, ByRef mh As Long, _
                    ByRef sc As Double, ByRef ox As Long, ByRef oy As Long)
    Dim mp As Range
    Dim sw As Long
    Dim sh As Long
    sc = 0
    On Error Resume Next
    Set mp = m_ws.Range("pb_map")
    If mp Is Nothing Then Exit Sub
    mx = PxOfCol(mp.Column): my = PxOfRow(mp.Row)
    mw = mp.Columns.Count * PB_UNIT: mh = mp.Rows.Count * PB_UNIT
    sw = m_scrR - m_scrL: sh = m_scrB - m_scrT
    If sw <= 0 Or sh <= 0 Then Exit Sub
    sc = mw / sw
    If mh / sh < sc Then sc = mh / sh
    ox = mx + (mw - CLng(sw * sc)) \ 2
    oy = my + (mh - CLng(sh * sc)) \ 2
End Sub

Private Sub PaintMiniMap()
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    Dim sig As String
    Dim i As Long
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    Dim edge As Long
    Dim face As Long
    Dim barC As Long

    On Error Resume Next
    If m_winN = 0 Then Exit Sub
    ' 窓の並びが前と同じなら塗り替えない
    For i = 0 To m_winN - 1
        sig = sig & m_winL(i) & "," & m_winT(i) & "," & m_winR(i) & "," & m_winB(i) & ";"
    Next i
    sig = sig & "|" & m_picked & "|" & m_winN
    If sig = m_mapSig Then Exit Sub
    m_mapSig = sig

    MapGeom mx, my, mw, mh, sc, ox, oy
    If sc <= 0 Then Exit Sub

    SetTxt "pb_res", (m_scrR - m_scrL) & "×" & (m_scrB - m_scrT) & " 自動判定"
    SetTxt "pb_scale", "×" & Format$(sc, "0.00") & " ･ " & CLng((m_scrR - m_scrL) * sc) & _
        "×" & CLng((m_scrB - m_scrT) * sc)

    ' ミニマップの中には結合セルを置かない。位置が変わるたびに結合し直す
    ' ことになり、その塗り替えの再描画だけでポンプが止まった（実測）。
    ' 窓の名前は下の対象バーに出す。
    Hold
    Fill mx, my, mw, mh, C_PANEL
    Fill ox, oy, CLng((m_scrR - m_scrL) * sc), CLng((m_scrB - m_scrT) * sc), C_BODY

    For i = m_winN - 1 To 0 Step -1
        x1 = ox + CLng((m_winL(i) - m_scrL) * sc)
        y1 = oy + CLng((m_winT(i) - m_scrT) * sc)
        x2 = ox + CLng((m_winR(i) - m_scrL) * sc)
        y2 = oy + CLng((m_winB(i) - m_scrT) * sc)
        If x2 - x1 > 20 And y2 - y1 > 20 Then
            If i = m_picked Then
                edge = C_KEY: face = C_SOFT: barC = C_KEY
            Else
                edge = C_WINOTHER: face = C_BODY: barC = C_SUB
            End If
            Fill x1, y1, x2 - x1, y2 - y1, edge
            Fill x1 + PB_UNIT, y1 + PB_UNIT, x2 - x1 - PB_UNIT * 2, y2 - y1 - PB_UNIT * 2, face
            Fill x1 + PB_UNIT, y1 + PB_UNIT, x2 - x1 - PB_UNIT * 2, 14, barC
        End If
    Next i
End Sub

' ミニマップのセル → 画面座標を逆算して、その点にある窓を選ぶ
Private Sub PickAt(ByVal c As Range)
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    Dim px As Long, py As Long
    Dim i As Long

    On Error Resume Next
    MapGeom mx, my, mw, mh, sc, ox, oy
    If sc <= 0 Then Exit Sub
    px = m_scrL + CLng((PxOfCol(c.Column) - ox) / sc)
    py = m_scrT + CLng((PxOfRow(c.Row) - oy) / sc)
    SetTxt "pb_pt_xy", "PT " & px & ", " & py

    For i = 0 To m_winN - 1
        If px >= m_winL(i) And px < m_winR(i) And py >= m_winT(i) And py < m_winB(i) Then
            m_picked = i
            m_mapSig = ""                      ' 選択が変わったので描き直す
            FillPointPanel i
            Exit Sub
        End If
    Next i
    SetTxt "pb_pt_win", "その点に窓はありません"
    SetTxt "pb_pt_rect", "―"
End Sub

Private Sub FillPointPanel(ByVal idx As Long)
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim tp As Object
    Dim canMove As String
    On Error GoTo Failed
    SetTxt "pb_pt_win", Left$(m_winLabel(idx), 30) & " ･ " & _
        IIf(m_winNp(idx) < 0, "XLMAIN", "Notepad") & " ･ pid " & m_winPid(idx)
    canMove = "―"
    If m_winNp(idx) < 0 Then
        canMove = "True（Excel 自身）"
    Else
        Set el = m_npWin(m_winNp(idx))
        If Not el Is Nothing Then
            Set tp = el.GetCurrentPattern(UIA_TransformPatternId)
            If tp Is Nothing Then
                canMove = "False（TransformPattern なし）"
            Else
                canMove = CStr(CBool(tp.CurrentCanMove))
            End If
        End If
    End If
    SetTxt "pb_pt_rect", m_winL(idx) & ", " & m_winT(idx) & " ･ " & _
        (m_winR(idx) - m_winL(idx)) & "×" & (m_winB(idx) - m_winT(idx)) & " ･ " & canMove
    Exit Sub
Failed:
    SetTxt "pb_pt_rect", "―"
End Sub

' ドラッグで選ばれた矩形へ、選択中の窓を一発で Move / Resize する
Private Sub PlaceAt(ByVal rg As Range)
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long

    On Error GoTo Failed
    MapGeom mx, my, mw, mh, sc, ox, oy
    If sc <= 0 Then Exit Sub
    x1 = m_scrL + CLng((PxOfCol(rg.Column) - ox) / sc)
    y1 = m_scrT + CLng((PxOfRow(rg.Row) - oy) / sc)
    x2 = m_scrL + CLng((PxOfCol(rg.Column + rg.Columns.Count) - ox) / sc)
    y2 = m_scrT + CLng((PxOfRow(rg.Row + rg.Rows.Count) - oy) / sc)
    If x2 - x1 < 160 Or y2 - y1 < 120 Then
        Note "小さすぎます（160×120 未満）"
        Exit Sub
    End If
    If MoveOne(m_picked, x1, y1, x2 - x1, y2 - y1) Then
        Note "Move / Resize 完了 ･ " & (x2 - x1) & "×" & (y2 - y1)
        m_mapSig = ""
        ScanWindows
    Else
        Note "動かせません（CanMove=False）"
    End If
    Exit Sub
Failed:
    Note "配置に失敗 " & Err.Number
End Sub

Private Sub Note(ByVal s As String)
    SetTxt "pb_drag", s
End Sub

'==============================================================================
' ウィンドウ整列（Win32 なし）
'==============================================================================
Private Sub ArrangeWindows(ByVal leftRight As Boolean)
    Dim halfW As Long, halfH As Long
    Dim nNp As Long
    Dim i As Long
    Dim moved As Long
    Dim step1 As Long

    On Error GoTo Failed
    If m_winN = 0 Then ScanWindows
    nNp = 0
    For i = 0 To PB_MAXNP - 1
        If m_npBound(i) Then nNp = nNp + 1
    Next i
    halfW = (m_scrR - m_scrL) \ 2
    halfH = (m_scrB - m_scrT) \ 2

    If leftRight Then
        If MoveOne(0, m_scrL, m_scrT, halfW, m_scrB - m_scrT) Then moved = moved + 1
        If nNp > 0 Then
            step1 = (m_scrB - m_scrT) \ nNp
            For i = 0 To nNp - 1
                If MoveOne(i + 1, m_scrL + halfW, m_scrT + i * step1, halfW, step1) Then moved = moved + 1
            Next i
        End If
    Else
        If MoveOne(0, m_scrL, m_scrT, m_scrR - m_scrL, halfH) Then moved = moved + 1
        If nNp > 0 Then
            step1 = (m_scrR - m_scrL) \ nNp
            For i = 0 To nNp - 1
                If MoveOne(i + 1, m_scrL + i * step1, m_scrT + halfH, step1, halfH) Then moved = moved + 1
            Next i
        End If
    End If
    Note IIf(leftRight, "左右 2 分割", "上下 2 分割") & "：" & moved & " 窓"
    m_mapSig = ""
    ScanWindows
    Exit Sub
Failed:
    Note "整列に失敗 " & Err.Number
End Sub

Private Function MoveOne(ByVal idx As Long, ByVal x As Long, ByVal y As Long, _
                         ByVal w As Long, ByVal h As Long) As Boolean
    If idx < 0 Or idx >= m_winN Then Exit Function
    If m_winNp(idx) < 0 Then
        MoveOne = MoveExcel(x, y, w, h)
    Else
        MoveOne = MoveWin(m_npWin(m_winNp(idx)), x, y, w, h)
    End If
End Function

' Excel 自身。UIA のピクセルで受け取り、実測した比でポイントへ戻す。
Private Function MoveExcel(ByVal x As Long, ByVal y As Long, _
                           ByVal w As Long, ByVal h As Long) As Boolean
    On Error GoTo Failed
    If m_pxPerPt <= 0 Then MeasurePxPerPt
    If m_pxPerPt <= 0 Then Exit Function
    If Application.WindowState <> xlNormal Then Application.WindowState = xlNormal
    Application.Left = x / m_pxPerPt
    Application.Top = y / m_pxPerPt
    Application.Width = w / m_pxPerPt
    Application.Height = h / m_pxPerPt
    MoveExcel = True
    Exit Function
Failed:
    MoveExcel = False
End Function

' ほかの窓は UIA のパターンで動かす。最大化されていると CanMove が False に
' なるので、先に WindowPattern で通常表示へ戻す。
Private Function MoveWin(ByVal el As UIAutomationClient.IUIAutomationElement, _
                         ByVal x As Long, ByVal y As Long, _
                         ByVal w As Long, ByVal h As Long) As Boolean
    Dim wp As Object
    Dim tp As Object

    On Error GoTo Failed
    If el Is Nothing Then Exit Function
    Set wp = el.GetCurrentPattern(UIA_WindowPatternId)
    If Not wp Is Nothing Then
        On Error Resume Next
        wp.SetWindowVisualState WindowVisualState_Normal
        On Error GoTo Failed
    End If
    Set tp = el.GetCurrentPattern(UIA_TransformPatternId)
    If tp Is Nothing Then Exit Function
    If Not CBool(tp.CurrentCanMove) Then Exit Function
    tp.Move CDbl(x), CDbl(y)
    If CBool(tp.CurrentCanResize) Then tp.Resize CDbl(w), CDbl(h)
    MoveWin = True
    Exit Function
Failed:
    MoveWin = False
End Function

'==============================================================================
' ショートカット
'==============================================================================
Private Sub PbBindKeys()
    On Error Resume Next
    Application.OnKey "^+M", Qual("PbKeySync")
    Application.OnKey "^+A", Qual("PbKeyAnim")
    Application.OnKey "^+F", Qual("PbKeyRunFe")
    Application.OnKey "^+B", Qual("PbKeyRunBe")
End Sub

Private Sub PbUnbindKeys()
    On Error Resume Next
    Application.OnKey "^+M"
    Application.OnKey "^+A"
    Application.OnKey "^+F"
    Application.OnKey "^+B"
End Sub

Public Sub PbKeySync()
    DoAction "sync"
    Release
End Sub

Public Sub PbKeyAnim()
    DoAction "anim"
    Release
End Sub

' 実行はボタンだけでなくキーからも。ボタンはセルなので、押すには盤面が
' 前に出ていないといけない。キーなら他の窓を見ていても始められる。
Public Sub PbKeyRunFe()
    DoAction "runfe"
    Release
End Sub

Public Sub PbKeyRunBe()
    DoAction "runbe"
    Release
End Sub

'==============================================================================
' 後始末
'==============================================================================
' 閉じる要求が来たその場で鎖を畳む。ThisWorkbook の BeforeClose から呼ばれる
' （ビルド時に注入）。Auto_Close まで待つと、その間に予約が発火して「まだ
' 動いている」ティックが次を仕掛け直し、Excel は閉じられないまま待たされる。
' 実測：閉じる要求から Auto_Close が走り出すまで 35 秒かかった。
' 実証済み実装の Rdv3AppPrepareClose と同じ役どころ。
Public Sub PbPrepareClose()
    Dim k As Long
    On Error Resume Next
    m_running = False
    ' 「保存しますか」を出させない。この盤面は保存しない前提で毎回作り直す
    ' ものなので、変更として数えられる意味がない。実測：ここを立てずに閉じる
    ' と Excel 自身の確認ダイアログ（クラス bosa_sdm_XL9）が出て、そこで
    ' 閉じる処理が止まり、90 秒待ってもブックが消えなかった。
    ThisWorkbook.Saved = True
    ' 予約した時刻そのもので取り消す。ここを Now + 1 秒 で呼んでいたときは
    ' 一度も当たらず、予約が残ったせいでブックが閉じなかった。
    PbDisarm
    PbLog "quit: pump disarmed"
    PbUnbindKeys
End Sub

Public Sub PbShutdown()
    Dim k As Long
    On Error Resume Next
    PbLog "quit: begin"
    PbPrepareClose

    ' BE に終われと言って、返事を待つ。相手のプロセスは殺さない。
    If m_beStarted Then
        m_reqSeq = m_reqSeq + 1
        WriteAll CmdPath(), "{""cmd"":""quit"",""seq"":" & m_reqSeq & _
            ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
        For k = 1 To 8
            If InStr(1, ReadAll(ProgPath()), """bye""") > 0 Then Exit For
            Application.Wait Now + TimeSerial(0, 0, 1)
        Next k
        PbLog "quit: be waited " & k
    End If
    KillQuiet CmdPath()
    KillQuiet SrcPath()
    KillQuiet CleanPath()
    KillQuiet ProgPath()
    KillQuiet FePath()
    RmDirQuiet Left$(TempDir(), Len(TempDir()) - 1)
    KillQuiet BeBookPath()

    ' 1px セルの盤面を前に出したまま表示設定を戻すと再配置が何度も走る。
    ' 先にふつうのシートへ移る。
    If Not OtherSheet() Is Nothing Then OtherSheet().Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(PB_SHEET).Delete
    Release
    ThisWorkbook.Saved = True
    ' UI Automation で掴んだ相手（メモ帳の窓と ValuePattern）を手放す。
    ' これはプロセスをまたぐ COM 参照なので、握ったままだと Excel は終われない。
    ' 実測：後始末は 2 秒で終わっているのに、プロセスだけが 45 秒たっても残った。
    For k = 0 To PB_MAXNP - 1
        m_npBound(k) = False
        Set m_npVal(k) = Nothing
        Set m_npWin(k) = Nothing
    Next k
    Set m_root = Nothing
    Set m_uia = Nothing
    PbLog "quit: done"
End Sub

Private Sub RmDirQuiet(ByVal path As String)
    On Error Resume Next
    RmDir path
End Sub

'==============================================================================
' 重い処理そのもの ── 原稿の校正と清書
'
' メモ帳 A（原稿）を 1 文字ずつメモ帳 B（清書）へ写す。1 文字書くたびに、
' そこまで書いた分を頭から読み直して原稿と突き合わせ、照合コードを取り直す。
' だから k 文字目のコストは k に比例し、全体は文字数の 2 乗で効く。
' 速くする工夫は入れない（要件どおり、意図的に重い）。
'==============================================================================
' 先頭から upto 文字目までを読み直して畳む。重さの中身はここ。
Private Function ProofFold(ByVal src As String, ByVal upto As Long) As Long
    Dim i As Long
    Dim h As Long

    h = 5381
    For i = 1 To upto
        ' 23 bit に丸めてから 33 倍する。Long のまま掛けると溢れる。
        h = (((h And &H7FFFFF) * 33) Xor AscW(Mid$(src, i, 1))) And &H7FFFFFFF
    Next i
    ProofFold = h
End Function

' 原稿が短いとすぐ終わってしまい、FE と BE の差が見えない。そこで「原稿を
' 何周読み直すか」で総量を揃える。周回数は画面にも確認ダイアログにも出すので、
' 水増しを隠しはしない。
Private Function RoundsFor(ByVal n As Long) As Long
    Dim ops As Double

    RoundsFor = 0
    If n <= 0 Then Exit Function
    ops = CDbl(n) * CDbl(n) / 2#
    RoundsFor = CLng(PB_OPS / ops)
    If RoundsFor < 1 Then RoundsFor = 1
End Function

' 進捗を受け取り、まだ塗っていないランプだけ塗る。ランプの色はその位置の文字
' そのものから決めるので、同じ字は同じ色になる。追記だけなので塗る面積は増分。
Private Sub PaintProgress(ByVal src As String, ByVal doneChars As Long)
    Dim rg As Range
    Dim lx As Long
    Dim ly As Long
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim n As Long
    Dim want As Long
    Dim at As Long

    On Error Resume Next
    Set rg = m_ws.Range("pb_lamps")
    If rg Is Nothing Then Exit Sub
    n = Len(src)
    If n <= 0 Then Exit Sub
    lx = PxOfCol(rg.Column)
    ly = PxOfRow(rg.Row)
    want = CLng(CDbl(doneChars) * LAMP_N / CDbl(n))
    If want > LAMP_N Then want = LAMP_N
    For i = m_lampShown To want - 1
        r = i \ LAMP_COLS
        c = i Mod LAMP_COLS
        at = CLng(CDbl(i) * n / CDbl(LAMP_N)) + 1
        If at < 1 Then at = 1
        If at > n Then at = n
        Fill lx + c * LAMP_PX, ly + r * LAMP_PX, LAMP_PX, LAMP_PX, _
             LampColor(Abs(AscW(Mid$(src, at, 1))) Mod 10)
    Next i
    If want > m_lampShown Then m_lampShown = want
End Sub

Private Sub ClearLamps()
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range("pb_lamps")
    If rg Is Nothing Then Exit Sub
    Fill PxOfCol(rg.Column), PxOfRow(rg.Row), LAMP_COLS * LAMP_PX, LAMP_ROWS * LAMP_PX, C_BODY
    m_lampShown = 0
End Sub

'==============================================================================
' 実行
'
' 原稿（メモ帳 A）を清書（メモ帳 B）へ写す。同じ仕事を、
'   FE：この Excel の中で同期に回す。描画は戻さないので Excel は本当に応答
'       しなくなり、Windows がタイトルバーに「(応答なし)」を付ける。メモ帳 B
'       は終わるまで 1 文字も増えない。演出は入れない。
'   BE：非表示の別プロセス Excel へ投げる。FE は 1 秒ポンプが生きたままなので、
'       清書がメモ帳 B へ少しずつ伸びていくのが見えるし、時計もミニマップも
'       UIA 同期も動き続ける。
' という 2 通りで走らせて見比べる。それがこのショーケースの主題。
'==============================================================================
Private Sub RunJob(ByVal side As String)
    Dim ans As VbMsgBoxResult
    Dim src As String
    Dim n As Long
    Dim a0 As Long
    Dim a1 As Long

    If Len(m_job) > 0 Then Exit Sub
    If Not m_npBound(0) Or Not m_npBound(1) Then
        SetTxt "pb_jobstat", "メモ帳が 2 枚要ります（原稿と清書）"
        MsgBox "このデモにはメモ帳の窓が 2 つ要ります。" & vbCrLf & vbCrLf & _
               "文章が入っているほうが「原稿」、空のほうが「清書」になります。" & vbCrLf & _
               "メモ帳を 2 枚（片方は空のまま）開いてから実行してください。" & vbCrLf & vbCrLf & _
               "アプリはメモ帳を起動も終了もしません。開いている窓につなぐだけです。", _
               vbExclamation, PB_APP
        Exit Sub
    End If

    ' 役割は「並び順」ではなく「中身」で決める。文章が入っているほうが原稿、
    ' 空のほうが清書。並び順で決めていたときは、たまたま先に見つかった窓が
    ' 清書になり、利用者の書きかけを上書きしかけた。清書側は必ず上書きする
    ' のだから、そこは空でなければならない。両方に文章があるなら実行しない。
    a0 = Len(Trim$(CStr(m_npVal(0).CurrentValue)))
    a1 = Len(Trim$(CStr(m_npVal(1).CurrentValue)))
    If a0 > 0 And a1 > 0 Then
        SetTxt "pb_jobstat", "両方に文章があります ･ 清書用に空のメモ帳を 1 枚"
        MsgBox "メモ帳が 2 枚とも埋まっています。" & vbCrLf & vbCrLf & _
               Left$(m_npTitle(0), 40) & "（" & a0 & " 文字）" & vbCrLf & _
               Left$(m_npTitle(1), 40) & "（" & a1 & " 文字）" & vbCrLf & vbCrLf & _
               "このデモは清書側を上書きします。消えては困る文章なので実行しません。" & vbCrLf & _
               "空のメモ帳を 1 枚用意して、そちらが清書になるようにしてください。", _
               vbExclamation, PB_APP
        Exit Sub
    End If
    If a0 = 0 And a1 = 0 Then
        SetTxt "pb_jobstat", "原稿が空です ･ どちらかに文章を書いてください"
        MsgBox "メモ帳が 2 枚とも空です。" & vbCrLf & _
               "どちらかに文章を書いてから実行してください。そちらが原稿になります。", _
               vbExclamation, PB_APP
        Exit Sub
    End If
    If a0 > 0 Then
        m_srcSlot = 0
        m_dstSlot = 1
    Else
        m_srcSlot = 1
        m_dstSlot = 0
    End If
    src = ReadManuscript()
    n = Len(src)
    If n < 8 Then
        SetTxt "pb_jobstat", "原稿が短すぎます（" & n & " 文字）"
        MsgBox "原稿（メモ帳 A）に文章を書いてから実行してください。" & vbCrLf & _
               "いまは " & n & " 文字です。", vbExclamation, PB_APP
        Exit Sub
    End If

    m_jobSrc = src
    m_jobRounds = RoundsFor(n)
    ans = MsgBox("原稿 " & n & " 文字を、1 文字ずつ清書します。" & vbCrLf & _
                 "1 文字書くたびに、そこまでを頭から読み直して突き合わせます" & vbCrLf & _
                 "（校正 " & m_jobRounds & " 周。高速化はしません＝意図的に重い処理です）。" & vbCrLf & vbCrLf & _
                 "原稿：" & Left$(m_npTitle(m_srcSlot), 30) & vbCrLf & _
                 "清書：" & Left$(m_npTitle(m_dstSlot), 30) & "（中身は上書きされます）" & vbCrLf & vbCrLf & _
                 IIf(side = "FE", "この Excel の中で回すので、終わるまで操作できません。" & vbCrLf & _
                                  "清書はいっぺんに出ます。", _
                                  "非表示の別プロセス Excel が回します。この画面は動き続け、" & vbCrLf & _
                                  "清書は少しずつ伸びていきます。") & vbCrLf & vbCrLf & _
                 "よろしいですか", vbOKCancel + vbQuestion, PB_APP)
    If ans <> vbOK Then Exit Sub

    m_job = side
    m_jobT0 = Timer
    m_repaints = 0
    m_jobDone = 0
    ClearLamps
    ClearClean
    If side = "FE" Then
        RunJobHere
    Else
        StartJobOnBe
    End If
End Sub

' 原稿はメモ帳 A から、そのとき見えている文字をそのまま取る。
Private Function ReadManuscript() As String
    On Error Resume Next
    ReadManuscript = ToLf(CStr(m_npVal(m_srcSlot).CurrentValue))
End Function

' 清書をメモ帳 B へ書く。書いたぶんは同期側の控えにも入れておく。そうしないと
' 次のティックで「B が勝手に変わった」と見なして Excel 側へ引き戻してしまう。
Private Sub WriteClean(ByVal s As String)
    On Error Resume Next
    If Not m_npBound(m_dstSlot) Then Exit Sub
    If s = m_cleanShown Then Exit Sub
    m_npVal(m_dstSlot).SetValue ToCrLf(s)
    m_npLastText(m_dstSlot) = ToCrLf(s)
    m_npLastCell(m_dstSlot) = s
    m_cleanShown = s
    SetTxt "pb_get", "→ SetValue " & Format$(Now, "hh:nn:ss")
End Sub

Private Sub ClearClean()
    On Error Resume Next
    m_cleanShown = ""
    WriteClean ""
End Sub

' FE 側。DoEvents も入れないので、この間 Excel は本当に固まる。
Private Sub RunJobHere()
    Dim k As Long
    Dim r As Long
    Dim ms As Double
    Dim n As Long
    Dim sink As Long

    On Error GoTo Failed
    n = Len(m_jobSrc)
    ' 固まる直前の状態を 1 度だけ確定させておく（固まったあとは何も描けない）
    SetTxt "pb_fe_state", "実行中（固まります）"
    SetTxt "pb_jobstat", "FE で校正中 ･ " & n & " 文字 × " & m_jobRounds & " 周"
    Release

    Hold
    For r = 1 To m_jobRounds
        For k = 1 To n
            sink = sink Xor ProofFold(m_jobSrc, k)
        Next k
    Next r
    ms = (Timer - m_jobT0) * 1000
    If ms < 0 Then ms = ms + 86400000#
    PaintProgress m_jobSrc, n
    WriteClean m_jobSrc
    SetTxt "pb_fe_state", "待機 pid " & PidText()
    SetTxt "pb_jobstat", "FE 完了 " & Format$(ms / 1000, "0.0") & " 秒 ･ " & n & " 文字 ･ 描画 " & _
        (m_repaints + 1) & " 回"
    m_job = ""
    Release
    MsgBox "FE（この Excel）で完了しました。" & vbCrLf & vbCrLf & _
           "所要時間：" & Format$(ms / 1000, "0.00") & " 秒" & vbCrLf & _
           "原稿：" & n & " 文字 × 校正 " & m_jobRounds & " 周" & vbCrLf & _
           "描画回数：" & m_repaints & " 回" & vbCrLf & vbCrLf & _
           "実行中この画面は固まっていて、清書も最後にまとめて出ました。", vbInformation, PB_APP
    Exit Sub
Failed:
    m_job = ""
    Release
    SetTxt "pb_jobstat", "FE エラー " & Err.Number & " " & Err.Description
End Sub

Private Function PidText() As String
    On Error Resume Next
    PidText = CStr(m_winPid(0))
End Function

' BE 側へ投げる。原稿そのものは JSON に載せない。長い本文を JSON へ押し込むと
' 引用符と改行の始末で必ず事故る。原稿は別ファイルに置き、JSON では「どこを
' 読め」だけを伝える。
Private Sub StartJobOnBe()
    Dim js As String

    If Not m_beStarted Then
        SetTxt "pb_jobstat", "BE が起動していません ･ " & m_beNote
        m_job = ""
        Exit Sub
    End If
    WriteAll SrcPath(), m_jobSrc
    m_reqSeq = m_reqSeq + 1
    js = "{""cmd"":""proof"",""seq"":" & m_reqSeq & ",""n"":" & Len(m_jobSrc) & _
         ",""rounds"":" & m_jobRounds & _
         ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    WriteAll CmdPath(), js
    Flash "pb_lane_out"
    SetTxt "pb_jobstat", "BE へ投げました ･ この画面は動き続けます"
    SetTxt "pb_be_state", "実行中"
End Sub

'==============================================================================
' FE と BE のあいだ ─ Temp の JSON だけで往復する
'==============================================================================
Private Function TempDir() As String
    TempDir = Environ$("TEMP")
    If Right$(TempDir, 1) <> "\" Then TempDir = TempDir & "\"
    TempDir = TempDir & "pixelbridge\"
End Function

Private Sub EnsureTempDir()
    On Error Resume Next
    If Dir$(TempDir(), vbDirectory) = "" Then MkDir Left$(TempDir(), Len(TempDir()) - 1)
End Sub

Private Function CmdPath() As String
    CmdPath = TempDir() & "command.json"
End Function

Private Function ProgPath() As String
    ProgPath = TempDir() & "progress.json"
End Function

Private Function FePath() As String
    FePath = TempDir() & "fe.json"
End Function

' 原稿と清書は JSON に載せず、素のテキストで置く。長い本文を JSON へ押し込むと
' 引用符と改行の始末で必ず事故る。
Private Function SrcPath() As String
    SrcPath = TempDir() & "manuscript.txt"
End Function

Private Function CleanPath() As String
    CleanPath = TempDir() & "clean.txt"
End Function

' BE 用ブックの名前に空白を入れない。実証済みの実装（src\app\vba）は
' ReaderDataViewer_worker_<sid>.xlsm という空白なしの名前を使っている。
' Application.Run は 'a b.xlsm'!Proc を解決できるが、Application.OnTime の
' 名前解決はそれとは別物で、空白入りだと仕掛けても発火しなかった（実測）。
Private Function BeBookPath() As String
    BeBookPath = ThisWorkbook.Path & "\pixelbridge_be.xlsm"
End Function

Private Sub WriteAll(ByVal path As String, ByVal s As String)
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
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

' ファイルはバイトで読む。
'
' 以前は Open ... For Input + Input$(LOF(f), f) だった。これは「LOF バイト分の
' 文字」を要求する読み方で、中身が ASCII のうちは動く（バイト数＝文字数）。
' 日本語を入れた瞬間に壊れた。CP932 では 1 文字 2 バイトなので、文字数で数える
' と必ずファイル末尾を越え、「Input past end of file」で読み取り全体が落ちる。
' 実測：原稿を渡した BE が毎回「原稿が読めません」を返した。
' バイトで読んで StrConv で戻せば、JSON も日本語の本文も同じ道で扱える。
Private Function ReadAll(ByVal path As String) As String
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

Private Function JVal(ByVal js As String, ByVal key As String) As String
    Dim a As Long
    Dim b As Long
    Dim c As String
    a = InStr(1, js, """" & key & """", vbTextCompare)
    If a = 0 Then Exit Function
    a = InStr(a, js, ":")
    If a = 0 Then Exit Function
    a = a + 1
    Do While a <= Len(js)
        If Mid$(js, a, 1) <> " " Then Exit Do
        a = a + 1
    Loop
    If Mid$(js, a, 1) = """" Then
        a = a + 1
        b = InStr(a, js, """")
        If b = 0 Then b = Len(js) + 1
    Else
        b = a
        Do While b <= Len(js)
            c = Mid$(js, b, 1)
            If c = "," Or c = "}" Then Exit Do
            b = b + 1
        Loop
    End If
    JVal = Trim$(Mid$(js, a, b - a))
End Function

Private Sub KillQuiet(ByVal path As String)
    On Error Resume Next
    If Dir$(path) <> "" Then Kill path
End Sub

'------------------------------------------------------------------ FE 側
' BE 用ブックは元ファイルと同じ（信頼済みの）フォルダへ置く。%TEMP% は JSON の
' やりとりだけに使う（要件 v2 §3.1）。
Private Sub PbEnsureBe()
    Dim beApp As Object
    Dim beBook As Object
    Dim prevSec As Long
    Dim copyPath As String

    On Error GoTo Failed
    EnsureTempDir
    KillQuiet CmdPath()
    KillQuiet SrcPath()
    KillQuiet CleanPath()
    KillQuiet ProgPath()
    WriteAll FePath(), "{""fe"":""alive"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"

    copyPath = BeBookPath()
    KillQuiet copyPath
    ThisWorkbook.SaveCopyAs copyPath

    Set beApp = CreateObject("Excel.Application")
    beApp.Visible = False
    beApp.DisplayAlerts = False
    beApp.ScreenUpdating = False
    beApp.EnableEvents = False
    prevSec = beApp.AutomationSecurity
    beApp.AutomationSecurity = 1                  ' msoAutomationSecurityLow
    Set beBook = beApp.Workbooks.Open(copyPath, 0, True)
    beApp.AutomationSecurity = prevSec
    ' 入口は名前で直接呼ぶ。RunAutoMacros で Auto_Open を起こす手もあるが、
    ' そちらは相手側で何かあったときに返ってこないまま止まった（実測）。名前で
    ' 呼べば失敗が COM のエラーとして返り、こちらで受け止められる。中身は
    ' OnTime を仕込むだけなので、この呼び出しはすぐ返る（要件 v2 §3.1）。
    PbLog "be: opened, calling bootstrap"
    beApp.Run "'" & Replace(beBook.Name, "'", "''") & "'!PbBeBootstrap"
    PbLog "be: bootstrap returned"
    m_beNote = "Visible=False"
    ' FE は BE への COM 参照を持ち続けない。忙しい BE を COM で叩くと FE が
    ' サーバービジーで固まる。ここで手放す。
    Set beBook = Nothing
    Set beApp = Nothing
    m_beStarted = True
    SetTxt "pb_be_state", "待機 Visible=False"
    Exit Sub
Failed:
    m_beStarted = False
    m_beNote = "起動に失敗 " & Err.Number
    PbLog "be: FAILED " & Err.Number & " " & Err.Description
    SetTxt "pb_be_state", m_beNote
End Sub

Private Sub Flash(ByVal nm As String)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If rg.Interior.Color = C_KEY Then Exit Sub
    Hold
    rg.Interior.Color = C_KEY
End Sub

Private Sub Unflash(ByVal nm As String)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If rg.Interior.Color = C_LINE Then Exit Sub
    Hold
    rg.Interior.Color = C_LINE
End Sub

Private Sub PumpChannel()
    Dim js As String
    Dim st As String
    Dim ms As Double
    Dim doneN As Long
    Dim total As Long
    Dim clean As String

    On Error Resume Next
    WriteAll FePath(), "{""fe"":""alive"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    js = ReadAll(ProgPath())
    If Len(js) = 0 Then Exit Sub
    If js = m_lastProg Then
        Unflash "pb_lane_in"
        Exit Sub
    End If
    m_lastProg = js
    Flash "pb_lane_in"
    Unflash "pb_lane_out"

    st = JVal(js, "state")
    ms = Val(JVal(js, "ms"))
    doneN = CLng(Val(JVal(js, "n")))
    total = CLng(Val(JVal(js, "total")))
    If total <= 0 Then total = Len(m_jobSrc)

    ' 届いた清書をメモ帳 B へ写す。ここが「FE が生きている」いちばんの証拠に
    ' なる。BE が回している間もこの 1 秒ポンプが動いているから、B が伸びる。
    If doneN > 0 Then
        clean = ReadAll(CleanPath())
        Do While Len(clean) > 0 And (Right$(clean, 1) = vbLf Or Right$(clean, 1) = vbCr)
            clean = Left$(clean, Len(clean) - 1)
        Loop
        If Len(clean) > 0 Then
            WriteClean clean
            PaintProgress m_jobSrc, Len(clean)
        End If
    End If

    Select Case st
        Case "running"
            SetTxt "pb_be_state", "実行中 " & doneN & " / " & total & " 文字"
            SetTxt "pb_jobstat", "BE 実行中 ･ 校正 " & JVal(js, "round") & " / " & _
                JVal(js, "rounds") & " 周 ･ 清書 " & doneN & " / " & total & " 文字 ･ " & _
                Format$(ms / 1000, "0.0") & " 秒 ･ 描画 " & m_repaints & " 回"
        Case "done"
            SetTxt "pb_be_state", "待機 Visible=False"
            SetTxt "pb_jobstat", "BE 完了 " & Format$(ms / 1000, "0.0") & " 秒 ･ " & _
                total & " 文字 ･ 描画 " & m_repaints & " 回"
            If m_job = "BE" Then
                m_job = ""
                Release
                MsgBox "BE（非表示の別プロセス Excel）で完了しました。" & vbCrLf & vbCrLf & _
                       "所要時間：" & Format$(ms / 1000, "0.00") & " 秒" & vbCrLf & _
                       "原稿：" & total & " 文字 × 校正 " & m_jobRounds & " 周" & vbCrLf & _
                       "描画回数：" & m_repaints & " 回" & vbCrLf & vbCrLf & _
                       "実行中もこの画面は動き続け、清書は少しずつ伸びていました。", _
                       vbInformation, PB_APP
            End If
        Case "error"
            SetTxt "pb_jobstat", "BE エラー " & JVal(js, "msg")
            m_job = ""
    End Select
End Sub

'==============================================================================
' BE 側（同じ .bas が、コピーされたブックの中では BE として動く）
'==============================================================================
' BE の入口。やるのは「自分の足で歩き出す」ことだけで、すぐ返す。だから
' FE→BE の呼び出しは即 return になる（要件 v2 §3.1）。実測 8 ms。
'
' 最初はここで progress.json も書いていたが、その呼び出しで BE の Excel が
' その場で落ちた（プロセス消滅／RPC 失敗。切り分け済み）。書くのは最初の
' ティックに回す。入口は軽いほど安全でもある。
'
' 開発中、ここで仕掛けた OnTime が一度も発火せず、不可視インスタンスの
' 制約かと疑った。違った。原因は調べ方のほうにあった。BE を覗くための
' スクリプトが、その Excel の VBE で「VBAProject のコンパイル」を実行して
' いた。自動化で起こした Excel の VBE を触ると、そのインスタンスは壊れる。
' 以後 OnTime は Err.Number = 0 のまま沈黙し、常駐ループも 1 周で無応答に
' なる。VBE に触らずに同じブックを走らせれば、この bootstrap は 8 ms で
' 返り、常駐ループは idle → running → done → bye まで通る（実測）。
' 教訓：不可視の Excel を測るとき、VBE には触らない。触れば測っているのは
' 「壊したあとの Excel」になる。
Public Function PbBeBootstrap() As String
    On Error Resume Next
    ' FE は返った直後に COM 参照を捨てる。参照ゼロの自動化サーバーは次の
    ' メッセージポンプで自分を畳むので、自前所有だと宣言しておく。
    Application.UserControl = True
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    m_beQuit = False
    m_beLastSeq = 0
    ' BE 側でもフォルダを確かめる。FE がまだ作っていないことがある。
    EnsureTempDir
    Application.OnTime Now, Qual("PbBeMain")
    PbBeBootstrap = "armed"
End Function

' BE の常駐ループ。ここは不可視なので、待つのにブロックして構わない。
' Application.Wait は 1 秒刻みだが、この用途にはちょうど良い。
Public Sub PbBeMain()
    Dim js As String
    Dim seq As Long
    Dim feNow As String

    On Error Resume Next
    EnsureTempDir
    PbBeWrite "{""state"":""idle"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Do
        If m_beQuit Then Exit Do
        js = ReadAll(TempDir() & "command.json")
        seq = CLng(Val(JVal(js, "seq")))
        If seq > m_beLastSeq Then
            m_beLastSeq = seq
            Select Case JVal(js, "cmd")
                Case "proof"
                    BeProofread CLng(Val(JVal(js, "n"))), CLng(Val(JVal(js, "rounds")))
                Case "quit"
                    Exit Do
            End Select
        End If
        ' FE が消えたら自分から終わる。置き去りの BE を残さない。
        '
        ' 「fe.json が無い」だけを見ていたのでは足りない。FE が正常に閉じた
        ' ときは消えるが、強制終了されるとファイルは残ったままになり、この
        ' BE はそれを「FE は生きている」と読んで永遠に回り続ける（実測：FE を
        ' kill したあと BE だけが残った）。だから中身が更新されているかで見る。
        ' FE は毎ティック時刻を書き直すので、変わらなければ止まっている。
        ' 待つ長さは長めに取る。FE は正常でも黙る時間がある。起動直後の最初の
        ' 再描画（1px なら数十秒）、そして「FE で清書」を選んだときは仕様どおり
        ' 15 秒ほど完全に固まる。短く見切ると、その間に BE が自分で終わってしまう
        ' （実測：8 秒にしていたら、FE の起動直後に BE が bye を書いて消えた）。
        feNow = ReadAll(FePath())
        If Len(feNow) = 0 Or feNow = m_beFeSeen Then
            m_beMissing = m_beMissing + 1
            If m_beMissing > 90 Then Exit Do
        Else
            m_beMissing = 0
        End If
        m_beFeSeen = feNow
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    BeQuit
End Sub

' BE 側の仕事。原稿は JSON ではなく別ファイルから読む。
' 進捗は progress.json、清書そのものは clean.txt。本文を JSON に入れないので、
' 引用符も改行も気にしなくていい。
Private Sub BeProofread(ByVal n As Long, ByVal rounds As Long)
    Dim src As String
    Dim k As Long
    Dim r As Long
    Dim t0 As Double
    Dim ms As Double
    Dim chunk As Long
    Dim done As Long
    Dim sink As Long

    On Error GoTo Failed
    src = ReadAll(SrcPath())
    ' WriteAll の Print # は末尾に改行を足す。原稿の長さがずれるので落とす。
    Do While Len(src) > 0 And (Right$(src, 1) = vbLf Or Right$(src, 1) = vbCr)
        src = Left$(src, Len(src) - 1)
    Loop
    If Len(src) = 0 Then
        PbBeWrite "{""state"":""error"",""msg"":""原稿が読めません"",""ts"":""" & _
            Format$(Now, "hh:nn:ss") & """}"
        Exit Sub
    End If
    n = Len(src)
    If rounds < 1 Then rounds = 1
    ' 進捗は 40 回くらいに割る。細かすぎると書き出しのほうが重くなる。
    chunk = n \ 40
    If chunk < 1 Then chunk = 1
    t0 = Timer

    For r = 1 To rounds
        For k = 1 To n
            sink = sink Xor ProofFold(src, k)
            ' メッセージを捌く。これは演出でも手加減でもなく、行儀の問題。
            '
            ' ここを入れずに測ったとき、FE は BE の実行中ずっと固まった。原因は
            ' FE 側の仕事ではない。FE の 1 秒ティックにある ScanWindows（UI
            ' Automation の列挙）が、この BE を待って返らなかった。実測：ティックの
            ' t:sel から t:scan までが 15.6 秒で、BE の処理時間とぴたり一致した。
            ' UIA の列挙は端から端まで全部のトップレベル窓に触りにいくので、
            ' その中に「メッセージを一切処理しないプロセス」が一つでもあると、
            ' そこで止まる。不可視でも窓は窓で、ここに BE 自身がいた。
            ' つまり「重い処理を別プロセスへ逃がす」だけでは足りず、逃がした先が
            ' 応答し続けないと、結局 FE が巻き添えで固まる。
            DoEvents
            If (k Mod chunk) = 0 Or k = n Then
                ' 清書は「全体の何割まで進んだか」で伸ばす。最終周だけで伸ばすと
                ' 実行時間のほとんどで 0 のままになり、動いて見えない（実測）。
                done = CLng((CDbl(r - 1) * n + k) / (CDbl(rounds) * n) * n)
                If done < 1 Then done = 1
                If done > n Then done = n
                WriteAll CleanPath(), Left$(src, done)
                ms = (Timer - t0) * 1000
                If ms < 0 Then ms = ms + 86400000#
                PbBeWrite "{""state"":""running"",""n"":" & done & ",""total"":" & n & _
                    ",""round"":" & r & ",""rounds"":" & rounds & ",""ms"":" & CLng(ms) & _
                    ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
            End If
        Next k
    Next r

    WriteAll CleanPath(), src
    ms = (Timer - t0) * 1000
    If ms < 0 Then ms = ms + 86400000#
    PbBeWrite "{""state"":""done"",""n"":" & n & ",""total"":" & n & _
        ",""rounds"":" & rounds & ",""ms"":" & CLng(ms) & _
        ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Exit Sub
Failed:
    PbBeWrite "{""state"":""error"",""msg"":""" & Err.Number & " " & _
        Replace(Err.Description, """", "'") & """,""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
End Sub

' 終わり方の順序が効く。ThisWorkbook.Close を先に呼ぶと、そこで自分のコードが
' 止まって Application.Quit に届かず、ブックを持たない不可視 Excel が残る（実測）。
Private Sub BeQuit()
    On Error Resume Next
    m_beQuit = True
    m_beAlive = False
    PbBeWrite "{""state"":""bye"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    Application.Quit
End Sub

Private Sub PbBeWrite(ByVal s As String)
    Dim pth As String
    pth = TempDir() & "progress.json"
    WriteAll pth, s
End Sub

'==============================================================================
' 書体
'==============================================================================
Private Sub PickFonts()
    m_fontJp = PickOne("Noto Sans JP", "NotoSansJP*", "Yu Gothic UI")
    m_fontMono = PickOne("Noto Sans Mono", "NotoSansMono*", "Consolas")
End Sub

Private Function PickOne(ByVal want As String, ByVal filePat As String, _
                         ByVal fallback As String) As String
    On Error Resume Next
    If Len(Dir$(Environ$("WINDIR") & "\Fonts\" & filePat)) > 0 Then
        PickOne = want
        Exit Function
    End If
    If Len(Dir$(Environ$("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\" & filePat)) > 0 Then
        PickOne = want
        Exit Function
    End If
    PickOne = fallback
End Function
