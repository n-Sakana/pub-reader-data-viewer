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
' 実測して決めたこと（測定は showcase\README.md に記録）
'   - Application.OnTime は秒単位に量子化される。0.5 秒は出ないので 1 秒。
'   - ScreenUpdating の False → True は、変えた量に関係なくビューポート全体を
'     無効化する。その再描画の値段は見えているセルの「色の切り替わり数」で
'     決まり、1px の実盤面では 1 回 2.4 秒（このノート、実測）。逆に、True の
'     まま少しだけ書けば汚した矩形しか描き直さず、1px でも 1 秒に収まる
'     （同じ盤で直書き 1016ms、トグル 1711ms、素の盤ならどちらも 1000ms）。
'     だから毎秒のティックはトグルせずに直書きし、Hold / Release は初期構築・
'     後始末・FE 清書（わざと固める）だけに使う。ここが軽さの本体。
'   - IUIAutomation.ElementFromPoint は POINT を ByVal で取るため VBA から
'     呼べない（コンパイルエラー）。点 → 窓は矩形ヒットテストで出す。
'   - デスクトップ直下の列挙は UIA が Excel 自身にも尋ねるので詰まり得る。
'     100 周 x 3 通りでは詰まらなかったが（平均 118ms）安全の保証にはならない。
'     だから 4 ティックに 1 度だけ、しかも ClassName="Notepad" に絞って呼ぶ。
'==============================================================================
Option Explicit

'------------------------------------------------------------------ identity
Public Const PB_APP As String = "PIXEL BRIDGE"
Public Const PB_SUB As String = "notepad.exe / UI Automation ･ FE ⇔ BE ･ pi 15s"
Private Const PB_SHEET As String = "PIXELBRIDGE"
Private Const PB_BE_MARK As String = "_be"

'------------------------------------------------------------------ geometry
Private Const PB_UNIT As Long = 4              ' ビルドが 1 / 2 / 4 で上書きする。単体で取り込むとこの値
Private Const PB_W As Long = 812
Private Const PB_H As Long = 812
Private Const PB_DOTS As Long = 8
' 追う窓は 2 つ：この Excel と、つなぐメモ帳 1 枚。メモ帳は 1 枚だけを見る
' （owner 指示 2026-08-19）。2 枚目のスロット・選択・切替は全部落としてある。
Private Const PB_MAXWIN As Long = 2
Private Const PB_SCANEVERY As Long = 4         ' 窓の探し直しは何ティックに 1 度か
Private Const PB_MAXBTN As Long = 8

' 進捗の盤。モックと同じ 67 x 12 = 804 ランプ。円周率の桁を左上から右下へ
' 1 桁 1 ランプで置き、桁の値（0～9）がそのままランプの色になる。ランプ 1 つは
' PB_UNIT の倍数にしてあるので、1px でも 2px でも 4px でもセルに割り切れる。
Private Const LAMP_COLS As Long = 67
Private Const LAMP_ROWS As Long = 12
Private Const LAMP_PX As Long = 4              ' ランプ 1 つの設計 px
Private Const LAMP_N As Long = LAMP_COLS * LAMP_ROWS

' ベンチの長さ（秒）。窓の配置と表示用 Excel の用意が終わってから測りはじめ、
' この時間で自動的に終わる。
Private Const PB_BENCHSECS As Double = 15#

' 表示用 Excel の並べ方。1 セルに 10 桁、1 行に 10 セル＝ 100 桁。
Private Const PI_PERCELL As Long = 10
Private Const PI_PERROW As Long = 10

'------------------------------------------------------------------ UIA ids
Private Const TS_CHILDREN As Long = 2
Private Const TS_DESCENDANTS As Long = 4
Private Const UIA_ControlTypePropertyId As Long = 30003
Private Const UIA_HasKeyboardFocusPropertyId As Long = 30008
Private Const UIA_ClassNamePropertyId As Long = 30012
Private Const UIA_NativeWindowHandlePropertyId As Long = 30020
Private Const UIA_IsValuePatternAvailablePropertyId As Long = 30043
Private Const UIA_ValuePatternId As Long = 10002
' TransformPattern は 10016。ここが 10003（= RangeValuePattern）だったせいで、
' メモ帳の窓に GetCurrentPattern を投げても常に Nothing が返り、整列もドラッグ
' 配置も毎回 "movewin: no transform" で落ちていた。同じ窓を .NET の管理 UIA で
' 見ると CanMove=True が取れる、という食い違いの正体はプラットフォームの差では
' なくこの定数の取り違え（2026-08-19 実測。この端末のメモ帳 11.2606 は窓要素で
' 10016 と 10009 の 2 つを公開している）。
Private Const UIA_TransformPatternId As Long = 10016
Private Const UIA_WindowPatternId As Long = 10009
Private Const UIA_IsTransformPatternAvailablePropertyId As Long = 30042
Private Const UIA_DocumentControlTypeId As Long = 50030
Private Const UIA_EditControlTypeId As Long = 50004
Private Const WindowVisualState_Normal As Long = 0
Private Const WindowVisualState_Minimized As Long = 2

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
Private m_devPt As Double                      ' 1 デバイスピクセルのポイント数（実測）
Private m_pxPerPt As Double                    ' 画面ピクセル / ポイント（= 1 / m_devPt）
Private m_running As Boolean
Private m_held As Boolean                      ' いま描画を止めているか
Private m_dirty As Boolean                     ' このティックで何か描いたか
Private m_startAt As Double
Private m_ticks As Long
Private m_armed As Boolean                     ' 予約が 1 本生きているか
Private m_nextTick As Date                     ' その予約の時刻（取り消しに要る）
Private m_inTick As Boolean                    ' ティックの再入防止
Private m_polls As Long
Private m_syncOn As Boolean
Private m_dotAt As Long
Private m_dotShown As Long
Private m_fontJp As String
Private m_fontMono As String
Private m_lastSel As String

Private m_uia As UIAutomationClient.IUIAutomation
Private m_root As UIAutomationClient.IUIAutomationElement
Private m_uiaNote As String
Private m_myPid As String                      ' 起動時に 1 回だけ読む（下記 ReadOwnPid）
Private m_scrL As Long, m_scrT As Long, m_scrR As Long, m_scrB As Long

' 追える窓：0 = この Excel、1 = メモ帳
Private m_winN As Long
Private m_winL(0 To PB_MAXWIN - 1) As Long
Private m_winT(0 To PB_MAXWIN - 1) As Long
Private m_winR(0 To PB_MAXWIN - 1) As Long
Private m_winB(0 To PB_MAXWIN - 1) As Long
Private m_winLabel(0 To PB_MAXWIN - 1) As String
Private m_winApp(0 To PB_MAXWIN - 1) As String
Private m_winPid(0 To PB_MAXWIN - 1) As String
Private m_winNp(0 To PB_MAXWIN - 1) As Long    ' -1 = Excel、0 = メモ帳
Private m_mapSig As String

' つなぐメモ帳。1 枚だけなので配列も選択状態も持たない。動かす相手は常にこれ。
Private m_npWin As UIAutomationClient.IUIAutomationElement
Private m_npDoc As UIAutomationClient.IUIAutomationElement
Private m_npVal As UIAutomationClient.IUIAutomationValuePattern
Private m_npTitle As String
Private m_npHwnd As Variant
Private m_npPid As String
Private m_npLastText As String
Private m_npLastCell As String
Private m_npBound As Boolean
Private m_npFresh As Boolean                   ' 結んだ直後＝まずメモ帳から取り込む
' 最小化されている（矩形が画面の外へ飛ぶ）。掴んだ要素は手放さない。手放すと
' 「最小化したら二度と戻せない」になる。
Private m_npMin As Boolean
Private m_frameOn As Long

' ボタン（セルで描いてあり、押下は選択セルで拾う）
Private m_btnN As Long
Private m_btnKey(0 To PB_MAXBTN - 1) As String
Private m_btnCell(0 To PB_MAXBTN - 1) As String

' FE / BE
Private m_beStarted As Boolean
Private m_beNote As String
Private m_reqSeq As Long
Private m_repaints As Long                     ' 実際に描画を戻した回数
Private m_lastProg As String
Private m_beMissing As Long
Private m_beFeSeen As String                   ' 前に見た fe.json の中身
Private m_beLastSeq As Long
Private m_beQuit As Boolean
Private m_beStopping As Boolean                ' quit を送って bye 待ち
Private m_beStopTicks As Long

' 円周率ベンチ
Private m_bench As Boolean                     ' 回っている最中か
Private m_benchT0 As Double                    ' 15 秒の起点（配置と用意が済んだ時刻）
Private m_benchMs As Double                    ' 終わったときの所要（表示に残す）
Private m_piCalc As Long                       ' BE が計算した桁数
Private m_piShown As Long                      ' 表示用 Excel へ実際に書いた桁数
Private m_piText As String                     ' 受け取った桁（"3." のあとの小数部）
Private m_lampShown As Long                    ' もう塗ったランプの数
' 表示用 Excel。BE とは別プロセスの、保存しない一時ブック。FE がここへ書く。
Private m_dispApp As Object
Private m_dispBook As Object
Private m_dispSheet As Object

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
    PbLog "show: begin"
    Hold
    PickFonts
    ' BE は画面を作る前に起こす。作ったあとだと SaveCopyAs が 16 万セルの盤面
    ' ごと複製してしまい（94KB が 557KB）、その読み込みで FE が詰まる（実測）。
    PbEnsureBe
    PbMakeSheet
    ReadOwnPid
    PbLog "show: sheet pitch=" & m_pitch & " pid=" & m_myPid
    ' 盤面を前に出すのは、組み終わってから一度だけ。組む前に出すと 1px では
    ' 表示設定の変更で再配置が走り続けて戻ってこない（実測：CPU 478 秒）。
    PbBuildScreen
    PbLog "show: built"
    PbScrollHome
    PbLog "show: sheet shown"
    PbCheckFit
    PbLog "show: fitted"
    Release
    SetTxt "pb_fe_state", "動作中"
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
' Hold / Release は「大きく描き換えるとき」だけの道具。初期構築、後始末、
' FE 清書（わざと固める演出）で使う。毎秒のティックでは使わない。
' False → True のトグルは変えた量に関係なくビューポート全体を無効化し、
' その再描画は盤面が複雑なほど高く付く（1px の実盤面で 1 回 2.4 秒、実測）。
' ティックの小さな更新は ScreenUpdating を True のまま直に書く。汚した矩形
' しか描き直されないので、1px でも 1 秒に収まる（実測 1016ms）。
'==============================================================================
Private Sub Hold()
    If m_held Then Exit Sub
    Application.ScreenUpdating = False
    m_held = True
End Sub

' Release 1 回 =「まとめ描き」1 回。ティック側の直書きは m_dirty で数え、
' どちらも m_repaints に足す。演出ではなく実測値として画面に出す。
Private Sub Release()
    If Not m_held Then Exit Sub
    Application.ScreenUpdating = True
    m_held = False
    m_repaints = m_repaints + 1
    m_dirty = False                            ' まとめ描きが吸収した
End Sub

'==============================================================================
' 疑似ピクセルの寸法
'==============================================================================
Private Sub PbMakeSheet()
    Set m_ws = EnsureSheet()
    ' 幅を詰める前に表示設定を済ませる（PbSheetChrome の説明を参照）
    PbSheetChrome
    ResetSheet
    ' 刻みを選ぶ前に、窓を最終の置き場所（画面の左半分）へ置いてしまう。
    ' UsableWidth / UsableHeight は「いまの窓」の広さなので、最大化したまま
    ' 測って刻みを選ぶと、左半分に収まらない盤面を作ってしまう。
    MeasureDevPt
    MeasurePxPerPt
    PbReadScreen
    PbPlaceExcel
    PbCalibrate
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

' 画面の大きさ（デバイス px）。UIA の root の矩形がそのまま画面全体になる。
' 取れない端末では、最大化した窓の大きさで代用する。
Private Sub PbReadScreen()
    Dim rc As UIAutomationClient.tagRECT
    On Error GoTo Fallback
    If Not EnsureUia() Then GoTo Fallback
    rc = m_root.CurrentBoundingRectangle
    If rc.Right - rc.Left < 320 Or rc.Bottom - rc.Top < 240 Then GoTo Fallback
    m_scrL = rc.Left: m_scrT = rc.Top: m_scrR = rc.Right: m_scrB = rc.Bottom
    Exit Sub
Fallback:
    On Error Resume Next
    Application.WindowState = xlMaximized
    m_scrL = 0: m_scrT = 0
    m_scrR = CLng(Application.Width * m_pxPerPt)
    m_scrB = CLng(Application.Height * m_pxPerPt)
    If m_scrR < 320 Then m_scrR = 1280
    If m_scrB < 240 Then m_scrB = 720
End Sub

' 起動時に一度だけ、この Excel を画面の左半分へ合わせる。**以後この窓は
' 位置も大きさも変えない**（owner 指示）。ミニマップのドラッグ配置も
' 左右 / 上下 2 分割も、動かすのは起動済みのメモ帳だけ。
Private Sub PbPlaceExcel()
    On Error Resume Next
    If m_pxPerPt <= 0 Then Exit Sub
    If Application.WindowState <> xlNormal Then Application.WindowState = xlNormal
    Application.Left = m_scrL / m_pxPerPt
    Application.Top = m_scrT / m_pxPerPt
    Application.Width = ((m_scrR - m_scrL) \ 2) / m_pxPerPt
    Application.Height = (m_scrB - m_scrT) / m_pxPerPt
    PbLog "  place: screen " & (m_scrR - m_scrL) & "x" & (m_scrB - m_scrT) & _
        " app " & Application.Left & "," & Application.Top & " " & _
        Application.Width & "x" & Application.Height
End Sub

Private Sub PbCalibrate()
    Dim devPt As Double
    Dim k As Long
    Dim best As Long
    Dim bestErr As Double
    Dim e As Double
    Dim cells As Long
    Dim kFit As Long
    Dim kf As Long

    devPt = m_devPt
    If devPt <= 0 Then
        MeasureDevPt
        devPt = m_devPt
    End If

    If GridAlreadyOk() Then Exit Sub
    cells = PB_W \ PB_UNIT

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

    ' ただし、その寸法で盤面がこの窓に収まらないなら意味がない。この端末
    ' （1920x1080 / 125%）では 0.75pt どおりだと盤面が 1015～1218 デバイス px に
    ' なり、下 4 割（実行ボタンとフッタ）が最初から画面の外だった（実測：
    ' fit visible 204x158 / need 203x203）。だから「窓に入る最大の整数
    ' デバイス px」を上限にする。窓はもう最終の場所（左半分）に置いてあり、
    ' 見出しもタブも横スクロールバーも消した後なので、UsableHeight /
    ' UsableWidth がそのままシートの見える広さになる。端の列を半分だけ
    ' 見せないための「必要数 + 1」まで入る大きさを選ぶ。
    '
    ' 以前はここで縦に 48 デバイス px の「沈み」を足していた。最大化した
    ' 作業領域（1020 デバイス px）で測ると 812 行が入らず、窓を画面の下へ
    ' はみ出させる前提だったから。いまは窓自身を画面の高さいっぱい（1080）に
    ' 置くので下駄は要らない。幅からは縦スクロールバーぶん（UsableWidth は
    ' これを含む。実測）の 24 デバイス px を引いておく。
    kFit = Int((ActiveWindow.UsableHeight / devPt) / (cells + 1))
    kf = Int((ActiveWindow.UsableWidth / devPt - 24) / (cells + 1))
    If kf < kFit Then kFit = kf
    If kFit < 1 Then kFit = 1
    If best > kFit Then best = kFit

    m_pitch = best * devPt
    m_colW = FindColWidth(m_pitch)
    m_ws.Range(m_ws.Columns(1), m_ws.Columns(cells)).ColumnWidth = m_colW
    m_ws.Range(m_ws.Rows(1), m_ws.Rows(cells)).RowHeight = m_pitch
    PbLog "  cal: devPt=" & devPt & " pitch=" & m_pitch & " colw=" & m_colW & _
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

' Excel の座標（ポイント）と UIA の座標（画面ピクセル）の比。
' 要件 v2 §7 の「単位系を UIA に統一（DPI 問題の回避）」はこれで満たす。
'
' 以前はここで ElementFromHandle(Application.hwnd) を使っていた。自分の窓の
' UIA 要素を作ると Excel 自身の UIA プロバイダがこのプロセスの中で目を覚まし、
' 以後（毎ティックのフォーカス取得でも同じことが起きて）セルへの書き込み
' 1 回ごとに約 130ms の同期コストが乗る（実測。素の盤なら 5ms）。
' 比は行高の実測（1 デバイスピクセル = m_devPt ポイント）から出せるので、
' 自分の窓には UIA で触らない。
Private Sub MeasurePxPerPt()
    m_pxPerPt = 0
    If m_devPt > 0 Then m_pxPerPt = 1# / m_devPt
End Sub

' 窓の大きさは起動時に一度決めたきり動かさない（PbPlaceExcel）。ここは
' 「入っているか」を見て記録するだけ。足りないときに窓を広げる（旧
' PbFitRefine）のは owner 指示の「初期配置後は Excel を一切動かさない」に
' 反するのでやらない。入る刻みを選ぶのは PbCalibrate の仕事で、そちらは
' 置いたあとの窓の実寸（UsableWidth / UsableHeight）で決めている。
'
' 見る値は UsableWidth / UsableHeight ではなく ActiveWindow.VisibleRange。
' UsableWidth は縦スクロールバーを含み、UsableHeight はシート見出しと横
' スクロールバーを含むので、そのまま信じると設計の右端と下端が窓の外へ出る
' （実測：右下のボタンが切れ、横スクロールバーが出た）。VisibleRange なら、
' そのとき本当に見えている範囲そのもの。
Private Sub PbCheckFit()
    Dim needC As Long
    Dim needR As Long
    Dim vis As Range

    On Error Resume Next
    needC = PB_W \ PB_UNIT
    needR = PB_H \ PB_UNIT
    Set vis = ActiveWindow.VisibleRange
    If vis Is Nothing Then Exit Sub
    PbLog "  fit: visible " & vis.Columns.Count & "x" & vis.Rows.Count & _
        " need " & needC & "x" & needR
    PaintMargin vis, needC, needR
End Sub

' 盤面は 812x812 だが、窓のほうが少し広い。その余白（塗っていないセル）に、
' 1px ビルドだけ細い黒線が何本か描かれる。行が 0.6pt まで潰れた盤で、
' 結合セルのある行の帯が余白側にはみ出して描かれるためで、セルの値・書式・
' 結合・図形のどれを問い合わせても空のまま出る（実測。この線は変更前の版にも
' 同じように出ていた）。塗りのあるセルの上には出ないので、見えている余白を
' 地の色で塗って隠す。盤の外なので設計には触れていない。
Private Sub PaintMargin(ByVal vis As Range, ByVal needC As Long, ByVal needR As Long)
    Dim lastC As Long
    Dim lastR As Long

    On Error Resume Next
    lastC = vis.Column + vis.Columns.Count - 1
    lastR = vis.Row + vis.Rows.Count - 1
    If lastC > needC Then
        m_ws.Range(m_ws.Cells(1, needC + 1), m_ws.Cells(lastR, lastC)).Interior.Color = C_BODY
    End If
    If lastR > needR Then
        m_ws.Range(m_ws.Cells(needR + 1, 1), m_ws.Cells(lastR, lastC)).Interior.Color = C_BODY
    End If
    PbLog "  margin: painted to " & lastC & "," & lastR
End Sub

' 表示設定は、列幅を詰める前の軽いシートに対して変える。
'
' グリッド線と見出しの切り替えは、そのシートの列と行を全部並べ直す。列幅が
' 既定のうちは一瞬だが、812 列を 0.08 まで詰めたあと（1px、659,344 セル）に
' やると戻ってこない。実測：盤面を組み終わってからこの 3 行を実行して 276 秒。
' 4px では 41,209 セルなのでどちらでも気づかない。1px で初めて牙をむく。
'
' 横スクロールバーとシートタブも消す。1080p の縦は薄く、この 1 段（約 33
' デバイス px）が「812 行が入るか入らないか」を分ける（1920x1080 / 125% で
' 実測）。リボン・数式バー・ステータスバーは owner 指示どおり残す。
' どれも終了時に PbShutdown が戻す。
Private Sub PbSheetChrome()
    On Error Resume Next
    m_ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayWorkbookTabs = False
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
    m_dirty = True
    Dim tFill As Double
    tFill = Timer
    Cel(x, y, w, h).Interior.Color = c
    If (Timer - tFill) > 0.03 Then PbLog "    fill slow " & Format$((Timer - tFill) * 1000, "0") & _
        "ms at " & x & "," & y & " " & w & "x" & h
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
    m_dirty = True
    Set rg = Cel(x, y, w, h)
    rg.Merge
    With rg
        ' 文字列書式を先に敷く。既定のままだと "01:07" が時刻として取り込まれ、
        ' 表示が "1:07" に化ける（実機で実測。稼働時間の先頭の 0 が消えていた）。
        ' しかも読み返した値が書いた値と一致しなくなるので、SetTxt が毎ティック
        ' 書き直して無駄に描画する。
        .NumberFormat = "@"
        .Value = s
        .Font.Name = IIf(mono, m_fontMono, m_fontJp)
        .Font.Size = RoundFont(sizePx * PxPt())
        .Font.Color = colr
        .Font.bold = bold
        .HorizontalAlignment = alignH
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    ' 文字列書式にした結果、"0" や "9" のような札に Excel の
    ' 「数値が文字列として保存されています」の緑三角が出る（実機で実測）。
    ' 見た目の汚れなので、このセルに限ってそのチェックを黙らせる。
    ' 利用者の Excel 全体の設定（ErrorCheckingOptions）には触らない。
    ' Range.Errors は 1 セルにしか無い。結合した範囲のまま呼ぶと落ちるので
    ' （On Error で握り潰されて効いていなかった。実機で実測）左上のセルへ。
    On Error Resume Next
    rg.Cells(1, 1).Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
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
    m_dirty = True
    Dim tSet As Double
    tSet = Timer
    rg.Cells(1, 1).Value = s
    If (Timer - tSet) > 0.03 Then PbLog "    settxt slow " & Format$((Timer - tSet) * 1000, "0") & _
        "ms " & nm
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
    ' 小さいピッチでは名目のポイント数より実描画が広くなり、中央揃えの札は
    ' 両端から欠ける（実測：セル 3 デバイス px で「左右 2 分割」の左右が消えた）。
    ' 幅に収まるまで Excel に縮めさせる。
    rg.ShrinkToFit = True
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
    BuildBenchCard
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
    Txt "", 14, 16, 152, 25, PB_APP, 19, C_ONBAND, True, xlLeft, False
    ' 副題は等幅 37 字ぶん（約 222 設計 px）。170px の枠では "⇔ BE" が枠の外で
    ' 切れていた（実機で実測）。結合セルは溢れた分を隣へ流さず、そこで切る。
    Txt "", 172, 21, 300, 13, PB_SUB, 10, C_ONBAND, False, xlLeft, True

    ' アニメーションは常時 1 秒。速さの切替は置かない（切替ボタンは廃止）。

    Txt "pb_uptime", 596, 18, 56, 21, "00:00", 15, C_ONBAND, True, xlLeft, True
    Fill 660, 24, 10, 10, C_LAMP
    NameIt "pb_lamp", Cel(660, 24, 10, 10)
    Txt "pb_status", 676, 18, 122, 21, "同期中 poll 0", 11, C_ONBAND, False, xlLeft, True
End Sub

'------------------------------------------------------------------ ミニマップ
Private Sub BuildMiniMapCard()
    CardBox 14, 66, 784, 450
    Txt "", 28, 82, 60, 15, "MINIMAP", 11, C_TEXT, True, xlLeft, True
    ' 解像度は自動判定。モックのセレクタは表示だけにする。
    Txt "pb_res", 95, 78, 150, 22, "―", 11, C_TEXT, False, xlLeft, True
    Txt "pb_scale", 254, 83, 200, 13, "―", 10, C_SUB, False, xlLeft, True
    ' セル 3 デバイス px（この端末の 4px ビルド）でも文字が欠けないよう、
    ' 幅は 108px・文字は 10px にしてある（91px / 11px では両端が欠けた。実測）。
    ' Excel は起動時に画面の左半分へ置いたきり動かさない（owner 指示）。この
    ' 2 つが動かすのは、つないでいるメモ帳 1 枚だけ。
    Button "npmax", 560, 76, 108, 28, "最大化", _
        "メモ帳を画面の右半分いっぱいへ移動・リサイズします（OS の最大化ではありません）", 0, 10
    Button "npmin", 676, 76, 108, 28, "最小化", _
        "メモ帳をタスクバーへ格納します。最大化か、ミニマップの範囲選択で戻せます", 0, 10

    ' 要素情報はカードをやめ、ミニマップ直下の 2 行ストリップに畳む（v3）
    Fill 16, 116, 780, 40, C_PANEL
    Txt "", 28, 120, 102, 14, "ElementFromPoint", 10, C_KEY, False, xlLeft, True
    Txt "pb_pt_win", 138, 120, 540, 14, "―", 10, C_TEXT, False, xlLeft, True
    Txt "pb_pt_xy", 680, 120, 104, 14, "PT ―", 10, C_SUB, False, xlRight, True
    Txt "", 28, 136, 100, 14, "Rect ･ CanMove", 10, C_SUB, False, xlLeft, True
    Txt "pb_pt_rect", 136, 136, 396, 14, "―", 10, C_TEXT, False, xlLeft, True
    ' 操作の結果はここに出る。"メモ帳 1 を移動 ･ 960,0 960×540" のように長く、
    ' 120px の枠では最後まで出せなかった（実機で実測）。
    Txt "pb_drag", 540, 136, 244, 14, "ミニマップの範囲でメモ帳を置く", 10, C_SUB, False, xlRight, False

    Fill 16, 156, 780, 358, C_PANEL
    NameIt "pb_map", Cel(16, 156, 780, 358)
    m_mapSig = ""
End Sub

'------------------------------------------------------------------ 同期カード
' つなぐメモ帳は 1 枚だけ。入力面もこの 1 つで、ここへ直接書いた内容がその窓へ
' 流れ、その窓で打った文字がここへ返る。
Private Sub BuildValueCard()
    CardBox 14, 528, 392, 224
    Txt "", 28, 544, 92, 16, "ValuePattern", 11, C_KEY, True, xlLeft, True
    Txt "pb_focus", 124, 546, 176, 14, "focus ―", 10, C_SUB, False, xlLeft, True
    Button "sync", 304, 540, 88, 28, "同期を停止", _
        "1 秒ごとに本文を取得し、差分があれば SetValue で書き戻します（Ctrl+Shift+M）", 2

    Txt "pb_np_title", 28, 572, 268, 16, "―", 10, C_SUB, False, xlLeft, False
    Txt "pb_np_state", 300, 572, 92, 16, "", 10, C_KEY, False, xlRight, False
    ' 入力面は結合セルそのもの。内寸は PB_UNIT に依らない固定値にする
    ' （PB_UNIT 倍のインセットだと 1px / 2px だけ最終行が半分描かれて切れた）。
    Fill 28, 592, 364, 124, C_LINE2
    NameIt "pb_np_frame", Cel(28, 592, 364, 124)
    Fill 28 + PB_UNIT, 592 + PB_UNIT, 364 - PB_UNIT * 2, 124 - PB_UNIT * 2, C_BODY
    With Txt("pb_np_text", 40, 600, 340, 108, "", 11, C_TEXT, False, xlLeft, False)
        .ShrinkToFit = False
        .WrapText = True
        .VerticalAlignment = xlTop
        .Interior.Color = C_BODY
        ' 文字列書式。既定のままだと「1-2」が日付、「007」が 7 に化け、次の
        ' ティックが「セルが編集された」と誤読して、化けた文字列をメモ帳へ
        ' 書き戻す（利用者が打った文字が勝手に変わって見える）。
        .NumberFormat = "@"
    End With
    m_frameOn = C_LINE2

    Txt "pb_get", 28, 724, 132, 16, "← GetValue --:--:--", 10, C_SUB, False, xlLeft, True
    BuildDots
End Sub


Private Sub BuildDots()
    Dim i As Long
    For i = 0 To PB_DOTS - 1
        Fill 166 + i * 29, 726, 27, 8, C_LINE
    Next i
    NameIt "pb_dots", Cel(166, 726, PB_DOTS * 29 - 2, 8)
    m_dotShown = -1
End Sub

'------------------------------------------------------------------ ベンチ
' 円周率ベンチのカード。FE / BE / 表示用 Excel の 3 つの状態、Temp のレーン、
' 桁の盤、経過と桁数、実行ボタン 1 つ。
Private Sub BuildBenchCard()
    Dim i As Long
    Dim lx As Long
    Dim ly As Long

    CardBox 418, 528, 380, 224
    Txt "", 432, 540, 120, 20, "円周率ベンチ", 11, C_KEY, True, xlLeft, True
    ' 桁の色見本。盤のランプ 1 つが 1 桁で、桁の値がそのまま色になる。
    Txt "", 580, 541, 12, 13, "0", 9, C_SUB, False, xlLeft, True
    For i = 0 To 9
        Fill 596 + i * 6, 544, 6, 8, LampColor(i)
    Next i
    Txt "", 660, 541, 12, 13, "9", 9, C_SUB, False, xlLeft, True
    Txt "", 676, 541, 108, 13, "%TEMP%\pixelbridge\", 9, C_SUB, False, xlRight, True

    ProcBox "fe", 432, 564, 116, 32, "FE"
    ProcBox "be", 552, 564, 116, 32, "BE"
    ProcBox "disp", 672, 564, 112, 32, "表示"

    Txt "", 432, 604, 80, 13, "command.json", 9, C_SUB, False, xlLeft, True
    Fill 512, 610, 74, 2, C_LINE
    NameIt "pb_lane_out", Cel(512, 610, 74, 2)
    Txt "", 588, 604, 12, 13, "→", 9, C_SUB, False, xlLeft, True
    Txt "", 612, 604, 12, 13, "←", 9, C_SUB, False, xlLeft, True
    Fill 626, 610, 74, 2, C_LINE
    NameIt "pb_lane_in", Cel(626, 610, 74, 2)
    Txt "", 704, 604, 80, 13, "progress.json", 9, C_SUB, False, xlRight, True

    ' 桁の盤。67 x 12 = 804 ランプを 432,624 352x64 の枠の中に置く。
    ' 枠と札は行を分ける（結合セルは 1 セル塗られただけで全体が塗り潰される）。
    Fill 432, 624, 352, 64, C_PANEL
    lx = 432 + (352 - LAMP_COLS * LAMP_PX) \ 2
    ly = 624 + (64 - LAMP_ROWS * LAMP_PX) \ 2
    Fill lx, ly, LAMP_COLS * LAMP_PX, LAMP_ROWS * LAMP_PX, C_BODY
    NameIt "pb_lamps", Cel(lx, ly, LAMP_COLS * LAMP_PX, LAMP_ROWS * LAMP_PX)
    Txt "pb_jobstat", 432, 692, 352, 16, "待機", 10, C_SUB, False, xlLeft, True

    Button "runpi", 432, 712, 352, 28, "円周率計算", _
        "非表示の別プロセス Excel が 15 秒だけ円周率を計算し、その桁を表示用 Excel へ流します（Ctrl+Shift+P）", 1, 11
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
    Txt "", 14, 780, 40, 16, "ONKEY", 10, C_SUB, False, xlLeft, True
    ' 札の幅は中身の実寸から取る。狭いと「切」「算」が枠の外で切れる（実測）。
    Chip 56, 776, 124, "^+M 同期 入 / 切"
    Chip 184, 776, 116, "^+P 円周率計算"
    Txt "pb_meta1", 420, 780, 240, 16, "", 10, C_SUB, False, xlRight, True
    Txt "pb_meta2", 668, 780, 130, 16, "", 10, C_SUB, False, xlRight, True
End Sub

' 札の中の文字は、枠の内寸いっぱいの高さで置く。旧版は上下に PB_UNIT*2 ずつ
' 空けていたので、4px ビルドでは文字の入る高さが 8 デバイス px しかなく、
' 10px の和文が縦に切れ、「同期」の下半分が欠けて別の字に見えた（実機で実測）。
Private Sub Chip(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal s As String)
    Fill x, y, w, 24, C_PANEL
    Fill x + PB_UNIT, y + PB_UNIT, w - PB_UNIT * 2, 24 - PB_UNIT * 2, C_BODY
    Txt "", x + PB_UNIT * 2, y + PB_UNIT, w - PB_UNIT * 4, 24 - PB_UNIT * 2, _
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
        SyncNotepad
        PbLog "  t: sync"
    End If
    PaintMiniMap
    PbLog "  t: map"
    PumpChannel
    PbLog "  t: chan"
    PaintHeader
    PbLog "  t: head"
    PaintDots
    PbLog "  t: dots"
    PaintFooter
    ' フォーカス取得は最後（PaintFocusTail の説明を参照）
    PaintFocusTail
    PbLog "  t: painted dirty=" & m_dirty

    ' ティックは直書きなので Release は通常なにもしない（FE 清書が Hold を
    ' 残して転んだときの保険）。描いたティックを「描画回数」として数える。
    Release
    If m_dirty Then
        m_repaints = m_repaints + 1
        m_dirty = False
    End If
    PbLog "tick " & m_ticks & ": end"
    m_inTick = False
    PbArm
    Exit Sub
Failed:
    ' 1 ティックの失敗でポンプを止めない
    Release
    m_dirty = False
    PbLog "tick " & m_ticks & ": err " & Err.Number & " " & Err.Description
    m_inTick = False
    PbArm
End Sub

' 時計とポーリング回数は毎ティック進める。直書きなので、この毎秒の更新が
' 汚すのはヘッダの数字の矩形だけで、1px でも 1 秒に収まる（実測）。
Private Sub PaintHeader()
    Dim s As Double
    Dim mm As Long

    If m_syncOn Then
        SetLamp C_LAMP
    Else
        SetLamp C_OFFLAMP
    End If
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
    m_dirty = True
    rg.Interior.Color = c
End Sub

Private Sub PaintFooter()
    SetTxt "pb_meta1", "1 セル = " & PB_UNIT & "px ･ " & _
        IIf(PB_UNIT >= 4, "角丸なし", "角丸 " & (8 \ PB_UNIT) & " セル")
    SetTxt "pb_meta2", "進捗盤 = " & (LAMP_COLS * LAMP_PX \ PB_UNIT) & "×" & _
        (LAMP_ROWS * LAMP_PX \ PB_UNIT) & " セル"
End Sub

' ドット列は「同期が生きている」表示。毎ティック 1 つ進む。
Private Sub PaintDots()
    Dim rg As Range
    Dim i As Long
    Dim x As Long
    Dim y As Long

    If Not m_syncOn Then Exit Sub
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
        Case "npmax": NpMaximize
        Case "npmin": NpMinimize
        Case "runpi": StartBench
    End Select
End Sub

Private Sub ToggleSync()
    m_syncOn = Not m_syncOn
    SetTxt "pb_btn_sync", IIf(m_syncOn, "同期を停止", "同期を開始")
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

' 追う窓：この Excel と、つなぐメモ帳 1 枚。
'
' メモ帳の探索はデスクトップ直下の列挙だが、ClassName で絞り、しかも
' PB_SCANEVERY ティックに 1 度しか呼ばない。全列挙と条件付き列挙は速さがほぼ
' 同じ（118.1ms 対 118.8ms、実測）なので、得なのは速度ではなく「他人のプロバイダ
' に触る数が 22 から 2 に減る」こと。
'
' 掴む相手は 1 枚だけ。**いま掴んでいる窓が残っていれば必ずそれを使い続ける**。
' FindAll が返す並びは Z 順で、利用者のクリック 1 つで入れ替わる。並び順で選ぶと、
' 入れ替わった先へ前の窓の本文を書き戻して利用者の文章を消す（この構図が実害と
' して報告された）。掴んでいる窓が居なくなったときだけ選び直し、そのときは
' 見えている窓を優先する。
'
' 最小化された窓も手放さない。矩形は (-32000,-32000) へ飛ぶがハンドルは生きて
' いるので、掴んだままにしておけば「最大化」で戻せる。手放すと二度と戻せない。
Private Sub ScanWindows()
    Dim el As UIAutomationClient.IUIAutomationElement
    Dim keepEl As UIAutomationClient.IUIAutomationElement
    Dim arr As UIAutomationClient.IUIAutomationElementArray
    Dim cond As UIAutomationClient.IUIAutomationCondition
    Dim rc As UIAutomationClient.tagRECT
    Dim i As Long
    Dim n As Long
    Dim h As Variant
    Dim keepH As Variant

    On Error GoTo Failed
    If Not EnsureUia() Then Exit Sub
    rc = m_root.CurrentBoundingRectangle
    m_scrL = rc.Left: m_scrT = rc.Top: m_scrR = rc.Right: m_scrB = rc.Bottom

    ' 自分の窓の矩形は Application の座標（ポイント）から作る。UIA の
    ' ElementFromHandle(自分の hwnd) は使わない。自分の窓の UIA 要素を作ると
    ' Excel 自身のプロバイダがこのプロセスで動き出し、以後の書き込みが
    ' 1 回 130ms 級になる（MeasurePxPerPt の説明を参照）。
    m_winN = 0
    If m_pxPerPt > 0 Then
        rc.Left = CLng(Application.Left * m_pxPerPt)
        rc.Top = CLng(Application.Top * m_pxPerPt)
        rc.Right = CLng((Application.Left + Application.Width) * m_pxPerPt)
        rc.Bottom = CLng((Application.Top + Application.Height) * m_pxPerPt)
        AddWin rc, "この Excel", "EXCEL.EXE", PidText(), -1
    End If

    Set cond = m_uia.CreatePropertyCondition(UIA_ClassNamePropertyId, "Notepad")
    Set arr = m_root.FindAll(TS_CHILDREN, cond)
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
        m_npBound = False
        m_npMin = False
        Set m_npWin = Nothing
        Set m_npDoc = Nothing
        Set m_npVal = Nothing
        m_npTitle = ""
        m_uiaNote = "メモ帳 0 窓"
        Exit Sub
    End If

    BindNotepad keepEl, keepH
    rc = keepEl.CurrentBoundingRectangle
    m_npMin = (rc.Right <= rc.Left) Or (rc.Bottom <= rc.Top) Or (rc.Left < -10000) _
              Or CBool(keepEl.CurrentIsOffscreen)
    If Not m_npMin Then AddWin rc, m_npTitle, "notepad.exe", m_npPid, 0
    m_uiaNote = "メモ帳 " & n & " 窓"
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

Private Sub BindNotepad(ByVal win As UIAutomationClient.IUIAutomationElement, _
                        ByVal h As Variant)
    Dim d As UIAutomationClient.IUIAutomationElement
    Dim c1 As UIAutomationClient.IUIAutomationCondition
    Dim c2 As UIAutomationClient.IUIAutomationCondition

    On Error GoTo Failed
    Set m_npWin = win
    m_npTitle = win.CurrentName
    m_npPid = CStr(win.CurrentProcessId)
    If m_npBound Then
        If CStr(m_npHwnd) = CStr(h) Then Exit Sub      ' 同じ窓なら結び直さない
    End If
    Set c1 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_DocumentControlTypeId)
    Set c2 = m_uia.CreatePropertyCondition(UIA_ControlTypePropertyId, UIA_EditControlTypeId)
    Set d = win.FindFirst(TS_DESCENDANTS, m_uia.CreateOrCondition(c1, c2))
    If d Is Nothing Then
        m_npBound = False
        Exit Sub
    End If
    Set m_npVal = d.GetCurrentPattern(UIA_ValuePatternId)
    If m_npVal Is Nothing Then
        m_npBound = False
        Exit Sub
    End If
    Set m_npDoc = d
    m_npHwnd = h
    m_npBound = True
    ' 結んだ直後の 1 回は必ずメモ帳 → セルの向きで揃える。ここで書き戻しの
    ' 向きに入ると、画面側に残っていた古い文字で相手の本文を潰す。
    m_npFresh = True
    m_npLastText = ""
    m_npLastCell = ""
    Exit Sub
Failed:
    m_npBound = False
End Sub

'==============================================================================
' メモ帳との双方向フリー同期
'
' 入力面が変わっていれば SetValue で書き戻し、変わっていなければ GetValue だけ。
' 窓ごとに独立。判定も 1 文字制限も送信ボタンもない。
'==============================================================================

' フォーカスの見せ方。GetFocusedElement は使わない。フォーカスが Excel 自身に
' あるとき、それは「自分の窓の UIA 要素を作る」ことになり、Excel 自身の
' プロバイダがこのプロセスで動き出して、以後のセル書き込みが毎ティック
' 1 回 130ms 級になる（実測。フォーカス取得をティックの最後へ移しても、
' 汚染は次のティックの書き込みに残った）。代わりに、既に掴んでいるメモ帳の
' 文書要素（外部プロセスの要素）に HasKeyboardFocus を尋ねる。自分の窓には
' 一切触らないので税が無い。フォーカスがメモ帳以外にあるときは「―」。
Private Sub PaintFocusTail()
    Dim hf As Boolean
    If Not m_syncOn Then Exit Sub
    If m_npBound Then
        On Error Resume Next
        hf = CBool(m_npDoc.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId))
        On Error GoTo 0
    End If
    FrameColor IIf(hf, C_KEY, C_LINE2)
    If hf Then
        SetTxt "pb_focus", "focus pid " & m_npPid & " ･ Len " & Len(m_npLastText)
    Else
        SetTxt "pb_focus", "focus ―"
    End If
End Sub

Private Sub SyncNotepad()
    Dim cur As String
    Dim np As String
    Dim rg As Range

    On Error GoTo Failed
    If Not m_npBound Then
        SetTxt "pb_np_title", "―"
        SetTxt "pb_np_state", ""
        Exit Sub
    End If
    SetTxt "pb_np_title", Left$(m_npTitle, 40)
    SetTxt "pb_np_state", IIf(m_npMin, "最小化中", "")

    Set rg = m_ws.Range("pb_np_text")
    If rg Is Nothing Then Exit Sub

    ' 結んだ直後はメモ帳が正。セル側を合わせてから通常の同期に入る。
    If m_npFresh Then
        np = ToLf(CStr(m_npVal.CurrentValue))
        MirrorToCell rg, np
        m_npFresh = False
        SetTxt "pb_get", "← GetValue " & Format$(Now, "hh:nn:ss")
        Exit Sub
    End If

    cur = CStr(rg.Cells(1, 1).Value)
    If cur <> m_npLastCell Then
        ' Excel 側が編集された → メモ帳へ書き戻す
        SetTxt "pb_np_state", "SetValue 待ち"
        m_npVal.SetValue ToCrLf(cur)
        m_npLastCell = cur
        m_npLastText = ToLf(cur)
        SetTxt "pb_np_state", IIf(m_npMin, "最小化中", "")
        SetTxt "pb_get", "→ SetValue " & Format$(Now, "hh:nn:ss")
        Exit Sub
    End If

    np = ToLf(CStr(m_npVal.CurrentValue))
    If np <> m_npLastText Then
        MirrorToCell rg, np
        SetTxt "pb_get", "← GetValue " & Format$(Now, "hh:nn:ss")
    End If
    Exit Sub
Failed:
    m_npBound = False
    SetTxt "pb_np_state", "切断 " & Err.Number
End Sub

' メモ帳の本文をセルへ写し、控えを揃える。控えは 2 つの正規形で持つ。
'   - メモ帳側（m_npLastText）は LF 形。メモ帳の ValuePattern は CRLF を
'     LF に直した文字列を返すので（実測）、CRLF で控えると書いた直後に
'     「メモ帳が変わった」と誤読して 1 回よけいに写し直す。
'   - セル側（m_npLastCell）は「書いた値」ではなく「書いた結果セルが返す値」。
'     Excel は代入で先頭の ' を落とすなど値を黙って直すことがあり、意図した
'     値で控えると次のティックが「セルが編集された」と誤読して、直した後の
'     文字列をメモ帳へ書き戻してしまう。
Private Sub MirrorToCell(ByVal rg As Range, ByVal np As String)
    m_dirty = True
    rg.Cells(1, 1).Value = ToLf(np)
    m_npLastText = ToLf(np)
    m_npLastCell = CStr(rg.Cells(1, 1).Value)
End Sub

Private Function ToLf(ByVal s As String) As String
    ToLf = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function ToCrLf(ByVal s As String) As String
    ToCrLf = Replace(ToLf(s), vbLf, vbCrLf)
End Function

' 枠の帯だけ塗り替える。変わっていなければ何もしない。
Private Sub FrameColor(ByVal c As Long)
    Dim rg As Range
    Dim x As Long, y As Long, w As Long, h As Long
    If m_frameOn = c Then Exit Sub
    m_frameOn = c
    On Error Resume Next
    Set rg = m_ws.Range("pb_np_frame")
    If rg Is Nothing Then Exit Sub
    x = PxOfCol(rg.Column): y = PxOfRow(rg.Row)
    w = rg.Columns.Count * PB_UNIT: h = rg.Rows.Count * PB_UNIT
    Fill x, y, w, PB_UNIT, c
    Fill x, y + h - PB_UNIT, w, PB_UNIT, c
    Fill x, y, PB_UNIT, h, c
    Fill x + w - PB_UNIT, y, PB_UNIT, h, c
End Sub

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
    Dim cx1 As Long, cy1 As Long, cx2 As Long, cy2 As Long

    On Error Resume Next
    If m_winN = 0 Then Exit Sub
    ' 窓の並びが前と同じなら塗り替えない
    For i = 0 To m_winN - 1
        sig = sig & m_winL(i) & "," & m_winT(i) & "," & m_winR(i) & "," & m_winB(i) & ";"
    Next i
    sig = sig & "|" & m_winN
    If sig = m_mapSig Then Exit Sub
    m_mapSig = sig

    MapGeom mx, my, mw, mh, sc, ox, oy
    If sc <= 0 Then Exit Sub

    SetTxt "pb_res", (m_scrR - m_scrL) & "×" & (m_scrB - m_scrT) & " 自動判定"
    SetTxt "pb_scale", "×" & Format$(sc, "0.00") & " ･ " & CLng((m_scrR - m_scrL) * sc) & _
        "×" & CLng((m_scrB - m_scrT) * sc)

    ' ミニマップの中には結合セルを置かない。位置が変わるたびに結合し直す
    ' ことになり、その塗り替えの再描画だけでポンプが止まった（実測）。
    ' 窓の名前は下の対象バーに出す。塗りは直書き。個々の Fill は汚した矩形
    ' だけの再描画で済み、1 ティック内の連続塗りは 1 回の描き直しに畳まれる。
    Fill mx, my, mw, mh, C_PANEL
    cx1 = ox: cy1 = oy
    cx2 = ox + CLng((m_scrR - m_scrL) * sc)
    cy2 = oy + CLng((m_scrB - m_scrT) * sc)
    Fill cx1, cy1, cx2 - cx1, cy2 - cy1, C_BODY

    ' 窓の矩形は画面の外へ出ていることがある（メモ帳を画面の端に置く、Excel の
    ' 窓がタスクバーの下へ沈む、など）。それをそのまま描くと、ミニマップの枠を
    ' 越えて隣のカードの上まで塗ってしまう。実機で「青がミニマップの外へはみ出す」
    ' 「下のカードが灰色で塗り潰される」として見えていたのがこれ。画面の矩形
    ' （＝ミニマップの白い面）で必ず切り取る。
    For i = m_winN - 1 To 0 Step -1
        x1 = ox + CLng((m_winL(i) - m_scrL) * sc)
        y1 = oy + CLng((m_winT(i) - m_scrT) * sc)
        x2 = ox + CLng((m_winR(i) - m_scrL) * sc)
        y2 = oy + CLng((m_winB(i) - m_scrT) * sc)
        If x1 < cx1 Then x1 = cx1
        If y1 < cy1 Then y1 = cy1
        If x2 > cx2 Then x2 = cx2
        If y2 > cy2 Then y2 = cy2
        If x2 - x1 > 20 And y2 - y1 > 20 Then
            ' 操作できる窓（メモ帳）を強調する。Excel は固定なので添え物。
            If m_winNp(i) >= 0 Then
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


' ミニマップのセル → 画面座標を逆算して、その点にある窓を調べる。
' これは ElementFromPoint の実演で、動かす相手を選ぶ操作ではない。つなぐ
' メモ帳は 1 枚だけなので、動かす相手は常にそれ。
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
            FillPointPanel i
            Exit Sub
        End If
    Next i
    SetTxt "pb_pt_win", "その点に窓はありません"
    SetTxt "pb_pt_rect", "―"
End Sub

Private Sub FillPointPanel(ByVal idx As Long)
    Dim tp As UIAutomationClient.IUIAutomationTransformPattern
    Dim canMove As String
    On Error GoTo Failed
    SetTxt "pb_pt_win", Left$(m_winLabel(idx), 30) & " ･ " & _
        IIf(m_winNp(idx) < 0, "XLMAIN", "Notepad") & " ･ pid " & m_winPid(idx)
    canMove = "―"
    If m_winNp(idx) < 0 Then
        canMove = "固定（起動時に画面の左半分へ配置）"
    ElseIf Not m_npWin Is Nothing Then
        Set tp = m_npWin.GetCurrentPattern(UIA_TransformPatternId)
        If tp Is Nothing Then
            canMove = "False（TransformPattern なし）"
        Else
            canMove = "Move " & CStr(CBool(tp.CurrentCanMove)) & _
                      " / Resize " & CStr(CBool(tp.CurrentCanResize))
        End If
    End If
    SetTxt "pb_pt_rect", m_winL(idx) & ", " & m_winT(idx) & " ･ " & _
        (m_winR(idx) - m_winL(idx)) & "×" & (m_winB(idx) - m_winT(idx)) & " ･ " & canMove
    Exit Sub
Failed:
    SetTxt "pb_pt_rect", "―"
End Sub

' ドラッグで選ばれた矩形へ、つないでいるメモ帳を Move / Resize する。
'
' 動かすのは **メモ帳だけ**。この Excel は起動時に画面の左半分へ置いたきり、
' 位置も大きさも変えない（owner 指示）。最小化されていても掴んだ要素は
' 残してあるので、ここから戻せる（MoveWin が先に通常表示へ戻す）。
Private Sub PlaceAt(ByVal rg As Range)
    Dim mx As Long, my As Long, mw As Long, mh As Long
    Dim ox As Long, oy As Long
    Dim sc As Double
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long

    On Error GoTo Failed
    If Not m_npBound Then
        Note "メモ帳が見つかりません"
        Exit Sub
    End If
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
    If MoveWin(m_npWin, x1, y1, x2 - x1, y2 - y1) Then
        Note "メモ帳を移動 ･ " & x1 & "," & y1 & " " & (x2 - x1) & "×" & (y2 - y1)
        m_npMin = False
        m_mapSig = ""
        ScanWindows
    Else
        Note "動かせません（Transform が取れませんでした）"
    End If
    Exit Sub
Failed:
    Note "配置に失敗 " & Err.Number
End Sub

Private Sub Note(ByVal s As String)
    SetTxt "pb_drag", s
End Sub

'==============================================================================
' メモ帳の最大化 / 最小化（Win32 なし）
'
' 「最大化」は OS の全画面最大化ではない。FE Excel を画面の左半分に固定した
' まま、メモ帳を右半分いっぱいへ合わせる（owner 指示）。だから使うのは
' WindowPattern の Maximized ではなく TransformPattern の Move / Resize。
' 「最小化」は WindowPattern.SetWindowVisualState(Minimized)。掴んだ要素は
' 手放さないので、そのあと「最大化」でもミニマップの範囲選択でも戻せる。
'==============================================================================
Private Sub NpMaximize()
    Dim fx As Long, fw As Long
    On Error GoTo Failed
    If Not m_npBound Then
        Note "最大化：メモ帳が見つかりません"
        Exit Sub
    End If
    If m_scrR <= m_scrL Then ScanWindows
    fx = m_scrL + (m_scrR - m_scrL) \ 2
    fw = m_scrR - fx
    PbLog "npmax: " & fx & ",0 " & fw & "x" & (m_scrB - m_scrT)
    If MoveWin(m_npWin, fx, m_scrT, fw, m_scrB - m_scrT) Then
        Note "最大化 ･ " & fx & "," & m_scrT & " " & fw & "×" & (m_scrB - m_scrT)
        m_npMin = False
    Else
        Note "最大化できません（Transform が取れませんでした）"
    End If
    m_mapSig = ""
    ScanWindows
    Exit Sub
Failed:
    Note "最大化に失敗 " & Err.Number
End Sub

Private Sub NpMinimize()
    Dim wp As UIAutomationClient.IUIAutomationWindowPattern
    On Error GoTo Failed
    If Not m_npBound Then
        Note "最小化：メモ帳が見つかりません"
        Exit Sub
    End If
    Set wp = m_npWin.GetCurrentPattern(UIA_WindowPatternId)
    If wp Is Nothing Then
        Note "最小化できません（WindowPattern が取れませんでした）"
        Exit Sub
    End If
    wp.SetWindowVisualState WindowVisualState_Minimized
    m_npMin = True
    PbLog "npmin: minimized"
    Note "最小化 ･ タスクバーへ格納しました"
    m_mapSig = ""
    ScanWindows
    Exit Sub
Failed:
    Note "最小化に失敗 " & Err.Number
End Sub


' ほかの窓は UIA のパターンで動かす。最大化されていると CanMove が False に
' なるので、先に WindowPattern で通常表示へ戻す。
'
' パターンは必ず型どおりに受ける（IUIAutomationWindowPattern /
' IUIAutomationTransformPattern）。UIA のパターンは IUnknown 系で IDispatch を
' 持たないため、As Object（遅延バインディング）で受けると Set の時点で失敗し、
' 「整列ボタンを押してもメモ帳が動かない」になる（この実機で実測。同じ理由で
' ValuePattern も最初から早期バインディングで受けてある）。
'
' そして、それを直してもまだ動かなかった本当の原因は **パターン ID の
' 取り違え**だった。UIA_TransformPatternId は 10016 で、10003 は
' RangeValuePattern。10003 を渡すと GetCurrentPattern は例外も出さずに
' Nothing を返すので、ログには "no transform" とだけ出て、あたかも
' 「このメモ帳は Transform を公開していない」ように見える。実際にはこの端末の
' メモ帳 11.2606 は窓要素で 10016（Transform）と 10009（Window）を公開して
' いて、CanMove / CanResize とも True（2026-08-19 実測）。
Private Function MoveWin(ByVal el As UIAutomationClient.IUIAutomationElement, _
                         ByVal x As Long, ByVal y As Long, _
                         ByVal w As Long, ByVal h As Long) As Boolean
    Dim wp As UIAutomationClient.IUIAutomationWindowPattern
    Dim tp As UIAutomationClient.IUIAutomationTransformPattern

    On Error GoTo Failed
    If el Is Nothing Then
        PbLog "  movewin: el nothing"
        Exit Function
    End If
    Set wp = el.GetCurrentPattern(UIA_WindowPatternId)
    If Not wp Is Nothing Then
        On Error Resume Next
        wp.SetWindowVisualState WindowVisualState_Normal
        On Error GoTo Failed
    End If
    Set tp = el.GetCurrentPattern(UIA_TransformPatternId)
    If tp Is Nothing Then
        ' 取れなかったときは「相手が公開していない」のか「こちらの訊き方が
        ' 悪い」のかをログで分ける。IsTransformPatternAvailable が True なのに
        ' パターンが Nothing なら、それは訊き方（ID・受け型）の側の問題。
        PbLog "  movewin: no transform (avail=" & _
            CStr(el.GetCurrentPropertyValue(UIA_IsTransformPatternAvailablePropertyId)) & ")"
        Exit Function
    End If
    If Not CBool(tp.CurrentCanMove) Then
        PbLog "  movewin: canmove false"
        Exit Function
    End If
    tp.Move CDbl(x), CDbl(y)
    If CBool(tp.CurrentCanResize) Then tp.Resize CDbl(w), CDbl(h)
    PbLog "  movewin: ok " & x & "," & y & " " & w & "x" & h
    MoveWin = True
    Exit Function
Failed:
    PbLog "  movewin: err " & Err.Number & " " & Err.Description
    MoveWin = False
End Function

'==============================================================================
' ショートカット
'==============================================================================
Private Sub PbBindKeys()
    On Error Resume Next
    Application.OnKey "^+M", Qual("PbKeySync")
    Application.OnKey "^+P", Qual("PbKeyRunPi")
End Sub

Private Sub PbUnbindKeys()
    On Error Resume Next
    Application.OnKey "^+M"
    Application.OnKey "^+P"
End Sub

Public Sub PbKeySync()
    DoAction "sync"
    Release
End Sub

' 実行はボタンだけでなくキーからも。ボタンはセルなので、押すには盤面が
' 前に出ていないといけない。キーなら他の窓を見ていても始められる。
Public Sub PbKeyRunPi()
    DoAction "runpi"
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
'
' **別プロセスを残さないための後始末も、必ず走るこの入口でやる。** Auto_Close
' は必ず走るとは限らない（実測：ベンチ直後に閉じたとき、BeforeClose のログは
' 出たのに Auto_Close が来ないまま 45 秒待たされたことがある）。表示用 Excel を
' 閉じ、BE へ quit を送るところまでは、ここで済ませてしまう。
Public Sub PbPrepareClose()
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
    ' ベンチが回っている最中に閉じられても、後始末は同じ道を通す。
    If m_bench Then FinishBench
    CloseDisplayExcel
    If m_beStarted Then
        m_reqSeq = m_reqSeq + 1
        WriteAll CmdPath(), "{""cmd"":""quit"",""seq"":" & m_reqSeq & _
            ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    End If
End Sub

Public Sub PbShutdown()
    Dim k As Long
    On Error Resume Next
    PbLog "quit: begin"
    ' quit の送信と表示用 Excel の後始末は PbPrepareClose が済ませている
    ' （BeforeClose は必ず走るが、Auto_Close は走らないことがある。実測）。
    PbPrepareClose

    ' BE の返事を待つ。相手のプロセスは殺さない。
    If m_beStarted Then
        For k = 1 To 8
            If InStr(1, ReadAll(ProgPath()), """bye""") > 0 Then Exit For
            Application.Wait Now + TimeSerial(0, 0, 1)
        Next k
        PbLog "quit: be waited " & k
    End If
    KillQuiet CmdPath()
    KillQuiet PiPath()
    KillQuiet ProgPath()
    KillQuiet FePath()
    ' BE のコピーもこのフォルダの中にある。先に消さないとフォルダが空にならず、
    ' RmDir が失敗して空の pixelbridge フォルダが残る。
    KillQuiet BeBookPath()
    RmDirQuiet Left$(TempDir(), Len(TempDir()) - 1)

    ' 1px セルの盤面を前に出したまま表示設定を戻すと再配置が何度も走る。
    ' 先にふつうのシートへ移る。
    If Not OtherSheet() Is Nothing Then OtherSheet().Activate
    ActiveWindow.DisplayGridlines = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(PB_SHEET).Delete
    Release
    ThisWorkbook.Saved = True
    ' UI Automation で掴んだ相手（メモ帳の窓と ValuePattern）を手放す。
    ' これはプロセスをまたぐ COM 参照なので、握ったままだと Excel は終われない。
    ' 実測：後始末は 2 秒で終わっているのに、プロセスだけが 45 秒たっても残った。
    m_npBound = False
    Set m_npVal = Nothing
    Set m_npDoc = Nothing
    Set m_npWin = Nothing
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




Private Sub ClearLamps()
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range("pb_lamps")
    If rg Is Nothing Then Exit Sub
    Fill PxOfCol(rg.Column), PxOfRow(rg.Row), LAMP_COLS * LAMP_PX, LAMP_ROWS * LAMP_PX, C_BODY
    m_lampShown = 0
End Sub

'==============================================================================
' 円周率ベンチ（15 秒）
'
' 主題は 1 つだけ。**非表示の別プロセス Excel が重い計算をしている最中も、
' FE の 1 秒アニメ・メモ帳との双方向同期・メモ帳の移動 / リサイズが止まらない**
' ことを、一画面と実操作で見せる。
'
' 登場人物は 4 つ。
'   FE Excel     … この疑似ピクセルアプリ。1 秒ポンプで全部を回す。
'   BE Excel     … 別プロセス・非表示。15 秒だけ円周率を実際に計算する。
'   表示用 Excel … BE とは別の、保存しない一時ブック。計算できた桁が増えて
'                  いく様子をセルへ出す。書くのは FE で、BE は触らない。
'   メモ帳       … つないでいる 1 枚。
' やりとりは従来どおり Temp の JSON（command.json / progress.json）と、
' 桁そのものを置く pi.txt。
'
' 「計算できた桁」と「表示できた桁」は別に数える。正式な結果は後者。内部で
' 何桁進んでいても、表示用 Excel まで実際に届いていない桁は数えない。
'==============================================================================
Private Sub StartBench()
    Dim js As String

    If m_bench Then Exit Sub
    On Error GoTo Failed
    SetTxt "pb_jobstat", "用意しています ･ 表示用 Excel を作成中"
    ClearLamps
    m_piText = ""
    m_piCalc = 0
    m_piShown = 0
    m_benchMs = 0

    If Not m_beStarted Then PbEnsureBe
    If Not m_beStarted Then
        SetTxt "pb_jobstat", "BE を起こせませんでした ･ " & m_beNote
        Exit Sub
    End If
    KillQuiet PiPath()
    KillQuiet ProgPath()
    m_lastProg = ""

    If Not EnsureDisplayExcel() Then
        SetTxt "pb_jobstat", "表示用 Excel を作れませんでした"
        Exit Sub
    End If
    BenchLayout

    ' ここまでが「配置と準備」。15 秒はこの下から測る。
    m_benchT0 = Timer
    m_bench = True
    m_reqSeq = m_reqSeq + 1
    js = "{""cmd"":""pi"",""secs"":" & CLng(PB_BENCHSECS) & ",""seq"":" & m_reqSeq & _
         ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    WriteAll CmdPath(), js
    Flash "pb_lane_out"
    PbLog "bench: started"
    SetTxt "pb_be_state", "計算中"
    SetTxt "pb_jobstat", "計算中 ･ 経過 0.0 秒"
    Exit Sub
Failed:
    m_bench = False
    PbLog "bench: start failed " & Err.Number & " " & Err.Description
    SetTxt "pb_jobstat", "開始に失敗 " & Err.Number
    CloseDisplayExcel
End Sub

' ベンチ開始時の配置。FE は左半分のまま動かさない（起動時に置いたきり）。
' 表示用 Excel を右上 4 分の 1、メモ帳を右下 4 分の 1 へ。三つが重ならずに
' 画面を埋めるので、どれも隠れない。
Private Sub BenchLayout()
    Dim hx As Long, hy As Long, hw As Long, hh As Long

    On Error Resume Next
    If m_scrR <= m_scrL Then ScanWindows
    hx = m_scrL + (m_scrR - m_scrL) \ 2
    hw = m_scrR - hx
    hy = m_scrT
    hh = (m_scrB - m_scrT) \ 2

    If Not m_dispApp Is Nothing Then
        m_dispApp.WindowState = xlNormal
        m_dispApp.Left = hx / m_pxPerPt
        m_dispApp.Top = hy / m_pxPerPt
        m_dispApp.Width = hw / m_pxPerPt
        m_dispApp.Height = hh / m_pxPerPt
        m_dispApp.Visible = True
    End If
    If m_npBound Then
        MoveWin m_npWin, hx, hy + hh, hw, (m_scrB - m_scrT) - hh
        m_npMin = False
        ' 前面へ。UIA の SetFocus は相手のプロセスの窓に対して使う（自分の窓の
        ' 要素は作らない、という原則は守っている）。
        m_npWin.SetFocus
    End If
    m_mapSig = ""
    ScanWindows
    PbLog "bench: layout disp=" & hx & "," & hy & " " & hw & "x" & hh
End Sub

' 表示用 Excel。BE とは別プロセスで、保存しない一時ブックを 1 つ持つだけ。
' 起こし方は XLToolRack の JobHost（CreateObject → Visible / DisplayAlerts を
' 決めてから使う → 済んだら Quit して参照を捨てる）と同じ形にしてある。
' 違いは 2 つ：この相手はマクロを持たない新規ブックなので開くファイルが無く
' AutomationSecurity を触る必要が無いこと、そして **こちらは忙しくないので
' COM 参照を持ち続けてよい**こと（BE の参照を手放すのは、忙しい相手を COM で
' 叩くと FE がサーバービジーで止まるからで、表示用 Excel は FE が叩く側）。
Private Function EnsureDisplayExcel() As Boolean
    On Error GoTo Failed
    If Not m_dispApp Is Nothing Then
        ' 生きているか確かめる。落ちていれば作り直す。
        If m_dispApp.Workbooks.Count > 0 Then
            PrepareDisplaySheet
            EnsureDisplayExcel = True
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
    SetTxt "pb_disp_state", "作成"
    EnsureDisplayExcel = True
    Exit Function
Failed:
    PbLog "disp: FAILED " & Err.Number & " " & Err.Description
    SetTxt "pb_disp_state", "失敗 " & Err.Number
    CloseDisplayExcel
    EnsureDisplayExcel = False
End Function

Private Sub PrepareDisplaySheet()
    On Error Resume Next
    With m_dispSheet
        .Cells.Clear
        .Cells.Font.Name = m_fontMono
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
    ' 全セルに出る。これは使い捨ての別プロセスなので、その Excel の
    ' エラーチェックだけ切る（利用者の Excel の設定には触っていない）。
    m_dispApp.ErrorCheckingOptions.NumberAsText = False
    m_dispApp.ActiveWindow.DisplayGridlines = False
End Sub

' 保存せずに閉じる。ブックを閉じてから Quit、そのあと参照を捨てる。
' 逆順にすると、ブックを持たない Excel だけが残ることがある（BeQuit の教訓）。
Private Sub CloseDisplayExcel()
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
    SetTxt "pb_disp_state", "終了"
End Sub

' 毎ティック。BE の進捗を読み、まだ表示していない桁を表示用 Excel へ流す。
Private Sub PumpBench()
    Dim el As Double
    Dim js As String
    Dim st As String

    If Not m_bench Then Exit Sub
    On Error GoTo Failed
    el = Timer - m_benchT0
    If el < 0 Then el = el + 86400
    ' 15 秒を過ぎたティックでは、もう表示しない。正式な結果は「15 秒以内に
    ' 表示用 Excel まで届いた桁」なので、境界を越えてから足すと数えすぎになる。
    ' 1 秒ポンプなので、最後に数えられるのは 14 秒台のティックまで。
    If el >= PB_BENCHSECS Then
        FinishBench
        Exit Sub
    End If

    js = ReadAll(ProgPath())
    If Len(js) > 0 And js <> m_lastProg Then
        m_lastProg = js
        Flash "pb_lane_in"
        Unflash "pb_lane_out"
        st = JVal(js, "state")
        m_piCalc = CLng(Val(JVal(js, "n")))
        If st = "error" Then
            SetTxt "pb_jobstat", "BE エラー " & JVal(js, "msg")
            FinishBench
            Exit Sub
        End If
    Else
        Unflash "pb_lane_in"
    End If

    ShowDigits ReadDigits()

    SetTxt "pb_jobstat", "計算中 ･ 経過 " & Format$(el, "0.0") & " 秒 ･ 計算 " & _
        PiDecimals(m_piCalc) & " 桁 ･ 表示 " & m_piShown & " 桁"
    Exit Sub
Failed:
    PbLog "bench: pump err " & Err.Number & " " & Err.Description
    FinishBench
End Sub

' 15 秒で自動終了。BE を止め、表示用 Excel を保存せずに閉じ、Temp を片づける。
' 例外や途中終了からも同じ道を通す。
Private Sub FinishBench()
    Dim el As Double

    On Error Resume Next
    If Not m_bench Then Exit Sub
    m_bench = False
    el = Timer - m_benchT0
    If el < 0 Then el = el + 86400
    m_benchMs = el * 1000

    m_reqSeq = m_reqSeq + 1
    WriteAll CmdPath(), "{""cmd"":""quit"",""seq"":" & m_reqSeq & _
        ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    ' 返事（bye）をここで待つと FE が数秒固まる。待つのは次のティック以降。
    m_beStopping = True
    m_beStopTicks = 0
    CloseDisplayExcel
    KillQuiet PiPath()
    Unflash "pb_lane_in"
    Unflash "pb_lane_out"
    SetTxt "pb_be_state", "終了中"
    SetTxt "pb_jobstat", "完了 " & Format$(el, "0.00") & " 秒 ･ 計算 " & _
        PiDecimals(m_piCalc) & " 桁 ･ 表示 " & m_piShown & " 桁（結果）"
    PbLog "bench: finished " & Format$(el, "0.00") & "s calc=" & m_piCalc & _
        " shown=" & m_piShown
End Sub

' BE が置いた桁の列を読む。末尾の改行は WriteAll の Print # が足すので落とす。
Private Function ReadDigits() As String
    Dim s As String
    s = ReadAll(PiPath())
    Do While Len(s) > 0 And (Right$(s, 1) = vbLf Or Right$(s, 1) = vbCr)
        s = Left$(s, Len(s) - 1)
    Loop
    ReadDigits = s
End Function

' 先頭の 1 桁は整数部の 3。数えるのは小数点以下だけ。
Private Function PiDecimals(ByVal n As Long) As Long
    If n > 0 Then PiDecimals = n - 1
End Function

' まだ表示していない桁を、10 桁ずつ表示用 Excel のセルへ流す。表示できた
' 桁数（m_piShown）だけが結果になるので、ここは「実際に書けた分」しか進めない。
Private Sub ShowDigits(ByVal digits As String)
    Dim dec As String
    Dim have As Long
    Dim g As Long
    Dim groups As Long
    Dim r As Long
    Dim c As Long
    Dim piece As String

    On Error GoTo Failed
    If Len(digits) < 2 Then Exit Sub
    If m_dispSheet Is Nothing Then Exit Sub
    m_piText = digits
    dec = Mid$(digits, 2)                       ' 整数部の 3 を落とす
    have = Len(dec)
    groups = have \ PI_PERCELL                  ' 埋まりきった 10 桁だけ出す
    g = m_piShown \ PI_PERCELL
    Do While g < groups
        r = 4 + g \ PI_PERROW
        c = 1 + (g Mod PI_PERROW)
        piece = Mid$(dec, g * PI_PERCELL + 1, PI_PERCELL)
        m_dispSheet.Cells(r, c).Value = "'" & piece
        g = g + 1
        m_piShown = g * PI_PERCELL
    Loop
    PaintDigits dec, m_piShown
    Exit Sub
Failed:
    PbLog "bench: show err " & Err.Number & " " & Err.Description
End Sub

' 盤のランプは 1 つが 1 桁。桁の値がそのまま色になる（0 = 薄い、9 = 濃い）。
Private Sub PaintDigits(ByVal dec As String, ByVal shown As Long)
    Dim rg As Range
    Dim lx As Long
    Dim ly As Long
    Dim i As Long
    Dim want As Long

    On Error Resume Next
    Set rg = m_ws.Range("pb_lamps")
    If rg Is Nothing Then Exit Sub
    lx = PxOfCol(rg.Column)
    ly = PxOfRow(rg.Row)
    want = shown
    If want > LAMP_N Then want = LAMP_N
    For i = m_lampShown To want - 1
        Fill lx + (i Mod LAMP_COLS) * LAMP_PX, ly + (i \ LAMP_COLS) * LAMP_PX, _
             LAMP_PX, LAMP_PX, LampColor(Val(Mid$(dec, i + 1, 1)))
    Next i
    If want > m_lampShown Then m_lampShown = want
End Sub




Private Function PidText() As String
    If Len(m_myPid) = 0 Then
        PidText = "―"
    Else
        PidText = m_myPid
    End If
End Function

' 自分のプロセス番号。Win32 なしで取れる唯一の口が UIA なので、起動時に
' 1 回だけ自分の窓の要素を作って読み、すぐ手放す。持ち続けたり毎ティック
' 作り直したりすると、その間ずっと書き込みが 130ms 級になる（実測）。
Private Sub ReadOwnPid()
    Dim el As UIAutomationClient.IUIAutomationElement
    On Error Resume Next
    m_myPid = ""
    If Not EnsureUia() Then Exit Sub
    Set el = m_uia.ElementFromHandle(ByVal Application.hwnd)
    If el Is Nothing Then Exit Sub
    m_myPid = CStr(el.CurrentProcessId)
    Set el = Nothing
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

' 桁そのものを置くファイル。JSON に本文を入れないので、引用符も改行も
' 気にしなくていい（原稿を JSON へ入れていた頃の教訓をそのまま使う）。
Private Function PiPath() As String
    PiPath = TempDir() & "pi.txt"
End Function


' BE 用の軽いコピーは %TEMP%\pixelbridge\ に置く。配布物のフォルダには
' 何も作らない。配布するのは 1px / 2px / 4px の FE ブック 1 冊ずつだけで、
' BE 用の別冊は配らない。
'
' 場所を TEMP にできるのは、開くときに AutomationSecurity = 1 を使っている
' から。あれは「プログラムから開くときはマクロの警告を出さない」設定で、
' 置き場所の信頼設定とは別の口。XLToolRack も同じやり方で %TEMP% のコピーを
' 開いている（JobHost.cls の WorkerCopy）。
Private Function BeBookPath() As String
    BeBookPath = TempDir() & "pixelbridge_be.xlsm"
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
' BE 用のコピーも、やりとりの JSON も、みんな %TEMP%\pixelbridge\ の中。
' 配布物のフォルダには何も書かない。異常終了でコピーが残っていても、次に
' 起動したときここで消してから作り直すので、溜まっていくことはない。
' reuseCopy: 途中で BE を起こし直すとき True。起動時に作った BE 用コピーを
' そのまま使う。ここで SaveCopyAs し直すと、組み上がった 16 万セルの盤面ごと
' 複製することになり（94KB が 557KB になり、その読み込みで FE が詰まる。実測）、
' 起動時に「画面を作る前に BE を起こす」と決めた理由が無に帰す。
Private Sub PbEnsureBe(Optional ByVal reuseCopy As Boolean = False)
    Dim beApp As Object
    Dim beBook As Object
    Dim prevSec As Long
    Dim copyPath As String

    On Error GoTo Failed
    EnsureTempDir
    KillQuiet CmdPath()
    KillQuiet PiPath()
    KillQuiet ProgPath()
    WriteAll FePath(), "{""fe"":""alive"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"

    copyPath = BeBookPath()
    If Not (reuseCopy And Len(Dir$(copyPath)) > 0) Then
        KillQuiet copyPath
        ThisWorkbook.SaveCopyAs copyPath
    End If

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
    m_beNote = "非表示"
    ' FE は BE への COM 参照を持ち続けない。忙しい BE を COM で叩くと FE が
    ' サーバービジーで固まる。ここで手放す。
    Set beBook = Nothing
    Set beApp = Nothing
    m_beStarted = True
    SetTxt "pb_be_state", "待機 非表示"
    Exit Sub
Failed:
    m_beStarted = False
    m_beNote = "失敗 " & Err.Number
    PbLog "be: FAILED " & Err.Number & " " & Err.Description
    SetTxt "pb_be_state", m_beNote
End Sub

Private Sub Flash(ByVal nm As String)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If rg.Interior.Color = C_KEY Then Exit Sub
    m_dirty = True
    rg.Interior.Color = C_KEY
End Sub

Private Sub Unflash(ByVal nm As String)
    Dim rg As Range
    On Error Resume Next
    Set rg = m_ws.Range(nm)
    If rg Is Nothing Then Exit Sub
    If rg.Interior.Color = C_LINE Then Exit Sub
    m_dirty = True
    rg.Interior.Color = C_LINE
End Sub

' 毎ティックの Temp まわり。BE が居るあいだだけ「FE は生きている」印を置き、
' ベンチ中はそのポンプへ回す。BE へ quit を送ったあとは bye を待ち、見えたら
' やりとりに使った一時ファイルをフォルダごと片づける。
Private Sub PumpChannel()
    Dim js As String

    On Error Resume Next
    If m_beStarted Then
        WriteAll FePath(), "{""fe"":""alive"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    End If
    If m_bench Then
        PumpBench
        Exit Sub
    End If
    js = ReadAll(ProgPath())
    If Len(js) > 0 And js <> m_lastProg Then
        m_lastProg = js
        ' BE は FE が黙って見えると自分から終わる（置き去り防止）。それを見たら、
        ' 次のベンチで起こし直せるようにこちら側の印も倒す。
        If JVal(js, "state") = "bye" Then
            m_beStarted = False
            If Not m_beStopping Then SetTxt "pb_be_state", "自動終了"
        End If
    End If
    If m_beStopping Then
        m_beStopTicks = m_beStopTicks + 1
        ' bye が見えたか、8 ティック待っても来ないなら、そこで片づける。
        If Not m_beStarted Or m_beStopTicks > 8 Then
            m_beStopping = False
            m_beStarted = False
            BenchCleanupTemp
            SetTxt "pb_be_state", "終了"
        End If
    End If
End Sub

' ベンチのやりとりに使った一時ファイルを消す。フォルダも空なら消す。
' 例外や途中終了からもここを通す（FinishBench → 次のティック）。
Private Sub BenchCleanupTemp()
    On Error Resume Next
    KillQuiet PiPath()
    KillQuiet ProgPath()
    KillQuiet CmdPath()
    KillQuiet FePath()
    KillQuiet BeBookPath()
    RmDirQuiet Left$(TempDir(), Len(TempDir()) - 1)
    m_lastProg = ""
    PbLog "bench: temp cleaned"
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
                Case "pi"
                    BePiRun Val(JVal(js, "secs")), seq
                Case "pistop"
                    ' 走っていなければ何もしない（走行中は BePiRun が自分で拾う）
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

' BE 側の仕事。円周率を **実際に計算する**。既知の桁を流し込むことは 1 桁も
' しない。使うのは Rabinowitz-Wagon の spigot（有限桁版）で、桁を頭から 1 つずつ
' 吐くので「増えていく」様子がそのまま出せる。
'
' 配列の長さ ln は「何桁まで正しく出せるか」を決める。途中で打ち切っても、
' そこまでに吐いた桁は正しい（足りなくなるのではなく、まだ吐いていないだけ）。
' ln = 10n/3 + 1 で、1 桁あたり内側ループが ln 回まわるから総コストは約
' 10n^2/3。端末の速さは決め打ちできない（開発機で決めた回数が非力なノートで
' 14 倍かかった、という実測がこのリポジトリにある）ので、走る前に 0.15 秒だけ
' 同じ内側ループを空回しして速さを測り、15 秒で届く n を選ぶ。1.15 倍だけ
' 多めに取って、15 秒の打ち切りまで手が空かないようにする。
Private Sub BePiRun(ByVal secs As Double, ByVal stopSeq As Long)
    Dim a() As Long
    Dim n As Long, ln As Long
    Dim i As Long, j As Long, k As Long
    Dim q As Long, x As Long
    Dim predigit As Long, nines As Long
    Dim first As Boolean
    Dim out As String
    Dim t0 As Double, ms As Double
    Dim lastDv As Double
    Dim chunk As Long
    Dim rate As Double
    Dim js As String

    On Error GoTo Failed
    If secs <= 0 Then secs = 15
    rate = BeInnerRate()
    n = CLng(Sqr(3# * rate * secs / 10#) * 1.15)
    If n < 200 Then n = 200
    If n > 30000 Then n = 30000
    ln = (10 * n) \ 3 + 1
    ReDim a(1 To ln)
    For i = 1 To ln
        a(i) = 2
    Next i
    predigit = 0
    nines = 0
    first = True
    out = ""
    chunk = n \ 60
    If chunk < 5 Then chunk = 5
    t0 = Timer
    lastDv = t0
    PbBeWrite "{""state"":""running"",""n"":0,""target"":" & n & ",""rate"":" & _
        CLng(rate) & ",""ms"":0,""ts"":""" & Format$(Now, "hh:nn:ss") & """}"

    For j = 1 To n
        q = 0
        For i = ln To 2 Step -1
            x = 10 * a(i) + q * i
            a(i) = x Mod (2 * i - 1)
            q = x \ (2 * i - 1)
        Next i
        x = 10 * a(1) + q
        a(1) = x Mod 10
        q = x \ 10
        If q = 9 Then
            nines = nines + 1
        ElseIf q = 10 Then
            out = out & CStr(predigit + 1)
            For k = 1 To nines
                out = out & "0"
            Next k
            predigit = 0
            nines = 0
        Else
            If Not first Then out = out & CStr(predigit)
            predigit = q
            For k = 1 To nines
                out = out & "9"
            Next k
            nines = 0
            first = False
        End If

        ' 応答を保つ。ここを入れないと FE の 1 秒ティックにある UIA の列挙が
        ' この BE を待って返らず、「別プロセスへ逃がしたのに FE が固まる」に
        ' なる（実測。README の該当節を参照）。ただし DoEvents は 1 回 1.6ms
        ' かかるので毎桁は高すぎる（1 桁の計算より重くなる端末がある）。
        ' 時間で間隔を決めて 30ms に 1 回にすると、応答は保てて損は数 % で済む。
        ms = Timer - lastDv
        If ms < 0 Then ms = ms + 86400
        If ms >= 0.03 Then
            DoEvents
            lastDv = Timer
        End If

        If (j Mod chunk) = 0 Or j = n Then
            ms = (Timer - t0) * 1000
            If ms < 0 Then ms = ms + 86400000#
            WriteAll PiPath(), out
            PbBeWrite "{""state"":""running"",""n"":" & Len(out) & ",""target"":" & n & _
                ",""ms"":" & CLng(ms) & ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
            If ms >= secs * 1000 Then Exit For
            ' FE から新しい命令（pistop / quit）が来ていたら、そこで畳む。
            js = ReadAll(TempDir() & "command.json")
            If CLng(Val(JVal(js, "seq"))) > stopSeq Then Exit For
        End If
    Next j

    ms = (Timer - t0) * 1000
    If ms < 0 Then ms = ms + 86400000#
    WriteAll PiPath(), out
    PbBeWrite "{""state"":""done"",""n"":" & Len(out) & ",""target"":" & n & _
        ",""ms"":" & CLng(ms) & ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Exit Sub
Failed:
    PbBeWrite "{""state"":""error"",""msg"":""" & Err.Number & " " & _
        Replace(Err.Description, """", "'") & """,""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
End Sub

' 内側ループを 0.15 秒だけ空回しして、1 秒あたり何回まわるかを測る。
' 桁を吐かないので円周率そのものには 1 桁も関与しない。
Private Function BeInnerRate() As Double
    Dim a() As Long
    Dim ln As Long
    Dim i As Long
    Dim q As Long, x As Long
    Dim t0 As Double, el As Double
    Dim loops As Double

    ln = 2000
    ReDim a(1 To ln)
    For i = 1 To ln
        a(i) = 2
    Next i
    t0 = Timer
    Do
        q = 0
        For i = ln To 2 Step -1
            x = 10 * a(i) + q * i
            a(i) = x Mod (2 * i - 1)
            q = x \ (2 * i - 1)
        Next i
        loops = loops + ln
        el = Timer - t0
        If el < 0 Then el = el + 86400
    Loop While el < 0.15
    If el <= 0 Then el = 0.15
    BeInnerRate = loops / el
End Function

' 終わり方の順序が効く。ThisWorkbook.Close を先に呼ぶと、そこで自分のコードが
' 止まって Application.Quit に届かず、ブックを持たない不可視 Excel が残る（実測）。
Private Sub BeQuit()
    On Error Resume Next
    m_beQuit = True
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
