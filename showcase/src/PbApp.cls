VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbApp - アプリケーション本体（部品を持ち、1 秒の鎖を回す）
'
' 責務：起動、1 秒ごとのティック、押されたボタンの割り振り、そして終了。
' 個々の仕事は部品が持つ。ここが持つのは「順序」と「寿命」。
'
' ポンプの不変条件（壊すと閉じられなくなる）
'
'   **予約は常に 1 本。** m_armed が真なら二重に張らない。ティックは入口で
'   m_armed を False に戻し（この予約は消費された）、本体の最後にもう 1 本
'   だけ張る。止めたいときは張らないだけでいい。取り消しは保険であって、
'   頼りにする仕組みではない。
'
'   一度これを破って「ダイアログを閉じた直後にもう 1 本張る」を足したところ、
'   予約が同時に 2 本生きる状態ができ、閉じるときに取り消しきれなくなった。
'   残った予約を走らせるため Excel はブックを開いたまま保ち、後始末は終わって
'   いるのに窓だけが消えない（150 秒待っても消えなかった）。
'
'   **次を張るのは本体の最後。** 先頭で張ると「予約が常に 1 つ残っている」
'   状態になり、閉じようとしても Excel がその予約を待つ（閉じる要求から
'   Auto_Close が走り出すまで 35 秒かかった）。最後に張れば、終了時に立てた
'   m_running = False へ最後の 1 発が落ちて、そこで鎖が終わる。
'
' 解放の順序（変えない）
'
'   1. ベンチが回っていれば畳む（表示用 Excel を閉じ、BE へ quit）
'   2. 掴んだ UIA 要素を手放す（メモ帳の窓と ValuePattern）
'   3. UIA クライアントを手放す
'
'   どれもプロセスをまたぐ COM 参照で、握ったままだと Excel は終われない
'   （後始末は 2 秒で終わっているのにプロセスだけが 45 秒残った）。
'   **Class_Terminate には任せない。** 破棄の順序は GC の都合で決まるため。
'==============================================================================
Option Explicit

Private m_log As PbLog
Private m_err As PbError
Private m_canvas As PbCanvas
Private m_screen As PbScreen
Private m_ws As PbWorkspace
Private m_be As PbBeSession
Private m_run As PbBenchmarkRun

Private m_running As Boolean
Private m_armed As Boolean
Private m_nextTick As Date
Private m_inTick As Boolean
Private m_ticks As Long
Private m_polls As Long
Private m_syncOn As Boolean
Private m_startAt As Double
Private m_lastSel As String

'==============================================================================
' 起動
'==============================================================================
Public Sub Startup()
    On Error GoTo Failed
    Set m_log = New PbLog
    Set m_err = New PbError
    Set m_canvas = New PbCanvas
    Set m_screen = New PbScreen
    Set m_ws = New PbWorkspace
    Set m_be = New PbBeSession
    Set m_run = New PbBenchmarkRun

    m_canvas.Init m_log
    m_screen.Init m_canvas, m_log
    m_ws.Init m_log, m_canvas
    m_run.Init m_log, m_screen, m_be, m_ws, m_canvas
    m_be.Init m_run, m_log

    m_startAt = Timer
    m_ticks = 0
    m_polls = 0
    m_syncOn = True
    m_log.Note "show: begin"
    m_canvas.Hold

    ' BE は画面を作る前に起こす。作ったあとだと SaveCopyAs が盤面ごと複製し、
    ' その読み込みで FE が詰まる。
    m_be.Start
    m_screen.SetText "pb_be_state", m_be.StatusText

    ' 刻みを選ぶ前に、窓を最終の置き場所（画面の左半分）へ置いてしまう。
    ' UsableWidth / UsableHeight は「いまの窓」の広さなので、最大化したまま
    ' 測って刻みを選ぶと、左半分に収まらない盤面を作ってしまう。
    m_canvas.Build
    m_ws.ReadScreen m_canvas.PxPerPt
    m_canvas.PlaceLeftHalf m_ws.ScreenLeft, m_ws.ScreenTop, _
                           m_ws.ScreenRight, m_ws.ScreenBottom
    m_canvas.Calibrate
    m_ws.ReadOwnPid
    m_log.Note "show: sheet pitch=" & m_canvas.Pitch & " pid=" & m_ws.OwnPidText

    ' 盤面を前に出すのは、組み終わってから一度だけ。組む前に出すと 1px では
    ' 表示設定の変更で再配置が走り続けて戻ってこない（CPU 478 秒）。
    m_screen.BuildAll
    m_log.Note "show: built"
    m_canvas.ScrollHome
    m_canvas.CheckFit
    m_canvas.Unhold
    m_screen.SetText "pb_fe_state", "動作中"
    m_screen.SetText "pb_be_state", m_be.StatusText
    BindKeys
    m_running = True
    m_log.Note "show: shown"
    m_log.Flush
    Arm
    Exit Sub
Failed:
    m_err.Capture "Startup", "pitch=" & m_canvas.Pitch
    m_log.Note m_err.Text
    m_log.Flush
    m_canvas.Unhold
    MsgBox "起動に失敗しました。" & vbCrLf & m_err.Number & " " & m_err.Description, _
           vbExclamation, PB_APP
End Sub

'==============================================================================
' ポンプ
'==============================================================================
Private Sub Arm()
    On Error Resume Next
    If m_armed Then Exit Sub
    m_nextTick = Now + TimeSerial(0, 0, 1)
    Application.OnTime m_nextTick, Qual("PbTick")
    m_armed = True
End Sub

Private Sub Disarm()
    On Error Resume Next
    If m_armed Then
        ' 予約した時刻そのもので取り消す。Now + 1 秒 で呼んでいたときは一度も
        ' 当たらず、残った予約のせいでブックが閉じなかった。
        Application.OnTime m_nextTick, Qual("PbTick"), , False
    End If
    m_armed = False
End Sub

' 見張り。ブック自身のイベント（選択の変化・シート切り替え・窓の復帰）から
' 呼ばれ、鎖が切れていたら張り直す。切れたかどうかは「予約の時刻を過ぎても
' ティックが来ていないか」で見る。5 秒の猶予を置くのは、Excel がセル編集中や
' ダイアログ表示中に OnTime を後回しにするからで、その遅れを故障と取り違えない
' ため。
Public Sub EnsureArmed()
    On Error Resume Next
    If Not m_running Then Exit Sub
    If m_inTick Then Exit Sub
    If m_armed Then
        If Now <= DateAdd("s", 5, m_nextTick) Then Exit Sub
        m_log.Note "watchdog: schedule overdue, re-arming"
    Else
        m_log.Note "watchdog: chain lost, re-arming"
    End If
    m_armed = False
    Arm
End Sub

Public Sub Tick()
    Dim tAll As Double
    Dim t0 As Double

    On Error GoTo Failed
    If m_inTick Then Exit Sub
    If Not m_running Then Exit Sub
    m_armed = False              ' この予約は消費された
    m_inTick = True
    m_ticks = m_ticks + 1
    tAll = m_log.Mark

    t0 = m_log.Mark
    HandleSelection
    m_log.Span "selection", t0

    If m_syncOn Then
        m_polls = m_polls + 1
        If m_ws.Count = 0 Or (m_ticks Mod PB_SCANEVERY) = 1 Then
            t0 = m_log.Mark
            m_ws.ReadScreen m_canvas.PxPerPt
            m_ws.Rescan
            m_ws.RebuildWindows
            m_log.Span "uia scan", t0
        End If
        t0 = m_log.Mark
        SyncNotepad
        m_log.Span "uia sync", t0
    End If

    t0 = m_log.Mark
    m_screen.PaintMiniMap m_ws
    m_log.Span "minimap", t0

    t0 = m_log.Mark
    PumpChannel
    m_log.Span "channel", t0

    m_screen.PaintHeader m_syncOn, m_startAt, m_polls
    m_screen.PaintDots m_syncOn
    m_screen.PaintFooter
    ' フォーカスの取得は最後（PaintFocusTail の理由を参照）
    PaintFocusTail

    ' ティックは直書きなので Unhold は通常なにもしない（大きく描き換える
    ' 処理が Hold を残して転んだときの保険）。描いたティックを「描画回数」と
    ' して数える。
    m_canvas.Unhold
    If m_canvas.Dirty Then
        m_canvas.CountRepaint
        m_canvas.Dirty = False
    End If
    m_log.SpanAlways "tick " & m_ticks, tAll
    m_log.Flush
    m_inTick = False
    Arm
    Exit Sub
Failed:
    ' 1 ティックの失敗でポンプを止めない。記録して次を張る。
    m_err.Capture "Tick", "tick=" & m_ticks
    m_log.Note m_err.Text
    m_log.Flush
    m_canvas.Unhold
    m_canvas.Dirty = False
    m_inTick = False
    Arm
End Sub

'==============================================================================
' メモ帳との双方向同期
'
' 入力面が変わっていれば SetValue で書き戻し、変わっていなければ GetValue だけ。
' 判定も 1 文字制限も送信ボタンもない。
'==============================================================================
Private Sub SyncNotepad()
    Dim cur As String
    Dim np As String
    Dim rg As Range

    On Error GoTo Failed
    If Not m_ws.IsBound Then
        m_screen.SetText "pb_np_title", "―"
        m_screen.SetText "pb_np_state", ""
        Exit Sub
    End If
    m_screen.SetText "pb_np_title", Left$(m_ws.Title, 40)
    m_screen.SetText "pb_np_state", IIf(m_ws.IsMinimized, "最小化中", "")

    Set rg = m_canvas.RangeOf("pb_np_text")
    If rg Is Nothing Then Exit Sub

    ' 結んだ直後はメモ帳が正。セル側を合わせてから通常の同期に入る。
    If m_ws.IsFresh Then
        MirrorToCell rg, m_ws.ReadText()
        m_screen.SetText "pb_get", "← GetValue " & Format$(Now, "hh:nn:ss")
        Exit Sub
    End If

    cur = CStr(rg.Cells(1, 1).Value)
    If cur <> m_ws.LastCell Then
        ' 入力面が編集された → メモ帳へ書き戻す
        m_screen.SetText "pb_np_state", "SetValue 待ち"
        If m_ws.WriteText(cur) Then
            m_screen.SetText "pb_np_state", IIf(m_ws.IsMinimized, "最小化中", "")
            m_screen.SetText "pb_get", "→ SetValue " & Format$(Now, "hh:nn:ss")
        Else
            m_screen.SetText "pb_np_state", "切断"
        End If
        Exit Sub
    End If

    np = m_ws.ReadText()
    If np <> m_ws.LastText Then
        MirrorToCell rg, np
        m_screen.SetText "pb_get", "← GetValue " & Format$(Now, "hh:nn:ss")
    End If
    Exit Sub
Failed:
    m_ws.MarkDisconnected
    m_screen.SetText "pb_np_state", "切断 " & Err.Number
End Sub

Private Sub MirrorToCell(ByVal rg As Range, ByVal npText As String)
    m_canvas.Dirty = True
    rg.Cells(1, 1).Value = ToLf(npText)
    ' 控えは「書いた値」ではなく「書いた結果セルが返す値」で持つ（PbWorkspace
    ' の NoteMirrored を参照）。
    m_ws.NoteMirrored npText, CStr(rg.Cells(1, 1).Value)
End Sub

' フォーカスの見せ方。GetFocusedElement は使わない。フォーカスが Excel 自身に
' あるとき、それは「自分の窓の UIA 要素を作る」ことになり、以後のセル書き込みが
' 毎ティック 130ms 級になる。代わりに、既に掴んでいるメモ帳の文書要素
' （外部プロセスの要素）に尋ねる。フォーカスがメモ帳以外にあるときは「―」。
Private Sub PaintFocusTail()
    Dim hf As Boolean
    If Not m_syncOn Then Exit Sub
    hf = m_ws.HasFocus()
    m_screen.SetFrameFocused hf
    If hf Then
        m_screen.SetText "pb_focus", "focus pid " & m_ws.Pid & " ･ Len " & Len(m_ws.LastText)
    Else
        m_screen.SetText "pb_focus", "focus ―"
    End If
End Sub

' Temp まわり。ベンチ中はベンチのポンプへ、そうでなければ BE の様子見へ。
Private Sub PumpChannel()
    On Error Resume Next
    If m_run.IsRunning Then
        m_run.Heartbeat
        m_run.Pump
        Exit Sub
    End If
    m_be.PumpIdle
    m_screen.SetText "pb_be_state", m_be.StatusText
End Sub

'==============================================================================
' 押されたことは「選択セルが変わった」で知る
'
' 1 セルなら「そこを指した」、複数セルなら「その矩形へ置け」。ドラッグ選択は
' 離した瞬間に矩形になるので、これがそのままドラッグ配置の入口になる。
' 連続追従はできない（マウスイベントが取れないため）。
'==============================================================================
Private Sub HandleSelection()
    Dim rg As Range
    Dim a As String
    Dim key As String

    On Error Resume Next
    If ActiveSheet Is Nothing Then Exit Sub
    If ActiveSheet.Name <> PB_SHEET Then Exit Sub
    Set rg = Selection
    If rg Is Nothing Then Exit Sub
    a = rg.Address
    If a = m_lastSel Then Exit Sub
    m_lastSel = a

    ' ボタンは結合セルなので、押すと選択は「複数セル」になる。だからセル数で
    ' 分ける前に、まずボタンかどうかを見る。
    key = m_screen.ButtonKeyAt(rg.Cells(1, 1).Address)
    If Len(key) > 0 Then
        DoAction key
        Exit Sub
    End If

    If Not m_screen.InMap(rg.Cells(1, 1)) Then Exit Sub
    If rg.Cells.Count = 1 Then
        PickAt rg.Cells(1, 1)
    Else
        PlaceAt rg
    End If
End Sub

Public Sub DoAction(ByVal key As String)
    Select Case key
        Case "sync": ToggleSync
        Case "npmax": NpMaximize
        Case "npmin": NpMinimize
        Case "runpi": m_run.Start
    End Select
End Sub

Private Sub ToggleSync()
    m_syncOn = Not m_syncOn
    m_screen.SetText "pb_btn_sync", IIf(m_syncOn, "同期を停止", "同期を開始")
    m_screen.PaintHeader m_syncOn, m_startAt, m_polls
End Sub

Private Sub PickAt(ByVal c As Range)
    Dim px As Long, py As Long
    On Error Resume Next
    m_screen.CellToScreen c, px, py, m_ws
    m_screen.ShowPointXY px, py
    m_screen.ShowPointInfo m_ws.HitTest(px, py), m_ws
End Sub

' ドラッグで選ばれた矩形へ、つないでいるメモ帳を移動・リサイズする。
' 動かすのはメモ帳だけ。最小化されていても掴んだ要素は残してあるので、
' ここから戻せる（MoveResize が先に通常表示へ戻す）。
Private Sub PlaceAt(ByVal rg As Range)
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long

    On Error GoTo Failed
    If Not m_ws.IsBound Then
        m_screen.Note "メモ帳が見つかりません"
        Exit Sub
    End If
    If Not m_screen.RangeToScreen(rg, x1, y1, x2, y2, m_ws) Then Exit Sub
    If x2 - x1 < 160 Or y2 - y1 < 120 Then
        m_screen.Note "小さすぎます（160×120 未満）"
        Exit Sub
    End If
    If m_ws.MoveResize(x1, y1, x2 - x1, y2 - y1) Then
        m_screen.Note "メモ帳を移動 ･ " & x1 & "," & y1 & " " & (x2 - x1) & "×" & (y2 - y1)
        RefreshWindows
    Else
        m_screen.Note "動かせません（Transform が取れませんでした）"
    End If
    Exit Sub
Failed:
    m_err.Capture "PlaceAt", ""
    m_screen.Note "配置に失敗 " & m_err.Number
End Sub

' 「最大化」は OS の全画面最大化ではない。FE を画面の左半分に固定したまま、
' メモ帳を右半分いっぱいへ合わせる。だから使うのは WindowPattern の Maximized
' ではなく TransformPattern の Move / Resize。
Private Sub NpMaximize()
    Dim fx As Long, fw As Long
    On Error GoTo Failed
    If Not m_ws.IsBound Then
        m_screen.Note "最大化：メモ帳が見つかりません"
        Exit Sub
    End If
    If Not m_ws.HasScreen Then m_ws.ReadScreen m_canvas.PxPerPt
    fx = m_ws.ScreenLeft + m_ws.ScreenWidth \ 2
    fw = m_ws.ScreenRight - fx
    If m_ws.MoveResize(fx, m_ws.ScreenTop, fw, m_ws.ScreenHeight) Then
        m_screen.Note "最大化 ･ " & fx & "," & m_ws.ScreenTop & " " & fw & "×" & m_ws.ScreenHeight
    Else
        m_screen.Note "最大化できません（Transform が取れませんでした）"
    End If
    RefreshWindows
    Exit Sub
Failed:
    m_err.Capture "NpMaximize", ""
    m_screen.Note "最大化に失敗 " & m_err.Number
End Sub

Private Sub NpMinimize()
    On Error GoTo Failed
    If Not m_ws.IsBound Then
        m_screen.Note "最小化：メモ帳が見つかりません"
        Exit Sub
    End If
    If m_ws.Minimize() Then
        m_screen.Note "最小化 ･ タスクバーへ格納しました"
    Else
        m_screen.Note "最小化できません（WindowPattern が取れませんでした）"
    End If
    RefreshWindows
    Exit Sub
Failed:
    m_err.Capture "NpMinimize", ""
    m_screen.Note "最小化に失敗 " & m_err.Number
End Sub

Private Sub RefreshWindows()
    m_screen.InvalidateMap
    m_ws.Rescan
    m_ws.RebuildWindows
End Sub

'==============================================================================
' ショートカット
'==============================================================================
Private Sub BindKeys()
    On Error Resume Next
    Application.OnKey "^+M", Qual("PbKeySync")
    Application.OnKey "^+P", Qual("PbKeyRunPi")
End Sub

Private Sub UnbindKeys()
    On Error Resume Next
    Application.OnKey "^+M"
    Application.OnKey "^+P"
End Sub

Public Sub AfterKey()
    m_canvas.Unhold
End Sub

'==============================================================================
' 後始末
'==============================================================================
' 閉じる要求が来たその場で鎖を畳む。ThisWorkbook の BeforeClose から呼ばれる。
'
' **別プロセスを残さないための後始末も、必ず走るこの入口でやる。** Auto_Close は
' 必ず走るとは限らない（ベンチ直後に閉じたとき、BeforeClose のログは出たのに
' Auto_Close が来ないまま 45 秒待たされたことがある）。
Public Sub PrepareClose()
    On Error Resume Next
    m_running = False
    ' 「保存しますか」を出させない。この盤面は保存しない前提で毎回作り直す
    ' ものなので、変更として数えられる意味がない。立てずに閉じると Excel 自身の
    ' 確認ダイアログが出て、そこで閉じる処理が止まる。
    ThisWorkbook.Saved = True
    Disarm
    m_log.Note "quit: pump disarmed"
    UnbindKeys
    If m_run.IsRunning Then m_run.Finish
    m_run.CloseDisplay
    If m_be.IsStarted And Not m_be.IsStopping Then m_be.RequestQuit
    m_log.Flush
End Sub

Public Sub Shutdown()
    On Error Resume Next
    m_log.Note "quit: begin"
    ' quit の送信と表示用 Excel の後始末は PrepareClose が済ませている。
    PrepareClose

    m_be.WaitForBye 8
    m_run.CleanAll

    ' 1px の盤面を前に出したまま表示設定を戻すと再配置が何度も走る。
    ' 先にふつうのシートへ移る。
    If Not m_canvas.OtherSheet() Is Nothing Then m_canvas.OtherSheet().Activate
    m_canvas.RestoreChrome
    m_canvas.DeleteSheet
    m_canvas.Unhold
    ThisWorkbook.Saved = True

    ' 外の窓を掴んでいた COM 参照を手放す。順序（掴んだ要素 → UIA クライアント）
    ' は PbWorkspace の中で守られる。
    m_ws.ReleaseAll
    m_log.Note "quit: done"
    m_log.Flush
End Sub

'==============================================================================
' 外からの問い合わせ（検証が実画面の中身を確かめるのに使う）
'==============================================================================
Public Function TextOf(ByVal nm As String) As String
    On Error Resume Next
    TextOf = m_canvas.TextOf(nm)
End Function
