Attribute VB_Name = "modPixelBridge"
'==============================================================================
' VBA Pixel Bridge - Excel 分離アーキテクチャ 技術ショーケース
'
' modPixelBridge - 公開の入口だけ
'
' 責務：Excel から名前で呼ばれるものを受け、アプリ本体（PbApp）へ渡す。
' 判断はここに書かない。ここが薄いままであることが、この構成の目印。
'
' 呼ばれ方は 3 通り。
'   1. Auto_Open / Auto_Close      … ブックを開いた / 閉じたとき
'   2. Application.OnTime / OnKey  … 1 秒のポンプとショートカット
'   3. ThisWorkbook のイベント     … 鎖が切れていないかの見張り（ビルド時に注入）
'
' 参照設定は UIAutomationClient だけ。Win32 API / Declare / Shell / PowerShell /
' cmd / WMI / 外部 helper / C# は製品の実行経路に一切持ち込まない。別 Excel の
' 起動は Excel COM のみ。図形もフォームコントロールも ActiveX も使わない。
'
' 構成（責務ごとに分けてある）
'   modPbDesign    寸法・色・状態を持たない座標計算
'   modPbCommon    UIA の識別子・文字列ユーティリティ
'   modPbBackend   BE 側の入口と円周率の計算（コピーされたブックの中で動く）
'   PbApp          本体。ポンプ、操作の割り振り、寿命
'   PbCanvas       疑似ピクセルの画布（シートと描く道具）
'   PbScreen       盤面の配置と差分描画
'   PbUia          UI Automation クライアントと画面
'   PbNotepad      つないでいるメモ帳 1 枚
'   PbWindows      ミニマップに映す窓の一覧
'   PbChannel      Temp のファイルだけで往復する FE ⇔ BE
'   PbBeSession    非表示の別プロセス Excel の一生
'   PbDisplayBook  表示用 Excel（保存しない一時ブック）
'   PbBench        15 秒の円周率ベンチ
'   PbLog          記録（イミディエイト / ファイル）
'   PbError        例外の文脈
'==============================================================================
Option Explicit

Private g_app As PbApp

' アプリ本体。最初に呼ばれたときに作る。
Private Function App() As PbApp
    If g_app Is Nothing Then Set g_app = New PbApp
    Set App = g_app
End Function

'------------------------------------------------------------------ 起動と終了
' Auto_Open は標準モジュールで動く唯一の起動フック。利用者がブックを開けば
' FE が立ち上がる。FE が Workbooks.Open で開いた BE 用コピーでは Auto_Open は
' 走らないので、FE 側から入口を名前で呼ぶ。その BE では自分が BE だと気づいて
' 別の道を通る。
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

' BE 用のコピーはファイル名で見分ける。
Private Function IsBeBook() As Boolean
    IsBeBook = (InStr(1, ThisWorkbook.Name, PB_BE_MARK & ".", vbTextCompare) > 0)
End Function

Public Sub PbShow()
    App().Startup
End Sub

Public Sub PbPrepareClose()
    If g_app Is Nothing Then Exit Sub
    g_app.PrepareClose
End Sub

Public Sub PbShutdown()
    If g_app Is Nothing Then Exit Sub
    g_app.Shutdown
    Set g_app = Nothing
End Sub

'------------------------------------------------------------------ ポンプ
Public Sub PbTick()
    If g_app Is Nothing Then Exit Sub
    g_app.Tick
End Sub

Public Sub PbEnsureArmed()
    If g_app Is Nothing Then Exit Sub
    g_app.EnsureArmed
End Sub

'------------------------------------------------------------------ 操作
Public Sub DoAction(ByVal key As String)
    App().DoAction key
End Sub

Public Sub PbKeySync()
    App().DoAction "sync"
    App().AfterKey
End Sub

' 実行はボタンだけでなくキーからも。ボタンはセルなので、押すには盤面が前に
' 出ていないといけない。キーなら他の窓を見ていても始められる。
Public Sub PbKeyRunPi()
    App().DoAction "runpi"
    App().AfterKey
End Sub

'------------------------------------------------------------------ 外からの確認
' ビルドはこれを呼んでコンパイルを通す。
Public Function PbPing() As String
    PbPing = PB_APP & " ok " & Application.Version
End Function

' 画面のどこかの文字を名前で読む。外からの検証がこれで実画面の中身を確かめる。
Public Function PbGet(ByVal nm As String) As String
    If g_app Is Nothing Then Exit Function
    PbGet = g_app.TextOf(nm)
End Function
