Attribute VB_Name = "modZipBench"
'==============================================================================
' modZipBench -- 司令塔。準備・実行・記録・取消・後始末。
'
' 正本の置き方
' ------------------------------------------------------------------
' データブック 1 冊の 3 シートが全方式の正本。実行時にファイルは一切通らない。
'   MASTER  1列目=郵便番号, 2列目=住所   (KEN_ALL から準備時に作る)
'   INPUT   1列目=郵便番号               (決定的に生成)
'   OUTPUT  1列目=住所                   (各方式が書く先)
' どのシートも 1 行目からデータ。見出し行は置かない。
'
' 計測の約束 (全方式で同一)
' ------------------------------------------------------------------
'   起動秒     押してからエンジンが指示を受けられるまで。処理E2Eとは別枠。
'   Master読込 / Input読込 / 辞書構築 / 変換 / 出力書込 / 通知
'   処理E2E    仕事を投げてから、前面ブックで完了を検知するまで
'   照合秒     計測の外。OUTPUT を読んで正解と突き合わせる
'
' 読み書きの原則: Input と Master は各 1 回、Output は 1 回の一括 Value2。
' 固定待機・ポーリング・中間CSV・分割書込・処理中の画面更新は入れない。
' 完了は「通知セルを 1 回更新 → Worksheet_Change」で全方式共通に受ける。
'
' 時刻は QueryPerformanceCounter。VBA の Timer は Single で 1ULP が約 3.9ms あり、
' 工程内訳には粗すぎる。
'
' 失敗した方式は失敗のまま記録する。別方式へ切り替えない。
'==============================================================================
Option Explicit

Public Const SHEET_BENCH  As String = "BENCH"
Public Const SHEET_LOG    As String = "LOG"
Public Const SHEET_SIGNAL As String = "SIGNAL"
Public Const SHEET_RUNS   As String = "RUNS"
Public Const SIGNAL_CELL  As String = "A1"

Public Const DATA_MASTER  As String = "MASTER"
Public Const DATA_INPUT   As String = "INPUT"
Public Const DATA_OUTPUT  As String = "OUTPUT"

Public Const FIRST_ROW    As Long = 1        ' 3 シートともデータは 1 行目から

'--- 1 方式 1 回ぶんの計測結果 -------------------------------------------------
Public Type ZbResult
    MethodNo    As Long
    MethodName  As String
    WorkerKind  As String
    Created     As String
    Launched    As String
    UiaOk       As String
    Converted   As String
    ErrNumber   As Long

    LaunchSec   As Double      ' 起動 (処理E2Eとは別枠)
    MasterSec   As Double      ' Master 読込
    InputSec    As Double      ' Input 読込
    DictSec     As Double      ' 辞書構築
    ConvertSec  As Double      ' 変換 + 出力配列作成
    WriteSec    As Double      ' 出力生成・転送 (COM一括書込 / TSV書出+atomic rename)
    ImportSec   As Double      ' Excel反映 (QueryTable取込。COM経路は書込と同一操作なので0)
    NotifySec   As Double      ' 通知 (+ 未分解の往復)
    TotalSec    As Double      ' 処理E2E
    VerifySec   As Double      ' 照合 (計測の外)

    Rows        As Long
    MatchText   As String
    MismatchRow As Long
    Hash        As String
    Outcome     As String
    Note        As String
End Type

'--- 準備で作られ、全方式が共有するもの ---------------------------------------
Public g_N        As Long        ' 件数
Public g_MasterN  As Long        ' MASTER の行数
Public g_Input    As Variant     ' 1 起点。入力
Public g_Expected As Variant     ' 1 起点。正解
Public g_Prepared As Boolean
Public g_Cancel   As Boolean
Public g_DataWb   As Workbook

'--- 完了通知 -----------------------------------------------------------------
' 各方式は最後に通知セルを 1 回だけ更新し、その Worksheet_Change でここが
' 呼ばれる。処理E2E の終了時刻はハンドラの中で取るので、待ち側の粗さは入らない。
Public g_SigArmed   As Boolean
Public g_SigFired   As Boolean
Public g_SigStamp   As Currency
Public g_SigPayload As String

'--- 計測前後で退避する Excel の状態 ------------------------------------------
Private m_ScreenUpdating As Boolean
Private m_EnableEvents   As Boolean
Private m_Calculation    As Long
Private m_InMeasure      As Boolean

'==============================================================================
' パス
'==============================================================================
Public Function ZB_Root() As String
    If Len(ThisWorkbook.Path) > 0 Then ZB_Root = ThisWorkbook.Path Else ZB_Root = CurDir$
End Function

' 1 行に詰めた手続き定義は書かないこと。
' 「End Sub、End Function または End Property 以降には、コメントのみが
'  記述できます」でプロジェクト全体がコンパイルできなくなる。
Public Function ZB_Data() As String
    ZB_Data = ZB_Root() & "\data"
End Function

Public Function ZB_Work() As String
    ZB_Work = ZB_Root() & "\work"
End Function

Public Function ZB_KenAllPath() As String
    ZB_KenAllPath = ZB_Data() & "\KEN_ALL.CSV"
End Function

Public Function ZB_DictPath() As String
    ZB_DictPath = ZB_Data() & "\zip_dict.csv"
End Function

'==============================================================================
' シート
'==============================================================================
Private Function SheetByName(ByVal nm As String, ByVal atEnd As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        If atEnd Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        End If
        ws.Name = nm
    End If
    Set SheetByName = ws
End Function

Public Function ZB_Sheet() As Worksheet
    Set ZB_Sheet = SheetByName(SHEET_BENCH, False)
End Function

Public Function ZB_LogSheet() As Worksheet
    Set ZB_LogSheet = SheetByName(SHEET_LOG, True)
End Function

Public Function ZB_SignalSheet() As Worksheet
    Set ZB_SignalSheet = SheetByName(SHEET_SIGNAL, True)
End Function

Public Function ZB_RunsSheet() As Worksheet
    Set ZB_RunsSheet = SheetByName(SHEET_RUNS, True)
End Function

Public Function ZB_RowOf(ByVal methodNo As Long) As Long
    ZB_RowOf = 11 + methodNo
End Function

Public Sub ZB_Log(ByVal msg As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ZB_LogSheet()
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).NumberFormat = "@"
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 1).Value = Format$(Now, "hh:nn:ss")
    ws.Cells(r, 2).Value = msg
    Debug.Print Format$(Now, "hh:nn:ss") & "  " & msg
End Sub

'==============================================================================
' 完了通知
'==============================================================================
Public Sub ZbSignalFired(ByVal payload As String)
    If Not g_SigArmed Then Exit Sub
    If Len(payload) = 0 Then Exit Sub          ' 空にしただけの変更は完了ではない
    g_SigStamp = modZipRule.QpcNow()
    g_SigPayload = payload
    g_SigFired = True
End Sub

' 通知セルを空にしてから武装する。順序が逆だと、その「空にする」こと自体が
' Worksheet_Change を起こし、走り出す前に完了したことになる (実測済み)。
Public Sub ZbArmSignal()
    Dim prev As Boolean
    On Error Resume Next
    g_SigArmed = False
    prev = Application.EnableEvents
    Application.EnableEvents = False
    ZB_SignalSheet().Range(SIGNAL_CELL).ClearContents
    Application.EnableEvents = prev
    g_SigFired = False
    g_SigStamp = 0
    g_SigPayload = ""
    g_SigArmed = True
    Err.Clear
End Sub

Public Sub ZbDisarmSignal()
    g_SigArmed = False
End Sub

' 前面プロセスの中で完結する方式 (1/2/7) が、自分で通知セルを更新するための入口。
' 完了検知の仕組みを全方式でそろえるために通す。
Public Sub ZbFireLocalSignal(ByVal payload As String)
    Dim prev As Boolean
    On Error Resume Next
    prev = Application.EnableEvents
    Application.EnableEvents = True
    ZB_SignalSheet().Range(SIGNAL_CELL).Value = payload
    Application.EnableEvents = prev
    Err.Clear
End Sub

'==============================================================================
' 画面の初期化
'==============================================================================
Public Sub ZbInitSheet()
    Dim ws As Worksheet
    Dim i As Long
    Dim hdr As Variant

    Set ws = ZB_Sheet()
    Application.ScreenUpdating = False

    ws.Cells.Clear
    ws.Range("A1").Value = "ZipBench -- 郵便番号100万件を住所へ変換する 起動動作テストと速度ベンチマーク"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 13

    ws.Range("A3").Value = "件数"
    ws.Range("B3").Value = 2000
    ws.Range("C3").Value = "<- 小件数で全方式の一致を確認してから 1000000 にする"
    ws.Range("A4").Value = "ワーカー版"
    ws.Range("B4").Value = "prebuilt"
    ws.Range("C4").Value = "<- prebuilt = あらかじめ用意した exe / emitted = Excel が BAT・PS1・C#ソースを作って建てる"
    ws.Range("A5").Value = "ワーカー表示"
    ws.Range("B5").Value = "offscreen"
    ws.Range("C5").Value = "<- offscreen / minimized / normal / hidden (hidden は UIA ツリーに出ないので失敗するのが正しい)"
    ws.Range("A6").Value = "準備状態"
    ws.Range("B6").Value = "未"
    ws.Range("A7").Value = "元データ"
    ws.Range("B7").Value = "日本郵便 KEN_ALL.CSV"
    ws.Range("A8").Value = "方式3の案内"
    ws.Range("B8").Value = "表示"
    ws.Range("C8").Value = "<- 表示 / 非表示。無人で流すときだけ 非表示 にする"
    ws.Range("A3:A8").Font.Bold = True

    hdr = Array("#", "方式", "ワーカー版", "作成", "起動", "UIA操作", "変換", "Err番号", _
                "起動秒", "Master読込", "Input読込", "辞書構築", "変換", "出力生成/転送", "Excel反映", _
                "通知", "処理E2E", "照合秒", "回数", "E2E最小", "E2E最大", _
                "行数", "結果一致", "不一致行", "出力ハッシュ", "完了結果", "メモ")
    For i = 0 To UBound(hdr)
        ws.Cells(11, i + 1).Value = hdr(i)
    Next i
    With ws.Range(ws.Cells(11, 1), ws.Cells(11, UBound(hdr) + 1))
        .Font.Bold = True
        .Interior.Color = RGB(230, 230, 230)
        .WrapText = True
    End With

    For i = 1 To 10
        ws.Cells(ZB_RowOf(i), 1).Value = i
        ws.Cells(ZB_RowOf(i), 2).Value = ZB_MethodName(i)
    Next i
    For i = 13 To 16
        ws.Cells(ZB_RowOf(i), 1).Value = i
        ws.Cells(ZB_RowOf(i), 2).Value = ZB_MethodName(i)
    Next i
    ws.Cells(ZB_RowOf(11), 1).Value = "F1"
    ws.Cells(ZB_RowOf(11), 2).Value = ZB_MethodName(11)
    ws.Cells(ZB_RowOf(12), 1).Value = "F2"
    ws.Cells(ZB_RowOf(12), 2).Value = ZB_MethodName(12)

    ws.Range("A29").Value = "注記"
    ws.Range("A29").Font.Bold = True
    ws.Range("A30").Value = "・正本は データブックの MASTER / INPUT / OUTPUT の3シート。実行時にファイルは通らない。"
    ws.Range("A31").Value = "・Input と Master は各1回読み、分割書込・中間CSV・固定待機・ポーリングは無し。"
    ws.Range("A32").Value = "・起動秒は処理E2Eとは別枠。処理E2E = 仕事を投げてから通知セルの Worksheet_Change を受けるまで。"
    ws.Range("A33").Value = "・通知秒は「処理E2E - 各工程の合計」。通知セル書込のほか、往復やスレッド起動を含む未分解の残り。"
    ws.Range("A34").Value = "・表の秒は保存済み全実行の中央値。各回の生値は RUNS シートにある。"
    ws.Range("A35").Value = "・時刻は QueryPerformanceCounter。C#側は Stopwatch。計測中は画面を一切更新しない。"
    ws.Range("A36").Value = "・方式1 は1セルずつ読んで1セルずつ書くので工程を分けられない。変換秒に合算する。"
    ws.Range("A37").Value = "・方式9/13/14/15 はワーカーが書いた同一ファイルを読む。違うのはExcel側の取込経路だけ。"
    ws.Range("A38").Value = "・失敗した方式は失敗のまま残す。別方式へ切り替えない。"

    ws.Columns("A").ColumnWidth = 5
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 10
    ws.Columns("D:G").ColumnWidth = 7
    ws.Columns("H").ColumnWidth = 8
    ws.Columns("I:Q").ColumnWidth = 9
    ws.Columns("R").ColumnWidth = 6
    ws.Columns("S:T").ColumnWidth = 9
    ws.Columns("U").ColumnWidth = 11
    ws.Columns("V").ColumnWidth = 9
    ws.Columns("W").ColumnWidth = 8
    ws.Columns("X").ColumnWidth = 11
    ws.Columns("Y").ColumnWidth = 20
    ws.Columns("Z").ColumnWidth = 70
    ws.Range("I12:R23").NumberFormat = "0.000"
    ws.Range("T12:U23").NumberFormat = "0.000"
    ws.Range("V12:V23").NumberFormat = "#,##0"

    With ZB_LogSheet()
        .Cells.Clear
        .Range("A1").Value = "時刻": .Range("B1").Value = "ログ"
        .Range("A1:B1").Font.Bold = True
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 150
    End With

    InitRunsSheet
    With ZB_SignalSheet()
        .Range("A3").Value = "各方式はここへ完了を1回書く。その Worksheet_Change が完了検知そのもの。"
        .Columns("A").ColumnWidth = 90
    End With

    ws.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub InitRunsSheet()
    Dim hdr As Variant, i As Long
    hdr = Array("時刻", "#", "方式", "ワーカー版", "件数", "回", _
                "起動秒", "Master読込", "Input読込", "辞書構築", "変換", "出力生成/転送", "Excel反映", _
                "通知", "処理E2E", "照合秒", "行数", "結果一致", "ハッシュ")
    With ZB_RunsSheet()
        .Cells.Clear
        For i = 0 To UBound(hdr)
            .Cells(1, i + 1).Value = hdr(i)
        Next i
        .Range(.Cells(1, 1), .Cells(1, UBound(hdr) + 1)).Font.Bold = True
        .Columns("A").ColumnWidth = 9
        .Columns("C").ColumnWidth = 30
        .Columns("G:O").ColumnWidth = 10
    End With
End Sub

Public Function ZB_MethodName(ByVal n As Long) As String
    Select Case n
        Case 1:  ZB_MethodName = "1 低速VBA (1セルずつ)"
        Case 2:  ZB_MethodName = "2 高速VBA (基準)"
        Case 3:  ZB_MethodName = "3 BAT手動起動 -> C#ワーカー"
        Case 4:  ZB_MethodName = "4 WScript.Shell.Run 起動"
        Case 5:  ZB_MethodName = "5 VBA Shell 起動"
        Case 6:  ZB_MethodName = "6 Task Scheduler COM 起動"
        Case 7:  ZB_MethodName = "7 不可視の別Excelプロセス"
        Case 8:  ZB_MethodName = "8 Shell.Application.ShellExecute 起動"
        Case 9:  ZB_MethodName = "9 候補2 TSV+rename -> QueryTable取込"
        Case 10: ZB_MethodName = "10 候補3 COM Range.Value2 一括"
        Case 11: ZB_MethodName = "ファイル作成だけ (prebuilt)"
        Case 12: ZB_MethodName = "ファイル作成だけ (emitted)"
        Case 13: ZB_MethodName = "13 ADO Recordset + CopyFromRecordset"
        Case 14: ZB_MethodName = "14 DAO Recordset + CopyFromRecordset"
        Case 15: ZB_MethodName = "15 TSV + Workbooks.OpenText 直接取込"
        Case 16: ZB_MethodName = "16 Excel-DNA/XLL (Excel内プロセス)"
        Case Else: ZB_MethodName = "?"
    End Select
End Function

'==============================================================================
' 元データの取得 (準備。どの計測にも入らない)
'==============================================================================
Public Sub ZbDownloadKenAll()
    Const KEN_ALL_URL As String = _
        "https://www.post.japanpost.jp/service/search/zipcode/download/kogaki/zip/ken_all.zip"
    Dim http As Object, st As Object, sh As Object
    Dim zipPath As String
    Dim t0 As Currency, rc As Long

    On Error GoTo Fail
    Application.StatusBar = "ZipBench | KEN_ALL を取得中..."
    modZipRule.EnsureFolder ZB_Data()
    zipPath = ZB_Data() & "\ken_all.zip"
    t0 = modZipRule.QpcNow()

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", KEN_ALL_URL, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (ZipBench)"
    http.Send
    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, "ZbDownloadKenAll", "HTTP " & http.Status & " " & http.statusText
    End If

    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write http.responseBody
    st.SaveToFile zipPath, 2
    st.Close

    Set sh = CreateObject("WScript.Shell")
    rc = sh.Run("powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
                """Expand-Archive -LiteralPath '" & zipPath & "' -DestinationPath '" & _
                ZB_Data() & "' -Force""", 0, True)
    If rc <> 0 Then Err.Raise vbObjectError + 2, "ZbDownloadKenAll", "展開に失敗 (exit " & rc & ")"
    If Not modZipRule.FileExists(ZB_KenAllPath()) Then
        Err.Raise 53, "ZbDownloadKenAll", "展開後に KEN_ALL.CSV が見つからない"
    End If

    ZB_Log "KEN_ALL 取得 " & Format$(modZipRule.QpcSince(t0), "0.00") & " 秒"
    Application.StatusBar = "ZipBench | KEN_ALL 取得完了"
    ZB_Sheet().Range("B7").Value = "日本郵便 KEN_ALL.CSV 取得済 " & Format$(Now, "yyyy-mm-dd hh:nn")
    Exit Sub
Fail:
    ZB_Log "KEN_ALL 取得失敗: Err " & Err.Number & " " & Err.Description
    Application.StatusBar = "ZipBench | KEN_ALL 取得失敗 Err " & Err.Number
    MsgBox "KEN_ALL の取得に失敗しました。" & vbCrLf & "Err " & Err.Number & ": " & Err.Description, _
           vbExclamation, "ZipBench"
End Sub

'==============================================================================
' 準備 -- 正本の 3 シートを作る。どの計測にも入らない。
'==============================================================================
Public Sub ZbPrepare()
    Dim ws As Worksheet
    Dim t0 As Currency
    Dim zips As Variant
    Dim d As Object
    Dim nDict As Long

    On Error GoTo Fail
    Set ws = ZB_Sheet()
    g_Cancel = False
    g_Prepared = False
    Application.StatusBar = "ZipBench | 準備中..."

    g_N = CLng(ws.Range("B3").Value)
    If g_N < 1 Then Err.Raise 5, "ZbPrepare", "件数が不正"
    If g_N > 1048576 Then Err.Raise 5, "ZbPrepare", "件数がワークシートの行数を超える"

    modZipRule.EnsureFolder ZB_Data()
    modZipRule.EnsureFolder ZB_Work()
    If Not modZipRule.FileExists(ZB_KenAllPath()) Then
        Err.Raise 53, "ZbPrepare", "KEN_ALL.CSV が無い。先に [KEN_ALL取得] を押す"
    End If

    ' --- 辞書 (KEN_ALL より古ければ作り直す) ---
    t0 = modZipRule.QpcNow()
    If NeedsRebuild(ZB_DictPath(), ZB_KenAllPath()) Then
        nDict = modZipRule.BuildDictionary(ZB_KenAllPath(), ZB_DictPath())
        ZB_Log "辞書を作成: " & nDict & " 件"
    End If
    Set d = modZipRule.LoadDictionary(ZB_DictPath(), zips)
    g_MasterN = d.Count
    ZB_Log "辞書読込 " & Format$(modZipRule.QpcSince(t0), "0.00") & " 秒 / " & g_MasterN & " 件"

    ' --- 入力の決定的生成 ---
    t0 = modZipRule.QpcNow()
    g_Input = modZipRule.GenerateInput(g_N, zips)
    ZB_Log "入力生成 " & Format$(modZipRule.QpcSince(t0), "0.00") & " 秒 / " & g_N & " 件"

    ' --- 正解 ---
    t0 = modZipRule.QpcNow()
    g_Expected = modZipRule.BuildExpected(g_Input, d)
    ZB_Log "正解作成 " & Format$(modZipRule.QpcSince(t0), "0.00") & " 秒"

    ' --- 正本の 3 シート ---
    t0 = modZipRule.QpcNow()
    PrepareDataWorkbook d, zips
    ZB_Log "正本シート (MASTER/INPUT/OUTPUT) を用意 " & Format$(modZipRule.QpcSince(t0), "0.00") & " 秒"

    g_Prepared = True
    ws.Range("B6").Value = "準備済 (" & Format$(g_N, "#,##0") & " 件 / Master " & Format$(g_MasterN, "#,##0") & " 件)"
    Application.StatusBar = "ZipBench | 準備完了 " & Format$(g_N, "#,##0") & " 件"
    ThisWorkbook.Activate
    ws.Activate
    Exit Sub
Fail:
    g_Prepared = False
    ws.Range("B6").Value = "準備失敗"
    ZB_Log "準備失敗: Err " & Err.Number & " " & Err.Description
    Application.StatusBar = "ZipBench | 準備失敗 Err " & Err.Number
    MsgBox "準備に失敗しました。" & vbCrLf & "Err " & Err.Number & ": " & Err.Description, vbExclamation, "ZipBench"
End Sub

Private Function NeedsRebuild(ByVal target As String, ByVal src As String) As Boolean
    If Not modZipRule.FileExists(target) Then NeedsRebuild = True: Exit Function
    NeedsRebuild = (FileDateTime(target) < FileDateTime(src))
End Function

Private Sub PrepareDataWorkbook(ByVal d As Object, ByRef zips As Variant)
    Dim wsM As Worksheet, wsI As Worksheet, wsO As Worksheet
    Dim buf() As String
    Dim i As Long, chunk As Long, upTo As Long

    If Not DataWbAlive() Then
        Set g_DataWb = Workbooks.Add(xlWBATWorksheet)
        g_DataWb.Windows(1).Visible = True
    End If
    Do While g_DataWb.Worksheets.Count < 3
        g_DataWb.Worksheets.Add After:=g_DataWb.Worksheets(g_DataWb.Worksheets.Count)
    Loop
    Set wsM = g_DataWb.Worksheets(1): wsM.Name = DATA_MASTER
    Set wsI = g_DataWb.Worksheets(2): wsI.Name = DATA_INPUT
    Set wsO = g_DataWb.Worksheets(3): wsO.Name = DATA_OUTPUT
    wsM.Cells.Clear: wsI.Cells.Clear: wsO.Cells.Clear
    wsM.Columns("A:B").NumberFormat = "@"
    wsI.Columns("A").NumberFormat = "@"
    wsO.Columns("A").NumberFormat = "@"

    ' MASTER: 郵便番号と住所を辞書の並び (KEN_ALL 順) で
    ReDim buf(1 To g_MasterN, 1 To 2)
    For i = 0 To g_MasterN - 1
        buf(i + 1, 1) = zips(i)
        buf(i + 1, 2) = d(zips(i))
    Next i
    wsM.Range(wsM.Cells(FIRST_ROW, 1), wsM.Cells(FIRST_ROW + g_MasterN - 1, 2)).Value = buf

    ' INPUT: 100 万件。10 万件ずつ置く (準備工程なので分割してよい)
    chunk = 100000
    i = 1
    Do While i <= g_N
        upTo = i + chunk - 1
        If upTo > g_N Then upTo = g_N
        ReDim buf(1 To upTo - i + 1, 1 To 1)
        Dim j As Long
        For j = i To upTo
            buf(j - i + 1, 1) = g_Input(j)
        Next j
        wsI.Range(wsI.Cells(FIRST_ROW + i - 1, 1), wsI.Cells(FIRST_ROW + upTo - 1, 1)).Value = buf
        i = upTo + 1
    Loop
End Sub

Public Function DataWbAlive() As Boolean
    Dim s As String
    On Error Resume Next
    If g_DataWb Is Nothing Then DataWbAlive = False: Exit Function
    s = g_DataWb.Name
    DataWbAlive = (Err.Number = 0)
    Err.Clear
End Function

Public Function DataSheet(ByVal nm As String) As Worksheet
    If Not DataWbAlive() Then Err.Raise 91, "DataSheet", "正本のブックが無い。先に [準備] を押す"
    Set DataSheet = g_DataWb.Worksheets(nm)
End Function

'==============================================================================
' 候補2 の「Excel反映」: 出来上がった TSV を Excel ネイティブの取込経路で
' OUTPUT へ一括で入れる。
'
' ワーカーは一時名で書いてから atomic rename しているので、ここへ来た時点で
' ファイルは必ず完成している。存在確認のポーリングも待機も要らない。
'
' 取り込みは QueryTables。列は xlTextFormat 固定にする。既定のままだと
' 住所の先頭が数字の行を Excel が数値と見なして壊す。
'==============================================================================
Public Function ZB_ImportTsv(ByVal path As String) As Double
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim t0 As Currency

    Set ws = DataSheet(DATA_OUTPUT)
    t0 = modZipRule.QpcNow()

    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & path, _
                                Destination:=ws.Cells(FIRST_ROW, 1))
    With qt
        .TextFilePlatform = 65001              ' UTF-8
        .TextFileParseType = xlDelimited
        .TextFileTabDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileConsecutiveDelimiter = False
        .TextFileColumnDataTypes = Array(xlTextFormat)
        .AdjustColumnWidth = False
        .PreserveFormatting = False
        .RefreshStyle = xlOverwriteCells
        .BackgroundQuery = False
        .SaveData = False
        .Refresh BackgroundQuery:=False
    End With
    ZB_ImportTsv = modZipRule.QpcSince(t0)

    ' 取込の痕跡 (接続オブジェクト) は残さない。値だけ残す。
    On Error Resume Next
    qt.Delete
    Err.Clear
End Function

'==============================================================================
' 同じ 1 本のテキストファイルを、別々の経路で OUTPUT へ入れる。
'
' 方式 9・13・14・15 はワーカーが書いたまったく同じファイルを読む。
' 違うのは Excel 側の取り込み経路だけで、ファイルの中身は 1 バイトも違わない。
'
' 拡張子を .txt にしてあるのは ACE のテキスト ドライバの都合。既定で認識するのは
' txt / csv / tab / asc だけで、.tsv は「不明な拡張子」として弾かれる。
' レジストリの DisabledExtensions をいじれば通せるが、比較のために利用者の環境を
' 書き換えるのは筋が悪いので、こちらを合わせている。中身はタブ区切りのまま。
'==============================================================================

' ADO も DAO も、テキストを表として読むには列の型定義が要る。無いとドライバが
' 先頭数行から型を推測し、郵便番号側で数値化が起きる。schema.ini は取り込み経路の
' 一部なので、計測の中に入れてある。
Private Sub WriteSchemaIni(ByVal folder As String, ByVal fileName As String)
    Dim f As Integer
    f = FreeFile
    Open folder & "\schema.ini" For Output As #f
    Print #f, "[" & fileName & "]"
    Print #f, "ColNameHeader=False"
    Print #f, "Format=TabDelimited"
    Print #f, "CharacterSet=65001"
    Print #f, "Col1=addr Text Width 255"
    Close #f
End Sub

Private Sub SplitPath(ByVal path As String, ByRef folder As String, ByRef fileName As String)
    Dim p As Long
    p = InStrRev(path, "\")
    folder = Left$(path, p - 1)
    fileName = Mid$(path, p + 1)
End Sub

' 方式 13  ADODB.Recordset + Range.CopyFromRecordset
Public Function ZB_ImportAdo(ByVal path As String) As Double
    Dim ws As Worksheet
    Dim cn As Object, rs As Object
    Dim folder As String, fileName As String
    Dim t0 As Currency

    Set ws = DataSheet(DATA_OUTPUT)
    SplitPath path, folder, fileName

    t0 = modZipRule.QpcNow()
    WriteSchemaIni folder, fileName

    Set cn = CreateObject("ADODB.Connection")
    cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & folder & ";" & _
            "Extended Properties=""text;HDR=NO;FMT=TabDelimited;CharacterSet=65001"";"

    Set rs = CreateObject("ADODB.Recordset")
    ' adOpenForwardOnly(0) / adLockReadOnly(1)。前へ 1 回流すだけなので
    ' カーソルを持つ必要がない。
    rs.Open "SELECT * FROM [" & fileName & "]", cn, 0, 1

    ws.Cells(FIRST_ROW, 1).CopyFromRecordset rs

    ZB_ImportAdo = modZipRule.QpcSince(t0)

    On Error Resume Next
    rs.Close
    cn.Close
    Err.Clear
End Function

' 方式 14  DAO.Recordset + Range.CopyFromRecordset
' DAO はテキスト ファイルを表として開くとき、拡張子の区切りに "." ではなく "#" を使う。
Public Function ZB_ImportDao(ByVal path As String) As Double
    Dim ws As Worksheet
    Dim de As Object, db As Object, rs As Object
    Dim folder As String, fileName As String
    Dim t0 As Currency

    Set ws = DataSheet(DATA_OUTPUT)
    SplitPath path, folder, fileName

    t0 = modZipRule.QpcNow()
    WriteSchemaIni folder, fileName

    Set de = CreateObject("DAO.DBEngine.120")
    ' OpenDatabase(名前, 排他, 読み取り専用, 接続文字列)
    Set db = de.OpenDatabase(folder, False, True, "Text;")
    Set rs = db.OpenRecordset("SELECT * FROM [" & Replace(fileName, ".", "#") & "]")

    ws.Cells(FIRST_ROW, 1).CopyFromRecordset rs

    ZB_ImportDao = modZipRule.QpcSince(t0)

    On Error Resume Next
    rs.Close
    db.Close
    Err.Clear
End Function

' 方式 15  Workbooks.OpenText で別ブックとして開き、範囲ごと OUTPUT へ複写
' QueryTable も Recordset も経由しない、Excel が自前でテキストを開く経路。
Public Function ZB_ImportOpenText(ByVal path As String) As Double
    Dim ws As Worksheet
    Dim src As Workbook
    Dim t0 As Currency

    Set ws = DataSheet(DATA_OUTPUT)

    t0 = modZipRule.QpcNow()

    Application.Workbooks.OpenText fileName:=path, Origin:=65001, StartRow:=1, _
        DataType:=xlDelimited, Tab:=True, Semicolon:=False, Comma:=False, _
        Space:=False, Other:=False, ConsecutiveDelimiter:=False, _
        FieldInfo:=Array(Array(1, xlTextFormat)), Local:=False
    Set src = Application.ActiveWorkbook

    ' VBA の配列を経由せず、Excel の中だけで範囲を移す。
    src.Worksheets(1).Range("A1").Resize(g_N, 1).Copy _
        Destination:=ws.Cells(FIRST_ROW, 1)

    ZB_ImportOpenText = modZipRule.QpcSince(t0)

    On Error Resume Next
    Application.CutCopyMode = False
    src.Close SaveChanges:=False
    Err.Clear
End Function

Public Sub ClearOutput()
    Dim ws As Worksheet
    Set ws = DataSheet(DATA_OUTPUT)
    ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + g_N - 1, 1)).ClearContents
End Sub

'==============================================================================
' 計測前後の Excel 設定。全方式でまったく同じにする。
'==============================================================================
Public Sub ZB_EnterMeasure(Optional ByVal useEvents As Boolean = False)
    m_ScreenUpdating = Application.ScreenUpdating
    m_EnableEvents = Application.EnableEvents
    m_Calculation = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = useEvents
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler
    m_InMeasure = True
End Sub

Public Sub ZB_LeaveMeasure()
    If Not m_InMeasure Then Exit Sub
    On Error Resume Next
    Application.Calculation = m_Calculation
    Application.EnableEvents = m_EnableEvents
    Application.ScreenUpdating = m_ScreenUpdating
    Application.EnableCancelKey = xlInterrupt
    m_InMeasure = False
End Sub

'==============================================================================
' 照合 -- 全方式まったく同じ経路。OUTPUT を読んで正解と突き合わせる。
'==============================================================================
Public Sub ZB_VerifyFromCells(ByRef r As ZbResult)
    Dim ws As Worksheet
    Dim v As Variant
    Dim arr() As String
    Dim i As Long, n As Long
    Dim t0 As Currency

    n = g_N
    t0 = modZipRule.QpcNow()
    On Error GoTo Fail
    Set ws = DataSheet(DATA_OUTPUT)
    v = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + n - 1, 1)).Value
    ReDim arr(1 To n)
    If n = 1 Then
        arr(1) = CStr(v)
    Else
        For i = 1 To n
            arr(i) = CStr(v(i, 1))
        Next i
    End If
    r.VerifySec = modZipRule.QpcSince(t0)
    ZB_Verify r, arr, n
    Exit Sub
Fail:
    r.VerifySec = modZipRule.QpcSince(t0)
    r.MatchText = "照合失敗"
    r.MismatchRow = -1
End Sub

Public Sub ZB_Verify(ByRef r As ZbResult, ByRef got As Variant, ByVal rowCount As Long)
    Dim bad As Long
    r.Rows = rowCount
    If rowCount <> g_N Then
        r.MatchText = "件数違い"
        r.MismatchRow = -1
        Exit Sub
    End If
    bad = modZipRule.CompareArrays(got, g_Expected, g_N)
    If bad = 0 Then
        r.MatchText = "一致"
        r.MismatchRow = 0
    Else
        r.MatchText = "不一致"
        r.MismatchRow = bad
    End If
End Sub

'==============================================================================
' 結果の確定
'==============================================================================
Public Sub ZB_Conclude(ByRef r As ZbResult)
    If r.ErrNumber = 18 Then
        r.Outcome = "取消 (Esc)"
    ElseIf r.ErrNumber <> 0 Then
        r.Outcome = "失敗 Err " & r.ErrNumber
    ElseIf r.MatchText = "一致" Then
        r.Outcome = "完了・一致"
    ElseIf r.MatchText = "不一致" Then
        r.Outcome = "完了・不一致 (行 " & r.MismatchRow & ")"
    ElseIf r.MatchText = "件数違い" Then
        r.Outcome = "完了・件数違い"
    ElseIf Len(r.MatchText) = 0 Then
        r.Outcome = "未変換"
    Else
        r.Outcome = r.MatchText
    End If
End Sub

' StatusBar は全方式この 1 か所だけが書く。書式も 1 つだけ。
Public Sub ZB_FinalStatus(ByRef r As ZbResult)
    Dim pct As Double
    If g_N > 0 Then pct = r.Rows / g_N * 100#
    Application.StatusBar = "ZipBench | " & r.MethodName & _
        " | " & Format$(r.Rows, "#,##0") & "/" & Format$(g_N, "#,##0") & _
        " (" & Format$(pct, "0.0") & "%)" & _
        " | 処理E2E " & Format$(r.TotalSec, "0.000") & " 秒" & _
        " | " & r.Outcome
End Sub

Public Function ZB_NewResult(ByVal methodNo As Long) As ZbResult
    Dim r As ZbResult
    r.MethodNo = methodNo
    r.MethodName = ZB_MethodName(methodNo)
    r.WorkerKind = "-"
    r.Created = "-": r.Launched = "-": r.UiaOk = "-": r.Converted = "-"
    r.MismatchRow = 0
    ZB_NewResult = r
End Function

Public Sub ZB_ClearRow(ByVal methodNo As Long)
    Dim ws As Worksheet, row As Long
    On Error Resume Next
    Set ws = ZB_Sheet()
    row = ZB_RowOf(methodNo)
    ws.Range(ws.Cells(row, 3), ws.Cells(row, 27)).ClearContents
    ws.Range(ws.Cells(row, 3), ws.Cells(row, 27)).Interior.ColorIndex = xlColorIndexNone
End Sub

' 1 回ぶんを RUNS へ積み、その方式の全実行から中央値を出して BENCH を書き直す。
Public Sub ZB_RecordRun(ByRef r As ZbResult)
    Dim ws As Worksheet, rw As Long
    On Error Resume Next
    Set ws = ZB_RunsSheet()
    rw = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1
    If rw < 2 Then rw = 2
    ws.Cells(rw, 1).NumberFormat = "@"
    ws.Cells(rw, 1).Value = Format$(Now, "hh:nn:ss")
    ws.Cells(rw, 2).Value = r.MethodNo
    ws.Cells(rw, 3).Value = r.MethodName
    ws.Cells(rw, 4).Value = r.WorkerKind
    ws.Cells(rw, 5).Value = g_N
    ws.Cells(rw, 6).Value = CountRuns(r.MethodNo) + 1
    ws.Cells(rw, 7).Value = r.LaunchSec
    ws.Cells(rw, 8).Value = r.MasterSec
    ws.Cells(rw, 9).Value = r.InputSec
    ws.Cells(rw, 10).Value = r.DictSec
    ws.Cells(rw, 11).Value = r.ConvertSec
    ws.Cells(rw, 12).Value = r.WriteSec
    ws.Cells(rw, 13).Value = r.ImportSec
    ws.Cells(rw, 14).Value = r.NotifySec
    ws.Cells(rw, 15).Value = r.TotalSec
    ws.Cells(rw, 16).Value = r.VerifySec
    ws.Cells(rw, 17).Value = r.Rows
    ws.Cells(rw, 18).Value = r.MatchText
    ws.Cells(rw, 19).NumberFormat = "@"
    ws.Cells(rw, 19).Value = r.Hash
End Sub

Private Function CountRuns(ByVal methodNo As Long) As Long
    Dim ws As Worksheet, last As Long, i As Long, c As Long
    On Error Resume Next
    Set ws = ZB_RunsSheet()
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For i = 2 To last
        If ws.Cells(i, 2).Value = methodNo Then c = c + 1
    Next i
    CountRuns = c
End Function

' 保存済み全実行の中央値で BENCH を書く。
Public Sub ZB_WriteRow(ByRef r As ZbResult)
    Dim ws As Worksheet, row As Long, col As Long
    On Error Resume Next
    Set ws = ZB_Sheet()
    row = ZB_RowOf(r.MethodNo)

    ws.Cells(row, 2).Value = r.MethodName
    ws.Cells(row, 3).Value = r.WorkerKind
    ws.Cells(row, 4).Value = r.Created
    ws.Cells(row, 5).Value = r.Launched
    ws.Cells(row, 6).Value = r.UiaOk
    ws.Cells(row, 7).Value = r.Converted
    ws.Cells(row, 8).Value = r.ErrNumber

    ' 11・12 は「ファイル作成だけ」で計測対象の工程を持たない。それ以外は
    ' 保存済み全実行の中央値を書く。
    If r.MethodNo <> 11 And r.MethodNo <> 12 Then
        For col = 7 To 16                              ' RUNS の 7..16 = 起動..照合
            ws.Cells(row, col + 2).Value = MedianOf(r.MethodNo, col)
        Next col
        ws.Cells(row, 19).Value = CountRuns(r.MethodNo)
        ws.Cells(row, 20).Value = MinOf(r.MethodNo, 15)
        ws.Cells(row, 21).Value = MaxOf(r.MethodNo, 15)
    End If

    ws.Cells(row, 22).Value = r.Rows
    ws.Cells(row, 23).Value = r.MatchText
    ws.Cells(row, 24).Value = IIf(r.MismatchRow > 0, r.MismatchRow, "")
    ws.Cells(row, 25).NumberFormat = "@"
    ws.Cells(row, 25).Value = r.Hash
    ws.Cells(row, 26).Value = r.Outcome
    ws.Cells(row, 27).NumberFormat = "@"
    ws.Cells(row, 27).Value = r.Note

    For col = 4 To 7
        ws.Cells(row, col).Interior.Color = StateColor(ws.Cells(row, col).Value)
    Next col
    If r.MatchText = "一致" Then
        ws.Cells(row, 23).Interior.Color = RGB(200, 240, 200)
    ElseIf Len(r.MatchText) > 0 Then
        ws.Cells(row, 23).Interior.Color = RGB(250, 200, 200)
    End If
End Sub

Private Function CollectRuns(ByVal methodNo As Long, ByVal col As Long, ByRef v() As Double) As Long
    Dim ws As Worksheet, last As Long, i As Long, n As Long
    Set ws = ZB_RunsSheet()
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    ReDim v(1 To 1)
    If last < 2 Then CollectRuns = 0: Exit Function
    ReDim v(1 To last)
    For i = 2 To last
        If ws.Cells(i, 2).Value = methodNo Then
            n = n + 1
            v(n) = CDbl(ws.Cells(i, col).Value)
        End If
    Next i
    CollectRuns = n
End Function

Private Function MedianOf(ByVal methodNo As Long, ByVal col As Long) As Double
    Dim v() As Double, n As Long, i As Long, j As Long, t As Double
    n = CollectRuns(methodNo, col, v)
    If n = 0 Then Exit Function
    For i = 1 To n - 1
        For j = i + 1 To n
            If v(j) < v(i) Then t = v(i): v(i) = v(j): v(j) = t
        Next j
    Next i
    If (n Mod 2) = 1 Then
        MedianOf = v((n + 1) \ 2)
    Else
        MedianOf = (v(n \ 2) + v(n \ 2 + 1)) / 2#
    End If
End Function

Private Function MinOf(ByVal methodNo As Long, ByVal col As Long) As Double
    Dim v() As Double, n As Long, i As Long
    n = CollectRuns(methodNo, col, v)
    If n = 0 Then Exit Function
    MinOf = v(1)
    For i = 2 To n
        If v(i) < MinOf Then MinOf = v(i)
    Next i
End Function

Private Function MaxOf(ByVal methodNo As Long, ByVal col As Long) As Double
    Dim v() As Double, n As Long, i As Long
    n = CollectRuns(methodNo, col, v)
    If n = 0 Then Exit Function
    MaxOf = v(1)
    For i = 2 To n
        If v(i) > MaxOf Then MaxOf = v(i)
    Next i
End Function

Private Function StateColor(ByVal v As String) As Long
    Select Case v
        Case "OK":  StateColor = RGB(200, 240, 200)
        Case "NG":  StateColor = RGB(250, 200, 200)
        Case Else:  StateColor = RGB(240, 240, 240)
    End Select
End Function

Public Function ZB_WorkerKind() As String
    Dim s As String
    s = LCase$(Trim$(CStr(ZB_Sheet().Range("B4").Value)))
    If s <> "emitted" Then s = "prebuilt"
    ZB_WorkerKind = s
End Function

Public Function ZB_WorkerMode() As String
    Dim s As String
    s = LCase$(Trim$(CStr(ZB_Sheet().Range("B5").Value)))
    Select Case s
        Case "offscreen", "minimized", "normal", "hidden"
        Case Else: s = "offscreen"
    End Select
    ZB_WorkerMode = s
End Function

'==============================================================================
' 実行
'==============================================================================
Public Sub ZbRunMethod(ByVal methodNo As Long)
    Dim r As ZbResult

    If Not g_Prepared Then
        MsgBox "先に [準備] を押してください。", vbExclamation, "ZipBench"
        Exit Sub
    End If

    g_Cancel = False
    ZB_Log "=== 方式 " & methodNo & " 開始 (" & ZB_MethodName(methodNo) & ", " & Format$(g_N, "#,##0") & " 件) ==="

    Select Case methodNo
        Case 1, 2:  r = modZipVba.RunVbaMethod(methodNo)
        Case 7:     r = modZipExcel.RunExcelMethod()
        Case 16:    r = modZipXll.RunXllMethod()
        Case Else:  r = modZipWorker.RunWorkerMethod(methodNo)   ' 3-6, 8, 9, 10, 13-15
    End Select

    ' ここから先は計測の外
    ZB_Conclude r
    ZB_RecordRun r
    ZB_WriteRow r
    ZB_FinalStatus r
    ZB_Log "方式 " & methodNo & " 終了: " & r.Outcome & _
           " / 処理E2E " & Format$(r.TotalSec, "0.000") & "s" & _
           " (起動 " & Format$(r.LaunchSec, "0.000") & _
           " / M " & Format$(r.MasterSec, "0.000") & _
           " / I " & Format$(r.InputSec, "0.000") & _
           " / 辞書 " & Format$(r.DictSec, "0.000") & _
           " / 変換 " & Format$(r.ConvertSec, "0.000") & _
           " / 出力 " & Format$(r.WriteSec, "0.000") & _
           " / 反映 " & Format$(r.ImportSec, "0.000") & _
           " / 通知 " & Format$(r.NotifySec, "0.000") & ")" & _
           " / 照合 " & Format$(r.VerifySec, "0.000") & "s" & _
           IIf(Len(r.Note) > 0, " / " & r.Note, "")
End Sub

Public Sub ZbRunAll()
    Dim i As Long
    If Not g_Prepared Then
        MsgBox "先に [準備] を押してください。", vbExclamation, "ZipBench"
        Exit Sub
    End If
    For i = 1 To 8
        ZbRunMethod i
        DoEvents
        If g_Cancel Then ZB_Log "取消により残りの方式を中止": Exit For
    Next i
End Sub

'==============================================================================
' 取消と後始末
'==============================================================================
Public Sub ZbCancel()
    g_Cancel = True
    ZB_Log "取消を要求"
    Application.StatusBar = "ZipBench | 取消を要求しました"
End Sub

Public Sub ZbCleanup()
    On Error Resume Next
    g_Cancel = True
    ZB_LeaveMeasure
    ZbDisarmSignal

    modZipWorker.ShutdownOwnWorker
    modZipExcel.QuitOwnExcel

    If DataWbAlive() Then
        g_DataWb.Saved = True
        g_DataWb.Close SaveChanges:=False
    End If
    Set g_DataWb = Nothing
    g_Prepared = False
    ZB_Sheet().Range("B6").Value = "未"

    Application.StatusBar = False
    ZB_Log "後始末を実行した (自分が作ったワーカー・別Excel・正本ブックのみ)"

    If modZipExcel.OwnExcelPid() <> 0 Then
        If modZipExcel.OwnExcelStillAlive() Then
            ZB_Log "注意: 方式7 で起こした Excel (pid " & modZipExcel.OwnExcelPid() & _
                   ") がまだ生きている。手で終了させること"
        Else
            ZB_Log "方式7 で起こした Excel (pid " & modZipExcel.OwnExcelPid() & ") は終了済み"
        End If
    End If
End Sub
