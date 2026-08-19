VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbDisplayBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbDisplayBook - 表示用 Excel（ベンチの桁を映すだけの一時ブック）
'
' 責務：別プロセスの Excel を 1 つ持ち、保存しないブックへ桁を書く。BE とは
' 別物で、計算はしない。書くのは FE で、BE はこのブックを知らない。
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
Option Explicit

Private m_app As Object
Private m_book As Object
Private m_sheet As Object
Private m_log As PbLog
Private m_fontMono As String

Public Sub Init(ByVal logger As PbLog, ByVal fontMono As String)
    Set m_log = logger
    m_fontMono = fontMono
End Sub

Public Property Get IsOpen() As Boolean
    IsOpen = Not (m_sheet Is Nothing)
End Property

Public Function Ensure() As Boolean
    On Error GoTo Failed
    If Not m_app Is Nothing Then
        ' 生きているか確かめる。落ちていれば作り直す。
        If m_app.Workbooks.Count > 0 Then
            PrepareSheet
            Ensure = True
            Exit Function
        End If
    End If
    Set m_app = CreateObject("Excel.Application")
    m_app.DisplayAlerts = False
    m_app.EnableEvents = False
    m_app.UserControl = False
    Set m_book = m_app.Workbooks.Add
    Set m_sheet = m_book.Worksheets(1)
    m_app.Visible = True
    PrepareSheet
    Ensure = True
    Exit Function
Failed:
    m_log.Note "disp: FAILED " & Err.Number & " " & Err.Description
    CloseQuiet
    Ensure = False
End Function

Private Sub PrepareSheet()
    On Error Resume Next
    With m_sheet
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
    ' 全セルに出る。使い捨ての別プロセスなので、その Excel のチェックだけ切る
    ' （利用者の Excel の設定には触っていない）。
    m_app.ErrorCheckingOptions.NumberAsText = False
    m_app.ActiveWindow.DisplayGridlines = False
End Sub

' 画面の指定の場所へ置く。座標はデバイス px、Excel はポイントで受けるので変換する。
Public Sub PlaceAt(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, _
                   ByVal pxPerPt As Double)
    On Error Resume Next
    If m_app Is Nothing Or pxPerPt <= 0 Then Exit Sub
    m_app.WindowState = xlNormal
    m_app.Left = x / pxPerPt
    m_app.Top = y / pxPerPt
    m_app.Width = w / pxPerPt
    m_app.Height = h / pxPerPt
    m_app.Visible = True
End Sub

' 10 桁ぶんを 1 セルへ。先頭に ' を付けて、数値へ化けるのを止める。
Public Sub WriteGroup(ByVal r As Long, ByVal c As Long, ByVal piece As String)
    On Error Resume Next
    If m_sheet Is Nothing Then Exit Sub
    m_sheet.Cells(r, c).Value = "'" & piece
End Sub

' 保存せずに閉じる。順序を変えない（上のコメントを参照）。
Public Sub CloseQuiet()
    On Error Resume Next
    If Not m_book Is Nothing Then
        m_book.Saved = True
        m_book.Close SaveChanges:=False
    End If
    If Not m_app Is Nothing Then
        m_app.DisplayAlerts = False
        m_app.Quit
    End If
    Set m_sheet = Nothing
    Set m_book = Nothing
    Set m_app = Nothing
End Sub
