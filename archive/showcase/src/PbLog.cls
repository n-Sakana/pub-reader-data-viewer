VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbLog - 記録
'
' 責務：どこで何ミリ秒かかったかを残す。行き先は 2 つだけ。
'
'   既定 … VBE のイミディエイトウィンドウ（Debug.Print）
'   切替 … テキストファイル。下の LOG_TO_FILE を True にするだけで変わる
'
' 運用中に開発者がいない端末でも記録を採れるように、%TEMP%\pixelbridge.debug
' というマーカーファイルがあれば、コードを触らずファイル出力へ倒れる。その
' ときの行き先は %TEMP%\pixelbridge-build.log で、検証スクリプトはこれを読む。
'
' **ログ自身をボトルネックにしない**ことがこのクラスの一番の制約。
'   - ファイル出力は 1 行ごとに開き直さない。溜めておいて区切りで書き出す。
'     1 行ごとに Open / Print / Close すると、1 秒のティックの中で数十回の
'     ファイル操作が起き、測っているものより測る側が重くなる。
'   - 区間の記録は「かかった時間が下限を超えたときだけ」。毎セル・毎桁のような
'     細かさで残すと、行数そのものが処理時間になる。
'   - Flush はティックの終わりのような自然な区切りで呼ぶ。そうすれば外から
'     ログを読む側も 1 ティック遅れまでで追いつける。
'==============================================================================
Option Explicit

'--- ここ 1 行でイミディエイト出力とファイル出力を切り替える
Private Const LOG_TO_FILE As Boolean = False

'--- ファイル出力にしたときの置き場所（リポジトリ直下）
Private Const LOG_FILE_NAME As String = "pixelbridge.log"

'--- 区間の記録を残す下限（ミリ秒）。これ未満は捨てる
Private Const SPAN_FLOOR_MS As Double = 30#

'--- 溜める行数の上限。超えたらその場で書き出す
Private Const FLUSH_LINES As Long = 64

Private m_toFile As Boolean
Private m_path As String
Private m_buf() As String
Private m_n As Long
Private m_on As Boolean

Private Sub Class_Initialize()
    ReDim m_buf(0 To FLUSH_LINES - 1)
    m_n = 0
    m_on = True
    ResolveSink
End Sub

' 行き先を決める。マーカーがあればファイル（検証と同じ場所）、
' 無ければ LOG_TO_FILE の指定どおり。
Private Sub ResolveSink()
    Dim tmp As String
    On Error Resume Next
    tmp = Environ$("TEMP")
    If Right$(tmp, 1) <> "\" Then tmp = tmp & "\"
    If Len(Dir$(tmp & "pixelbridge.debug")) > 0 Then
        m_toFile = True
        m_path = tmp & "pixelbridge-build.log"
        Exit Sub
    End If
    m_toFile = LOG_TO_FILE
    If m_toFile Then m_path = RepoRoot() & LOG_FILE_NAME
End Sub

' リポジトリ直下。成果物は <repo>\showcase\dist\ に置かれるので 2 つ上。
' そこが無ければブックの隣へ落とす（配布物を別の場所へ置いた場合）。
Private Function RepoRoot() As String
    Dim p As String
    On Error Resume Next
    p = ThisWorkbook.Path
    If Len(p) = 0 Then
        RepoRoot = Environ$("TEMP") & "\"
        Exit Function
    End If
    If Len(Dir$(p & "\..\..\.git", vbDirectory)) > 0 Then
        RepoRoot = p & "\..\..\"
    Else
        RepoRoot = p & "\"
    End If
End Function

Public Property Get Enabled() As Boolean
    Enabled = m_on
End Property

Public Property Let Enabled(ByVal v As Boolean)
    m_on = v
End Property

' 1 行残す。
Public Sub Note(ByVal s As String)
    Dim row As String
    If Not m_on Then Exit Sub
    On Error Resume Next
    row = Format$(Now, "hh:nn:ss") & " " & Format$(Timer, "0.000") & " " & s
    If Not m_toFile Then
        Debug.Print row
        Exit Sub
    End If
    If m_n > UBound(m_buf) Then Flush
    m_buf(m_n) = row
    m_n = m_n + 1
    If m_n >= FLUSH_LINES Then Flush
End Sub

' 区間の開始時刻。戻り値をそのまま Span へ渡す。
Public Function Mark() As Double
    Mark = Timer
End Function

' 区間の終わり。下限を超えたときだけ残す。
Public Sub Span(ByVal name As String, ByVal t0 As Double)
    Dim ms As Double
    If Not m_on Then Exit Sub
    ms = ElapsedSince(t0) * 1000
    If ms < SPAN_FLOOR_MS Then Exit Sub
    Note "  span " & name & " " & Format$(ms, "0") & "ms"
End Sub

' 区間の終わり（下限に関係なく残す）。ティック全体のように、
' 毎回の値そのものが意味を持つものだけに使う。
Public Sub SpanAlways(ByVal name As String, ByVal t0 As Double)
    If Not m_on Then Exit Sub
    Note "  span " & name & " " & Format$(ElapsedSince(t0) * 1000, "0") & "ms"
End Sub

' 溜めた行を書き出す。ティックの終わりなど、自然な区切りで呼ぶ。
Public Sub Flush()
    Dim f As Integer
    Dim i As Long
    If m_n = 0 Then Exit Sub
    If Not m_toFile Then
        m_n = 0
        Exit Sub
    End If
    On Error Resume Next
    f = FreeFile
    Open m_path For Append As #f
    If Err.Number = 0 Then
        For i = 0 To m_n - 1
            Print #f, m_buf(i)
        Next i
        Close #f
        m_n = 0
    Else
        ' 書けないなら黙って捨てる。記録できないことで処理を止めない。
        m_n = 0
    End If
End Sub
