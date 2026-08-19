VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbChannel - FE と BE のあいだ（Temp のファイルだけで往復する）
'
' 責務：%TEMP%\pixelbridge\ の下だけを触る。置き場所を知っているのはこの
' クラスだけで、他はパスを組み立てない。
'
'   command.json  FE → BE の指示（cmd と seq）
'   progress.json BE → FE の状態
'   fe.json       FE が生きている印。BE はこれが更新されなくなったら自分で終わる
'   pi.txt        計算できた桁そのもの
'   pixelbridge_be.xlsm  BE 用のブックのコピー
'
' 守っている約束
'   - **本文のような自由な文字列を JSON へ入れない。** 引用符も改行もそのまま
'     置ける別ファイルにする。JSON に入れていた頃、日本語の本文が読めずに
'     BE が毎回失敗した。
'   - **ファイルはバイトで読む。** Input$(LOF(f), f) は「LOF バイト分の文字」を
'     要求する読み方で、ASCII のうちは動くが CP932 の日本語では必ず末尾を越えて
'     読み取り全体が落ちる。
'   - **配布物のフォルダには何も書かない。** BE 用のコピーも Temp に置く。
'     置き場所を Temp にできるのは、開くときに AutomationSecurity を落として
'     いるから（信頼された場所かどうかとは別の口）。
'==============================================================================
Option Explicit

Private m_log As PbLog
Private m_seq As Long

Public Sub Init(ByVal logger As PbLog)
    Set m_log = logger
End Sub

'--- 置き場所 -----------------------------------------------------------------
Public Property Get Folder() As String
    Dim p As String
    p = Environ$("TEMP")
    If Right$(p, 1) <> "\" Then p = p & "\"
    Folder = p & "pixelbridge\"
End Property

Public Property Get CmdPath() As String
    CmdPath = Folder & "command.json"
End Property

Public Property Get ProgPath() As String
    ProgPath = Folder & "progress.json"
End Property

Public Property Get FePath() As String
    FePath = Folder & "fe.json"
End Property

Public Property Get PiPath() As String
    PiPath = Folder & "pi.txt"
End Property

Public Property Get BeBookPath() As String
    BeBookPath = Folder & "pixelbridge_be.xlsm"
End Property

Public Sub EnsureFolder()
    On Error Resume Next
    If Dir$(Folder, vbDirectory) = "" Then MkDir Left$(Folder, Len(Folder) - 1)
End Sub

'--- 読み書き -----------------------------------------------------------------
Public Sub WriteAll(ByVal path As String, ByVal s As String)
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
    ' 相手が同じ瞬間に読んでいると開けないことがある。少し粘る。
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

Public Function ReadAll(ByVal path As String) As String
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

Public Sub KillQuiet(ByVal path As String)
    On Error Resume Next
    If Dir$(path) <> "" Then Kill path
End Sub

'--- やりとり -----------------------------------------------------------------
' FE が生きている印。BE はこの中身が変わらなくなったら置き去りと判断して
' 自分から終わる。「ファイルが有るか」では足りない（強制終了で残るため）。
Public Sub Heartbeat()
    WriteAll FePath(), "{""fe"":""alive"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
End Sub

' 指示を 1 つ送る。seq は単調増加で、BE は「前より大きい seq」だけを実行する。
' 走行中の BE も seq を見て、新しい指示が来たことに気づいて畳む。
Public Function SendCommand(ByVal cmd As String, ByVal extraJson As String) As Long
    Dim js As String
    m_seq = m_seq + 1
    js = "{""cmd"":""" & cmd & """,""seq"":" & m_seq
    If Len(extraJson) > 0 Then js = js & "," & extraJson
    js = js & ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    WriteAll CmdPath(), js
    SendCommand = m_seq
End Function

Public Function ReadProgress() As String
    ReadProgress = ReadAll(ProgPath())
End Function

' BE が置いた桁の列。WriteAll の Print # が末尾へ改行を足すので落とす。
Public Function ReadDigits() As String
    Dim s As String
    s = ReadAll(PiPath())
    Do While Len(s) > 0 And (Right$(s, 1) = vbLf Or Right$(s, 1) = vbCr)
        s = Left$(s, Len(s) - 1)
    Loop
    ReadDigits = s
End Function

'--- 後始末 -------------------------------------------------------------------
' やりとりに使ったものを全部消し、フォルダも畳む。BE 用のコピーもこの中に
' あるので、先に消さないとフォルダが空にならず RmDir が失敗して残る。
Public Sub CleanAll()
    On Error Resume Next
    KillQuiet PiPath()
    KillQuiet ProgPath()
    KillQuiet CmdPath()
    KillQuiet FePath()
    KillQuiet BeBookPath()
    RmDir Left$(Folder, Len(Folder) - 1)
    If Not m_log Is Nothing Then m_log.Note "channel: temp cleaned"
End Sub
