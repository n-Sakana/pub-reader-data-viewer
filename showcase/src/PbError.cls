VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PbError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' PbError - 例外の文脈
'
' 責務：Err の中身は次の 1 行で消える。捕まえたその場で写し取り、どの手続きの
' どういう状態で起きたかを添えて、外側の境界まで運ぶ。
'
' 方針
'   - **想定内の状態は例外にしない。** メモ帳がまだ無い、BE がまだ起きていない、
'     といったものは戻り値（Boolean / 状態コード）で表す。ここへ来るのは
'     「起こるはずのないこと」だけ。
'   - 外側の境界（ティック 1 回、起動、終了）で必ず捕まえる。捕まえたら
'     記録し、明示的に終了・参照解放へ流す。**Class_Terminate に片付けを
'     任せない。** 解放の順序が意味を持つ場面があり（掴んだ UIA 要素 →
'     UIA クライアント → 別プロセスの Excel）、GC の都合で順序が決まると
'     プロセスが残る。
'   - 1 ティックの失敗でポンプを止めない。記録して次のティックを張り直す。
'==============================================================================
Option Explicit

Private m_has As Boolean
Private m_number As Long
Private m_desc As String
Private m_source As String
Private m_proc As String
Private m_state As String
Private m_at As Date

' Err の中身をその場で写す。呼ぶのはエラーハンドラの先頭。
Public Sub Capture(ByVal procName As String, ByVal stateText As String)
    m_has = True
    m_number = Err.Number
    m_desc = Err.Description
    m_source = Err.Source
    m_proc = procName
    m_state = stateText
    m_at = Now
End Sub

Public Sub Clear()
    m_has = False
    m_number = 0
    m_desc = ""
    m_source = ""
    m_proc = ""
    m_state = ""
End Sub

Public Property Get HasError() As Boolean
    HasError = m_has
End Property

Public Property Get Number() As Long
    Number = m_number
End Property

Public Property Get Description() As String
    Description = m_desc
End Property

Public Property Get Source() As String
    Source = m_source
End Property

Public Property Get Procedure() As String
    Procedure = m_proc
End Property

Public Property Get StateText() As String
    StateText = m_state
End Property

' 記録用の 1 行。画面に出すには長いので、そちらは短い要約を別に作る。
Public Property Get Text() As String
    If Not m_has Then
        Text = "(no error)"
        Exit Property
    End If
    Text = "err " & m_number & " in " & m_proc & " : " & m_desc & _
           IIf(Len(m_source) > 0, " [src " & m_source & "]", "") & _
           IIf(Len(m_state) > 0, " {" & m_state & "}", "")
End Property

' 画面の狭い欄に出す用。
Public Property Get Brief() As String
    If Not m_has Then
        Brief = ""
    Else
        Brief = m_proc & " " & m_number
    End If
End Property
