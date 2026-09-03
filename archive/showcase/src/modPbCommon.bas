Attribute VB_Name = "modPbCommon"
'==============================================================================
' modPbCommon - UI Automation の識別子と、状態を持たない文字列ユーティリティ
'
' 責務：どのクラスからも使う「値」と「純粋関数」だけ。状態も副作用も持たない。
'
' UIA の識別子をここへ集めてあるのには理由がある。**パターン ID とプロパティ
' ID は、間違えても例外にならない。** 存在しない組み合わせを渡すと
' GetCurrentPattern は Nothing を返し、GetCurrentPropertyValue は空を返すので、
' 呼び出し側からは「相手がその機能を持っていない」と区別がつかない。
' 実際にそれで長く誤診した：TransformPattern を 10003（正しくは 10016）で
' 訊いていたため、メモ帳が Transform を公開していないという誤った結論に至り、
' 「Win32 なしでは窓を動かせない」と文書にまで書いた。数字は 1 か所に集め、
' 隣にコメントで名前を書いておく。
'==============================================================================
Option Explicit

'--- TreeScope
Public Const TS_CHILDREN As Long = 2
Public Const TS_DESCENDANTS As Long = 4

'--- プロパティ ID
Public Const UIA_ControlTypePropertyId As Long = 30003
Public Const UIA_HasKeyboardFocusPropertyId As Long = 30008
Public Const UIA_ClassNamePropertyId As Long = 30012
Public Const UIA_NativeWindowHandlePropertyId As Long = 30020
Public Const UIA_IsTransformPatternAvailablePropertyId As Long = 30042
Public Const UIA_IsValuePatternAvailablePropertyId As Long = 30043

'--- パターン ID。10003 は RangeValue であって Transform ではない。
Public Const UIA_ValuePatternId As Long = 10002
Public Const UIA_WindowPatternId As Long = 10009
Public Const UIA_TransformPatternId As Long = 10016

'--- コントロール型 ID
Public Const UIA_DocumentControlTypeId As Long = 50030
Public Const UIA_EditControlTypeId As Long = 50004

'--- 窓の表示状態
Public Const WindowVisualState_Normal As Long = 0
Public Const WindowVisualState_Minimized As Long = 2

' 改行の正規形。メモ帳の ValuePattern は CRLF を LF に直した文字列を返すので、
' 控えは LF 形で持つ。CRLF で控えると、書いた直後に「相手が変わった」と誤読して
' 1 往復よけいに写し直す。書き戻すときだけ CRLF へ戻す。
Public Function ToLf(ByVal s As String) As String
    ToLf = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

Public Function ToCrLf(ByVal s As String) As String
    ToCrLf = Replace(ToLf(s), vbLf, vbCrLf)
End Function

' JSON から 1 つの値を取り出す。やりとりするのは自分で書いた平たい JSON だけ
' なので、パーサは持たない。本文のような「引用符や改行が入りうるもの」は
' JSON へ入れず別ファイルに置く、という約束とセットで成り立っている。
Public Function JVal(ByVal js As String, ByVal key As String) As String
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

' Application.OnTime / OnKey へ渡す「このブックの手続き」の書き方。
' ブック名に ' が入っていても壊れないよう二重化する。
Public Function Qual(ByVal procName As String) As String
    Qual = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
End Function

' 経過秒。Timer は真夜中で 0 に戻るので、またいだときの負値をならす。
Public Function ElapsedSince(ByVal t0 As Double) As Double
    Dim v As Double
    v = Timer - t0
    If v < 0 Then v = v + 86400
    ElapsedSince = v
End Function
