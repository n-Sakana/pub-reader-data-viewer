Attribute VB_Name = "modPbPi"
'==============================================================================
' modPbPi -- VERIFICATION ONLY. v3 モックの piDigits を VBA に移して測る。
'
' モックは 804 桁（67 桁 x 12 行）を、1 桁ごとに先頭から計算し直す（最適化なし
' ＝要件どおり意図的に重い）。同じ書き方が VBA で何秒になるかは未知なので、
' 桁数を決める前に測る。FE が固まる時間として実演に耐える長さ（十秒前後）に
' したい。
'==============================================================================
Option Explicit

Private m_res As String

Private Sub Say(ByVal s As String)
    Dim f As Integer
    Dim i As Long
    On Error Resume Next
    For i = 1 To 60
        Err.Clear
        f = FreeFile
        Open m_res For Append As #f
        If Err.Number = 0 Then
            Print #f, s
            Close #f
            Exit For
        End If
        DoEvents
    Next i
End Sub

Public Sub PbPiPing()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "PONG"
End Sub

Public Sub PbPiArm()
    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Application.OnTime Now, "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!PbPiRun"
End Sub

' モックと同じ spigot。k 桁目だけを、毎回まっさらから求める。
Public Function PiDigit(ByVal k As Long) As Long
    Dim m As Long
    Dim ln As Long
    Dim a() As Long
    Dim nines As Long
    Dim pre As Long
    Dim first As Boolean
    Dim n As Long
    Dim got As Long
    Dim j As Long
    Dim i As Long
    Dim q As Long
    Dim x As Long
    Dim z As Long

    m = k + 2
    ln = (10 * m) \ 3 + 4
    ReDim a(0 To ln - 1)
    For i = 0 To ln - 1
        a(i) = 2
    Next i
    first = True
    For j = 1 To m + 3
        If got >= k Then Exit For
        q = 0
        For i = ln To 1 Step -1
            x = 10 * a(i - 1) + q * i
            a(i - 1) = x Mod (2 * i - 1)
            q = x \ (2 * i - 1)
        Next i
        a(0) = q Mod 10
        q = q \ 10
        If q = 9 Then
            nines = nines + 1
        ElseIf q = 10 Then
            If first Then
                first = False
            Else
                n = pre + 1
                got = got + 1
            End If
            For z = 1 To nines
                If got >= k Then Exit For
                n = 0
                got = got + 1
            Next z
            pre = 0
            nines = 0
        Else
            If first Then
                first = False
            Else
                n = pre
                got = got + 1
            End If
            For z = 1 To nines
                If got >= k Then Exit For
                n = 9
                got = got + 1
            Next z
            nines = 0
            pre = q
        End If
    Next j
    PiDigit = n
End Function

Public Sub PbPiRun()
    Dim t0 As Double
    Dim ms As Double
    Dim k As Long
    Dim d As Long
    Dim s As String
    Dim marks As Variant
    Dim mi As Long

    m_res = CStr(ThisWorkbook.Worksheets(1).Range("A1").Value)
    Say "=== pi probe start ==="

    ' 正しさの確認：最初の 20 桁は 14159265358979323846
    s = ""
    For k = 1 To 20
        s = s & CStr(PiDigit(k))
    Next k
    Say "first20 = " & s & "  (expect 14159265358979323846)"

    ' 累積の所要時間。67 桁 = 1 行ぶんの区切りで測る。
    marks = Array(67, 134, 201, 268, 335, 402, 536, 670, 804)
    mi = 0
    t0 = Timer
    For k = 1 To 804
        d = PiDigit(k)
        If mi <= UBound(marks) Then
            If k = marks(mi) Then
                ms = (Timer - t0) * 1000
                If ms < 0 Then ms = ms + 86400000#
                Say "digits " & Format$(k, "@@@@") & "  = " & Format$(ms, "0") & " ms  (" & _
                    Format$(ms / k, "0.0") & " ms/digit)"
                mi = mi + 1
            End If
        End If
    Next k
    Say "=== pi probe done ==="
End Sub
