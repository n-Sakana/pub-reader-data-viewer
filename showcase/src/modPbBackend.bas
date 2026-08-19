Attribute VB_Name = "modPbBackend"
'==============================================================================
' modPbBackend - BE 側（同じコードが、コピーされたブックの中では BE として動く）
'
' 責務：FE から起こされ、指示を待ち、円周率を実際に計算して桁を置く。画面は
' 持たない（このプロセスは不可視）。FE の画面は一切触らない。
'
' 入口が標準モジュールにあるのは、FE が `Application.Run "'book'!PbBeBootstrap"`
' と **名前で** 呼ぶから。名前で呼べば失敗が COM のエラーとして返り、FE 側で
' 受け止められる。入口は OnTime を仕込むだけなのですぐ返る。
'
' 開発中、ここで仕掛けた OnTime が一度も発火せず、不可視インスタンスの制約かと
' 疑ったことがある。違った。原因は調べ方のほうだった。**自動化で起こした Excel の
' VBE を触ると、そのインスタンスは壊れる。** 以後 OnTime は Err.Number = 0 の
' まま沈黙し、常駐ループも 1 周で無応答になる。VBE に触らずに同じブックを
' 走らせれば、この入口は 8ms で返り、常駐ループは idle → running → bye まで通る。
' 教訓：不可視の Excel を測るとき、VBE には触らない。
'==============================================================================
Option Explicit

Private m_comms As PbBenchmarkRun
Private m_quit As Boolean
Private m_lastSeq As Long
Private m_feMissing As Long
Private m_feSeen As String

' やりとりの置き場所と読み書きは FE 側と同じ約束でなければ噛み合わないので、
' 同じクラスをこちらでも作る。BE が使うのはその Temp のやりとりの部分だけで、
' 表示用 Excel も計測も動かさない（Init に記録係だけ渡すのはそのため）。
Private Function Comms() As PbBenchmarkRun
    Dim lg As PbLog
    If m_comms Is Nothing Then
        Set lg = New PbLog
        Set m_comms = New PbBenchmarkRun
        m_comms.Init lg
    End If
    Set Comms = m_comms
End Function

'------------------------------------------------------------------ 入口
Public Function PbBeBootstrap() As String
    On Error Resume Next
    ' FE は返った直後に COM 参照を捨てる。参照ゼロの自動化サーバーは次の
    ' メッセージポンプで自分を畳むので、自前所有だと宣言しておく。
    Application.UserControl = True
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    m_quit = False
    m_lastSeq = 0
    ' FE がまだフォルダを作っていないことがある。
    Comms().EnsureFolder
    Application.OnTime Now, Qual("PbBeMain")
    PbBeBootstrap = "armed"
End Function

' 常駐ループ。ここは不可視なので、待つのにブロックして構わない。
Public Sub PbBeMain()
    Dim js As String
    Dim seq As Long
    Dim feNow As String

    On Error Resume Next
    Comms().EnsureFolder
    BeWrite "{""state"":""idle"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Do
        If m_quit Then Exit Do
        js = Comms().ReadAll(Comms().CmdPath)
        seq = CLng(Val(JVal(js, "seq")))
        If seq > m_lastSeq Then
            m_lastSeq = seq
            Select Case JVal(js, "cmd")
                Case "pi"
                    BePiRun Val(JVal(js, "secs")), seq
                Case "quit"
                    Exit Do
            End Select
        End If
        ' FE が消えたら自分から終わる。置き去りの BE を残さない。
        '
        ' 「fe.json が無い」だけを見ていたのでは足りない。FE が正常に閉じた
        ' ときは消えるが、強制終了されるとファイルは残ったままになり、それを
        ' 「FE は生きている」と読んで永遠に回り続ける。だから中身が更新されて
        ' いるかで見る。待つ長さは長めに取る。FE は正常でも黙る時間があり
        ' （起動直後の最初の再描画は 1px なら数十秒）、短く見切ると、その間に
        ' BE が自分で終わってしまう。
        feNow = Comms().ReadAll(Comms().FePath)
        If Len(feNow) = 0 Or feNow = m_feSeen Then
            m_feMissing = m_feMissing + 1
            If m_feMissing > 90 Then Exit Do
        Else
            m_feMissing = 0
        End If
        m_feSeen = feNow
        Application.Wait Now + TimeSerial(0, 0, 1)
    Loop
    BeQuit
End Sub

'------------------------------------------------------------------ 円周率
' 円周率を **実際に計算する**。既知の桁を流し込むことは 1 桁もしない。
'
' 使うのは Rabinowitz-Wagon の spigot（有限桁版）。桁を頭から 1 つずつ吐くので
' 「増えていく」様子がそのまま出せる。配列の長さ ln が「何桁まで正しく出せるか」
' を決めるので、**途中で打ち切っても出した桁は正しい**（足りなくなるのではなく、
' まだ吐いていないだけ）。
'
' ln = 10n/3 + 1 で、1 桁あたり内側ループが ln 回まわるから総コストは約 10n^2/3。
' 端末の速さは決め打ちできない（開発機で決めた回数が非力なノートで 14 倍かかった
' 実測がある）ので、走る前に 0.15 秒だけ同じ内側ループを空回しして速さを測り、
' 与えられた秒数で届く n を選ぶ。1.15 倍だけ多めに取って、打ち切りの瞬間まで
' 手が空かないようにする。
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
    If secs <= 0 Then secs = PB_BENCHSECS
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
    ' 何桁ごとに桁と進捗を置くか。**FE は 1 秒に 1 回しか読まないので、ここが
    ' 1 秒あたり 2 回を割ると、FE から見た「その 1 秒に届いた桁」が 1 回ぶんか
    ' 2 回ぶんかで揺れる。** 実測：n\60（1 秒あたり 1.7 回）だと計器の目盛が
    ' 高い・低いを交互に繰り返した。1 秒あたり 5 回まで上げるとその折り返しは
    ' 消え、増えた書き込みの損は 30 秒のうち 1% 未満で済む。
    chunk = n \ 180
    If chunk < 5 Then chunk = 5
    t0 = Timer
    lastDv = t0
    BeWrite "{""state"":""running"",""n"":0,""target"":" & n & ",""rate"":" & _
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
        ' 9 が続くあいだは繰り上がりが確定しないので、吐かずに数えておく。
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
        ' なる。ただし DoEvents は 1 回 1.6ms かかるので毎桁は高すぎる（1 桁の
        ' 計算より重くなる端末がある）。時間で間隔を決めて 30ms に 1 回にすると、
        ' 応答は保てて損は数 % で済む。
        ms = Timer - lastDv
        If ms < 0 Then ms = ms + 86400
        If ms >= 0.03 Then
            DoEvents
            lastDv = Timer
        End If

        If (j Mod chunk) = 0 Or j = n Then
            ms = ElapsedSince(t0) * 1000
            Comms().WriteAll Comms().PiPath, out
            BeWrite "{""state"":""running"",""n"":" & Len(out) & ",""target"":" & n & _
                ",""ms"":" & CLng(ms) & ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
            If ms >= secs * 1000 Then Exit For
            ' FE から新しい指示（quit）が来ていたら、そこで畳む。
            js = Comms().ReadAll(Comms().CmdPath)
            If CLng(Val(JVal(js, "seq"))) > stopSeq Then Exit For
        End If
    Next j

    ms = ElapsedSince(t0) * 1000
    Comms().WriteAll Comms().PiPath, out
    BeWrite "{""state"":""done"",""n"":" & Len(out) & ",""target"":" & n & _
        ",""ms"":" & CLng(ms) & ",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Exit Sub
Failed:
    BeWrite "{""state"":""error"",""msg"":""" & Err.Number & " " & _
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
        el = ElapsedSince(t0)
    Loop While el < 0.15
    If el <= 0 Then el = 0.15
    BeInnerRate = loops / el
End Function

'------------------------------------------------------------------ 終わり方
' 順序が効く。ThisWorkbook.Close を先に呼ぶと、そこで自分のコードが止まって
' Application.Quit に届かず、ブックを持たない不可視 Excel が残る。
Private Sub BeQuit()
    On Error Resume Next
    m_quit = True
    BeWrite "{""state"":""bye"",""ts"":""" & Format$(Now, "hh:nn:ss") & """}"
    Application.DisplayAlerts = False
    ThisWorkbook.Saved = True
    Set m_comms = Nothing
    Application.Quit
End Sub

Private Sub BeWrite(ByVal s As String)
    Comms().WriteAll Comms().ProgPath, s
End Sub
