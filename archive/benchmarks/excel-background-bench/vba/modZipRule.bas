Attribute VB_Name = "modZipRule"
'==============================================================================
' modZipRule -- ZipBench の共有基盤。
'
' ここに置くものは「全方式が同じであること」を保証しなければならないものだけ:
'   * 変換規則 (郵便番号の正規化と、辞書の作り方)
'   * 入力100万件の決定的な生成規則
'   * 正解 (オラクル) の作り方
'   * UTF-8 テキストの読み書き
'
' 変換規則 ZipBench rule v1
' ------------------------------------------------------------------
'  A. 辞書づくり (KEN_ALL.CSV -> zip_dict.csv)   ... 準備時に一度だけ
'       1. 町域 = KEN_ALL の第9フィールド (漢字町域)
'       2. 全角開き括弧 U+FF08 があれば、そこで町域を打ち切る
'       3. 残りが「場合」で終わるなら、それは日本郵便の注記行であって
'          地名ではないので町域ごと捨てる
'          (「以下に掲載がない場合」「◯◯の次に番地がくる場合」を全部拾える)
'       4. 住所 = 都道府県 + 市区町村 + 町域
'       5. 同じ郵便番号が複数行に出たら、ファイル順で最初の行を採用する
'  B. 変換 (郵便番号 -> 住所)                     ... 各方式が計測対象として行う
'       1. 入力を正規化する (NormalizeZip)
'       2. 辞書を引く
'       3. 無ければ "該当なし"
'
' A を準備時に一度だけ行い、全方式が同じ zip_dict.csv を読む。
' したがって「変換規則が方式ごとに違う」ことは原理的に起こらない。
' 方式ごとに実装が要るのは B の正規化と辞書引きだけで、正規化の文字集合は
' worker\ZipWorker.cs の ZipRule.NormalizeZip と 1 文字も違わない。
'==============================================================================
Option Explicit

'--- 変換規則の定数 -----------------------------------------------------------
Public Const ZB_NOTFOUND As String = "該当なし"

'==============================================================================
' 高分解能タイマー
'
' VBA の Timer は Single で秒を返す。昼どき (42,600 秒) の値では 1 ULP が
' 3.90625 ms あり、工程内訳を測るには粗すぎる。QueryPerformanceCounter を使う。
' 64bit のカウンタは Currency で受ける (Currency は 64bit 整数を 1/10000 倍で
' 持つ型なので、そのまま入れて差を周波数で割れば秒になる。倍率は約分される)。
'==============================================================================
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef c As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef f As Currency) As Long
#End If

' 関数の本体はこの下の宣言部が終わってから。VBA はモジュールレベルの宣言を
' 手続きより後ろに置けない。置くと「End Sub、End Function または End Property
' 以降には、コメントのみが記述できます」でプロジェクト全体が止まる。

Private Const CH_PAREN   As Long = &HFF08&   ' 全角開き括弧
Private Const CH_BA      As Long = &H5834&   ' 「場」
Private Const CH_AI      As Long = &H5408&   ' 「合」

'--- 入力生成の定数 (C# / PowerShell の参照実装と同じ) ------------------------
Private Const GEN_NOTFOUND_MOD  As Long = 997      ' i mod 997  = 996 -> 辞書に無い番号
Private Const GEN_NOTFOUND_HIT  As Long = 996
Private Const GEN_INVALID_MOD   As Long = 1499     ' i mod 1499 = 1498 -> そもそも数字でない
Private Const GEN_INVALID_HIT   As Long = 1498
Private Const GEN_SHAPES        As Long = 5

'==============================================================================
' 高分解能タイマーの入口 (宣言部より後ろ、手続きの先頭に置く)
'==============================================================================
Public Function QpcNow() As Currency
    QueryPerformanceCounter QpcNow
End Function

Public Function QpcSec(ByVal fromTick As Currency, ByVal toTick As Currency) As Double
    Dim f As Currency
    QueryPerformanceFrequency f
    If f = 0 Then Exit Function
    QpcSec = (toTick - fromTick) / f
End Function

Public Function QpcSince(ByVal fromTick As Currency) As Double
    QpcSince = QpcSec(fromTick, QpcNow())
End Function

'==============================================================================
' 郵便番号の正規化
'
' 全角数字は半角へ畳み、決められた区切り文字だけを捨て、それ以外の文字が
' 1 つでも混ざれば入力全体を不正とする。結果がちょうど半角数字 7 桁のときだけ
' その 7 桁を返し、それ以外は "" を返す。
'
' 100万回呼ばれるので Mid$/AscW の 1 文字ずつではなくバイト配列で回す。
' VBA の String は UTF-16 なので、b(i) と b(i+1) でコードユニットが作れる。
'==============================================================================
Public Function NormalizeZip(ByVal s As String) As String
    Dim b() As Byte
    Dim i As Long, n As Long, c As Long
    Dim out As String

    If LenB(s) = 0 Then NormalizeZip = "": Exit Function

    b = s
    out = "01234567"                      ' 8 文字ぶんの器 (7 桁を超えたら不正)

    For i = 0 To UBound(b) - 1 Step 2
        c = b(i) + b(i + 1) * 256&

        If c >= &HFF10& And c <= &HFF19& Then c = 48 + (c - &HFF10&)   ' 全角数字

        If c >= 48 And c <= 57 Then
            If n >= 8 Then NormalizeZip = "": Exit Function
            n = n + 1
            Mid$(out, n, 1) = ChrW$(c)
        Else
            Select Case c
                Case &H2D&, &H2010&, &H2012&, &H2013&, &H2014&, &H2015&, _
                     &H2212&, &HFF0D&, &H30FC&, &H3012&, _
                     &H20&, &H3000&, &H9&, &HD&, &HA&
                    ' 区切り・空白・〒 は捨てる
                Case Else
                    NormalizeZip = ""
                    Exit Function
            End Select
        End If
    Next i

    If n = 7 Then NormalizeZip = Left$(out, 7) Else NormalizeZip = ""
End Function

'==============================================================================
' 変換の中身。方式2 (高速VBA) と 方式7 (別Excel) は、この 2 つの関数を
' そのまま共有する。同じファイルの同じ関数なので、アルゴリズム・辞書・
' ループが食い違う余地がない。違うのは「どのプロセスで走るか」だけ。
'
' どちらにも DoEvents は入れない。片方だけ 50 回 DoEvents する、といった
' 非対称を作らないため。取消は Application.EnableCancelKey (Esc) で拾う。
'==============================================================================

' MASTER シート (1列目:郵便番号, 2列目:住所) を 1 回で読んだ 2 次元配列から
' 辞書を作る。同じ郵便番号が複数あれば先勝ち。
Public Function BuildDictFromMaster(ByRef master As Variant, ByVal n As Long) As Object
    Dim d As Object
    Dim i As Long
    Dim k As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0                     ' vbBinaryCompare
    For i = 1 To n
        k = CStr(master(i, 1))
        If Len(k) > 0 Then
            If Not d.Exists(k) Then d.Add k, CStr(master(i, 2))
        End If
    Next i
    Set BuildDictFromMaster = d
End Function

' Input の 2 次元配列を変換して、出力の 2 次元配列を作る。
' 出力はそのまま Range.Value へ渡せる形 (1 To n, 1 To 1)。
Public Sub ConvertBlock(ByRef src As Variant, ByVal d As Object, _
                        ByRef dst() As String, ByVal n As Long)
    Dim i As Long
    Dim k As String

    ReDim dst(1 To n, 1 To 1)
    For i = 1 To n
        k = NormalizeZip(CStr(src(i, 1)))
        If Len(k) = 7 Then
            If d.Exists(k) Then
                dst(i, 1) = d(k)
            Else
                dst(i, 1) = ZB_NOTFOUND
            End If
        Else
            dst(i, 1) = ZB_NOTFOUND
        End If
    Next i
End Sub

'==============================================================================
' 辞書づくり: KEN_ALL.CSV -> zip_dict.csv
'
' KEN_ALL は Shift-JIS で、引用符の中にカンマを含まない (全行 15 フィールドで
' あることを確認済み)。だから素朴な Split(",") で正しく割れる。
' 戻り値は書き出した郵便番号の件数。
'==============================================================================
Public Function BuildDictionary(ByVal kenAllPath As String, ByVal outDictPath As String) As Long
    Dim raw As String, lines() As String
    Dim i As Long, p As Long
    Dim f() As String
    Dim zip As String, town As String, addr As String
    Dim seen As Object
    Dim outArr() As String
    Dim n As Long

    raw = ReadTextShiftJis(kenAllPath)
    raw = Replace$(raw, vbCrLf, vbLf)
    lines = Split(raw, vbLf)

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 0                      ' vbBinaryCompare
    ReDim outArr(0 To UBound(lines))

    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            f = Split(lines(i), ",")
            If UBound(f) >= 8 Then
                zip = Unquote(f(2))
                If Not seen.Exists(zip) Then
                    seen.Add zip, 1
                    town = Unquote(f(8))

                    ' 規則 2: 全角開き括弧から後ろを捨てる
                    p = InStr(1, town, ChrW$(CH_PAREN), vbBinaryCompare)
                    If p > 0 Then town = Left$(town, p - 1)

                    ' 規則 3: 「場合」で終わるなら注記行なので町域ごと捨てる
                    If Len(town) >= 2 Then
                        If Right$(town, 2) = ChrW$(CH_BA) & ChrW$(CH_AI) Then town = ""
                    End If

                    addr = Unquote(f(6)) & Unquote(f(7)) & town
                    outArr(n) = zip & "," & addr
                    n = n + 1
                End If
            End If
        End If
    Next i

    ReDim Preserve outArr(0 To n - 1)
    WriteTextUtf8 outDictPath, Join(outArr, vbCrLf) & vbCrLf
    BuildDictionary = n
End Function

'==============================================================================
' 辞書の読み込み。
'   戻り値      Scripting.Dictionary (郵便番号 -> 住所)
'   zips()      郵便番号を KEN_ALL の順で並べた 0 起点配列 (入力生成に使う)
' 各方式が計測区間の中で呼ぶので、余計なことは一切しない。
'==============================================================================
Public Function LoadDictionary(ByVal dictPath As String, Optional ByRef zips As Variant) As Object
    Dim raw As String, lines() As String
    Dim i As Long, c As Long, n As Long
    Dim d As Object
    Dim k As String
    Dim keys() As String
    Dim wantKeys As Boolean

    wantKeys = Not IsMissing(zips)

    raw = ReadTextUtf8(dictPath)
    raw = Replace$(raw, vbCrLf, vbLf)
    lines = Split(raw, vbLf)

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0
    If wantKeys Then ReDim keys(0 To UBound(lines))

    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then
            c = InStr(1, lines(i), ",", vbBinaryCompare)
            If c > 1 Then
                k = Left$(lines(i), c - 1)
                If Not d.Exists(k) Then
                    d.Add k, Mid$(lines(i), c + 1)
                    If wantKeys Then keys(n) = k
                    n = n + 1
                End If
            End If
        End If
    Next i

    If wantKeys Then
        ReDim Preserve keys(0 To n - 1)
        zips = keys
    End If
    Set LoadDictionary = d
End Function

'==============================================================================
' 入力の決定的生成。
' 乱数を使わない。同じ件数なら常に同じ内容になるので、方式をまたいだ比較でも、
' 日を跨いだ再測定でも、他言語の参照実装 (build\make_input_reference.ps1) でも
' 同じバイト列が出る。
'
'   i mod 997  = 996  -> "9999999"  辞書に無い番号
'   i mod 1499 = 1498 -> "ABC-DEFG" 数字ですらない
'   それ以外は zips(i mod M) を 5 通りの書式のどれかで出す
'       0: 1000001        3桁+4桁そのまま
'       1: 100-0001       半角ハイフン入り
'       2: 〒100-0001     郵便記号つき
'       3: １０００００１ 全角数字
'       4: "1000001 "     末尾に半角スペース
' 5 通りあるのは正規化を実際に働かせるため。逃げ道 2 つは「該当なし」を
' 必ず通すため。どの方式もここで手を抜けない。
'==============================================================================
Public Function GenerateInput(ByVal n As Long, ByRef zips As Variant) As Variant
    Dim arr() As String
    Dim i As Long, m As Long, j As Long
    Dim z As String, t As String

    m = UBound(zips) - LBound(zips) + 1
    If m <= 0 Then Err.Raise 5, "GenerateInput", "empty dictionary"

    ReDim arr(1 To n)
    For i = 0 To n - 1
        If (i Mod GEN_NOTFOUND_MOD) = GEN_NOTFOUND_HIT Then
            arr(i + 1) = "9999999"
        ElseIf (i Mod GEN_INVALID_MOD) = GEN_INVALID_HIT Then
            arr(i + 1) = "ABC-DEFG"
        Else
            z = zips(LBound(zips) + (i Mod m))
            Select Case i Mod GEN_SHAPES
                Case 0
                    arr(i + 1) = z
                Case 1
                    arr(i + 1) = Left$(z, 3) & "-" & Mid$(z, 4)
                Case 2
                    arr(i + 1) = ChrW$(&H3012&) & Left$(z, 3) & "-" & Mid$(z, 4)
                Case 3
                    t = z
                    For j = 1 To 7
                        Mid$(t, j, 1) = ChrW$(&HFF10& + (Asc(Mid$(z, j, 1)) - 48))
                    Next j
                    arr(i + 1) = t
                Case 4
                    arr(i + 1) = z & " "
            End Select
        End If
    Next i
    GenerateInput = arr
End Function

'==============================================================================
' 正解 (オラクル)。全方式の結果はこれと 1 件ずつ突き合わせる。
' inputs は 1 起点、戻り値も 1 起点。
'==============================================================================
Public Function BuildExpected(ByRef inputs As Variant, ByVal d As Object) As Variant
    Dim arr() As String
    Dim i As Long, n As Long
    Dim k As String

    n = UBound(inputs)
    ReDim arr(1 To n)
    For i = 1 To n
        k = NormalizeZip(inputs(i))
        If Len(k) = 7 Then
            If d.Exists(k) Then arr(i) = d(k) Else arr(i) = ZB_NOTFOUND
        Else
            arr(i) = ZB_NOTFOUND
        End If
    Next i
    BuildExpected = arr
End Function

'==============================================================================
' 突き合わせ。一致なら 0、不一致ならその最初の行番号 (1 起点)、
' 件数違いなら -1 を返す。全方式で同じ関数を使う。
'==============================================================================
Public Function CompareArrays(ByRef got As Variant, ByRef want As Variant, ByVal n As Long) As Long
    Dim i As Long

    If UBound(got) - LBound(got) + 1 < n Then CompareArrays = -1: Exit Function
    If UBound(want) - LBound(want) + 1 < n Then CompareArrays = -1: Exit Function

    Dim gOff As Long, wOff As Long
    gOff = LBound(got) - 1
    wOff = LBound(want) - 1
    For i = 1 To n
        If got(gOff + i) <> want(wOff + i) Then CompareArrays = i: Exit Function
    Next i
    CompareArrays = 0
End Function

'==============================================================================
' テキスト入出力
'
' 出力は BOM なしの UTF-8。C# 側は BOM 検出を有効にしてあるので付いていても
' 読めるが、付けない方が「Excel が書いた行数 = ファイルの行数」が素直になる。
'==============================================================================
Public Sub WriteTextUtf8(ByVal path As String, ByVal s As String)
    Dim st As Object, bin As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                      ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText s

    ' ADODB.Stream は必ず BOM を書くので、バイナリに切り替えて先頭 3 バイトを捨てる
    st.Position = 0
    st.Type = 1                      ' adTypeBinary
    st.Position = 3

    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    st.CopyTo bin
    bin.SaveToFile path, 2           ' adSaveCreateOverWrite
    bin.Close
    st.Close
End Sub

' BOM つき UTF-8。csc.exe に渡す .cs だけはこちらで書く
' (BOM があれば csc は文字コードを取り違えようがない)。
Public Sub WriteTextUtf8Bom(ByVal path As String, ByVal s As String)
    Dim st As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.SaveToFile path, 2
    st.Close
End Sub

' .bat / .ps1 は ASCII のみ。cmd.exe と Windows PowerShell 5.1 が
' BOM なしファイルを OEM/ANSI で読むので、非 ASCII を書くと壊れる。
Public Sub WriteTextAscii(ByVal path As String, ByVal s As String)
    Dim st As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "us-ascii"
    st.Open
    st.WriteText s
    st.SaveToFile path, 2
    st.Close
End Sub

Public Function ReadTextUtf8(ByVal path As String) As String
    Dim st As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile path
    ReadTextUtf8 = st.ReadText(-1)
    st.Close
End Function

Public Function ReadTextShiftJis(ByVal path As String) As String
    Dim st As Object

    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "shift_jis"
    st.Open
    st.LoadFromFile path
    ReadTextShiftJis = st.ReadText(-1)
    st.Close
End Function

'==============================================================================
' 出力ファイル -> 1 起点配列。照合のためだけに使う (計測区間の外)。
' 末尾の空行は数えない。
'==============================================================================
Public Function ReadOutputLines(ByVal path As String, ByRef rowCount As Long) As Variant
    Dim raw As String, parts() As String
    Dim i As Long, n As Long
    Dim arr() As String

    raw = ReadTextUtf8(path)
    raw = Replace$(raw, vbCrLf, vbLf)
    parts = Split(raw, vbLf)

    n = UBound(parts) + 1
    Do While n > 0
        If Len(parts(n - 1)) > 0 Then Exit Do
        n = n - 1
    Loop

    If n = 0 Then
        rowCount = 0
        ReDim arr(1 To 1)
        ReadOutputLines = arr
        Exit Function
    End If

    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = parts(i - 1)
    Next i
    rowCount = n
    ReadOutputLines = arr
End Function

'==============================================================================
' 小物
'==============================================================================
Private Function Unquote(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) >= 2 Then
        If Left$(t, 1) = """" And Right$(t, 1) = """" Then t = Mid$(t, 2, Len(t) - 2)
    End If
    Unquote = t
End Function

Public Sub EnsureFolder(ByVal p As String)
    Dim fso As Object
    Dim parts() As String, i As Long, cur As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    parts = Split(Replace$(p, "/", "\"), "\")
    cur = parts(0)
    For i = 1 To UBound(parts)
        cur = cur & "\" & parts(i)
        If Not fso.FolderExists(cur) Then fso.CreateFolder cur
    Next i
End Sub

Public Function FileExists(ByVal p As String) As Boolean
    FileExists = (Len(Dir$(p, vbNormal)) > 0)
End Function
