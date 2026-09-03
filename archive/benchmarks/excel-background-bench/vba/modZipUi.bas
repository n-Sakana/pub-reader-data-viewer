Attribute VB_Name = "modZipUi"
'==============================================================================
' modZipUi -- ボタンから呼ばれる入口だけを集めた module。
'
' なぜ入口だけ別の module に分けてあるか
' ------------------------------------------------------------------
' Application.Run と図形の OnAction は、その手続きが入っている module の
' ディスパッチ面を通して呼ぶ。ところが module に「公開されたユーザー定義型を
' 使う公開手続き」が 1 つでもあると、その module はオートメーションに
' 公開できなくなり、同じ module の中のどの手続きも Application.Run から
' 呼べなくなる。
'
' ZipBench は 1 方式ぶんの計測結果を ZbResult という型で持ち回っていて、
' modZipBench・modZipVba・modZipWorker・modZipExcel はどれもその型を
' 引数か戻り値に使っている。だからそれらの module は Application.Run の
' 相手にできない。
'
' 実測した挙動 (Excel 16.0 / Win11 26200):
'     UDT を持たない module の Sub  -> Application.Run 成功
'     UDT を持つ module の Sub      -> 「マクロを実行できません。このブックで
'                                        マクロが使用できないか...」で失敗
'
' そこで、型を一切使わない薄い入口をここに置き、中身は各 module に委ねる。
' ボタンの OnAction はすべてこの module の名前を指す。
'==============================================================================
Option Explicit

'--- 動作確認 -----------------------------------------------------------------
' ボタンからも Application.Run からも本当に呼べるかを確かめる最小の入口。
' 何かがおかしいとき、まずこれが通るかどうかで切り分けられる。
Public Function ZB_Ping() As String
    ZB_Ping = "ZipBench ok " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function

Public Sub ZB_Nop()
End Sub

'--- 準備 ---------------------------------------------------------------------
Public Sub ZB_DownloadKenAll()
    modZipBench.ZbDownloadKenAll
End Sub

Public Sub ZB_Prepare()
    modZipBench.ZbPrepare
End Sub

Public Sub ZB_InitSheet()
    modZipBench.ZbInitSheet
End Sub

' Workbook_Open から呼ばれる。まだ何も描かれていないときだけ画面を作る。
' ビルド時にマクロを走らせないための入口 (xltoolrack の Build-Addin.ps1 が
' 同じ考え方で、ビルダは Application.Run を一切使わず、ブックは開かれたときに
' 自分で初期化する)。
Public Sub ZB_InitSheetIfEmpty()
    On Error Resume Next
    If Len(CStr(modZipBench.ZB_Sheet().Range("A1").Value)) = 0 Then
        modZipBench.ZbInitSheet
    End If
End Sub

'--- 8 方式 -------------------------------------------------------------------
' 1 行に詰めて書かないこと。「Sub X(): 処理: End Sub」の 1 行形式は
' コンパイルは通るが、その module 全体が Application.Run から呼べなくなる。
' 実測 (Excel 16.0 / Win11 26200): 1 行形式を 1 つでも含む module は、
' 同じ module の空の Sub すら「マクロを実行できません」で弾かれる。
Public Sub ZB_Run1()
    modZipBench.ZbRunMethod 1
End Sub

Public Sub ZB_Run2()
    modZipBench.ZbRunMethod 2
End Sub

Public Sub ZB_Run3()
    modZipBench.ZbRunMethod 3
End Sub

Public Sub ZB_Run4()
    modZipBench.ZbRunMethod 4
End Sub

Public Sub ZB_Run5()
    modZipBench.ZbRunMethod 5
End Sub

Public Sub ZB_Run6()
    modZipBench.ZbRunMethod 6
End Sub

Public Sub ZB_Run7()
    modZipBench.ZbRunMethod 7
End Sub

Public Sub ZB_Run8()
    modZipBench.ZbRunMethod 8
End Sub

'--- 結果の届け方の比較 (起動手段は両方とも WScript.Shell.Run で固定) ---------
Public Sub ZB_Run9()
    modZipBench.ZbRunMethod 9
End Sub

Public Sub ZB_Run10()
    modZipBench.ZbRunMethod 10
End Sub

'--- 出力フェッチ / 書戻し方式の比較 (起動は 13-15 とも WScript.Shell.Run で固定) ---
Public Sub ZB_Run13()
    modZipBench.ZbRunMethod 13
End Sub

Public Sub ZB_Run14()
    modZipBench.ZbRunMethod 14
End Sub

Public Sub ZB_Run15()
    modZipBench.ZbRunMethod 15
End Sub

Public Sub ZB_Run16()
    modZipBench.ZbRunMethod 16
End Sub

Public Sub ZB_RunAll()
    modZipBench.ZbRunAll
End Sub

'--- ファイル作成だけ (起動も変換もしない) ------------------------------------
Public Sub ZB_EmitOnlyPrebuilt()
    modZipWorker.ZbEmitOnlyPrebuilt
End Sub

Public Sub ZB_EmitOnlyEmitted()
    modZipWorker.ZbEmitOnlyEmitted
End Sub

'--- 取消と後始末 -------------------------------------------------------------
Public Sub ZB_Cancel()
    modZipBench.ZbCancel
End Sub

Public Sub ZB_Cleanup()
    modZipBench.ZbCleanup
End Sub
