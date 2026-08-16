# ============================================================================
# build_workbooks2.ps1 -- the two one-to-many workbooks, from src\v2\vba.
#
#   dist\ReaderDataViewer-VBA-Hash.xlsm   hand-built hash index
#   dist\ReaderDataViewer-VBA-Dict.xlsm   Scripting.Dictionary, late bound
#
# The two books differ by exactly one imported file. clsRdv2IdxHash.cls and
# clsRdv2IdxDict.cls both declare a class called clsRdv2Idx with the same four
# members, so modRdv2Engine says New clsRdv2Idx and never knows which it got.
#
# It does not touch dist\ReaderDataViewer-VBA.xlsm or the hybrid book: those are
# the frozen 1:1 build and build_workbooks.ps1 still owns them.
#
# Same four traps as build_workbooks.ps1, checked the same way before Excel is
# ever opened: CP932 transcoding, trust access to the VBA project, references
# added after the imports, and VBA's three silent compile traps. Plus the
# inline-hash check, which now covers the two copies inside clsRdv2IdxHash.
#
# The candidate dialog is a UserForm, built here through the VBIDE designer
# rather than shipped as a file, because a .frm carries a binary .frx twin that
# would not survive being kept in a text repository.
#
# It only ever touches the Excel instance it created itself.
# ============================================================================
[CmdletBinding()]
param([string] $Root = "")
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'excel_own.ps1')   # exact Excel ownership, never a pid diff
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }

$xlWBATWorksheet = -4167
$xlOpenXMLWorkbookMacroEnabled = 52

$buildLog = Join-Path $Root 'work\build_workbooks2.log'
$workDir = Split-Path -Parent $buildLog
if (-not (Test-Path -LiteralPath $workDir)) { New-Item -ItemType Directory -Path $workDir | Out-Null }
Set-Content -LiteralPath $buildLog -Value '' -Encoding utf8 -Force
function Step([string] $m) {
  $line = "{0}  {1}" -f (Get-Date -Format 'HH:mm:ss.fff'), $m
  Add-Content -LiteralPath $buildLog -Value $line -Encoding utf8
  Write-Output $line
}

$modules = @('modRdv2Spec', 'modRdv2Uia', 'modRdv2Engine', 'modRdv2App')
$classes = @('clsRdv2IdxHash', 'clsRdv2IdxDict')
$src = Join-Path $Root 'src\v2\vba'

function Src([string] $name) {
  if ($classes -contains $name) { return (Join-Path $src "$name.cls") }
  return (Join-Path $src "$name.bas")
}

# --- preflight: trust -------------------------------------------------------
$vbom = 0
foreach ($v in @('16.0', '15.0', '14.0')) {
  $p = "HKCU:\Software\Microsoft\Office\$v\Excel\Security"
  if (Test-Path $p) { $x = (Get-ItemProperty $p).AccessVBOM; if ($x) { $vbom = $x } }
}
if ($vbom -ne 1) {
  throw @"
Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が無効です。
  Excel > ファイル > オプション > トラスト センター > トラスト センターの設定
        > マクロの設定 > VBA プロジェクト オブジェクト モデルへのアクセスを信頼する
にチェックを入れてから、もう一度このスクリプトを実行してください。
"@
}

# --- preflight: name collisions, one-liners, declaration order ---------------
$declRe = '^\s*(?:Public\s+|Private\s+|Global\s+)?(?:Static\s+)?(?:Declare\s+(?:PtrSafe\s+)?)?(?:Sub|Function|Property\s+(?:Get|Let|Set)|Const|Type|Enum)\s+([A-Za-z_][A-Za-z0-9_]*)'
$varRe = '(?<![A-Za-z0-9_])([A-Za-z_][A-Za-z0-9_]*)\s*(?:\([^)]*\))?\s+As\s'
$collisions = @()
$oneLiners = @()
$misplaced = @()
foreach ($m in ($modules + $classes)) {
  $seen = @{}
  $ln = 0
  $seenProc = $false
  foreach ($line in [IO.File]::ReadAllLines((Src $m))) {
    $ln++
    if ($line -match "^\s*'") { continue }
    $names = @()
    if ($line -match $declRe) { $names += $Matches[1] }
    if ($line -match '(?i)^\s*(Public|Private|Global|Dim)\s' -and
        $line -notmatch '(?i)^\s*(Public|Private|Global)?\s*(Declare|Sub|Function|Property|Const|Type|Enum)\b') {
      foreach ($mm in [regex]::Matches($line, $varRe)) { $names += $mm.Groups[1].Value }
    }
    foreach ($name in $names) {
      $k = $name.ToLowerInvariant()
      if ($seen.ContainsKey($k) -and $seen[$k] -ne $name) {
        $collisions += "$m : '$($seen[$k])' と '$name' は VBA では同じ名前 (大文字小文字を区別しない)"
      }
      $seen[$k] = $name
    }
    if ($line -match '(?i)^\s*(Public|Private|Friend)?\s*(Sub|Function|Property)\b.*:\s*End\s+(Sub|Function|Property)\s*$') {
      $oneLiners += "$m line ${ln}: $($line.Trim())"
    }
    if ($line -match '(?i)^\s*(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property)\s') { $seenProc = $true; continue }
    if (-not $seenProc) { continue }
    if ($line -match '(?i)^\s*(Public|Private|Global)\s+(Const|Type|Declare|WithEvents)\s' -or
        $line -match '(?i)^\s*(Public|Private|Global)\s+[A-Za-z_][A-Za-z0-9_]*\s+As\s' -or
        $line -match '(?i)^\s*Declare\s') {
      $misplaced += "$m line ${ln}: $($line.Trim())"
    }
  }
}
if ($collisions.Count) { throw ("VBA の名前衝突:`n  " + ($collisions -join "`n  ")) }
if ($oneLiners.Count) { throw ("1 行に詰めた手続き定義:`n  " + ($oneLiners -join "`n  ")) }
if ($misplaced.Count) { throw ("宣言が手続きより後ろ:`n  " + ($misplaced -join "`n  ")) }
Write-Output 'name / one-line / declaration-order preflight: ok'

# --- preflight: the two class members must match across the flavours ---------
function Signatures([string] $path) {
  $sig = @()
  foreach ($l in [IO.File]::ReadAllLines($path)) {
    if ($l -match '(?i)^\s*Public\s+(Sub|Function|Property\s+Get)\s+') { $sig += $l.Trim() }
  }
  return $sig
}
$sigH = Signatures (Src 'clsRdv2IdxHash')
$sigD = Signatures (Src 'clsRdv2IdxDict')
if (($sigH -join '|') -ne ($sigD -join '|')) {
  throw ("索引クラス 2 種の公開メンバーが一致しません:`n  hash: " + ($sigH -join ' / ') + "`n  dict: " + ($sigD -join ' / '))
}
foreach ($c in $classes) {
  $nm = ([IO.File]::ReadAllLines((Src $c)) | Where-Object { $_ -match '^Attribute VB_Name' })
  if ($nm -notmatch 'clsRdv2Idx') { throw "$c の VB_Name が clsRdv2Idx ではありません: $nm" }
}
Write-Output ("index-class preflight: ok (" + $sigH.Count + " shared members)")

# --- preflight: the inlined hash still matches the canonical one -------------
function Normalize-Hash([string[]] $lines) {
  $out = @()
  foreach ($l in $lines) {
    $t = $l.Trim()
    if ($t.Length -eq 0) { continue }
    if ($t.StartsWith("'")) { continue }
    $t = [regex]::Replace($t, '(?<![A-Za-z0-9_])[A-Za-z_][A-Za-z0-9_]*\(p', 'ARR(p')
    $out += $t
  }
  return ($out -join '|')
}
$canon = @()
$inFn = $false
foreach ($l in [IO.File]::ReadAllLines((Src 'modRdv2Spec'))) {
  if ($l -match '^\s*Public Function Rdv2HashBytes') { $inFn = $true; continue }
  if ($inFn) {
    if ($l -match '^\s*End Function') { break }
    if ($l -match '^\s*Dim ' -or $l -match '^\s*Rdv2HashBytes\s*=') { continue }
    $canon += $l
  }
}
$canonKey = Normalize-Hash $canon
if ($canonKey.Length -lt 50) { throw 'canonical Rdv2HashBytes body not found' }
$copies = 0; $block = @(); $inBlock = $false
foreach ($l in [IO.File]::ReadAllLines((Src 'clsRdv2IdxHash'))) {
  if ($l -match "Rdv2HashBytes, inlined") { $inBlock = $true; $block = @(); continue }
  if ($inBlock) {
    if ($l -match 'end inline') {
      $inBlock = $false
      $copies++
      if ((Normalize-Hash $block) -ne $canonKey) {
        throw "clsRdv2IdxHash の inline ハッシュ #$copies が modRdv2Spec.Rdv2HashBytes と一致しません"
      }
      continue
    }
    $block += $l
  }
}
if ($copies -ne 2) { throw "inline ハッシュは 2 か所のはずが $copies か所でした" }
Write-Output "inline-hash preflight: ok ($copies copies identical to Rdv2HashBytes)"

# --- transcode UTF-8 -> CP932 ----------------------------------------------
$stage = Join-Path $env:TEMP ("rdv2_bas_" + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $stage | Out-Null
$cp932 = [Text.Encoding]::GetEncoding(932)
foreach ($m in ($modules + $classes)) {
  $p = Src $m
  $t = [IO.File]::ReadAllText($p)
  if ($t -ne $cp932.GetString($cp932.GetBytes($t))) {
    throw "$p has characters that cannot survive CP932; VBE would import them as '?'"
  }
  # CRLF, always: VBE only recognises the VERSION/BEGIN/END header of a .cls
  # when the lines end in CRLF. With LF it imports the header as code and the
  # project fails to compile at line 1, behind a modal dialog.
  $t = $t -replace "`r`n", "`n"
  $t = $t -replace "`n", "`r`n"
  $ext = [IO.Path]::GetExtension($p)
  [IO.File]::WriteAllText((Join-Path $stage "$m$ext"), $t, $cp932)
}
Step "staged modules as CP932 in $stage"

# --- the shared screen ------------------------------------------------------
function Paint-View2($ws, [string] $method) {
  $ws.Name = 'VIEW'
  $ws.Cells.Font.Name = 'Yu Gothic UI'
  $ws.Cells.Font.Size = 10

  $ws.Columns('A').ColumnWidth = 2
  $ws.Columns('B').ColumnWidth = 15
  $ws.Columns('C').ColumnWidth = 26
  $ws.Columns('D').ColumnWidth = 15
  $ws.Columns('E').ColumnWidth = 26
  $ws.Columns('F').ColumnWidth = 15
  $ws.Columns('G').ColumnWidth = 26
  $ws.Columns('H').ColumnWidth = 15
  $ws.Columns('I').ColumnWidth = 14
  $ws.Columns('J').ColumnWidth = 26
  $ws.Columns('K').ColumnWidth = 13
  $ws.Columns('L').ColumnWidth = 9

  $ws.Range('B1').Value2 = 'Reader Data Viewer'
  $ws.Range('B1').Font.Size = 18
  $ws.Range('B1').Font.Bold = $true
  $ws.Range('E1').Value2 = $method
  $ws.Range('E1').Font.Size = 12
  $ws.Range('E1').Font.Bold = $true
  $ws.Range('E1').Font.Color = 8210719

  $ws.Range('B2').Value2 = '状態'
  $ws.Range('B3').Value2 = 'メモ帳'
  $ws.Range('B4').Value2 = 'データ'
  $ws.Range('B5').Value2 = '索引方式'
  $ws.Range('B2:B5').Font.Color = 8210719
  $ws.Range('C2').Font.Bold = $true
  $ws.Range('C2').Font.Size = 12

  $ws.Range('E2').Value2 = '手動実行の番号'
  $ws.Range('E2').Font.Color = 8210719
  $ws.Range('F2').NumberFormat = '@'
  $ws.Range('F2').Interior.Color = 15921906
  $ws.Range('F2').HorizontalAlignment = -4108

  $ws.Range('B7').Value2 = '現在の番号 (番号1)'
  $ws.Range('B7').Font.Color = 8210719
  $ws.Range('C7').Font.Size = 22
  $ws.Range('C7').Font.Bold = $true
  $ws.Range('C7').NumberFormat = '@'
  $ws.Range('E7').Font.Size = 12
  $ws.Range('E7').Font.Bold = $true
  $ws.Range('B7:H7').Interior.Color = 16250871

  $hdr = @('#', '表A 項目', '表A 値', '表B 項目', '表B 値', '表C 項目', '表C 値')
  for ($i = 0; $i -lt $hdr.Count; $i++) {
    $c = $ws.Cells.Item(9, 2 + $i)
    $c.Value2 = $hdr[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 14994616
  }
  $ws.Range('B10:H19').NumberFormat = '@'
  $ws.Range('B10:B19').HorizontalAlignment = -4152
  $ws.Range('B9:H19').Borders.LineStyle = 1
  $ws.Range('B9:H19').Borders.Color = 13421772

  $ws.Range('J9').Value2 = '工程'
  $ws.Range('K9').Value2 = 'ミリ秒'
  $ws.Range('L9').Value2 = '割合'
  $ws.Range('J9:L9').Font.Bold = $true
  $ws.Range('J9:L9').Interior.Color = 14994616
  $ws.Range('K10:K26').NumberFormat = '#,##0.0'
  $ws.Range('L10:L26').NumberFormat = '0.0"%"'
  $ws.Range('J9:L26').Borders.LineStyle = 1
  $ws.Range('J9:L26').Borders.Color = 13421772
  $ws.Range('J21:K21').Font.Bold = $true

  $h2 = @('#', '時刻', '番号1', '候補数', '総時間ms', '選択後ms', '検知ms', '照合')
  for ($i = 0; $i -lt $h2.Count; $i++) {
    $c = $ws.Cells.Item(22, 2 + $i)
    $c.Value2 = $h2[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 14994616
  }
  $ws.Range('D23:D42').NumberFormat = '@'
  $ws.Range('F23:H42').NumberFormat = '#,##0.0'
  $ws.Range('B22:I42').Borders.LineStyle = 1
  $ws.Range('B22:I42').Borders.Color = 13421772

  $ws.Range('B43').Value2 = '候補一覧 (番号1 が複数件ヒットしたとき、選択ダイアログに出したもの)'
  $ws.Range('B43').Font.Color = 8210719
  $h3 = @('番号2', '行番号', '伝票番号', '日付', '数量', '状態', '品目コード', 'メーカー')
  for ($i = 0; $i -lt $h3.Count; $i++) {
    $c = $ws.Cells.Item(44, 2 + $i)
    $c.Value2 = $h3[$i]
    $c.Font.Bold = $true
    $c.Interior.Color = 14994616
  }
  $ws.Range('B45:I54').NumberFormat = '@'
  $ws.Range('B44:I54').Borders.LineStyle = 1
  $ws.Range('B44:I54').Borders.Color = 13421772

  $ws.Range('B57').Value2 = 'エラー'
  $ws.Range('B58').Value2 = '停止フラグ'
  $ws.Range('B59').Value2 = 'ログ出力先'
  $ws.Range('B60').Value2 = 'データフォルダ'
  $ws.Range('B61').Value2 = '自動選択 (計測用)'
  $ws.Range('B57:B61').Font.Color = 8210719
  $ws.Range('C57').Font.Color = 2701998
  $ws.Range('C58:C61').Font.Color = 8210719
  $ws.Range('B63').Value2 = 'メモ帳に 8 桁の番号を入力すると、3 本の CSV を読み直して結合し直し、この画面を更新します。'
  $ws.Range('B64').Value2 = '番号1 が複数件ヒットしたときは候補ダイアログが開きます。1 件選ぶと、その統合結果を表示します。'
  $ws.Range('B65').Value2 = '同じ番号を続けて読むときは、いったん欄を空にしてから読み取ってください。'
  $ws.Range('B63:B65').Font.Color = 8210719

  $ws.Range('N1').Value2 = '操作'
  $ws.Range('N1').Font.Bold = $true
  $anchor = $ws.Range('N2')
  $left = [double]$anchor.Left
  $y = [double]$anchor.Top
  $buttons = @(
    @{ c = '監視開始'; m = 'RDV2_StartMonitor' },
    @{ c = '監視停止'; m = 'RDV2_StopMonitor' },
    @{ c = '手動実行 (F2 の番号)'; m = 'RDV2_ManualRun' },
    @{ c = '画面クリア'; m = 'RDV2_Reset' }
  )
  foreach ($b in $buttons) {
    $btn = $ws.Buttons().Add($left, $y, 170.0, 24.0)
    $btn.Caption = $b.c
    $btn.OnAction = $b.m
    $y += 28.0
  }
  $ws.Range('A1').Select() | Out-Null
}

# --- the candidate dialog ---------------------------------------------------
# The eight identification columns are the documented set; README.md lists the
# same eight and src\v2\csharp\Rdv2Ui.cs shows them in the same order.
$pickCols = @(
  @{ h = '#';          w = 26 },
  @{ h = '番号2';      w = 62 },
  @{ h = '行番号';     w = 44 },
  @{ h = '伝票番号';   w = 76 },
  @{ h = '日付';       w = 62 },
  @{ h = '数量';       w = 42 },
  @{ h = '状態';       w = 46 },
  @{ h = '品目コード'; w = 74 },
  @{ h = 'メーカー';   w = 76 }
)

function Add-PickForm($wb) {
  $frm = $wb.VBProject.VBComponents.Add(3)      # vbext_ct_MSForm
  $frm.Name = 'frmRdv2Pick'
  # a parameterised COM property: PowerShell needs Item(..).Value, not (..)=
  $frm.Properties.Item('Caption').Value = '候補を選択'
  $frm.Properties.Item('Width').Value = 552
  $frm.Properties.Item('Height').Value = 348
  $d = $frm.Designer

  $hint = $d.Controls.Add('Forms.Label.1', 'lblHint', $true)
  $hint.Left = 10; $hint.Top = 8; $hint.Width = 520; $hint.Height = 16
  $hint.Caption = ''

  $x = 10.0
  for ($i = 0; $i -lt $pickCols.Count; $i++) {
    $lb = $d.Controls.Add('Forms.Label.1', ("lblH" + $i), $true)
    $lb.Left = $x + 2
    $lb.Top = 28
    $lb.Width = [double]$pickCols[$i].w
    $lb.Height = 14
    $lb.Caption = $pickCols[$i].h
    $x += [double]$pickCols[$i].w
  }

  $lst = $d.Controls.Add('Forms.ListBox.1', 'lstCand', $true)
  $lst.Left = 8; $lst.Top = 44; $lst.Width = 528; $lst.Height = 230
  $lst.ColumnCount = $pickCols.Count
  $lst.ColumnWidths = (($pickCols | ForEach-Object { $_.w }) -join ';')
  # fonts are set in UserForm_Initialize: a design-time MSForms control has no
  # Font object to set here

  $ok = $d.Controls.Add('Forms.CommandButton.1', 'cmdOk', $true)
  $ok.Left = 300; $ok.Top = 282; $ok.Width = 130; $ok.Height = 26
  $ok.Caption = 'この候補を表示'
  $ok.Default = $true

  $cn = $d.Controls.Add('Forms.CommandButton.1', 'cmdCancel', $true)
  $cn.Left = 438; $cn.Top = 282; $cn.Width = 98; $cn.Height = 26
  $cn.Caption = '選択しない'
  $cn.Cancel = $true

  $frm.CodeModule.AddFromString(@'
Option Explicit

' Filled by modRdv2App while the clock is still running -- building this list is
' the "選択肢作成" the measurement stops at. Show comes afterwards, and however
' long the operator then takes is not in any number.

Public Chosen As Long
Public AutoPick As Long

Private Sub UserForm_Initialize()
    Dim ctl As Control
    Chosen = -1
    AutoPick = -1
    lstCand.Font.Name = "Consolas"
    lstCand.Font.Size = 9
    For Each ctl In Me.Controls
        If Left$(ctl.Name, 4) = "lblH" Then
            ctl.Font.Size = 8
            ctl.Font.Bold = True
        End If
    Next ctl
End Sub

Public Sub Fill(ByVal key As String, ByVal n As Long)
    Dim arr() As String
    Dim i As Long, c As Long
    Me.Caption = "番号1 " & key & " の候補 " & n & " 件"
    lblHint.Caption = "番号1 " & key & " に " & n & " 件の候補があります。表示する 1 件を選んでください。"
    If n <= 0 Then Exit Sub
    ReDim arr(0 To n - 1, 0 To 8)
    For i = 0 To n - 1
        arr(i, 0) = CStr(i + 1)
        For c = 0 To 7
            arr(i, c + 1) = Rdv2CandCell(i, c)
        Next c
    Next i
    lstCand.List = arr
    lstCand.ListIndex = 0
End Sub

Private Sub cmdOk_Click()
    If lstCand.ListIndex >= 0 Then
        Chosen = lstCand.ListIndex
        Me.Hide
    End If
End Sub

Private Sub cmdCancel_Click()
    Chosen = -1
    Me.Hide
End Sub

Private Sub lstCand_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOk_Click
End Sub

' -autopick for an unattended measurement run. It takes the same path a click
' does, after the clock has already stopped.
Private Sub UserForm_Activate()
    If AutoPick >= 0 Then
        If AutoPick < lstCand.ListCount Then
            lstCand.ListIndex = AutoPick
        Else
            lstCand.ListIndex = 0
        End If
        cmdOk_Click
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Chosen = -1
        Cancel = True
        Me.Hide
    End If
End Sub
'@)
}

# --- Excel ------------------------------------------------------------------
# identity is settled before anything is done to it (excel_own.ps1):
# a reused or unidentifiable instance throws here and is never driven
$rdvOwn = New-OwnedExcel
$xl = $rdvOwn.App
$mine = @($rdvOwn.Pid)
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.AlertBeforeOverwriting = $false
Step ("started Excel (new pid: {0})" -f $(if ($mine.Count) { $mine -join ',' } else { 'reused' }))

$dist = Join-Path $Root 'dist'
if (-not (Test-Path -LiteralPath $dist)) { New-Item -ItemType Directory -Path $dist | Out-Null }

$builds = @(
  @{ cls = 'clsRdv2IdxHash'; name = 'ReaderDataViewer-VBA-Hash.xlsm'; method = 'VBA 自前ハッシュ' },
  @{ cls = 'clsRdv2IdxDict'; name = 'ReaderDataViewer-VBA-Dict.xlsm'; method = 'VBA 標準 Dictionary' }
)

try {
  foreach ($b in $builds) {
    $path = Join-Path $dist $b.name
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    Step ("{0}: Workbooks.Add" -f $b.cls)
    $wb = $xl.Workbooks.Add($xlWBATWorksheet)
    # save first: Application.Run resolves a macro through the workbook name,
    # and an unsaved book called "Sheet1" collides with its own sheet name
    $wb.SaveAs($path, $xlOpenXMLWorkbookMacroEnabled)

    Step '    add UserForm frmRdv2Pick'
    Add-PickForm $wb
    # adding a form opens the VBE window; leave it open and any compile problem
    # appears as a modal dialog with nobody to answer it
    try { $xl.VBE.MainWindow.Visible = $false } catch { }

    Step '    import modules'
    foreach ($m in $modules) {
      [void]$wb.VBProject.VBComponents.Import((Join-Path $stage "$m.bas"))
      Step "      imported $m"
    }
    [void]$wb.VBProject.VBComponents.Import((Join-Path $stage ($b.cls + '.cls')))
    Step ("      imported {0} as clsRdv2Idx" -f $b.cls)

    # AFTER the imports, never before
    Step '    AddFromGuid UIAutomationClient'
    [void]$wb.VBProject.References.AddFromGuid('{944DE083-8FB8-45CF-BCB7-C477ACB2F897}', 1, 0)

    Paint-View2 $wb.Worksheets.Item(1) $b.method
    $wb.Worksheets.Item(1).Range('C5').Value2 = $(if ($b.cls -eq 'clsRdv2IdxHash') { 'hash' } else { 'dict' })

    $tw = $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule
    if ($tw.CountOfLines -gt 0) { $tw.DeleteLines(1, $tw.CountOfLines) }
    $tw.AddFromString(@'
Option Explicit

' The builder never runs a macro. The screen is set up the first time the book
' is opened normally, which is the only way anyone ever sees it.
Private Sub Workbook_Open()
    On Error Resume Next
    modRdv2App.RDV2_Reset
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    modRdv2App.RDV2_StopMonitor
    Me.Saved = True
End Sub
'@)

    Step '    final Save'
    $wb.Save()
    $wb.Close($false)
    Step ("    built {0}" -f $path)
  }
}
catch {
  Step ("FAILED: " + $_.Exception.Message)
  Step ("        at " + $_.InvocationInfo.PositionMessage.Trim())
  throw
}
finally {
  Step 'quitting Excel'
  try { $xl.Quit() } catch { }
  try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Milliseconds 500
  foreach ($p in $mine) {
    $q = Get-Process -Id $p -ErrorAction SilentlyContinue
    if ($q) { Step "closing my excel $p"; $q.Kill() }
  }
  Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue
}
Step 'done'
