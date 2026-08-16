# ============================================================================
# ui_grid_app.ps1 -- the VBA build's main screen, painted from the SAME measured
# rectangle table the C# build lays out from.
#
# WHY A GRID
# The C# screen owns its pixels: Rdv3Geom.cs holds the rectangles read out of
# "Reader Data Viewer_ver3.html" and Layout1() places every element from that
# table. A worksheet cannot place anything at a pixel -- it can only place
# things in CELLS. So the sheet is given cells small enough to be used AS
# pixels: one calibrated column width and one row height, repeated, and every
# element of the reference is a MERGE of the block its rectangle covers.
#
# The unit is not 1 px. A 1240 x 689 card at 1 px per cell is 1240 columns and
# 689 rows of merges, and Excel spends real time on that for a fidelity nobody
# can see. GRID_PX below is the unit; at 4 px the worst rounding error is 2 px,
# which is the same tolerance the C# acceptance check allows.
#
# WHAT IS DIFFERENT FROM THE C# SCREEN, AND WHY
#   - the error row is ALWAYS present, blank when there is no error. The
#     reference inserts it and lets the card grow; a sheet cannot reflow, and
#     repainting every row height on an error would be a full-sheet reformat on
#     a hot path. So the band is reserved once and filled when it is used, and
#     the card is that much taller than 689.
#   - the candidate table shows the rows it has. The reference scrolls a 152 px
#     viewport; a sheet scrolls itself.
#
# Everything the code writes at run time is a DEFINED NAME (rdv*). The builder
# creates them from the table below and modRdv3Ui writes through them, so a
# moved rectangle never means hunting for an address in VBA.
#
# No Shape, no UserForm, no ActiveX, no Forms control. Buttons are hyperlink
# cells, dispatched through Worksheet_FollowHyperlink, exactly as before.
# ============================================================================

$GRID_PX = 4.0                     # one cell, in reference px
$CARD_W = 1240.0
$ERR_BAND = 38.0                   # the reserved error row (see the note above)
$CARD_H = 689.0 + $ERR_BAND

# ---- the design tokens, as Excel BGR ---------------------------------------
function RGBx([int]$r, [int]$g, [int]$b) { return $r + $g * 256 + $b * 65536 }
$C_BG        = RGBx 242 242 243    # --color-bg        #f2f2f3
$C_SURFACE   = RGBx 233 233 234    # --color-surface   #e9e9ea
$C_INK       = RGBx 29 31 32       # --color-text      #1d1f20
$C_DIVIDER   = RGBx 205 205 206    # text at 16% on bg
$C_ACCENT    = RGBx 89 128 166     # --color-accent    #5980a6
$C_ACC100    = RGBx 238 246 255    # accent-100        #eef6ff
$C_ACC700    = RGBx 65 97 128      # accent-700        #416180
$C_ACC800    = RGBx 44 69 93       # accent-800        #2c455d
$C_STATUSBAR = RGBx 29 45 61       # accent-900        #1d2d3d
$C_N100      = RGBx 245 245 248
$C_N300      = RGBx 212 212 215
$C_N700      = RGBx 93 93 96
$C_WHITE     = RGBx 255 255 255
$C_ERRINK    = RGBx 176 74 62      # the input-error brick, ui-spec 11

# ui-spec 6.3: only fonts Windows already has, never shipped with the product
$F_BODY = 'Meiryo UI'
$F_HEAD = 'Bahnschrift SemiBold Condensed'
$F_NUM  = 'Bahnschrift SemiBold'

$xlCenter = -4108; $xlLeft = -4131; $xlRight = -4152; $xlTop = -4160; $xlBottom = -4107
$xlContinuous = 1; $xlNone = -4142

# ---------------------------------------------------------------------------
# One calibrated column width. Excel stores column width in "characters" of the
# Normal font, and the pixel it lands on is quantised, so the width is not
# computed -- it is MEASURED, once, by asking Excel what it actually produced.
# ---------------------------------------------------------------------------
function Set-GridWidth($ws, [int]$firstCol, [int]$lastCol, [double]$targetPx) {
  $targetPt = $targetPx * 0.75
  $best = 0.1; $bestErr = [double]::MaxValue
  $probe = $ws.Columns.Item($firstCol)
  for ($w = 0.05; $w -le 2.0; $w += 0.01) {
    $probe.ColumnWidth = $w
    $err = [Math]::Abs($probe.Width - $targetPt)
    if ($err -lt $bestErr) { $bestErr = $err; $best = $w }
    if ($probe.Width -gt $targetPt + 3) { break }
  }
  $probe.ColumnWidth = $best
  $got = $probe.Width
  $ws.Range((ColLetter $firstCol) + ':' + (ColLetter $lastCol)).ColumnWidth = $best
  return [pscustomobject]@{ Width = $best; Px = $got / 0.75 }
}

# ---------------------------------------------------------------------------
# a reference rectangle -> the cell block that covers it
# ---------------------------------------------------------------------------
$script:ColBase = 2                # one spare column/row outside the card
$script:RowBase = 2

function ColLetter([int]$n) {
  $s = ''
  while ($n -gt 0) {
    $m = ($n - 1) % 26
    $s = [string][char](65 + $m) + $s
    $n = [int](($n - $m - 1) / 26)
  }
  return $s
}

# The range is asked for by A1 address, not as Range(cellA, cellB). Measured:
# the two-Range form comes back from PowerShell as an Object[] rather than a
# Range -- .Interior then has no .Color and the build dies with a property that
# "cannot be found" a long way from the cause. An address string binds to exactly
# one overload and cannot do that.
function Rect($ws, [double]$x, [double]$y, [double]$w, [double]$h) {
  $c0 = [int][Math]::Round($x / $GRID_PX) + $script:ColBase
  $r0 = [int][Math]::Round($y / $GRID_PX) + $script:RowBase
  $c1 = [int][Math]::Round(($x + $w) / $GRID_PX) + $script:ColBase - 1
  $r1 = [int][Math]::Round(($y + $h) / $GRID_PX) + $script:RowBase - 1
  if ($c1 -lt $c0) { $c1 = $c0 }
  if ($r1 -lt $r0) { $r1 = $r0 }
  $a = (ColLetter $c0) + [string]$r0 + ':' + (ColLetter $c1) + [string]$r1
    # the comma matters: PowerShell UNROLLS an enumerable on the way out, and
  # a COM Range enumerates into its cells -- returning one plainly handed
  # the caller 3,410 single-cell objects instead of the block. ,$x makes it
  # one element, and the single unroll on output leaves the Range itself.
  return ,$ws.Range($a)
}

# Style a rectangle, and merge it ONLY when it has to behave as one cell.
#
# MERGING A BACKGROUND SWALLOWS EVERYTHING DRAWN ON IT. The panels and the title
# bar are rectangles the width of the card; merging them made every label and
# control inside them part of a bigger merged cell, and Excel DISCARDS a write to
# a cell that belongs to a merge whose anchor is elsewhere -- silently, with no
# error. The screen kept its merges, its names, its borders and its geometry, and
# lost every caption and every value: a blank sheet that passes a geometry check.
#
# So only a box that carries text or a name is merged -- those are the labels,
# the readouts and the buttons, and they have to be one cell to centre their text
# and to answer a click. A fill or a rule is painted across its cells as it is.
function Box($ws, [double]$x, [double]$y, [double]$w, [double]$h, [hashtable]$o) {
  $r = Rect $ws $x $y $w $h
  $isPart = ($o.ContainsKey('text') -or $o.ContainsKey('name'))
  if ($isPart -and $r.Cells.Count -gt 1) { $r.Merge() | Out-Null }
  if ($o.ContainsKey('fill')) { $r.Interior.Color = [int]$o['fill'] }
  if ($o.ContainsKey('text')) { $r.Value2 = [string]$o['text'] }
  # each value settled into a typed local first: an inline if that yields $null
  # reaches Excel as an invalid cast, a long way from the element that caused it
  $fName = [string]$F_BODY
  if ($o.ContainsKey('font')) { $fName = [string]$o['font'] }
  $fSize = [double]9
  if ($o.ContainsKey('size')) { $fSize = [double]$o['size'] }
  $fInk = [int]$C_INK
  if ($o.ContainsKey('ink')) { $fInk = [int]$o['ink'] }
  $hA = [int]$xlLeft
  if ($o.ContainsKey('halign')) { $hA = [int]$o['halign'] }
  $vA = [int]$xlCenter
  if ($o.ContainsKey('valign')) { $vA = [int]$o['valign'] }
  $r.Font.Name = $fName
  $r.Font.Size = $fSize
  $r.Font.Bold = [bool]$o['bold']
  $r.Font.Color = $fInk
  $r.HorizontalAlignment = $hA
  $r.VerticalAlignment = $vA
  if ($o.ContainsKey('wrap')) { $r.WrapText = [bool]$o['wrap'] }
  if ($o.ContainsKey('fmt')) { $r.NumberFormat = [string]$o['fmt'] }
  if ($o.ContainsKey('border')) {
    # around the block, not around every cell in it: an unmerged rectangle on a
    # 4 px grid would otherwise be drawn as a mesh of hundreds of little boxes
    $r.BorderAround($xlContinuous, 2, 0, [int]$o['border']) | Out-Null
  }
  if ($o.ContainsKey('name')) {
    $ws.Parent.Names.Add([string]$o['name'], $r) | Out-Null
  }
  return ,$r
}

# a 1px rule, drawn as a thin filled band
function Rule($ws, [double]$x, [double]$y, [double]$w, [int]$color) {
  # a filled band, never merged: a rule runs the width of the card and merging it
  # would swallow whatever is drawn across it (see Box)
  $r = Rect $ws $x $y $w $GRID_PX
  $r.Interior.Color = $color
  return ,$r
}

# a hyperlink cell that behaves as a button. Same mechanism the build has always
# used: the click arrives in Worksheet_FollowHyperlink and modRdv3App dispatches.
function Button($ws, [double]$x, [double]$y, [double]$w, [double]$h,
                [string]$caption, [string]$tip, [string]$name, [bool]$primary) {
  $r = Rect $ws $x $y $w $h
  if ($r.Cells.Count -gt 1) { $r.Merge() | Out-Null }
  $addr = $r.Cells.Item(1).Address(0, 0)
  $r.Value2 = $caption
  [void]$ws.Hyperlinks.Add($r.Cells.Item(1), "", ($ws.Name + "!" + $addr), $tip, $caption)
  $r.Font.Name = $F_HEAD
  $r.Font.Size = 10.5
  $r.Font.Bold = $true
  $r.Font.Underline = $false
  $r.HorizontalAlignment = $xlCenter
  $r.VerticalAlignment = $xlCenter
  $r.Borders.LineStyle = $xlContinuous
  $r.Borders.Weight = 2
  if ($primary) {
    $r.Interior.Color = $C_ACCENT
    $r.Font.Color = $C_BG
    $r.Borders.Color = $C_ACCENT
  } else {
    $r.Interior.Color = $C_BG
    $r.Font.Color = $C_INK
    $r.Borders.Color = $C_DIVIDER
  }
  $ws.Parent.Names.Add($name, $r) | Out-Null
  return ,$r
}

# ---------------------------------------------------------------------------
# THE SCREEN. Coordinates are the reference's own (Rdv3Geom.cs / ui-ref-geom.json).
#
# One deliberate departure: the reference scrolls the candidate table inside a
# 152 px viewport. A worksheet has no sub-viewport -- it scrolls itself -- so the
# list panel is given the height its rows actually need and everything below it
# moves down by the same amount. Sizes and horizontal positions are unchanged;
# this is the reference's own rule for extra height (ui-spec 6.1), applied in the
# one direction a sheet can apply it.
# ---------------------------------------------------------------------------
function Paint-Ui($ws) {
  $ws.Name = 'UI'
  $wb = $ws.Parent

  $cols = [int][Math]::Ceiling($CARD_W / $GRID_PX) + $script:ColBase + 1
  $rowsN = [int][Math]::Ceiling(($CARD_H + 320) / $GRID_PX) + $script:RowBase + 1

  $ws.Cells.Font.Name = $F_BODY
  $ws.Cells.Font.Size = 9
  $ws.Cells.Interior.Color = $C_BG
  $cal = Set-GridWidth $ws $script:ColBase $cols $GRID_PX
  $ws.Columns.Item(1).ColumnWidth = 0.6
  $ws.Rows("1:" + [string]$rowsN).RowHeight = $GRID_PX * 0.75

  # how far everything under the candidate list moves down (see the header)
  $rowsShown = 10
  $listNeed = 31.0 + $rowsShown * 38.6
  $grow = $listNeed - 152.0

  # ---- 1. title bar ------------------------------------------------------
  Box $ws 0.7 0.7 1238.7 43.7 @{ fill = $C_BG } | Out-Null
  Box $ws 14.3 7.5 158.0 29.4 @{ text = 'Reader Data Viewer'; font = $F_HEAD; size = 14.5; bold = $true } | Out-Null
  Box $ws 178.0 10.7 126.0 23.0 @{ text = 'VBA 標準 Dictionary'; size = 8.5; fill = $C_ACC100; ink = $C_ACC800; halign = $xlCenter } | Out-Null
  Box $ws 312.0 10.0 72.0 24.4 @{ text = '統合台帳'; size = 8.5; halign = $xlCenter; border = $C_ACCENT; ink = $C_ACC800 } | Out-Null
  Button $ws 952.0 8.7 126.7 26.9 '設定' '監視対象や桁数、ファイルの場所を変更します' 'rdvBtnSettings' $false | Out-Null
  Button $ws 1099.0 8.7 126.7 26.9 '監視対象を再検出' '監視対象の欄を探し直します' 'rdvBtnRebind' $false | Out-Null
  Rule $ws 0.7 44.4 1238.7 $C_DIVIDER | Out-Null

  # ---- 2. summary --------------------------------------------------------
  Box $ws 14.3 58.0 1211.5 97.2 @{ fill = $C_WHITE; border = $C_DIVIDER } | Out-Null
  Box $ws 28.5 68.8 145.5 15.5 @{ text = '現在の番号 (番号1)'; size = 7.5; ink = $C_ACCENT } | Out-Null
  Box $ws 28.5 86.3 145.5 37.4 @{ name = 'rdvKeyValue'; font = $F_NUM; size = 22; bold = $true; fmt = '@' } | Out-Null
  Box $ws 28.5 125.7 145.5 18.6 @{ name = 'rdvKeySub'; size = 8.5; ink = $C_N700 } | Out-Null

  Rule $ws 187.6 64.8 1.2 $C_DIVIDER | Out-Null
  Box $ws 201.9 68.8 119.6 15.5 @{ text = '代表ステータス'; size = 7.5; ink = $C_ACCENT } | Out-Null
  Box $ws 201.9 86.3 119.6 37.4 @{ name = 'rdvStatusValue'; font = $F_NUM; size = 22; bold = $true } | Out-Null
  Box $ws 201.9 125.7 119.6 18.6 @{ name = 'rdvStatusSub'; size = 8.5; ink = $C_N700 } | Out-Null

  Rule $ws 335.0 64.8 1.2 $C_DIVIDER | Out-Null
  Box $ws 349.3 68.8 123.5 15.5 @{ text = '台帳総件数'; size = 7.5; ink = $C_ACCENT } | Out-Null
  Box $ws 349.3 86.3 170.0 37.4 @{ name = 'rdvRowsValue'; font = $F_NUM; size = 22; bold = $true } | Out-Null
  Box $ws 349.3 125.7 170.0 18.6 @{ name = 'rdvRowsSub'; size = 8.5; ink = $C_N700 } | Out-Null

  Box $ws 633.0 68.8 110.0 18.6 @{ text = '番号1 を入力'; size = 8.5; ink = $C_N700 } | Out-Null
  Box $ws 633.0 92.4 110.0 36.0 @{ name = 'rdvInput'; fill = $C_SURFACE; border = $C_DIVIDER; fmt = '@'; font = $F_NUM; size = 12; halign = $xlCenter } | Out-Null
  Box $ws 633.0 130.0 220.0 14.0 @{ name = 'rdvInputErr'; size = 7.5; ink = $C_ERRINK } | Out-Null
  Button $ws 749.8 96.7 73.8 31.7 '検索' '入力した番号1 で統合台帳を検索します' 'rdvBtnSearch' $true | Out-Null
  Button $ws 830.4 96.7 85.8 31.7 '内容クリア' '入力と結果表示を消します (台帳と処理済みは消えません)' 'rdvBtnClear' $false | Out-Null
  Button $ws 923.0 96.7 100.5 31.7 '処理済み' '表示中の統合レコードを処理済みにします (確認あり)' 'rdvBtnProcessed' $false | Out-Null

  Rule $ws 1037.1 64.8 1.2 $C_DIVIDER | Out-Null
  Box $ws 1051.4 68.8 160.1 20.0 @{ name = 'rdvIdentName'; size = 9.5; bold = $true } | Out-Null
  Box $ws 1051.4 90.0 160.1 17.0 @{ name = 'rdvIdentSub'; size = 8; ink = $C_N700 } | Out-Null
  Box $ws 1051.4 110.0 60.0 23.0 @{ name = 'rdvIdentTag'; size = 8; halign = $xlCenter; border = $C_ACCENT; ink = $C_ACC800 } | Out-Null

  # ---- 3. candidate list -------------------------------------------------
  Box $ws 14.3 172.2 1211.5 (195.7 + $grow) @{ fill = $C_WHITE; border = $C_DIVIDER } | Out-Null
  Box $ws 28.5 183.0 92.0 20.2 @{ text = '候補一覧'; font = $F_HEAD; size = 13; bold = $true } | Out-Null
  Box $ws 128.0 187.0 210.0 18.6 @{ text = '行の # をクリックすると表示します'; size = 8; ink = $C_N700 } | Out-Null
  Box $ws 344.0 186.0 690.0 20.0 @{ name = 'rdvVerdict'; size = 9.5; bold = $true; ink = $C_ACC700 } | Out-Null
  Box $ws 1151.8 184.7 59.7 23.0 @{ name = 'rdvListTag'; size = 8; halign = $xlCenter; border = $C_ACCENT; ink = $C_ACC800 } | Out-Null
  Rule $ws 14.9 214.6 1210.1 $C_DIVIDER | Out-Null

  $head = @('#', '番号2', '行番号', '伝票番号', '日付', '数量', '状態', '品目コード', 'メーカー', '処理済み')
  $colw = @(44.0, 94.7, 69.6, 103.8, 88.8, 57.8, 83.3, 85.9, 107.4, 80.5)
  $x = 28.5
  for ($i = 0; $i -lt $head.Count; $i++) {
    $ha = $xlLeft
    if ($i -eq 0 -or $i -eq 2 -or $i -eq 5) { $ha = $xlRight }
    Box $ws $x 215.2 $colw[$i] 31.0 @{ text = $head[$i]; size = 7.5; bold = $true; ink = $C_N700; fill = $C_BG; halign = $ha } | Out-Null
    $x += $colw[$i]
  }
  $rowTop = 246.2
  for ($r = 0; $r -lt $rowsShown; $r++) {
    $y = $rowTop + $r * 38.6
    $x = 28.5
    for ($i = 0; $i -lt $head.Count; $i++) {
      $ha = $xlLeft
      if ($i -eq 0 -or $i -eq 2 -or $i -eq 5) { $ha = $xlRight }
      $o = @{ name = ('rdvCand_' + $r + '_' + $i); size = 9; fmt = '@'; halign = $ha; fill = $C_WHITE }
      if ($i -eq 1) { $o.bold = $true }
      # 38.6 less the rule band: the cells are merged, and a merged block
      # cannot be painted over, so a separator drawn INSIDE them shows only
      # where there are no cells -- past the last column
      $cell = Box $ws $x $y $colw[$i] (38.6 - $GRID_PX) $o
      if ($i -eq 0) {
        # the # cell is the row's affordance, the way the reference selects a row
        $addr = $cell.Cells.Item(1, 1).Address(0, 0)
        # TextToDisplay is a SPACE, not '': given an empty string Excel shows the
        # link's own target, so the empty rows read "UI!I64" before the macros
        # run and the operator's first sight of the product is its internals
        [void]$ws.Hyperlinks.Add($cell.Cells.Item(1, 1), "", ($ws.Name + "!" + $addr), 'この候補の統合レコードを表示します', ' ')
        $cell.Font.Name = $F_BODY
        $cell.Font.Size = 9
        $cell.Font.Underline = $false
        $cell.Font.Color = $C_ACCENT
        $cell.Font.Bold = $true
        $cell.HorizontalAlignment = $xlRight
      }
      $x += $colw[$i]
    }
    Rule $ws 28.5 ($y + 38.6 - $GRID_PX) 1183.0 $C_N300 | Out-Null
  }

  # ---- 4. merged record ---------------------------------------------------
  $ry = 384.9 + $grow
  Box $ws 14.3 $ry 1211.5 266.8 @{ fill = $C_WHITE; border = $C_DIVIDER } | Out-Null
  Box $ws 28.5 ($ry + 10.8) 106.0 20.2 @{ text = '統合レコード'; font = $F_HEAD; size = 13; bold = $true } | Out-Null
  Box $ws 123.0 ($ry + 14.1) 430.0 18.6 @{ text = '候補一覧から行を選ぶと、その 1 件をここに出します'; size = 8; ink = $C_N700 } | Out-Null
  Box $ws 1071.2 ($ry + 10.8) 27.8 23.0 @{ name = 'rdvRecTagStatus'; size = 8; halign = $xlCenter; border = $C_ACCENT; ink = $C_ACC800 } | Out-Null
  Box $ws 1109.2 ($ry + 10.8) 27.8 23.0 @{ name = 'rdvRecTagProc'; size = 8; halign = $xlCenter; border = $C_ACCENT; ink = $C_ACC800 } | Out-Null
  Box $ws 1147.1 ($ry + 10.8) 64.4 23.0 @{ name = 'rdvRecTagKey2'; size = 8; halign = $xlCenter; border = $C_ACCENT; ink = $C_ACC800 } | Out-Null
  Rule $ws 14.9 ($ry + 40.7) 1210.1 $C_DIVIDER | Out-Null

  # ui-spec 4: a_name / a_code / a_grade / a_dept / a_date / a_amount / a_rate . a_flag
  $kv = @('取引先名', '取引先コード', 'グレード', '部門', '登録日', '金額', 'レート ・ フラグ')
  for ($i = 0; $i -lt $kv.Count; $i++) {
    $y = $ry + 51.5 + $i * 28.8
    Box $ws 28.5 $y 200.0 28.8 @{ text = $kv[$i]; size = 9; ink = $C_N700 } | Out-Null
    Box $ws 232.0 $y 310.2 28.8 @{ name = ('rdvKv' + $i); size = 9; bold = $true; halign = $xlRight; fmt = '@' } | Out-Null
    if ($i -lt $kv.Count - 1) { Rule $ws 28.5 ($y + 28.8 - $GRID_PX) 513.7 $C_N300 | Out-Null }
  }
  Box $ws 569.4 ($ry + 51.5) 642.1 17.0 @{ text = '摘要'; size = 7.5; ink = $C_ACCENT } | Out-Null
  Box $ws 569.4 ($ry + 72.6) 642.1 74.4 @{ name = 'rdvMemo'; size = 9; fill = $C_N100; border = $C_DIVIDER; wrap = $true; valign = $xlTop; fmt = '@' } | Out-Null
  Box $ws 569.4 ($ry + 157.1) 642.1 17.0 @{ text = '備考'; size = 7.5; ink = $C_ACCENT } | Out-Null
  Box $ws 569.4 ($ry + 178.2) 642.1 74.4 @{ name = 'rdvRemark'; size = 9; fill = $C_N100; border = $C_DIVIDER; wrap = $true; valign = $xlTop; fmt = '@' } | Out-Null

  # ---- 5. error row (reserved; blank until it is used) --------------------
  $ey = $ry + 266.8 + 8.0
  Box $ws 0.7 $ey 1238.7 38.0 @{ fill = $C_BG } | Out-Null
  Box $ws 14.3 ($ey + 7.0) 60.0 23.0 @{ name = 'rdvErrTag'; size = 8; halign = $xlCenter } | Out-Null
  Box $ws 84.0 ($ey + 7.0) 1140.0 23.0 @{ name = 'rdvErrText'; size = 9.5; bold = $true; ink = $C_ACC800 } | Out-Null

  # ---- 6. status bar ------------------------------------------------------
  $sy = $ey + 46.0
  Box $ws 0.7 $sy 1238.7 28.6 @{ fill = $C_STATUSBAR } | Out-Null
  Box $ws 30.0 ($sy + 5.0) 150.0 18.6 @{ name = 'rdvStState'; size = 8.5; bold = $true; fill = $C_STATUSBAR; ink = $C_BG } | Out-Null
  Box $ws 190.0 ($sy + 5.0) 330.0 18.6 @{ name = 'rdvStWatch'; size = 8.5; fill = $C_STATUSBAR; ink = $C_BG } | Out-Null
  Box $ws 528.0 ($sy + 5.0) 300.0 18.6 @{ name = 'rdvStLedger'; size = 8.5; fill = $C_STATUSBAR; ink = $C_BG } | Out-Null
  Box $ws 836.0 ($sy + 5.0) 126.0 18.6 @{ name = 'rdvStMerge'; size = 8.5; bold = $true; fill = $C_STATUSBAR; ink = $C_BG } | Out-Null
  Box $ws 966.0 ($sy + 5.0) 118.0 18.6 @{ name = 'rdvStSearch'; size = 8.5; bold = $true; fill = $C_STATUSBAR; ink = $C_BG } | Out-Null
  Box $ws 1088.0 ($sy + 5.0) 138.0 18.6 @{ name = 'rdvStId'; size = 8; fill = $C_STATUSBAR; ink = $C_BG; halign = $xlRight } | Out-Null

  # The grid is the layout, so Excel's own gridlines and headings would draw
  # straight through the card. Both are window settings and both are saved with
  # the workbook.
  $ws.Activate()
  $win = $ws.Parent.Windows.Item(1)
  $win.DisplayGridlines = $false
  $win.DisplayHeadings = $false
  $win.Zoom = 100

  # THE WHOLE CARD, as one name. The app fits the window to this at boot rather
  # than to a number written down twice (see modRdv3Ui.Rdv3UiFitWindow).
  $card = Rect $ws 0.0 0.0 $CARD_W ($sy + 28.6)
  $ws.Parent.Names.Add('rdvCard', $card) | Out-Null

  $ws.Range('A1').Select() | Out-Null
  return [pscustomobject]@{
    ColWidth = $cal.Width
    ColPx = [Math]::Round($cal.Px, 2)
    Cols = $cols
    Rows = $rowsN
    Grow = $grow
    CardH = ($sy + 28.6)
  }
}

# ---------------------------------------------------------------------------
# THE SETTINGS SCREEN, on the same pseudo-pixel grid.
#
# The C# build opens a modal window (ui-spec 12). A workbook has no second window
# worth having, so this is a sheet -- hidden until 設定 is pressed -- carrying the
# same three groups the C# dialog has: 監視対象 / 動作 / ファイル. It is laid out
# as three sections down the sheet rather than three tabs, because a sheet
# scrolls and tabs would need a control the build does not allow itself.
#
# Every editable cell is a defined name (set*) and modRdv3Set reads and writes
# through those names only.
# ---------------------------------------------------------------------------
function Paint-Settings($ws) {
  $ws.Name = 'SETTINGS'
  $wb = $ws.Parent

  $cols = [int][Math]::Ceiling($CARD_W / $GRID_PX) + $script:ColBase + 1
  $rowsN = [int][Math]::Ceiling(950.0 / $GRID_PX) + $script:RowBase + 1

  $ws.Cells.Font.Name = $F_BODY
  $ws.Cells.Font.Size = 9
  $ws.Cells.Interior.Color = $C_BG
  $cal = Set-GridWidth $ws $script:ColBase $cols $GRID_PX
  $ws.Columns.Item(1).ColumnWidth = 0.6
  $ws.Rows("1:" + [string]$rowsN).RowHeight = $GRID_PX * 0.75

  # ---- header -------------------------------------------------------------
  Box $ws 0.7 0.7 1238.7 60.0 @{ fill = $C_BG } | Out-Null
  Box $ws 14.3 12.0 300.0 28.0 @{ text = '設定'; font = $F_HEAD; size = 16; bold = $true } | Out-Null
  Box $ws 14.3 38.0 700.0 16.0 @{ text = '監視対象・番号の決まり・ファイルの場所。保存すると ReaderDataViewer.json に書き戻します'; size = 8; ink = $C_N700 } | Out-Null
  Button $ws 1060.0 14.0 166.0 30.0 '主画面へ戻る' '設定を閉じて主画面に戻ります (保存はしません)' 'setBtnBack' $false | Out-Null
  Rule $ws 0.7 60.7 1238.7 $C_DIVIDER | Out-Null

  # ---- 1. 監視対象 ---------------------------------------------------------
  Box $ws 14.3 74.0 300.0 20.0 @{ text = '監視対象 ― どの画面のどの欄を読むか'; size = 8; ink = $C_ACCENT; bold = $true } | Out-Null
  $head = @('有効', '名前', 'ウィンドウ:クラス名', 'ウィンドウ:プロセス', 'ウィンドウ:名前(部分)',
            '欄:AutomationId', '欄:種類', '取り方', '状態')
  $colw = @(56.0, 130.0, 150.0, 130.0, 150.0, 150.0, 110.0, 74.0, 261.0)
  $x = 14.3
  for ($i = 0; $i -lt $head.Count; $i++) {
    Box $ws $x 100.0 $colw[$i] 24.0 @{ text = $head[$i]; size = 7.5; bold = $true; ink = $C_N700; fill = $C_BG } | Out-Null
    $x += $colw[$i]
  }
  Rule $ws 14.3 124.0 1211.5 $C_DIVIDER | Out-Null
  $keys = @('on', 'name', 'class', 'proc', 'like', 'fid', 'ftype', 'read', 'why')
  for ($r = 0; $r -lt 6; $r++) {
    $y = 128.0 + $r * 30.0
    $x = 14.3
    for ($i = 0; $i -lt $head.Count; $i++) {
      $o = @{ name = ('setT' + $r + '_' + $keys[$i]); size = 9; fmt = '@' }
      if ($i -eq 8) {
        # the reason a target cannot be watched: written by the code, not typed
        $o.fill = $C_BG
        $o.ink = $C_ERRINK
        $o.size = 8
      } else {
        $o.fill = $C_WHITE
        $o.border = $C_DIVIDER
      }
      Box $ws $x $y $colw[$i] 30.0 $o | Out-Null
      $x += $colw[$i]
    }
  }
  Box $ws 14.3 312.0 1211.5 18.0 @{ text = '「有効」は はい / いいえ / 削除。この画面に列が無い項目 (中間パス・欄のクラス名など) は保存しても消えません。「欄:種類」は Edit や Document (複数はカンマ区切り)、「取り方」は value / text / name。ウィンドウか欄のどちらも空の行は、どの画面にも一致するので監視しません'; size = 7.5; ink = $C_N700 } | Out-Null
  # the element picker: the operator focuses the field and the BE reports it
  Button $ws 14.3 336.0 150.0 26.0 '画面から選ぶ' '対象アプリの欄をクリックするとプレビューします。反映は「この要素を使う」だけです' 'setBtnPick' $false | Out-Null
  Box $ws 174.0 336.0 120.0 26.0 @{ text = '取込み先の行'; size = 8; ink = $C_N700 } | Out-Null
  Box $ws 296.0 336.0 60.0 26.0 @{ name = 'setPickRow'; size = 9; bold = $true; fmt = '@'; fill = $C_WHITE; border = $C_DIVIDER; halign = $xlCenter } | Out-Null
  Box $ws 366.0 336.0 860.0 26.0 @{ text = '対象アプリの欄をクリックするとプレビューが追随します。クリックしても設定は変わりません。反映は「この要素を使う」だけです'; size = 8; ink = $C_N700 } | Out-Null

  # THE PREVIEW. What was picked, beside what the chosen row holds now. Nothing
  # here is a change: 採用 is what changes the row, and 保存 is what changes the
  # file. A click in the other application must never be able to rewrite a
  # target by itself.
  Box $ws 14.3 370.0 1211.5 96.0 @{ fill = $C_N100; border = $C_DIVIDER } | Out-Null
  Box $ws 28.5 376.0 300.0 18.0 @{ text = 'プレビュー ― 見ているだけで、まだ何も変えていません'; size = 8; ink = $C_ACCENT; bold = $true } | Out-Null
  $pvHead = @('', 'ウィンドウ:クラス名', 'ウィンドウ:プロセス', '欄:AutomationId', '欄:種類', '値の取り方', 'ウィンドウ名')
  $pvW = @(96.0, 180.0, 150.0, 180.0, 130.0, 100.0, 320.0)
  $x = 28.5
  for ($i = 0; $i -lt $pvHead.Count; $i++) {
    Box $ws $x 396.0 $pvW[$i] 18.0 @{ text = $pvHead[$i]; size = 7.5; bold = $true; ink = $C_N700; fill = $C_N100 } | Out-Null
    $x += $pvW[$i]
  }
  $pvNames = @('setPvClass', 'setPvProc', 'setPvFid', 'setPvFType', 'setPvRead', 'setPvWName')
  $x = 28.5
  Box $ws $x 416.0 $pvW[0] 22.0 @{ text = '取得'; size = 8; bold = $true; ink = $C_ACC800; fill = $C_N100 } | Out-Null
  $x += $pvW[0]
  for ($i = 0; $i -lt $pvNames.Count; $i++) {
    Box $ws $x 416.0 $pvW[$i + 1] 22.0 @{ name = $pvNames[$i]; size = 9; bold = $true; fmt = '@'; fill = $C_WHITE; border = $C_DIVIDER } | Out-Null
    $x += $pvW[$i + 1]
  }
  $nowNames = @('setPvNowClass', 'setPvNowProc', 'setPvNowFid', 'setPvNowFType')
  $x = 28.5
  Box $ws $x 440.0 $pvW[0] 22.0 @{ text = '現在'; size = 8; ink = $C_N700; fill = $C_N100 } | Out-Null
  $x += $pvW[0]
  for ($i = 0; $i -lt $nowNames.Count; $i++) {
    Box $ws $x 440.0 $pvW[$i + 1] 22.0 @{ name = $nowNames[$i]; size = 9; ink = $C_N700; fmt = '@'; fill = $C_N100 } | Out-Null
    $x += $pvW[$i + 1]
  }
  Button $ws 860.0 438.0 170.0 26.0 'この要素を使う' 'プレビュー中の要素を上の行へ反映します (設定ファイルはまだ変わりません)' 'setBtnAdopt' $true | Out-Null
  Button $ws 1040.0 438.0 130.0 26.0 '閉じる' 'ピッカーを閉じます。何も反映しません' 'setBtnPickCancel' $false | Out-Null
  Box $ws 28.5 466.0 1180.0 18.0 @{ name = 'setPvNote'; size = 8; ink = $C_ACC800 } | Out-Null

  # ---- 2. 動作 -------------------------------------------------------------
  Rule $ws 14.3 500.0 1211.5 $C_DIVIDER | Out-Null
  Box $ws 14.3 510.0 400.0 20.0 @{ text = '動作 ― 番号の決まりと読み取りのタイミング'; size = 8; ink = $C_ACCENT; bold = $true } | Out-Null
  $fields = @(
    @('桁数 (次回起動から)', 'setKeyLen', 0),
    @('数字のみ (はい/いいえ)', 'setDigits', 1),
    @('監視間隔 ms', 'setPoll', 0),
    @('確定待ち ms', 'setStable', 1),
    @('再接続間隔 ms', 'setRebind', 0),
    @('候補の表示行数', 'setCand', 1),
    @('前面のウィンドウを優先 (はい/いいえ)', 'setFocus', 0)
  )
  for ($i = 0; $i -lt $fields.Count; $i++) {
    $row = [int][Math]::Floor($i / 2)
    $col = [int]$fields[$i][2]
    $y = 538.0 + $row * 30.0
    $lx = $(if ($col -eq 0) { 14.3 } else { 620.0 })
    Box $ws $lx $y 250.0 30.0 @{ text = $fields[$i][0]; size = 9; ink = $C_N700 } | Out-Null
    Box $ws ($lx + 256.0) $y 140.0 30.0 @{ name = $fields[$i][1]; size = 9; bold = $true; fmt = '@'; fill = $C_WHITE; border = $C_DIVIDER; halign = $xlCenter } | Out-Null
  }

  # ---- 3. ファイル ---------------------------------------------------------
  Rule $ws 14.3 660.0 1211.5 $C_DIVIDER | Out-Null
  Box $ws 14.3 670.0 400.0 20.0 @{ text = 'ファイル ― どこを読むか (すべて次回起動から)'; size = 8; ink = $C_ACCENT; bold = $true } | Out-Null
  $files = @(
    @('データ (CSV) フォルダー', 'setDataDir'),
    @('統合台帳', 'setLedger'),
    @('ログ', 'setLog')
  )
  for ($i = 0; $i -lt $files.Count; $i++) {
    $y = 698.0 + $i * 32.0
    Box $ws 14.3 $y 250.0 32.0 @{ text = $files[$i][0]; size = 9; ink = $C_N700 } | Out-Null
    Box $ws 270.0 $y 700.0 32.0 @{ name = $files[$i][1]; size = 9; fmt = '@'; fill = $C_WHITE; border = $C_DIVIDER } | Out-Null
  }
  Box $ws 14.3 794.0 250.0 24.0 @{ text = '設定ファイル'; size = 8; ink = $C_N700 } | Out-Null
  Box $ws 270.0 794.0 700.0 24.0 @{ name = 'setFile'; size = 8; ink = $C_N700; fmt = '@' } | Out-Null

  # ---- footer --------------------------------------------------------------
  Rule $ws 0.7 832.0 1238.7 $C_DIVIDER | Out-Null
  Box $ws 0.7 836.0 1238.7 56.0 @{ fill = $C_N100 } | Out-Null
  Button $ws 14.3 850.0 130.0 31.7 '保存' '入力した内容を ReaderDataViewer.json に書き戻します' 'setBtnSave' $true | Out-Null
  Button $ws 156.0 850.0 130.0 31.7 '取消' 'ファイルを読み直して、編集を捨てます' 'setBtnCancel' $false | Out-Null
  Box $ws 300.0 850.0 926.0 31.7 @{ name = 'setNote'; size = 9; bold = $true; ink = $C_ACC800; fill = $C_N100 } | Out-Null

  $ws.Activate()
  $win = $ws.Parent.Windows.Item(1)
  $win.DisplayGridlines = $false
  $win.DisplayHeadings = $false
  $win.Zoom = 100
  $ws.Range('A1').Select() | Out-Null
  return [pscustomobject]@{ ColPx = [Math]::Round($cal.Px, 2); Cols = $cols; Rows = $rowsN }
}
