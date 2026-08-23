# ============================================================================
# gen_samples.ps1 -- realistic dummy data for the three sample definitions
# under src\samples: a set of CSVs + the settings.json + expected.txt (the
# facts a test can check the product against).
#
#   samples\sales      受注明細 / 得意先 / 商品        Shift_JIS, 日本語の列名, CRLF
#   samples\factory    検査記録 / ロット / 製品 / 設備  UTF-8 BOM, 英語の列名, 4 表
#   samples\booking    予約 / 会員 / 施設              UTF-8, LF, 日本語の保存値
#   samples\sales-wide       the sales data under a wide 3-column screen (title bar,
#                            two bands, Meiryo, 1400 wide)
#   samples\factory-compact  the factory data under a one-column compact screen
#
# Everything is generated from a fixed seed, so the files and the facts are
# the same on every run. Nothing here is real: the names are assembled from
# word lists, the phone numbers and mails are made up (example.com).
#
#   powershell -File build\gen_samples.ps1            # into samples\
#   powershell -File build\gen_samples.ps1 -Only sales
#   powershell -File build\gen_samples.ps1 -Force     # regenerate
# ============================================================================
[CmdletBinding()]
param(
  [string] $Root = "",
  [string] $OutDir = "",
  [string] $Only = "",
  [switch] $Force
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
if ([string]::IsNullOrEmpty($Root)) { $Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
if ([string]::IsNullOrEmpty($OutDir)) { $OutDir = Join-Path $Root 'samples' }
$srcSamples = Join-Path $Root 'src\samples'

# ---- word lists --------------------------------------------------------------
$Sei     = @('佐藤','鈴木','高橋','田中','伊藤','渡辺','山本','中村','小林','加藤','吉田','山田','佐々木','山口','松本','井上','木村','林','斎藤','清水')
$SeiKana = @('サトウ','スズキ','タカハシ','タナカ','イトウ','ワタナベ','ヤマモト','ナカムラ','コバヤシ','カトウ','ヨシダ','ヤマダ','ササキ','ヤマグチ','マツモト','イノウエ','キムラ','ハヤシ','サイトウ','シミズ')
$SeiRoma = @('sato','suzuki','takahashi','tanaka','ito','watanabe','yamamoto','nakamura','kobayashi','kato','yoshida','yamada','sasaki','yamaguchi','matsumoto','inoue','kimura','hayashi','saito','shimizu')
$Mei     = @('太郎','花子','健一','美咲','翔太','陽菜','大輔','結衣','拓也','さくら','亮','愛','直樹','彩','健太','恵','悠人','真由美','誠','香織')
$MeiKana = @('タロウ','ハナコ','ケンイチ','ミサキ','ショウタ','ヒナ','ダイスケ','ユイ','タクヤ','サクラ','リョウ','アイ','ナオキ','アヤ','ケンタ','メグミ','ユウト','マユミ','マコト','カオリ')
$Place   = @('東京','大阪','名古屋','横浜','札幌','福岡','仙台','広島','神戸','京都','千葉','埼玉','静岡','新潟','岡山')
$Trade   = @('商事','工業','産業','電機','物産','製作所','建設','運輸','食品','精機','化学','機工','システム','金属','器材')
$Pref    = @('東京都','神奈川県','大阪府','愛知県','福岡県','北海道','埼玉県','千葉県','兵庫県','京都府','静岡県','広島県','宮城県','新潟県','岡山県')
$City    = @('中央区','港区','北区','西区','南区','緑区','青葉区','旭町','本町','栄町','大手町','桜台','新町','若葉町','山手')
# a part and the specs that fit it, so a name reads like a real catalogue line
$Parts = @(
  @('六角ボルト', @('M6×20','M8×25','M10×30','M12×40','M6×15','M8×40')), @('六角ナット', @('M6','M8','M10','M12','M16')),
  @('平ワッシャー', @('M6','M8','M10','M12')), @('ばね座金', @('M6','M8','M10')), @('ベアリング', @('6204ZZ','6305','6001','608ZZ','6203')),
  @('Oリング', @('P-24','P-30','G-45','S-10')), @('オイルシール', @('φ25×40×7','φ30×47×8','φ20×35×7')), @('銅パイプ', @('15A','20A','25A','φ12.7')),
  @('樹脂チューブ', @('φ6×4','φ8×6','φ10×7.5')), @('制御基板', @('CB-120','CB-240','CB-360')), @('LED表示器', @('7SEG 4桁','7SEG 6桁','ドットマトリクス')),
  @('近接センサー', @('PR-18','PR-12','PR-30')), @('圧力計', @('0-1MPa','0-2.5MPa','0-10MPa')), @('流量計', @('DN25','DN40','DN50')),
  @('電磁弁', @('24V 2ポート','100V 2ポート','24V 3ポート')), @('カップリング', @('φ8-φ10','φ10-φ12','φ14-φ16')),
  @('タイミングベルト', @('L=500','L=600','L=800')), @('スプロケット', @('RS40 15T','RS50 20T','RS40 30T')), @('ヒューズ', @('3A','5A','10A')), @('端子台', @('TB-10','TB-20','TB-30'))
)
$Std     = @('JIS B 1180','JIS B 1181','JIS B 1256','JIS B 1251','JIS B 1521','JIS B 2401','SUS304','SS400','C1220T','PTFE','RoHS対応','IP67','CE','UL','φ25 ×1.2t')
$UnitList    = @('個','個','個','箱','m','kg','セット','本')
$Wh      = @('第1倉庫','第2倉庫','東京DC','大阪DC','外部倉庫')
$ShipList    = @('出荷済','出荷済','出荷済','出荷済','出荷済','出荷済','出荷済','出荷済','出荷済','出荷済','出荷済','未出荷','未出荷','未出荷','未出荷','未出荷','一部出荷','一部出荷','キャンセル','保留')
$NoteJ   = @('至急','分納可','納品書2部','検収後請求','直送','立会検査あり','梱包指定あり')
$NoteT   = @('要与信確認','新規開拓','休眠','本社決裁要')
$Machine = @('NC旋盤','マシニングセンタ','三次元測定機','投影機','硬度計','粗さ計','画像検査機','研削盤','プレス機','洗浄機')
$Loc     = @('第1工場 A棟','第1工場 B棟','第2工場','検査室','精密測定室')
$MStatus = @('稼働','稼働','稼働','稼働','休止','校正中')
$PName   = @('ブラケット','シャフト','ハウジング','ギア','フランジ','ピン','カバー','ベース','アーム','スリーブ','ローラー','ホルダー')
$Material = @('SUS304','S45C','A5052','SS400','C3604','POM','SCM435','A6061')
$Shift   = @('昼','昼','昼','夜')
$ResultCode = @('OK','OK','OK','OK','OK','OK','OK','OK','OK','OK','OK','OK','OK','OK','NG-DIM','NG-DIM','NG-APP','RETEST','RETEST','','','HOLD')
$CommentI = @('再測定済','端部にバリ','寸法上限付近','治具交換後','初品','ロット末尾','外観に微細な傷')
$Fac     = @(@('FC01','第1会議室',12,1500,'プロジェクター ・ ホワイトボード','3階'), @('FC02','第2会議室',8,1200,'モニター ・ ホワイトボード','3階'),
            @('FC03','大会議室',40,4000,'プロジェクター ・ マイク ・ 演台','4階'), @('FC04','研修室A',24,2500,'プロジェクター ・ 可動机','2階'),
            @('FC05','研修室B',24,2500,'モニター ・ 可動机','2階'), @('FC06','多目的ホール',120,9000,'音響 ・ 照明 ・ ステージ','1階'),
            @('FC07','音楽スタジオ',6,2000,'ドラム ・ アンプ ・ 防音','地下1階'), @('FC08','調理室',16,3000,'調理台8 ・ オーブン','1階'),
            @('FC09','和室',10,1000,'座卓 ・ 座布団','2階'), @('FC10','テニスコート',4,1800,'ナイター照明','屋外'),
            @('FC11','体育館',60,6000,'バスケットゴール ・ 更衣室','1階'), @('FC12','屋上テラス',30,2200,'ベンチ ・ 日よけ','屋上'))
$MType   = @('一般','一般','一般','ゴールド','法人','学生')
$PayList     = @('支払済','支払済','支払済','支払済','支払済','支払済','未払','未払','未払','返金','保留')
$Route   = @('WEB','WEB','WEB','電話','窓口')
$MemoR   = @('駐車場1台','延長の可能性あり','マイク追加','机レイアウト変更','請求書払い希望','初回利用')

# ---- helpers -----------------------------------------------------------------
function Pick([System.Random] $rng, [object[]] $arr) { return $arr[$rng.Next($arr.Count)] }
function Pad([int] $v, [int] $w) { return $v.ToString().PadLeft($w, '0') }
function DateIn([System.Random] $rng, [int] $y0, [int] $y1) {
  $d = (Get-Date -Year $y0 -Month 1 -Day 1).AddDays($rng.Next(0, 365 * ($y1 - $y0 + 1)))
  return $d.ToString('yyyyMMdd')
}
function AddDays([string] $ymd, [int] $days) {
  return [DateTime]::ParseExact($ymd, 'yyyyMMdd', $null).AddDays($days).ToString('yyyyMMdd')
}
function Person([System.Random] $rng) { $i = $rng.Next($Sei.Count); $j = $rng.Next($Mei.Count); return @(($Sei[$i] + ' ' + $Mei[$j]), ($SeiKana[$i] + ' ' + $MeiKana[$j]), ($SeiRoma[$i] + '.' + $j)) }
function Company([System.Random] $rng) {
  $n = (Pick $rng $Place) + (Pick $rng $Trade)
  switch ($rng.Next(4)) { 0 { return '株式会社' + $n } 1 { return $n + '株式会社' } 2 { return $n + '有限会社' } default { return $n + '株式会社' } }
}
function Phone([System.Random] $rng, [bool] $mobile) {
  if ($mobile) { return ('090-{0}-{1}' -f (Pad $rng.Next(10000) 4), (Pad $rng.Next(10000) 4)) }
  return ('0{0}-{1}-{2}' -f $rng.Next(3, 10), (Pad $rng.Next(1000, 10000) 4), (Pad $rng.Next(10000) 4))
}
function Write-Csv([string] $path, [System.Collections.Generic.List[string]] $lines, [Text.Encoding] $enc, [string] $newline) {
  $fs = New-Object IO.FileStream $path, 'Create', 'Write', 'None', 65536
  $w = New-Object IO.StreamWriter $fs, $enc, 65536
  $w.NewLine = $newline
  foreach ($l in $lines) { $w.WriteLine($l) }
  $w.Flush(); $w.Close()
}
function Write-Facts([string] $path, [System.Collections.Specialized.OrderedDictionary] $facts) {
  $sb = New-Object Text.StringBuilder
  foreach ($k in $facts.Keys) { [void]$sb.Append($k).Append('=').Append([string]$facts[$k]).Append("`r`n") }
  [IO.File]::WriteAllText($path, $sb.ToString(), (New-Object Text.UTF8Encoding($false)))
}
function Prepare([string] $name) {
  $dir = Join-Path $OutDir $name
  $data = Join-Path $dir 'data'
  if (Test-Path -LiteralPath $dir) { Remove-Item -LiteralPath $dir -Recurse -Force }
  New-Item -ItemType Directory -Path $data | Out-Null
  Copy-Item -LiteralPath (Join-Path (Join-Path $srcSamples $name) 'settings.json') -Destination (Join-Path $dir 'settings.json')
  return $dir
}

# ==============================================================================
function Gen-Sales() {
  $dir = Prepare 'sales'
  $data = Join-Path $dir 'data'
  $rng = New-Object System.Random 20240801
  $sjis = [Text.Encoding]::GetEncoding(932)
  $facts = New-Object System.Collections.Specialized.OrderedDictionary

  # 得意先
  $nT = 3000
  $rowsT = New-Object 'System.Collections.Generic.List[string]'
  $rowsT.Add('得意先コード,得意先名,担当者,電話番号,都道府県,与信限度額,取引開始日,締日,備考')
  $tName = @{}
  for ($i = 1; $i -le $nT; $i++) {
    $code = 'T' + (Pad $i 5)
    $name = Company $rng
    $tName[$code] = $name
    $p = Person $rng
    $credit = ($rng.Next(10, 5000) * 10000)
    $note = $(if ($rng.Next(10) -lt 8) { '' } else { Pick $rng $NoteT })
    $rowsT.Add(($code, $name, $p[0], (Phone $rng $false), (Pick $rng $Pref), $credit, (DateIn $rng 2005 2024), (Pick $rng @('末日','20日','25日')), $note) -join ',')
  }
  Write-Csv (Join-Path $data 'tokuisaki.csv') $rowsT $sjis "`r`n"

  # 商品
  $nS = 5000
  $rowsS = New-Object 'System.Collections.Generic.List[string]'
  $rowsS.Add('商品コード,商品名,規格,単位,標準単価,在庫数,倉庫,仕入先,備考')
  $sName = @{}; $sPrice = @{}
  for ($i = 1; $i -le $nS; $i++) {
    $code = 'P' + (Pad $i 7)
    $part = $Parts[$rng.Next($Parts.Count)]
    $name = $part[0] + ' ' + (Pick $rng $part[1])
    $price = $rng.Next(5, 5000) * 10
    $sName[$code] = $name; $sPrice[$code] = $price
    $note = $(if ($rng.Next(10) -lt 9) { '' } else { Pick $rng @('廃番予定','受注生産','長納期') })
    $rowsS.Add(($code, $name, (Pick $rng $Std), (Pick $rng $UnitList), $price, $rng.Next(0, 5000), (Pick $rng $Wh), (Company $rng), $note) -join ',')
  }
  Write-Csv (Join-Path $data 'shohin.csv') $rowsS $sjis "`r`n"

  # 受注明細
  $nO = 4000
  $rowsJ = New-Object 'System.Collections.Generic.List[string]'
  $rowsJ.Add('受注番号,明細番号,得意先コード,商品コード,受注日,数量,単価,金額,納期,出荷状況,備考')
  $matchT = 0; $matchS = 0; $rows = 0
  $probe = $null; $probeBlank = $null; $multiKey = ''; $singleKey = ''
  $probeUndef = $null
  for ($o = 1; $o -le $nO; $o++) {
    $order = 'JU' + (Pad $o 7)
    $r = $rng.Next(100)
    $lines = $(if ($r -lt 45) { 1 } elseif ($r -lt 70) { 2 } elseif ($r -lt 85) { 3 } elseif ($r -lt 95) { 4 } elseif ($r -lt 99) { 5 } else { 6 })
    if ($lines -eq 5 -and $multiKey -eq '') { $multiKey = $order }
    if ($lines -eq 1 -and $singleKey -eq '' -and $o -gt 100) { $singleKey = $order }
    $tcode = $(if ($rng.Next(100) -lt 2) { 'T' + (Pad ($nT + 1 + $rng.Next(200)) 5) } else { 'T' + (Pad ($rng.Next($nT) + 1) 5) })
    $date = DateIn $rng 2024 2025
    for ($l = 1; $l -le $lines; $l++) {
      $id = $order + '-' + (Pad $l 2)
      $scode = $(if ($rng.Next(100) -lt 3) { 'P' + (Pad ($nS + 1 + $rng.Next(500)) 7) } else { 'P' + (Pad ($rng.Next($nS) + 1) 7) })
      $qty = $rng.Next(1, 500)
      $unit = $(if ($sPrice.ContainsKey($scode)) { [int]($sPrice[$scode] * (0.8 + $rng.Next(0, 41) / 100.0)) } else { $rng.Next(5, 5000) * 10 })
      $ship = Pick $rng $ShipList
      $note = $(if ($rng.Next(10) -lt 7) { '' } else { Pick $rng $NoteJ })
      $rowsJ.Add(($order, $id, $tcode, $scode, $date, $qty, $unit, ($qty * $unit), (AddDays $date $rng.Next(7, 60)), $ship, $note) -join ',')
      $rows++
      $hasT = $tName.ContainsKey($tcode); $hasS = $sName.ContainsKey($scode)
      if ($hasT) { $matchT++ }
      if ($hasS) { $matchS++ }
      if ($null -eq $probe -and $order -eq $multiKey -and $hasT -and $hasS) {
        $probe = @{ id = $id; date = $date; qty = $qty; amount = ($qty * $unit); ship = $ship; tname = $tName[$tcode]; sname = $sName[$scode] }
      }
      if ($null -eq $probeBlank -and -not $hasS -and $hasT -and $o -gt 20) { $probeBlank = @{ id = $id; key = $order; index = ($l - 1) } }
      if ($null -eq $probeUndef -and $ship -eq '保留' -and $hasT -and $hasS -and $o -gt 20) { $probeUndef = @{ id = $id; key = $order; index = ($l - 1) } }
    }
  }
  if ($null -eq $probe) { throw 'sales: no probe row (change the seed)' }
  Write-Csv (Join-Path $data 'juchu_meisai.csv') $rowsJ $sjis "`r`n"

  $facts['rows'] = $rows
  $facts['match.T'] = $matchT
  $facts['match.S'] = $matchS
  $facts['probe.multi.key'] = $multiKey
  $facts['probe.multi.hits'] = 5
  $facts['probe.single.key'] = $singleKey
  $facts['probe.single.hits'] = 1
  $facts['probe.none.key'] = 'JU9999999'
  $facts['probe.id'] = $probe.id
  $facts['probe.col.J.受注日'] = $probe.date
  $facts['probe.col.J.数量'] = $probe.qty
  $facts['probe.col.J.金額'] = $probe.amount
  $facts['probe.col.J.出荷状況'] = $probe.ship
  $facts['probe.col.T.得意先名'] = $probe.tname
  $facts['probe.col.S.商品名'] = $probe.sname
  $facts['probe.judge'] = $(switch ($probe.ship) { '出荷済' { 'ok' } '未出荷' { 'ng' } '一部出荷' { 'ng' } 'キャンセル' { 'cancel' } default { 'undefined' } })
  $facts['probe.blank.id'] = $probeBlank.id
  $facts['probe.blank.key'] = $probeBlank.key
  $facts['probe.blank.index'] = $probeBlank.index
  $facts['probe.blank.col'] = 'S.商品名'
  $facts['probe.undef.id'] = $probeUndef.id
  $facts['probe.undef.key'] = $probeUndef.key
  $facts['probe.undef.index'] = $probeUndef.index
  $facts['initial.stored'] = '0'
  Write-Facts (Join-Path $dir 'expected.txt') $facts
  Write-Output ("  sales    rows={0} match.T={1} match.S={2} multi={3} probe={4}" -f $rows, $matchT, $matchS, $multiKey, $probe.id)
}

# ==============================================================================
function Gen-Factory() {
  $dir = Prepare 'factory'
  $data = Join-Path $dir 'data'
  $rng = New-Object System.Random 20240802
  $bom = New-Object Text.UTF8Encoding($true)
  $facts = New-Object System.Collections.Specialized.OrderedDictionary

  # 設備
  $nM = 40
  $rowsM = New-Object 'System.Collections.Generic.List[string]'
  $rowsM.Add('machine_id,machine_name,maker,location,calibrated_on,status')
  $mName = @{}
  for ($i = 1; $i -le $nM; $i++) {
    $id = 'MC' + (Pad $i 2)
    $name = (Pick $rng $Machine) + ' #' + (1 + $rng.Next(6))
    $mName[$id] = $name
    $rowsM.Add(($id, $name, (Company $rng), (Pick $rng $Loc), (DateIn $rng 2023 2025), (Pick $rng $MStatus)) -join ',')
  }
  Write-Csv (Join-Path $data 'machines.csv') $rowsM $bom "`r`n"

  # 製品
  $nP = 300
  $rowsP = New-Object 'System.Collections.Generic.List[string]'
  $rowsP.Add('part_no,part_name,material,drawing_no,spec_lower,spec_upper,unit,remark')
  $partName = @{}; $pLower = @{}; $pUpper = @{}
  for ($i = 1; $i -le $nP; $i++) {
    $pn = 'PN' + (Pad (10000 + $i) 5)
    $name = (Pick $rng $PName) + ' ' + (Pick $rng @('A','B','C','D','L','R','S','M')) + '-' + (Pad $rng.Next(1000) 3)
    $lower = [Math]::Round(5 + $rng.Next(0, 4000) / 100.0, 3)
    $upper = [Math]::Round($lower + 0.02 + $rng.Next(0, 80) / 1000.0, 3)
    $partName[$pn] = $name; $pLower[$pn] = $lower.ToString('0.000'); $pUpper[$pn] = $upper.ToString('0.000')
    $remark = $(if ($rng.Next(10) -lt 8) { '' } else { Pick $rng @('重要保安部品','図面改訂中','客先指定材') })
    $rowsP.Add(($pn, $name, (Pick $rng $Material), ('DWG-' + (Pad $rng.Next(1000000) 6)), $pLower[$pn], $pUpper[$pn], 'mm', $remark) -join ',')
  }
  Write-Csv (Join-Path $data 'products.csv') $rowsP $bom "`r`n"

  # ロット
  $nL = 1500
  $rowsL = New-Object 'System.Collections.Generic.List[string]'
  $rowsL.Add('lot_no,part_no,produced_on,shift,operator,qty,line,remark')
  $lotPart = @{}; $lotDate = @{}; $lotOp = @{}
  for ($i = 1; $i -le $nL; $i++) {
    $lot = 'LT' + (Pad (240000 + $i) 6)
    $pn = 'PN' + (Pad (10000 + 1 + $rng.Next($nP)) 5)
    $lotPart[$lot] = $pn
    $lotDate[$lot] = DateIn $rng 2024 2025
    $op = (Person $rng)[0]
    $lotOp[$lot] = $op
    $remark = $(if ($rng.Next(10) -lt 8) { '' } else { Pick $rng @('材料ロット変更','段取り替え後','試作混在') })
    $rowsL.Add(($lot, $pn, $lotDate[$lot], (Pick $rng $Shift), $op, $rng.Next(50, 2000), ('LINE-' + (1 + $rng.Next(4))), $remark) -join ',')
  }
  Write-Csv (Join-Path $data 'lots.csv') $rowsL $bom "`r`n"

  # 検査記録
  $rowsI = New-Object 'System.Collections.Generic.List[string]'
  $rowsI.Add('inspection_id,lot_no,part_no,machine_id,inspected_at,inspector,measured,lower,upper,result_code,comment')
  $n = 0; $matchL = 0; $matchP = 0; $matchM = 0
  $probe = $null; $probeEmpty = $null; $probeBlank = $null; $probeUndef = $null; $multiKey = ''; $singleKey = ''
  for ($i = 1; $i -le $nL; $i++) {
    $lot = 'LT' + (Pad (240000 + $i) 6)
    $r = $rng.Next(100)
    $count = $(if ($r -lt 30) { 1 } elseif ($r -lt 55) { 2 } elseif ($r -lt 75) { 3 } elseif ($r -lt 88) { 4 } elseif ($r -lt 96) { 6 } else { 8 })
    if ($count -eq 6 -and $multiKey -eq '') { $multiKey = $lot }
    if ($count -eq 1 -and $singleKey -eq '' -and $i -gt 50) { $singleKey = $lot }
    for ($k = 1; $k -le $count; $k++) {
      $n++
      $id = 'INS' + (Pad $n 6)
      $pn = $(if ($rng.Next(100) -lt 2) { 'PN' + (Pad (99000 + $rng.Next(900)) 5) } else { $lotPart[$lot] })
      $mid = $(if ($rng.Next(100) -lt 2) { 'MC99' } else { 'MC' + (Pad (1 + $rng.Next($nM)) 2) })
      $at = $lotDate[$lot] + (Pad (8 + $rng.Next(12)) 2) + (Pad ($rng.Next(0, 12) * 5) 2)
      $hasP = $partName.ContainsKey($pn); $hasM = $mName.ContainsKey($mid)
      $lower = $(if ($hasP) { $pLower[$pn] } else { '0.000' }); $upper = $(if ($hasP) { $pUpper[$pn] } else { '0.000' })
      $measured = ([double]$lower + ([double]$upper - [double]$lower) * ($rng.Next(-20, 121) / 100.0)).ToString('0.000')
      $code = Pick $rng $ResultCode
      $comment = $(if ($rng.Next(100) -lt 75) { '' } else { Pick $rng $CommentI })
      $insp = (Person $rng)[0]
      $rowsI.Add(($id, $lot, $pn, $mid, $at, $insp, $measured, $lower, $upper, $code, $comment) -join ',')
      $matchL++
      if ($hasP) { $matchP++ }
      if ($hasM) { $matchM++ }
      if ($null -eq $probe -and $lot -eq $multiKey -and $hasP -and $hasM -and $code -ne '') {
        $probe = @{ id = $id; at = $at; measured = $measured; code = $code; op = $lotOp[$lot]; pname = $partName[$pn]; mname = $mName[$mid] }
      }
      if ($null -eq $probeEmpty -and $code -eq '' -and $i -gt 10) { $probeEmpty = @{ id = $id; key = $lot; index = ($k - 1) } }
      if ($null -eq $probeBlank -and -not $hasM -and $hasP -and $i -gt 10) { $probeBlank = @{ id = $id; key = $lot; index = ($k - 1) } }
      if ($null -eq $probeUndef -and $code -eq 'HOLD' -and $hasM -and $hasP -and $i -gt 10) { $probeUndef = @{ id = $id; key = $lot; index = ($k - 1) } }
    }
  }
  if ($null -eq $probe) { throw 'factory: no probe row (change the seed)' }
  Write-Csv (Join-Path $data 'inspections.csv') $rowsI $bom "`r`n"

  $facts['rows'] = $n
  $facts['match.L'] = $matchL
  $facts['match.P'] = $matchP
  $facts['match.M'] = $matchM
  $facts['probe.multi.key'] = $multiKey
  $facts['probe.multi.hits'] = 6
  $facts['probe.single.key'] = $singleKey
  $facts['probe.single.hits'] = 1
  $facts['probe.none.key'] = 'LT999999'
  $facts['probe.id'] = $probe.id
  $facts['probe.col.I.inspected_at'] = $probe.at
  $facts['probe.col.I.measured'] = $probe.measured
  $facts['probe.col.I.result_code'] = $probe.code
  $facts['probe.col.L.operator'] = $probe.op
  $facts['probe.col.P.part_name'] = $probe.pname
  $facts['probe.col.M.machine_name'] = $probe.mname
  $facts['probe.judge'] = $(if ($probe.code -like 'OK*') { 'ok' } elseif ($probe.code -like 'NG*') { 'ng' } elseif ($probe.code -eq 'RETEST') { 'retest' } else { 'undefined' })
  $facts['probe.empty.id'] = $probeEmpty.id
  $facts['probe.empty.key'] = $probeEmpty.key
  $facts['probe.empty.index'] = $probeEmpty.index
  $facts['probe.empty.judge'] = 'none'
  $facts['probe.blank.id'] = $probeBlank.id
  $facts['probe.blank.key'] = $probeBlank.key
  $facts['probe.blank.index'] = $probeBlank.index
  $facts['probe.blank.col'] = 'M.machine_name'
  $facts['probe.undef.id'] = $probeUndef.id
  $facts['probe.undef.key'] = $probeUndef.key
  $facts['probe.undef.index'] = $probeUndef.index
  $facts['initial.stored'] = 'OPEN'
  Write-Facts (Join-Path $dir 'expected.txt') $facts
  Write-Output ("  factory  rows={0} match.L={1} match.P={2} match.M={3} multi={4} probe={5}" -f $n, $matchL, $matchP, $matchM, $multiKey, $probe.id)
}

# ==============================================================================
function Gen-Booking() {
  $dir = Prepare 'booking'
  $data = Join-Path $dir 'data'
  $rng = New-Object System.Random 20240803
  $utf8 = New-Object Text.UTF8Encoding($false)
  $facts = New-Object System.Collections.Specialized.OrderedDictionary

  # 施設
  $rowsF = New-Object 'System.Collections.Generic.List[string]'
  $rowsF.Add('施設コード,施設名,定員,時間料金,設備,階,備考')
  $fName = @{}; $fCap = @{}; $fRate = @{}
  foreach ($f in $Fac) {
    $fName[$f[0]] = $f[1]; $fCap[$f[0]] = $f[2]; $fRate[$f[0]] = $f[3]
    $rowsF.Add(($f[0], $f[1], $f[2], $f[3], $f[4], $f[5], '') -join ',')
  }
  Write-Csv (Join-Path $data 'facilities.csv') $rowsF $utf8 "`n"

  # 会員
  $nM = 2500
  $rowsM = New-Object 'System.Collections.Generic.List[string]'
  $rowsM.Add('会員番号,氏名,フリガナ,会員種別,入会日,電話番号,メール,住所,備考')
  $mName = @{}
  for ($i = 1; $i -le $nM; $i++) {
    $id = 'M' + (Pad $i 8)
    $p = Person $rng
    $mName[$id] = $p[0]
    $addr = (Pick $rng $Pref) + (Pick $rng $City) + (1 + $rng.Next(9)) + '-' + (1 + $rng.Next(30)) + '-' + (1 + $rng.Next(20))
    $note = $(if ($rng.Next(10) -lt 9) { '' } else { Pick $rng @('要配慮','団体代表','割引適用') })
    $rowsM.Add(($id, $p[0], $p[1], (Pick $rng $MType), (DateIn $rng 2015 2025), (Phone $rng $true), ($p[2] + (Pad $rng.Next(1000) 3) + '@example.com'), $addr, $note) -join ',')
  }
  Write-Csv (Join-Path $data 'members.csv') $rowsM $utf8 "`n"

  # 予約
  $nR = 9000
  $rowsR = New-Object 'System.Collections.Generic.List[string]'
  $rowsR.Add('予約番号,会員番号,施設コード,利用日,開始時刻,終了時刻,人数,料金,支払状況,予約経路,メモ')
  $matchM = 0; $matchF = 0
  $hits = @{}
  $seen = @{}
  $probe = $null; $probeBlank = $null; $probeUndef = $null
  for ($i = 1; $i -le $nR; $i++) {
    $id = 'R24-' + (Pad $i 6)
    $mid = $(if ($rng.Next(100) -lt 3) { 'M0999' + (Pad $rng.Next(10000) 4) } else { 'M' + (Pad (1 + $rng.Next($nM)) 8) })
    $fc = $(if ($rng.Next(100) -lt 1) { 'FC99' } else { $Fac[$rng.Next($Fac.Count)][0] })
    $start = 9 + $rng.Next(11)
    $hours = 1 + $rng.Next(4)
    $hasM = $mName.ContainsKey($mid); $hasF = $fName.ContainsKey($fc)
    $people = $(if ($hasF) { 1 + $rng.Next([Math]::Max(1, [int]$fCap[$fc])) } else { 1 + $rng.Next(10) })
    $fee = $(if ($hasF) { [int]$fRate[$fc] * $hours } else { 1000 * $hours })
    $pay = Pick $rng $PayList
    $memo = $(if ($rng.Next(10) -lt 7) { '' } else { Pick $rng $MemoR })
    $rowsR.Add(($id, $mid, $fc, (DateIn $rng 2024 2025), ((Pad $start 2) + ':00'), ((Pad ($start + $hours) 2) + ':00'), $people, $fee, $pay, (Pick $rng $Route), $memo) -join ',')
    $idx = $(if ($seen.ContainsKey($mid)) { $seen[$mid] } else { 0 })
    $seen[$mid] = $idx + 1
    if ($hasM) { $matchM++; if ($hits.ContainsKey($mid)) { $hits[$mid]++ } else { $hits[$mid] = 1 } }
    if ($hasF) { $matchF++ }
    if ($null -eq $probe -and $hasM -and $hasF -and $memo -ne '') { $probe = @{ id = $id; mid = $mid; fee = $fee; pay = $pay; mname = $mName[$mid]; fname = $fName[$fc]; people = $people } }
    if ($null -eq $probeBlank -and -not $hasM -and $hasF -and $i -gt 20) { $probeBlank = @{ id = $id; key = $mid; index = $idx } }
    if ($null -eq $probeUndef -and $pay -eq '保留' -and $hasM -and $hasF -and $i -gt 20) { $probeUndef = @{ id = $id; key = $mid; index = $idx } }
  }
  Write-Csv (Join-Path $data 'reservations.csv') $rowsR $utf8 "`n"

  # a member with many reservations, one with exactly one
  $multiKey = ''; $multiHits = 0; $singleKey = ''
  foreach ($k in ($hits.Keys | Sort-Object)) {
    if ($hits[$k] -ge 7 -and $multiKey -eq '') { $multiKey = $k; $multiHits = $hits[$k] }
    if ($hits[$k] -eq 1 -and $singleKey -eq '') { $singleKey = $k }
  }
  if ($multiKey -eq '') { foreach ($k in ($hits.Keys | Sort-Object)) { if ($hits[$k] -gt $multiHits) { $multiKey = $k; $multiHits = $hits[$k] } } }

  $facts['rows'] = $nR
  $facts['match.M'] = $matchM
  $facts['match.F'] = $matchF
  $facts['probe.multi.key'] = $multiKey
  $facts['probe.multi.hits'] = $multiHits
  $facts['probe.single.key'] = $singleKey
  $facts['probe.single.hits'] = 1
  $facts['probe.none.key'] = 'M99999999'
  $facts['probe.id'] = $probe.id
  $facts['probe.col.R.料金'] = $probe.fee
  $facts['probe.col.R.人数'] = $probe.people
  $facts['probe.col.R.支払状況'] = $probe.pay
  $facts['probe.col.M.氏名'] = $probe.mname
  $facts['probe.col.F.施設名'] = $probe.fname
  $facts['probe.judge'] = $(switch ($probe.pay) { '支払済' { 'ok' } '未払' { 'ng' } '返金' { 'refund' } default { 'undefined' } })
  $facts['probe.blank.id'] = $probeBlank.id
  $facts['probe.blank.key'] = $probeBlank.key
  $facts['probe.blank.index'] = $probeBlank.index
  $facts['probe.blank.col'] = 'M.氏名'
  $facts['probe.undef.id'] = $probeUndef.id
  $facts['probe.undef.key'] = $probeUndef.key
  $facts['probe.undef.index'] = $probeUndef.index
  $facts['initial.stored'] = '未'
  Write-Facts (Join-Path $dir 'expected.txt') $facts
  Write-Output ("  booking  rows={0} match.M={1} match.F={2} multi={3}({4}) probe={5}" -f $nR, $matchM, $matchF, $multiKey, $multiHits, $probe.id)
}

# a screen variant over another sample's data: the same CSVs and facts, a
# different settings.json
function Gen-Variant([string] $name, [string] $base) {
  $baseDir = Join-Path $OutDir $base
  if (-not (Test-Path -LiteralPath (Join-Path $baseDir 'expected.txt'))) { throw ($name + ' needs ' + $base + ' first') }
  $dir = Prepare $name
  Copy-Item -Path (Join-Path (Join-Path $baseDir 'data') '*') -Destination (Join-Path $dir 'data')
  Copy-Item -LiteralPath (Join-Path $baseDir 'expected.txt') -Destination (Join-Path $dir 'expected.txt')
  Write-Output ("  {0,-16} the {1} data under its own screen" -f $name, $base)
}

# ==============================================================================
$all = @('sales', 'factory', 'booking', 'sales-wide', 'factory-compact')
$want = $(if ([string]::IsNullOrEmpty($Only)) { $all } else { @($Only) })
Write-Output ("samples -> " + $OutDir)
$sw = [Diagnostics.Stopwatch]::StartNew()
foreach ($name in $want) {
  $marker = Join-Path (Join-Path $OutDir $name) 'expected.txt'
  if ((Test-Path -LiteralPath $marker) -and -not $Force) { Write-Output ("  {0,-16} already generated (-Force to redo)" -f $name); continue }
  switch ($name) {
    'sales' { Gen-Sales }
    'factory' { Gen-Factory }
    'booking' { Gen-Booking }
    'sales-wide' { Gen-Variant 'sales-wide' 'sales' }
    'factory-compact' { Gen-Variant 'factory-compact' 'factory' }
    default { throw ('unknown sample ' + $name) }
  }
}
Write-Output ("done in {0:N1} s" -f $sw.Elapsed.TotalSeconds)
