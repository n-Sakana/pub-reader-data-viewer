# settings.jsonの早見表

設定を新規に作るときは、[READMEの最小例](../README.md#quick-start)を写してください。[全項目索引](../README.md#全項目索引)と設定内のK番号で、出荷設定の425箇所へ辿れます。415箇所は可変値と入れ物の出現数であり、独立した調整項目の数ではありません。ソースを読む必要はありません。

設定を書く人がValidateOnlyとRunUpdateを実行し、台帳の値と業務フローを照合してから提出します。手順は[README冒頭](../README.md)。要件を実現できない場合は、理由を返し、実運用用のJSONを提出しません。

## 何をどこへ書くか

| 領域 | 役割 |
|---|---|
| `paths` | 入力フォルダー、共有台帳、PC別ログの場所 |
| `search` | 検索入力の正規表現と候補件数。キー検証とは別 |
| `watch` | 外部ウィンドウをUI Automationで読む設定。CSV到着の監視ではない |
| `jobs` | 処理期限、共有ロック再試行、他PCの変更確認の間隔。時間を過ぎても未確定の保存を強制停止しない |
| `data.tables` | 入力表ID、ファイル、一意になるキー列、表別文字コード、キー検証 |
| `data.types / labels` | 列の型と、人向けの名前。形から型を推測しない |
| `data.jobs` | 一般操作を実行順に並べる。先頭のupdateジョブが起動時とRunUpdateの対象 |
| `data.ledger` | 一意な識別列、検索列、保存列、入力変更時の状態の扱い |
| `screen` | 表示項目、確認状態と遷移、候補一覧、出力初期選択、寸法と書体 |

`schema:3`のJSONCです。UTF-8で保存し、`//`、`/* ... */`、末尾カンマが使えます。存在しないキーでコメントを表さないでください。省略値がある項目は省略できますが、未知のキー、型違い、不明な参照先、壊れたJSONは理由を表示して拒否します。

## 相対パスの基準

`paths.dataDir / ledger / log`は**アプリ一式のフォルダー**が基準です。表とジョブ入力の`file`は**dataDir**が基準です。コマンド引数の`-Config / -DataDir / -Output / -BaselineLedger`は**コマンドを実行する作業フォルダー**が基準です。絶対パスも指定できます。

台帳は全PCから同じ共有ファイルへ向け、ログはPCごとに分けます。入力・設定・プログラム・ログを台帳や出力で上書きする配置にはできません。共有先の別名、ハードリンク、シンボリックリンクの完全な同一性判定はしません。同じファイルへ別名でアクセスする構成を避けてください。

## tablesとkeyValidation

```json
"encoding": "utf-8",
"tables": {
  "A": {"file":"A.csv", "encoding":"shift_jis", "key":["id","part"],
        "keyValidation":{"characters":"unicode","length":"variable"}},
  "B": {"file":"B.csv", "key":["number","item"],
        "keyValidation":{"characters":"unicode","length":"variable"}}
}
```

文字コードは`tables.<ID>.encoding`で表ごとに上書きし、省略すると`data.encoding`、それも省略するとutf-8です。BOMは指定との一致を確かめるもので、自動切替の根拠にはしません。XLSXには適用しません。

見出しより前に表題や出力日の行があるファイルは、`tables.<ID>.headerRow`に見出しの行番号を書きます（例 `"headerRow": 3`）。それより前の行は読み飛ばして件数を通知します。CSVもXLSXも、ジョブの外部入力も同じです。省略すると1行目が見出しです。

BOMなしUTF-16LEは`utf-16`、BEは`utf-16BE`と明示できます。読取失敗時に十分なバイト配列の特徴があれば、UTF-16の候補と設定キーをエラーに添えます。自動切替はしません。[バイト列の確認方法](../README.md#inputs)もREADMEにあります。

`key`は1列の文字列、または列名の配列。複合キーは組合せ全体で一意、文字種と長さは列ごとに確認します。省略規則はASCII、固定長、内容が違う重複キーはエラー、キー空は除外です。

- `characters:unicode`は日本語等のキーを許す。`length:variable`は可変長を許す。
- `empty:skip`（省略値）は空キー行を除外して件数を通知する。複合キーはどれか1列が空なら対象。`error`なら場所と理由を示して停止する。
- 正規化後の全セルが一致する同一キーの行は1行を採り、重複除外数を通知する。
- `duplicates:distinct`は値が異なっても同じキーの先頭行だけ採る。合算ではない。合計するなら明細を区別できるキーを定義し、aggregateを使う。

見出しの前後の空白、値の末尾の空白、全角数字、BOM有無と改行の違いを扱います。数値は全角数字、`"¥66,131"`、`(500)`にも対応。空の型付きセルは空のまま検証を続けますが、数値演算で0を仮定はしません。

日付書式は`data.types`へ`{"type":"date","format":"yyyy-MM-dd"}`等と宣言します。`yyyyMMdd`、`yyyy-MM-dd`、`yyyy/MM/dd`を列ごとに指定できます。1列に複数形式を自動適用しません。

`yyyy年M月d日`も指定できます。例は`2026年9月9日`、`2026年12月31日`。`M`と`d`は1桁も許し、`MM`と`dd`は2桁です。表示を別書式にする場合は`value.format`の`from`と`to`を設定します。

表名は`data.tables.B.label`へ書きます。`data.labels.B`には書けません。labelを省略してもBは表名として登録されるため同じ制限です。列名`B.id`と、新しく作る結果名`joined_AB`・`ledger`は`data.labels`へ書きます。

`ledger`の名前も必須です。`data.labels`に`"ledger":"台帳"`を残してください。組込みの行き先だからラベル不要、とはなりません。

重複見出しは未使用列でも拒否します。labelsやselectでは読取り前の列を選別できません。元データの未使用の重複列を削除するか、必要な列を改名して参照も直します。入力を変更できずその表が必要なら、設定だけで実現できません。

先頭ゼロが消えた値は自動復元しません。ただしゼロの有無で同じ番号だと確定していれば、calculateで照合用の値を作れます。元が固定桁だと分かれば式でゼロ埋めも可能です。[識別規則を確認する条件と実際の式](../README.md#leading-zero)を参照してください。規則が不明な番号を推測しません。

## jobsとledger

[READMEの操作一覧と実行例](../README.md#jobs)に、全12操作、結合・複合keys・抽出・集計の記法があります。例えば複合結合は`keys:[["A.id","A.part"],["B.number","B.item"]]`。入力の見出しが違っても対応づけられます。

入力表のkeyは各行の一意性、結合のkeysは対応づける単位です。入力をid+partで識別しても、要件がidごとの照合ならidで集計してidだけで結合します。件数が多い候補を自動採用せず、業務フローと実際の一致・未一致行で確かめます。

合計は`aggregate`、`groupBy:["B.id"]`、`aggregates:[{"function":"sum","column":"B.amount","as":"amount"}]`。全件を1組にするならgroupByは空配列。countにはcolumnを書きません。集計結果は`output`と`as`で決まる参照（`B.amount`、`output:"totals"`なら`totals.amount`）を`ledger.columns.source`に書いて、その名前のまま台帳へ保存できます。`calculate`の列や`select`の`as`も同じです。保存する列は最後のmerge/replaceの入力に無ければならず、結合していない表の列を書くとエラーになります。

同じ列構成の月次ファイルを1つの台帳にまとめるのは`append`（縦に足す）です。`join`は別表の列を横に付ける操作で、片方にしか無い行は増えません。[READMEの完成例](../README.md#append-files)にappend → distinct → extract → delete → mergeの形があります。明細ファイルだけから伝票単位の台帳を作る例は[こちら](../README.md#detail-only)です。

[4表の完成例](../README.md#four-tables)は、CSV4本・設定全文・期待する数値を含み、集計、左結合、取消フラグによる確認状態のリセットまでそのまま実行できます。

`ledger.identity`は保存列のうち行を一意に識別する列です。入力表のkeyでも、`aggregate`の`groupBy`列のように更新ジョブが作る列でもよく、最後のmerge/replaceのkeysと同じ組合せにします。更新のたびに一意で空欄が無いことを確かめ、重複や空があれば止まります。単一列の既存設定はそのまま使えます。検索列はidentityとは別に指定できます。

`source`は入力が持つ列、`application`のworkStateはアプリが持つ確認状態です。マージでsourceの内容が変わった場合、`onSourceChange:reset`なら状態をinitialへ戻して通知し、preserveなら保持します。取消等を表す入力値も、sourceへ保存し更新ジョブを実行すればこの規則の対象です。**取消専用ではなく、保存するsource列のどれが変わっても対象**です。到着ファイルの自動監視、グループ単位の一括確認はありません。

## 設定・実行・実窓の三段で確認する

配布フォルダーで実行します。出力ファイル名は未使用の名前にします。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\ReaderDataViewer.ps1 -ValidateOnly -Config C:\config\settings.json -DataDir C:\input
powershell -NoProfile -ExecutionPolicy Bypass -File .\src\ReaderDataViewer.ps1 -RunUpdate -Config C:\config\settings.json -DataDir C:\input -Output C:\check\result.json
```

ValidateOnlyは設定と入力の構造・型を確認。RunUpdateは既存の処理本体で更新ジョブを実行し、列・値・状態をJSONへ出します。新規台帳として評価するのが既定で、共有台帳への書込み・送信はしません。既存状態との差を見るときは`-BaselineLedger`で台帳の控えを読みます。そのファイルも変更しません。

`summary.rows / skippedEmpty / skippedDuplicate / resetRows`、各`joins`の`unmatchedLeft / unmatchedRight`、`rows`の値を期待結果と比較してください。中間の抽出結果も`values.<出力名>.rows`で確認できます。exit 0だけで、意図する結果だったとは判定しません。

ValidateOnlyは同じ段階のエラーを集め、先頭に件数を出します。STOPは止まった段階、NOT CHECKEDは未検査の範囲です。失敗exit 3、成功exit 0、未知の引数exit 2。英語の本文に設定キーや入力箇所、期待する値、直し方が出ます。

診断ファイルは`%LOCALAPPDATA%/ReaderDataViewer/logs/feedback.log`。書けなければ`%TEMP%/ReaderDataViewer/logs/feedback.log`へ退避します。現在分と3世代、各4 MiBで循環します。問い合わせには当該実行を含むファイルを渡してください。[記録内容・停止時の読み方・書けない条件](../README.md#feedback-log)も参照してください。

最後にアプリの検証用コピーへ設定とデータを配置し、通常起動して検索・候補選択・状態変更・送信・出力を操作します。窓なしの検査はIME、DPI、監視対象アプリとの相性を確認しません。

## 保存と表示の制約

設定画面はpaths/search/watchを再生成し、その領域のコメントは消えます。他領域とそのコメントは保持します。コメント入りの委託用正本は別に残してください。保存前に再読込と内容照合を行い、別の編集を黙って上書きしません。構造や保存先を変えたら再起動し、共有台帳の契約を全PCで揃えます。

表示寸法はCSS px、文字サイズはptです。文字やラベルを大きくしたら必要な幅と高さも確認します。columns.gapの省略値は17、stackBelowは760で、card.gapとは別です。狭い窓・最長値・使用するDPIで見切れを確認してください。

XLSX入力は先頭のワークシートと保存済み計算値を読み、シート名の指定や再計算はしません。日付セルは`data.types`でdateを宣言した列だけ、その書式の文字列へ変換して台帳に保存します。宣言のない列はシリアル値の数字のままで、台帳は文字列で書くためExcelの書式設定では日付に戻りません。共有台帳はLEDGERの1枚をアプリが管理します。追加シート・非空の余分な列を拒否し、書式や図形の保持を約束しません。CSV出力の数式無効化は、Excelによる番号・日付の自動型変換を防ぎません。
