# 設定ファイル `settings.json`

配布物のうち、人が編集するのは 1 つです。

```
ReaderDataViewer.vbs        プログラム (入口。コンソール窓は出ません)
settings.json               設定 — この 1 枚に全部入っています
```

`.vbs` と同じフォルダーに置きます (`.cmd` から起動しても同じ)。`-config <path>` で別の場所を指定できます。
6 つの部分からなります。

| 部分 | 中身 | 誰が書くか |
|---|---|---|
| `paths` | CSV フォルダー、統合台帳、ログの場所 | 人、または「設定」ボタン |
| `search` | 番号の形式 (正規表現)、候補一覧の最大行数 | 人、または「設定」ボタン |
| `watch` | 監視する窓と欄 (UI Automation) | 人、または「設定」ボタン |
| `jobs` | バックグラウンド処理の時間上限 | 人 |
| `data` | CSV、表示名、更新／削除ジョブ、台帳の列 | 人 |
| `screen` | 画面に出すもの (部品、値、判定、作業状態、候補一覧) — [ui-spec.md](ui-spec.md) | 人 |

- `//` 行コメントと `/* ... */` ブロックコメントを書けます。
- 相対パスは `.cmd` のあるフォルダーからの相対です。
- 「設定」ボタンの保存は **`paths` / `search` / `watch` の 3 つだけ**を書き換えます。それ以外の部分と
  コメントは 1 バイトも変わりません (その 3 つの中に書いたコメントだけは残りません)。

## 壊れたファイルでは起動しない

ファイルは**厳格に**読みます。次のどれかに当たると**起動せず**、ファイル名・行・理由をダイアログとログ
(`ReaderDataViewer.log`) に出して終了します。既定値で代わりに動くことも、その値だけ無視して続けることも
ありません。「編集したつもりが効いていない」状態を作らないためです。

| 状況 | 例 |
|---|---|
| ファイルが無い、読めない、構文が壊れている | `settings.json の 74 行目: expected , or }` |
| プログラムの知らないメンバー (typo) | `watch.polMs: is not a member this program knows (pollMs, ...)` |
| 種類や範囲が違う値 | `watch.pollMs: 0 is out of range 5..5000`、`"pollMs": "40"` |
| 正規表現として読めない `search.pattern` | `search.pattern: is not a usable regular expression (...)` |
| 名前が何も指していない | 無い表・列・状態・判定を名指しした、台帳の列に無い列を画面が使った |
| `schema` がこの版と違う | `schema: this program reads schema 3; the file says 2` |
| 定義が CSV と合わない | `tableA.csv に列 a_notes がありません (data.ledger.columns)` — 窓を開く前にヘッダー行で検査 |

「設定」ボタンの保存も同じです。保存の直前にファイルを読み直し、読めなければ**保存を拒否して**理由を
出します (手で壊したファイルに上書きしません)。形式 (正規表現) が不正なとき、パスが空のときは、
モーダルがその欄の下に理由を出して開いたままになります。

## 見本: 3 つの現実的な定義

出荷する `settings.json` は脱色したサンプル (表A/B/C) ですが、`src/samples/` に業務らしい定義が 3 つ
あります。`build/gen_samples.ps1` がそれぞれの CSV を `samples/<name>/` に生成し、`build/test_samples.ps1`
が製品コードに通します (定義が通ること、結合の件数、xlsx の往復、検索、画面の値、判定、作業状態、更新)。

| 名前 | 表 | 特徴 |
|---|---|---|
| `sales` 販売 | 受注明細 (spine) / 得意先 / 商品 | Shift_JIS、日本語の列名、受注番号で 1 対多、日付と金額の書式、判定 3 結果 |
| `factory` 製造 | 検査記録 (spine) / ロット / 製品 / 設備 | UTF-8 BOM、英語の列名、4 表、正規表現の判定、空値の判定、作業状態 3 段階 |
| `booking` 施設予約 | 予約 (spine) / 会員 / 施設 | UTF-8 / LF、日本語の保存値 (未 / 済)、退会済み会員の空欄表示 |
| `sales-wide` | 販売と同じ CSV | 画面の組み方を変えたもの: タイトルバー、3 カラム (重み 2:1:1、中に textBox)、2 段目、判定の帯 2 本、計時のセグメント、Meiryo、幅 1400 |
| `factory-compact` | 製造と同じ CSV | 1 カラムの詰めた画面: 全幅 12 行、余白 8、セクションの margin、縦長の起動サイズ、狭い候補一覧 |

## `data` — CSV と台帳

```jsonc
"data": {
  "encoding": "utf-8",                       // または "shift_jis" など Windows が知る名前
  "tables": {
    "A": { "label": "表A", "file": "tableA.csv", "key": "key1" },
    "B": { "label": "表B", "file": "tableB.csv", "key": "key2" },
    "C": { "label": "表C", "file": "tableC.csv", "key": "key2" }
  },
  "labels": {
    "A.key1": "番号1", "B.key1": "番号1", "B.key2": "番号2", "C.key2": "番号2",
    "ledger": "統合台帳",
    "middle1": "中間1", "middle2": "中間2"
  },
  "jobs": [
    {
      "id": "merge-ledger", "name": "3表を統合して台帳へマージする", "kind": "update",
      "inputs": [ { "table": "A" }, { "table": "B" }, { "table": "C" } ],
      "steps": [
        { "operation": "join", "target1": "B", "target2": "A", "keys": ["B.key1", "A.key1"], "condition": "left", "output": "middle1" },
        { "operation": "join", "target1": "middle1", "target2": "C", "keys": ["B.key2", "C.key2"], "condition": "left", "output": "middle2" },
        { "operation": "merge", "target1": "middle2", "target2": "ledger", "keys": ["B.key2", "B.key2"], "condition": "", "output": "ledger",
          "sourceOnly": "add", "both": "update", "targetOnly": "keep" }
      ]
    }
  ],
  "ledger": {
    "identity": "B.key2",
    "search": { "columns": ["B.key1"], "match": "exact" },
    "columns": {
      "application": [ { "name": "workState", "onSourceChange": "reset" } ],
      "source": [ "B.key1", "B.key2", "A.a_code", ..., "C.c_remark" ]
    }
  }
}
```

| メンバー | 意味 |
|---|---|
| `tables.<id>.label` | ダイアログに出す表の名前 |
| `tables.<id>.file` | `paths.dataDir` の中のファイル名 |
| `tables.<id>.key` | その表の 1 行を特定する列 (**一意**。重複があれば止まる) |
| `labels` | 処理内容とテーブル出力に出す画面向けの名前。内部名は画面へ出さない |
| `jobs[]` | ボタンから参照する更新／削除ジョブ。配列だが、利用者が選ぶ一覧は画面に出さない |
| `jobs[].inputs` | 取り込む表、または外部ファイル。ダイアログはファイルから分かる行数と列一致だけを表示 |
| `jobs[].steps` | 順序どおりの操作。各 `output` は、そのまま後段の `target1` / `target2` にできる |
| マージ段の3指定 | `sourceOnly`: `add` / `ignore`、`both`: `update` / `keep`、`targetOnly`: `keep` / `delete`。`targetOnly` の省略値は `keep` |
| `ledger.identity` | 台帳の1行を特定する、入力表の一意キー |
| `ledger.columns.source` | 取り込み側が持つ列 `<表>.<列>`。xlsx の内容列はこの順 |
| `ledger.columns.application` | アプリが持つ列。`name` は現在 `workState`。`onSourceChange` は入力側の列が変わったとき `reset` (初期値へ戻す) / `preserve` (保持) |
| `ledger.search.columns` | 検索欄が照合する1列以上の配列 (`columns.source` のどれか) |
| `ledger.search.match` | `exact` (完全一致、辞書索引) または `contains` (部分一致) |

操作は、対象の種類が合う順ならジョブの `kind` にかかわらず組み合わせられます。対象は表、ある表に結び付いた
行集合、台帳の3種類です。たとえば `append` や `update` も更新ジョブと削除ジョブのどちらにも置けます。

| `operation` | 対象と指定 | 出力 |
|---|---|---|
| `join` | 2表。`keys` は対象1・対象2の列を順に2つ。`condition` は `match` (内部) / `left` (左外部) / `full` (完全外部) | 結合した表 |
| `append` | 同じ列名・順序の2表 | 対象1の下へ対象2を足した表 |
| `extract` | 表と `where`、表と表＋`keys`、または同じ表から得た2行集合 | 行集合。行集合どうしの条件は `either` / `both` / `exclude` |
| `select` | 1表と `columns`。各列の `as` で改名でき、配列順が出力順 | 選択した列の表 |
| `calculate` | 1表、追加する `column`、`expression` | 計算列を足した表 |
| `aggregate` | 1表、`groupBy`、`aggregates` (`sum` / `count` と `as`) | 集計表 |
| `sort` | 1表、`orders` (`ascending` / `descending`、`text` / `number`) | 並べ替えた表。台帳なら作業状態も同じ行とともに並べ替える |
| `distinct` | 1表と、同一判定に使う `columns` | 最初の行を残して重複を除いた表。台帳なら残した行の作業状態も残す |
| `update` | 1表と、その表から得た行集合。`set` に列と式 | 該当行の列を書き換えた表 |
| `delete` | 1表と、その表から得た行集合 | 該当行を除いた表 |
| `merge` | 元表と台帳、双方の `keys`、3方向の指定 | 追加・更新・保持／削除を適用した台帳 |
| `replace` | 元表と台帳、双方の `keys` | 元表だけで置き換えた台帳 |

`where` は `column / operator / value` です。`operator` は `equals` / `notEquals` / `contains` /
`startsWith` / `endsWith` / `empty` / `notEmpty` / `greater` / `atLeast` / `less` / `atMost`。
式は列名、数値、シングルクォートで囲んだ文字、丸括弧と `+ - * /` だけです。任意のコードや関数は実行しません。
改名列と新しい列は `<output>.<asまたはcolumn>` という名前になり、`labels` に画面向けの名前を置きます。

`merge` は、元だけの行を追加し、両方にある行の取り込み側列を更新し、先だけの行を残す設定にできます。
`targetOnly` を `delete` と明記したときだけ先だけの行を消します。`replace` も更新段として使え、元の行だけで
台帳を作り直します。入力側の列が変わった行でアプリ所有列をどうするかは、列ごとの `onSourceChange` が決めます。

入力列名は CSV のヘッダー行そのものです。名前はファイルを読んだ時点で定義の中で、
起動時に CSV のヘッダー行と突き合わせます。

### CSV の読み方 (厳格)

CSV はそのまま読みます。次のものは**ファイル名と行番号を挙げて拒否**し、起動しません (起動後の
「台帳更新」で見つかれば、理由を出して終了します)。

| 拒否するもの | 理由 |
|---|---|
| 列数がヘッダー行と違う行 | 列がずれる |
| 引用符 `"` で始まる列 | この読み方は引用符を外さないので、黙ってずれるより止める |
| タブ・改行などの制御文字を含む行 | タブは台帳の内部形式 (タブ区切り) を壊し、ほかの制御文字は台帳 (xlsx) に保存できない |
| キー列が空、ASCII 以外、1 行目と幅が違う行 | キーの比較はバイト幅固定。Excel に先頭ゼロを落とされた CSV を止める |
| 同じキーが 2 行にある表 | `key` は一意の約束。どちらを結合するか決められない |
| ヘッダーに同じ列名が 2 つ | どちらの列か決められない |

ヘッダー行の空白は取り除きます。UTF-8 の BOM、CRLF / LF はどちらも読めます。定義が使わない列が CSV に
余分にあっても構いません。

## `search` — 番号の形式

```jsonc
"search": { "pattern": "^[0-9]{8}$", "candidateRowsShown": 100 }
```

入力欄に打った値も監視対象から読んだ値も、**全体が `pattern` に一致したときだけ**番号として確定し
検索します。一致しなければ通知「番号1 が形式 … に一致しません」(「番号1」は `ledger.search.columns` の
先頭列の画面向けラベル)。`candidateRowsShown` は候補一覧に
出す最大行数です (件数は必ず表示)。保存すると即座に有効です。

## `watch` — 監視は UI Automation だけで行う

対象アプリの探索・要素の特定・値の取得は**すべて UI Automation** です。ウィンドウハンドル探索、
メッセージフック、キーボード／マウスのポーリングは使いません (自分の窓の DPI、自分の窓の文字描画、
起動用コンソールを隠すところだけが Win32 です)。

**複数指定できます。有効なものは同時に監視され**、最初に確定した番号が検索を動かします。
1 つの対象は次の経路です。

```
window  ->  path[0]  ->  path[1]  ->  ...  ->  field
```

各段で指定できるもの:

| メンバー | 意味 |
|---|---|
| `automationId` | アプリが付けている安定した識別子。**あるならこれが最良** |
| `className` | ウィンドウ／コントロールのクラス名 |
| `name` | 完全一致 |
| `nameLike` | `*` `?` のワイルドカード |
| `processName` | 実行ファイル名（`.exe` なし） |
| `controlTypes` | `Edit` `Document` `DataGrid` `Pane` `Custom` など (UI Automation が知る名前だけ) |
| `index` | 複数一致したとき何番目を採るか（0 起点） |
| `scope` | `descendants`（既定）または `children` |

`read` は値の取り出し方: `value` (ValuePattern、既定) / `text` (TextPattern) / `name` (Name プロパティ)。

`targets` は**書いたとおりに監視します**。無効 (`enabled: false`) にした対象も設定として保持され、
空の配列なら何も監視しません (画面は「監視対象なし」)。窓か欄のどちらも絞れていない対象は、起動しません。

### 手で書かずに済ませる

設定モーダルの**「画面から選ぶ」**を押すと右下に小さなパネルが出ます。対象アプリの欄に
**カーソルを合わせる**と 種類 / AutomationId / クラス名 / 名前 / プロセス / 読み取れる値が
リアルタイムで出て、**Ctrl + Shift** で取り込みます (Esc で中止)。その要素からウィンドウまで遡り、
AutomationId を持つ段だけを中間パスに残して、ウィンドウ・中間パス・欄・値の取り方を組み立てます。
**見ているだけでは何も変わらず、ファイルに書くのは「保存」だけ**です。

## そのほかのメンバー

| メンバー | 既定 | 意味 |
|---|---:|---|
| `schema` | (必須) | この版は `3` |
| `paths.dataDir` | `data` | CSV のあるフォルダー |
| `paths.ledger` | `ReaderDataViewer-Ledger.xlsx` | 統合台帳 |
| `paths.log` | `ReaderDataViewer.log` | 実行ログ |
| `data.encoding` | `utf-8` | CSV の文字コード |
| `watch.pollMs` | 40 | 監視間隔 |
| `watch.stableMs` | 120 | この時間だけ値が変わらなければ確定 |
| `watch.rebindMs` | 400 | 未接続の対象を探し直す間隔 |
| `watch.preferFocusedWindow` | true | 同じ条件の窓が複数あるとき前面を優先 |
| `jobs.checkTimeoutMs` | 180000 | 更新確認ジョブの上限 |
| `jobs.searchTimeoutMs` | 30000 | 検索ジョブの上限 |
| `jobs.saveTimeoutMs` | 60000 | 作業状態の保存ジョブの上限 |
| `jobs.markOverdueMs` | 180000 | 保存が長引いていると告げるまでの時間 |
| `jobs.pumpMs` | 1000 | 画面側の見張り間隔 |

「既定」は書かなかったときの値です。書いた値が範囲外なら起動しません。

`paths.*` と `data` の変更は**次回の起動から**有効です (台帳・CSV は起動時に読むため)。`search` と
`watch` は設定モーダルの保存で**即座に反映**されます (監視は張り直します)。
