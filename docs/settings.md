# 設定ファイル `ReaderDataViewer.json`

配布は **2 枚**です。

```
ReaderDataViewer.vbs     プログラム (入口。コンソール窓は出ません)
ReaderDataViewer.json    この設定
```

`.vbs` と同じフォルダーに置きます (`.cmd` も同梱してあり、そちらから起動しても同じです)。**無くても動きます**（下に書いた既定値で動作し、
ログにそう残ります）。`-config <path>` で別の場所を指定することもできます。

- `//` 行コメントと `/* ... */` ブロックコメントを書けます。
- 値が範囲外なら**その値だけ**既定に落ちて、理由がログに出ます。ファイル全体は生きます。
- 相対パスは `.cmd` のあるフォルダーからの相対です。
- アプリの右上「設定」ボタンから GUI で編集して、このファイルに書き戻せます。

## 監視は UI Automation だけで行う

対象アプリの探索・要素の特定・値の取得は**すべて UI Automation** です。ウィンドウハンドル探索、
メッセージフック、キーボード／マウスのポーリングは使いません（アプリ自身のウィンドウの
DPI 設定と、起動用コンソールを隠す 2 か所だけが Win32 です）。

## 監視対象 (`watch.targets`)

**複数指定できます。有効なものは同時に監視され**、最初に確定した番号が検索を動かします。
1 つの対象は次の経路です。

```
window  ->  path[0]  ->  path[1]  ->  ...  ->  field
```

業務アプリの欄はたいてい数階層下（タブ → パネル → グリッド → セル）にあるので、
`path` で段階的に絞ります。各段で指定できるもの:

| メンバー | 意味 |
|---|---|
| `automationId` | アプリが付けている安定した識別子。**あるならこれが最良** |
| `className` | ウィンドウ／コントロールのクラス名 |
| `name` | 完全一致 |
| `nameLike` | `*` `?` のワイルドカード |
| `processName` | 実行ファイル名（`.exe` なし） |
| `controlTypes` | `Edit` `Document` `DataGrid` `Pane` `Custom` など |
| `index` | 複数一致したとき何番目を採るか（0 起点） |
| `scope` | `descendants`（既定）または `children` |

`read` は値の取り出し方です。業務アプリの欄は 3 通りに割れます。

| 値 | 使う UIA | 典型 |
|---|---|---|
| `value` | ValuePattern | 一般的な入力欄（既定） |
| `text` | TextPattern | リッチエディット、ドキュメント |
| `name` | Name プロパティ | 読み取り専用のラベル表示 |

### 例: 業務アプリの特定欄

```jsonc
{
  "enabled": true,
  "name": "在庫照会",
  "window": { "processName": "LobApp", "nameLike": "在庫照会*", "scope": "children" },
  "path": [ { "automationId": "tabMain" }, { "automationId": "pnlSearch" } ],
  "field": { "automationId": "txtBarcode", "controlTypes": ["Edit"] },
  "read": "value"
}
```

### 手で書かずに済ませる

この JSON は画面右上の**「設定」**から GUI で編集できます。タブは 3 枚:

| タブ | 中身 |
|---|---|
| **監視対象** | 対象の一覧 (追加 / 複製 / 削除) と、選んだ対象のウィンドウ・欄・値の取り方・中間パス |
| **動作** | 番号の桁数と数字のみ、監視間隔・確定待ち・再接続間隔・候補の表示行数、前面優先 |
| **ファイル** | データ (CSV) フォルダー、統合台帳、ログ |

書き方が分からないときは、**「画面から選ぶ」**を押すと右下に小さなパネルが出ます。

1. 対象アプリの欄に**カーソルを合わせる** → 種類 / AutomationId / クラス名 / 名前 /
   プロセス / 読み取れる値がリアルタイムで表示されます
2. **Ctrl + Shift** で取り込み (パネルの「この要素を対象にする」でも同じ。Esc で中止)

その要素からウィンドウまで遡り、AutomationId を持つ段だけを中間パスに残して、
ウィンドウ・中間パス・欄・値の取り方を自動で組み立てます。手で書くより確実です。

## そのほかのメンバー

| メンバー | 既定 | 意味 |
|---|---:|---|
| `paths.dataDir` | `data` | CSV 3 本のあるフォルダー |
| `paths.ledger` | `ReaderDataViewer-Ledger.xlsx` | 統合台帳 |
| `paths.log` | `ReaderDataViewer.log` | 実行ログ |
| `key.length` | 8 | 番号の桁数（画面の入力欄もこれに従う） |
| `key.digitsOnly` | true | false なら英数字も可 |
| `search.candidateRowsShown` | 10 | 候補一覧に出す最大行数（件数は必ず表示） |
| `watch.pollMs` | 40 | 監視間隔 |
| `watch.stableMs` | 120 | この時間だけ値が変わらなければ確定 |
| `watch.rebindMs` | 400 | 未接続の対象を探し直す間隔 |
| `watch.preferFocusedWindow` | true | 同じ条件の窓が複数あるとき前面を優先 |
| `jobs.checkTimeoutMs` | 180000 | 起動時の更新確認ジョブの上限 |
| `jobs.searchTimeoutMs` | 30000 | 検索ジョブの上限 |
| `jobs.saveTimeoutMs` | 60000 | 処理済み保存ジョブの上限 |
| `jobs.markOverdueMs` | 180000 | 保存が長引いていると告げるまでの時間 |
| `jobs.pumpMs` | 1000 | 画面側の見張り間隔 |

`paths.*` の変更は**次回の起動から**有効です（台帳と CSV は起動時に読むため）。
それ以外は設定ダイアログの保存で**即座に反映**されます（監視は張り直します）。
