# 画面定義 — `settings.json` の `screen`

画面はプログラムに書いていません。`settings.json` の `screen` が並べた部品を、標準の WinForms
コントロールで組み立てます。ここは `screen` の書き方と、それが画面でどう振る舞うかの資料です。
`settings.json` の残り (`paths` / `search` / `watch` / `jobs` / `data`) は [settings.md](settings.md)。

定義は厳格に読みます。知らないメンバー、型の違う値、何も指していない名前 (無い列・状態・判定・ジョブ)
があればアプリは起動せず、行番号と理由を出します。組込みの既定画面はありません。

## 1. 部品 (sections)

`screen.sections` は表示順の配列で、部品は 8 種類です。

| type | 何が出るか | 主な項目 |
|---|---|---|
| `titleBar` | 窓のキャプション (`brand`) と、下部のボタン列へ足すボタン | `brand`, `tags[]`, `buttons[]` |
| `keyPanel` | 図版 (ラベルと大きな値) と、入力欄 + ボタン | `title`, `figure` {label, value}, `input` {label, placeholder, width, maxLength}, `buttons[]` |
| `columns` | 横並び (中は `fieldList` / `textBox` だけ) | `weights[]`, `gap`, `items[]` |
| `fieldList` | ラベルと値の行 | `title`, `labelWidth`, `rowHeight`, `rows[]` {label, value} |
| `textBox` | 複数行の読み取り専用の枠。`lines` 行ぶんの高さ | `title`, `lines`, `value` |
| `statusBand` | 判定の帯。中央に判定結果、右に副題 | `label`, `judgment`, `sub[]`, `joiner`, `height` |
| `sendBar` | 「未送信 N 件」と送信ボタンの帯。ボタンは `sendChanges` 1 つ | `height`, `value`, `buttons[]` |
| `statusBar` | 下端のステータスバー (区画) と、下部のボタン列 | `height`, `segments[]` {prefix, value, bold, dot}, `buttons[]` |

共通の `margin` (数値 1〜4 個、CSS 風) で部品の外側の余白を上書きできます。`card` は
{ `width`, `startSize` [幅, 高さ], `gap`, `padding`, `font`, `fontSize`, `keyValueFontSize`,
`judgmentFontSize`, `unsearchedFontSize` } で、`startSize` が起動時の窓のクライアント寸法、`font` は
名前で指定する Windows の書体です。`gap` は枠どうし、枠と中身、下部ボタンの上下に共通する間隔、
`padding` は画面外周の余白です。後ろの3サイズは、現在番号・判定結果・未検索の強調に使います。

ボタンの `action` は 9 つです。押したときの処理はプログラム側にあります。

| action | 何をするか |
|---|---|
| `search` | 入力欄の値で検索する。`primary: true` にすると Enter で押せる |
| `clear` | 入力と結果表示を消す |
| `workState` | 表示中のレコードの作業状態を切り替える (文言は `workState.button`) |
| `tableExport` | テーブル出力ダイアログを開く |
| `updateRecords` | `job` に書いた更新ジョブのダイアログを開く |
| `deleteRecords` | `job` に書いた削除ジョブのダイアログを開く |
| `sendChanges` | 手元の未送信の作業状態を共有台帳へ送る |
| `refreshLedger` | 更新の確認を今すぐやり直す |
| `settings` | 設定ダイアログを開く |

`job` を書けるのは `updateRecords` / `deleteRecords` だけで、`data.jobs[]` の `id` を指します。
利用者がジョブを選ぶ一覧は画面に出ません。

## 2. 値 (value)

```jsonc
{ "field": "A.a_name" }                                   // 1 列: <表>.<入力表の列名>
{ "fields": ["A.a_rate", "A.a_flag"], "joiner": " ・ " }  // 複数列を連結 (空の列は飛ばす)
{ "state": "searchKey" }                                  // アプリの値 (下表)
```

共通オプションは `format` ({"kind":"number","group":true} / {"kind":"date","from":"yyyyMMdd","to":"yyyy-MM-dd"})
と `empty` (レコードはあるが値が空のときの表示。既定 `N/A`、薄い色)。`<表>.<列>` は
`data.ledger.columns.source` のどれかでなければ起動しません。列名は入力表のヘッダー行そのものです。

`state` の名前: `searchKey`, `candidateCount`, `workState`, `workStateShort`, `rowNumber`,
`appState`, `watchLabel`, `watchDetail`, `ledgerFile`, `ledgerRows`, `ledgerSaved`, `pendingCount`,
`mergeMs`, `searchMs`, `pid`, `logName`, `clock`, `userName`, `hostName`。

## 3. 判定 (judgments)

```jsonc
"status1": {
  "source": { "field": "B.b_status" },
  "rules": [ { "equals": ["DONE", "OPEN"], "result": "ok" },
             { "equals": ["HOLD", "VOID"], "result": "ng" } ],
  "results": { "ok": { "text": "OK", "look": "ok", "icon": "check" },
               "ng": { "text": "NG", "look": "ng" },
               "undefined": { "text": "未定義", "look": "undefined" },
               "error": { "text": "エラー", "look": "error" } }
}
```

規則は上から順に `equals` (完全一致の列挙) / `pattern` (正規表現) / `empty` (true で空値)。どれにも
一致しなければ **undefined**、参照した列が台帳に無ければ **error**。どちらも OK にはなりません。
帯は ok を緑、ng と error を赤系、undefined を通常色で出します。レコード未選択のときは「未検索」です。

## 4. 作業状態 (workState)

```jsonc
"workState": {
  "trigger": "automatic",
  "store": { "column": "処理済み" },
  "states": [ { "id": "todo", "text": "未処理", "short": "未", "look": "neutral", "stored": "FALSE" },
              { "id": "done", "text": "処理済", "short": "済", "look": "accent",  "stored": "TRUE" } ],
  "initial": "todo",
  "transitions": [ { "from": "todo", "to": "done",
                     "confirm": "表示中のレコード (番号2 = {B.key2}) を処理済にします。よろしいですか?",
                     "done": "処理済の保存が完了しました" },
                   { "from": "done", "to": "todo", "done": "未処理に戻しました" } ],
  "button": { "text": "{state}(&W)", "tip": "..." }
}
```

- `store.column` は共有台帳 (xlsx) の先頭列の見出し。`stored` はその列に書く値です。
- `initial` は新しい行と、入力側の内容が変わった行 (`onSourceChange: reset` のとき) の状態です。
- ボタンは現在の状態から出る遷移を 1 つだけ探し、`confirm` があればモーダルで確認してから、手元の
  控え (LOCALAPPDATA) へ保存します。共有台帳へ書くのは「送信」を押したときだけです。
- 遷移が無ければ「このレコードはすでに{state}です」。台帳の列に定義外の値があれば、その値を赤系で
  そのまま出し、遷移を拒否します。
- `confirm` / `done` の中では `{state}`、`{key}`、`{<表>.<列>}` が使えます。
- 状態は 3 つ以上でも、循環する遷移でも定義できます。

## 5. テーブル出力 (export)

`export.defaultFields` は、テーブル出力ダイアログを開いたときと「既定に戻す」を押したときに選ぶ列を、
表示順で並べます。`data.labels` に画面名を持つ台帳列と、作業状態を表す `$work` を指定できます。
存在しない列、出力一覧にない列、重複した列は設定エラーになります。

```jsonc
"export": {
  "defaultFields": ["B.key1", "B.key2", "A.a_name", "B.b_status", "$work"]
}
```

出力ダイアログでは条件を複数追加でき、すべてに一致する行だけを書き出します。型指定のない文字列列は
「を含む／と等しい／で始まる／を含まない」、`data.types` の日付・数値列は両端を含む「範囲」です。
日付は `DateTimePicker` から選びます。出力項目の選択と絞り込み項目は独立しています。

`trigger` は `automatic`（既定。監視対象からの1件検索で自動遷移）または `manual`（ボタン操作だけ）です。

## 6. 候補一覧 (candidates)

`title`, `hint`, `width`, `maxHeight`, `rowHeight`, `headerHeight`, `columns[]`
{header, width, align, value, bold, muted, render: "text" | "tag", looks {値: look, "*": look}}。
候補が 2 件以上のとき、この列でモーダルの一覧 (`ListView`) を出します。同じ列定義を、レコード更新で
未処理へ戻った行の一覧にも使います。

## 7. 振る舞い

- 検索: 入力欄の値も監視対象から読んだ値も、全体が `search.pattern` に一致したときだけ番号として
  確定します。一致しなければ通知「{検索列のラベル} が形式 … に一致しません」。照合先は
  `data.ledger.search.columns[]`、方法は `exact` / `contains`。0 件は「見つかりません」、1 件は
  そのまま表示、2 件以上は候補一覧が開き、行を選ぶまで作業状態は変わりません。
- `workState.trigger` が `automatic` で、監視対象から読んだ番号が 1 件だけ当たったときは、確認なしで
  最初の遷移 (未処理 → 処理済) を自動で行います。`manual` または手検索では行いません。
- 作業状態ボタン: 上記。保存中は状態語「{状態}を保存中」、帯の副題に「(保存中...)」。保存の結果が
  出るまで次の変更と終了を受け付けません。
- 送信: 確認モーダルのあと、共有のロックファイルを取り、共有台帳を読み直してから、手元の変更を
  同じ番号の行へ当てて書き、符丁ファイルの版を進めます。当てられなかった変更 (行が無い、中身が変わった)
  は一覧で知らせ、手元に残します。
- レコード更新 / レコード削除: ジョブの入力ファイル・処理内容・書き出し先を表示するだけの
  ダイアログのあと、同じロックの下で実行します。入力側の内容が変わって初期状態へ戻った行は一覧で示します。
- 他の端末が台帳を書き換えたとき: 符丁ファイルの版で気づき、送信なら screen.work.states の状態名を
  使って「{誰} が N 件を{初期状態からの遷移先}、M 件を{初期状態}にしました」とステータスバーへ
  一時表示します。続けて届いた通知は同じ区画を更新するため、窓を増やしません。更新なら
  「切り替えますか」と尋ねてから読み直します。
- テーブル出力: 出す列とAND条件を選んで、統合台帳を CSV (UTF-8、BOM 付き) に書き出します。
- 値が無い表示欄は空白のままにし、代替のハイフンは表示しません。
- 設定: `paths` / `search` / `watch` だけを書き戻し、監視を張り直します。`data` と `screen` と
  コメントは触りません。
- 起動できない / 続行できない: 設定ファイルが読めないときは Windows のダイアログ (ファイル、行、理由)、
  入力表や台帳が読めないときは「データを読めません」(ファイル名と行番号) を出して終了します。
- 通知: エラーは親画面に結び付いた Windows の警告ダイアログで出します。完了や該当なしなどの
  お知らせは、操作を止めないようステータスバーへ一時表示します。
- ステータスバーの状態語: 起動中 / 更新を確認中 / 台帳を更新中 / レコードを削除中 / 送信中 /
  台帳を読み直し中 / 台帳が空くのを待機中 / 監視中 / {対象} を待機中 / 監視対象なし /
  {状態}を保存中 / 台帳がありません。

## 8. 描画と検査

画面は標準の WinForms 部品だけで組み、自前の描画 (`OnPaint`) はしません。
`Application.EnableVisualStyles` も呼ばず、Windows の既定の描画に任せます。

| 画面 | WinForms |
|---|---|
| 検索・番号・表の項目・メモ・備考の枠 | `GroupBox` |
| 入力欄 | `TextBox` |
| 読み取り専用の欄 | `Label` + `Fixed3D`、複数行は `TextBox` (`ReadOnly`) |
| 検索・クリア・下部のボタン | `Button` |
| 作業状態 | `CheckBox` (`Appearance = Button`) |
| 候補一覧・取り込むデータ・処理内容・戻った行の一覧 | `ListView` (`View = Details`) |
| ステータスバー | `StatusStrip` (区画は `ToolStripStatusLabel`) |

検査は窓を出さずに行います。

```
powershell -File build\test_settings_geometry.ps1      # 部品の種類と構成 (36 件)
powershell -File build\capture_headless.ps1            # 主画面を DrawToBitmap で撮る
powershell -File build\capture_dialogs_headless.ps1    # 各ダイアログを撮る
powershell -File build\test_settings_contract.ps1      # 定義の読込みと拒否 (209 件)
powershell -File build\test_samples.ps1                # 見本定義 5 組 (158 件)
powershell -File build\test_exit_guard.ps1             # 実配布物で起動・検索・保存・送信・終了 (35 件)
```
