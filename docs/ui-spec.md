# UI 仕様 (v2) — 画面定義 (`settings.json` の `screen`)

正本は **`docs/ui-reference/v2.html`** (= リポジトリ直下の `Reader Data Viewer v2 (standalone).html`)。
ブラウザで描画して実測した値が `docs/ui-ref-v2-geom.json` で、`build/test_ui_geometry.ps1` はこの表と
製品の実描画を要素ごとに突き合わせる。見本は読むだけで、変更・整形しない。v2 より前の案は
`archive/ui-prototypes/` にある。

v2 の見本から、オーナー指示で変えた点は 3 つだけ。

| 変更 | 理由 |
|---|---|
| ブランド行 (タイトルバー) を出さない | WinForms の枠がそのままアプリの枠。名前は窓のキャプションにある |
| 「台帳更新」「設定」はステータスバーの右端 | ブランド行を外した置き場 |
| 時計を出さない | 不要 |
| 長文の枠の「候補一覧から行を選択すると表示されます。」と、帯の「—」を出さない | 未選択のときは空欄・ラベルだけ |

「内容クリア」は「クリア」。起動時の窓は 840 × 830 CSS px (`card.startSize`) で、見本の 1240 px は
設計幅 (寸法の基準) であって窓の幅ではない。出荷定義のラベルと値は業務の色を抜いたもの (「表A の項目」
「名称」「コード」「参照番号」、`SAMPLE-A-0000001`) で、見本の「取引先」「伝票」「品目」「メーカー」は
使わない (オーナー指示)。

## 1. 画面は JSON から組む

画面は `settings.json` の `screen` (画面定義) の **7 種類の部品** を縦に並べたもの。プログラムに
座標は書いていない。定義は厳格に読まれ、知らないメンバー・型の違う値・何も指していない名前があれば
アプリは起動しない ([settings.md](settings.md))。組込みの既定画面は無い。部品ごとの高さは定義の値 (余白、行高、枠の高さ) から決まり、出荷定義では
見本の CSS 値そのものなので、設計幅では見本の座標に載る (±2 px、`test_ui_geometry.ps1`)。

| type | 内容 | 主な項目 |
|---|---|---|
| `titleBar` | 白いバー 48px (出荷定義では使わない) | `brand`, `tags[]`, `buttons[]` |
| `keyPanel` | 図版 (ラベル + 34px の数値) と入力欄 + ボタン | `figure`, `input` {placeholder, width, maxLength}, `buttons[]` |
| `columns` | 横並び | `weights[]`, `gap`, `stackBelow`, `items[]` (fieldList / textBox) |
| `fieldList` | ラベルと値の行 | `title`, `labelWidth`, `rowHeight`, `rows[]` {label, value} |
| `textBox` | 長文の枠。`lines` 行ぶんの高さで、カードはその周りに自動で合う。`columns` の中でカードが隣より高くなれば枠がカードいっぱいに伸びる。レコード未選択は空欄 | `title`, `lines`, `value` |
| `statusBand` | 判定の帯 84px | `label`, `judgment`, `sub[]`, `joiner`, `height` |
| `statusBar` | 下端のバー 48px | `segments[]` {prefix, value, bold, dot}, `buttons[]` |

カード共通は `card` { `width` 1240, `startSize` [840, 830], `gap` 17, `padding` [14.3, 14.3, 14.2], `font` }。
セクションの `margin` で上書きできる。表示順は配列順。ボタンの `action` は
`search` / `clear` / `workState` / `refreshLedger` / `settings` の 5 つだけで、押したときの処理は C# 側。

### 値 (value binding) は 3 通り

```jsonc
{ "field": "A.a_name" }                                   // 1 列: <表>.<CSV の列名>
{ "fields": ["A.a_rate", "A.a_flag"], "joiner": " ・ " }  // 複数列を連結 (空の列は飛ばす)
{ "state": "searchKey" }                                  // アプリの値 (下表)
```

共通オプション `format` ({"kind":"number","group":true} / {"kind":"date","from":"yyyyMMdd","to":"yyyy-MM-dd"})
と `empty` (レコードはあるが値が空のときの表示。既定 `N/A`、薄色)。
`<表>.<列>` は `data.ledger.columns` のどれかでなければならず、そうでなければ起動しない。列名は CSV の
ヘッダー行そのもので、プログラムは列名を持たない。

`state` の名前: `searchKey`, `candidateCount`, `workState`, `workStateShort`, `rowNumber`,
`appState`, `watchLabel`, `watchDetail`, `ledgerFile`, `ledgerRows`, `ledgerSaved`, `mergeMs`,
`searchMs`, `pid`, `logName`, `clock`, `userName`, `hostName`。

### 判定 (judgments)

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
一致しなければ **undefined**、列が解決できなければ **error**。どちらも暗黙に OK にはならず、
帯は `undefined` (薄灰) / `error` (煉瓦色) の見た目で出る。レコード未選択のときは値を出さない (ラベルだけ。
複数ヒットの副題だけは出る)。

### 作業状態 (workState)

```jsonc
"workState": {
  "store": { "column": "処理済み" },
  "states": [ { "id": "todo", "text": "未処理", "short": "未", "look": "neutral", "stored": "FALSE" },
              { "id": "done", "text": "処理済", "short": "済", "look": "accent",  "stored": "TRUE" } ],
  "initial": "todo",
  "transitions": [ { "from": "todo", "to": "done",
                     "confirm": "表示中のレコード (番号2 = {B.key2}) を処理済にします。よろしいですか?",
                     "done": "処理済の保存が完了しました" } ],
  "button": { "text": "{state}", "tip": "..." }
}
```

判定ではなく作業の状態。共有台帳 xlsx の `store.column` 列へ送る値を `stored` で定義する。ボタンの文言は
現在の状態 (`{state}`)、押すと現在の状態から出る遷移を C# が探し、`confirm` (省略可) を
モーダルで確認してから状態を手元の控えへ保存する。「送信」を押したときだけ共有台帳へ反映する。
遷移が無ければ「このレコードはすでに{state}です」。
台帳の列に定義外の値が入っていたら、その値を赤系でそのまま表示し、遷移は拒否する。
マージでは内容列が一致する行と先にだけある行の保存値を保ち、新規・変更行は `initial` になる。
状態は 3 つ以上・循環する遷移も定義できる (`test_settings_contract.ps1`)。

### 候補一覧 (candidates)

`title`, `hint`, `width` 980, `maxHeight` 340, `rowHeight` 46, `headerHeight` 38, `columns[]`
{header, width (0 = 残り), align, value, bold, muted, render: "text" | "tag", looks {値: look, "*": look}}。
`{state:"workStateShort"}` の列を `tag` で出すと、作業状態の look (未 = neutral, 済 = accent) になる。

## 2. 振る舞い (見本の demo と同じ)

- 検索: 入力が `search.pattern` に一致しないと通知「番号1 が形式 … に一致しません」。照合先は
  `data.ledger.search.columns[]` の複数列、方法は `exact` / `contains`。0 件は「見つかりません」。
  1 件は自動選択。N 件は候補一覧モーダルが開き、行を選ぶまで作業状態を変えない。行を選ぶと閉じて表示する。
  Esc/× なら未選択のまま (帯の副題「複数ヒット（候補 N 件）」)。
- 作業状態ボタン: 上記。保存中は状態語「処理済を保存中」、帯の副題に「(保存中...)」。
- 台帳更新: CSV を再結合して保存済み台帳と比較 (worker)。差異なしは通知、差異ありは「更新の確認」
  モーダル → 更新 → 通知。
- 設定: モーダル (監視対象カード、画面から選ぶ、形式 (正規表現)、パス 3 つ、取消/保存)。保存で
  `settings.json` の `paths` / `search` / `watch` だけを書き戻し、監視を張り直す。`data` と `screen` と
  コメントは触らない。形式が不正、パスが空なら、その欄の下に理由を赤で出して閉じない。
- 起動できない / 続行できない: 設定ファイルが読めないときは Windows のダイアログ (ファイル、行、理由)、
  CSV が読めないときは「データを読めません」モーダル (ファイル名と行番号) を出し、閉じると終了する。
- 通知 (トースト): 右下、タグ「完了」/「エラー」、`toast.durationMs` (3.6 秒) で消える、× で閉じる。
  終了拒否 (保存中) もトーストとログ。
- ステータスバー: `監視中 | メモ帳 接続中（タイトル） | 台帳 ファイル名 ・ N 件` + 右端に台帳更新 / 設定。
  状態語は 起動中 / 更新を確認中 / 台帳を更新中 / 監視中 / {対象} を待機中 / 監視対象なし /
  {状態}を保存中 / 台帳がありません。

## 3. レスポンシブ

窓の大きさにそのまま追従する (`Rdv3Form.Layout1` → `Rdv3Card.Fit`)。

| 窓 | 画面 |
|---|---|
| 広い | 全セクションが横に伸びる |
| 狭い (`columns.stackBelow` 760 未満) | 2 カラムが縦に積まれる |
| さらに狭く図版と入力群が並ばない | 入力群が図版の下の行へ折り返す |
| 高い | 余った高さは長文の枠へ |
| 低い | 長文の枠 (2 行まで。1 行の枠は 1 行) → 行の高さ (30 まで) → 帯 (60 まで) の順に縮み、それでも足りなければ本文だけスクロール。ステータスバーは下端に固定 |
| 入力群が入りきらない幅 | 入力欄が残り幅まで縮む (左へはみ出さない) |
| 帯の副題が長い | 副題を省略記号で詰め、それでも入らなければ副題を出さない。ラベルに重ならず、帯の外に出ない |
| 480 CSS px 未満 | 横スクロール |

モーダルは開いたときの窓に収まる縮尺で出る (下限 0.7)。最小化・最大化は WinForms の枠のまま。

## 4. 描画

- 書体は `card.font` (出荷は **Yu Gothic UI**、600 = Semibold、700 = Bold)。Windows 標準書体だけを名前で
  参照し、フォントファイルは同梱しない。Meiryo も指定できる。
- 文字は **GDI** (`CreateFontIndirect` + `ExtTextOut`, ClearType) で描き、`GetTextExtentPoint32` で
  測る。lfHeight をピクセルで与えるので、DPI スケールや描画先 (窓 / ビットマップ) によらず同じ大きさになる。
  GDI+ の文字描画は横に潰れて見え、TextRenderer は DPI-aware スレッドで点数換算が狂うため使わない。
- 角丸の面は GDI+ (`PixelOffsetMode.Half`) で塗る。モーダルの角は DWM (`DwmSetWindowAttribute`) に
  丸めさせ、Region で切らない (ギザギザと縁の黒ずみの原因)。
- 配色は見本の CSS トークン (`--color-*`)。エラー系だけ見本に無い煉瓦色 `#b04a3e`。

## 5. 受け入れ検査

```
powershell -File build\test_ui_geometry.ps1            # メイン画面 396 通り (約 10 秒)
powershell -File build\test_ui_geometry.ps1 -Quick
powershell -File build\test_ui_geometry.ps1 -Png       # 画像も残す (work\ui-check)
powershell -File build\test_ui_geometry.ps1 -Settings samples\sales-wide\settings.json   # 任意の定義で健全性だけ
powershell -File build\test_settings_geometry.ps1      # 設定・候補一覧・確認・ピッカー 39 通り
```

1. **忠実度** — 設計サイズ (1240 × 974、100%) で 53 要素を `docs/ui-ref-v2-geom.json` と比較。構造要素は
   x/y/w/h、文字に追従する要素 (図版、ボタン、帯の中身) は y/h、右寄せの値は右端と中心線、±2 px。
   設定モーダル 22 要素・ピッカー 9 要素も同じ。
2. **健全性** — サイズ 6 種 (狭い・低い・640 幅を含む) × 拡大率 3 種 × 状態 11 種 × データ 2 種 = 396 通りで
   **交差 0 / はみ出し 0 / 意図しない省略 0**。モーダルは拡大率 3 種 × データ数種 = 39 通り。
   `-Settings` で別の定義を与えると、その定義の列から見本行を作り (書式に合う日付・数値、規則から逆引きした
   ok / ng / 未定義の生値、定義の保存値)、同じ 396 通りを回す。`src/samples` の 5 定義 (販売、製造、施設予約、
   販売ワイド = タイトルバー・3 カラム・帯 2 本・Meiryo、製造コンパクト = 1 カラム・余白詰め) で全部通る。

検査はウィンドウを一度も表示せずに行う (`DrawToBitmap`)。現実的な定義 3 つ (`src/samples`) は
`build/test_samples.ps1` が製品コードに通し、`work/ui-v2/live_samples.ps1` が実配布物で起動・検索・
候補選択・状態保存まで通す (`work/ui-v2/live/sample-*.png`)。実配布物を通常の入口 (`.vbs`) で起動しての
確認は `work/ui-v2/live/*.png` (150% DPI: 起動直後、1 件/0 件/複数ヒット、候補一覧、行選択、確認、
保存、通知、設定、ピッカー、台帳更新、狭い/低い/最大化の窓、`strict-*.png` = typo の設定・無い列・
引用符付き CSV・重複キーで止まるところ、`21-settings-badpattern.png` = 不正な正規表現の拒否)。
5 つの見本定義は `work/ui-v2/live_scenario.ps1` が実配布物で利用動作を一通り通す (`scn-<name>-NN-*.png`、
1 定義 27 枚): 初回起動の台帳作成、形式違いの番号の通知、0 件、候補一覧と Esc、行選択、作業状態の全遷移
(確認モーダル・保存の通知)、最終状態でのもう一押し (拒否の通知)、未定義の判定・結合先なし・空値の行、
クリア、狭い/下限未満/低い/最大化/復帰の窓、台帳更新 (変更なしの通知、CSV を書き換えてからの確認モーダル
→ 列所有規則: 変更行は定義どおり初期化または保持し、未変更行は保持)、設定モーダル、終了、xlsx の中身。

### 結果 (2026-08-23)

```
メイン画面     396 通り: 交差 0 / はみ出し 0 / 省略 0     → PASS
               53 要素 : 許容差超え 0                     → PASS
見本定義 5 つ  各 396 通り: 交差 0 / はみ出し 0 / 省略 0  → PASS
モーダル        39 通り: 交差 0 / はみ出し 0 / 省略 0     → PASS
               設定 22 要素・ピッカー 9 要素: 許容差超え 0 → PASS
settings.json の契約  144 件                              → PASS
現実的な定義 5 組     158 件                              → PASS
実配布物の利用動作 5 定義 × 27 場面 (scn-*.png)           → 通過
終了ガード (実配布物)  12/12                              → PASS
```
