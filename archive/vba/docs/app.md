# 実用版 (app) — 設計と実測

比較ベンチ (`docs/results.md`, `docs/results2.md`) の次段階として作った**実用アプリ 2 方式**の記録。
比較版 (`src/`, `src/v2/`, `benchmarks/`) は凍結のまま一切変更していない。実用版の要件が
旧比較契約 (毎回再読込・無キャッシュ等) と衝突する箇所は、実用版側にのみ新契約を適用している。

本書は **v2 (台帳分離版)** の正本。v1 (台帳を親ブック内シートに保持していた版) の設計と実測は
末尾の「v1 の記録 (退役)」に保全してある。v1 の成果物・生ログは `work\archive-v1\` に退避済み。

| 方式 | 配布ルート | 索引 |
|---|---|---|
| VBA 標準版 | `dist\app-vba\` | `Scripting.Dictionary` (遅延バインド) |
| C# 標準版 | `dist\app-csharp\` | `Dictionary<string, List<int>>` |

自前ハッシュ版は**採用していない**。一対多比較 (`docs/results2.md`) の結論のとおり 10 万件では
標準の連想配列で実用域であり、実用版は保守性を優先して標準版のみとする。自前ハッシュ版・
4 方式比較の成果物とソースは比較証拠としてそのまま残してある。失敗時に別方式へ fallback する
経路は存在しない。

## 配布物

```
dist\app-vba\
  ReaderDataViewer.xlsm            小さな FE (UI シート + META シートのみ、~240 KB)
  ReaderDataViewer-Ledger.xlsx     統合台帳ブック (LEDGER + META シート。BE だけが読み書き)
  ReaderDataViewer-Ledger.state    台帳の sidecar ミラー (UTF-16LE。保存のたびに原子的に併記)
  data\tableA.csv / tableB.csv / tableC.csv

dist\app-csharp\
  ReaderDataViewer.cmd             自己完結 1 本 (PowerShell 起動 → in-box csc でコンパイル)
  ReaderDataViewer-Ledger.xlsx     統合台帳 (閲覧用。更新承認・処理済みでアプリが書き換える)
  data\tableA.csv / tableB.csv / tableC.csv
```

- 実行時の追加要件なし: 別 exe・常設 script・XLL・COM 登録・管理者権限は使わない。
  **VBA 版の実行コードは Win32 API (`Declare`) も `Shell` も一切使わない** — VBA と COM だけ
  で動く (根拠と実測は「Win32 / Shell を使わない実装」節)。子プロセスは自分で起動する
  不可視 Excel (BE) 1 本だけで、それは COM の `CreateObject` で作る。
- 実行ログ `ReaderDataViewer.log` (と VBA 版の `ReaderDataViewer.log.worker.log`) は
  各ルート直下に実行時生成される (配布部品ではない)。
- CSV は既存の検証済み `data-100k` (10 万行 × 3 表、`gen_data2.ps1` 生成、`expected.txt` が oracle)。

## 統合台帳のデータモデル

台帳 = A+B+C を全部統合した最終台帳。**表 B の明細 1 行 = 統合レコード 1 行** (10 万行)。
A-B は番号1、B-C は番号2 で結合する既存モデルのまま。

列は 29: `処理済み` + 内容 28 列。

```
処理済み | key1 key2 | a_code..a_note (9) | b_slip..b_memo (8) | c_item..c_remark (9)
```

key1 / key2 は結合で一致が保証されるので 1 回だけ持つ。30 フィールド表示 (表A/B/C 各 10 項目)
は key1 / key2 を各表へ再掲して復元する。A / C が不一致の行は該当フィールドを空にする
(この合成データでは全行一致)。

### 安定 identity = 番号2 (key2)

processed 引継ぎと「その正確な 1 行」の特定には行 identity が要る。

- 番号1 は**正規に一対多** (伝票 1 つに明細 1..5 行) なので identity にならない。
- 番号2 は実データモデル上、**表 B の明細行そのもの**を指す: 生成器 (`gen_data2.ps1`) は
  key2 を 1..N の置換 (全単射) として明細行に払い出し、表 C は key2 で 1:1。既存 4 方式も
  「番号2 は一意という前提」で候補選択を key2 で確定している (`docs/results2.md`)。
- よって **統合レコードの identity = key2**。コード上も carry / 処理済み確定 / 候補選択の
  キーはすべて key2。

### processed 引継ぎ規則

CSV 更新で台帳を再構築するとき、新しい統合レコードの processed は:

```
引継ぎ = 旧台帳に同じ key2 の行があり、かつ内容 28 列が完全一致し、かつ旧行が TRUE
内容が変わった行・新規の行 = FALSE (再処理が必要な行として扱う)
```

carried / reset / new / dropped の件数をログに残す。検索しただけでは processed は変わらない。
「処理済み」操作は確認ダイアログの後、表示中の 1 行 (key2 で特定) だけを TRUE にして永続化する。

## 台帳の保存構造 (v2: 台帳ブック + sidecar)

VBA 版の台帳は親ブックから完全に分離した。**正本は台帳ブック** (`ReaderDataViewer-Ledger.xlsx`、
人が開いて読める閲覧物) で、その**完全ミラーが sidecar** (`ReaderDataViewer-Ledger.state`、
ヘッダ 1 行 + 「processed TAB 内容 28 列」× 全行の TSV。文字コードは **UTF-16LE**
— VBA が変換なしで読み書きできる唯一の符号化で、理由は「Win32 / Shell を使わない実装」)。BE は保存のたびに
「台帳ブック Save → sidecar を tmp→rename で原子的に書換」の順で両方を書く。

起動時の更新判定は sidecar の高速路で行う:

```
[高速路]  merge+compose した新内容  vs  sidecar の全行 (行単位の完全文字列比較)
          — sidecar を信用する条件: 台帳ブックの mtime が sidecar に記録された
            保存時 mtime と一致すること。mtime はこの「保守側へ倒すガード」だけに
            使い、内容判定そのものには使わない (時刻で「同じ」とは決して言わない)。
[低速路]  sidecar が無い/壊れている/mtime 不一致 → 台帳ブック本体を ReadOnly で
          開いて全読み → 比較 → sidecar を書き直して自己修復 (ログに明示)。
[異常]    台帳ブックが無い/読めない → ledgerbad (作り直すか明示確認。無言の再構築はしない)
```

差分なしの起動では**台帳ブックを一切開かない** (v1 で起動の支配項だった 16.8 秒の
Excel ブック全読みが消える)。processed 状態と検索対象の内容は sidecar から採用する —
sidecar はブックとペアで書かれた保存済み状態そのもので、差分なし = merge 結果と完全一致を
行単位で確認済み、である。

crash 窓 (ブック保存後・sidecar 書換前に落ちた場合) は次回起動の mtime 不一致が拾い、
低速路がブックから状態を回復する。ドリフトが黙って採用される経路はない。

## 起動時の更新フロー

```
起動 → 「更新を確認中」アニメーション (FE は即 idle へ、pump が更新)
     → [FE] 最初の pump tick で BE を起動: CreateObject("Excel.Application") →
       worker book を開く → bootstrap を Run。bootstrap は OnTime を仕掛けて即 return し、
       以後の作業は BE 自身のメッセージポンプで走る (FE 占有 = spawn_ms 約 1.0 秒のみ)
     → [BE] CSV3 本 → 読込 → 標準索引 → A-B/B-C 統合 → 29 列 compose = 新しい統合内容
     → [BE] 保存済み状態 (sidecar 高速路 / ブック低速路) と行単位で内容比較
       ※ CSV の時刻・サイズは見ない。内容の比較だけ (mtime は sidecar の自己検証のみ)。
差分なし → ステータスに「更新はありません」(非モーダル通知) → READY
差分あり → MsgBox / MessageBox で更新確認
   承認 → [BE] processed を key2+内容一致で引継ぎ → [BE] 台帳ブックを書換・Save →
          [BE] sidecar 書換 → READY (FE は一切書かない)
   拒否 → 保存済み台帳のまま → READY
台帳が読めない/無い → エラー表示 + 作り直すか明示確認 (無言の作り直しはしない)
確認エラー/タイムアウト/worker 死亡/spawn 失敗 → エラー表示。fallback はしない
```

READY 後はメモ帳監視 (既存 UIA 検知) が自動で始まり、手動検索も使える。
検索は BE メモリ (C# はプロセス内メモリ) の現在台帳に対して行う。

## アーキテクチャ

### C# 版 (`src/app/csharp`) — WinForms + プロセス内 worker スレッド

| ファイル | 責務 |
|---|---|
| `Rdv3Core.cs` | CSV 読込 (バイト列 + 行オフセット)・時計・定数 |
| `Rdv3Index.cs` | 標準 `Dictionary<string, List<int>>` 索引 (結合用・検索用の唯一の索引) |
| `Rdv3Ledger.cs` | 台帳モデル: 統合行の生成、内容比較、processed 引継ぎ、key1 検索 |
| `Rdv3Xlsx.cs` | 台帳 xlsx の読み書き (ZipArchive 直書き、temp 書き→Replace の安全置換) |
| `Rdv3Jobs.cs` | worker スレッド + ジョブ列、run ID、タイムアウト検出 |
| `Rdv3Watch.cs` | UIA でメモ帳検知 (v2 からの意図的コピー、検知規則同一) |
| `Rdv3Ui.cs` | WinForms 画面: ステータス / 結果・候補 / 操作、更新確認アニメーション |
| `Rdv3Text.cs` | 画面文字列 (非 ASCII はここだけ) |
| `Rdv3App.cs` | 入口・状態機械・計時境界・実行ログ |

UI スレッドは重処理を持たない。読込・索引・統合・比較・引継ぎ・永続化・検索は worker
スレッドで実行し、結果だけ `Invoke` で UI に反映する。run ID の合わない古い結果は破棄して
ログに残す。子プロセスは起動しない。二重起動は台帳パス単位の Mutex で拒否する。
v2 での変更はロジックなし・計測の明示のみ (`startup boot_to_ready_ms` 行、processed の
`e2e_ms` = 確認クリック → 永続化+画面反映完了)。

### VBA 版 (`src/app/vba`) — FE/BE 分離 + 台帳分離

別プロジェクトで実証済みの非同期骨格 (FE/BE 分離) をそのまま採用し、v2 で台帳を
親ブックから外へ出した。根拠はそちらで取った実測:
**worker → FE のクロスプロセス同期 COM 書込みは、セル編集中の FE にパークし、OLE リトライが
キュー入力を飢餓させ、ESC の効かない恒久フリーズを起こす**。worker 側の対策では原理的に
直らないため、書込み方向を反転する。

```
  BE (不可視の別 Excel プロセス)               FE (利用者の Excel、~240 KB の小さな UI ブック)
  ─────────────────────────────               ─────────────────────────
  常駐自己ペーシングループ (40ms + 時間基準)     Application.OnTime 1 秒 pump
   ├ UIA でメモ帳検知 (検知規則は既存同一)        ├ 1 tick = aggregate 1 回読み + 小描画のみ
   ├ CSV 読込・Scripting.Dictionary 索引         ├ tick 間は VBA スタック完全解放
   ├ A-B/B-C 統合・compose・内容比較・carry      ├ セル編集中は Excel が OnTime を保留 →
   ├ 台帳ブック + sidecar の読み書き・Save       │   編集終了後に追いつく (正常仕様)
   ├ 「処理済み」の永続化 (セル+Save+sidecar)    ├ 起動 spawn は最初の tick で同期実行
   ├ key1 検索・候補抽出                         │   (bootstrap が OnTime を仕掛けて即 return
   └ 結果を atomic にファイルへ publish          │    するので占有は spawn_ms だけ)
       (tmp 書き → Kill → Name)                  └ FE ローカル SheetChange は
                                                     pump 再アームの watchdog のみ
  BE → FE の COM 呼出し: ゼロ / FE → BE の COM 参照: ゼロ (bootstrap 完了時に手放す)
  FE → BE: 種別ごとの小さい request ファイル (latest-wins, version 付き) + stop flag
  死活: FE リース lock を BE が probe / BE リース lock を FE が probe (対称・PID 不使用)
```

| モジュール | 役割 |
|---|---|
| `modRdv3Chan.bas` | ファイルチャネル: atomic publish、aggregate、request、FE/BE リース、be_done フラグ、session 掃除 |
| `modRdv3Be.bas` | BE 本体: 常駐ループ (時間基準の cadence)、マージ、状態ロード (sidecar/ブック)、比較、carry、台帳ブック書換+Save、sidecar、mark 永続化、検索、監視、publish |
| `modRdv3Host.bas` | FE 側の BE 所有: 埋込 worker book の展開、`CreateObject` による BE 起動と bootstrap 呼出し、リース probe による死活、stop (flag → BE の self-Quit を待つ。kill 経路は無い) |
| `modRdv3App.bas` | FE 本体: 短い boot、OnTime pump、dispatch、描画呼出し、ログ (台帳への書込・Save は存在しない) |
| `modRdv3Ui.bas` / `modRdv3Uia.bas` / `modRdv3Engine.bas` / `modRdv3Spec.bas` | 描画 / UIA 検知 / マージエンジン / 定数 (v2 由来) |
| `ThisWorkbook` / UI シートのイベント | 自ブックの SheetChange・SheetActivate → pump 再アーム watchdog、Hyperlink クリック → dispatch。入口だけを持ち、処理は `.bas` 側 (BE 通知には使わない) |

FE ブックのモジュールは Spec/Chan/Host/Ui/App の 5 本のみ (クラスモジュールは無い)。Engine/Be/Uia は
worker book (META に base64 埋込、実行時に `%TEMP%\rdv3\` へ session 固有名で展開) だけが持つ。

要点:

- **spawn は「短い同期呼び出し + BE 側 OnTime」**: FE の最初の pump tick が
  `CreateObject("Excel.Application")` で自分専用の不可視 Excel を作り、worker book を
  AutomationSecurity=low で開き、`Rdv3BeBootstrap` を Run する。bootstrap は
  `Application.UserControl = True` を立て、BE リースを開き、`Application.OnTime` を
  仕掛けて即 return する — 重い仕事は一つもしない。だから FE の占有は
  「プロセス起動 + worker book を開く」の spawn_ms (実測 1,000〜1,109 ms、n=8) だけで、以後は
  BE 自身のポンプで走る。実証済みの FE/BE プロジェクト (xltoolrack `JobHost.StartJob`) と
  同じ形。**FE は返ってきた時点で COM 参照を手放す**: 忙しい BE への COM 呼出しは
  Excel の「サーバービジー」で FE をブロックし得るため、参照は一切保持しない
  (参照ゼロの不可視 Excel が自己終了しないのは UserControl=True のおかげ)。
- **死活はリースファイル (両方向)**: FE が `fe_lease.lock` を deny-all で開いたまま保持し
  BE が lock 試行で判定する / BE も `be_lease.lock` を同じように保持し FE が probe する。
  どちらのプロセスが落ちても OS が即 lock を解放するので、PID も COM 呼出しも要らない。
  tombstone (`fe_gone`) と stop flag も併用。SessionId は timestamp 生成 (hWnd は不使用)。
- **承認後の台帳反映は BE が完結**: carry → 台帳ブックの全面書換 (16,384 行ブロックの
  Value2、テキスト書式で前ゼロ保持) → Save → sidecar。FE には統計値の APPLY 通知と READY
  だけが届く。v1 の part 転送 / FE materialize / FE Save は全廃した。
- **「処理済み」も BE が完結**: FE は要求を出し「処理済み: TRUE (保存中...)」を表示、
  BE がブックのセル書換 + Save + sidecar 後に MARK (marked) で確定を返し、FE が
  「処理済み: TRUE」に確定させる。失敗は markerr で明示し表示も FALSE へ戻す。
  最初の mark は BE が台帳ブックを開くぶん遅い (以後はブックを開いたまま保持し、
  BE 終了時に閉じる。保存はそのたび行う)。
- **FE ブックは何も保存しない**: 台帳が外に出たので FE に永続状態がない。閉じるときは
  `Saved = True` で無音 (UI セルの書込で保存プロンプトを出さない)。
- **待機カーソル**: pump callback 実行中に一瞬 WAIT カーソルが出ることは許容する
  (カーソル抑止の作り込みは適用範囲外)。

## 未確定の 1 件保存を黙って落とさない終了保護 (両版)

「処理済み」は 1 件ずつその場で永続化する。**結果が決まっていない唯一の窓**は、書込みが走って
いる間 — C# は worker スレッドの `Rdv3Xlsx.Write` (実測 641〜790 ms)、VBA は BE の
セル書換 + `Workbook.Save` + sidecar (実測 1.2〜9 秒) — で、その間に終了すると
「押したのに保存されたか分からない」まま消える。そこだけを塞ぐ。

規則は 3 つだけで、両版で同じ:

1. **保存中は新しい処理済み操作を受け付けない** (同じファイルを自分の保存の下から書き換えない)。
2. **保存中の終了要求はキャンセルし、理由を出す** — 画面のエラー行と、`保存中のため終了できません`
   というダイアログ。何も後片付けしないので、セッションはそのまま生きている。
3. **成功または失敗として確定した後に終了できる**。確定は「保存できた」と「保存に失敗した」の
   両方を含む。どちらでも解放する。

溜めない・キューにしない・一括保存しない。**保護するのは既に走っている 1 件の保存だけ**で、
終了プロトコルは足していない (2 回目の閉じる操作で普通に閉じる)。

| | C# 版 | VBA 版 |
|---|---|---|
| 状態 | `savingMark` (UI スレッドだけが触る) | `m_markPending` (FE、pump だけが触る) |
| 保存中の表示 | 状態「処理済みを保存中」+ 「処理済み: TRUE (保存中...)」+ ボタン無効 | 状態「処理済みを保存中」+ 「処理済み: TRUE (保存中...)」 |
| 終了要求 | `FormClosing` で `e.Cancel = True` | `Workbook_BeforeClose` で `Cancel = True` (BE 停止も session 掃除もしない) |
| 確定 | 書込みの成功/例外 + job が投げた場合も解放 | BE の **MARK** `marked` / `markerr` (専用スロット + `req=` で要求と対応) |
| 確定できない場合 | 60 秒で「保存が長引いています」と表示し、**保持は続ける** (managed job は中断できず、書けたかどうかを騙って言えない) | BE 死亡 (リース消失) または 180 秒で**失敗として確定**し、「保存未確定」と表示して解放 (ブックが永久に閉じられなくならないように) |

VBA 側で「BE 死亡・タイムアウトは失敗として確定」まで入れているのは、FE が閉じられなくなる
状態を作らないため。C# 側は同一プロセス内なので、その経路自体が無い。

### 実動作の確認 (2026-08-15、両版とも実物で)

`build\test_exit_guard.ps1`。dist の scratch コピー上で、アプリを起動 → 検索 → 処理済みを
確認 → **保存が走っている最中に終了要求 (WM_CLOSE)** → 拒否と表示を確認 → 保存確定を待つ →
もう一度終了 → 台帳に反映が残っていることを確認。**23 チェック全通過** (C# 12 / VBA 11)。
生ログ `work\guard-20260815-072128.tsv`。

```
C# 版
  07:21:33.170  P1 processed  save started key2=00089897 (exit held until it is decided)
  07:21:33      WM_CLOSE (保存中)
  07:21:33.205  -  exit  close refused: a processed save is still in flight
                ダイアログ「保存中のため終了できません」/ ウィンドウ生存
  07:21:33.814  P1 processed  key2=00089897 value=TRUE persist_ms=637.68 e2e_ms=650.74
  07:21:33.812  P1 exit  processed save decided (saved); exit released
                2 回目の終了要求 → 終了。台帳の processed=TRUE 行数 = 1

VBA 版
  07:21:46  - processed  dispatched key2=00089897 (exit held until it is decided)
  07:21:46    WM_CLOSE を Excel メインウィンドウへ (= 利用者と同じ閉じ方)
  07:21:46  - exit  close refused: a processed save is still in flight (key2=00089897)
              ダイアログ / ブックは開いたまま (BE 停止も session 掃除も走らない)
  07:21:50  P1 processed  key2=00089897 value=TRUE save_ms=1197.27 e2e_ms=4052.78
  07:21:50  P1 exit  processed save decided (saved); exit released
              その後の Close は受理。台帳の processed=TRUE 行数 = 1
```

### 保存中に検索すると確定通知が消える競合 — 原因と修正 (VBA 版、2026-08-15)

**症状 (修正前)**: 実経路ベンチで**保存要求の 200 ms 後に検索を 1 回**投げると、3 方式 3/3 で
**FE が保存完了を受け取らなかった**。保存できているのに 180 秒 (`MARK_TIMEOUT_S`) 終了を拒み続け、
その後「確定しません」と失敗側で確定する。証拠一式は `work\race-evidence-before\` に保全してある。

```
67869.2  req|mark|ver=23    29998|1|...        ← [処理済み] の保存要求
68337.4  req|search|ver=24  00066644|...       ← 保存の 200 ms 後に検索 1 回
69468.2  agg|RESULT|ver=46  res=marked;...     ← 保存完了 (BE は正しく返している)
69909.6  agg|RESULT|ver=47  res=single;...     ← 441 ms 後、同じ種別が上書き
```

**原因**: 終了保護ではなく**チャネルの構造**。`agg.dat` は**種別ごとに最新 1 件**しか持たないので、
種別が配送スロットそのものになっている。検索応答と保存完了が同じ `RESULT` を共有していたため、
後に publish された方が先の方を消していた (逆向きも同じで、保存完了が未読の検索応答を消すこともある)。
`modRdv3Chan` のヘッダにある「BE が返したのに FE が描画しなかった検索が 39 回中 2 回」と同じ根で、
終了保護はそれを**可視化しただけ**。

**修正**: 種別を配送クラスで分けた。

- **latest-wins** (`STATE` / `CHECK` / `READY` / `APPLY` / `BEERR` / `RESULT`) — 最新だけに意味がある。
  古い方が消えても損はない (心拍、最新の検索応答)。
- **must-deliver** (`MARK`、新設) — 「その 1 件がディスクに載った」という確定通知。**専用スロット**なので
  検索応答に場所を奪われない。次に書き換わるのは**次の保存要求のとき**だけで、それは終了保護が
  「前の確定を FE が受け取るまで」認めない。よって未読の確定が消える経路が**構造的に無い**。
- 確定通知は `req=<答える要求の版番号>` を載せ、FE は**待っている要求と一致するときだけ**確定として扱う
  (`HandleMark`)。一致しないものはログに残して捨てる (別要求への誤対応・重複の防止)。
- 検索は**抑制も遅延もしていない**。待ち時間の延長やリトライで隠してもいない。

**検証 (2026-08-15、実配布物・実クリック・現行 book 方式)**: `build\bench_e2e.ps1 -Mode race`。
生ログ `work\race-20260815-213005.tsv`。3 つの順序 (保存直前 / 保存中 200 ms / 保存完了直後) ×
2 セッション × 2 回 = 12 回 + 対照 (保存のみ 2 回・検索のみ 2 回) = **16 回、72 チェック全通過**。

| 順序 | 回数 | チャネルで後から publish されたもの | 結果 |
|---|---:|---|---|
| 保存中 (200 ms 後) | 4 | 検索応答 (4/4) | 全通過 |
| 保存直前 | 4 | **保存完了** (4/4) | 全通過 |
| 保存完了直後 | 4 | 検索応答 (4/4) | 全通過 |
| 保存のみ / 検索のみ | 2 / 2 | — | 全通過 |

毎回、(a) 検索結果が正しいキーで描画される (b) 保存完了が描画される (c) 確定通知から**最大 1,463 ms**で
終了保留が解除される (180 秒経路には一度も落ちない) (d) 独立リーダーで対象行のフラグが TRUE
(e) 確定通知はちょうど 1 件・別要求扱いの破棄ゼロ、を機械確認している。
回帰: 終了保護テスト **23/23 通過**、連続 10 件保存も従来どおり (最大 3,177.2 ms、10/10 読み直し成功)。
同じ `cont-busy` 検査で、修正前に 3 方式 3/3 で出ていた「FE が確定を描画しなかった」注記は**消えた**。

### 起動と登録の実測は `docs/save-methods.md` へ

「更新確認ジョブ」と「処理済み登録ジョブ」を、実配布物・実プロセス・実クリックで外から
計測した結果 (3 方式 × 3 条件 × 10 回、および C# 版 10 回) は `docs/save-methods.md`。
なお、その実測で得た同じ段階の値は 2026-08-14 の表 (下記) と**大きく違う**:

| 段階 (VBA 版) | 2026-08-14 (中央値) | 2026-08-15 (中央値) | 比 |
|---|---:|---:|---:|
| マージ 8 工程 | 1,295 ms | 1,137 ms | ×1.1 |
| compose | 1,835 ms | 422 ms | ×4.3 |
| 保存状態ロード (sidecar) | 395 ms | 122 ms | ×3.2 |
| 処理済み永続化 (BE の save_ms) | 9,000 ms | 1,195 ms | ×7.5 |
| 台帳ブックを開く | 約 11,000 ms | 約 1,900 ms | ×5.8 |

**CPU 律速のマージだけがほぼ同じ**で、確保・大バッファ・I/O が絡む段階が軒並み 3〜7 倍速い。
コードも測定境界も変えていないので、差は機械の状態 (空きメモリ・キャッシュ・電源) 側にある。
2026-08-14 当時の状態は再現できないため、**どちらが「正しい」かは決められない**。同じ日に
測った値どうしだけを比べること。

## Win32 / Shell を使わない実装 (VBA 版)

実用 VBA 版の実行コードから `Declare` (Win32 API) と `Shell` を全廃した。**FE/BE 分離・
非同期・Excel COM 利用の構成はそのまま**で、Win32 が担っていた機能を COM/VBA の手段に
置き換えてある。凍結中の比較版 (`src\vba`, `src\v2\vba`) は対象外 (Declare のまま)。

| 旧 (Win32 / Shell) | 新 (VBA + COM) | 実測 |
|---|---|---|
| `QueryPerformanceCounter` | VBA `Timer` (Double 差分・日跨ぎ補正) | 分解能 3.906 ms (= 1/256 秒) |
| `Sleep` (40ms 等) | WMI イベントソースの `NextEvent(ms)` タイムアウト待ち | 40ms 要求で 45〜54ms・スピンなし |
| `MultiByteToWideChar` / `WideCharToMultiByte` (sidecar UTF-8) | sidecar を **UTF-16LE** にして VBA の `Byte()`⇄`String` + `Get`/`Put` | 同一条件の読み比べで UTF-16 23.4ms 対 UTF-8 26.0ms (差なし)。代わりにファイルは 2 倍 |
| `GetForegroundWindow` + `FindWindowEx` + `IsWindowVisible` (メモ帳探索) | UIA `GetFocusedElement` → `ControlViewWalker` で親を辿り ClassName=`Notepad` | 取得 11.7ms・以後の poll は 25 回で 0.0ms |
| `GetWindowTextW` (タイトル) | UIA `IUIAutomationElement.CurrentName` | — |
| `GetCurrentProcessId` (ログの同定) | `Application.hWnd` (`Rdv3SelfId`) | — |
| `OpenProcess`+`GetExitCodeProcess` (BE 死活) | BE 側リースファイルの lock probe (FE リースの鏡像) | OS が即解放 |
| `TerminateProcess` (最後の手段の kill) | **廃止**。stop flag → BE の self-Quit を待ち、駄目なら報告する | — |
| `Shell("cscript ... spawn.vbs")` | FE 最初の tick で `CreateObject("Excel.Application")` → worker book → bootstrap Run | spawn_ms 1.00〜1.11 秒 (n=8) |

判断の根拠 (すべてこの機で実測):

- **`Application.Wait` は 40ms 待機に使えない**。粒度は 1 秒で秒境界に丸まる (+40ms 目標は
  即 return、+400ms 目標は 861ms 待った)。`Application.OnTime` はさらに悪く、1 秒未満の
  目標は約 1ms で発火する = スピンになる。WMI の `NextEvent(ms)` はタイムアウトまで
  スレッドを寝かせる本物の待機なので、これだけが要件を満たした。**待機が効かなくなったら
  空回りに落ちず、明示エラーで停止する** (BE ループ開始時に自己テストし、以後は毎回の待機で検出)。
- **sidecar を UTF-8 のままにする道は無かった**。`ADODB.Stream.ReadText` は 22MB を
  一括読みしても **22.2 秒**、mscorlib の `System.Text.UTF8Encoding` は VBA の
  `CreateObject` で作れない (0x80131500)。UTF-16LE にすると変換そのものが消える —
  同じ本文を同じプロセスで読み比べると **UTF-16 23.4ms 対 UTF-8 (Get+MB2WC) 26.0ms** で
  速度は互角、つまり Win32 を落としても変換コストは増えない。代わりにファイルが 2 倍
  (22.3MB → 44.6MB) になり、その分だけアプリ内のロード段が重くなる (後述の +203 ms)。
- **UIA だけでメモ帳を探せる**。デスクトップを UIA で歩くと Excel 自身の provider で
  デッドロックする (凍結版 modRdv2Uia の実測) ため、走査ではなく `GetFocusedElement` から
  親を数段辿る形にした。デスクトップ列挙を一度もしないのでデッドロックの入口が無い。
  操作者の選び方は従来どおり「そのウィンドウで作業している方が勝つ」。
  **ただし一度もフォーカスを持っていないメモ帳は掴まない**のが唯一の挙動差で、
  ステータス行に「未接続 -- メモ帳の入力欄をクリックすると接続します」と出す。

## 設定 (C# 版・`ReaderDataViewer.json`)

配布は `.cmd` と `.json` の 2 枚。**サイトごとに変わるもの**を設定ファイルへ出してある:
監視対象 (ウィンドウと欄を UI Automation の性質で指定、**複数同時**)、番号の桁数と字種、
監視間隔と確定待ち、候補の表示行数、各ジョブの上限、データ・台帳・ログの場所。
画面の寸法や配色 (UI 正本)、マージの手順、保存の契約は**設定ではない**ので出していない。

- 監視は `window -> path[...] -> field` の経路で指定する。業務アプリの欄はたいてい数階層
  下にあるため、`automationId` を持つ中間段を並べて絞り込める。値の取り方は
  `value` (ValuePattern) / `text` (TextPattern) / `name` (Name) から選ぶ。
- 右上の「設定」ボタンで GUI 編集し、そのまま JSON に書き戻す。同じダイアログの
  「画面を調べる」は生きている UIA ツリーを表示し、選んだ要素から
  ウィンドウ・中間パス・欄・取得方法を**自動で組み立てる**。
- 無い/壊れている場合も起動する。壊れた値は**その値だけ**既定に落ち、理由がログに出る。
- 詳細は [settings.md](settings.md)。

## 監視は UI Automation だけ

対象アプリの探索・要素の特定・値の取得はすべて UIA。ウィンドウハンドル探索
(`FindWindowEx`)、フォアグラウンド判定 (`GetForegroundWindow`)、メッセージフック、
キーボード／マウスのポーリングは**使っていない** (前面判定は
`AutomationElement.FocusedElement` から親を辿る)。Win32 が残るのは**自分の**ウィンドウの
DPI 設定だけ。

### 起動用コンソール — 入口を `.vbs` にして解決した

コンソールの有無は設定ではなく、**Windows が起動する実行体のサブシステム**で決まる。
この端末の PE ヘッダを実際に読んだ値:

| 実行体 | サブシステム |
|---|---|
| `cmd.exe` / `powershell.exe` / `cscript.exe` | コンソール |
| **`wscript.exe`** / `mshta.exe` / `explorer.exe` | **GUI (コンソールを持たない)** |

`.cmd` は `cmd.exe` に渡るので、**こちらのコードが動く前に**窓ができる。縮められるだけで
ゼロにはできない (実測: 出したまま = ずっと / PowerShell 経由で親即終了 = 224→535 ms /
wscript 経由 = 197→334 ms)。しかも**一瞬の出入りはフォーカスを奪い、入力中のテキストを
飛ばす**ので、点滅させるくらいなら出したままの方がまし、というのがオーナー判断だった。

そこで**入口そのものを `.vbs` にした**。`wscript.exe` は GUI 実行体なので窓が存在しない。

```
ReaderDataViewer.vbs    入口 (これを起動する)。中身は .cmd と同じ payload
ReaderDataViewer.json   設定
ReaderDataViewer.cmd    同梱。コンソールを見たいとき用で、無くても動く
```

- `.vbs` は `WScript.ScriptFullName` を `RDV_SELF` に置き、PowerShell を**非表示**で起動する
  (`WScript.Shell.Run(..., 0, False)` + `-WindowStyle Hidden`)。あとは今までと同じで、
  PowerShell が自分自身を UTF-8 で読み、マーカー以降の C# を in-box の csc でコンパイルする。
- 環境変数はこのプロセスにだけ渡す一時的なもので、レジストリにもプロファイルにも書かない。
  **利用者が設定するものは何もない** (`.cmd` 版も以前から同じ仕組み)。
- VBScript は実行前にファイル全体を構文解析するので、payload は 1 行ずつ `'` でコメント化して
  埋め込み、ブートストラップ側で外している。
- 検証 (アプリを起動せずに実施): 先頭部が ASCII / payload 全行がコメント化されている /
  外すと元の 2 ソースと**完全一致** / VBScript が構文解析できる (cscript で構文だけ実行) /
  PowerShell 側の取り出しが動き、C# は `.cmd` 版と同一。

## UI (VBA 版はシート描画のみ)

- UserForm / ActiveX / フォームコントロール / ListBox / Shape ボタンは不使用。
  枠線非表示、セルの塗り・罫線・結合だけのフラットな画面。
- セルボタン (検索 / 内容クリア / 処理済み / 監視停止・再開) は **Hyperlink** (自セルへの
  SubAddress) で実装。hover 説明は Hyperlink の **ScreenTip**、クリックは
  `Worksheet_FollowHyperlink` で dispatch。Win32 mouse polling / hook は使わない。
- 検索入力セルはデータ検証の入力メッセージで案内する装飾セル。
- 複数候補はシート内の候補行 (8 識別列 = v2 と同じ列) + 行頭の「選択」Hyperlink セルで
  選ぶ。候補スロットは 10 行 (このデータの最大は 5 件)。溢れたら「候補 n 件中 10 件表示」を
  明示する。
- hover でモーダルは出さない。「処理済み」だけ永続変更前に MsgBox で確認する。
  確定までの間は「処理済み: TRUE (保存中...)」と表示し、保存完了の確定通知で
  「処理済み: TRUE」になる (保存がまだなのに完了と見せない)。
- クリアは入力と結果/候補表示だけを消す。台帳・processed・計時表示は消さない。

C# 版は WinForms のネイティブ画面で、**`Reader Data Viewer_ver3.html` を正本として
作り直してある** (寸法・色・書体・状態別スタイルまで `docs/ui-spec.md`)。1 枚のカードが
クライアント領域いっぱいに広がる 6 行構成 — タイトルバー / サマリー (キー1・代表ステータス・
台帳総件数・入力と 3 ボタン・担当者) / 候補一覧 / 統合レコード / エラー行 (条件表示) /
ステータスバー。候補はメイン画面内の一覧に出し、クリックで選択する (モーダルにはしない)。
意味論は VBA 版と同一。

- 座標は見本の実測表 (`Rdv3Geom.cs`) から引く。描画・コントロール配置・受け入れ検査の
  ダンプはすべて同じ辞書を読むので、測ったものと描いたものが必ず一致する。
- ウィンドウは設計サイズ (1240×689 論理) より小さくできない。内容がそれ以上を要求する
  場合 (長いユーザー名など) は最小幅がそのぶん上がる。拡大分は候補一覧と長文ボックスへ。
- 100/125/150/175% の拡大率で比率を保つ。プロセスは PowerShell ホストの都合で
  DPI 非対応なので、スレッド DPI コンテキストを上げて自前で拡大する。
- 受け入れ検査は**ウィンドウを表示せずに**製品ソースを描画して行う (メイン画面 360 通り)。
  手順と結果は `docs/ui-spec.md` 7 章。
- 設定モーダルとピッカーは **`Reader Data Viewer_ver4.html`** を正本に、同じ手順で
  作り直してある (実測表 `Rdv3SetGeom.cs` / `docs/ui-ref-settings-geom.json`、
  検査 `build/test_settings_geometry.ps1` 45 通り)。設計の根拠は `docs/ui-spec.md` 12–13 章。
  モーダルは 1060×764 で詰まっているため伸縮させず、作業領域に収まらないときは
  **全体の縮尺を下げて**収める。

## 計時境界 (両版で同一の定義)

UI に出す性能値は 2 つ (マージ時間 / 検索時間)。ログとベンチはそれより多くの区間を持つ。

```
マージ時間     = 表A読込 + 表B読込 + 表C読込 + 索引A + 索引B + 索引C + A-B結合 + B-C結合
                 (8 工程の合計。compose = 29 列台帳行の組立は含まれず、
                  compose_ms として常にログへ併記し、ベンチは merge_total = マージ+compose も出す)
検索時間       = 検索確定 (ボタン t0 / BE 検知 t0) → 単一表示 or 候補一覧の描画完了
                 (人が候補を選ぶ時間は含まない。検知そのものの待ち = 打鍵安定→確定は
                  detect latency_ms として別掲)
startup        = boot (アプリ入口) → 操作可能 (READY)。ログの startup 行
full E2E       = プロセス起動 (Excel / cmd) → READY。ベンチ側の時計で計測
                 (VBA は COM 経由起動。素のダブルクリックは未署名マクロの信頼設定が別途要る)
processed E2E  = 確認クリック → 台帳保存完了 + 画面確定 (e2e_ms)
承認 → 操作可能 = 更新確認の「はい」クリック → READY (ベンチ計測)
```

境界の**定義**は同一だが、実装経路が違うため中身は異なる: C# の検索はプロセス内 worker →
`Invoke` 描画 (数 ms)。VBA の検索は request ファイル → BE 検索 → publish → **FE pump が
拾って描画**で、pump 待ちが数百 ms 含まれる。これはアーキテクチャの実コストであり、
そのまま実測として報告する。

## 実行ログ

各配布ルート直下 `ReaderDataViewer.log` にタブ区切りで追記 (VBA 版はさらに BE 診断ログ
`ReaderDataViewer.log.worker.log`)。UI には細目を出さない。

```
時刻  runID  区間  詳細
      R#     boot / worker / read / index / join / merge(compose_ms 付) / verify /
             load(src=sidecar|book 付) / compare / decision / carry / persist /
             startup / pump / watch / timeout / error
      S#/P#  search / candidate / display / processed(save_ms, e2e_ms 付) / clear
      (BE)   detect latency_ms / serve_ms / mark / ledger written / self-quit
```

各行に件数・結果・ms・例外を含む。merge 検算 (行数 + join checksum vs `expected.txt`、
存在する場合のみ) は既存 4 方式と同じ検証をログと NG 時のエラー表示で維持している。

## ビルド

```
build.bat                                                           # dist\ 全部 (ダブルクリック可)
```

```powershell
powershell -ExecutionPolicy Bypass -File build\build_dist.ps1       # build.bat の中身
powershell -ExecutionPolicy Bypass -File build\build_app.ps1        # 実用版の両配布ルートだけ
powershell -ExecutionPolicy Bypass -File build\bench_app.ps1 -Method vba    -Launches 7
powershell -ExecutionPolicy Bypass -File build\bench_app.ps1 -Method csharp -Launches 7
```

- `build.bat` は**管理者権限不要・レジストリ書込なし・実行ポリシー変更なし**で、比較 2 種
  (8 ファイル) と実用版 2 ルートを既存の C# / VBA ソースから生成する。Excel と、ユーザー単位の
  設定「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」(HKCU、管理者不要) が要る —
  `.xlsm` に VBA を入れる唯一の経路が `VBProject.VBComponents.Import` だから。設定が無い場合は
  **黙って減らさず**、有効化手順を出して停止する。
- `data-100k\` は無いときだけ生成する。`data\` (100 万行 × 3 表、241MB) は**生成しない** —
  そこから作る配布物は無く、1:1 比較版が実行時に読むデータなので、必要なときに
  `build\gen_data.ps1` を明示的に走らせる (build.bat の最後にもそう出る)。
- 連続ビルドで一度だけ、META へ base64 を書く行 (`Cells.Item(r,5).Value2 = ...`) が
  Excel から「十分なメモリがありません」を返して失敗した (物理メモリは 6GB 空き、
  他の Excel が 5 プロセス生きている状態)。**再実行で成功**する類の automation ゆらぎで、
  失敗は行番号つきで出て exit 1 になる (黙って減った成果物は残らない)。

- `pack_app.ps1` が `src\app\csharp` + `src\app\cmd` → `ReaderDataViewer.cmd` (既存 packer と
  同じ規則: using 統合、非 ASCII の \uXXXX 化、C# 5、verbatim 禁止)。既存 `pack_cmd.ps1` は
  変更していない。
- `build_workbook_app.ps1` が `src\app\vba` → 小さな FE xlsm + 台帳ブック + sidecar。
  CP932 検査・CRLF 正規化・宣言順検査は既存 build_workbooks2 と同じ。worker book
  (Spec/Chan/Uia/Engine/Be + UIAutomationClient 参照) を組み立てて base64 で META へ埋め込み、
  **初期台帳ブック + sidecar はその worker book の `Rdv3BeBuildInitial` (= 製品の apply と
  同一の書込コード) で生成**する。FE/worker 双方のコンパイル検査を通す。
- `build_app.ps1` が C# 側の初期 `ReaderDataViewer-Ledger.xlsx` を app と同じ書込コードで
  生成し、CSV を各配布ルートへ配置する。
- ベンチ v2 は dist を scratch (`work\bench-run\`) へコピーして走り、**全 run (cold・失敗
  含む) を TSV に記録**して min/median/max/母数を出す。処理済みクリック・更新承認は
  実マウス / 実ダイアログ応答で行う。

## 実測で判明した落とし穴 (実用版)

1. **worker → FE の同期 COM は編集モードで恒久フリーズする。** 本 repo では踏む前に
   別プロジェクトで取った実測に従って回避した。「BE が FE のシートへ直接書く」構成は
   採用していない。
2. **BE で大きいブックを開いた後は同じ VBA が数倍遅くなる。** v1 で BE が 24 MB の FE ブックを
   ReadOnly で開いて閉じた後にマージを走らせると、同一コードが ~5 倍遅く計測された
   (7.8 秒 vs 1.4 秒)。ヒープ汚染によるもので、マージを先に走らせて計測を保護した。
   v2 は差分なし起動でブックを開かないため、この汚染経路自体が消えた。
3. **due 直前の `Application.OnTime` はキャンセルできない** (成功と報告しながら発火し得る)。
   閉じたブックの tick が発火すると Excel がブックを**再オープン**し、セキュリティ通知
   モーダルの裏でプロセスが永久にゾンビ化する (実測)。対策はキャンセル全廃: 閉じるときは
   armed なら `Cancel = True` で 1 周期 defer し、生きている tick 自身が
   `Rdv3FinishClose` で閉じ直す決定論方式 (modRdv3App)。
4. **`ADODB.Stream.ReadText` は大きいファイルで壊滅的に遅い。** 22 MB の sidecar 読取が
   **124.5 秒**かかった (BSTR の逐次伸長)。`MultiByteToWideChar` の一括変換に替えて
   **402.7 ms** (309 倍)。
   (この 402.7 ms は v1 当時のアプリ内実測。今回 Win32 を外す際に空プロセスで測り直すと
   同じ本文の読取は Get+MB2WC で 26.0 ms、UTF-16LE の Get+`Byte()`→`String` で 23.4 ms
   だった — 差はヒープ状態で、順位は変わらない。ADODB が桁違いに遅いのは一括読みでも
   同じで、今回も 22.2 秒だった。項目 17 も参照。)
5. **BE 常駐ループの cadence は反復回数でなく時間で刻む。** 1 反復 = 40ms + UIA poll は
   負荷で 200ms 超に伸び、「25 反復 ≒ 1 秒」の想定が崩れて stop flag の検知が FE の
   kill 猶予より遅れた (forced=True を実測)。`Timer` 基準へ変更して根治。
6. **参照ゼロの不可視 automation Excel は、メッセージポンプが回った瞬間に自動終了する。**
   spawn script が参照を放しても、VBA が Sleep ループで走り続ける間は生きている — が、
   `Workbooks.Open` (大きいブック) のようなポンプする操作が自動シャットダウンに入口を
   与え、ダイアログもクラッシュイベントも無しにウィンドウが畳まれてプロセスが消えた
   (実測: 台帳ブックの slow path で 18 秒後に死亡、ウィンドウ列挙で消滅過程を捕捉)。
   対策は bootstrap で `Application.UserControl = True` — インスタンスを自己所有と宣言し、
   自動終了を無効化する。以後の終了は自前の SelfQuit だけ。v1 で問題が出なかったのは
   FE が COM 参照を session の間ずっと保持していたから。
6. **モジュールレベル宣言はプロシージャより前**。builder の
   preflight が検査して止める。
7. **Excel を強制 kill し続けると対象ファイルが Resiliency\DisabledItems に登録され、
   以後そのパスは自動化から開けなくなる** (COM は総称エラーしか返さない。コピーは開けるのに
   元パスだけ失敗する)。テストハーネスは自パスのエントリだけを掃除してから起動する。
   製品側の対策は「強制 kill をしない」こと: stop flag → BE self-Quit → 猶予 → 最後の手段。
8. **WM_SETTEXT を SendMessageA 束縛で送ると UTF-16 文字列が先頭 1 文字に切れる**
   (テストハーネス側の罠)。SendMessageW を明示する。
9. **モーダルダイアログを開くボタンへ BM_CLICK を SendMessage すると呼び出し側が
   ダイアログ応答までブロックされる** (テストハーネス側の罠)。PostMessage にするか、
   別プロセスの watcher に答えさせる。
10. **合成 WM_LBUTTONDOWN/UP では WinForms の MouseClick は発火しない** (capture 状態機械)。
    実クリック検証は SendInput / mouse_event の実入力で行う。
11. **Win11 のメモ帳は全ウィンドウ共有 1 プロセス。** 自分のウィンドウを閉じるつもりの
    process kill は他人のウィンドウを道連れにする。閉じるのは WM_CLOSE をウィンドウ単位で。
    さらに: (a) **ファイル引数つきの起動は既存ウィンドウのタブに吸収され、新しいウィンドウが
    出ない** (素の notepad.exe は既存があっても新ウィンドウを開く)。(b) **dirty なタブは
    WM_CLOSE を無言で無視する** — 閉じる前にタブ内容をクリアしてから WM_CLOSE (数秒遅延あり)。
12. **実マウスクリックを一度でも受けた Excel は、ブックを全部閉じた後の COM `Quit()` を
    無視してプロセスが残る** (A/B 実測: クリック無しは Quit で終了、クリック有りは 17 秒
    待っても生存)。ハーネス側の対処は「ユーザーと同じ閉じ方」— メインウィンドウへ WM_CLOSE
    を post すればクリーンに終了する。kill は最後の手段のまま。
13. **`Window.PointsToScreenPixels` はペイン原点しか信用できない** (スケールは 1:1 を返す
    嘘)。セルのスクリーン座標 = `P2SP(0)` の原点 + ポイント値 × (`GetDpiForWindow`/72)。
    さらに呼び出しプロセスが DPI-unaware だと SetCursorPos 側の座標系まで仮想化されて
    二重にずれる — 必ず `SetProcessDPIAware` してから計算する (125% 表示のこの機で実測)。
14. **自動化 (不可視) Excel の VBA コンパイルエラーは「見えないモーダル」= 永久停止**。
    `Application.Run` は返らず、標準出力にも何も出ない。ウィンドウを列挙して初めて
    `#32770 "Microsoft Visual Basic for Applications"` が見つかる。ビルドの compile probe が
    これを CI 的に潰しているのはこのため。
15. **UIA の `IUIAutomationElement.CurrentNativeWindowHandle` は VBA でコンパイルできない**
    (戻り値 `UIA_HWND` = `void*` → 「サポートされていないオートメーション タイプ」)。
    14 と組み合わさると原因不明の永久停止になる。**`GetCurrentPropertyValue(30020)`**
    (戻り値 VARIANT) を使えば同じ hwnd が取れる。`GetRootElement` / `CompareElements` /
    `ControlViewWalker` / `CurrentClassName` / `CurrentName` は VBA でも問題ない。
16. **1 秒未満を待つ VBA の手段は「見つからない」ではなく「実測で選ぶ」**。
    `Application.Wait` は 1 秒粒度で秒境界に丸まり (+40ms 目標 = 即 return、+400ms 目標 =
    861ms)、`Application.OnTime` の 1 秒未満目標は約 1ms で発火する (スピン)。WMI
    イベントソースの `NextEvent(ms)` だけがタイムアウトまで本当に寝る (40ms 要求で 45〜54ms)。
17. **`ADODB.Stream.ReadText` は一括読みでも遅い** — 22MB で **22.2 秒** (chunk 読みの
    問題ではない)。また **mscorlib のクラスは VBA の `CreateObject` から作れない**
    (`System.Text.UTF8Encoding` / `System.Collections.ArrayList` とも 0x80131500。
    PowerShell では作れるので「動く」と誤認しやすい)。VBA で大きな UTF-8 を扱う手段は
    Win32 変換か、**ファイル自体を UTF-16LE にして `Byte()`⇄`String` で運ぶ**かの二択。

## 実測 (v2)

計測: `build\bench_app.ps1` (2026-08-14)。各 7 launch + 承認サイクル 1 回、**全 run 記録・
cold (launch 1) 込み・失敗 run も TSV に FAIL 行で記録** (最終走は両方式とも失敗 0)。
引き金は本物: 自分で開いたメモ帳ウィンドウへの WM_CHAR 打鍵 → アプリの UIA 検知、
処理済みは実マウスクリック (VBA はシートの G9 セル、C# は WinForms ボタン) + 実確認
ダイアログ応答、更新承認も実ダイアログ。両方式とも同じ CSV・同じ 4 キー
(候補 5/3/2/1)・同じ計測境界。結合 checksum **46629685 が全 run 一致** (両方式の実ログから
抽出)。生データ: `work\bench-app-csharp-20260814-103654.tsv` /
`work\bench-app-vba-20260814-110211.tsv` + 同名 console ログ。

すべて min / median / max (ms)、母数付き:

| 指標 | C# (n=7) | VBA v2 (n=7) |
|---|---|---|
| **完全 E2E** (プロセス起動→操作可能) | 3,741 / **3,884** / 5,273 | 8,015 / **8,662** / 9,413 |
| うち cold (launch 1、集計に含む) | 3,884 | 9,110 |
| boot→ready (アプリ入口→操作可能) | 2,622 / 2,670 / 3,934 | 7,091 / 7,682 / 8,573 |
| マージ (8 工程) | 94 / 116 / 130 | 1,267 / 1,295 / 1,481 |
| compose (別掲) | 249 / 275 / 277 | 1,632 / 1,835 / 2,671 |
| **マージ+compose 総計** | 349 / 387 / 395 | 2,903 / 3,130 / 4,033 |
| 保存状態ロード | 2,168 / 2,235 / 3,463 (xlsx 全読) | 366 / **395** / 500 (sidecar) |
| 検索 全キー (検知→表示、n=28) | 1.3 / 4.2 / 16.9 | 53 / 177 / 1,045 |
| 検知 latency (打鍵安定→確定、n=28、別掲) | 134 / 146 / 157 | 140 / 144 / 152 |
| **処理済み E2E** (確認→保存完了+画面確定) | 778 / **792** / 870 | 20,358 / **21,181** / 21,476 |
| うち永続化 (persist/save) | 777 / 790 / 868 | 8,512 / 9,000 / 9,074 |
| **承認→操作可能** (n=1) | **1,258** | **39,633** |
| 承認サイクル完全 E2E (n=1) | 5,857 | 47,073 |

読み方:

- **VBA の完全 E2E 中央値 8.7 秒 (cold 9.1 秒)** — v1 の boot→ready 実測 29〜33 秒から
  約 1/4。支配項だった台帳ブック全読み (16.8 秒) が sidecar 照合 0.4 秒になり、FE ブックが
  24 MB → 238 KB になったぶん Open も速い。BE が毎回新品プロセスなのでマージも安定
  (1.27〜1.48 秒 — v1 の 2.9〜7.9 秒の振れが消えた。ヒープ汚染経路の消滅)。
- **C# の完全 E2E 3.9 秒**は cmd の csc コンパイル + WinForms 起動 + xlsx 台帳全読 2.2 秒
  込み。「~3 秒」といった丸めではなくプロセス起動からの実測。
- **検索**: C# 4.2 ms はプロセス内、VBA 177 ms はファイルチャネル + pump 経路込み (定義同一)。
  VBA の max 1,045 ms は各 launch の初回検索 (BE 側の初回 serve のウォームアップ)。
  検知 latency (~145 ms) は両方式ほぼ同一 — 同じ検知規則 (打鍵後の安定待ち) の実測で、
  検索時間とは別掲。
- **処理済み (VBA) 21.2 秒**: BE が台帳ブックを開く (初回 ~11 秒) + セル書換 + Save ~9 秒 +
  sidecar 書換。ベンチは launch ごとに新しい BE なので毎回「初回 mark」= 最悪値に相当する
  (同一セッション 2 回目以降はブックが開いたままなので Save ~9 秒 + sidecar のみ。
  suite V4b で確認、ベンチでは未計測)。**この間 FE は自由** (占有はゼロ、表示は
  「TRUE (保存中...)」→ 確定通知で「TRUE」)。
- **承認→操作可能 (VBA) 39.6 秒**: BE が carry + 台帳ブック全面書換 + Save + sidecar を
  すべて自分で行う実コスト。v1 (FE 占有 25 秒) より壁時計は長いが、**FE 占有はゼロ**
  (v1 は materialize 13 tick + Save 6.2 秒で FE がブロックしていた)。
- **FE の連続占有 (v2)**: 起動 spawn の同期占有は廃止 (Shell 即 return)。残る占有は
  pump tick のみ — idle max ~10 ms、CHECK 消費 ~20 ms、描画 tick 数十 ms
  (suite V11a: BE マージ中の FE probe n=32 max 26.6 ms / p95 6.0 ms)。materialize /
  Save / 台帳シート書込は FE に存在しない。F2 実編集 20 秒中に BE が CHECK を公開し、
  FE の消費が編集終了直後の tick になることをタイムスタンプで再確認 (repro-v11c 5/5)。

### Win32 / Shell 除去の前後 (VBA 版、同じ機・同じ日・同じ手順で連続測定)

`build\bench_app.ps1 -Method vba -Launches 7` を除去前 → 除去後の順に走らせた実測。
結合 checksum は全 run **46629685** で一致 (= 出力は同一)。除去後は 2 回測っている
(中間ビルドと、`build.bat` が出した最終成果物) — この機の run 間ばらつきが起動系の
差より大きいので両方載せる。生データ:
除去前 `work\bench-app-vba-20260814-160934.tsv` / `work\baseline-before-vba.txt`、
除去後 `work\bench-app-vba-20260814-163423.tsv` / `work\after-vba.txt`、
最終 `work\bench-app-vba-20260814-164716.tsv` / `work\after-vba-final.txt`。

| 指標 (中央値、n=7) | 除去前 | 除去後 (最終) | 除去後 (中間) |
|---|---:|---:|---:|
| **完全 E2E** (プロセス起動→操作可能) | 9,036.5 | 9,921.2 | 9,049.8 |
| うち cold (launch 1、集計に含む) | 12,511.8 | 10,545.5 | 8,654.2 |
| boot→ready | 8,160.6 | 9,054.7 | 8,085.9 |
| **マージ (8 工程)** | 1,754.3 | **1,347.7** | **1,316.4** |
| **compose (別掲)** | 2,316.9 | 2,242.2 | **1,765.6** |
| **マージ+compose 総計** | 4,113.4 | **3,621.1** | **3,125.0** |
| **保存状態ロード (sidecar)** | **386.3** | 589.8 | 589.8 |
| 検索 全キー | 179.8 (n=28) | 156.3 (n=27) | 140.6 (n=28) |
| 検知 latency | 150.1 (n=28) | 144.5 (n=27) | 140.6 (n=28) |
| 処理済み永続化 | 8,995.2 | 9,070.3 | 9,121.1 |
| 処理済み E2E | 21,166.6 | 21,183.6 | 20,703.1 |
| 承認→操作可能 (n=1) | 39,676.5 | 41,886.3 | 38,750.4 |
| 承認サイクル完全 E2E (n=1) | 47,188.8 | 50,386.6 | 46,514.4 |
| FE 占有 spawn_ms | (Shell で非同期) | 945〜1,047 (n=8) | 1,000〜1,109 (n=8) |

読み方:

- **極端な低下は無い**。繰り返し出る差は 3 つだけで、うち 2 つは改善:
  マージ -23%、マージ+compose -12〜-24%、そして**保存状態ロードが +203 ms** (唯一の悪化)。
  起動系 (完全 E2E / boot→ready / 承認→操作可能) は除去後の 2 回が 9,050 と 9,921 ms で
  割れており、**この機の run 間ばらつきのほうが変更の効果より大きい** — どちらとも言えない、
  というのが正直な読み。
- **ロードが遅くなった理由は符号化ではなく sidecar が 2 倍 (22.3MB → 44.6MB) になったこと**。
  同一条件で読み比べると UTF-16 23.4ms 対 UTF-8 26.0ms で差は無い。効いているのは
  マージ+compose でヒープが埋まった直後に 22MB 余分を first-touch する分で、単体計測
  (空のプロセス) では読み+split+parse 合計 80 ms のところがアプリ内では約 590 ms になる。
- **マージの改善は今回の変更の狙いではない**。エンジンのコードは 1 行も変えていない。
  除去前は BE のマージ中に spawn 用 cscript が pid フラグを 200 ms ごとにポーリングしながら
  生きていた (それが消えた) こと、および測定日の機械状態が効いていると見ている。除去前の
  merge には 4,853.9 ms の外れ値が 1 件あり中央値を押し上げている (min 同士では
  1,488.7 → 1,300.8 ms)。時計は QPC から Timer に変わったが、どちらも実時間で系統誤差は
  無く、量子化は 3.906 ms なのでこの桁の差の説明にはならない。
- **検知 latency は 140.6 / 144.5 ms に量子化されて見える**。Timer の刻みで値が丸まって
  いるだけで、実測範囲は除去前 (132.2〜155.4 ms) と同じ帯にある。
- FE 占有は spawn の同期化で増えた: 起動時に **spawn_ms 945〜1,109 ms** (実行ログの
  `worker spawn ok ... spawn_ms=`) が最初の pump tick に乗る。ただし壁時計は上表のとおり
  変わっていない — 除去前も同じ時間を cscript が使っていて、FE が待っていた。

### 取りこぼしていた置換の修正: publish された RESULT がまれに描画されない

**症状** — 3 回の 7 launch で BE が serve した検索は各 39 件。FE が描画したのは
**Win32 除去前 37 / 除去後 (中間) 39 / 除去後 (最終) 37** で、**除去前から同じ率で 2 件
落ちていた** (Win32/Shell 除去で入ったものではない)。最終 run ではその 1 件がベンチの
検査対象キーに当たり、`search_00021001` が FAIL 1 件として TSV に記録されている
(ベンチは全 run 記録なので消していない)。落ちた側の BE ログには `search ... serve_ms=` が
あり、FE ログには対応する `search` 行が無く、`publish skipped` も `result skipped` も
出ていなかった — つまり**どこにも報告されずに消えていた**。

**機序** — 旧 `modRdv3Chan.AtomicReplace` は `Kill dest` → `Name tmp As dest` の順だった。
`Kill` が成功した直後に `Name` が失敗すると、その時点で aggregate ファイルが消え、
**まだ FE が読んでいない他種別のレコードごと道連れ**になる。しかも失敗経路は tmp まで
`Kill` していたので、**新しいレコードも同時に失われた**。関数は False を返すだけで、
呼び出し側は「今の 1 件」を書き直すため、消えた RESULT は二度と現れない。
`Kill` が失敗する側 (FE が読んでいる最中) は retry で救われるが、この順の失敗は救えない。

**修正** — `Rdv3ChReplaceFile` に置き換え、**既存ファイルを消してから改名する経路を廃止**した:

```
path -> path.bak   live を「消さずに退避」(失敗しても何も動いていない)
tmp  -> path       失敗したら bak を即座に戻す
Kill path.bak      新しい版が live になってから初めて捨てる
```

どの失敗でも **live と tmp の両方がディスクに残り** False を返すので、呼び出し側はそのまま
再試行できる。2 つの Name の間でプロセスが落ちた場合は live が `.bak` に退避されたまま
残るが、**次に書く側が復元する** (`HealParked`。aggregate は他種別を引き継ぐために publish の
読み込み前にも復元する)。失敗は握りつぶさず `Rdv3ChLastReplaceErr()` に残り、BE が 6 回の
retry で諦めたときの `publish skipped kind=... last replace err=` に出る。
sidecar の書換 (`modRdv3Be.WriteStateFileAtomic`) も同じ関数を通すようにした。

**確認** — 実モジュール (`modRdv3Spec` + `modRdv3Chan`) を読み込んだ VBA で 24 チェック全通過:
成功経路 (live 置換・tmp 消費・`.bak` を残さない) / live が無い初回書込 /
**置換失敗時**(live を deny-all で開いた状態: False を返す・エラー番号 55 を報告・live の内容は
`OLD` のまま・tmp は `NEW` のまま・孤児 `.bak` なし) / 障害解消後の再試行で成功 /
クラッシュ窓 (live 不在・`.bak` あり) からの復元 / **aggregate 単位** (STATE 公開 → ロック下の
RESULT publish が False → その後も aggregate に STATE が残っている → 再試行で STATE と RESULT の
両方が揃う) / 退避されたままの aggregate を次の publish が復元して 3 種別とも揃う。

### 実テストの通過状況 (v2、scratch コピー上、自起動プロセスのみ使用)### 実テストの通過状況 (v2、scratch コピー上、自起動プロセスのみ使用)

- C#: **50 チェック全通過** (v2 計測コード込みで再実行) — 更新なし / 単一検索 (30 項目照合) /
  複数候補 + 実マウス選択 / 処理済み確認ダイアログ + 永続化 / クリア / 再起動後の保持 /
  CSV 更新 → 承認 (carried=1 reset=1) / 拒否 (ファイル無変更をハッシュで確認) / 台帳破損 →
  拒否で BLOCKED・承認で再構築 / 二重起動拒否 (Mutex) / 不正キー / 0 件表示。
- VBA: **64 チェック中 62 通過 + 差し替え済み検証 5/5 全通過**。新規シナリオ含む:
  V9 台帳 META 破損 → mtime ガード → 低速路 → ledgerbad → 拒否 BLOCKED / 承認再構築、
  **V9b sidecar 欠損 → src=book で起動 + sidecar 自己修復**、V4/V4b 処理済みの BE 永続化
  (save_ms / e2e_ms 付き確定、「保存中...」→確定表示)、V8 拒否で台帳ブック + sidecar の
  ハッシュ無変更。FAIL 2 件はいずれもハーネス起因を特定済み: V11b の COM probe 方式
  (v1 から既知の不安定) は決定論版 repro-v11c (F2 実編集 20 秒、BE 公開 = 編集中 / FE 消費 =
  編集終了直後をタイムスタンプ証明、打鍵 12 文字の実確定) で置換して全通過。V12 の残骸検出は
  デバッグ用スパイクが kill で残した自セッションのファイル (sid 一致で特定、掃除済み) で、
  suite 自体の全セッションはクリーン (worker exit forced=False)。

## v1 の記録 (退役: 台帳を親ブック内シートに保持していた版)

v1 の構成: FE xlsm (24 MB、UI + LEDGER 10 万行 + META) が台帳を保持し、BE が FE ブックを
ReadOnly で開いて保存状態を読み、承認時は BE が part ファイル (8,192 行 × 13) を publish、
FE pump が 1 tick 1 part でシートへ反映して `ThisWorkbook.Save` (~6.2 秒 FE 占有)。
処理済みも FE がセルを書いて次 tick で Save していた (mark ごとに ~6 秒級の FE 占有 —
これは監査 (#54/#55) で判明した報告漏れ)。

v1 実測 (2026-08-14、`work\archive-v1\` に生ログ保全):

| 項目 | C# | VBA (v1) |
|---|---:|---:|
| マージ中央値 (compose 除く) | 106.3 ms (6/7 launch、cold 208.6 ms は集計外だった) | 4,827.6 ms (7 launch、2,900.8〜5,227.5) |
| 検索中央値 (検知→表示) | 4.8 ms | 217.3 ms |
| 起動 (boot→ready) | 2.78〜3.77 秒 (+cmd compile 0.76〜0.87 秒) | 実測保全分 29 / 32 / 33 秒 (32 秒 run の支配項 = 台帳ロード 16.8 秒) |
| 承認 → 操作可能 | 1.08 秒 | ~25 秒 (materialize 13 tick + Save 6.2 秒、原本ログ喪失) |
| FE 占有 | — | idle tick 9〜58 ms / materialize 248〜423 ms×13 / Save 6,218 ms |

v1 の監査 (Room #56〜#58) で判明した計測上の注意 — compose がマージ時間の外 (ログには
明示)、C# ベンチの cold 除外と checksum 欄未取得、VBA マージの実測 spread は
1.46〜7.85 秒 — は v2 のベンチ設計 (全 run 記録・compose/total 併記・checksum 実抽出) に
反映した。v1 のテスト通過状況: C# 50/50、VBA 54/56 + 差し替え 2 シナリオ再実行で全通過。
