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
  ReaderDataViewer-Ledger.state    台帳の sidecar ミラー (保存のたびに原子的に併記)
  data\tableA.csv / tableB.csv / tableC.csv

dist\app-csharp\
  ReaderDataViewer.cmd             自己完結 1 本 (PowerShell 起動 → in-box csc でコンパイル)
  ReaderDataViewer-Ledger.xlsx     統合台帳 (閲覧用。更新承認・処理済みでアプリが書き換える)
  data\tableA.csv / tableB.csv / tableC.csv
```

- 実行時の追加要件なし: 別 exe・常設 script・XLL・COM 登録・管理者権限は使わない。
  (VBA 版は起動の一瞬だけ `cscript.exe` (in-box) が worker Excel を立ち上げて退出する。)
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
ヘッダ 1 行 + 「processed TAB 内容 28 列」× 全行の UTF-8 TSV)。BE は保存のたびに
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
     → [FE] spawn script を Shell して即 return (同期 CreateObject はしない)
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
   ├ 「処理済み」の永続化 (セル+Save+sidecar)    ├ 起動 spawn は Shell(cscript) で非同期
   ├ key1 検索・候補抽出                         │   (FE 同期の CreateObject 占有 ~2 秒を廃止)
   └ 結果を atomic にファイルへ publish          └ FE ローカル SheetChange は
       (tmp 書き → Kill → Name)                     pump 再アームの watchdog のみ
  BE → FE の COM 呼出し: ゼロ / FE → BE の COM 参照: ゼロ (v2 で spawn からも消えた)
  FE → BE: 種別ごとの小さい request ファイル (latest-wins, version 付き) + stop flag
```

| モジュール | 役割 |
|---|---|
| `modRdv3Chan.bas` | ファイルチャネル: atomic publish、aggregate、request、FE リース、be_pid/spawn_err フラグ、session 掃除 |
| `modRdv3Be.bas` | BE 本体: 常駐ループ (時間基準の cadence)、マージ、状態ロード (sidecar/ブック)、比較、carry、台帳ブック書換+Save、sidecar、mark 永続化、検索、監視、publish |
| `modRdv3Host.bas` | FE 側の BE 所有: 埋込 worker book の展開、spawn script 生成 + Shell、stop (flag → self-Quit 確認 → 最後の手段の kill は自 PID のみ) |
| `modRdv3App.bas` | FE 本体: 短い boot、OnTime pump、dispatch、描画呼出し、ログ (台帳への書込・Save は存在しない) |
| `modRdv3Ui.bas` / `modRdv3Uia.bas` / `modRdv3Engine.bas` / `modRdv3Spec.bas` | 描画 / UIA 検知 / マージエンジン / 定数 (v2 由来) |
| `clsRdv3AppEvents.cls` | FE ローカルイベント → pump 再アーム watchdog (BE 通知には使わない) |

FE ブックのモジュールは Spec/Chan/Host/Ui/App + class の 5+1 のみ。Engine/Be/Uia は
worker book (META に base64 埋込、実行時に `%TEMP%\rdv3\` へ session 固有名で展開) だけが持つ。

要点:

- **非同期 spawn**: FE は spawn script (`rdv3_<sid>_spawn.vbs`) を生成して `Shell` で
  cscript に渡し、即 return する。script が不可視 Excel を作り worker book を開いて
  bootstrap を Run し、BE ループが自分の PID を `be_pid` フラグへ書くまで参照を保持してから
  退出する (参照ゼロの idle 非表示 Excel は自己終了するため。ループが走り出せばもう落ちない)。
  失敗は `spawn_err` フラグで FE に届き、無ければ 30 秒でタイムアウト表示。FE は BE への
  COM 参照を一切持たないので、BE の self-Quit を妨げる要因も消えた。
- **FE 死活はリースファイル**: FE が開きっぱなし + deny-all lock で保持し、BE は lock 試行で
  判定する。FE がクラッシュすると OS が即 lock を解放し、BE は自己終了する。tombstone
  (`fe_gone`) と stop flag も併用。SessionId は timestamp 生成 (hWnd は不使用)。
- **承認後の台帳反映は BE が完結**: carry → 台帳ブックの全面書換 (16,384 行ブロックの
  Value2、テキスト書式で前ゼロ保持) → Save → sidecar。FE には統計値の APPLY 通知と READY
  だけが届く。v1 の part 転送 / FE materialize / FE Save は全廃した。
- **「処理済み」も BE が完結**: FE は要求を出し「処理済み: TRUE (保存中...)」を表示、
  BE がブックのセル書換 + Save + sidecar 後に RESULT (marked) で確定を返し、FE が
  「処理済み: TRUE」に確定させる。失敗は markerr で明示し表示も FALSE へ戻す。
  最初の mark は BE が台帳ブックを開くぶん遅い (以後はブックを開いたまま保持し、
  BE 終了時に閉じる。保存はそのたび行う)。
- **FE ブックは何も保存しない**: 台帳が外に出たので FE に永続状態がない。閉じるときは
  `Saved = True` で無音 (UI セルの書込で保存プロンプトを出さない)。
- **待機カーソル**: pump callback 実行中に一瞬 WAIT カーソルが出ることは許容する
  (カーソル抑止の作り込みは適用範囲外)。

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

C# 版は WinForms で同じ 3 部構成 (ステータス / 結果・候補 / 操作)。候補はメイン画面内の
一覧に出し、クリックで選択する (モーダルダイアログにはしない)。意味論は VBA 版と同一。

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

```powershell
powershell -ExecutionPolicy Bypass -File build\build_app.ps1        # 両配布ルートを生成
powershell -ExecutionPolicy Bypass -File build\bench_app.ps1 -Method vba    -Launches 7
powershell -ExecutionPolicy Bypass -File build\bench_app.ps1 -Method csharp -Launches 7
```

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
   **402.7 ms** (309 倍)。大きな UTF-8 テキストを VBA で読むときは必ず Win32 変換にする。
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

### 実テストの通過状況 (v2、scratch コピー上、自起動プロセスのみ使用)

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
