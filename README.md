# Reader Data Viewer

読み取り → 三 CSV 統合 → 検索 → 閲覧表示を、**同じ仕事を 3 通りの方式で実装して比べられる**
形で作ったアプリです。

メモ帳を実業務アプリの代役にして、そこへ入力・貼り付けされた番号を UI Automation の
ポーリングで検知します。検知するたびに、1,000,000 行 × 10 列の CSV 3 本を**毎回読み直して**
2 つの結合をやり直し、1 件を引いて画面に出します。キャッシュはしません。
工程ごとの時間が画面に出るので、どこで時間を使っているかがその場で分かります。

この repo には**比較が 2 つ**あります。どちらも配布物・ソース・計測が別で、片方をビルド
しても、もう片方の成果物は書き換わりません。

| | 対象 | データ | 配布物 | 記録 |
|---|---|---|---|---|
| **1 対 1 版** | VBA 単独 / C# 単独 / 併用 の **3 方式** | 100 万行 × 3 表、番号1 は一意 | 4 ファイル | [docs/results.md](docs/results.md) |
| **一対多版** | VBA・C# × 自前ハッシュ・標準連想配列 の **4 方式** | 10 万行 × 3 表、番号1 は正常に複数件 | 4 ファイル | [docs/results2.md](docs/results2.md) |

以下は 1 対 1 版の説明です。一対多版は [「一対多版 — 4 方式」](#一対多版--4-方式) に分けて
書いてあります。

計測結果は [docs/results.md](docs/results.md) にあります。3 行でまとめると:

- **方式差は 5.3 倍。** 併用 (C#+Excel) 0.59 秒、C# 単独 0.66 秒、VBA 単独 3.14 秒。
- **効くのは読み込みと結合だけ。** どの方式でも合計の 99% 以上がそこです。検索は 0.1 ms 未満、
  表示はいちばん重い方式でも 34 ms で、順位に寄与しません。
- **表示だけは順位が逆。** 同一プロセスのシート書き込み 1.9 ms < COM 越し 10.4 ms <
  WinForms 34.1 ms。30 項目でも経路の差は 18 倍あります。

## 配布物 — 4 ファイル (1 対 1 版)

| ファイル | 方式 | 中身 |
|---|---|---|
| `dist\ReaderDataViewer-VBA.xlsm` | **VBA 単独** | 監視・読込・結合・検索・表示・計測を 1 冊の中で完結 |
| `dist\ReaderDataViewer-CSharp.cmd` | **C# 単独** | 通常権限で Windows PowerShell を起動し、自分の中の C# をコンパイルして実行。WinForms で表示 |
| `dist\ReaderDataViewer-Hybrid.cmd` + `dist\ReaderDataViewer-Hybrid.xlsm` | **併用** | `.cmd` が監視・読込・結合・検索・計測、`.xlsm` が同じ内容の Excel 画面 |

`.cmd` はどちらも 1 ファイルで完結します。別の `.ps1` / `.cs` / `.exe` / `.dll` は要りません。
XLL、タスク スケジューラ、COM 登録、管理者権限も使いません。

**この 4 つは git に入っていません。** ビルド生成物なので、下の手順で作り直してください。

## 使う

```powershell
# 1. 合成データ (241 MB) を作り、4 つの配布物をビルドする
powershell -ExecutionPolicy Bypass -File build\build_all.ps1

# 2. メモ帳を開いておく (アプリが自動で見つけます)

# 3. どれか 1 つを起動する
dist\ReaderDataViewer-CSharp.cmd
dist\ReaderDataViewer-Hybrid.cmd
#   または dist\ReaderDataViewer-VBA.xlsm を開いて [監視開始]
```

あとはメモ帳に **8 桁の番号**を打つか貼り付けるだけです。番号 1 は `00000001` 〜 `01000000`。
先頭の 0 は意味のある文字として扱います。

同じ番号を続けて読むときは、**いったん欄を空にしてから**もう一度読ませてください
(空になった後の同一番号は許可、空にせずの連続重複は無視、という規則です)。

### 画面に出るもの

- 現在の番号、監視状態、つないでいるメモ帳のウィンドウ
- 統合レコード 1 件 (表A・表B・表C の 10 項目ずつ、計 30 項目)
- 工程別の時間と合計、検知遅延 (参考値)、データ行数、プローブ数
- 検算結果 (生成側が独立に計算した値との突合)
- エラー

### 引数

```
ReaderDataViewer-CSharp.cmd [データフォルダ] [-log <出力先.tsv>]
ReaderDataViewer-Hybrid.cmd [データフォルダ] [-log <出力先.tsv>]
```

データフォルダを省略すると `..\data` → `.\data` の順に探します。`-log` を付けたときだけ、
1 回 1 行の TSV を追記します (既定では何も書きません)。VBA 版は同じ設定をシートの
`C46` (ログ出力先) と `C47` (データフォルダ) で指定します。

## 何を 1 回として測っているか

**merge-select** = 番号が確定した瞬間から、画面にその 1 件が出るまで。

```
1 表A 読込・索引 → 2 表B 読込・索引 → 3 表C 読込・索引
  → 4 A-B 結合 (番号1) → 5 B-C 結合 (番号2) → 6 番号1 検索 → 7 画面表示
```

- 結合は**全 1,000,000 行**についてやります。「番号 1 の行だけ引く」近道はしません。
- 前回の統合結果も、読み込んだファイルも、索引も**持ち越しません**。
- ポーリングが番号を確定するまでの**検知遅延は主計測から分離**して、参考値として出します。
- C# のコンパイルと Excel の起動は**監視を始める前に**終わらせ、merge-select には混ぜません。

## 3 方式が同じ仕事をしている確認 (1 対 1 版)

方式差を「方式の差」と呼ぶには、中身が同じでなければ意味がありません。毎回の実行が
次を報告し、24 回すべてで一致しています。

| 検証項目 | 値 |
|---|---|
| 結合チェックサム | `658725490` (生成側が独立に計算した `data/expected.txt` と一致) |
| プローブ数 | `3,639,849` (索引構築時の線形探査の総回数) |
| 突合数 | A-B `1,000,000` / B-C `1,000,000` |

プローブ数まで一致するということは、VBA と C# が**同じハッシュ表を同じ順序で作り、
同じ場所で衝突している**ということです。実際、両者は同じ手順を実装しています
(生バイトのまま保持 / 行オフセット走査 / 同一の多項式ハッシュ / オープンアドレス法)。

## データ (1 対 1 版)

`build/gen_data.ps1` が決定的に生成します。同じスクリプトなら毎回同じバイトです。

| 表 | 行数 | 列 | キー |
|---|---|---|---|
| A | 1,000,000 | 10 | `key1` (8 桁、先頭 0 あり) |
| B | 1,000,000 | 10 | `key1` と `key2` の両方を持つ |
| C | 1,000,000 | 10 | `key2` |

A は `key1` 順、C は `key2` 順、B は種を固定したシャッフル順で書き出します。
`key2` は `key1` の固定置換なので、B→C は A→B と同じ並びになりません。
どの実装も、ファイルを順に舐めるだけでは正解に辿り着けません。

## 動かすのに要るもの

- Windows 10 / 11 (64bit)
- Excel デスクトップ版 (VBA が使えるもの、**64bit**) — VBA 単独方式と併用方式で必要
- .NET Framework 4.x の `csc.exe` (Windows 同梱) — `.cmd` のコンパイルに使用
- Windows PowerShell 5.1 (64bit)
- 空きメモリ 1 GB 程度

**ビルド時だけ** Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が
必要です (ファイル → オプション → トラスト センター → マクロの設定)。
出来上がった `.xlsm` を動かすのには要りません。

## 既知の制約

- **32bit Excel / 32bit PowerShell では動きません。** 100 万行 × 3 表を保持できません。
  起動時に検査して、その旨を出して止まります。
- **データは ASCII 前提**です。VBA 側は `StrConv(bytes, vbUnicode)` で行走査用の文字列を
  作るため、CP932 に無い文字が入ると壊れます。合成データは ASCII だけで作ってあります。
- **キーは 8 バイト固定**です。長さが違う行を見つけたらエラーとして報告し、素通りしません。
- **メモ帳のウィンドウは選べますが、閉じたり起動したりはしません。** 複数開いている場合は
  最前面のものを使い、無ければ最後に見つかったものを使います。画面に接続先のタイトルと
  ウィンドウ ハンドルが出るので、[メモ帳を再検出] で切り替えられます。
- **Windows 11 の新しいメモ帳と従来のメモ帳の両方**に対応します (Document / Edit のうち
  ValuePattern を持つ要素を探します)。
- **併用方式は Excel を 1 つ占有します。** 自分で起動した 1 インスタンスにだけ触り、
  すでに開いている他の Excel には触りません。終了時に閉じるのも自分が起動した 1 つだけです。
- **アプリは CSV にも Excel ファイルにも書き込みません。** ブックは読み取り専用で開きます。
  画面表示のためのセル書き込みだけを行い、保存はできません。
- PowerShell が制限言語モード / AppLocker 下にある環境では `Add-Type` が使えないため、
  `.cmd` の 2 方式は動きません。

## 作り直す

**`build.bat` をダブルクリックすれば `dist\` の配布物が全部 (比較 2 種 8 ファイル +
実用版 2 ルート) ソースから再生成されます。** 管理者権限は不要・レジストリは書きません・
実行ポリシーも変えません。Excel と、ユーザー単位の設定「VBA プロジェクト オブジェクト
モデルへのアクセスを信頼する」(トラスト センター、管理者不要) が要ります — `.xlsm` に
VBA を入れる唯一の経路が `VBProject.VBComponents.Import` だからです。設定が無いときは
**黙って減らさず**、有効化手順を出して止まります。

個別に走らせるとき:

```powershell
powershell -ExecutionPolicy Bypass -File build\build_dist.ps1         # build.bat の中身
powershell -ExecutionPolicy Bypass -File build\build_all.ps1          # 1 対 1 版 4 ファイル
powershell -ExecutionPolicy Bypass -File build\gen_data.ps1 -Force    # データだけ
powershell -ExecutionPolicy Bypass -File build\pack_cmd.ps1 -Variant csharp
powershell -ExecutionPolicy Bypass -File build\pack_cmd.ps1 -Variant hybrid
powershell -ExecutionPolicy Bypass -File build\build_workbooks.ps1    # ブック 2 冊
```

計測:

```powershell
powershell -ExecutionPolicy Bypass -File build\run_bench.ps1 -Method csharp -Repeat 7
powershell -ExecutionPolicy Bypass -File build\run_bench.ps1 -Method hybrid -Repeat 7
powershell -ExecutionPolicy Bypass -File build\run_bench.ps1 -Method vba    -Repeat 7
powershell -ExecutionPolicy Bypass -File build\summarize.ps1 -Log work\bench-vba-<日時>.tsv
```

`run_bench.ps1` はメモ帳へ `WM_CHAR` を 1 文字ずつ送ります。**引き金は本物**で、
エンジンを直接呼ぶ経路では測っていません。自分が開いたウィンドウだけを閉じ、
すでに開いていたメモ帳や Excel には触りません。

## 一対多版 — 4 方式

番号1 が**正常に複数件ヒットする**ときの版です。1 対 1 版とは配布物もソースも計測も別で、
`build\build_all2.ps1` は 1 対 1 版の 4 ファイルに触れません。

比べているのは**索引の作り方**です。言語 2 つ × 索引 2 つ の 4 方式:

| ファイル | 方式 | 索引 |
|---|---|---|
| `dist\ReaderDataViewer-VBA-Hash.xlsm` | VBA 自前ハッシュ | バケット連鎖 (自作) |
| `dist\ReaderDataViewer-VBA-Dict.xlsm` | VBA 標準 Dictionary | `Scripting.Dictionary` (遅延バインド) の値に `Collection` |
| `dist\ReaderDataViewer-CSharp-Hash.cmd` | C# 自前ハッシュ | バケット連鎖 (自作) |
| `dist\ReaderDataViewer-CSharp-Dict.cmd` | C# 標準 Dictionary | `Dictionary<string, List<int>>` |

同じ言語の 2 つは、**索引のファイル 1 本だけ**が違います。呼び出し側は同じで、実行時の
分岐もありません (C# は同名クラスを差し替え、VBA は同名クラス モジュールを差し替え)。

### データの形

表B は伝票の明細行です。1 つの番号1 が 1・2・3・5 行を持ちます。

| 表 | 行数 | キー | 重複 |
|---|---|---|---|
| A | 100,000 | `key1` | 一意 |
| B | 100,000 | `key1` と `key2` | **`key1` は重複する** / `key2` は一意 |
| C | 100,000 | `key2` | 一意 |

候補数はキーから決まります: `00000001`〜`00001000` が 5 件、`00001001`〜`00006000` が 3 件、
`00006001`〜`00021000` が 2 件、`00021001`〜`00071000` が 1 件。

### 複数件ヒットしたとき

候補ダイアログが開き、1 件選ぶと、その候補の統合レコードを通常画面に出します。
一覧に出す識別列は 10 フィールドから選んだ次の 8 列で、4 方式とも同じです:

| 列 | 出所 | なぜこれか |
|---|---|---|
| 番号2 | 表B `key2` | 候補を一意に決めるキー |
| 行番号 | 表B `b_line` | 同じ伝票の何行目か |
| 伝票番号 | 表B `b_slip` | どの伝票か |
| 日付 | 表B `b_date` | いつのものか |
| 数量 | 表B `b_qty` | 量で見分ける |
| 状態 | 表B `b_status` | OPEN / HOLD / DONE / VOID |
| 品目コード | 表C `c_item` | 番号2 で結合した先の品目 |
| メーカー | 表C `c_maker` | 同上 |

並びは表B のファイル順です。番号2 は一意という前提で扱います。

### 使う

```powershell
powershell -ExecutionPolicy Bypass -File build\build_all2.ps1   # データ + 4 配布物

dist\ReaderDataViewer-CSharp-Hash.cmd
dist\ReaderDataViewer-CSharp-Dict.cmd
#   または dist\ReaderDataViewer-VBA-Hash.xlsm / -Dict.xlsm を開いて [監視開始]
```

```
ReaderDataViewer-CSharp-Hash.cmd [データフォルダ] [-log <出力先.tsv>] [-autopick <n>]
```

データフォルダを省略すると `..\data-100k` → `.\data-100k` の順に探します。
`-autopick <n>` は候補ダイアログを自動で答える計測用の指定で、**計測が終わった後**に
効きます (人が押すのと同じ場所)。人が操作するときは付けないか `-1` にします。
VBA 版は同じ設定をシートの `C59` (ログ) `C60` (データフォルダ) `C61` (自動選択) で指定します。

### 何を 1 回として測っているか

```
1..3 表A/B/C 読込 → 4..6 表A/B/C 索引作成 → 7 A-B 結合 → 8 B-C 結合
  → 9 番号1 検索 (候補抽出) → 10 表示準備
       候補 1 件 : 統合レコードを画面へ
       複数件    : 候補ダイアログを組み立てて一覧を流し込むまで
── ここで計測終了。人が選ぶ時間は入りません。
選択後の表示は別項目として計測します。
```

読込と索引作成を**分けて**測ります。索引に何を使うかがこの比較の全部だからです。

### 結果 (各 8 回・中央値)

| | VBA 自前 | VBA Dict | C# 自前 | C# Dict |
|---|---:|---:|---:|---:|
| 合計 | 283.1 ms | 1,450.6 ms | 89.2 ms | 149.0 ms |
| うち索引まわり (作成 + 結合) | 181.6 ms | 1,323.0 ms | 22.1 ms | 64.9 ms |
| 選択後の表示 (計測外) | 1.3 ms | 1.2 ms | 41.5 ms | 44.3 ms |

4 方式とも目安の 5 秒以内。標準の連想配列に替えると索引まわりは VBA で 7.3 倍、C# で 2.9 倍
になります。全 32 回で結合チェックサム `46629685`・候補数 5/3/2/1 が一致しています。
内訳と、そうなる理由は [docs/results2.md](docs/results2.md) にあります。

### 一対多版の制約

- 番号2 は**一意である前提**で扱います。
- 索引は `key → 行番号の集合`です。`key → 単一行`の上書き辞書にはしていないので、
  同じ番号1 の行が失われることはありません。
- 候補ダイアログに出すのは先頭 256 件までです (このデータの最大は 5 件)。
- 1 対 1 版と同じく、CSV にも Excel ファイルにも書き込みません。

## 実用版 — 標準 2 方式

比較で選定した**標準の連想配列だけ**を使い、実運用の形に組んだアプリです。比較版と違い
統合結果を**台帳として永続化**し、起動時にバックグラウンドで更新を確認します。
設計の全体と実測は [docs/app.md](docs/app.md) にあります。

| 配布ルート | 中身 | 方式 |
|---|---|---|
| `dist\app-vba\` | 小さな FE `ReaderDataViewer.xlsm` (~240 KB) + 台帳ブック `ReaderDataViewer-Ledger.xlsx` + sidecar `.state` + `data\` (CSV 3 本) | VBA + 遅延バインド Scripting.Dictionary |
| `dist\app-csharp\` | `ReaderDataViewer.cmd` + 閲覧用 `ReaderDataViewer-Ledger.xlsx` + `data\` | C# + `Dictionary<string, List<int>>` |

- **C# 版の配布は 2 枚**: プログラム (`.cmd`) と設定 (`.json`)。監視するアプリのウィンドウと
  その中の欄 (UI Automation の AutomationId 等で指定・**複数可**)、番号の桁数、監視間隔、
  台帳やログの場所などは設定ファイル側にあります。画面右上の「設定」から GUI で編集でき、
  「画面を調べる」で対象アプリの UIA ツリーを見ながら欄を選べます →
  [docs/settings.md](docs/settings.md)。設定ファイルが無くても既定値で動きます。
- 統合台帳は A+B+C 統合の 10 万行 (統合レコード 1 件 = 1 行、identity は番号2)。
  `処理済み` 列を持ち、CSV 更新後も変更のないレコードへ引き継ぎます。
- 起動時に「更新を確認中」を表示しながら新しい統合結果を作り、**保存済み台帳の内容と比較**
  します (ファイル時刻やサイズで内容は判定しません)。差分があれば更新するか確認します。
- VBA 版 (v2) は**台帳を親ブックから分離**: FE は UI だけの小さなブックで、台帳ブックと
  その sidecar ミラーは**別プロセスの不可視 Excel (BE)** だけが読み書き・保存します。
  差分なしの起動は sidecar 照合だけで済み、台帳ブックを開きません。承認時の書換・保存も
  「処理済み」の永続化も BE が行い、FE は要求と表示だけ (FE の materialize / Save は v2 で
  全廃)。
- **VBA 版の実行コードは Win32 API (`Declare`) も `Shell` も使いません** — VBA と COM だけです。
  BE の起動は `CreateObject("Excel.Application")` + bootstrap (bootstrap は `OnTime` を
  仕掛けて即 return するので FE 占有は起動処理そのものだけ)、時計は `Timer`、1 秒未満の
  待機は WMI イベントの timeout 待ち、メモ帳の特定は UIA の `GetFocusedElement` から親を
  辿る方式、死活は双方向のリースファイル lock。理由と実測は `docs/app.md` の
  「Win32 / Shell を使わない実装」。
- 画面の性能値は**マージ時間と検索時間の 2 つだけ**。細目 (compose、台帳ロード、
  processed 保存、startup、検知 latency) はテキスト実行ログへ出ます。
- VBA 版の UI はシート描画だけ (UserForm / ActiveX / Shape 不使用)。ボタンは Hyperlink セル
  (hover 説明は ScreenTip)。
- 自前ハッシュ版は実用版には入れていません (比較証拠として `src/v2` に凍結)。
- **未確定の 1 件保存を黙って落とさない終了保護** (両版)。「処理済み」は 1 件ずつその場で
  永続化しますが、書込み中は結果がまだ決まっていません。その窓の中では (1) 新しい処理済み
  操作を受け付けず、(2) 終了要求をキャンセルして「保存中のため終了できません」と理由を出し、
  (3) 成功か失敗として**確定した後に**終了できます。溜めない・キューにしない・一括保存しない。
  実測は `build\test_exit_guard.ps1` (両版とも実物の保存中に終了要求を出して確認)。
  保存中でも**画面はいつもどおり操作でき**、検索もできます。保存完了の通知は検索応答とは別の
  チャネルスロットで配送され、答える要求の版番号を名乗るので、どちらかがもう一方を消すことは
  ありません (`build\bench_e2e.ps1 -Mode race` で 3 つの順序を実測)。
- **VBA 版は 1 件保存の方式を 3 つ持ちます** (現行 = 子 Excel で開いて `Workbook.Save` /
  閉じた台帳への ADO 更新 / **閉じた台帳の元シート XML の 1 バイトだけを純 VBA で書き換える**
  = deflate の復号・対象ブロックの再符号化・ビット接合・CRC-32 の 1 バイト差分補正)。
  **配布既定は現行方式のまま**で、切替は台帳の隣に `<台帳>.savemethod` を置いたときだけ
  (配布物にそのファイルはありません)。比較は 2 種類あり、どちらも
  [docs/save-methods.md](docs/save-methods.md) — **実配布物を実際にクリックして操作し、
  更新確認ジョブと処理済み登録ジョブを外から計測した実経路の測定** (3 方式 × 3 条件 × 10 回、
  `build\bench_e2e.ps1`) が正本で、保存関数だけを測ったマイクロベンチ (`build\bench_save.ps1`)
  は参考値です。毎回「そのフラグ 1 バイト以外は 1 バイトも変わっていない」ことを機械検査した
  結果、それを満たすのは 3 方式のうち 1 つだけでした。

実測 (v2、各 7 launch + 承認 1 回、cold 込み全 run 記録、失敗 0、中央値。正本は
`docs/app.md` の「実測 (v2)」— min/max/母数と読み方付き。v1 実測は同書の退役セクション):

| 中央値 | C# 標準 | VBA 標準 v2 |
|---|---:|---:|
| 完全 E2E (プロセス起動→操作可能) | 3.88 秒 | 8.66 秒 (v1: 約 29〜33 秒) |
| マージ+compose 総計 | 387 ms | 3,130 ms |
| 検索 (検知→表示) | 4.2 ms | 177 ms |
| 処理済み E2E (確認→保存完了) | 792 ms | 21.2 秒 (BE 内、FE 占有ゼロ) |
| 承認→操作可能 | 1.26 秒 | 39.6 秒 (BE 内、FE 占有ゼロ) |

```powershell
powershell -ExecutionPolicy Bypass -File build\build_app.ps1     # 両配布ルートを生成
dist\app-csharp\ReaderDataViewer.cmd                             # C# 版を起動
#   VBA 版は dist\app-vba\ReaderDataViewer.xlsm を開くだけ (マクロ有効化が必要)
```

## 構成

```
src\csharp\RdvCore.cs        エンジン (読込・索引・結合・検索・計測)   ← 2 つの .cmd が共有
src\csharp\RdvWatch.cs       UIA でメモ帳を見張る (検知規則もここ)
src\csharp\RdvRunner.cs      検知確定から画面表示までの順序と計測境界
src\csharp\RdvText.cs        画面に出る文字列 (非 ASCII はここだけ)
src\csharp\RdvUiForms.cs     WinForms 画面           ← C# 単独
src\csharp\RdvApp.cs         C# 単独の入口
src\csharp\RdvUiExcel.cs     Excel 画面 (COM 遅延バインド) ← 併用
src\csharp\RdvAppHybrid.cs   併用の入口とスレッド構成
src\cmd\header.cmd           .cmd の頭 (自分の中身を取り出して PowerShell へ渡す)
src\cmd\boot-common.ps1      参照解決とコンパイル
src\cmd\boot-csharp.ps1      / boot-hybrid.ps1   それぞれの起動

src\vba\modRdvSpec.bas       定数・時計・ハッシュ・画面レイアウトの住所
src\vba\modRdvEngine.bas     エンジン (C# 版と同じ手順)
src\vba\modRdvUia.bas        UIA クライアント (早期バインド必須)
src\vba\modRdvApp.bas        監視ループ・画面・計測境界・ログ

build\gen_data.ps1           合成 CSV 3 本 + 検算用 expected.txt
build\pack_cmd.ps1           src\csharp または src\v2\csharp → 1 個の .cmd
build\build_workbooks.ps1    src\vba → .xlsm 2 冊 (取り込み前検査つき)
build\build_all.ps1          上をまとめて実行
build\run_bench.ps1          方式を N 回まわして計測
build\summarize.ps1          ログから中央値・最小・最大
```

一対多版 (上の 1 対 1 版のソースには手を入れていません):

```
src\v2\csharp\Rdv2Core.cs      エンジン (読込・索引・結合・候補抽出・計測)
src\v2\csharp\Rdv2IdxHash.cs   索引 1: 自前バケット連鎖        ← どちらか 1 本だけ入る
src\v2\csharp\Rdv2IdxDict.cs   索引 2: Dictionary<string,List<int>>
src\v2\csharp\Rdv2Watch.cs     UIA でメモ帳を見張る
src\v2\csharp\Rdv2Text.cs      画面に出る文字列 (非 ASCII はここだけ)
src\v2\csharp\Rdv2Ui.cs        WinForms 画面 + 候補ダイアログ
src\v2\csharp\Rdv2App.cs       入口と計測境界
src\v2\cmd\boot-common2.ps1    参照解決とコンパイル
src\v2\cmd\boot-hash.ps1       / boot-dict.ps1

src\v2\vba\modRdv2Spec.bas     定数・時計・ハッシュ・画面レイアウトの住所
src\v2\vba\modRdv2Engine.bas   エンジン (C# 版と同じ手順)
src\v2\vba\clsRdv2IdxHash.cls  索引 1: 自前バケット連鎖        ← どちらか 1 本だけ入る
src\v2\vba\clsRdv2IdxDict.cls  索引 2: Scripting.Dictionary + Collection
src\v2\vba\modRdv2Uia.bas      UIA クライアント
src\v2\vba\modRdv2App.bas      監視ループ・画面・候補ダイアログ・計測境界・ログ

build\gen_data2.ps1            一対多の合成 CSV + expected.txt
build\build_workbooks2.ps1     src\v2\vba → .xlsm 2 冊 (UserForm もここで作る)
build\build_all2.ps1           一対多版だけをまとめてビルド
build\run_bench2.ps1           4 方式を N 回まわして計測
build\summarize2.ps1           ログから中央値・最小・最大
```

実用版 (比較版のソースには手を入れていません):

```
src\app\csharp\Rdv3*.cs        C# 実用版 (worker スレッド + xlsx 台帳 + WinForms)
src\app\cmd\boot-app.ps1       .cmd のブートストラップ
src\app\vba\modRdv3*.bas       VBA 実用版 v2 (小 FE + 別プロセス BE + ファイルチャネル +
                               台帳ブック/sidecar は BE 専有)

build\pack_app.ps1             src\app\csharp → dist\app-csharp\ReaderDataViewer.cmd
build\build_workbook_app.ps1   src\app\vba → 小 FE xlsm + 台帳ブック + sidecar (worker book 埋込み)
build\build_app.ps1            実用版 2 ルートをまとめてビルド (初期台帳の生成込み)
build\bench_app.ps1            実用版 2 方式の E2E 計測 v2 (全 run 記録、実メモ帳・実ダイアログ)
```

## リポジトリに入っていないもの

再生成できるものは追跡していません。

| 入っていないもの | 作り方 |
|---|---|
| `data\` の CSV 3 本 + `expected.txt` (241 MB) | `build\gen_data.ps1` |
| `data-100k\` / `data-tiny\` の CSV 3 本 + `expected.txt` | `build\gen_data2.ps1` |
| `dist\` の 1 対 1 版 4 ファイル | `build\build_all.ps1` |
| `dist\` の一対多版 4 ファイル | `build\build_all2.ps1` |
| `dist\app-vba\` / `dist\app-csharp\` の実用版 2 ルート | `build\build_app.ps1` |
| **`dist\` 全部を一度に** | **`build.bat` (= `build\build_dist.ps1`)** |
| `work\` の計測ログとビルド ログ | `build\run_bench.ps1` / `run_bench2.ps1` / `bench_app.ps1` |

## benchmarks/ — 方式選定のための検証

### [benchmarks/excel-background-bench](benchmarks/excel-background-bench)

このアプリを作る前に、「重い処理をどこで走らせ、結果をどうやって画面へ戻すか」だけを
14 通り変えて測ったものです。Excel から 1,000,000 件の郵便番号を住所へ変換します。

- 起動手段 (BAT / WScript.Shell.Run / VBA Shell / タスク スケジューラ COM) の差は
  E2E で 0.03 秒以内。**どうやって起動するかはほとんど効きません。**
- 効くのは結果の戻し方です。同じ 1,000,000 件で **0.77 秒から 10.95 秒まで 14.2 倍**開きます。
- 最速は Excel 自身のプロセスで走る Excel-DNA / XLL の 0.773 秒。

詳しくは [benchmarks/excel-background-bench/docs/results.md](benchmarks/excel-background-bench/docs/results.md)。
この検証は**方式選定の資料**であって、成果物ではありません。読み取り専用の扱いです。

本アプリの結果は、その結論と矛盾しません。ただし**返す結果が 1 件しかないので、
戻し方の差は 34 ms 対 2 ms にしかならず**、支配的なのは 241 MB を毎回読み直して
2,000,000 回の突合を行う部分でした。

## ライセンス

**CC0 1.0 Universal** (パブリックドメイン提供)。`SPDX-License-Identifier: CC0-1.0`
全文は [LICENSE](LICENSE)。著作権と関連する権利を可能な範囲で放棄しています。
帰属表示は不要です。

同梱していないもの: 画面が使う書体は Windows 標準搭載のもの (Meiryo UI / Yu Gothic UI /
MS UI Gothic / Bahnschrift) だけで、フォントファイルの再配布もインストールもしません。
