# excel-background-bench

Excel から 1,000,000 件の郵便番号を住所へ変換する、起動動作テストと速度ベンチマークです。

同じ入力・同じ変換規則・同じ出力・同じ正解確認・同じ計測範囲のまま、**処理をどこで走らせ、
どうやって起動し、結果をどうやって Excel に戻すか**だけを変えて 14 方式を比較します。

計測結果は [docs/results.md](docs/results.md) にあります。要約すると:

- 起動手段 (BAT / WScript.Shell.Run / VBA Shell / タスク スケジューラ COM) の差は
  E2E で 0.03 秒以内。**どうやって起動するかはほとんど効きません。**
- 効くのは結果の戻し方です。同じ 1,000,000 件を同じ規則で変換して同じシートへ入れるのに、
  経路だけで **0.77 秒から 10.95 秒まで 14.2 倍**開きます。変換そのものはどの方式でも
  0.05 秒前後で、順位に何も寄与しません。
- 最速は Excel-DNA/XLL の **0.773 秒**。Excel 自身のプロセスで走るのでプロセス境界が無く、
  読み込みの時点ですでに COM 経由の 4 倍速い。
- ファイル経由なら **ADO が 1.027 秒、DAO が 7.379 秒**。同じファイル・同じ `schema.ini`・
  同じ `CopyFromRecordset` で **7.2 倍**違います。ADO と DAO は別物として測る必要があります。
- `QueryTable` は最下位の 10.322 秒。TSV の書き出し自体は 0.038 秒で COM 書き込みの
  22 分の 1 なのに、**Excel に読み直させた時点で負けます**。

## 何を比べているのか

| # | 方式 | 処理する場所 | 起動 | 結果の戻し方 |
|---|---|---|---|---|
| 1 | セル単位 VBA | 前面 Excel | - | その場 |
| 2 | Dictionary + Variant 配列 VBA | 前面 Excel | - | 一括 Value2 |
| 3 | BAT を手動で開く | C# ワーカー | 人が実行 | COM |
| 4 | WScript.Shell.Run | C# ワーカー | WSH | COM |
| 5 | VBA Shell | C# ワーカー | Shell 関数 | COM |
| 6 | タスク スケジューラ COM | C# ワーカー | Schedule.Service | COM |
| 7 | 不可視の別 Excel プロセス | 別 Excel | CreateObject | 一括 Value2 |
| 8 | Shell.Application.ShellExecute | C# ワーカー | ShellExecute | COM |
| 9 | QueryTable 取込 | C# ワーカー | WScript.Shell.Run | TSV + rename → QueryTable |
| 10 | COM 一括書込 | C# ワーカー | WScript.Shell.Run | COM Range.Value2 |
| 13 | ADO 経由 | C# ワーカー | WScript.Shell.Run | TSV + rename → ADO Recordset → CopyFromRecordset |
| 14 | DAO 経由 | C# ワーカー | WScript.Shell.Run | TSV + rename → DAO Recordset → CopyFromRecordset |
| 15 | OpenText 取込 | C# ワーカー | WScript.Shell.Run | TSV + rename → Workbooks.OpenText |
| 16 | Excel-DNA / XLL | Excel 自身のプロセス | RegisterXLL | XLL の C API で直接 |

方式 3〜6・8 は起動手段だけが違い、ワーカーの中身は完全に同じ 1 本のコードです。
方式 9・10・13・14・15 は起動手段を固定して、結果の届け方だけを変えた対照実験です。
うち 9・13・14・15 は**ワーカーが書いたまったく同じ 1 本のファイル**を読みます。
中身は 1 バイトも違わず、違うのは Excel 側の取り込み経路だけです。

方式 16 だけはプロセス境界がありません。Excel-DNA の XLL が Excel 自身のプロセスへ
読み込まれ、C API (XLOPER12) でセル配列を直接やり取りします。COM のマーシャリングも
RPC も通りません。変換規則は他方式と同じものを `xll/ZbXll.cs` に置いてあります。

方式 8 は起動時間が他方式から大きく外れており (中央値 53.9 秒)、原因は未特定のため
今回は保留しています。変換結果は一致しますが、起動時間の数値は使えません。

C# ワーカーは 2 通り用意してあります。**あらかじめビルドした版 (prebuilt)** と、
**Excel が実行時に BAT・PS1・C# ソースを書き出してビルドする版 (emitted)** です。
「ファイルを作るだけ」と「作って起動して変換する」は別のボタンに分けてあります。

## 動かす

### 必要なもの

- Windows 10 / 11
- Excel (デスクトップ版, VBA が使えるもの)
- .NET Framework 4.x の `csc.exe` (Windows に同梱)
- PowerShell 7 (`pwsh`) または Windows PowerShell 5.1

Excel の設定で **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」** を
有効にしてください (ファイル → オプション → トラスト センター → マクロの設定)。
ブックを組み立てるときにだけ必要です。

### 手順

```powershell
# 1. C# ワーカーをビルド (prebuilt 版の置き場所を明示する)
mkdir prebuilt -Force
pwsh -File worker\build_worker.ps1 -Out (Resolve-Path .\prebuilt).Path\ZipWorker.exe

# 2. Excel-DNA の XLL を組み立てる (方式16 用。nuget.org から取得する)
pwsh -File xll\fetch_exceldna.ps1
pwsh -File xll\build_xll.ps1

# 3. 日本郵便の公式データを取得してブックを組み立てる
pwsh -File build\build_workbooks.ps1 -Root .

# 4. ZipBench.xlsm を開いて、シートのボタンから実行
```

`-Out` を省略すると `build_worker.ps1` は自分と同じ場所に `ZipWorker.exe` を作ります。
これは emitted 版がその場でビルドするときの動きで、prebuilt 版は `prebuilt\` に
置く必要があります (無いと方式 3〜6・8〜10 がエラー 53 で失敗します)。

`build_workbooks.ps1` は `ZipBench.xlsm` (操作盤) と `ZipData.xlsx` (データ) の
2 冊を作ります。データは日本郵便の `ken_all.zip` を初回だけダウンロードし、
`data\` に置きます。取得にかかる時間は計測範囲外です。

### コマンドラインから計測する

```powershell
# 小件数で全方式の一致を確認
pwsh -File build\run_bench.ps1 -Root . -Count 2000 -Methods 1,2,4,5,6,7

# 1,000,000 件を 5 回ずつ
pwsh -File build\run_bench.ps1 -Root . -Count 1000000 -Methods 9,10 -Repeat 5
```

主なオプション: `-WorkerKind prebuilt|emitted`、`-ManualBat` (方式 3)、
`-EmitOnly` (ファイル作成だけ)、`-Repeat N`。

`run_bench.ps1` は自分で起動した Excel だけを操作し、終了時にその 1 つだけを閉じます。
**すでに開いている他の Excel には触れません。**

### 取り消す

Esc で止まります (`Application.EnableCancelKey = xlErrorHandler` でエラー 18 を捕捉)。
ワーカーを起動済みの方式では、ワーカーにも停止を伝えてから後始末します。
`build\test_cancel.ps1` が取り消し経路の動作確認です。

## 仕組み

### 変換規則

`KEN_ALL.CSV` の 3 列目 (郵便番号 7 桁) をキー、`都道府県 + 市区町村 + 町域` を値とする
辞書を作り、入力の郵便番号を引きます。見つからなければ `該当なし`。

郵便番号は数字以外を落として 7 桁に揃えます (`modZipRule.NormalizeZip`)。
全角数字・ハイフン・空白のどれが来ても同じキーになります。

この規則は VBA (`vba\modZipRule.bas`)、C# (`worker\ZipWorker.cs`)、
PowerShell (`build\make_dict_reference.ps1`) の 3 か所に独立に実装してあり、
出力が 1,000,000 行バイト一致することを確認しています。3 つ目は正解の照合用です。

### Excel と C# ワーカーの間

ワーカーは UI Automation の **プロバイダ側**として振る舞います。不可視のウィンドウを
1 つ持ち、`AutomationProperties` に自分の状態を載せます。Excel からは
`UIAutomationClient` を**事前バインディング**で使って読み書きします
(UIA のクライアント インターフェイスは `IDispatch` を持たない `IUnknown` 派生なので、
`CreateObject` による遅延バインディングでは呼べません)。

UIA で渡すのは **jobId・対象 Excel のウィンドウ ハンドル・ブック名・
Input/Master/Output の範囲・通知セルの位置**だけです。1,000,000 件の配列は
UIA を通しません。ワーカーは受け取ったハンドルから対象の Excel に自分で結び付きます。

```
Application.Hwnd → XLMAIN → XLDESK → EXCEL7 → AccessibleObjectFromWindow → Window → Application
```

`GetActiveObject` や実行中オブジェクト テーブルは使いません。それらは
**他人の Excel を掴む可能性がある**ためです。ハンドルを起点にすれば、
指定された 1 つのインスタンスにしか繋がりません。

呼び出し先の Excel が操作中でも落ちないように、ワーカー側で `IOleMessageFilter` を
登録し、`RetryRejectedCall` で再試行します。ジョブは STA スレッドで走ります。

### 完了の伝え方

ワーカーは処理が終わったら通知セルに 1 回だけ書きます。Excel は `Worksheet_Change` で
受け取り、**ハンドラの中で最初に終了時刻を打刻**します。ポーリングも固定待機も
ありません。通知の中身は次の 1 行です。

```
DONE|jobId|rows|hash|bind|master|input|dict|convert|write|workerE2E
```

ワーカーは UIA 側の結果を先に確定させてから通知セルを書きます。逆順にすると、
Excel が通知を見て UIA を読みに行ったときにまだ空、という競合が起きます。

### 別 Excel プロセス方式 (方式 7)

前面 Excel が `CreateObject("Excel.Application")` で不可視の Excel を起こし、
そちらに変換させます。実装は casedesk / xltoolrack の FE/BE 構成をそのまま踏襲しています。

- `CreateObject` の前後で Excel の PID 差分を取り、**自分が起こした 1 つだけ**を覚える
- `AutomationSecurity` を退避してから変更し、必ず戻す
- `Workbooks.Open(ReadOnly:=True, UpdateLinks:=0)`
- `workerApp.Run "Module.Proc"` で呼ぶ
- 失敗したら `Quit`

変換本体は方式 2 とまったく同じ関数 (`BuildDictFromMaster` / `ConvertBlock`) を呼びます。
固定待機・1 秒刻みの待ち・進捗ポーリング・中間ファイルはありません。

### 計測

- 時刻は `QueryPerformanceCounter` を `Currency` で受けます。VBA の `Timer` は `Single`
  なので、正午過ぎには 1 ULP が 3.9 ミリ秒になり、0.05 秒の区間を測れません。
- **処理中は `Application.StatusBar` を一切更新しません。** 件数も割合も経過秒も出しません。
  終了時刻を確定 → 計測値を確定 → 最後に 1 回だけ表示、の順です。シートへの書き込みも
  計測が終わってからです。表示の書式は全方式で共通です。
- 区間は 起動 / Master 読込 / Input 読込 / 辞書構築 / 変換 / 出力生成・転送 /
  Excel 反映 / 通知 / 処理 E2E に分けて取ります。
- 日本郵便データの取得と 1,000,000 件入力の生成は計測範囲外です。
- 失敗した方式から別方式へ切り替えることはしません。失敗はそのまま記録します。

### 正解確認

全 1,000,000 行を突き合わせます。抜き取りではありません。あわせて FNV-1a 32bit
ハッシュを C# と PowerShell で同じ規則 (UTF-16 コード単位 + 行末 `\n`) で計算し、
一致を確認しています。値は `22E2C28C` で、設計を 3 回組み直しても変わっていません。

## 構成

```
worker\ZipWorker.cs          C# ワーカー本体 (UIA プロバイダ / Excel COM / 変換)
worker\build_worker.ps1      csc でビルド
worker\run_worker.bat        起動用 BAT のひな形

vba\modZipRule.bas           変換規則・辞書構築・変換ループ・QPC タイマ (方式 2 と 7 が共有)
vba\modZipBench.bas          計測の枠組み・結果の記録・照合・取り込み経路 4 種
vba\modZipVba.bas            方式 1, 2
vba\modZipWorker.bas         方式 3, 4, 5, 6, 8, 9, 10, 13, 14, 15
vba\modZipExcel.bas          方式 7 (別 Excel プロセス)
vba\modZipEngine.bas         別 Excel 側で走る本体
vba\modZipUia.bas            UIA クライアント (事前バインディング)
vba\modZipXll.bas            方式 16 (Excel-DNA / XLL)
vba\modZipUi.bas             ボタンの入口だけ (後述)

xll\ZbXll.cs                 Excel のプロセスの中で走る変換 (C API 経由)
xll\fetch_exceldna.ps1       Excel-DNA を nuget.org から取得
xll\build_xll.ps1            XLL を組み立てる

build\build_workbooks.ps1    ブック 2 冊を組み立てる
build\gen_emitter.ps1        emitted 版のソース埋め込みモジュールを生成
build\run_bench.ps1          コマンドラインから計測を回す
build\test_cancel.ps1        取り消し経路の確認
build\make_*_reference.ps1   正解データの独立実装
build\UiaProbe.cs            UIA ツリーの確認用
build\ComWriteProbe.cs       COM 一括書き込みの単体計測用
```

`vba\modZipEmit.bas` は `build_workbooks.ps1` が毎回生成するため、リポジトリには
含めていません。

## 作るときに引っかかったこと

同じところで詰まる人がいると思うので、実測した挙動を残しておきます。

**UDT を引数か戻り値に持つ公開手続きが 1 つでもある module は、`Application.Run` から
呼べなくなります。** その module の中の、型を一切使わない空の `Sub` すら
「マクロを実行できません。このブックでマクロが使用できないか…」で弾かれます。
だから `modZipUi.bas` に型を使わない薄い入口だけを置いて、中身は各 module に委ねています。
図形の `OnAction` もすべてこの module を指します。

**`Sub X(): 処理: End Sub` の 1 行形式も同じ症状を起こします。** コンパイルは通るのに、
その module 全体が `Application.Run` から呼べなくなります。
`build_workbooks.ps1` に検査を入れてあります。

**VBA の識別子は大文字小文字を区別しません。** `ZB_SHEET` という定数と `ZB_Sheet()` という
関数が同居すると、プロジェクト全体がコンパイルできなくなります。しかも表に出る症状は
「マクロが使用できないか…」だけで、どこが悪いのか何も言いません。これも検査を入れました。

**module レベルの宣言は、どの手続きよりも前に置く必要があります。** `Declare` を
`Const` ブロックより上に書いただけで、同じ不親切なエラーになります。これも検査対象です。

**`.bas` はシステムの ANSI コードページで読まれます。** 日本語環境では CP932 なので、
ソースに CP932 に無い文字 (U+2012、U+2013、U+2014、U+2212、U+301C など) を書くと
そこで壊れます。ビルド時に CP932 の往復チェックをして、失われる文字があれば止めます。

**`Shell.Application.ShellExecute` は .bat に引数を渡しません。** 渡したつもりの引数が
消えるので、起動手段ごとに token・mode・pid を埋め込んだ `launch.bat` を毎回書き出す形に
統一しました。

**ADO と DAO のテキスト ドライバは UTF-8 の BOM をデータとして読みます。**
先頭行だけが化けて、他の行は正しい、という形で出ます。`QueryTable` と
`Workbooks.OpenText` は符号化を明示して渡すので BOM を必要としません。
4 経路が同じ 1 本のファイルを読めるよう、BOM を付けない側へ揃えました。

**ACE のテキスト ドライバは `.tsv` を認識しません。** 既定で表として開くのは
txt / csv / tab / asc だけです。レジストリの `DisabledExtensions` をいじれば通せますが、
比較のために利用者の環境を書き換えるのは筋が悪いので、拡張子を `.txt` にしています
(中身はタブ区切りのまま)。

**ADO も DAO も `schema.ini` が無いと列の型を推測します。** 先頭数行から判断するので、
郵便番号側で数値化が起きます。`Format=TabDelimited` / `ColNameHeader=False` /
`CharacterSet=65001` / `Col1=addr Text Width 255` を書いて型を固定しています。
これは取り込み経路の一部なので、計測の中に入れてあります。

**不可視のウィンドウは UIA ツリーに現れません。** ワーカーは `Show()` の後に
`ShowWindow(SW_SHOWNOACTIVATE)` を呼んで、画面には出さずにツリーには載る状態にします。

**ワークブックが 1 つも開いていない Excel に `Calculation` を設定すると 1004 です。**
ブックを開いた後に設定します。

## 他人の Excel に触れないこと

このベンチマークは、自分が起動したプロセスと、明示的にハンドルを渡された Excel にしか
触れません。

- `GetActiveObject` / 実行中オブジェクト テーブルは使いません
- 別 Excel 方式は `CreateObject` の前後の PID 差分で自分の 1 つだけを特定します
- `run_bench.ps1` は自分が起動した Excel の PID だけを覚えて、それだけを閉じます
- ワーカーは UIA で渡されたウィンドウ ハンドルからしか Excel に繋ぎません

## このディレクトリの位置づけ

これは [Reader Data Viewer](../../README.md) の**方式選定のための検証資料**です。
リポジトリの主役はこれから作る完成アプリのほうで、こちらは成果物ではありません。

## ライセンス

MIT License。リポジトリ直下の [LICENSE](../../LICENSE) を参照してください。

日本郵便の郵便番号データと Excel-DNA は本リポジトリには含まれていません。どちらも
ビルド時に取得します (`build/build_workbooks.ps1`, `xll/fetch_exceldna.ps1`)。
それぞれの扱いは配布元の定めるところに従ってください。
