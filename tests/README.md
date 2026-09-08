# 検証コードと実行結果

## 2026-09-09 構造入力の除外と計算結果の説明

C#回帰は94/94（既存91件と追加3件）。列不足・空行の件数、元の行番号、未参照の重複列の除外と列位置、参照する重複の拒否、台帳生成・処理ダイアログの検証内容・headlessの一致を確認します。保存済み台帳や共有・終了ガードの検査は保持しています。

`python tests/test_input_boundaries.py` は7/7、`python tests/test_feedback_hints.py` は8/8、`python tests/test_readme_example.py` は5/5です。前者の重複見出しは台帳に保存する列なので引き続き拒否し、案内文の期待値を更新しました。4表例のsummaryは追加した3つの除外件数を0として照合します。既存検査の削除・skip・期待値の緩和はしていません。

実素材を使った単表のRunUpdateでは、列不足7行と空行3行を除外して29,993行。別表は未参照の重複2列を除外して18行、同じ重複名を保存する指定はexit 3でした。READMEの計算例はselectで未使用列を外してからcalculateし、保存結果6と250を確認しました。これらは業務全体の委託検収や、本番共有先・複数PCの確認を意味しません。

## 以前の検証記録

2026-09-09、共有更新と終了保留の組み替え後にWindowsで再実行しました。C#は30ファイルです。ブラウザー検査は模擬ブリッジを使います。これとは別に、実際のWPF/WebView2の窓をUI Automationから操作し、下記の範囲を確認しました。

| 種別 | この点検での実施状況 | 記録 |
|---|---|---|
| ブラウザーのJS／DOM | Chromiumで20項目通過 | `results/web-results.json` |
| 元ZIPの再現比較 | 過去の選定3項目の記録。今回は再実行していない | `results/original-web-reproductions.json` |
| 静的・サンプル整合 | 36項目通過 | `results/static-results.json` |
| Windows C#コンパイル | PowerShell 5.1で全ソースをコンパイル、exit 0 | 配布処理の`build-summary.json` |
| C#回帰試験63項目 | 63 / 63 PASS | `results/core-results.txt` |
| Win98の外観・ソース照合 | 53 / 53 PASS | `results/theme-static-results.json` |
| Win98のブラウザー回帰 | 29 / 29 PASS。旧Win98との主画面画素比較を含む | `results/theme-results.json` |
| WPF／WebView2／UI Automation | 144 DPI、同梱データで台帳作成・検索・候補選択・状態変更・両方向の送信・終了を確認 | `results/native-smoke.json`。IME、他のDPI、全ダイアログ等は未実施 |
| 実際の複数PC・SMB・障害回復 | **未実施** | `../docs/shared-ledger.md`の試験表 |

静的チェックが通ることはC#のコンパイル成功を意味しません。ブラウザーのブリッジを模擬した検査が通ることも、Windowsホストや共有台帳への保存が成功する証明ではありません。

## Windows上の検査

アプリフォルダーに移動したコマンドプロンプトで実行します。

```bat
build.bat compile
build.bat test
```

`compile`は起動時と同じソース統合・参照アセンブリーでコンパイルのみを行います。`test`は同じコンパイルに`RegressionTests.cs`を加え、テストメソッドを呼びます。通常のアプリ画面は起動しません。Windows PowerShellと.NET Framework／WPF、および同梱のWebView2管理DLLが読める環境が必要です。

63項目はJSON、正規表現、未送信の競合判定、再送、CSV、XLSX、出力保護、設定保存競合、ローカルセッション、共有ロック、8スレッドの競合書き込み等を含みます。元の50項目とJSON文字列の1項目に、共有更新・戻し通知・終了保留の12項目を追加しました。確認後の台帳変更、保存失敗、保存後の通知失敗、ロック待ち中の他者送信、終了とロック取得の競合、期限超過後の処理継続を再現します。**8スレッド試験は同一Windows環境の一時フォルダー上での試験**であり、複数PCやSMBの代替ではありません。

C#試験の設定と入力データは`tests/fixtures/`の隔離コピーを使います。ルートの運用設定や本番台帳を読み書きするテストではありません。試験用のCSV・XLSX・未送信データは新しい一時フォルダーに生成し、最後に削除を試みます。削除できなかった場合はその旨を表示します。試験結果はアプリ配下の`tests/results/core-results.txt`へ出力するため、その場所への書込権限は必要です。

終了コード0で全項目PASSです。失敗時には該当ケースと例外が出ます。コンパイルに失敗した場合はテスト自体が未開始であり、以前の結果ファイルを今回の合格結果と解釈しないでください。`native-status.json`はこの監査時点の固定記録で、自動更新しません。

## ブラウザー側の検査

Python 3とPlaywright、Chromiumを使います。例えば別の検証用Python環境で準備して実行します。

```text
python -m pip install playwright
python -m playwright install chromium
python tests/test_web.py
```

既存のChromium実行ファイルを指定する場合は次の形式です。

```text
python tests/test_web.py --chromium /usr/bin/chromium
```

HTML・CSS・JSの同梱ファイルをブラウザーへ注入し、架空のレコードとWebView2ブリッジを使って検査します。ファイルURLやHTTPサーバーのアクセス可否には依存しません。実アプリでのリソース読み込み・ナビゲーション制限・C#メッセージ処理の検証ではありません。

入力DOMの正規化、IME Enter、検索イベント数、無効化、重複ボタンID、モーダル、競合変更の保持／破棄、HTML文字列の扱い、出力の数式対策選択・拡張子等を確認します。IMEは`isComposing`等を設定した合成キーボードイベントを用いており、実際の日本語IME製品との相性検証ではありません。

出力ダイアログの状態を列選択・絞り込み・出力先へ分けた後、6項目を追加しました。左右の移動、キーボード選択と既定復帰、文字列条件と削除、数値・日付条件、検証エラーからの再試行、閉じた画面や別tokenへの遅い応答を確認します。追加検査の初回は19/20で、同じ検査を変更前の製品へ実行しても同じ1件が失敗しました。`f3`は表を包む要素のクラスであるため、追加検査のセレクタを1回訂正し、20/20となりました。既存14項目は削除・変更していません。

元ZIPを別に展開して比較する場合：

```text
python tests/test_web.py --baseline --root PATH_TO_ORIGINAL --output baseline.json
```

`--baseline`は元版の3件を再現するための選定モードです。元版に対しては3件のFAILと終了コード1が想定されます。これは修正版テストが失敗しているという意味ではありません。対象を区別せず集計しないでください。

## 静的・サンプル整合検査

Pythonの`pygments`、`json5`、PATH上のNode.jsを使います。

```text
python -m pip install pygments json5
python tests/test_static.py
python tests/test_static.py --original PATH_TO_ORIGINAL
```

C#の字句／括弧、文字列定数と追加ヘルパーの参照名、NodeのJS構文、同梱JSONC、入力CSVの形・キー・列参照を調べます。Pygmentsが誤認する有効なUnicodeエスケープ文字リテラルは字句検査時のみ置き換えます。C#の型検査、オーバーロード解決、IL生成、API互換性検査ではありません。

`--original`を指定すると元の全ファイルの残存と、入力・DLL・ライセンス類のバイト一致を確認します。現在はオプションなしで36項目、指定すると37項目になります。今回はオプションなしで実行しました。共有更新・戻し通知・終了保留の3ファイルを加えたため、ソースごとの字句検査が3項目増えています。ルートの設定・入力を書き換えた後はサンプル整合検査の条件や結果も変わります。

## 同梱の結果ファイルの時刻

結果の時刻はUTCです。日本時間の2026-09-08に行った検証が、JSON内では2026-09-07のUTC時刻になっています。各検証の実時刻・ブラウザー版・プラットフォームは結果ファイルを参照してください。
