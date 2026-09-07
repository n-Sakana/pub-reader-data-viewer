# 検証コードと実行結果

## この配布物の検証状況

| 種別 | この点検での実施状況 | 記録 |
|---|---|---|
| ブラウザーのJS／DOM | Chromiumで14項目通過 | `results/web-results.json` |
| 元ZIPの再現比較 | 選定3項目で旧挙動を再現。修正版では同項目通過 | `results/original-web-reproductions.json` |
| 静的・サンプル整合 | 30項目通過 | `results/static-results.json` |
| Windows C#コンパイル | **未実施** | `results/native-status.json` |
| C#回帰試験50項目 | **コードを提供、未実施** | 実施すると`results/core-results.txt`を生成 |
| WPF／WebView2／UI Automation | **未実施** | 実機受入試験が必要 |
| 実際の複数PC・SMB・障害回復 | **未実施** | `../docs/shared-ledger.md`の試験表 |

静的チェックが通ることはC#のコンパイル成功を意味しません。ブラウザーのブリッジを模擬した検査が通ることも、Windowsホストや共有台帳への保存が成功する証明ではありません。

## Windows上の検査

アプリフォルダーに移動したコマンドプロンプトで実行します。

```bat
build.bat compile
build.bat test
```

`compile`は起動時と同じソース統合・参照アセンブリーでコンパイルのみを行います。`test`は同じコンパイルに`RegressionTests.cs`を加え、テストメソッドを呼びます。通常のアプリ画面は起動しません。Windows PowerShellと.NET Framework／WPF、および同梱のWebView2管理DLLが読める環境が必要です。

50項目はJSON、正規表現、未送信の競合判定、再送、CSV、XLSX、出力保護、設定保存競合、ローカルセッション、共有ロック、8スレッドの競合書き込み等を含みます。**8スレッド試験は同一Windows環境の一時フォルダー上での試験**であり、複数PCやSMBの代替ではありません。

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

`--original`を指定すると元の全ファイルの残存と、入力・DLL・ライセンス類のバイト一致を確認します。この1項目を含め、この監査では30項目です。オプションなしでは29項目です。ルートの設定・入力を書き換えた後はサンプル整合検査の条件や結果も変わります。

## 同梱の結果ファイルの時刻

結果の時刻はUTCです。日本時間の2026-09-08に行った検証が、JSON内では2026-09-07のUTC時刻になっています。各検証の実時刻・ブラウザー版・プラットフォームは結果ファイルを参照してください。
