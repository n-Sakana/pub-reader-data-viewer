# マルチデザイン版の検証記録

検証日：2026-09-08（日本時間）。対象は、アップロードされた `ReaderDataViewer_fixed(1).zip` を基準にした今回の外観／ビルド改修です。

## 実施した確認

| 検査 | 結果 | 実際に確認した範囲 |
|---|---:|---|
| ６テーマのブラウザー回帰 | 194 / 194 PASS | 模擬 WebView ブリッジ、Chromium 144.0.7559.96、Linux |
| 既存の静的検査 | 30 / 30 PASS | C# の字句・括弧整合性、限定的な参照名、JS 構文、設定とCSVの整合性 |
| 新しい外観・保全検査 | 58 / 58 PASS | 元ファイル48件のSHA-256、テーマ一覧、JS構文、文字色の組合せ、スクリプトの文字コードなど |
| Win98 の画素比較 | 一致 | 上の194件に含む。同じ模擬データ・920×780・同じブラウザー条件で元版と差分なし |

**これらの合格は Windows でのコンパイル・実行成功を意味しません。** C# の静的検査はコンパイラーではありません。ブラウザー試験の名前に `native-dialog` が付く項目も、WebView の副画面モードをブラウザーで模擬したもので、WPF ウィンドウそのものの試験ではありません。

ブラウザー回帰では、元の14件の操作試験を各テーマで繰り返しました。入力正規化、IME、Enterの二重実行防止、最大長、更新／削除ジョブの振り分け、無効操作、競合保持／破棄、HTMLエスケープ、CSVの安全な既定値などを確認しています。さらに818 / 480 / 1100px幅の表示、９種類の副画面、候補一覧のキーボード操作を確認しました。モダン版ではフォーカス、短い演出、演出無効化、reduced-motion、forced-colors、表示通知後のフェードが測定サイズを変えないことも確認しています。

文字色の検査は、通常文字・補助文字・主要ボタン・選択ラベルの４組合せに対して4.5:1以上を確認する**限定的な数値検査**です。全画面・全状態のアクセシビリティ適合認証ではありません。

`web/app.js`、`web/app.css`、`settings.json`、既存の業務C#、起動ランチャー、DLL、元の入力データ、元の回帰試験など48ファイルは、アップロードされたZIPの同名ファイルとSHA-256が一致します。既存C#の変更は `src/02_MainWindow.cs` の背景・タイトルバー配色と外観専用の表示通知に限定し、`src/03_Theme.cs` を追加しています。

## 未実施の確認

Windows PowerShell 5.1 の実行環境がないため、**選択TUIの実キー操作、PowerShellによる実パッケージ生成、C#コンパイル、WPF／WebView2の実機起動は未実施**です。実際のSMB共有、複数PC同時運用、競合時の復旧、既存台帳を使う移行、DPI・日本語フォント・OSテーマとの組合せも未確認です。動作保証済みバイナリーではなく、改修済みソースとして提供しています。

`tests/Test-Build.ps1` は Windows 向けの追加統合試験です。PowerShellパーサー、６テーマの出力、ハッシュ、ZIP内容、複数回出力時の保全、演出off、サンプルなし、不正指定拒否を試験します。このスクリプト自体は今回未実行です。生成された合格報告を装っていません。

元ZIPの `AUDIT_REPORT.md`、`docs/source-changes.patch`、`tests/results/native-status.json` などは過去の業務修正履歴として残してあり、今回の追加ネイティブ試験が完了した証拠ではありません。

## Windows での確認手順

まず開発用フォルダーで次を実行します。本番台帳を指定せず、同梱の隔離試験から始めてください。

```bat
build.bat compile
build.bat test
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-Build.ps1
rem 上２つのネイティブ検査も統合試験とまとめて走らせる場合：
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-Build.ps1 -Native
```

次に `build.bat` を引数なしで開き、上下キー、Space、1〜6、A、M、F、D、T、Enter、Escを確認します。少なくとも全選択、２テーマだけ選択、空選択の拒否、キャンセル、新規フォルダーへの２回連続出力を確認してください。

各出力の `.cmd` を試験用の入力／台帳パスで起動し、検索、IME、候補選択、作業状態、送信、競合、設定、CSV出力、終了・再起動を確認します。Win98は無演出、各モダン版は短いフィードバック、`-Motion off`とWindowsの演出軽減設定では無演出になることを確認します。Win10/11のタイトルバー、100 / 150 / 200% DPI、大きな設定フォント、狭い画面、ハイコントラストも確認対象です。長い内容にスクロールが必要なこと自体は従来の仕様です。

共有台帳の複数PC試験と本番移行は、従来の `shared-ledger.md` の注意事項に従って別途実施してください。生成パッケージには業務設定がそのまま入るため、別テーマの試運転が本番パスを参照しないように必ず確認します。

## 再現用ファイル

開発用ZIP内に以下を含めています。配布用アプリのZIPには開発テスト一式を入れません。

```text
tests/test_static.py                   元の静的検査
tests/test_theme_static.py             元ファイル保全・外観の静的検査
tests/test_themes.py                   ６テーマのブラウザー試験
tests/theme_support.py                 副作用のない模擬WebView・表示データ
tests/Test-Build.ps1                   Windows用ビルド統合試験（未実施）
tests/results/static-results.json      静的検査の実測結果
tests/results/theme-static-results.json 外観・保全検査の実測結果
tests/results/theme-results.json       ブラウザー試験の実測結果
design/preserved-files.json            元ZIPのハッシュ基準
docs/theme-gallery.html                模擬データによる画面見本
docs/previews/                         実際のブラウザー描画PNG
docs/theme-changes.md                  今回の変更範囲
docs/theme-source.patch                今回のコード・設定の差分
```

Pythonの試験には Python 3.9以降、Playwright、Pillow、json5、Pygments、Node.js、Chromium が必要です。これらは**開発検査用**で、通常ビルドやアプリ起動の追加要件ではありません。

```sh
python tests/test_static.py
python tests/test_theme_static.py
python tests/test_themes.py --chromium /path/to/chromium --screenshots
# 元ZIPを別フォルダーに展開して画素比較も実行する場合：
python tests/test_themes.py --chromium /path/to/chromium --original /path/to/original/ReaderDataViewer --screenshots
```

既存の `tests/test_web.py` の14件は `test_themes.py` が各テーマに対して呼び出します。実画面のスクリーンショットは模擬データの見本であり、本番の顧客・業務台帳やWindowsのネイティブタイトルバーを撮影したものではありません。
