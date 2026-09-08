# Win98版のビルドと配布

画面はWindows 98 Classicだけです。文字・枠・色・密度と無演出の動作を維持します。タイトルバー、移動・リサイズ・スナップ、ファイル選択はOSが提供します。タイトルバーの色指定に非対応のWindowsでは、その部分だけOS標準表示になります。

## 作成する

開発用フォルダーで `build.bat` を実行します。Fで出力形式（フォルダー＋ZIP／フォルダー／ZIP）、DでサンプルCSVの同梱、Tで隔離C#回帰試験の有無を選び、Enterで作成します。EscまたはQでキャンセルします。通常は先にC#のコンパイル検査を通します。

入出力がリダイレクトされている場合は、対話入力を待たずエラーになります。次のコマンド方式を使ってください。

```bat
build.bat package -Theme win98
build.bat package -Format zip -Data none -RunTests
build.bat package -OutputRoot "C:\RDV releases"
build.bat compile
build.bat test
build.bat list
```

`package` のテーマ指定は省略できます。`-Theme win98` と `-All` は従来の呼出しとの互換用で、いずれもWin98を1組だけ作ります。両方の同時指定、Win98以外のテーマ、不正な値はエラーになります。

`-Format` は `both / folder / zip`、`-Data` は `sample / none`、`-OutputRoot` は出力先、`-RunTests` はC#回帰試験の実行です。通常のビルドにNode.js、npm、外部SDK、オンラインCDNは不要です。

ビルド処理の検査専用に `-SkipValidation` があります。明示した場合だけコンパイル検査を省略し、マニフェストに `not_run_explicit_skip` と記録します。本番配布に向けた合格を意味しません。`-RunTests` との同時指定はエラーです。

## 出力

```text
releases/
  build-YYYYMMDD-HHMMSS-xxxxxxxx/
    ReaderDataViewer-win98/
    ReaderDataViewer-win98.zip
    build-summary.json
```

各アプリには `PACKAGE-README.txt` と `package-manifest.json` が入ります。Win98・無演出であること、検証の実施状況、全ペイロードファイルのSHA-256を記録します。マニフェスト自身はハッシュ一覧に含めません。

開発用の `design/preserved-files.json` は、現行ソースの期待ハッシュです。意図して変更した既存ファイルは元ZIPの値を `baselineSha256` に残し、新しいファイルも検査対象へ加えます。実行時の `data/` とGit追跡外のファイルを除く規則は従来どおりです。

新しい一時フォルダーで準備し、新しい出力先へ移します。過去のビルドを上書きしません。準備中に失敗した場合は、そのビルド専用の一時フォルダーだけを片づけます。`src / web / lib / data / tests / tools / design / docs` など入力木の内部は出力先にできず、リンクされた入力も拒否します。

## 起動とデータの保護

64ビットWindows、Windows PowerShell 5.1、.NET FrameworkのWPF、WebView2 Runtimeを使います。ソースを起動時にコンパイルする配布方式です。単体EXEではありません。ZIPをすべて展開してから `ReaderDataViewer.vbs` を起動します。コンソールで診断する場合は `ReaderDataViewer.cmd` を使います。

`settings.json` は開発用フォルダーの内容をバイト単位でコピーします。サンプル同梱時は `tests/fixtures/data/` の `tableA.csv / tableB.csv / tableC.csv / delete.csv` だけを使い、稼働中の `data/`、台帳、ログ、出力、未送信変更、WebView2キャッシュはコピーしません。サンプルが欠けている場合は警告を出し、残っているサンプルとともに配布物を作ります。`-Data none` は空の `data/` と `output/` を作ります。

独自設定とサンプルが一致するとは限りません。本番の絶対パスや共有パスを記入している場合、生成物の試運転でも本番データを指し得ます。試運転前に入力と台帳を試験用のパスへ変えてください。サンプル同梱は本番環境からの隔離を保証しません。

既存環境へ導入するときは、アプリを終了し、台帳とローカルの未送信変更をバックアップします。既存の設定とデータをサンプルで置き換えず、同じ版の `src/` と `web/` を一組で更新します。配置先と台帳パスを維持し、複数の配布フォルダーを同じ作業者の本番環境として同時に使わないでください。共有台帳を複数PCで利用する条件は [shared-ledger.md](shared-ledger.md) を参照してください。

## 確認

```bat
build.bat compile
build.bat test
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-Build.ps1
python tests/test_static.py
python tests/test_web.py
python tests/test_theme_static.py
python tests/test_themes.py --chromium <chrome.exe>
```

ブラウザ検査は模擬WebViewブリッジによるDOMと操作の確認です。WPFと実際の共有台帳、実際のUIA対象を確認した結果とは区別してください。コンパイル・検査・配布の成否は、それぞれ実際に実行した結果と `tests/results/` を参照します。
