# Reader Data Viewer — マルチデザイン・選択式ビルド版

2026-09-08 改訂。CSV／XLSXを設定に従って統合し、検索、作業状態の変更、共有台帳への送信を行うWindows用アプリです。ソースを起動時にWindows PowerShellでコンパイルし、WPFとWebView2で画面を表示します。

> **Windows 98 の外観を既定のまま維持し、Apple / Material / Fluent（Win10・11風）/ Carbon / Spectrum の５系統を追加しました。ビルドは複数選択式です。** Windows での C# コンパイル、TUI 実機操作、WPF / WebView2 起動、実際の共有台帳試験はこの環境では未実施です。今回の検証範囲は [docs/theme-validation.md](docs/theme-validation.md) を確認してください。

## デザインを選んで出力する

ZIP を展開して **`build.bat`** を実行します。上下キー＋Spaceで複数選択、Aで全選択、Enterで出力します。Mで演出の有効／無効を選べます。旧 Win98 版は常に無演出です。通常はコンパイル検査後、`releases/build-日時-ID/` に選択した版のフォルダーと ZIP を生成します。

```bat
build.bat
build.bat package -All
build.bat package -Theme "apple,fluent" -Motion off
```

詳しい操作・配布時の注意は [docs/design-build.md](docs/design-build.md)、画面見本は [docs/theme-gallery.html](docs/theme-gallery.html) にあります。検索・状態変更・送信・設定などの業務処理は共通です。各社の公式ライブラリーではなく、設計指針を参照した独自の外観実装です。単体 EXE 化ではなく、元の起動時コンパイル方式を維持します。

**生成パッケージは現在の業務設定をそのままコピーしますが、実運用の台帳・ログ・未送信変更はコピーしません。** 同梱できるのは隔離試験用のサンプルCSVだけです。本番パスを設定したまま別テーマの試運転をしないでください。

## 維持した修正版の業務動作

同じレコードの作業状態を別の人が変更していた場合、後から送信した値で無条件に上書きせず、未送信の競合として残すようにしました。ロック取得後の最新台帳に対して、元の内容と変更前の状態を照合します。送信済みの状態への再送信は二重に計上しません。

**監視による作業状態変更は既定で手動です。** `screen.workState.trigger` を `manual` としています。外部画面の検出は検索に使用し、処理済みにする操作は作業状態ボタンで行います。従来の自動変更が業務上必要な場合は、テスト後に `automatic` を明記してください。

**CSV出力は新規ファイルのみです。** 既存ファイル、入力表、削除用入力、設定、台帳、プログラムなどへの上書きを拒否します。数式として解釈され得る値の先頭にアポストロフィを付ける「Excel向けに数式を無効化」は既定で有効です。機械連携で元の値が必要な場合は解除します。CSVの同名列は必要に応じて `A.id` のように識別可能な見出しへ変わります。

詳細・残存制約は [AUDIT_REPORT.md](AUDIT_REPORT.md) にあります。

## 起動前の準備

64ビットWindows、Windows PowerShell 5.1系の.NET Framework／WPF環境、WebView2 Runtimeを使用する構成です。同梱の `lib/` はそのまま配置してください。今回DLLの更新や組織端末の実行許可設定は行っていません。

ZIPを**展開してから**利用します。共有運用ではアプリ一式を各PCのローカルフォルダーへ置き、`paths.ledger` だけを同一の共有ファイルへ向ける構成を基本とします。ログはPC別のローカル保存にします。ローカルの未送信変更を複数PCへコピーして使い回さないでください。

`settings.json` の相対パスはアプリのフォルダーが基準です。`data.tables` と処理ジョブの相対入力ファイル名は `paths.dataDir` が基準です。設定に列挙された入力表は起動時に存在する必要があります。添付の `data/` は元ZIPの内容を変えていません。

最初はコマンドプロンプトで次を実行します。

```bat
build.bat compile
build.bat test
ReaderDataViewer.cmd
```

`compile` はコンパイルのみ、`test` は隔離したサンプルを使うC#の回帰試験です。どちらも本番台帳を更新しません。ただし、この報告時点でWindows上での合否は未確認です。`build.bat` の引数省略はデザイン選択 TUI です（`compile` の明示指定は従来どおりです）。元READMEにあった `build.bat data` や未同梱のデータ生成・過去版復元機能は提供していません。

通常起動は `ReaderDataViewer.vbs`、起動メッセージをコンソールで確認する場合は `ReaderDataViewer.cmd` を利用します。32ビットの起動元から呼ばれた場合も、存在すればSysnative経由で64ビットPowerShellを選択します。組織の実行制御により起動が制限される場合は管理者に確認してください。

## 基本操作と「保存」の意味

起動すると最初の更新ジョブで入力表と台帳を確認します。台帳の新規作成・更新は確認後に実行します。読み込めない既存台帳を、そのまま上書き再作成することはしません。

検索キーを入力して検索し、複数候補がある場合は対象を選びます。作業状態ボタンで変更した値は、まず**そのWindowsユーザーのローカル未送信変更**として保存されます。これだけでは他の人の共有台帳には反映されません。**「変更を送信」して共有台帳へ反映**してください。アプリを閉じても未送信分はローカルに残ります。別PCで起動しても元PCの未送信分は引き継がれません。

競合一覧の通常の「OK」やEscは未送信変更を保持します。不要な変更は「一覧の未送信変更を破棄」を選び、追加確認します。破棄前のバックアップファイルの場所が表示されます。破棄操作は共有台帳を変更しません。必要な変更をやり直す際は、最新台帳を読み直し、対象行を再確認してください。

CSV出力は**その時点で当該PCに見えている台帳とローカル未送信状態**のスナップショットです。他PCの未送信分や、まだ読み直していない共有変更は含みません。確定版の一覧が必要な場合は、各PCの送信と最新台帳への切替を確認してから出力してください。

## 共有運用・設定変更

共有運用の準備、移行、障害時の確認は [docs/shared-ledger.md](docs/shared-ledger.md) を参照してください。OneDrive等の同期コピーを「同時に開ける1つの台帳」として使う構成は対象外です。動作中に共有台帳をExcel等で直接編集しないでください。

JSONで変更できる範囲と注意点は [docs/settings.md](docs/settings.md) にあります。列構成・識別キー・状態の保存値を変える場合は、全PCを停止してバックアップと移行を行います。旧版と修正版を混在させないでください。

## 配布物

```text
ReaderDataViewer.cmd / .vbs   起動入口
settings.json                 設定（JSONC、schema 3）
src/                          アプリC# 24ファイルと起動PowerShell
web/                          HTML、JavaScript、CSS
lib/                          元ZIPのWebView2 DLL・表示文書
data/                         元ZIPの入力CSVとログ（内容維持）
build.bat / tools/            デザイン選択TUI / package / compile / test
design/                       ６種類のカタログ
AUDIT_REPORT.md                点検・修正・検証結果
docs/                         設定、共有運用、原README、変更一覧
tests/                        C#、ブラウザー、静的検査、隔離サンプル、検証結果
```

検査の詳細は [tests/README.md](tests/README.md) にあります。`tests/results/native-status.json` は未実施の記録であり、合格証跡ではありません。Windows試験後の実際の結果は `core-results.txt` に出力されます。

## ライセンス

元ZIPの `LICENSE`、`THIRD-PARTY-NOTICES.md`、`lib/` 内の表示文書を維持しています。元のREADMEも `docs/README.original.md` に保全しました。
