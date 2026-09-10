# ビルドと配布

ビルドの正本は `tools/Build.ps1`、入口はルートの `build.bat` です。旧 `build/build_dist.ps1` も同じ処理へ渡します。ソースのブランチで版が決まり、一回のビルドで一方を作ります。

| ブランチ | 製品 | 出力名 |
|---|---|---|
| main | レイアウトJSON設定版 | ReaderDataViewer-json-layout |
| semifixed | レイアウト固定版 | ReaderDataViewer-fixed-layout |

## 作成

```bat
build.bat package -RunTests
build.bat package -Format zip -Data none
build.bat compile
build.bat test
```

`-Format` は `both / folder / zip`、`-Data` は `sample / none`、`-OutputRoot` は出力先です。`-Theme win98` と `-All` は以前の呼出しとの互換指定です。画面の見た目は一種類で、32bit版やWindows 98用の製品を選ぶ指定ではありません。

既定の出力先は `build/packages/build-日時-識別子/` です。検査途中のビルドと利用者へ渡すリリースを分け、検収したZIPを配布します。出力先にはアプリフォルダ、ZIP、`package-manifest.json`、`build-summary.json` を作ります。`zip` では展開フォルダを残しません。各回は新しい場所へ作り、既存の設定・データ・過去ビルドを上書きしません。

## 同梱するもの

`tools/package-files.json` が必要な実行ファイル・説明書の一覧です。`src`・`web`・`lib` 全体を無条件にコピーしません。追加ファイルは用途を確認して一覧へ登録します。ライセンス表示も一覧に含みます。

設定の配布元は `configs/sample/settings.json` です。`-Data none` も同じ公開用設定を使い、入力と練習用データだけを省きます。作業ルートの `settings.json` や稼働中の `data` は配布へ取り込みません。個別環境用JSONは別途渡します。

サンプル入力は `tests/fixtures/sample-v4` のCSV2本・XLSX2本です。期間混在の入力は `samples/next-period` と `samples/partial-pay` の登録ファイルです。見出し・形式・件数を保ったままバイト単位でコピーします。必要な入力が欠ける場合はビルドを失敗させます。

`manual/release.json` は製品の種類・作成時刻・ソースcommit・未commit差分の有無を記録します。ローカルの作業パスや端末名は入れません。ビルド検査結果と全同梱ファイルのSHA-256はZIPの外の `package-manifest.json` に残します。

## 検査

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Test-Build.ps1
```

通常の配布作成はWindows PowerShell 5.1でC#をコンパイルし、`-RunTests` で中核検査も実行します。作ったZIPを新しい場所へ展開して起動・操作する確認は、ビルドのコンパイル検査とは別です。

`-SkipValidation` はビルド処理の検査専用です。`-RunTests` と同時には指定できず、スキップを合格とは記録しません。未commitの成果や検査を省いた成果を、検収済みリリースと取り違えないでください。

旧WinForms・VBA・UI試作は `archive/` に保管し、現行のビルド対象に含めません。
