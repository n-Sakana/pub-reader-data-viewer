# デザイン選択・ビルド手順

## 構成

従来の **Windows 98 Classic** を残し、次の５系統を追加しています。合計６種類の外観です。

| ID | 参照した設計体系 | 実装上の特徴 |
|---|---|---|
| `win98` | 既存の Win98 風 UI | 元の枠・色・太字・密度。アニメーションなし |
| `apple` | Apple Human Interface Guidelines / macOS | ニュートラルな面、控えめな丸み・影、青い主要操作 |
| `material` | Google Material Design 3 | トーンの異なる面、丸いボタンとグループ、短い状態変化 |
| `fluent` | Microsoft Fluent / Windows 10・11 | Windows 系の角丸・奥行き、明瞭な入力線とフォーカス |
| `carbon` | IBM Carbon | 角形、強い区切り、情報密度を重視する productive motion |
| `spectrum` | Adobe Spectrum | 中立色の作業面、輪郭を明確にした操作、ピル形ボタン |

いずれも既存の HTML コンポーネントに対する**独自の外観実装**です。各社の公式コンポーネントライブラリーを移植したもの、完全再現、認証済み実装ではありません。Windows アプリが macOS ネイティブアプリになるわけではありません。Apple 固有フォント、各社ロゴ、外部 Web フォントは同梱せず、ネットワーク接続を増やしていません。

Fluent は Win10/11 風のクライアント領域です。タイトルバー、移動・サイズ変更・最大化・スナップ・ファイル選択などは引き続き OS が提供します。Win11 のタイトルバー色指定に非対応の Windows では、その部分だけ OS 標準表示になります。５種類のモダン UI のために OS のシステムボタンを置き換えてはいません。

## 最初に

開発用 ZIP を展開し、`build.bat` をダブルクリックします。引数省略の動作は、旧版の「コンパイルのみ」から **選択式 TUI** に変わりました。アプリ自体を起動する `ReaderDataViewer.vbs` / `.cmd` は従来どおりです。開発用フォルダーをそのまま起動した場合は Win98 版です。

ビルドとアプリ起動の対象環境は、元の Windows PowerShell 5.1 / .NET Framework / WPF / 64 ビット Windows 構成です。アプリの画面表示には WebView2 Runtime が必要です。Node.js、npm、各社 SDK、オンライン CDN は実行時・通常ビルド時ともに不要です。

## TUI 操作

```text
 READER DATA VIEWER / DESIGN BUILD

 > [x] 1. Windows 98 Classic
   [ ] 2. Apple HIG / macOS inspired
   [ ] 3. Google Material Design 3 inspired
   [ ] 4. Microsoft Fluent / Windows 10-11 inspired
   [ ] 5. IBM Carbon inspired
   [ ] 6. Adobe Spectrum inspired

 Up/Down: move    Space/1-6: select    A: all/clear
 M: Motion       F: Format           D: Data
 T: Core tests   Enter: build        Esc/Q: cancel
```

上下キーで移動し、Space または 1〜6 で選択状態を切り替えます。複数選択できます。`A` で全選択／全解除、Enter で出力します。選択が空の場合は実行しません。Esc / Q でキャンセルするとファイルを書き込みません。

`M` は演出を `auto`／`off`、`F` は `both`（フォルダー＋ZIP）／`folder`／`zip`、`D` はサンプル CSV を同梱する `sample`／同梱しない `none`、`T` は元の隔離 C# 回帰試験の実施有無を切り替えます。Win98 の演出は常に off です。通常ビルドでは選択内容にかかわらず、まず全 C# ソースをコンパイル検査します。

小さいコンソールでは番号をカンマ区切りで入力する方式に切り替わります。この場合、演出・形式などには既定値またはコマンドラインで指定した値を使用します。入出力がリダイレクトされている場合は対話入力せずエラーで終了するので、次のコマンド方式を使用してください。

## コマンド方式

```bat
rem 一覧
build.bat list

rem ６種類すべて。既定でフォルダーと ZIP を出力
build.bat package -All

rem Apple と Fluent のみ。演出を無効化
build.bat package -Theme "apple,fluent" -Motion off

rem 全６種類 + 元の隔離 C# 回帰試験
build.bat package -All -RunTests

rem ZIP のみ、入力 CSV なし、保存先を指定
build.bat package -Theme "carbon,spectrum" -Format zip -Data none -OutputRoot "C:\RDV releases"

rem 旧版と同じ検査コマンド
build.bat compile
build.bat test
```

`-All` と `-Theme` は同時指定できません。未知の ID、テーマ未指定の `package`、不正な値はエラーになります。空白を含むパスは引用符で囲んでください。

ビルド処理の検査専用に `-SkipValidation` があります。この場合は Windows の C# コンパイル検査をせず、マニフェストに `not_run_explicit_skip` と明記します。**本番配布に向けた合格を意味しません。** `-RunTests` との同時指定はエラーです。通常は使用しないでください。

## 出力物

```text
releases/
  build-YYYYMMDD-HHMMSS-xxxxxxxx/
    ReaderDataViewer-win98/       アプリ一式
    ReaderDataViewer-win98.zip
    ReaderDataViewer-apple/
    ReaderDataViewer-apple.zip
    ...選択したテーマだけ...
    build-summary.json
```

各アプリには `PACKAGE-README.txt` と `package-manifest.json` が入ります。マニフェストにはテーマ、演出、検証の実施状況、全ペイロードファイルの SHA-256 を記録します。マニフェスト自身はハッシュ一覧に含めません。

出力は一時フォルダー内で準備した後、新規フォルダーとして公開します。過去のビルド・ソース・設定・業務データを上書きしません。準備中に失敗した場合は、そのビルド専用の一時フォルダーだけを削除します。`src` / `web` / `lib` / `data` などの入力ディレクトリー内部は出力先に指定できません。

**元アプリと同じ「起動時にソースをコンパイルする配布方式」であり、単体 EXE を作る機能ではありません。** ZIP 内の `.cmd` / `.vbs` を直接実行せず、必ず全体を展開してから起動します。

## データと既存環境を守るために

`settings.json` は開発用フォルダーの内容を**バイト単位でそのまま**コピーします。テーマごとに業務設定を書き換えたり、新しい設定項目を業務 JSON に追加したりしません。

サンプル同梱時にコピーするのは `tests/fixtures/data/` の `tableA.csv`、`tableB.csv`、`tableC.csv`、`delete.csv` の４ファイルだけです。実行中の `data/` ディレクトリー、台帳、ログ、CSV 出力結果、ユーザー別の未送信変更、WebView2 キャッシュはコピーしません。`-Data none` は空の `data/` と `output/` を作ります。

設定を独自に変更している場合、同梱サンプルとその設定が一致するとは限りません。設定が本番の絶対パスや共有パスを指す場合は、**生成した別テーマの試運転でも本番データを指し得ます。起動前に設定を確認し、試験用入力と試験用台帳のパスに変更してください。** サンプル同梱は本番データからの隔離を保証する機能ではありません。

既存の運用フォルダーへ導入するときは、先にアプリを終了し、台帳とローカル未送信変更をバックアップします。可能な未送信変更は送信を完了してから切り替えます。既存の `settings.json` と業務用 `data/` をサンプルで置き換えず、アプリのプログラムと画面ファイルだけを更新してください。`src` と `web` は同じ版を一組として更新します。

６個の出力フォルダーは６個の別配置です。未送信変更を自動移行・共有するものではありません。運用中に外観だけを差し替える場合も作業フォルダーと台帳のパスを維持し、複数の配布フォルダーを同じ作業者の本番環境として同時運用しないでください。共有台帳自体の複数 PC 運用に関する従来の注意事項は `shared-ledger.md` にあります。

## 演出と見やすさ

Win98 では旧 CSS をそのまま使用します。モダン版ではホバーと押下を中心に、90〜120ms の状態変化、80ms の小さな押下移動、120〜160ms のダイアログ・カレンダーの短いフェードを使用します。閉じる操作と業務処理は即時です。アニメーション終了を待つ処理、ループ、点滅、派手なリップル、バウンド、画面全体のスライド、ぼかしは追加していません。

`auto` は OS / ブラウザーの `prefers-reduced-motion` を尊重します。演出を減らす設定または forced-colors が有効なら動きを止めます。`off` は OS 設定に関係なく動きを止めます。設定変更中のネイティブダイアログの演出もキャンセルします。読み込み済みのネイティブ副画面は、サイズ確定と表示の通知を受けてから短くフェードします。サイズ報告が繰り返されてもフェードを繰り返しません。

操作の意味をアニメーションだけに頼らず、既存の文字・状態表示を保持します。入力可能な欄、主要操作、選択行、未送信数、エラーを見分けやすくし、モダン版のキーボードフォーカスは 2px の輪郭で表示します。長い内容は従来どおりスクロールできます。元のウィンドウ開始サイズを維持するため、小さい画面や大きいフォント設定ではスクロールが必要です。

## 実装の境界

- `design/themes.json`：ビルド画面用の６種類の一覧とネイティブ配色。
- `web/themes.css`：モダン版だけに適用する見た目・密度・演出。
- `web/theme.json` と `web/index.html` のテーマ属性：ビルド時に一組で生成。
- `web/theme-motion.js`：ネイティブ副画面の表示後フェードだけを担当。
- `src/03_Theme.cs` と `src/02_MainWindow.cs` の限定差分：背景・タイトルバー色と表示通知。
- `tools/Build.ps1` と `build.bat`：選択とパッケージ作成。

`web/app.js`、`web/app.css`、既存の検索・状態遷移・台帳・監視・CSV / XLSX・共有・設定処理の C# は変更していません。新しい GUI 操作やテーマ選択メニューを業務画面へ追加してはいません。テーマ選択はビルド時です。

## 検証

今回の検査結果と Windows で必要な追加確認は `theme-validation.md`、機械可読結果は `../tests/results/theme-results.json` などを参照してください。ブラウザー試験は模擬 WebView ブリッジによるもので、WPF と実際の共有台帳試験の代わりにはなりません。

開発用 ZIP では次を利用できます。

```bat
build.bat compile
build.bat test
powershell -NoProfile -ExecutionPolicy Bypass -File tests\Test-Build.ps1
```

## 設計資料

参照日：2026-09-08。各社の方針を参照しつつ、速度上限と演出の範囲は本アプリの業務用途向けに独自に制限しています。

1. Apple, Human Interface Guidelines / Motion: https://developer.apple.com/design/human-interface-guidelines/motion
2. Google, Material Design 3 / Motion: https://m3.material.io/styles/motion/overview/how-it-works
3. Microsoft, Fluent 2 / Motion: https://fluent2.microsoft.design/motion
4. IBM, Carbon / Motion: https://carbondesignsystem.com/elements/motion/overview/
5. Adobe, Spectrum / Motion: https://spectrum.adobe.com/page/motion/

Material 3 の現行資料には physics / expressive motion もありますが、この実装では業務用途を優先し、ばねの振動や長い演出は使用しません。Apple の透過・ガラス効果や Fluent の OS 素材を WebView 内で擬似的に全面再現することも避けています。
