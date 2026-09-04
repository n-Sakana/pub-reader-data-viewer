# Reader Data Viewer

Reader Data Viewer は C# / WPF の窓に WebView2 を載せる Windows デスクトップアプリです。

`ReaderDataViewer.vbs` から黒いコンソールを出さずに起動し、承認済み v13 の HTML 画面を表示します。検索、候補選択、処理状態、台帳更新・削除、送信、CSV 出力、設定、UI Automation 監視は WebView2 の JSON ブリッジから既存の処理本体へ接続されています。画面の構成と寸法はルートの `settings.json` にある `screen` 定義から組み立てます。

通常起動は `ReaderDataViewer.vbs` をダブルクリックします。起動ログをコンソールで確認するときだけ `ReaderDataViewer.cmd` を使います。どちらも管理者権限やインストールを必要としません。

```text
ReaderDataViewer.vbs     コンソールを出さない通常入口
ReaderDataViewer.cmd     起動失敗をコンソールで確認する入口
ReaderDataViewer.ps1     WPF / WebView2 の起動と C# の Add-Type コンパイル
src/                     現役版の C# ソース
web/index.html            v13 モックから取り出した窓とダイアログの正本
web/app.js                screen 定義の描画と WebView2 操作ブリッジ
web/app.css               settings.json の寸法を受ける v13 補助スタイル
settings.json             データ、処理、画面の設定
lib/                     WebView2 の再配布 DLL と表示文書
archive/winforms/        フェーズ17完了時点で凍結した WinForms 版
```

WinForms 版は [archive/winforms](archive/winforms/) に、内容を変えず一式で保全しています。現役版はそこからコードを参照しません。

## ライセンス

本体は [CC0 1.0 Universal](LICENSE) です。同梱する WebView2 の表示は [第三者ソフトウェアの表示](THIRD-PARTY-NOTICES.md) を参照してください。
