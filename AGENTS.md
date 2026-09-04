# AGENTS.md - pub/reader-data-viewer

## 現在の製品

- 現役版は C# / WPF の窓に WebView2 を載せる版。起動は `ReaderDataViewer.vbs`、実体は `ReaderDataViewer.ps1` と `src/*.cs`。
- `ReaderDataViewer.ps1` は `src/*.cs` を名前順に一度だけ連結し、Windows PowerShell の `Add-Type` でコンパイルする。
- WebView2 の managed DLL はパスからではなくバイト列で読み、zip 由来の Mark of the Web の影響を受けないようにする。
- UI 正本は `C:\repos\fin\deck\data\workspaces\iris\reader-mock\v13.html`。`web/index.html` の `.stage` は正本から内容を変えず抽出したもの。上のデモ操作列と下の変更一覧は製品へ入れない。
- WinForms 版は `archive/winforms/` に凍結済み。明示依頼がない限り、中身の修正、整形、参照切れ修理をしない。現役版との共通ライブラリも作らない。

## ガードレール

- 機能は凍結した WinForms 版から増減させない。画面は v13 にないものを足さず、あるものを省かない。
- 業務固有の判断や言葉をソース、設定、画面へ埋め込まない。動作と寸法は最終的に `settings.json` から組み立てる。
- C# は C# 5 の範囲に置き、C# ソースは ASCII のみとする。
- `.vbs` は ASCII のみ。通常起動でコンソールを出さず、管理者権限やインストールを要求しない。
- 先生の画面へ検証用ウィンドウを出さない。`RDV_WEBVIEW2_PROBE_OUTPUT` を指定した検査では画面外に置き、PNG を保存して確認する。
- `archive/`、`dist/`、`data-1k/`、`data-100k/`、`samples/`、`work/` の既存内容を削除しない。
- commit / push は明示指示があるまで行わない。

## 現在の区切り

フェーズ18の JSON ブリッジまで実装済み。`web/index.html` の `.stage` は v13 正本のまま保持し、`web/app.js` が `settings.json` の `screen.sections` を描画して `src/Rdv3App.cs` の既存業務ロジックへ操作を返す。次の検収では複数の設定値を振った実画面確認と、機能・DOM 回帰を行う。
