# Reader Data Viewer

Windows 上で動く C# / WinForms アプリです。UI Automation で監視対象の入力欄を読み、CSV に JSON で並べた表操作を適用して検索結果を表示します。画面も JSON の画面定義から組み立てます。

現在の製品版は **C# 版 1 種類**です。VBA 版、方式比較、方式選定ベンチ、旧 UI 案、VBA Pixel Bridge 展示は
すべて [`archive/`](archive/) に保全してあり、通常のビルドには入りません。

## 配布物を作る

リポジトリ直下の `build.bat` をダブルクリックします。管理者権限、Excel、インストール作業は不要です。

生成先は `dist\app-csharp\` です。

```text
ReaderDataViewer.vbs          通常の入口。コンソールを表示しない
ReaderDataViewer.cmd          コンソールを確認したいときの入口
settings.json                 設定 (場所、監視対象、番号の形式、CSV の定義、画面定義)
ReaderDataViewer-Ledger.xlsx  統合台帳
data\tableA.csv
data\tableB.csv
data\tableC.csv
```

Windows PowerShell 5.1 と、Windows に含まれる .NET Framework の C# コンパイラを使います。`dist\` と入力データは生成物なので Git には含めません。

同梱のデータと画面は脱色したサンプルです (表A / 表B / 表C、`SAMPLE-A-0000001` のような値)。現実的な
ダミーで動かしてみるには `build\gen_samples.ps1` が `samples\` に 3 組 (販売 / 製造 / 施設予約) を
作るので、その `settings.json` と `data\` を `.vbs` の隣に置いて起動します。

## 使う

1. `dist\app-csharp\ReaderDataViewer.vbs` を起動します。
2. 監視対象の入力欄へ番号を入れるか、画面の検索欄へ直接入力します。
3. 複数候補がある場合は候補一覧から 1 件を選びます。
4. 判定 (OK / NG / 未定義) は画面定義の規則で CSV の値から決まります。
5. 作業状態 (未処理 → 処理済) は手元へ控えられ、「送信」を押したときだけ共有台帳へ反映されます。

ジョブは結合、抽出、追加、更新、削除、列操作、計算、集計、並べ替え、重複除去、マージ、置換を順に組み合わせます。
台帳へマージするときの3方向の行き先、入力側の列が変わったときのアプリ所有列、検索列と完全一致／部分一致も JSON で指定します。

アプリはローカルだけで動き、実行時にネットワークへ接続しません。監視対象は UI Automation だけで読み取り、対象アプリへ書き込みません。

## 現行フォルダ

```text
src/
  csharp/       製品コード
  config/       出荷設定と画面定義
  launcher/     .vbs / .cmd のヘッダーとブートストラップ
  samples/      検査用の見本設定 (販売 / 製造 / 施設予約 + 画面を変えた 2 変種)
build/          C# 製品の生成と受け入れ検査
docs/           現行 C# 版の設計・設定・UI 資料
archive/        退役物の保全。現行製品からは参照しない
```

現行製品はこの 4 つだけで成り立ちます。`archive/` の中身は [アーカイブ案内](archive/README.md) を見てください。

詳しい入口は次のとおりです。

- [現行ドキュメント一覧](docs/README.md)
- [アーキテクチャ](docs/architecture.md)
- [設定ファイル](docs/settings.md)
- [旧画面定義と UI 仕様（v2・履歴）](docs/ui-spec.md)
- [アーカイブ案内](archive/README.md)

## 開発時の確認

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File build\build_dist.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\compile_check.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_settings_contract.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_settings_geometry.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_samples.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_exit_guard.ps1
```

`build.bat` は C# 製品だけを作り、`archive/` には触れません。

## ライセンス

ソースコードは [CC0 1.0 Universal](LICENSE) です。
