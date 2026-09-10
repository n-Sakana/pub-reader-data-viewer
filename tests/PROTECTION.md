# 台帳保護・復元の検査

Windows PowerShell 5.1で、開発用repoのルートから実行します。実データや本番台帳は使いません。実窓の検査は画面を使用するため、従来版と固定版を順番に実行します。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File src/ReaderDataViewer.ps1 -TestCore
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Test-ProtectionWorkflow.ps1 -Evidence C:\RDV-test\workflow
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Test-LedgerMigration.ps1 -Evidence C:\RDV-test\migration
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File tests/Test-ProtectionNative.ps1 -Evidence C:\RDV-test\native
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Test-Build.ps1
python tests/test_static.py
```

`-Evidence`には版・検査ごとに別の新しいフォルダを指定します。Native検査はNode.jsとPlaywrightも必要です。Nodeのパスが異なる場合は `-NodePath` を指定します。自動前面化の再検査が不要なら、`Test-ProtectionNative.ps1 -GuardsOnly` で復元UIと4種類の書込みガードを確認できます。

従来版の既存ブラウザー回帰は `python tests/test_web.py`、固定版の設定契約は `powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Test-SemifixedConfig.ps1` です。従来版の動的レイアウト専用テストを、固定版のレイアウト検査として扱いません。

## 確認すること

- 中核120件。保護全内容・未送信との差・削除の全行保管・復元衝突・定義正規化・旧形式読取/移行・共有ロック・保存失敗時の不変性・自動付与条件を含みます。
- 期間混在XLSX。100件から21件を送信、20件削除、入力差替え後81件、部分差替え後81件、1件復元で82件、全復元で101件。各段階の実ファイルを読み戻し、全内容と状態を比べます。
- 移行コマンド。確認指定なし・別定義・ロック競合・既存出力・元台帳と同じ出力先を拒否。同一定義の移行は元台帳を残し、100件の全内容と保存済み状態を別XLSXへ保ちます。
- 実窓。削除20件の全内容表示、1件復元、再表示19件、プロセス終了後のXLSX読戻し。実行中にJSONの保護値を変え、入力差分のない更新・送信・削除・復元を拒否してファイルのバイト不変を確認。表示文言だけの変更では送信が可能です。
- 自動読取。単一・複数・0件、最小化から通常/最大化への復帰、同値の連続読取抑止、空欄を挟んだ再読取、手動検索との区別。OS窓座標、親UIA状態、実前面ハンドル、Topmost、PrintWindow画像を記録します。
- `Test-ProtectionNative.ps1 -AutomaticOnly` は、決済済みだけ自動処理済みにする追加条件を実窓で確認します。未決済単一・混在候補未選択/キャンセル・未決済候補には付与せず、決済済み単一と選択した決済済み候補だけ付与します。手動変更を往復してから送信し、実XLSXの100行中2行だけTRUE、残る98行FALSEを読み戻します。
- 配布。設定と4入力、初期リセット用コピー、次期間と部分差替えの入力、移行スクリプト、説明書、ZIPと全ファイルのSHA-256を照合します。静的検査は運用用のroot設定ではなく、隔離した `tests/fixtures` の汎用見本を使用します。

## 前面化の条件と限界

検査用読取元は、OS許可のない通常条件と、前面の読取元自身が `AllowSetForegroundWindow` を呼ぶ条件を分けます。後者は成功経路の検査条件であり、製品に追加した強制前面化機能ではありません。通常条件で拒否されても検索を継続し、結果とログの一致を確認します。他アプリや操作者が途中で前面を変更した走行は、全体PASSとして扱いません。

会社PC・実カードリーダー・本番共有先・複数PC・他DPIは未確認です。所要時間を性能合格の根拠にはしません。試験終了後、検査が作ったPIDの終了記録を確認してください。
