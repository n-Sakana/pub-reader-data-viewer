# 候補窓を再び開く検査

Windows PowerShell 5.1、WebView2 Runtime、Node を使う実窓検査。架空サンプルのコピーを作り、同じ会員番号を持つ 2 行を用意する。既存台帳・設定は変更しない。

`powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/Test-CandidateReopen.ps1 -Evidence C:/temp/rdv-candidate-new-run -NodePath C:/path/to/node.exe`

Evidence には未使用のディレクトリを指定する。Root を省略すると、この検査を置いた製品を使用する。Root を別の版へ向ければ、同じ検査で修正前後を比較できる。

検索→複数候補→選択→表示→クリアを 8 回行う。同じキーと別キー、検索ボタンと Enter、候補のダブルクリック・表示ボタン・Enter を含む。候補をクリックする前に、別の PowerShell 側で OS 窓の座標が実モニタ内に収まることを確認する。選択とクリアの後は WPF の UI Automation IsEnabled=true を確認する。最後に古い token のサイズ通知が窓を変えないことも検査する。

CDP の DOM 表示だけでは画面外の窓を直接操作できてしまう。Win32 IsWindowEnabled と UIA IsOffscreen も今回の画面外・WPF 無効状態を正しく区別しなかったため、実座標と親の UIA IsEnabled を組み合わせる。

`results.json` に PID・UTC 時刻・各窓の座標・親の有効状態・判定を保存する。試験が起動した Reader と子プロセスだけを PID ごとに終了する。これは機能検査であり、会社配布 PC の性能検収ではない。
