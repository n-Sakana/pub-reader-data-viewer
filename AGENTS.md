# AGENTS.md - pub/reader-data-viewer

詳細な利用方法は [README.md](README.md) を参照する。

## 現在の製品

- 現役の Reader Data Viewer は **C# / WinForms 版 1 種類**。
- 製品ソースは `src/csharp/`、設定 (1 枚の `settings.json`: paths / search / watch / jobs / data / screen) は `src/config/`、起動部は `src/launcher/`。
- CSV の定義 (`data`: 表、キー、結合、台帳の列) と画面 (`screen`: 部品 7 種、値の取り方 3 種、判定規則 3 種) は JSON の名前だけで書き、式やコードは書かせない。
- `settings.json` は厳格に読む。無い・壊れている・知らないメンバー・範囲外・何も指していない名前は、既定で動かず起動しない (理由をダイアログとログへ)。組込み既定を持ち込まない。
- `build.bat` と `build/build_dist.ps1` は `dist/app-csharp/` だけを生成する。出荷するデータと画面は**脱色したサンプル** (表A/B/C、`SAMPLE-A-0000001` のような値、中立なラベル) で、業務の色を付けない。
- 現実的なダミー (販売 / 製造 / 施設予約 + 画面を変えた 2 変種) は `src/samples/<name>/settings.json` + `build/gen_samples.ps1` が `samples/` に生成する検査用の組で、`build/test_samples.ps1` が製品コードに、`build/test_ui_geometry.ps1 -Settings` が画面の健全性に、`work/ui-v2/live_scenario.ps1` が実配布物の利用動作に通す。配布物には入れない。
- 画面の定義を変える・増やすときは、出荷定義の忠実度 (`test_ui_geometry.ps1`) と見本 5 定義の健全性 (`-Settings`) を両方回す。レイアウトの不具合は見本定義のほうで先に出る。
- 退役物はすべて `archive/` の下に集約済み。現行製品へ混ぜず、通常ビルドから参照しない。
  - `archive/vba/` — 退役した実用 VBA 版。
  - `archive/comparisons/` — 1 対 1・一対多の方式比較。
  - `archive/benchmarks/` — 方式選定の凍結済み証拠で read-only。配下を変更しない。
  - `archive/showcase/` — VBA Pixel Bridge 展示。Reader 本体の変更に巻き込まない。
  - `archive/ui-prototypes/` — v2 より前の UI 案。
- 現行製品はリポジトリ直下の `src/`・`build/`・`docs/` の 3 つだけで成り立つ。ここへ退役物を戻さない。

## 動作契約

- ローカル Windows 上だけで動き、実行時にネットワーク接続しない。
- UI Automation で対象の入力欄を読み、対象アプリには書き込まない。
- CSV は `data` の定義どおりに結合する (出荷定義: 3 本、A-B は key1、B-C は key2)。
- 統合台帳の identity は spine 表のキー (出荷定義では key2)。作業状態 (処理済み列) は identity と内容列が全部一致するときだけ引き継ぐ。
- CSV は厳格に読む。列数違い・引用符付きの列・キーの空や幅違い・一意キーの重複は、ファイル名と行番号を挙げて拒否し、続行しない。黙ってずらさない。
- 判定 (OK/NG) は画面定義の規則で CSV の生データから決め、一致なしは未定義、列なしはエラーとして明示する。暗黙に OK にしない。
- 検索・結合・保存は worker スレッドで行い、UI スレッドを塞がない。
- 作業状態の保存中は次の保存と終了を拒否し、成功または失敗が確定した後に解除する。キューや終了時一括保存へ変えない。
- 失敗時に別方式へフォールバックしない。失敗理由を記録して表示する。
- メモ帳を起動・終了・kill しない。すでに存在する対象へ接続して読むだけにする。

## 読む場所

- `docs/README.md` — 現行ドキュメントの索引
- `docs/architecture.md` — 現行 C# 版の構成とデータ契約
- `docs/settings.md` — `settings.json` (paths / search / watch / jobs / data、厳格な読込み)
- `docs/ui-spec.md` — `settings.json` の `screen`、画面の振る舞い、レスポンシブ、検査
- `docs/ui-reference/v2.html` — UI の正本 HTML。単体でブラウザで開ける。ここが唯一の正本で、複製を作らない
- `archive/README.md` — 退役物の境界と所在

## 開発コマンド

```powershell
build.bat
powershell -NoProfile -ExecutionPolicy Bypass -File build\build_dist.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\compile_check.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_settings_contract.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_ui_geometry.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_settings_geometry.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_exit_guard.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File build\test_samples.ps1
```

## ガードレール

- `data/`、`data-100k/`、`dist/`、`samples/`、`work/` は生成物。コミットしない。
- `archive/` の履歴資料を現行設計の根拠にしない。変更依頼がない限り整理・修正しない。
- key1 は一対多なので、索引は必ず `key -> 行の集合` とし、単一行で上書きしない。
- xlsx は temp 書込みから安全に置換する。保存完了前に成功表示しない。
- UI の寸法は画面定義 (`src/config/settings.json` の `screen`) の値で、`docs/ui-ref-v2-geom.json` (正本 HTML の実測) と検査で突き合わせる。コードに座標表を持たない。
- 文字は GDI (`Rdv3Skin.Draw/Measure`) で描き測る。GDI+ の文字描画や `TextRenderer` に戻さない (横潰れ・DPI 換算の狂い)。モーダルの角丸は DWM に任せ、Region で切らない。
- C# は Windows 同梱の .NET Framework `csc` でコンパイルするため **C# 5** に限定する。verbatim string を使わない。非 ASCII を含む C# は `Rdv3Text.cs` だけ。
- 製品ソースの一覧は `build/sources.ps1` だけに持つ。
- 非 ASCII を含む `.ps1` は UTF-8 BOM を維持する。
- 「設定」ボタンの書戻しは `paths` / `search` / `watch` の 3 メンバーの範囲だけを差し替える (`Rdv3Config.Save`)。ファイル全体を生成し直す形に戻さない。
- `build.bat` は CRLF を維持し、`&`・`(`・`)` を含む配置パスでもダブルクリック実行できる引用を崩さない。
