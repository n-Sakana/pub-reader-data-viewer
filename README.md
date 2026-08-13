# Reader Data Viewer

読み取り → 三 CSV 統合 → 検索 → 閲覧表示を行うアプリを作るためのリポジトリです。

**このリポジトリの主役は、これから作る完成アプリです。**
`benchmarks/` に入っているものは、そのアプリをどの方式で実装するかを決めるための
検証資料であって、成果物そのものではありません。

## これから作るもの

Notepad / UI Automation による読み取り検知を入口に、

1. 1,000,000 行 × 10 列の CSV を 3 本 (A / B / C) 受け取る
2. A-B は番号 1、B-C は番号 2 で 1 対 1 結合する
3. 番号 1 で検索する
4. 結果を表示する

これを 3 通りの実装で比べながら、工程ごとに時間を計測できる形にします。

- VBA / Excel
- C# / WinForms
- C# + Excel

現時点ではまだ実装していません。このリポジトリには、下の検証資料だけが入っています。

## benchmarks/ — 方式選定のための検証

### [benchmarks/excel-background-bench](benchmarks/excel-background-bench)

Excel から 1,000,000 件の郵便番号を住所へ変換する、起動動作テストと速度ベンチマークです。
同じ入力・同じ変換規則・同じ出力・同じ正解確認・同じ計測範囲のまま、**処理をどこで走らせ、
どうやって起動し、結果をどうやって Excel に戻すか**だけを変えて 14 方式を比較しています。

上のアプリでも「重い処理をどこで走らせ、結果をどうやって画面へ戻すか」が同じ問題になるため、
先にこちらで測りました。分かったことを 3 行でまとめると:

- **起動手段はほとんど効きません。** BAT / WScript.Shell.Run / VBA Shell /
  タスク スケジューラ COM の差は E2E で 0.03 秒以内でした。
- **効くのは結果の戻し方です。** 同じ 1,000,000 件を同じ規則で変換して同じシートへ入れるのに、
  経路だけで **0.77 秒から 10.95 秒まで 14.2 倍**開きます。変換そのものはどの方式でも
  0.05 秒前後で、順位に何も寄与しません。
- **最速は Excel 自身のプロセスで走らせる形 (Excel-DNA / XLL) で 0.773 秒。**
  プロセス境界を越えないぶん、書き戻しだけでなく読み込みの時点ですでに COM 経由の 4 倍速い。

詳しい数字と、そこから何が読み取れるかは
[benchmarks/excel-background-bench/docs/results.md](benchmarks/excel-background-bench/docs/results.md)
にあります。ビルド方法と再現手順は
[benchmarks/excel-background-bench/README.md](benchmarks/excel-background-bench/README.md)
を見てください。

## 動かすのに要るもの

`benchmarks/excel-background-bench` を動かす場合:

- Windows 10 / 11
- Excel (デスクトップ版, VBA が使えるもの)
- .NET Framework 4.x の `csc.exe` (Windows に同梱)
- PowerShell 7 (`pwsh`) または Windows PowerShell 5.1

## リポジトリに入っていないもの

再生成できるもの、配布物、他者の著作物は追跡していません。手順で作り直せます。

| 入っていないもの | 作り方 |
|---|---|
| 日本郵便の郵便番号データ (`KEN_ALL.CSV`) | `build/build_workbooks.ps1` が公式サイトから取得 |
| 1,000,000 件の入力・出力・辞書 | `ZipBench.xlsm` の [準備] ボタン、または `build/run_bench.ps1` |
| `ZipBench.xlsm` / `ZipData.xlsx` | `build/build_workbooks.ps1` |
| `ZipWorker.exe` | `worker/build_worker.ps1` |
| Excel-DNA と `ZbXll64.xll` | `xll/fetch_exceldna.ps1` → `xll/build_xll.ps1` |
| `vba/modZipEmit.bas` | `build/build_workbooks.ps1` が毎回生成 |

日本郵便の郵便番号データと Excel-DNA は本リポジトリに含まれていません。
それぞれの配布元の定めるところに従ってください。

## ライセンス

MIT License。[LICENSE](LICENSE) を参照してください。
このリポジトリで書かれたコードにのみ適用されます。ビルド時に取得する第三者の成果物
(日本郵便のデータ、Excel-DNA) は含まれず、それぞれの条件に従います。
