> 元ZIPのREADMEの保全コピーです。現在の起動・検査手順はルートのREADME.mdを参照してください。以下に記載された未同梱スクリプトや過去のリポジトリ状況は今回検証していません。

# Reader Data Viewer

Reader Data Viewer は C# / WPF の窓に WebView2 を載せる Windows デスクトップアプリです。

`ReaderDataViewer.vbs` から黒いコンソールを出さずに起動し、承認済み v13 の HTML 画面を表示します。検索、候補選択、処理状態、台帳更新・削除、送信、CSV 出力、設定、UI Automation 監視は WebView2 の JSON ブリッジから既存の処理本体へ接続されています。画面の構成と寸法はルートの `settings.json` にある `screen` 定義から組み立てます。

通常起動は `ReaderDataViewer.vbs` をダブルクリックします。起動ログをコンソールで確認するときだけ `ReaderDataViewer.cmd` を使います。ビルドと検査は `build.bat` をダブルクリックします。いずれも管理者権限やインストールを必要としません。

## どこに何があるか

人が書くものだけがリポジトリ直下に出ます。

```text
ReaderDataViewer.vbs     コンソールを出さない通常入口
ReaderDataViewer.cmd     起動失敗をコンソールで確認する入口
build.bat                ビルド・検査・データ生成の入口
settings.json            データ、処理、画面の設定
src/                     C# ソース 21 本と ReaderDataViewer.ps1 (起動とコンパイル)
web/index.html            v13 モックから取り出した窓とダイアログの正本
web/app.js                screen 定義の描画と WebView2 操作ブリッジ
web/app.css               settings.json の寸法を受ける v13 補助スタイル
lib/                     WebView2 の再配布 DLL と表示文書
build/                   生成・検査・配布のスクリプト。build/samples/ はサンプル定義 5 種
archive/winforms/        フェーズ17完了時点で凍結した WinForms 版
```

## どこに何ができるか

生成物の置き場は **2 か所だけ**です。どちらも追跡していません。

```text
data/                        入力表 3 枚と delete.csv、expected.txt
                             アプリが読む唯一の実体。台帳・ロック・版・動作ログもここ
build/out/ReaderDataViewer/  配布物。丸ごとコピーして配る
build/out/samples/           サンプル 5 種
build/out/work/              検査の一時領域。検査が通ったらその場で消える
build/out/data-100k/         -OutDir を付けて生成したときの行き先
```

## 作り直す

`build.bat` がすべて作り直します。

```text
build.bat            build\out\ReaderDataViewer\ を組み立て、動かして検証する
build.bat test       検査を全部 (サンプル、設定の契約、終了時の守り、画面 DOM)
build.bat data       data\ と build\out\samples\ を生成する
build.bat compile    src\*.cs を起動時と同じ手順でコンパイルだけ行う
```

個別に走らせるときは `build\` の中を直接呼びます。

```text
powershell -File build\gen_data2.ps1        data\ の入力表 3 枚と delete.csv、expected.txt
powershell -File build\gen_data2.ps1 -Rows 100000 -OutDir data-100k   大きい方 (build\out\ の下へ)
powershell -File build\gen_samples.ps1      build\out\samples\ の 5 種 (定義は build\samples\)
powershell -File build\compile_check.ps1    コンパイルだけ
powershell -File build\build_dist.ps1       配布物の組み立てと検証
powershell -File build\test_samples.ps1            サンプル定義と製品の突き合わせ
powershell -File build\test_settings_contract.ps1  settings.json の契約
powershell -File build\test_exit_guard.ps1         共有台帳のロックと終了時の守り
powershell -File build\test_ui_dom.ps1             WebView2 の DOM を読む画面回帰
```

検査はいずれも窓を先生の画面に出しません。検査は 1 回ごとに `src/`・`lib/`・`web/`・`data/` を作業ディレクトリへ複製しますが、**通れば自分で消します**。失敗したときだけ残り、その場所を出力の最後に書きます。

台帳 (`ReaderDataViewer-Ledger.xlsx`)、そのロックと版、動作ログは `settings.json` の `paths` に従って `data/` の中へ書かれます。リポジトリ直下には出ません。

配布は `build\out\ReaderDataViewer\` を丸ごと配ります。`ReaderDataViewer.ps1` が起動のたびに `src\*.cs` をコンパイルするため、事前のビルド成果物はありません。

WinForms 版は [archive/winforms](archive/winforms/) に、内容を変えず一式で保全しています。現役版はそこからコードを参照しません。

## 履歴を後ろへ動かすとき

このリポジトリは複数の端末に置かれ、機械が定期的に commit して端末どうしで同期しています（`fin-berth …` という作者の commit がそれです）。そのため **片方の端末だけで `git reset` して履歴を後ろへ戻しても、その戻しは残りません。**

戻らない理由は commit 側ではなく取り込み側にあります。相手の端末はまだ古い（＝こちらより先の）先端を持っており、同期はそれを見て次のどちらかを行います。

- 戻したこちらの先端が相手の先端の祖先になる → **早送りで元の位置へ戻される**
- 双方が分岐している → **こちらの commit を相手の先端の上へ載せ替える**（rebase）

後者は同じ変更が別のハッシュで二重に生えるので、`git log` に同じ日時・同じ件名の commit が 2 つ並びます。2026-09-07 の時点で、このリポジトリの `main` と GitHub 側にその二重が 5 組ありました（`git log --cherry-pick --right-only` で片側固有の変更が 0 件であることを確認済み）。

**後ろへ動かす操作をするなら、先に相手の端末の先端を揃えること。**片側だけで完結する操作ではありません。二重が既にできてしまった場合は、履歴を書き換えず `git merge -s ours` で取り込めば、中身を変えずに枝を 1 本へ戻せます。

## ライセンス

本体は [CC0 1.0 Universal](LICENSE) です。同梱する WebView2 の表示は [第三者ソフトウェアの表示](THIRD-PARTY-NOTICES.md) を参照してください。
