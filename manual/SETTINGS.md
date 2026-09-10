# 設定リファレンス

本文の画面定義を含む完成例はレイアウトJSON設定版向けです。固定版で業務定義を変えるときは、同梱JSONの `screen.bindings` を保持し、必要な `data` の定義を変更します。表示対応は [業務設定の説明](PAYMENT-GUIDE.md#レイアウト固定版と従来版の違い) を参照してください。



CSV / XLSXを表として読み、結合・集計・抽出等から作った共有台帳を検索し、行ごとに確認状態を付けて送信するWindowsアプリです。業務の流れは設定済みの一般操作を組み合わせて表します。自由なプログラム実行や、あらゆる業務の自動化を約束するものではありません。

設定を作る人は、このREADME、コメント入りの`settings.json`、対象データ、実現したい操作の説明から始められます。ソースコードを読む必要はありません。**出荷設定を少し変える場合は、下の最小例へ入れ替えず、既存の表示項目・状態・ジョブを残して編集してください。** 最小例は新規作成の出発点です。

**設定の作成と検査は、書く人の仕事です。検査を「受け渡し先で今後実施」として終えないでください。** アプリが使える環境では、提出前に次の2本を自分で実行し、結果の行・値・未一致数まで業務フローと照合します。パスは実際の配置に直し、2本目の出力は未使用の名前にします。

```powershell
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File C:/app/src/ReaderDataViewer.ps1 -ValidateOnly -Config C:/trial/settings.json -DataDir C:/trial/data
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File C:/app/src/ReaderDataViewer.ps1 -RunUpdate -Config C:/trial/settings.json -DataDir C:/trial/data -Output C:/trial/result-01.json
```

失敗したら`%LOCALAPPDATA%/ReaderDataViewer/logs/feedback.log`（書けない場合は`%TEMP%/ReaderDataViewer/logs/feedback.log`）を読み、直して再実行します。`FAIL`の件数、`STOP`、`NOT CHECKED`を確認してください。[検査と結果JSONの読み方](#verify)、[ログの詳細](#feedback-log)へ続きます。アプリを実行できない環境なら、その事実と未検証範囲を明記し、「正しく動くことを確認済み」とは書きません。要件を表せないと判明した場合は、できない理由を返し、実運用用の設定ファイルを提出しません。

まず読む順は、[最小設定](#quick-start) → [入力とキー](#inputs) → [結合の単位](#join-grain) → [4表の完成例](#four-tables) → [実行検査](#verify)です。**表名はtablesのlabel、列と中間結果名はdata.labels、台帳は`data.labels.ledger`に名前が必要です。** 最小例の`"ledger":"台帳"`を削らないでください。`has no screen label`が出たら、[処理に使う名前と画面名](#process-labels)の記入先を確認します。

**要件を受けられるか、最初に確認してください。** 想定外のレコードは除外して続け、件数と元の行番号または識別値を警告します。停止するのはファイル・設定の構造が違う場合と、保存済み台帳が壊れている場合です。CSVの列不足行と空行もこの扱いです。CSV/XLSXの重複見出しは、その名前を設定から参照しない場合だけ全列を除外します。**参照する重複見出しは、片方の全セルが空でも拒否します。** 値のある方を自動で選ぶ機能ではありません。[入力の除外と拒否の境界](#input-shape)と[提供しない機能](#limits)を先に確認します。

**計算した差や積、集計した合計や件数は、作った名前のまま台帳へ保存できます。** `calculate`の`column`、`aggregate`の`as`、`select`の`as`で作った列を`<結果名>.<列名>`（例 `T.合計金額`）で`data.ledger.columns.source`に書きます。台帳の`identity`にも集計後のグループ列を使えるので、明細ファイルだけから伝票単位の台帳も作れます。[計算結果を台帳へ保存する](#calculation-storage)と[明細だけから伝票単位](#detail-only)を読んでください。

**同じ列構成のファイルを1つの台帳にまとめるのは`append`（縦に足す）です。** `join`は別の表の列を横に付ける操作で、左の表に無い行は増えません。[複数ファイルを縦に足す](#append-files)に完成例があります。帳票の表題が見出しより前にあるCSV/XLSXは`headerRow`で見出しの行番号を指定します。XLSXの日付セルは`data.types`でdateを宣言した列だけ、その書式の文字列へ変換して保存します。

<a id="quick-start"></a>

次は、UTF-8の`data/rows.csv`に`id,name`という見出しがある場合の、そのまま読める設定です。`id`は一意な文字列、`name`は表示する値です。先頭ゼロを保ちます。状態は「未確認 ↔ 確認済」、外部ウィンドウ監視は無効です。

```csv
id,name
001,例1
002,例2
```

```json
{
  "schema": 3,
  "paths": {
    "dataDir": "data",
    "ledger": "data/ReaderDataViewer-Ledger.xlsx",
    "log": "ReaderDataViewer.log"
  },
  "search": {
    "pattern": ".+",
    "candidateRowsShown": 100
  },
  "watch": {
    "targets": []
  },
  "data": {
    "encoding": "utf-8",
    "tables": {
      "A": {
        "file": "rows.csv",
        "key": "id",
        "keyValidation": {
          "characters": "unicode",
          "length": "variable"
        }
      }
    },
    "labels": {
      "A.id": "識別子",
      "A.name": "名称",
      "ledger": "台帳"
    },
    "jobs": [
      {
        "id": "update",
        "kind": "update",
        "inputs": [
          {
            "table": "A"
          }
        ],
        "steps": [
          {
            "operation": "merge",
            "target1": "A",
            "target2": "ledger",
            "keys": [
              "A.id",
              "A.id"
            ],
            "sourceOnly": "add",
            "both": "update",
            "targetOnly": "keep",
            "output": "ledger"
          }
        ]
      }
    ],
    "ledger": {
      "identity": "A.id",
      "search": {
        "columns": [
          "A.id"
        ],
        "match": "exact"
      },
      "columns": {
        "source": [
          "A.id",
          "A.name"
        ],
        "application": [
          {
            "name": "workState",
            "onSourceChange": "reset"
          }
        ]
      }
    }
  },
  "screen": {
    "card": {
      "startSize": [
        818,
        636
      ],
      "font": "Meiryo UI",
      "fontSize": 10,
      "gap": 8,
      "padding": [
        8
      ]
    },
    "workState": {
      "trigger": "manual",
      "store": {
        "column": "確認状態"
      },
      "states": [
        {
          "id": "todo",
          "text": "未確認",
          "stored": "FALSE"
        },
        {
          "id": "done",
          "text": "確認済",
          "look": "accent",
          "stored": "TRUE"
        }
      ],
      "initial": "todo",
      "transitions": [
        {
          "from": "todo",
          "to": "done"
        },
        {
          "from": "done",
          "to": "todo"
        }
      ]
    },
    "export": {
      "defaultFields": [
        "A.id",
        "A.name",
        "$work"
      ]
    },
    "candidates": {
      "columns": [
        {
          "header": "id",
          "value": {
            "field": "A.id"
          }
        },
        {
          "header": "name",
          "value": {
            "field": "A.name"
          }
        }
      ]
    },
    "sections": [
      {
        "type": "keyPanel",
        "title": "検索",
        "figure": {
          "label": "識別",
          "value": {
            "field": "A.id"
          }
        },
        "input": {
          "label": "検索値",
          "maxLength": 64
        },
        "buttons": [
          {
            "action": "search",
            "text": "検索"
          },
          {
            "action": "clear",
            "text": "クリア"
          },
          {
            "action": "workState"
          }
        ]
      },
      {
        "type": "fieldList",
        "title": "レコード",
        "rowHeight": 30,
        "rows": [
          {
            "label": "id",
            "value": {
              "field": "A.id",
              "empty": ""
            }
          },
          {
            "label": "name",
            "value": {
              "field": "A.name",
              "empty": ""
            }
          }
        ]
      },
      {
        "type": "sendBar",
        "value": {
          "state": "pendingCount"
        },
        "buttons": [
          {
            "action": "sendChanges",
            "text": "変更を送信"
          }
        ]
      },
      {
        "type": "statusBar",
        "segments": [
          {
            "value": {
              "state": "appState"
            }
          },
          {
            "value": {
              "state": "ledgerRows"
            }
          }
        ],
        "buttons": [
          {
            "action": "updateRecords",
            "text": "レコード更新",
            "job": "update"
          },
          {
            "action": "tableExport",
            "text": "テーブル出力"
          },
          {
            "action": "settings",
            "text": "設定"
          }
        ]
      }
    ]
  }
}
```


まず次の順に変更します。数値に見える識別子を勝手に数値型にしないでください。

1. ファイルごとに、見出し、文字コード、一意になる列の組合せ、空欄、同じキーで内容が異なる行の有無を確認する。全件を読んで確認し、数行だけから一意性を決めない。
2. `data.tables`の`file`、`key`、必要なら表別`encoding`を変える。複合キーは配列。異なる見出しの対応は`steps[].keys`で書く。
3. 台帳の1行を何で識別するか決め、`data.ledger.identity`と`source`、最後の`merge.keys`を揃える。検索する列は`data.ledger.search.columns`へ。
4. 結合・集計・抽出を`data.jobs[].steps`へ実行順に書く。**左に何を残したいか**を先に決める。全レコードを残して相手のないものも見るなら`left`結合。
5. `data.labels`を揃え、画面の`value.field`、候補列、出力列を保存済みの列へ合わせる。不要な項目は新規定義から省略し、使わない元設定の列参照を残さない。
6. 後述の三段の検査を行う。**exit 0はエラーなく実行した印です。意図どおりの行かどうかは結果の列・値・件数を期待値と比較してください。**

必要な入力・処理をこのアプリで表せない場合は、制約と未達の要求を明記し、使える設定として提出しないでください。無関係な表を黙って外す、壊れた必要データを除外する、根拠のない結合キーを仮定する方法で「成功」にしません。入力修正や業務上の識別規則が必要なら、その条件が満たされるまで実運用用JSONは返さず、できない理由を返します。

`//`と`/* ... */`のコメント、末尾カンマを使えます。設定ファイルはUTF-8で保存します。コメントを表すために`comment`等の未定義キーを足さないでください。真偽値は`true/false`、数値は引用符なし、Windowsパスの`\`はJSON文字列では二重にします。`C:/data`のような`/`も使用できます。

<a id="verify"></a>

## 三段で確かめる

展開済みのアプリフォルダーで実行します。これらの窓なし検査にも同梱の`src`と`lib`は必要です。設定を書くためにソースを読む必要はありません。

```powershell
# ① 設定、全入力表の見出し・キー・宣言型、各ジョブの接続を検査。台帳は作らない。
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File src/ReaderDataViewer.ps1 -ValidateOnly -Config C:/trial/settings.json -DataDir C:/trial/data

# ② 最初の更新ジョブを実際に実行。結果は新しいJSONへ書き出す。
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File src/ReaderDataViewer.ps1 -RunUpdate -Config C:/trial/settings.json -DataDir C:/trial/data -Output C:/trial/result-01.json

# 既存台帳に対するマージ・状態の戻りを調べるときだけ、試験用の台帳の控えを渡す。
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File src/ReaderDataViewer.ps1 -RunUpdate -Config C:/trial/settings.json -DataDir C:/trial/data -BaselineLedger C:/trial/before.xlsx -Output C:/trial/result-02.json
```

`-Config`省略時はアプリ直下の`settings.json`、`-DataDir`省略時はその設定の`paths.dataDir`です。コマンド引数として渡した相対パスは実行中の作業フォルダーが基準です。JSON内の相対パスの基準とは異なるため、委託の検査では絶対パスが確実です。

`-ValidateOnly`は入力を読み、列・一意性・型・手順の参照を確認します。実データの計算結果、既存台帳の状態、画面の収まりは確認しません。`-RunUpdate`は同じパーサーと処理本体で更新ジョブを実行します。`paths.ledger`や共有ロック、未送信の控えに書き込みません。`-BaselineLedger`を渡さなければ空の台帳から作り、戻した状態の件数は0です。渡した場合も、そのファイルは読み取りだけです。

結果JSONはBOMなしUTF-8で、次の形です。保存値は文字列のままで、例えば`"001"`を`1`にはしません。

| 出力 | 意味 |
|---|---|
| `columns`、`rows` | 列参照の配列、その順の値配列。出来た台帳の内容 |
| `states` | `rows`と同じ順の確認状態の保存値。表示名や真偽値ではない |
| `summary.rows` | 出来た台帳の行数。検査だけのときは`null` |
| `inputs[]` | 更新ジョブが読む各入力のID、ファイル、採用行数と各除外数 |
| `summary.skippedEmpty/skippedDuplicate` | 更新ジョブの各入力の除外数の合計。同じファイルを異なる定義で2回読むなら2回の読取りとして数える |
| `summary.skippedShort/skippedBlank/skippedColumns` | CSVの列不足行数、空行数、未参照の重複見出しに属する全列数。各入力にも同名の件数を出す。列数は行数と合算しない |
| `joins[]` | 各結合の左右の入力行数、出力行数、`unmatchedLeft`と`unmatchedRight`。出力名でどの結合か識別する |
| `summary.skippedInvalid` | 型・キー規則・XLSXセル・計算等で除外した行数。入力ごとの同名項目は読取り時の除外だけ。処理中は段ごとの除外を加えるため、同じ元行を別の分岐で除外すればそれぞれ数える。空キー・重複・列不足・空行は別項目で重ねて数えない |
| `summary.baselineRows` | 比較のために読み込んだ既存台帳の行数 |
| `summary.resetRows`、`resetRows` | 内容変更によって初期状態へ戻った行の件数と内容。取消という特定の値を自動判定した件数ではない |
| `values` | 操作の結果名ごとの種類、件数、列、行内容。抽出した行集合も答え合わせできる。同じ名前を再使用すると最後の値になる |
| `warnings` | 列不足行・空行・未参照の重複列・キー規則違反・内容が違う重複グループ全行・型違反・計算等の不成立による除外、`headerRow`で読み飛ばした行数、0件の入力等の説明。ファイルログと画面の入力検証・処理後の警告にも出す |

結合が複数あると、同じ元行が複数段で未一致になることがあります。`unmatchedLeft`を全段合計して「未一致のユニーク行数」としないでください。左に相手が複数ある場合は1対多で行が増えます。`extract`で選んだ件数は`values.<出力名>.count`で確認します。処理が0件でも、0件になった段階と除外・未一致の件数を見られます。

成功はexit 0、設定・入力・実行エラーはexit 3、未知のコマンド引数はexit 2です。理由は標準エラー、除外の通知と要約は標準出力へ出し、同じ内容を[ファイルログ](#feedback-log)にも残します。既存の出力・入力・台帳・設定等への上書きは拒否します。毎回新しい`-Output`を指定してください。

`-ValidateOnly`は同じ検査段階のエラーを集め、先頭に`FAIL 4 errors`のように件数を出します。JSON構文、設定定義、入力ファイル、列・型、ジョブ準備の順で進み、その先を安全に検査できない段階で止めます。`STOP`は止まった段階、`NOT CHECKED`は未検査の範囲です。出ているエラーをまとめて直してから再実行してください。後の段のエラーが次回初めて出る場合があります。`-RunUpdate`と窓の処理も、レコードの問題では止まらず、除外した行の警告を集めて続けます。構造の問題では実行を止めます。`-ValidateOnly`は計算・集計を実行せず、既存台帳も読まないため、実行時の除外・値・台帳の検証には`-RunUpdate`が必要です。

③は試験用フォルダーへアプリ一式と設定・入力を置き、`ReaderDataViewer.vbs`を開きます。**通常起動は常にアプリ直下の`settings.json`です。** 作成・更新の確認を経て、検索値を入れ、候補と表示列、状態の変更と送信、CSV出力を操作します。窓なし検査だけでは、実際のWPF/WebView2、IME、DPI、監視対象アプリとの相性は確認できません。

<a id="paths"></a>

## 配置と共有の約束

`paths.dataDir`、`paths.ledger`、`paths.log`の相対パスは**アプリ一式のフォルダー**が基準です。表とジョブ入力の`file`は**dataDir**が基準です。絶対パスも使えます。

```json
"paths": {
  "dataDir": "data",
  "ledger": "\\\\server\\share\\ReaderDataViewer-Ledger.xlsx",
  "log": "ReaderDataViewer.log"
}
```

各PCにアプリを置き、台帳だけを全PCから同じ共有ファイルへ向け、ログはPC別にします。実行ログ`paths.log`はアプリ一式のフォルダー直下に置き、入力フォルダー`data`の中には置きません（出荷時の設定も`ReaderDataViewer.log`）。OneDrive等の同期コピーを同じ1本のファイルとして扱う運用は対象外です。台帳をExcel等で同時編集しないでください。パス比較は通常の絶対パス比較であり、共有先の別名・シンボリックリンク・ハードリンクの完全な同一性判定ではありません。

出荷設定は`watch.targets: []`で、監視は無効です。「メモ帳」という表示名でデスクトップを読む見本は入れていません。監視する場合は設定画面で実際の対象を選んで追加します。

起動時の入力除外はファイルログと画面の入力検証で確認できます。除外警告だけで毎回OKを求めるダイアログは出しません。手動の更新・削除では、除外した行の警告を表示します。送信後も、それまで表示していた識別値のレコードを最新台帳から探して表示し続けます。共有側でその行自体が削除されていた場合は表示を消します。

台帳は初回には同梱されません。起動後の作成確認は正常です。読めない既存台帳を無条件に作り直す動作はしません。

<a id="operation-log"></a>

### 端末別の操作ログ

共有台帳を置き換えるたびに、台帳と同じフォルダーの端末別CSVへ1行追記します。ファイル名は`<台帳名>-操作ログ-<端末名>.csv`（例 `ReaderDataViewer-Ledger-操作ログ-PC01.csv`）です。端末ごとに別ファイルなので、複数のPCと利用者が同時に使っても互いの記録を壊しません。UTF-8（BOM付き）、CRLF、列は`操作日時,端末名,ユーザー名,操作,台帳の行数,内容`で、Excelでそのまま開けます。

| 操作 | いつ追記するか | 内容の例 |
|---|---|---|
| 作成 | 台帳が無い状態で更新ジョブが台帳を新しく作ったとき | `更新: 追加 990 件、更新 0 件、削除 0 件、未確認に戻した 0 件` |
| 更新 | 更新ジョブ（merge / replace 等）が共有台帳を置き換えたとき | 同上。先頭はdata.jobsのname、戻した件数の状態名はscreen.workStateのinitialのtext |
| 削除 | 削除ジョブが共有台帳を置き換えたとき | `削除: 削除 4 件` |
| 送信 | 送信で確認状態を共有台帳へ書いたとき | `確認済 12 件、未確認 3 件`（状態名はscreen.workState.statesのtext） |

追記は共有台帳を書き換えた直後、同じ共有ロックの中で行います。差分が無く台帳を書き換えなかった更新・送信は記録しません。`-ValidateOnly`と`-RunUpdate`は共有台帳を書かないので記録しません。共有フォルダーへ追記できなかったとき（Excelがそのファイルを開いている、共有先が一時的に見えない等）は台帳の書き換えを取り消さず、その行を`%LOCALAPPDATA%/ReaderDataViewer/`の控えへ残し、次の起動・更新・削除・送信のときに先頭から再送します。控えへ回った事実は`paths.log`と[ファイルログ](#feedback-log)に`oplog`として残ります。このCSVは切り替えも自動削除もしません。入力ファイル・ログ・テーブル出力にこの名前のファイルは指定できません。

<a id="inputs"></a>

## 入力、文字コード、キー、型

`data.tables`はIDごとに`file`と`key`を指定します。`label`を省略するとIDが表名になります。CSVはカンマ区切り、引用符内のカンマ・改行、二重引用符、CRLF / LF / CRに対応します。BOMのあるUTF-8も、ないUTF-8も`utf-8`で読めます。

```json
"encoding": "utf-8",
"tables": {
  "A": { "file": "A.csv", "encoding": "shift_jis", "key": "id" },
  "B": { "file": "B.csv", "key": ["number", "item"],
         "keyValidation": { "characters": "unicode", "length": "variable" } },
  "C": { "file": "C.csv", "encoding": "utf-8", "key": "id" }
}
```

`data.encoding`の省略値は`utf-8`。**`data.tables.<ID>.encoding`があればその表だけ上書きし、なければ共通設定を使います。** `shift_jis`（WindowsのCP932）、`utf-8`、`utf-16`（UTF-16LE）、`utf-16BE`等を指定できます。UTF-16はBOMなしも、指定すれば読めます。BOMと指定が違えば、実際と期待した文字コード、変更すべき設定パスを示して拒否します。別の文字コードを自動選択して読み進めることはしません。XLSXの文字コードはこの設定の対象外です。

読めないバイトがあり、BOMなし・片側のバイト位置に0が偏る・UTF-16として厳格に読める・区切りと改行を含む文字列になる、という条件を満たすと、エラーに`Possible UTF-16LE`または`Possible UTF-16BE`と、変更するencodingのキー・値を添えます。**候補の案内であり、確定した自動判定ではありません。** 元の出力設定を確認して明示的に直し、再実行してください。0の偏りが乏しいファイルや他の文字コードには候補を出せない場合があります。

判断材料を見るには、PowerShellで`Format-Hex -Path C:/input/table.csv | Select-Object -First 4`を実行できます。`FF FE`はUTF-16LE、`FE FF`はUTF-16BE、`EF BB BF`はUTF-8のBOMです。BOMなしでもASCIIの区切り等が`2C 00`・`0D 00 0A 00`ならLE、`00 2C`・`00 0D 00 0A`ならBEの手掛かりになります。単一のバイトだけで確定せず、正しい見出しが読めることも確認します。UTF-8とShift_JISを理由なく交互に指定し続けないでください。

見出しより前に表題や出力日の行がある帳票CSVは、`data.tables.<ID>.headerRow`に見出しの行番号を書きます。それより前の行は読み飛ばし、飛ばした行数を`warnings`に出します。飛ばした行は列不足行・空行には数えません。XLSXも同じ指定で見出しより上の行を飛ばします。指定が無ければ1行目（XLSXは最初の行）が見出しです。ジョブの外部入力（`file`/`column`形式）にも同じ`headerRow`を書けます。

```json
"R": { "file": "拠点別売上.csv", "key": "拠点コード", "headerRow": 3 }
```

区切りがカンマでないファイルは`delimiter`で区切り文字を指定します。Excelの「Unicode テキスト (*.txt)」はUTF-16のタブ区切りなので`"encoding": "utf-16"`と`"delimiter": "tab"`です。`"semicolon"`、`"pipe"`、または1文字（`";"`等）も書けます。省略はカンマです。引用符の扱いは同じで、外部入力にも書けます。指定の無いタブ区切りのファイルは「見出しにタブ文字があり、指定した区切り文字がありません」で止まります。

```json
"S": { "file": "出荷予定.txt", "encoding": "utf-16", "key": "出荷番号", "delimiter": "tab" }
```

<a id="process-labels"></a>

### 処理に使う名前と画面名

同じ意味の列が`A.id`と`B.number`という異なる名前でも、結合・抽出の`keys`に両方を書けば対応づけられます。CSVを加工して見出しを揃える必要はありません。`data.labels`は人向けの別名であり、元CSVの列名を書き換えるものではありません。

**表そのものの名前は`data.tables.B.label`に書きます。`data.labels.B`は書けません。** `label`を省略しても表IDの`B`が登録されるため、同じ名前を`data.labels`で再定義できません。同じ表示名を両方へ書いた場合もエラーです。列の名前`B.id`、新しい中間結果名`joined_AB`、台帳名`ledger`は`data.labels`に書きます。

**`ledger`は組込みの行き先ですが、画面向けの名前は自動登録されません。** `data.labels`へ`"ledger":"台帳"`を入れてください。`ledger has no screen label`が出たときの修正箇所もここです。`tables.ledger`を作る方法ではありません。

```json
"tables": {"B": {"label":"表B", "file":"B.csv", "key":"id"}},
"labels": {"B.id":"識別子", "joined_AB":"結合結果", "ledger":"台帳"}
```

これは`data`内の部分例です。`B already has a table label`が出たら`data.labels.B`を削除し、表名の変更は`data.tables.B.label`へ移してください。`B.id`等の列ラベルは残します。

**手順の`output`に名前を書くだけでは、画面名は登録されません。** 結合後の表、抽出した行集合など、ファイルへ保存しない中間結果も「処理内容」の対象・出力として表示するので名前が必要です。

| 処理で使う名前の例 | 画面名の記入先 |
|---|---|
| 入力表`B` | `data.tables.B.label`（省略時は`B`。明示するなら空にしない） |
| 結合の`output: "joined_AB"` | `data.labels`の`"joined_AB": "結合結果"` |
| 抽出の`output: "selected"` | `data.labels`の`"selected": "選択した行"` |
| 台帳`ledger` | `data.labels`の`"ledger": "台帳"` |
| キーや条件に使う列`B.id` | `data.labels`の`"B.id": "識別子"` |
| `output: "totals"`へ作る集計列`as: "amount"` | `data.labels`の`"totals.amount": "合計値"`。結果名`totals`にも別途画面名が必要 |

`joined_AB has no screen label`が手順の`output`と次の`target1`の2か所に出ても、追加するのは`data.labels`内の`"joined_AB"`1件です。エラーに示した名前をキーとしてそのまま使い、値には用途に合う空でない表示名を書きます。`"B.id"`や`"totals.amount"`は、ドットを含む1個のキーです。`"B": {"id": ...}`のように入れ子にしません。

既存の`labels`オブジェクトへ追記し、他の名前を残してください。別の`labels`オブジェクトを作ったり、中間結果を`data.tables`へ登録したりする必要はありません。エラー位置は「その名前を使った場所」、修正先は上表の場所です。名前の綴りが誤っている場合は参照を正し、修正後はもう一度`-ValidateOnly`と`-RunUpdate`を実行します。

| 入力の揺れ | 実際の扱い |
|---|---|
| 見出しの前後の空白 | 取り除いて列名を照合 |
| 値の末尾の半角空白・全角空白・NBSP | 取り除く。値の先頭の空白や途中の改行は一律には消さない |
| 全角数字 | 読んだ値と検索入力の数字を半角化。全角英字等を一括変換する機能ではない |
| `¥66,131` / `￥66,131` | 数値として66131。CSV内の桁区切りカンマは引用符で囲む：`"¥66,131"` |
| `(500)` / 全角の数字・括弧・小数点等 | 数値として-500等を読む。`(-500)`のような二重の符号は拒否 |
| 型宣言のある列の空セル | 空のまま型検証を続行。**計算・sumに渡すと0を仮定せず、その行を除外する場合がある**。空欄を集計に含めるかは手順で決める |
| キー空の行 | 除外し、ファイル・キー列・元の行番号・除外件数・採用件数を警告。旧`empty:error`も同じ。複合キーはどれか1列が空なら除外 |
| 完全に同じ行の再送 | 正規化後に全セルが一致する同一キーの行は1行を採用し、除外件数を通知 |
| 同じキーで内容が違う行 | 既定ではそのキーの全行を除外。キー・件数・全行の元の行番号を警告。先に同一内容が続いてから違う内容が来た場合も全行が対象。枝番を足すか値を直すかを入力の管理者が決める |
| 見出しより前の表題行・出力日の行 | `headerRow`に見出しの行番号を書いて読み飛ばす。指定が無ければ1行目を見出しとして読み、列数が合わない行で止まる |
| タブ区切り・セミコロン区切り | `delimiter`で指定する。指定が無いタブ区切りは見出しのタブを検出して止まる |

キー検証の省略値は`characters:ascii`、`length:fixed`、`duplicates:error`、`empty:skip`です。`error`という綴りは既存JSONとの互換のため残していますが、現在はレコードの問題でファイル全体を止めません。ASCII・固定長・制御文字の違反行は除外し、`duplicates:error`は内容が違う同一キーの全行を除外します。新しい表が可変長や日本語の識別子を持つ場合は、最小例のように`characters:unicode`、`length:variable`を指定します。複合キーでは**列ごとに**文字種と長さを検証し、**組合せ全体で**一意性を検証します。単純な文字連結ではなく組合せを保持するため、`["AB","C"]`と`["A","BC"]`は別のキーです。

`duplicates:distinct`は**キーが同じなら最初の1行だけ採る**指定です。値が違う後続行も除外します。**合計したい行にdistinctを指定してはいけません。** 明細等を保って集計するには、行を区別できる単一キーまたは複合キーを宣言し、`aggregate`でグループ化します。一意になる組合せがないのに、全行を異なる明細として取り込む指定はありません。

重複キーの除外警告が出たら、複数の明細が正常なデータなのかを先に確認します。正常なら、明細を区別する連番・日付等も含めて入力のkeyを決め、**全件を保持して読んだあと**に`aggregate/groupBy`で照合したい単位へ集めます。入力のキー検査は集計より先です。「集計するので重複を許す」という省略指定はなく、distinctでは合算されません。[入力行と結合の単位](#join-grain)も参照してください。

<a id="input-shape"></a>

**CSVの列不足行と空行は除外し、件数と元の行番号を警告します。** 見出しが`id,name`なら`001`の行は列不足で除外、`001,`は2列なので空のnameとして採用します。途中の区切りが失われた行も列数が少なければ除外します。セルの補完や列ずれの修復はしません。引用符付きCSV、複数行セル、UTF-16でも同じです。引用符内の改行を空行として数えません。列不足の判定は重複列の除外より前の、元の見出し列数で行います。

**空の見出し、参照する重複見出し、不正な引用符、CSVの見出しより多い列、不正な文字コードのバイト列は、構造を読み取れないため停止します。** キーの制御文字は、その行を除外し、列名と制御文字を`\u0009`等で表した実際の値を警告します。 非キーのタブ等の一部制御文字は警告して`?`へ置換し、引用符内のCR/LFは保持します。意味が壊れる入力を、別の正常値へ推測して読み替えません。XLSXの省略セルはセル番地で位置が決まるため、従来どおり空セルとして読みます。CSVの列不足行とは扱いが異なります。XLSXのデータ行にエラーセル、計算結果が未保存の数式、見出しの範囲外のセルがある場合は、その行を除外し、ファイル・元の行番号・セル番地・値を警告します。見出しのエラー、壊れた共有文字列参照、ワークブック構造の不正は停止します。保存済み台帳では行を勝手に除外しません。

**CSV/XLSXの重複見出しは、その名前を設定から参照しない場合だけ、同名の全列を除外して件数と名前を警告します。** `id,name,unused,unused`でidとnameだけ使うなら、unusedの2列を除外して読めます。前後の空白を除いた`unused`と`unused `も同名です。判定対象は入力キー、型宣言（textも含む）、全ジョブの列・式・条件、台帳の保存列です。表示名だけを`data.labels`へ登録したことや、式の文字列定数はデータ参照に数えません。

**参照する名前が重複していたら、どちらの列か決まらないため拒否します。** `id,value,value`でvalueを計算や保存に使う場合は、片方の値が全行空でも読み分けません。元データの管理者へ`value1/value2`等への改名を依頼し、設定参照も合わせてください。列番号の指定はなく、`data.labels`や読取り後の`select`では解消できません。**元データを変更できず、この重複名の値が要件に必要なら、設定だけではできません。** 必要な表や列を外したJSONを完成品として返さないでください。

型は列参照に対して宣言します。省略列は文字列で、数字に見えても数値・日付へ自動判定しません。

```json
"types": {
  "A.compactDate": { "type": "date", "format": "yyyyMMdd" },
  "B.isoDate":     { "type": "date", "format": "yyyy-MM-dd" },
  "C.slashDate":   { "type": "date", "format": "yyyy/MM/dd" },
  "D.localDate":   { "type": "date", "format": "yyyy年M月d日" },
  "B.amount":     { "type": "number" },
  "B.number":     { "type": "text" }
}
```

これらの日付書式を列ごとに指定できます。`yyyy年M月d日`なら`2026年9月9日`や`2026年12月31日`を日付として扱えます。`M`と`d`は1桁の月・日も許す指定です。固定2桁に限定する場合は`MM`と`dd`を使います。1列に指定できる書式は1つです。**同じ列に複数の日付書式が混在する場合の自動判別はありません。** `data.types`は検証と数値・日付の絞り込みのための宣言です。見せる形だけを変える`screen`の`value.format`とは別です。表示を変える例は`{"field":"D.localDate","format":{"kind":"date","from":"yyyy年M月d日","to":"yyyy-MM-dd"}}`です。

XLSX入力は、ブックで最初に列挙されたワークシートを読みます。任意のシート名を指定する機能はありません。数式は保存済みの計算値を読み、未計算・エラーセルは拒否します。再計算やExcelの見た目の再現はしません。**日付セルはExcelの内部値（シリアル値、例 `46246`）で読みます。その列を`data.types`で`date`と`format`で宣言すると、シリアル値を宣言した書式の文字列（例 `2026/08/12`）へ変換して台帳に保存します。** 宣言しない列は数字のまま保存され、台帳は文字列で書くのでExcelの書式設定でも日付には戻りません。日付として台帳に持つ列は必ず宣言してください。文字列の日付セルは宣言した書式に合っていればそのまま読みます。時刻の部分は捨てます。ブックが1904年日付方式（`workbookPr date1904`）で保存されていれば、その方式で同じ暦日に変換します。変換した列がその表のキーなら、変換後の値で文字種と桁数のキー規則を確かめ直します（`yyyy/M/d`のように桁数が変わる書式は`keyValidation.length`を`variable`に、漢字を含む書式は`characters`を`unicode`にします）。失われた先頭ゼロを入力時に自動復元はしません。識別規則が確定している場合の式による対応は[先頭ゼロの節](#leading-zero)を参照してください。

<a id="jobs"></a>

## 処理をつなぐ

[4表の完成例](#four-tables)では、CSV4本と設定全文、集計・左結合・取消フラグによる状態リセットまでの操作と期待値を一緒に確認できます。

`data.jobs[]`の`id`は全ジョブで一意、`name`は省略するとid、`kind`は`update`か`delete`です。**先頭のupdateジョブが起動時の確認と`-RunUpdate`に使われ、最後の結果はledgerでなければなりません。** 画面から実行するジョブはボタンの`job`で結びます。利用者にジョブを選ばせる画面はありません。

`inputs`には登録表を`{"table":"A"}`で指定できます。この形には他のキーを足しません。外部ファイルを1列の条件値一覧として読む場合は次の形です。省略時は`data.encoding`、明記すればその入力だけ上書きします。

```json
{"id":"D", "label":"条件値", "file":"values.csv", "encoding":"shift_jis",
 "column":"元の列名", "key":"A.id",
 "keyValidation":{"characters":"unicode","length":"variable","duplicates":"distinct","empty":"skip"}}
```

これは元の`column`を論理的な`A.id`という列として比較する定義です。任意の表を読む場合は`data.tables`を使います。同じCSVを2列の値一覧として2回読む方法は、**それぞれの列に値が含まれるか**を調べる方法で、**同じ元行で2列が一致する複合キー照合とは違います**。組合せを守る照合は表と複合`keys`を使ってください。

段の`target1/target2`は入力ID、前段の`output`、または`ledger`です。`output`は結果につける名前で、ファイル名ではありません。新しい結果名と列参照は`data.labels`に画面向け名を付けます。存在する名前を上書きできるのは`output`が自分の`target1`と同じ場合、または`ledger`の場合です。

値には「表」「その表から選んだ行集合」「台帳」の3種類があります。`extract`の結果は行集合なので、**そのままmergeへ渡せません**。元の表へ`delete`または`update`で適用します。行集合を作った後に表を別の結果へ更新したら、改めてその表から選んでください。

| operation | 入力 → 出力 | 操作固有の指定 |
|---|---|---|
| `join` | 表と表 → 表 | `keys`は左右の対応列。`condition:match`は一致だけ、`left`は左の全行、`full`は両方の全行を残す。相手なしの列は空文字。1対多・多対多では行が増える |
| `extract` | 表と条件 → 行集合 | `where`で列の条件、または相手表と`keys`で`match/exclude`。両形式は同時指定しない |
| `extract` | 同じ表の行集合2つ → 行集合 | `condition:both`はAND、`either`はOR、`exclude`は左から右を除外。keys/whereは書かない |
| `delete` | 表とその行集合 → 表 | 選んだ行を除く。台帳を主対象にすれば台帳。conditionは省略 |
| `update` | 表とその行集合 → 表 | `set:[{"column":"A.amount","expression":"A.amount + 1"}]`。複数列可。式は変更前の行の値を使う。アプリ所有の状態列を直接指定する操作ではない |
| `append` | 表と表 → 表 | 縦につなぐ。列数と末尾の列名が同じ順に揃うこと。出力の列参照は左の表を使う |
| `select` | 表 → 表 | `columns:["A.id",{"column":"B.name","as":"name"}]`。順序・採用列・別名を決める。別名を付けると`output.name`になる |
| `calculate` | 表 → 表 | `column:"newValue",expression:"A.amount * 2"`。作った列は`output.newValue`。既存列と同名は不可 |
| `aggregate` | 表 → 表 | `groupBy:["B.id"]`と`aggregates`。`sum`と`count`のみ。`groupBy:[]`は全体を1組にする |
| `sort` | 表 → 表 | `orders:[{"column":"A.amount","direction":"descending","type":"number"}]`。direction省略ascending、type省略text。複数指定で優先順 |
| `distinct` | 表 → 表 | `columns:["A.id","A.part"]`の組合せが同じ行の先頭だけを残す。合算しない |
| `merge` | 表とledger → ledger | `keys`、`sourceOnly:add/ignore`、`both:update/keep`、`targetOnly:keep/delete`。後者の省略はkeep |
| `replace` | 表とledger → ledger | `keys`を指定。入力にない行まで除き、入力の順序に置換。初期化だけを目的に気軽に選ばない |

**`join`と`append`の使い分け。** 別の表の列を横に付ける（受注に得意先名を付ける）のは`join`で、左の表の行数は減りません（`left`）が、相手にしか無い行は増えません。同じ列構成のファイルを縦に足す（4月分と5月分をまとめる）のは`append`です。月次ファイルを`join`でまとめると片方にしか無い受注が落ちるので、[複数ファイルを縦に足す](#append-files)の形にしてください。

`where.operator`は`equals/notEquals/contains/startsWith/endsWith/empty/notEmpty/greater/atLeast/less/atMost`。前6種等の文字比較は値をそのまま比較し、大小4種は数値です。`empty/notEmpty`には`value`を書かず、それ以外には文字列の`value`を書きます。空欄と0は同じではありません。

`expression`は列参照、数値、単一引用符の文字列、括弧、`+ - * /`、`substring(列,開始,長さ)`、`splitPart(列,'区切り',位置)`、`regexExtract(列,'正規表現')`です。開始位置は0。`'it''s'`のように単一引用符を重ねて文字自体を表します。`+`は両辺が数値なら加算、どちらかが数値でなければ文字列連結です。例えば`A.id + '-' + A.part`は区切り付きの文字列になりますが、数字だけの`A.id + A.part`は加算され得ます。行の識別には連結の代わりに複合キーを使ってください。

任意のC#/SQL、IFやCASE、`concat`等の未定義関数は使えません。上の3つの文字列抽出関数は、空入力・空の抽出結果・不一致・範囲外の行を除外し、関数・入力値・条件・元の行を警告します。`splitPart`の位置も0始まりなので、`p*q`の`q`は位置1です。空欄を別の値へ変える、コードを名前に置き換えるなど条件つきの置き換えは、`extract`で選んで`update.set.expression`に定数を書く形で表せます（[完成例](#conditional-replace)）。

**計算・集計・数値比較の不成立。** `- * /`の数値でない値、0除算、桁あふれ、`sum`の数値でない値、数値並べ替えの不正値は、処理対象からその行を除外します。集計は行単位で確定するため、1セルが不正ならその行を他のsumやcountにも加えません。数値条件に適合しない値は選択されません。元の表を別の段で使う場合、選択から外れたこと自体は元の表の削除を意味しません。

`update`で複数列を指定した場合は、1行の代入がすべて成立してから反映します。台帳への更新で成立しない行は、元の値と状態を保持して更新対象から除きます。保存済み台帳自体に数値として並べ替えられない値がある場合は、行を消さず停止します。

`data.types`をtextへ変えるのは、その値を数値・日付として扱わないときの対応です。数値の差や合計が必要な値をtextへ変えただけでは計算できません。単位付きの値は、形式が決まっていればregexExtract等で必要な部分を計算列へ取り出し、その列を計算・集計に使えます。形式や単位の意味が混在していて規則を決められない場合は、未対応のまま報告してください。

<a id="calculation-storage"></a>

### 計算結果を台帳へ保存する

`calculate`の`column`、`aggregate`の`as`、`select`の`as`で作った列は、後段の計算・結合・抽出に使えるだけでなく、**その名前のまま`data.ledger.columns.source`に書いて台帳へ保存できます。** 参照は`<結果名>.<列名>`です。結果名は手順の`output`に書いた名前で、入力表のIDへ上書きした（`output:"A"`）ならその表IDです。`data.labels`には結果名と各列の表示名を書きます。

例えば表Aの入力見出しが`id,name,quantity,completed,unitValue`で、差と積を保存する更新ジョブの手順は次の3段です。

```json
[
  {"operation":"calculate","target1":"A","column":"diff","expression":"A.quantity - A.completed","output":"A"},
  {"operation":"calculate","target1":"A","column":"product","expression":"A.quantity * A.unitValue","output":"A"},
  {"operation":"merge","target1":"A","target2":"ledger","keys":["A.id","A.id"],"sourceOnly":"add","both":"update","targetOnly":"keep","output":"ledger"}
]
```

`ledger.columns.source`は`["A.id","A.name","A.diff","A.product"]`、`data.labels`に`"A.diff":"数量差"`、`"A.product":"積"`を足します。計算元の3列は`number`を宣言します。`A.diff`や`A.product`にも`number`を宣言すると、出力の絞り込みで数値として扱えます（作った列の型は入力の検査対象ではありません）。quantity=10、completed=4、unitValue=25なら保存結果は6と250です。RunUpdateの`columns`と`rows`で答え合わせします。台帳（Excel）の見出しは列名の部分（`diff`、`product`）です。

**保存する列は、更新ジョブの最後の書込み段（merge / replace）の入力に無ければなりません。** 結合していない表の列や、`select`で外した列、綴りの違う名前を`source`に書くと、`台帳の保存列 … は、更新ジョブの最後の書込み段 (…) の入力にありません`で止まります。無い列を空欄で通すことはしません。`calculate`は同名の列を上書きしないので、既存の見出しと同じ名前へ置くときは先に`select`で外します。従来どおり、元の値を使わない既存列へ`select`と同名の`calculate`で置く書き方も動きます。集計結果の保存は[複数の行を単位ごとに合計する](#aggregate)、集計後のキーを台帳の識別に使う例は[明細だけから伝票単位](#detail-only)にあります。

<a id="join-grain"></a>

### 入力行のキーと、結合する単位を分ける

`data.tables.B.key`は**CSVの各行を失わず識別する組合せ**、`steps[].keys`は**二つの表を何の単位で対応づけたいか**です。同じである必要はありません。例えばBが`id,part,value`で1つのidに複数のpartを持つ場合、入力キーは`["id","part"]`にします。それでも要件が「idごとに、相手が1件でもあるか」なら、Bを`groupBy:["B.id"]`で集計してから`keys:["A.id","B.id"]`で結合します。partまで結合条件に加えると、同じidに別partがある行を「相手なし」にしてしまいます。

一方、要件が「idとpartの両方が同じ行」なら複合の結合キーが必要です。**一致件数が多い組合せを無条件に選んではいけません。** 次の順に決めます。

1. 最終台帳の1行の単位と、何を満たせば一致と呼ぶかを業務フローから言葉にする。
2. 入力全件の一意性を数え、CSVのkeyを決める。結合する両側の候補キーも重複数と1対多を確認する。
3. フローが許す候補ごとにRunUpdateで実行し、未一致・出力行数・集計値を比較する。列を足して一致が急減したなら、その追加条件を要件が本当に求めているか確認する。
4. 一致例と未一致例をそれぞれ実データへ戻って確かめる。件数だけで正しいと決めない。フローが曖昧なら、必要な識別規則を質問して確定する。

4表の完成例では、Bの入力キーはid+part、集計・結合の単位はidです。照合用の計算列を使う場合も、その同じ単位で集計のgroupByと結合keysを揃えます。

<a id="detail-only"></a>

### 明細ファイルだけから伝票単位の台帳を作る

明細しか出力できないシステムから、伝票ごとに1行の台帳を作る例です。`data/出荷実績.csv`の見出しは`伝票番号,行番号,出荷日,得意先,数量,単価`で、伝票番号＋行番号で一意、出荷日と得意先は同じ伝票の明細行ですべて同じです。台帳は伝票番号ごとに1行で、合計金額＝Σ数量×単価を持ちます。

入力表の`key`は各行を識別する`["伝票番号","行番号"]`にします。`calculate`で明細金額を作り、`aggregate`の`groupBy`に伝票番号と、伝票の中で同じ値の列（出荷日、得意先）を並べると、伝票ごとに1行の表`T`になります。`groupBy`の列は元の参照（`L.伝票番号`など）のまま残り、集計列は`T.合計金額`と`T.明細数`になります。台帳の`identity`は集計後のグループ列`L.伝票番号`です。入力表の`key`と一致している必要はなく、更新のたびに一意で空欄が無いことを確かめます。

```jsonc
{
  "tables": {"L": {"label":"出荷明細", "file":"出荷実績.csv", "key":["伝票番号","行番号"], "keyValidation":{"length":"variable"}}},
  "types": {"L.出荷日":{"type":"date","format":"yyyy/MM/dd"}, "L.数量":{"type":"number"}, "L.単価":{"type":"number"}, "T.合計金額":{"type":"number"}},
  "labels": {"L.伝票番号":"伝票番号", "L.行番号":"行番号", "L.出荷日":"出荷日", "L.得意先":"得意先", "L.数量":"数量", "L.単価":"単価",
             "L.金額":"明細金額", "T":"伝票集計", "T.合計金額":"合計金額", "T.明細数":"明細数", "ledger":"台帳"},
  "jobs": [{"id":"update", "kind":"update", "inputs":[{"table":"L"}], "steps":[
    {"operation":"calculate", "target1":"L", "column":"金額", "expression":"L.数量 * L.単価", "output":"L"},
    {"operation":"aggregate", "target1":"L", "groupBy":["L.伝票番号","L.出荷日","L.得意先"],
     "aggregates":[{"function":"sum","column":"L.金額","as":"合計金額"},{"function":"count","as":"明細数"}], "output":"T"},
    {"operation":"merge", "target1":"T", "target2":"ledger", "keys":["L.伝票番号","L.伝票番号"],
     "sourceOnly":"add", "both":"update", "targetOnly":"keep", "output":"ledger"}
  ]}],
  "ledger": {"identity":"L.伝票番号", "search":{"columns":["L.伝票番号"],"match":"exact"},
             "columns":{"source":["L.伝票番号","L.出荷日","L.得意先","T.合計金額","T.明細数"],
                        "application":[{"name":"workState","onSourceChange":"reset"}]}}
}
```

これは`data`の中身の例です。画面の`field`や出力の`defaultFields`にも同じ参照（`T.合計金額`）を書きます。伝票の中で値が違う列（商品コードなど）を`groupBy`に入れると伝票が複数行に割れるので、台帳に要らない列は`groupBy`に入れません。単価に`"1,250"`のようなカンマ入りの表記があっても、`number`の宣言で数値として計算します。

<a id="leading-zero"></a>

### 先頭ゼロが消えた番号を扱う範囲

読取りは先頭ゼロを保ちますが、消えた桁数は推測しません。`123`が`00123`なのか`00000123`なのか、また別の番号なのかは値だけでは決まりません。`data.types`をnumberにしても結合キーは自動的に数値比較になりません。表示の書式も保存値や結合キーを直しません。

**先頭ゼロの有無で同一の番号だと業務上確定している場合だけ**、既存のcalculateで照合用の数値文字列を作れます。次は両側とも1～8桁の数字という契約の部分例です。data.labelsに`A.matchId/B.matchId`も追加します。元のA.id/B.idは残ります。

```json
{"operation":"calculate","target1":"A","column":"matchId","expression":"regexExtract(A.id, '^[0-9]{1,8}$') + 0","output":"A"},
{"operation":"calculate","target1":"B","column":"matchId","expression":"regexExtract(B.id, '^[0-9]{1,8}$') + 0","output":"B"},
{"operation":"join","target1":"A","target2":"B","keys":["A.matchId","B.matchId"],"condition":"left","output":"joined"}
```

`00000001`と`1`の照合値がどちらも`1`になります。数字以外・9桁以上は内側のregexExtractで止めるため、別の正常値へ丸めません。照合値が重複する場合は、元は別レコードなのか、同じ単位の明細として集計すべきなのかを確認してください。番号のゼロが意味を持つ業務ではこの正規化を使いません。上の式は先頭ゼロを戻して保存するものではなく、照合用の値を作るものです。

**元の番号は必ず8桁だったと確定していて、8桁へ戻した計算列が必要な場合**は、次の既存式でも表せます。専用のpadLeft関数はありません。

```json
{"operation":"calculate","target1":"A","column":"paddedId","expression":"regexExtract('x00000000' + regexExtract(A.id, '^[0-9]{1,8}$'), '[0-9]{8}$')","output":"A"}
```

非数字の`x`は、`+`を数値加算ではなく文字列連結にするためです。内側で1～8桁を検証してから末尾8桁を取り、`23`は`00000023`、`00000000`はそのままになります。桁数の分からない番号へこの例をそのまま当ててはいけません。`data.labels.A.paddedId`も必要です。計算列は後のkeysで使えますが、直接`ledger.columns.source`へ増やせるのは登録入力表の実在する列なので、元の値を台帳に残す使い方が明確です。

**同一性や元の桁数が不明なら、元システムの仕様または対応表なしに復元できません。** その行を黙って捨てたり、未一致として済ませたりせず、未達の要求と必要な情報を報告してください。すべての行を正しく扱う要求を満たせない段階で、推測した設定を完成品として提出しません。

### 複合キーで結合する

`A`の見出しが`id,part,amount`、`B`が`number,item,value`なら、表の`key`をそれぞれ`["id","part"]`、`["number","item"]`にします。次の段で同じ組合せ同士を結合します。

```json
{"operation":"join", "target1":"A", "target2":"B",
 "keys":[["A.id","A.part"],["B.number","B.item"]],
 "condition":"left", "output":"joined"}
```

台帳の行も`A.id + A.part`で区別するなら、`data.ledger.identity:["A.id","A.part"]`、`source`にも両列を指定し、書込み段の`keys`を`[["A.id","A.part"],["A.id","A.part"]]`とします。検索は1列ごとの一致なので、`A.id`で探して複数候補から枝番を選べます。複合キーを単一の文字列へ連結して検索する設定ではありません。

### 左にあって相手の表にない行を確認する

上の`left`結合なら左は残り、相手のキー列が空の行が未一致です。相手のキーを保存・出力列へ含めて確認できます。単なる説明列は元から空のことがあるので、**未一致の判定には相手のキー列**を使います。

行集合を直接確認するなら次の段です。`-RunUpdate`の`values.missing.rows`へ、選んだ元表の行が出ます。この段を置くなら`missing`にも`data.labels`で表示名を付け、更新ジョブの最後には台帳を書き込む段を置いてください。

```json
{"operation":"extract", "target1":"A", "target2":"B",
 "keys":[["A.id","A.part"],["B.number","B.item"]],
 "condition":"exclude", "output":"missing"}
```

<a id="aggregate"></a>

### 複数の行を単位ごとに合計する

例：`A.csv`は1行1IDの`id`列、`B.csv`は`id,part,amount`。Bのキーは`["id","part"]`とし、同じidの異なるpartを失わずに読ませます。`data.labels`には`A.id/B.id/B.part/B.amount/joined/ledger`、台帳のidentityは`A.id`、sourceは`["A.id","B.amount"]`を指定します。次が更新ジョブの全stepsです。

```json
[
  {"operation":"aggregate", "target1":"B", "groupBy":["B.id"],
   "aggregates":[{"function":"sum","column":"B.amount","as":"amount"}], "output":"B"},
  {"operation":"join", "target1":"A", "target2":"B",
   "keys":["A.id","B.id"], "condition":"left", "output":"joined"},
  {"operation":"merge", "target1":"joined", "target2":"ledger",
   "keys":["A.id","A.id"], "sourceOnly":"add", "both":"update", "targetOnly":"keep", "output":"ledger"}
]
```

`sum`は`column`必須、`count`は`{"function":"count","as":"count"}`としcolumnを書きません。複数集計を並べられます。集計結果は`output.as`という列参照になります。上例の`output:B`、`as:amount`は既存の`B.amount`へ結果を置き、台帳へ保存できる形にしています。

集計結果は`output`と`as`で決まる参照で台帳の`source`に書けます。上例なら`B.amount`で、`as`を`total`にすれば`B.total`です。既存の見出しと同じ名前へ置いても、新しい名前でも構いません。新しい名前には`data.labels`で表示名を付けます。結果名を入力表と別の名前（`output:"totals"`）にした場合は、`groupBy`の列は元の参照（`B.id`）のまま、集計列は`totals.amount`になります。[計算結果を台帳へ保存する](#calculation-storage)に条件をまとめています。

例えば同じidの`"¥66,131"`と`(500)`の合計は65631、別idの`２５０`は250。Aにだけあるidの合計欄は空です。完全重複の再送行は入力時点で除かれるので、再送を二重加算する動作にはなりません。

<a id="append-files"></a>

### 複数ファイルを縦に足す

月ごとに出力される同じ列構成のファイルを1つの台帳にまとめる例です。`受注_4月.csv`と`受注_5月.csv`の見出しはどちらも`受注番号,受注日,得意先,金額,状態`で、5月のファイルには4月の受注が再掲されることがあり、再掲された行は5月の内容が最新です。状態が`取消`の受注は台帳に載せません。

手順は4段です。`append`で縦に足し、`distinct`で受注番号の重複を除き、`extract`の`where`で取消の行を選んで`delete`で除き、`merge`で台帳へ書きます。`append`の出力は左の表の参照（`A.受注番号`など）を使うので、**最新のファイルを左（`target1`）に置き**、`distinct`は先に現れた行を残します。台帳の`identity`は左の表のキー`A.受注番号`です。

```jsonc
{
  "tables": {
    "A": {"label":"受注（5月）", "file":"受注_5月.csv", "key":"受注番号"},
    "B": {"label":"受注（4月）", "file":"受注_4月.csv", "key":"受注番号"}
  },
  "labels": {"A.受注番号":"受注番号", "A.受注日":"受注日", "A.得意先":"得意先", "A.金額":"金額", "A.状態":"状態",
             "B.受注番号":"受注番号", "B.受注日":"受注日", "B.得意先":"得意先", "B.金額":"金額", "B.状態":"状態",
             "all":"4月と5月", "uniq":"重複を除いた行", "cancelled":"取消の行", "kept":"取消を除いた行", "ledger":"台帳"},
  "jobs": [{"id":"update", "kind":"update", "inputs":[{"table":"A"},{"table":"B"}], "steps":[
    {"operation":"append", "target1":"A", "target2":"B", "output":"all"},
    {"operation":"distinct", "target1":"all", "columns":["A.受注番号"], "output":"uniq"},
    {"operation":"extract", "target1":"uniq", "where":{"column":"A.状態","operator":"equals","value":"取消"}, "output":"cancelled"},
    {"operation":"delete", "target1":"uniq", "target2":"cancelled", "output":"kept"},
    {"operation":"merge", "target1":"kept", "target2":"ledger", "keys":["A.受注番号","A.受注番号"],
     "sourceOnly":"add", "both":"update", "targetOnly":"keep", "output":"ledger"}
  ]}],
  "ledger": {"identity":"A.受注番号", "search":{"columns":["A.受注番号"],"match":"exact"},
             "columns":{"source":["A.受注番号","A.受注日","A.得意先","A.金額","A.状態"],
                        "application":[{"name":"workState","onSourceChange":"reset"}]}}
}
```

`extract`と`delete`は同じ表の値（ここでは`uniq`）に対して使います。`distinct`の前の`all`から選んだ行集合を`uniq`から`delete`することはできません。RunUpdateの`summary.rows`が「4月の件数＋5月だけの件数－取消の件数」になることを確かめてください。`join`でまとめると片方のファイルにしか無い受注が台帳に入りません。

<a id="conditional-replace"></a>

### コード表がファイルに無いとき、値を条件で置き換える

状態が`1`、`2`、`9`のコードで出力され、その意味（1＝受注、2＝出荷済、9＝取消）がマニュアルにしか無い場合の例です。IFやCASEの式はありませんが、**`calculate`で列を複製し、コードごとに`extract`の`where`で行を選んで`update`の`set`に定数を書く**ことで、条件つきの置き換えを表せます。定数は単一引用符で囲みます。

```jsonc
{
  "tables": {"A": {"label":"受注", "file":"受注.csv", "key":"受注番号"}},
  "types": {"A.受注日":{"type":"date","format":"yyyy/MM/dd"}, "A.金額":{"type":"number"}},
  "labels": {"A.受注番号":"受注番号", "A.受注日":"受注日", "A.得意先":"得意先", "A.金額":"金額",
             "A.状態コード":"状態コード", "A.状態名":"状態名",
             "c1":"コード1の行", "c2":"コード2の行", "c9":"コード9の行", "ledger":"台帳"},
  "jobs": [{"id":"update", "kind":"update", "inputs":[{"table":"A"}], "steps":[
    {"operation":"calculate", "target1":"A", "column":"状態名", "expression":"A.状態コード", "output":"A"},
    {"operation":"extract", "target1":"A", "where":{"column":"A.状態コード","operator":"equals","value":"1"}, "output":"c1"},
    {"operation":"update", "target1":"A", "target2":"c1", "set":[{"column":"A.状態名","expression":"'受注'"}], "output":"A"},
    {"operation":"extract", "target1":"A", "where":{"column":"A.状態コード","operator":"equals","value":"2"}, "output":"c2"},
    {"operation":"update", "target1":"A", "target2":"c2", "set":[{"column":"A.状態名","expression":"'出荷済'"}], "output":"A"},
    {"operation":"extract", "target1":"A", "where":{"column":"A.状態コード","operator":"equals","value":"9"}, "output":"c9"},
    {"operation":"update", "target1":"A", "target2":"c9", "set":[{"column":"A.状態名","expression":"'取消'"}], "output":"A"},
    {"operation":"merge", "target1":"A", "target2":"ledger", "keys":["A.受注番号","A.受注番号"],
     "sourceOnly":"add", "both":"update", "targetOnly":"keep", "output":"ledger"}
  ]}],
  "ledger": {"identity":"A.受注番号", "search":{"columns":["A.受注番号"],"match":"exact"},
             "columns":{"source":["A.受注番号","A.受注日","A.得意先","A.金額","A.状態コード","A.状態名"],
                        "application":[{"name":"workState","onSourceChange":"reset"}]}}
}
```

最初の`calculate`はコードをそのまま複製しているので、表に無いコードは名前に変わらずコードのまま残り、画面で気づけます。`update`の行集合は、その直前の`A`（更新済みの表）から選びます。古い表の値から選んだ行集合を新しい表へ適用することはできません。置き換えが多い場合や表が変わる場合は、コード表をCSVにして`join`で付けるほうが変更に強く、依頼元にコード表の提供を求める価値があります。`extract`は`contains`や`startsWith`、数値の`atLeast`等も使えるので、範囲による区分（金額が1万円以上なら「大口」）も同じ形で表せます。

<a id="four-tables"></a>

## 4表を集計・左結合し、内容変更で確認状態を戻す完成例

特定の業務に依存しない、4件のレコードとその補助表の例です。**この節の設定は単独で起動できます。** 新しい試験用のアプリ一式へ配置してください。既存の出荷設定や実運用データへ上書きしません。

4本ともUTF-8で保存します。表AのIDを台帳の1行にし、表Bの数値をIDごとに集計、表Cの分類名、表Dの取消フラグを左結合します。相手の無いAの行も残ります。取消フラグは保存する普通の入力列です。空から`1`へ変わるとsourceの変更になり、確認済みの行を初期状態へ戻します。取消以外の保存列が変わった場合も同じ規則です。

`data/A.csv` — 基準の4行。日付は日本語の年月日書式です。

```csv
id,group,date
R01,G1,2026年9月9日
R02,G2,2026年10月1日
R03,G1,2026年12月31日
R04,G2,2026年1月2日
```

`data/B.csv` — 集計前は`id`と`part`の組合せで一意です。R01の合計は80、R03は1200、R02とR04には行がありません。

```csv
id,part,amount
R01,01,100
R01,02,(20)
R03,01,"¥1,200"
```

`data/C.csv` — 分類名を対応づけます。

```csv
group,name
G1,分類1
G2,分類2
```

`data/D.csv` — 最初は見出しだけです。データ0件の通知は出ますが、左結合なのでAの行は残ります。

```csv
id,cancel
```

アプリ直下の`settings.json`へ次を保存します。省略した項目は既定値で動きます。`tables.*.label`で表名を定義しているので、`labels`へ`A/B/C/D`を重ねて書きません。

```jsonc
{
  "schema": 3,
  "paths": {"dataDir":"data", "ledger":"data/ReaderDataViewer-Ledger.xlsx", "log":"ReaderDataViewer.log"},
  "search": {"pattern":"R[0-9]{2}"},
  "watch": {"targets":[]},
  "data": {
    "encoding":"utf-8",
    "tables": {
      "A":{"label":"基準表", "file":"A.csv", "key":"id", "keyValidation":{"length":"variable"}},
      "B":{"label":"数値表", "file":"B.csv", "key":["id","part"], "keyValidation":{"length":"variable"}},
      "C":{"label":"分類表", "file":"C.csv", "key":"group", "keyValidation":{"length":"variable"}},
      "D":{"label":"変更表", "file":"D.csv", "key":"id", "keyValidation":{"length":"variable"}}
    },
    "types": {"A.date":{"type":"date","format":"yyyy年M月d日"}, "B.amount":{"type":"number"}},
    "labels": {
      "A.id":"識別子", "A.group":"分類番号", "A.date":"日付",
      "B.id":"対応する識別子", "B.part":"枝番", "B.amount":"合計値",
      "C.group":"分類番号", "C.name":"分類名", "D.id":"対象識別子", "D.cancel":"取消フラグ",
      "AB":"数値を結合", "ABC":"分類を結合", "ABCD":"変更を結合", "ledger":"台帳"
    },
    "jobs":[{
      "id":"update", "kind":"update",
      "inputs":[{"table":"A"},{"table":"B"},{"table":"C"},{"table":"D"}],
      "steps":[
        // Bの行をidごとに1行へ。既存のamount列名を使うので台帳へB.amountとして保存できる。
        {"operation":"aggregate", "target1":"B", "groupBy":["B.id"],
         "aggregates":[{"function":"sum","column":"B.amount","as":"amount"}], "output":"B"},
        {"operation":"join", "target1":"A", "target2":"B", "keys":["A.id","B.id"], "condition":"left", "output":"AB"},
        {"operation":"join", "target1":"AB", "target2":"C", "keys":["A.group","C.group"], "condition":"left", "output":"ABC"},
        {"operation":"join", "target1":"ABC", "target2":"D", "keys":["A.id","D.id"], "condition":"left", "output":"ABCD"},
        // 元にない台帳の行は保持。入力が持つsource列を更新する。
        {"operation":"merge", "target1":"ABCD", "target2":"ledger", "keys":["A.id","A.id"],
         "sourceOnly":"add", "both":"update", "targetOnly":"keep", "output":"ledger"}
      ]
    }],
    "ledger": {
      "identity":"A.id", "search":{"columns":["A.id"],"match":"exact"},
      "columns": {
        "source":["A.id","A.group","A.date","B.id","B.amount","C.name","D.cancel"],
        "application":[{"name":"workState","onSourceChange":"reset"}]
      }
    }
  },
  "screen": {
    "card":{"startSize":[818,636],"font":"Meiryo UI","fontSize":10,"gap":8,"padding":[8]},
    "workState": {
      "trigger":"manual", "store":{"column":"確認状態"},
      "states":[{"id":"todo","text":"未確認","stored":"FALSE"},{"id":"done","text":"確認済","stored":"TRUE","look":"accent"}],
      "initial":"todo", "transitions":[{"from":"todo","to":"done"},{"from":"done","to":"todo"}]
    },
    "export":{"defaultFields":["A.id","A.group","A.date","B.id","B.amount","C.name","D.cancel","$work"]},
    "candidates":{"columns":[{"header":"識別子","value":{"field":"A.id"}},{"header":"分類名","value":{"field":"C.name"}}]},
    "sections":[
      {"type":"keyPanel","title":"検索","figure":{"label":"識別子","value":{"field":"A.id"}},
       "input":{"label":"検索値","maxLength":3},
       "buttons":[{"action":"search","text":"検索"},{"action":"clear","text":"クリア"},{"action":"workState"}]},
      {"type":"fieldList","title":"レコード","rowHeight":30,"rows":[
        {"label":"分類名","value":{"field":"C.name"}},
        {"label":"日付","value":{"field":"A.date","format":{"kind":"date","from":"yyyy年M月d日","to":"yyyy-MM-dd"}}},
        {"label":"合計値","value":{"field":"B.amount"}},
        {"label":"数値表の識別子","value":{"field":"B.id","empty":"相手なし"}},
        {"label":"取消フラグ","value":{"field":"D.cancel","empty":""}}
      ]},
      {"type":"sendBar","value":{"state":"pendingCount"},"buttons":[{"action":"sendChanges","text":"変更を送信"}]},
      {"type":"statusBar","segments":[{"value":{"state":"appState"}},{"value":{"state":"ledgerRows"}}],
       "buttons":[{"action":"updateRecords","text":"レコード更新","job":"update"},{"action":"tableExport","text":"テーブル出力"},{"action":"settings","text":"設定"}]}
    ]
  }
}
```

最初の答え合わせは次です。相手の有無は`B.amount`の0ではなく`B.id`の空欄で判定します。合計0の行が存在する場合と、相手がいない場合を区別するためです。

| A.id | B.id | B.amount | C.name | D.cancel | 状態 |
|---|---|---|---|---|---|
| R01 | R01 | 80 | 分類1 | 空 | FALSE |
| R02 | 空 | 空 | 分類2 | 空 | FALSE |
| R03 | R03 | 1200 | 分類1 | 空 | FALSE |
| R04 | 空 | 空 | 分類2 | 空 | FALSE |

`-ValidateOnly`、次に`-RunUpdate -Output C:/trial/first.json`を実行します（配置に合わせて`-Config/-DataDir`も指定）。期待値は`summary.rows=4`、`skippedEmpty=0`、`skippedDuplicate=0`、`resetRows=0`。`joins`の`AB/ABC/ABCD`は順に`unmatchedLeft=2/0/4`です。

次に、試験用のアプリを起動して台帳を作成し、R01・R02を検索してそれぞれ確認済みにして送信します。台帳の控えを`C:/trial/before.xlsx`として残します。手元で状態を倒しただけでは台帳へ入らないので、**控えを取る前に送信してください。** アプリを終了してから、試験用D.csvだけを次へ入れ替えます。

```csv
id,cancel
R02,1
```

`-RunUpdate -BaselineLedger C:/trial/before.xlsx -Output C:/trial/second.json`を実行すると、行数は4のまま、`resetRows=1`、戻った行はR02だけです。R01のTRUEは残り、R02のD.cancelは`1`、状態はFALSEになります。`ABCD.unmatchedLeft=3`です。窓でも「レコード更新」を実行し、戻った行の通知とR02の表示を確認できます。同じ内容をもう一度更新しても、新たな戻りは0件です。

この例は取消の行を台帳から削除しません。取消を表す値を保存して再確認を促す構成です。削除したい場合は一般操作の`delete`を明示して別の手順にします。


<a id="ledger"></a>

## 台帳の更新と確認状態

`data.ledger.columns.source`は入力の内容、`application`の`workState`はアプリが持つ状態です。マージはsourceの内容だけを上書きし、内容が同じなら確認状態を保ちます。`sourceOnly:add`、`both:update`、`targetOnly:keep`なら、新しい行を足しつつ台帳を育てます。入力にない行も消す同期は`targetOnly:delete`を明記した場合だけです。

**変更・取消等を表す値が、台帳に保存するsource列へ反映される構成なら**、`onSourceChange:reset`により、内容が変わったその行の確認状態を`screen.workState.initial`へ戻し、戻した行を通知します。例えば状態を表す入力列を台帳に保存し、その値が変わったときに更新ジョブを実行すれば、この規則が働きます。

これは取消専用処理ではありません。**保存するsource列のどれが変わってもresetの対象**です。取消時だけ戻し、他の列の変更では保持するという条件別の規則はありません。`preserve`なら内容変更でも状態を保持します。ファイルが到着しただけでは更新せず、起動時の確認またはレコード更新で取り込みます。

状態は2つに限りません。`states`の各`id`は内部名、`stored`は台帳へ書く文字列、`text`と`short`は人向けの長い名前・短い名前、`look`は見せ方です。`transitions`でfromからtoへの1つの行先を定義します。状態の保存値・識別列・列順を変更する場合は、全PCを止めて未送信分と台帳の控えを取り、移行後の契約を揃えてください。

検索後の状態ボタン操作は、まずそのWindowsユーザーのローカル未送信変更へ保存されます。**変更を送信して初めて共有台帳へ反映**します。未確認へ戻した値も送れます。閉じても未送信分は残りますが、別PCへ自動移動しません。

マージと送信は同じ共有ロックを使い、ロック取得後に最新台帳を読み直します。送信では内容の指紋と変更前の状態を照合し、別の人の変更を黙って上書きしません。既に同じ値を送信済みなら再送を二重計上しません。競合した未送信変更は保持され、明示的な破棄と確認の操作でのみ取り除きます。

保存が遅れても、その結果が確定するまでは終了を保留します。`jobs`の期限を短くしてもI/Oを強制終了する設定にはなりません。台帳はアプリ用の`LEDGER`シート1枚で管理し、追加のシートや非空の余分な列を拒否します。書式・数式・図形・コメント等を保持する用途ではありません。

<a id="screen"></a>

## 表示・監視・省略値

画面はWin98の1種類です。表示項目は`screen.sections`、候補一覧は`screen.candidates`、初期出力列は`screen.export.defaultFields`で決めます。状態表示と判定帯の色は、実際の検索・更新条件と別です。見た目だけを変えて処理の意味が変わったと思わないようにしてください。

表示値は`{"field":"A.id"}`、`{"fields":["A.id","A.name"],"joiner":" / "}`、`{"state":"pendingCount"}`のいずれか1方式です。field/fieldsはsourceへ保存する台帳列だけ。`empty:""`で空欄をそのまま見せ、`empty:"未入力"`で代わりの文言を指定できます。形式は次のとおりです。

```json
{"field":"A.amount", "format":{"kind":"number","group":true}, "empty":""}
{"field":"A.date", "format":{"kind":"date","from":"yyyyMMdd","to":"yyyy-MM-dd"}}
```

`format.kind`省略はtext、numberのgroup省略はtrue、dateのfrom省略はyyyyMMdd、to省略はyyyy-MM-dd。formatは表示の変換で、元の台帳値を書き換えません。

| state名 | 表示する情報 |
|---|---|
| searchKey / candidateCount / rowNumber | 検索入力、該当件数、候補行の番号 |
| workState / workStateShort | 現在の確認状態の表示名／短い表示名 |
| appState / watchLabel / watchDetail | アプリの状況、監視先の名前／詳細 |
| ledgerFile / ledgerRows / ledgerSaved | 台帳ファイル、台帳件数、保存状況 |
| pendingCount | このPCの未送信件数 |
| mergeMs / searchMs | 更新確認・検索の計測時間 |
| pid / logName / clock / userName / hostName | プロセス番号、ログ名、時計、利用者、端末 |

`sections`の部品は次の8種類です。寸法はCSS px、fontSize等の文字サイズはptです。DPIが変われば画面上の画素数も変わります。**寸法を小さくしたり、文言・フォントを大きくしたりすると文字が収まらなくなります。** 最長の実データと候補、狭い窓と使用するDPIで確認してください。設定できる範囲内の数値であることと、読める画面になることは別です。

| type | 中身・必須要素 | 任意の調整 |
|---|---|---|
| titleBar | 表題。必須はtypeだけ | brand、tags:[{text,look}]、buttons |
| keyPanel | `figure:{label,value}` | title、input:{label,width,maxLength}、buttons |
| columns | `items`にfieldList/textBoxを1個以上 | weights（個数をitemsと同じに）、gap、stackBelow（この幅以下は縦並び、0は切替なし） |
| fieldList | `rows:[{label,value}]`を1行以上 | title、labelWidth、rowHeight |
| textBox | value、lines（1～60） | title |
| statusBand | judgment（judgmentsのID） | label、height、sub（表示値配列）、joiner |
| sendBar | value、sendChangesボタン1つ | height |
| statusBar | 下部区画 | segments:[{prefix,value,dot}]、buttons、height |

全部品に`margin`を指定できます。padding/marginは数値1つ、またはCSSと同じ1～4要素の配列です。columnsのgap省略は17、stackBelow省略は760 CSS pxです。card.gapとcolumns.gapは別の値です。候補列のwidth省略は自動幅、align省略はleft、render省略はtext。`render:tag`と`looks:{"値":"accent","*":"neutral"}`で値別の見せ方を選べます。

`judgments`はsourceの値に対してrulesを上から試し、最初に当たるresultを採用します。equals配列／pattern正規表現／empty:trueは、同じ規則内ではORです。各resultの`text`と`look:ok/ng/undefined/error`をresultsへ書きます。規則不一致はundefined、読めない値はerror。意味のある「要確認」等の文言を、省略のために成功表示へ置き換えないでください。

表示値の定義には、省略可能な`requires`配列で表示に必要な保存列を指定できます。例：`{"field":"A.name","requires":["B.id"],"empty":""}`は、`B.id`が空なら画面の値を空欄にし、台帳の`A.name`は保持します。複数指定した場合は全列に値があるときだけ表示します。判定の`source`に指定した場合、必要な列が空なら判定元を空として`empty:true`の規則へ進みます。存在しない列は設定エラーになり、欠損値として隠しません。

`paths/search/watch/jobs`の領域全体は省略できます。`screen.card`、judgments、watch.targetsも省略できます。全て省略して画面を自動生成する機能ではありません。schema、dataの入力・ジョブ・台帳、screenのworkState/export/candidates/sectionsは必要です。全項目索引に省略値と制約を載せています。

外部監視は`watch.targets`に窓から対象欄までのUI Automationの経路を指定します。window→path配列→fieldの順で探し、readでvalue/text/nameを選びます。各段のautomationId/className/name/nameLike/processName/controlTypes/requireValuePattern/index/scopeを使えます。省略属性では絞りません。nameLikeは`*`と`?`、indexは0から、scopeはdescendantsまたはchildren。対象アプリごとにUIAの公開内容・読取権限が違うため、設定画面の「画面から選ぶ」と実窓で確認してください。

画面の設定保存はpaths/search/watchを編集します。その領域内のコメントは再生成で消えるため、委託元はコメント入りの正本を別に残してください。他の領域とコメントは保持します。保存前に外部編集を照合して競合時に上書きを拒否します。JSONを外部編集しただけでは全項目が実行中へ自動反映されません。構造を変更した場合は再起動してください。

<a id="limits"></a>

## 設定では提供しないこと

- グループ・場所等の単位での確認状態の一括変更。人は行ごとに確認します。
- 任意SQL/C#、任意の関数、データベース接続、CSV到着を合図にした自動更新。IFやCASEの式も無いが、条件つきの置き換えは`extract`と`update`の組合せで表せる（[例](#conditional-replace)）。
- 取消の列だけを特別扱いする状態リセット。source内容の変更に対するreset/preserveを使います。
- 日付形式やBOMなし文字コードの自動選択、同じ列で複数日付形式を試す機能。読取り失敗時のUTF-16候補表示は案内だけで、自動切替はしません。
- 更新ジョブが作らない列を台帳に増設すること。保存できるのは入力表の列と、calculate / aggregate / select が作った列だけです（[計算結果の保存](#calculation-storage)）。
- 任意シート選択、Excel数式の再計算、識別規則の不明な先頭ゼロの推測復元、ブックの装飾の保存。識別規則が既知の番号には上記の計算列を使えます。XLSXの日付セルは`data.types`でdateを宣言した列だけ変換し、宣言のない列はシリアル値の数字のままです。
- 見出し行の自動判定。表題行のあるファイルは`headerRow`で見出しの行番号を指定します。
- 操作ログの場所・名前・列の変更。台帳と同じフォルダーの端末別CSVへ固定の形式で追記します（[操作ログ](#operation-log)）。
- 実物のカードリーダー制御。責務は設定した外部ウィンドウをUIAで読み取るところまでです。

CSV出力は現在このPCに見えている台帳と未送信の状態です。他PCの未送信分は含みません。出力は新規CSVのみ。数式として解釈される値を抑えるExcel向け設定は既定で有効ですが、Excelが番号や日付を自動変換することまでは防ぎません。機械向けに生の値を出すときは出力画面で数式無効化を解除します。

<a id="feedback-log"></a>

## エラーを返すファイルログ

通常起動、`-ValidateOnly`、`-RunUpdate`とも、設定を読む前から次へ記録します。設定の`paths.log`が壊れていても場所は変わりません。

```text
%LOCALAPPDATA%/ReaderDataViewer/logs/feedback.log
```

主な記録は、起動時の設定・入力・出力パス、PID、処理段階、設定の全検出エラー、入力・ジョブの例外、除外件数・結果の要約、WPFとWebView2の例外、終了コードです。台帳の全行は入れず、`-RunUpdate`の結果JSONに出します。従来の`paths.log`にも処理ログを残し、その内容をこのローカルログへも記録します。

この場所に書けないときは`%TEMP%/ReaderDataViewer/logs/feedback.log`へ退避します。問い合わせ・AIへの再提出では、**該当する実行の`BEGIN BOOTSTRAP`または`BEGIN`から`END`までを含むファイル**を渡してください。時刻と`pid=`で実行を区別できます。複数回起動した場合も同じファイルへ追記します。ローテーションをまたいだときは`.1`等も渡し、退避先へ出た場合はそのファイルも含めます。ファイルには入力や保存先のパス、問題のある値が含まれるため、共有範囲を確認してください。

1ファイルは4 MiBで切り替え、`feedback.log.1`（直前）、`.2`、`.3`の3世代を残します。次の切替時に最古の`.3`を削除します。新しく記録するログは現在分を含む最大4ファイル・16 MiBで、`paths.log`も同じ規則です。1件だけで4 MiBを超える記録は切り詰めた印を残します。導入前から上限超過のログがある場合は初回に旧内容を丸ごと`.1`へ移すため、その旧世代が残る間は16 MiBを超える場合があります。旧版の日付付き起動ログは自動削除しません。

各記録をディスクへ確定してから進みます。通常の例外には原因とスタックを残します。実行中は5秒ごとに最後の処理段階を記録し、WPFが10秒以上応答しないと`UI NOT RESPONDING`を記録します。これは停止の観測であり、原因の断定や処理の強制中断はしません。ヘッドレス処理の`HEARTBEAT`は「まだ動いていて完了未確認」の意味で、時間だけで異常と判定しません。

**強制終了・電源断では、終了した後の原因は書けません。** 残った開始記録と最後の段階に対して`END`が無ければ、正常終了を確認できなかった実行です。理由なく通常のプロセス終了へ進んだ場合は可能なら`UNEXPECTED EXIT`も記録します。WPFの応答とWebView2の描画プロセスは別なので、描画プロセス障害は`WEBVIEW PROCESS FAILED`を見ます。

OSがPowerShellの開始を拒む場合、スクリプト自体の構文・引数束縛で実行前に止まる場合、ローカルとTEMPの両方が書込み不能の場合には、このログを作れません。最後の場合は標準エラーへ書込み失敗を出します。ログが見当たらなければ両方のフォルダーを確認し、`.cmd`または上のPowerShellコマンドから起動してOS側のエラーも確認してください。

## 起動

起動と導入は [利用者向けREADME](../README.md)、同梱サンプルは [業務設定の説明](PAYMENT-GUIDE.md) を参照してください。

## 全項目索引

出荷設定の413箇所を一覧にしています。入れ物や繰り返しの出現も数えるため、独立した調整項目が413個あるという意味ではありません。出荷時の監視見本を外したためK015〜K026は欠番です。監視機能自体は設定画面から対象を追加して使用できます。内部名と互換項目も、効かない設定を調整しないよう明記します。追加の表・列・状態には同じ規則を繰り返します。新しい表別encoding・複合キー等は本文の該当節を参照。

| 番号・JSONパス | 意味・変えると起きること |
|---|---|
| <a id="k001"></a>K001 `schema` | 設定形式の版。必須で3。この数字は調整しません。 |
| <a id="k002"></a>K002 `paths` | 配置を決める領域。全体を省略すると入力data、台帳data/ReaderDataViewer-Ledger.xlsx、ログReaderDataViewer.log。変更は再起動後。 |
| <a id="k003"></a>K003 `paths.dataDir` | 入力フォルダー。相対パスはアプリ一式のフォルダーが基準。入力のfileはこの中を探します。省略data。 |
| <a id="k004"></a>K004 `paths.ledger` | 複数PCで共有する1本の.xlsx。相対パスはアプリ基準。入力・ログと同じファイルにはできません。省略data/ReaderDataViewer-Ledger.xlsx。同じフォルダーへ端末別の操作ログ <台帳名>-操作ログ-<端末名>.csv を追記します。 |
| <a id="k005"></a>K005 `paths.log` | このPCの実行ログ。相対パスはアプリ基準で、省略するとアプリ一式のフォルダー直下のReaderDataViewer.log（dataの中ではない）。共有台帳と分け、PCごとに別の場所へ。 |
| <a id="k006"></a>K006 `search` | 検索入力の形式と候補数。テーブルのキー検証とは別で、設定画面の保存後から効きます。 |
| <a id="k007"></a>K007 `search.pattern` | 入力全体に一致させる正規表現。省略は数字8桁。英字等を使う入力なら例 [A-Z0-9-]+ へ。台帳の検索対象列はdata.ledger.search.columns。 |
| <a id="k008"></a>K008 `search.candidateRowsShown` | 候補一覧へ載せる上限1～1000、省略100。小さくすると一覧から選べる行も減るので、検索条件を絞れることを確認。 |
| <a id="k009"></a>K009 `watch` | 外部ウィンドウの値を監視して検索する設定。省略すると監視対象なし。CSV到着の監視ではありません。 |
| <a id="k010"></a>K010 `watch.pollMs` | 値を読みに行く間隔、5～5000ms、省略40。短くすると反応が速くなり対象アプリへの問い合わせが増えます。 |
| <a id="k011"></a>K011 `watch.stableMs` | 値が変わらず続いたら検索へ渡すまでの時間、0～60000ms、省略120。短すぎると入力途中の値を拾います。 |
| <a id="k012"></a>K012 `watch.rebindMs` | 未接続の窓を探し直す間隔、50～60000ms、省略400。長くすると窓を開いた後の接続が遅れます。 |
| <a id="k013"></a>K013 `watch.preferFocusedWindow` | 一致する窓が複数あるとき前面を優先するか。省略true。別窓を拾わないようwindowも絞ってください。 |
| <a id="k014"></a>K014 `watch.targets` | 監視先の配列。複数指定可、[]または省略なら監視しません。各対象はwindow→pathの順→fieldをたどります。 |
| <a id="k027"></a>K027 `jobs` | 処理時間の監視と共有ファイルの再確認間隔。単位ms。OSのI/Oを強制終了する設定ではありません。省略可。 |
| <a id="k028"></a>K028 `jobs.checkTimeoutMs` | 更新確認が遅れたと判定する時間、1000～3600000ms、省略180000。完了前に結果を成功扱いしません。 |
| <a id="k029"></a>K029 `jobs.searchTimeoutMs` | 検索が遅れたと判定する時間、1000～3600000ms、省略30000。遅れた検索結果の採用を防ぐための期限。 |
| <a id="k030"></a>K030 `jobs.saveTimeoutMs` | 保存が遅れたと判定する時間、1000～3600000ms、省略60000。期限で保存済みとは決めず、実際の完了を待ちます。 |
| <a id="k031"></a>K031 `jobs.markOverdueMs` | 保存待ちの長期化を知らせる時間、1000～3600000ms、省略180000。過ぎても未確定の書込みがある間は終了を保留します。 |
| <a id="k032"></a>K032 `jobs.pumpMs` | 画面側の処理状況確認間隔、100～60000ms、省略1000。長いと完了や遅延の表示が遅れます。 |
| <a id="k033"></a>K033 `jobs.lockRetryMs` | 共有ロック取得を再試行する間隔、50～5000ms、省略250。短いと共有先へのアクセスが増えます。 |
| <a id="k034"></a>K034 `jobs.lockStaleMs` | 古いロックか調べる経過時間、60000～86400000ms、省略600000。生きている書込みを時刻だけで解除する設定ではありません。 |
| <a id="k035"></a>K035 `jobs.markerPollMs` | 他PCの更新通知を確認する間隔、500～60000ms、省略3000。長いと最新台帳への切替に気づくのが遅れます。 |
| <a id="k036"></a>K036 `data` | 入力表と操作順、台帳の列を定義します。画面の判定や作業状態とは別の領域です。必須。 |
| <a id="k037"></a>K037 `data.encoding` | CSVの既定文字コード、省略utf-8。各tables.<ID>.encodingで上書き可。BOMは自動切替ではなく一致確認に使います。XLSXには適用しません。 |
| <a id="k038"></a>K038 `data.tables` | 入力表を短いIDで登録。1表以上必須。IDにドットや空白は使わず、ledgerは予約名。ここに登録した全表を起動時に確認します。 |
| <a id="k039"></a>K039 `data.tables.A` | この入力表の定義。fileは入力元、keyは一意性を確かめる列。encodingを追加するとこの表だけ文字コードを上書きでき、見出しが1行目でなければheaderRowに見出しの行番号を書きます。 |
| <a id="k040"></a>K040 `data.tables.A.label` | 表の画面向け名。省略すると表ID。data.labelsに同じ表IDを重ねて書けません（省略時も不可）。列名や結合条件は変わりません。 |
| <a id="k041"></a>K041 `data.tables.A.file` | 入力CSVまたはXLSXのファイル名、必須。相対名はpaths.dataDir基準、絶対パスも可。 |
| <a id="k042"></a>K042 `data.tables.A.key` | 一意な行を識別する列名、必須。複数列なら ["id","part"] の配列。組合せ全体で重複判定し、各列へkeyValidationが適用されます。 |
| <a id="k043"></a>K043 `data.tables.B` | この入力表の定義。fileは入力元、keyは一意性を確かめる列。encodingを追加するとこの表だけ文字コードを上書きでき、見出しが1行目でなければheaderRowに見出しの行番号を書きます。 |
| <a id="k044"></a>K044 `data.tables.B.label` | 表の画面向け名。省略すると表ID。data.labelsに同じ表IDを重ねて書けません（省略時も不可）。列名や結合条件は変わりません。 |
| <a id="k045"></a>K045 `data.tables.B.file` | 入力CSVまたはXLSXのファイル名、必須。相対名はpaths.dataDir基準、絶対パスも可。 |
| <a id="k046"></a>K046 `data.tables.B.key` | 一意な行を識別する列名、必須。複数列なら ["id","part"] の配列。組合せ全体で重複判定し、各列へkeyValidationが適用されます。 |
| <a id="k047"></a>K047 `data.tables.B.keyValidation` | キーの検証規則。省略はASCII・列ごとの固定長。規則違反・空キーの行は除外し、内容が違う同一キーは全行を除外して警告します。文字列を整形して別のキーを作る指定ではありません。 |
| <a id="k048"></a>K048 `data.tables.B.keyValidation.empty` | skip（省略時）と旧指定errorは、いずれも空キーの行を除外し、件数とファイル・行・列を警告します。複合キーではどれか1列が空なら対象。 |
| <a id="k049"></a>K049 `data.tables.C` | この入力表の定義。fileは入力元、keyは一意性を確かめる列。encodingを追加するとこの表だけ文字コードを上書きでき、見出しが1行目でなければheaderRowに見出しの行番号を書きます。 |
| <a id="k050"></a>K050 `data.tables.C.label` | 表の画面向け名。省略すると表ID。data.labelsに同じ表IDを重ねて書けません（省略時も不可）。列名や結合条件は変わりません。 |
| <a id="k051"></a>K051 `data.tables.C.file` | 入力CSVまたはXLSXのファイル名、必須。相対名はpaths.dataDir基準、絶対パスも可。 |
| <a id="k052"></a>K052 `data.tables.C.key` | 一意な行を識別する列名、必須。複数列なら ["id","part"] の配列。組合せ全体で重複判定し、各列へkeyValidationが適用されます。 |
| <a id="k053"></a>K053 `data.types` | 数値・日付の型宣言。省略列は文字列、形からの型推測なし。表示だけの書式はscreenのvalue.formatで指定します。calculate/aggregate/selectで作った列にも宣言でき、その値は生成した段階で検査し、違反のある行を除外して値・列・元の行を警告します。XLSXの日付セルはdateの宣言でformatの文字列に変換します。 |
| <a id="k054"></a>K054 `data.types["A.a_date"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k055"></a>K055 `data.types["A.a_date"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k056"></a>K056 `data.types["A.a_date"].format` | 日付入力の正確な書式。例 yyyyMMdd / yyyy-MM-dd / yyyy/MM/dd / yyyy年M月d日（2026年9月9日）。1列に1書式で、自動判別や複数書式混在ではありません。 |
| <a id="k057"></a>K057 `data.types["B.b_date"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k058"></a>K058 `data.types["B.b_date"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k059"></a>K059 `data.types["B.b_date"].format` | 日付入力の正確な書式。例 yyyyMMdd / yyyy-MM-dd / yyyy/MM/dd / yyyy年M月d日（2026年9月9日）。1列に1書式で、自動判別や複数書式混在ではありません。 |
| <a id="k060"></a>K060 `data.types["C.c_exp"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k061"></a>K061 `data.types["C.c_exp"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k062"></a>K062 `data.types["C.c_exp"].format` | 日付入力の正確な書式。例 yyyyMMdd / yyyy-MM-dd / yyyy/MM/dd / yyyy年M月d日（2026年9月9日）。1列に1書式で、自動判別や複数書式混在ではありません。 |
| <a id="k063"></a>K063 `data.types["A.a_rate"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k064"></a>K064 `data.types["A.a_rate"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k065"></a>K065 `data.types["A.a_amount"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k066"></a>K066 `data.types["A.a_amount"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k067"></a>K067 `data.types["B.b_qty"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k068"></a>K068 `data.types["B.b_qty"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k069"></a>K069 `data.types["B.b_total"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k070"></a>K070 `data.types["B.b_total"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k071"></a>K071 `data.types["C.c_price"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k072"></a>K072 `data.types["C.c_price"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k073"></a>K073 `data.types["C.c_stock"]` | この実在する入力列の型宣言。ラベルではなく表ID.見出しで指定。番号の先頭ゼロを保つ列はtextまたは省略。 |
| <a id="k074"></a>K074 `data.types["C.c_stock"].type` | date / number / text。dateはformat必須。numberは全角数字・¥・桁区切り・括弧マイナスを読めます。空セルは型検証を止めません。XLSXの日付セル（シリアル値）はdate宣言の列だけformatの文字列へ変換します。 |
| <a id="k075"></a>K075 `data.labels` | 内部参照から画面向け表示名への対応。手順の出力名、キー、集計列にも必要。表自体のラベルはtables側に書きます。 |
| <a id="k076"></a>K076 `data.labels["B.key1"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k077"></a>K077 `data.labels["B.key2"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k078"></a>K078 `data.labels["B.b_line"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k079"></a>K079 `data.labels["B.b_ref"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k080"></a>K080 `data.labels["B.b_date"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k081"></a>K081 `data.labels["B.b_qty"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k082"></a>K082 `data.labels["B.b_unit"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k083"></a>K083 `data.labels["B.b_total"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k084"></a>K084 `data.labels["B.b_status"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k085"></a>K085 `data.labels["B.b_memo"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k086"></a>K086 `data.labels["A.a_name"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k087"></a>K087 `data.labels["A.a_code"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k088"></a>K088 `data.labels["A.a_grade"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k089"></a>K089 `data.labels["A.a_dept"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k090"></a>K090 `data.labels["A.a_amount"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k091"></a>K091 `data.labels["C.c_code"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k092"></a>K092 `data.labels["C.c_name"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k093"></a>K093 `data.labels["C.c_cat"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k094"></a>K094 `data.labels["C.c_price"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k095"></a>K095 `data.labels["C.c_stock"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k096"></a>K096 `data.labels["A.key1"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k097"></a>K097 `data.labels["C.key2"]` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k098"></a>K098 `data.labels.ledger` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k099"></a>K099 `data.labels.middle1` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k100"></a>K100 `data.labels.middle2` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k101"></a>K101 `data.labels.target1` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k102"></a>K102 `data.labels.target2` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k103"></a>K103 `data.labels.target` | この内部参照を処理内容・出力一覧で見せる名前。変えてもCSVの見出しや値は変わりません。空文字は不可。 |
| <a id="k104"></a>K104 `data.jobs` | 操作を実行順に並べたジョブ。1件以上必須。最初のkind:updateが起動時と-RunUpdateの仕事で、最後はledgerを作る必要があります。 |
| <a id="k105"></a>K105 `data.jobs[0].id` | ボタンから指定するジョブID。全ジョブで一意。必須。 |
| <a id="k106"></a>K106 `data.jobs[0].name` | 画面で読むジョブ名、省略id。自然文で処理を別に定義せず、stepsが実処理と説明の共通の元です。 |
| <a id="k107"></a>K107 `data.jobs[0].kind` | updateまたはdelete。最初のupdateが自動更新。操作の固定順序を選ぶ値ではありません。 |
| <a id="k108"></a>K108 `data.jobs[0].inputs` | このジョブが読む表または外部の値一覧。[]も可。table形式とfile/column形式を混ぜられます。 |
| <a id="k109"></a>K109 `data.jobs[0].inputs[0].table` | data.tablesのIDを参照。列・キー・文字コードをその表から引き継ぐため、この入力オブジェクトにはtableだけ書きます。 |
| <a id="k110"></a>K110 `data.jobs[0].inputs[1].table` | data.tablesのIDを参照。列・キー・文字コードをその表から引き継ぐため、この入力オブジェクトにはtableだけ書きます。 |
| <a id="k111"></a>K111 `data.jobs[0].inputs[2].table` | data.tablesのIDを参照。列・キー・文字コードをその表から引き継ぐため、この入力オブジェクトにはtableだけ書きます。 |
| <a id="k112"></a>K112 `data.jobs[0].steps` | 操作を上から実行、1段以上必須。行集合はそれを選んだ同じ表へdelete/updateで適用します。 |
| <a id="k113"></a>K113 `data.jobs[0].steps[0].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k114"></a>K114 `data.jobs[0].steps[0].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k115"></a>K115 `data.jobs[0].steps[0].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k116"></a>K116 `data.jobs[0].steps[0].keys` | 左と右の対応列。単一は["A.id","B.id"]、複合は[["A.id","A.part"],["B.id","B.part"]]。左右同じ個数・対応順。 |
| <a id="k117"></a>K117 `data.jobs[0].steps[0].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k118"></a>K118 `data.jobs[0].steps[0].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k119"></a>K119 `data.jobs[0].steps[1].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k120"></a>K120 `data.jobs[0].steps[1].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k121"></a>K121 `data.jobs[0].steps[1].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k122"></a>K122 `data.jobs[0].steps[1].keys` | 左と右の対応列。単一は["A.id","B.id"]、複合は[["A.id","A.part"],["B.id","B.part"]]。左右同じ個数・対応順。 |
| <a id="k123"></a>K123 `data.jobs[0].steps[1].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k124"></a>K124 `data.jobs[0].steps[1].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k125"></a>K125 `data.jobs[0].steps[2].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k126"></a>K126 `data.jobs[0].steps[2].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k127"></a>K127 `data.jobs[0].steps[2].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k128"></a>K128 `data.jobs[0].steps[2].keys` | 左と右の対応列。単一は["A.id","B.id"]、複合は[["A.id","A.part"],["B.id","B.part"]]。左右同じ個数・対応順。 |
| <a id="k129"></a>K129 `data.jobs[0].steps[2].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k130"></a>K130 `data.jobs[0].steps[2].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k131"></a>K131 `data.jobs[0].steps[2].sourceOnly` | 入力にだけある行をadd（追加）またはignore（無視）。mergeでは必須。 |
| <a id="k132"></a>K132 `data.jobs[0].steps[2].both` | 両方にある行をupdate（source列更新）またはkeep（保持）。mergeでは必須。 |
| <a id="k133"></a>K133 `data.jobs[0].steps[2].targetOnly` | 台帳にだけある行をkeep（省略時は保持）またはdelete（削除）。keepなら月を追加して台帳を育てられます。 |
| <a id="k134"></a>K134 `data.jobs[1].id` | ボタンから指定するジョブID。全ジョブで一意。必須。 |
| <a id="k135"></a>K135 `data.jobs[1].name` | 画面で読むジョブ名、省略id。自然文で処理を別に定義せず、stepsが実処理と説明の共通の元です。 |
| <a id="k136"></a>K136 `data.jobs[1].kind` | updateまたはdelete。最初のupdateが自動更新。操作の固定順序を選ぶ値ではありません。 |
| <a id="k137"></a>K137 `data.jobs[1].inputs` | このジョブが読む表または外部の値一覧。[]も可。table形式とfile/column形式を混ぜられます。 |
| <a id="k138"></a>K138 `data.jobs[1].inputs[0].id` | 外部の値一覧につけるID。ジョブ内で一意、ledger不可。後段のtarget2から参照します。 |
| <a id="k139"></a>K139 `data.jobs[1].inputs[0].label` | 外部入力の画面向け名、省略id。同じIDを別の場所に使う場合は同じ表示名にします。 |
| <a id="k140"></a>K140 `data.jobs[1].inputs[0].file` | 値一覧のCSV/XLSX。相対名はpaths.dataDir基準。この入力にもencodingを追加して既定文字コードを上書きできます。 |
| <a id="k141"></a>K141 `data.jobs[1].inputs[0].column` | 値一覧から読む1列の実見出し。keyとは別です。同じファイルを異なるcolumnで複数回定義できます。 |
| <a id="k142"></a>K142 `data.jobs[1].inputs[0].key` | 読み取った値を比較するときの論理列参照。例B.id。元CSVの見出しはcolumnに書きます。 |
| <a id="k143"></a>K143 `data.jobs[1].inputs[1].id` | 外部の値一覧につけるID。ジョブ内で一意、ledger不可。後段のtarget2から参照します。 |
| <a id="k144"></a>K144 `data.jobs[1].inputs[1].label` | 外部入力の画面向け名、省略id。同じIDを別の場所に使う場合は同じ表示名にします。 |
| <a id="k145"></a>K145 `data.jobs[1].inputs[1].file` | 値一覧のCSV/XLSX。相対名はpaths.dataDir基準。この入力にもencodingを追加して既定文字コードを上書きできます。 |
| <a id="k146"></a>K146 `data.jobs[1].inputs[1].column` | 値一覧から読む1列の実見出し。keyとは別です。同じファイルを異なるcolumnで複数回定義できます。 |
| <a id="k147"></a>K147 `data.jobs[1].inputs[1].key` | 読み取った値を比較するときの論理列参照。例B.id。元CSVの見出しはcolumnに書きます。 |
| <a id="k148"></a>K148 `data.jobs[1].steps` | 操作を上から実行、1段以上必須。行集合はそれを選んだ同じ表へdelete/updateで適用します。 |
| <a id="k149"></a>K149 `data.jobs[1].steps[0].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k150"></a>K150 `data.jobs[1].steps[0].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k151"></a>K151 `data.jobs[1].steps[0].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k152"></a>K152 `data.jobs[1].steps[0].keys` | 左と右の対応列。単一は["A.id","B.id"]、複合は[["A.id","A.part"],["B.id","B.part"]]。左右同じ個数・対応順。 |
| <a id="k153"></a>K153 `data.jobs[1].steps[0].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k154"></a>K154 `data.jobs[1].steps[0].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k155"></a>K155 `data.jobs[1].steps[1].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k156"></a>K156 `data.jobs[1].steps[1].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k157"></a>K157 `data.jobs[1].steps[1].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k158"></a>K158 `data.jobs[1].steps[1].keys` | 左と右の対応列。単一は["A.id","B.id"]、複合は[["A.id","A.part"],["B.id","B.part"]]。左右同じ個数・対応順。 |
| <a id="k159"></a>K159 `data.jobs[1].steps[1].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k160"></a>K160 `data.jobs[1].steps[1].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k161"></a>K161 `data.jobs[1].steps[2].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k162"></a>K162 `data.jobs[1].steps[2].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k163"></a>K163 `data.jobs[1].steps[2].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k164"></a>K164 `data.jobs[1].steps[2].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k165"></a>K165 `data.jobs[1].steps[2].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k166"></a>K166 `data.jobs[1].steps[3].operation` | 実行する一般操作。join/extract/delete/append/update/merge/replace/select/calculate/aggregate/sort/distinct。操作ごとに使える項目が異なります。同じ列構成のファイルを縦に足すのはappend、別表の列を横に付けるのはjoinです。 |
| <a id="k167"></a>K167 `data.jobs[1].steps[3].target1` | 操作の主対象。inputsのID、前段のoutput、またはledger。未作成の名前は参照できません。 |
| <a id="k168"></a>K168 `data.jobs[1].steps[3].target2` | 結合する相手、比較対象、選んだ行集合、または保存先ledger。1対象の操作には書きません。 |
| <a id="k169"></a>K169 `data.jobs[1].steps[3].condition` | joinはmatch/left/full、2表のextractはmatch/exclude、行集合のextractはboth/either/exclude。merge/deleteの空文字は互換項目で省略可。 |
| <a id="k170"></a>K170 `data.jobs[1].steps[3].output` | この操作の結果名。次のtargetに使えます。初出ならdata.labelsにも表示名を記入。台帳書込みの結果名はledger。 |
| <a id="k171"></a>K171 `data.ledger` | 共有台帳の契約。識別・検索・保存列を指定。変更する際は全PCを止めて控えを取り、既存台帳との移行を確認します。 |
| <a id="k172"></a>K172 `data.ledger.identity` | 台帳の各行を識別する列参照。1列の文字列または配列。sourceに含めます。入力表のkeyのほか、aggregateのgroupBy列など更新ジョブが作る列も指定でき、更新のたびに一意で空欄なしを確かめます。 |
| <a id="k173"></a>K173 `data.ledger.search` | 検索する列と一致方法。identityと別の列でも可。複数該当は人が候補を選ぶまで状態を変更しません。 |
| <a id="k174"></a>K174 `data.ledger.search.columns` | 検索する台帳列参照の配列、1列以上。いずれか1列が合えば候補。配列の値を連結した複合検索ではありません。 |
| <a id="k175"></a>K175 `data.ledger.search.match` | exact（省略時）は各列の完全一致、containsは部分一致。部分一致は全行を調べるため件数が増えるほど時間がかかります。 |
| <a id="k176"></a>K176 `data.ledger.columns` | 入力が持つsourceとアプリが持つapplicationを分離。作業状態をCSVの列で上書きしないための境界。 |
| <a id="k177"></a>K177 `data.ledger.columns.application` | アプリが持つ列の規則。現在はworkStateの1種類。sourceへ同じ状態列を混ぜません。 |
| <a id="k178"></a>K178 `data.ledger.columns.application[0].name` | 内部の列種別名。workState固定。利用者へ出す名前やExcel見出しはscreen.workState側で変更します。 |
| <a id="k179"></a>K179 `data.ledger.columns.application[0].onSourceChange` | resetは保存する入力内容が変わった行の状態をinitialへ戻し通知、preserveは保持。特定の列の変更だけを選ぶ規則ではありません。 |
| <a id="k180"></a>K180 `data.ledger.columns.source` | 台帳に保存する列を順に列挙。入力表の列に加え、calculateのcolumn・aggregateのas・selectのasで作った列を<結果名>.<列名>で書けます（README「計算結果を台帳へ保存する」）。identityと検索・表示・出力する列を含めます。最後の書込み段の入力に無い列はエラー。未保存列は画面から参照できません。 |
| <a id="k181"></a>K181 `screen` | 部品と表示値の定義。業務の表示項目を決めます。見た目はWin98のみ。必須。 |
| <a id="k182"></a>K182 `screen.card` | 窓全体の書体・起動寸法・隙間。省略可。大きな文字や長いラベルに変えたら行の高さと列幅も実窓で確認。 |
| <a id="k183"></a>K183 `screen.card.width` | 互換項目。現在の窓幅とレイアウトには使いません。窓の起動寸法はstartSizeを変更。省略可。 |
| <a id="k184"></a>K184 `screen.card.startSize` | [幅,高さ]、窓内のCSS px。枠・影・DPIを掛けた画面上の画素数とは別。省略[840,830]、最低480×300。 |
| <a id="k185"></a>K185 `screen.card.gap` | 枠間・内側等へ渡す共通の隙間、0～200 CSS px、省略17。狭めると文字やボタンが枠へ近づきます。 |
| <a id="k186"></a>K186 `screen.card.padding` | 画面外周の余白、CSS px。[全辺] / [上下,左右] / [上,左右,下] / [上,右,下,左]。省略[14.3,14.3,14.2,14.3]。 |
| <a id="k187"></a>K187 `screen.card.font` | 端末にある書体名、省略Yu Gothic UI。書体が変わると文字幅・必要な高さも変わるため長い値を確認。 |
| <a id="k188"></a>K188 `screen.card.fontSize` | 本文文字の大きさ、6～24pt、省略9。ptはCSS pxではありません。拡大時はラベル幅やボタンの収まりも確認。 |
| <a id="k189"></a>K189 `screen.card.keyValueFontSize` | 主キー表示の文字、6～36pt、省略15。大きくすると必要な高さも増え、長い番号が横に収まりにくくなります。 |
| <a id="k190"></a>K190 `screen.card.judgmentFontSize` | 判定文字の大きさ、6～36pt、省略15。文言を長くした場合は帯の高さ・窓幅も確認。 |
| <a id="k191"></a>K191 `screen.card.unsearchedFontSize` | 未検索表示の文字、6～36pt、省略12。本文と独立しています。 |
| <a id="k192"></a>K192 `screen.judgments` | 元の値を判定名へ変える定義。複数の判定を置けます。省略可。状態ボタンの遷移とは独立。 |
| <a id="k193"></a>K193 `screen.judgments.status1` | この名前の判定を定義。statusBand.judgmentから参照します。複数作って別の帯へ出せます。 |
| <a id="k194"></a>K194 `screen.judgments.status1.source` | 判定する値の出所。台帳列field等を1方式だけ指定。表示されたラベルでなく値を調べます。 |
| <a id="k195"></a>K195 `screen.judgments.status1.source.field` | 判定元の台帳列参照。data.ledger.columns.sourceへ保存した列だけ指定できます。 |
| <a id="k196"></a>K196 `screen.judgments.status1.rules` | 上から最初に当たる規則を採用。境界が重なる場合は順番が結果を変えます。 |
| <a id="k197"></a>K197 `screen.judgments.status1.rules[0].equals` | この配列のいずれかと完全一致なら当たり。pattern/emptyを併記するとOR。上から最初に当たった規則を採ります。 |
| <a id="k198"></a>K198 `screen.judgments.status1.rules[0].result` | 当たったときのresults内のID。undefined/errorは予約結果なので規則の行先に指定しません。 |
| <a id="k199"></a>K199 `screen.judgments.status1.rules[1].equals` | この配列のいずれかと完全一致なら当たり。pattern/emptyを併記するとOR。上から最初に当たった規則を採ります。 |
| <a id="k200"></a>K200 `screen.judgments.status1.rules[1].result` | 当たったときのresults内のID。undefined/errorは予約結果なので規則の行先に指定しません。 |
| <a id="k201"></a>K201 `screen.judgments.status1.results` | 規則のresult IDに対応する表示。undefined/errorの説明もここで業務に合う意味を記入できます。 |
| <a id="k202"></a>K202 `screen.judgments.status1.results.ok` | この判定結果の文言と見せ方。undefinedは規則不一致、errorは値を取得できない結果で、成功へ読み替えません。 |
| <a id="k203"></a>K203 `screen.judgments.status1.results.ok.text` | この結果を人に伝える文言。未定義・要確認等の意味を保ち、長くする際は判定帯の収まりを確認。 |
| <a id="k204"></a>K204 `screen.judgments.status1.results.ok.look` | 判定の見せ方ok/ng/undefined/error。文言や状態変更の可否とは独立。色だけでなくtextにも意味を書きます。 |
| <a id="k205"></a>K205 `screen.judgments.status1.results.ok.icon` | 互換項目。現在の判定帯ではアイコンを描かないため省略可。 |
| <a id="k206"></a>K206 `screen.judgments.status1.results.ng` | この判定結果の文言と見せ方。undefinedは規則不一致、errorは値を取得できない結果で、成功へ読み替えません。 |
| <a id="k207"></a>K207 `screen.judgments.status1.results.ng.text` | この結果を人に伝える文言。未定義・要確認等の意味を保ち、長くする際は判定帯の収まりを確認。 |
| <a id="k208"></a>K208 `screen.judgments.status1.results.ng.look` | 判定の見せ方ok/ng/undefined/error。文言や状態変更の可否とは独立。色だけでなくtextにも意味を書きます。 |
| <a id="k209"></a>K209 `screen.judgments.status1.results.undefined` | この判定結果の文言と見せ方。undefinedは規則不一致、errorは値を取得できない結果で、成功へ読み替えません。 |
| <a id="k210"></a>K210 `screen.judgments.status1.results.undefined.text` | この結果を人に伝える文言。未定義・要確認等の意味を保ち、長くする際は判定帯の収まりを確認。 |
| <a id="k211"></a>K211 `screen.judgments.status1.results.undefined.look` | 判定の見せ方ok/ng/undefined/error。文言や状態変更の可否とは独立。色だけでなくtextにも意味を書きます。 |
| <a id="k212"></a>K212 `screen.judgments.status1.results.error` | この判定結果の文言と見せ方。undefinedは規則不一致、errorは値を取得できない結果で、成功へ読み替えません。 |
| <a id="k213"></a>K213 `screen.judgments.status1.results.error.text` | この結果を人に伝える文言。未定義・要確認等の意味を保ち、長くする際は判定帯の収まりを確認。 |
| <a id="k214"></a>K214 `screen.judgments.status1.results.error.look` | 判定の見せ方ok/ng/undefined/error。文言や状態変更の可否とは独立。色だけでなくtextにも意味を書きます。 |
| <a id="k215"></a>K215 `screen.workState` | 人が行ごとに付ける状態の保存値・表示名・遷移。2状態に限定せず3状態以上も指定できます。必須。 |
| <a id="k216"></a>K216 `screen.workState.trigger` | manual（省略時）はボタンで状態変更。automaticは監視入力の検索で対象を確定したとき初期状態を進めます。automaticWhenを指定すると、指定した判定IDの結果IDが一致した案件だけが対象です。構文は[台帳保護の説明](PROTECTION-GUIDE.md#自動読取と前面化)を参照します。手入力検索を自動確定する指定ではありません。 |
| <a id="k217"></a>K217 `screen.workState.store` | 状態の保存先見出し。台帳に既にある場合は変更前に移行が必要。 |
| <a id="k218"></a>K218 `screen.workState.store.column` | Excel台帳で状態を保存する見出し。必須。source列の見出しと同じにしません。 |
| <a id="k219"></a>K219 `screen.workState.states` | 状態を列挙。idとstoredはそれぞれ一意。どちらも運用途中で意味を変えると既存台帳・未送信変更に影響します。 |
| <a id="k220"></a>K220 `screen.workState.states[0].id` | 状態を参照する内部ID。initial、transitionsのfrom/toからこのIDを使います。必須・一意。 |
| <a id="k221"></a>K221 `screen.workState.states[0].text` | 状態の表示名、省略id。人に伝える意味を変えないように決めます。 |
| <a id="k222"></a>K222 `screen.workState.states[0].short` | 狭い欄での短い状態名、省略text。省略しすぎて別状態と区別できなくしないこと。 |
| <a id="k223"></a>K223 `screen.workState.states[0].look` | 状態の見せ方neutral/accent/outline/faded/error、省略neutral。保存値とは別。 |
| <a id="k224"></a>K224 `screen.workState.states[0].stored` | 台帳と未送信データに保存する文字列、省略id。空文字・重複・制御文字不可。変更には既存データの移行が必要。 |
| <a id="k225"></a>K225 `screen.workState.states[1].id` | 状態を参照する内部ID。initial、transitionsのfrom/toからこのIDを使います。必須・一意。 |
| <a id="k226"></a>K226 `screen.workState.states[1].text` | 状態の表示名、省略id。人に伝える意味を変えないように決めます。 |
| <a id="k227"></a>K227 `screen.workState.states[1].short` | 狭い欄での短い状態名、省略text。省略しすぎて別状態と区別できなくしないこと。 |
| <a id="k228"></a>K228 `screen.workState.states[1].look` | 状態の見せ方neutral/accent/outline/faded/error、省略neutral。保存値とは別。 |
| <a id="k229"></a>K229 `screen.workState.states[1].stored` | 台帳と未送信データに保存する文字列、省略id。空文字・重複・制御文字不可。変更には既存データの移行が必要。 |
| <a id="k230"></a>K230 `screen.workState.initial` | 新規行とreset対象行に付ける状態のid。必須。textやstoredではなくstatesのidを書きます。 |
| <a id="k231"></a>K231 `screen.workState.transitions` | ボタンを押したときのfrom→to。省略なら遷移なし。1つのfromにつき行先は1つ。未処理へ戻す遷移も明記します。 |
| <a id="k232"></a>K232 `screen.workState.transitions[0].from` | 変更前の状態ID。statesで定義した値。同じfromから複数の行先は作れません。 |
| <a id="k233"></a>K233 `screen.workState.transitions[0].to` | 変更後の状態ID。statesで定義した値。未処理へ戻す操作もこの遷移で表します。 |
| <a id="k234"></a>K234 `screen.workState.transitions[0].confirm` | 変更前に人へ確認する文。省略または空文字なら確認なし。{state}/{key}/{A.id}等を利用可。 |
| <a id="k235"></a>K235 `screen.workState.transitions[0].done` | 変更後の通知文。省略または空文字なら空。{state}/{key}/{A.id}等を利用可。 |
| <a id="k236"></a>K236 `screen.workState.transitions[1].from` | 変更前の状態ID。statesで定義した値。同じfromから複数の行先は作れません。 |
| <a id="k237"></a>K237 `screen.workState.transitions[1].to` | 変更後の状態ID。statesで定義した値。未処理へ戻す操作もこの遷移で表します。 |
| <a id="k238"></a>K238 `screen.workState.transitions[1].done` | 変更後の通知文。省略または空文字なら空。{state}/{key}/{A.id}等を利用可。 |
| <a id="k239"></a>K239 `screen.workState.button` | 状態ボタンの表示。省略すると現在の状態名を表示します。 |
| <a id="k240"></a>K240 `screen.workState.button.text` | 状態ボタンの文言、省略{state}。{state}は現在の状態表示、{key}は検索入力、{A.id}等は台帳値に置換。 |
| <a id="k241"></a>K241 `screen.workState.button.tip` | 互換項目。現在の状態ボタンにはこの説明は表示されません。省略可。 |
| <a id="k242"></a>K242 `screen.export` | テーブル出力を開いたときの列選択。保存先パスはJSON設定ではなく出力時に選びます。 |
| <a id="k243"></a>K243 `screen.export.defaultFields` | 初期選択の台帳列参照、表示順、1列以上。$workは状態の保存値。出力画面で追加・削除・並べ替えできます。 |
| <a id="k244"></a>K244 `screen.sections` | 表示する部品の配列、1個以上。列の一覧・判定・送信・ステータスの内容をここで決めます。8種類の部品の記法はREADMEへ。 |
| <a id="k245"></a>K245 `screen.sections[0].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k246"></a>K246 `screen.sections[0].title` | この枠・一覧の見出し。省略時の既定名または空文字。長くする場合は幅と高さを確認。 |
| <a id="k247"></a>K247 `screen.sections[0].figure` | 検索欄の横に大きく表示する値とラベル。入力そのものと選択中のレコード値を混同しないように指定。 |
| <a id="k248"></a>K248 `screen.sections[0].figure.label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k249"></a>K249 `screen.sections[0].figure.value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k250"></a>K250 `screen.sections[0].figure.value.state` | 表示するアプリ値 searchKey。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k251"></a>K251 `screen.sections[0].input` | 手入力の検索欄。見た目と最大入力長。検索を許す形式はsearch.patternで別に設定します。 |
| <a id="k252"></a>K252 `screen.sections[0].input.label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k253"></a>K253 `screen.sections[0].input.placeholder` | 互換項目。現在の検索入力にはこの値を表示しません。省略可。 |
| <a id="k254"></a>K254 `screen.sections[0].input.width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k255"></a>K255 `screen.sections[0].input.maxLength` | 手入力欄の最大文字数1～256、省略64。これを超えるキーを入れられなくなるため、最長の検索値以上にします。 |
| <a id="k256"></a>K256 `screen.sections[0].buttons` | この領域から実行するボタンの配列。actionで役割、jobでジョブを指定。送信帯はsendChangesを1つ。 |
| <a id="k257"></a>K257 `screen.sections[0].buttons[0].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k258"></a>K258 `screen.sections[0].buttons[0].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k259"></a>K259 `screen.sections[0].buttons[0].primary` | 主な操作として強調するか、省略false。実行内容や確認の有無は変わりません。 |
| <a id="k260"></a>K260 `screen.sections[0].buttons[0].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k261"></a>K261 `screen.sections[0].buttons[1].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k262"></a>K262 `screen.sections[0].buttons[1].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k263"></a>K263 `screen.sections[0].buttons[1].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k264"></a>K264 `screen.sections[0].buttons[2].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k265"></a>K265 `screen.sections[1].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k266"></a>K266 `screen.sections[1].weights` | 横並び各部品へ配る幅の比率。省略は全部1。個数はitemsと同じ、正の数。長いラベルがある側の幅も確保します。 |
| <a id="k267"></a>K267 `screen.sections[1].gap` | 横並び部品間の隙間、0～200 CSS px、省略17。card.gapとは別。縮めると枠や文字が密集します。 |
| <a id="k268"></a>K268 `screen.sections[1].items` | 横並びに置くfieldList/textBoxの配列。1個以上、weightsの個数と揃えます。 |
| <a id="k269"></a>K269 `screen.sections[1].items[0].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k270"></a>K270 `screen.sections[1].items[0].title` | この枠・一覧の見出し。省略時の既定名または空文字。長くする場合は幅と高さを確認。 |
| <a id="k271"></a>K271 `screen.sections[1].items[0].labelWidth` | 項目名の幅、20～1000 CSS px、省略104。増やすと値側が狭くなり、減らすと長い項目名が収まりません。 |
| <a id="k272"></a>K272 `screen.sections[1].items[0].rowHeight` | 一覧の1行の高さ、20～200 CSS px。文字の大きさと改行が収まる値へ。省略はfieldList44、候補46。 |
| <a id="k273"></a>K273 `screen.sections[1].items[0].rows` | 項目名labelと表示値valueを上から並べる配列。1行以上。列の保存順とは独立。 |
| <a id="k274"></a>K274 `screen.sections[1].items[0].rows[0].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k275"></a>K275 `screen.sections[1].items[0].rows[0].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k276"></a>K276 `screen.sections[1].items[0].rows[0].value.field` | 表示する台帳列 B.key1。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k277"></a>K277 `screen.sections[1].items[0].rows[1].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k278"></a>K278 `screen.sections[1].items[0].rows[1].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k279"></a>K279 `screen.sections[1].items[0].rows[1].value.field` | 表示する台帳列 B.key2。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k280"></a>K280 `screen.sections[1].items[0].rows[2].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k281"></a>K281 `screen.sections[1].items[0].rows[2].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k282"></a>K282 `screen.sections[1].items[0].rows[2].value.field` | 表示する台帳列 B.b_line。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k283"></a>K283 `screen.sections[1].items[1].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k284"></a>K284 `screen.sections[1].items[1].title` | この枠・一覧の見出し。省略時の既定名または空文字。長くする場合は幅と高さを確認。 |
| <a id="k285"></a>K285 `screen.sections[1].items[1].labelWidth` | 項目名の幅、20～1000 CSS px、省略104。増やすと値側が狭くなり、減らすと長い項目名が収まりません。 |
| <a id="k286"></a>K286 `screen.sections[1].items[1].rowHeight` | 一覧の1行の高さ、20～200 CSS px。文字の大きさと改行が収まる値へ。省略はfieldList44、候補46。 |
| <a id="k287"></a>K287 `screen.sections[1].items[1].rows` | 項目名labelと表示値valueを上から並べる配列。1行以上。列の保存順とは独立。 |
| <a id="k288"></a>K288 `screen.sections[1].items[1].rows[0].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k289"></a>K289 `screen.sections[1].items[1].rows[0].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k290"></a>K290 `screen.sections[1].items[1].rows[0].value.field` | 表示する台帳列 A.a_name。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k291"></a>K291 `screen.sections[1].items[1].rows[1].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k292"></a>K292 `screen.sections[1].items[1].rows[1].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k293"></a>K293 `screen.sections[1].items[1].rows[1].value.field` | 表示する台帳列 A.a_code。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k294"></a>K294 `screen.sections[1].items[1].rows[2].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k295"></a>K295 `screen.sections[1].items[1].rows[2].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k296"></a>K296 `screen.sections[1].items[1].rows[2].value.field` | 表示する台帳列 A.a_grade。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k297"></a>K297 `screen.sections[1].items[1].rows[3].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k298"></a>K298 `screen.sections[1].items[1].rows[3].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k299"></a>K299 `screen.sections[1].items[1].rows[3].value.field` | 表示する台帳列 A.a_dept。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k300"></a>K300 `screen.sections[1].items[1].rows[4].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k301"></a>K301 `screen.sections[1].items[1].rows[4].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k302"></a>K302 `screen.sections[1].items[1].rows[4].value.field` | 表示する台帳列 A.a_date。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k303"></a>K303 `screen.sections[1].items[1].rows[5].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k304"></a>K304 `screen.sections[1].items[1].rows[5].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k305"></a>K305 `screen.sections[1].items[1].rows[5].value.field` | 表示する台帳列 A.a_amount。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k306"></a>K306 `screen.sections[1].items[1].rows[6].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k307"></a>K307 `screen.sections[1].items[1].rows[6].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k308"></a>K308 `screen.sections[1].items[1].rows[6].value.fields` | 複数の台帳列をこの順で表示。すべてsourceへの保存が必要。 |
| <a id="k309"></a>K309 `screen.sections[1].items[1].rows[6].value.joiner` | 複数列の間に挟む文字。省略は中点。長くすると表示欄で使える文字幅が減ります。 |
| <a id="k310"></a>K310 `screen.sections[2].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k311"></a>K311 `screen.sections[2].title` | この枠・一覧の見出し。省略時の既定名または空文字。長くする場合は幅と高さを確認。 |
| <a id="k312"></a>K312 `screen.sections[2].lines` | テキスト枠に割り当てる行数1～60、必須。内容が長い場合の読める範囲を決めます。 |
| <a id="k313"></a>K313 `screen.sections[2].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k314"></a>K314 `screen.sections[2].value.field` | 表示する台帳列 B.b_memo。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k315"></a>K315 `screen.sections[3].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k316"></a>K316 `screen.sections[3].title` | この枠・一覧の見出し。省略時の既定名または空文字。長くする場合は幅と高さを確認。 |
| <a id="k317"></a>K317 `screen.sections[3].lines` | テキスト枠に割り当てる行数1～60、必須。内容が長い場合の読める範囲を決めます。 |
| <a id="k318"></a>K318 `screen.sections[3].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k319"></a>K319 `screen.sections[3].value.field` | 表示する台帳列 C.c_remark。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k320"></a>K320 `screen.sections[4].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k321"></a>K321 `screen.sections[4].label` | 人に表示する欄名。省略は空文字。labelWidthや窓幅に収まる言葉にします。 |
| <a id="k322"></a>K322 `screen.sections[4].judgment` | この帯へ出すscreen.judgmentsのID。未定義の名前はエラー。別の帯で別の判定も表示できます。 |
| <a id="k323"></a>K323 `screen.sections[4].height` | この帯の高さ、CSS px。statusBandは30～500、sendBarは24～200、statusBarは20～200。大きな文字や複数行が収まるか確認。 |
| <a id="k324"></a>K324 `screen.sections[4].sub` | 判定に添える値の配列。保存列fieldかアプリstateをそれぞれ指定。省略なら補足なし。 |
| <a id="k325"></a>K325 `screen.sections[4].sub[0].field` | 表示する台帳列 B.b_status。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k326"></a>K326 `screen.sections[4].sub[1].state` | 表示するアプリ値 workState。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k327"></a>K327 `screen.sections[4].joiner` | 補足の値をつなぐ文字。長くすると本文が使える幅を消費します。 |
| <a id="k328"></a>K328 `screen.sections[5].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k329"></a>K329 `screen.sections[5].height` | この帯の高さ、CSS px。statusBandは30～500、sendBarは24～200、statusBarは20～200。大きな文字や複数行が収まるか確認。 |
| <a id="k330"></a>K330 `screen.sections[5].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k331"></a>K331 `screen.sections[5].value.state` | 表示するアプリ値 pendingCount。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k332"></a>K332 `screen.sections[5].buttons` | この領域から実行するボタンの配列。actionで役割、jobでジョブを指定。送信帯はsendChangesを1つ。 |
| <a id="k333"></a>K333 `screen.sections[5].buttons[0].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k334"></a>K334 `screen.sections[5].buttons[0].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k335"></a>K335 `screen.sections[5].buttons[0].primary` | 主な操作として強調するか、省略false。実行内容や確認の有無は変わりません。 |
| <a id="k336"></a>K336 `screen.sections[5].buttons[0].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k337"></a>K337 `screen.sections[6].type` | 部品の種類。titleBar/keyPanel/columns/fieldList/textBox/statusBand/sendBar/statusBar。columnsのitemsにはfieldList/textBoxだけ置けます。 |
| <a id="k338"></a>K338 `screen.sections[6].height` | この帯の高さ、CSS px。statusBandは30～500、sendBarは24～200、statusBarは20～200。大きな文字や複数行が収まるか確認。 |
| <a id="k339"></a>K339 `screen.sections[6].segments` | 下部ステータス欄の区画配列。数・順・表示値を選べます。4区画固定ではありません。 |
| <a id="k340"></a>K340 `screen.sections[6].segments[0].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k341"></a>K341 `screen.sections[6].segments[0].value.state` | 表示するアプリ値 appState。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k342"></a>K342 `screen.sections[6].segments[0].bold` | 互換項目。現在のWin98は共通の太字表示のため、この指定で太さは変わりません。省略可。 |
| <a id="k343"></a>K343 `screen.sections[6].segments[1].prefix` | この区画の値の前に付ける文字、省略空文字。区画の意味が分かる短い名前にします。 |
| <a id="k344"></a>K344 `screen.sections[6].segments[1].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k345"></a>K345 `screen.sections[6].segments[1].value.state` | 表示するアプリ値 watchLabel。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k346"></a>K346 `screen.sections[6].segments[2].prefix` | この区画の値の前に付ける文字、省略空文字。区画の意味が分かる短い名前にします。 |
| <a id="k347"></a>K347 `screen.sections[6].segments[2].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k348"></a>K348 `screen.sections[6].segments[2].value.state` | 表示するアプリ値 ledgerFile。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k349"></a>K349 `screen.sections[6].segments[3].prefix` | この区画の値の前に付ける文字、省略空文字。区画の意味が分かる短い名前にします。 |
| <a id="k350"></a>K350 `screen.sections[6].segments[3].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k351"></a>K351 `screen.sections[6].segments[3].value.state` | 表示するアプリ値 clock。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k352"></a>K352 `screen.sections[6].buttons` | この領域から実行するボタンの配列。actionで役割、jobでジョブを指定。送信帯はsendChangesを1つ。 |
| <a id="k353"></a>K353 `screen.sections[6].buttons[0].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k354"></a>K354 `screen.sections[6].buttons[0].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k355"></a>K355 `screen.sections[6].buttons[0].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k356"></a>K356 `screen.sections[6].buttons[1].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k357"></a>K357 `screen.sections[6].buttons[1].job` | このボタンが実行するdata.jobsのid。updateRecordsはkind:update、deleteRecordsはkind:deleteを指定。 |
| <a id="k358"></a>K358 `screen.sections[6].buttons[1].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k359"></a>K359 `screen.sections[6].buttons[1].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k360"></a>K360 `screen.sections[6].buttons[2].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k361"></a>K361 `screen.sections[6].buttons[2].job` | このボタンが実行するdata.jobsのid。updateRecordsはkind:update、deleteRecordsはkind:deleteを指定。 |
| <a id="k362"></a>K362 `screen.sections[6].buttons[2].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k363"></a>K363 `screen.sections[6].buttons[2].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k364"></a>K364 `screen.sections[6].buttons[3].action` | 押したときの動作。search/clear/workState/tableExport/updateRecords/deleteRecords/sendChanges/refreshLedger/settings。更新・削除はjobも指定。 |
| <a id="k365"></a>K365 `screen.sections[6].buttons[3].text` | 人に表示するボタン等の文言。長くすると必要幅が増えます。状態ボタン以外のボタンでは空文字不可。 |
| <a id="k366"></a>K366 `screen.sections[6].buttons[3].tip` | ボタンの補足説明。省略は空文字。見ただけで分からない操作の効き方を記入。 |
| <a id="k367"></a>K367 `screen.candidates` | 複数該当したときに人が選ぶ一覧。キー以外の見分ける情報も載せます。必須。 |
| <a id="k368"></a>K368 `screen.candidates.title` | この枠・一覧の見出し。省略時の既定名または空文字。長くする場合は幅と高さを確認。 |
| <a id="k369"></a>K369 `screen.candidates.hint` | 一覧で選ぶ人への補足説明、省略空文字。何を見比べて選ぶかを短く書きます。 |
| <a id="k370"></a>K370 `screen.candidates.width` | 候補ダイアログの幅、300～4000 CSS px、省略980。各列の必要幅も確保し、実際の窓幅・DPIで収まりを確認。 |
| <a id="k371"></a>K371 `screen.candidates.maxHeight` | 候補一覧本体の最大高さ、60～4000 CSS px、省略340。超えた分はスクロール。表示件数上限とは別。 |
| <a id="k372"></a>K372 `screen.candidates.rowHeight` | 一覧の1行の高さ、20～200 CSS px。文字の大きさと改行が収まる値へ。省略はfieldList44、候補46。 |
| <a id="k373"></a>K373 `screen.candidates.headerHeight` | 候補一覧の見出し行の高さ、16～200 CSS px、省略38。大きな文字を使う場合は高さも確認。 |
| <a id="k374"></a>K374 `screen.candidates.columns` | 候補一覧の列、1列以上。valueで台帳値またはアプリ状態を参照。行選択の判断に必要な情報を残します。 |
| <a id="k375"></a>K375 `screen.candidates.columns[0].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k376"></a>K376 `screen.candidates.columns[0].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k377"></a>K377 `screen.candidates.columns[0].align` | 列の文字寄せleft/right、省略left。数値を比較する列にはrightが使えます。 |
| <a id="k378"></a>K378 `screen.candidates.columns[0].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k379"></a>K379 `screen.candidates.columns[0].value.state` | 表示するアプリ値 rowNumber。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k380"></a>K380 `screen.candidates.columns[0].muted` | 候補列を控えめに表示するか、省略false。読みにくくしないよう必要な識別情報には注意。 |
| <a id="k381"></a>K381 `screen.candidates.columns[1].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k382"></a>K382 `screen.candidates.columns[1].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k383"></a>K383 `screen.candidates.columns[1].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k384"></a>K384 `screen.candidates.columns[1].value.field` | 表示する台帳列 B.key2。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k385"></a>K385 `screen.candidates.columns[1].bold` | 互換項目。現在のWin98は共通の太字表示のため、この指定で太さは変わりません。省略可。 |
| <a id="k386"></a>K386 `screen.candidates.columns[2].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k387"></a>K387 `screen.candidates.columns[2].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k388"></a>K388 `screen.candidates.columns[2].align` | 列の文字寄せleft/right、省略left。数値を比較する列にはrightが使えます。 |
| <a id="k389"></a>K389 `screen.candidates.columns[2].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k390"></a>K390 `screen.candidates.columns[2].value.field` | 表示する台帳列 B.b_line。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k391"></a>K391 `screen.candidates.columns[3].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k392"></a>K392 `screen.candidates.columns[3].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k393"></a>K393 `screen.candidates.columns[3].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k394"></a>K394 `screen.candidates.columns[3].value.field` | 表示する台帳列 B.b_ref。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k395"></a>K395 `screen.candidates.columns[4].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k396"></a>K396 `screen.candidates.columns[4].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k397"></a>K397 `screen.candidates.columns[4].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k398"></a>K398 `screen.candidates.columns[4].value.field` | 表示する台帳列 B.b_date。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k399"></a>K399 `screen.candidates.columns[5].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k400"></a>K400 `screen.candidates.columns[5].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k401"></a>K401 `screen.candidates.columns[5].align` | 列の文字寄せleft/right、省略left。数値を比較する列にはrightが使えます。 |
| <a id="k402"></a>K402 `screen.candidates.columns[5].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k403"></a>K403 `screen.candidates.columns[5].value.field` | 表示する台帳列 B.b_qty。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k404"></a>K404 `screen.candidates.columns[6].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k405"></a>K405 `screen.candidates.columns[6].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k406"></a>K406 `screen.candidates.columns[6].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k407"></a>K407 `screen.candidates.columns[6].value.field` | 表示する台帳列 B.b_status。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k408"></a>K408 `screen.candidates.columns[6].render` | 候補列の表示をtext（省略時）またはtagにする。tagの色はlooks、状態のlookと組み合わせます。 |
| <a id="k409"></a>K409 `screen.candidates.columns[6].looks` | 値から見せ方への対応。render:tagで使います。*はどの値にも個別指定がない場合の見せ方。 |
| <a id="k410"></a>K410 `screen.candidates.columns[6].looks.DONE` | この値のタグの見せ方。neutral/accent/outline/faded/error。*は既定。元の値や判定結果を変える指定ではありません。 |
| <a id="k411"></a>K411 `screen.candidates.columns[6].looks.HOLD` | この値のタグの見せ方。neutral/accent/outline/faded/error。*は既定。元の値や判定結果を変える指定ではありません。 |
| <a id="k412"></a>K412 `screen.candidates.columns[6].looks.VOID` | この値のタグの見せ方。neutral/accent/outline/faded/error。*は既定。元の値や判定結果を変える指定ではありません。 |
| <a id="k413"></a>K413 `screen.candidates.columns[6].looks["*"]` | この値のタグの見せ方。neutral/accent/outline/faded/error。*は既定。元の値や判定結果を変える指定ではありません。 |
| <a id="k414"></a>K414 `screen.candidates.columns[7].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k415"></a>K415 `screen.candidates.columns[7].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k416"></a>K416 `screen.candidates.columns[7].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k417"></a>K417 `screen.candidates.columns[7].value.field` | 表示する台帳列 C.c_code。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k418"></a>K418 `screen.candidates.columns[8].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k419"></a>K419 `screen.candidates.columns[8].width` | この入力欄・候補列・一覧の幅、CSS px。候補列の0または省略は自動幅。狭いと文字が収まらなくなるため実データで確認。 |
| <a id="k420"></a>K420 `screen.candidates.columns[8].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k421"></a>K421 `screen.candidates.columns[8].value.field` | 表示する台帳列 C.c_name。sourceに保存した列だけ使用可。ラベルはこの参照ではなく別のlabel/headerで変更。 |
| <a id="k422"></a>K422 `screen.candidates.columns[9].header` | 候補一覧の列見出し。省略は空文字。幅を超える長い見出しにすると読みづらくなります。 |
| <a id="k423"></a>K423 `screen.candidates.columns[9].value` | 表示する値。field / fields / state の1方式を選択。値なし時の表示や書式はこのオブジェクト内で指定できます。 |
| <a id="k424"></a>K424 `screen.candidates.columns[9].value.state` | 表示するアプリ値 workStateShort。利用可能な状態名と意味はREADMEの「表示値」。 |
| <a id="k425"></a>K425 `screen.candidates.columns[9].render` | 候補列の表示をtext（省略時）またはtagにする。tagの色はlooks、状態のlookと組み合わせます。 |
