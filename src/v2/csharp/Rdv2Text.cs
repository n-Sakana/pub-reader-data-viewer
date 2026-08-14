// ============================================================================
// Rdv2Text.cs -- every string the operator reads in the one-to-many build.
//
// The only file here with non-ASCII in it. build\pack_cmd.ps1 rewrites each
// non-ASCII character as \uXXXX on the way into the .cmd, because Add-Type
// hands the source to the in-box csc, which reads a BOM-less file in the
// system ANSI code page. No verbatim strings anywhere in the packed sources.
// ============================================================================

public static class Rdv2Text
{
    public const string AppTitle = "Reader Data Viewer";

    public const string MethodHash = "C# 自前ハッシュ";
    public const string MethodDict = "C# 標準 Dictionary";

    public const string StateBoot     = "起動中";
    public const string StateWatching = "監視中";
    public const string StateWaiting  = "メモ帳を待機中";
    public const string StateBusy     = "処理中";
    public const string StatePick     = "候補を選択してください";

    public const string LabelData    = "データ:";
    public const string LabelIndex   = "索引方式:";
    public const string LabelNotepad = "メモ帳:";
    public const string LabelCurrent = "現在の番号 (番号1)";
    public const string LabelField   = "入力欄:";
    public const string NotepadNone  = "メモ帳: 未接続 -- メモ帳を開くと自動で接続します";

    public const string BtnRebind = "メモ帳を再検出";
    public const string BtnManual = "手動実行";
    public const string BtnPick   = "この候補を表示";
    public const string BtnCancel = "選択しない";

    public const string BoxRecord  = "統合レコード (A-B-C 結合結果 1 件)";
    public const string BoxStages  = "工程別時間";
    public const string BoxHistory = "実行履歴";

    public const string ColFieldA = "表A 項目";
    public const string ColValueA = "表A 値";
    public const string ColFieldB = "表B 項目";
    public const string ColValueB = "表B 値";
    public const string ColFieldC = "表C 項目";
    public const string ColValueC = "表C 値";
    public const string ColStage  = "工程";
    public const string ColMs     = "ミリ秒";
    public const string ColShare  = "割合";
    public const string ColTime   = "時刻";
    public const string ColKey    = "番号1";
    public const string ColCand   = "候補数";
    public const string ColTotal  = "総時間ms";
    public const string ColPick   = "選択後ms";
    public const string ColDetect = "検知ms";
    public const string ColCheck  = "照合";

    // the candidate dialog. These column names are the documented identification
    // set; README.md lists the same eight.
    public const string PickTitle  = "番号1 に複数の候補があります";
    public const string PickHint   = "件の候補が見つかりました。表示する 1 件を選んでください。";
    public const string PickNo     = "#";
    public const string PickKey2   = "番号2";
    public const string PickLine   = "行番号";
    public const string PickSlip   = "伝票番号";
    public const string PickDate   = "日付";
    public const string PickQty    = "数量";
    public const string PickStatus = "状態";
    public const string PickItem   = "品目コード";
    public const string PickMaker  = "メーカー";

    public const string VerdictOne  = "一致 1 件";
    public const string VerdictMany = "候補 {n} 件から選択";
    public const string VerdictNone = "該当なし";
    public const string VerdictErr  = "エラー";

    public const string OracleOk  = "検算 OK";
    public const string OracleBad = "検算 NG";

    public static readonly string[] StageName =
    {
        "1 表A 読込",
        "2 表B 読込",
        "3 表C 読込",
        "4 表A 索引作成",
        "5 表B 索引作成",
        "6 表C 索引作成",
        "7 A-B 結合 (番号1)",
        "8 B-C 結合 (番号2)",
        "9 番号1 検索 (候補抽出)",
        "10 表示準備"
    };

    public const string StageOther  = "その他 (差分)";
    public const string StageTotal  = "合計 (選択待ちを除く)";
    public const string StagePick   = "選択後の表示 (計測外)";
    public const string StageDetect = "検知遅延 (参考・計測外)";
    public const string StageRows   = "データ行数";
    public const string StageCand   = "候補数";
    public const string StageProbes = "索引走査回数";

    public const string BootCompiled = "起動: C# コンパイル";
    public const string BootData     = "行 x 10 列 x 3 表";

    public const string ErrNo64   = "64 ビットのプロセスが必要です。";
    public const string ErrNoData = "データが見つかりません。build\\gen_data2.ps1 を先に実行してください: ";
}
