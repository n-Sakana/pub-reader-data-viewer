// ============================================================================
// RdvText.cs -- every string the operator reads, in one place.
//
// This is the only file with non-ASCII in it. build\pack_cmd.ps1 rewrites each
// non-ASCII character as \uXXXX on the way into the .cmd, because Add-Type
// hands the source to the in-box csc, which reads a file with no BOM in the
// system ANSI code page and would mangle it. So: no verbatim strings anywhere
// in the packed sources -- a verbatim literal keeps the backslash instead of
// turning the escape back into the character. The packer refuses to build one.
// ============================================================================

public static class RdvText
{
    public const string AppTitle = "Reader Data Viewer";

    public const string MethodCs     = "C# 単独 (WinForms)";

    public const string StateBoot     = "起動中";
    public const string StateWatching = "監視中";
    public const string StateWaiting  = "メモ帳を待機中";
    public const string StateBusy     = "処理中";

    public const string LabelData    = "データ:";
    public const string LabelNotepad = "メモ帳:";
    public const string LabelCurrent = "現在の番号 (番号1)";
    public const string LabelField   = "入力欄:";
    public const string NotepadNone  = "メモ帳: 未接続 -- メモ帳を開くと自動で接続します";

    public const string BtnRebind = "メモ帳を再検出";
    public const string BtnManual = "手動実行";

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
    public const string ColTotal  = "総時間ms";
    public const string ColDetect = "検知ms";
    public const string ColCheck  = "照合";

    public const string VerdictHit   = "一致 1 件";
    public const string VerdictMiss  = "該当なし";
    public const string VerdictError = "エラー";

    public const string OracleOk  = "検算 OK";
    public const string OracleBad = "検算 NG";

    public static readonly string[] StageName =
    {
        "1 表A 読込・索引",
        "2 表B 読込・索引",
        "3 表C 読込・索引",
        "4 A-B 結合 (番号1)",
        "5 B-C 結合 (番号2)",
        "6 番号1 検索",
        "7 画面表示"
    };

    public const string StageOther  = "その他 (差分)";
    public const string StageTotal  = "合計 (merge-select)";
    public const string StageDetect = "検知遅延 (参考・計測外)";
    public const string StageRows   = "データ行数";
    public const string StageProbes = "プローブ数";

    public const string BootCompiled = "起動: C# コンパイル";
    public const string BootExcel    = "起動: Excel 起動・ブック表示";
    public const string BootData     = "行 x 10 列 x 3 表";

    public const string ErrNo64      = "64 ビットのプロセスが必要です (32 ビットでは 1,000,000 行を保持できません)。";
    public const string ErrNoData    = "データが見つかりません。build\\gen_data.ps1 を先に実行してください: ";
    public const string ErrNoExcel   = "Excel を起動できませんでした: ";
    public const string ErrNoBook    = "表示用ブックが見つかりません: ";

}
