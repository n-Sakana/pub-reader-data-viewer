// ============================================================================
// Rdv3Text.cs -- every string the operator reads in the practical C# build.
//
// The only file here with non-ASCII in it (Rdv3Xlsx carries one header word of
// the ledger file format on purpose). build\pack_app.ps1 rewrites each
// non-ASCII character as \uXXXX on the way into the .cmd, because Add-Type
// hands the source to the in-box csc, which reads a BOM-less file in the
// system ANSI code page. No verbatim strings anywhere in the packed sources.
// ============================================================================

public static class Rdv3Text
{
    public const string AppTitle = "Reader Data Viewer";
    public const string Method = "C# 標準 Dictionary";

    public const string StateBoot = "起動中";
    public const string StateChecking = "更新を確認中";
    public const string StateApplying = "台帳を更新中";
    public const string StateReady = "監視中";
    public const string StateWaitingNotepad = "メモ帳を待機中";
    public const string StateBusy = "検索中";
    public const string StateSavingMark = "処理済みを保存中";
    public const string StateBlocked = "台帳がありません";

    public const string NoteNoDiff = "更新はありません (台帳は最新です)";
    public const string NoteUpdated = "台帳を更新しました";
    public const string NoteRejected = "更新を見送りました (保存済み台帳のまま)";

    public const string ConfirmUpdateTitle = "更新の確認";
    public const string ConfirmUpdateBody = "CSV に変更があります。統合台帳を更新しますか?\r\n(処理済みは変更のないレコードへ引き継がれます)";
    public const string ConfirmRebuildBody = "保存済みの統合台帳が読めません:\r\n{err}\r\nCSV から作り直しますか? (処理済み状態は失われます)";
    public const string ConfirmCreateBody = "保存済みの統合台帳がありません。CSV から新しく作成しますか?";
    public const string ConfirmProcessedTitle = "処理済みの確認";
    public const string ConfirmProcessedBody = "表示中の統合レコード (番号2 = {key2}) を処理済みにします。よろしいですか?";

    public const string LabelState = "状態";
    public const string LabelNotepad = "メモ帳";
    public const string LabelLedger = "台帳";
    public const string LabelMergeMs = "マージ時間";
    public const string LabelSearchMs = "検索時間";
    public const string LabelKey = "現在の番号 (番号1)";
    public const string LabelSearchBox = "番号1（8桁）を入力";
    public const string LabelProcessed = "処理済み";
    public const string NotepadNone = "未接続 — メモ帳を開くと自動で接続します";
    public const string LedgerRows = "{n} 件";

    public const string BtnSearch = "検索";
    public const string BtnClear = "内容クリア";
    public const string BtnProcessed = "処理済み";
    public const string BtnRebind = "メモ帳を再検出";
    public const string TipSearch = "入力した番号1 で統合台帳を検索します";
    public const string TipClear = "入力と結果表示を消します (台帳と処理済みは消えません)";
    public const string TipProcessed = "表示中の統合レコードを処理済みにします (確認あり)";

    public const string BoxRecord = "統合レコード (A-B-C 統合 1 件)";
    public const string BoxCand = "候補一覧 -- 行をクリックすると表示します";

    public const string ColFieldA = "表A 項目";
    public const string ColValueA = "表A 値";
    public const string ColFieldB = "表B 項目";
    public const string ColValueB = "表B 値";
    public const string ColFieldC = "表C 項目";
    public const string ColValueC = "表C 値";

    public const string PickNo = "#";
    public const string PickKey2 = "番号2";
    public const string PickLine = "行番号";
    public const string PickSlip = "伝票番号";
    public const string PickDate = "日付";
    public const string PickQty = "数量";
    public const string PickStatus = "状態";
    public const string PickItem = "品目コード";
    public const string PickMaker = "メーカー";

    public const string VerdictOne = "一致 1 件   番号2 = {key2}";
    public const string VerdictPicked = "一致 1 件   番号2 = {key2}   (候補 {n} 件中 {i} 件目)";
    public const string VerdictMany = "候補 {n} 件 — 一覧から 1 件選んでください";
    public const string VerdictManyCut = "候補 {n} 件中 {m} 件を表示 -- 一覧から 1 件選んでください";
    public const string VerdictNone = "該当なし";
    public const string VerdictErr = "エラー";

    public const string ErrNo64 = "64 ビットのプロセスが必要です。";
    public const string ErrNoData = "CSV が見つかりません: ";
    public const string ErrNoLedger = "統合台帳がありません。検索できません。";
    public const string ErrCheckFailed = "更新確認に失敗しました: ";
    public const string ErrCheckTimeout = "更新確認がタイムアウトしました (保存済み台帳のまま続行します)";
    public const string ErrPersist = "台帳を書き込めませんでした: ";
    public const string ErrLedgerRead = "台帳を読み込めませんでした: ";
    public const string ErrOracle = "検算 NG: 統合結果が expected.txt と一致しません";
    public const string ErrNotReady = "更新確認が終わるまで操作できません";
    public const string ErrSaveInFlight = "処理済みを保存中です。確定するまで次の操作はできません";
    public const string ErrCloseWhileSaving = "処理済みを保存中です。確定するまで終了できません";
    public const string ErrSaveOverdue = "処理済みの保存が想定より長引いています (確定するまで終了しません)";
    public const string CloseBlockedTitle = "保存中のため終了できません";
    public const string CloseBlockedBody = "処理済みの保存がまだ確定していません。\r\n保存が成功または失敗として確定してから、もう一度閉じてください。";
    public const string NoteSaveDoneCanClose = "処理済みの保存が完了しました。終了できます";
    public const string NoteSaveFailedCanClose = "処理済みの保存は失敗として確定しました。終了できます";
    public const string ErrBadKey = "番号1 は 8 桁の数字で入力してください";
    public const string ErrNoRecordShown = "処理済みにする統合レコードが表示されていません";

    // Nothing chosen yet is BLANK: the sub-line already says so, and a row of
    // dashes reads as noise. A field that IS on screen but empty in the ledger
    // says so in words, the way a spreadsheet does.
    public const string NoValue = "N/A";

    public const string MsUnit = " ms";
    public const string NotYet = "--";
    public const string SavingSuffix = " (保存中...)";

    // ---- the reference screen (docs\ui-spec.md) ----------------------------
    public const string TagLedger = "統合台帳";
    public const string LabelKeyNow = "現在の番号（番号1）";
    public const string LabelRepStatus = "代表ステータス";
    public const string LabelLedgerRows = "台帳総件数";
    public const string LabelLastSaved = "最終更新";
    public const string UnitRows = "件";
    public const string LabelUnprocessed = "未処理";
    public const string RoleNormal = "一般";
    public const string RoleReadOnly = "読み取り専用";

    public const string PanelCand = "候補一覧";
    public const string HintCand = "行をクリックすると統合レコードを表示します";
    public const string PanelRec = "統合レコード";
    public const string HintRec = "選択中の候補に紐づく 表A（取引先）の情報と 表B/表C の長文項目";
    public const string CandCount = "候補 {n} 件";
    public const string TagKey2 = "番号2";
    public const string LabelMemo = "摘要（表B: b_memo）";
    public const string LabelRemark = "備考（表C: c_remark）";
    public const string PickToSee = "候補一覧から行を選択すると表示されます。";

    public const string SubIdle = "未検索";
    public const string SubNone = "該当なし";
    public const string SubUnselected = "未選択（候補 {n} 件）";
    public const string SubOne = "一致 1 件 ・ 番号2 = {key2}";
    public const string SubPicked = "候補 {i}/{n} ・ 番号2 = {key2}";
    public const string SubMulti = "複数ヒット（候補 {n} 件）";
}
