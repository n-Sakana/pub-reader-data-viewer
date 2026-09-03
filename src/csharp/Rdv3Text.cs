// ============================================================================
// Rdv3Text.cs -- the operator-facing strings the PROGRAM owns.
//
// Everything that names a part of the screen (section titles, row labels,
// button captions, state and result names) lives in settings.json ("screen").
// What stays here is what the program says on its own behalf: the words of its
// state machine, its notices and errors, the fixed modals (update check,
// settings, picker), the reasons it refuses data, and a few placeholders.
// "{...}" markers are filled in by the caller.
//
// The only file here with non-ASCII in it. build\pack_app.ps1 rewrites each
// non-ASCII character as \uXXXX on the way into the .cmd, because Add-Type
// hands the source to the in-box csc, which reads a BOM-less file in the
// system ANSI code page. No verbatim strings anywhere in the packed sources.
// ============================================================================

public static class Rdv3Text
{
    public const string AppTitle = "Reader Data Viewer";

    // ---- the state word in the status bar -----------------------------------
    public const string StateBoot = "起動中";
    public const string StateChecking = "更新を確認中";
    public const string StateApplying = "台帳を更新中";
    public const string StateDeleting = "レコードを削除中";
    public const string StateSending = "送信中";
    public const string StateReloading = "台帳を読み直し中";
    public const string StateLockWaiting = "台帳が空くのを待機中";
    public const string StateReady = "監視中";
    public const string StateSavingFmt = "{state}を保存中";
    public const string StateBlocked = "台帳がありません";
    public const string StateWaitingFmt = "{name} を待機中";
    public const string StateNoTarget = "監視対象なし";

    // ---- the watch segment ----------------------------------------------------
    public const string LabelNotepad = "メモ帳";
    public const string LabelWatch = "監視";
    public const string WatchConnectedFmt = "接続中（{title}）";
    public const string WatchNone = "未接続";
    public const string WatchNoTarget = "監視対象がありません (設定で追加してください)";

    // ---- the ledger segment ---------------------------------------------------
    public const string LedgerSegFmt = "{file} / {n} 件";
    public const string PendingCountFmt = "未送信 {n} 件";
    public const string NotYet = "--";
    public const string MsUnit = " ms";

    // ---- notices (the 完了 toast) --------------------------------------------
    public const string NoteNoDiff = "更新はありません (台帳は最新です)";
    public const string NoteUpdated = "台帳を更新しました";
    public const string NoteRejected = "更新を見送りました (保存済み台帳のまま)";
    public const string NoteSettingsApplied = "設定を保存しました";
    public const string NoteSaveDoneCanClose = "状態の保存が完了しました。終了できます";
    public const string NoteSaveFailedCanClose = "状態の保存は失敗として確定しました。終了できます";
    public const string NoteStateSaved = "{state}として保存しました";
    public const string NoteSendDone = "{n} 件を送信しました";
    public const string NoteDeleteDone = "{n} 件を削除しました";
    public const string NoteNoPending = "未送信の変更はありません";
    public const string NoteNotFound = "見つかりません";
    public const string TagDone = "完了";
    public const string TagError = "エラー";

    // ---- confirmations ---------------------------------------------------------
    public const string ConfirmUpdateTitle = "更新の確認";
    public const string ConfirmRebuildBody = "保存済みの統合台帳が読めません:\n{err}\nCSV から作り直しますか? (作業状態は失われます)";
    public const string ConfirmCreateBody = "保存済みの統合台帳がありません。CSV から新しく作成しますか?";
    public const string ConfirmStateTitleFmt = "{state}の確認";
    public const string SendTitle = "送信";
    public const string ConfirmSendBody = "未送信の {n} 件を送信します。よろしいですか?\n\n送信した行は未送信から外れます。取り込んだデータは書き換えません。";
    public const string BtnYes = "はい";
    public const string BtnNo = "いいえ";

    public static string UpdateConfirmBody(string onSourceChange, string initialState)
    {
        string first = "定義された処理で台帳に変更があります。更新しますか?\n";
        if (onSourceChange == "preserve")
        {
            return first + "(入力側の列が変わっても、作業状態は現在値を保ちます)";
        }
        return first + "(入力側の列が変わった行は、作業状態を「" + initialState + "」へ戻します)";
    }

    // ---- errors (the エラー toast) ------------------------------------------
    public const string ErrNo64 = "64 ビットのプロセスが必要です。";
    public const string ErrAlreadyRunning = "同じ統合台帳を開いている Reader Data Viewer が、すでに起動しています。";
    public const string ErrNoData = "CSV が見つかりません: ";
    public const string ErrBadLedgerPath = "統合台帳のパスが不正です: ";
    public const string ErrNoLedger = "統合台帳がありません。検索できません。";
    public const string ErrCheckFailed = "更新確認に失敗しました: ";
    public const string ErrCheckTimeout = "更新確認がタイムアウトしました (保存済み台帳のまま続行します)";
    public const string ErrPersist = "台帳を書き込めませんでした: ";
    public const string ErrLedgerRead = "台帳を読み込めませんでした: ";
    public const string ErrOracle = "検算 NG: 統合結果が expected.txt と一致しません";
    public const string ErrNotReady = "更新確認が終わるまで操作できません";
    public const string ErrSaveInFlight = "状態を保存中です。確定するまで次の操作はできません";
    public const string ErrCloseWhileSaving = "状態を保存中です。確定するまで終了できません";
    public const string ErrCloseWhileWriting = "書き込み中です。結果が確定するまで終了できません";
    public const string ErrSaveOverdue = "状態の保存が想定より長引いています (確定するまで終了しません)";
    public const string ErrBadKeyFmt = "番号1 が形式 {pattern} に一致しません";
    public const string ErrBadPattern = "番号の形式（正規表現）が不正です。設定を確認してください";
    public const string ErrNoRecordShown = "{state}にするレコードが表示されていません";
    public const string ErrNoTransition = "このレコードはすでに{state}です";
    public const string ErrUnknownState = "台帳の状態 {stored} は画面定義にありません。遷移できません";
    public const string ErrSettingsSave = "設定を保存できませんでした: ";
    public const string ErrPatternTyped = "形式（正規表現）が不正です: ";
    public const string ErrPathBlank = "パスは 3 つとも必要です";
    public const string ErrDataDir = "データフォルダーが見つかりません: ";
    public const string ErrLogWrite = "ログを書けません: ";
    public const string ErrExport = "CSV を出力できませんでした: ";
    public const string ErrDelete = "レコードを削除できませんでした: ";
    public const string ErrPendingRead = "手元の未送信データを読めません: ";
    public const string ErrPendingWrite = "手元の未送信データを書けませんでした: ";
    public const string ErrSend = "送信できませんでした: ";
    public const string ErrSendOverdue = "送信が長引いています。台帳が空くまで待っています";
    public const string ErrSharedWriteOverdue = "台帳の書き込みが長引いています。結果が確定するまで待っています";
    public const string ErrReloadOverdue = "台帳の読み直しが長引いています";
    public const string ErrSharedMarker = "台帳の更新通知を読めませんでした: ";
    public const string ErrSharedReload = "更新された台帳を読み直せませんでした: ";
    public const string LockWaitingFmt = "{user}（{host}）が {minutes} 分前から台帳を使用しています。空くまで待ちます。";

    // ---- the app does not start / cannot go on (one modal, then it stops) ----
    public const string FatalTitle = "起動できません";
    public const string FatalSettings = "設定ファイルに問題があるため起動できません。\n\nファイル: {file}\n{reason}\n\n直してから起動し直してください。";
    public const string FatalDataTitle = "データを読めません";
    public const string FatalData = "データに問題があるため続行できません。\n\n{reason}\n\n直してから起動し直してください。";

    // ---- why a CSV or the ledger is refused ({file} / {row} / {name} ...) ----
    public const string DataNoRows = "{file}: ヘッダー行とデータ行がありません";
    public const string DataBlankHeader = "{file}: ヘッダー行に空の列名があります";
    public const string DataDupHeader = "{file}: 列名 {name} がヘッダー行に 2 回あります";
    public const string DataNoColumn = "{file} に列 {name} がありません";
    public const string DataQuoted = "{file} の {row} 行目: 引用符 (\") で始まる列があります。引用符付きの CSV は読めません";
    public const string DataColumnCount = "{file} の {row} 行目: 列数が {n} です (ヘッダー行は {cols} 列)";
    public const string DataEmptyKey = "{file} の {row} 行目: キー列 {name} が空です";
    public const string DataKeyNotAscii = "{file} の {row} 行目: キー列 {name} に ASCII 以外の文字があります";
    public const string DataKeyWidth = "{file} の {row} 行目: キー列 {name} の幅が 1 行目 ({n} 文字) と違います";
    public const string DataDupKey = "{file}: キー列 {name} の値 {key} が {row1} 行目と {row2} 行目にあります (キーは一意である必要があります)";
    public const string DataLedgerHeader = "{file} の見出し行が、作業状態の列と画面定義の台帳列に一致しません";

    // ---- placeholders and fixed words on the screen ---------------------------
    public const string PanelCand = "候補一覧";
    public const string CandCount = "候補 {n} 件";
    public const string SubMulti = "複数ヒット（候補 {n} 件）";
    public const string FieldUnresolved = "列なし";
    public const string StateBlank = "(空)";
    public const string JudgeUndefined = "未定義";
    public const string JudgeError = "エラー";
    public const string SavingSuffix = " (保存中...)";
    public const string BtnClose = "閉じる";
    public const string BtnOk = "OK";
    public const string BtnBrowse = "参照...";
    public const string Dash = "—";
    public const string Unsearched = "未検索";
    public const string CandidateHitsFmt = "該当 {n} 件";

    // ---- shared-ledger notices ------------------------------------------------
    public const string SharedSendTitle = "台帳が更新されました";
    public const string SharedSendBody = "{user} が {done} 件を処理済、{todo} 件を未処理にしました";
    public const string SharedUpdateTitle = "台帳の更新";
    public const string SharedUpdateBody = "台帳が更新されました。切り替えますか";
    public const string SharedResetFmt = "中身が変わったため未処理に戻ったレコード: {n} 件";
    public const string UnmatchedTitle = "送信できなかったレコード";
    public const string UnmatchedBodyFmt = "共有台帳へ当てられなかった変更が {n} 件あります。未送信のまま残しました。";
    public const string UnmatchedMissing = "行がありません";
    public const string UnmatchedChanged = "中身が変わっています";
    public const string ColReason = "理由";

    // ---- process job dialogs --------------------------------------------------
    public const string UpdateRecordsTitle = "レコード更新";
    public const string DeleteRecordsTitle = "レコード削除";
    public const string UpdateRecordsHint = "定義されたファイルだけを読み、JSON の手順で統合台帳を更新します。";
    public const string DeleteRecordsHint = "定義されたファイルだけを読み、条件に一致する行を統合台帳から取り除きます。";
    public const string SecInputs = "取り込むデータ（data＼）";
    public const string SecProcess = "処理内容";
    public const string SecOutput = "書き出し先";
    public const string ColInput = "表";
    public const string ColDeleteInput = "指定";
    public const string ColFile = "ファイル";
    public const string ColKey = "キー";
    public const string ColRows = "行数";
    public const string ColValidation = "検証";
    public const string ColNumber = "#";
    public const string ColOperation = "操作";
    public const string ColTarget1 = "対象1";
    public const string ColTarget2 = "対象2";
    public const string ColCondition = "条件";
    public const string ColOutput = "出力";
    public const string ValidationColumnsMatch = "列一致";
    public const string ValidationMissing = "ファイルなし";
    public const string ValidationError = "不一致";
    public const string LblPath = "パス";
    public const string LblFileName = "ファイル名";
    public const string LblLastWrite = "最終更新";
    public const string LblNeverWritten = "未作成";
    public const string BtnExecute = "実行";
    public const string BtnDelete = "削除する";
    public const string ProcessNotRun = "入力ファイルを確認できないため実行できません。";

    public static string OperationLabel(string operation)
    {
        switch (operation)
        {
            case "join": return "結合";
            case "extract": return "抽出";
            case "delete": return "削除";
            case "append": return "追加";
            case "update": return "更新";
            case "merge": return "マージ";
            case "replace": return "置換";
            case "select": return "選択";
            case "calculate": return "計算";
            case "aggregate": return "集計";
            case "sort": return "並べ替え";
            case "distinct": return "重複除去";
        }
        return operation;
    }

    public static string MergeDestinations(string sourceOnly, string both, string targetOnly)
    {
        return "元のみ:" + RowDestination(sourceOnly)
            + " / 両方:" + RowDestination(both)
            + " / 先のみ:" + RowDestination(targetOnly);
    }

    private static string RowDestination(string value)
    {
        switch (value)
        {
            case "add": return "追加";
            case "ignore": return "無視";
            case "update": return "更新";
            case "keep": return "保持";
            case "delete": return "削除";
        }
        return value;
    }

    public static string ConditionLabel(string condition)
    {
        switch (condition)
        {
            case "match": return "一致";
            case "either": return "どちらか";
            case "both": return "両方";
            case "exclude": return "除く";
            case "content": return "内容も一致";
        }
        return (condition == null || condition.Length == 0) ? Dash : condition;
    }

    public static string JoinConditionLabel(string condition)
    {
        switch (condition)
        {
            case "match": return "内部";
            case "left": return "左外部";
            case "full": return "完全外部";
        }
        return ConditionLabel(condition);
    }

    public static string PredicateLabel(string operation)
    {
        switch (operation)
        {
            case "equals": return "等しい";
            case "notEquals": return "等しくない";
            case "contains": return "含む";
            case "startsWith": return "で始まる";
            case "endsWith": return "で終わる";
            case "empty": return "空";
            case "notEmpty": return "空でない";
            case "greater": return "より大きい";
            case "atLeast": return "以上";
            case "less": return "より小さい";
            case "atMost": return "以下";
        }
        return operation;
    }

    public static string AggregateLabel(string function)
    {
        if (function == "sum") { return "合計"; }
        if (function == "count") { return "件数"; }
        return function;
    }

    public static string DirectionLabel(string direction)
    {
        if (direction == "ascending") { return "昇順"; }
        if (direction == "descending") { return "降順"; }
        return direction;
    }

    public static string SortTypeLabel(string type)
    {
        if (type == "text") { return "文字"; }
        if (type == "number") { return "数値"; }
        return type;
    }

    // ---- table export ---------------------------------------------------------
    public const string ExportTitle = "テーブル出力";
    public const string ExportHint = "統合台帳から、選んだ項目だけを CSV に書き出します。";
    public const string ExportAvailable = "出力できる項目";
    public const string ExportSelectedFmt = "出力する項目（{n}）";
    public const string ExportDefault = "既定に戻す";
    public const string ExportDestination = "出力先";
    public const string ExportDefaultPath = "data\\export-{yyyyMMdd-HHmmss}.csv";
    public const string ExportNeedField = "出力する項目を 1 つ以上選んでください。";
    public const string ExportDoneFmt = "CSV を出力しました: {file}";
    public const string BtnMoveRight = "▶";
    public const string BtnMoveLeft = "◀";
    public const string WorkStateColumn = "処理済";

    // ---- the settings modal -----------------------------------------------------
    public const string SettingsTitle = "設定";
    public const string SettingsHint = "書き戻すのは paths / search / watch の 3 つだけです。";
    public const string SecPlaces = "場所";
    public const string SecSearch = "検索";
    public const string LblDataShort = "データ";
    public const string LblKeyPatternShort = "番号の形式";
    public const string LblCandidateRows = "候補の表示件数";
    public const string LblTarget = "対象";
    public const string LblRead = "読み取り";
    public const string ReadSummaryFmt = "{mode} / {poll}ms 間隔";
    public const string ReadValuePattern = "ValuePattern";
    public const string ReadTextPattern = "TextPattern";
    public const string ReadNameProperty = "Name property";
    public const string SecTargets = "監視対象";
    public const string SecTarget = "対象の指定";
    public const string BtnInspect = "画面から選ぶ";
    public const string NoteReadTarget = "読み取り対象:  {sum}（「画面から選ぶ」で自動設定されます）";
    public const string NoteTargetSummary = "{name}（{kind} = {value}）";
    public const string NoteNoTargetShort = "監視対象がありません";
    public const string LblDisabled = "（無効）";
    public const string SecKeyPattern = "形式（正規表現）";
    public const string NoteKeyPattern = "入力がこの正規表現に一致したときだけ、番号として確定し検索します。";
    public const string SecPaths = "ファイル";
    public const string LblDataDir = "データ (CSV) フォルダー";
    public const string LblLedger = "統合台帳";
    public const string LblLog = "ログ";
    public const string NoteFilesBase = "相対パスはプログラムと同じフォルダーが基準です。変更は次回の起動から有効になります。";
    public const string BtnSave = "保存";
    public const string BtnCancel = "キャンセル";
    public const string NoValue = "N/A";

    // ---- the element picker ------------------------------------------------------
    public const string PickTitle = "画面から選ぶ";
    public const string PickHow = "対象の欄にカーソルを合わせて  Ctrl + Shift  を押す";
    public const string PickEsc = "Esc で中止";
    public const string TagTopMost = "最前面";
    public const string LblControlTypes = "種類";
    public const string LblAutomationId = "AutomationId";
    public const string LblClassName = "クラス名";
    public const string LblName = "名前";
    public const string LblProcessOf = "プロセス";
    public const string PickReading = "読み取り";
    public const string PickNoRead = "この要素からは値を読めません";
}
