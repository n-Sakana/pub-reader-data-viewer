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
    public const string LedgerSegFmt = "{file} ・ {n} 件";
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
    public const string TagDone = "完了";
    public const string TagError = "エラー";

    // ---- confirmations ---------------------------------------------------------
    public const string ConfirmUpdateTitle = "更新の確認";
    public const string ConfirmUpdateBody = "CSV に変更があります。統合台帳を更新しますか?\n(作業状態は変更のないレコードへ引き継がれます)";
    public const string ConfirmRebuildBody = "保存済みの統合台帳が読めません:\n{err}\nCSV から作り直しますか? (作業状態は失われます)";
    public const string ConfirmCreateBody = "保存済みの統合台帳がありません。CSV から新しく作成しますか?";
    public const string ConfirmStateTitleFmt = "{state}の確認";
    public const string BtnYes = "はい";
    public const string BtnNo = "いいえ";

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

    // ---- the settings modal -----------------------------------------------------
    public const string SettingsTitle = "設定";
    public const string SecTargets = "監視対象";
    public const string SecTarget = "対象の指定";
    public const string BtnInspect = "画面から選ぶ";
    public const string NoteReadTarget = "読み取り対象:  {sum}（「画面から選ぶ」で自動設定されます）";
    public const string NoteTargetSummary = "{win} → {field}";
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
    public const string BtnCancel = "取消";
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
