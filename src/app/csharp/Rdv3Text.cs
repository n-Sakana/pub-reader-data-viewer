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
    public const string LabelSearchBox = "番号1（{n}桁）を入力";
    public const string LabelProcessed = "処理済み";
    public const string NotepadNone = "未接続 — メモ帳を開くと自動で接続します";
    // {name} is the watch target from the settings file, so the screen says
    // what is actually being watched rather than always saying "notepad"
    public const string WatchNoneFmt = "未接続 — {name} を開くと自動で接続します";
    public const string StateWaitingFmt = "{name} を待機中";
    public const string LabelWatch = "監視";
    // "no targets" is a SETTING, not a window that has yet to appear. Saying
    // "waiting for Notepad" for a config that deliberately watches nothing sent
    // the operator looking for an application the file never asked for.
    public const string WatchNoTarget = "監視対象がありません (設定で追加してください)";
    public const string StateNoTarget = "監視対象なし";
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
    // the same row, said as a completion. "設定を保存しました" and "処理済みの保存が
    // 完了しました" were both going out tagged エラー, which reads as "it did not
    // work" for two things that did.
    public const string VerdictNote = "完了";

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
    public const string ErrSaveInFlight = "処理済みを保存中です。確定するまで次の操作はできません";
    public const string ErrCloseWhileSaving = "処理済みを保存中です。確定するまで終了できません";
    public const string ErrSaveOverdue = "処理済みの保存が想定より長引いています (確定するまで終了しません)";
    public const string CloseBlockedTitle = "保存中のため終了できません";
    public const string CloseBlockedBody = "処理済みの保存がまだ確定していません。\r\n保存が成功または失敗として確定してから、もう一度閉じてください。";
    public const string NoteSaveDoneCanClose = "処理済みの保存が完了しました。終了できます";
    public const string NoteSaveFailedCanClose = "処理済みの保存は失敗として確定しました。終了できます";
    public const string ErrBadKey = "番号1 は {n} 桁の数字で入力してください";
    public const string ErrBadKeyAny = "番号1 は {n} 文字で入力してください";
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

    // ---- 設定モーダル (docs\settings.md) ----------------------------------
    public const string BtnSettings = "設定";
    public const string TipSettings = "監視対象や動作の設定を変更します";
    public const string SettingsTitle = "設定";
    public const string SecTargets = "監視対象";
    public const string SecTarget = "対象の指定";
    public const string SecWindow = "ウィンドウ";
    public const string SecField = "欄";
    public const string SecMisc = "そのほか";
    public const string SecKey = "番号";
    public const string SecWatch = "監視";
    public const string SecBehaviour = "動作";
    public const string SecPaths = "ファイル";
    public const string BtnAdd = "追加";
    public const string BtnCopy = "複製";
    public const string BtnRemove = "削除";
    public const string BtnInspect = "画面から選ぶ";
    public const string PickTitle = "画面から選ぶ";
    public const string PickHow = "対象の欄にカーソルを合わせて  Ctrl + Shift  を押す";
    public const string PickEsc = "Esc で中止";
    public const string PickTaken = "取り込みました";
    public const string PickNothing = "カーソルの下に要素がありません";
    public const string PickReading = "読み取り";
    public const string PickNoRead = "この要素からは値を読めません";
    public const string BtnSave = "保存";
    public const string BtnCancel = "取消";
    public const string BtnClose = "閉じる";
    public const string BtnRefresh = "再取得";
    public const string BtnUseElement = "この要素を対象にする";
    public const string LblName = "名前";
    public const string LblEnabled = "有効";
    public const string LblWindow = "ウィンドウ";
    public const string LblPath = "中間パス";
    public const string LblField = "欄";
    public const string LblRead = "値の取り方";
    public const string LblAutomationId = "AutomationId";
    public const string LblClassName = "クラス名";
    public const string LblNameLike = "名前 (* ?)";
    public const string LblProcess = "プロセス名";
    public const string LblControlTypes = "種類";
    public const string LblIndex = "何番目";
    public const string LblScope = "探索範囲";
    public const string LblScopeChildren = "直下のみ";
    public const string LblScopeDesc = "子孫すべて";
    public const string LblUnitMs = "ms";
    public const string LblUnitRows = "行";
    public const string LblUnitDigits = "桁";
    public const string NoteTargetSummary = "{win} → {field}";
    public const string LblKeyLen = "桁数";
    public const string LblDigitsOnly = "数字のみ";
    public const string LblPollMs = "監視間隔 (ms)";
    public const string LblStableMs = "確定待ち (ms)";
    public const string LblRebindMs = "再接続間隔 (ms)";
    public const string LblPreferFocus = "前面のウィンドウを優先";
    public const string LblCandRows = "候補の表示行数";
    public const string LblDataDir = "データ (CSV) フォルダー";
    public const string LblLedger = "統合台帳";
    public const string LblLog = "ログ";
    public const string NoteRestart = "ファイルの変更は次回の起動から有効になります";
    public const string NoteSavedTo = "保存先";
    public const string NoteInspectHint = "対象の画面を開いてから「再取得」。要素を選んで下のボタンへ。";
    public const string NoteNoTarget = "監視対象がありません。「追加」か「画面を調べる」で指定してください。";
    public const string ErrSettingsSave = "設定を保存できませんでした: ";
    public const string ErrSettingsRead = "設定を読めませんでした (既定値で動作します): ";
    public const string NoteSettingsApplied = "設定を保存しました";
    public const string ColElement = "要素";
    public const string LblPatterns = "取得可能";
    public const string LblProcessOf = "プロセス";

    // ---- the settings modal, rebuilt to the ver4 reference ------------------
    // Section headings carry the question they answer, so a heading is worth
    // reading: "ウィンドウ" alone says nothing the fields below do not.
    public const string SetSubtitle = "監視対象や動作の設定を変更します";
    public const string SetMetaSchema = " ・ SCHEMA ";
    public const string SecTargetList = "対象リスト";
    public const string SecWindowLong = "ウィンドウ — どの画面か";
    public const string SecFieldLong = "欄 — どの入力欄を読むか";
    public const string SecKeyLong = "番号 — 何を検索キーとするか";
    public const string SecWatchLong = "監視 — 読み取りのタイミング";
    public const string SecFilesLong = "ファイル — 場所の指定";
    public const string NoteKeyForm = "この形式に一致した入力だけを番号として確定し、検索します。";
    public const string NoteWatchTiming = "確定待ちの間、値が変化しなければ検索を実行します。対象を見失ったときは再接続間隔ごとに探し直します。";
    public const string NoteFilesBase = "相対パスはプログラムと同じフォルダーが基準です。ファイルの変更は次回の起動から有効になります。";
    public const string HintPick = "対象の欄にカーソルを合わせて Ctrl + Shift で取り込み。Esc で中止します。";
    public const string LblStepNone = "なし — ウィンドウ直下から探索します";
    public const string LblValuePattern = "ValuePattern 必須";
    public const string LblDisabled = "無効";
    public const string TagTopMost = "最前面";
    public const string NoteNoTargetShort = "監視対象がありません";
}
