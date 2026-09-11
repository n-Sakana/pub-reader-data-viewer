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
    public const string MigrationNeedsConfirmation = "旧台帳の作成時設定であることを確認し、-ConfirmOriginalDefinition を指定してください。旧台帳には結合定義の履歴がなく、アプリだけでは証明できません。";
    public const string MigrationDefinitionMismatch = "作成時設定と移行先設定の入力・結合・保存・削除定義が一致しません。旧台帳と移行先は変更していません。";
    public const string MigrationBusy = "旧台帳をほかの端末が処理中です。全端末の処理と未送信変更を終えてから移行してください。";
    public const string MigrationAlreadyBound = "この台帳は既に定義と削除保管情報を持っています。旧台帳の移行対象ではありません。";
    public const string MigrationDone = "旧台帳の全内容・状態を保持して別ファイルへ移しました。移行先: ";
    public const string LegacyLedgerNeedsMigration = "この旧台帳には業務定義の記録がありません。自動で現在の設定を登録せず、書込みを停止しました。作成時の設定を用意し、Migrate-Ledger.ps1 で別名の台帳へ移行してください。元の台帳と状態は保持しています。";
    public const string BusinessDefinitionMismatch = "台帳作成時と入力形式・結合・保存・削除の定義が違います。同じ台帳への書込みはできません。作成時の設定へ戻すか、別の台帳を指定してください。";
    public const string ProtectionInvalid = "台帳の定義または削除保管データが不正です。書込みを停止しました。正常な控えから台帳を復旧してください。";
    public const string ArchiveIdentityConflict = "復元対象が変更されたか、同じ識別キーが既に通常台帳にあります。上書きせず停止しました。台帳を読み直して確認してください。";
    public const string ArchiveTitle = "削除済みレコード";
    public const string OpRestore = "復元";
    public const string ArchiveHint = "削除時の内容と確認状態を保管しています。選択したレコードだけを復元します。";
    public const string RestoreDone = "{0} 件を削除時の内容と状態で復元しました。";
    public const string RestoreConfirm = "選択した {0} 件を、削除時の内容と確認状態で共有台帳へ復元します。よろしいですか。";
    public const string DeleteStateHint = "\u78BA\u8A8D\u72B6\u614B\u306F\u5171\u6709\u53F0\u5E33\u306B\u4FDD\u5B58\u6E08\u307F\u306E\u5024\u3067\u5224\u5B9A\u3057\u307E\u3059\u3002\u672A\u9001\u4FE1\u306E\u5909\u66F4\u306F\u9001\u4FE1\u3057\u3066\u304B\u3089\u524A\u9664\u3057\u3066\u304F\u3060\u3055\u3044\u3002";
    public const string RecordXmlValue = "値「{0}」には XLSX に保存できない文字があります。";
    public static string Format(string text, params object[] values)
    { return string.Format(System.Globalization.CultureInfo.InvariantCulture, text, values); }

    public const string SourceRow = "{0} の {1} 行目";
    public const string RecordExcluded = "1 行を除外しました: {0}。{1}";
    public const string RecordColumns = "列数が見出しと違います。必要な列数 {0}、実際の列数 {1}。";
    public const string RecordControlKey = "キー列 {0} の値「{1}」に制御文字があります。";
    public const string RecordEmptyKey = "キー列 {0} の値が空です。";
    public const string RecordDuplicate = "{0}: キー列 {1}、値「{2}」の {3} 行を除外しました（元の行番号: {4}）。{5}";
    public const string RecordConflict = "同じキーで内容が違うため、どの行も採用しません。キーに枝番を付けるか、入力データを訂正してください。複数の明細を集計する場合は、一意な明細key（複合keyも可）で全件を読んでから、aggregateのgroupByで集めます。入力キーの検査はaggregateより先です。duplicates:distinctは後続行を捨て、合計しません。";
    public const string RecordIdentical = "同じキーの重複行です。";
    public const string RecordStep = "手順 {0}（{1} → {2}）、{3}";
    public const string RecordValues = "列と値: {0}";
    public const string RecordNumber = "数値として読めない値「{0}」です。";
    public const string RecordDivisionZero = "0 で割る計算です（左の値「{0}」、右の値「{1}」）。";
    public const string RecordOverflow = "計算結果が数値として扱える範囲を超えました。";
    public const string RecordFunction = "{0} の入力「{1}」から指定された文字列を取得できません（条件: {2}）。";
    public const string RecordRegexTimeout = "正規表現の照合が制限時間を超えました。";
    public const string RecordXlsxCell = "{0}、セル {1}: {2}";
    public const string RecordXlsxError = "エラー値「{0}」が入っています。元のブックの値を訂正してください。";
    public const string RecordXlsxFormula = "数式「{0}」の保存済み計算結果がありません。元のブックを再計算して保存してください。";
    public const string RecordXlsxLong = "セルの値が上限の 32767 文字を超えています（{0} 文字）。";
    public const string InputWarningLog = "入力の除外があります。対象の行と理由は実行ログを確認してください。";
    public const string SettingsLabelExample = "画面に表示する名前";
    public const string SettingsFixLabel = " settings.json の data.labels に {entry} を設定してください（表示名は用途に合わせ、空にしません）。中間結果と ledger も画面名が必要です。同じ名前の登録は1回で足ります。README.md の「処理に使う名前と画面名」を参照してください。";
    public const string SettingsFixTableLabel = " settings.json の data.tables[{id}].label に空でない画面名を設定してください。この表IDは予約済みなので、data.labels へ同じIDを追加できません。";
    public const string InputRowsSkipped = "{file}: キー列 {column} が空の {empty} 行と重複の {duplicate} 行を除き、{kept} 行を読みました。";
    public const string InputNoData = "{file}: 見出しだけを読みました。データ行は 0 件です。入力ファイルを確認してください。";
    public const string InputNoKeptData = "{file}: 除外後のデータ行は 0 件です。入力ファイルと除外件数を確認してください。";
    public const string InputError = "{file} の {row} 行目、列 {column}: 必要なのは「{expected}」、実際は「{actual}」です。{fix}";
    public const string InputExpectKey = "空でないキー";
    public const string InputExpectWidth = "最初の有効行と同じ {n} 文字のキー";
    public const string InputFixKey = "入力値を直してください。この形を許す定義なら、settings.json の {path} を \"{choice}\" にします。";
    public const string InputFixConflict = " 同じキーで他の列の内容が違うため、どちらを採るか決められません。両方の行を照合して内容を直すか、{path} の key（外部入力では column）を各行を識別できる列名に直してください。複数の明細を残す場合は、一意な明細keyで読み、aggregateのgroupByで集めます。入力キーの検査はaggregateより先です。duplicates:distinctは最初の行を残すだけで、合計しません。明細を識別できなければ、入力の変更が必要な理由を返してください。必要な行を捨てた設定を完成品にしないでください。";
    public const string InputFixType = " 値を指定した型・日付書式に直してください。数値・日付として扱わない列なら、settings.json の data.types[\"{ref}\"].type を \"text\" にします。";
    public const string InputFixEncoding = "settings.json の {setting} を \"{encoding}\" に合わせるか、入力ファイルを指定した文字コードで保存してください。";
    public const string InputFixUnknownEncoding = "このバイト列は指定した文字コードでは読めません。元ファイルの文字コードを確認して settings.json の {setting} に指定するか、ファイルをその文字コードで保存し直してください。";
    public const string InputUnknownEncoding = "文字コード「{value}」を利用できません。encoding に \"utf-8\"、\"shift_jis\"、\"utf-16\" など、この端末で利用できる文字コード名を指定してください。";
    public const string InputExpectHeader = "空でなく重複しない列名";
    public const string InputFixHeader = "Use non-empty source column names. Duplicate headers used by a key, type, job or ledger reference are ambiguous: rename them in the source and update the references in data.tables, data.jobs, data.types, data.ledger and screen. JSON labels or select cannot disambiguate repeated source headers. Unreferenced duplicate name groups are excluded automatically with a warning; empty cells do not make a referenced duplicate safe.";
    public const string InputShapeSkipped = "{file}: 見出しより列が少ない {short} 行と空行 {blank} 行を除外しました。不足セルの補完はしていません。";
    public const string InputHeaderOffset = "{file}: headerRow の指定により、見出し行より前の {n} 行を読み飛ばしました。";
    public const string DataNoDerivedColumn = "更新ジョブの結果に列 {name} がありません ({where})。入力表の見出し、または calculate の column / aggregate の as / select の as で作った参照を指定してください。";
    public const string LedgerColumnNotProduced = "台帳の保存列 {name} は、更新ジョブの最後の書込み段 ({step}) の入力にありません。その列を持つ表を結合していないか、select で外したか、綴りが違います。";
    public const string InputHeadersSkipped = "{file}: 設定から参照されていない重複見出し「{names}」の {n} 列をすべて除外しました。";
    public const string InputExpectQuote = "セル全体を囲む二重引用符と閉じ引用符";
    public const string InputExpectDelimiter = "閉じ引用符の後のカンマまたは改行";
    public const string InputColumnCount = "{n} 列";
    public const string InputFixCsv = "CSV の各行を見出しと同じ列数にしてください。セル内の区切り文字・改行はセル全体を二重引用符で囲み、セル内の二重引用符は2個重ねます。区切り文字は既定でカンマです（settings.json の delimiter で tab 等に変更できます）。";
    public const string InputFixTab = " 見出しにタブ文字があり、指定した区切り文字がありません。タブ区切りのファイル（Excel の「Unicode テキスト」等）なら、settings.json の data.tables.<ID>.delimiter（外部入力なら inputs[].delimiter）を \"tab\" にしてください。";
    public const string StorageContractMismatch = "\u3053\u306e\u53f0\u5e33\u3068\u8a2d\u5b9a\u306e\u5217\u30fb\u8b58\u5225\u30ad\u30fc\u30fb\u4f5c\u696d\u72b6\u614b\u306e\u5b9a\u7fa9\u304c\u4e00\u81f4\u3057\u307e\u305b\u3093\u3002\u540c\u3058\u53f0\u5e33\u3092\u4f7f\u3046\u5168PC\u306e\u5b9a\u7fa9\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002";

    public const string UpdateChangedDuringCheck = "\u78ba\u8a8d\u4e2d\u306b\u5171\u6709\u53f0\u5e33\u306e\u5185\u5bb9\u304c\u5909\u308f\u308a\u307e\u3057\u305f\u3002\u4e0a\u66f8\u304d\u305b\u305a\u4e2d\u6b62\u3057\u307e\u3057\u305f\u3002\u3082\u3046\u4e00\u5ea6\u300c\u66f4\u65b0\u300d\u3067\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002";
    public const string ErrSearchTimeout = "\u691c\u7d22\u304c\u6642\u9593\u5185\u306b\u5b8c\u4e86\u3057\u307e\u305b\u3093\u3067\u3057\u305f\u3002\u9045\u308c\u3066\u5c4a\u3044\u305f\u7d50\u679c\u306f\u8868\u793a\u3057\u307e\u305b\u3093\u3002";

    public const string SettingsChangedExternally = "\u8a2d\u5b9a\u30d5\u30a1\u30a4\u30eb\u304c\u5225\u306e\u64cd\u4f5c\u3067\u5909\u66f4\u3055\u308c\u3066\u3044\u307e\u3059\u3002\u753b\u9762\u3092\u958b\u304d\u76f4\u3057\u3066\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002";
    public const string StateConflict = "\u4ed6\u306e\u64cd\u4f5c\u3067\u5171\u6709\u306e\u72b6\u614b\u304c\u5909\u66f4\u3055\u308c\u307e\u3057\u305f\u3002";
    public const string LegacyPending = "\u65e7\u7248\u306e\u672a\u9001\u4fe1\u5909\u66f4\u3067\u3059\u3002\u5909\u66f4\u524d\u306e\u72b6\u614b\u3092\u78ba\u8a8d\u3067\u304d\u307e\u305b\u3093\u3002";
    public const string DiscardPendingConfirm = "\u3053\u306e\u4e00\u89a7\u306e\u672a\u9001\u4fe1\u5909\u66f4\u3092\u7834\u68c4\u3057\u307e\u3059\u304b\uff1f \u5171\u6709\u53f0\u5e33\u306e\u5185\u5bb9\u306f\u5909\u66f4\u3057\u307e\u305b\u3093\u3002\u5fc5\u8981\u306a\u5909\u66f4\u306f\u3001\u6700\u65b0\u306e\u5185\u5bb9\u3092\u78ba\u8a8d\u3057\u3066\u304b\u3089\u767b\u9332\u3057\u76f4\u3057\u3066\u304f\u3060\u3055\u3044\u3002";
    public const string DiscardPendingButton = "\u4e00\u89a7\u306e\u672a\u9001\u4fe1\u5909\u66f4\u3092\u7834\u68c4";
    public const string LedgerStateInvalid = "\u53f0\u5e33\u306b\u8a2d\u5b9a\u3067\u5b9a\u7fa9\u3055\u308c\u3066\u3044\u306a\u3044\u4f5c\u696d\u72b6\u614b\u304c\u3042\u308a\u307e\u3059\u3002\u884c ";
    public const string ExportExists = "\u540c\u540d\u30d5\u30a1\u30a4\u30eb\u3078\u306e\u4e0a\u66f8\u304d\u306f\u884c\u3044\u307e\u305b\u3093\u3002\u5225\u306e\u30d5\u30a1\u30a4\u30eb\u540d\u3092\u6307\u5b9a\u3057\u3066\u304f\u3060\u3055\u3044\u3002";
    public const string ExportProtected = "\u5165\u529b\u30fb\u53f0\u5e33\u30fb\u8a2d\u5b9a\u30fb\u30d7\u30ed\u30b0\u30e9\u30e0\u95a2\u9023\u306e\u30d5\u30a1\u30a4\u30eb\u306b\u306f\u51fa\u529b\u3067\u304d\u307e\u305b\u3093\u3002";
    public const string ExportCsvOnly = "\u51fa\u529b\u5148\u306b\u306f .csv \u30d5\u30a1\u30a4\u30eb\u3092\u6307\u5b9a\u3057\u3066\u304f\u3060\u3055\u3044\u3002";
    public const string StaleLedger = "\u8868\u793a\u306f\u66f4\u65b0\u524d\u306e\u53f0\u5e33\u3067\u3059\u3002\u300c\u66f4\u65b0\u300d\u3067\u6700\u65b0\u306e\u53f0\u5e33\u3092\u78ba\u8a8d\u3057\u3066\u304f\u3060\u3055\u3044\u3002";
    public const string ExcelSafe = "Excel\u5411\u3051\u306b\u6570\u5f0f\u3068\u3057\u3066\u89e3\u91c8\u3055\u308c\u308b\u6587\u5b57\u5217\u3092\u7121\u52b9\u5316\uff08\u5148\u982d\u306b\u30a2\u30dd\u30b9\u30c8\u30ed\u30d5\u30a3\u3092\u8ffd\u52a0\uff09";
    public const string PendingBackup = "\u7834\u68c4\u524d\u306e\u672a\u9001\u4fe1\u5909\u66f4\u3092\u30d0\u30c3\u30af\u30a2\u30c3\u30d7\u3057\u307e\u3057\u305f: ";

    public const string AppTitle = "Reader Data Viewer\uFF08\u4EEE\uFF09";

    // ---- the state word in the status bar -----------------------------------
    public const string StateBoot = "\u8D77\u52D5\u4E2D";
    public const string StateChecking = "\u66F4\u65B0\u3092\u78BA\u8A8D\u4E2D";
    public const string StateApplying = "\u53F0\u5E33\u3092\u66F4\u65B0\u4E2D";
    public const string StateDeleting = "\u30EC\u30B3\u30FC\u30C9\u3092\u524A\u9664\u4E2D";
    public const string StateSending = "\u9001\u4FE1\u4E2D";
    public const string StateReloading = "\u53F0\u5E33\u3092\u8AAD\u307F\u76F4\u3057\u4E2D";
    public const string StateLockWaiting = "\u53F0\u5E33\u304C\u7A7A\u304F\u306E\u3092\u5F85\u6A5F\u4E2D";
    public const string StateReady = "\u76E3\u8996\u4E2D";
    public const string StateSavingFmt = "{state}\u3092\u4FDD\u5B58\u4E2D";
    public const string StateBlocked = "\u53F0\u5E33\u304C\u3042\u308A\u307E\u305B\u3093";
    public const string StateWaitingFmt = "{name} \u3092\u5F85\u6A5F\u4E2D";
    public const string StateNoTarget = "\u76E3\u8996\u5BFE\u8C61\u306A\u3057";

    // ---- the watch segment ----------------------------------------------------
    public const string LabelWatch = "\u76E3\u8996";
    public const string WatchConnectedFmt = "\u63A5\u7D9A\u4E2D\uFF08{title}\uFF09";
    public const string WatchNone = "\u672A\u63A5\u7D9A";
    public const string WatchNoTarget = "\u76E3\u8996\u5BFE\u8C61\u304C\u3042\u308A\u307E\u305B\u3093 (\u8A2D\u5B9A\u3067\u8FFD\u52A0\u3057\u3066\u304F\u3060\u3055\u3044)";

    // ---- the ledger segment ---------------------------------------------------
    public const string LedgerSegFmt = "{file} / {n} \u4EF6";
    public const string PendingCountFmt = "\u672A\u9001\u4FE1 {n} \u4EF6";
    public const string NotYet = "--";
    public const string MsUnit = " ms";

    // ---- notices (shown in the status bar) ----------------------------------
    public const string NoteNoDiff = "\u66F4\u65B0\u306F\u3042\u308A\u307E\u305B\u3093 (\u53F0\u5E33\u306F\u6700\u65B0\u3067\u3059)";
    public const string NoteUpdated = "\u53F0\u5E33\u3092\u66F4\u65B0\u3057\u307E\u3057\u305F";
    public const string NoteRejected = "\u66F4\u65B0\u3092\u898B\u9001\u308A\u307E\u3057\u305F (\u4FDD\u5B58\u6E08\u307F\u53F0\u5E33\u306E\u307E\u307E)";
    public const string NoteSettingsApplied = "\u8A2D\u5B9A\u3092\u4FDD\u5B58\u3057\u307E\u3057\u305F";
    public const string NoteSaveDoneCanClose = "\u72B6\u614B\u306E\u4FDD\u5B58\u304C\u5B8C\u4E86\u3057\u307E\u3057\u305F\u3002\u7D42\u4E86\u3067\u304D\u307E\u3059";
    public const string NoteSaveFailedCanClose = "\u72B6\u614B\u306E\u4FDD\u5B58\u306F\u5931\u6557\u3068\u3057\u3066\u78BA\u5B9A\u3057\u307E\u3057\u305F\u3002\u7D42\u4E86\u3067\u304D\u307E\u3059";
    public const string NoteStateSaved = "{state}\u3068\u3057\u3066\u4FDD\u5B58\u3057\u307E\u3057\u305F";
    public const string NoteSendDone = "{n} \u4EF6\u3092\u9001\u4FE1\u3057\u307E\u3057\u305F";
    public const string NoteDeleteDone = "{n} \u4EF6\u3092\u524A\u9664\u3057\u307E\u3057\u305F";
    public const string NoteNoPending = "\u672A\u9001\u4FE1\u306E\u5909\u66F4\u306F\u3042\u308A\u307E\u305B\u3093";
    public const string NoteNotFound = "\u898B\u3064\u304B\u308A\u307E\u305B\u3093";

    // ---- confirmations ---------------------------------------------------------
    public const string ConfirmUpdateTitle = "\u66F4\u65B0\u306E\u78BA\u8A8D";
    public const string ConfirmRebuildBody = "\u4FDD\u5B58\u6E08\u307F\u306E\u7D71\u5408\u53F0\u5E33\u304C\u8AAD\u3081\u307E\u305B\u3093:\n{err}\nCSV \u304B\u3089\u4F5C\u308A\u76F4\u3057\u307E\u3059\u304B? (\u4F5C\u696D\u72B6\u614B\u306F\u5931\u308F\u308C\u307E\u3059)";
    public const string ConfirmCreateBody = "\u4FDD\u5B58\u6E08\u307F\u306E\u7D71\u5408\u53F0\u5E33\u304C\u3042\u308A\u307E\u305B\u3093\u3002CSV \u304B\u3089\u65B0\u3057\u304F\u4F5C\u6210\u3057\u307E\u3059\u304B?";
    public const string ConfirmStateTitleFmt = "{state}\u306E\u78BA\u8A8D";
    public const string SendTitle = "\u9001\u4FE1";
    public const string ConfirmSendBody = "{n} 件の確認状態を統合台帳に反映します。よろしいですか？";
    public const string BtnYes = "\u306F\u3044";
    public const string BtnNo = "\u3044\u3044\u3048";

    public static string UpdateConfirmBody(string onSourceChange, string initialState)
    {
        string first = "CSVから読み取った内容と、保存済みの統合台帳に違いがあります。統合台帳を更新しますか？\n"
            + "「はい」で反映し、「いいえ」で現在の台帳を保ちます。\n";
        if (onSourceChange == "preserve")
        {
            return first + "確認状態は現在のまま保ちます。";
        }
        return first + "CSVの内容が更新されたデータは、再確認が必要なため、確認状態を「" + initialState + "」に戻します。\n送信済みの確認済データと削除済みデータは変更しません。";
    }

    // ---- errors (shown in a warning dialog) ---------------------------------
    public const string ErrNo64 = "64 \u30D3\u30C3\u30C8\u306E\u30D7\u30ED\u30BB\u30B9\u304C\u5FC5\u8981\u3067\u3059\u3002";
    public const string ErrAlreadyRunning = "\u540C\u3058\u7D71\u5408\u53F0\u5E33\u3092\u958B\u3044\u3066\u3044\u308B Reader Data Viewer \u304C\u3001\u3059\u3067\u306B\u8D77\u52D5\u3057\u3066\u3044\u307E\u3059\u3002";
    public const string ErrNoData = "\u30C7\u30FC\u30BF\u30D5\u30A1\u30A4\u30EB\u304C\u898B\u3064\u304B\u308A\u307E\u305B\u3093: ";
    public const string ErrBadLedgerPath = "\u7D71\u5408\u53F0\u5E33\u306E\u30D1\u30B9\u304C\u4E0D\u6B63\u3067\u3059: ";
    public const string ErrNoLedger = "\u7D71\u5408\u53F0\u5E33\u304C\u3042\u308A\u307E\u305B\u3093\u3002\u691C\u7D22\u3067\u304D\u307E\u305B\u3093\u3002";
    public const string ErrCheckFailed = "\u66F4\u65B0\u78BA\u8A8D\u306B\u5931\u6557\u3057\u307E\u3057\u305F: ";
    public const string ErrCheckTimeout = "\u66F4\u65B0\u78BA\u8A8D\u304C\u30BF\u30A4\u30E0\u30A2\u30A6\u30C8\u3057\u307E\u3057\u305F (\u4FDD\u5B58\u6E08\u307F\u53F0\u5E33\u306E\u307E\u307E\u7D9A\u884C\u3057\u307E\u3059)";
    public const string ErrPersist = "\u53F0\u5E33\u3092\u66F8\u304D\u8FBC\u3081\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrLedgerRead = "\u53F0\u5E33\u3092\u8AAD\u307F\u8FBC\u3081\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrNotReady = "\u66F4\u65B0\u78BA\u8A8D\u304C\u7D42\u308F\u308B\u307E\u3067\u64CD\u4F5C\u3067\u304D\u307E\u305B\u3093";
    public const string ErrSaveInFlight = "\u72B6\u614B\u3092\u4FDD\u5B58\u4E2D\u3067\u3059\u3002\u78BA\u5B9A\u3059\u308B\u307E\u3067\u6B21\u306E\u64CD\u4F5C\u306F\u3067\u304D\u307E\u305B\u3093";
    public const string ErrCloseWhileWriting = "\u66F8\u304D\u8FBC\u307F\u4E2D\u3067\u3059\u3002\u7D50\u679C\u304C\u78BA\u5B9A\u3059\u308B\u307E\u3067\u7D42\u4E86\u3067\u304D\u307E\u305B\u3093";
    public const string ErrSaveOverdue = "\u72B6\u614B\u306E\u4FDD\u5B58\u304C\u60F3\u5B9A\u3088\u308A\u9577\u5F15\u3044\u3066\u3044\u307E\u3059 (\u78BA\u5B9A\u3059\u308B\u307E\u3067\u7D42\u4E86\u3057\u307E\u305B\u3093)";
    public const string ErrBadKeyFmt = "{label} \u304C\u5F62\u5F0F {pattern} \u306B\u4E00\u81F4\u3057\u307E\u305B\u3093";
    public const string ErrBadPattern = "\u756A\u53F7\u306E\u5F62\u5F0F\uFF08\u6B63\u898F\u8868\u73FE\uFF09\u304C\u4E0D\u6B63\u3067\u3059\u3002\u8A2D\u5B9A\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044";
    public const string ErrNoRecordShown = "{state}\u306B\u3059\u308B\u30EC\u30B3\u30FC\u30C9\u304C\u8868\u793A\u3055\u308C\u3066\u3044\u307E\u305B\u3093";
    public const string ErrNoTransition = "\u3053\u306E\u30EC\u30B3\u30FC\u30C9\u306F\u3059\u3067\u306B{state}\u3067\u3059";
    public const string ErrUnknownState = "\u53F0\u5E33\u306E\u72B6\u614B {stored} \u306F\u753B\u9762\u5B9A\u7FA9\u306B\u3042\u308A\u307E\u305B\u3093\u3002\u9077\u79FB\u3067\u304D\u307E\u305B\u3093";
    public const string ErrSettingsSave = "\u8A2D\u5B9A\u3092\u4FDD\u5B58\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrPatternTyped = "\u5F62\u5F0F\uFF08\u6B63\u898F\u8868\u73FE\uFF09\u304C\u4E0D\u6B63\u3067\u3059: ";
    public const string ErrPathBlank = "\u30D1\u30B9\u306F 3 \u3064\u3068\u3082\u5FC5\u8981\u3067\u3059";
    public const string ErrDataDir = "\u30C7\u30FC\u30BF\u30D5\u30A9\u30EB\u30C0\u30FC\u304C\u898B\u3064\u304B\u308A\u307E\u305B\u3093: ";
    public const string ErrLogWrite = "\u30ED\u30B0\u3092\u66F8\u3051\u307E\u305B\u3093: ";
    public const string ErrExport = "CSV \u3092\u51FA\u529B\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrDelete = "\u30EC\u30B3\u30FC\u30C9\u3092\u524A\u9664\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrPendingRead = "\u624B\u5143\u306E\u672A\u9001\u4FE1\u30C7\u30FC\u30BF\u3092\u8AAD\u3081\u307E\u305B\u3093: ";
    public const string ErrPendingWrite = "\u624B\u5143\u306E\u672A\u9001\u4FE1\u30C7\u30FC\u30BF\u3092\u66F8\u3051\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrSend = "\u9001\u4FE1\u3067\u304D\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrSendOverdue = "\u9001\u4FE1\u304C\u9577\u5F15\u3044\u3066\u3044\u307E\u3059\u3002\u53F0\u5E33\u304C\u7A7A\u304F\u307E\u3067\u5F85\u3063\u3066\u3044\u307E\u3059";
    public const string ErrSharedWriteOverdue = "\u53F0\u5E33\u306E\u66F8\u304D\u8FBC\u307F\u304C\u9577\u5F15\u3044\u3066\u3044\u307E\u3059\u3002\u7D50\u679C\u304C\u78BA\u5B9A\u3059\u308B\u307E\u3067\u5F85\u3063\u3066\u3044\u307E\u3059";
    public const string ErrReloadOverdue = "\u53F0\u5E33\u306E\u8AAD\u307F\u76F4\u3057\u304C\u9577\u5F15\u3044\u3066\u3044\u307E\u3059";
    public const string ErrSharedMarker = "\u53F0\u5E33\u306E\u66F4\u65B0\u901A\u77E5\u3092\u8AAD\u3081\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string ErrSharedReload = "\u66F4\u65B0\u3055\u308C\u305F\u53F0\u5E33\u3092\u8AAD\u307F\u76F4\u305B\u307E\u305B\u3093\u3067\u3057\u305F: ";
    public const string LockWaitingFmt = "{user}\uFF08{host}\uFF09\u304C {minutes} \u5206\u524D\u304B\u3089\u53F0\u5E33\u3092\u4F7F\u7528\u3057\u3066\u3044\u307E\u3059\u3002\u7A7A\u304F\u307E\u3067\u5F85\u3061\u307E\u3059\u3002";

    // ---- the app does not start / cannot go on (one modal, then it stops) ----
    public const string FatalTitle = "\u8D77\u52D5\u3067\u304D\u307E\u305B\u3093";
    public const string FatalSettings = "\u8A2D\u5B9A\u30D5\u30A1\u30A4\u30EB\u306B\u554F\u984C\u304C\u3042\u308B\u305F\u3081\u8D77\u52D5\u3067\u304D\u307E\u305B\u3093\u3002\n\n\u30D5\u30A1\u30A4\u30EB: {file}\n{reason}\n\n\u76F4\u3057\u3066\u304B\u3089\u8D77\u52D5\u3057\u76F4\u3057\u3066\u304F\u3060\u3055\u3044\u3002";
    public const string FatalDataTitle = "\u30C7\u30FC\u30BF\u3092\u8AAD\u3081\u307E\u305B\u3093";
    public const string FatalData = "\u30C7\u30FC\u30BF\u306B\u554F\u984C\u304C\u3042\u308B\u305F\u3081\u7D9A\u884C\u3067\u304D\u307E\u305B\u3093\u3002\n\n{reason}\n\n\u76F4\u3057\u3066\u304B\u3089\u8D77\u52D5\u3057\u76F4\u3057\u3066\u304F\u3060\u3055\u3044\u3002";

    // ---- why a CSV or the ledger is refused ({file} / {row} / {name} ...) ----
    public const string DataNoRows = "{file}: \u30D8\u30C3\u30C0\u30FC\u884C\u3068\u30C7\u30FC\u30BF\u884C\u304C\u3042\u308A\u307E\u305B\u3093";
    public const string DataBlankHeader = "{file}: \u30D8\u30C3\u30C0\u30FC\u884C\u306B\u7A7A\u306E\u5217\u540D\u304C\u3042\u308A\u307E\u3059";
    public const string DataDupHeader = "{file}: \u5217\u540D {name} \u304C\u30D8\u30C3\u30C0\u30FC\u884C\u306B 2 \u56DE\u3042\u308A\u307E\u3059";
    public const string DataNoColumn = "{file} \u306B\u5217 {name} \u304C\u3042\u308A\u307E\u305B\u3093";
    public const string DataQuoted = "{file} \u306E {row} \u884C\u76EE: \u5F15\u7528\u7B26 (\") \u3067\u59CB\u307E\u308B\u5217\u304C\u3042\u308A\u307E\u3059\u3002\u5F15\u7528\u7B26\u4ED8\u304D\u306E CSV \u306F\u8AAD\u3081\u307E\u305B\u3093";
    public const string DataColumnCount = "{file} \u306E {row} \u884C\u76EE: \u5217\u6570\u304C {n} \u3067\u3059 (\u30D8\u30C3\u30C0\u30FC\u884C\u306F {cols} \u5217)";
    public const string DataEmptyKey = "{file} \u306E {row} \u884C\u76EE: \u30AD\u30FC\u5217 {name} \u304C\u7A7A\u3067\u3059";
    public const string DataKeyNotAscii = "{file} \u306E {row} \u884C\u76EE: \u30AD\u30FC\u5217 {name} \u306B ASCII \u4EE5\u5916\u306E\u6587\u5B57\u304C\u3042\u308A\u307E\u3059";
    public const string DataKeyWidth = "{file} \u306E {row} \u884C\u76EE: \u30AD\u30FC\u5217 {name} \u306E\u5E45\u304C\u6700\u521D\u306E\u6709\u52B9\u884C ({n} \u6587\u5B57) \u3068\u9055\u3044\u307E\u3059";
    public const string DataDupKey = "{file}: \u30AD\u30FC\u5217 {name} \u306E\u5024 {key} \u304C {row1} \u884C\u76EE\u3068 {row2} \u884C\u76EE\u306B\u3042\u308A\u307E\u3059 (\u30AD\u30FC\u306F\u4E00\u610F\u3067\u3042\u308B\u5FC5\u8981\u304C\u3042\u308A\u307E\u3059)";
    public const string DataLedgerHeader = "{file} \u306E\u898B\u51FA\u3057\u884C\u304C\u3001\u4F5C\u696D\u72B6\u614B\u306E\u5217\u3068\u753B\u9762\u5B9A\u7FA9\u306E\u53F0\u5E33\u5217\u306B\u4E00\u81F4\u3057\u307E\u305B\u3093";
    public const string DataControlChar = "{file} \u306E {row} \u884C\u76EE: \u5236\u5FA1\u6587\u5B57 (0x{code}) \u3092\u4EE3\u66FF\u6587\u5B57 (?) \u306B\u7F6E\u304D\u63DB\u3048\u3066\u8AAD\u307F\u9032\u3081\u307E\u3057\u305F";
    public const string DataLedgerTab = "{file} \u306E {row} \u884C\u76EE: \u30BB\u30EB\u306E\u30BF\u30D6\u6587\u5B57\u3092\u4EE3\u66FF\u6587\u5B57 (?) \u306B\u7F6E\u304D\u63DB\u3048\u3066\u8AAD\u307F\u9032\u3081\u307E\u3057\u305F";
    public const string DataLedgerBlankIdentity = "{file} \u306E {row} \u884C\u76EE: {name} \u304C\u7A7A\u3067\u3059";
    public const string DataLedgerDupIdentity = "{file}: {name} \u306E\u5024 {key} \u304C {row1} \u884C\u76EE\u3068 {row2} \u884C\u76EE\u306B\u3042\u308A\u307E\u3059 (\u53F0\u5E33\u306E 1 \u884C\u306F {name} \u3067\u7279\u5B9A\u3057\u307E\u3059)";
    public const string DataTypedValue = "{file} \u306E {row} \u884C\u76EE: \u5217 {name} \u306E\u5B9F\u969B\u306E\u5024\u300C{value}\u300D\u3092 {type} \u3068\u3057\u3066\u8AAD\u3081\u307E\u305B\u3093";
    public const string TypeDateFormat = "\u65E5\u4ED8\uFF08{format}\uFF09";
    public const string TypeNumber = "\u6570\u5024";
    public const string DataTypedResult = "\u51E6\u7406\u7D50\u679C\u306E\u5217 {name}\uFF08\u8B58\u5225 {identity}\uFF09\u306E\u5024\u300C{value}\u300D\u3092 {type} \u3068\u3057\u3066\u8AAD\u3081\u307E\u305B\u3093\u3002\u5F0F\u3084\u96C6\u8A08\u306E\u7D50\u679C\u3092\u78BA\u8A8D\u3057\u3066\u304F\u3060\u3055\u3044\u3002";
    public const string ProcessBlankIdentity = "\u30B8\u30E7\u30D6\u300C{job}\u300D\u306E\u51FA\u529B\u3067\u3001\u8B58\u5225\u5217\u300C{column}\u300D\u304C\u7A7A\u3067\u3059";
    public const string ProcessDuplicateIdentity = "\u30B8\u30E7\u30D6\u300C{job}\u300D\u306E\u51FA\u529B\u3067\u3001\u8B58\u5225\u5217\u300C{column}\u300D\u306E\u5024\u300C{value}\u300D\u304C\u91CD\u8907\u3057\u3066\u3044\u307E\u3059";

    // ---- placeholders and fixed words on the screen ---------------------------
    public const string PanelCand = "\u5019\u88DC\u4E00\u89A7";
    public const string FieldUnresolved = "\u5217\u306A\u3057";
    public const string StateBlank = "(\u7A7A)";
    public const string JudgeUndefined = "\u672A\u5B9A\u7FA9";
    public const string JudgeError = "\u30A8\u30E9\u30FC";
    public const string SavingSuffix = " (\u4FDD\u5B58\u4E2D...)";
    public const string BtnClose = "\u9589\u3058\u308B";
    public const string BtnOk = "OK";
    public const string BtnBrowse = "\u53C2\u7167...";
    public const string Unsearched = "\u672A\u691C\u7D22";
    public const string CandidateHitsFmt = "\u8A72\u5F53 {n} \u4EF6";

    // ---- shared-ledger notices ------------------------------------------------
    public const string SharedSendBodyMulti = "{user} が {changed} 件を{initialState}以外の状態に、{initial} 件を{initialState}にしました";
    public const string SharedSendBody = "{user} \u304C {changed} \u4EF6\u3092{changedState}\u3001{initial} \u4EF6\u3092{initialState}\u306B\u3057\u307E\u3057\u305F";
    public const string SharedUpdateTitle = "\u53F0\u5E33\u306E\u66F4\u65B0";
    public const string SharedUpdateBody = "\u53F0\u5E33\u304C\u66F4\u65B0\u3055\u308C\u307E\u3057\u305F\u3002\u5207\u308A\u66FF\u3048\u307E\u3059\u304B";
    public const string SharedResetFmt = "\u4E2D\u8EAB\u304C\u5909\u308F\u3063\u305F\u305F\u3081{state}\u306B\u623B\u3063\u305F\u30EC\u30B3\u30FC\u30C9: {n} \u4EF6";
    public const string UnmatchedTitle = "\u9001\u4FE1\u3067\u304D\u306A\u304B\u3063\u305F\u30EC\u30B3\u30FC\u30C9";
    public const string UnmatchedBodyFmt = "\u5171\u6709\u53F0\u5E33\u3078\u5F53\u3066\u3089\u308C\u306A\u304B\u3063\u305F\u5909\u66F4\u304C {n} \u4EF6\u3042\u308A\u307E\u3059\u3002\u672A\u9001\u4FE1\u306E\u307E\u307E\u6B8B\u3057\u307E\u3057\u305F\u3002";
    public const string UnmatchedMissing = "\u884C\u304C\u3042\u308A\u307E\u305B\u3093";
    public const string UnmatchedChanged = "\u4E2D\u8EAB\u304C\u5909\u308F\u3063\u3066\u3044\u307E\u3059";
    public const string ColReason = "\u7406\u7531";

    // ---- the per-terminal operation log beside the shared ledger -------------
    public const string OpLogInfix = "-\u64cd\u4f5c\u30ed\u30b0-";
    public const string OpLogHeader = "\u64cd\u4f5c\u65e5\u6642,\u7aef\u672b\u540d,\u30e6\u30fc\u30b6\u30fc\u540d,\u64cd\u4f5c,\u53f0\u5e33\u306e\u884c\u6570,\u5185\u5bb9";
    public const string OpCreate = "\u4f5c\u6210";
    public const string OpUpdate = "\u66f4\u65b0";
    public const string OpDelete = "\u524a\u9664";
    public const string OpSend = "\u9001\u4fe1";
    public const string OpUpdateDetailFmt = "{job}: \u8ffd\u52a0 {added} \u4ef6\u3001\u66f4\u65b0 {updated} \u4ef6\u3001\u524a\u9664 {deleted} \u4ef6\u3001{state}\u306b\u623b\u3057\u305f {reset} \u4ef6";
    public const string OpDeleteDetailFmt = "{job}: \u524a\u9664 {n} \u4ef6";
    public const string OpSendDetailItemFmt = "{state} {n} \u4ef6";
    public const string OpSendDetailSeparator = "\u3001";

    // ---- process job dialogs --------------------------------------------------
    public const string UpdateRecordsTitle = "\u30EC\u30B3\u30FC\u30C9\u66F4\u65B0";
    public const string DeleteRecordsTitle = "\u30EC\u30B3\u30FC\u30C9\u524A\u9664";
    public const string UpdateRecordsHint = "\u5B9A\u7FA9\u3055\u308C\u305F\u30D5\u30A1\u30A4\u30EB\u3060\u3051\u3092\u8AAD\u307F\u3001JSON \u306E\u624B\u9806\u3067\u7D71\u5408\u53F0\u5E33\u3092\u66F4\u65B0\u3057\u307E\u3059\u3002";
    public const string DeleteRecordsHint = "\u5B9A\u7FA9\u3055\u308C\u305F\u30D5\u30A1\u30A4\u30EB\u3060\u3051\u3092\u8AAD\u307F\u3001\u6761\u4EF6\u306B\u4E00\u81F4\u3059\u308B\u884C\u3092\u7D71\u5408\u53F0\u5E33\u304B\u3089\u53D6\u308A\u9664\u304D\u307E\u3059\u3002";
    public const string SecInputs = "\u53D6\u308A\u8FBC\u3080\u30C7\u30FC\u30BF\uFF08{dir}\uFF3C\uFF09";
    public const string SecProcess = "\u51E6\u7406\u5185\u5BB9";
    public const string SecOutput = "\u66F8\u304D\u51FA\u3057\u5148";
    public const string ColInput = "\u8868";
    public const string ColDeleteInput = "\u6307\u5B9A";
    public const string ColFile = "\u30D5\u30A1\u30A4\u30EB";
    public const string ColKey = "\u30AD\u30FC";
    public const string ColRows = "\u884C\u6570";
    public const string ColValidation = "\u691C\u8A3C";
    public const string ColNumber = "#";
    public const string ColOperation = "\u64CD\u4F5C";
    public const string ColTarget1 = "\u5BFE\u8C611";
    public const string ColTarget2 = "\u5BFE\u8C612";
    public const string ColCondition = "\u6761\u4EF6";
    public const string ColOutput = "\u51FA\u529B";
    public const string ValidationColumnsMatch = "\u5217\u4E00\u81F4";
    public const string ValidationMissing = "\u30D5\u30A1\u30A4\u30EB\u306A\u3057";
    public const string ValidationError = "\u4E0D\u4E00\u81F4";
    public const string ValidationEncodingMismatch = "\u6587\u5B57\u30B3\u30FC\u30C9\u4E0D\u4E00\u81F4\uFF08{row} \u884C\u76EE\uFF09";
    public const string LblPath = "\u30D1\u30B9";
    public const string LblFileName = "\u30D5\u30A1\u30A4\u30EB\u540D";
    public const string LblLastWrite = "\u6700\u7D42\u66F4\u65B0";
    public const string LblNeverWritten = "\u672A\u4F5C\u6210";
    public const string BtnExecute = "\u5B9F\u884C";
    public const string BtnDelete = "\u524A\u9664\u3059\u308B";
    public const string ProcessNotRun = "\u5165\u529B\u30D5\u30A1\u30A4\u30EB\u3092\u78BA\u8A8D\u3067\u304D\u306A\u3044\u305F\u3081\u5B9F\u884C\u3067\u304D\u307E\u305B\u3093\u3002";

    public static string OperationLabel(string operation)
    {
        switch (operation)
        {
            case "join": return "\u7D50\u5408";
            case "extract": return "\u62BD\u51FA";
            case "delete": return "\u524A\u9664";
            case "append": return "\u8FFD\u52A0";
            case "update": return "\u66F4\u65B0";
            case "merge": return "\u30DE\u30FC\u30B8";
            case "replace": return "\u7F6E\u63DB";
            case "select": return "\u9078\u629E";
            case "calculate": return "\u8A08\u7B97";
            case "aggregate": return "\u96C6\u8A08";
            case "sort": return "\u4E26\u3079\u66FF\u3048";
            case "distinct": return "\u91CD\u8907\u9664\u53BB";
        }
        return operation;
    }

    public static string MergeDestinations(string sourceOnly, string both, string targetOnly)
    {
        return "\u5143\u306E\u307F:" + RowDestination(sourceOnly)
            + " / \u4E21\u65B9:" + RowDestination(both)
            + " / \u5148\u306E\u307F:" + RowDestination(targetOnly);
    }

    private static string RowDestination(string value)
    {
        switch (value)
        {
            case "add": return "\u8FFD\u52A0";
            case "ignore": return "\u7121\u8996";
            case "update": return "\u66F4\u65B0";
            case "keep": return "\u4FDD\u6301";
            case "delete": return "\u524A\u9664";
        }
        return value;
    }

    public static string ConditionLabel(string condition)
    {
        switch (condition)
        {
            case "match": return "\u4E00\u81F4";
            case "either": return "\u3069\u3061\u3089\u304B";
            case "both": return "\u4E21\u65B9";
            case "exclude": return "\u9664\u304F";
        }
        return (condition == null) ? "" : condition;
    }

    public static string JoinConditionLabel(string condition)
    {
        switch (condition)
        {
            case "match": return "\u5185\u90E8";
            case "left": return "\u5DE6\u5916\u90E8";
            case "full": return "\u5B8C\u5168\u5916\u90E8";
        }
        return ConditionLabel(condition);
    }

    public static string PredicateLabel(string operation)
    {
        switch (operation)
        {
            case "equals": return "\u7B49\u3057\u3044";
            case "notEquals": return "\u7B49\u3057\u304F\u306A\u3044";
            case "contains": return "\u542B\u3080";
            case "startsWith": return "\u3067\u59CB\u307E\u308B";
            case "endsWith": return "\u3067\u7D42\u308F\u308B";
            case "empty": return "\u7A7A";
            case "notEmpty": return "\u7A7A\u3067\u306A\u3044";
            case "greater": return "\u3088\u308A\u5927\u304D\u3044";
            case "atLeast": return "\u4EE5\u4E0A";
            case "less": return "\u3088\u308A\u5C0F\u3055\u3044";
            case "atMost": return "\u4EE5\u4E0B";
        }
        return operation;
    }

    public static string AggregateLabel(string function)
    {
        if (function == "sum") { return "\u5408\u8A08"; }
        if (function == "count") { return "\u4EF6\u6570"; }
        return function;
    }

    public static string DirectionLabel(string direction)
    {
        if (direction == "ascending") { return "\u6607\u9806"; }
        if (direction == "descending") { return "\u964D\u9806"; }
        return direction;
    }

    public static string SortTypeLabel(string type)
    {
        if (type == "text") { return "\u6587\u5B57"; }
        if (type == "number") { return "\u6570\u5024"; }
        return type;
    }

    // ---- table export ---------------------------------------------------------
    public const string ExportTitle = "\u30C6\u30FC\u30D6\u30EB\u51FA\u529B";
    public const string ExportHint = "\u7D71\u5408\u53F0\u5E33\u304B\u3089\u3001\u9078\u3093\u3060\u9805\u76EE\u3060\u3051\u3092 CSV \u306B\u66F8\u304D\u51FA\u3057\u307E\u3059\u3002";
    public const string ExportAvailable = "\u51FA\u529B\u3067\u304D\u308B\u9805\u76EE";
    public const string ExportSelectedFmt = "\u51FA\u529B\u3059\u308B\u9805\u76EE\uFF08{n}\uFF09";
    public const string ExportDefault = "\u65E2\u5B9A\u306B\u623B\u3059";
    public const string ExportDestination = "\u51FA\u529B\u5148";
    public const string ExportDefaultPath = "output\\export-{yyyyMMdd-HHmmss}.csv";
    public const string ExportNeedField = "\u51FA\u529B\u3059\u308B\u9805\u76EE\u3092 1 \u3064\u4EE5\u4E0A\u9078\u3093\u3067\u304F\u3060\u3055\u3044\u3002";
    public const string ExportDoneFmt = "CSV \u3092\u51FA\u529B\u3057\u307E\u3057\u305F: {file}";
    public const string ExportFilterGroup = "\u7D5E\u308A\u8FBC\u307F\u6761\u4EF6\uFF08\u3059\u3079\u3066\u306B\u4E00\u81F4\uFF09";
    public const string ExportFilterField = "\u9805\u76EE";
    public const string ExportFilterCondition = "\u6761\u4EF6";
    public const string ExportFilterValue = "\u5024";
    public const string ExportFilterAdd = "\u8FFD\u52A0";
    public const string ExportFilterRemove = "\u524A\u9664";
    public const string ExportFilterContains = "\u3092\u542B\u3080";
    public const string ExportFilterEquals = "\u3068\u7B49\u3057\u3044";
    public const string ExportFilterStarts = "\u3067\u59CB\u307E\u308B";
    public const string ExportFilterNotContains = "\u3092\u542B\u307E\u306A\u3044";
    public const string ExportFilterRange = "\u306E\u7BC4\u56F2";
    public const string ExportFilterRangeMark = "\uFF5E";
    public const string ExportFilterNeedValue = "\u7D5E\u308A\u8FBC\u3080\u5024\u3092\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044\u3002";
    public const string ExportFilterNeedNumber = "\u7BC4\u56F2\u306E\u4E21\u7AEF\u306B\u6570\u5024\u3092\u5165\u529B\u3057\u3066\u304F\u3060\u3055\u3044\u3002";
    public const string ExportFilterOrder = "\u7BC4\u56F2\u306E\u5148\u982D\u306F\u672B\u5C3E\u4EE5\u4E0B\u306B\u3057\u3066\u304F\u3060\u3055\u3044\u3002";
    public const string BtnMoveRight = "\u25B6";
    public const string BtnMoveLeft = "\u25C0";

    // ---- the settings modal -----------------------------------------------------
    public const string SettingsTitle = "\u8A2D\u5B9A";
    public const string SettingsHint = "\u66F8\u304D\u623B\u3059\u306E\u306F paths / search / watch \u306E 3 \u3064\u3060\u3051\u3067\u3059\u3002";
    public const string SecPlaces = "\u5834\u6240";
    public const string SecSearch = "\u691C\u7D22";
    public const string LblDataShort = "\u30C7\u30FC\u30BF";
    public const string LblKeyPatternShort = "\u756A\u53F7\u306E\u5F62\u5F0F";
    public const string LblCandidateRows = "\u5019\u88DC\u306E\u8868\u793A\u4EF6\u6570";
    public const string LblTarget = "\u5BFE\u8C61";
    public const string LblRead = "\u8AAD\u307F\u53D6\u308A";
    public const string ReadSummaryFmt = "{mode} / {poll}ms \u9593\u9694";
    public const string ReadValuePattern = "ValuePattern";
    public const string ReadTextPattern = "TextPattern";
    public const string ReadNameProperty = "Name property";
    public const string SecTargets = "\u76E3\u8996\u5BFE\u8C61";
    public const string SecTarget = "\u5BFE\u8C61\u306E\u6307\u5B9A";
    public const string BtnInspect = "\u753B\u9762\u304B\u3089\u9078\u3076";
    public const string NoteTargetSummary = "{name}\uFF08{kind} = {value}\uFF09";
    public const string NoteNoTargetShort = "\u76E3\u8996\u5BFE\u8C61\u304C\u3042\u308A\u307E\u305B\u3093";
    public const string LblLedger = "\u7D71\u5408\u53F0\u5E33";
    public const string LblLog = "\u30ED\u30B0";
    public const string BtnCancel = "\u30AD\u30E3\u30F3\u30BB\u30EB";
    public const string NoValue = "N/A";

    // ---- the element picker ------------------------------------------------------
    public const string PickTitle = "\u753B\u9762\u304B\u3089\u9078\u3076";
    public const string PickHow = "\u5BFE\u8C61\u306E\u6B04\u306B\u30AB\u30FC\u30BD\u30EB\u3092\u5408\u308F\u305B\u3066  Ctrl + Shift  \u3092\u62BC\u3059";
    public const string PickEsc = "Esc \u3067\u4E2D\u6B62";
    public const string LblControlTypes = "\u7A2E\u985E";
    public const string LblAutomationId = "AutomationId";
    public const string LblClassName = "\u30AF\u30E9\u30B9\u540D";
    public const string LblName = "\u540D\u524D";
    public const string LblProcessOf = "\u30D7\u30ED\u30BB\u30B9";
    public const string PickReading = "\u8AAD\u307F\u53D6\u308A";
    public const string PickNoRead = "\u3053\u306E\u8981\u7D20\u304B\u3089\u306F\u5024\u3092\u8AAD\u3081\u307E\u305B\u3093";

    public const string ProcessMissingColumn = "表に列「{0}」がありません。設定の参照列と入力の見出しを確認してください。";
    public const string ProcessUnknownOperation = "未対応の操作「{0}」です。";
    public const string LedgerRowColumns = "台帳の {0} 行目の列数が違います。必要 {1} 列、実際 {2} 列。";
    public const string ProcessDuplicateColumn = "出力の列「{0}」が重複しています。設定の列名を区別してください。";
    public const string ProcessAppendCount = "縦結合する表の列数が違います。対象1は {0} 列、対象2は {1} 列です。";
    public const string ProcessAppendColumn = "縦結合の {0} 列目の見出しが違います。対象1「{1}」、対象2「{2}」。";
    public const string ProcessSelectionSource = "対象の表と行の選択元が違います。抽出した行と同じ表を対象に指定してください。";
    public const string ProcessLedgerKey = "書き込み先のキーが data.ledger.identity と違います。設定の keys の対象2を台帳の識別列に合わせてください。";
    public const string LedgerLengths = "台帳の内容行数と作業状態の行数が一致しません。台帳の整合性を確認してください。";
    public const string LedgerNotUpdate = "レコード更新のジョブが指定されていません。";
    public const string LedgerNoResult = "更新ジョブの結果が台帳ではありません。最後の出力を確認してください。";
    public const string LedgerNoWrite = "更新ジョブに台帳へ書き込む手順がありません。";
    public const string LedgerNoChangeRule = "作業状態の onSourceChange に reset または preserve が指定されていません。";
    public const string LedgerNotDelete = "レコード削除のジョブが指定されていません。";
    public const string LedgerNoDeleteResult = "削除ジョブの結果が台帳ではありません。最後の出力を確認してください。";
    public const string LedgerBlankRow = "{0} の {1} 行目の識別列が空です。台帳の処理を停止しました。";
    public const string LedgerDuplicateRows = "{0} の識別値「{1}」が {2} 行目と {3} 行目で重複しています。台帳の処理を停止しました。";
    public const string XlsxNoSheetPart = "{0}: ワークシートのデータがありません。ブックの構造を確認してください。";
    public const string XlsxDuplicateCell = "{0}: セル {1} の位置が重複しているか、XFD 列の範囲を超えています。";
    public const string XlsxLedgerExtra = "{0}: 台帳のセル {1} に未定義の列の値「{2}」があります。";
    public const string XlsxLedgerNoHead = "{0}: 台帳に見出し行がありません。";
    public const string XlsxStringIndex = "共有文字列を指す番号「{0}」が不正です。ブックの文字列データが壊れています。";
    public const string XlsxNoMetadata = "ブックの構成情報（workbook.xml または関連付け）がありません。";
    public const string XlsxLedgerSheets = "共有台帳には LEDGER シートだけを置いてください。追加のシートは別のブックへ保存してください。";
    public const string XlsxNoLedgerSheet = "LEDGER シートがありません。";
    public const string XlsxSheetNotFound = "\u30d6\u30c3\u30af\u306b\u30b7\u30fc\u30c8\u300c{name}\u300d\u304c\u3042\u308a\u307e\u305b\u3093\u3002\u3042\u308b\u306e\u306f: {sheets}";
    public const string XlsxNoSheet = "ブックにワークシートがありません。";
    public const string XlsxExternalSheet = "シート「{0}」の参照先がブック内のワークシートではありません。参照先「{1}」。";
    public const string XlsxBadTarget = "シートの参照先「{0}」が不正です。";
    public const string XlsxMissingTarget = "シート「{0}」の参照先「{1}」がブック内にありません。";
    public const string XlsxMissingRelation = "シート「{0}」の関連付け「{1}」がありません。";
    public const string XlsxInvalidWriteCell = "セルの値が存在しないか、上限の 32767 文字を超えています。";
    public const string XlsxInvalidDimensions = "台帳の行数・列数・作業状態の件数が不正です。行は 1048575 件、内容列は 16383 列までで、状態は各行に1つ必要です。";
    public const string XlsxNullRow = "台帳の {0} 行目の内容が存在しません。";
    public const string XlsxColumnRange = "セル番地「{0}」が XFD 列の範囲を超えています。";
    public const string ExpressionExpected = "数値、引用符で囲んだ文字列、列名、関数、または左括弧が必要です。";
    public const string ExpressionParen = "右括弧が必要です。";
    public const string ExpressionEmptyPattern = "regexExtract の正規表現は空にできません。";
    public const string ExpressionEmptySeparator = "splitPart の区切り文字は空にできません。";
    public const string ExpressionQuote = "文字列の閉じる引用符がありません。";
    public const string ExpressionUnexpected = "式の末尾に解釈できない文字があります。";
    public const string ExpressionColumn = "列「{0}」がありません。";
    public const string ExpressionFunction = "関数「{0}」は未対応です。";
    public const string ExpressionArgumentSeparator = "関数 {0} の引数の後にカンマまたは右括弧が必要です。";
    public const string ExpressionArguments = "関数 {0} には {1} 個の引数が必要です。実際は {2} 個です。";
    public const string ExpressionBadPattern = "regexExtract の正規表現「{0}」が不正です。{1}";
    public const string ExpressionQuotedArgument = "関数 {0} の引数 {1} は引用符で囲んだ文字列で指定してください。";
    public const string ExpressionWholeArgument = "関数 {0} の引数 {1} は {2} 以上の整数で指定してください。";
    public const string ExpressionLocation = "式「{0}」の {1} 文字目: {2}";
}
