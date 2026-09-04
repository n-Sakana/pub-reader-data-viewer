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
    public const string ConfirmSendBody = "\u672A\u9001\u4FE1\u306E {n} \u4EF6\u3092\u9001\u4FE1\u3057\u307E\u3059\u3002\u3088\u308D\u3057\u3044\u3067\u3059\u304B?\n\n\u9001\u4FE1\u3057\u305F\u884C\u306F\u672A\u9001\u4FE1\u304B\u3089\u5916\u308C\u307E\u3059\u3002\u53D6\u308A\u8FBC\u3093\u3060\u30C7\u30FC\u30BF\u306F\u66F8\u304D\u63DB\u3048\u307E\u305B\u3093\u3002";
    public const string BtnYes = "\u306F\u3044";
    public const string BtnNo = "\u3044\u3044\u3048";

    public static string UpdateConfirmBody(string onSourceChange, string initialState)
    {
        string first = "\u5B9A\u7FA9\u3055\u308C\u305F\u51E6\u7406\u3067\u53F0\u5E33\u306B\u5909\u66F4\u304C\u3042\u308A\u307E\u3059\u3002\u66F4\u65B0\u3057\u307E\u3059\u304B?\n";
        if (onSourceChange == "preserve")
        {
            return first + "(\u5165\u529B\u5074\u306E\u5217\u304C\u5909\u308F\u3063\u3066\u3082\u3001\u4F5C\u696D\u72B6\u614B\u306F\u73FE\u5728\u5024\u3092\u4FDD\u3061\u307E\u3059)";
        }
        return first + "(\u5165\u529B\u5074\u306E\u5217\u304C\u5909\u308F\u3063\u305F\u884C\u306F\u3001\u4F5C\u696D\u72B6\u614B\u3092\u300C" + initialState + "\u300D\u3078\u623B\u3057\u307E\u3059)";
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
    public const string SharedSendBody = "{user} \u304C {changed} \u4EF6\u3092{changedState}\u3001{initial} \u4EF6\u3092{initialState}\u306B\u3057\u307E\u3057\u305F";
    public const string SharedUpdateTitle = "\u53F0\u5E33\u306E\u66F4\u65B0";
    public const string SharedUpdateBody = "\u53F0\u5E33\u304C\u66F4\u65B0\u3055\u308C\u307E\u3057\u305F\u3002\u5207\u308A\u66FF\u3048\u307E\u3059\u304B";
    public const string SharedResetFmt = "\u4E2D\u8EAB\u304C\u5909\u308F\u3063\u305F\u305F\u3081{state}\u306B\u623B\u3063\u305F\u30EC\u30B3\u30FC\u30C9: {n} \u4EF6";
    public const string UnmatchedTitle = "\u9001\u4FE1\u3067\u304D\u306A\u304B\u3063\u305F\u30EC\u30B3\u30FC\u30C9";
    public const string UnmatchedBodyFmt = "\u5171\u6709\u53F0\u5E33\u3078\u5F53\u3066\u3089\u308C\u306A\u304B\u3063\u305F\u5909\u66F4\u304C {n} \u4EF6\u3042\u308A\u307E\u3059\u3002\u672A\u9001\u4FE1\u306E\u307E\u307E\u6B8B\u3057\u307E\u3057\u305F\u3002";
    public const string UnmatchedMissing = "\u884C\u304C\u3042\u308A\u307E\u305B\u3093";
    public const string UnmatchedChanged = "\u4E2D\u8EAB\u304C\u5909\u308F\u3063\u3066\u3044\u307E\u3059";
    public const string ColReason = "\u7406\u7531";

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
}
