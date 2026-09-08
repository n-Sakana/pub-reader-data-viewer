// Windows/.NET Framework regression tests. Never write the supplied input data
// or the configured ledger. All ledger/CSV/pending fixtures use a fresh temp dir.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;

public static class Rdv3RegressionTests
{
    private static string root, temp;
    private static int passed, failed;
    private static readonly List<string> results = new List<string>();

    private static void Check(bool condition, string message)
    { if (!condition) { throw new Exception(message); } }
    private static void Throws<T>(Action action) where T : Exception
    {
        try { action(); }
        catch (T) { return; }
        throw new Exception("Expected " + typeof(T).Name);
    }
    private static void Test(string name, Action action)
    {
        try { action(); passed++; results.Add("PASS " + name); }
        catch (Exception ex) { failed++; results.Add("FAIL " + name + ": " + ex); }
        Console.WriteLine(results[results.Count - 1]);
    }
    private static string NewPath(string suffix)
    { return Path.Combine(temp, Guid.NewGuid().ToString("N") + suffix); }
    private static string Csv(string text, Encoding encoding)
    {
        string path = NewPath(".csv");
        File.WriteAllText(path, text, encoding);
        return path;
    }
    private static Rdv3Table Table(string text)
    { return Rdv3Table.Read(Csv(text, new UTF8Encoding(false)), "T", Encoding.UTF8, "id"); }
    private static Rdv3PendingStore Pending()
    { return new Rdv3PendingStore(NewPath(".pending")); }
    private static string[] Lines() { return new string[] { "001\tA", "002\tB" }; }
    private static string[] States() { return new string[] { "0", "0" }; }
    private static void Entry(ZipArchive zip, string name, string text)
    {
        using (StreamWriter writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)))
        { writer.Write(text); }
    }
    private static string Worksheet(string cell)
    {
        return "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
            + "<row><c t=\"inlineStr\"><is><t>id</t></is></c></row><row>" + cell + "</row></sheetData></worksheet>";
    }
    private static string Fixture(string cell, bool reordered)
    {
        string path = NewPath(".xlsx");
        using (FileStream stream = File.Create(path))
        using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            Entry(zip, "xl/workbook.xml", "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>"
                + (reordered ? "<sheet name=\"first\" sheetId=\"2\" r:id=\"r2\"/>" : "")
                + "<sheet name=\"LEDGER\" sheetId=\"1\" r:id=\"r1\"/></sheets></workbook>");
            Entry(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "<Relationship Id=\"r2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/></Relationships>");
            Entry(zip, "xl/worksheets/sheet1.xml", Worksheet(cell));
            Entry(zip, "xl/worksheets/sheet2.xml", Worksheet("<c t=\"inlineStr\"><is><t>SECOND-PART-FIRST-TAB</t></is></c>"));
            Entry(zip, "xl/sharedStrings.xml", "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><t>one</t></si></sst>");
        }
        return path;
    }
    private static string[][] ReadSource(string path)
    {
        string[] head; string[][] rows; string warning;
        Rdv3Xlsx.ReadTable(path, out head, out rows, out warning);
        return rows;
    }
    private static void RewriteSheet(string path, Func<string, string> change)
    {
        using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Update))
        {
            ZipArchiveEntry old = zip.GetEntry("xl/worksheets/sheet1.xml");
            string text;
            using (StreamReader reader = new StreamReader(old.Open())) { text = reader.ReadToEnd(); }
            old.Delete(); Entry(zip, "xl/worksheets/sheet1.xml", change(text));
        }
    }
    private static void ReadLedger(string path, string contract)
    {
        string[] lines, states; string warning;
        Rdv3Xlsx.Read(path, new string[] { "id", "value" }, "state", out lines, out states, out warning, contract);
    }

    public static int Run(string directory)
    {
        root = Path.Combine(Path.GetFullPath(directory), "tests", "fixtures");
        temp = Path.Combine(Path.GetTempPath(), "rdv-regression-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(temp);
        passed = failed = 0; results.Clear();
        results.Add("UTC " + DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
        results.Add("OS " + Environment.OSVersion + "; CLR " + Environment.Version);
        try
        {
            Test("json-comments-trailing-comma", delegate { Check(Rdv3Json.Parse("{/*x*/\"a\":[1,],}").Member("a").Count == 1, "JSONC"); });
            Test("json-string-roundtrip", delegate {
                char[] controls = new char[32];
                for (int i = 0; i < controls.Length; i++) { controls[i] = (char)i; }
                string text = "quote=\" slash=\\ <tag> \uD83D\uDE00 " + new string(controls);
                foreach (string json in new string[] {
                    Rdv3Json.Quote(text), Rdv3Config.Q(text), Rdv3WebJson.Q(text) })
                { Check(Rdv3Json.Parse(json).Str == text, "quoted text changed"); }
                Check(Rdv3Json.Parse(Rdv3Json.Quote(null)).Str == "", "null string changed");
                Check(Rdv3Json.Parse(Rdv3Json.Quote("")).Str == "", "empty string changed");
                Check(Rdv3Config.Q("\b\f") == "\"\\u0008\\u000c\"", "settings escape spelling changed");
            });
            Test("json-invalid-number-forms", delegate {
                foreach (string value in new string[] { "01", "1.", "-.1", "1e", "1e9999", "+1" })
                { Throws<Rdv3LoadError>(delegate { Rdv3Json.Parse(value); }); }
            });
            Test("json-unclosed-comment", delegate { Throws<Rdv3LoadError>(delegate { Rdv3Json.Parse("{} /*oops"); }); });
            Test("json-unescaped-control", delegate { Throws<Rdv3LoadError>(delegate { Rdv3Json.Parse("\"a\tb\""); }); });
            Test("json-surrogate-pairs", delegate {
                Throws<Rdv3LoadError>(delegate { Rdv3Json.Parse("\"\\uD800\""); });
                Check(Rdv3Json.Parse("\"\\uD83D\\uDE00\"").Str.Length == 2, "valid pair");
            });
            Test("json-depth-limit", delegate { Throws<Rdv3LoadError>(delegate { Rdv3Json.Parse(new string('[', 140) + "0" + new string(']', 140)); }); });
            Test("json-duplicate-member", delegate { Throws<Rdv3LoadError>(delegate { Rdv3Json.Parse("{\"a\":1,\"a\":2}"); }); });
            Test("whole-input-regex", delegate {
                Rdv3Config c = new Rdv3Config(); c.KeyPattern = "[0-9]{3}";
                Check(c.IsKey("123") && !c.IsKey("x123x") && !c.IsKey("123\n"), "full match");
                c.KeyPattern = "a|ab"; Check(c.IsKey("a") && c.IsKey("ab"), "alternative backtracking");
            });
            Test("bounded-regex-time", delegate {
                Rdv3Config c = new Rdv3Config(); c.KeyPattern = "(a+)+$";
                Stopwatch clock = Stopwatch.StartNew(); Check(!c.IsKey(new string('a', 36) + "!"), "no match");
                Check(clock.ElapsedMilliseconds < 5000, "regex timeout was not bounded");
            });
            Test("pending-disjoint-merge", delegate {
                Rdv3PendingStore a = Pending(), b = Pending();
                a.Set("001", "1", Lines()[0], "0"); b.Set("002", "1", Lines()[1], "0");
                Rdv3PendingApply first = a.PrepareSend(Lines(), States(), 0, "0");
                Rdv3PendingApply next = b.PrepareSend(Lines(), first.States, 0, "0");
                Check(next.States[0] == "1" && next.States[1] == "1", "lost disjoint change");
            });
            Test("pending-same-row-conflict", delegate {
                Rdv3PendingStore p = Pending(); p.Set("001", "2", Lines()[0], "0");
                Rdv3PendingApply a = p.PrepareSend(Lines(), new string[] { "1", "0" }, 0, "0");
                Check(a.States[0] == "1" && a.Resolved.Count == 0 && a.Unmatched[0].Reason == "state-conflict" && p.Count == 1, "overwrote conflict");
                Check(p.Overlay(Lines(), new string[] { "1", "0" }, 0)[0] == "1", "overlay concealed conflict");
            });
            Test("pending-idempotent-retry", delegate {
                Rdv3PendingStore p = Pending(); p.Set("001", "1", Lines()[0], "0");
                Rdv3PendingApply a = p.PrepareSend(Lines(), new string[] { "1", "0" }, 0, "0");
                Check(a.Resolved.Count == 1 && a.FromInitial + a.ToInitial == 0, "double counted retry");
            });
            Test("pending-content-change", delegate {
                Rdv3PendingStore p = Pending(); p.Set("001", "1", Lines()[0], "0");
                Rdv3PendingApply a = p.PrepareSend(new string[] { "001\tNEW", Lines()[1] }, States(), 0, "0");
                Check(a.Unmatched[0].Reason == "changed" && a.States[0] == "0", "content mismatch");
            });
            Test("pending-deleted-row", delegate {
                Rdv3PendingStore p = Pending(); p.Set("001", "1", Lines()[0], "0");
                Check(p.PrepareSend(new string[0], new string[0], 0, "0").Unmatched[0].Reason == "missing", "missing");
            });
            Test("pending-baseline-not-rebased", delegate {
                Rdv3PendingStore p = Pending(); p.Set("001", "1", Lines()[0], "0"); p.Set("001", "2", Lines()[0], "1");
                Check(p.PrepareSend(Lines(), new string[] { "1", "0" }, 0, "0").Unmatched[0].Reason == "state-conflict", "rebased stale intent");
            });
            Test("pending-restart-and-discard", delegate {
                Rdv3PendingStore p = Pending(); p.Set("001", "1", Lines()[0], "0");
                p = new Rdv3PendingStore(p.Path); Check(p.Snapshot()[0].BaselineKnown, "baseline missing after restart");
                p.Remove(new List<string>(new string[] { "001" })); Check(new Rdv3PendingStore(p.Path).Count == 0, "discard not persisted");
            });
            Test("legacy-pending-fails-closed", delegate {
                string path = NewPath(".pending");
                File.WriteAllText(path, "RDV-PENDING-1\r\nMDAx\tMQ==\t" + Rdv3PendingStore.DigestOf(Lines()[0]) + "\r\n", new UTF8Encoding(false));
                Rdv3PendingStore p = new Rdv3PendingStore(path);
                Check(p.PrepareSend(Lines(), States(), 0, "0").Unmatched[0].Reason == "legacy", "legacy overwritten");
                Check(p.PrepareSend(Lines(), new string[] { "1", "0" }, 0, "0").Resolved.Count == 1, "legacy retry not idempotent");
            });
            Test("pending-invalid-digest", delegate {
                string path = NewPath(".pending"); File.WriteAllText(path, "RDV-PENDING-1\r\nMDAx\tMQ==\tbad\r\n");
                Throws<InvalidDataException>(delegate { new Rdv3PendingStore(path); });
            });
            Test("local-session-exclusive", delegate {
                string path = NewPath(".pending");
                using (FileStream a = Rdv3Files.AcquireLocalSession(path))
                { Throws<IOException>(delegate { using (FileStream b = Rdv3Files.AcquireLocalSession(path)) { } }); }
                using (FileStream c = Rdv3Files.AcquireLocalSession(path)) { }
            });
            Test("shared-lock-live-not-stale", delegate {
                Rdv3SharedFiles shared = new Rdv3SharedFiles(NewPath(".xlsx"), "host", "user", "test"); Rdv3LockInfo info;
                using (Rdv3LedgerLock a = shared.TryAcquire(out info))
                {
                    Check(a != null, "first lock");
                    Check(shared.TryAcquire(out info) == null, "second lock acquired");
                    Thread.Sleep(20); long age; Check(!shared.TryRemoveStaleLock(0, out age), "deleted live lock");
                }
                using (Rdv3LedgerLock b = shared.TryAcquire(out info)) { Check(b != null, "lock not released"); }
            });
            Test("csv-plain-and-blank-lines", delegate {
                Rdv3Table t = Table("id,note\r\n\r\n001,A\r\n002,B\r\n");
                Check(t.Rows == 2 && t.SourceRow(0) == 3, "physical row number");
            });
            Test("csv-quotes-comma-doublequote-multiline", delegate {
                Rdv3Table t = Table("\"id\",\"note\"\r\n\"001\",\"a,b\"\r\n\"002\",\"a\"\"b\r\nline\"\r\n");
                Check(t.Rows == 2 && t.Field(0, 1) == "a,b" && t.Field(1, 1) == "a\"b\r\nline", "quoted CSV");
            });
            Test("csv-header-only", delegate { Check(Table("id,note\r\n").Rows == 0 && Table("\"id\",note\r\n").Rows == 0, "empty table"); });
            Test("csv-cr-only", delegate { Check(Table("id,note\r001,A\r002,B\r").Rows == 2, "CR records"); });
            Test("csv-declared-utf16", delegate {
                string path = Csv("id,note\r\n001,\u65e5\u672c\r\n", Encoding.Unicode);
                Check(Rdv3Table.Read(path, "T", Encoding.Unicode, "id").Field(0, 1) == "\u65e5\u672c", "UTF16 decoding");
                Throws<Rdv3DataError>(delegate { Rdv3Table.Read(path, "T", Encoding.UTF8, "id"); });
            });
            Test("csv-invalid-encoding", delegate {
                string path = NewPath(".csv"); File.WriteAllBytes(path, new byte[] { 105,100,44,110,10,48,48,49,44,255,10 });
                Throws<Exception>(delegate { Rdv3Table.Read(path, "T", new UTF8Encoding(false), "id"); });
            });
            Test("csv-malformed-and-control-key", delegate {
                Throws<Rdv3DataError>(delegate { Table("id,note\n001,\"unterminated"); });
                Throws<Rdv3DataError>(delegate { Table("id,note\n001\t,A\n"); });
                Throws<Rdv3DataError>(delegate { Table("id,note\n001,A,B\n"); });
            });
            Test("csv-unicode-variable-keys", delegate {
                Rdv3KeyValidation v = new Rdv3KeyValidation(); v.Ascii = false; v.FixedLength = false;
                string path = Csv("id,note\n\u65e5,A\n\u65e5\u672c,B\n", Encoding.UTF8);
                Check(Rdv3Table.Read(path, "T", Encoding.UTF8, "id", v).Rows == 2, "Unicode config");
            });
            Test("index-no-ascii-replacement-match", delegate {
                Rdv3Index index = new Rdv3Index(Table("id,note\n?,A\n")); List<int> hits;
                Check(index.FindBytes(new byte[] { 255 }, 0, 1, out hits) == 0, "non-ASCII matched question mark");
            });
            Test("xlsx-roundtrip-whitespace-crlf-zero", delegate {
                string path = NewPath(".xlsx"); string value = "001\t  A\r\nB  ";
                Rdv3Xlsx.Write(path, new string[] { "id", "value" }, "state", new string[] { value }, new string[] { "0" }, "t");
                string[] lines, states; Rdv3Xlsx.Read(path, new string[] { "id", "value" }, "state", out lines, out states);
                Check(lines[0] == value && states[0] == "0", "roundtrip text changed");
            });
            Test("xlsx-write-shape-fails-before-replace", delegate {
                string path = NewPath(".xlsx"); File.WriteAllText(path, "KEEP");
                Throws<InvalidDataException>(delegate { Rdv3Xlsx.Write(path, new string[] { "id", "v" }, "state", new string[] { "001" }, new string[] { "0" }, "t"); });
                Check(File.ReadAllText(path) == "KEEP", "destroyed previous target");
            });
            Test("xlsx-invalid-xml-character", delegate {
                Throws<Exception>(delegate { Rdv3Xlsx.Write(NewPath(".xlsx"), new string[] { "id" }, "state", new string[] { "A\0" }, new string[] { "0" }, "t"); });
            });
            Test("xlsx-logical-sheet-order", delegate {
                Check(ReadSource(Fixture("<c t=\"inlineStr\"><is><t>WRONG</t></is></c>", true))[0][0] == "SECOND-PART-FIRST-TAB", "physical sheet1 was used");
            });
            Test("xlsx-header-only-source", delegate {
                string path = Fixture("", false);
                Check(ReadSource(path).Length == 0, "header-only XLSX was not empty");
            });
            Test("xlsx-ledger-additional-sheets-refused", delegate {
                Throws<InvalidDataException>(delegate { ReadLedger(Fixture("", true), null); });
            });
            Test("xlsx-phonetic-text-not-concatenated", delegate {
                string cell = "<c t=\"inlineStr\"><is><r><t>AB</t></r><rPh sb=\"0\" eb=\"2\"><t>PHONETIC</t></rPh></is></c>";
                Check(ReadSource(Fixture(cell, false))[0][0] == "AB", "phonetic text appended");
            });
            Test("xlsx-invalid-shared-index", delegate { Throws<InvalidDataException>(delegate { ReadSource(Fixture("<c t=\"s\"><v>999</v></c>", false)); }); });
            Test("xlsx-no-formula-cache-and-error", delegate {
                Throws<InvalidDataException>(delegate { ReadSource(Fixture("<c><f>1+1</f></c>", false)); });
                Throws<InvalidDataException>(delegate { ReadSource(Fixture("<c t=\"e\"><v>#DIV/0!</v></c>", false)); });
            });
            Test("xlsx-oversized-column", delegate { Throws<InvalidDataException>(delegate { ReadSource(Fixture("<c r=\"ZZZZZZ999\"><v>1</v></c>", false)); }); });
            Test("xlsx-extra-ledger-column", delegate {
                string path = NewPath(".xlsx"); Rdv3Xlsx.Write(path, new string[] { "id", "value" }, "state", Lines(), States(), "t");
                RewriteSheet(path, delegate(string xml) { return xml.Replace("</row></sheetData>", "<c r=\"D3\" t=\"inlineStr\"><is><t>KEEP-ME</t></is></c></row></sheetData>"); });
                Throws<InvalidDataException>(delegate { ReadLedger(path, null); });
            });
            Test("xlsx-storage-contract", delegate {
                string path = NewPath(".xlsx"); string contract = Rdv3PendingStore.DigestOf("schema-a");
                Rdv3Xlsx.Write(path, new string[] { "id", "value" }, "state", Lines(), States(), "t", contract);
                ReadLedger(path, contract);
                Throws<InvalidDataException>(delegate { ReadLedger(path, Rdv3PendingStore.DigestOf("schema-b")); });
            });
            Test("export-no-overwrite-and-atomic-cleanup", delegate {
                string path = NewPath(".csv"); Rdv3Files.WriteNewText(path, "KEEP");
                Throws<IOException>(delegate { Rdv3Files.WriteNewText(path, "REPLACE"); });
                Check(File.ReadAllText(path) == "KEEP" && Directory.GetFiles(temp, ".rdv-export-*.tmp").Length == 0, "export changed target/temp leak");
            });
            Test("export-duplicate-headers-qualified", delegate {
                Rdv3Data d = new Rdv3Data();
                Rdv3ColumnRef a = new Rdv3ColumnRef(); a.Ref = "A.id"; a.Column = "id"; d.Columns.Add(a);
                Rdv3ColumnRef b = new Rdv3ColumnRef(); b.Ref = "B.id"; b.Column = "id"; d.Columns.Add(b);
                string[] headers = Rdv3Files.ExportHeaders(new List<string> { "A.id", "B.id", "$work" }, d, "id");
                Check(headers[0] == "A.id" && headers[1] == "B.id" && headers[2] == "id ($work)", "ambiguous export headings");
            });
            Test("export-formula-mode-explicit", delegate {
                Check(Rdv3Files.CsvCell(" =1+1", true) == "' =1+1", "unsafe formula");
                Check(Rdv3Files.CsvCell("=1+1", false) == "=1+1", "raw export changed");
                Check(Rdv3Files.CsvCell("a,b", false) == "\"a,b\"", "CSV quoting");
            });
            Test("config-sample-and-unlabelled-export", delegate {
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                string reference = c.Data.Columns[c.Data.Columns.Count - 1].Ref;
                c.Data.LabelOrder.Remove(reference); c.Screen.ExportDefaultFields = new string[] { reference };
                c.Screen.Check(c.Data); Check(c.Screen.Work.Trigger == "manual", "unsafe default watch");
            });
            Test("config-save-cas", delegate {
                string path = NewPath(".json"); File.Copy(Path.Combine(root, "settings.json"), path);
                Rdv3Config a = Rdv3Config.Load(path), b = Rdv3Config.Load(path);
                a.KeyPattern = "a|ab"; Check(a.Save(path) == null, "first save failed");
                string after = File.ReadAllText(path); b.KeyPattern = "[0-9]+";
                Check(b.Save(path) != null && File.ReadAllText(path) == after, "stale settings overwrote newer settings");
            });
            Test("config-protected-paths", delegate {
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json")); c.Log = c.Ledger;
                Throws<IOException>(delegate { Rdv3Files.ValidateLayout(c, root, c.SourcePath); });
                c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                string data = Rdv3Files.Full(c.DataDir, root);
                Throws<IOException>(delegate { Rdv3Files.ExportPath(Path.Combine(data, c.Data.Tables[0].File), root, data, Rdv3Files.Full(c.Ledger, root), Rdv3Files.Full(c.Log, root), c.SourcePath, c.Data); });
            });
            Test("config-file-only-job-input-protected", delegate {
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                c.Log = Path.Combine(c.DataDir, "delete.csv");
                Throws<IOException>(delegate { Rdv3Files.ValidateLayout(c, root, c.SourcePath); });
            });
            Test("supplied-data-pipeline-readonly", delegate {
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                Rdv3MergeResult m = Rdv3Ledger.BuildFromCsv(c.Data, c.Data.UpdateJob, Rdv3Files.Full(c.DataDir, root));
                Check(m.Lines != null && m.Lines.Length > 0, "sample pipeline produced no rows");
                Rdv3Ledger.RowMap(m.Lines, c.Data.IdentityCol, "test");
            });
            Test("input-padding-wide-digits-and-quoted-newlines", delegate {
                foreach (Encoding encoding in new Encoding[] { new UTF8Encoding(false), new UTF8Encoding(true), Encoding.GetEncoding(932) })
                {
                    string path = Csv("id ,note \r\n\uff10\uff10\uff11 ,A\u3000\n002,B \r\n", encoding);
                    Rdv3Table t = Rdv3Table.Read(path, "T", encoding, "id");
                    Check(t.Rows == 2 && t.Key(0) == "001" && t.Field(0, 1) == "A" && t.Field(1, 1) == "B", "normalized source");
                }
                Rdv3Table quoted = Table("id,note\n\"001\" ,\"A\r\nB  \" \n");
                Check(quoted.Field(0, 1) == "A\r\nB", "embedded newline lost");
            });
            Test("input-identical-retransmission-counts-and-source-rows", delegate {
                foreach (string quote in new string[] { "", "\"" })
                {
                    Rdv3Table t = Table("id,note\n" + quote + "001" + quote + ",A\n001,A\n,B\n002,C\n");
                    Check(t.Rows == 2 && t.SkippedDuplicateRows == 1 && t.SkippedEmptyRows == 1, "discard counts");
                    Check(t.SourceRow(1) == 5 && t.Key(1) == "002" && new Rdv3Index(t).Keys == 2, "compacted source rows/index");
                }
                Rdv3Table conflict = Table("id,note\n001,A\n001,A\n001,B\n");
                Throws<Rdv3DataError>(delegate { new Rdv3Index(conflict); });
            });
            Test("input-key-errors-identify-source-and-fix", delegate {
                Rdv3KeyValidation v = new Rdv3KeyValidation(); v.SkipEmpty = false;
                v.SettingsPath = "data.tables.T.keyValidation";
                foreach (string row in new string[] { ",B", "0002,B", "\u65e5,B" })
                {
                    string path = Csv("id,note\n001,A\n" + row + "\n", Encoding.UTF8);
                    try { Rdv3Table.Read(path, "T", Encoding.UTF8, "id", v); throw new Exception("bad key accepted"); }
                    catch (Rdv3DataError ex)
                    {
                        Check(ex.Message.Contains(Path.GetFileName(path)) && ex.Message.Contains("3")
                            && ex.Message.Contains("id") && ex.Message.Contains(v.SettingsPath), "missing source or setting");
                    }
                }
            });
            Test("input-invalid-bytes-and-shape-explain-repair", delegate {
                foreach (string prefix in new string[] { "id,note\n001,", "id,note\n\"001\",\"" })
                {
                    string path = NewPath(".csv"); byte[] start = Encoding.UTF8.GetBytes(prefix);
                    byte[] bytes = new byte[start.Length + 1]; Array.Copy(start, bytes, start.Length); bytes[start.Length] = 255;
                    File.WriteAllBytes(path, bytes);
                    try { Rdv3Table.Read(path, "T", Encoding.UTF8, "id"); throw new Exception("invalid encoding accepted"); }
                    catch (Rdv3DataError ex) { Check(ex.Message.Contains("data.encoding") && ex.Message.Contains("FF") && ex.Message.Contains("2"), "encoding repair context"); }
                }
                try { Table("id,note\n001,A,B\n"); throw new Exception("extra column accepted"); }
                catch (Rdv3DataError ex) { Check(ex.Message.Contains("CSV") && ex.Message.Contains("3"), "CSV repair context"); }
            });
            Test("input-search-normalization-without-changing-stored-snapshot", delegate {
                string line = "\uff10\uff10\uff11  \tKEEP";
                Rdv3Index exact = new Rdv3Index(new string[] { line }, new int[] { 0 }, "exact");
                Rdv3Index contains = new Rdv3Index(new string[] { line }, new int[] { 0 }, "contains");
                Check(exact.Find("001")[0] == 0 && exact.Find("\uff10\uff10\uff11 ")[0] == 0, "exact search");
                Check(contains.Find("\uff11")[0] == 0 && line == "\uff10\uff10\uff11  \tKEEP", "snapshot changed");
                Rdv3Config c = new Rdv3Config(); c.KeyPattern = "[0-9]{3}";
                Check(c.IsKey("\uff10\uff10\uff11 ") && !c.IsKey("001\n"), "pattern boundary");
                Check(Rdv3Watch.Candidate("previous\r\n\uff10\uff10\uff11 ") == "001", "UIA input");
            });
            Test("input-numeric-values-reach-expression-aggregate-and-export", delegate {
                string[] values = { "\u00a51,234", "\uffe5\uff11\uff0c\uff12\uff13\uff14", "(500)", "\uff08\uff15\uff10\uff10\uff09" };
                decimal[] expected = { 1234m, 1234m, -500m, -500m };
                for (int i = 0; i < values.Length; i++)
                {
                    decimal value; Check(Rdv3Input.TryNumber(values[i], out value) && value == expected[i], "money conversion");
                    Check(Rdv3Expression.Compile("T.amount + 1", new string[] { "T.amount" }).Evaluate(new string[] { values[i] })
                        == (expected[i] + 1m).ToString(CultureInfo.InvariantCulture), "expression value");
                }
                decimal unused;
                Check(!Rdv3Input.TryNumber("(-500)", out unused) && !Rdv3Input.TryNumber("12oops", out unused), "ambiguous numeric input");
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                Rdv3PreparedProcess p = new Rdv3PreparedProcess(); p.Data = c.Data; p.Job = new Rdv3ProcessJobDef();
                Rdv3Relation input = new Rdv3Relation(); input.Columns = new string[] { "T.amount" };
                input.Rows.Add(new string[] { values[0] }); input.Rows.Add(new string[] { values[2] }); p.Inputs.Add("T", input);
                Rdv3ProcessStepDef sum = new Rdv3ProcessStepDef(); sum.Operation = "aggregate"; sum.Target1 = "T"; sum.Output = "S";
                sum.Aggregates.Add(new Rdv3ProcessAggregateDef { Function = "sum", Column = "T.amount", As = "total" }); p.Job.Steps.Add(sum);
                Check(Rdv3Process.Execute(p, new string[0], new string[0], "", false).Lines[0] == "734", "aggregate sum");
                Rdv3ExportFilter filter = new Rdv3ExportFilter { Field = "B.b_qty", Operator = "range", First = "(600)", Last = "\u00a52,000" };
                string[] row = new string[c.Data.Columns.Count]; row[c.Data.IndexOf("B.b_qty")] = "(500)";
                Check(filter.Matches(c.Data, row, ""), "export range lost accepted money");
            });
            Test("input-dirty-join-preserves-990-records-and-reports-discards", delegate {
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                Rdv3MergeResult clean = Rdv3Ledger.BuildFromCsv(c.Data, Path.Combine(root, "data"));
                string dir = NewPath("-data"); Directory.CreateDirectory(dir);
                foreach (string file in Directory.GetFiles(Path.Combine(root, "data"), "*.csv")) { File.Copy(file, Path.Combine(dir, Path.GetFileName(file))); }
                string bPath = Path.Combine(dir, "tableB.csv"); string[] lines = File.ReadAllLines(bPath, Encoding.UTF8);
                for (int i = 1; i < lines.Length; i++)
                {
                    string[] fields = lines[i].Split(',');
                    for (int col = 0; col < 2; col++)
                    {
                        char[] digits = fields[col].ToCharArray();
                        for (int k = 0; k < digits.Length; k++) { if (digits[k] >= '0' && digits[k] <= '9') { digits[k] = (char)('\uff10' + digits[k] - '0'); } }
                        fields[col] = new string(digits) + " ";
                    }
                    lines[i] = string.Join(",", fields);
                }
                string blank = new string(',', lines[0].Split(',').Length - 1);
                File.WriteAllText(bPath, string.Join("\r\n", lines) + "\n" + lines[1] + "\n" + blank + "\n" + blank + "\n", Encoding.UTF8);
                Rdv3MergeResult dirty = Rdv3Ledger.BuildFromCsv(c.Data, dir);
                Check(dirty.Rows == 990 && string.Join("\n", dirty.Lines) == string.Join("\n", clean.Lines), "dirty joins changed records");
                Check(dirty.Warnings.Exists(delegate(string s) { return s.Contains("tableB.csv") && s.Contains("12") && s.Contains("990"); }), "missing discard report");
                object[] args = { c.Data, c.Data.UpdateJob, dir, NewPath(".xlsx"), false };
                string body = (string)typeof(Rdv3ProcessForm).GetMethod("Build", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static).Invoke(null, args);
                Check((bool)args[4] && body.Contains("12") && body.Contains("990"), "process dialog lost accepted input report");
                string deletePath = Path.Combine(dir, "delete.csv"); File.AppendAllText(deletePath, ",\n,\n", Encoding.UTF8);
                Rdv3ProcessJobDef deletion = c.Data.Jobs.Find(delegate(Rdv3ProcessJobDef j) { return j.Kind == "delete"; });
                Rdv3DeleteResult deleted = Rdv3Ledger.ApplyDelete(c.Data, deletion, dir, clean.Lines, new string[clean.Rows], c.Screen.Work.InitialStored);
                Check(deleted.Warnings.Exists(delegate(string s) { return s.Contains("delete.csv") && s.Contains("2"); }), "delete discarded input silently");
            });
            Test("config-optional-sections-save-and-external-edit-guard", delegate {
                Rdv3Json json = Rdv3Json.Parse(File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8));
                string text = "{/*keep-root-comment*/\"schema\":3,\"data\":" + File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8).Substring(json.Member("data").Start, json.Member("data").End - json.Member("data").Start)
                    + ",\"screen\":" + File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8).Substring(json.Member("screen").Start, json.Member("screen").End - json.Member("screen").Start) + ",}";
                string path = NewPath(".json"); File.WriteAllText(path, text, Encoding.UTF8);
                Rdv3Config a = Rdv3Config.Load(path), b = Rdv3Config.Load(path);
                Check(a.Targets.Count == 0 && a.PollMs == 40 && a.DataDir == "data" && a.CheckTimeoutMs == 180000, "missing defaults");
                a.KeyPattern = ".+";
                Check(a.Save(path) == null && Rdv3Config.Load(path).KeyPattern == ".+", "save missing objects");
                string saved = File.ReadAllText(path, Encoding.UTF8);
                Check(saved.Contains("/*keep-root-comment*/") && b.Save(path) != null && File.ReadAllText(path, Encoding.UTF8) == saved, "comment/CAS lost");
                Check(saved.Contains("\"paths\"") && saved.Contains("\"search\"") && saved.Contains("\"watch\""), "objects not persisted");
            });
            Test("config-text-type-and-enum-padding", delegate {
                string text = File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8).Replace("\"type\": \"number\"", "\"type\": \" Text \"");
                Rdv3Config c = Rdv3Config.Parse(text);
                Check(c.Data.TypeOf("B.b_qty") == null, "text uses numeric filter");
                Check(Rdv3Ledger.BuildFromCsv(c.Data, Path.Combine(root, "data")).Rows == 990, "explicit text rejected");
                Check(Rdv3Json.Parse("{\"v\":\" Merge \"}").Word("v", "", "merge", "replace") == "merge", "enum normalization");
            });
            Test("apply-reloads-state-after-lock-wait", delegate {
                ApplyFixture f = new ApplyFixture();
                f.Save(Lines(), States());
                Rdv3LockInfo owner;
                Rdv3LedgerLock other = f.Shared.TryAcquire(out owner);
                Check(other != null, "other writer lock");
                Rdv3ApplyOutcome outcome = null;
                using (ManualResetEvent waiting = new ManualResetEvent(false))
                {
                    Thread update = new Thread(delegate() {
                        outcome = f.Store.Apply(f.Source(new string[] { "001\tNEW", "002\tB" }), Lines(), "apply",
                            delegate {
                                Rdv3LedgerLock lease = null;
                                Stopwatch clock = Stopwatch.StartNew();
                                while (lease == null && clock.ElapsedMilliseconds < 5000) {
                                    lease = f.Shared.TryAcquire(out owner);
                                    if (lease == null) { waiting.Set(); Thread.Sleep(10); }
                                }
                                return lease;
                            }, f.Trace, f.Warn);
                    });
                    update.IsBackground = true;
                    try {
                        update.Start();
                        Check(waiting.WaitOne(5000), "update did not wait for the other writer");
                        f.Save(Lines(), new string[] { "0", "2" });
                    }
                    finally { other.Dispose(); Check(update.Join(5000), "update worker did not return"); }
                }
                Check(outcome != null && outcome.Error == null && outcome.Committed, "apply did not commit");
                Rdv3LedgerSnapshot saved = f.Store.Read(f.Head);
                Check(saved.Lines[0] == "001\tNEW" && saved.States[1] == "2", "lost the send made while waiting");
            });
            Test("apply-rejects-changed-preview-and-creation-races", delegate {
                ApplyFixture f = new ApplyFixture();
                f.Save(new string[] { "001\tOTHER", "002\tB" }, States());
                byte[] before = File.ReadAllBytes(f.Path);
                Rdv3ApplyOutcome changed = f.Apply(f.Source(Lines()), Lines());
                Check(changed.Error != null && !changed.CanAdopt && !changed.Committed, "stale preview was accepted");
                Check(Convert.ToBase64String(File.ReadAllBytes(f.Path)) == Convert.ToBase64String(before), "stale preview changed file");
                Check(!f.Apply(f.Source(Lines()), null).CanAdopt, "created-after-preview file was replaced");
                ApplyFixture missing = new ApplyFixture();
                Check(!missing.Apply(missing.Source(Lines()), Lines()).CanAdopt && !File.Exists(missing.Path), "deleted-after-preview file was rebuilt");
            });
            Test("apply-read-failure-never-overwrites-ledger", delegate {
                ApplyFixture f = new ApplyFixture();
                File.WriteAllText(f.Path, "unreadable ledger");
                Rdv3ApplyOutcome outcome = f.Apply(f.Source(Lines()), Lines());
                Check(!outcome.Committed && !outcome.CanAdopt && outcome.Error != null, "corrupt ledger adopted");
                Check(File.ReadAllText(f.Path) == "unreadable ledger", "corrupt ledger overwritten");
                Rdv3LockInfo owner;
                using (Rdv3LedgerLock lease = f.Shared.TryAcquire(out owner)) { Check(lease != null, "failed read leaked lease"); }
            });
            Test("apply-write-failure-keeps-old-file", delegate {
                ApplyFixture f = new ApplyFixture();
                f.Save(Lines(), States());
                string bytes = Convert.ToBase64String(File.ReadAllBytes(f.Path));
                using (FileStream held = new FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    Rdv3ApplyOutcome outcome = f.Apply(f.Source(new string[] { "001\tNEW", "002\tB" }), Lines());
                    Check(outcome.Error != null && !outcome.Committed && !outcome.CanAdopt, "failed replacement reported saved");
                    Check(Convert.ToBase64String(File.ReadAllBytes(f.Path)) == bytes, "failed replacement damaged file");
                }
            });
            Test("apply-marker-failure-retains-committed-result", delegate {
                ApplyFixture f = new ApplyFixture();
                f.Save(Lines(), States());
                File.WriteAllText(f.Shared.MarkerPath, "unreadable marker");
                Rdv3ApplyOutcome outcome = f.Apply(f.Source(new string[] { "001\tNEW", "002\tB" }), Lines());
                Check(outcome.Committed && outcome.CanAdopt && outcome.Error != null && outcome.Marker == null, "lost committed outcome");
                Check(f.Store.Read(f.Head).Lines[0] == "001\tNEW", "commit was rolled back after marker failure");
                Rdv3LockInfo owner;
                using (Rdv3LedgerLock lease = f.Shared.TryAcquire(out owner)) { Check(lease != null, "marker failure leaked lease"); }
            });
            Test("apply-unchanged-does-not-write-or-notify", delegate {
                ApplyFixture f = new ApplyFixture();
                f.Save(Lines(), new string[] { "1", "2" });
                File.WriteAllText(f.Shared.MarkerPath, "unreadable marker");
                string bytes = Convert.ToBase64String(File.ReadAllBytes(f.Path));
                Rdv3ApplyOutcome outcome = f.Apply(f.Source(Lines()), Lines());
                Check(outcome.Error == null && outcome.CanAdopt && !outcome.Committed && outcome.Marker == null, "unchanged apply tried to notify");
                Check(Convert.ToBase64String(File.ReadAllBytes(f.Path)) == bytes, "unchanged apply rewrote ledger");
            });
            Test("ledger-read-keeps-state-identity-and-contract-checks", delegate {
                ApplyFixture f = new ApplyFixture();
                f.Save(Lines(), new string[] { "unknown", "0" });
                Throws<InvalidDataException>(delegate { f.Store.Read(f.Head); });
                f.Save(new string[] { "001\tA", "001\tB" }, States());
                Throws<Rdv3DataError>(delegate { f.Store.Read(f.Head); });
                f.Save(Lines(), States());
                f.Work.States.Add(new Rdv3StateDef { Id = "extra", Stored = "3" });
                Rdv3LedgerStore incompatible = new Rdv3LedgerStore(f.Path, f.Data, f.Work, f.Shared);
                Throws<InvalidDataException>(delegate { incompatible.Read(f.Head); });
            });
            Test("apply-prepared-pipeline-keeps-general-path", delegate {
                Rdv3Config c = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                string path = NewPath(".xlsx");
                Rdv3SharedFiles shared = new Rdv3SharedFiles(path, "test", "user", "pipeline");
                Rdv3LedgerStore store = new Rdv3LedgerStore(path, c.Data, c.Screen.Work, shared);
                string dataDir = Rdv3Files.Full(c.DataDir, root);
                Rdv3MergeResult source = Rdv3Ledger.BuildFromCsv(c.Data, c.Data.UpdateJob, dataDir);
                source.Prepared = Rdv3Process.Prepare(c.Data, c.Data.UpdateJob, dataDir);
                Rdv3LockInfo owner;
                Rdv3ApplyOutcome outcome = store.Apply(source, null, "pipeline",
                    delegate { return shared.TryAcquire(out owner); }, delegate(string stage, string detail) { }, delegate(string warning) { });
                Check(outcome.Error == null && outcome.Committed && store.Read(source.Head).Lines.Length == source.Lines.Length, "general pipeline failed");
            });
            Test("reset-notice-unites-shared-and-local-without-dropping-rows", delegate {
                string[] before = { "001\told", "002\told", "003\told", "004\told", "005\tsame", "006\told", "007\tremoved" };
                Rdv3PendingStore pending = Pending();
                pending.Set("002", "2", before[1], "0");
                string[] effectiveBefore = pending.Overlay(before, new string[] { "0", "0", "2", "0", "2", "2", "2" }, 0);
                Rdv3UpdateResult update = new Rdv3UpdateResult();
                update.Lines = new string[] { "003\tnew", "002\tnew", "001\tnew", "004\tnew", "005\tsame", "006\tnew", "008\tadded" };
                update.States = new string[] { "0", "0", "0", "0", "0", "2", "0" };
                update.ResetLines.Add("001\tnew"); update.ResetLines.Add("003\tnew");
                string[] effectiveAfter = pending.Overlay(update.Lines, update.States, 0);
                Rdv3ResetNotice notice = new Rdv3ResetNotice(0, "0");
                List<string> rows = notice.AfterUpdate(update, before, effectiveBefore, effectiveAfter);
                Check(string.Join("|", rows.ToArray()) == "001\tnew|002\tnew|003\tnew", "lost, duplicated or invented reset notification");
                Check(string.Join("|", notice.ChangedRows(before, effectiveBefore, update.Lines, effectiveAfter).ToArray()) == "003\tnew|002\tnew", "reload/deletion notification order changed");
                Check(pending.Count == 1 && pending.PrepareSend(update.Lines, update.States, 0, "0").Unmatched[0].Reason == "changed", "notification consumed conflicting pending");
            });
            Test("write-guard-local-outcome-controls-close", delegate {
                for (int i = 0; i < 2; i++) {
                    Rdv3WriteGuard guard = new Rdv3WriteGuard();
                    guard.Begin(false);
                    Check(!guard.TryClose() && guard.Pending && !guard.Closing, "closed before local outcome");
                    bool requested;
                    Check(guard.Complete(out requested) && requested && !guard.Pending, "decided write did not release close");
                    Check(guard.TryClose() && guard.Closing, "close still blocked after outcome");
                }
            });
            Test("write-guard-close-and-acquisition-are-exclusive", delegate {
                for (int i = 0; i < 100; i++) {
                    Rdv3WriteGuard guard = new Rdv3WriteGuard(); guard.Begin(true);
                    bool entered = false, closed = false;
                    using (ManualResetEvent start = new ManualResetEvent(false)) {
                        Thread writer = new Thread(delegate() { start.WaitOne(); entered = guard.TryEnterSharedWrite(); });
                        Thread closer = new Thread(delegate() { start.WaitOne(); closed = guard.TryClose(); });
                        writer.Start(); closer.Start(); start.Set();
                        Check(writer.Join(5000) && closer.Join(5000), "guard race stalled");
                    }
                    Check(entered != closed, "close and shared write both won or both failed");
                    Check(!entered || (guard.Pending && !guard.Closing), "write lost its close hold");
                    Check(!closed || !guard.TryEnterSharedWrite(), "acquired a shared write after close");
                }
            });
            Test("write-guard-overdue-worker-remains-held-until-return", delegate {
                Rdv3WriteGuard guard = new Rdv3WriteGuard(); guard.Begin(false);
                Rdv3Worker worker = new Rdv3Worker();
                using (ManualResetEvent started = new ManualResetEvent(false))
                using (ManualResetEvent release = new ManualResetEvent(false))
                using (ManualResetEvent returned = new ManualResetEvent(false)) {
                    Rdv3Job job = new Rdv3Job(); job.Kind = "apply"; job.RunId = "late"; job.TimeoutMs = 1;
                    job.Work = delegate { started.Set(); release.WaitOne(); returned.Set(); };
                    worker.Start(); worker.Post(job);
                    try {
                        Check(started.WaitOne(5000), "worker did not start");
                        Thread.Sleep(20);
                        Check(worker.TakeOverdue() == job && worker.TakeOverdue() == null, "deadline did not report exactly once");
                        Check(!guard.TryClose() && guard.Pending && !returned.WaitOne(0), "deadline released or aborted unfinished write");
                        release.Set(); Check(returned.WaitOne(5000), "late write did not return");
                        bool requested; guard.Complete(out requested);
                        Check(requested && guard.TryClose(), "late return did not release close");
                    }
                    finally { release.Set(); worker.Stop(); }
                }
            });
            Test("eight-thread-shared-ledger-simulation", ConcurrentWriters);
            ConfigurationInputs();
        }
        finally
        {
            results.Add("TOTAL passed=" + passed + " failed=" + failed);
            string resultPath = Path.Combine(Path.GetFullPath(directory), "tests", "results", "core-results.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(resultPath));
            File.WriteAllLines(resultPath, results.ToArray(), new UTF8Encoding(false));
            try { Directory.Delete(temp, true); } catch (Exception ex) { Console.WriteLine("TEMP CLEANUP: " + ex.Message); }
        }
        Console.WriteLine("TOTAL passed=" + passed + " failed=" + failed);
        return failed == 0 ? 0 : 1;
    }

    private static void ConfigurationInputs()
    {
        Test("encoding-per-input-fast-and-general-join", delegate {
            string dir = NewPath("-mixed"); Directory.CreateDirectory(dir);
            string clean = NewPath("-utf8"); Directory.CreateDirectory(clean);
            foreach (string file in new string[] { "tableA.csv", "tableB.csv", "tableC.csv", "delete.csv" })
            {
                string[] rows = File.ReadAllLines(Path.Combine(root, "data", file), Encoding.UTF8);
                for (int r = 0; r < rows.Length; r++)
                {
                    string[] cells = rows[r].Split(',');
                    if (r == 0 && file == "tableA.csv") { cells[2] = "\u540d\u79f0"; }
                    if (r > 0)
                    {
                        if ((file == "tableA.csv" || file == "tableB.csv") && cells[0].Length > 0) { cells[0] = "\u7fa4" + cells[0]; }
                        if (file == "tableA.csv") { cells[2] = "\u65e5\u672c" + cells[2]; }
                        if (file == "tableB.csv") { cells[2] = "\u53c2\u7167" + cells[2]; }
                        if (file == "delete.csv") { cells[1] = "\u53c2\u7167" + cells[1]; }
                    }
                    rows[r] = string.Join(",", cells);
                }
                File.WriteAllLines(Path.Combine(clean, file), rows, new UTF8Encoding(false));
                File.WriteAllLines(Path.Combine(dir, file), rows, file == "tableA.csv" || file == "delete.csv" ? Encoding.GetEncoding(932) : new UTF8Encoding(file == "tableC.csv"));
            }
            string text = File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8)
                .Replace("A.a_name", "A.\u540d\u79f0")
                .Replace("\"key\": \"key1\"", "\"key\": \"key1\", \"keyValidation\": { \"characters\": \"unicode\" }")
                .Replace("\"column\": \"b_ref\"", "\"column\": \"b_ref\", \"keyValidation\": { \"characters\": \"unicode\" }");
            Rdv3Config canonical = Rdv3Config.Parse(text);
            Rdv3Config mixed = Rdv3Config.Parse(text.Replace("\"file\": \"tableA.csv\"", "\"file\": \"tableA.csv\", \"encoding\": \"shift_jis\"")
                .Replace("\"file\": \"delete.csv\"", "\"file\": \"delete.csv\", \"encoding\": \"shift_jis\""));
            Check(mixed.Data.Tables[0].Enc.CodePage == 932 && mixed.Data.Tables[1].Enc.CodePage == 65001, "table override or fallback");
            Check(mixed.Data.Jobs[0].Inputs[0].Enc.CodePage == 932 && mixed.Data.Jobs[1].Inputs[0].Enc.CodePage == 932, "job inherited/file encoding");
            Check(Rdv3Table.ReadHead(Path.Combine(dir, "tableA.csv"), mixed.Data.Tables[0].Enc, mixed.Data.Tables[0].EncodingSetting)[2] == "\u540d\u79f0", "startup header encoding");
            Rdv3MergeResult expected = Rdv3Ledger.BuildFromCsv(canonical.Data, clean);
            Rdv3MergeResult actual = Rdv3Ledger.BuildFromCsv(mixed.Data, dir);
            Check(mixed.Data.UpdateJob.FastJoinPlan && actual.Rows == 990, "fast path not exercised");
            Check(string.Join("\n", actual.Lines) == string.Join("\n", expected.Lines), "mixed byte-key join changed records");
            mixed.Data.UpdateJob.FastJoinPlan = false;
            Check(string.Join("\n", Rdv3Ledger.BuildFromCsv(mixed.Data, dir).Lines) == string.Join("\n", expected.Lines), "general join encoding");
            string[] initial = new string[actual.Rows]; for (int i = 0; i < initial.Length; i++) { initial[i] = "FALSE"; }
            Rdv3DeleteResult a = Rdv3Ledger.ApplyDelete(canonical.Data, canonical.Data.Jobs[1], clean, expected.Lines, initial, "FALSE");
            Rdv3DeleteResult b = Rdv3Ledger.ApplyDelete(mixed.Data, mixed.Data.Jobs[1], dir, actual.Lines, initial, "FALSE");
            Check(a.Deleted > 0 && a.Deleted == b.Deleted && string.Join("\n", a.Lines) == string.Join("\n", b.Lines), "file-only mixed delete");
            object[] args = { mixed.Data, mixed.Data.Jobs[1], dir, NewPath(".xlsx"), false };
            typeof(Rdv3ProcessForm).GetMethod("Build", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static).Invoke(null, args);
            Check((bool)args[4], "process preview rejected encoded file inputs");
        });
        Test("encoding-errors-point-to-the-effective-setting", delegate {
            string text = File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8);
            try { Rdv3Config.Parse(text.Replace("\"file\": \"tableA.csv\"", "\"file\": \"tableA.csv\", \"encoding\": \"not-an-encoding\"")); throw new Exception("invalid encoding accepted"); }
            catch (Rdv3LoadError ex) { Check(ex.Message.Contains("data.tables.A.encoding") && ex.Message.Contains("utf-8"), "missing encoding repair"); }
            string path = Csv("id,value\n001,\u65e5\u672c\n", new UTF8Encoding(true));
            try { Rdv3Table.Read(path, "A", Encoding.GetEncoding(932), "id", null, "data.tables.A.encoding"); throw new Exception("BOM mismatch accepted"); }
            catch (Rdv3DataError ex) { Check(ex.Message.Contains("data.tables.A.encoding"), "BOM repair points to global default"); }
        });
        Test("date-three-explicit-column-formats", delegate {
            string[] formats = { "yyyyMMdd", "yyyy-MM-dd", "yyyy/MM/dd" };
            string[] values = { "20260909", "2026-09-09", "2026/09/09" };
            for (int i = 0; i < formats.Length; i++)
            {
                DateTime date;
                Rdv3ColumnTypeDef type = new Rdv3ColumnTypeDef { Type = "date", Format = formats[i] };
                Check(type.TryDate(values[i], out date) && date == new DateTime(2026, 9, 9), "declared date format");
                Check(!type.TryDate(values[(i + 1) % 3], out date), "date format inferred instead of declared");
            }
        });
        Test("composite-table-identities-discard-only-identical-rows", delegate {
            CompositeFixture f = new CompositeFixture();
            Rdv3Table t = Rdv3Table.Read(Path.Combine(f.Dir, "T.csv"), "T", Encoding.GetEncoding(932), new string[] { "id", "part" }, f.Config.Data.Tables[0].KeyValidation);
            Check(t.Rows == 3 && t.SkippedEmptyRows == 1 && t.SkippedDuplicateRows == 1 && new Rdv3Index(t).Keys == 3, "tuple counts");
            Check(t.Key(0) != t.Key(1) && t.Key(0) != t.Key(2), "tuple concatenation collided");
            string conflict = Csv("id,part,value\nA,B,one\nA,B,two\n", Encoding.UTF8);
            try { new Rdv3Index(Rdv3Table.Read(conflict, "T", Encoding.UTF8, new string[] { "id", "part" }, null)); throw new Exception("conflicting tuple accepted"); }
            catch (Rdv3DataError ex) { Check(ex.Message.Contains("id / part") && ex.Message.Contains("2") && ex.Message.Contains("3"), "tuple conflict location"); }
        });
        Test("composite-key-rules-apply-to-each-component", delegate {
            Rdv3KeyValidation rule = new Rdv3KeyValidation();
            string skipped = Csv("id,part\nTOO-LONG,\nA,01\nB,02\n", Encoding.UTF8);
            Check(Rdv3Table.Read(skipped, "T", Encoding.UTF8, new string[] { "id", "part" }, rule).Rows == 2, "skipped row set fixed width");
            string bad = Csv("id,part\nA,01\nB,002\n", Encoding.UTF8);
            try { Rdv3Table.Read(bad, "T", Encoding.UTF8, new string[] { "id", "part" }, rule); throw new Exception("part width accepted"); }
            catch (Rdv3DataError ex) { Check(ex.Message.Contains("part") && ex.Message.Contains("variable"), "component width repair"); }
            Throws<Rdv3DataError>(delegate { Rdv3Table.Read(Csv("id,part\nA,\u3042\n", Encoding.UTF8), "T", Encoding.UTF8, new string[] { "id", "part" }, rule); });
        });
        Test("composite-join-and-headless-result-show-unmatched-sides", delegate {
            CompositeFixture f = new CompositeFixture();
            Rdv3Json report = Rdv3Json.Parse(Rdv3Headless.Evaluate(f.Config, f.Dir, f.Dir, true, ""));
            Check(report.Member("summary").Member("rows").Num == 3 && report.Member("summary").Member("skippedEmpty").Num == 1 && report.Member("summary").Member("skippedDuplicate").Num == 1, "report counts");
            Rdv3Json join = report.Member("joins").At(0);
            Check(join.Member("unmatchedLeft").Num == 1 && join.Member("unmatchedRight").Num == 1, "join counts");
            Check(report.Member("rows").At(0).At(3).Str == "first" && report.Member("rows").At(1).At(3).Str == "second" && report.Member("rows").At(2).At(3).Str == "", "wrong pair joined");
            Check(!File.Exists(f.Config.Ledger) && !File.Exists(f.Config.Ledger + ".lock"), "headless wrote a shared file");
            Check(Rdv3Ledger.BuildFromCsv(f.Config.Data, f.Dir).Rows == 3, "window update preparation disagrees");
        });
        Test("composite-extract-tests-the-whole-pair", delegate {
            CompositeFixture f = new CompositeFixture();
            Rdv3ProcessJobDef job = f.Config.Data.UpdateJob;
            job.Steps.Clear();
            job.Steps.Add(new Rdv3ProcessStepDef { Operation = "extract", Target1 = "T", Target2 = "U", Output = "missing", Condition = "exclude", Keys = new string[] { "T.id", "U.id" }, KeyGroups = new string[][] { new string[] { "T.id", "T.part" }, new string[] { "U.id", "U.part" } } });
            Rdv3ProcessResult result = Rdv3Process.Run(f.Config.Data, job, f.Dir, new string[0], new string[0], "FALSE");
            Check(result.Lines.Length == 1 && result.Lines[0].StartsWith("AB\tD\t"), "cross-pair exclusion");
        });
        Test("composite-pending-keeps-digest-baseline-and-idempotence", delegate {
            int[] ids = { 0, 1 }; string[] lines = { "AB\tC\tone", "A\tBC\ttwo", "AB\tD\tthree" };
            string[] states = { "FALSE", "HOLD", "FALSE" };
            Rdv3PendingStore pending = Pending();
            pending.Set(Rdv3Key.FromLine(lines[0], ids), "TRUE", lines[0], "FALSE");
            pending.Set(Rdv3Key.FromLine(lines[1], ids), "FALSE", lines[1], "HOLD");
            Check(pending.Overlay(lines, states, ids)[2] == "FALSE", "third tuple changed");
            Rdv3PendingApply sent = pending.PrepareSend(lines, states, ids, "FALSE");
            Check(sent.Resolved.Count == 2 && sent.Unmatched.Count == 0 && sent.States[0] == "TRUE" && sent.States[1] == "FALSE", "distinct tuple sends");
            Check(pending.PrepareSend(lines, sent.States, ids, "FALSE").Resolved.Count == 2, "tuple retry not idempotent");
            string[] changed = (string[])lines.Clone(); changed[0] += "changed";
            Check(pending.PrepareSend(changed, states, ids, "FALSE").Unmatched[0].Reason == "changed", "content conflict lost");
            states[0] = "HOLD";
            Check(pending.PrepareSend(lines, states, ids, "FALSE").Unmatched[0].Reason == "state-conflict", "baseline conflict lost");
        });
        Test("composite-update-reset-and-contract-readback", delegate {
            CompositeFixture f = new CompositeFixture();
            Rdv3ProcessResult first = Rdv3Process.Run(f.Config.Data, f.Config.Data.UpdateJob, f.Dir, new string[0], new string[0], "FALSE");
            string path = NewPath(".xlsx"); string[] states = { "TRUE", "HOLD", "FALSE" };
            Rdv3Xlsx.Write(path, f.Config.Data.Head, f.Config.Screen.Work.Column, first.Lines, states, "baseline", Rdv3Files.StorageContract(f.Config.Data, f.Config.Screen.Work));
            byte[] before = File.ReadAllBytes(path);
            string u = Path.Combine(f.Dir, "U.csv"); File.WriteAllText(u, File.ReadAllText(u).Replace("first", "changed"), new UTF8Encoding(true));
            Rdv3Json report = Rdv3Json.Parse(Rdv3Headless.Evaluate(f.Config, f.Dir, f.Dir, true, path));
            Check(report.Member("summary").Member("resetRows").Num == 1 && report.Member("states").At(0).Str == "FALSE" && report.Member("states").At(1).Str == "HOLD", "wrong tuple reset");
            Check(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(path)), "headless overwrote baseline");
            string contract = Rdv3Files.StorageContract(f.Config.Data, f.Config.Screen.Work);
            Array.Reverse(f.Config.Data.IdentityCols);
            Check(contract != Rdv3Files.StorageContract(f.Config.Data, f.Config.Screen.Work), "tuple order not in contract");
            Throws<InvalidDataException>(delegate { new Rdv3LedgerStore(path, f.Config.Data, f.Config.Screen.Work, null).Read(f.Config.Data.Head); });
        });
        Test("single-key-array-preserves-existing-storage-contract", delegate {
            string text = File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8);
            Rdv3Config one = Rdv3Config.Parse(text);
            Rdv3Config array = Rdv3Config.Parse(text.Replace("\"key\": \"key2\"", "\"key\": [\"key2\"]").Replace("\"identity\": \"B.key2\"", "\"identity\": [\"B.key2\"]"));
            Check(Rdv3Files.StorageContract(one.Data, one.Screen.Work) == Rdv3Files.StorageContract(array.Data, array.Screen.Work), "single array changed stored contract");
            Check(Rdv3Key.FromLine("00000001\tvalue", new int[] { 0 }) == "00000001", "legacy pending representation changed");
        });
        Test("headless-validation-output-and-input-protection", delegate {
            CompositeFixture f = new CompositeFixture(); string output = NewPath(".json");
            Check(Rdv3Headless.Run(f.Dir, f.Config.SourcePath, f.Dir, false, "", "") == 0, "validate failed");
            Check(Rdv3Headless.Run(f.Dir, f.Config.SourcePath, f.Dir, true, output, "") == 0, "run failed");
            string original = File.ReadAllText(output);
            Check(Rdv3Headless.Run(f.Dir, f.Config.SourcePath, f.Dir, true, output, "") != 0 && File.ReadAllText(output) == original, "report overwritten");
            Check(Rdv3Headless.Run(f.Dir, f.Config.SourcePath, f.Dir, true, f.Config.SourcePath, "") != 0, "config writable as report");
            File.WriteAllText(Path.Combine(f.Dir, "U.csv"), "wrong,header\n1,2\n");
            string failed = NewPath(".json");
            Check(Rdv3Headless.Run(f.Dir, f.Config.SourcePath, f.Dir, true, failed, "") != 0 && !File.Exists(failed), "invalid inputs published success");
        });
    }

    private sealed class CompositeFixture
    {
        public readonly string Dir;
        public readonly Rdv3Config Config;
        public CompositeFixture()
        {
            Dir = NewPath("-tuple"); Directory.CreateDirectory(Dir);
            File.WriteAllText(Path.Combine(Dir, "T.csv"), "id,part,amount\nAB,C,100\nA,BC,200\nAB,D,300\nAB,C,100\nLONG,,400\n", Encoding.GetEncoding(932));
            File.WriteAllText(Path.Combine(Dir, "U.csv"), "id,part,value\nAB,C,first\nA,BC,second\nAB,E,extra\n", new UTF8Encoding(true));
            string rule = "\"keyValidation\":{\"characters\":\"unicode\",\"length\":\"variable\"}";
            string data = "{\"tables\":{\"T\":{\"file\":\"T.csv\",\"encoding\":\"shift_jis\",\"key\":[\"id\",\"part\"]," + rule + "},\"U\":{\"file\":\"U.csv\",\"key\":[\"id\",\"part\"]," + rule + "}},"
                + "\"labels\":{\"T.id\":\"ID\",\"T.part\":\"Part\",\"T.amount\":\"Amount\",\"U.id\":\"ID\",\"U.part\":\"Part\",\"U.value\":\"Value\",\"joined\":\"Joined\",\"ledger\":\"Ledger\"},"
                + "\"jobs\":[{\"id\":\"update\",\"kind\":\"update\",\"inputs\":[{\"table\":\"T\"},{\"table\":\"U\"}],\"steps\":["
                + "{\"operation\":\"join\",\"target1\":\"T\",\"target2\":\"U\",\"keys\":[[\"T.id\",\"T.part\"],[\"U.id\",\"U.part\"]],\"condition\":\"left\",\"output\":\"joined\"},"
                + "{\"operation\":\"merge\",\"target1\":\"joined\",\"target2\":\"ledger\",\"keys\":[[\"T.id\",\"T.part\"],[\"T.id\",\"T.part\"]],\"sourceOnly\":\"add\",\"both\":\"update\",\"output\":\"ledger\"}]}],"
                + "\"ledger\":{\"identity\":[\"T.id\",\"T.part\"],\"search\":{\"columns\":[\"T.id\"]},\"columns\":{\"source\":[\"T.id\",\"T.part\",\"T.amount\",\"U.value\"],\"application\":[{\"name\":\"workState\",\"onSourceChange\":\"reset\"}]}}}";
            string path = Path.Combine(Dir, "settings.json");
            string screen = "{\"workState\":{\"trigger\":\"manual\",\"store\":{\"column\":\"state\"},\"states\":[{\"id\":\"todo\",\"stored\":\"FALSE\"},{\"id\":\"done\",\"stored\":\"TRUE\"},{\"id\":\"hold\",\"stored\":\"HOLD\"}],\"initial\":\"todo\"},\"export\":{\"defaultFields\":[\"T.id\"]},\"candidates\":{\"columns\":[{\"value\":{\"field\":\"T.id\"}}]},\"sections\":[{\"type\":\"titleBar\"}]}";
            File.WriteAllText(path, "{\"schema\":3,\"data\":" + data + ",\"screen\":" + screen + "}", new UTF8Encoding(false));
            Config = Rdv3Config.Load(path);
        }
    }

    private sealed class ApplyFixture
    {
        public readonly string Path = NewPath(".xlsx");
        public readonly string[] Head = { "id", "value" };
        public readonly Rdv3Data Data = new Rdv3Data();
        public readonly Rdv3WorkState Work = new Rdv3WorkState();
        public readonly Rdv3SharedFiles Shared;
        public readonly Rdv3LedgerStore Store;
        public ApplyFixture()
        {
            Data.IdentityCol = 0;
            Data.Columns.Add(new Rdv3ColumnRef { Ref = "T.id", Column = "id" });
            Data.Columns.Add(new Rdv3ColumnRef { Ref = "T.value", Column = "value" });
            Work.Column = "state"; Work.Initial = "initial";
            Work.States.Add(new Rdv3StateDef { Id = "initial", Stored = "0" });
            Work.States.Add(new Rdv3StateDef { Id = "first", Stored = "1" });
            Work.States.Add(new Rdv3StateDef { Id = "second", Stored = "2" });
            Shared = new Rdv3SharedFiles(Path, "test", "user", "apply");
            Store = new Rdv3LedgerStore(Path, Data, Work, Shared);
        }
        public void Save(string[] lines, string[] states)
        { Rdv3Xlsx.Write(Path, Head, Work.Column, lines, states, "fixture", Rdv3Files.StorageContract(Data, Work)); }
        public Rdv3MergeResult Source(string[] lines)
        {
            Rdv3MergeResult source = new Rdv3MergeResult();
            source.Lines = lines; source.Head = Head;
            source.Job = new Rdv3ProcessJobDef { OnSourceChange = "reset",
                ApplyStep = new Rdv3ProcessStepDef { Operation = "merge", SourceOnly = "add", Both = "update", TargetOnly = "keep" } };
            return source;
        }
        public void Trace(string stage, string detail) { }
        public void Warn(string warning) { }
        public Rdv3ApplyOutcome Apply(Rdv3MergeResult source, string[] checkedLines)
        {
            Rdv3LockInfo owner;
            return Store.Apply(source, checkedLines, "apply", delegate { return Shared.TryAcquire(out owner); }, Trace, Warn);
        }
    }

    private static void ConcurrentWriters()
    {
        string path = NewPath(".xlsx");
        string[] head = { "id", "value" }, lines = new string[8], states = new string[8];
        for (int i = 0; i < 8; i++) { lines[i] = i.ToString("D3", CultureInfo.InvariantCulture) + "\tvalue"; states[i] = "0"; }
        Rdv3Xlsx.Write(path, head, "state", lines, states, "init");
        List<Exception> errors = new List<Exception>(); List<Thread> threads = new List<Thread>();
        for (int i = 0; i < 8; i++)
        {
            int row = i;
            Thread thread = new Thread(delegate()
            {
                try
                {
                    Rdv3PendingStore pending = Pending();
                    pending.Set(row.ToString("D3", CultureInfo.InvariantCulture), "1", lines[row], "0");
                    Rdv3SharedFiles shared = new Rdv3SharedFiles(path, "test", "worker", row.ToString(CultureInfo.InvariantCulture));
                    Rdv3LockInfo info; Rdv3LedgerLock lease = null; Stopwatch clock = Stopwatch.StartNew();
                    while (lease == null && clock.ElapsedMilliseconds < 30000)
                    { lease = shared.TryAcquire(out info); if (lease == null) { Thread.Sleep(10); } }
                    if (lease == null) { throw new Exception("lock wait timed out"); }
                    using (lease)
                    {
                        string[] latest, stored; Rdv3Xlsx.Read(path, head, "state", out latest, out stored);
                        Rdv3PendingApply apply = pending.PrepareSend(latest, stored, 0, "0");
                        Check(apply.Unmatched.Count == 0, "unexpected conflict");
                        Rdv3Xlsx.Write(path, head, "state", latest, apply.States, "send");
                        shared.WriteMarker("send", latest.Length, apply.FromInitial, apply.ToInitial);
                        pending.Remove(apply.Resolved);
                    }
                }
                catch (Exception ex) { lock (errors) { errors.Add(ex); } }
            });
            threads.Add(thread); thread.Start();
        }
        for (int i = 0; i < threads.Count; i++) { threads[i].Join(); }
        if (errors.Count > 0) { throw new Exception("parallel writer failed", errors[0]); }
        string[] finalLines, finalStates; Rdv3Xlsx.Read(path, head, "state", out finalLines, out finalStates);
        for (int i = 0; i < 8; i++) { Check(finalStates[i] == "1", "lost update at row " + i); }
    }
}
