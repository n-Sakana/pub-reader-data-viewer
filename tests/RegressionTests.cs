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
            Test("eight-thread-shared-ledger-simulation", ConcurrentWriters);
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
