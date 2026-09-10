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
            ProtectionCases();
            AutomaticCases();
            Test("json-comments-trailing-comma", delegate { Check(Rdv3Json.Parse("{/*x*/\"a\":[1,],}").Member("a").Count == 1, "JSONC"); });
            Test("screen-requires-preserves-saved-record", delegate {
                Rdv3Bind bind = Rdv3Bind.Read(Rdv3Json.Parse("{\"field\":\"A.name\",\"requires\":[\"B.id\",\"C.id\"]}"));
                Rdv3Fields fields = new Rdv3Fields(new string[] { "A.name", "B.id", "C.id" });
                Rdv3View view = new Rdv3View(); view.Record = new string[] { "kept", "b", "c" };
                Check(Rdv3Eval.Evaluate(bind, view, fields, null).Text == "kept", "populated dependencies hid value");
                view.Record[2] = "";
                Check(Rdv3Eval.Evaluate(bind, view, fields, null).Text == "", "missing dependency did not blank display");
                Check(view.Record[0] == "kept", "display condition erased saved data");
                view.Record = null;
                Check(Rdv3Eval.Evaluate(bind, view, fields, null).Text == "", "no selection displayed a value");
            });
            Test("screen-requires-rejects-unknown-fields", delegate {
                Rdv3Bind bind = Rdv3Bind.Read(Rdv3Json.Parse("{\"field\":\"A.name\",\"requires\":[\"B.id\",\"missing\"]}"));
                Rdv3View view = new Rdv3View(); view.Record = new string[] { "kept", "" };
                Check(Rdv3Eval.Evaluate(bind, view, new Rdv3Fields(new string[] { "A.name", "B.id" }), null).Tone == Rdv3Value.Error,
                    "an empty dependency concealed an unknown field");
                Throws<Rdv3LoadError>(delegate { Rdv3Bind.Read(Rdv3Json.Parse("{\"field\":\"A.name\",\"requires\":[\" \" ]}")); });
                Rdv3Config cfg = Rdv3Config.Load(Path.Combine(root, "settings.json"));
                foreach (Rdv3Judgment judgment in cfg.Screen.Judgments.Values) { judgment.Source.Requires = new string[] { "missing" }; }
                Throws<Rdv3LoadError>(delegate { cfg.Screen.Check(cfg.Data); });
            });
            Test("judgment-requires-tests-empty-and-unresolved", delegate {
                Rdv3Judgment judgment = Rdv3Judgment.Read("state", Rdv3Json.Parse("{\"source\":{\"field\":\"B.id\",\"requires\":[\"C.id\"]},\"rules\":[{\"pattern\":\".+\",\"result\":\"yes\"},{\"empty\":true,\"result\":\"none\"}],\"results\":{\"yes\":{\"text\":\"yes\"},\"none\":{\"text\":\"none\"}}}"));
                Rdv3Fields fields = new Rdv3Fields(new string[] { "B.id", "C.id" });
                Rdv3View view = new Rdv3View(); view.Record = new string[] { "b", "c" };
                Check(Rdv3Eval.Judge(judgment, view, fields).Result.Id == "yes", "populated judgment");
                view.Record[1] = "";
                Check(Rdv3Eval.Judge(judgment, view, fields).Result.Id == "none", "missing dependency must reach empty rule");
                judgment.Source.Requires = new string[] { "missing" };
                Check(Rdv3Eval.Judge(judgment, view, fields).Result.Id == "error", "unknown dependency is not an empty result");
            });
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
                Rdv3Table invalid = Table("id,note\n001\t,A\n");
                Check(invalid.Rows == 0 && invalid.InputCounts.InvalidRows == 1
                    && string.Join(" / ", invalid.InputCounts.RowWarnings.ToArray()).Contains("001\\u0009"), "control key was adopted or not identified");
                Throws<Rdv3DataError>(delegate { Table("id,note\n001,A,B\n"); });
            });
            Test("short-csv-counts-values-and-physical-rows", delegate {
                foreach (bool quoted in new bool[] { false, true })
                {
                    Rdv3Table t = Table((quoted ? "\"id\",note\n" : "id,note\n") + "001,A\nshort\n\n002,B\n003,\n");
                    Check(t.Rows == 3 && t.SourceRow(1) == 5 && t.Field(2, 1) == "", "record values or positions shifted");
                    Check(t.InputCounts.ShortRows == 1 && t.InputCounts.BlankRows == 1, "exclusion counts");
                    List<string> warnings = new List<string>(); t.AddWarnings(warnings);
                    string shortRow = Rdv3Text.Format(Rdv3Text.SourceRow, Path.GetFileName(t.Path), 3);
                    string blankRow = Rdv3Text.Format(Rdv3Text.SourceRow, Path.GetFileName(t.Path), 4);
                    Check(warnings.Count == 3 && warnings.Exists(delegate(string warning) { return warning.Contains(shortRow); })
                        && warnings.Exists(delegate(string warning) { return warning.Contains(blankRow); }), "shape summary or excluded physical rows missing");
                }
            });
            Test("unused-header-projection-and-key-ambiguity", delegate {
                HashSet<string> refs = new HashSet<string>(new string[] { "id", "note" }, StringComparer.Ordinal);
                string path = Csv("unused,note,id,unused\nleft,A,001,right\nshort,B,002\nleft,C,003,right\n", Encoding.UTF8);
                Rdv3Table t = Rdv3Table.Read(path, "T", Encoding.UTF8, "id", null, "data.encoding", refs);
                Check(t.Rows == 2 && t.Key(1) == "003" && t.Field(1, 0) == "C", "projection lost field alignment");
                Check(t.InputCounts.HeaderColumns == 2 && t.InputCounts.ShortRows == 1, "source width/counts");
                Check(string.Join(",", Rdv3Table.ReadHead(path, Encoding.UTF8, "data.encoding", refs)) == "note,id", "startup head differs");
                refs.Add("unused");
                Throws<Rdv3DataError>(delegate { Rdv3Table.Read(path, "T", Encoding.UTF8, "id", null, "data.encoding", refs); });
                refs.Remove("unused");
                Throws<Rdv3DataError>(delegate { Rdv3Table.Read(path, "T", Encoding.UTF8, "unused", null, "data.encoding", refs); });
            });
            Test("structural-exclusions-match-ledger-and-process-preview", delegate {
                CompositeFixture f = new CompositeFixture();
                Rdv3MergeResult expected = Rdv3Ledger.BuildFromCsv(f.Config.Data, f.Dir);
                string path = Path.Combine(f.Dir, "T.csv");
                string[] lines = File.ReadAllLines(path, Encoding.GetEncoding(932));
                for (int i = 0; i < lines.Length; i++) { lines[i] += i == 0 ? ",unused,unused" : ",left,right"; }
                File.WriteAllText(path, string.Join("\n", lines) + "\nshort\n\n", Encoding.GetEncoding(932));
                Rdv3MergeResult actual = Rdv3Ledger.BuildFromCsv(f.Config.Data, f.Dir);
                Check(string.Join("\n", actual.Lines) == string.Join("\n", expected.Lines), "ledger values changed");
                Check(actual.Warnings.Count >= 2, "ledger warnings missing");
                object[] args = { f.Config.Data, f.Config.Data.UpdateJob, f.Dir, NewPath(".xlsx"), false };
                string preview = (string)typeof(Rdv3ProcessForm).GetMethod("Build", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static).Invoke(null, args);
                Check((bool)args[4] && preview.Contains("unused"), "preview refused or hid ignored columns");
                Rdv3Json report = Rdv3Json.Parse(Rdv3Headless.Evaluate(f.Config, f.Dir, f.Dir, true, ""));
                Check(report.Member("summary").Member("skippedShort").Num == 1
                    && report.Member("summary").Member("skippedBlank").Num == 1
                    && report.Member("summary").Member("skippedColumns").Num == 2, "headless summary differs");
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
                foreach (string cell in new string[] { "<c><f>1+1</f></c>", "<c t=\"e\"><v>#DIV/0!</v></c>" })
                {
                    Rdv3Table t = Rdv3Table.Read(Fixture(cell, false), "T", Encoding.UTF8, "id");
                    Check(t.Rows == 0 && t.InputCounts.InvalidRows == 1 && t.InputCounts.RowWarnings[0].Contains("A2"), "bad XLSX cell was adopted or not located");
                    Check(t.InputCounts.RowWarnings[0].Contains(cell.Contains("#DIV/0!") ? "#DIV/0!" : "1+1"), "bad XLSX value missing");
                }
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
                Check(conflict.Rows == 0 && conflict.SkippedDuplicateRows == 3 && new Rdv3Index(conflict).Keys == 0, "a conflicting copy was retained");
                string conflictWarning = string.Join(" / ", conflict.InputCounts.RowWarnings.ToArray());
                Check(conflictWarning.Contains("001") && conflictWarning.Contains("2, 3, 4"), "conflicting source group missing");
            });
            Test("input-key-errors-identify-source-and-fix", delegate {
                Rdv3KeyValidation v = new Rdv3KeyValidation(); v.SkipEmpty = false;
                v.SettingsPath = "data.tables.T.keyValidation";
                foreach (string row in new string[] { ",B", "0002,B", "\u65e5,B" })
                {
                    string path = Csv("id,note\n001,A\n" + row + "\n", Encoding.UTF8);
                    Rdv3Table t = Rdv3Table.Read(path, "T", Encoding.UTF8, "id", v);
                    string warning = string.Join(" / ", t.InputCounts.RowWarnings.ToArray());
                    Check(t.Rows == 1 && t.Key(0) == "001" && t.SkippedEmptyRows + t.InputCounts.InvalidRows == 1, "invalid key was retained");
                    Check(warning.Contains(Path.GetFileName(path)) && warning.Contains("3") && warning.Contains("id"), "missing source or column");
                    if (!row.StartsWith(",")) { Check(warning.Contains(v.SettingsPath), "missing key repair setting"); }
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
            Test("operation-log-path-name-and-protection", delegate {
                string ledger = Path.Combine(temp, "Ledger.xlsx");
                string expected = Path.Combine(temp, "Ledger" + Rdv3Text.OpLogInfix + "PC-01.csv");
                Check(Rdv3OperationLog.PathFor(ledger, "PC-01") == expected, "log path");
                Check(Rdv3OperationLog.FileNameFor(ledger, "a/b:c") == "Ledger" + Rdv3Text.OpLogInfix + "a_b_c.csv", "invalid characters not replaced");
                Check(Rdv3OperationLog.IsOperationLog(expected, ledger) && Rdv3OperationLog.IsOperationLog(expected.ToUpperInvariant(), ledger), "own file not recognised");
                Check(!Rdv3OperationLog.IsOperationLog(Path.Combine(temp, "Ledger.csv"), ledger)
                    && !Rdv3OperationLog.IsOperationLog(Path.Combine(temp, "x", Path.GetFileName(expected)), ledger), "unrelated file recognised");
                Check(Rdv3OperationLog.SpoolPathFor(ledger).EndsWith(".spool.csv", StringComparison.Ordinal), "spool path");
                Throws<IOException>(delegate { Rdv3Files.ExportPath(expected, temp, temp, ledger, Path.Combine(temp, "x.log"), Path.Combine(temp, "settings.json"), new Rdv3Data()); });
            });
            Test("operation-log-header-once-quoting-and-spool-flush", delegate {
                string ledger = NewPath(".xlsx");
                Rdv3OperationLog log = new Rdv3OperationLog(ledger, "PC-01", "taro", NewPath(".spool.csv"));
                Check(log.Record(Rdv3Text.OpCreate, 2, "a,b \"q\"\r\nc") == null, "first append failed");
                Check(log.Record(Rdv3Text.OpSend, 2, "plain") == null, "second append failed");
                byte[] bytes = File.ReadAllBytes(log.Path);
                Check(bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "no BOM");
                string[] lines = File.ReadAllText(log.Path, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                Check(lines.Length == 3 && lines[0] == Rdv3Text.OpLogHeader, "header not written exactly once");
                Check(lines[1].EndsWith(",PC-01,taro," + Rdv3Text.OpCreate + ",2,\"a,b \"\"q\"\"  c\"", StringComparison.Ordinal), "line not quoted: " + lines[1]);
                Check(lines[2].EndsWith("," + Rdv3Text.OpSend + ",2,plain", StringComparison.Ordinal), "second line: " + lines[2]);
                using (FileStream held = new FileStream(log.Path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    Check(log.Record(Rdv3Text.OpUpdate, 3, "while locked") != null, "locked file reported as written");
                    Check(log.HasSpool(), "spool missing");
                }
                Check(File.ReadAllText(log.Path, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length == 3, "locked append changed the file");
                Check(log.Flush() == null && !log.HasSpool(), "flush failed");
                lines = File.ReadAllText(log.Path, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                Check(lines.Length == 4 && lines[3].EndsWith(",while locked", StringComparison.Ordinal), "spooled line not delivered in order");
                Check(log.Flush() == null, "empty flush");
            });
            Test("operation-log-keeps-one-file-per-terminal-under-parallel-writers", delegate {
                string ledger = NewPath(".xlsx");
                List<Exception> errors = new List<Exception>(); List<Thread> threads = new List<Thread>();
                for (int i = 0; i < 8; i++)
                {
                    int n = i;
                    Thread thread = new Thread(delegate() {
                        try
                        {
                            Rdv3OperationLog log = new Rdv3OperationLog(ledger, n % 2 == 0 ? "PC-A" : "PC-B", "u" + n, NewPath(".spool.csv"));
                            for (int k = 0; k < 25; k++)
                            {
                                string failure = log.Record(Rdv3Text.OpSend, k, "t" + n + "-" + k);
                                if (failure != null) { throw new Exception(failure); }
                            }
                        }
                        catch (Exception ex) { lock (errors) { errors.Add(ex); } }
                    });
                    threads.Add(thread); thread.Start();
                }
                for (int i = 0; i < threads.Count; i++) { threads[i].Join(); }
                if (errors.Count > 0) { throw new Exception("parallel operation log failed", errors[0]); }
                string[] files = Directory.GetFiles(temp, Path.GetFileNameWithoutExtension(ledger) + Rdv3Text.OpLogInfix + "*.csv");
                Check(files.Length == 2, "expected one file per terminal, found " + files.Length);
                for (int i = 0; i < files.Length; i++)
                {
                    string[] lines = File.ReadAllText(files[i], Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                    Check(lines.Length == 1 + 4 * 25 && lines[0] == Rdv3Text.OpLogHeader, "lost or duplicated lines in " + files[i] + ": " + lines.Length);
                    for (int k = 1; k < lines.Length; k++) { Check(lines[k].Split(',').Length == 6, "interleaved line: " + lines[k]); }
                }
            });
            Test("apply-records-create-and-update-only-when-the-file-was-replaced", delegate {
                ApplyFixture f = new ApplyFixture();
                Rdv3ApplyOutcome created = f.Apply(f.Source(Lines()), null);
                Check(created.Error == null && created.Committed, "creation failed");
                string logPath = f.Shared.Operations.Path;
                Check(Rdv3OperationLog.IsOperationLog(logPath, f.Path), "store log path");
                string[] lines = File.ReadAllText(logPath, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                Check(lines.Length == 2 && lines[1].Contains(",test,user," + Rdv3Text.OpCreate + ",2,"), "creation not recorded: " + string.Join("|", lines));
                Check(!f.Apply(f.Source(Lines()), Lines()).Committed, "unchanged apply committed");
                Check(File.ReadAllText(logPath, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length == 2, "unchanged apply recorded");
                File.WriteAllText(f.Shared.MarkerPath, "unreadable marker");
                Rdv3ApplyOutcome updated = f.Apply(f.Source(new string[] { "001\tNEW", "002\tB" }), Lines());
                Check(updated.Committed && updated.Error != null, "marker failure setup");
                lines = File.ReadAllText(logPath, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                Check(lines.Length == 3 && lines[2].Contains("," + Rdv3Text.OpUpdate + ",2,"), "update before marker failure not recorded: " + string.Join("|", lines));
                Check(lines[2].Contains(" 1 ") && lines[2].Contains(" 0 "), "counts missing from detail: " + lines[2]);
                File.Delete(f.Shared.MarkerPath);
                using (FileStream held = new FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.Read))
                { Check(!f.Apply(f.Source(new string[] { "001\tX", "002\tB" }), new string[] { "001\tNEW", "002\tB" }).Committed, "write failure committed"); }
                Check(File.ReadAllText(logPath, Encoding.UTF8).Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries).Length == 3, "failed write recorded");
            });
            Test("xlsx-date-key-joins-on-fast-path-and-rechecks-key-rules", delegate {
                string dir = NewPath("-datekey"); Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "rows.csv"), "id,name,day\n001,Plain,20260812\n", new UTF8Encoding(false));
                File.Copy(WorkbookFile(new string[] { "day", "label" }, new object[][] { new object[] { 46246, "MATCHED" } }, false), Path.Combine(dir, "dates.xlsx"));
                Rdv3Config cfg = DateJoinConfig(dir, "yyyyMMdd", "");
                Check(cfg.Data.UpdateJob.FastJoinPlan, "fast plan expected for a plain table join");
                Rdv3MergeResult fast = Rdv3Ledger.BuildFromCsv(cfg.Data, dir);
                Check(fast.Matched[0] == 1 && fast.Lines[0] == "001\tPlain\t20260812\t20260812\tMATCHED", "fast path lost the converted date key: " + fast.Lines[0]);
                Check(Rdv3Headless.Evaluate(cfg, dir, dir, true, "").Contains("MATCHED"), "general path lost the join");
                // the converted keys must still obey the table's key rules
                File.Copy(WorkbookFile(new string[] { "day", "label" }, new object[][] { new object[] { 46246, "A" }, new object[] { 46027, "B" } }, false), Path.Combine(dir, "dates.xlsx"), true);
                Rdv3MergeResult ascii = Rdv3Ledger.BuildFromCsv(DateJoinConfig(dir, "yyyy年M月d日", "").Data, dir);
                Check(ascii.Keys[1] == 0 && ascii.Rows == 1 && string.Join(" / ", ascii.Warnings.ToArray()).Contains("2026年8月12日"), "converted non-ASCII key was retained");
                Rdv3MergeResult fixedWidth = Rdv3Ledger.BuildFromCsv(DateJoinConfig(dir, "yyyy/M/d", "").Data, dir);
                Check(fixedWidth.Keys[1] == 1 && fixedWidth.Rows == 1 && string.Join(" / ", fixedWidth.Warnings.ToArray()).Contains("2026/1/5"), "converted variable width was retained");
                Rdv3MergeResult variable = Rdv3Ledger.BuildFromCsv(DateJoinConfig(dir, "yyyy/M/d", ",\"keyValidation\":{\"length\":\"variable\"}").Data, dir);
                Check(variable.Rows == 1 && variable.Matched[0] == 0, "variable-width date keys rejected");
            });
            Test("xlsx-1904-date-system-converts-to-the-same-calendar-day", delegate {
                foreach (bool old in new bool[] { false, true })
                {
                    string dir = NewPath(old ? "-d1904" : "-d1900"); Directory.CreateDirectory(dir);
                    File.Copy(WorkbookFile(new string[] { "id", "name", "day" }, new object[][] { new object[] { "001", "Same day", old ? 44784 : 46246 } }, old), Path.Combine(dir, "rows.xlsx"));
                    Rdv3Config cfg = ConfigOf(dir, SingleTableData("rows.xlsx", "\"A.day\":{\"type\":\"date\",\"format\":\"yyyyMMdd\"}", "\"A.day\":\"Day\",", MergeStep("A"), "\"A.id\",\"A.name\",\"A.day\""));
                    Rdv3MergeResult merge = Rdv3Ledger.BuildFromCsv(cfg.Data, dir);
                    Check(merge.Lines[0] == "001\tSame day\t20260812", (old ? "1904" : "1900") + " system: " + merge.Lines[0]);
                }
            });
            Test("reload-asks-when-content-changed-even-if-the-last-marker-is-a-send", delegate {
                string dir = NewPath("-reload"); Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "rows.csv"), "id,name\n001,Old\n002,Other\n", new UTF8Encoding(false));
                Rdv3Config cfg = ConfigOf(dir, SingleTableData("rows.csv", "", "", MergeStep("A"), "\"A.id\",\"A.name\""));
                string ledger = Path.Combine(dir, "ledger.xlsx");
                Rdv3Ledger.BuildFromCsv(cfg.Data, dir);
                Rdv3Xlsx.Write(ledger, cfg.Data.Head, cfg.Screen.Work.Column, new string[] { "001\tNew", "002\tOther" }, new string[] { "FALSE", "TRUE" }, "test", Rdv3Files.StorageContract(cfg.Data, cfg.Screen.Work), Rdv3LedgerProtection.Create(cfg.Data));
                new Rdv3SharedFiles(ledger, "A", "a", "A", Path.Combine(dir, "sa.csv")).WriteMarker("update", 2, 0, 0);
                new Rdv3SharedFiles(ledger, "B", "b", "B", Path.Combine(dir, "sb.csv")).WriteMarker("send", 2, 1, 0);
                Rdv3SharedFiles reader = new Rdv3SharedFiles(ledger, "C", "c", "C", Path.Combine(dir, "sc.csv"));
                Rdv3SharedMarker marker = reader.ReadMarker();
                Check(marker.Kind == "send", "setup: the last notification must be a send");
                ReaderDataViewer.MainWindow window = new ReaderDataViewer.MainWindow(cfg.Screen);
                try
                {
                    Rdv3Form form = new Rdv3Form(window, cfg.Screen);
                    Rdv3App app = new Rdv3App(form, root, dir, ledger, Path.Combine(dir, "reload.log"), cfg, new Rdv3PendingStore(Path.Combine(dir, "pending.dat")), reader, null);
                    System.Reflection.BindingFlags flags = System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic;
                    typeof(Rdv3App).GetField("ledHead", flags).SetValue(app, cfg.Data.Head);
                    typeof(Rdv3App).GetField("ledLines", flags).SetValue(app, new string[] { "001\tOld", "002\tOther" });
                    typeof(Rdv3App).GetField("ledStates", flags).SetValue(app, new string[] { "TRUE", "FALSE" });
                    typeof(Rdv3App).GetField("sharedStates", flags).SetValue(app, new string[] { "TRUE", "FALSE" });
                    typeof(Rdv3App).GetMethod("ReloadSharedJob", flags).Invoke(app, new object[] { "test-reload", marker, "", "" });
                    string[] active = (string[])typeof(Rdv3App).GetField("ledLines", flags).GetValue(app);
                    string log = File.ReadAllText(Path.Combine(dir, "reload.log"), Encoding.UTF8);
                    Check(log.Contains("changed=true") && log.Contains("reset=1"), "content change not detected behind a send notification: " + log);
                    Check(active[0] == "001\tOld" && log.Contains("switch declined"), "content change adopted without confirmation: " + log);
                }
                finally { window.Close(); }
            });
            Test("send-detail-and-notice-name-the-actual-states", delegate {
                string ken = "件", sep = "、";
                Rdv3WorkState two = new Rdv3WorkState();
                two.Column = "state"; two.Initial = "todo";
                two.States.Add(new Rdv3StateDef { Id = "todo", Stored = "FALSE", Text = "Todo" });
                two.States.Add(new Rdv3StateDef { Id = "done", Stored = "TRUE", Text = "Checked" });
                two.Transitions.Add(new Rdv3Transition { From = "todo", To = "done" });
                two.Transitions.Add(new Rdv3Transition { From = "done", To = "todo" });
                Check(Rdv3OperationLog.SendDetail(two, new string[] { "FALSE", "TRUE", "FALSE" }, new string[] { "TRUE", "FALSE", "FALSE" }) == "Checked 1 " + ken + sep + "Todo 1 " + ken, "two-state detail");
                Rdv3WorkState three = new Rdv3WorkState();
                three.Column = "state"; three.Initial = "todo";
                three.States.Add(new Rdv3StateDef { Id = "todo", Stored = "FALSE", Text = "Todo" });
                three.States.Add(new Rdv3StateDef { Id = "done", Stored = "TRUE", Text = "Checked" });
                three.States.Add(new Rdv3StateDef { Id = "approved", Stored = "APPROVED", Text = "Approved" });
                three.Transitions.Add(new Rdv3Transition { From = "todo", To = "done" });
                three.Transitions.Add(new Rdv3Transition { From = "done", To = "approved" });
                three.Transitions.Add(new Rdv3Transition { From = "approved", To = "todo" });
                Check(Rdv3OperationLog.SendDetail(three, new string[] { "TRUE" }, new string[] { "APPROVED" }) == "Checked 0 " + ken + sep + "Approved 1 " + ken + sep + "Todo 0 " + ken, "three-state detail");
                Check(Rdv3App.SendNoticeText(two, "taro", 1, 0).Contains("Checked"), "two-state notice names the only destination");
                string notice = Rdv3App.SendNoticeText(three, "taro", 1, 0);
                Check(!notice.Contains("Checked") && !notice.Contains("Approved") && notice.Contains("Todo") && notice.Contains("taro"), "three-state notice guessed a state: " + notice);
            });
            Test("operation-log-never-resends-a-delivered-spool", delegate {
                string ledger = NewPath(".xlsx");
                Rdv3OperationLog log = new Rdv3OperationLog(ledger, "PC-01", "taro", NewPath(".spool.csv"));
                using (FileStream held = new FileStream(log.Path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None))
                { Check(log.Record(Rdv3Text.OpSend, 1, "one").StartsWith("spooled", StringComparison.Ordinal), "expected a spooled line"); }
                // a reader that shares nothing else: the spool can neither be deleted nor truncated
                using (FileStream reader = new FileStream(log.SpoolPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    string first = log.Flush();
                    string second = log.Flush();
                    Check(first != null && first.StartsWith("written", StringComparison.Ordinal), "first flush outcome: " + first);
                    Check(File.ReadAllLines(log.Path).Length == 2, "delivered line appended twice: " + first + " / " + second);
                    Check(!log.HasSpool(), "delivered spool still counted as pending");
                }
                Rdv3OperationLog again = new Rdv3OperationLog(ledger, "PC-01", "taro", log.SpoolPath);
                Check(!again.HasSpool() && again.Flush() == null && File.ReadAllLines(log.Path).Length == 2, "a new process resent the delivered spool");
                Check(!File.Exists(log.SpoolPath) || new FileInfo(log.SpoolPath).Length == 0, "spool not settled after release");
                Check(log.Record(Rdv3Text.OpSend, 1, "two") == null && File.ReadAllLines(log.Path).Length == 3, "later line not appended once");
            });
            Test("derived-typed-column-values-are-checked-before-the-ledger-write", delegate {
                string dir = NewPath("-derived"); Directory.CreateDirectory(dir);
                File.WriteAllText(Path.Combine(dir, "rows.csv"), "id,name\n001,Example\n", new UTF8Encoding(false));
                foreach (string expression in new string[] { "'oops'", "'12'" })
                {
                    string steps = "{\"operation\":\"calculate\",\"target1\":\"A\",\"column\":\"amount\",\"expression\":\"" + expression + "\",\"output\":\"D\"}," + MergeStep("D");
                    Rdv3Config cfg = ConfigOf(dir, SingleTableData("rows.csv", "\"D.amount\":{\"type\":\"number\"}", "\"D\":\"Calc\",\"D.amount\":\"Amount\",", steps, "\"A.id\",\"A.name\",\"D.amount\""));
                    if (expression == "'oops'")
                    {
                        Rdv3MergeResult excluded = Rdv3Ledger.BuildFromCsv(cfg.Data, dir);
                        string warning = string.Join(" / ", excluded.Warnings.ToArray());
                        Check(excluded.Rows == 0 && warning.Contains("D.amount") && warning.Contains("oops") && warning.Contains("001"), "typed result was retained or unidentified: " + warning);
                    }
                    else { Check(Rdv3Ledger.BuildFromCsv(cfg.Data, dir).Lines[0] == "001\tExample\t12", "valid number rejected"); }
                }
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
            ValidationFeedback();
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

    private static Rdv3ValidationError ValidationFailure(Action action)
    {
        try { action(); }
        catch (Rdv3ValidationError error) { return error; }
        throw new Exception("Expected collected validation errors");
    }

    private static void ValidationFeedback()
    {
        Test("validation-collects-duplicate-label-and-all-missing-references", delegate {
            CompositeFixture f = new CompositeFixture();
            string text = File.ReadAllText(f.Config.SourcePath, Encoding.UTF8)
                .Replace("\"labels\":{", "\"labels\":{\"T\":\"duplicate\",")
                .Replace("\"joined\":\"Joined\",", "").Replace("\"U.part\":\"Part\",", "");
            Rdv3ValidationError e = ValidationFailure(delegate { Rdv3Config.Parse(text, true); });
            Check(e.Errors.Length == 4 && e.Stage == "settings", "expected duplicate plus three reference locations");
            string errors = string.Join("\n", e.Errors);
            Check(errors.Contains("T already has a table label") && errors.Contains("joined has no screen label under data.labels or tables.*.label")
                && errors.Contains("steps[1].target1") && errors.Contains("U.part"), "lost reason or location");
            Throws<Rdv3LoadError>(delegate { Rdv3Config.Parse(text); });
        });
        Test("validation-independent-settings-fields", delegate {
            CompositeFixture f = new CompositeFixture();
            string text = File.ReadAllText(f.Config.SourcePath, Encoding.UTF8)
                .Replace("\"schema\":3,", "\"schema\":3,\"watch\":{\"pollMs\":0,\"stableMs\":-1},\"search\":{\"pattern\":\"[\",\"candidateRowsShown\":0},")
                .Replace("\"screen\":{", "\"screen\":{\"card\":{\"width\":1,\"fontSize\":1},");
            Rdv3ValidationError e = ValidationFailure(delegate { Rdv3Config.Parse(text, true); });
            Check(e.Errors.Length == 6, "independent fields stopped each other");
            Check(string.Join("\n", e.Unchecked).Contains("input files"), "input stage not explicitly skipped");
        });
        Test("validation-independent-table-definitions", delegate {
            CompositeFixture f = new CompositeFixture();
            string text = File.ReadAllText(f.Config.SourcePath, Encoding.UTF8)
                .Replace("\"file\":\"T.csv\"", "\"file\":42").Replace("\"file\":\"U.csv\"", "\"file\":false");
            Rdv3ValidationError e = ValidationFailure(delegate { Rdv3Config.Parse(text, true); });
            Check(e.Errors.Length == 2 && string.Join("\n", e.Errors).Contains("tables.U.file"), "second table not checked");
            Check(string.Join("\n", e.Unchecked).Contains("incomplete table definitions"), "dependent stage not identified");
        });
        Test("validation-independent-type-definitions", delegate {
            CompositeFixture f = new CompositeFixture();
            string text = File.ReadAllText(f.Config.SourcePath, Encoding.UTF8).Replace("\"jobs\":[", "\"types\":{\"T.amount\":{\"type\":\"bad\"},\"U.value\":{\"type\":\"bad\"}},\"jobs\":[");
            Check(ValidationFailure(delegate { Rdv3Config.Parse(text, true); }).Errors.Length == 2, "type definitions stopped each other");
        });
        Test("validation-independent-input-files", delegate {
            CompositeFixture f = new CompositeFixture();
            f.Config.Data.Tables[0].File = "missing-T.csv"; f.Config.Data.Tables[1].File = "missing-U.csv";
            Rdv3ValidationError e = ValidationFailure(delegate { Rdv3Headless.Evaluate(f.Config, f.Dir, f.Dir, false, ""); });
            Check(e.Errors.Length == 2 && e.Stage == "input files", "missing files not collected");
            Check(string.Join("\n", e.Unchecked).Contains("input columns/types"), "claimed dependent validation");
        });
        Test("validation-reports-each-invalid-typed-cell", delegate {
            CompositeFixture f = new CompositeFixture();
            string text = File.ReadAllText(f.Config.SourcePath, Encoding.UTF8).Replace("\"jobs\":[", "\"types\":{\"T.amount\":{\"type\":\"number\"},\"U.value\":{\"type\":\"number\"}},\"jobs\":[");
            File.WriteAllText(Path.Combine(f.Dir, "T.csv"), "id,part,amount\nAB,C,bad1\nA,BC,bad2\n", Encoding.GetEncoding(932));
            File.WriteAllText(f.Config.SourcePath, text, new UTF8Encoding(false));
            Rdv3Config c = Rdv3Config.Load(f.Config.SourcePath, true);
            Rdv3Json report = Rdv3Json.Parse(Rdv3Headless.Evaluate(c, f.Dir, f.Dir, false, ""));
            Check(report.Member("summary").Member("skippedInvalid").Num == 5, "invalid rows were hidden");
            string warnings = "";
            foreach (Rdv3Json warning in report.Member("warnings").Items) { warnings += warning.Str + "\n"; }
            Check(warnings.Contains("bad1") && warnings.Contains("bad2") && warnings.Contains("first")
                && warnings.Contains("second") && warnings.Contains("extra"), "lost actual values");
        });
        Test("validation-syntax-failure-count-and-exit", delegate {
            string path = NewPath(".json"); File.WriteAllText(path, "{\"schema\":", Encoding.UTF8);
            TextWriter previous = Console.Error; StringWriter captured = new StringWriter(); int code;
            try { Console.SetError(captured); code = Rdv3Headless.Run(temp, path, temp, false, "", ""); }
            finally { Console.SetError(previous); }
            Check(code == 3 && captured.ToString().StartsWith("FAIL 1 error"), "failure count or exit changed");
            Check(captured.ToString().Contains("subsequent checks were not performed"), "syntax failure implies complete validation");
        });
        Test("validation-valid-config-and-results-remain-equivalent", delegate {
            CompositeFixture f = new CompositeFixture();
            Rdv3Config c = Rdv3Config.Load(f.Config.SourcePath, true);
            Check(c.Data.Describe() == f.Config.Data.Describe() && c.Screen.Describe() == f.Config.Screen.Describe(), "collection changed valid definitions");
            Check(Rdv3Headless.Evaluate(c, f.Dir, f.Dir, false, "") == Rdv3Headless.Evaluate(f.Config, f.Dir, f.Dir, false, ""), "validation results differ");
            Check(!File.Exists(Path.Combine(f.Dir, c.Ledger)), "validation created a ledger");
        });
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
            Rdv3Table excluded = Rdv3Table.Read(conflict, "T", Encoding.UTF8, new string[] { "id", "part" }, null);
            string warning = string.Join(" / ", excluded.InputCounts.RowWarnings.ToArray());
            Check(excluded.Rows == 0 && excluded.SkippedDuplicateRows == 2 && new Rdv3Index(excluded).Keys == 0, "conflicting tuple retained");
            Check(warning.Contains("id / part") && warning.Contains("2, 3"), "tuple conflict location");
        });
        Test("composite-key-rules-apply-to-each-component", delegate {
            Rdv3KeyValidation rule = new Rdv3KeyValidation();
            string skipped = Csv("id,part\nTOO-LONG,\nA,01\nB,02\n", Encoding.UTF8);
            Check(Rdv3Table.Read(skipped, "T", Encoding.UTF8, new string[] { "id", "part" }, rule).Rows == 2, "skipped row set fixed width");
            string bad = Csv("id,part\nA,01\nB,002\n", Encoding.UTF8);
            Rdv3Table width = Rdv3Table.Read(bad, "T", Encoding.UTF8, new string[] { "id", "part" }, rule);
            Check(width.Rows == 1 && width.InputCounts.InvalidRows == 1 && width.InputCounts.RowWarnings[0].Contains("part")
                && width.InputCounts.RowWarnings[0].Contains("variable"), "component width repair");
            Rdv3Table unicode = Rdv3Table.Read(Csv("id,part\nA,\u3042\n", Encoding.UTF8), "T", Encoding.UTF8, new string[] { "id", "part" }, rule);
            Check(unicode.Rows == 0 && unicode.InputCounts.InvalidRows == 1 && unicode.InputCounts.RowWarnings[0].Contains("part"), "non-ASCII component retained");
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
            Rdv3Ledger.BuildFromCsv(f.Config.Data, f.Dir);
            Rdv3ProcessResult first = Rdv3Process.Run(f.Config.Data, f.Config.Data.UpdateJob, f.Dir, new string[0], new string[0], "FALSE");
            string path = NewPath(".xlsx"); string[] states = { "TRUE", "HOLD", "FALSE" };
            Rdv3Xlsx.Write(path, f.Config.Data.Head, f.Config.Screen.Work.Column, first.Lines, states, "baseline", Rdv3Files.StorageContract(f.Config.Data, f.Config.Screen.Work), Rdv3LedgerProtection.Create(f.Config.Data));
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

    // A small workbook: header row plus data rows; an int cell is written as a
    // numeric <v> (a serial date), anything else as an inline string.
    private static string WorkbookFile(string[] head, object[][] rows, bool date1904)
    {
        string path = NewPath(".xlsx");
        string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        StringBuilder sheet = new StringBuilder();
        sheet.Append("<worksheet xmlns=\"" + ns + "\"><sheetData>");
        for (int r = -1; r < rows.Length; r++)
        {
            object[] cells = (r < 0) ? head : rows[r];
            string number = (r + 2).ToString(CultureInfo.InvariantCulture);
            sheet.Append("<row r=\"" + number + "\">");
            for (int c = 0; c < cells.Length; c++)
            {
                string address = ((char)('A' + c)).ToString() + number;
                if (cells[c] is int) { sheet.Append("<c r=\"" + address + "\"><v>" + ((int)cells[c]).ToString(CultureInfo.InvariantCulture) + "</v></c>"); }
                else { sheet.Append("<c r=\"" + address + "\" t=\"inlineStr\"><is><t>" + cells[c] + "</t></is></c>"); }
            }
            sheet.Append("</row>");
        }
        sheet.Append("</sheetData></worksheet>");
        using (FileStream stream = File.Create(path))
        using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            Entry(zip, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/><Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>");
            Entry(zip, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + rel + "/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
            Entry(zip, "xl/workbook.xml", "<workbook xmlns=\"" + ns + "\" xmlns:r=\"" + rel + "\"><workbookPr date1904=\"" + (date1904 ? "1" : "0") + "\"/><sheets><sheet name=\"Input\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            Entry(zip, "xl/_rels/workbook.xml.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"" + rel + "/worksheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>");
            Entry(zip, "xl/worksheets/sheet1.xml", sheet.ToString());
        }
        return path;
    }

    // A configuration whose data section is given; the screen is the two-state minimum.
    private static Rdv3Config ConfigOf(string dir, string data)
    {
        string screen = "{\"workState\":{\"trigger\":\"manual\",\"store\":{\"column\":\"state\"},\"states\":[{\"id\":\"todo\",\"stored\":\"FALSE\",\"text\":\"Todo\"},{\"id\":\"done\",\"stored\":\"TRUE\",\"text\":\"Done\"}],\"initial\":\"todo\"},\"export\":{\"defaultFields\":[\"A.id\"]},\"candidates\":{\"columns\":[{\"value\":{\"field\":\"A.id\"}}]},\"sections\":[{\"type\":\"titleBar\"}]}";
        string path = Path.Combine(dir, "settings.json");
        File.WriteAllText(path, "{\"schema\":3,\"data\":" + data + ",\"screen\":" + screen + "}", new UTF8Encoding(false));
        return Rdv3Config.Load(path);
    }

    private static string SingleTableData(string file, string types, string extraLabels, string steps, string source)
    {
        return "{\"tables\":{\"A\":{\"file\":\"" + file + "\",\"key\":\"id\"}},\"types\":{" + types + "},"
            + "\"labels\":{\"A.id\":\"ID\",\"A.name\":\"Name\"," + extraLabels + "\"ledger\":\"Ledger\"},"
            + "\"jobs\":[{\"id\":\"update\",\"kind\":\"update\",\"inputs\":[{\"table\":\"A\"}],\"steps\":[" + steps + "]}],"
            + "\"ledger\":{\"identity\":\"A.id\",\"search\":{\"columns\":[\"A.id\"]},\"columns\":{\"source\":[" + source + "],\"application\":[{\"name\":\"workState\",\"onSourceChange\":\"reset\"}]}}}";
    }

    private static string MergeStep(string target1)
    {
        return "{\"operation\":\"merge\",\"target1\":\"" + target1 + "\",\"target2\":\"ledger\",\"keys\":[[\"A.id\"],[\"A.id\"]],\"sourceOnly\":\"add\",\"both\":\"update\",\"output\":\"ledger\"}";
    }

    private static Rdv3Config DateJoinConfig(string dir, string format, string keyRule)
    {
        string data = "{\"tables\":{\"A\":{\"file\":\"rows.csv\",\"key\":\"id\"},\"B\":{\"file\":\"dates.xlsx\",\"key\":\"day\"" + keyRule + "}},"
            + "\"types\":{\"B.day\":{\"type\":\"date\",\"format\":\"" + format + "\"}},"
            + "\"labels\":{\"A.id\":\"ID\",\"A.name\":\"Name\",\"A.day\":\"Left\",\"B.day\":\"Right\",\"B.label\":\"Label\",\"J\":\"Joined\",\"ledger\":\"Ledger\"},"
            + "\"jobs\":[{\"id\":\"update\",\"kind\":\"update\",\"inputs\":[{\"table\":\"A\"},{\"table\":\"B\"}],\"steps\":["
            + "{\"operation\":\"join\",\"target1\":\"A\",\"target2\":\"B\",\"keys\":[[\"A.day\"],[\"B.day\"]],\"condition\":\"left\",\"output\":\"J\"},"
            + MergeStep("J") + "]}],"
            + "\"ledger\":{\"identity\":\"A.id\",\"search\":{\"columns\":[\"A.id\"]},\"columns\":{\"source\":[\"A.id\",\"A.name\",\"A.day\",\"B.day\",\"B.label\"],\"application\":[{\"name\":\"workState\",\"onSourceChange\":\"reset\"}]}}}";
        return ConfigOf(dir, data);
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

    private static void AutomaticCases()
    {
        Test("automatic-condition-uses-result-id-and-requires-a-record", delegate {
            Rdv3Config config = Rdv3Config.Load(Path.Combine(root, "settings.json"));
            Rdv3WorkState work = config.Screen.Work;
            work.Trigger = "automatic"; work.AutomaticJudgment = "settlementCheck"; work.AutomaticResult = "eligible";
            Rdv3Judgment judgment = new Rdv3Judgment { Id = "settlementCheck", Source = new Rdv3Bind { Fields = new string[] { "arbitrary.flag" } } };
            judgment.Results.Add("eligible", new Rdv3Result { Id = "eligible", Text = "same display" });
            judgment.Results.Add("hold", new Rdv3Result { Id = "hold", Text = "same display" });
            judgment.Rules.Add(new Rdv3Rule { EqualsAny = new string[] { "yes" }, Result = "eligible" });
            judgment.Rules.Add(new Rdv3Rule { Empty = true, Result = "hold" });
            config.Screen.Judgments.Add(judgment.Id, judgment);
            Rdv3Fields fields = new Rdv3Fields(new string[] { "arbitrary.flag" });
            Rdv3View view = new Rdv3View();
            Check(!work.AllowsAutomatic(config.Screen, view, fields), "no selection was completed");
            view.Record = new string[] { "" };
            Check(!work.AllowsAutomatic(config.Screen, view, fields), "same display text allowed an ineligible result");
            view.Record[0] = "other";
            Check(!work.AllowsAutomatic(config.Screen, view, fields), "undefined result was completed");
            view.Record[0] = "yes";
            Check(!work.AllowsAutomatic(config.Screen, view, Rdv3Fields.Empty), "unresolved field was completed");
            Check(work.AllowsAutomatic(config.Screen, view, fields), "named eligible result was not completed");
            work.Trigger = "manual";
            Check(!work.AllowsAutomatic(config.Screen, view, fields), "manual setting completed automatically");
        });
        Test("automatic-condition-parsing-and-invalid-references", delegate {
            Rdv3Config config = Rdv3Config.Load(Path.Combine(root, "settings.json"));
            string reference = config.Data.ColumnRefs[0];
            Rdv3Judgment judgment = new Rdv3Judgment { Id = "generic", Source = new Rdv3Bind { Fields = new string[] { reference } } };
            judgment.Results.Add("allow", new Rdv3Result { Id = "allow" });
            config.Screen.Judgments.Add(judgment.Id, judgment);
            string definition = "{\"store\":{\"column\":\"state\"},\"states\":[{\"id\":\"new\"}],\"initial\":\"new\",\"trigger\":\"automatic\",\"automaticWhen\":{\"judgment\":\"generic\",\"result\":\"allow\"}}";
            Rdv3WorkState work = Rdv3WorkState.Read(Rdv3Json.Parse(definition));
            work.CheckAutomatic(config.Screen);
            Check(work.AutomaticJudgment == "generic" && work.AutomaticResult == "allow", "condition not parsed");
            foreach (string result in new string[] { "missing", "undefined", "error" })
            { work.AutomaticResult = result; Throws<Rdv3LoadError>(delegate { work.CheckAutomatic(config.Screen); }); }
            work.AutomaticResult = "allow"; work.AutomaticJudgment = "missing";
            Throws<Rdv3LoadError>(delegate { work.CheckAutomatic(config.Screen); });
            config.Screen.Work.AutomaticJudgment = "missing"; config.Screen.Work.AutomaticResult = "allow";
            Throws<Rdv3LoadError>(delegate { config.Screen.Check(config.Data); });
            Throws<Rdv3LoadError>(delegate { Rdv3WorkState.Read(Rdv3Json.Parse(definition.Replace("\"result\":\"allow\"", "\"typo\":\"allow\""))); });
        });
        Test("automatic-without-condition-keeps-existing-behavior", delegate {
            Rdv3Config config = Rdv3Config.Load(Path.Combine(root, "settings.json"));
            config.Screen.Work.Trigger = "automatic";
            Check(config.Screen.Work.AllowsAutomatic(config.Screen, new Rdv3View { Record = new string[] { "" } }, Rdv3Fields.Empty), "legacy automatic mode changed");
        });
    }

    private static void ProtectionCases()
    {
        Test("protection-preserves-saved-done-content-and-updates-blanks", delegate {
            ApplyFixture f = new ApplyFixture(); f.Data.ProtectedStates = new string[] { "1" };
            f.Save(Lines(), new string[] { "1", "0" });
            Rdv3ApplyOutcome result = f.Apply(f.Source(new string[] { "001\tNEW", "002\t", "003\tC" }), Lines());
            Check(result.Error == null && result.Committed, "protected update failed");
            Rdv3LedgerSnapshot saved = f.Store.Read(f.Head);
            Check(saved.Lines.Length == 3 && saved.Lines[0] == "001\tA" && saved.States[0] == "1", "sent row overwritten");
            Check(saved.Lines[1] == "002\t" && saved.States[1] == "0" && saved.Lines[2] == "003\tC", "blank/new values lost");
            Check(result.Update.Protected == 1 && result.Update.Updated == 1 && result.Update.Added == 1, "counts wrong");
        });
        Test("protection-rechecks-state-at-lease-and-allows-sent-reset", delegate {
            ApplyFixture f = new ApplyFixture(); f.Data.ProtectedStates = new string[] { "1" }; f.Save(Lines(), States());
            Rdv3LockInfo owner;
            Rdv3ApplyOutcome result = f.Store.Apply(f.Source(new string[] { "001\tNEW", "002\tB" }), Lines(), "race",
                delegate { Rdv3LedgerLock lease = f.Shared.TryAcquire(out owner); f.Save(Lines(), new string[] { "1", "0" }); return lease; }, f.Trace, f.Warn);
            Check(result.Error == null && f.Store.Read(f.Head).Lines[0] == "001\tA", "state sent before lease read ignored");
            f.Save(Lines(), States());
            result = f.Apply(f.Source(new string[] { "001\tNEW", "002\tB" }), Lines());
            Check(result.Error == null && f.Store.Read(f.Head).Lines[0] == "001\tNEW", "sent reset stayed permanently locked");
        });
        Test("protection-does-not-use-pending-overlay", delegate {
            ApplyFixture f = new ApplyFixture(); f.Data.ProtectedStates = new string[] { "1" }; f.Save(Lines(), States());
            Rdv3PendingStore pending = Pending(); pending.Set("001", "1", Lines()[0], "0");
            Check(pending.Overlay(Lines(), States(), f.Data.IdentityCols)[0] == "1", "fixture pending");
            Rdv3ApplyOutcome result = f.Apply(f.Source(new string[] { "001\tNEW" }), Lines());
            Rdv3LedgerSnapshot saved = f.Store.Read(f.Head);
            Check(result.Error == null && saved.Lines[0] == "001\tNEW" && saved.Lines[1] == "002\tB", "pending protected or absent row lost");
            Check(pending.PrepareSend(saved.Lines, saved.States, f.Data.IdentityCols, "0").Unmatched.Count == 1, "pending content conflict lost");
        });
        Test("archive-roundtrip-reimport-and-explicit-full-restore", delegate {
            ApplyFixture f = new ApplyFixture(); f.Data.ProtectedStates = new string[] { "1" };
            string[] original = { "001\tA & <tag> \r\n", "002\tB" }; f.Save(original, new string[] { "1", "0" });
            Rdv3LedgerSnapshot snapshot = f.Store.Read(f.Head);
            snapshot.Protection.ArchiveRemoved(f.Data, snapshot.Lines, snapshot.States, new string[] { original[1] });
            f.Store.Write(f.Head, new string[] { original[1] }, new string[] { "0" }, "delete", snapshot.Protection);
            Rdv3ApplyOutcome result = f.Apply(f.Source(new string[] { "001\tREPLACEMENT", "002\tUPDATED" }), new string[] { original[1] });
            Rdv3LedgerSnapshot saved = f.Store.Read(f.Head);
            Check(result.Error == null && saved.Lines.Length == 1 && saved.Protection.Deleted.Count == 1, "deleted row resurrected");
            Check(saved.Protection.Deleted[0].Line == original[0] && saved.Protection.Deleted[0].State == "1", "archive changed");
            string[] lines = saved.Lines, states = saved.States;
            saved.Protection.Restore(f.Data, new Rdv3DeletedRecord[] { saved.Protection.Deleted[0] }, ref lines, ref states);
            f.Store.Write(f.Head, lines, states, "restore", saved.Protection);
            saved = f.Store.Read(f.Head);
            Check(saved.Lines.Length == 2 && saved.Lines[1] == original[0] && saved.States[1] == "1" && saved.Protection.Deleted.Count == 0, "full restore failed");
        });
        Test("archive-rejects-active-and-stale-restore-without-mutation", delegate {
            ApplyFixture f = new ApplyFixture(); Rdv3LedgerProtection p = Rdv3LedgerProtection.Create(f.Data);
            p.ArchiveRemoved(f.Data, Lines(), States(), new string[] { Lines()[1] });
            string[] lines = Lines(), states = States();
            Throws<InvalidDataException>(delegate { p.Restore(f.Data, new Rdv3DeletedRecord[] { p.Deleted[0] }, ref lines, ref states); });
            Check(lines.Length == 2 && p.Deleted.Count == 1, "conflict mutated archive");
            lines = new string[] { Lines()[1] }; states = new string[] { "0" };
            Rdv3DeletedRecord stale = new Rdv3DeletedRecord { Line = "001\tSTALE", State = "0", DeletedAt = p.Deleted[0].DeletedAt };
            Throws<InvalidDataException>(delegate { p.Restore(f.Data, new Rdv3DeletedRecord[] { stale }, ref lines, ref states); });
            Check(lines.Length == 1 && p.Deleted.Count == 1, "stale restore mutated archive");
        });
        Test("archive-write-failure-retains-live-and-archive-together", delegate {
            ApplyFixture f = new ApplyFixture(); f.Save(Lines(), States()); byte[] before = File.ReadAllBytes(f.Path);
            Rdv3LedgerSnapshot s = f.Store.Read(f.Head); s.Protection.ArchiveRemoved(f.Data, s.Lines, s.States, new string[] { s.Lines[1] });
            using (FileStream held = new FileStream(f.Path, FileMode.Open, FileAccess.Read, FileShare.Read))
            { Throws<IOException>(delegate { f.Store.Write(f.Head, new string[] { s.Lines[1] }, new string[] { "0" }, "blocked", s.Protection); }); }
            Check(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(f.Path)), "half deletion committed");
            Check(f.Store.Read(f.Head).Protection.Deleted.Count == 0, "archive committed despite failure");
        });
        Test("definition-rejects-changes-and-missing-metadata", delegate {
            ApplyFixture f = new ApplyFixture(); f.Save(Lines(), States()); byte[] before = File.ReadAllBytes(f.Path);
            f.Data.Definition = "{\"changed\":true}";
            Check(f.Apply(f.Source(Lines()), Lines()).Error is InvalidDataException, "changed definition accepted");
            Check(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(f.Path)), "definition failure changed file");
            f.Data.Definition = "{}";
            using (FileStream file = new FileStream(f.Path, FileMode.Open, FileAccess.ReadWrite))
            using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Update)) { zip.GetEntry("rdv-protection.xml").Delete(); }
            Throws<InvalidDataException>(delegate { f.Store.Read(f.Head); });
        });
        Test("definition-normalizes-json-and-permits-presentation-location", delegate {
            string text = File.ReadAllText(Path.Combine(root, "settings.json"), Encoding.UTF8);
            Rdv3Config a = Rdv3Config.Parse(text);
            Rdv3Config b = Rdv3Config.Parse(text.Replace("\"file\": \"tableA.csv\"", "\"file\": \"elsewhere/tableA.csv\""));
            Check(a.Data.Definition == b.Data.Definition, "location changed definition");
            string x = "{\"tables\":{\"T\":{\"label\":\"A\",\"file\":\"one.csv\",\"key\":\"id\"}},\"labels\":{\"a\":\"A\"}}";
            string y = "{ /* comment */ \"labels\":{},\"tables\":{\"T\":{\"key\":[\"id\"],\"file\":\"sub/two.CSV\",\"label\":\"B\"}}}";
            Check(Rdv3BusinessDefinition.Normalize(Rdv3Json.Parse(x), true) == Rdv3BusinessDefinition.Normalize(Rdv3Json.Parse(y), true), "JSON representation rejected");
            Check(Rdv3BusinessDefinition.Normalize(Rdv3Json.Parse(y.Replace(".CSV", ".xlsx")), true) != Rdv3BusinessDefinition.Normalize(Rdv3Json.Parse(x), true), "input format change accepted");
        });
        Test("legacy-ledger-readable-but-no-writes-without-migration", delegate {
            ApplyFixture f = new ApplyFixture();
            Rdv3Xlsx.Write(f.Path, f.Head, f.Work.Column, Lines(), States(), "legacy", Rdv3Files.LegacyStorageContract(f.Data, f.Work));
            Check(f.Store.Read(f.Head).Protection.Legacy, "legacy not recognized");
            byte[] before = File.ReadAllBytes(f.Path);
            Rdv3ApplyOutcome result = f.Apply(f.Source(new string[] { "001\tNEW" }), Lines());
            Check(result.Error is InvalidDataException && !result.Committed, "legacy overwritten");
            Check(Convert.ToBase64String(before) == Convert.ToBase64String(File.ReadAllBytes(f.Path)), "legacy file changed");
        });
        Test("migration-preserves-all-content-and-states-in-new-file", delegate {
            CompositeFixture f = new CompositeFixture(); Rdv3MergeResult first = Rdv3Ledger.BuildFromCsv(f.Config.Data, f.Dir);
            string oldPath = NewPath(".xlsx"), newPath = NewPath(".xlsx"); string[] states = { "TRUE", "HOLD", "FALSE" };
            Rdv3Xlsx.Write(oldPath, first.Head, f.Config.Screen.Work.Column, first.Lines, states, "legacy", Rdv3Files.LegacyStorageContract(f.Config.Data, f.Config.Screen.Work));
            byte[] oldBytes = File.ReadAllBytes(oldPath);
            Rdv3Config current = Rdv3Config.Load(f.Config.SourcePath);
            Rdv3Migration.Migrate(f.Config, current, f.Dir, f.Dir, oldPath, newPath);
            Rdv3LedgerSnapshot saved = new Rdv3LedgerStore(newPath, current.Data, current.Screen.Work, null).Read(first.Head);
            Check(!saved.Protection.Legacy && Rdv3Ledger.SameLedger(first.Lines, states, saved.Lines, saved.States), "migration changed rows or states");
            Check(Convert.ToBase64String(oldBytes) == Convert.ToBase64String(File.ReadAllBytes(oldPath)), "migration touched old file");
            Throws<IOException>(delegate { Rdv3Migration.Migrate(f.Config, current, f.Dir, f.Dir, oldPath, newPath); });
            string badPath = NewPath(".xlsx"); current.Data.LegacyDefinition += "changed";
            Throws<InvalidDataException>(delegate { Rdv3Migration.Migrate(f.Config, current, f.Dir, f.Dir, oldPath, badPath); });
            Check(!File.Exists(badPath), "mismatched definition created output");
        });
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
        { Rdv3Xlsx.Write(Path, Head, Work.Column, lines, states, "fixture", Rdv3Files.StorageContract(Data, Work), Rdv3LedgerProtection.Create(Data)); }
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
