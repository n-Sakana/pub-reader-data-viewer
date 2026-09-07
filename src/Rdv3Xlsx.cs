// ============================================================================
// Rdv3Xlsx.cs -- the persisted ledger: a real .xlsx, read and written directly.
//
// The distribution ships one Excel ledger file next to the .cmd. It is a
// viewing copy for the operator AND the persistence of the app: the app reads
// it at startup (saved content + work states) and rewrites it after an
// approved update or a work-state change. The first column holds the work
// state as the stored string the screen definition maps to a state name; its
// heading and the content headings are whatever the caller passes (the CSV
// column names in ledger order), and a file whose header row differs from
// them is refused.
//
// No Excel process and no library is involved: an .xlsx is a zip package, the
// in-box .NET Framework has ZipArchive, and every cell is written as an inline
// string so leading zeros survive. Reading accepts what Excel itself would
// write back (shared strings, r="D5" cell addresses, omitted empty cells), so
// a ledger a user re-saved from Excel still loads. A file that does not parse
// or whose header row is wrong is an error the caller must surface -- it is
// never silently replaced.
//
// Writing is write-temp-then-replace: the target is either the old file or the
// new one, never a torn half. A locked target (the operator has it open in
// Excel) fails the replace, and that failure is reported, not swallowed.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

public static class Rdv3Xlsx
{
    public const string SheetName = "LEDGER";

    // A workbook named by data.tables is an ordinary input table: its first
    // non-empty row is the header and the remaining rows are source records.
    // The saved ledger reader below deliberately remains separate because it
    // has the additional application-owned state column contract.
    public static string[] ReadTableHead(string path)
    {
        string[] head;
        string[][] rows;
        string warning;
        ReadTableCore(path, true, out head, out rows, out warning);
        return head;
    }

    public static void ReadTable(string path, out string[] head, out string[][] rows, out string warning)
    {
        ReadTableCore(path, false, out head, out rows, out warning);
    }

    private static void ReadTableCore(string path, bool headOnly, out string[] head,
                                      out string[][] rows, out string warning)
    {
        warning = "";
        head = null;
        List<string[]> result = new List<string[]>();
        string file = Path.GetFileName(path);
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read,
                                               FileShare.ReadWrite | FileShare.Delete))
        using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            string[] shared = ReadSharedStrings(z);
            ZipArchiveEntry sheet = FindSheet(z, false);
            if (sheet == null) { throw new InvalidDataException("no worksheet part in " + path); }
            using (XmlReader xr = Xml(sheet.Open()))
            {
                List<string> cells = null;
                bool inRow = false;
                int col = 0;
                int rowNumber = 0;
                HashSet<int> seenColumns = new HashSet<int>();
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "row")
                    {
                        inRow = true;
                        cells = new List<string>();
                        seenColumns.Clear();
                        col = 0;
                        rowNumber++;
                        int stated;
                        string rowRef = xr.GetAttribute("r");
                        if (rowRef != null && int.TryParse(rowRef, NumberStyles.Integer,
                            CultureInfo.InvariantCulture, out stated) && stated > 0) { rowNumber = stated; }
                        if (xr.IsEmptyElement) { inRow = false; }
                        continue;
                    }
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "c" && inRow)
                    {
                        string cellRef = xr.GetAttribute("r");
                        if (cellRef != null && cellRef.Length > 0)
                        {
                            int referenced = ColOf(cellRef);
                            if (referenced >= 0) { col = referenced; }
                        }
                        if (col >= 16384 || !seenColumns.Add(col)) { throw new InvalidDataException("duplicate/out-of-range XLSX cell"); }
                        while (cells.Count <= col) { cells.Add(""); }
                        cells[col] = ReadCellValue(xr, xr.GetAttribute("t"), shared);
                        col++;
                        continue;
                    }
                    if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "row" && inRow)
                    {
                        inRow = false;
                        if (cells.Count == 0) { continue; }
                        string[] values = cells.ToArray();
                        for (int c = 0; c < values.Length; c++)
                        {
                            values[c] = SafeSourceCell(values[c], file, rowNumber, ref warning);
                        }
                        if (head == null)
                        {
                            head = CheckTableHead(values, file);
                            if (headOnly) { break; }
                            continue;
                        }
                        if (values.Length > head.Length)
                        {
                            throw new Rdv3DataError(Rdv3Text.DataColumnCount.Replace("{file}", file)
                                .Replace("{row}", rowNumber.ToString(CultureInfo.InvariantCulture))
                                .Replace("{n}", values.Length.ToString(CultureInfo.InvariantCulture))
                                .Replace("{cols}", head.Length.ToString(CultureInfo.InvariantCulture)));
                        }
                        Array.Resize(ref values, head.Length);
                        for (int c = 0; c < values.Length; c++) { if (values[c] == null) { values[c] = ""; } }
                        result.Add(values);
                    }
                }
            }
        }
        if (head == null)
        {
            throw new Rdv3DataError(Rdv3Text.DataNoRows.Replace("{file}", file).Replace("{row}", "0"));
        }
        rows = result.ToArray();
    }

    private static string[] CheckTableHead(string[] values, string file)
    {
        for (int i = 0; i < values.Length; i++)
        {
            values[i] = values[i].Trim();
            if (values[i].Length == 0)
            {
                throw new Rdv3DataError(Rdv3Text.DataBlankHeader.Replace("{file}", file).Replace("{row}", "1"));
            }
            for (int k = 0; k < i; k++)
            {
                if (string.Equals(values[k], values[i], StringComparison.Ordinal))
                {
                    throw new Rdv3DataError(Rdv3Text.DataDupHeader.Replace("{file}", file)
                        .Replace("{row}", "1").Replace("{name}", values[i]));
                }
            }
        }
        return values;
    }

    private static string SafeSourceCell(string value, string file, int row, ref string warning)
    {
        if (value == null || value.Length == 0) { return ""; }
        char[] safe = null;
        for (int i = 0; i < value.Length; i++)
        {
            if (value[i] >= ' ' || value[i] == '\r' || value[i] == '\n') { continue; }
            if (warning.Length == 0)
            {
                warning = Rdv3Text.DataControlChar.Replace("{file}", file)
                    .Replace("{row}", row.ToString(CultureInfo.InvariantCulture))
                    .Replace("{code}", ((int)value[i]).ToString("X2", CultureInfo.InvariantCulture));
            }
            if (safe == null) { safe = value.ToCharArray(); }
            safe[i] = '?';
        }
        return (safe == null) ? value : new string(safe);
    }

    // ---- write -------------------------------------------------------------
    public static void Write(string path, string[] head, string stateHead, string[] lines, string[] states, string runId, string contract = null)
    {
        ValidateWrite(head, stateHead, lines, states);
        string dir = Path.GetDirectoryName(Path.GetFullPath(path));
        string tmp = Path.Combine(dir, Path.GetFileName(path) + ".tmp-" + Guid.NewGuid().ToString("N"));
        try
        {
            using (FileStream fs = new FileStream(tmp, FileMode.CreateNew, FileAccess.Write))
            using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                if (!string.IsNullOrEmpty(contract))
                { AddEntry(z, "rdv-contract.xml", "<contract>" + contract + "</contract>"); }
                AddEntry(z, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "</Types>");
                AddEntry(z, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                AddEntry(z, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheets><sheet name=\"" + SheetName + "\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                AddEntry(z, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                    "</Relationships>");
                AddEntry(z, "xl/styles.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>" +
                    "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
                    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>" +
                    "<fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                    "<borders count=\"1\"><border/></borders>" +
                    "<cellStyleXfs count=\"1\"><xf/></cellStyleXfs>" +
                    "<cellXfs count=\"2\"><xf/><xf fontId=\"1\" applyFont=\"1\"/></cellXfs>" +
                    "</styleSheet>");

                ZipArchiveEntry e = z.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Fastest);
                using (StreamWriter w = new StreamWriter(e.Open(), new UTF8Encoding(false), 1 << 16))
                {
                    w.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                    w.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                    w.Write("<sheetData>");

                    w.Write("<row>");
                    Cell(w, stateHead, true);
                    for (int c = 0; c < head.Length; c++) { Cell(w, head[c], true); }
                    w.Write("</row>");

                    for (int i = 0; i < lines.Length; i++)
                    {
                        w.Write("<row>");
                        Cell(w, states[i], false);
                        string line = lines[i];
                        int p = 0;
                        int n = line.Length;
                        for (int c = 0; c < head.Length; c++)
                        {
                            int q = line.IndexOf('\t', p);
                            if (q < 0) { q = n; }
                            CellSpan(w, line, p, q - p);
                            p = (q < n) ? q + 1 : n;
                        }
                        w.Write("</row>");
                    }
                    w.Write("</sheetData></worksheet>");
                }
            }

            if (File.Exists(path))
            {
                File.Replace(tmp, path, null);
            }
            else
            {
                File.Move(tmp, path);
            }
        }
        finally
        {
            if (File.Exists(tmp))
            {
                try { File.Delete(tmp); }
                catch (Exception) { }
            }
        }
    }

    private static void Cell(StreamWriter w, string v, bool head)
    {
        w.Write(head ? "<c t=\"inlineStr\" s=\"1\"><is><t xml:space=\"preserve\">" : "<c t=\"inlineStr\"><is><t xml:space=\"preserve\">");
        Esc(w, v, 0, v.Length);
        w.Write("</t></is></c>");
    }

    private static void CellSpan(StreamWriter w, string s, int off, int len)
    {
        w.Write("<c t=\"inlineStr\"><is><t xml:space=\"preserve\">");
        Esc(w, s, off, len);
        w.Write("</t></is></c>");
    }

    private static void Esc(StreamWriter w, string s, int off, int len)
    {
        int end = off + len;
        for (int i = off; i < end; i++)
        {
            char ch = s[i];
            if (ch == '\r') { w.Write("&#13;"); }
            else if (ch == '&') { w.Write("&amp;"); }
            else if (ch == '<') { w.Write("&lt;"); }
            else if (ch == '>') { w.Write("&gt;"); }
            else { w.Write(ch); }
        }
    }

    private static void AddEntry(ZipArchive z, string name, string content)
    {
        ZipArchiveEntry e = z.CreateEntry(name, CompressionLevel.Fastest);
        using (StreamWriter w = new StreamWriter(e.Open(), new UTF8Encoding(false)))
        {
            w.Write(content);
        }
    }

    // ---- read --------------------------------------------------------------
    public static void Read(string path, string[] head, string stateHead, out string[] lines, out string[] states)
    {
        string ignored;
        Read(path, head, stateHead, out lines, out states, out ignored);
    }

    public static void Read(string path, string[] head, string stateHead, out string[] lines, out string[] states,
                            out string warning, string expectedContract = null)
    {
        warning = "";
        // ReadWrite | Delete: another copy of the app may replace the file
        // (write-temp-then-replace) while this one is still reading it. The
        // replace then goes through, and this reader finishes the version it
        // opened. With Read alone the other copy's send would fail with a
        // sharing violation for as long as this read takes.
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            string[] shared = ReadSharedStrings(z);
            CheckContract(z, expectedContract);
            ZipArchiveEntry sheet = FindSheet(z, true);
            if (sheet == null) { throw new InvalidDataException("no worksheet part in " + path); }

            List<string> outLines = new List<string>(131072);
            List<string> outStates = new List<string>(131072);
            string[] cells = new string[head.Length + 1];
            bool sawHeader = false;

            using (XmlReader xr = Xml(sheet.Open()))
            {
                int col = 0;
                bool inRow = false;
                HashSet<int> seenColumns = new HashSet<int>();
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "row")
                    {
                        inRow = true;
                        col = 0;
                        seenColumns.Clear();
                        for (int i = 0; i < cells.Length; i++) { cells[i] = ""; }
                        if (xr.IsEmptyElement) { inRow = false; }
                        continue;
                    }
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "c" && inRow)
                    {
                        string r = xr.GetAttribute("r");
                        if (r != null && r.Length > 0)
                        {
                            int rc = ColOf(r);
                            if (rc >= 0) { col = rc; }
                        }
                        if (col >= 16384 || !seenColumns.Add(col)) { throw new InvalidDataException("duplicate/out-of-range XLSX cell"); }
                        string t = xr.GetAttribute("t");
                        string v = ReadCellValue(xr, t, shared);
                        if (col < cells.Length) { cells[col] = v; }
                        else if (v.Length > 0) { throw new InvalidDataException("ledger has unexpected non-empty columns: " + path); }
                        col++;
                        continue;
                    }
                    if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "row" && inRow)
                    {
                        inRow = false;
                        if (!sawHeader)
                        {
                            CheckHeader(cells, head, stateHead, path);
                            sawHeader = true;
                            continue;
                        }
                        // A ledger line is tab-separated. Replace a tab typed
                        // into a cell before materialising that internal row.
                        for (int c = 0; c < cells.Length; c++)
                        {
                            if (cells[c].IndexOf('\t') >= 0)
                            {
                                if (warning.Length == 0)
                                {
                                    warning = Rdv3Text.DataLedgerTab.Replace("{file}", Path.GetFileName(path))
                                        .Replace("{row}", (outLines.Count + 2).ToString(CultureInfo.InvariantCulture));
                                }
                                cells[c] = cells[c].Replace('\t', '?');
                            }
                        }
                        StringBuilder sb = new StringBuilder(256);
                        for (int c = 0; c < head.Length; c++)
                        {
                            if (c > 0) { sb.Append('\t'); }
                            sb.Append(cells[c + 1]);
                        }
                        outLines.Add(sb.ToString());
                        outStates.Add(cells[0]);
                    }
                }
            }
            if (!sawHeader) { throw new InvalidDataException("ledger has no header row: " + path); }
            lines = outLines.ToArray();
            states = outStates.ToArray();
        }
    }

    private static void CheckHeader(string[] cells, string[] head, string stateHead, string path)
    {
        bool ok = string.Equals(cells[0], stateHead, StringComparison.Ordinal);
        for (int c = 0; ok && c < head.Length; c++)
        {
            if (!string.Equals(cells[c + 1], head[c], StringComparison.Ordinal)) { ok = false; }
        }
        if (!ok)
        {
            throw new Rdv3DataError(Rdv3Text.DataLedgerHeader.Replace("{file}", Path.GetFileName(path)));
        }
    }

    // xr is positioned on a non-empty <c> element. Collects the text of every
    // v / t inside it (a rich-text inline string has several t runs), and
    // leaves xr on the </c> so the caller's next Read() moves cleanly on.
    private static string ReadCellValue(XmlReader xr, string t, string[] shared)
    {
        if (xr.IsEmptyElement) { return ""; }
        StringBuilder v = new StringBuilder();
        bool formula = false, cached = false;
        using (XmlReader sub = xr.ReadSubtree())
        {
            bool moved = sub.Read();
            while (moved)
            {
                if (sub.NodeType == XmlNodeType.Element && sub.LocalName == "rPh")
                { sub.Skip(); moved = !sub.EOF; continue; }
                if (sub.NodeType == XmlNodeType.Element && sub.LocalName == "f") { formula = true; }
                if (sub.NodeType == XmlNodeType.Element && sub.LocalName == "v") { cached = true; }
                if (sub.NodeType == XmlNodeType.Element &&
                    (sub.LocalName == "v" || sub.LocalName == "t") && !sub.IsEmptyElement)
                {
                    v.Append(sub.ReadElementContentAsString());
                    // the reader is already on the node after the element
                    moved = !sub.EOF;
                    continue;
                }
                moved = sub.Read();
            }
        }
        if (formula && !cached) { throw new InvalidDataException("XLSX formula has no cached result; recalculate and save the source workbook"); }
        return Resolve(t, v.ToString(), shared);
    }

    private static string Resolve(string t, string s, string[] shared)
    {
        if (t == "s")
        {
            int idx;
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out idx) &&
                shared != null && idx >= 0 && idx < shared.Length)
            {
                if (shared[idx].Length > 32767) { throw new InvalidDataException("XLSX shared string exceeds 32767 characters"); }
                return shared[idx];
            }
            throw new InvalidDataException("invalid shared-string index in XLSX");
        }
        if (t == "e") { throw new InvalidDataException("XLSX contains a formula/error cell: " + s); }
        if (s.Length > 32767) { throw new InvalidDataException("XLSX cell exceeds 32767 characters"); }
        return s;
    }

    private static string[] ReadSharedStrings(ZipArchive z)
    {
        ZipArchiveEntry e = z.GetEntry("xl/sharedStrings.xml");
        if (e == null) { return null; }
        List<string> all = new List<string>();
        using (XmlReader xr = Xml(e.Open()))
        {
            StringBuilder cur = null;
            bool moved = xr.Read();
            while (moved)
            {
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "rPh")
                { xr.Skip(); moved = !xr.EOF; continue; }
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "si")
                {
                    if (cur != null) { all.Add(cur.ToString()); }
                    cur = new StringBuilder();
                    if (xr.IsEmptyElement) { all.Add(""); cur = null; }
                    moved = xr.Read();
                    continue;
                }
                if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "t" && cur != null && !xr.IsEmptyElement)
                {
                    cur.Append(xr.ReadElementContentAsString());
                    // the reader is already on the node after the element
                    moved = !xr.EOF;
                    continue;
                }
                moved = xr.Read();
            }
            if (cur != null) { all.Add(cur.ToString()); }
        }
        return all.ToArray();
    }

    private static XmlReader Xml(Stream stream)
    {
        XmlReaderSettings settings = new XmlReaderSettings();
        settings.DtdProcessing = DtdProcessing.Prohibit;
        settings.XmlResolver = null;
        settings.CloseInput = true;
        settings.MaxCharactersInDocument = 512L * 1024L * 1024L;
        return XmlReader.Create(stream, settings);
    }

    private static XmlDocument Document(ZipArchiveEntry entry)
    {
        if (entry == null) { throw new InvalidDataException("XLSX workbook metadata is missing"); }
        XmlDocument doc = new XmlDocument();
        doc.XmlResolver = null;
        using (XmlReader reader = Xml(entry.Open())) { doc.Load(reader); }
        return doc;
    }

    private static ZipArchiveEntry FindSheet(ZipArchive z, bool ledger)
    {
        XmlDocument workbook = Document(z.GetEntry("xl/workbook.xml"));
        XmlDocument relations = Document(z.GetEntry("xl/_rels/workbook.xml.rels"));
        XmlElement selected = null;
        int sheetCount = 0;
        foreach (XmlNode node in workbook.GetElementsByTagName("*"))
        {
            XmlElement sheet = node as XmlElement;
            if (sheet == null || sheet.LocalName != "sheet") { continue; }
            sheetCount++;
            if ((!ledger && selected == null) || (ledger && sheet.GetAttribute("name") == SheetName)) { selected = sheet; }
        }
        // The writer creates a single managed sheet. Do not silently discard
        // additional sheets someone added to the shared ledger.
        if (ledger && sheetCount != 1) { throw new InvalidDataException("shared ledger must contain only the LEDGER worksheet; preserve additional sheets separately"); }
        if (selected == null) { throw new InvalidDataException(ledger ? "LEDGER worksheet is missing" : "no worksheet in workbook"); }
        string id = "";
        foreach (XmlAttribute attribute in selected.Attributes)
        { if (attribute.LocalName == "id") { id = attribute.Value; } }
        foreach (XmlNode node in relations.GetElementsByTagName("*"))
        {
            XmlElement relation = node as XmlElement;
            if (relation == null || relation.LocalName != "Relationship" || relation.GetAttribute("Id") != id) { continue; }
            if (relation.GetAttribute("TargetMode") == "External"
                || !relation.GetAttribute("Type").EndsWith("/worksheet", StringComparison.Ordinal))
            { throw new InvalidDataException("selected XLSX sheet is not a local worksheet"); }
            Uri target = new Uri(new Uri("http://rdv.local/xl/workbook.xml"), relation.GetAttribute("Target").Replace('\\', '/'));
            if (target.Host != "rdv.local" || target.Scheme != "http") { throw new InvalidDataException("invalid worksheet target"); }
            ZipArchiveEntry entry = z.GetEntry(Uri.UnescapeDataString(target.AbsolutePath).TrimStart('/'));
            if (entry == null) { throw new InvalidDataException("worksheet relationship points to a missing part"); }
            return entry;
        }
        throw new InvalidDataException("worksheet relationship is missing");
    }

    private static void CheckContract(ZipArchive zip, string expected)
    {
        if (string.IsNullOrEmpty(expected)) { return; }
        ZipArchiveEntry entry = zip.GetEntry("rdv-contract.xml");
        if (entry == null) { return; } // Legacy ledger: header validation still applies.
        XmlDocument doc = Document(entry);
        if (doc.DocumentElement == null || doc.DocumentElement.LocalName != "contract"
            || doc.DocumentElement.InnerText != expected)
        { throw new InvalidDataException(Rdv3Text.StorageContractMismatch); }
    }

    private static void ValidateCell(string value)
    {
        if (value == null || value.Length > 32767) { throw new InvalidDataException("invalid/oversized XLSX cell"); }
        XmlConvert.VerifyXmlChars(value);
    }

    private static void ValidateWrite(string[] head, string stateHead, string[] lines, string[] states)
    {
        if (head == null || lines == null || states == null || lines.Length != states.Length
            || head.Length == 0 || head.Length >= 16384 || lines.Length >= 1048576)
        { throw new InvalidDataException("invalid XLSX dimensions or state count"); }
        ValidateCell(stateHead);
        for (int c = 0; c < head.Length; c++) { ValidateCell(head[c]); }
        for (int r = 0; r < lines.Length; r++)
        {
            ValidateCell(states[r]);
            if (lines[r] == null) { throw new InvalidDataException("null ledger row"); }
            string[] cells = lines[r].Split('\t');
            if (cells.Length != head.Length) { throw new InvalidDataException("ledger row column count differs from header"); }
            for (int c = 0; c < cells.Length; c++) { ValidateCell(cells[c]); }
        }
    }

    // "D5" -> 3. Letters only matter; the row digits are ignored.
    private static int ColOf(string cellRef)
    {
        int col = 0;
        bool any = false;
        for (int i = 0; i < cellRef.Length; i++)
        {
            char ch = cellRef[i];
            if (ch >= 'A' && ch <= 'Z') { col = col * 26 + (ch - 'A' + 1); any = true; }
            else if (ch >= 'a' && ch <= 'z') { col = col * 26 + (ch - 'a' + 1); any = true; }
            else { break; }
            if (col > 16384) { throw new InvalidDataException("XLSX column reference exceeds XFD"); }
        }
        return any ? col - 1 : -1;
    }
}
