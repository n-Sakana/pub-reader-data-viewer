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

    // ---- write -------------------------------------------------------------
    public static void Write(string path, string[] head, string stateHead, string[] lines, string[] states, string runId)
    {
        string dir = Path.GetDirectoryName(Path.GetFullPath(path));
        string tmp = Path.Combine(dir, Path.GetFileName(path) + ".tmp-" + runId);
        try
        {
            using (FileStream fs = new FileStream(tmp, FileMode.Create, FileAccess.Write))
            using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Create))
            {
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
        w.Write(head ? "<c t=\"inlineStr\" s=\"1\"><is><t>" : "<c t=\"inlineStr\"><is><t>");
        Esc(w, v, 0, v.Length);
        w.Write("</t></is></c>");
    }

    private static void CellSpan(StreamWriter w, string s, int off, int len)
    {
        w.Write("<c t=\"inlineStr\"><is><t>");
        Esc(w, s, off, len);
        w.Write("</t></is></c>");
    }

    private static void Esc(StreamWriter w, string s, int off, int len)
    {
        int end = off + len;
        for (int i = off; i < end; i++)
        {
            char ch = s[i];
            if (ch == '&') { w.Write("&amp;"); }
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
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Read))
        {
            string[] shared = ReadSharedStrings(z);
            ZipArchiveEntry sheet = FindSheet(z);
            if (sheet == null) { throw new InvalidDataException("no worksheet part in " + path); }

            List<string> outLines = new List<string>(131072);
            List<string> outStates = new List<string>(131072);
            string[] cells = new string[head.Length + 1];
            bool sawHeader = false;

            using (XmlReader xr = XmlReader.Create(sheet.Open()))
            {
                int col = 0;
                bool inRow = false;
                while (xr.Read())
                {
                    if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "row")
                    {
                        inRow = true;
                        col = 0;
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
                        string t = xr.GetAttribute("t");
                        string v = ReadCellValue(xr, t, shared);
                        if (col < cells.Length) { cells[col] = v; }
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
        using (XmlReader sub = xr.ReadSubtree())
        {
            bool moved = sub.Read();
            while (moved)
            {
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
                return shared[idx];
            }
            return "";
        }
        return s;
    }

    private static string[] ReadSharedStrings(ZipArchive z)
    {
        ZipArchiveEntry e = z.GetEntry("xl/sharedStrings.xml");
        if (e == null) { return null; }
        List<string> all = new List<string>();
        using (XmlReader xr = XmlReader.Create(e.Open()))
        {
            StringBuilder cur = null;
            bool moved = xr.Read();
            while (moved)
            {
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

    private static ZipArchiveEntry FindSheet(ZipArchive z)
    {
        ZipArchiveEntry e = z.GetEntry("xl/worksheets/sheet1.xml");
        if (e != null) { return e; }
        foreach (ZipArchiveEntry x in z.Entries)
        {
            if (x.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) &&
                x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                return x;
            }
        }
        return null;
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
        }
        return any ? col - 1 : -1;
    }
}

