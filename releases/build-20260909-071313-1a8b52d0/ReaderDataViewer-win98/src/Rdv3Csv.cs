// CSV fallback for quoting/multiline fields and non-byte-compatible encodings.
// The ordinary unquoted UTF-8/Shift-JIS/single-byte path remains in Rdv3Core.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

public static class Rdv3Csv
{
    public static bool NeedsDecoded(byte[] bytes, Encoding encoding)
    {
        int cp = encoding.CodePage;
        if (!encoding.IsSingleByte && cp != 65001 && cp != 932) { return true; }
        if (bytes.Length >= 2 && ((bytes[0] == 255 && bytes[1] == 254) || (bytes[0] == 254 && bytes[1] == 255))) { return true; }
        if (bytes.Length >= 3 && bytes[0] == 239 && bytes[1] == 187 && bytes[2] == 191 && cp != 65001) { return true; }
        for (int i = 0; i < bytes.Length; i++)
        {
            if (bytes[i] == (byte)'"' || (bytes[i] == 13 && (i + 1 == bytes.Length || bytes[i + 1] != 10))) { return true; }
        }
        return false;
    }

    public static void Read(string path, Encoding encoding, bool headOnly,
                            out string[] head, out string[][] rows, out int[] rowNumbers, string encodingSetting = "data.encoding",
                            HashSet<string> references = null, Rdv3InputCounts counts = null)
    {
        head = null;
        if (counts == null) { counts = new Rdv3InputCounts(); }
        Rdv3InputColumns columns = null;
        List<string[]> data = new List<string[]>();
        List<int> numbers = new List<int>();
        Encoding strict = (Encoding)encoding.Clone();
        strict.DecoderFallback = DecoderFallback.ExceptionFallback;
        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete))
        {
            byte[] bom = new byte[4];
            int n = stream.Read(bom, 0, bom.Length);
            int cp = 0, skip = 0;
            if (n >= 4 && bom[0] == 255 && bom[1] == 254 && bom[2] == 0 && bom[3] == 0) { cp = 12000; skip = 4; }
            else if (n >= 4 && bom[0] == 0 && bom[1] == 0 && bom[2] == 254 && bom[3] == 255) { cp = 12001; skip = 4; }
            else if (n >= 3 && bom[0] == 239 && bom[1] == 187 && bom[2] == 191) { cp = 65001; skip = 3; }
            else if (n >= 2 && bom[0] == 255 && bom[1] == 254) { cp = 1200; skip = 2; }
            else if (n >= 2 && bom[0] == 254 && bom[1] == 255) { cp = 1201; skip = 2; }
            if (cp != 0 && cp != encoding.CodePage)
            { throw Rdv3Input.Error(path, 1, "BOM", encoding.WebName, Encoding.GetEncoding(cp).WebName,
                Rdv3Text.InputFixEncoding.Replace("{encoding}", Encoding.GetEncoding(cp).WebName).Replace("{setting}", encodingSetting)); }
            stream.Position = skip;
            using (StreamReader reader = new StreamReader(stream, strict, false, 65536))
            {
                int physical = 1;
                while (true)
                {
                    int first = physical;
                    bool blank;
                    string[] cells;
                    try { cells = Record(reader, path, ref physical, out blank); }
                    catch (DecoderFallbackException)
                    {
                        // Use the same open file to locate the invalid byte; a
                        // decoder's buffered read may run ahead of this record.
                        stream.Position = 0;
                        using (MemoryStream copy = new MemoryStream())
                        { stream.CopyTo(copy); Rdv3Input.ValidateEncoding(copy.ToArray(), encoding, path, encodingSetting); }
                        throw;
                    }
                    if (cells == null) { break; }
                    if (blank) { counts.BlankRows++; continue; }
                    if (head == null)
                    {
                        head = cells;
                        for (int i = 0; i < head.Length; i++)
                        {
                            head[i] = head[i].Trim();
                            for (int k = 0; k < head[i].Length; k++)
                            { if (head[i][k] < ' ') { throw Failure(path, first, i + 1, Rdv3Text.InputExpectHeader, head[i], Rdv3Text.InputFixHeader); } }
                        }
                        columns = Rdv3InputColumns.Read(head, path, first, references, counts);
                        head = columns.Head;
                        if (headOnly) { break; }
                    }
                    else
                    {
                        if (cells.Length < columns.SourceCount) { counts.ShortRows++; continue; }
                        if (cells.Length > columns.SourceCount)
                        { throw Failure(path, first, columns.SourceCount + 1,
                            Rdv3Text.InputColumnCount.Replace("{n}", columns.SourceCount.ToString(CultureInfo.InvariantCulture)),
                            Rdv3Text.InputColumnCount.Replace("{n}", cells.Length.ToString(CultureInfo.InvariantCulture)), Rdv3Text.InputFixCsv); }
                        data.Add(columns.Project(cells));
                        numbers.Add(first);
                    }
                }
            }
        }
        if (head == null) { throw Failure(path, 1, 1, Rdv3Text.InputExpectHeader, "", Rdv3Text.InputFixHeader); }
        rows = data.ToArray();
        rowNumbers = numbers.ToArray();
    }

    private static string[] Record(TextReader reader, string path, ref int line, out bool blank)
    {
        blank = true;
        List<string> cells = new List<string>();
        StringBuilder field = new StringBuilder();
        bool quoted = false, afterQuote = false, any = false;
        int recordLine = line;
        while (true)
        {
            int read = reader.Read();
            if (read < 0)
            {
                if (quoted) { throw Failure(path, recordLine, cells.Count + 1, Rdv3Text.InputExpectQuote, "EOF", Rdv3Text.InputFixCsv); }
                if (!any) { return null; }
                cells.Add(field.ToString());
                return cells.ToArray();
            }
            char ch = (char)read;
            any = true;
            if (quoted)
            {
                if (ch == '"')
                {
                    if (reader.Peek() == '"') { reader.Read(); field.Append('"'); }
                    else { quoted = false; afterQuote = true; }
                }
                else
                {
                    field.Append(ch);
                    if (ch == '\r')
                    {
                        if (reader.Peek() == '\n') { reader.Read(); field.Append('\n'); }
                        line++;
                    }
                    else if (ch == '\n') { line++; }
                }
                continue;
            }
            if (ch == '\r' || ch == '\n')
            {
                if (ch == '\r' && reader.Peek() == '\n') { reader.Read(); }
                line++;
                cells.Add(field.ToString());
                return cells.ToArray();
            }
            blank = false;
            if (ch == ',') { cells.Add(field.ToString()); field.Length = 0; afterQuote = false; continue; }
            if (afterQuote && Rdv3Input.IsPadding(ch)) { continue; }
            if (afterQuote) { throw Failure(path, recordLine, cells.Count + 1, Rdv3Text.InputExpectDelimiter, ch.ToString(), Rdv3Text.InputFixCsv); }
            if (ch == '"')
            {
                if (field.Length != 0) { throw Failure(path, recordLine, cells.Count + 1, Rdv3Text.InputExpectQuote, field.ToString() + ch, Rdv3Text.InputFixCsv); }
                quoted = true;
            }
            else { field.Append(ch); }
        }
    }

    private static Rdv3DataError Failure(string path, int line, int column, string expected, string actual, string fix)
    { return Rdv3Input.Error(path, line, column.ToString(CultureInfo.InvariantCulture), expected, actual, fix); }
}
