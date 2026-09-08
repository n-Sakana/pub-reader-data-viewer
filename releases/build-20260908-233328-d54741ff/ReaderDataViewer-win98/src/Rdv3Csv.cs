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
                            out string[] head, out string[][] rows, out int[] rowNumbers)
    {
        head = null;
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
            { throw Error(path, 1, "BOM does not match data.encoding: " + encoding.WebName); }
            stream.Position = skip;
            using (StreamReader reader = new StreamReader(stream, strict, false, 65536))
            {
                int physical = 1;
                while (true)
                {
                    int first = physical;
                    bool blank;
                    string[] cells = Record(reader, path, ref physical, out blank);
                    if (cells == null) { break; }
                    if (blank) { continue; }
                    if (head == null)
                    {
                        head = cells;
                        HashSet<string> names = new HashSet<string>(StringComparer.Ordinal);
                        for (int i = 0; i < head.Length; i++)
                        {
                            head[i] = head[i].Trim();
                            if (head[i].Length == 0 || !names.Add(head[i])) { throw Error(path, first, "blank or duplicate header"); }
                            for (int k = 0; k < head[i].Length; k++)
                            { if (head[i][k] < ' ') { throw Error(path, first, "control character in header"); } }
                        }
                        if (headOnly) { break; }
                    }
                    else
                    {
                        if (cells.Length != head.Length) { throw Error(path, first, "column count differs from header"); }
                        data.Add(cells);
                        numbers.Add(first);
                    }
                }
            }
        }
        if (head == null) { throw Error(path, 1, "CSV has no header"); }
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
                if (quoted) { throw Error(path, recordLine, "quoted field is not closed"); }
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
            if (afterQuote) { throw Error(path, recordLine, "unexpected text after closing quote"); }
            if (ch == '"')
            {
                if (field.Length != 0) { throw Error(path, recordLine, "quote inside an unquoted field"); }
                quoted = true;
            }
            else { field.Append(ch); }
        }
    }

    private static Rdv3DataError Error(string path, int line, string reason)
    { return new Rdv3DataError(Path.GetFileName(path) + ": row " + line.ToString(CultureInfo.InvariantCulture) + ": " + reason); }
}
