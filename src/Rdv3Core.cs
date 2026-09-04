// ============================================================================
// Rdv3Core.cs -- clock, CSV tables and the data error.
//
// The CSVs are read once, in the background, when the app starts or the
// operator asks for an update: their merge is the integrated ledger, which is
// persisted and searched. Searching never re-reads the CSVs.
//
// The CSVs are RAW data and the tables are described by settings.json
// ("data": which files, which column is the key, how they join -- Rdv3Data).
// A table's column names are its header row. Key validation is strict by
// default; a definition can opt into Unicode, variable width, distinct values
// and skipped blank keys for condition-list inputs.
//
// Structural errors are refused: a wrong column count, a quoted field (this
// reader does not unquote), or an unusable key. A control character inside a
// field is different: refusing a whole input for one old byte is too broad.
// The first occurrence is reported and each such byte is replaced with '?',
// keeping both the ledger's tab-separated rows and its XML safe. Bytes that
// are not valid in the declared encoding are likewise decoded with the
// replacement character, and the first such row is remembered for the UI.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// See build\pack_app.ps1.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

// the input (a CSV or the saved ledger) cannot be used as it is; the message
// is the operator's, naming the file and the row
public sealed class Rdv3DataError : Exception
{
    public Rdv3DataError(string msg) : base(msg) { }
}

public static class Rdv3Clock
{
    private static readonly double TickMs = 1000.0 / (double)Stopwatch.Frequency;

    public static long Now()
    {
        return Stopwatch.GetTimestamp();
    }

    public static double MsSince(long t0)
    {
        return (double)(Stopwatch.GetTimestamp() - t0) * TickMs;
    }

    public static double MsBetween(long t0, long t1)
    {
        return (double)(t1 - t0) * TickMs;
    }

    public static string Fmt(double ms)
    {
        return ms.ToString("N1", CultureInfo.InvariantCulture);
    }
}

public sealed class Rdv3KeyValidation
{
    public bool Ascii = true;
    public bool FixedLength = true;
    public bool Unique = true;
    public bool SkipEmpty;

    public bool UsesFixedAsciiPath { get { return Ascii && FixedLength; } }
}

// One CSV: the bytes off the disk plus where each row starts and ends.
public sealed class Rdv3Table
{
    public string Name;                      // the table id of the definition ("A")
    public string Path;
    public byte[] Buf;
    public Encoding Enc;
    public int Rows;
    public int[] Start;
    public int[] End;
    // the header row, one name per column
    public string[] Head;
    // the key column: its index in Head, the byte offset of the key in each
    // row, and its width (taken from the first data row)
    public int KeyCol;
    public int[] KeyAt;
    public int KeyLen;
    public int[] KeyLengths;
    public int[] SourceRows;
    public Rdv3KeyValidation KeyValidation = new Rdv3KeyValidation();
    public int InvalidEncodingRow;
    public string ControlCharacterWarning = "";
    // Workbook sources have already been decoded into cells. CSV sources keep
    // their byte slices above so the common large-file path stays unchanged.
    public string[][] Cells;

    private static string Fmt(string text, string file, int row)
    {
        return text.Replace("{file}", file).Replace("{row}", row.ToString(CultureInfo.InvariantCulture));
    }

    // the header row only (for the start-up check of the definition against
    // the data, before the full read in the worker)
    public static string[] ReadHead(string path, Encoding enc)
    {
        if (string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return Rdv3Xlsx.ReadTableHead(path);
        }
        string file = System.IO.Path.GetFileName(path);
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
        using (StreamReader r = new StreamReader(fs, enc, true))
        {
            string first = r.ReadLine();
            if (first == null || first.Trim().Length == 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoRows, file, 0)); }
            string ignored;
            return SplitHead(first, file, out ignored);
        }
    }

    private static string[] SplitHead(string line, string file, out string warning)
    {
        warning = "";
        if (line.Length > 0 && line[0] == '\uFEFF') { line = line.Substring(1); }
        char[] safe = line.ToCharArray();
        bool changed = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] >= ' ') { continue; }
            if (warning.Length == 0) { warning = ControlChar(file, 1, (int)line[i]); }
            safe[i] = '?';
            changed = true;
        }
        if (changed) { line = new string(safe); }
        string[] h = line.Split(',');
        for (int i = 0; i < h.Length; i++)
        {
            h[i] = h[i].Trim();
            if (h[i].Length == 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataBlankHeader, file, 1)); }
            if (h[i][0] == '"') { throw new Rdv3DataError(Fmt(Rdv3Text.DataQuoted, file, 1)); }
            for (int k = 0; k < i; k++)
            {
                if (string.Equals(h[k], h[i], StringComparison.Ordinal))
                {
                    throw new Rdv3DataError(Fmt(Rdv3Text.DataDupHeader, file, 1).Replace("{name}", h[i]));
                }
            }
        }
        return h;
    }

    // the whole table, validated row by row
    public static Rdv3Table Read(string path, string name, Encoding enc, string keyName)
    {
        return Read(path, name, enc, keyName, null);
    }

    public static Rdv3Table Read(string path, string name, Encoding enc, string keyName,
                                 Rdv3KeyValidation validation)
    {
        if (validation == null) { validation = new Rdv3KeyValidation(); }
        if (string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return ReadWorkbook(path, name, enc, keyName, validation);
        }
        Rdv3Table t = new Rdv3Table();
        t.Name = name;
        t.Path = path;
        t.Enc = enc;
        t.KeyValidation = validation;
        t.Buf = File.ReadAllBytes(path);
        t.InvalidEncodingRow = FindInvalidEncodingRow(t.Buf, enc);
        string file = System.IO.Path.GetFileName(path);

        byte[] b = t.Buf;
        int n = b.Length;
        int cap = 65536;
        int[] st = new int[cap];
        int[] en = new int[cap];
        int rows = 0;
        int pos = 0;
        if (n >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) { pos = 3; }

        while (pos < n)
        {
            int nl = pos;
            while (nl < n && b[nl] != (byte)'\n') { nl++; }
            int e = nl;
            if (e > pos && b[e - 1] == (byte)'\r') { e--; }
            if (e > pos)
            {
                if (rows == cap)
                {
                    cap = cap << 1;
                    Array.Resize(ref st, cap);
                    Array.Resize(ref en, cap);
                }
                st[rows] = pos;
                en[rows] = e;
                rows++;
            }
            pos = nl + 1;
        }
        if (rows < 2) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoRows, file, 0)); }

        string headerWarning;
        t.Head = SplitHead(enc.GetString(b, st[0], en[0] - st[0]), file, out headerWarning);
        t.ControlCharacterWarning = headerWarning;
        int cols = t.Head.Length;
        t.KeyCol = -1;
        for (int i = 0; i < cols; i++) { if (t.Head[i] == keyName) { t.KeyCol = i; } }
        if (t.KeyCol < 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoColumn, file, 1).Replace("{name}", keyName)); }

        int sourceRows = rows - 1;
        t.Rows = 0;
        t.Start = new int[sourceRows];
        t.End = new int[sourceRows];
        t.KeyAt = new int[sourceRows];
        if (!validation.UsesFixedAsciiPath) { t.KeyLengths = new int[sourceRows]; }
        if (validation.SkipEmpty || !validation.Unique) { t.SourceRows = new int[sourceRows]; }
        HashSet<string> distinct = validation.Unique ? null : new HashSet<string>(StringComparer.Ordinal);
        int fixedLength = -1;
        for (int source = 0; source < sourceRows; source++)
        {
            int rs = st[source + 1];
            int re = en[source + 1];
            int row = source + 2;              // the row number a person sees (header = 1)
            // one pass over the row: count the columns, refuse a quoted field,
            // and find the key
            int field = 0;
            int p = rs;
            int keyAt = -1;
            int keyEnd = -1;
            while (true)
            {
                if (p < re && b[p] == (byte)'"') { throw new Rdv3DataError(Fmt(Rdv3Text.DataQuoted, file, row)); }
                int q = p;
                while (q < re && b[q] != (byte)',')
                {
                    if (b[q] < 0x20)
                    {
                        if (t.ControlCharacterWarning.Length == 0)
                        {
                            t.ControlCharacterWarning = ControlChar(file, row, (int)b[q]);
                        }
                        b[q] = (byte)'?';
                    }
                    q++;
                }
                if (field == t.KeyCol) { keyAt = p; keyEnd = q; }
                field++;
                if (q >= re) { break; }
                p = q + 1;
            }
            if (field != cols)
            {
                throw new Rdv3DataError(Fmt(Rdv3Text.DataColumnCount, file, row)
                    .Replace("{n}", field.ToString(CultureInfo.InvariantCulture))
                    .Replace("{cols}", cols.ToString(CultureInfo.InvariantCulture)));
            }
            int klen = keyEnd - keyAt;
            if (klen <= 0)
            {
                if (validation.SkipEmpty) { continue; }
                throw new Rdv3DataError(Fmt(Rdv3Text.DataEmptyKey, file, row).Replace("{name}", keyName));
            }
            if (validation.Ascii)
            {
                for (int k = keyAt; k < keyEnd; k++)
                {
                    if (b[k] > 127) { throw new Rdv3DataError(Fmt(Rdv3Text.DataKeyNotAscii, file, row).Replace("{name}", keyName)); }
                }
            }
            string key = null;
            int logicalLength = klen;
            if (!validation.Ascii || distinct != null)
            {
                key = enc.GetString(b, keyAt, klen);
                if (!validation.Ascii) { logicalLength = key.Length; }
            }
            if (validation.FixedLength)
            {
                if (fixedLength < 0) { fixedLength = logicalLength; }
                else if (logicalLength != fixedLength)
                {
                    throw new Rdv3DataError(Fmt(Rdv3Text.DataKeyWidth, file, row)
                        .Replace("{name}", keyName)
                        .Replace("{n}", fixedLength.ToString(CultureInfo.InvariantCulture)));
                }
            }
            if (distinct != null && !distinct.Add(key)) { continue; }

            int kept = t.Rows;
            t.Start[kept] = rs;
            t.End[kept] = re;
            t.KeyAt[kept] = keyAt;
            if (t.KeyLengths != null) { t.KeyLengths[kept] = klen; }
            if (t.SourceRows != null) { t.SourceRows[kept] = row; }
            t.Rows++;
        }
        t.KeyLen = validation.UsesFixedAsciiPath && fixedLength > 0 ? fixedLength : 0;
        Array.Resize(ref t.Start, t.Rows);
        Array.Resize(ref t.End, t.Rows);
        Array.Resize(ref t.KeyAt, t.Rows);
        if (t.KeyLengths != null) { Array.Resize(ref t.KeyLengths, t.Rows); }
        if (t.SourceRows != null) { Array.Resize(ref t.SourceRows, t.Rows); }
        return t;
    }

    private static Rdv3Table ReadWorkbook(string path, string name, Encoding enc, string keyName,
                                          Rdv3KeyValidation validation)
    {
        Rdv3Table t = new Rdv3Table();
        t.Name = name;
        t.Path = path;
        t.Enc = enc;
        t.KeyValidation = validation;
        string warning;
        Rdv3Xlsx.ReadTable(path, out t.Head, out t.Cells, out warning);
        t.ControlCharacterWarning = warning;
        t.KeyCol = t.ColumnOf(keyName);
        if (t.KeyCol < 0)
        {
            throw new Rdv3DataError(Fmt(Rdv3Text.DataNoColumn, System.IO.Path.GetFileName(path), 1)
                .Replace("{name}", keyName));
        }
        string[][] source = t.Cells;
        List<string[]> kept = new List<string[]>(source.Length);
        List<int> sourceRows = (validation.SkipEmpty || !validation.Unique) ? new List<int>(source.Length) : null;
        HashSet<string> distinct = validation.Unique ? null : new HashSet<string>(StringComparer.Ordinal);
        int fixedLength = -1;
        string file = System.IO.Path.GetFileName(path);
        for (int i = 0; i < source.Length; i++)
        {
            string key = source[i][t.KeyCol];
            int row = i + 2;
            if (key.Length == 0)
            {
                if (validation.SkipEmpty) { continue; }
                throw new Rdv3DataError(Fmt(Rdv3Text.DataEmptyKey, file, row).Replace("{name}", keyName));
            }
            if (validation.Ascii)
            {
                for (int k = 0; k < key.Length; k++)
                {
                    if (key[k] > 127)
                    {
                        throw new Rdv3DataError(Fmt(Rdv3Text.DataKeyNotAscii, file, row).Replace("{name}", keyName));
                    }
                }
            }
            if (validation.FixedLength)
            {
                if (fixedLength < 0) { fixedLength = key.Length; }
                else if (key.Length != fixedLength)
                {
                    throw new Rdv3DataError(Fmt(Rdv3Text.DataKeyWidth, file, row)
                        .Replace("{name}", keyName)
                        .Replace("{n}", fixedLength.ToString(CultureInfo.InvariantCulture)));
                }
            }
            if (distinct != null && !distinct.Add(key)) { continue; }
            kept.Add(source[i]);
            if (sourceRows != null) { sourceRows.Add(row); }
        }
        t.Cells = kept.ToArray();
        t.Rows = t.Cells.Length;
        t.KeyAt = new int[t.Rows];
        t.SourceRows = (sourceRows == null) ? null : sourceRows.ToArray();
        t.KeyLen = validation.UsesFixedAsciiPath && fixedLength > 0 ? fixedLength : 0;
        return t;
    }

    // A tab, a carriage return or any other control character, named by code.
    private static string ControlChar(string file, int row, int code)
    {
        return Fmt(Rdv3Text.DataControlChar, file, row).Replace("{code}", code.ToString("X2", CultureInfo.InvariantCulture));
    }

    private static int FindInvalidEncodingRow(byte[] bytes, Encoding enc)
    {
        Encoding strict = (Encoding)enc.Clone();
        strict.DecoderFallback = DecoderFallback.ExceptionFallback;
        try
        {
            strict.GetCharCount(bytes);
            return 0;
        }
        catch (DecoderFallbackException ex)
        {
            int stop = Math.Max(0, Math.Min(ex.Index, bytes.Length));
            int row = 1;
            for (int i = 0; i < stop; i++) { if (bytes[i] == (byte)'\n') { row++; } }
            return row;
        }
    }

    // byte offset of field f in row i, -1 if the row has no such field
    public int FieldAt(int i, int f, out int len)
    {
        if (Cells != null) { len = 0; return -1; }
        int p = Start[i];
        int e = End[i];
        for (int k = 0; k < f; k++)
        {
            while (p < e && Buf[p] != (byte)',') { p++; }
            if (p >= e) { len = 0; return -1; }
            p++;
        }
        int q = p;
        while (q < e && Buf[q] != (byte)',') { q++; }
        len = q - p;
        return p;
    }

    public string Field(int i, int f)
    {
        if (Cells != null)
        {
            return (i < 0 || i >= Cells.Length || f < 0 || f >= Cells[i].Length) ? "" : Cells[i][f];
        }
        int len;
        int p = FieldAt(i, f, out len);
        if (p < 0 || len <= 0) { return ""; }
        return Enc.GetString(Buf, p, len);
    }

    public string Key(int i)
    {
        if (Cells != null) { return Cells[i][KeyCol]; }
        int len = (KeyLengths == null) ? KeyLen : KeyLengths[i];
        return Enc.GetString(Buf, KeyAt[i], len);
    }

    public int SourceRow(int i)
    {
        return (SourceRows == null) ? i + 2 : SourceRows[i];
    }

    public int ColumnOf(string name)
    {
        for (int i = 0; i < Head.Length; i++) { if (Head[i] == name) { return i; } }
        return -1;
    }
}
