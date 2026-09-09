// ============================================================================
// Rdv3Core.cs -- clock, CSV tables and the data error.
//
// The CSVs are read once, in the background, when the app starts or the
// operator asks for an update: their merge is the integrated ledger, which is
// persisted and searched. Searching never re-reads the CSVs.
//
// The CSVs are RAW data and the tables are described by settings.json
// ("data": which files, which column is the key, how they join -- Rdv3Data).
// A table's column names are its header row. Key rules select ASCII or
// Unicode, fixed or variable width, and unique or explicit distinct values.
// Invalid records and conflicting key groups are reported and excluded.
//
// Short CSV records and blank lines are counted and skipped. Extra columns,
// referenced duplicate headers, malformed quoting and invalid encodings
// still stop a structurally unusable input. Quoted CSV uses Rdv3Csv.
// A control character outside a key is reported and sanitized:
// The first occurrence is reported and each such byte is replaced with '?',
// keeping both the ledger's tab-separated rows and its XML safe. Bytes that
// are not valid in the declared encoding cause an error rather than
// silently replacing parts of identities or values.
//
// C# 5 only, no verbatim strings, ASCII only outside Rdv3Text.cs.
// See build.bat and tests/README.md.
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
    public bool SkipEmpty = true;
    public string SettingsPath = "";

    public bool UsesFixedAsciiPath { get { return Ascii && FixedLength; } }
}

// One CSV: the bytes off the disk plus where each row starts and ends.
public sealed class Rdv3Table
{
    public string Name;                      // the table id of the definition ("A")
    public string Path;
    public byte[] Buf;
    public Encoding Enc;
    public char Delimiter = ',';             // the field separator of a CSV source
    public int Rows;
    public int[] Start;
    public int[] End;
    // the header row, one name per column
    public string[] Head;
    // the key column: its index in Head, the byte offset of the key in each
    // row, and its width (taken from the first data row)
    public int KeyCol;
    public int[] KeyCols;
    public int[] KeyAt;
    public int KeyLen;
    public int[] KeyLengths;
    public int[] SourceRows;
    public Rdv3KeyValidation KeyValidation = new Rdv3KeyValidation();
    public int InvalidEncodingRow;
    public string ControlCharacterWarning = "";
    public int SkippedEmptyRows;
    public int SkippedDuplicateRows;
    public readonly Rdv3InputCounts InputCounts = new Rdv3InputCounts();
    // Workbook sources have already been decoded into cells. CSV sources keep
    // their byte slices above so the common large-file path stays unchanged.
    public string[][] Cells;

    private static string Fmt(string text, string file, int row)
    {
        return text.Replace("{file}", file).Replace("{row}", row.ToString(CultureInfo.InvariantCulture));
    }

    // the header row only (for the start-up check of the definition against
    // the data, before the full read in the worker)
    public static string[] ReadHead(string path, Encoding enc, string encodingSetting = "data.encoding",
                                    HashSet<string> references = null, int headerRow = 1, char delimiter = ',',
                                    string sheet = "")
    {
        if (string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return Rdv3Xlsx.ReadTableHead(path, references, headerRow, sheet);
        }
        string[] head;
        string[][] rows;
        int[] rowNumbers;
        Rdv3Csv.Read(path, enc, true, out head, out rows, out rowNumbers, encodingSetting, references, null, headerRow, delimiter);
        return head;
    }

    private static string[] SplitHead(string line, string file, int row, out string warning, char delimiter)
    {
        warning = "";
        if (line.Length > 0 && line[0] == '\uFEFF') { line = line.Substring(1); }
        // A header with tabs and none of the configured separator is a
        // tab-separated file; name that instead of reading one odd column.
        if (delimiter != '\t' && line.IndexOf('\t') >= 0 && line.IndexOf(delimiter) < 0)
        { throw Rdv3Input.Error(file, row, "1", Rdv3Text.InputExpectHeader, line.Replace('\t', ' '), Rdv3Text.InputFixTab); }
        char[] safe = line.ToCharArray();
        bool changed = false;
        for (int i = 0; i < line.Length; i++)
        {
            if (line[i] >= ' ') { continue; }
            if (warning.Length == 0) { warning = ControlChar(file, row, (int)line[i]); }
            safe[i] = '?';
            changed = true;
        }
        if (changed) { line = new string(safe); }
        string[] h = line.Split(delimiter);
        for (int i = 0; i < h.Length; i++)
        {
            h[i] = h[i].Trim();
            if (h[i].Length == 0)
            { throw Rdv3Input.Error(file, row, (i + 1).ToString(CultureInfo.InvariantCulture), Rdv3Text.InputExpectHeader, h[i], Rdv3Text.InputFixHeader); }
            if (h[i][0] == '"') { throw new Rdv3DataError(Fmt(Rdv3Text.DataQuoted, file, row)); }
        }
        return h;
    }

    // the whole table, validated row by row
    public static Rdv3Table Read(string path, string name, Encoding enc, string keyName)
    {
        return Read(path, name, enc, keyName, null);
    }

    // headerRow: the physical line of the header (1 = first line); lines above
    // it are a report title or similar and are skipped, not counted.
    public static Rdv3Table Read(string path, string name, Encoding enc, string keyName,
                                 Rdv3KeyValidation validation, string encodingSetting = "data.encoding",
                                 HashSet<string> references = null, int headerRow = 1, char delimiter = ',',
                                 string sheet = "")
    {
        if (validation == null) { validation = new Rdv3KeyValidation(); }
        if (references != null) { references = new HashSet<string>(references, StringComparer.Ordinal); references.Add(keyName); }
        if (string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            return ReadDecoded(path, name, enc, new string[] { keyName }, validation, true, encodingSetting, references, headerRow, delimiter, sheet);
        }
        Rdv3Table t = new Rdv3Table();
        t.Name = name;
        t.Path = path;
        t.Enc = enc;
        t.Delimiter = delimiter;
        t.KeyValidation = validation;
        t.Buf = File.ReadAllBytes(path);
        if (Rdv3Csv.NeedsDecoded(t.Buf, enc)) { return ReadDecoded(path, name, enc, new string[] { keyName }, validation, false, encodingSetting, references, headerRow, delimiter); }
        Rdv3Input.ValidateEncoding(t.Buf, enc, path, encodingSetting, delimiter);
        string file = System.IO.Path.GetFileName(path);

        byte[] b = t.Buf;
        byte sep = (byte)delimiter;
        int n = b.Length;
        int cap = 65536;
        int[] st = new int[cap];
        int[] en = new int[cap];
        int rows = 0;
        int physicalRow = 1;
        List<int> physicalRows = new List<int>();
        int pos = 0;
        if (n >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) { pos = 3; }

        if (headerRow > 1) { t.InputCounts.HeaderOffset = headerRow - 1; }
        while (pos < n)
        {
            int nl = pos;
            while (nl < n && b[nl] != (byte)'\n') { nl++; }
            int e = nl;
            if (e > pos && b[e - 1] == (byte)'\r') { e--; }
            if (physicalRow < headerRow) { /* a title line above the header: neither a row nor a blank */ }
            else if (e > pos)
            {
                if (rows == cap)
                {
                    cap = cap << 1;
                    Array.Resize(ref st, cap);
                    Array.Resize(ref en, cap);
                }
                physicalRows.Add(physicalRow);
                st[rows] = pos;
                en[rows] = e;
                rows++;
            }
            else { t.InputCounts.Shape(path, physicalRow, 0, 0, true); }
            pos = nl + 1;
            physicalRow++;
        }
        if (rows < 1) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoRows, file, 0)); }

        string headerWarning;
        t.Head = SplitHead(enc.GetString(b, st[0], en[0] - st[0]), file, physicalRows[0], out headerWarning, delimiter);
        Rdv3InputColumns columns = Rdv3InputColumns.Read(t.Head, path, physicalRows[0], references, t.InputCounts);
        if (t.InputCounts.HeaderColumns > 0)
        { return ReadDecoded(path, name, enc, new string[] { keyName }, validation, false, encodingSetting, references, headerRow, delimiter); }
        t.Head = columns.Head;
        t.ControlCharacterWarning = headerWarning;
        int cols = t.Head.Length;
        t.KeyCol = -1;
        for (int i = 0; i < cols; i++) { if (t.Head[i] == keyName) { t.KeyCol = i; } }
        if (t.KeyCol < 0) { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoColumn, file, 1).Replace("{name}", keyName)); }
        t.KeyCols = new int[] { t.KeyCol };

        int sourceRows = rows - 1;
        t.Rows = 0;
        t.Start = new int[sourceRows];
        t.End = new int[sourceRows];
        t.KeyAt = new int[sourceRows];
        if (!validation.UsesFixedAsciiPath) { t.KeyLengths = new int[sourceRows]; }
        t.SourceRows = new int[sourceRows];
        HashSet<string> distinct = validation.Unique ? null : new HashSet<string>(StringComparer.Ordinal);
        int fixedLength = -1;
        for (int source = 0; source < sourceRows; source++)
        {
            int rs = st[source + 1];
            int re = en[source + 1];
            int row = physicalRows[source + 1];              // the row number a person sees (header = 1)
            // one pass over the row: count the columns, refuse a quoted field,
            // and find the key
            int field = 0;
            int p = rs;
            int keyAt = -1;
            int keyEnd = -1;
            bool controlKey = false;
            int controlCode = -1;
            while (true)
            {
                if (p < re && b[p] == (byte)'"') { throw new Rdv3DataError(Fmt(Rdv3Text.DataQuoted, file, row)); }
                int q = p;
                while (q < re && b[q] != sep)
                {
                    if (b[q] < 0x20)
                    {
                        if (field == t.KeyCol) { controlKey = true; }
                        if (controlCode < 0) { controlCode = (int)b[q]; }
                        if (field != t.KeyCol) { b[q] = (byte)'?'; }
                    }
                    q++;
                }
                if (field == t.KeyCol) { keyAt = p; keyEnd = q; }
                field++;
                if (q >= re) { break; }
                p = q + 1;
            }
            if (field < cols) { t.InputCounts.Shape(path, row, field, cols, false); continue; }
            if (field > cols)
            {
                throw Rdv3Input.Error(path, row, (Math.Min(field, cols) + 1).ToString(CultureInfo.InvariantCulture),
                    Rdv3Text.InputColumnCount.Replace("{n}", cols.ToString(CultureInfo.InvariantCulture)),
                    Rdv3Text.InputColumnCount.Replace("{n}", field.ToString(CultureInfo.InvariantCulture)), Rdv3Text.InputFixCsv);
            }
            if (controlKey)
            {
                t.InputCounts.Exclude(path, row, Rdv3Text.Format(Rdv3Text.RecordControlKey,
                    keyName, Rdv3Input.Display(enc.GetString(b, keyAt, keyEnd - keyAt))));
                continue;
            }
            if (controlCode >= 0 && t.ControlCharacterWarning.Length == 0) { t.ControlCharacterWarning = ControlChar(file, row, controlCode); }
            while (keyEnd > keyAt && b[keyEnd - 1] == (byte)' ') { keyEnd--; }
            for (int k = keyAt; k < keyEnd; k++)
            {
                // Unicode digits occupy several bytes. Validate their normalized
                // characters, rather than guessing widths from the encoded bytes.
                if (b[k] > 127) { return ReadDecoded(path, name, enc, new string[] { keyName }, validation, false, encodingSetting, references, headerRow, delimiter); }
            }
            int klen = keyEnd - keyAt;
            if (klen <= 0)
            {
                t.SkippedEmptyRows++;
                t.InputCounts.RowWarnings.Add(Rdv3Text.Format(Rdv3Text.RecordExcluded,
                    Rdv3Text.Format(Rdv3Text.SourceRow, file, row), Rdv3Text.Format(Rdv3Text.RecordEmptyKey, keyName)));
                continue;
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
                    t.InputCounts.Exclude(path, row, t.KeyError(row, enc.GetString(b, keyAt, klen),
                        Rdv3Text.InputExpectWidth.Replace("{n}", fixedLength.ToString(CultureInfo.InvariantCulture)), "length", "variable").Message);
                    continue;
                }
            }
            if (distinct != null && !distinct.Add(key)) { t.DuplicateWarning(key, new List<int> { row }, false); continue; }

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
        t.RemoveIdenticalRows();
        return t;
    }

    public static Rdv3Table Read(string path, string name, Encoding enc, string[] keyNames,
                                 Rdv3KeyValidation validation, string encodingSetting = "data.encoding",
                                 HashSet<string> references = null, int headerRow = 1, char delimiter = ',',
                                 string sheet = "")
    {
        if (keyNames == null || keyNames.Length == 0) { throw new ArgumentException("key names are empty"); }
        if (keyNames.Length == 1) { return Read(path, name, enc, keyNames[0], validation, encodingSetting, references, headerRow, delimiter, sheet); }
        if (references != null) { references = new HashSet<string>(references, StringComparer.Ordinal); references.UnionWith(keyNames); }
        return ReadDecoded(path, name, enc, keyNames, validation ?? new Rdv3KeyValidation(),
            string.Equals(System.IO.Path.GetExtension(path), ".xlsx", StringComparison.OrdinalIgnoreCase), encodingSetting, references, headerRow, delimiter, sheet);
    }

    private static Rdv3Table ReadDecoded(string path, string name, Encoding enc, string[] keyNames,
                                          Rdv3KeyValidation validation, bool workbook, string encodingSetting,
                                          HashSet<string> references, int headerRow, char delimiter,
                                          string sheet = "")
    {
        Rdv3Table t = new Rdv3Table();
        t.Name = name;
        t.Path = path;
        t.Enc = enc;
        t.Delimiter = delimiter;
        t.KeyValidation = validation;
        string warning = "";
        int[] originalRows = null;
        if (workbook)
        {
            Rdv3Xlsx.ReadTable(path, out t.Head, out t.Cells, out warning, references, t.InputCounts, headerRow, sheet);
            originalRows = t.InputCounts.SourceRows.ToArray();
        }
        else { Rdv3Csv.Read(path, enc, false, out t.Head, out t.Cells, out originalRows, encodingSetting, references, t.InputCounts, headerRow, delimiter); }
        t.ControlCharacterWarning = warning;
        t.KeyCols = new int[keyNames.Length];
        for (int k = 0; k < keyNames.Length; k++)
        {
            t.KeyCols[k] = t.ColumnOf(keyNames[k]);
            if (t.KeyCols[k] < 0)
            { throw new Rdv3DataError(Fmt(Rdv3Text.DataNoColumn, System.IO.Path.GetFileName(path), 1).Replace("{name}", keyNames[k])); }
        }
        t.KeyCol = t.KeyCols[0];
        string[][] source = t.Cells;
        List<string[]> kept = new List<string[]>(source.Length);
        List<int> sourceRows = new List<int>(source.Length);
        HashSet<string> distinct = validation.Unique ? null : new HashSet<string>(StringComparer.Ordinal);
        int[] fixedLengths = new int[keyNames.Length];
        for (int k = 0; k < fixedLengths.Length; k++) { fixedLengths[k] = -1; }
        string file = System.IO.Path.GetFileName(path);
        for (int i = 0; i < source.Length; i++)
        {
            for (int c = 0; c < source[i].Length; c++) { source[i][c] = Rdv3Input.Cell(source[i][c]); }
            int row = originalRows == null ? i + 2 : originalRows[i];
            bool empty = false;
            foreach (int col in t.KeyCols) { if (source[i][col].Length == 0) { empty = true; break; } }
            if (empty)
            {
                t.SkippedEmptyRows++;
                t.InputCounts.RowWarnings.Add(Rdv3Text.Format(Rdv3Text.RecordExcluded,
                    Rdv3Text.Format(Rdv3Text.SourceRow, file, row), Rdv3Text.Format(Rdv3Text.RecordEmptyKey, t.KeyLabel)));
                continue;
            }
            int[] rowLengths = (int[])fixedLengths.Clone();
            try
            {
            for (int part = 0; part < t.KeyCols.Length; part++)
            {
                int col = t.KeyCols[part];
                string value = source[i][col];
                if (value.Length == 0)
                {
                    throw t.KeyError(row, "", Rdv3Text.InputExpectKey, "empty", "skip", col);
                }
                for (int k = 0; k < value.Length; k++)
                {
                    if (value[k] < ' ')
                    { throw Rdv3Input.Error(path, row, t.Head[col], Rdv3Text.InputExpectKey, value, Rdv3Text.InputFixCsv); }
                    if (validation.Ascii && value[k] > 127)
                    { throw t.KeyError(row, value, "ASCII", "characters", "unicode", col); }
                }
                if (validation.FixedLength)
                {
                    if (rowLengths[part] < 0) { rowLengths[part] = value.Length; }
                    else if (value.Length != rowLengths[part])
                    { throw t.KeyError(row, value, Rdv3Text.InputExpectWidth.Replace("{n}", rowLengths[part].ToString(CultureInfo.InvariantCulture)), "length", "variable", col); }
                }
            }
            }
            catch (Rdv3DataError error) { t.InputCounts.Exclude(path, row, error.Message); continue; }
            fixedLengths = rowLengths;
            for (int c = 0; c < source[i].Length; c++)
            {
                string value = source[i][c];
                char[] safe = null;
                for (int k = 0; k < value.Length; k++)
                {
                    if (value[k] >= ' ' || value[k] == '\r' || value[k] == '\n') { continue; }
                    if (t.ControlCharacterWarning.Length == 0) { t.ControlCharacterWarning = ControlChar(file, row, (int)value[k]); }
                    if (safe == null) { safe = value.ToCharArray(); }
                    safe[k] = '?';
                }
                if (safe != null) { source[i][c] = new string(safe); }
            }
            string key = Rdv3Key.FromCells(source[i], t.KeyCols);
            if (distinct != null && !distinct.Add(key)) { t.DuplicateWarning(key, new List<int> { row }, false); continue; }
            kept.Add(source[i]);
            if (sourceRows != null) { sourceRows.Add(row); }
        }
        t.Cells = kept.ToArray();
        t.Rows = t.Cells.Length;
        t.KeyAt = new int[t.Rows];
        t.SourceRows = (sourceRows == null) ? null : sourceRows.ToArray();
        t.KeyLen = keyNames.Length == 1 && validation.UsesFixedAsciiPath && fixedLengths[0] > 0 ? fixedLengths[0] : 0;
        t.RemoveIdenticalRows();
        return t;
    }

    // A workbook key column converted to its declared date format has new key
    // strings. The character and width rules are checked again on the converted
    // values, and the fixed width the byte index relies on is taken from them:
    // a width left over from the serial numbers makes every fast-path lookup miss.
    public void RevalidateKeys()
    {
        if (Cells == null || KeyCols == null) { return; }
        int[] fixedLengths = new int[KeyCols.Length];
        for (int k = 0; k < fixedLengths.Length; k++) { fixedLengths[k] = -1; }
        HashSet<int> rejected = new HashSet<int>();
        for (int i = 0; i < Rows; i++)
        {
            int row = SourceRow(i);
            int[] rowLengths = (int[])fixedLengths.Clone();
            try
            {
            for (int part = 0; part < KeyCols.Length; part++)
            {
                int col = KeyCols[part];
                string value = Cells[i][col];
                if (value.Length == 0) { throw KeyError(row, "", Rdv3Text.InputExpectKey, "empty", "skip", col); }
                if (KeyValidation.Ascii)
                {
                    for (int k = 0; k < value.Length; k++)
                    { if (value[k] > 127) { throw KeyError(row, value, "ASCII", "characters", "unicode", col); } }
                }
                if (KeyValidation.FixedLength)
                {
                    if (rowLengths[part] < 0) { rowLengths[part] = value.Length; }
                    else if (value.Length != rowLengths[part])
                    {
                        throw KeyError(row, value, Rdv3Text.InputExpectWidth.Replace("{n}",
                            rowLengths[part].ToString(CultureInfo.InvariantCulture)), "length", "variable", col);
                    }
                }
            }
            }
            catch (Rdv3DataError error) { InputCounts.Exclude(Path, row, error.Message); rejected.Add(i); continue; }
            fixedLengths = rowLengths;
        }
        RemoveRows(rejected);
        KeyLen = KeyCols.Length == 1 && KeyValidation.UsesFixedAsciiPath && fixedLengths[0] > 0 ? fixedLengths[0] : 0;
        RemoveIdenticalRows();
    }

    private void RemoveIdenticalRows()
    {
        if (!KeyValidation.Unique) { return; }
        Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>(Rows, StringComparer.Ordinal);
        for (int row = 0; row < Rows; row++)
        {
            string key = Key(row);
            List<int> group;
            if (!groups.TryGetValue(key, out group)) { group = new List<int>(); groups.Add(key, group); }
            group.Add(row);
        }
        HashSet<int> rejected = new HashSet<int>();
        foreach (KeyValuePair<string, List<int>> pair in groups)
        {
            List<int> group = pair.Value;
            if (group.Count < 2) { continue; }
            bool conflict = false;
            for (int i = 1; i < group.Count && !conflict; i++)
            {
                for (int c = 0; c < Head.Length; c++)
                { if (Field(group[0], c) != Field(group[i], c)) { conflict = true; break; } }
            }
            List<int> numbers = new List<int>();
            // Decide on the whole group before compacting: A,A,B must exclude
            // all three, including the earlier identical retransmission.
            for (int i = conflict ? 0 : 1; i < group.Count; i++)
            { rejected.Add(group[i]); numbers.Add(SourceRow(group[i])); }
            DuplicateWarning(pair.Key, numbers, conflict);
        }
        RemoveRows(rejected);
    }

    private void DuplicateWarning(string key, List<int> rows, bool conflict)
    {
        SkippedDuplicateRows += rows.Count;
        List<string> numbers = new List<string>();
        foreach (int row in rows) { numbers.Add(row.ToString(CultureInfo.InvariantCulture)); }
        InputCounts.RowWarnings.Add(Rdv3Text.Format(Rdv3Text.RecordDuplicate,
            System.IO.Path.GetFileName(Path), KeyLabel, Rdv3Input.Display(key), rows.Count,
            string.Join(", ", numbers.ToArray()), conflict ? Rdv3Text.RecordConflict : Rdv3Text.RecordIdentical));
    }

    public void RemoveRows(HashSet<int> rejected)
    {
        if (rejected.Count == 0) { return; }
        int kept = 0;
        for (int row = 0; row < Rows; row++)
        {
            if (rejected.Contains(row)) { continue; }
            if (Cells != null) { Cells[kept] = Cells[row]; }
            else { Start[kept] = Start[row]; End[kept] = End[row]; KeyAt[kept] = KeyAt[row]; }
            if (KeyLengths != null) { KeyLengths[kept] = KeyLengths[row]; }
            if (SourceRows != null) { SourceRows[kept] = SourceRows[row]; }
            kept++;
        }
        Rows = kept;
        if (Cells != null) { Array.Resize(ref Cells, kept); }
        else { Array.Resize(ref Start, kept); Array.Resize(ref End, kept); Array.Resize(ref KeyAt, kept); }
        if (KeyLengths != null) { Array.Resize(ref KeyLengths, kept); }
        if (SourceRows != null) { Array.Resize(ref SourceRows, kept); }
    }

    public string InputNotice()
    {
        if (Rows == 0 && SkippedEmptyRows == 0 && SkippedDuplicateRows == 0)
        { return (InputCounts.ShortRows + InputCounts.BlankRows + InputCounts.InvalidRows == 0 ? Rdv3Text.InputNoData : Rdv3Text.InputNoKeptData)
            .Replace("{file}", System.IO.Path.GetFileName(Path)); }
        if (SkippedEmptyRows == 0 && SkippedDuplicateRows == 0) { return ""; }
        return Rdv3Text.InputRowsSkipped.Replace("{file}", System.IO.Path.GetFileName(Path))
            .Replace("{column}", KeyLabel)
            .Replace("{empty}", SkippedEmptyRows.ToString(CultureInfo.InvariantCulture))
            .Replace("{duplicate}", SkippedDuplicateRows.ToString(CultureInfo.InvariantCulture))
            .Replace("{kept}", Rows.ToString(CultureInfo.InvariantCulture));
    }

    public void AddWarnings(List<string> warnings)
    {
        InputCounts.AddWarnings(Path, warnings);
        if (ControlCharacterWarning.Length > 0) { warnings.Add(ControlCharacterWarning); }
        string notice = InputNotice();
        if (notice.Length > 0) { warnings.Add(notice); }
    }

    private Rdv3DataError KeyError(int row, string actual, string expected, string rule, string choice, int column = -1)
    {
        string path = KeyValidation.SettingsPath.Length == 0 ? "data.tables." + Name + ".keyValidation" : KeyValidation.SettingsPath;
        return Rdv3Input.Error(Path, row, Head[column < 0 ? KeyCol : column], expected, actual,
            Rdv3Text.InputFixKey.Replace("{path}", path + "." + rule).Replace("{choice}", choice));
    }

    // A tab, a carriage return or any other control character, named by code.
    private static string ControlChar(string file, int row, int code)
    {
        return Fmt(Rdv3Text.DataControlChar, file, row).Replace("{code}", code.ToString("X2", CultureInfo.InvariantCulture));
    }

    // byte offset of field f in row i, -1 if the row has no such field
    public int FieldAt(int i, int f, out int len)
    {
        if (Cells != null) { len = 0; return -1; }
        int p = Start[i];
        int e = End[i];
        byte sep = (byte)Delimiter;
        for (int k = 0; k < f; k++)
        {
            while (p < e && Buf[p] != sep) { p++; }
            if (p >= e) { len = 0; return -1; }
            p++;
        }
        int q = p;
        while (q < e && Buf[q] != sep) { q++; }
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
        return Rdv3Input.Cell(Enc.GetString(Buf, p, len));
    }

    public string Key(int i)
    {
        if (Cells != null) { return Rdv3Key.FromCells(Cells[i], KeyCols); }
        int len = (KeyLengths == null) ? KeyLen : KeyLengths[i];
        return Enc.GetString(Buf, KeyAt[i], len);
    }

    public string KeyLabel
    {
        get
        {
            string[] names = new string[KeyCols.Length];
            for (int i = 0; i < names.Length; i++) { names[i] = Head[KeyCols[i]]; }
            return string.Join(" / ", names);
        }
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
