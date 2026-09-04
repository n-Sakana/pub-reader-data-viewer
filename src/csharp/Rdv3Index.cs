// ============================================================================
// Rdv3Index.cs -- the one index: the .NET standard Dictionary<string, List<int>>.
//
// A key owns a set of rows, so the value is a list and nothing is ever
// overwritten. Strict table keys refuse duplicates. A table configured as a
// condition-value set has already kept only the first row for each key. The
// ledger's search column is one-to-many by nature.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public sealed class Rdv3Index
{
    private readonly Dictionary<string, List<int>> map;
    private readonly string[] scanLines;
    private readonly int[] scanColumns;
    private readonly bool contains;
    // the width of the key this index was built on (0 for a string index)
    private readonly int keyLen;
    private readonly bool fixedAscii;
    private readonly Encoding keyEncoding;

    public int Keys { get { return (map == null) ? 0 : map.Count; } }
    public int KeyLen { get { return keyLen; } }

    // index a CSV table by its key column, at the width the table itself
    // declares; the key is unique, and a duplicate is a data error
    public Rdv3Index(Rdv3Table t)
    {
        scanLines = null;
        scanColumns = null;
        contains = false;
        fixedAscii = t.KeyValidation.UsesFixedAsciiPath;
        keyLen = fixedAscii ? t.KeyLen : 0;
        keyEncoding = t.Enc;
        map = new Dictionary<string, List<int>>(t.Rows, StringComparer.Ordinal);
        byte[] b = t.Buf;
        for (int i = 0; i < t.Rows; i++)
        {
            string k = (fixedAscii && t.Cells == null)
                ? Encoding.ASCII.GetString(b, t.KeyAt[i], keyLen)
                : t.Key(i);
            List<int> rows;
            if (map.TryGetValue(k, out rows))
            {
                if (t.KeyValidation.Unique)
                {
                    throw new Rdv3DataError(Rdv3Text.DataDupKey
                        .Replace("{file}", System.IO.Path.GetFileName(t.Path))
                        .Replace("{name}", t.Head[t.KeyCol])
                        .Replace("{key}", k)
                        .Replace("{row1}", t.SourceRow(rows[0]).ToString(CultureInfo.InvariantCulture))
                        .Replace("{row2}", t.SourceRow(i).ToString(CultureInfo.InvariantCulture)));
                }
                continue;
            }
            rows = new List<int>(1);
            rows.Add(i);
            map.Add(k, rows);
        }
    }

    // index arbitrary row keys (the ledger's search column: one key, many rows)
    public Rdv3Index(string[] keys)
    {
        scanLines = null;
        scanColumns = null;
        contains = false;
        fixedAscii = false;
        keyEncoding = null;
        map = new Dictionary<string, List<int>>(keys.Length, StringComparer.Ordinal);
        for (int i = 0; i < keys.Length; i++)
        {
            List<int> rows;
            if (!map.TryGetValue(keys[i], out rows))
            {
                rows = new List<int>(1);
                map.Add(keys[i], rows);
            }
            rows.Add(i);
        }
    }

    // A ledger search can name several columns. Exact matching keeps the
    // dictionary path; contains matching deliberately scans the configured
    // columns and is measured by the caller on real data.
    public Rdv3Index(string[] lines, int[] columns, string match)
    {
        keyLen = 0;
        fixedAscii = false;
        keyEncoding = null;
        scanLines = lines ?? new string[0];
        scanColumns = columns ?? new int[0];
        contains = string.Equals(match, "contains", StringComparison.Ordinal);
        if (contains)
        {
            map = null;
            return;
        }
        map = new Dictionary<string, List<int>>(scanLines.Length, StringComparer.Ordinal);
        for (int row = 0; row < scanLines.Length; row++)
        {
            List<string> seen = new List<string>(scanColumns.Length);
            for (int c = 0; c < scanColumns.Length; c++)
            {
                string value = Rdv3Ledger.FieldOf(scanLines[row], scanColumns[c]);
                if (seen.Contains(value)) { continue; }
                seen.Add(value);
                List<int> rows;
                if (!map.TryGetValue(value, out rows))
                {
                    rows = new List<int>(1);
                    map.Add(value, rows);
                }
                rows.Add(row);
            }
        }
    }

    public List<int> Find(string key)
    {
        if (contains)
        {
            List<int> hits = new List<int>();
            for (int row = 0; row < scanLines.Length; row++)
            {
                for (int c = 0; c < scanColumns.Length; c++)
                {
                    string value = Rdv3Ledger.FieldOf(scanLines[row], scanColumns[c]);
                    if (value.IndexOf(key, StringComparison.Ordinal) < 0) { continue; }
                    hits.Add(row);
                    break;
                }
            }
            return (hits.Count == 0) ? null : hits;
        }
        List<int> rows;
        if (!map.TryGetValue(key, out rows)) { return null; }
        return rows;
    }

    // Fixed ASCII tables retain the width check and ASCII decoder. Relaxed
    // tables decode the actual slice with the declared data encoding.
    public int FindBytes(byte[] buf, int off, int len, out List<int> rows)
    {
        if (fixedAscii && len != keyLen) { rows = null; return 0; }
        if (!fixedAscii && keyEncoding == null) { rows = null; return 0; }
        string k = fixedAscii ? Encoding.ASCII.GetString(buf, off, len) : keyEncoding.GetString(buf, off, len);
        if (!map.TryGetValue(k, out rows)) { return 0; }
        return rows.Count;
    }
}
