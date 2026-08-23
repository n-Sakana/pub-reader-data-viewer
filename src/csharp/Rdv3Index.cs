// ============================================================================
// Rdv3Index.cs -- the ONE index of the practical build: the .NET standard
// Dictionary<string, List<int>>.
//
// Same structure the archived one-to-many comparison measured as "C# Dict"
// (archive\comparisons\one-to-many\src\csharp\Rdv2IdxDict.cs). The hand-built
// hash measured next to it is evidence only and is deliberately not in this build:
// there is no second method here and no fallback.
//
// A key owns a set of rows, so the value is a list and nothing is ever
// overwritten. A TABLE's key column is declared unique by the definition, so
// a table index refuses a duplicate (two rows for one key would make the join
// pick one of them in silence). The ledger's search column is one-to-many by
// nature and is indexed without that check.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

public sealed class Rdv3Index
{
    private readonly Dictionary<string, List<int>> map;
    // the width of the key this index was built on (0 for a string index)
    private readonly int keyLen;

    public int Keys { get { return map.Count; } }
    public int KeyLen { get { return keyLen; } }

    // index a CSV table by its key column, at the width the table itself
    // declares; the key is unique, and a duplicate is a data error
    public Rdv3Index(Rdv3Table t)
    {
        keyLen = t.KeyLen;
        map = new Dictionary<string, List<int>>(t.Rows, StringComparer.Ordinal);
        byte[] b = t.Buf;
        for (int i = 0; i < t.Rows; i++)
        {
            string k = Encoding.ASCII.GetString(b, t.KeyAt[i], keyLen);
            List<int> rows;
            if (map.TryGetValue(k, out rows))
            {
                throw new Rdv3DataError(Rdv3Text.DataDupKey
                    .Replace("{file}", System.IO.Path.GetFileName(t.Path))
                    .Replace("{name}", t.Head[t.KeyCol])
                    .Replace("{key}", k)
                    .Replace("{row1}", (rows[0] + 2).ToString(CultureInfo.InvariantCulture))
                    .Replace("{row2}", (i + 2).ToString(CultureInfo.InvariantCulture)));
            }
            rows = new List<int>(1);
            rows.Add(i);
            map.Add(k, rows);
        }
    }

    // index arbitrary row keys (the ledger's search column: one key, many rows)
    public Rdv3Index(string[] keys)
    {
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

    public List<int> Find(string key)
    {
        List<int> rows;
        if (!map.TryGetValue(key, out rows)) { return null; }
        return rows;
    }

    // look up `len` bytes at `off`: a key of another width can never match
    public int FindBytes(byte[] buf, int off, int len, out List<int> rows)
    {
        if (len != keyLen) { rows = null; return 0; }
        string k = Encoding.ASCII.GetString(buf, off, len);
        if (!map.TryGetValue(k, out rows)) { return 0; }
        return rows.Count;
    }
}
