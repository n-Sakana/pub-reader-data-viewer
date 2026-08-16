// ============================================================================
// Rdv3Index.cs -- the ONE index of the practical build: the .NET standard
// Dictionary<string, List<int>>.
//
// Same structure the one-to-many comparison measured as "C# Dict"
// (src\v2\csharp\Rdv2IdxDict.cs). The hand-built hash that was measured next
// to it is comparison evidence only and is deliberately NOT in this build:
// there is no second method here and no fallback.
//
// A key owns a set of rows -- key 1 is legitimately one-to-many -- so the value
// is a list and nothing is ever overwritten.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Text;

public sealed class Rdv3Index
{
    private readonly Dictionary<string, List<int>> map;

    public int Keys { get { return map.Count; } }

    // index a CSV table by its key column (field 0), whose width is the one key
    // length the whole program uses (Rdv3Spec.KeyLength, settled before the CSVs
    // were read)
    public Rdv3Index(Rdv3Table t)
    {
        map = new Dictionary<string, List<int>>(t.Rows, StringComparer.Ordinal);
        byte[] b = t.Buf;
        for (int i = 0; i < t.Rows; i++)
        {
            string k = Encoding.ASCII.GetString(b, t.KeyAt[i], Rdv3Spec.KeyLength);
            List<int> rows;
            if (!map.TryGetValue(k, out rows))
            {
                rows = new List<int>(1);
                map.Add(k, rows);
            }
            rows.Add(i);
        }
    }

    // index arbitrary row keys (used for the ledger's key1 search index)
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

    public int FindBytes(byte[] buf, int off, out List<int> rows)
    {
        string k = Encoding.ASCII.GetString(buf, off, Rdv3Spec.KeyLength);
        if (!map.TryGetValue(k, out rows)) { return 0; }
        return rows.Count;
    }
}
