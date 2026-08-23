// ============================================================================
// Rdv2IdxDict.cs -- index method 2 of 2: the .NET standard associative array.
//
// Rdv2IdxHash.cs defines a class with the same name and the same three
// members. Exactly one of the two is packed into a given .cmd.
//
// Dictionary<string, List<int>>: the ordinary thing to write. A key that owns
// several rows has several entries in its list, so one-to-many is the natural
// shape here too.
//
// What it costs, and what the comparison is really about: a key has to become
// a string. Every row allocates one on the way in, and every probe allocates
// one on the way out, because the caller only ever has bytes at an offset.
// That is not a handicap invented for the benchmark -- it is what using the
// stock dictionary on a byte buffer actually costs, and the hand-built index
// exists to show the size of it.
// ============================================================================

using System;
using System.Collections.Generic;
using System.Text;

public sealed class RdvIndex
{
    private readonly Dictionary<string, List<int>> map;
    private long probes;

    public static string Kind { get { return "dict"; } }

    public long Probes { get { return probes; } }

    public RdvIndex(Rdv2Table t)
    {
        map = new Dictionary<string, List<int>>(t.Rows, StringComparer.Ordinal);
        byte[] b = t.Buf;
        for (int i = 0; i < t.Rows; i++)
        {
            string k = Encoding.ASCII.GetString(b, t.KeyAt[i], Rdv2Spec.KeyLen);
            List<int> rows;
            if (!map.TryGetValue(k, out rows))
            {
                rows = new List<int>(1);
                map.Add(k, rows);
            }
            rows.Add(i);
        }
    }

    public int Find(byte[] other, int off, int[] result)
    {
        string k = Encoding.ASCII.GetString(other, off, Rdv2Spec.KeyLen);
        List<int> rows;
        if (!map.TryGetValue(k, out rows)) { probes++; return 0; }
        probes += rows.Count;
        int n = rows.Count;
        if (n > result.Length) { n = result.Length; }
        for (int i = 0; i < n; i++) { result[i] = rows[i]; }
        return n;
    }
}
