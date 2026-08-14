// ============================================================================
// Rdv2IdxHash.cs -- index method 1 of 2: hand-built, bucket-chained hash.
//
// Rdv2IdxDict.cs defines a class with the same name and the same three
// members. Exactly one of the two is packed into a given .cmd, so the call
// sites in Rdv2Core.cs are identical and nothing dispatches at run time.
//
// Structure: one bucket array over the raw file bytes plus a next-row array.
// A key that owns several rows simply has several rows on its chain, so
// one-to-many falls out of the design instead of being bolted on.
//
//   head[hash & mask]  first row in the bucket, +1 so 0 means empty
//   next[row]          next row in the same bucket, +1, 0 ends the chain
//
// Nothing is copied out of the file: a key is eight bytes at a known offset,
// and comparison is eight byte compares. There are no strings and no
// allocations per row, which is exactly the thing being compared against
// Dictionary<string, List<int>>.
// ============================================================================

using System;

public sealed class RdvIndex
{
    private readonly byte[] buf;
    private readonly int[] keyAt;
    private readonly int[] head;
    private readonly int[] next;
    private readonly int mask;
    private long probes;

    public static string Kind { get { return "hash"; } }

    public long Probes { get { return probes; } }

    public RdvIndex(Rdv2Table t)
    {
        buf = t.Buf;
        keyAt = t.KeyAt;
        int cap = Rdv2Spec.Capacity(t.Rows);
        mask = cap - 1;
        head = new int[cap];
        next = new int[t.Rows];

        // Insert at the head of the bucket. Rows of the same key end up in
        // reverse file order on the chain; Find hands them back as it walks and
        // the caller sorts, so the visible order is file order either way.
        for (int i = 0; i < t.Rows; i++)
        {
            int h = Rdv2Spec.Hash(buf, keyAt[i]) & mask;
            next[i] = head[h];
            head[h] = i + 1;
        }
    }

    // every row whose key equals the eight bytes at other[off].
    // Returns how many were found; fills result up to its length.
    public int Find(byte[] other, int off, int[] result)
    {
        int n = 0;
        int steps = 0;
        int r = head[Rdv2Spec.Hash(other, off) & mask];
        while (r != 0)
        {
            steps++;
            int row = r - 1;
            if (Rdv2Spec.SameKey(buf, keyAt[row], other, off))
            {
                if (n < result.Length) { result[n] = row; }
                n++;
            }
            r = next[row];
        }
        probes += steps;
        return (n > result.Length) ? result.Length : n;
    }
}
