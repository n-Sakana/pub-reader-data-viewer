using System;
using System.Collections.Generic;

internal sealed class Rdv3ResetNotice
{
    private readonly int[] identityColumn;
    private readonly string initial;

    public Rdv3ResetNotice(int identityCol, string initialStored) : this(new int[] { identityCol }, initialStored) { }

    public Rdv3ResetNotice(int[] identityCol, string initialStored)
    {
        identityColumn = identityCol;
        initial = initialStored;
    }

    public List<string> ChangedRows(string[] beforeLines, string[] beforeStates, string[] afterLines, string[] afterStates)
    {
        List<string> result = new List<string>();
        if (beforeLines == null || beforeStates == null) { return result; }
        Dictionary<string, int> before = new Dictionary<string, int>(StringComparer.Ordinal);
        for (int i = 0; i < beforeLines.Length; i++)
        {
            string identity = Rdv3Key.FromLine(beforeLines[i], identityColumn);
            if (!before.ContainsKey(identity)) { before.Add(identity, i); }
        }
        for (int i = 0; i < afterLines.Length; i++)
        {
            string identity = Rdv3Key.FromLine(afterLines[i], identityColumn);
            int old;
            if (!before.TryGetValue(identity, out old)) { continue; }
            if (string.Equals(beforeLines[old], afterLines[i], StringComparison.Ordinal)) { continue; }
            if (string.Equals(beforeStates[old], initial, StringComparison.Ordinal)) { continue; }
            if (!string.Equals(afterStates[i], initial, StringComparison.Ordinal)) { continue; }
            result.Add(afterLines[i]);
        }
        return result;
    }

    public List<string> AfterUpdate(Rdv3UpdateResult update, string[] beforeLines, string[] beforeStates, string[] effectiveStates)
    {
        // The merger sees shared states; the screen also sees pending states.
        // Either may have been reset. Keep both sources and notify each identity once.
        Dictionary<string, string> rows = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (string line in update.ResetLines) { rows[Rdv3Key.FromLine(line, identityColumn)] = line; }
        foreach (string line in ChangedRows(beforeLines, beforeStates, update.Lines, effectiveStates))
        {
            string identity = Rdv3Key.FromLine(line, identityColumn);
            if (!rows.ContainsKey(identity)) { rows.Add(identity, line); }
        }
        List<string> identities = new List<string>(rows.Keys);
        identities.Sort(StringComparer.Ordinal);
        List<string> result = new List<string>();
        foreach (string identity in identities) { result.Add(rows[identity]); }
        return result;
    }
}
