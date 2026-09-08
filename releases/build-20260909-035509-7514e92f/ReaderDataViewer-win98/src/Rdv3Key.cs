using System;
using System.Globalization;
using System.Text;

// Single-column identities retain their stored representation. A tuple uses
// lengths, not a separator that might also occur in a source value.
public static class Rdv3Key
{
    public static string FromCells(string[] row, int[] columns)
    {
        if (columns.Length == 1) { return row[columns[0]]; }
        string[] values = new string[columns.Length];
        for (int i = 0; i < columns.Length; i++) { values[i] = row[columns[i]]; }
        return Pack(values);
    }

    public static string FromLine(string line, int[] columns)
    {
        if (columns.Length == 1) { return Rdv3Ledger.FieldOf(line, columns[0]); }
        string[] values = new string[columns.Length];
        for (int i = 0; i < columns.Length; i++) { values[i] = Rdv3Ledger.FieldOf(line, columns[i]); }
        return Pack(values);
    }

    private static string Pack(string[] values)
    {
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < values.Length; i++)
        {
            if (string.IsNullOrEmpty(values[i])) { return ""; }
            result.Append(values[i].Length.ToString(CultureInfo.InvariantCulture)).Append(':').Append(values[i]);
        }
        return result.ToString();
    }
}
