using System;
using System.Globalization;

// Normalize imported cells and lookup input, never persisted ledger/pending
// snapshots: changing those would invalidate their conflict baselines.
public static class Rdv3Input
{
    public static bool IsPadding(char c)
    { return c == ' ' || c == '\u3000' || c == '\u00a0'; }

    public static string Cell(string value)
    {
        if (string.IsNullOrEmpty(value)) { return value ?? ""; }
        int end = value.Length;
        while (end > 0 && IsPadding(value[end - 1])) { end--; }
        char[] changed = null;
        for (int i = 0; i < end; i++)
        {
            char c = value[i];
            if (c < '\uff10' || c > '\uff19') { continue; }
            if (changed == null) { changed = value.Substring(0, end).ToCharArray(); }
            changed[i] = (char)('0' + c - '\uff10');
        }
        return changed != null ? new string(changed) : (end == value.Length ? value : value.Substring(0, end));
    }

    public static bool TryNumber(string text, out decimal value)
    {
        text = Cell(text).Trim().Replace('\uff0c', ',').Replace('\uff0e', '.')
            .Replace('\uff0d', '-').Replace('\uff0b', '+').Replace('\uff08', '(').Replace('\uff09', ')');
        bool negative = text.Length >= 2 && text[0] == '(' && text[text.Length - 1] == ')';
        if (negative) { text = text.Substring(1, text.Length - 2).Trim(); }
        if (text.StartsWith("\u00a5", StringComparison.Ordinal) || text.StartsWith("\uffe5", StringComparison.Ordinal))
        { text = text.Substring(1).TrimStart(); }
        if (negative && (text.StartsWith("-", StringComparison.Ordinal) || text.StartsWith("+", StringComparison.Ordinal)))
        { value = 0; return false; }
        if (!decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out value)) { return false; }
        if (negative) { value = -value; }
        return true;
    }
}
