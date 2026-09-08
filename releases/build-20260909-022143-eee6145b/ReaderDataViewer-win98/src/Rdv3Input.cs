using System;
using System.Globalization;
using System.Text;

// Normalize imported cells and lookup input, never persisted ledger/pending
// snapshots: changing those would invalidate their conflict baselines.
public static class Rdv3Input
{
    public static void ValidateEncoding(byte[] bytes, Encoding encoding, string path)
    {
        Encoding strict = (Encoding)encoding.Clone();
        strict.DecoderFallback = DecoderFallback.ExceptionFallback;
        try { strict.GetCharCount(bytes); }
        catch (DecoderFallbackException ex)
        {
            int stop = Math.Max(0, Math.Min(ex.Index, bytes.Length));
            string prefix = encoding.GetString(bytes, 0, stop);
            int row = 1, column = 1;
            bool quoted = false;
            for (int i = 0; i < prefix.Length; i++)
            {
                char c = prefix[i];
                if (c == '"') { quoted = !quoted; }
                else if (c == ',' && !quoted) { column++; }
                else if (c == '\r' || c == '\n')
                {
                    if (c == '\r' && i + 1 < prefix.Length && prefix[i + 1] == '\n') { i++; }
                    row++;
                    if (!quoted) { column = 1; }
                }
            }
            throw Error(path, row, column.ToString(CultureInfo.InvariantCulture), encoding.WebName,
                "bytes " + BitConverter.ToString(ex.BytesUnknown), Rdv3Text.InputFixUnknownEncoding);
        }
    }

    public static Rdv3DataError Error(string file, int row, string column, string expected, string actual, string fix)
    {
        return new Rdv3DataError(Rdv3Text.InputError.Replace("{file}", System.IO.Path.GetFileName(file))
            .Replace("{row}", row.ToString(CultureInfo.InvariantCulture)).Replace("{column}", column)
            .Replace("{expected}", expected).Replace("{actual}", actual).Replace("{fix}", fix));
    }

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
