using System.Globalization;
using System.Text.RegularExpressions;

namespace ExcelMergeFork.Core.Excel;

public static class KeyNormalizer
{
    private static readonly Regex LeadingZero = new(@"^0\d+$", RegexOptions.Compiled);
    private static readonly Regex IntegralText = new(@"^[+-]?(?:\d+)(?:\.0+)?$", RegexOptions.Compiled);
    private static readonly Regex DecimalText = new(@"^[+-]?\d+\.\d+$", RegexOptions.Compiled);

    public static string Normalize(object? value)
    {
        if (value is null)
        {
            return "";
        }

        switch (value)
        {
            case sbyte or byte or short or ushort or int or uint or long or ulong:
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            case float or double or decimal:
                return NormalizeDecimal(Convert.ToDecimal(value, CultureInfo.InvariantCulture));
        }

        var text = CellText.From(value);
        if (text.Length == 0)
        {
            return "";
        }

        if (LeadingZero.IsMatch(text))
        {
            return text;
        }

        if (IntegralText.IsMatch(text) &&
            decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var integral) &&
            integral == decimal.Truncate(integral))
        {
            return decimal.Truncate(integral).ToString(CultureInfo.InvariantCulture);
        }

        if (DecimalText.IsMatch(text) &&
            decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out var dec))
        {
            return NormalizeDecimal(dec);
        }

        return text;
    }

    public static string HeaderForCompare(string? value) => (value ?? "").Trim().ToLowerInvariant();

    public static string Prefix(string? key)
    {
        var text = (key ?? "").Trim();
        if (text.Length == 0)
        {
            return "";
        }

        foreach (var sep in new[] { "-", "_", " ", "\t" })
        {
            if (text.Contains(sep, StringComparison.Ordinal))
            {
                var part = text.Split(sep, 2)[0].Trim();
                return part.Length == 0 ? text : part;
            }
        }

        return text;
    }

    public static int InsertionIndex(IReadOnlyList<string> merged, string newKey)
    {
        var prefix = Prefix(newKey);
        var last = -1;
        for (var i = 0; i < merged.Count; i++)
        {
            if (Prefix(merged[i]) == prefix)
            {
                last = i;
            }
        }

        if (last >= 0)
        {
            return last + 1;
        }

        for (var i = merged.Count - 1; i >= 0; i--)
        {
            if (CommonPrefixLength(prefix, Prefix(merged[i])) >= AppConstants.FuzzyPrefixMinLen)
            {
                return i + 1;
            }
        }

        return merged.Count;
    }

    public static List<string> MergeOrdered(IEnumerable<string> baseOrdered, IEnumerable<string> newKeys)
    {
        var merged = baseOrdered.ToList();
        var seen = new HashSet<string>(merged);
        foreach (var key in newKeys)
        {
            if (!seen.Add(key))
            {
                continue;
            }

            merged.Insert(InsertionIndex(merged, key), key);
        }

        return merged;
    }

    private static string NormalizeDecimal(decimal value)
    {
        if (value == decimal.Truncate(value))
        {
            return decimal.Truncate(value).ToString(CultureInfo.InvariantCulture);
        }

        var text = value.ToString(CultureInfo.InvariantCulture);
        return text.TrimEnd('0').TrimEnd('.') is { Length: > 0 } trimmed ? trimmed : "0";
    }

    private static int CommonPrefixLength(string a, string b)
    {
        var n = Math.Min(a.Length, b.Length);
        var i = 0;
        while (i < n && a[i] == b[i])
        {
            i++;
        }

        return i;
    }
}
