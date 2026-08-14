using System.Globalization;
using ClosedXML.Excel;

namespace ExcelMergeFork.Core.Excel;

public static class CellText
{
    public static string From(object? value)
    {
        if (value is null || value is DBNull)
        {
            return "";
        }

        switch (value)
        {
            case string s:
                return s.Trim();
            case bool b:
                return b ? "True" : "False";
            case DateTime dt:
                return dt.ToString(CultureInfo.InvariantCulture).Trim();
            case sbyte or byte or short or ushort or int or uint or long or ulong:
                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            case float or double or decimal:
                return FormatNumber(Convert.ToDecimal(value, CultureInfo.InvariantCulture));
            case XLCellValue cell:
                return From(ToObject(cell));
            default:
                return Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? "";
        }
    }

    public static object? ToObject(XLCellValue value)
    {
        if (value.IsBlank)
        {
            return null;
        }

        if (value.IsBoolean)
        {
            return value.GetBoolean();
        }

        if (value.IsNumber)
        {
            var number = value.GetNumber();
            if (double.IsFinite(number) && Math.Abs(number - Math.Round(number)) < 1e-9)
            {
                return Convert.ToInt64(Math.Round(number));
            }

            return number;
        }

        if (value.IsDateTime)
        {
            return value.GetDateTime();
        }

        if (value.IsTimeSpan)
        {
            return value.GetTimeSpan();
        }

        if (value.IsText)
        {
            return value.GetText();
        }

        if (value.IsError)
        {
            return value.GetError().ToString();
        }

        return value.ToString();
    }

    public static bool RowsEqual(IReadOnlyList<object?>? left, IReadOnlyList<object?>? right)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        if (left is null || right is null)
        {
            return left == right;
        }

        var n = Math.Max(left.Count, right.Count);
        for (var i = 0; i < n; i++)
        {
            var a = i < left.Count ? From(left[i]) : "";
            var b = i < right.Count ? From(right[i]) : "";
            if (a != b)
            {
                return false;
            }
        }

        return true;
    }

    public static bool ColumnsEqual(IReadOnlyList<string>? left, IReadOnlyList<string>? right)
    {
        if (ReferenceEquals(left, right))
        {
            return true;
        }

        if (left is null || right is null)
        {
            return left == right;
        }

        var n = Math.Max(left.Count, right.Count);
        for (var i = 0; i < n; i++)
        {
            var a = i < left.Count ? left[i] : "";
            var b = i < right.Count ? right[i] : "";
            if (a != b)
            {
                return false;
            }
        }

        return true;
    }

    private static string FormatNumber(decimal value)
    {
        if (value == decimal.Truncate(value))
        {
            return decimal.Truncate(value).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToString(CultureInfo.InvariantCulture).Trim();
    }
}
