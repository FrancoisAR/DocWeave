using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

/// <summary>
/// Turns caller-supplied names into names Excel accepts for sheets, pivot tables and native tables.
/// </summary>
internal static class ExcelNames
{
    internal static string SanitizeSheetName(string value)
    {
        var invalid = new[] { ':', '\\', '/', '?', '*', '[', ']' };
        var safe = new string(value.Select(character => invalid.Contains(character) ? '_' : character).ToArray()).Trim();
        return string.IsNullOrWhiteSpace(safe) ? "Sheet" : safe[..Math.Min(31, safe.Length)];
    }

    internal static string SanitizePivotName(string value)
    {
        var safe = new string(value.Select(character => char.IsLetterOrDigit(character) || character == '_' ? character : '_').ToArray()).Trim('_');
        return string.IsNullOrWhiteSpace(safe) ? "PivotTable" : safe[..Math.Min(255, safe.Length)];
    }

    internal static string SanitizeTableName(string value)
    {
        var safe = new string(value.Select(character => char.IsLetterOrDigit(character) || character == '_' ? character : '_').ToArray()).Trim('_');
        if (string.IsNullOrWhiteSpace(safe))
        {
            safe = "Table";
        }

        if (char.IsDigit(safe[0]))
        {
            safe = $"Table_{safe}";
        }

        return safe[..Math.Min(255, safe.Length)];
    }
}
