using System.Text.RegularExpressions;

namespace DocWeave.Excel;

internal sealed partial record ExcelCellAddress(int Row, int Column)
{
    public static ExcelCellAddress Parse(string cellReference)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(cellReference);

        var match = CellReferenceRegex().Match(cellReference.Trim());
        if (!match.Success)
        {
            throw new ArgumentException($"Invalid Excel cell reference '{cellReference}'.", nameof(cellReference));
        }

        return new ExcelCellAddress(
            int.Parse(match.Groups["row"].Value),
            ColumnLetterToNumber(match.Groups["column"].Value));
    }

    /// <summary>
    /// Parses a rectangular range such as A1:D5. Absolute markers (for example $A$1:$D$5) are accepted.
    /// </summary>
    public static (ExcelCellAddress Start, ExcelCellAddress End) ParseRange(string range)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(range);

        var parts = range.Split(':', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 2)
        {
            throw new ArgumentException("Range must be in A1:D5 format.", nameof(range));
        }

        return (Parse(parts[0].Replace("$", string.Empty, StringComparison.Ordinal)), Parse(parts[1].Replace("$", string.Empty, StringComparison.Ordinal)));
    }

    public string ToReference()
    {
        return $"{ColumnNumberToLetter(Column)}{Row}";
    }

    public static string ColumnNumberToLetter(int column)
    {
        return ColumnLetters.FromNumber(column);
    }

    public static int ColumnLetterToNumber(string letters)
    {
        return ColumnLetters.ToNumber(letters);
    }

    public ExcelCellAddress Offset(int rows, int columns)
    {
        return new ExcelCellAddress(Row + rows, Column + columns);
    }

    [GeneratedRegex("^(?<column>[A-Za-z]+)(?<row>[1-9][0-9]*)$")]
    private static partial Regex CellReferenceRegex();
}
