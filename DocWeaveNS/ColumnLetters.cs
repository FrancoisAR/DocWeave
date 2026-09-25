namespace DocWeave;

/// <summary>
/// Converts between spreadsheet column letters (A, B, ... Z, AA, ...) and one-based column numbers.
/// Shared by the Excel and CSV readers, which both let callers pick columns by letter.
/// </summary>
internal static class ColumnLetters
{
    /// <summary>
    /// Converts letters such as "A" or "AA" to a one-based column number (1 and 27).
    /// </summary>
    public static int ToNumber(string letters)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(letters, nameof(letters));

        var column = 0;
        foreach (var character in letters.Trim().ToUpperInvariant())
        {
            if (character is < 'A' or > 'Z')
            {
                throw new ArgumentException($"Invalid Excel column letters '{letters}'.", nameof(letters));
            }

            column = column * 26 + character - 'A' + 1;
        }

        return column;
    }

    /// <summary>
    /// Converts a one-based column number to letters (1 gives "A", 27 gives "AA").
    /// </summary>
    public static string FromNumber(int column)
    {
        if (column < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(column), "Column number must be one or greater.");
        }

        var value = column;
        var letters = string.Empty;
        while (value > 0)
        {
            value--;
            letters = (char)('A' + value % 26) + letters;
            value /= 26;
        }

        return letters;
    }
}
