namespace DocWeave.Csv;

/// <summary>
/// Chooses which CSV columns to read. Included columns come back in the order they were requested, and exclusions are applied after inclusions. A header that does not exist fails the read with an <see cref="InvalidOperationException"/>.
/// </summary>
public sealed class CsvColumnSelectionBuilder
{
    private readonly ColumnSelection selection;

    internal CsvColumnSelectionBuilder(ColumnSelection selection)
    {
        this.selection = selection;
    }

    /// <summary>
    /// Includes columns by Excel-style letters, such as A, B or AA.
    /// </summary>
    public CsvColumnSelectionBuilder IncludeLetters(params string[] letters)
    {
        selection.IncludeLetters.AddRange(letters);
        return this;
    }

    /// <summary>
    /// Includes columns by header. Headers are matched ignoring case.
    /// </summary>
    public CsvColumnSelectionBuilder IncludeHeaders(params string[] headers)
    {
        selection.IncludeHeaders.AddRange(headers);
        return this;
    }

    /// <summary>
    /// Excludes columns by Excel-style letters.
    /// </summary>
    public CsvColumnSelectionBuilder ExcludeLetters(params string[] letters)
    {
        selection.ExcludeLetters.AddRange(letters);
        return this;
    }

    /// <summary>
    /// Excludes columns by header.
    /// </summary>
    public CsvColumnSelectionBuilder ExcludeHeaders(params string[] headers)
    {
        selection.ExcludeHeaders.AddRange(headers);
        return this;
    }
}
