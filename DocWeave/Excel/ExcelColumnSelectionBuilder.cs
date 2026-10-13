namespace DocWeave.Excel;

/// <summary>
/// Configures Excel import columns by worksheet letters and/or header captions.
/// </summary>
public sealed class ExcelColumnSelectionBuilder
{
    private readonly ColumnSelection selection;

    internal ExcelColumnSelectionBuilder(ColumnSelection selection)
    {
        this.selection = selection;
    }

    /// <summary>
    /// Includes columns by Excel letters, such as A, B, or AA.
    /// </summary>
    public ExcelColumnSelectionBuilder IncludeLetters(params string[] letters)
    {
        selection.IncludeLetters.AddRange(letters);
        return this;
    }

    /// <summary>
    /// Includes columns by header captions from the imported header row.
    /// </summary>
    public ExcelColumnSelectionBuilder IncludeHeaders(params string[] headers)
    {
        selection.IncludeHeaders.AddRange(headers);
        return this;
    }

    /// <summary>
    /// Excludes columns by Excel letters after includes have been resolved.
    /// </summary>
    public ExcelColumnSelectionBuilder ExcludeLetters(params string[] letters)
    {
        selection.ExcludeLetters.AddRange(letters);
        return this;
    }

    /// <summary>
    /// Excludes columns by header captions after includes have been resolved.
    /// </summary>
    public ExcelColumnSelectionBuilder ExcludeHeaders(params string[] headers)
    {
        selection.ExcludeHeaders.AddRange(headers);
        return this;
    }
}
