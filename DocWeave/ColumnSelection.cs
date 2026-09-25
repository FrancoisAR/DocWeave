namespace DocWeave;

/// <summary>
/// Which columns of a header row to keep, chosen by letter and/or header caption, with optional exclusions.
/// Shared by the Excel and CSV readers.
/// </summary>
internal sealed class ColumnSelection
{
    private readonly string sourceName;
    private readonly bool keepRequestedOrder;

    /// <param name="sourceName">Used in error messages, for example "CSV" or "Excel".</param>
    /// <param name="keepRequestedOrder">
    /// When true the selected columns come back in the order they were requested (IncludeHeaders("Status", "ID")
    /// gives Status first). When false they come back in source order.
    /// </param>
    public ColumnSelection(string sourceName, bool keepRequestedOrder)
    {
        this.sourceName = sourceName;
        this.keepRequestedOrder = keepRequestedOrder;
    }

    public List<string> IncludeLetters { get; } = [];
    public List<string> IncludeHeaders { get; } = [];
    public List<string> ExcludeLetters { get; } = [];
    public List<string> ExcludeHeaders { get; } = [];

    /// <summary>
    /// Returns the zero-based indexes of the selected columns. Nothing included means every column.
    /// </summary>
    public IReadOnlyList<int> Resolve(IReadOnlyList<string> headers)
    {
        // Includes come first: letters, then headers. With nothing included every column is kept.
        var indexes = new List<int>();

        if (IncludeLetters.Count == 0 && IncludeHeaders.Count == 0)
        {
            indexes.AddRange(Enumerable.Range(0, headers.Count));
        }
        else
        {
            indexes.AddRange(IncludeLetters.Select(letter => ColumnLetters.ToNumber(letter) - 1));
            indexes.AddRange(IncludeHeaders.Select(header => FindHeaderIndex(headers, header)));
        }

        // Exclusions are applied after inclusions and always win.
        var excluded = new HashSet<int>(ExcludeLetters.Select(letter => ColumnLetters.ToNumber(letter) - 1));
        foreach (var header in ExcludeHeaders)
        {
            excluded.Add(FindHeaderIndex(headers, header));
        }

        // A letter beyond the last column is ignored rather than treated as an error, because the same selection is often
        // reused on files with different numbers of columns.
        var selected = indexes
            .Where(index => index >= 0 && index < headers.Count)
            .Where(index => !excluded.Contains(index))
            .Distinct();

        return (keepRequestedOrder ? selected : selected.OrderBy(index => index)).ToList();
    }

    private int FindHeaderIndex(IReadOnlyList<string> headers, string header)
    {
        for (var i = 0; i < headers.Count; i++)
        {
            if (string.Equals(headers[i], header, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        throw new InvalidOperationException($"{sourceName} header '{header}' was not found.");
    }
}
