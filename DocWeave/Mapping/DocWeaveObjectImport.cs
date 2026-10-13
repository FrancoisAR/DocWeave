namespace DocWeave;

/// <summary>
/// One problem found while reading rows into objects.
/// </summary>
/// <param name="RowNumber">
/// The row in the source (the worksheet row number, or the CSV record number). Zero when the problem is with the columns as a whole,
/// for example a required column that is missing.
/// </param>
/// <param name="Column">The header of the column, or null when the problem is not about one column.</param>
/// <param name="Property">The name of the property that could not be filled.</param>
/// <param name="Value">The text that could not be used, or null when the cell was blank.</param>
/// <param name="Message">A description of the problem.</param>
public sealed record DocWeaveImportError(int RowNumber, string? Column, string? Property, string? Value, string Message);

/// <summary>
/// The objects read from a source, and the rows that could not be read.
/// </summary>
/// <typeparam name="T">The type of one object.</typeparam>
/// <param name="Items">The objects made from rows that had no errors, in source order.</param>
/// <param name="Errors">Every problem found. A row with an error contributes no object, and each of its problems is listed.</param>
/// <param name="RowCount">The number of data rows read, whether or not they produced an object.</param>
public sealed record DocWeaveObjectImportResult<T>(IReadOnlyList<T> Items, IReadOnlyList<DocWeaveImportError> Errors, int RowCount)
{
    /// <summary>
    /// True when at least one problem was found.
    /// </summary>
    public bool HasErrors => Errors.Count > 0;
}
