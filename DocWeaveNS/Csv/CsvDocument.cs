namespace DocWeave.Csv;

internal sealed record CsvDocument(IReadOnlyList<string> Headers, IReadOnlyList<CsvRow> Rows);

/// <summary>
/// One data row read from a CSV source.
/// </summary>
/// <param name="SourceLabel">The name of the source, from <see cref="CsvReadBuilder.Label"/>, or the file name for a file.</param>
/// <param name="RowNumber">The one-based record number in the source. It counts the header row and any blank lines that were skipped, so it matches what an editor shows.</param>
/// <param name="Values">The field values by header, matched ignoring case. A value is null when the row has too few fields.</param>
public sealed record CsvRow(string? SourceLabel, int RowNumber, IReadOnlyDictionary<string, string?> Values);
