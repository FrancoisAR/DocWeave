namespace DocWeave.Excel;

/// <summary>
/// Imported Excel worksheet data.
/// </summary>
public sealed record ExcelDocument(IReadOnlyList<string> Headers, IReadOnlyList<ExcelRow> Rows);

/// <summary>
/// One imported Excel row with source metadata and cell values.
/// </summary>
public sealed record ExcelRow(string? SourceLabel, int RowNumber, IReadOnlyDictionary<string, string?> Values);
