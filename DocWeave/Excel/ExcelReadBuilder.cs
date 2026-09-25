using System.Data;
using DocWeave;

namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for importing rows from one Excel worksheet.
/// </summary>
public sealed class ExcelReadBuilder
{
    private readonly Func<Stream> streamFactory;
    private readonly ColumnSelection selection = new("Excel", keepRequestedOrder: false);
    private string? label;
    private string? sheetName;
    private int? sheetIndex;
    private ExcelCellAddress startCell = new(1, 1);
    private bool hasHeaderRow = true;
    private readonly DocWeaveOperationContext operationContext = new();

    internal ExcelReadBuilder(Func<Stream> streamFactory, string? label)
    {
        this.streamFactory = streamFactory;
        this.label = label;
    }

    /// <summary>
    /// Sets a friendly source label that is copied into imported row metadata.
    /// </summary>
    public ExcelReadBuilder Label(string value)
    {
        label = value;
        return this;
    }

    /// <summary>
    /// Selects a sheet by name. When omitted, the first worksheet is imported.
    /// </summary>
    public ExcelReadBuilder Sheet(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        sheetName = name;
        sheetIndex = null;
        return this;
    }

    /// <summary>
    /// Selects a sheet by zero-based workbook order. When omitted, the first worksheet is imported.
    /// </summary>
    public ExcelReadBuilder SheetAt(int index)
    {
        if (index < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(index), "Sheet index cannot be negative.");
        }

        sheetIndex = index;
        sheetName = null;
        return this;
    }

    /// <summary>
    /// Sets the top-left cell for the import range.
    /// </summary>
    public ExcelReadBuilder StartAt(string cellReference)
    {
        startCell = ExcelCellAddress.Parse(cellReference);
        return this;
    }

    /// <summary>
    /// Controls whether the first imported row should be treated as column headers.
    /// </summary>
    public ExcelReadBuilder HasHeaderRow(bool value = true)
    {
        hasHeaderRow = value;
        return this;
    }

    /// <summary>
    /// Configures included or excluded columns by Excel letters and/or header captions.
    /// </summary>
    public ExcelReadBuilder Columns(Action<ExcelColumnSelectionBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        configure(new ExcelColumnSelectionBuilder(selection));
        return this;
    }

    /// <summary>
    /// Sends progress reports to the caller when a read starts and finishes.
    /// </summary>
    public ExcelReadBuilder WithProgress(IProgress<DocWeaveOperationProgress> progress)
    {
        operationContext.Progress = progress ?? throw new ArgumentNullException(nameof(progress));
        return this;
    }

    /// <summary>
    /// Lets the caller cancel a read. The token is checked before the workbook is opened, and a cancelled read throws <see cref="OperationCanceledException"/>.
    /// </summary>
    public ExcelReadBuilder WithCancellation(CancellationToken cancellationToken)
    {
        operationContext.CancellationToken = cancellationToken;
        return this;
    }

    /// <summary>
    /// Sends an audit event when a read completes. It holds the row and column counts, never the cell values.
    /// </summary>
    public ExcelReadBuilder Audit(Action<DocWeaveAuditEvent> audit)
    {
        operationContext.Audit = audit ?? throw new ArgumentNullException(nameof(audit));
        return this;
    }

    /// <summary>
    /// Reads the selected worksheet rows as dictionaries keyed by imported header name.
    /// </summary>
    public IReadOnlyList<ExcelRow> ReadRows()
    {
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("ExcelImport", "Starting", 0);
        using var stream = streamFactory();
        var document = ExcelWorkbookImporter.Import(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label);
        operationContext.Report("ExcelImport", "Completed", document.Rows.Count, document.Rows.Count);
        operationContext.AuditEvent("ExcelImport", "Excel worksheet imported", new Dictionary<string, object?>
        {
            ["RowCount"] = document.Rows.Count,
            ["ColumnCount"] = document.Headers.Count
        });
        return document.Rows;
    }

    /// <summary>
    /// Reads the selected worksheet rows and returns them with the headers.
    /// </summary>
    /// <returns>The headers that were kept, the rows, and any warnings. Warnings are currently always empty.</returns>
    public DocWeaveImportResult<ExcelRow> ReadResult()
    {
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("ExcelImport", "Starting", 0);
        using var stream = streamFactory();
        var document = ExcelWorkbookImporter.Import(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label);
        operationContext.Report("ExcelImport", "Completed", document.Rows.Count, document.Rows.Count);
        operationContext.AuditEvent("ExcelImport", "Excel worksheet imported", new Dictionary<string, object?>
        {
            ["RowCount"] = document.Rows.Count,
            ["ColumnCount"] = document.Headers.Count
        });
        return new DocWeaveImportResult<ExcelRow>(document.Headers, document.Rows, []);
    }

    /// <summary>
    /// Reads the selected worksheet rows as plain dictionaries.
    /// </summary>
    public IReadOnlyList<IReadOnlyDictionary<string, string?>> ReadDictionaries()
    {
        return ReadRows()
            .Select(row => (IReadOnlyDictionary<string, string?>)row.Values)
            .ToList();
    }

    /// <summary>
    /// Like <see cref="ReadDictionaries"/>, but also returns the headers and any warnings.
    /// </summary>
    /// <returns>The headers that were kept, one dictionary per row, and any warnings.</returns>
    public DocWeaveImportResult<IReadOnlyDictionary<string, string?>> ReadDictionaryResult()
    {
        var result = ReadResult();
        return new DocWeaveImportResult<IReadOnlyDictionary<string, string?>>(
            result.Headers,
            result.Rows.Select(row => (IReadOnlyDictionary<string, string?>)row.Values).ToList(),
            result.Warnings);
    }

    /// <summary>
    /// Reads the selected worksheet rows into a DataTable.
    /// </summary>
    public DataTable ReadDataTable(string tableName = "ExcelData")
    {
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("ExcelImport", "Starting", 0);
        using var stream = streamFactory();
        var document = ExcelWorkbookImporter.Import(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label);
        var table = new DataTable(tableName);

        foreach (var header in document.Headers)
        {
            table.Columns.Add(header, typeof(string));
        }

        foreach (var row in document.Rows)
        {
            var dataRow = table.NewRow();
            foreach (var header in document.Headers)
            {
                dataRow[header] = row.Values.TryGetValue(header, out var value) ? value ?? string.Empty : string.Empty;
            }

            table.Rows.Add(dataRow);
        }

        operationContext.Report("ExcelImport", "Completed", document.Rows.Count, document.Rows.Count);
        operationContext.AuditEvent("ExcelImport", "Excel worksheet imported", new Dictionary<string, object?>
        {
            ["RowCount"] = document.Rows.Count,
            ["ColumnCount"] = document.Headers.Count
        });
        return table;
    }
}
