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
        NetStandardGuard.ThrowIfNullOrWhiteSpace(name, nameof(name));
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
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
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
    /// Lets the caller cancel a read. The token is checked before the workbook is opened and again for each worksheet row, and a cancelled read throws <see cref="OperationCanceledException"/>.
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
    /// Reads the worksheet one row at a time, so a large sheet is never held in memory as a whole. The workbook stays open while
    /// the sequence is being enumerated and is closed when enumeration ends or the enumerator is disposed.
    /// </summary>
    /// <returns>
    /// One <see cref="ExcelRow"/> per data row. The header row is the first row found at or below the start cell, and its width
    /// sets the columns, because later rows have not been read yet. A cell to the right of the header row is kept as
    /// <c>Column N</c> when no column selection is configured, and left out otherwise.
    /// </returns>
    /// <remarks>
    /// The source must be seekable, because an .xlsx file is a zip archive. Progress and audit are not reported by this method.
    /// </remarks>
    public IEnumerable<ExcelRow> EnumerateRows()
    {
        return EnumerateCore(_ => { });
    }

    /// <summary>
    /// Reads the worksheet one row at a time, turning each row into an object of type <typeparamref name="T"/>. A row with a
    /// problem is not returned; each of its problems is passed to <paramref name="onError"/> and reading carries on.
    /// </summary>
    /// <typeparam name="T">
    /// The object type. A class with a public parameterless constructor is filled property by property. A type without one, such as a
    /// positional record, is built through its public constructor with the most parameters. Public settable properties and constructor parameters are matched to columns;
    /// see <see cref="DocWeaveColumnAttribute"/>.
    /// </typeparam>
    /// <param name="onError">Called for each problem, including a required column that is missing (row number 0). Null to ignore problems.</param>
    /// <returns>The objects, one at a time. The workbook is closed when enumeration ends or is disposed.</returns>
    /// <remarks>
    /// A date or time property accepts the number Excel stores for a date, as well as text such as <c>2026-01-07</c>.
    /// </remarks>
    /// <exception cref="NotSupportedException">A property has a type that a cell cannot be converted to.</exception>
    public IEnumerable<T> EnumerateObjects<T>(Action<DocWeaveImportError>? onError = null)
    {
        return EnumerateObjectsCore<T>(null, onError);
    }

    /// <summary>
    /// Like <see cref="EnumerateObjects{T}(Action{DocWeaveImportError})"/>, with the column mapping written in code instead of with
    /// attributes. A setting made in <paramref name="configureMap"/> replaces the same property's attributes.
    /// </summary>
    /// <typeparam name="T">The object type. See <see cref="DocWeaveColumnAttribute"/> for how a type is built and matched.</typeparam>
    /// <param name="configureMap">Describes the mapping, for example <c>map => map.Column(x => x.Rep, "Sales Rep")</c>.</param>
    /// <param name="onError">Called for each problem. Pass null to ignore problems.</param>
    public IEnumerable<T> EnumerateObjects<T>(Action<DocWeaveObjectMap<T>> configureMap, Action<DocWeaveImportError>? onError)
    {
        return EnumerateObjectsCore<T>(DocWeaveObjectMap<T>.Build(configureMap), onError);
    }

    private IEnumerable<T> EnumerateObjectsCore<T>(ObjectMapSettings? settings, Action<DocWeaveImportError>? onError)
    {
        RowObjectMapper<T>? mapper = null;
        var rows = EnumerateCore(headers =>
        {
            mapper = RowObjectMapper<T>.Create(headers, excelSerialDates: true, settings);
            foreach (var error in mapper.ColumnErrors)
            {
                onError?.Invoke(error);
            }
        });

        foreach (var row in rows)
        {
            if (mapper!.ColumnErrors.Count > 0)
            {
                yield break;
            }

            if (mapper.TryMap(row.RowNumber, row.Values, onError, out var item))
            {
                yield return item;
            }
        }
    }

    /// <summary>
    /// Reads the worksheet into objects of type <typeparamref name="T"/>, collecting the problems instead of stopping at the first.
    /// </summary>
    /// <typeparam name="T">
    /// The object type. A class with a public parameterless constructor is filled property by property. A type without one, such as a
    /// positional record, is built through its public constructor with the most parameters. Public settable properties and constructor parameters are matched to columns;
    /// see <see cref="DocWeaveColumnAttribute"/>.
    /// </typeparam>
    /// <returns>
    /// The objects made from rows without problems, and a list of every problem found. A row with a problem contributes no object.
    /// </returns>
    /// <exception cref="NotSupportedException">A property has a type that a cell cannot be converted to.</exception>
    public DocWeaveObjectImportResult<T> ReadObjects<T>()
    {
        return ReadObjectsCore<T>(null);
    }

    /// <summary>
    /// Like <see cref="ReadObjects{T}()"/>, with the column mapping written in code instead of with attributes. A setting made in
    /// <paramref name="configureMap"/> replaces the same property's attributes.
    /// </summary>
    /// <typeparam name="T">The object type. See <see cref="DocWeaveColumnAttribute"/> for how a type is built and matched.</typeparam>
    /// <param name="configureMap">Describes the mapping, for example <c>map => map.Column(x => x.Rep, "Sales Rep")</c>.</param>
    /// <returns>The objects made from rows without problems, and a list of every problem found.</returns>
    public DocWeaveObjectImportResult<T> ReadObjects<T>(Action<DocWeaveObjectMap<T>> configureMap)
    {
        return ReadObjectsCore<T>(DocWeaveObjectMap<T>.Build(configureMap));
    }

    private DocWeaveObjectImportResult<T> ReadObjectsCore<T>(ObjectMapSettings? settings)
    {
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("ExcelImport", "Starting", 0);

        var errors = new List<DocWeaveImportError>();
        var items = new List<T>();
        foreach (var item in EnumerateObjectsCore<T>(settings, errors.Add))
        {
            items.Add(item);
        }

        var rowCount = items.Count + errors.Where(error => error.RowNumber > 0).Select(error => error.RowNumber).Distinct().Count();
        operationContext.Report("ExcelImport", "Completed", rowCount, rowCount);
        operationContext.AuditEvent("ExcelImport", "Excel worksheet imported", new Dictionary<string, object?>
        {
            ["RowCount"] = rowCount,
            ["ErrorCount"] = errors.Count
        });
        return new DocWeaveObjectImportResult<T>(items, errors, rowCount);
    }

    private IEnumerable<ExcelRow> EnumerateCore(Action<IReadOnlyList<string>> onHeaders)
    {
        operationContext.ThrowIfCancellationRequested();
        using var stream = streamFactory();
        foreach (var row in ExcelWorkbookImporter.Enumerate(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label, onHeaders, operationContext.CancellationToken))
        {
            yield return row;
        }
    }

    /// <summary>
    /// Reads the selected worksheet rows as dictionaries keyed by imported header name.
    /// </summary>
    public IReadOnlyList<ExcelRow> ReadRows()
    {
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("ExcelImport", "Starting", 0);
        using var stream = streamFactory();
        var document = ExcelWorkbookImporter.Import(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label, operationContext.CancellationToken);
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
        var document = ExcelWorkbookImporter.Import(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label, operationContext.CancellationToken);
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
        var document = ExcelWorkbookImporter.Import(stream, sheetName, sheetIndex, startCell, hasHeaderRow, selection, label, operationContext.CancellationToken);
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
