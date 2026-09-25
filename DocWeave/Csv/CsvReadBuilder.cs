using System.Data;
using System.Text;
using DocWeave;

namespace DocWeave.Csv;

/// <summary>
/// Fluent builder for reading one CSV source. Start with <see cref="CsvReader"/>, set options, then call a Read method or <see cref="EnumerateRows"/>.
/// Malformed input throws <see cref="FormatException"/> from the Read methods: a quoted field that never closes, a duplicate header (when set to be rejected), or a row of the wrong length (when set to be rejected).
/// </summary>
public sealed class CsvReadBuilder
{
    private readonly Func<Stream> streamFactory;
    private readonly bool ownsStream;
    private string? label;
    private readonly ColumnSelection selection = new("CSV", keepRequestedOrder: true);
    private string delimiter = ",";
    private char quote = '"';
    private Encoding encoding = System.Text.Encoding.UTF8;
    private bool hasHeaderRow = true;
    private bool trimFields;
    private bool skipEmptyLines = true;
    private CsvDuplicateHeaderMode duplicateHeaders = CsvDuplicateHeaderMode.Rename;
    private CsvRowLengthMode rowLengthMode = CsvRowLengthMode.Ignore;
    private readonly DocWeaveOperationContext operationContext = new();

    internal CsvReadBuilder(Func<Stream> streamFactory, bool ownsStream, string? label)
    {
        this.streamFactory = streamFactory;
        this.ownsStream = ownsStream;
        this.label = label;
    }

    /// <summary>
    /// Names the source. The name is used in error messages and stored in <see cref="CsvRow.SourceLabel"/>.
    /// </summary>
    public CsvReadBuilder Label(string value)
    {
        label = value;
        return this;
    }

    /// <summary>
    /// Sets the text between fields. It may be more than one character. The default is a comma.
    /// </summary>
    /// <exception cref="System.ArgumentException">The delimiter is null or empty.</exception>
    public CsvReadBuilder DelimitedBy(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            throw new ArgumentException("Delimiter cannot be empty.", nameof(value));
        }

        delimiter = value;
        return this;
    }

    /// <summary>
    /// Sets the quote character. A quote inside a quoted field is written as two quotes. The default is a double quote.
    /// </summary>
    public CsvReadBuilder Quote(char value)
    {
        quote = value;
        return this;
    }

    /// <summary>
    /// Sets the text encoding. The default is UTF-8. A byte order mark at the start of the file takes precedence over this setting.
    /// </summary>
    public CsvReadBuilder Encoding(Encoding value)
    {
        encoding = value ?? throw new ArgumentNullException(nameof(value));
        return this;
    }

    /// <summary>
    /// Says whether the first record holds column headers. The default is true. Without headers the columns are named by letter: A, B, C, and so on.
    /// </summary>
    public CsvReadBuilder HasHeaderRow(bool value = true)
    {
        hasHeaderRow = value;
        return this;
    }

    /// <summary>
    /// Removes leading and trailing whitespace from every field, including quoted ones. The default is to keep fields exactly as written.
    /// </summary>
    public CsvReadBuilder TrimFields(bool value = true)
    {
        trimFields = value;
        return this;
    }

    /// <summary>
    /// Skips lines that contain nothing at all. On by default. Row numbers still count skipped lines.
    /// </summary>
    public CsvReadBuilder SkipEmptyLines(bool value = true)
    {
        skipEmptyLines = value;
        return this;
    }

    /// <summary>
    /// Chooses whether duplicate header names are renamed (the default) or rejected.
    /// </summary>
    public CsvReadBuilder DuplicateHeaders(CsvDuplicateHeaderMode mode)
    {
        duplicateHeaders = mode;
        return this;
    }

    /// <summary>
    /// Chooses what happens when a row has more or fewer fields than the header row.
    /// </summary>
    public CsvReadBuilder OnRowLengthMismatch(CsvRowLengthMode mode)
    {
        rowLengthMode = mode;
        return this;
    }

    /// <summary>
    /// Chooses which columns to read, by letter or by header. With nothing chosen every column is read.
    /// </summary>
    public CsvReadBuilder Columns(Action<CsvColumnSelectionBuilder> configure)
    {
        configure(new CsvColumnSelectionBuilder(selection));
        return this;
    }

    /// <summary>
    /// Sends progress reports to the caller when a read starts and finishes. Reports are not sent by <see cref="EnumerateRows"/>.
    /// </summary>
    public CsvReadBuilder WithProgress(IProgress<DocWeaveOperationProgress> progress)
    {
        operationContext.Progress = progress ?? throw new ArgumentNullException(nameof(progress));
        return this;
    }

    /// <summary>
    /// Lets the caller cancel a read. The token is checked before reading starts, and a cancelled read throws <see cref="OperationCanceledException"/>.
    /// </summary>
    public CsvReadBuilder WithCancellation(CancellationToken cancellationToken)
    {
        operationContext.CancellationToken = cancellationToken;
        return this;
    }

    /// <summary>
    /// Sends an audit event when a read completes. It holds the row and column counts, never the values.
    /// </summary>
    public CsvReadBuilder Audit(Action<DocWeaveAuditEvent> audit)
    {
        operationContext.Audit = audit ?? throw new ArgumentNullException(nameof(audit));
        return this;
    }

    /// <summary>
    /// Reads rows one at a time without loading the whole file. The source stays open while the
    /// sequence is being enumerated and is released when enumeration ends or the enumerator is disposed.
    /// </summary>
    public IEnumerable<CsvRow> EnumerateRows()
    {
        using var session = Open();
        foreach (var row in session.ReadRows())
        {
            yield return row;
        }
    }

    /// <summary>
    /// Reads rows one at a time, turning each into an object of type <typeparamref name="T"/>. A row with a problem is not
    /// returned; each of its problems is passed to <paramref name="onError"/> and reading carries on. The source stays open while
    /// the sequence is being enumerated and is released when enumeration ends or the enumerator is disposed.
    /// </summary>
    /// <typeparam name="T">
    /// The object type. A class with a public parameterless constructor is filled property by property. A type without one, such as a
    /// positional record, is built through its public constructor with the most parameters. Public settable properties and constructor parameters are matched to columns;
    /// see <see cref="DocWeaveColumnAttribute"/>.
    /// </typeparam>
    /// <param name="onError">Called for each problem, including a required column that is missing (row number 0). Null to ignore problems.</param>
    /// <exception cref="NotSupportedException">A property has a type that a text value cannot be converted to.</exception>
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
        using var session = Open();
        var mapper = RowObjectMapper<T>.Create(session.Headers, excelSerialDates: false, settings);
        foreach (var error in mapper.ColumnErrors)
        {
            onError?.Invoke(error);
        }

        if (mapper.ColumnErrors.Count > 0)
        {
            yield break;
        }

        foreach (var row in session.ReadRows())
        {
            if (mapper.TryMap(row.RowNumber, row.Values, onError, out var item))
            {
                yield return item;
            }
        }
    }

    /// <summary>
    /// Reads every row into objects of type <typeparamref name="T"/>, collecting the problems instead of stopping at the first.
    /// </summary>
    /// <typeparam name="T">
    /// The object type. A class with a public parameterless constructor is filled property by property. A type without one, such as a
    /// positional record, is built through its public constructor with the most parameters. Public settable properties and constructor parameters are matched to columns;
    /// see <see cref="DocWeaveColumnAttribute"/>.
    /// </typeparam>
    /// <returns>
    /// The objects made from rows without problems, and a list of every problem found. A row with a problem contributes no object.
    /// </returns>
    /// <exception cref="NotSupportedException">A property has a type that a text value cannot be converted to.</exception>
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
        operationContext.Report("CsvImport", "Starting", 0);

        var errors = new List<DocWeaveImportError>();
        var items = new List<T>();
        foreach (var item in EnumerateObjectsCore<T>(settings, errors.Add))
        {
            items.Add(item);
        }

        var rowCount = items.Count + errors.Where(error => error.RowNumber > 0).Select(error => error.RowNumber).Distinct().Count();
        operationContext.Report("CsvImport", "Completed", rowCount, rowCount);
        operationContext.AuditEvent("CsvImport", "CSV imported", new Dictionary<string, object?>
        {
            ["RowCount"] = rowCount,
            ["ErrorCount"] = errors.Count
        });
        return new DocWeaveObjectImportResult<T>(items, errors, rowCount);
    }

    /// <summary>
    /// Reads every row into memory.
    /// </summary>
    /// <returns>One <see cref="CsvRow"/> per data row. Blank lines are skipped unless <see cref="SkipEmptyLines"/> was turned off.</returns>
    public IReadOnlyList<CsvRow> ReadRows()
    {
        return ReadDocument().Rows;
    }

    /// <summary>
    /// Reads every row into memory and returns them with the headers.
    /// </summary>
    /// <returns>The headers that were kept, the rows, and any warnings. Warnings are currently always empty.</returns>
    public DocWeaveImportResult<CsvRow> ReadResult()
    {
        var document = ReadDocument();
        return new DocWeaveImportResult<CsvRow>(document.Headers, document.Rows, []);
    }

    /// <summary>
    /// Reads every row into memory as dictionaries.
    /// </summary>
    /// <returns>One dictionary per data row, keyed by header (ignoring case). A value is null when the row has too few fields.</returns>
    public IReadOnlyList<IReadOnlyDictionary<string, string?>> ReadDictionaries()
    {
        return ReadDocument().Rows
            .Select(row => (IReadOnlyDictionary<string, string?>)row.Values)
            .ToList();
    }

    /// <summary>
    /// Like <see cref="ReadDictionaries"/>, but also returns the headers and any warnings.
    /// </summary>
    /// <returns>The headers that were kept, one dictionary per row, and any warnings.</returns>
    public DocWeaveImportResult<IReadOnlyDictionary<string, string?>> ReadDictionaryResult()
    {
        var document = ReadDocument();
        return new DocWeaveImportResult<IReadOnlyDictionary<string, string?>>(
            document.Headers,
            document.Rows.Select(row => (IReadOnlyDictionary<string, string?>)row.Values).ToList(),
            []);
    }

    /// <summary>
    /// Reads every row into a <see cref="System.Data.DataTable"/>.
    /// </summary>
    /// <param name="tableName">The name to give the table.</param>
    /// <returns>A table whose columns are all text. A missing value is stored as <see cref="DBNull"/> rather than an empty string.</returns>
    public DataTable ReadDataTable(string tableName = "CsvData")
    {
        var document = ReadDocument();
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
                // A missing value stays null rather than becoming an empty string.
                dataRow[header] = row.Values.TryGetValue(header, out var value) && value is not null
                    ? value
                    : DBNull.Value;
            }

            table.Rows.Add(dataRow);
        }

        return table;
    }

    private CsvDocument ReadDocument()
    {
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("CsvImport", "Starting", 0);
        using var session = Open();
        var rows = session.ReadRows().ToList();
        operationContext.Report("CsvImport", "Completed", rows.Count, rows.Count);
        operationContext.AuditEvent("CsvImport", "CSV imported", new Dictionary<string, object?>
        {
            ["RowCount"] = rows.Count,
            ["ColumnCount"] = session.Headers.Count
        });
        return new CsvDocument(session.Headers, rows);
    }

    private ReadSession Open()
    {
        var stream = streamFactory();
        try
        {
            return new ReadSession(this, stream);
        }
        catch
        {
            if (ownsStream)
            {
                stream.Dispose();
            }

            throw;
        }
    }

    /// <summary>
    /// One pass over the source. Owns the reader, and the stream too when the library opened it.
    /// </summary>
    private sealed class ReadSession : IDisposable
    {
        private readonly CsvReadBuilder builder;
        private readonly Stream stream;
        private readonly StreamReader reader;
        private readonly IEnumerator<CsvParsedRecord> records;
        private readonly IReadOnlyList<string> sourceHeaders;
        private readonly IReadOnlyList<int> selectedIndexes;
        private readonly CsvParsedRecord? firstDataRecord;
        private readonly bool hasData;

        public ReadSession(CsvReadBuilder builder, Stream stream)
        {
            this.builder = builder;
            this.stream = stream;
            reader = new StreamReader(stream, builder.encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
            var parser = new CsvParser(builder.delimiter, builder.quote, builder.trimFields, builder.skipEmptyLines);
            records = parser.Parse(reader).GetEnumerator();

            if (!records.MoveNext())
            {
                sourceHeaders = [];
                selectedIndexes = [];
                Headers = [];
                return;
            }

            var first = records.Current;
            if (builder.hasHeaderRow)
            {
                sourceHeaders = CsvHeaderNormalizer.Normalize(first.Fields, builder.duplicateHeaders);
            }
            else
            {
                sourceHeaders = Enumerable.Range(0, first.Fields.Count).Select(index => ColumnLetters.FromNumber(index + 1)).ToList();
                firstDataRecord = first;
            }

            hasData = true;
            selectedIndexes = builder.selection.Resolve(sourceHeaders);
            Headers = selectedIndexes.Select(index => sourceHeaders[index]).ToList();
        }

        public IReadOnlyList<string> Headers { get; }

        public IEnumerable<CsvRow> ReadRows()
        {
            if (!hasData)
            {
                yield break;
            }

            if (firstDataRecord is { } pending)
            {
                yield return ToRow(pending);
            }

            while (records.MoveNext())
            {
                yield return ToRow(records.Current);
            }
        }

        private CsvRow ToRow(CsvParsedRecord record)
        {
            builder.operationContext.ThrowIfCancellationRequested();
            var values = record.Fields;
            if (builder.rowLengthMode == CsvRowLengthMode.Throw && values.Count != sourceHeaders.Count)
            {
                var source = builder.label is null ? "CSV" : $"CSV '{builder.label}'";
                throw new FormatException($"{source} row {record.Number} has {values.Count} fields but {sourceHeaders.Count} were expected.");
            }

            var rowValues = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            foreach (var index in selectedIndexes)
            {
                rowValues[sourceHeaders[index]] = index < values.Count ? values[index] : null;
            }

            return new CsvRow(builder.label, record.Number, rowValues);
        }

        public void Dispose()
        {
            records.Dispose();
            reader.Dispose();
            if (builder.ownsStream)
            {
                stream.Dispose();
            }
        }
    }
}
