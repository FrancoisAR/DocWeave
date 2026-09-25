using System.Text;
using DocWeave;

namespace DocWeave.Csv;

/// <summary>
/// Fluent builder for writing CSV data to text, bytes, streams, or files.
/// </summary>
public sealed class CsvWriteBuilder
{
    private readonly TabularData data;
    private readonly List<string> columnOrder = [];
    private readonly Dictionary<string, string> headerAliases = new(StringComparer.OrdinalIgnoreCase);
    private string delimiter = ",";
    private char quote = '"';
    private string newLine = Environment.NewLine;
    private Encoding encoding = System.Text.Encoding.UTF8;
    private bool includeHeaders = true;
    private bool neutralizeFormulas = true;
    private readonly DocWeaveOperationContext operationContext = new();

    internal CsvWriteBuilder(TabularData data)
    {
        this.data = data;
    }

    /// <summary>
    /// Sets the delimiter between fields. Common values are comma, semicolon, tab, or pipe.
    /// </summary>
    public CsvWriteBuilder DelimitedBy(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            throw new ArgumentException("Delimiter cannot be empty.", nameof(value));
        }

        delimiter = value;
        return this;
    }

    /// <summary>
    /// Sets the quote character used when escaping fields.
    /// </summary>
    public CsvWriteBuilder Quote(char value)
    {
        quote = value;
        return this;
    }

    /// <summary>
    /// Sets the text encoding used by ToBytes, SaveAs, and SaveTo.
    /// </summary>
    public CsvWriteBuilder Encoding(Encoding value)
    {
        encoding = value ?? throw new ArgumentNullException(nameof(value));
        return this;
    }

    /// <summary>
    /// Sets the line ending used between CSV records.
    /// </summary>
    public CsvWriteBuilder NewLine(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            throw new ArgumentException("New line cannot be empty.", nameof(value));
        }

        newLine = value;
        return this;
    }

    /// <summary>
    /// Controls whether the first CSV record contains headers.
    /// </summary>
    public CsvWriteBuilder IncludeHeaders(bool value = true)
    {
        includeHeaders = value;
        return this;
    }

    /// <summary>
    /// Sets the exported column order. Unlisted columns are appended afterwards in source order.
    /// </summary>
    public CsvWriteBuilder ColumnOrder(params string[] columns)
    {
        ArgumentNullException.ThrowIfNull(columns);
        columnOrder.Clear();
        columnOrder.AddRange(columns.Where(column => !string.IsNullOrWhiteSpace(column)));
        return this;
    }

    /// <summary>
    /// Replaces exported header captions without changing source column names.
    /// </summary>
    public CsvWriteBuilder HeaderAliases(IReadOnlyDictionary<string, string> aliases)
    {
        ArgumentNullException.ThrowIfNull(aliases);
        foreach (var alias in aliases)
        {
            headerAliases[alias.Key] = alias.Value;
        }

        return this;
    }

    /// <summary>
    /// Replaces one exported header caption without changing the source column name.
    /// </summary>
    public CsvWriteBuilder HeaderAlias(string columnName, string caption)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        ArgumentException.ThrowIfNullOrWhiteSpace(caption);
        headerAliases[columnName] = caption;
        return this;
    }

    /// <summary>
    /// Writes text that starts with =, +, -, @, tab, or carriage return exactly as supplied.
    /// By default such text values (and headers) get a leading apostrophe so a spreadsheet
    /// opening the file treats them as text instead of running them as a formula.
    /// Numbers are never changed, only values that are strings.
    /// </summary>
    public CsvWriteBuilder AllowFormulaText(bool value = true)
    {
        neutralizeFormulas = !value;
        return this;
    }

    /// <summary>
    /// Sends progress reports to the caller while writing: at the start, after the header, every 1,000 rows, and at the end.
    /// </summary>
    public CsvWriteBuilder WithProgress(IProgress<DocWeaveOperationProgress> progress)
    {
        operationContext.Progress = progress ?? throw new ArgumentNullException(nameof(progress));
        return this;
    }

    /// <summary>
    /// Lets the caller cancel a write. The token is checked before writing starts and before each row, and a cancelled write throws <see cref="OperationCanceledException"/>.
    /// </summary>
    public CsvWriteBuilder WithCancellation(CancellationToken cancellationToken)
    {
        operationContext.CancellationToken = cancellationToken;
        return this;
    }

    /// <summary>
    /// Sends audit events when writing starts and finishes. They hold the row and column counts, never the values.
    /// </summary>
    public CsvWriteBuilder Audit(Action<DocWeaveAuditEvent> audit)
    {
        operationContext.Audit = audit ?? throw new ArgumentNullException(nameof(audit));
        return this;
    }

    /// <summary>
    /// Builds the CSV as text.
    /// </summary>
    public string ToText()
    {
        using var writer = new StringWriter(System.Globalization.CultureInfo.InvariantCulture);
        WriteTo(writer);
        return writer.ToString();
    }

    /// <summary>
    /// Streams the CSV to a TextWriter without building the full output string first.
    /// </summary>
    public void WriteTo(TextWriter writer)
    {
        operationContext.ThrowIfCancellationRequested();
        var columns = ResolveColumns();
        var total = data.Rows.Count + (includeHeaders ? 1 : 0);
        var completed = 0;
        var writtenAnyRecord = false;
        operationContext.Report("CsvExport", "Starting", completed, total);
        operationContext.AuditEvent("CsvExport", "Writing CSV", new Dictionary<string, object?>
        {
            ["RowCount"] = data.Rows.Count,
            ["ColumnCount"] = columns.Count
        });

        if (includeHeaders)
        {
            WriteRecord(writer, columns.Select(column => NeutralizeFormula(headerAliases.TryGetValue(column, out var alias) ? alias : column)), ref writtenAnyRecord);
            completed++;
            operationContext.Report("CsvExport", "Writing", completed, total);
        }

        // Look each output column up once instead of once per row.
        var columnIndexes = columns.Select(data.IndexOf).ToList();
        foreach (var row in data.Rows)
        {
            operationContext.ThrowIfCancellationRequested();
            WriteRecord(writer, columnIndexes.Select(index => FormatCell(index >= 0 && index < row.Count ? row[index] : null)), ref writtenAnyRecord);
            completed++;
            if (completed % 1000 == 0 || completed == total)
            {
                operationContext.Report("CsvExport", "Writing", completed, total);
            }
        }

        operationContext.Report("CsvExport", "Completed", completed, total);
        operationContext.AuditEvent("CsvExport", "CSV written", new Dictionary<string, object?>
        {
            ["RowCount"] = data.Rows.Count,
            ["ColumnCount"] = columns.Count
        });
    }

    /// <summary>
    /// Builds the CSV as bytes using the configured encoding.
    /// </summary>
    public byte[] ToBytes()
    {
        return encoding.GetBytes(ToText());
    }

    /// <summary>
    /// Writes the CSV to a caller-owned stream.
    /// </summary>
    public void SaveTo(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        using var writer = new StreamWriter(stream, encoding, bufferSize: 65536, leaveOpen: true);
        WriteTo(writer);
    }

    /// <summary>
    /// Writes the CSV to a file, replacing it if it already exists.
    /// </summary>
    public void SaveAs(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        using var stream = File.Create(path);
        SaveTo(stream);
    }

    private IReadOnlyList<string> ResolveColumns()
    {
        if (columnOrder.Count == 0)
        {
            return data.Columns;
        }

        var ordered = new List<string>();
        foreach (var column in columnOrder)
        {
            var sourceColumn = data.Columns.FirstOrDefault(item => string.Equals(item, column, StringComparison.OrdinalIgnoreCase));
            if (sourceColumn is null)
            {
                throw new InvalidOperationException($"CSV column '{column}' was not found.");
            }

            ordered.Add(sourceColumn);
        }

        ordered.AddRange(data.Columns.Where(column => ordered.All(item => !string.Equals(item, column, StringComparison.OrdinalIgnoreCase))));
        return ordered;
    }

    private void AppendRecord(StringBuilder builder, IEnumerable<string?> values)
    {
        if (builder.Length > 0)
        {
            builder.Append(newLine);
        }

        builder.Append(string.Join(delimiter, values.Select(Escape)));
    }

    private void WriteRecord(TextWriter writer, IEnumerable<string?> values, ref bool writtenAnyRecord)
    {
        if (writtenAnyRecord)
        {
            writer.Write(newLine);
        }

        writer.Write(string.Join(delimiter, values.Select(Escape)));
        writtenAnyRecord = true;
    }

    private string Escape(string? value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        var mustQuote = value.Contains(delimiter, StringComparison.Ordinal)
            || value.Contains(quote)
            || value.Contains('\r')
            || value.Contains('\n');

        if (!mustQuote)
        {
            return value;
        }

        return quote + value.Replace(quote.ToString(), new string(quote, 2), StringComparison.Ordinal) + quote;
    }

    private string FormatCell(object? value)
    {
        return value is string text ? NeutralizeFormula(text) : FormatValue(value);
    }

    private string NeutralizeFormula(string text)
    {
        if (!neutralizeFormulas || text.Length == 0)
        {
            return text;
        }

        return text[0] is '=' or '+' or '-' or '@' or '\t' or '\r' or '\n' ? "'" + text : text;
    }

    private static string FormatValue(object? value)
    {
        return value switch
        {
            null => string.Empty,
            DateTime dateTime => dateTime.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            DateOnly dateOnly => dateOnly.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            TimeOnly timeOnly => timeOnly.ToString("O", System.Globalization.CultureInfo.InvariantCulture),
            IFormattable formattable => formattable.ToString(null, System.Globalization.CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty
        };
    }
}
