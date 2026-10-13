using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json;

namespace DocWeave.Csv;

/// <summary>
/// Entry point for importing CSV content from declarative JSON templates.
/// </summary>
public static class CsvTemplateReader
{
    /// <summary>
    /// Starts a CSV import from a JSON template. The template can set reading options ("input"), choose columns ("columns"), and describe the typed result ("output").
    /// </summary>
    /// <param name="json">The template document.</param>
    public static CsvTemplateImportBuilder FromJson(string json)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(json, nameof(json));
        return new CsvTemplateImportBuilder(json);
    }
}

/// <summary>
/// Binds a CSV source to an import template and materializes typed rows or a DataTable.
/// </summary>
public sealed class CsvTemplateImportBuilder
{
    private readonly string json;
    private Func<CsvReadBuilder>? readerFactory;

    internal CsvTemplateImportBuilder(string json)
    {
        this.json = json;
    }

    /// <summary>
    /// Reads from a file.
    /// </summary>
    /// <param name="path">The file to read.</param>
    public CsvTemplateImportBuilder FromFile(string path)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(path, nameof(path));
        readerFactory = () => CsvReader.FromFile(path);
        return this;
    }

    /// <summary>
    /// Reads from text, treated as UTF-8.
    /// </summary>
    /// <param name="text">The CSV content.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public CsvTemplateImportBuilder FromText(string text, string? label = null)
    {
        NetStandardGuard.ThrowIfNull(text, nameof(text));
        readerFactory = () => CsvReader.FromText(text, label);
        return this;
    }

    /// <summary>
    /// Reads from bytes held in memory.
    /// </summary>
    /// <param name="bytes">The CSV content.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public CsvTemplateImportBuilder FromBytes(byte[] bytes, string? label = null)
    {
        NetStandardGuard.ThrowIfNull(bytes, nameof(bytes));
        readerFactory = () => CsvReader.FromBytes(bytes, label);
        return this;
    }

    /// <summary>
    /// Reads from a stream the caller owns. The stream is left open and can be read only once.
    /// </summary>
    /// <param name="stream">The stream to read.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public CsvTemplateImportBuilder FromStream(Stream stream, string? label = null)
    {
        NetStandardGuard.ThrowIfNull(stream, nameof(stream));
        readerFactory = () => CsvReader.FromStream(stream, label);
        return this;
    }

    /// <summary>
    /// Checks the template and the source setup without reading any data. Problems are returned, not thrown.
    /// </summary>
    /// <returns>Whether the template is valid and, if not, what is wrong.</returns>
    public CsvTemplateValidationResult Validate()
    {
        var errors = new List<CsvTemplateValidationError>();
        try
        {
            using var document = JsonDocument.Parse(json, new JsonDocumentOptions { AllowTrailingCommas = true });
            _ = CreateReader(document.RootElement);
            if (document.RootElement.TryGetProperty("output", out var output)
                && output.TryGetProperty("columns", out var columns))
            {
                foreach (var column in columns.EnumerateArray())
                {
                    _ = column.GetRequiredString("name");
                    _ = TemplateImportConverter.ResolveType(column.GetOptionalString("type") ?? "string");
                }
            }
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException)
        {
            errors.Add(new CsvTemplateValidationError("$", exception.GetType().Name, exception.Message));
        }

        return new CsvTemplateValidationResult(errors.Count == 0, errors);
    }

    /// <summary>
    /// Reads the source and shapes each row the way the template says: only the declared output columns (or every column when none are declared), values converted to their declared types, then any derived columns added. A blank value becomes null unless the column is required.
    /// </summary>
    /// <returns>One dictionary per row.</returns>
    /// <exception cref="System.InvalidOperationException">A required value is missing or a value cannot be converted to its declared type.</exception>
    /// <exception cref="System.FormatException">The CSV itself is malformed.</exception>
    public IReadOnlyList<IReadOnlyDictionary<string, object?>> ReadDictionaries()
    {
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        var rows = CreateReader(root).ReadDictionaries();
        return TemplateImportConverter.ProjectRows(rows, root);
    }

    /// <summary>
    /// Like <see cref="ReadDictionaries"/>, but returns a <see cref="System.Data.DataTable"/> whose columns have the declared types.
    /// </summary>
    /// <param name="tableName">The table name. When null the template's "tableName" is used, and failing that a default.</param>
    /// <returns>The table. A missing value is stored as <see cref="DBNull"/>.</returns>
    /// <exception cref="System.InvalidOperationException">A required value is missing or a value cannot be converted to its declared type.</exception>
    public DataTable ReadDataTable(string? tableName = null)
    {
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        var rows = ReadDictionaries();
        var columns = TemplateImportConverter.ReadOutputColumns(root, rows);
        var table = new DataTable(tableName ?? root.GetOptionalString("tableName") ?? "CsvTemplateData");

        foreach (var column in columns)
        {
            table.Columns.Add(column.Name, column.Type);
        }

        foreach (var row in rows)
        {
            var dataRow = table.NewRow();
            foreach (var column in columns)
            {
                dataRow[column.Name] = row.TryGetValue(column.Name, out var value) && value is not null
                    ? value
                    : DBNull.Value;
            }

            table.Rows.Add(dataRow);
        }

        return table;
    }

    private CsvReadBuilder CreateReader(JsonElement root)
    {
        if (readerFactory is null)
        {
            throw new InvalidOperationException("A CSV source has not been supplied.");
        }

        var reader = readerFactory();
        if (root.TryGetProperty("input", out var input))
        {
            var delimiter = input.GetOptionalString("delimiter");
            if (!string.IsNullOrEmpty(delimiter))
            {
                reader.DelimitedBy(delimiter);
            }

            var quote = input.GetOptionalString("quote");
            if (!string.IsNullOrEmpty(quote))
            {
                reader.Quote(quote[0]);
            }

            var encoding = input.GetOptionalString("encoding");
            if (!string.IsNullOrWhiteSpace(encoding))
            {
                reader.Encoding(System.Text.Encoding.GetEncoding(encoding));
            }

            reader.HasHeaderRow(input.GetOptionalBool("hasHeaderRow") ?? true);
            reader.TrimFields(input.GetOptionalBool("trimFields") ?? false);
            reader.SkipEmptyLines(input.GetOptionalBool("skipEmptyLines") ?? true);
        }

        TemplateImportConverter.ApplyCsvColumnSelection(reader, root);
        return reader;
    }
}

internal sealed record TemplateImportColumn(string Name, string Source, Type Type, bool Required, bool Nullable);

internal static class TemplateImportConverter
{
    public static void ApplyCsvColumnSelection(CsvReadBuilder reader, JsonElement root)
    {
        if (!root.TryGetProperty("columns", out var columns))
        {
            return;
        }

        reader.Columns(selection =>
        {
            ApplySelection(columns, selection.IncludeLetters, selection.IncludeHeaders, selection.ExcludeLetters, selection.ExcludeHeaders);
        });
    }

    public static IReadOnlyList<IReadOnlyDictionary<string, object?>> ProjectRows(
        IReadOnlyList<IReadOnlyDictionary<string, string?>> sourceRows,
        JsonElement root)
    {
        var outputColumns = ReadDeclaredOutputColumns(root);
        // With no declared output columns, every column of the source is passed through as text.
        if (outputColumns.Count == 0 && sourceRows.Count > 0)
        {
            outputColumns = sourceRows[0].Keys
                .Select(key => new TemplateImportColumn(key, key, typeof(string), Required: false, Nullable: true))
                .ToList();
        }

        var projected = new List<IReadOnlyDictionary<string, object?>>();
        for (var rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++)
        {
            var sourceRow = sourceRows[rowIndex];
            var output = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in outputColumns)
            {
                sourceRow.TryGetValue(column.Source, out var value);
                // A blank cell counts as missing: an error if the column is required or not nullable, otherwise the column is left out of this row.
                if (value is null or "")
                {
                    if (column.Required || !column.Nullable)
                    {
                        throw new InvalidOperationException($"Imported field '{column.Source}' row {rowIndex + 1} is required.");
                    }

                    output[column.Name] = null;
                    continue;
                }

                output[column.Name] = ConvertValue(value, column.Type, column.Name, rowIndex + 1);
            }

            // Derived columns are added last so they can use the converted values.
            AddDerivedColumns(output, root);
            projected.Add(output);
        }

        return projected;
    }

    public static IReadOnlyList<TemplateImportColumn> ReadOutputColumns(
        JsonElement root,
        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows)
    {
        var columns = ReadDeclaredOutputColumns(root);
        if (columns.Count == 0)
        {
            columns = rows.Count == 0
                ? []
                : rows[0].Select(item => new TemplateImportColumn(item.Key, item.Key, item.Value?.GetType() ?? typeof(string), Required: false, Nullable: true)).ToList();
        }

        if (rows.Count == 0)
        {
            return columns;
        }

        // Columns the template did not declare (derived ones, for example) are still part of the result. They take their type
        // from the value in the first row.
        var known = columns.Select(column => column.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var item in rows[0])
        {
            if (known.Add(item.Key))
            {
                columns.Add(new TemplateImportColumn(item.Key, item.Key, item.Value?.GetType() ?? typeof(string), Required: false, Nullable: true));
            }
        }

        return columns;
    }

    public static Type ResolveType(string type)
    {
        // Type names are matched ignoring case.
        return type.ToLowerInvariant() switch
        {
            "string" or "text" => typeof(string),
            "int" or "integer" => typeof(int),
            "long" => typeof(long),
            "decimal" or "number" => typeof(decimal),
            "double" => typeof(double),
            "date" or "datetime" => typeof(DateTime),
            "bool" or "boolean" => typeof(bool),
            _ => throw new InvalidOperationException($"Import data type '{type}' is not supported.")
        };
    }

    private static List<TemplateImportColumn> ReadDeclaredOutputColumns(JsonElement root)
    {
        if (!root.TryGetProperty("output", out var output) || !output.TryGetProperty("columns", out var columns))
        {
            return [];
        }

        return columns.EnumerateArray()
            .Select(column =>
            {
                var source = column.GetOptionalString("source") ?? column.GetRequiredString("name");
                var name = column.GetRequiredString("name");
                return new TemplateImportColumn(
                    name,
                    source,
                    ResolveType(column.GetOptionalString("type") ?? "string"),
                    column.GetOptionalBool("required") ?? false,
                    column.GetOptionalBool("nullable") ?? true);
            })
            .ToList();
    }

    private static object ConvertValue(string value, Type targetType, string columnName, int rowNumber)
    {
        try
        {
            if (targetType == typeof(string))
            {
                return value;
            }

            if (targetType == typeof(int))
            {
                return int.Parse(value, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(long))
            {
                return long.Parse(value, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(decimal))
            {
                return decimal.Parse(value, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(double))
            {
                return double.Parse(value, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(DateTime))
            {
                return DateTime.Parse(value, CultureInfo.InvariantCulture);
            }

            if (targetType == typeof(bool))
            {
                return bool.Parse(value);
            }

            return value;
        }
        // Numbers and dates are parsed with the invariant culture, so the result does not depend on the machine's locale. A bad
        // value is reported with its column and row number.
        catch (Exception exception) when (exception is FormatException or OverflowException)
        {
            throw new InvalidOperationException($"Imported field '{columnName}' row {rowNumber} could not be converted to {targetType.Name}.", exception);
        }
    }

    private static void AddDerivedColumns(Dictionary<string, object?> row, JsonElement root)
    {
        if (!root.TryGetProperty("output", out var output) || !output.TryGetProperty("derivedColumns", out var derivedColumns))
        {
            return;
        }

        foreach (var derived in derivedColumns.EnumerateArray())
        {
            var name = derived.GetRequiredString("name");
            if (derived.TryGetProperty("constant", out var constant))
            {
                row[name] = ReadJsonScalar(constant);
                continue;
            }

            // Each part of a concat is either the name of a column already in the row, whose value is inserted,
            // or literal text.
            if (derived.TryGetProperty("concat", out var concat))
            {
                row[name] = string.Concat(concat.EnumerateArray().Select(part =>
                {
                    var token = part.GetString() ?? string.Empty;
                    return row.TryGetValue(token, out var value) ? Convert.ToString(value, CultureInfo.InvariantCulture) : token;
                }));
            }
        }
    }

    private static object? ReadJsonScalar(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.Null => null,
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number when value.TryGetInt64(out var integer) => integer,
            JsonValueKind.Number => value.GetDecimal(),
            _ => value.GetString()
        };
    }

    private static void ApplySelection(
        JsonElement columns,
        Func<string[], object> includeLetters,
        Func<string[], object> includeHeaders,
        Func<string[], object> excludeLetters,
        Func<string[], object> excludeHeaders)
    {
        // The four callbacks add to the selection; their return values are not needed here.
        if (columns.TryGetProperty("includeLetters", out var includeLetterValues))
        {
            includeLetters(ReadStringArray(includeLetterValues));
        }

        if (columns.TryGetProperty("includeHeaders", out var includeHeaderValues))
        {
            includeHeaders(ReadStringArray(includeHeaderValues));
        }

        if (columns.TryGetProperty("excludeLetters", out var excludeLetterValues))
        {
            excludeLetters(ReadStringArray(excludeLetterValues));
        }

        if (columns.TryGetProperty("excludeHeaders", out var excludeHeaderValues))
        {
            excludeHeaders(ReadStringArray(excludeHeaderValues));
        }
    }

    private static string[] ReadStringArray(JsonElement array)
    {
        return array.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToArray();
    }
}
