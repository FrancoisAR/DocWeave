using System.Data;
using System.Text.Json;
using DocWeave.Csv;

namespace DocWeave.Excel;

/// <summary>
/// Entry point for importing Excel worksheet content from declarative JSON templates.
/// </summary>
public static class ExcelTemplateReader
{
    /// <summary>
    /// Starts an Excel import from a JSON template. The template can pick the worksheet and start cell ("input"), choose columns ("columns"), and describe the typed result ("output").
    /// </summary>
    /// <param name="json">The template document.</param>
    public static ExcelTemplateImportBuilder FromJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);
        return new ExcelTemplateImportBuilder(json);
    }
}

/// <summary>
/// Binds an Excel source to an import template and materializes typed rows or a DataTable.
/// </summary>
public sealed class ExcelTemplateImportBuilder
{
    private readonly string json;
    private Func<ExcelReadBuilder>? readerFactory;

    internal ExcelTemplateImportBuilder(string json)
    {
        this.json = json;
    }

    /// <summary>
    /// Reads from a workbook file.
    /// </summary>
    /// <param name="path">The file to read.</param>
    public ExcelTemplateImportBuilder FromFile(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        readerFactory = () => ExcelReader.FromFile(path);
        return this;
    }

    /// <summary>
    /// Reads from workbook bytes held in memory.
    /// </summary>
    /// <param name="bytes">The workbook content.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public ExcelTemplateImportBuilder FromBytes(byte[] bytes, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        readerFactory = () => ExcelReader.FromBytes(bytes, label);
        return this;
    }

    /// <summary>
    /// Reads from a stream the caller owns.
    /// </summary>
    /// <param name="stream">The stream to read.</param>
    /// <param name="label">A name for the source, used in messages.</param>
    public ExcelTemplateImportBuilder FromStream(Stream stream, string? label = null)
    {
        ArgumentNullException.ThrowIfNull(stream);
        readerFactory = () => ExcelReader.FromStream(stream, label);
        return this;
    }

    /// <summary>
    /// Checks the template and the source setup without reading any data. Problems are returned, not thrown.
    /// </summary>
    /// <returns>Whether the template is valid and, if not, what is wrong.</returns>
    public ExcelTemplateValidationResult Validate()
    {
        var errors = new List<ExcelTemplateValidationError>();
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
            errors.Add(new ExcelTemplateValidationError("$", exception.GetType().Name, exception.Message));
        }

        return new ExcelTemplateValidationResult(errors.Count == 0, errors);
    }

    /// <summary>
    /// Reads the worksheet and shapes each row the way the template says: only the declared output columns (or every column when none are declared), values converted to their declared types, then any derived columns added. A blank value becomes null unless the column is required.
    /// </summary>
    /// <returns>One dictionary per row.</returns>
    /// <exception cref="System.InvalidOperationException">A required value is missing or a value cannot be converted to its declared type.</exception>
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
        var table = new DataTable(tableName ?? root.GetOptionalString("tableName") ?? "ExcelTemplateData");

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

    private ExcelReadBuilder CreateReader(JsonElement root)
    {
        if (readerFactory is null)
        {
            throw new InvalidOperationException("An Excel source has not been supplied.");
        }

        var reader = readerFactory();
        if (root.TryGetProperty("input", out var input))
        {
            var sheet = input.GetOptionalString("sheet");
            if (!string.IsNullOrWhiteSpace(sheet))
            {
                reader.Sheet(sheet);
            }

            if (input.TryGetProperty("sheetIndex", out var sheetIndex))
            {
                reader.SheetAt(sheetIndex.GetInt32());
            }

            var startCell = input.GetOptionalString("startCell");
            if (!string.IsNullOrWhiteSpace(startCell))
            {
                reader.StartAt(startCell);
            }

            reader.HasHeaderRow(input.GetOptionalBool("hasHeaderRow") ?? true);
        }

        ApplyColumnSelection(reader, root);
        return reader;
    }

    private static void ApplyColumnSelection(ExcelReadBuilder reader, JsonElement root)
    {
        if (!root.TryGetProperty("columns", out var columns))
        {
            return;
        }

        reader.Columns(selection =>
        {
            if (columns.TryGetProperty("includeLetters", out var includeLetterValues))
            {
                selection.IncludeLetters(ReadStringArray(includeLetterValues));
            }

            if (columns.TryGetProperty("includeHeaders", out var includeHeaderValues))
            {
                selection.IncludeHeaders(ReadStringArray(includeHeaderValues));
            }

            if (columns.TryGetProperty("excludeLetters", out var excludeLetterValues))
            {
                selection.ExcludeLetters(ReadStringArray(excludeLetterValues));
            }

            if (columns.TryGetProperty("excludeHeaders", out var excludeHeaderValues))
            {
                selection.ExcludeHeaders(ReadStringArray(excludeHeaderValues));
            }
        });
    }

    private static string[] ReadStringArray(JsonElement array)
    {
        return array.EnumerateArray().Select(item => item.GetString() ?? string.Empty).ToArray();
    }
}
