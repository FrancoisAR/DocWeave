using System.Text.Json;

namespace DocWeave.Csv;

/// <summary>
/// Entry point for CSV exports from declarative JSON templates.
/// </summary>
public static class CsvTemplateWriter
{
    /// <summary>
    /// Starts a CSV export from a JSON template. The template names a data source, the columns to write, and output options such as the delimiter.
    /// </summary>
    /// <param name="json">The template document.</param>
    public static CsvTemplateExportBuilder FromJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);
        return new CsvTemplateExportBuilder(json);
    }
}

/// <summary>
/// Binds data to a CSV template and writes the result. Get one from <see cref="CsvTemplateWriter.FromJson"/>.
/// </summary>
public sealed class CsvTemplateExportBuilder
{
    private readonly string json;
    private readonly Dictionary<string, TabularData> sources = new(StringComparer.OrdinalIgnoreCase);

    internal CsvTemplateExportBuilder(string json)
    {
        this.json = json;
    }

    /// <summary>
    /// Binds a collection of objects to a data source named in the template. Public properties become columns.
    /// </summary>
    /// <typeparam name="T">The type of each object.</typeparam>
    /// <param name="name">The data source name used by the template.</param>
    /// <param name="items">The objects to write.</param>
    public CsvTemplateExportBuilder WithData<T>(string name, IEnumerable<T> items)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(items);
        sources[name] = TabularData.FromObjects(items);
        return this;
    }

    /// <summary>
    /// Binds dictionary rows to a data source named in the template. Keys become columns.
    /// </summary>
    /// <param name="name">The data source name used by the template.</param>
    /// <param name="rows">The rows to write.</param>
    public CsvTemplateExportBuilder WithRows(string name, IEnumerable<IReadOnlyDictionary<string, object?>> rows)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(rows);
        sources[name] = TabularData.FromDictionaries(rows);
        return this;
    }

    /// <summary>
    /// Checks the template against the bound data without writing anything. Problems are returned, not thrown.
    /// </summary>
    /// <returns>Whether the template is valid and, if not, what is wrong.</returns>
    public CsvTemplateValidationResult Validate()
    {
        var errors = new List<CsvTemplateValidationError>();
        try
        {
            Compile();
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException)
        {
            errors.Add(new CsvTemplateValidationError("$", exception.GetType().Name, exception.Message));
        }

        return new CsvTemplateValidationResult(errors.Count == 0, errors);
    }

    /// <summary>
    /// Writes the CSV as text.
    /// </summary>
    /// <exception cref="System.InvalidOperationException">The template is invalid, for example a data source has not been bound.</exception>
    public string ToText()
    {
        return Compile().ToText();
    }

    /// <summary>
    /// Writes the CSV as UTF-8 bytes.
    /// </summary>
    /// <exception cref="System.InvalidOperationException">The template is invalid, for example a data source has not been bound.</exception>
    public byte[] ToBytes()
    {
        return Compile().ToBytes();
    }

    private CsvWriteBuilder Compile()
    {
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        var sourceName = root.GetRequiredString("source");
        if (!sources.TryGetValue(sourceName, out var source))
        {
            throw new InvalidOperationException($"CSV template source '{sourceName}' has not been bound.");
        }

        var builder = new CsvWriteBuilder(ProjectSource(source, root));
        if (root.TryGetProperty("output", out var output))
        {
            builder.DelimitedBy(output.GetOptionalString("delimiter") ?? ",");
            builder.IncludeHeaders(output.GetOptionalBool("includeHeaders") ?? true);
            if (output.GetOptionalBool("neutralizeFormulaText") == false)
            {
                builder.AllowFormulaText();
            }
        }

        if (root.TryGetProperty("columns", out var columns))
        {
            foreach (var column in columns.EnumerateArray())
            {
                var sourceColumn = column.GetRequiredString("source");
                if (column.TryGetProperty("caption", out var caption))
                {
                    builder.HeaderAlias(sourceColumn, caption.GetString() ?? sourceColumn);
                }
            }
        }

        return builder;
    }

    private static TabularData ProjectSource(TabularData source, JsonElement root)
    {
        if (!root.TryGetProperty("columns", out var columns))
        {
            return source;
        }

        var selected = columns.EnumerateArray().Select(column => column.GetRequiredString("source")).ToList();
        var indexes = selected
            .Select(column => source.IndexOf(column) is var index and >= 0
                ? index
                : throw new InvalidOperationException($"CSV source field '{column}' was not found."))
            .ToList();

        var rows = source.Rows
            .Select(row => (IReadOnlyList<object?>)indexes.Select(index => index < row.Count ? row[index] : null).ToList())
            .ToList();

        return new TabularData(selected, rows);
    }
}

/// <summary>
/// The outcome of checking a CSV template against its bound data.
/// </summary>
/// <param name="IsValid">True when no problem was found.</param>
/// <param name="Errors">What is wrong. Empty when the template is valid.</param>
public sealed record CsvTemplateValidationResult(bool IsValid, IReadOnlyList<CsvTemplateValidationError> Errors);

/// <summary>
/// One problem found in a CSV template.
/// </summary>
/// <param name="Path">Where the problem is. Currently always "$", meaning the whole template.</param>
/// <param name="Code">The name of the kind of problem, taken from the exception type that reported it.</param>
/// <param name="Message">A description of the problem.</param>
public sealed record CsvTemplateValidationError(string Path, string Code, string Message);
