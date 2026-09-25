using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// Entry point for creating Excel workbooks from declarative JSON templates.
/// </summary>
public static class ExcelTemplateWriter
{
    /// <summary>
    /// Starts an Excel template export from a JSON template document.
    /// </summary>
    public static ExcelTemplateExportBuilder FromJson(string json)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(json);
        return new ExcelTemplateExportBuilder(json);
    }

    /// <summary>
    /// Starts an Excel template export from a stored template provider.
    /// </summary>
    public static ExcelTemplateExportBuilder FromTemplateId(IExcelTemplateProvider provider, string templateId)
    {
        ArgumentNullException.ThrowIfNull(provider);
        ArgumentException.ThrowIfNullOrWhiteSpace(templateId);
        return FromJson(provider.GetTemplateJson(templateId));
    }
}

/// <summary>
/// Binds data to a JSON template and renders it through the same workbook renderer as the fluent API.
/// </summary>
public sealed class ExcelTemplateExportBuilder
{
    private static readonly JsonDocumentOptions JsonOptions = new() { AllowTrailingCommas = true };

    private readonly string json;
    private readonly Dictionary<string, TabularData> sources = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, object?> parameters = new(StringComparer.OrdinalIgnoreCase);

    internal ExcelTemplateExportBuilder(string json)
    {
        this.json = json;
    }

    /// <summary>
    /// Binds an object collection to a named template source.
    /// </summary>
    public ExcelTemplateExportBuilder WithData<T>(string name, IEnumerable<T> items)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(items);
        sources[name] = TabularData.FromObjects(items);
        return this;
    }

    /// <summary>
    /// Binds dictionary rows to a named template source.
    /// </summary>
    public ExcelTemplateExportBuilder WithRows(string name, IEnumerable<IReadOnlyDictionary<string, object?>> rows)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(rows);
        sources[name] = TabularData.FromDictionaries(rows);
        return this;
    }

    /// <summary>
    /// Binds a scalar parameter for template placeholders such as {{parameters.reportDate}}.
    /// </summary>
    public ExcelTemplateExportBuilder WithParameter(string name, object? value)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        parameters[name] = value;
        return this;
    }

    /// <summary>
    /// Writes the rendered template workbook to a caller-owned stream.
    /// </summary>
    public void SaveTo(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        ExcelWorkbookRenderer.Render(Compile(), stream);
    }

    /// <summary>
    /// Builds the rendered template workbook into a byte array.
    /// </summary>
    public byte[] ToBytes()
    {
        using var stream = new MemoryStream();
        SaveTo(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Validates the template and currently bound sources without rendering the workbook.
    /// </summary>
    public ExcelTemplateValidationResult Validate()
    {
        var errors = new List<ExcelTemplateValidationError>();
        try
        {
            Compile();
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException or NotSupportedException)
        {
            errors.Add(new ExcelTemplateValidationError("$", exception.GetType().Name, exception.Message));
        }

        return new ExcelTemplateValidationResult(errors.Count == 0, errors);
    }

    /// <summary>
    /// Reports the cells each cell-occupying region would take up, without building the workbook.
    /// </summary>
    public ExcelTemplatePreviewResult Preview()
    {
        using var document = JsonDocument.Parse(json, JsonOptions);
        var root = document.RootElement;
        ExcelTemplateDataSources.Validate(root, sources);

        // Bounds do not depend on styles, so previews do not read the style catalogue.
        var templateContext = new ExcelTemplateContext(sources, parameters, new Dictionary<string, ExcelStyleSpec>());
        var sheets = new List<ExcelTemplateSheetPreview>();

        foreach (var (sheetName, regionElements) in ReadSheets(root))
        {
            var regions = new List<ExcelTemplateRegionPreview>();
            TemplateRegionBounds? previous = null;
            foreach (var region in regionElements.EnumerateArray())
            {
                var compiler = ExcelTemplateRegions.Find(region.GetRequiredString("type"));
                if (compiler?.OccupiesCells != true)
                {
                    continue;
                }

                var placement = ExcelTemplateLayout.ResolvePlacement(region, compiler, templateContext, previous);
                regions.Add(new ExcelTemplateRegionPreview(
                    compiler.Type,
                    ExcelCellAddress.ColumnNumberToLetter(placement.Bounds.StartColumn) + placement.Bounds.StartRow,
                    ExcelCellAddress.ColumnNumberToLetter(placement.Bounds.EndColumn) + placement.Bounds.EndRow));
                previous = placement.Bounds;
            }

            sheets.Add(new ExcelTemplateSheetPreview(sheetName, regions));
        }

        return new ExcelTemplatePreviewResult(sheets);
    }

    private ExcelWorkbookSpec Compile()
    {
        using var document = JsonDocument.Parse(json, JsonOptions);
        var root = document.RootElement;
        var templateContext = new ExcelTemplateContext(sources, parameters, ExcelTemplateStyles.Read(root));
        ExcelTemplateDataSources.Validate(root, sources);

        var sheets = new List<ExcelSheetSpec>();
        foreach (var (sheetName, regionElements) in ReadSheets(root))
        {
            var sheet = new ExcelSheetSpec(sheetName);
            CompileRegions(sheet, regionElements, templateContext);
            sheets.Add(sheet);
        }

        var workbook = new ExcelWorkbookSpec(sheets);
        if (root.TryGetProperty("workbookProtection", out var workbookProtection))
        {
            workbook.Protection = new ExcelWorkbookProtectionSpec(
                workbookProtection.GetOptionalBool("lockStructure") ?? true,
                workbookProtection.GetOptionalString("password"));
        }

        return workbook;
    }

    /// <summary>
    /// Pairs each output sheet with the "regions" array of the sheet template it names.
    /// </summary>
    private static IEnumerable<(string Name, JsonElement Regions)> ReadSheets(JsonElement root)
    {
        var sheetTemplates = root.GetProperty("sheetTemplates");
        foreach (var sheetElement in root.GetProperty("sheets").EnumerateArray())
        {
            var sheetName = sheetElement.GetRequiredString("name");
            var templateName = sheetElement.GetRequiredString("template");
            if (!sheetTemplates.TryGetProperty(templateName, out var templateElement))
            {
                throw new InvalidOperationException($"Sheet template '{templateName}' was not found.");
            }

            yield return (sheetName, templateElement.GetProperty("regions"));
        }
    }

    private static void CompileRegions(ExcelSheetSpec sheet, JsonElement regions, ExcelTemplateContext templateContext)
    {
        var occupied = new List<TemplateRegionBounds>();
        TemplateRegionBounds? previousBounds = null;
        foreach (var region in regions.EnumerateArray())
        {
            var type = region.GetRequiredString("type");
            var compiler = ExcelTemplateRegions.Find(type);
            var placement = ExcelTemplateLayout.ResolvePlacement(region, compiler, templateContext, previousBounds);
            if (compiler?.OccupiesCells == true)
            {
                ExcelTemplateLayout.ValidateNoOverlap(occupied, placement.Bounds);
                previousBounds = placement.Bounds;
            }

            if (compiler is null)
            {
                throw new NotSupportedException($"Template region type '{type}' is not supported by the initial Excel template compiler.");
            }

            compiler.Compile(sheet, region, templateContext, placement.StartCell);
        }
    }
}
