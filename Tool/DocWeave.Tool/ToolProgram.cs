using System.Text.Json;
using System.Text.Json.Nodes;
using DocWeave.Csv;
using DocWeave.Excel;
using DocWeave.WebApi;

namespace DocWeave.Tool;

public static class ToolProgram
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static int Run(string[] args)
    {
        try
        {
            if (args.Length == 0 || Has(args, "--help") || Has(args, "-h"))
            {
                WriteHelp();
                return 0;
            }

            return args[0].ToLowerInvariant() switch
            {
                "template" => RunTemplate(args.Skip(1).ToArray()),
                "request" => RunRequest(args.Skip(1).ToArray()),
                _ => Fail($"Unknown command '{args[0]}'.")
            };
        }
        catch (Exception exception) when (exception is IOException or JsonException or InvalidOperationException or ArgumentException)
        {
            Console.Error.WriteLine($"error: {exception.Message}");
            return 1;
        }
    }

    private static int RunTemplate(string[] args)
    {
        if (args.Length == 0)
        {
            return Fail("Missing template command.");
        }

        return args[0].ToLowerInvariant() switch
        {
            "new" => CreateTemplate(args.Skip(1).ToArray()),
            "save" or "format" => SaveTemplate(args.Skip(1).ToArray()),
            "validate" => ValidateTemplate(args.Skip(1).ToArray()),
            "preview" => PreviewTemplate(args.Skip(1).ToArray()),
            "render" => RenderTemplate(args.Skip(1).ToArray()),
            _ => Fail($"Unknown template command '{args[0]}'.")
        };
    }

    private static int RunRequest(string[] args)
    {
        if (args.Length == 0)
        {
            return Fail("Missing request command.");
        }

        return args[0].ToLowerInvariant() switch
        {
            "build" => BuildRequest(args.Skip(1).ToArray()),
            _ => Fail($"Unknown request command '{args[0]}'.")
        };
    }

    private static int CreateTemplate(string[] args)
    {
        var kind = Require(args, "--kind");
        var output = Require(args, "--out");
        var template = kind.ToLowerInvariant() switch
        {
            "excel-export" or "xlsx-export" => NewExcelExportTemplate(),
            "csv-export" => NewCsvExportTemplate(),
            "excel-import" or "xlsx-import" => NewExcelImportTemplate(),
            "csv-import" => NewCsvImportTemplate(),
            _ => throw new ArgumentException($"Unsupported template kind '{kind}'.")
        };

        WriteJson(output, template);
        Console.WriteLine($"Created {kind} template: {output}");
        return 0;
    }

    private static int SaveTemplate(string[] args)
    {
        var templatePath = Require(args, "--template");
        var output = Value(args, "--out") ?? templatePath;
        var node = JsonNode.Parse(File.ReadAllText(templatePath))
            ?? throw new InvalidOperationException("Template JSON is empty.");

        WriteJson(output, node);
        Console.WriteLine($"Saved formatted template: {output}");
        return 0;
    }

    private static int ValidateTemplate(string[] args)
    {
        var templatePath = Require(args, "--template");
        var format = ResolveFormat(args, templatePath);
        var template = File.ReadAllText(templatePath);

        return format switch
        {
            "xlsx" or "excel" => ValidateExcel(template, args),
            "csv" => ValidateCsv(template, args),
            _ => Fail($"Unsupported format '{format}'.")
        };
    }

    private static int ValidateExcel(string template, string[] args)
    {
        if (IsImportTemplate(template))
        {
            var source = Require(args, "--source");
            var result = ExcelTemplateReader.FromJson(template).FromFile(source).Validate();
            return WriteValidation(result.IsValid, result.Errors.Select(error => Notification("error", error.Code, error.Message, error.Path)));
        }

        var validation = BindExcelExport(template, Value(args, "--data")).Validate();
        return WriteValidation(validation.IsValid, validation.Errors.Select(error => Notification("error", error.Code, error.Message, error.Path)));
    }

    private static int ValidateCsv(string template, string[] args)
    {
        if (IsImportTemplate(template))
        {
            var source = Require(args, "--source");
            var result = CsvTemplateReader.FromJson(template).FromText(File.ReadAllText(source), Path.GetFileName(source)).Validate();
            return WriteValidation(result.IsValid, result.Errors.Select(error => Notification("error", error.Code, error.Message, error.Path)));
        }

        var validation = BindCsvExport(template, Value(args, "--data")).Validate();
        return WriteValidation(validation.IsValid, validation.Errors.Select(error => Notification("error", error.Code, error.Message, error.Path)));
    }

    private static int PreviewTemplate(string[] args)
    {
        var templatePath = Require(args, "--template");
        var template = File.ReadAllText(templatePath);
        var preview = BindExcelExport(template, Value(args, "--data")).Preview();
        Console.WriteLine(JsonSerializer.Serialize(preview, JsonOptions));
        return 0;
    }

    private static int RenderTemplate(string[] args)
    {
        var templatePath = Require(args, "--template");
        var output = Require(args, "--out");
        var format = ResolveFormat(args, templatePath);
        var template = File.ReadAllText(templatePath);

        if (format is "xlsx" or "excel")
        {
            File.WriteAllBytes(output, BindExcelExport(template, Value(args, "--data")).ToBytes());
        }
        else if (format == "csv")
        {
            File.WriteAllText(output, BindCsvExport(template, Value(args, "--data")).ToText());
        }
        else
        {
            return Fail($"Unsupported format '{format}'.");
        }

        Console.WriteLine($"Rendered sample output: {output}");
        return 0;
    }

    private static int BuildRequest(string[] args)
    {
        var templatePath = Require(args, "--template");
        var output = Require(args, "--out");
        var format = ResolveFormat(args, templatePath);
        var templateId = Value(args, "--template-id");
        var templateJson = templateId is null ? File.ReadAllText(templatePath) : null;
        var data = ReadRequestSection(Value(args, "--data"), "dataSources");
        var parameters = ReadRequestSection(Value(args, "--parameters"), "parameters");
        var fileName = Value(args, "--file-name") ?? DefaultFileName(format);

        var request = new DocWeaveExportRequest(
            Format: NormalizeFormat(format),
            TemplateId: templateId,
            TemplateJson: templateJson,
            DataSources: data,
            Parameters: parameters.Count == 0 ? null : parameters,
            Delivery: new DocWeaveExportDeliveryRequest("download", fileName),
            Audit: new DocWeaveApiAuditMetadata(ClientId: "docweave-tool"));

        WriteJson(output, JsonSerializer.SerializeToNode(request, JsonOptions)!);
        Console.WriteLine($"Built export request: {output}");
        return 0;
    }

    private static ExcelTemplateExportBuilder BindExcelExport(string template, string? dataPath)
    {
        var builder = ExcelTemplateWriter.FromJson(template);
        foreach (var (name, rows) in ReadRows(dataPath))
        {
            builder.WithRows(name, rows);
        }

        foreach (var (name, value) in ReadScalars(dataPath))
        {
            builder.WithParameter(name, value);
        }

        return builder;
    }

    private static CsvTemplateExportBuilder BindCsvExport(string template, string? dataPath)
    {
        var builder = CsvTemplateWriter.FromJson(template);
        foreach (var (name, rows) in ReadRows(dataPath))
        {
            builder.WithRows(name, rows);
        }

        return builder;
    }

    private static IReadOnlyDictionary<string, JsonElement> ReadRequestSection(string? path, string wrapperName)
    {
        if (path is null)
        {
            return new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
        }

        using var document = JsonDocument.Parse(File.ReadAllText(path), new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty(wrapperName, out var wrapped))
        {
            root = wrapped;
        }

        return root.EnumerateObject().ToDictionary(
            property => property.Name,
            property => property.Value.Clone(),
            StringComparer.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> ReadRows(string? path)
    {
        if (path is null)
        {
            return new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
        }

        using var document = JsonDocument.Parse(File.ReadAllText(path), new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("dataSources", out var wrapped))
        {
            root = wrapped;
        }

        var result = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in root.EnumerateObject())
        {
            if (source.Value.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            result[source.Name] = source.Value.EnumerateArray()
                .Select(ReadRow)
                .Cast<IReadOnlyDictionary<string, object?>>()
                .ToList();
        }

        return result;
    }

    private static IReadOnlyDictionary<string, object?> ReadScalars(string? path)
    {
        if (path is null)
        {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        }

        using var document = JsonDocument.Parse(File.ReadAllText(path), new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("parameters", out var wrapped))
        {
            root = wrapped;
        }

        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in root.EnumerateObject())
        {
            if (property.Value.ValueKind != JsonValueKind.Array && property.Value.ValueKind != JsonValueKind.Object)
            {
                result[property.Name] = ConvertJsonValue(property.Value);
            }
        }

        return result;
    }

    private static Dictionary<string, object?> ReadRow(JsonElement item)
    {
        if (item.ValueKind != JsonValueKind.Object)
        {
            throw new InvalidOperationException("Data source arrays must contain JSON objects.");
        }

        return item.EnumerateObject().ToDictionary(
            property => property.Name,
            property => ConvertJsonValue(property.Value),
            StringComparer.OrdinalIgnoreCase);
    }

    private static object? ConvertJsonValue(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString(),
            JsonValueKind.Number when value.TryGetInt64(out var integer) => integer,
            JsonValueKind.Number => value.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => value.GetRawText()
        };
    }

    private static JsonObject NewExcelExportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "excelWorkbook",
            ["output"] = new JsonObject { ["format"] = "xlsx", ["fileName"] = "report.xlsx" },
            ["dataSources"] = new JsonArray
            {
                new JsonObject
                {
                    ["name"] = "rows",
                    ["kind"] = "array",
                    ["required"] = true,
                    ["columns"] = new JsonArray
                    {
                        Column("Name", "string", required: true),
                        Column("Amount", "number", required: false)
                    }
                }
            },
            ["styles"] = new JsonObject
            {
                ["title"] = new JsonObject { ["font"] = new JsonObject { ["size"] = 16, ["bold"] = true } },
                ["header"] = new JsonObject
                {
                    ["font"] = new JsonObject { ["bold"] = true, ["color"] = "#FFFFFF" },
                    ["fill"] = new JsonObject { ["color"] = "#1F4E78" }
                }
            },
            ["sheetTemplates"] = new JsonObject
            {
                ["main"] = new JsonObject
                {
                    ["regions"] = new JsonArray
                    {
                        new JsonObject { ["type"] = "cell", ["cell"] = "A1", ["value"] = "Report", ["style"] = "title" },
                        new JsonObject
                        {
                            ["type"] = "table",
                            ["source"] = "rows",
                            ["startCell"] = "A3",
                            ["columns"] = new JsonArray
                            {
                                new JsonObject { ["source"] = "Name", ["caption"] = "Name" },
                                new JsonObject { ["source"] = "Amount", ["caption"] = "Amount" }
                            },
                            ["styles"] = new JsonObject { ["header"] = "header" }
                        }
                    }
                }
            },
            ["sheets"] = new JsonArray { new JsonObject { ["name"] = "Report", ["template"] = "main" } }
        };
    }

    private static JsonObject NewCsvExportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "csv",
            ["output"] = new JsonObject
            {
                ["format"] = "csv",
                ["fileName"] = "report.csv",
                ["delimiter"] = ",",
                ["includeHeaders"] = true,
                ["neutralizeFormulaText"] = true
            },
            ["source"] = "rows",
            ["columns"] = new JsonArray
            {
                new JsonObject { ["source"] = "Name", ["caption"] = "Name" },
                new JsonObject { ["source"] = "Amount", ["caption"] = "Amount" }
            }
        };
    }

    private static JsonObject NewExcelImportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "excelImport",
            ["input"] = new JsonObject { ["sheet"] = "Sheet1", ["startCell"] = "A1", ["hasHeaderRow"] = true },
            ["columns"] = new JsonObject { ["includeHeaders"] = new JsonArray("Name", "Amount") },
            ["output"] = new JsonObject
            {
                ["columns"] = new JsonArray
                {
                    ImportColumn("Name", "Name", "string", required: true),
                    ImportColumn("Amount", "Amount", "decimal", required: false)
                }
            }
        };
    }

    private static JsonObject NewCsvImportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "csvImport",
            ["input"] = new JsonObject
            {
                ["delimiter"] = ",",
                ["hasHeaderRow"] = true,
                ["trimFields"] = true,
                ["skipEmptyLines"] = true
            },
            ["columns"] = new JsonObject { ["includeHeaders"] = new JsonArray("Name", "Amount") },
            ["output"] = new JsonObject
            {
                ["columns"] = new JsonArray
                {
                    ImportColumn("Name", "Name", "string", required: true),
                    ImportColumn("Amount", "Amount", "decimal", required: false)
                }
            }
        };
    }

    private static JsonObject Column(string name, string type, bool required)
    {
        return new JsonObject { ["name"] = name, ["type"] = type, ["required"] = required };
    }

    private static JsonObject ImportColumn(string name, string source, string type, bool required)
    {
        return new JsonObject { ["name"] = name, ["source"] = source, ["type"] = type, ["required"] = required };
    }

    private static bool IsImportTemplate(string template)
    {
        using var document = JsonDocument.Parse(template, new JsonDocumentOptions { AllowTrailingCommas = true });
        return document.RootElement.TryGetProperty("kind", out var kind)
            && (kind.GetString()?.Contains("import", StringComparison.OrdinalIgnoreCase) ?? false);
    }

    private static string ResolveFormat(string[] args, string templatePath)
    {
        var format = Value(args, "--format");
        if (!string.IsNullOrWhiteSpace(format))
        {
            return format;
        }

        using var document = JsonDocument.Parse(File.ReadAllText(templatePath), new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.TryGetProperty("output", out var output)
            && output.TryGetProperty("format", out var outputFormat))
        {
            return outputFormat.GetString() ?? "xlsx";
        }

        if (root.TryGetProperty("kind", out var kind))
        {
            var kindValue = kind.GetString() ?? string.Empty;
            if (kindValue.Contains("csv", StringComparison.OrdinalIgnoreCase))
            {
                return "csv";
            }
        }

        return "xlsx";
    }

    private static string NormalizeFormat(string format)
    {
        return format is "excel" ? "xlsx" : format;
    }

    private static string DefaultFileName(string format)
    {
        return NormalizeFormat(format) == "csv" ? "report.csv" : "report.xlsx";
    }

    private static object Notification(string severity, string code, string message, string? path)
    {
        return new { severity, code, message, path };
    }

    private static int WriteValidation(bool isValid, IEnumerable<object> notifications)
    {
        Console.WriteLine(JsonSerializer.Serialize(new { isValid, notifications = notifications.ToList() }, JsonOptions));
        return isValid ? 0 : 2;
    }

    private static void WriteJson(string path, JsonNode node)
    {
        var directory = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(path, node.ToJsonString(JsonOptions));
    }

    private static string? Value(string[] args, string name)
    {
        for (var i = 0; i < args.Length; i++)
        {
            if (string.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
            {
                return i + 1 < args.Length ? args[i + 1] : throw new ArgumentException($"Missing value for {name}.");
            }
        }

        return null;
    }

    private static string Require(string[] args, string name)
    {
        return Value(args, name) ?? throw new ArgumentException($"Missing required option {name}.");
    }

    private static bool Has(string[] args, string name)
    {
        return args.Any(arg => string.Equals(arg, name, StringComparison.OrdinalIgnoreCase));
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine($"error: {message}");
        return 1;
    }

    private static void WriteHelp()
    {
        Console.WriteLine("""
        DocWeave template tool

        Commands:
          template new      --kind excel-export|csv-export|excel-import|csv-import --out template.json
          template save     --template template.json [--out template.json]
          template validate --template template.json [--format xlsx|csv] [--data sample-data.json] [--source source-file]
          template preview  --template template.json [--data sample-data.json]
          template render   --template template.json --out report.xlsx [--format xlsx|csv] [--data sample-data.json]
          request build     --template template.json --out request.json [--format xlsx|csv] [--data sample-data.json]

        Sample data may be either:
          { "rows": [ { "Name": "North", "Amount": 10 } ] }
        or:
          { "dataSources": { "rows": [ ... ] }, "parameters": { "reportDate": "2026-10-08" } }
        """);
    }
}
