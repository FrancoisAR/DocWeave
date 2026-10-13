using System.Collections.ObjectModel;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Controls;
using DocWeave.Csv;
using DocWeave.Excel;
using DocWeave.WebApi;
using Microsoft.Win32;

namespace DocWeave.Tool.Ui;

public partial class MainWindow
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    private readonly ObservableCollection<ColumnModel> columns = [];
    private string? currentTemplatePath;
    private bool loaded;

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        if (loaded)
        {
            return;
        }

        loaded = true;
        ColumnsGrid.ItemsSource = columns;
        columns.Add(new ColumnModel("Name", "string", true));
        columns.Add(new ColumnModel("Amount", "number", false));
        TemplateJsonBox.Text = BuildTemplateJson().ToJsonString(JsonOptions);
        SampleDataBox.Text = """
        {
          "rows": [
            { "Name": "North", "Amount": 10 },
            { "Name": "South", "Amount": 20 }
          ]
        }
        """;
    }

    private void NewTemplate_Click(object sender, RoutedEventArgs e)
    {
        currentTemplatePath = null;
        CurrentTemplatePathText.Text = "Unsaved";
        TemplateJsonBox.Text = BuildTemplateJson().ToJsonString(JsonOptions);
        ClearResults();
    }

    private void OpenTemplate_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*" };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        currentTemplatePath = dialog.FileName;
        CurrentTemplatePathText.Text = currentTemplatePath;
        TemplateJsonBox.Text = File.ReadAllText(dialog.FileName);
        SyncDesignerFromJson();
    }

    private void SaveTemplate_Click(object sender, RoutedEventArgs e)
    {
        if (currentTemplatePath is null)
        {
            SaveTemplateAs_Click(sender, e);
            return;
        }

        File.WriteAllText(currentTemplatePath, FormatJson(TemplateJsonBox.Text));
        TemplateJsonBox.Text = File.ReadAllText(currentTemplatePath);
    }

    private void SaveTemplateAs_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new SaveFileDialog { Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*", FileName = "template.json" };
        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        currentTemplatePath = dialog.FileName;
        CurrentTemplatePathText.Text = currentTemplatePath;
        SaveTemplate_Click(sender, e);
    }

    private void OpenData_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog { Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*" };
        if (dialog.ShowDialog(this) == true)
        {
            SampleDataBox.Text = File.ReadAllText(dialog.FileName);
        }
    }

    private void ApplyDesigner_Click(object sender, RoutedEventArgs e)
    {
        TemplateJsonBox.Text = BuildTemplateJson().ToJsonString(JsonOptions);
    }

    private void SyncDesigner_Click(object sender, RoutedEventArgs e)
    {
        SyncDesignerFromJson();
    }

    private void ValidateTemplate_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var notifications = ValidateCurrentTemplate();
            ValidationGrid.ItemsSource = notifications;
            var errors = notifications.Count(item => item.Severity == "Error");
            ValidationSummaryText.Text = errors == 0 ? "Template is valid." : $"{errors} validation issue(s).";
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException or IOException)
        {
            ShowException(exception);
        }
    }

    private void PreviewTemplate_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            PreviewBox.Text = JsonSerializer.Serialize(BindExcelExport().Preview(), JsonOptions);
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException)
        {
            ShowException(exception);
        }
    }

    private void RenderTemplate_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var format = SelectedFormat();
            var dialog = new SaveFileDialog
            {
                Filter = format == "csv" ? "CSV files (*.csv)|*.csv" : "Excel workbooks (*.xlsx)|*.xlsx",
                FileName = format == "csv" ? "report.csv" : "report.xlsx"
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            if (format == "csv")
            {
                File.WriteAllText(dialog.FileName, BindCsvExport().ToText());
            }
            else
            {
                File.WriteAllBytes(dialog.FileName, BindExcelExport().ToBytes());
            }
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException or IOException)
        {
            ShowException(exception);
        }
    }

    private void BuildRequest_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var request = new DocWeaveExportRequest(
                SelectedFormat(),
                TemplateId: null,
                TemplateJson: TemplateJsonBox.Text,
                DataSources: ReadJsonSection(SampleDataBox.Text, "dataSources"),
                Parameters: ReadJsonSection(SampleDataBox.Text, "parameters"),
                Delivery: new DocWeaveExportDeliveryRequest("download", FileNameBox.Text),
                Audit: new DocWeaveApiAuditMetadata(ClientId: "docweave-tool-ui"));
            RequestJsonBox.Text = JsonSerializer.Serialize(request, JsonOptions);
        }
        catch (Exception exception) when (exception is JsonException or InvalidOperationException or ArgumentException)
        {
            ShowException(exception);
        }
    }

    private JsonObject BuildTemplateJson()
    {
        return SelectedKind() switch
        {
            "csv-export" => BuildCsvExportTemplate(),
            "excel-import" => BuildExcelImportTemplate(),
            "csv-import" => BuildCsvImportTemplate(),
            _ => BuildExcelExportTemplate()
        };
    }

    private JsonObject BuildExcelExportTemplate()
    {
        var source = SourceNameBox.Text.Trim();
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "excelWorkbook",
            ["output"] = new JsonObject { ["format"] = "xlsx", ["fileName"] = FileNameBox.Text },
            ["dataSources"] = new JsonArray
            {
                new JsonObject
                {
                    ["name"] = source,
                    ["kind"] = "array",
                    ["required"] = true,
                    ["columns"] = new JsonArray(columns.Where(IsNamed).Select(ToDataSourceColumn).ToArray<JsonNode?>())
                }
            },
            ["styles"] = DefaultStyles(),
            ["sheetTemplates"] = new JsonObject
            {
                ["main"] = new JsonObject
                {
                    ["regions"] = new JsonArray
                    {
                        new JsonObject { ["type"] = "cell", ["cell"] = "A1", ["value"] = TitleBox.Text, ["style"] = "title" },
                        new JsonObject
                        {
                            ["type"] = "table",
                            ["source"] = source,
                            ["startCell"] = "A3",
                            ["columns"] = new JsonArray(columns.Where(IsNamed).Select(ToTableColumn).ToArray<JsonNode?>()),
                            ["styles"] = new JsonObject { ["header"] = "header" }
                        }
                    }
                }
            },
            ["sheets"] = new JsonArray { new JsonObject { ["name"] = SheetNameBox.Text, ["template"] = "main" } }
        };
    }

    private JsonObject BuildCsvExportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "csv",
            ["output"] = new JsonObject
            {
                ["format"] = "csv",
                ["fileName"] = FileNameBox.Text,
                ["delimiter"] = ",",
                ["includeHeaders"] = true,
                ["neutralizeFormulaText"] = true
            },
            ["source"] = SourceNameBox.Text.Trim(),
            ["columns"] = new JsonArray(columns.Where(IsNamed).Select(ToTableColumn).ToArray<JsonNode?>())
        };
    }

    private JsonObject BuildExcelImportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "excelImport",
            ["input"] = new JsonObject { ["sheet"] = SheetNameBox.Text, ["startCell"] = "A1", ["hasHeaderRow"] = true },
            ["columns"] = new JsonObject { ["includeHeaders"] = new JsonArray(columns.Where(IsNamed).Select(column => (JsonNode?)column.Name).ToArray()) },
            ["output"] = new JsonObject { ["columns"] = new JsonArray(columns.Where(IsNamed).Select(ToImportColumn).ToArray<JsonNode?>()) }
        };
    }

    private JsonObject BuildCsvImportTemplate()
    {
        return new JsonObject
        {
            ["schemaVersion"] = "1.0",
            ["kind"] = "csvImport",
            ["input"] = new JsonObject { ["delimiter"] = ",", ["hasHeaderRow"] = true, ["trimFields"] = true, ["skipEmptyLines"] = true },
            ["columns"] = new JsonObject { ["includeHeaders"] = new JsonArray(columns.Where(IsNamed).Select(column => (JsonNode?)column.Name).ToArray()) },
            ["output"] = new JsonObject { ["columns"] = new JsonArray(columns.Where(IsNamed).Select(ToImportColumn).ToArray<JsonNode?>()) }
        };
    }

    private void SyncDesignerFromJson()
    {
        using var document = JsonDocument.Parse(TemplateJsonBox.Text, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.TryGetProperty("kind", out var kind))
        {
            SelectCombo(KindBox, kind.GetString() ?? "excel-export");
        }

        if (root.TryGetProperty("output", out var output))
        {
            if (output.TryGetProperty("format", out var format))
            {
                SelectCombo(FormatBox, format.GetString() ?? "xlsx");
            }

            if (output.TryGetProperty("fileName", out var fileName))
            {
                FileNameBox.Text = fileName.GetString() ?? FileNameBox.Text;
            }
        }

        columns.Clear();
        if (TryReadColumns(root, out var readColumns))
        {
            foreach (var column in readColumns)
            {
                columns.Add(column);
            }
        }
    }

    private IReadOnlyList<ValidationNotification> ValidateCurrentTemplate()
    {
        if (SelectedKind().Contains("import", StringComparison.OrdinalIgnoreCase))
        {
            return [new ValidationNotification("Info", "SourceRequired", "Import validation needs a source file; this first UI validates export templates locally.", "$")];
        }

        if (SelectedFormat() == "csv")
        {
            var result = BindCsvExport().Validate();
            return result.Errors.Select(error => new ValidationNotification("Error", error.Code, error.Message, error.Path)).ToList();
        }

        var excel = BindExcelExport().Validate();
        return excel.Errors.Select(error => new ValidationNotification("Error", error.Code, error.Message, error.Path)).ToList();
    }

    private ExcelTemplateExportBuilder BindExcelExport()
    {
        var builder = ExcelTemplateWriter.FromJson(TemplateJsonBox.Text);
        foreach (var (name, rows) in ReadRows(SampleDataBox.Text))
        {
            builder.WithRows(name, rows);
        }

        foreach (var (name, value) in ReadScalars(SampleDataBox.Text))
        {
            builder.WithParameter(name, value);
        }

        return builder;
    }

    private CsvTemplateExportBuilder BindCsvExport()
    {
        var builder = CsvTemplateWriter.FromJson(TemplateJsonBox.Text);
        foreach (var (name, rows) in ReadRows(SampleDataBox.Text))
        {
            builder.WithRows(name, rows);
        }

        return builder;
    }

    private static IReadOnlyDictionary<string, JsonElement> ReadJsonSection(string sampleData, string wrapperName)
    {
        using var document = JsonDocument.Parse(sampleData, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty(wrapperName, out var wrapped))
        {
            root = wrapped;
        }

        return root.ValueKind == JsonValueKind.Object
            ? root.EnumerateObject().ToDictionary(property => property.Name, property => property.Value.Clone(), StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
    }

    private static IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> ReadRows(string sampleData)
    {
        using var document = JsonDocument.Parse(sampleData, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("dataSources", out var wrapped))
        {
            root = wrapped;
        }

        var result = new Dictionary<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>(StringComparer.OrdinalIgnoreCase);
        foreach (var source in root.EnumerateObject())
        {
            if (source.Value.ValueKind == JsonValueKind.Array)
            {
                result[source.Name] = source.Value.EnumerateArray().Select(ReadRow).Cast<IReadOnlyDictionary<string, object?>>().ToList();
            }
        }

        return result;
    }

    private static IReadOnlyDictionary<string, object?> ReadScalars(string sampleData)
    {
        using var document = JsonDocument.Parse(sampleData, new JsonDocumentOptions { AllowTrailingCommas = true });
        var root = document.RootElement;
        if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("parameters", out var wrapped))
        {
            root = wrapped;
        }

        return root.ValueKind == JsonValueKind.Object
            ? root.EnumerateObject()
                .Where(property => property.Value.ValueKind is not (JsonValueKind.Array or JsonValueKind.Object))
                .ToDictionary(property => property.Name, property => ConvertJsonValue(property.Value), StringComparer.OrdinalIgnoreCase)
            : new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
    }

    private static Dictionary<string, object?> ReadRow(JsonElement item)
    {
        return item.EnumerateObject().ToDictionary(property => property.Name, property => ConvertJsonValue(property.Value), StringComparer.OrdinalIgnoreCase);
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

    private static bool TryReadColumns(JsonElement root, out IReadOnlyList<ColumnModel> result)
    {
        result = [];
        if (root.TryGetProperty("columns", out var columnsElement) && columnsElement.ValueKind == JsonValueKind.Array)
        {
            result = columnsElement.EnumerateArray().Select(ReadTemplateColumn).ToList();
            return true;
        }

        if (root.TryGetProperty("dataSources", out var dataSources)
            && dataSources.ValueKind == JsonValueKind.Array
            && dataSources.EnumerateArray().FirstOrDefault().TryGetProperty("columns", out columnsElement))
        {
            result = columnsElement.EnumerateArray().Select(ReadDataSourceColumn).ToList();
            return true;
        }

        return false;
    }

    private static ColumnModel ReadTemplateColumn(JsonElement column)
    {
        var name = column.TryGetProperty("source", out var source) ? source.GetString() : column.GetProperty("name").GetString();
        return new ColumnModel(name ?? "Column", "string", false);
    }

    private static ColumnModel ReadDataSourceColumn(JsonElement column)
    {
        return new ColumnModel(
            column.GetProperty("name").GetString() ?? "Column",
            column.TryGetProperty("type", out var type) ? type.GetString() ?? "string" : "string",
            column.TryGetProperty("required", out var required) && required.GetBoolean());
    }

    private static JsonObject DefaultStyles()
    {
        return new JsonObject
        {
            ["title"] = new JsonObject { ["font"] = new JsonObject { ["size"] = 16, ["bold"] = true } },
            ["header"] = new JsonObject
            {
                ["font"] = new JsonObject { ["bold"] = true, ["color"] = "#FFFFFF" },
                ["fill"] = new JsonObject { ["color"] = "#1F4E78" }
            }
        };
    }

    private static bool IsNamed(ColumnModel column) => !string.IsNullOrWhiteSpace(column.Name);

    private static JsonObject ToDataSourceColumn(ColumnModel column) => new() { ["name"] = column.Name, ["type"] = column.Type, ["required"] = column.Required };

    private static JsonObject ToTableColumn(ColumnModel column) => new() { ["source"] = column.Name, ["caption"] = column.Name };

    private static JsonObject ToImportColumn(ColumnModel column) => new() { ["name"] = column.Name, ["source"] = column.Name, ["type"] = column.Type, ["required"] = column.Required };

    private string SelectedKind() => ((ComboBoxItem)KindBox.SelectedItem).Content.ToString() ?? "excel-export";

    private string SelectedFormat() => ((ComboBoxItem)FormatBox.SelectedItem).Content.ToString() ?? "xlsx";

    private static void SelectCombo(ComboBox comboBox, string value)
    {
        foreach (ComboBoxItem item in comboBox.Items)
        {
            if (string.Equals(item.Content.ToString(), value, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedItem = item;
                return;
            }
        }
    }

    private static string FormatJson(string json) => (JsonNode.Parse(json) ?? throw new InvalidOperationException("JSON is empty.")).ToJsonString(JsonOptions);

    private void ClearResults()
    {
        ValidationGrid.ItemsSource = null;
        ValidationSummaryText.Text = string.Empty;
        PreviewBox.Clear();
        RequestJsonBox.Clear();
    }

    private void ShowException(Exception exception)
    {
        ValidationGrid.ItemsSource = new[] { new ValidationNotification("Error", exception.GetType().Name, exception.Message, "$") };
        ValidationSummaryText.Text = "Validation failed.";
    }

    public sealed record ColumnModel(string Name, string Type, bool Required);

    public sealed record ValidationNotification(string Severity, string Code, string Message, string? Path);
}
