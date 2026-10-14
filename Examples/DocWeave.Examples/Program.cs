using System.Data;
using System.Text.Json;
using DocWeave.Csv;
using DocWeave.Excel;
using DocWeave.WebApi;

var outputRoot = args.Length > 0
    ? Path.GetFullPath(args[0])
    : Path.Combine(AppContext.BaseDirectory, "DocWeave.Examples.Output");

Directory.CreateDirectory(outputRoot);

Console.WriteLine("DocWeave examples");
Console.WriteLine($"Writing examples to: {outputRoot}");
Console.WriteLine();

// Template examples are grouped by format. Each one writes:
// - template.json: the declarative template used by DocWeave.
// - api-request.json: the equivalent Web API request payload.
// - sample input/output files that make the scenario runnable.
Run("template-xlsx-01-styled-sales-export", TemplateXlsxStyledSalesExport);
Run("template-xlsx-02-dataset-multiple-sheets", TemplateXlsxDataSetMultipleSheets);
Run("template-xlsx-03-parameters-summary-chart", TemplateXlsxParametersSummaryChart);
Run("template-xlsx-04-import-typed-derived-columns", TemplateXlsxImportTypedDerivedColumns);
Run("template-xlsx-05-import-offset-selected-columns", TemplateXlsxImportOffsetSelectedColumns);

Run("template-csv-01-selected-columns-aliases-delimiter", TemplateCsvSelectedColumnsAliasesDelimiter);
Run("template-csv-02-datatable-without-headers", TemplateCsvDataTableWithoutHeaders);
Run("template-csv-03-dataset-selected-table", TemplateCsvDataSetSelectedTable);
Run("template-csv-04-import-typed-derived-columns", TemplateCsvImportTypedDerivedColumns);
Run("template-csv-05-import-pipe-headerless", TemplateCsvImportPipeHeaderless);

// Direct examples show the fluent/read APIs beside the template examples.
Run("direct-xlsx-01-fluent-styled-export", DirectXlsxFluentStyledExport);
Run("direct-xlsx-02-datatable-workbook-export", DirectXlsxDataTableWorkbookExport);
Run("direct-xlsx-03-selected-column-import", DirectXlsxSelectedColumnImport);
Run("direct-csv-01-datatable-pipe-export", DirectCsvDataTablePipeExport);
Run("direct-csv-02-object-import-with-mapping", DirectCsvObjectImportWithMapping);

Console.WriteLine();
BorderExamples.Run(outputRoot);
Console.WriteLine("Done.");

void Run(string name, Action<string> scenario)
{
    var directory = Path.Combine(outputRoot, name);
    Directory.CreateDirectory(directory);
    scenario(directory);
    Console.WriteLine($"  {name}");
}

// ---------------------------------------------------------------------------
// Template XLSX examples
// ---------------------------------------------------------------------------

void TemplateXlsxStyledSalesExport(string directory)
{
    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "excelWorkbook",
      "styles": {
        "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
        "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#000000" } },
        "data": { "border": { "style": "thin", "color": "#B4C6E7" } },
        "alternate": { "fill": { "color": "#DDEBF7" }, "border": { "style": "thin", "color": "#B4C6E7" } },
        "currency": { "numberFormat": "currency", "border": { "style": "thin", "color": "#70AD47" } },
        "totals": { "font": { "bold": true }, "fill": { "color": "#D9EAD3" }, "border": { "style": "thick", "color": "#000000" } }
      },
      "sheetTemplates": {
        "salesSheet": {
          "regions": [
            { "type": "cell", "cell": "A1", "value": "Template Sales Export", "style": "title" },
            {
              "type": "table",
              "source": "sales",
              "startCell": "A3",
              "columns": [
                { "source": "Region", "caption": "Region" },
                { "source": "Income", "caption": "Income", "style": "currency" },
                { "source": "Costs", "caption": "Costs", "style": "currency" }
              ],
              "styles": { "header": "header", "data": "data", "alternateRow": "alternate" },
              "totals": { "position": "bottom", "label": "Totals", "style": "totals" }
            }
          ]
        }
      },
      "sheets": [ { "name": "Sales", "template": "salesSheet" } ]
    }
    """;

    var rows = SalesRows();
    var bytes = ExcelTemplateWriter.FromJson(template)
        .WithData("sales", rows)
        .ToBytes();

    WriteTemplateExportScenario(directory, "xlsx", "styled-sales.xlsx", template, new Dictionary<string, object?> { ["sales"] = rows });
    File.WriteAllBytes(Path.Combine(directory, "styled-sales.xlsx"), bytes);
}

void TemplateXlsxDataSetMultipleSheets(string directory)
{
    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "excelWorkbook",
      "styleCatalog": {
        "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#385723" }, "border": { "style": "thin" } },
        "data": { "border": { "style": "thin", "color": "#A9D18E" } }
      },
      "sheetTemplates": {
        "salesSheet": {
          "regions": [
            { "type": "table", "source": "Sales", "startCell": "A1", "styles": { "header": "header", "data": "data" } }
          ]
        },
        "customerSheet": {
          "regions": [
            { "type": "table", "source": "Customers", "startCell": "A1", "styles": { "header": "header", "data": "data" } }
          ]
        }
      },
      "sheets": [
        { "name": "Sales", "template": "salesSheet" },
        { "name": "Customers", "template": "customerSheet" }
      ]
    }
    """;

    var dataSet = BusinessDataSet();
    var bytes = ExcelTemplateWriter.FromJson(template)
        .WithDataSet(dataSet)
        .ToBytes();

    WriteTemplateExportScenario(directory, "xlsx", "dataset-sheets.xlsx", template, DataSetSources(dataSet));
    File.WriteAllBytes(Path.Combine(directory, "dataset-sheets.xlsx"), bytes);
}

void TemplateXlsxParametersSummaryChart(string directory)
{
    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "excelWorkbook",
      "styles": {
        "title": { "font": { "size": 16, "bold": true, "color": "#385723" } },
        "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#385723" }, "border": { "style": "thin" } },
        "money": { "numberFormat": "currency", "border": { "style": "thin", "color": "#A9D18E" } },
        "label": { "font": { "bold": true }, "fill": { "color": "#DDEBF7" }, "border": { "style": "thin" } }
      },
      "sheetTemplates": {
        "dashboard": {
          "regions": [
            { "type": "cell", "cell": "A1", "value": "Monthly dashboard for {{parameters.reportDate}}", "style": "title" },
            {
              "type": "table",
              "source": "monthly",
              "startCell": "A3",
              "columns": [
                { "source": "Month", "caption": "Month" },
                { "source": "Income", "caption": "Income", "style": "money" }
              ],
              "styles": { "header": "header", "data": "money" }
            },
            {
              "type": "summary",
              "source": "monthly",
              "startCell": "D3",
              "labelStyle": "label",
              "valueStyle": "money",
              "fields": [
                { "caption": "Total income", "aggregate": "sum", "source": "Income" },
                { "caption": "Average income", "aggregate": "avg", "source": "Income" }
              ]
            },
            {
              "type": "chart",
              "title": "Income {{parameters.reportDate}}",
              "chartType": "column",
              "dataSource": { "sheet": "Dashboard", "categoryRange": "$A$4:$A$9", "valueRange": "$B$4:$B$9" },
              "position": { "topLeft": "A12", "bottomRight": "J28" },
              "style": 10,
              "seriesColor": "#70AD47",
              "showDataLabels": true
            }
          ]
        }
      },
      "sheets": [ { "name": "Dashboard", "template": "dashboard" } ]
    }
    """;

    var rows = MonthlyRows();
    var parameters = new Dictionary<string, object?> { ["reportDate"] = "2026-10" };
    var bytes = ExcelTemplateWriter.FromJson(template)
        .WithData("monthly", rows)
        .WithParameter("reportDate", parameters["reportDate"])
        .ToBytes();

    WriteTemplateExportScenario(directory, "xlsx", "dashboard.xlsx", template, new Dictionary<string, object?> { ["monthly"] = rows }, parameters);
    File.WriteAllBytes(Path.Combine(directory, "dashboard.xlsx"), bytes);
}

void TemplateXlsxImportTypedDerivedColumns(string directory)
{
    var sourcePath = Path.Combine(directory, "import-source.xlsx");
    File.WriteAllBytes(sourcePath, ImportWorkbook());

    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "excelImport",
      "input": {
        "sheet": "Import Source",
        "startCell": "A1",
        "hasHeaderRow": true
      },
      "columns": {
        "includeHeaders": [ "ID", "Name", "Amount" ]
      },
      "output": {
        "columns": [
          { "source": "ID", "name": "Identifier", "type": "int", "required": true, "nullable": false },
          { "source": "Name", "name": "Name", "type": "string", "required": true, "nullable": false },
          { "source": "Amount", "name": "Amount", "type": "decimal", "required": true, "nullable": false }
        ],
        "derivedColumns": [
          { "name": "DisplayName", "concat": [ "Identifier", " - ", "Name" ] },
          { "name": "Source", "constant": "XLSX" }
        ]
      }
    }
    """;

    var table = ExcelTemplateReader.FromJson(template)
        .FromFile(sourcePath)
        .ReadDataTable("ImportedExcel");

    WriteTemplateImportScenario(directory, "xlsx", template, sourcePath);
    WriteJson(Path.Combine(directory, "import-result.json"), DataTableRows(table));
}

void TemplateXlsxImportOffsetSelectedColumns(string directory)
{
    var sourcePath = Path.Combine(directory, "offset-source.xlsx");
    var bytes = ExcelWriter.Create()
        .AddSheet("Offset", sheet =>
        {
            sheet.AddCell("B2", "Region Import", style => style.Bold().FontSize(14));
            sheet.AddTable(new[]
                {
                    new { Region = "North", Manager = "Asha", InternalCode = "N-1", Amount = 125.25m },
                    new { Region = "South", Manager = "Ben", InternalCode = "S-1", Amount = 220.00m }
                })
                .StartAt("B4")
                .HeaderAlias("Amount", "Sales Amount");
        })
        .ToBytes();
    File.WriteAllBytes(sourcePath, bytes);

    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "excelImport",
      "input": {
        "sheet": "Offset",
        "startCell": "B4",
        "hasHeaderRow": true
      },
      "columns": {
        "includeHeaders": [ "Region", "Manager", "Sales Amount" ]
      },
      "output": {
        "columns": [
          { "source": "Region", "name": "Region", "type": "string", "required": true },
          { "source": "Manager", "name": "Owner", "type": "string", "required": true },
          { "source": "Sales Amount", "name": "Amount", "type": "decimal", "required": true }
        ]
      }
    }
    """;

    var rows = ExcelTemplateReader.FromJson(template)
        .FromFile(sourcePath)
        .ReadDictionaries();

    WriteTemplateImportScenario(directory, "xlsx", template, sourcePath);
    WriteJson(Path.Combine(directory, "import-result.json"), rows);
}

// ---------------------------------------------------------------------------
// Template CSV examples
// ---------------------------------------------------------------------------

void TemplateCsvSelectedColumnsAliasesDelimiter(string directory)
{
    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "csv",
      "source": "currencies",
      "output": {
        "delimiter": ";",
        "includeHeaders": true
      },
      "columns": [
        { "source": "Code", "caption": "Currency" },
        { "source": "Country", "caption": "Country" },
        { "source": "LastHighRate", "caption": "High" }
      ]
    }
    """;

    var rows = CurrencyRows();
    var csv = CsvTemplateWriter.FromJson(template)
        .WithData("currencies", rows)
        .ToText();

    WriteTemplateExportScenario(directory, "csv", "currencies.csv", template, new Dictionary<string, object?> { ["currencies"] = rows });
    File.WriteAllText(Path.Combine(directory, "currencies.csv"), csv);
}

void TemplateCsvDataTableWithoutHeaders(string directory)
{
    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "csv",
      "source": "salesTable",
      "output": {
        "delimiter": ",",
        "includeHeaders": false
      },
      "columns": [
        { "source": "Region" },
        { "source": "Amount" }
      ]
    }
    """;

    var table = SalesDataTable("salesTable");
    var csv = CsvTemplateWriter.FromJson(template)
        .WithDataTable("salesTable", table)
        .ToText();

    WriteTemplateExportScenario(directory, "csv", "datatable-no-headers.csv", template, new Dictionary<string, object?> { ["salesTable"] = DataTableRows(table) });
    File.WriteAllText(Path.Combine(directory, "datatable-no-headers.csv"), csv);
}

void TemplateCsvDataSetSelectedTable(string directory)
{
    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "csv",
      "source": "Customers",
      "output": {
        "delimiter": "|",
        "includeHeaders": true
      },
      "columns": [
        { "source": "Name", "caption": "Customer" },
        { "source": "Tier" },
        { "source": "Balance", "caption": "Outstanding" }
      ]
    }
    """;

    var dataSet = BusinessDataSet();
    var csv = CsvTemplateWriter.FromJson(template)
        .WithDataSet(dataSet)
        .ToText();

    WriteTemplateExportScenario(directory, "csv", "dataset-customers.csv", template, DataSetSources(dataSet));
    File.WriteAllText(Path.Combine(directory, "dataset-customers.csv"), csv);
}

void TemplateCsvImportTypedDerivedColumns(string directory)
{
    var sourcePath = Path.Combine(directory, "import-source.csv");
    File.WriteAllText(sourcePath, "ID,Name,Amount,Ignored\n1,Alice,12.50,x\n2,Bob,30.00,y");

    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "csvImport",
      "input": {
        "delimiter": ",",
        "hasHeaderRow": true,
        "trimFields": true
      },
      "columns": {
        "includeHeaders": [ "ID", "Name", "Amount" ]
      },
      "output": {
        "columns": [
          { "source": "ID", "name": "Identifier", "type": "int", "required": true, "nullable": false },
          { "source": "Name", "name": "Name", "type": "string", "required": true, "nullable": false },
          { "source": "Amount", "name": "Amount", "type": "decimal", "required": true, "nullable": false }
        ],
        "derivedColumns": [
          { "name": "DisplayName", "concat": [ "Identifier", " - ", "Name" ] },
          { "name": "Source", "constant": "CSV" }
        ]
      }
    }
    """;

    var table = CsvTemplateReader.FromJson(template)
        .FromFile(sourcePath)
        .ReadDataTable("ImportedCsv");

    WriteTemplateImportScenario(directory, "csv", template, sourcePath);
    WriteJson(Path.Combine(directory, "import-result.json"), DataTableRows(table));
}

void TemplateCsvImportPipeHeaderless(string directory)
{
    var sourcePath = Path.Combine(directory, "pipe-headerless.csv");
    File.WriteAllText(sourcePath, "North|Asha|125.25\nSouth|Ben|220.00");

    var template = """
    {
      "schemaVersion": "1.0",
      "kind": "csvImport",
      "input": {
        "delimiter": "|",
        "hasHeaderRow": false,
        "trimFields": true
      },
      "columns": {
        "includeLetters": [ "A", "B", "C" ]
      },
      "output": {
        "columns": [
          { "source": "A", "name": "Region", "type": "string", "required": true },
          { "source": "B", "name": "Owner", "type": "string", "required": true },
          { "source": "C", "name": "Amount", "type": "decimal", "required": true }
        ]
      }
    }
    """;

    var rows = CsvTemplateReader.FromJson(template)
        .FromFile(sourcePath)
        .ReadDictionaries();

    WriteTemplateImportScenario(directory, "csv", template, sourcePath);
    WriteJson(Path.Combine(directory, "import-result.json"), rows);
}

// ---------------------------------------------------------------------------
// Direct XLSX and CSV examples
// ---------------------------------------------------------------------------

void DirectXlsxFluentStyledExport(string directory)
{
    var bytes = ExcelWriter.Create()
        .AddSheet("Direct Sales", sheet =>
        {
            sheet.AddCell("B1", "Direct Sales Export", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
            sheet.AddTable(SalesRows())
                .StartAt("B3")
                .HeaderAlias("Income", "Income Total")
                .HeaderAlias("Costs", "Cost Total")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                .AlternateRowStyle(style => style.BackgroundColor("#DDEBF7").Border())
                .ColumnStyle("Income", style => style.NumberFormat(ExcelNumberFormat.Currency))
                .ColumnStyle("Costs", style => style.NumberFormat(ExcelNumberFormat.Currency))
                .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick));
        })
        .ToBytes();

    File.WriteAllBytes(Path.Combine(directory, "direct-styled-sales.xlsx"), bytes);
    WriteNotes(directory, "Direct XLSX export using the fluent worksheet/table API with headings, styles, currency columns and totals.");
}

void DirectXlsxDataTableWorkbookExport(string directory)
{
    var dataSet = BusinessDataSet();
    var bytes = ExcelWriter.Create()
        .AddSheet("Sales", sheet => sheet.AddTable(dataSet.Tables["Sales"]!).StartAt("A1"))
        .AddSheet("Customers", sheet => sheet.AddTable(dataSet.Tables["Customers"]!).StartAt("A1"))
        .ToBytes();

    File.WriteAllBytes(Path.Combine(directory, "direct-dataset.xlsx"), bytes);
    WriteNotes(directory, "Direct XLSX export from DataTable instances, one sheet per table.");
}

void DirectXlsxSelectedColumnImport(string directory)
{
    var sourcePath = Path.Combine(directory, "source.xlsx");
    File.WriteAllBytes(sourcePath, ImportWorkbook());

    var rows = ExcelReader.FromFile(sourcePath)
        .Sheet("Import Source")
        .StartAt("A1")
        .Columns(columns => columns.IncludeHeaders("Name", "Amount"))
        .ReadDictionaries();

    WriteJson(Path.Combine(directory, "selected-columns.json"), rows);
    WriteNotes(directory, "Direct XLSX import that reads selected columns by header name.");
}

void DirectCsvDataTablePipeExport(string directory)
{
    var table = SalesDataTable("Sales");
    var csv = CsvWriter.FromDataTable(table)
        .DelimitedBy("|")
        .ColumnOrder("Region", "Amount")
        .HeaderAlias("Amount", "Sales Amount")
        .ToText();

    File.WriteAllText(Path.Combine(directory, "direct-sales.csv"), csv);
    WriteNotes(directory, "Direct CSV export from a DataTable with pipe delimiter, column order and header alias.");
}

void DirectCsvObjectImportWithMapping(string directory)
{
    var sourcePath = Path.Combine(directory, "people.csv");
    File.WriteAllText(sourcePath, "Full Name,Years,Ignored\nAlice,31,x\nBob,42,y");

    var result = CsvReader.FromFile(sourcePath)
        .ReadObjects<PersonImportRow>(map => map
            .Column(row => row.Name, "Full Name", required: true)
            .Column(row => row.Age, "Years", required: true));

    WriteJson(Path.Combine(directory, "people.json"), result.Items);
    WriteNotes(directory, "Direct CSV import into typed objects using code-based column mapping.");
}

// ---------------------------------------------------------------------------
// Shared sample data
// ---------------------------------------------------------------------------

static IReadOnlyList<SalesRow> SalesRows() =>
[
    new("London", 210000m, 125000m),
    new("United States", 325000m, 190000m),
    new("Europe", 260000m, 155000m),
    new("Asia", 295000m, 175000m)
];

static IReadOnlyList<CurrencyRow> CurrencyRows() =>
[
    new("USD", "United States", 1.2875m, 1.2420m),
    new("GBP", "United Kingdom", 1.0000m, 1.0000m),
    new("EUR", "Eurozone", 1.1764m, 1.1325m)
];

static IReadOnlyList<MonthlyRow> MonthlyRows() =>
[
    new("Jan", 120000m),
    new("Feb", 132000m),
    new("Mar", 141000m),
    new("Apr", 138000m),
    new("May", 155000m),
    new("Jun", 168000m)
];

static DataTable SalesDataTable(string tableName)
{
    var table = new DataTable(tableName);
    table.Columns.Add("Region", typeof(string));
    table.Columns.Add("Amount", typeof(decimal));
    table.Rows.Add("North", 125.25m);
    table.Rows.Add("South", 220.00m);
    return table;
}

static DataSet BusinessDataSet()
{
    var dataSet = new DataSet("Business");

    var sales = new DataTable("Sales");
    sales.Columns.Add("Region", typeof(string));
    sales.Columns.Add("Amount", typeof(decimal));
    sales.Rows.Add("North", 125.25m);
    sales.Rows.Add("South", 220.00m);
    dataSet.Tables.Add(sales);

    var customers = new DataTable("Customers");
    customers.Columns.Add("Name", typeof(string));
    customers.Columns.Add("Tier", typeof(string));
    customers.Columns.Add("Balance", typeof(decimal));
    customers.Rows.Add("Contoso", "Gold", 1200.00m);
    customers.Rows.Add("Fabrikam", "Silver", 875.50m);
    dataSet.Tables.Add(customers);

    return dataSet;
}

static byte[] ImportWorkbook()
{
    return ExcelWriter.Create()
        .AddSheet("Import Source", sheet => sheet
            .AddTable(new[]
            {
                new { ID = 1, Name = "Alice", Amount = 12.50m },
                new { ID = 2, Name = "Bob", Amount = 30.00m }
            })
            .StartAt("A1"))
        .ToBytes();
}

// ---------------------------------------------------------------------------
// API-equivalent writers and JSON helpers
// ---------------------------------------------------------------------------

static void WriteTemplateExportScenario(
    string directory,
    string format,
    string outputFileName,
    string templateJson,
    IReadOnlyDictionary<string, object?> dataSources,
    IReadOnlyDictionary<string, object?>? parameters = null)
{
    WriteJson(Path.Combine(directory, "template.json"), JsonDocument.Parse(templateJson).RootElement);
    WriteJson(Path.Combine(directory, "sample-data.json"), new { dataSources, parameters });

    var request = new DocWeaveExportRequest(
        Format: format,
        TemplateId: null,
        TemplateJson: templateJson,
        DataSources: ToJsonElementDictionary(dataSources),
        Parameters: parameters is null ? null : ToJsonElementDictionary(parameters),
        Delivery: new DocWeaveExportDeliveryRequest("download", outputFileName),
        Audit: DemoAudit());

    WriteJson(Path.Combine(directory, "api-request.json"), request);
    WriteNotes(directory, $"Template {format.ToUpperInvariant()} export. api-request.json is the Web API equivalent of the local template call.");
}

static void WriteTemplateImportScenario(string directory, string format, string templateJson, string sourcePath)
{
    WriteJson(Path.Combine(directory, "template.json"), JsonDocument.Parse(templateJson).RootElement);

    var request = new DocWeaveImportRequest(
        Format: format,
        Source: new DocWeaveImportSource(
            Kind: "path",
            FileName: Path.GetFileName(sourcePath),
            Path: Path.GetFullPath(sourcePath)),
        TemplateId: null,
        TemplateJson: templateJson,
        Audit: DemoAudit());

    WriteJson(Path.Combine(directory, "api-request.json"), request);
    WriteNotes(directory, $"Template {format.ToUpperInvariant()} import. api-request.json is the Web API equivalent of the local template read.");
}

static DocWeaveApiAuditMetadata DemoAudit()
{
    return new DocWeaveApiAuditMetadata(
        ClientId: "docweave-examples",
        RequestedBy: "developer",
        CorrelationId: "example-run",
        Tags: new Dictionary<string, string> { ["purpose"] = "how-to" });
}

static Dictionary<string, object?> DataSetSources(DataSet dataSet)
{
    return dataSet.Tables
        .Cast<DataTable>()
        .ToDictionary(table => table.TableName, table => (object?)DataTableRows(table), StringComparer.OrdinalIgnoreCase);
}

static List<Dictionary<string, object?>> DataTableRows(DataTable table)
{
    return table.Rows
        .Cast<DataRow>()
        .Select(row => table.Columns
            .Cast<DataColumn>()
            .ToDictionary(column => column.ColumnName, column => row[column] is DBNull ? null : row[column]))
        .ToList();
}

static Dictionary<string, JsonElement> ToJsonElementDictionary(IReadOnlyDictionary<string, object?> values)
{
    return values.ToDictionary(
        item => item.Key,
        item => JsonSerializer.SerializeToElement(item.Value, CreateJsonOptions()),
        StringComparer.OrdinalIgnoreCase);
}

static void WriteJson(string path, object? value)
{
    File.WriteAllText(path, JsonSerializer.Serialize(value, CreateJsonOptions()));
}

static void WriteNotes(string directory, string notes)
{
    File.WriteAllText(Path.Combine(directory, "README.txt"), notes + Environment.NewLine);
}

static JsonSerializerOptions CreateJsonOptions()
{
    return new JsonSerializerOptions
    {
        WriteIndented = true
    };
}

public sealed record SalesRow(string Region, decimal Income, decimal Costs);

public sealed record CurrencyRow(string Code, string Country, decimal LastHighRate, decimal LastLowRate);

public sealed record MonthlyRow(string Month, decimal Income);

public sealed record PersonImportRow(string Name, int Age);
