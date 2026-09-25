# Getting Started With DocWeave

DocWeave gives you two ways to build imports and exports:

- Fluent API: best for .NET callers that want compile-time code and full control.
- Template/API approach: best for Web API or service-to-service use, where the report layout should be sent or stored as JSON and data is supplied separately.

Use the fluent API inside trusted .NET code. Use templates when the caller is remote, when layouts need to be approved/stored, or when a Web API should accept data without exposing the whole C# builder surface.

## Basic Excel Export With The Fluent API

```csharp
using DocWeave.Excel;

var invoices = new[]
{
    new { Id = 1, Customer = "Northwind", Net = 100m, Vat = 20m },
    new { Id = 2, Customer = "Contoso", Net = 250m, Vat = 50m }
};

byte[] bytes = ExcelWriter.Create()
    .AddSheet("Invoices", sheet =>
    {
        sheet.AddCell("B1", "Invoice Export", s => s.Bold().FontSize(16));
        sheet.MergeCells("B1:F1");

        sheet.AddTable(invoices)
            .StartAt("B4")
            .HeaderAlias("Id", "Invoice No")
            .HeaderAlias("Net", "Net Amount")
            .HeaderAlias("Vat", "VAT")
            .FormulaColumn("Gross", ctx => $"D{ctx.WorksheetRow}+E{ctx.WorksheetRow}")
            .HeaderStyle(s => s.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF").Border())
            .DataStyle(s => s.Border().NumberFormat(ExcelNumberFormat.Currency))
            .AlternateRowStyle(s => s.BackgroundColor("#EAF3F8"))
            .ColumnStyle("Customer", s => s.FontFamily("Aptos").WrapText())
            .AddTotals.Bottom("Totals", s => s.Bold().Border(ExcelBorderLineStyle.Thick));

        sheet.Column("B").Width(14);
        sheet.Column("C").Width(28);
        sheet.FreezePanes(rows: 4);
        sheet.AutoFilter("B4:F6");
    })
    .ToBytes();

File.WriteAllBytes("invoices.xlsx", bytes);
```

Use `SaveAs(path)` when writing directly to disk:

```csharp
ExcelWriter.Create()
    .AddSheet("Invoices", sheet => sheet.AddTable(invoices).StartAt("A1"))
    .SaveAs("invoices.xlsx");
```

## Exporting Multiple Sheets

```csharp
byte[] bytes = ExcelWriter.Create()
    .AddSheet("Regions", sheet =>
    {
        sheet.AddTable(regionRows)
            .StartAt("A1")
            .HeaderStyle(s => s.Bold().BackgroundColor("#5B9BD5").FontColor("#FFFFFF"));
    })
    .AddSheet("Invoices", sheet =>
    {
        sheet.AddTable(invoiceRows)
            .StartAt("B3")
            .AddTotals.Bottom();
    })
    .AddSheet("Summary", sheet =>
    {
        sheet.AddCell("A1", "Summary", s => s.Bold().FontSize(16));
        sheet.AddFormulaCell("A3", "SUM(Invoices!F:F)", s => s.Bold().NumberFormat(ExcelNumberFormat.Currency));
    })
    .ToBytes();
```

## Repeater-Style Excel Export

Use a repeater when one source item should become a small block/card rather than a single row.

```csharp
byte[] bytes = ExcelWriter.Create()
    .AddSheet("Rates", sheet =>
    {
        sheet.AddRepeater(currencyRows)
            .StartAt("B3")
            .ItemsPerRow(3)
            .ItemSize(columns: 5, rows: 6)
            .Gap(columns: 1, rows: 2)
            .TitleField("Currency")
            .Field("Country")
            .Field("LastHighRate", "High")
            .Field("LastLowRate", "Low")
            .FormulaField("Spread", ctx => $"D{ctx.WorksheetRow + 3}-D{ctx.WorksheetRow + 4}")
            .CardStyle(s => s.Border(ExcelBorderLineStyle.Thick).BackgroundColor("#F7FBFF"))
            .TitleStyle(s => s.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF"))
            .LabelStyle(s => s.Bold())
            .ValueStyle(s => s.NumberFormat(ExcelNumberFormat.Number));
    })
    .ToBytes();
```

## Excel Import With The Fluent API

Read a sheet into dictionaries:

```csharp
var rows = ExcelReader.FromFile("invoices.xlsx")
    .Sheet("Invoices")
    .StartAt("B4")
    .Columns(c => c.IncludeHeaders("Invoice No", "Customer", "Gross"))
    .ReadDictionaries();
```

Read a sheet into a `DataTable`:

```csharp
DataTable table = ExcelReader.FromBytes(uploadedBytes, "upload.xlsx")
    .SheetAt(0)
    .StartAt("A1")
    .ReadDataTable("ImportedRows");
```

Use `ReadDictionaryResult()` or `ReadResult()` when you also want metadata such as headers and import result details.

## CSV Export With The Fluent API

```csharp
using DocWeave.Csv;

string csv = CsvWriter.FromObjects(invoices)
    .DelimitedBy(";")
    .ColumnOrder("Id", "Customer", "Net", "Vat")
    .HeaderAlias("Id", "Invoice No")
    .HeaderAlias("Net", "Net Amount")
    .IncludeHeaders()
    .ToText();
```

Write to a stream or file:

```csharp
CsvWriter.FromObjects(invoices)
    .DelimitedBy("|")
    .SaveAs("invoices.csv");
```

CSV export neutralizes formula-like text by default. This helps prevent CSV formula injection when a file is opened in Excel. Only call `AllowFormulaText()` for trusted data and trusted recipients.

## CSV Import With The Fluent API

Select columns by Excel-style letters:

```csharp
var rows = CsvReader.FromFile("input.csv")
    .Columns(c => c.IncludeLetters("A", "B", "E"))
    .ReadDictionaries();
```

Select columns by header:

```csharp
DataTable table = CsvReader.FromText(csvText)
    .DelimitedBy("|")
    .HasHeaderRow()
    .Columns(c => c
        .IncludeHeaders("Id", "Name", "Amount")
        .ExcludeHeaders("InternalNotes"))
    .ReadDataTable("ImportedCsv");
```

For large files, use `EnumerateRows()` to stream rows lazily:

```csharp
foreach (var row in CsvReader.FromFile("large.csv").EnumerateRows())
{
    // Process one row at a time.
}
```

## Reading Rows Into Objects

`ReadObjects<T>()` turns each row into an object, for both Excel and CSV. Problems are collected in a list; a bad value does not stop the import.

```csharp
public sealed class Invoice
{
    public string? InvoiceNo { get; set; }            // matches the header "Invoice No"
    public string? Customer { get; set; }
    public decimal NetAmount { get; set; }            // matches "Net Amount", "net_amount" or "NetAmount"
    public DateTime InvoiceDate { get; set; }
    public Status Status { get; set; }                // an enum, matched by name ignoring case
    public int? Quantity { get; set; }                // blank gives null

    [DocWeaveColumn("Sales Rep", Required = true)]    // a different header, and a blank is an error
    public string? Rep { get; set; }

    [DocWeaveColumn(Ordinal = 9)]                     // the 9th imported column, whatever its header says
    public string? Notes { get; set; }

    [DocWeaveIgnore]                                  // never filled from the file
    public string? Scratch { get; set; }
}

var result = ExcelReader.FromFile("invoices.xlsx")
    .Sheet("Invoices")
    .StartAt("B4")
    .ReadObjects<Invoice>();

foreach (var invoice in result.Items)
{
    // Only rows with no problems.
}

foreach (var error in result.Errors)
{
    Console.WriteLine($"Row {error.RowNumber}, column '{error.Column}': {error.Message}");
}
```

The same call works on CSV: `CsvReader.FromFile("invoices.csv").ReadObjects<Invoice>()`.

How columns are matched, in order:

1. `[DocWeaveColumn(Ordinal = n)]`: the n-th imported column, counting from 1. The count starts at the start cell and applies after any column selection.
2. `[DocWeaveColumn("header")]`: the column with that header.
3. The property name. Case, spaces, underscores and hyphens are ignored, so `NetAmount` matches `Net Amount`.

What is converted: text, whole numbers (`int`, `long`, `short`, `byte` and the unsigned kinds), `decimal`, `double`, `float`, `bool` (`true/false`, `yes/no`, `1/0`), `DateTime`, `DateOnly`, `TimeOnly`, `DateTimeOffset`, `TimeSpan`, `Guid`, enums, and the nullable form of each. Numbers and dates are read with the invariant culture, so `1.5` is one and a half and `1,5` is reported as an error rather than read as 15. In an Excel file a date property also accepts the number Excel stores for a date. Other property types throw `NotSupportedException` when the read starts; mark them `[DocWeaveIgnore]`.

Rules for blanks and errors:

- A blank cell leaves the property as its constructor or initializer set it. If the property is a plain value type such as `int`, or is marked `Required = true`, a blank is an error instead.
- A row with any error is left out of `Items`, and every problem in that row is listed in `Errors`. `RowNumber` is the worksheet row, or the CSV record number, so it matches what you see in the file.
- A required column that is missing from the file is one error with `RowNumber` 0, and no objects are returned.

Records and constructors:

- A class with a public parameterless constructor is created first and then filled property by property. `init` setters work.
- A type without one, such as a positional record, is built through its public constructor (the one with the most parameters). Parameters are matched to columns exactly as properties are, and a value the file does not supply becomes the parameter's default value.
- Properties the constructor does not take (for example an `init` property added to a record) are filled after it is built.
- Put `[DocWeaveColumn]` on the parameter, or on the property with `[property: DocWeaveColumn(...)]`.
- If the constructor or a setter throws, the row is not returned and the exception message is listed as an error for that row.
- A type that cannot be built (abstract, no public constructor, or several equal constructors) throws `NotSupportedException` when the read starts.

```csharp
public sealed record Invoice(
    string InvoiceNo,                                    // matches "Invoice No"
    decimal NetAmount,
    [DocWeaveColumn("Sales Rep", Required = true)] string Rep,
    int Quantity = 1);                                   // 1 when the column is missing or blank

var invoices = CsvReader.FromFile("invoices.csv").ReadObjects<Invoice>().Items;
```

Mapping in code, for a class you cannot put attributes on, or to read one class from differently laid out files:

```csharp
var result = ExcelReader.FromFile("invoices.xlsx")
    .Sheet("Invoices")
    .ReadObjects<Invoice>(map => map
        .Column(x => x.InvoiceNo, "Number")             // by header
        .Column(x => x.NetAmount, 3)                    // by position (the 3rd imported column)
        .Column(x => x.Rep, "Sales Rep", required: true)
        .Ignore(x => x.Quantity));
```

- `Column(x => x.Property, "header")`, `Column(x => x.Property, position)`, `Required(x => x.Property)` and `Ignore(x => x.Property)` are the four settings.
- A setting made in code replaces the attribute on that property. A property that is not mentioned keeps its attribute, or is matched by name.
- `Column(...)` on a property marked `[DocWeaveIgnore]` brings it back.
- `EnumerateObjects<T>(map => ..., onError)` takes the same map. Pass `null` for `onError` to ignore problems.

To read lazily and report problems as they happen, use `EnumerateObjects<T>(onError)`:

```csharp
foreach (var invoice in ExcelReader.FromFile("big.xlsx").EnumerateObjects<Invoice>(error => log.Warn(error.Message)))
{
    // One object at a time.
}
```

## Streaming Large Excel Files

`EnumerateRows()` reads a worksheet one row at a time instead of loading it all, as CSV already did.

```csharp
foreach (var row in ExcelReader.FromFile("large.xlsx").Sheet("Data").EnumerateRows())
{
    // row.Values["Header"]
}
```

On a 4.6 MB sheet of 200,000 rows, measured on the development machine, `ReadRows()` needed about 180 MB of extra memory and `EnumerateRows()` about 8 MB. These numbers depend on the machine and the data.

Differences from `ReadRows()`:

- The header row sets the columns, because later rows have not been read yet. A cell to the right of the header row is returned as `Column N` when no column selection is configured, and left out otherwise. `ReadRows()` sees the whole sheet first, so it gives such a column its own header.
- The source must be seekable, because an `.xlsx` file is a zip archive.
- The workbook is closed when the loop ends or is stopped early.
- Cancellation is checked for each row. Progress and audit events are not sent.
- The shared string table (where Excel keeps its text) is still read into memory once.

## Progress, Cancellation And Audit

All main fluent builders expose progress, cancellation and audit hooks.

```csharp
using DocWeave;

using var cts = new CancellationTokenSource();

var progress = new Progress<DocWeaveOperationProgress>(p =>
{
    Console.WriteLine($"{p.Operation}: {p.Stage} {p.Completed}/{p.Total}");
});

byte[] bytes = ExcelWriter.Create()
    .WithProgress(progress)
    .WithCancellation(cts.Token)
    .Audit(e => Console.WriteLine($"{e.Operation}: {e.Message}"))
    .AddSheet("Data", sheet => sheet.AddTable(rows).StartAt("A1"))
    .ToBytes();
```

Keep audit events about operation metadata. Do not put sensitive cell values into logs.

## Basic Excel Export With A Template

Templates describe the output workbook. The data is bound separately by name.

```json
{
  "schemaVersion": "1.0",
  "kind": "excelWorkbook",
  "styleCatalog": {
    "title": {
      "font": { "bold": true, "size": 16 }
    },
    "header": {
      "font": { "bold": true, "color": "#FFFFFF" },
      "fill": { "color": "#1F4E78" },
      "border": { "style": "thin" }
    },
    "money": {
      "numberFormat": "currency",
      "border": { "style": "thin" }
    }
  },
  "sheetTemplates": {
    "invoiceSheet": {
      "regions": [
        { "type": "cell", "cell": "B1", "value": "Invoice Export", "style": "title" },
        { "type": "mergeCells", "range": "B1:F1" },
        {
          "type": "table",
          "source": "invoices",
          "startCell": "B4",
          "columns": [
            { "source": "Id", "caption": "Invoice No" },
            { "source": "Customer" },
            { "source": "Net", "caption": "Net Amount", "style": "money" },
            { "source": "Vat", "caption": "VAT", "style": "money" }
          ],
          "styles": {
            "header": "header",
            "data": "money"
          },
          "totals": {
            "label": "Totals",
            "style": "header"
          }
        },
        { "type": "freezePane", "rows": 4, "columns": 0 },
        { "type": "autoFilter", "range": "B4:E100" }
      ]
    }
  },
  "sheets": [
    {
      "name": "Invoices",
      "template": "invoiceSheet"
    }
  ]
}
```

Render the template:

```csharp
byte[] bytes = ExcelTemplateWriter.FromJson(templateJson)
    .WithData("invoices", invoices)
    .ToBytes();
```

Validate before rendering:

```csharp
var builder = ExcelTemplateWriter.FromJson(templateJson)
    .WithData("invoices", invoices);

var validation = builder.Validate();
if (!validation.IsValid)
{
    foreach (var error in validation.Errors)
    {
        Console.WriteLine($"{error.Path}: {error.Message}");
    }
}
```

## Template Excel Chart Example

```json
{
  "schemaVersion": "1.0",
  "kind": "excelWorkbook",
  "sheetTemplates": {
    "dataSheet": {
      "regions": [
        { "type": "table", "source": "monthly", "startCell": "A1" }
      ]
    },
    "chartSheet": {
      "regions": [
        {
          "type": "chart",
          "title": "Income by Month",
          "chartType": "column",
          "dataSource": {
            "sheet": "Data",
            "categoryRange": "A2:A13",
            "valueRange": "B2:B13"
          },
          "position": {
            "topLeft": "B2",
            "bottomRight": "J20"
          },
          "showLegend": false,
          "axisTitles": {
            "category": "Month",
            "value": "Income"
          }
        }
      ]
    }
  },
  "sheets": [
    {
      "name": "Data",
      "template": "dataSheet"
    },
    {
      "name": "Charts",
      "template": "chartSheet"
    }
  ]
}
```

## Basic CSV Export With A Template

```json
{
  "source": "invoices",
  "output": {
    "delimiter": "|",
    "includeHeaders": true
  },
  "columns": [
    { "source": "Id", "caption": "Invoice No" },
    { "source": "Customer" },
    { "source": "Net", "caption": "Net Amount" },
    { "source": "Vat", "caption": "VAT" }
  ]
}
```

Render the CSV:

```csharp
string csv = CsvTemplateWriter.FromJson(templateJson)
    .WithData("invoices", invoices)
    .ToText();
```

## Using DocWeave Behind A Web API

For API scenarios, prefer accepting a template and named data sources rather than exposing the fluent API remotely.

Recommended export request shape:

```json
{
  "format": "xlsx",
  "templateId": "monthly-invoice-export",
  "templateJson": null,
  "dataSources": {
    "invoices": [
      { "id": 1, "customer": "Northwind", "net": 100, "vat": 20 }
    ]
  },
  "parameters": {
    "requestedBy": "finance-service"
  },
  "delivery": {
    "kind": "download",
    "fileName": "invoices.xlsx"
  },
  "audit": {
    "clientId": "finance-service",
    "requestedBy": "francois",
    "correlationId": "abc-123"
  }
}
```

The DTOs for this are in `DocWeave.WebApi`:

- `DocWeaveExportRequest`
- `DocWeaveImportRequest`
- `DocWeaveImportSource`
- `DocWeaveExportDeliveryRequest`
- `DocWeaveApiAuditMetadata`
- `DocWeaveOperationAcceptedResponse`
- `DocWeaveOperationStatusResponse`

A simple controller action can translate the API request into the template builder:

```csharp
using DocWeave.Excel;
using DocWeave.WebApi;

[HttpPost("exports/excel")]
public IActionResult ExportExcel(DocWeaveExportRequest request)
{
    var builder = ExcelTemplateWriter.FromJson(request.TemplateJson!);

    foreach (var source in request.DataSources)
    {
        var rows = JsonSerializer.Deserialize<List<Dictionary<string, object?>>>(source.Value.GetRawText())!;
        builder.WithRows(source.Key, rows);
    }

    foreach (var parameter in request.Parameters ?? new Dictionary<string, JsonElement>())
    {
        builder.WithParameter(parameter.Key, ToTemplateValue(parameter.Value));
    }

    var bytes = builder.ToBytes();
    var fileName = request.Delivery?.FileName ?? "export.xlsx";

    return File(
        bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileName);
}

static object? ToTemplateValue(JsonElement value)
{
    return value.ValueKind switch
    {
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number when value.TryGetInt64(out var longValue) => longValue,
        JsonValueKind.Number => value.GetDouble(),
        JsonValueKind.True => true,
        JsonValueKind.False => false,
        JsonValueKind.Null => null,
        _ => JsonSerializer.Deserialize<object?>(value.GetRawText())
    };
}
```

In production, the API host should also:

- Load approved templates by `TemplateId`.
- Reject unknown or unapproved templates where required.
- Validate the template before rendering.
- Enforce size limits on uploaded data.
- Apply authorization per template and per destination.
- Record audit metadata and correlation ids.
- Use background jobs for long-running exports/imports.
- Return `DocWeaveOperationAcceptedResponse` and expose operation status for long jobs.

## Choosing Fluent Or Template/API

Use fluent when:

- The caller is a .NET application.
- The export layout is owned by code.
- Compile-time refactoring and typed models matter.
- The consumer is trusted and internal.

Use templates when:

- A Web API receives requests from other services.
- Layouts need to be stored, versioned or approved.
- The caller should provide data, not execute layout code.
- Multiple clients need the same report shape.
- You want safer long-running operation handling and auditing at the API boundary.

Both approaches can coexist. Internally, the goal is for both the fluent builders and the template layer to translate into the same workbook/table/CSV specifications before rendering.

## Useful Test Outputs

The test project writes example files under:

```text
DocWeave.Tests\TestOutput\Excel
DocWeave.Tests\TestOutput\Csv
```

Those files are useful when checking how a feature renders in Excel or how CSV output is shaped.

Large Excel volume variants can be enabled with:

```powershell
$env:DOCWEAVE_RUN_LARGE_EXCEL_VOLUME = "1"
dotnet test DocWeave.Tests
```
