# DocWeave Pre-Release Functionality Guide

DocWeave is a .NET library for importing and exporting Excel `.xlsx` and CSV files. It can be used directly through a fluent C# API, or indirectly through JSON templates that are easier to expose from a Web API or another service boundary.

This guide lists the current pre-release functionality in a user-friendly form. The API is still marked as preview, so method names and template shapes may change before a stable `1.0` release.

## Sorted Index

- [API and Service Usage](#api-and-service-usage)
- [Audit Events](#audit-events)
- [Cancellation](#cancellation)
- [Charts](#charts)
- [Column Selection](#column-selection)
- [CSV Export](#csv-export)
- [CSV Import](#csv-import)
- [CSV Templates](#csv-templates)
- [Data Sources](#data-sources)
- [Data Validation](#data-validation)
- [Excel Export](#excel-export)
- [Excel Import](#excel-import)
- [Excel Templates](#excel-templates)
- [File Formats](#file-formats)
- [Formulas](#formulas)
- [Known Preview Limitations](#known-preview-limitations)
- [Layout and Sheet Features](#layout-and-sheet-features)
- [Output Targets](#output-targets)
- [Pivot Tables](#pivot-tables)
- [Progress Reporting](#progress-reporting)
- [Protection](#protection)
- [Reading Rows Into Objects](#reading-rows-into-objects)
- [Repeaters](#repeaters)
- [Streaming Large Excel Files](#streaming-large-excel-files)
- [Styling](#styling)
- [Tables](#tables)
- [Testing and Generated Examples](#testing-and-generated-examples)

## File Formats

Current formats:

- Excel `.xlsx` import.
- Excel `.xlsx` export.
- CSV import.
- CSV export.

Not currently supported:

- `.xls`, because it is the older binary Excel format.
- Creating macros.
- Full workbook encryption.
- Word and PowerPoint output.

## Data Sources

DocWeave can export tabular data from:

- `DataTable`
- `IEnumerable<T>` object collections
- `IEnumerable<IReadOnlyDictionary<string, object?>>` dictionary rows

The library normalises these sources into an internal tabular model so Excel, CSV and templates can share the same source handling.

## Output Targets

Excel and CSV exports can be written to:

- A byte array for API responses, storage or attachments.
- A file path.
- A caller-owned stream.
- Text output for CSV.
- A `TextWriter` for CSV.

Example:

```csharp
byte[] bytes = ExcelWriter.Create()
    .AddSheet("Data", sheet => sheet.AddTable(rows).StartAt("A1"))
    .ToBytes();
```

## Column Selection

Import flows can select or exclude columns by:

- Excel-style column letters, such as `A`, `B`, `E`.
- Header names, such as `Id`, `Name`, `Amount`.

Example:

```csharp
var rows = ExcelReader.FromFile("input.xlsx")
    .Sheet("Data")
    .StartAt("B4")
    .Columns(c => c
        .IncludeHeaders("Id", "Name", "Amount")
        .ExcludeHeaders("InternalNotes"))
    .ReadDictionaries();
```

CSV supports requested order when selecting columns. Excel import reads in worksheet order.

## Excel Export

Excel export starts with `ExcelWriter.Create()`.

Current workbook options:

- Add one or more worksheets.
- Save to a path.
- Save to a stream.
- Render to bytes.
- Protect workbook structure.
- Report progress.
- Accept a cancellation token.
- Emit audit events.

Example:

```csharp
ExcelWriter.Create()
    .AddSheet("Invoices", sheet =>
    {
        sheet.AddTable(invoices)
            .StartAt("B4")
            .HeaderAlias("Id", "Invoice No")
            .AddTotals.Bottom("Totals");
    })
    .SaveAs("invoices.xlsx");
```

## Excel Import

Excel import starts with `ExcelReader`.

Supported sources:

- `FromFile(path)`
- `FromStream(stream, label)`
- `FromBytes(bytes, label)`
- `FromBase64(base64, label)`
- `FromBase64Text(base64Text, encoding, label)`

Supported options:

- Choose a sheet by name.
- Choose a sheet by index.
- Set the starting cell.
- Read with or without a header row.
- Include or exclude columns.
- Report progress.
- Accept cancellation.
- Emit audit events.

Supported result shapes:

- `ReadRows()`
- `ReadResult()`
- `ReadDictionaries()`
- `ReadDictionaryResult()`
- `ReadDataTable(tableName)`
- `EnumerateRows()`, which reads one row at a time
- `ReadObjects<T>()` and `EnumerateObjects<T>()`, which fill your own class

Example:

```csharp
DataTable table = ExcelReader.FromBytes(uploadedBytes, "upload.xlsx")
    .Sheet("Sales")
    .StartAt("A1")
    .ReadDataTable("Sales");
```

Excel import currently focuses on reading worksheet data. It does not yet reverse-engineer a full exported workbook layout with charts, styles, comments, protection, pivots or templates.

## Tables

Excel table regions can be exported from `DataTable`, object collections or dictionary rows.

Current table options:

- Start cell, for example `StartAt("B4")`.
- Include or suppress headers.
- Header aliases.
- Column locations.
- Header styling.
- Data styling.
- Alternate row styling.
- Column styling.
- Specific row styling.
- Conditional row styling.
- Conditional cell styling.
- Formula columns.
- Conditional value styling.
- Native Excel table output.
- Totals row at the bottom.

Example:

```csharp
sheet.AddTable(invoices)
    .StartAt("B4")
    .HeaderAlias("Id", "Invoice No")
    .ColumnLocation("Id", "B")
    .ColumnLocation("Customer", "C")
    .HeaderStyle(s => s.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF").Border())
    .DataStyle(s => s.Border())
    .AlternateRowStyle(s => s.BackgroundColor("#EAF3F8"))
    .AddTotals.Bottom("Totals", s => s.Bold().Border(ExcelBorderLineStyle.Thick));
```

## Styling

Current style options:

- Bold.
- Italic.
- Underline.
- Font family.
- Font size.
- Font colour.
- Background colour.
- Pattern fills.
- Wrap text.
- Rotate text.
- Number format.
- Border style and border colour.

Supported number formats:

- `General`
- `Number`
- `Currency`
- `Percent`
- `Date`
- `Text`

Supported border styles:

- `None`
- `Thin`
- `Medium`
- `Thick`
- `Dashed`
- `Dotted`
- `Double`
- `Hair`
- `DashDot`
- `DashDotDot`
- `SlantDashDot`

Supported fill patterns:

- `Solid`
- `Gray125`
- `DarkGray`
- `MediumGray`
- `LightGray`
- `DarkHorizontal`
- `DarkVertical`
- `DarkGrid`
- `LightGrid`

## Layout and Sheet Features

Current worksheet layout options:

- Static cells.
- Formula cells.
- Merged cells.
- Cell comments.
- Range styles.
- Column width.
- Column auto-fit metadata.
- Hidden columns.
- Column styles.
- Row height.
- Row auto-fit metadata.
- Hidden rows.
- Freeze panes.
- Auto filters.
- Data validation.
- Sheet protection.

Example:

```csharp
sheet.AddCell("B1", "Sales Export", s => s.Bold().FontSize(16));
sheet.MergeCells("B1:F1");
sheet.Column("B").Width(20);
sheet.Row(1).Height(28);
sheet.FreezePanes(rows: 4);
sheet.AutoFilter("B4:F100");
```

## Formulas

DocWeave can write formulas into:

- Individual cells.
- Table formula columns.
- Repeater formula fields.
- Template formula cells.

Formula text can be supplied with or without the leading equals sign.

Example:

```csharp
sheet.AddFormulaCell("F2", "SUM(F5:F100)", s => s.Bold().NumberFormat(ExcelNumberFormat.Currency));
```

Formula columns can use the generated worksheet row:

```csharp
sheet.AddTable(invoices)
    .FormulaColumn("Gross", ctx => $"D{ctx.WorksheetRow}+E{ctx.WorksheetRow}");
```

## Repeaters

Repeaters export each data row as a card-like block instead of one worksheet row. This is useful for grouped layouts such as regions, currencies, departments, summaries or repeating panels.

Current repeater options:

- Start cell.
- Items per row.
- Item width and height.
- Column and row gaps.
- Title field.
- Explicit fields.
- Formula fields.
- Card style.
- Title style.
- Label style.
- Value style.
- Per-field label/value styles.

Example:

```csharp
sheet.AddRepeater(currencies)
    .StartAt("B3")
    .ItemsPerRow(3)
    .ItemSize(columns: 5, rows: 7)
    .Gap(columns: 1, rows: 2)
    .TitleField("Currency")
    .Field("Country")
    .Field("LastHighRate", "High")
    .Field("LastLowRate", "Low")
    .CardStyle(s => s.Border(ExcelBorderLineStyle.Thick))
    .TitleStyle(s => s.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF"));
```

## Charts

Charts can be exported to Excel sheets.

Supported chart types:

- Column.
- Bar.
- Line.
- Area.
- Pie.
- Doughnut.
- Scatter.

Current chart options:

- Chart title.
- Chart type.
- Source sheet.
- Category range.
- Value range.
- Top-left and bottom-right position.
- Built-in Excel chart style number.
- Series colour.
- Legend on/off.
- Legend position.
- Data labels.
- Axis titles.

Example:

```csharp
sheet.AddChart("Income by Month", chart => chart
    .Type(ExcelChartType.Column)
    .DataSource("Data", "A2:A13", "B2:B13")
    .Position("H3", "P20")
    .ShowLegend(false)
    .AxisTitles("Month", "Income"));
```

## Pivot Tables

DocWeave includes a pivot table builder for Excel export.

Current pivot options:

- Source sheet and range.
- Start cell.
- Row fields.
- Column fields.
- Filter fields.
- Value fields.
- Sum, count, average, min and max aggregation.
- Style name.
- Row and column grand totals.
- Refresh-on-open flag.

Example:

```csharp
sheet.AddPivotTable("SalesByRegion", pivot => pivot
    .Source("Data", "A1:F101")
    .StartAt("A3")
    .Rows("Region")
    .Columns("Quarter")
    .Values(v => v.Sum("Amount", "Total Amount")));
```

Preview note: pivot generation is one of the more sensitive Open XML areas. The current test suite covers a number of row, column, value and filter scenarios, but generated pivot workbooks should still be opened in Excel during acceptance testing.

## Data Validation

Excel exports can include data validation rules.

Supported validation rules:

- Fixed list.
- List from range.
- Whole number between.
- Decimal between.
- Date between.
- Text length between.
- Custom formula.
- Allow blank flag.
- Input message.
- Error message.

Example:

```csharp
sheet.AddDataValidation("E4:E100", validation => validation
    .DecimalBetween(0, 1000000)
    .InputMessage("Amount", "Enter a positive amount.")
    .ErrorMessage("Invalid amount", "Enter a value between 0 and 1,000,000."));
```

## Protection

Current protection features:

- Workbook structure protection.
- Sheet protection.
- Optional sheet/workbook protection passwords.

Important: Excel sheet and workbook structure protection is not the same as file encryption. It prevents casual edits but does not securely encrypt the document.

## CSV Export

CSV export starts with `CsvWriter`.

Supported sources:

- `FromDataTable(dataTable)`
- `FromObjects(items)`
- `FromRows(dictionaryRows)`

Current export options:

- Delimiter.
- Quote character.
- Encoding.
- New line.
- Include or suppress headers.
- Column order.
- Header aliases.
- Formula-text safety.
- Progress reporting.
- Cancellation.
- Audit events.

Supported outputs:

- Text.
- Bytes.
- File.
- Stream.
- `TextWriter`.

Example:

```csharp
string csv = CsvWriter.FromObjects(invoices)
    .DelimitedBy(";")
    .ColumnOrder("Id", "Customer", "Net", "Vat")
    .HeaderAlias("Id", "Invoice No")
    .ToText();
```

CSV export neutralises formula-like text by default. Use `AllowFormulaText()` only when the source data and destination are trusted.

## CSV Import

CSV import starts with `CsvReader`.

Supported sources:

- `FromFile(path)`
- `FromStream(stream, label)`
- `FromBytes(bytes, label)`
- `FromText(text, label)`

Current import options:

- Delimiter.
- Quote character.
- Encoding.
- Header row on/off.
- Trim fields.
- Skip empty lines.
- Duplicate header handling.
- Row length mismatch handling.
- Include/exclude by letters.
- Include/exclude by headers.
- Progress reporting.
- Cancellation.
- Audit events.

Supported result shapes:

- Lazy row enumeration.
- Objects of your own class (`ReadObjects<T>()`, `EnumerateObjects<T>()`).
- Row list.
- Import result.
- Dictionaries.
- Dictionary import result.
- `DataTable`.

Example:

```csharp
var table = CsvReader.FromFile("input.csv")
    .DelimitedBy("|")
    .HasHeaderRow()
    .Columns(c => c.IncludeLetters("A", "B", "E"))
    .ReadDataTable("ImportedCsv");
```

## Excel Templates

Excel templates are JSON documents that describe workbook layout. They are useful for Web API and service-to-service scenarios where a caller should provide data, not execute layout code.

Current Excel template features:

- Style catalog.
- Named sheet templates.
- Output sheets based on templates.
- Bound data sources.
- Bound scalar parameters.
- Template validation.
- Template preview.
- Stored template provider.

Supported region types:

- `cell`
- `formulaCell`
- `table`
- `repeater`
- `summary`
- `chart`
- `pivotTable`
- `rangeStyle`
- `mergeCells`
- `comment`
- `column`
- `row`
- `dataValidation`
- `freezePane`
- `autoFilter`
- `sheetProtection`

Example:

```csharp
byte[] bytes = ExcelTemplateWriter.FromJson(templateJson)
    .WithData("sales", salesRows)
    .WithParameter("createdBy", "Finance API")
    .ToBytes();
```

## Reading Rows Into Objects

Both `ExcelReader` and `CsvReader` can fill your own class:

```csharp
var result = ExcelReader.FromFile("invoices.xlsx").Sheet("Invoices").ReadObjects<Invoice>();
// result.Items   the objects made from rows with no problems
// result.Errors  every problem: RowNumber, Column, Property, Value, Message
```

- Properties are matched to columns by header name (ignoring case, spaces and underscores), by `[DocWeaveColumn("header")]`, or by position with `[DocWeaveColumn(Ordinal = n)]`. `[DocWeaveIgnore]` skips a property.
- Problems are collected, not thrown. A row with a problem contributes no object, and each of its problems is listed.
- `Required = true` makes a blank value, or a missing column, an error.
- Numbers and dates use the invariant culture. In Excel a date property also accepts the number Excel stores for a date.
- `EnumerateObjects<T>(onError)` does the same lazily, one object at a time.
- Positional records and other types with only a constructor are supported: the values go through the constructor, and a parameter default is used when the file has no value.
- The mapping can also be written in code, without attributes: `ReadObjects<Invoice>(map => map.Column(x => x.Rep, "Sales Rep", required: true).Ignore(x => x.Notes))`.

Details and the list of supported property types are in [getting-started.md](getting-started.md#reading-rows-into-objects).

## Streaming Large Excel Files

`ExcelReader...EnumerateRows()` (and `EnumerateObjects<T>()`) read a worksheet one row at a time, so memory does not grow with the number of rows. The source must be seekable. See [getting-started.md](getting-started.md#streaming-large-excel-files) for the differences from `ReadRows()`.

## CSV Templates

CSV templates describe CSV output as JSON.

Current CSV template features:

- Named source.
- Delimiter.
- Header inclusion.
- Formula-text neutralisation setting.
- Column projection.
- Header captions.
- Validation.
- Output to text or bytes.

Example:

```csharp
string csv = CsvTemplateWriter.FromJson(templateJson)
    .WithData("sales", salesRows)
    .ToText();
```

## API and Service Usage

DocWeave includes API contract records in `DocWeave.WebApi`. These are DTOs only; they do not include controllers.

Current Web API contract records:

- `DocWeaveExportRequest`
- `DocWeaveImportRequest`
- `DocWeaveImportSource`
- `DocWeaveExportDeliveryRequest`
- `DocWeaveApiAuditMetadata`
- `DocWeaveOperationAcceptedResponse`
- `DocWeaveOperationStatusResponse`

Recommended API model:

- Accept a stored template id or inline template JSON.
- Accept named data sources.
- Accept scalar parameters.
- Return bytes directly for short operations.
- Return an accepted response and operation id for long-running operations.
- Expose progress/status polling for background jobs.

The recommended service boundary is template plus data, not the entire fluent API over HTTP.

## Progress Reporting

Main import/export builders can report progress through `IProgress<DocWeaveOperationProgress>`.

Progress records contain:

- Operation.
- Stage.
- Completed count.
- Optional total count.

## Cancellation

Main import/export builders accept a `CancellationToken`.

Current cancellation support:

- CSV import/export checks cancellation during processing.
- Excel import supports cancellation.
- Excel export currently checks cancellation before rendering starts.

## Audit Events

Main import/export builders can emit `DocWeaveAuditEvent` records.

Audit events are intended for operational metadata such as:

- Operation name.
- Stage.
- Sheet count.
- Row count.
- Client id.
- User id.
- Correlation id.

Do not write sensitive cell values to audit logs.

## Testing and Generated Examples

The test project is `DocWeave.Tests`.

Generated examples are written under:

```text
DocWeave.Tests\TestOutput\Excel
DocWeave.Tests\TestOutput\Csv
```

The normal suite uses moderate data sizes. Larger volume tests can be enabled with:

```powershell
$env:DOCWEAVE_RUN_LARGE_EXCEL_VOLUME = "1"
$env:DOCWEAVE_RUN_VOLUME_TESTS = "1"
dotnet test DocWeave.Tests
```

## Known Preview Limitations

Current limitations to keep in mind:

- The public API can still change before `1.0`.
- Excel import reads worksheet data, not full workbook layout.
- Excel import does not yet reverse styles, charts, pivots, comments or protections into a rich object model.
- `.xls` is not supported.
- Creating macros is not supported.
- File encryption is not implemented.
- `.xlsm` macro preservation is not a guaranteed feature.
- Excel auto-fit metadata may be recalculated visually by Excel only when the workbook is opened.
- Pivot output should be acceptance-tested in Excel for each important report shape.
- Web API contracts exist, but no Web API host project is included yet.

