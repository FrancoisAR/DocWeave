# DocWeave Current Functionality

This document describes the functionality currently implemented in `DocWeave`. It is a snapshot of the present code, not the full long-term design.

## Supported Formats

DocWeave currently focuses on:

- Excel `.xlsx` import and export.
- CSV import and export.
- Template-driven Excel and CSV generation for API/service scenarios.

Not currently supported:

- `.xls`, because it requires the older binary Excel format and usually Office interop.
- Creating new macros.
- Reliable native PivotTable generation for every scenario. See the Pivot Tables section.
- Word and PowerPoint generation. These remain later-phase modules.

## Main Entrypoints

Fluent Excel export:

```csharp
using DocWeave.Excel;

byte[] bytes = ExcelWriter.Create()
    .AddSheet("Sales", sheet =>
    {
        sheet.AddTable(rows).StartAt("A3");
    })
    .ToBytes();
```

Fluent Excel import:

```csharp
using DocWeave.Excel;

var table = ExcelReader.FromFile(path)
    .Sheet("Sales")
    .StartAt("A3")
    .ReadDataTable("Sales");
```

Fluent CSV export:

```csharp
using DocWeave.Csv;

string csv = CsvWriter.FromObjects(rows)
    .DelimitedBy("|")
    .ColumnOrder("Id", "Name", "Amount")
    .HeaderAlias("Amount", "Invoice Amount")
    .ToText();
```

Fluent CSV import:

```csharp
using DocWeave.Csv;

var rows = CsvReader.FromFile(path)
    .DelimitedBy("|")
    .Columns(c => c.IncludeHeaders("Id", "Name", "Amount"))
    .ReadDictionaries();
```

Template Excel export:

```csharp
using DocWeave.Excel;

byte[] bytes = ExcelTemplateWriter.FromJson(templateJson)
    .WithData("sales", salesRows)
    .WithParameter("createdBy", "Finance API")
    .ToBytes();
```

Template CSV export:

```csharp
using DocWeave.Csv;

string csv = CsvTemplateWriter.FromJson(templateJson)
    .WithData("sales", salesRows)
    .ToText();
```

## Excel Export

Excel export is built around `ExcelWriter.Create()`, one or more sheets, and sheet regions such as tables, repeaters, cells, charts and pivots.

### Workbook Features

- Create `.xlsx` workbooks in memory with `ToBytes()`.
- Save workbooks to a file with `SaveAs(path)`.
- Save workbooks to a caller-owned stream with `SaveTo(stream)`.
- Add multiple sheets.
- Protect workbook structure with `ProtectStructure(password)`.
- Receive progress with `WithProgress`.
- Support cancellation before rendering starts with `WithCancellation`.
- Emit audit events with `Audit`.

Workbook structure protection is not encryption. It uses Excel's legacy workbook protection mechanism and is intended to prevent casual sheet add/remove/rename/reorder changes.

### Worksheet Features

Each sheet can contain:

- Static cells.
- Formula cells.
- One or more tables.
- One or more repeater/card layouts.
- Charts.
- Pivot table requests.
- Comments.
- Merged cells.
- Range styles.
- Row and column settings.
- Freeze panes.
- Auto filters.
- Data validation.
- Sheet protection.

### Table Export

Tables can be created from:

- `DataTable`.
- `IEnumerable<T>` object collections.
- `IEnumerable<IReadOnlyDictionary<string, object?>>` dictionary rows.

Current table options include:

- `StartAt("B4")` placement.
- Include or suppress headers.
- Header aliases, so source column names can be replaced with friendly captions.
- Specific column locations, for example place `Id` in column `B` and `Name` in column `C`.
- Header style.
- Data style.
- Alternate row style.
- Column styles.
- Specific row styles by data row number.
- Conditional row styles with `RowStyleWhen`.
- Conditional cell styles with `CellStyle`.
- Formula columns using the current data index and worksheet row.
- Conditional styles for a column value.
- Native Excel tables with `AsNativeTable`.
- Totals rows at the bottom with `AddTotals.Bottom(...)`.

Example:

```csharp
byte[] bytes = ExcelWriter.Create()
    .AddSheet("Invoices", sheet =>
    {
        sheet.AddCell("B1", "Invoice Export", s => s.Bold().FontSize(16));
        sheet.AddCell("B2", DateTime.Today, s => s.NumberFormat(ExcelNumberFormat.Date));

        sheet.AddTable(invoices)
            .StartAt("B4")
            .HeaderAlias("Id", "Invoice No")
            .HeaderAlias("CustomerName", "Customer")
            .ColumnLocation("Id", "B")
            .ColumnLocation("CustomerName", "C")
            .ColumnLocation("Net", "D")
            .ColumnLocation("Vat", "E")
            .ColumnLocation("Gross", "F")
            .HeaderStyle(s => s.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF").Border())
            .DataStyle(s => s.Border().NumberFormat(ExcelNumberFormat.Number))
            .AlternateRowStyle(s => s.BackgroundColor("#EAF3F8"))
            .ColumnStyle("Gross", s => s.Bold().NumberFormat(ExcelNumberFormat.Currency))
            .FormulaColumn("Gross", ctx => $"D{ctx.WorksheetRow}+E{ctx.WorksheetRow}")
            .AddTotals.Bottom("Totals", s => s.Bold().Border(ExcelBorderLineStyle.Thick));
    })
    .ToBytes();
```

### Repeater Export

Repeaters are for card-like or group-like layouts where each object is exported into a rectangular block rather than a simple row. This supports scenarios such as multiple currencies, regions, departments or summary cards repeated across the sheet.

Current options include:

- Source from object collections or dictionary rows.
- `StartAt`.
- `ItemsPerRow`.
- `ItemSize(columns, rows)`.
- Horizontal and vertical `Gap`.
- Title field.
- Explicit fields.
- Formula fields.
- Card, title, label and value styles.
- Per-field label and value styles.

Example:

```csharp
byte[] bytes = ExcelWriter.Create()
    .AddSheet("Currencies", sheet =>
    {
        sheet.AddRepeater(currencies)
            .StartAt("B3")
            .ItemsPerRow(3)
            .ItemSize(columns: 5, rows: 7)
            .Gap(columns: 1, rows: 2)
            .TitleField("Currency")
            .Field("Country")
            .Field("LastHighRate", "High")
            .Field("LastLowRate", "Low")
            .FormulaField("Spread", ctx => $"D{ctx.WorksheetRow + 3}-D{ctx.WorksheetRow + 4}")
            .CardStyle(s => s.Border(ExcelBorderLineStyle.Thick).BackgroundColor("#F7FBFF"))
            .TitleStyle(s => s.Bold().FontSize(13).BackgroundColor("#1F4E78").FontColor("#FFFFFF"))
            .LabelStyle(s => s.Bold())
            .ValueStyle(s => s.NumberFormat(ExcelNumberFormat.Number));
    })
    .ToBytes();
```

### Styling

Current style options are intentionally simple and reusable:

- Bold, italic and underline.
- Font size.
- Font family.
- Font color.
- Background color.
- Pattern fills.
- Wrap text.
- Rotate text.
- Number format.
- Border style and border color.

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

Supported number formats:

- `General`
- `Number`
- `Currency`
- `Percent`
- `Date`
- `Text`

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

### Worksheet Layout

Current sheet layout options include:

- `Column("B").Width(20)`
- `Column("B").AutoFitWidth()`
- `Column("B").Hidden()`
- `Column("B").Style(...)`
- `Row(3).Height(28)`
- `Row(3).AutoFitHeight()`
- `Row(3).Hidden()`
- `RangeStyle("B4:F20", ...)`
- `MergeCells("B1:F1")`
- `FreezePanes(rows: 1, columns: 0)`
- `AutoFilter("B4:F104")`

### Comments

Comments can be added to individual cells:

```csharp
sheet.AddComment("B4", "Imported from finance service.", "DocWeave");
```

### Data Validation

The sheet builder can add validation rules to a range:

```csharp
sheet.AddDataValidation("F4:F100", v => v
    .DecimalBetween(0, 1000000)
    .InputMessage("Amount", "Enter a positive amount.")
    .ErrorMessage("Invalid amount", "The value must be between 0 and 1,000,000."));
```

Supported validation shapes:

- Fixed list values.
- List from another range.
- Whole number between.
- Decimal between.
- Date between.
- Text length between.
- Custom formula.
- Allow blank flag.
- Input and error messages.

### Charts

Charts can be placed on the same sheet as the data or on another sheet. The data source is an Excel range, so the chart can refer to any sheet.

Supported chart types:

- Column
- Bar
- Line
- Area
- Pie
- Doughnut
- Scatter

Current chart options include:

- Data source sheet.
- Category range.
- Value range.
- Top-left and bottom-right placement.
- Built-in Excel chart style number.
- Series color.
- Legend visibility.
- Legend position.
- Data labels.
- Axis titles.

Example:

```csharp
sheet.AddChart("Income by Month", chart => chart
    .Type(ExcelChartType.Column)
    .DataSource("Data", "A2:A13", "B2:B13")
    .Position("H3", "P20")
    .Style(10)
    .SeriesColor("#4472C4")
    .ShowLegend(false)
    .AxisTitles("Month", "Income"));
```

### Pivot Tables

The fluent API has a pivot table builder:

```csharp
sheet.AddPivotTable("SalesByRegion", pivot => pivot
    .Source("Data", "A1:F101")
    .StartAt("A3")
    .Rows("Region")
    .Columns("Quarter")
    .Values(v => v.Sum("Amount", "Total Amount"))
    .ShowGrandTotals(rows: true, columns: true));
```

The pivot API currently supports:

- Source sheet and range.
- Row fields.
- Column fields.
- Filter fields.
- Value fields with sum, count, average, min and max.
- Style name.
- Grand total options.
- Refresh-on-open flag.

Important current limitation: native PivotTable generation is still limited. Some complex native pivot combinations have previously produced workbooks that Excel repairs or closes unexpectedly. To keep generated workbooks safe, complex/multi-value/grouped pivot-style scenarios are currently rendered as static summary tables, and grouped output can use Excel row outlines for expand/collapse behavior. Treat native PivotTable output as experimental until it is rebuilt and tested against Excel-authored reference files.

## Excel Import

Excel import is built around `ExcelReader`.

Input sources:

- `FromFile(path)`
- `FromStream(stream, label)`
- `FromBytes(bytes, label)`
- `FromBase64(base64, label)`
- `FromBase64Text(base64Text, encoding, label)`

Options:

- Select sheet by name with `Sheet("Sales")`.
- Select sheet by zero-based index with `SheetAt(0)`.
- Set a starting cell with `StartAt("B4")`.
- Enable or disable header row handling.
- Include columns by Excel letters.
- Include columns by header names.
- Exclude columns by Excel letters.
- Exclude columns by header names.
- Progress, cancellation and audit.

Output shapes:

- `ReadRows()`
- `ReadResult()`
- `ReadDictionaries()`
- `ReadDictionaryResult()`
- `ReadDataTable(tableName)`
- `EnumerateRows()` for one row at a time.
- `ReadObjects<T>()` and `EnumerateObjects<T>()` for typed objects, with problems collected in a list.

Example:

```csharp
var result = ExcelReader.FromBytes(uploadedBytes, "sales.xlsx")
    .Sheet("Sales")
    .StartAt("B4")
    .Columns(c => c
        .IncludeHeaders("Invoice No", "Customer", "Gross")
        .ExcludeHeaders("Internal Notes"))
    .ReadDictionaryResult();
```

## CSV Export

CSV export is built around `CsvWriter`.

Data sources:

- `FromDataTable(dataTable)`
- `FromObjects(items)`
- `FromRows(dictionaryRows)`

Options:

- Custom delimiter with `DelimitedBy`.
- Custom quote character.
- Custom encoding.
- Custom newline.
- Include or suppress headers.
- Column order.
- Header aliases.
- Formula text safety.
- Progress, cancellation and audit.

By default, text that could be interpreted as a spreadsheet formula is neutralized for safety. Use `AllowFormulaText()` only when the caller explicitly trusts the output target.

Output shapes:

- `ToText()`
- `WriteTo(TextWriter)`
- `ToBytes()`
- `SaveTo(Stream)`
- `SaveAs(path)`

Example:

```csharp
CsvWriter.FromDataTable(table)
    .DelimitedBy(";")
    .ColumnOrder("Id", "Name", "Amount")
    .HeaderAlias("Id", "Reference")
    .HeaderAlias("Amount", "Invoice Amount")
    .SaveAs(path);
```

## CSV Import

CSV import is built around `CsvReader`.

Input sources:

- `FromFile(path)`
- `FromStream(stream, label)`
- `FromBytes(bytes, label)`
- `FromText(text, label)`

Options:

- Custom delimiter.
- Custom quote character.
- Custom encoding.
- Header row on/off.
- Trim fields.
- Skip empty lines.
- Duplicate header behavior.
- Row length mismatch behavior.
- Include/exclude columns by letters.
- Include/exclude columns by headers.
- Progress, cancellation and audit.

Output shapes:

- Lazy `EnumerateRows()`.
- `ReadRows()`.
- `ReadResult()`.
- `ReadDictionaries()`.
- `ReadDictionaryResult()`.
- `ReadDataTable(tableName)`.
- `ReadObjects<T>()` and `EnumerateObjects<T>()` for typed objects, with problems collected in a list.

Example:

```csharp
var table = CsvReader.FromFile(path)
    .DelimitedBy(",")
    .HasHeaderRow()
    .Columns(c => c
        .IncludeLetters("A", "B", "E")
        .ExcludeHeaders("InternalOnly"))
    .ReadDataTable("ImportedCsv");
```

## Template-Driven Exports

Templates are the API-friendly alternative to the fluent API. A template describes the workbook or CSV output as data. The caller binds named data sources and optional parameters to that template.

This is the preferred shape for Web API usage because it avoids exposing the whole fluent C# API over HTTP.

### Excel Template Concepts

An Excel template can include:

- `schemaVersion`
- `kind`
- `styleCatalog`
- `dataSources`
- `parameters`
- `sheetTemplates`
- `sheets`
- Workbook protection.

Supported Excel template region types:

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
    "salesSheet": {
      "regions": [
        { "type": "cell", "cell": "B1", "value": "Sales Export", "style": "title" },
        {
          "type": "table",
          "source": "sales",
          "startCell": "B4",
          "columns": [
            { "source": "Id", "caption": "Reference" },
            { "source": "Region" },
            { "source": "Amount", "caption": "Invoice Amount", "style": "money" }
          ],
          "styles": {
            "header": "header",
            "data": "money"
          },
          "totals": {
            "label": "Totals",
            "style": "header"
          }
        }
      ]
    }
  },
  "sheets": [
    {
      "name": "Sales",
      "template": "salesSheet"
    }
  ]
}
```

### CSV Template Concepts

A CSV template names a single data source, output settings and optional projected columns.

Example:

```json
{
  "source": "sales",
  "output": {
    "delimiter": "|",
    "includeHeaders": true
  },
  "columns": [
    { "source": "Id", "caption": "Reference" },
    { "source": "Name", "caption": "Customer" },
    { "source": "Amount", "caption": "Invoice Amount" }
  ]
}
```

## Template-Driven Imports

Excel and CSV template import builders exist for API-style imports:

- `ExcelTemplateReader.FromJson(templateJson)`
- `CsvTemplateReader.FromJson(templateJson)`

They support binding a source file and returning dictionaries or a `DataTable`.

Excel import template sources:

- `FromFile(path)`
- `FromBytes(bytes, label)`
- `FromStream(stream, label)`

CSV import template sources:

- `FromFile(path)`
- `FromText(text, label)`
- `FromBytes(bytes, label)`
- `FromStream(stream, label)`

## Web API Contracts

The library includes DTOs in `DocWeave.WebApi` for a host application to expose import/export services.

Current records:

- `DocWeaveExportRequest`
- `DocWeaveImportRequest`
- `DocWeaveImportSource`
- `DocWeaveExportDeliveryRequest`
- `DocWeaveApiAuditMetadata`
- `DocWeaveOperationAcceptedResponse`
- `DocWeaveOperationStatusResponse`

The intended API model is:

- For exports, accept a format, a stored template id or inline template JSON, named data sources, parameters, delivery settings and audit metadata.
- For imports, accept a format, an import source, an optional template id or inline template JSON, and audit metadata.
- For long-running jobs, return an accepted response and expose status polling.

The Web API layer should generally expose templates and data, not the entire fluent interface.

## Telemetry, Audit and Cancellation

Both Excel and CSV flows include an `DocWeaveOperationContext` through public builder methods:

- `WithProgress(IProgress<DocWeaveOperationProgress>)`
- `WithCancellation(CancellationToken)`
- `Audit(Action<DocWeaveAuditEvent>)`

Progress records contain:

- Operation
- Stage
- Completed
- Total

Audit records contain:

- Operation
- Message
- Properties

Audit events are intended for operational facts such as file name, sheet count, row count, caller and correlation id. They should not contain sensitive cell values.

## Performance and Volume Coverage

The test project includes CSV and Excel volume tests.

Excel volume tests include:

- Tall styled tables with thousands of rows and 40+ columns.
- Very wide tables with 120+ columns.
- Multi-sheet workbooks with varied column counts.
- Styling, borders, formulas and totals rows.

Large Excel volume tests use conservative default sizes so the normal test suite stays fast. Set `DOCWEAVE_RUN_LARGE_EXCEL_VOLUME=1` to run larger variants.

CSV volume tests cover large row counts, including million-row scenarios.

## Known Gaps and Cautions

- Native PivotTable output is not yet production-ready for complex cases.
- Export file encryption is not currently implemented. Workbook/sheet protection is different from encryption.
- Password-opening encrypted imports are represented in API contracts but not fully implemented in the fluent readers.
- Macro creation is not supported.
- `.xlsm` macro preservation is not implemented as a general API guarantee.
- `.xls` is intentionally excluded.
- Auto-fit settings are written as workbook metadata; Excel may only visually recalculate exact widths/heights when the workbook is opened.
- Cancellation for Excel export is currently checked before rendering starts, not continuously during every row/cell write.
