# DocWeave Fluent API Capabilities

> **Status: design document.** This describes the intended API surface. It is a wish list, not a reference: many of the methods and types it shows (for example `ExcelReport`, `Sparklines`, `Slicers`, `RequirePasswordToOpen`) are not implemented. The current API is in [getting-started.md](getting-started.md) and [pre-release-functionality-guide.md](pre-release-functionality-guide.md), and what is missing is summarised in [roadmap.md](roadmap.md).

## Purpose

This document describes the intended fluent API surface for DocWeave as a new library. It focuses on the import/export capabilities the library should support, including Excel `.xlsx`, CSV, streams, byte arrays, object collections, `DataTable`, `DataSet`, formatting, formulas, conditional formatting, templates, and other structured data options.

The goal is to design the fluent API around clear user intent:

```csharp
ExcelWriter.Create()
    .AddSheet("Report", sheet => sheet
        .AddTable(rows)
        .StartAt("A1"))
    .SaveAs(path);
```

Rather than around OpenXML implementation details.

## Multi-Format Output Strategy

Excel should be the first and richest output target, but the library should not be designed as an Excel-only writer. The public model should separate data, layout intent, and rendering so that additional output formats can be added without redesigning the fluent API.

Conceptually:

```text
Source data -> Data adapter -> Document/table spec -> Renderer -> Output format
```

Primary output formats:

- `.xlsx` for fully formatted Excel workbooks.
- `.csv` for simple tabular interchange.
- `.tsv` and custom-delimited text for system integration.
- `.json` and JSON Lines for structured data pipelines.
- `.xml` for systems that still require XML payloads.

Later output formats:

- Word documents generated from source data, templates, report sections, summary tables, and repeated content blocks.
- PowerPoint decks generated from source data, slide templates, charts, KPI panels, and repeated slide groups.
- PDF output, likely through a dedicated renderer or through document conversion when available.
- HTML reports for previewing, email bodies, or web delivery.

Delivery targets are separate from output formats. The renderer should create a file, stream, or byte array first; then a delivery adapter can send or store it.

Potential delivery targets:

- HTTP response.
- File or network path.
- Blob/object storage.
- Email attachment or email link.
- Microsoft Teams chat/channel attachment or link.
- SharePoint or OneDrive document library.
- Webhook callback with a file reference.
- Message queue with a file reference.

This implies a shared core model with format-specific renderers:

```text
WorkbookSpec / ReportSpec
    -> ExcelRenderer
    -> CsvRenderer
    -> JsonRenderer
    -> WordRenderer
    -> PowerPointRenderer
    -> HtmlRenderer
```

Not every renderer should support every layout feature. CSV, for example, should ignore fonts, borders, merged cells, formulas, and conditional formatting, while Excel should support the full table and layout model. Word and PowerPoint should reuse the same data binding and schema ideas, but expose document and slide concepts rather than pretending everything is a worksheet.

Initial Excel file format scope:

- Support `.xlsx` import/export.
- Support `.csv` import/export.
- Support `.xltx` as an Excel template input that produces `.xlsx`.
- Support `.xlsm` and `.xltm` only for preserving existing macro-enabled workbooks/templates when the target output is `.xlsm`.
- Do not support creating, editing, inspecting, or signing VBA macros in the initial version.
- Do not support legacy `.xls`.
- Do not use Excel Interop/COM automation.

## Export Entry Points

### File Output

```csharp
ExcelReport.Create()
    .ToFile(@"C:\Exports\Report.xlsx")
    .Sheet("Data", sheet => sheet.From(rows))
    .Save();
```

### Stream Output

```csharp
using var stream = new MemoryStream();

ExcelReport.Create()
    .ToStream(stream)
    .Sheet("Data", sheet => sheet.From(rows))
    .Save();
```

### Byte Array Output

```csharp
byte[] content = ExcelReport.Create()
    .ToByteArray()
    .Sheet("Data", sheet => sheet.From(rows))
    .Save();
```

### Template-Based Output

```csharp
ExcelReport.Create()
    .UseTemplate(@"C:\Templates\Audit.xlsx")
    .ToFile(@"C:\Exports\Audit.xlsx")
    .Sheet("Audit Trail", sheet => sheet
        .From(auditRows)
        .StartAt("A2")
        .Headers(h => h.Hide()))
    .Save();
```

## Export Data Sources

### `IEnumerable<T>`

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Sheet("Cities", sheet => sheet
        .From(cities));
```

Supported behavior:

- Infer columns from public properties.
- Honor `[DisplayName]` or custom attributes.
- Preserve explicit column ordering if configured.
- Format values by type or column.
- Support nullable values.

### `IEnumerable<object>`

```csharp
.Sheet("Objects", sheet => sheet
    .From(objects));
```

Useful for dynamically built object lists where each object has public properties.

### `DataTable`

```csharp
.Sheet("DataTable", sheet => sheet
    .From(dataTable));
```

Expected behavior:

- Use `DataColumn.ColumnName` as default header.
- Use `DataColumn.DataType` for formatting.
- Preserve table column order.

### `DataSet`

Option 1: one worksheet per table.

```csharp
ExcelReport.Create()
    .ToFile(path)
    .SheetsFrom(dataSet);
```

Option 2: multiple tables on one worksheet.

```csharp
.Sheet("All Tables", sheet => sheet
    .Sources(sources => sources
        .Add(dataSet.Tables[0])
        .Add(dataSet.Tables[1])
        .Layout(l => l.StackVertically(spacingRows: 2))));
```

Option 3: named tables and relationships.

```csharp
ExcelReport.Create()
    .ToFile(path)
    .DataSet(dataSet, data => data
        .Table("Regions").ToSheet("Regions")
        .Table("Costs").ToSheet("Costs")
        .Table("Income").ToSheet("Income")
        .UseRelations())
    .Save();
```

Expected behavior:

- Export all tables or a selected subset.
- Use `DataTable.TableName` as a default sheet/table name.
- Preserve `DataColumn` order and type metadata.
- Optionally use `DataRelation` metadata for grouped layouts.
- Allow each table to become an Excel table so it can feed formulas, validation, charts, or pivot tables.

### `IEnumerable<DataRow>`

```csharp
.Sheet("Rows", sheet => sheet
    .From(dataTable.AsEnumerable()));
```

Expected behavior:

- Infer columns from the parent `DataTable`.
- Preserve `DataColumn` metadata where available.

### Dictionaries

```csharp
IEnumerable<IDictionary<string, object?>> records = GetRecords();

.Sheet("Dictionary", sheet => sheet
    .From(records));
```

Expected behavior:

- Use dictionary keys as column names.
- Infer column order from first row unless configured.
- Fill missing keys with blank/default values.

### Dynamic / ExpandoObject

```csharp
.Sheet("Dynamic", sheet => sheet
    .From(expandoRows));
```

This should normalize through the dictionary adapter.

### Custom Source Adapter

```csharp
.Sheet("Custom", sheet => sheet
    .From(new MyRowSource()));
```

Suggested interface:

```csharp
public interface IRowSource
{
    IReadOnlyList<ColumnDefinition> Columns { get; }
    IEnumerable<RowValueSet> Rows { get; }
}
```

## Export Sheet Options

```csharp
.Sheet("Report", sheet => sheet
    .StartAt(row: 2, column: 1)
    .TabColor("004D17")
    .Freeze(row: 1)
    .Filter()
    .AutoFitColumns()
    .HideGridlines()
    .Landscape());
```

Options:

- `StartAt("A1")`
- `StartAt(row, column)`
- `TabColor(hex)`
- `SheetFill(hex)`
- `Freeze(row)`
- `Freeze(row, column)`
- `Filter()`
- `AutoFitColumns()`
- `UseSourceColumnWidths()`
- `HideGridlines()`
- `PageSetup(...)`
- `PrintArea("A1:H50")`
- `Landscape()`
- `Portrait()`

## Export Column Options

```csharp
.Columns(columns => columns
    .Include(x => x.Name)
    .Include(x => x.Country)
    .Ignore(x => x.InternalId)
    .Rename(x => x.Name, "City Name")
    .Order("City Name", "Country", "Population")
    .Width("Country", 20)
    .AutoFit()
    .UseDisplayNameAttributes());
```

Options:

- Include/exclude columns.
- Rename headers.
- Explicit column order.
- DisplayName attribute support.
- Width per column.
- Auto-fit all columns.
- Minimum/maximum column width.
- Hidden columns.
- Type-based formats.
- Column-specific formats.
- Column-specific styles.

## Export Header Options

```csharp
.Headers(headers => headers
    .Show()
    .Freeze()
    .Filter()
    .Style(style => style
        .Foreground("FFFFFF")
        .Background("004D17")
        .Bold()));
```

Options:

- Show/hide headers.
- Freeze header row.
- Apply auto-filter.
- Header row style.
- Custom header names.
- Multi-row headers later if needed.

## Multiple Data Sources on a Sheet

The workbook should support multiple independently configured export regions. A region is any exported source or layout unit placed on a sheet, such as a table, summary range, repeated block, pivot table, chart, or template-bound area.

Sources can come from different types in the same export:

- one or more `DataTable` instances from a `DataSet`
- standalone `DataTable` instances
- `IEnumerable<T>` object collections
- dictionaries or dynamic rows
- custom `IRowSource` implementations
- summary/scalar values

Each source should be configurable independently, even when several sources appear on the same sheet.

### Vertical Layout

```csharp
.Sheet("Summary", sheet => sheet
    .Sources(sources => sources
        .Add("Customers", customers)
        .Add("Transactions", transactions)
        .Layout(layout => layout.StackVertically(spacingRows: 2))));
```

### Horizontal Layout

```csharp
.Sheet("Side By Side", sheet => sheet
    .Sources(sources => sources
        .Add("Customers", customers)
        .Add("Transactions", transactions)
        .Layout(layout => layout.StackHorizontally(spacingColumns: 2))));
```

### Absolute Placement

```csharp
.Sheet("Dashboard", sheet => sheet
    .Sources(sources => sources
        .Add("Summary", summary).At("B2")
        .Add("Detail", details).At("B10")));
```

### Mixed Sources With Independent Formatting

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Sheet("Operational Summary", sheet => sheet
        .Sources(sources => sources
            .Add("Regions", dataSet.Tables["Regions"])
                .At("A2")
                .AsTable("RegionsTable")
                .Headers(h => h.Style("greenHeader"))
                .Columns(c => c
                    .Width("RegionName", 28)
                    .Format<decimal>("#,##0.00"))
                .Style("greenTable")
            .Add("Forecast", forecastRows)
                .At("H2")
                .Headers(h => h.Style("blueHeader"))
                .Columns(c => c
                    .Include(x => x.Month)
                    .Include(x => x.ExpectedIncome)
                    .Include(x => x.ExpectedCost)
                    .Format<decimal>("#,##0.00"))
                .Totals(t => t
                    .Sum("ExpectedIncome")
                    .Sum("ExpectedCost"))
                .Style("blueTable")
            .Add("RunInfo", runSummary)
                .At("A20")
                .AsKeyValueBlock()
                .Style("summaryBlock"))))
    .Save();
```

### DataSet Selection And Placement

```csharp
ExcelReport.Create()
    .ToFile(path)
    .DataSet(dataSet, data => data
        .Table("Regions")
            .ToSheet("Reference Data")
            .At("A1")
            .Style("referenceTable")
        .Table("Costs")
            .ToSheet("Financials")
            .At("A2")
            .Style("costTable")
            .Totals(t => t.Sum("Amount"))
        .Table("Income")
            .ToSheet("Financials")
            .At("H2")
            .Style("incomeTable")
            .Totals(t => t.Sum("Amount"))
        .Ignore("InternalAuditScratch")))
    .Save();
```

Export region options:

- Source name.
- Source object.
- Sheet target.
- Absolute placement.
- Relative layout after another region.
- Stack vertically or horizontally.
- Convert to Excel table.
- Header behavior.
- Column inclusion, ordering, naming, width, and formatting.
- Styles per source/region.
- Totals and formulas per source/region.
- Conditional formatting per source/region.
- Empty-source behavior per source/region.
- Named ranges per source/region.
- Visibility and print behavior per region.

## Data Groups

Data groups allow multiple sources to share a style, layout, totals, formulas, and other behavior.

```csharp
.Sheet("Risk", sheet => sheet
    .Group("Exposure", group => group
        .Style(ReportStyles.Green)
        .Source("FX", fxRows)
        .Source("Money", moneyRows)
        .Headers()
        .AlternatingRows()
        .Totals("Exposure", "Limit")));
```

Group options:

- Shared style.
- Shared headers.
- Shared alternating row behavior.
- Shared totals.
- Shared formulas.
- Shared first row/column.
- Template range mapping.
- Default placement/layout for sources in the group.
- Per-source overrides for style, headers, columns, totals, formulas, and placement.
- Mixed source types in the same group.
- Group-level title, spacing, borders, and print behavior.

Example with group defaults and per-source overrides:

```csharp
.Sheet("Financial Pack", sheet => sheet
    .Group("Financials", group => group
        .StartAt("A2")
        .StackVertically(spacingRows: 3)
        .Style("financialDefault")
        .Headers(h => h.Show().Style("financialHeader"))
        .Source("Costs", dataSet.Tables["Costs"], source => source
            .Style("costTable")
            .Totals(t => t.Sum("Amount")))
        .Source("Income", incomeRows, source => source
            .Style("incomeTable")
            .Columns(c => c.Format<decimal>("#,##0.00"))
            .Totals(t => t.Sum("Amount")))
        .Source("Commentary", commentaryRows, source => source
            .Style("commentaryTable")
            .Columns(c => c.Width("Comment", 60).Wrap("Comment")))));
```

## Repeating Data Blocks

Repeating data blocks support report layouts where each source item is rendered as a shaped block instead of a single table row. This is useful for data-repeater style output, such as one block per region, site, department, customer, project, or account.

Example source:

```csharp
public sealed class RegionSummary
{
    public string Region { get; init; } = "";
    public decimal Costs { get; init; }
    public decimal Income { get; init; }
    public decimal Profit => Income - Costs;
}
```

Example output pattern:

```text
London               United States          Europe
Costs   120,000      Costs   450,000        Costs   310,000
Income  180,000      Income  610,000        Income  480,000
Profit   60,000      Profit  160,000        Profit  170,000

Asia                 Africa                 Australia
Costs   210,000      Costs    80,000        Costs   160,000
Income  330,000      Income  120,000        Income  240,000
Profit  120,000      Profit   40,000        Profit   80,000
```

Fluent example:

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Sheet("Regional Summary", sheet => sheet
        .RepeatBlocks(regions, blocks => blocks
            .StartAt("A2")
            .PerRow(3)
            .BlockSize(rows: 4, columns: 3)
            .Spacing(rows: 1, columns: 2)
            .ForEach((block, region) => block
                .Cell("A1").Value(region.Region).Style("regionTitle").MergeAcross(3)
                .Cell("A2").Value("Costs").Style("metricLabel")
                .Cell("B2").Value(region.Costs).Format("#,##0.00").Style("metricValue")
                .Cell("A3").Value("Income").Style("metricLabel")
                .Cell("B3").Value(region.Income).Format("#,##0.00").Style("metricValue")
                .Cell("A4").Value("Profit").Style("metricLabel")
                .Cell("B4").Formula("=B3-B2").Format("#,##0.00").Style("metricValue"))))
    .Save();
```

Template-based block example:

```csharp
.Sheet("Regional Summary", sheet => sheet
    .RepeatBlocks(regions, blocks => blocks
        .UseTemplateRange("Templates!A1:C4")
        .StartAt("A2")
        .PerRow(2)
        .Spacing(rows: 2, columns: 1)
        .Bind(bind => bind
            .Value("RegionName", x => x.Region)
            .Value("Costs", x => x.Costs)
            .Value("Income", x => x.Income)
            .Formula("Profit", "{Income}-{Costs}"))));
```

Repeating block options:

- `PerRow(count)` to control how many blocks appear across before wrapping.
- Fixed `BlockSize(rows, columns)`.
- Automatic block size from a template range.
- Row and column spacing between blocks.
- Start cell.
- Page breaks after N block rows.
- Repeat section headers.
- Keep each block together for printing.
- Bind scalar properties.
- Bind nested collections as mini-tables inside a block.
- Use formulas relative to the current block.
- Use named placeholders or template cells.
- Apply styles, borders, wrapping, merges, and conditional formatting within each block.
- Hide template sheets after rendering.

Repeating blocks are different from normal tables:

- Table export maps each item to a row.
- Repeating block export maps each item to a rectangular layout region.
- A block can contain labels, values, formulas, mini-tables, charts, and styled ranges.

## Styling

### Style Presets

```csharp
.Style(ReportStyles.Green)
.Style(ReportStyles.Plain)
.Style(ReportStyles.Audit)
.Style(ReportStyles.Financial)
```

### Inline Style Builder

```csharp
.Style(style => style
    .Title(t => t
        .Foreground("FFFFFF")
        .Background("004D17")
        .FontSize(12)
        .Bold())
    .Header(h => h
        .Foreground("FFFFFF")
        .Background("004D17")
        .BorderBottom(BorderWeight.Thick))
    .Rows(r => r
        .Foreground("000000")
        .Background("FFFFFF")
        .Alternating("FFFFFF", "CCE1C1"))
    .Totals(t => t
        .Foreground("FFFFFF")
        .Background("004D17")
        .Bold()));
```

Style areas:

- Title.
- Header.
- Data rows.
- Alternating rows.
- Totals.
- Group blocks.
- Empty-source messages.
- Conditional formats.

Style properties:

- Font family.
- Font size.
- Bold/italic/underline.
- Foreground color.
- Background/fill color.
- Borders.
- Alignment.
- Number format.
- Date format.
- Wrap text.

## Number and Date Formatting

```csharp
.Formats(formats => formats
    .Date("yyyy-MM-dd")
    .DateTime("yyyy-MM-dd HH:mm:ss")
    .Decimal("#,##0.00")
    .Currency("£#,##0.00")
    .Percentage("0.00%")
    .Column("Amount", "#,##0.00")
    .Column("RunDate", "yyyy-MM-dd"));
```

Options:

- Type-level defaults.
- Column-level overrides.
- Culture-aware formatting.
- Excel number formats.
- Preserve raw numeric/date cell values rather than strings.

## Totals and Subtotals

```csharp
.Totals(totals => totals
    .Sum("Amount")
    .Average("Rate")
    .Count("Account")
    .Label("Total"));
```

Options:

- Sum.
- Average.
- Count.
- Min/max.
- Custom formula.
- Total row style.
- Per-group totals.

## Excel Formulas

### Column Formulas

```csharp
.Formulas(formulas => formulas
    .Column("Total", "=SUM([Amount],[Tax])")
    .Column("Variance", row => $"C{row}-D{row}"));
```

### Row Formulas

```csharp
.Formulas(formulas => formulas
    .RowTotal("Total", columns: new[] { "Jan", "Feb", "Mar" }));
```

### Total Row Formulas

```csharp
.Totals(totals => totals
    .Formula("Amount", ctx => $"SUM({ctx.Column}{ctx.StartRow}:{ctx.Column}{ctx.EndRow})"));
```

Supported placeholders:

- `{row}`
- `{start}`
- `{end}`
- `{column:Name}`
- `{cell:ColumnName}`

Formula goals:

- Avoid callers manually calculating cell references where possible.
- Support custom formulas for advanced cases.
- Keep generated formulas valid Excel formulas.

## Pivot Tables

Pivot tables should be supported as an Excel-specific export feature. A common scenario is exporting the source data to one worksheet and creating one or more pivot tables on separate summary worksheets.

Basic example:

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Sheet("Data", sheet => sheet
        .From(sales)
        .AsTable("SalesData")
        .Hide())
    .Sheet("Summary", sheet => sheet
        .PivotTable("SalesByRegion", pivot => pivot
            .SourceTable("SalesData")
            .StartAt("A3")
            .Rows("Region")
            .Columns("Month")
            .Values(values => values
                .Sum("Amount", "Total Sales")
                .Count("OrderId", "Orders"))
            .Filters("Year")
            .Style("PivotStyleMedium9")
            .RefreshOnOpen()))
    .Save();
```

Alternative source-range example:

```csharp
.Sheet("Summary", sheet => sheet
    .PivotTable("DepartmentSpend", pivot => pivot
        .SourceRange("Data!A1:H500")
        .StartAt("B4")
        .Rows("Department", "CostCentre")
        .Columns("Quarter")
        .Values(v => v.Sum("Spend"))
        .ReportLayout(PivotReportLayout.Tabular)
        .ShowGrandTotals(rows: true, columns: true)));
```

Pivot options:

- Source from an exported Excel table.
- Source from a named range.
- Source from an explicit sheet range.
- Place pivot tables on the same sheet or another sheet.
- Hide or keep visible the data/source sheet.
- Row fields.
- Column fields.
- Value fields.
- Page/filter fields.
- Sum, count, average, min, max, product, count numbers, standard deviation, and variance aggregations.
- Custom captions for value fields.
- Number formats on value fields.
- Tabular, outline, or compact report layout.
- Row and column grand totals.
- Subtotals per row field.
- Repeat item labels.
- Expand/collapse defaults where OpenXML supports it.
- Pivot table style names.
- Refresh on open.
- Preserve or discard cached records, depending on file size and compatibility needs.

Pivot tables should prefer a table-backed source such as `.AsTable("SalesData")`, because the source can grow without recalculating hard-coded cell ranges. Range-backed pivots are still useful for template scenarios.

## Charts, Graphs, And Visual Elements

Excel export should support chart and visual regions that can be placed on the same sheet as the data, on dashboard sheets, or beside pivot tables. Charts should be modeled as export regions so they can use the same placement and layout behavior as tables, repeated blocks, and pivots.

Basic chart example:

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Sheet("Data", sheet => sheet
        .From(sales)
        .AsTable("SalesData"))
    .Sheet("Dashboard", sheet => sheet
        .Chart("SalesTrend", chart => chart
            .SourceTable("SalesData")
            .Type(ChartType.Line)
            .Category("Month")
            .Series("Amount", "Sales")
            .At("B3")
            .Size(width: 720, height: 320)
            .Title("Sales Trend")
            .Legend(ChartLegendPosition.Bottom)))
    .Save();
```

Pivot chart example:

```csharp
.Sheet("Dashboard", sheet => sheet
    .PivotTable("SalesByRegion", pivot => pivot
        .SourceTable("SalesData")
        .Rows("Region")
        .Values(v => v.Sum("Amount")))
    .Chart("SalesByRegionChart", chart => chart
        .SourcePivot("SalesByRegion")
        .Type(ChartType.Column)
        .At("H3")
        .Title("Sales by Region")));
```

Useful chart types:

- Column and stacked column.
- Bar and stacked bar.
- Line.
- Area.
- Pie and doughnut.
- Scatter.
- Bubble.
- Combo chart with secondary axis.
- Waterfall later if practical.
- Histogram/Pareto later if practical.
- Stock charts later if practical.

Chart options:

- Source from table, range, named range, pivot table, or explicit series.
- Category axis.
- Multiple series.
- Secondary axis.
- Chart title.
- Axis titles.
- Legend position.
- Data labels.
- Number formats.
- Colors and theme styles.
- Size and placement.
- Move and resize behavior.
- Hidden source sheet support.

### Sparklines

Sparklines are useful for compact dashboards and table rows.

```csharp
.Sparklines(s => s
    .Line("Trend", sourceRange: "B2:M2", targetCell: "N2")
    .ForEachRow(sourceColumns: "B:M", targetColumn: "N"));
```

Sparkline options:

- Line, column, and win/loss sparklines.
- Per-row sparkline generation.
- Group styling.
- High/low/negative markers.

### Slicers And Timelines

Slicers are useful when exported workbooks include tables or pivot tables.

```csharp
.Slicers(slicers => slicers
    .ForTable("SalesData")
    .Field("Region")
    .At("B2"))
.Timeline("SalesTimeline", timeline => timeline
    .ForPivot("SalesByRegion")
    .Field("OrderDate")
    .At("E2"));
```

Support for slicers and timelines should be later-phase because OpenXML support is more complex and compatibility-sensitive.

### Images, Logos, And Icons

```csharp
.Image("CompanyLogo", image => image
    .FromFile(logoPath)
    .At("A1")
    .Size(width: 160, height: 48)
    .AltText("Company logo"));
```

Image options:

- File, stream, or byte array source.
- PNG, JPEG, GIF, BMP where supported by Excel.
- Cell anchor.
- Absolute size.
- Preserve aspect ratio.
- Alt text.
- Header/footer logo later if needed.

### Shapes, Diagrams, And Connectors

For diagrams, the library can support a practical subset of Excel drawing shapes rather than trying to become a full drawing tool.

```csharp
.Diagram("ApprovalFlow", diagram => diagram
    .At("A8")
    .Shape("Request", shape => shape
        .Text("Request")
        .Rectangle()
        .Size(120, 50)
        .Style("processShape"))
    .Shape("Review", shape => shape
        .Text("Review")
        .Rectangle()
        .RightOf("Request", spacing: 40)
        .Style("processShape"))
    .Connector("Request", "Review", ConnectorType.Arrow));
```

Diagram options:

- Rectangles, rounded rectangles, lines, arrows, callouts, and text boxes.
- Connectors between named shapes.
- Basic flowchart layout helpers.
- Manual placement.
- Style presets.
- Grouping shapes.
- Text wrapping and alignment.

SmartArt should be considered out of scope for early versions unless a strong use case appears.

### Hyperlinks, Comments, And Notes

```csharp
.Hyperlinks(links => links
    .Cell("A2").ToUrl("https://example.com")
    .Column("DocumentPath").ToFilePath())
.Comments(comments => comments
    .Cell("C4").Text("Reviewed by Finance"));
```

Options:

- URL links.
- File links.
- Internal sheet/range links.
- Cell comments/notes.
- Author metadata.

### Data Validation And Drop-Down Lists

```csharp
.Validation(v => v
    .Column("Status").List("Open", "Closed", "Pending")
    .Column("Amount").DecimalBetween(0, 1000000)
    .Column("RunDate").DateBetween(startDate, endDate));
```

Validation options:

- List validation.
- Whole number and decimal ranges.
- Date ranges.
- Text length.
- Custom formula validation.
- Input messages.
- Error messages.

### Workbook Protection And Review Features

```csharp
.Protect(protection => protection
    .ProtectWorkbook(password: "structure-password")
    .ProtectSheet("Data", password: "sheet-password", allowFiltering: true, allowSorting: true)
    .UnlockRange("Data", "InputCells"));
```

Options:

- Protect workbook structure.
- Protect worksheets.
- Sheet/workbook protection passwords.
- Locked/unlocked cells.
- Allow sorting, filtering, formatting, inserting rows, or selecting unlocked cells.
- Hide formulas.
- Very hidden sheets for internal source data.

Protection note:

- Worksheet/workbook protection helps prevent accidental or casual editing.
- It is not the same as encrypting the file.
- Sensitive exports should use file encryption when the workbook must require a password to open.

### Workbook File Encryption

For sensitive exports, callers should be able to request full workbook encryption so the generated `.xlsx` requires a password to open.

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Encrypt(encryption => encryption
        .OpenPassword(openPassword)
        .Algorithm(ExcelEncryptionAlgorithm.Aes256)
        .RequirePasswordToOpen())
    .Sheet("Sensitive Data", sheet => sheet
        .From(rows)
        .Protect(p => p.ProtectSheet(password: sheetPassword)))
    .Save();
```

Encryption options:

- Password to open the workbook.
- Optional password to modify, if supported.
- Encryption algorithm policy, such as AES-256 where available.
- Encrypt the whole workbook package.
- Combine file encryption with sheet/workbook protection.
- Mark selected sheets as hidden or very hidden before encryption.

Implementation note:

- Excel file encryption is different from normal OpenXML sheet protection.
- Full `.xlsx` package encryption may require a dedicated encryption service or dependency because plain OpenXML package writing does not by itself encrypt the workbook for opening.
- Partial sheet encryption is not an Excel feature in the same way full file encryption is. For sheet-level confidentiality, use separate encrypted files, hidden/very-hidden sheets plus workbook encryption, or split sensitive data into separate outputs.

## Conditional Formatting

```csharp
.ConditionalFormatting(cf => cf
    .Range("E2:E100")
    .GreaterThan(0)
    .Fill("CCE1C1"));
```

### Value Rules

```csharp
.ConditionalFormatting(cf => cf
    .Column("Exposure").GreaterThan(1000000).Fill("FFC7CE")
    .Column("Limit").LessThan(0).Fill("FFEB9C")
    .Column("Status").TextContains("FAIL").Fill("FFC7CE"));
```

### Built-In Rule Types

- Greater than.
- Less than.
- Between.
- Equal/not equal.
- Text contains.
- Starts with / ends with.
- Top N.
- Bottom N.
- Above average.
- Below average.
- Duplicate values.
- Unique values.

### Visual Rules

- Color scales.
- Data bars.
- Icon sets.

Example:

```csharp
.ConditionalFormatting(cf => cf
    .Column("Utilisation").DataBar("63C384")
    .Column("Score").ColorScale("F8696B", "FFEB84", "63BE7B"));
```

## Filters, Sorting, and Freeze Panes

```csharp
.Sheet("Data", sheet => sheet
    .Filter()
    .Freeze(row: 1)
    .Sort(sort => sort
        .By("Department")
        .ThenByDescending("Amount")));
```

Sorting can be implemented in two ways:

- Sort source data before rendering.
- Add workbook sort metadata later if needed.

The first version should sort data before rendering because it is simpler and deterministic.

## Merges and Titles

```csharp
.Title("Monthly Risk Report", title => title
    .At("A1")
    .MergeAcross(8)
    .Style(ReportStyles.Title));
```

Options:

- Add sheet title.
- Merge title cells.
- Style title.
- Place title above data.

## Empty Source Behavior

```csharp
.EmptySource(empty => empty
    .ShowMessage("No records found")
    .Style(ReportStyles.Empty));
```

Options:

- Show message.
- Hide sheet.
- Create headers only.
- Skip source.
- Throw error.

## Template Export

### File Template

```csharp
ExcelReport.Create()
    .UseTemplate(@"C:\Templates\Risk.xlsx")
    .ToFile(outputPath)
    .Template(template => template
        .CopySheet("Risk Template").As("Risk Report")
        .MapCell("Report_Date").To(DateTime.Today)
        .MapRange("FX_Table").To(fxRows))
    .Save();
```

### Stream Template

```csharp
ExcelReport.Create()
    .UseTemplate(templateStream)
    .ToStream(outputStream)
    .Template(template => template
        .MapRange("Data_Table").To(rows))
    .Save();
```

### Template Features

- Copy whole sheet.
- Copy source sheet to target sheet name.
- Populate named cells.
- Populate named ranges.
- Populate Excel tables.
- Insert rows into ranges.
- Preserve source formatting.
- Repeat card/region layouts.
- Map multiple data sources into one sheet.
- Map one data source into multiple regions.

## Import Entry Points

Import should support both direct library usage and WebAPI/server usage. Source files may arrive as paths, network paths, streams, byte arrays, uploaded files, links, or batches of files.

### XLSX from File

```csharp
var rows = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

### XLSX from Password-Protected or Encrypted File

```csharp
var rows = ExcelReader.FromFile(path)
    .Password(openPassword)
    .Sheet("Sensitive Data")
    .UseHeaderRow(1)
    .Read<SensitiveRecord>();
```

If a workbook is protected but not encrypted, a password may be needed only for protected structures or sheets:

```csharp
var rows = ExcelReader.FromFile(path)
    .WorkbookPassword(workbookPassword)
    .Sheet("Protected Sheet", sheet => sheet
        .Password(sheetPassword)
        .UseHeaderRow(1))
    .Read<ProtectedRecord>();
```

Import password options:

- Password to open an encrypted workbook.
- Workbook structure password where relevant.
- Sheet protection password where relevant.
- Per-source password for batch imports.
- Password provider callback so the password is not stored in schema files.

### XLSX from Network Path

```csharp
var rows = ExcelReader.FromFile(@"\\server\share\imports\cities.xlsx")
    .Label("CitiesImport")
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

### XLSX from Stream

```csharp
var rows = ExcelReader.FromStream(stream)
    .Label("uploaded-cities.xlsx")
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

### XLSX from Byte Array

```csharp
var rows = ExcelReader.FromBytes(bytes)
    .Label("cities.xlsx")
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

### XLSX from Uploaded File

For ASP.NET Core, the WebAPI wrapper can expose helpers for uploaded files while the core library still consumes a stream.

```csharp
var rows = ExcelReader.FromUpload(file.FileName, file.OpenReadStream())
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

### XLSX from URI or Link

The core library should not hide network/security policy. It can accept a stream directly, or use a caller-provided file resolver.

```csharp
var rows = ExcelReader.FromUri(uri, resolver)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

Resolver examples:

- HTTP/HTTPS URL resolver.
- Internal document service resolver.
- SharePoint or file-store resolver.
- Blob storage resolver.
- SFTP resolver.
- Database/file-id resolver.

### Batch Import From Multiple Files

```csharp
var result = ImportBatch.Create()
    .AddExcel("cities", cityBytes, fileName: "cities.xlsx", password: cityPassword)
    .AddCsv("rates", ratesStream, fileName: "rates.csv")
    .AddExcelPath("products", @"\\server\share\products.xlsx", password: productPassword)
    .Read(batch => batch
        .Source("cities").Sheet("Cities").ReadDataTable()
        .Source("rates").ReadDictionaries()
        .Source("products").Sheet(0).Read<Product>());
```

### CSV from File

```csharp
var rows = CsvReader.FromFile(path)
    .HasHeaderRow()
    .DelimitedBy(",")
    .Read<City>();
```

### CSV from Stream

```csharp
var rows = CsvReader.FromStream(stream)
    .Encoding(Encoding.UTF8)
    .HasHeaderRow()
    .Read<City>();
```

### CSV from Bytes, Upload, URI, or Network Path

CSV import should mirror Excel import source handling:

```csharp
CsvReader.FromBytes(bytes).Label("rates.csv");
CsvReader.FromUpload(file.FileName, file.OpenReadStream());
CsvReader.FromFile(@"\\server\share\imports\rates.csv");
CsvReader.FromUri(uri, resolver);
```

## Import Sheet and Range Options

```csharp
ExcelReader.FromFile(path)
    .Sheet("Rates")
    .Range(range => range
        .FromRow(5)
        .ToRow(110)
        .FromColumn("B")
        .ToColumn("H"))
    .Read<DataTable>();
```

Options:

- Sheet by name.
- Sheet by index.
- First row.
- Last row.
- First column.
- Last column.
- Named range.
- Excel table.
- Used range.
- Skip hidden rows.
- Skip hidden columns.
- Include empty rows.
- Stop after N empty rows.

## Import Column Mapping

Import should allow callers to choose exactly which columns are imported, which columns are ignored, how columns are renamed, what data types are expected, and whether derived columns should be created from imported values.

### Include Selected Columns

```csharp
var table = ExcelReader.FromFile(path)
    .Sheet("Transactions")
    .UseHeaderRow(1)
    .Columns(columns => columns
        .Include("TradeDate")
        .Include("Account")
        .Include("Amount")
        .Include("Currency"))
    .ReadDataTable();
```

### Exclude Columns

```csharp
var table = ExcelReader.FromFile(path)
    .Sheet("Transactions")
    .UseHeaderRow(1)
    .Columns(columns => columns
        .Ignore("InternalId")
        .Ignore("AuditScratch")
        .IgnoreHidden())
    .ReadDataTable();
```

### Rename Imported Columns

```csharp
var table = ExcelReader.FromFile(path)
    .Sheet("Transactions")
    .UseHeaderRow(1)
    .Columns(columns => columns
        .Rename("Trade Date", "TradeDate")
        .Rename("Acct", "Account")
        .Rename("Value", "Amount"))
    .ReadDataTable();
```

### Specify Data Types

```csharp
var table = ExcelReader.FromFile(path)
    .Sheet("Transactions")
    .UseHeaderRow(1)
    .Columns(columns => columns
        .Type("TradeDate", typeof(DateTime))
        .Type("Amount", typeof(decimal))
        .Type("Quantity", typeof(int))
        .Type("IsActive", typeof(bool))
        .Nullable("SettlementDate"))
    .ReadDataTable();
```

### Derived Columns

```csharp
var table = ExcelReader.FromFile(path)
    .Sheet("Transactions")
    .UseHeaderRow(1)
    .Columns(columns => columns
        .Include("Income")
        .Include("Cost")
        .Add("Profit", row => row.Decimal("Income") - row.Decimal("Cost"))
        .Add("ImportStatus", row => row.Decimal("Profit") >= 0 ? "Profit" : "Loss")
        .AddFormula("Margin", "Profit / Income"))
    .ReadDataTable();
```

Derived column options:

- Add a constant value.
- Add a value from a row callback.
- Add a value using a typed row accessor.
- Add a formula-like expression evaluated during import.
- Add source metadata such as file name, sheet name, row number, or import batch id.
- Add validation status columns.

### By Header

```csharp
.MapTo<City>(map => map
    .Property(x => x.Name).FromHeader("City Name")
    .Property(x => x.Country).FromHeader("Country"));
```

### By Column Letter

```csharp
.MapTo<Rate>(map => map
    .Property(x => x.Code).FromColumn("B")
    .Property(x => x.Buy).FromColumn("C")
    .Property(x => x.Sell).FromColumn("D"));
```

### By Column Index

```csharp
.MapTo<Rate>(map => map
    .Property(x => x.Code).FromIndex(1)
    .Property(x => x.Buy).FromIndex(2));
```

### By Attribute

```csharp
.MapTo<City>(map => map
    .UseDisplayNameAttributes());
```

### Typed Object Mapping With Include/Ignore

```csharp
var rows = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .MapTo<City>(map => map
        .Include(x => x.Name)
        .Include(x => x.Country)
        .Ignore(x => x.InternalId)
        .Property(x => x.Population).FromHeader("Population").AsInt32()
        .Property(x => x.IsCapital).FromHeader("Capital").AsBoolean("Y", "N"))
    .Read<City>();
```

## Import Return Types

### Typed Objects

```csharp
List<City> cities = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Read<City>();
```

### DataTable

```csharp
DataTable table = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .ReadDataTable();
```

### Dictionaries

```csharp
IReadOnlyList<IDictionary<string, object?>> rows = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .ReadDictionaries();
```

### Raw Cell Rows

```csharp
IReadOnlyList<RowValueSet> rows = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .ReadRows();
```

### Validation Result

```csharp
ImportResult<City> result = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .Validate<City>();
```

## Import Conversion Options

```csharp
.Conversion(conversion => conversion
    .Culture("en-GB")
    .DateFormat("dd/MM/yyyy")
    .EmptyStringAsNull()
    .TrimStrings()
    .Boolean("Y", "N"));
```

Options:

- Culture.
- Date formats.
- Number formats.
- Boolean true/false values.
- Empty string handling.
- Whitespace trimming.
- Default values.
- Custom converters.

Custom converter:

```csharp
.MapTo<City>(map => map
    .Property(x => x.Population)
    .FromHeader("Population")
    .ConvertWith(value => int.Parse(value.ToString())));
```

## Import Error Handling

```csharp
.OnError(ErrorMode.Throw)
.OnError(ErrorMode.Collect)
.OnError(ErrorMode.UseDefault)
```

Error details should include:

- Source file or stream label.
- Sheet name.
- Row number.
- Column number.
- Column header.
- Raw value.
- Target property.
- Target type.
- Error message.

## CSV Options

```csharp
CsvReader.FromFile(path)
    .DelimitedBy("|")
    .Quote('"')
    .Escape('\\')
    .Encoding(Encoding.UTF8)
    .HasHeaderRow()
    .TrimFields()
    .Read<MyRecord>();
```

Options:

- Delimiter.
- Quote character.
- Escape character.
- Encoding.
- Header row.
- No header row with explicit columns.
- Skip blank lines.
- Trim fields.
- Culture-specific conversion.
- Strict/permissive malformed row handling.

CSV export:

```csharp
CsvWriter.Create()
    .ToFile(path)
    .DelimitedBy(",")
    .Headers()
    .From(rows)
    .Save();
```

## Other Structured File Options

Potential later readers/writers:

- TSV.
- Pipe-delimited.
- Fixed-width.
- JSON array.
- JSON Lines.
- XML flat records.
- YAML records for configuration-style data exchange.
- Parquet.
- Markdown tables.
- HTML tables.

Generic entry point:

```csharp
StructuredData.FromFile(path)
    .AutoDetect()
    .Read<DataTable>();
```

Structured export should use the same idea:

```csharp
StructuredData.From(rows)
    .ToFile(path)
    .AsJsonLines()
    .Save();
```

Suggested export formats:

- `AsCsv()` for comma-delimited files.
- `AsDelimited("|")` for pipe-delimited or custom-delimited files.
- `AsFixedWidth(widths => ...)` for legacy integration files.
- `AsJsonArray()` for one JSON array.
- `AsJsonLines()` for streaming-friendly record output.
- `AsXmlRecords()` for flat XML record sets.
- `AsHtmlTable()` for lightweight report preview.
- `AsMarkdownTable()` for documentation or human-readable output.
- `AsParquet()` only if a dependency is accepted later, because this is outside the OpenXML domain.

Format-specific capabilities should be explicit. CSV and delimited files support headers, encoding, culture, quoting, escaping, null handling, and value conversion. They do not support cell styling, formulas, merges, borders, or conditional formatting.

## Document And Presentation Outputs

Word and PowerPoint generation should be planned as later renderers that reuse the same data binding, schema, and style-token concepts.

Word document scenarios:

- Generate report documents from `IEnumerable<T>`, `DataTable`, `DataSet`, dictionaries, or composed data models.
- Bind scalar values into named placeholders.
- Repeat sections for grouped records.
- Render tables with headers, totals, styles, borders, wrapping, and column widths.
- Insert generated summaries, images, charts, page breaks, headers, footers, and table of contents markers.
- Use `.docx` templates with named content controls or placeholders.

Potential Word API:

```csharp
WordReport.Create()
    .UseTemplate(templatePath)
    .ToFile(outputPath)
    .Value("ReportTitle", "Monthly Sales")
    .Table("SalesTable", rows)
    .Repeat("RegionSection", regions)
    .Save();
```

PowerPoint scenarios:

- Generate decks from data, templates, slide masters, and repeatable slide layouts.
- Bind scalar values into title, subtitle, KPI, and text placeholders.
- Render tables, charts, images, and grouped slide sections.
- Repeat one slide per entity, group, or data slice.
- Keep theme colors, fonts, and layout tokens consistent with the schema model.

Potential PowerPoint API:

```csharp
PowerPointDeck.Create()
    .UseTemplate(templatePath)
    .ToFile(outputPath)
    .Slide("Summary", slide => slide
        .Value("Title", "Monthly Sales")
        .Table("TopCustomers", customers)
        .Chart("SalesTrend", salesByMonth))
    .RepeatSlide("RegionDetail", regions)
    .Save();
```

These should be separate entry points rather than extensions of `ExcelReport`, but they can share internal concepts:

- `DataSourceSpec`
- `BindingSpec`
- `StyleToken`
- `ThemeSpec`
- `TableSpec`
- `ChartSpec`
- `TemplateSpec`
- `ValidationResult`

## Suggested Public Types

Export:

- `ExcelReport`
- `ExcelWorkbookBuilder`
- `ExcelSheetBuilder`
- `ExcelRepeatingBlockBuilder`
- `ExcelBlockTemplateBuilder`
- `ExcelPivotTableBuilder`
- `ExcelPivotFieldBuilder`
- `ExcelChartBuilder`
- `ExcelDrawingBuilder`
- `ExcelImageBuilder`
- `ExcelValidationBuilder`
- `ExcelProtectionBuilder`
- `ExcelSourceBuilder`
- `ExcelColumnBuilder`
- `ExcelStyleBuilder`
- `ExcelTemplateBuilder`
- `CsvWriter`

Import:

- `ExcelReader`
- `ExcelReadBuilder`
- `ExcelSheetReadBuilder`
- `ExcelMappingBuilder<T>`
- `CsvReader`
- `ImportResult<T>`
- `ImportError`

Shared:

- `IRowSource`
- `ColumnDefinition`
- `RowValueSet`
- `WorkbookSpec`
- `SheetSpec`
- `DataSourceSpec`
- `RepeatingBlockSpec`
- `BlockTemplateSpec`
- `PivotTableSpec`
- `PivotFieldSpec`
- `ChartSpec`
- `DrawingSpec`
- `ImageSpec`
- `ValidationSpec`
- `HyperlinkSpec`
- `CommentSpec`
- `ProtectionSpec`
- `SlicerSpec`
- `StyleSpec`
- `TemplateSpec`

## Minimum Useful First Version

The first implementation should support:

- Export to `.xlsx` file.
- Export to `.csv` file.
- `IEnumerable<T>`, `DataTable`, and `DataSet` sources.
- Sheet name, start cell, headers, auto-fit columns.
- Basic styles with presets.
- Basic number/date formatting.
- Basic totals and formulas.
- Import from `.xlsx` file to `List<T>` and `DataTable`.
- CSV import/export for simple files, including delimiter, headers, encoding, quote, and escape options.

After that, add:

- Streams and byte arrays.
- Conditional formatting.
- Templates.
- Named ranges and Excel tables.
- Repeating data blocks for data-repeater style report layouts.
- Pivot tables with source data on a separate sheet.
- Charts sourced from tables, ranges, and pivots.
- Images, hyperlinks, comments, data validation, and sheet/workbook protection.
- Sparklines, slicers, timelines, and drawing shapes after the core chart path is stable.
- Validation result mode.
- More structured formats such as TSV, fixed-width, JSON, JSON Lines, XML, HTML tables, Markdown tables, and Parquet.
- Word document generation from templates and source data.
- PowerPoint deck generation from templates and source data.
- PDF or HTML report output if there is a reliable rendering path.

## Declarative Layout Schema

The fluent API should not be the only way to define workbook layout. DocWeave should also support a declarative schema that describes the workbook, sheets, tables, ranges, styles, fonts, borders, wrapping, formulas, pivot tables, and conditional formatting.

This schema should compile into the same internal `WorkbookSpec` as the fluent API:

```text
Fluent API  -> WorkbookSpec -> Renderer
Schema      -> WorkbookSpec -> Renderer
```

That keeps the rendering engine shared while giving callers two ways to describe output:

- fluent code for dynamic or simple reports
- schema/layout documents for standardized report packs

### Schema Use Cases

Schema-driven layout is useful when:

- Reports need consistent corporate styling.
- Non-developers or report designers need to adjust layout.
- The same layout is reused with different data.
- A workbook has multiple sections, ranges, tables, and summary areas.
- Layout needs to be versioned independently from code.
- The output resembles a report document more than a simple data dump.

### Schema Formats

Potential schema formats:

- JSON
- YAML
- XML
- C# object model serialized to JSON/YAML

JSON is the best first target because it is easy to validate, serialize, diff, and generate.

Example:

```json
{
  "workbook": {
    "defaultStyle": "corporate",
    "sheets": [
      {
        "name": "Active Directory",
        "tabColor": "004D17",
        "gridlines": false,
        "sections": [
          {
            "type": "title",
            "text": "Active Directory Report",
            "range": "A1:Q1",
            "style": "reportTitle"
          },
          {
            "type": "table",
            "name": "users",
            "source": "activeDirectory",
            "startCell": "A3",
            "style": "greenTable",
            "headers": {
              "show": true,
              "freeze": true,
              "filter": true,
              "wrap": true
            },
            "columns": [
              { "key": "ID", "header": "AD ID", "width": 14 },
              { "key": "EMail", "header": "Mail", "width": 30 },
              { "key": "TelephoneNumber", "header": "Phone", "width": 20, "format": "phone" },
              { "key": "SAMAccountName", "header": "Account", "width": 18 },
              { "key": "Department", "header": "Department", "width": 24 },
              { "key": "Office", "header": "Office", "width": 24 }
            ],
            "totals": {
              "show": false
            }
          }
        ]
      },
      {
        "name": "Summary",
        "sections": [
          {
            "type": "pivotTable",
            "name": "UsersByDepartment",
            "sourceTable": "users",
            "startCell": "A3",
            "rows": [ "Department" ],
            "columns": [ "Office" ],
            "values": [
              {
                "field": "ID",
                "caption": "Users",
                "aggregate": "count"
              }
            ],
            "filters": [ "Company" ],
            "layout": "tabular",
            "style": "PivotStyleMedium9",
            "refreshOnOpen": true
          }
        ]
      }
    ],
    "styles": {
      "reportTitle": {
        "font": {
          "name": "Calibri",
          "size": 14,
          "bold": true,
          "color": "FFFFFF"
        },
        "fill": "004D17",
        "alignment": {
          "horizontal": "center",
          "vertical": "center",
          "wrap": true
        },
        "border": {
          "bottom": {
            "style": "thick",
            "color": "004D17"
          }
        }
      },
      "greenTable": {
        "header": {
          "font": { "bold": true, "color": "FFFFFF" },
          "fill": "004D17",
          "wrap": true,
          "border": {
            "bottom": { "style": "thick", "color": "004D17" }
          }
        },
        "row": {
          "font": { "color": "000000" },
          "fill": "FFFFFF",
          "wrap": false
        },
        "alternateRow": {
          "fill": "CCE1C1"
        }
      }
    }
  }
}
```

### Schema Concepts

The schema should support these top-level concepts:

- Workbook.
- Sheets.
- Sections.
- Tables.
- Ranges.
- Cells.
- Styles.
- Named data sources.
- Layout rules.
- Formulas.
- Conditional formatting.
- Validation rules.

### Workbook Schema

Workbook-level options:

- Default font.
- Default style.
- Theme colors.
- Culture.
- Date/number formats.
- Document properties.
- Template path or template stream reference.

Example:

```json
{
  "workbook": {
    "defaultFont": {
      "name": "Calibri",
      "size": 11
    },
    "culture": "en-GB",
    "theme": "corporateGreen"
  }
}
```

### Sheet Schema

Sheet-level options:

- Sheet name.
- Tab color.
- Gridlines.
- Freeze panes.
- Print settings.
- Page orientation.
- Sections.

Example:

```json
{
  "name": "Summary",
  "tabColor": "004D17",
  "gridlines": false,
  "freeze": { "row": 3 },
  "page": {
    "orientation": "landscape",
    "fitToWidth": 1
  }
}
```

### Section Schema

Sections define visible layout blocks. Section types can include:

- `title`
- `text`
- `table`
- `range`
- `summary`
- `spacer`
- `image` later if needed

Example:

```json
{
  "type": "summary",
  "name": "riskSummary",
  "startCell": "B4",
  "layout": "twoColumn",
  "items": [
    { "label": "Run Date", "value": "{{runDate}}" },
    { "label": "Generated By", "value": "{{userName}}" }
  ],
  "style": "summaryPanel"
}
```

### Table Schema

Tables are the most important schema object.

```json
{
  "type": "table",
  "name": "transactions",
  "source": "transactions",
  "startCell": "A5",
  "headers": {
    "show": true,
    "filter": true,
    "freeze": true,
    "wrap": true
  },
  "columns": [
    {
      "key": "TradeDate",
      "header": "Trade Date",
      "format": "yyyy-MM-dd",
      "width": 14,
      "style": "dateColumn"
    },
    {
      "key": "Amount",
      "header": "Amount",
      "format": "#,##0.00",
      "width": 16,
      "style": "amountColumn"
    }
  ],
  "totals": {
    "show": true,
    "label": "Total",
    "columns": [
      { "key": "Amount", "function": "sum" }
    ]
  }
}
```

Table options:

- Source name.
- Start cell.
- Header behavior.
- Column order.
- Column names.
- Column widths.
- Column formats.
- Column styles.
- Totals.
- Formulas.
- Conditional formatting.
- Empty-source behavior.
- Alternating rows.
- Filters.

### DataTable Report Sheet Schema

This scenario covers a common v1 export: a `DataTable` with a variable number of rows is exported to a new sheet, starts at `B4`, has report details above the table, replaces raw column names with friendlier headers, applies alternating row colors and fonts, adds totals, and gives one important column a thick border while the rest of the table uses normal borders.

```json
{
  "workbook": {
    "sheets": [
      {
        "name": "Invoice Export",
        "sections": [
          {
            "type": "range",
            "name": "exportDetails",
            "cells": [
              {
                "cell": "B1",
                "value": "Invoice Export",
                "style": "reportTitle"
              },
              {
                "cell": "B2",
                "value": "Generated:",
                "style": "detailLabel"
              },
              {
                "cell": "C2",
                "value": "{{generatedDate}}",
                "format": "yyyy-MM-dd HH:mm",
                "style": "detailValue"
              },
              {
                "cell": "B3",
                "value": "Requested By:",
                "style": "detailLabel"
              },
              {
                "cell": "C3",
                "value": "{{requestedBy}}",
                "style": "detailValue"
              }
            ]
          },
          {
            "type": "table",
            "name": "invoiceTable",
            "source": "invoiceRows",
            "startCell": "B4",
            "style": "standardTable",
            "headers": {
              "show": true,
              "filter": true,
              "freeze": true,
              "wrap": true,
              "style": "tableHeader"
            },
            "columns": [
              {
                "key": "InvoiceNo",
                "header": "Invoice No.",
                "width": 16
              },
              {
                "key": "CustomerName",
                "header": "Customer",
                "width": 32
              },
              {
                "key": "InvoiceDate",
                "header": "Date",
                "format": "yyyy-MM-dd",
                "width": 14
              },
              {
                "key": "Status",
                "header": "Status",
                "width": 16,
                "style": "highlightColumn"
              },
              {
                "key": "Amount",
                "header": "Amount",
                "format": "#,##0.00",
                "width": 16,
                "style": "amountColumn"
              }
            ],
            "rows": {
              "style": "tableRow",
              "alternateStyle": "tableAlternateRow"
            },
            "totals": {
              "show": true,
              "label": "Total",
              "style": "totalRow",
              "columns": [
                { "key": "Amount", "function": "sum" }
              ]
            },
            "emptySource": {
              "mode": "message",
              "message": "No invoice records were available for export."
            }
          }
        ]
      }
    ],
    "styles": {
      "reportTitle": {
        "font": { "name": "Calibri", "size": 14, "bold": true },
        "alignment": { "horizontal": "left" }
      },
      "detailLabel": {
        "font": { "bold": true },
        "alignment": { "horizontal": "left" }
      },
      "detailValue": {
        "font": { "name": "Calibri", "size": 11 }
      },
      "standardTable": {
        "border": {
          "all": { "style": "thin", "color": "B7B7B7" }
        }
      },
      "tableHeader": {
        "font": { "bold": true, "color": "FFFFFF" },
        "fill": "1F4E78",
        "alignment": { "horizontal": "center", "wrap": true },
        "border": {
          "all": { "style": "thin", "color": "808080" }
        }
      },
      "tableRow": {
        "font": { "name": "Calibri", "size": 11, "color": "000000" },
        "fill": "FFFFFF",
        "border": {
          "all": { "style": "thin", "color": "D9E2F3" }
        }
      },
      "tableAlternateRow": {
        "font": { "name": "Calibri", "size": 11, "color": "000000" },
        "fill": "EAF2F8",
        "border": {
          "all": { "style": "thin", "color": "D9E2F3" }
        }
      },
      "highlightColumn": {
        "border": {
          "left": { "style": "thick", "color": "000000" },
          "right": { "style": "thick", "color": "000000" },
          "top": { "style": "thin", "color": "D9E2F3" },
          "bottom": { "style": "thin", "color": "D9E2F3" }
        }
      },
      "amountColumn": {
        "numberFormat": "#,##0.00",
        "alignment": { "horizontal": "right" }
      },
      "totalRow": {
        "font": { "bold": true },
        "fill": "D9EAF7",
        "border": {
          "top": { "style": "double", "color": "1F4E78" }
        }
      }
    }
  }
}
```

Fluent equivalent:

```csharp
ExcelReport.Create()
    .ToFile(path)
    .WithValue("generatedDate", DateTime.Now)
    .WithValue("requestedBy", currentUser)
    .Sheet("Invoice Export", sheet => sheet
        .Sources(sources =>
        {
            sources.AddKeyValueBlock("exportDetails", new
            {
                Generated = DateTime.Now,
                RequestedBy = currentUser
            })
            .At("B1")
            .Style("detailBlock");

            sources.Add("invoiceTable", invoiceDataTable)
            .At("B4")
            .AsTable("InvoiceTable")
            .Headers(h => h.Show().Filter().Freeze().Wrap().Style("tableHeader"))
            .Columns(c => c
                .Rename("InvoiceNo", "Invoice No.")
                .Rename("CustomerName", "Customer")
                .Rename("InvoiceDate", "Date")
                .Rename("Amount", "Amount")
                .Format("InvoiceDate", "yyyy-MM-dd")
                .Format("Amount", "#,##0.00")
                .Style("Status", "highlightColumn")
                .Width("CustomerName", 32))
            .Style("standardTable")
            .Totals(t => t
                .Label("Total")
                .Sum("Amount")
                .Style("totalRow"));
        }))
    .Save();
```

### Repeating Block Schema

Repeating blocks describe a rectangular layout that is rendered once per item in a source collection. They are useful for dashboards, grouped summaries, labels, cards, account packs, project snapshots, and region-by-region financial blocks.

```json
{
  "type": "repeatingBlock",
  "name": "regionSummaryBlocks",
  "source": "regions",
  "startCell": "A2",
  "perRow": 3,
  "blockSize": {
    "rows": 4,
    "columns": 3
  },
  "spacing": {
    "rows": 1,
    "columns": 2
  },
  "keepTogether": true,
  "cells": [
    {
      "cell": "A1:C1",
      "value": "{{Region}}",
      "merge": true,
      "style": "regionTitle"
    },
    {
      "cell": "A2",
      "value": "Costs",
      "style": "metricLabel"
    },
    {
      "cell": "B2",
      "value": "{{Costs}}",
      "format": "#,##0.00",
      "style": "metricValue"
    },
    {
      "cell": "A3",
      "value": "Income",
      "style": "metricLabel"
    },
    {
      "cell": "B3",
      "value": "{{Income}}",
      "format": "#,##0.00",
      "style": "metricValue"
    },
    {
      "cell": "A4",
      "value": "Profit",
      "style": "metricLabel"
    },
    {
      "cell": "B4",
      "formula": "{{Income}}-{{Costs}}",
      "format": "#,##0.00",
      "style": "metricValue"
    }
  ]
}
```

Repeating block schema options:

- Source collection.
- Start cell.
- Number of blocks per row.
- Block size.
- Row and column spacing.
- Template range instead of inline cells.
- Cells, ranges, and merges inside each block.
- Bound values and formulas.
- Nested mini-tables inside a block.
- Conditional formatting inside each block.
- Page breaks and print options.
- Keep-together behavior.

### Pivot Table Schema

Pivot tables can be declared as sheet sections. The source should usually be a named table section from another sheet, but explicit ranges and named ranges should also be supported.

```json
{
  "type": "pivotTable",
  "name": "SalesByRegion",
  "sourceTable": "salesData",
  "startCell": "A3",
  "rows": [ "Region", "SalesPerson" ],
  "columns": [ "Month" ],
  "values": [
    {
      "field": "Amount",
      "caption": "Total Sales",
      "aggregate": "sum",
      "format": "#,##0.00"
    },
    {
      "field": "OrderId",
      "caption": "Orders",
      "aggregate": "count"
    }
  ],
  "filters": [ "Year" ],
  "layout": "tabular",
  "style": "PivotStyleMedium9",
  "grandTotals": {
    "rows": true,
    "columns": true
  },
  "refreshOnOpen": true
}
```

Pivot table options:

- Source table, named range, or explicit source range.
- Start cell.
- Row fields.
- Column fields.
- Value fields.
- Filter/page fields.
- Aggregation per value field.
- Caption and number format per value field.
- Report layout.
- Grand totals and subtotals.
- Pivot style.
- Refresh on open.
- Source sheet visibility.
- Cache behavior for file size versus compatibility.

### Chart And Visual Schema

Charts and visual elements can be declared as sheet sections.

```json
{
  "type": "chart",
  "name": "SalesTrend",
  "chartType": "line",
  "sourceTable": "salesData",
  "category": "Month",
  "series": [
    {
      "field": "Amount",
      "caption": "Sales",
      "format": "#,##0.00"
    }
  ],
  "position": {
    "cell": "B3",
    "width": 720,
    "height": 320
  },
  "title": "Sales Trend",
  "legend": "bottom",
  "style": "salesChart"
}
```

Image schema:

```json
{
  "type": "image",
  "name": "CompanyLogo",
  "source": "companyLogo",
  "position": {
    "cell": "A1",
    "width": 160,
    "height": 48
  },
  "altText": "Company logo"
}
```

Diagram/schema shape example:

```json
{
  "type": "diagram",
  "name": "ApprovalFlow",
  "startCell": "A8",
  "shapes": [
    { "name": "Request", "shape": "rectangle", "text": "Request", "width": 120, "height": 50 },
    { "name": "Review", "shape": "rectangle", "text": "Review", "rightOf": "Request", "spacing": 40 }
  ],
  "connectors": [
    { "from": "Request", "to": "Review", "type": "arrow" }
  ]
}
```

Visual schema options:

- Chart source table, range, named range, or pivot.
- Chart type.
- Categories and series.
- Titles, axes, legends, data labels, and number formats.
- Position and size.
- Images from named binary sources.
- Basic shapes, text boxes, arrows, connectors, and callouts.
- Sparklines.
- Slicers and timelines later.

### Range and Cell Schema

Useful for summaries, report parameters, and template-like layouts.

```json
{
  "type": "range",
  "name": "reportInfo",
  "startCell": "A1",
  "cells": [
    {
      "cell": "A1",
      "value": "Report Date",
      "style": "label"
    },
    {
      "cell": "B1",
      "value": "{{runDate}}",
      "format": "yyyy-MM-dd",
      "style": "value"
    }
  ]
}
```

Cell options:

- Static value.
- Bound value.
- Formula.
- Format.
- Style.
- Merge.
- Comment/note later if useful.

### Protection And Encryption Schema

Schemas can declare that a workbook or sheet should be protected or encrypted, but schemas should reference password keys rather than store raw passwords.

```json
{
  "workbook": {
    "encryption": {
      "openPasswordRef": "sensitiveExportPassword",
      "algorithm": "aes256"
    },
    "protection": {
      "protectStructure": true,
      "passwordRef": "workbookStructurePassword"
    },
    "sheets": [
      {
        "name": "Sensitive Data",
        "protection": {
          "protectSheet": true,
          "passwordRef": "sensitiveSheetPassword",
          "allowFiltering": true,
          "allowSorting": true
        },
        "sections": []
      }
    ]
  }
}
```

Runtime binding:

```csharp
ExcelReport.FromSchema(schema)
    .WithSecret("sensitiveExportPassword", openPassword)
    .WithSecret("workbookStructurePassword", workbookPassword)
    .WithSecret("sensitiveSheetPassword", sheetPassword)
    .WithSource("sensitiveRows", rows)
    .ToFile(path)
    .Save();
```

Schema guidance:

- Store password references, not password values.
- Resolve passwords from the caller, a secret store, or a secure WebAPI request context.
- Use workbook encryption for sensitive data that must require a password to open.
- Use sheet/workbook protection for editing controls.
- Do not treat sheet protection as data confidentiality.

### Style Schema

Styles should be reusable and composable.

```json
{
  "styles": {
    "currencyColumn": {
      "format": "£#,##0.00",
      "alignment": {
        "horizontal": "right"
      }
    },
    "warningCell": {
      "font": {
        "bold": true,
        "color": "9C0006"
      },
      "fill": "FFC7CE"
    }
  }
}
```

Style options:

- Font name.
- Font size.
- Bold.
- Italic.
- Underline.
- Font color.
- Fill/background.
- Number format.
- Date format.
- Horizontal alignment.
- Vertical alignment.
- Wrap text.
- Indent.
- Borders.
- Locked/hidden protection flags later if needed.

### Border Schema

```json
{
  "border": {
    "top": { "style": "thin", "color": "D9EAD3" },
    "right": { "style": "thin", "color": "D9EAD3" },
    "bottom": { "style": "thick", "color": "004D17" },
    "left": { "style": "thin", "color": "D9EAD3" }
  }
}
```

Border styles:

- none
- hair
- dotted
- dashed
- thin
- medium
- thick
- double

### Formula Schema

Formulas can be attached to columns, totals, or specific cells.

```json
{
  "formulas": [
    {
      "target": "Variance",
      "type": "column",
      "formula": "={column:Actual}{row}-{column:Budget}{row}"
    },
    {
      "target": "Amount",
      "type": "total",
      "formula": "=SUM({column}{start}:{column}{end})"
    }
  ]
}
```

Formula placeholders:

- `{row}`
- `{start}`
- `{end}`
- `{column}`
- `{column:Name}`
- `{cell:Name}`

### Conditional Formatting Schema

```json
{
  "conditionalFormatting": [
    {
      "target": "Amount",
      "rule": "greaterThan",
      "value": 1000000,
      "style": "warningCell"
    },
    {
      "range": "E2:E500",
      "rule": "dataBar",
      "color": "63C384"
    }
  ]
}
```

Supported rule types:

- greaterThan
- lessThan
- between
- equals
- notEquals
- textContains
- top
- bottom
- duplicate
- unique
- colorScale
- dataBar
- iconSet

### Data Binding

The schema should refer to named sources supplied by code:

```csharp
ExcelReport.FromSchema(schema)
    .WithSource("activeDirectory", activeDirectoryRows)
    .WithSource("transactions", transactions)
    .WithValue("runDate", DateTime.Today)
    .ToFile(outputPath)
    .Save();
```

DataSet tables should also be bindable either as a whole `DataSet` or as individual named sources:

```csharp
ExcelReport.FromSchema(schema)
    .WithDataSet("financialData", dataSet)
    .WithSource("forecast", forecastRows)
    .WithSource("runSummary", runSummary)
    .ToFile(outputPath)
    .Save();
```

Mixed schema example:

```json
{
  "workbook": {
    "sheets": [
      {
        "name": "Financials",
        "sections": [
          {
            "type": "table",
            "name": "costs",
            "source": "financialData.Costs",
            "startCell": "A2",
            "style": "costTable",
            "columns": [
              { "key": "Region", "width": 20 },
              { "key": "Amount", "format": "#,##0.00" }
            ],
            "totals": {
              "show": true,
              "columns": [
                { "key": "Amount", "function": "sum" }
              ]
            }
          },
          {
            "type": "table",
            "name": "income",
            "source": "financialData.Income",
            "startCell": "H2",
            "style": "incomeTable"
          },
          {
            "type": "table",
            "name": "forecast",
            "source": "forecast",
            "startCell": "A20",
            "style": "forecastTable"
          },
          {
            "type": "keyValueBlock",
            "name": "runSummary",
            "source": "runSummary",
            "startCell": "H20",
            "style": "summaryBlock"
          }
        ]
      },
      {
        "name": "Reference Data",
        "sections": [
          {
            "type": "table",
            "name": "regions",
            "source": "financialData.Regions",
            "startCell": "A1",
            "style": "referenceTable"
          }
        ]
      }
    ]
  }
}
```

Binding syntax in schema:

```text
{{runDate}}
{{user.name}}
{{parameters.asOfDate}}
```

### Fluent Equivalent

Every schema object should have a fluent equivalent. For example this schema:

```json
{
  "type": "table",
  "source": "transactions",
  "startCell": "A3",
  "style": "greenTable"
}
```

Should map to:

```csharp
.Sheet("Transactions", sheet => sheet
    .From(transactions)
    .StartAt("A3")
    .Style("greenTable"))
```

### Schema Validation

Schema loading should validate:

- Required fields.
- Invalid cell references.
- Duplicate sheet names.
- Duplicate style names.
- Missing data sources.
- Missing style references.
- Unsupported format strings.
- Invalid formulas/placeholders.
- Invalid range references.

Validation result:

```csharp
SchemaValidationResult result = LayoutSchema.Load(json).Validate();
```

### Suggested API

```csharp
var schema = LayoutSchema.FromJsonFile("risk-report.layout.json");

ExcelReport.FromSchema(schema)
    .WithSource("activeDirectory", activeDirectoryRows)
    .WithSource("transactions", transactions)
    .WithValue("runDate", DateTime.Today)
    .ToFile(outputPath)
    .Save();
```

Or inline:

```csharp
ExcelReport.Create()
    .UseLayout(layout => layout
        .Sheet("Summary", sheet => sheet
            .Title("Risk Report")
            .Table("transactions", table => table
                .At("A3")
                .Style("greenTable"))))
    .WithSource("transactions", transactions)
    .ToFile(outputPath)
    .Save();
```

### Implementation Notes

The schema should not directly drive OpenXML rendering. It should be parsed into the same specs as fluent calls:

```text
JSON/YAML/XML schema
  -> LayoutSchema
  -> WorkbookSpec
  -> WorkbookRenderer
```

This keeps one rendering engine and makes behavior consistent between fluent and declarative layouts.
