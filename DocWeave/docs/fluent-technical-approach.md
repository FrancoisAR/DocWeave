# DocWeave Fluent API Technical Approach

> **Status: design document.** It describes the intended approach, and parts of it are not implemented. The current API is in [getting-started.md](getting-started.md) and [pre-release-functionality-guide.md](pre-release-functionality-guide.md).

## Purpose

DocWeave should provide a fluent, strongly typed API for importing and exporting structured data, with Excel `.xlsx` as the primary target and CSV or other tabular formats as first-class companions.

The initial implementation should support `.xlsx` and `.csv` import/export. Legacy `.xls` is out of scope because it is an older binary format and tends to require different libraries or Excel Interop/COM automation, which is not suitable for server-side use.

Macro-enabled workbooks should have a narrow initial scope:

- Do not create macros.
- Do not edit VBA.
- Do not inspect VBA.
- Do not sign macros.
- Preserve existing macros when an `.xlsm` or `.xltm` template/workbook is supplied and the target output is `.xlsm`.
- Validate extension combinations so a macro-enabled template is not accidentally saved as `.xlsx` while silently dropping macro parts.

The architecture should also leave room for later document and presentation generation from the same source data. Word and PowerPoint support do not need to be part of the first implementation, but the core model should avoid assumptions that only worksheets can be rendered.

The initial feature target should cover the useful report and import scenarios already identified:

- Export workbooks with one or more sheets.
- Export `IEnumerable<T>`, `DataTable`, `DataSet`, `DataRow`, dictionaries, and mixed data sources.
- Export `DataSet` instances as one sheet per table, selected tables, multiple tables on a sheet, or related/grouped layouts.
- Compose a sheet from many independently configured export regions, mixing `DataTable`, `DataSet` tables, object collections, dictionaries, summary ranges, repeated blocks, pivots, and charts.
- Export object collections as repeating data blocks where each object renders into a rectangular layout region.
- Configure sheet names, starting rows and columns, headers, filters, totals, formulas, conditional formatting, alternating rows, tab colors, column widths, and style colors.
- Use Excel templates, named ranges, copied sheets, region mapping, cards, and grouped data.
- Import from files, streams, and byte arrays.
- Import into `List<T>`, `DataTable`, and dictionary-like structures.
- Map columns by headers or explicit column letters.
- Handle CSV-style import alongside XLSX import.
- Export simple tabular data to CSV and other delimited formats.
- Leave a future path for Word document and PowerPoint deck renderers that use the same data binding, templates, styles, and validation concepts.

The redesign should preserve that expressive caller experience while separating the public fluent API from OpenXML mechanics.

## Design Goals

- Keep simple exports simple.
- Make advanced reports possible without forcing every caller into low-level configuration.
- Avoid global mutable state.
- Support file, stream, and byte-array workflows consistently.
- Make the fluent API compile-time discoverable where possible.
- Represent export and import intent as testable model objects before rendering.
- Keep OpenXML details behind a renderer layer.
- Prefer a clean, greenfield API over preserving previous call shapes.

## Proposed Architecture

### Public API Layer

This layer contains fluent builders only. Builders collect configuration and produce immutable specs. They should not write OpenXML directly.

Primary entry points:

```csharp
ExcelReport.Create()
ExcelReader.FromFile(path)
CsvReader.FromFile(path)
CsvWriter.Create()
StructuredData.Read(path)
```

Export example:

```csharp
ExcelReport.Create()
    .ToFile(outputPath)
    .Sheet("Active Directory", sheet => sheet
        .From(activeDirectoryRows)
        .StartAt("A1")
        .Headers(headers => headers.Show().Freeze())
        .Columns(columns => columns.AutoFit())
        .Style(ReportStyles.Green))
    .Save();
```

Import example:

```csharp
var cities = ExcelReader.FromFile(inputPath)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .MapTo<City>()
    .Read();
```

### Specification Layer

Fluent calls should build plain model objects:

- `WorkbookSpec`
- `SheetSpec`
- `DataSourceSpec`
- `ExportRegionSpec`
- `DataGroupSpec`
- `RegionPlacementSpec`
- `RepeatingBlockSpec`
- `BlockTemplateSpec`
- `BlockPlacementSpec`
- `ColumnSpec`
- `HeaderSpec`
- `PivotTableSpec`
- `PivotFieldSpec`
- `ChartSpec`
- `DrawingSpec`
- `ImageSpec`
- `ValidationSpec`
- `HyperlinkSpec`
- `CommentSpec`
- `ProtectionSpec`
- `StyleSpec`
- `TemplateSpec`
- `ImportSpec`
- `CsvSpec`

These specs should be immutable after build, or at least treated as read-only by renderers.

This enables unit tests like:

```csharp
Assert.Equal("Active Directory", workbook.Sheets[0].Name);
Assert.Equal("A1", workbook.Sheets[0].StartCell);
Assert.True(workbook.Sheets[0].Columns.AutoFit);
```

### Data Adapter Layer

Every supported source should be normalized into the same internal shape.

Suggested interfaces:

```csharp
public interface IRowSource
{
    IReadOnlyList<ColumnDefinition> Columns { get; }
    IEnumerable<RowValueSet> Rows { get; }
}

public sealed record ColumnDefinition(string Key, string Header, Type DataType);
public sealed record RowValueSet(IReadOnlyDictionary<string, object?> Values);
```

Adapters:

- `ObjectEnumerableAdapter<T>`
- `DataTableAdapter`
- `DataSetAdapter`
- `DataRowEnumerableAdapter`
- `DictionaryAdapter`
- `DynamicObjectAdapter`
- `CsvRecordAdapter`

The renderer should not care whether the original source was `DataTable` or `List<T>`.

### Rendering Layer

Format-specific logic should live in renderer services. Excel rendering will use OpenXML first, while CSV and later structured formats should use their own focused renderers.

- `WorkbookRenderer`
- `WorksheetRenderer`
- `StyleRenderer`
- `TableRenderer`
- `RepeatingBlockRenderer`
- `PivotTableRenderer`
- `PivotCacheRenderer`
- `ChartRenderer`
- `DrawingRenderer`
- `ImageRenderer`
- `ValidationRenderer`
- `ProtectionRenderer`
- `TemplateRenderer`
- `SharedStringWriter`
- `ColumnWidthCalculator`
- `CsvRenderer`
- `DelimitedTextRenderer`
- `JsonRenderer`
- `WordDocumentRenderer`
- `PowerPointDeckRenderer`

Excel rendering pipeline:

1. Validate `WorkbookSpec`.
2. Create or copy workbook/template.
3. Build style registry.
4. Render sheets.
5. Resolve export region layout for each sheet.
6. Render tables/data groups.
7. Render repeating blocks and template ranges.
8. Render pivot caches and pivot table definitions after source sheets/tables exist.
9. Render charts, images, diagrams, shapes, and drawing anchors after source ranges exist.
10. Apply formulas, filters, freezes, conditional formats, data validation, hyperlinks, comments, widths, merges, and protection.
11. Save to file, stream, or byte array.

An export region is the unit of placement and formatting. It should wrap a source plus all source-specific options: target sheet, start cell, layout behavior, table name, columns, headers, styles, formulas, totals, conditional formatting, empty-source behavior, named range/table creation, and print behavior. Sheets can contain many regions, and each region can be independently configured.

Region layout should support:

- Absolute placement by cell.
- Relative placement after another region.
- Vertical stacking.
- Horizontal stacking.
- Grid placement.
- Template/range placement.
- Collision validation so two regions do not unintentionally overlap.
- Optional overlap mode for advanced template scenarios.

Chart and drawing rendering should be separate from table rendering. Charts may depend on rendered tables, named ranges, explicit ranges, or pivot tables, so they should be resolved after data regions and pivots have been written. Diagram, shape, image, and connector support should use OpenXML drawing anchors and expose a practical subset of Excel drawing features.

Visual rendering scope:

- Charts from tables, ranges, named ranges, and pivots.
- Chart titles, axes, legends, series, data labels, and number formats.
- Images from files, streams, or byte arrays.
- Basic shapes, text boxes, arrows, callouts, and connectors.
- Sparklines after table/range rendering is stable.
- Slicers and timelines in a later phase because they are more compatibility-sensitive.

Repeating block rendering needs a layout engine separate from table rendering. It should calculate the target top-left cell for each source item using start cell, block size, spacing, and the configured block count per row. Each block can then render bound values, formulas, ranges, styles, borders, merges, nested mini-tables, and conditional formatting relative to that block origin.

Repeating block placement formula:

```text
blockRow = itemIndex / perRow
blockColumn = itemIndex % perRow

targetRow = startRow + blockRow * (blockHeight + rowSpacing)
targetColumn = startColumn + blockColumn * (blockWidth + columnSpacing)
```

This feature should support two authoring modes:

- Inline block definitions, where fluent calls or schema cells define labels, values, formulas, and styles.
- Template range definitions, where a source range is copied for each item and placeholders are bound inside the copied range.

Pivot table rendering needs a dedicated path because OpenXML pivots are not just formatted ranges. The renderer must create or update pivot cache definitions, pivot cache records when needed, worksheet pivot table definitions, source relationships, and field mappings. The public API should hide those mechanics behind `PivotTableSpec`.

Pivot table source strategy:

- Prefer source data exported as an Excel table, because table-backed pivots are resilient when row counts change.
- Support explicit source ranges for template and advanced scenarios.
- Support source data on a visible sheet, hidden sheet, or very hidden sheet.
- Allow one source table to feed multiple pivot tables.
- Allow pivot output on a different worksheet from the source data.
- Support refresh-on-open so Excel recalculates the pivot when the workbook opens.

Plain data export pipeline:

1. Validate source and column definitions.
2. Normalize rows through `IRowSource`.
3. Apply column selection, ordering, captions, conversion, and null handling.
4. Write the selected format, such as CSV, TSV, fixed-width, JSON, JSON Lines, XML, HTML table, or Markdown table.
5. Save to file, stream, byte array, or text writer.

Document and presentation renderers should be later additions over a more general `ReportSpec` or `DocumentSpec`. They should support templates, placeholder binding, repeated sections, tables, charts, images, and theme/style tokens without forcing those concepts into the Excel-specific builder.

### Import Layer

Import should mirror export but in reverse:

1. Open structured source.
2. Select sheet, table, or range.
3. Read raw cells or fields.
4. Build normalized rows.
5. Apply header and column mapping.
6. Convert values.
7. Return typed objects, `DataTable`, dictionaries, or validation results.

Suggested API:

```csharp
var rows = ExcelReader.FromStream(stream)
    .Sheet("Rates")
    .Range(range => range
        .FromRow(5)
        .ToRow(110)
        .FromColumn("B"))
    .Columns(columns => columns
        .Map("Code").From("B")
        .Map("Buy").From("C")
        .Map("Sell").From("D"))
    .Read<DataTable>();
```

## Export Fluent Surface

### Workbook

```csharp
ExcelReport.Create()
    .ToFile(path)
    .UseTemplate(templatePath)
    .Properties(p => p
        .Author("...")
        .Title("..."))
    .Sheet(...)
    .Save();
```

Output targets:

- `ToFile(string path)`
- `ToStream(Stream stream)`
- `ToByteArray()`
- `Save()`
- `Build()` returning a `WorkbookSpec`

### Sheet

```csharp
.Sheet("Summary", sheet => sheet
    .StartAt(row: 2, column: 1)
    .TabColor("004D17")
    .Freeze(row: 1)
    .Filter()
    .From(source))
```

Features:

- Sheet name.
- Tab color.
- Sheet background or fill where supported.
- Start row and column.
- Freeze panes.
- Auto filter.
- Page setup.
- Print area.
- Landscape or portrait.
- Hide gridlines.

### Data Sources

Simple source:

```csharp
.From(rows)
```

Named source:

```csharp
.Source("Audit Trail", rows)
```

Multiple sources:

```csharp
.Sources(sources => sources
    .Add("Table 1", table1)
    .Add("Table 2", table2)
    .Layout(layout => layout.StackVertically(spacingRows: 2)))
```

Mixed independently formatted regions:

```csharp
.Sheet("Summary", sheet => sheet
    .Sources(sources => sources
        .Add("Costs", dataSet.Tables["Costs"])
            .At("A2")
            .AsTable("CostsTable")
            .Style("costTable")
            .Columns(c => c.Format<decimal>("#,##0.00"))
            .Totals(t => t.Sum("Amount"))
        .Add("Forecast", forecastRows)
            .At("H2")
            .Style("forecastTable")
            .Headers(h => h.Style("blueHeader"))
            .Columns(c => c
                .Include(x => x.Month)
                .Include(x => x.ExpectedIncome)
                .Include(x => x.ExpectedCost))
        .Add("RunInfo", runSummary)
            .At("A20")
            .AsKeyValueBlock()
            .Style("summaryBlock")))
```

Groups:

```csharp
.Group("Risk", group => group
    .Style(ReportStyles.Green)
    .Source("FX", fxRows)
    .Source("Rates", rateRows)
    .Totals("Amount", "Exposure"))
```

Repeating blocks:

```csharp
.RepeatBlocks(regions, blocks => blocks
    .StartAt("A2")
    .PerRow(3)
    .BlockSize(rows: 4, columns: 3)
    .Spacing(rows: 1, columns: 2)
    .ForEach((block, region) => block
        .Cell("A1:C1").Value(region.Region).Merge().Style("regionTitle")
        .Cell("A2").Value("Costs")
        .Cell("B2").Value(region.Costs).Format("#,##0.00")
        .Cell("A3").Value("Income")
        .Cell("B3").Value(region.Income).Format("#,##0.00")
        .Cell("A4").Value("Profit")
        .Cell("B4").Formula("=B3-B2").Format("#,##0.00")))
```

Supported sources:

- `IEnumerable<T>`
- `DataTable`
- `DataSet`
- `IEnumerable<DataRow>`
- `IEnumerable<IDictionary<string, object?>>`
- custom `IRowSource`

Dataset export modes:

- One sheet per `DataTable`.
- Selected `DataTable` instances.
- Multiple tables stacked or placed on one sheet.
- Mixed with non-DataSet sources on the same sheet.
- Related tables rendered as grouped sections.
- Tables exported as named Excel tables for formulas, validation, charts, and pivots.

Export region features:

- Independent placement.
- Independent style.
- Independent columns.
- Independent headers.
- Independent formulas and totals.
- Independent conditional formatting.
- Independent empty-source behavior.
- Optional group-level defaults with per-region overrides.

Repeating block features:

- One rectangular block per source item.
- Configurable blocks per row.
- Configurable row and column spacing.
- Inline cell/range definitions.
- Template-range copying.
- Property bindings.
- Formulas relative to the current block.
- Nested mini-tables for child collections.
- Optional page breaks and print-friendly keep-together behavior.

### Columns

```csharp
.Columns(columns => columns
    .AutoFit()
    .UseDisplayNameAttributes()
    .Include(x => x.Name)
    .Ignore(x => x.InternalId)
    .Rename("SAMAccountName", "Account")
    .Format<DateTime>("yyyy-MM-dd")
    .Format<decimal>("#,##0.00"))
```

Column features:

- Include and exclude.
- Rename and custom headers.
- DisplayName attribute support.
- Explicit ordering.
- Widths.
- Auto-fit.
- Type-based formats.
- Per-column styles.
- Formulas.
- Totals.

### Headers

```csharp
.Headers(headers => headers
    .Show()
    .Style(ReportStyles.Header)
    .Freeze()
    .Filter())
```

Features:

- Show or hide headers.
- Header row index.
- Header styles.
- Auto-filter binding.
- Freeze header row.

### Styles

Avoid long positional style methods for styling. Prefer named style objects:

```csharp
.Style(style => style
    .Header(h => h
        .Foreground("FFFFFF")
        .Background("004D17")
        .BorderBottom(BorderWeight.Thick))
    .Rows(r => r
        .Alternating("FFFFFF", "CCE1C1"))
    .Totals(t => t
        .Bold()
        .Background("004D17")))
```

Also support reusable presets:

```csharp
.Style(ReportStyles.Green)
.Style(ReportStyles.Plain)
.Style(ReportStyles.Audit)
```

### Conditional Formatting

```csharp
.ConditionalFormatting(cf => cf
    .Range("E2:E100")
    .GreaterThan(0)
    .Fill("CCE1C1"))
```

Options:

- Greater than, less than, and between.
- Text contains.
- Top or bottom N.
- Data bars.
- Color scales.
- Icon sets where practical.

### Formulas

```csharp
.Formulas(formulas => formulas
    .Column("Total", "=SUM([Amount],[Tax])")
    .Total("Amount", Formula.Sum)
    .Custom("Variance", row => $"C{row}-D{row}"))
```

Formula support should understand row placeholders:

- `{row}`
- `{start}`
- `{end}`
- `{column:Name}`

### Templates

Template API:

```csharp
ExcelReport.Create()
    .UseTemplate(t => t
        .FromFile(templatePath)
        .CopySheet("Risk Template").As("Risk Report")
        .MapCell("ReportDate").To(DateTime.Today)
        .MapRange("FX_Table").To(fxRows)
        .MapRegion("Header").To(headerValues))
    .SaveAs(outputPath);
```

Template features:

- Copy whole sheet.
- Copy named range.
- Populate specific cells or named ranges.
- Insert rows into named table or range.
- Preserve source formatting.
- Card or repeated-region support.
- Map multiple sources to multiple named ranges.

## Import Fluent Surface

### Excel Import

```csharp
var rows = ExcelReader.FromFile(path)
    .Sheet("Cities")
    .UseHeaderRow(1)
    .FirstDataRow(2)
    .SkipHiddenRows()
    .SkipHiddenColumns()
    .Read<City>();
```

Options:

- Source: file, stream, byte array.
- Sheet by name or index.
- Used range or explicit range.
- Header row.
- First and last row.
- First and last column.
- Include or exclude columns.
- Map by header name.
- Map by column letter.
- Read hidden rows and columns or skip them.
- Shared string handling.
- Date and number format handling.
- Validation/error collection.

### Typed Mapping

```csharp
.MapTo<City>(map => map
    .Property(x => x.Name).FromHeader("City Name")
    .Property(x => x.Country).FromColumn("B")
    .Property(x => x.Population).ConvertWith(value => int.Parse(value)))
```

Return modes:

- `Read<T>()`
- `ReadDataTable()`
- `ReadDictionaries()`
- `ReadRows()`
- `Validate<T>()`

### CSV Import

```csharp
var rows = CsvReader.FromFile(path)
    .DelimitedBy(",")
    .HasHeaderRow()
    .MapTo<City>()
    .Read();
```

CSV options:

- Delimiter.
- Quote character.
- Escape handling.
- Encoding.
- Header row or no header row.
- Culture-specific parsing.
- Streaming large files.
- Bad row handling.

## Other Structured Formats

The same normalized data model can support:

- TSV.
- Pipe-delimited files.
- JSON arrays.
- JSON Lines.
- XML flat records.
- Parquet later, if a dependency is acceptable.

Possible generic entry point:

```csharp
StructuredData.FromFile(path)
    .AutoDetect()
    .Read<DataTable>();
```

## Validation and Error Handling

Avoid silent conversion failures. Import should support strict and permissive modes:

```csharp
.OnError(ErrorMode.Throw)
.OnError(ErrorMode.Collect)
.OnError(ErrorMode.UseDefault)
```

Validation result:

```csharp
public sealed record ImportResult<T>(
    IReadOnlyList<T> Rows,
    IReadOnlyList<ImportError> Errors);
```

Errors should include:

- Sheet name.
- Row number.
- Column.
- Header.
- Raw value.
- Target type.
- Exception or message.

## Performance Considerations

- Support streaming reads for large XLSX sheets using `OpenXmlReader`.
- Support streaming CSV reads with `IAsyncEnumerable<T>` later.
- Avoid materializing `IEnumerable<T>` multiple times.
- Cache reflection metadata per type.
- Cache styles by value or hash.
- Use shared strings correctly and avoid duplicates.
- Avoid global singleton state for render contexts.

## Greenfield API Position

DocWeave should be treated as a new program rather than a compatibility rewrite. Existing report scenarios can guide feature coverage, but public naming, builder structure, object models, and error behavior should be designed for clarity.

The preferred shape is:

```csharp
ExcelReport.Create()
    .ToFile(path)
    .Sheet("Cities", s => s.From(cities))
    .Save();
```

Avoid adding legacy aliases unless there is a specific migration need later. This keeps the first implementation smaller and prevents internal design choices from being constrained by older abstractions.

## Suggested Implementation Phases

### Phase 1: Core Export Model

- Create specs: workbook, sheet, data source, columns, style.
- Create adapters for `IEnumerable<T>`, `DataTable`, `DataSet`.
- Render simple non-template XLSX files.
- Support `ToFile`, `ToStream`, `ToByteArray`.

### Phase 2: Fluent Export API

- Add workbook/sheet/source builders.
- Add headers, start cell, auto-fit, column ordering, simple styles.
- Add basic formulas and totals.

### Phase 3: Import API

- Read XLSX to normalized rows.
- Map to `T`, `DataTable`, dictionaries.
- Add column/header mapping and conversion options.
- Add validation result mode.

### Phase 4: CSV and Structured Files

- Add CSV reader/writer.
- Add delimiter, encoding, and culture options.
- Add JSON/JSONL readers if useful.

### Phase 5: Templates

- Add template specs.
- Copy sheets and ranges.
- Populate cells and named ranges.
- Render repeated ranges/cards.

### Phase 6: Developer Experience and Hardening

- Add high-level recipes for common report shapes.
- Add diagnostics for invalid sheet names, missing sources, mapping conflicts, and unsupported output targets.
- Add sample projects for exports, imports, templates, and CSV.
- Add tests comparing generated workbook contents for representative reports.

## Initial Folder Proposal

```text
DocWeave/
  Export/
    ExcelReport.cs
    Builders/
    Rendering/
    Styles/
  Import/
    ExcelReader.cs
    CsvReader.cs
    Mapping/
  Data/
    IRowSource.cs
    Adapters/
  Specs/
    WorkbookSpec.cs
    SheetSpec.cs
    DataSourceSpec.cs
  Templates/
  Validation/
  docs/
```

## Summary

DocWeave should use a clear greenfield pipeline:

```text
Fluent API -> Specs -> Data Adapters -> Renderer/Reader -> File/Stream/Objects
```

That gives the project room to support rich Excel reporting features while making imports, templates, CSV, and future structured formats easier to reason about and test.
