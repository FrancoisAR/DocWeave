# Excel Import/Export Implementation Plan

> **Status: design document.** This is the original implementation plan. It does not track which phases are finished. For what works today see [current-functionality.md](current-functionality.md), and for what is missing see [roadmap.md](roadmap.md).

## Purpose

This plan describes how to implement `.xlsx` import and export in DocWeave. Start with simple, reliable table import/export, then expand into richer combinations such as multiple sheets, multiple regions, styles, formulas, templates, pivots, charts, protection, and encryption.

The implementation must be server-safe:

- Use OpenXML-compatible package manipulation.
- Do not use Excel Interop or COM automation.
- Do not require Excel to be installed.
- Do not support legacy `.xls`.
- Preserve existing macros in `.xlsm`/`.xltm` template workflows later, but do not create or edit macros.

## Implementation Principles

- Build simple table export first.
- Keep public fluent/schema APIs separate from rendering.
- Compile fluent/schema input into specs.
- Validate specs before rendering.
- Resolve layout before writing cells.
- Keep renderers small and focused.
- Add comments where code handles non-obvious OpenXML rules.
- Avoid comments that merely repeat the code.

## Commenting Standards

Use comments to explain intent, constraints, or OpenXML-specific behavior.

Good comments:

```csharp
// Shared strings keep repeated text values small and avoid inline string duplication.
var sharedStringIndex = sharedStrings.GetOrAdd(value);
```

```csharp
// OpenXML stores dates as OADate numbers. The cell style carries the date format.
cell.CellValue = new CellValue(date.ToOADate());
```

Avoid comments like:

```csharp
// Set the value.
cell.CellValue = value;
```

Every non-trivial renderer should have a short summary comment explaining which OpenXML part it owns.

## Target Architecture

```text
Fluent API / Schema
    -> WorkbookSpec
    -> validation
    -> WorkbookPlan
    -> ExcelRenderer
    -> .xlsx
```

Import:

```text
.xlsx source
    -> ExcelReader
    -> sheet/table/range selector
    -> normalized rows
    -> column selection/conversion
    -> DataTable / dictionaries / typed objects
```

## Proposed Excel Modules

```text
DocWeave.Excel
  ExcelReader
  ExcelReport
  OpenXmlWorkbookRenderer
  OpenXmlWorksheetRenderer
  OpenXmlCellWriter
  OpenXmlSharedStringTable
  OpenXmlStyleRegistry
  OpenXmlTableRenderer
  OpenXmlFormulaWriter
  OpenXmlImportReader
  OpenXmlTypeConverter
```

Later:

```text
  OpenXmlTemplateRenderer
  OpenXmlPivotRenderer
  OpenXmlChartRenderer
  OpenXmlDrawingRenderer
  OpenXmlProtectionRenderer
  OpenXmlEncryptionService
```

## Phase 1: Basic Excel Export

Goal: export one data source to one worksheet.

Supported:

- `IEnumerable<T>`
- `DataTable`
- dictionaries later in this phase
- file output
- stream output
- byte array output
- sheet name
- start cell
- headers
- strings, numbers, decimals, dates, booleans
- simple column widths

Internal steps:

1. Convert source to `IRowSource`.
2. Build `WorkbookSpec`.
3. Validate sheet name, start cell, and source.
4. Create workbook package.
5. Create worksheet.
6. Write header row.
7. Write data rows.
8. Save.

Tests:

- Export `DataTable` with 0 rows.
- Export `DataTable` with 10 rows.
- Export `IEnumerable<InvoiceRow>`.
- Export to file, stream, and byte array.
- Verify generated `.xlsx` can be opened as an OpenXML package.
- Verify cell values in expected addresses.

## Phase 2: Basic Excel Import

Goal: import worksheet data into `DataTable`, dictionaries, and raw row values.

Supported:

- file, stream, byte array
- sheet by name or index
- used range
- explicit range
- header row
- include columns by letters
- include columns by headers
- exclude columns
- strings, numbers, dates, booleans

Internal steps:

1. Open workbook package.
2. Resolve shared strings and styles.
3. Find target sheet.
4. Resolve target range.
5. Read rows with `OpenXmlReader` where practical.
6. Resolve headers.
7. Apply column selection/exclusion.
8. Convert cells to normalized values.
9. Return `DataTable`, dictionaries, or raw rows.

Tests:

- Import generated basic workbook.
- Select columns by letters.
- Select columns by headers.
- Exclude columns.
- Import empty sheet.
- Import missing sheet gives useful error.
- Import date/decimal/bool values.

## Phase 3: Styles And Formatting

Goal: support professional basic formatting.

Supported:

- named styles
- header style
- row style
- alternate row style
- column style
- total row style
- font, fill, alignment, borders
- number/date formats
- wrap text

Style precedence:

```text
workbook defaults
  -> region/table style
  -> header/row/alternate/total style
  -> column style
  -> conditional style later
  -> explicit cell/range style
```

Tests:

- Header font/fill applied.
- Alternating rows applied.
- Column-specific style overrides row style.
- Thick border on one column while rest uses thin border.
- Date stored as numeric OADate with date format.

## Phase 4: Multi-Sheet And Multi-Region Export

Goal: compose real workbooks.

Supported:

- multiple sheets
- multiple sources on one sheet
- absolute region placement
- vertical stacking
- horizontal placement
- collision validation
- independent style per region
- independent headers/columns/totals per region

Tests:

- Two tables on one sheet.
- Two tables on separate sheets.
- Region collision detected.
- Vertical stacking after variable row count.

## Phase 5: Formulas And Totals

Goal: support generated calculations.

Supported:

- total row
- sum/count/average/min/max
- formula columns
- formulas with row placeholders
- formulas using Excel table structured references
- workbook calculation mode/full calculate on open

Tests:

- Total row formula references correct range.
- Formula column points to correct row.
- Structured reference formula emitted for Excel table.

## Phase 6: Excel Tables

Goal: add real Excel table objects, not just styled ranges.

Supported:

- table name
- table range
- header row
- total row
- table style
- filters

Implementation order:

1. Render cells first.
2. Compute table range.
3. Add table definition part.
4. Add worksheet relationship.
5. Register table id/name.

## Phase 7: Schema / Shadow Sheet Export

Goal: let WebAPI callers use declarative shadow sheet templates.

Supported:

- workbook schema
- `sheetTemplates`
- output sheets referencing templates
- range sections
- table sections
- styles
- source binding
- parameters

Tests:

- One shadow sheet.
- Multiple shadow sheets.
- Same template used for multiple output sheets.
- Missing source validation.
- Unknown style validation.

## Phase 8: Template Workbook Support

Goal: populate existing `.xlsx` or `.xltx` templates.

Supported:

- open template workbook
- copy sheet
- populate named cells
- populate named ranges
- populate template tables
- preserve formatting

Macro policy:

- `.xltx` template can produce `.xlsx`.
- `.xlsm` or `.xltm` can produce `.xlsm` with macro parts preserved.
- Do not create or edit macros.
- Validate that macro-enabled templates are not accidentally saved as `.xlsx` if preservation is requested.

## Phase 9: Repeating Blocks

Goal: support data-repeater style sheet layouts.

Supported:

- one rectangular block per object
- configurable blocks per row
- row/column spacing
- inline cell definitions
- template range copying later
- nested mini-tables later

## Phase 10: Protection And Encryption

Goal: protect editing and optionally encrypt output.

Supported:

- sheet protection
- workbook structure protection
- locked/unlocked cells
- hide formulas
- full workbook encryption via pluggable service

## Phase 11: Pivot Tables

Goal: add pivot tables after table rendering is stable.

Supported:

- source from Excel table
- source from range
- pivot on same or different sheet
- rows/columns/values/filters
- sum/count/average/min/max
- refresh on open

Implementation note: pivot tables require pivot cache definitions, relationships, and pivot table parts. Isolate this in `OpenXmlPivotRenderer`.

## Phase 12: Charts And Visuals

Goal: add dashboard-style output.

Supported:

- chart from table/range
- chart from pivot later
- images/logos
- basic shapes/connectors later
- sparklines later

Charts depend on rendered ranges, so render them after tables, ranges, formulas, and pivots.

## Phase 13: Import Enhancements

Goal: make import production-friendly.

Supported:

- typed object mapping
- data type declarations
- derived columns
- custom converters
- validation result mode
- password-protected/encrypted workbooks
- batch imports
- progress/cancellation

## Phase 14: Performance And Streaming

Goal: handle large files predictably.

Export:

- use streaming row writing where possible
- avoid holding unnecessary cell state
- avoid shared string explosion where inline strings are better for huge exports

Import:

- use `OpenXmlReader` for large sheets
- avoid loading entire worksheet DOM where possible
- allow cancellation
- report progress by rows/sheets

Tests:

- 10k rows export.
- 100k rows export later.
- 10k rows import.
- cancellation test with artificial slow source.

## First Implementation Checklist

Start with this order:

1. Add OpenXML package dependency.
2. Create `ExcelReport` facade.
3. Create `WorkbookSpec`, `SheetSpec`, `ExportRegionSpec`.
4. Create `IRowSource` and adapters for `DataTable` and `IEnumerable<T>`.
5. Implement single-sheet export.
6. Implement stream/file/byte-array outputs.
7. Implement basic import to `DataTable`.
8. Add column selection by letters/headers for import.
9. Add styles.
10. Add multi-sheet export.
11. Add multi-region export.

## Documentation And Comments Required During Implementation

Every new non-trivial class should include either:

- a concise XML summary when public, or
- a short internal comment near complex OpenXML behavior.

Minimum public XML docs:

- public facades
- public builders
- public result types
- public exceptions
- public options

Minimum internal comments:

- shared string handling
- style id caching
- OADate/date handling
- cell reference translation
- table part creation
- relationship creation
- streaming reader/writer behavior
- macro preservation
- encryption service boundary
