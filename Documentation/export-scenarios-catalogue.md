# DocWeave Export Scenarios Catalogue

## Purpose

This catalogue lists export scenarios the initial CSV and Excel-focused version should be able to support, or should be designed to support soon after the first table export path is stable. The examples intentionally combine different sources, layouts, formatting, and output targets so the interface design can be tested against real usage patterns.

## Scenario List

### 1. Basic Object List To Excel

Export `IEnumerable<Customer>` to one sheet with inferred columns, default headers, auto-fit widths, and a table style.

### 2. Basic DataTable To Excel

Export a single `DataTable` to a new worksheet, using `DataColumn.ColumnName` for headers and `DataColumn.DataType` for value formatting.

### 3. Friendly Headers From DataTable

Export a `DataTable` starting at `B4`, replacing raw column names with friendly headers, applying filter headers, alternating rows, and a total row.

### 4. Multiple DataTables To Separate Sheets

Export every table in a `DataSet`, creating one worksheet per `DataTable`, naming each sheet from `DataTable.TableName`.

### 5. Selected DataSet Tables Only

Export only selected tables from a `DataSet`, ignoring internal or temporary tables.

### 6. DataSet Tables On Same Sheet

Export `Costs`, `Income`, and `Forecast` tables from one `DataSet` onto a single `Financials` worksheet, stacked vertically with spacing rows.

### 7. Mixed DataTable And Object Collection

Place a `DataTable` at `A2` and an `IEnumerable<ForecastRow>` at `H2` on the same sheet, with different styles and number formats.

### 8. Dashboard Summary With Detail Table

Render a summary range at the top of a sheet, then a detailed transaction table below it with filters, frozen headers, totals, and conditional formatting.

### 9. Key-Value Metadata Block Above Table

Export report metadata such as run date, user, parameters, and environment in a two-column block above the main data table.

### 10. Horizontal Tables

Place two small tables side by side, such as current-period values on the left and prior-period values on the right.

### 11. Absolute Region Placement

Place named export regions at exact cells: title at `B1`, summary at `B3`, main table at `B8`, notes at `J8`.

### 12. Grouped Financial Pack

Use group defaults for shared financial formatting, then export `Costs`, `Income`, and `Adjustments` with per-source overrides.

### 13. Repeating Regional Blocks

Render one block per region, three blocks per row, wrapping to the next row after every third region.

### 14. Repeating Blocks From Template Range

Copy a designed template range for each customer/account and bind placeholders inside each copied block.

### 15. Repeating Blocks With Child Tables

Render one project block per project, with a nested mini-table of milestones inside each block.

### 16. Empty Source Message

If a source has no rows, render a styled message such as "No records were available" instead of an empty table.

### 17. Empty Source Skip

Skip an empty optional source and close up surrounding layout spacing.

### 18. Empty Source Headers Only

Render only the table headers when there are no rows, useful for standard template outputs.

### 19. Column Selection And Ordering

Export a source with only selected columns, using an explicit order different from the source object or `DataTable`.

### 20. Hidden Columns

Include technical columns in the sheet but hide them, allowing formulas, validation, pivots, or audit checks to use them.

### 21. Column-Specific Borders

Apply normal thin borders to the table, but thick borders around a status/risk/category column.

### 22. Column-Specific Formatting

Apply date, currency, percentage, integer, and text formats per column.

### 23. Alternating Rows

Apply alternating row fill colors while preserving column-specific number formats and borders.

### 24. Totals Row

Add a total row with sum, count, average, min, and max functions across selected columns.

### 25. Formula Columns

Add calculated columns such as `Profit = Income - Cost` or `Variance = Actual - Budget`.

### 26. Cross-Column Formula With Placeholders

Generate formulas using placeholders so callers do not manually calculate Excel cell references.

### 27. Named Excel Table

Export a table as `SalesData` so formulas, charts, pivots, validation, or downstream users can reference it.

### 28. Named Range

Export a summary block or data range as a named range for formulas and template integration.

### 29. Hidden Source Sheet

Export raw data to a hidden sheet and render user-facing pivots/charts on a visible dashboard sheet.

### 30. Very Hidden Source Sheet

Export internal source data to a very hidden worksheet for templates that should not expose raw data casually.

### 31. Pivot Table From Exported Table

Export source data as an Excel table and create a pivot table on a separate summary sheet.

### 32. Multiple Pivot Tables From One Source

Use one exported table to create several pivots: by region, by month, by product, and by salesperson.

### 33. Pivot Table With Hidden Source Data

Place source data on a hidden sheet and render only the pivot table output to users.

### 34. Pivot Table Plus Detail Table

Render a pivot summary at the top of a sheet and the filtered detail table below it.

### 35. Chart From Table

Create a line, bar, or column chart from an exported table on a dashboard sheet.

### 36. Chart From Pivot Table

Create a pivot chart based on a generated pivot table.

### 37. Combo Chart

Create a chart with columns for sales and a line on a secondary axis for margin percentage.

### 38. Table With Sparklines

Add a sparkline column showing month-by-month trend for each row.

### 39. Conditional Formatting Heatmap

Apply color scale formatting to score, risk, or utilization columns.

### 40. Conditional Formatting Data Bars

Apply data bars to percentage or progress columns.

### 41. Status Highlighting

Highlight failed, overdue, blocked, or high-risk rows using text or value rules.

### 42. Data Validation Drop-Downs

Export a sheet with editable status columns that use drop-down list validation.

### 43. Protected Sheet With Editable Inputs

Protect the worksheet while unlocking only specific input cells or ranges.

### 44. Hyperlink Columns

Convert URL, file path, or document reference columns into clickable hyperlinks.

### 45. Comments And Notes

Add comments to specific cells explaining calculations, assumptions, or review notes.

### 46. Branded Report Header

Add a logo image, title, subtitle, run metadata, and styled divider above the exported data.

### 47. Simple Diagram

Render a basic flow diagram using shapes, text boxes, arrows, and connectors.

### 48. Template Workbook Population

Open an existing `.xlsx` template, populate named ranges and tables, preserve formatting, and save a new output file.

### 49. Template Sheet Copy Per Entity

Copy a template sheet once per region/customer/project and bind that entity's data into the copied sheet.

### 50. CSV Export With Selected Columns

Export selected columns to CSV, with custom headers, delimiter, encoding, and culture-specific formatting.

### 51. CSV Export From DataTable

Export a `DataTable` to CSV with header row, quote/escape handling, and configured null value output.

### 52. CSV Export From Object Collection

Export `IEnumerable<T>` to CSV using selected properties and friendly headers.

### 53. Pipe-Delimited Export

Export data to a pipe-delimited text file for integration with another system.

### 54. Excel And CSV From Same Source

Generate an `.xlsx` report and a `.csv` extract from the same source data using different output configurations.

### 55. Schema-Driven Excel Export

Load a JSON schema defining sheets, table regions, styles, and totals, then bind named sources and render the workbook.

### 56. Schema-Driven CSV Export

Load a CSV schema defining delimiter, headers, encoding, selected columns, and formats, then bind one source and export.

### 57. Dynamic WebAPI Export

Receive schema/spec plus data from a trusted internal WebAPI caller, validate the request, and stream back the generated file.

### 58. TemplateId WebAPI Export

Receive a `templateId`, source data, and parameters from another service, load the approved server-side template, and return the generated `.xlsx`.

### 59. Validation-Only Export Request

Validate that the provided sources match a schema without generating a file, returning missing fields, invalid formats, and layout issues.

### 60. Large Dataset Streaming Export

Export a large source to Excel or CSV with controlled memory use, avoiding unnecessary materialization where possible.

## Coverage Notes

These scenarios test the main design requirements:

- Multiple source types.
- Multiple sheets.
- Multiple regions per sheet.
- Per-region formatting.
- Repeating block layout.
- Table layout.
- Template layout.
- Pivot/chart readiness.
- CSV and Excel outputs.
- Schema-driven and fluent-driven usage.
- WebAPI-friendly invocation.
