# DocWeave Examples

This project is a runnable how-to for the main DocWeave APIs.

Run it from the repository root:

```powershell
dotnet run --project Examples\DocWeave.Examples\DocWeave.Examples.csproj -- Examples\Output
```

The command creates one folder per scenario. Template scenarios include:

- `template.json`, the DocWeave template.
- `api-request.json`, the equivalent Web API payload.
- sample source files or generated export files.
- `import-result.json` for import scenarios.

The examples are grouped in `Program.cs`:

- 5 XLSX template scenarios.
- 5 CSV template scenarios.
- direct XLSX and CSV fluent/read scenarios for comparison.


## Border Examples

The additional borders folder contains direct-financials.xlsx, border-combinations.xlsx,
template-financials.xlsx, template.json and api-request.json. See BorderExamples.cs for
commented direct and template calls. The report has 20 data rows, currency columns,
a thin top and double bottom on totals, and first-column styling through the totals label.

BorderTop/Left/Bottom/Right accept independent styles and colors. Border sets all sides
and resets previous overrides; later side calls override it. None removes one side.
In JSON, border.top/left/bottom/right are objects with style and optional color.
A global border.style applies first. With only side objects, unspecified sides have no border.
RangeStyle replaces the complete cell style, so the totals label needs its own complete style.
## Reporting Cookbook

See ReportingExamples.cs for seven additional commented scenarios, generated under reporting/:

| File | Demonstrates |
| --- | --- |
| 01-formulas-totals.xlsx | Profit formula column, separate summary formula, currency, totals borders, frozen headings and filter |
| 02-conditional-styles.xlsx | Loss-making row highlighting, revenue cell rules, review-status value rules, alternating rows |
| 03-pivot-quarterly.xlsx | Native source table, region/quarter pivot, SUM and grand totals |
| 04-pivot-multiple-values.xlsx | Channel report filter, SUM and average value fields |
| 05-repeater-formulas.xlsx | Two cards across, field headings and profit formulas relative to each card |
| 06-calculated-csv-totals.csv | Precomputed profit and totals, semicolon delimiter and custom heading |
| 07-template-pivot-totals.xlsx | JSON styles, totals and quarterly pivot; includes 07-template.json and 07-api-request.json |

The runner validates every new XLSX with OpenXmlValidator before saving it.
This checks document structure, not Excel rendering or calculated results.
Excel calculates formulas; DocWeave does not evaluate them. Pivot source ranges
exclude the totals row. Complex pivot combinations remain subject to the library's
renderer limitations; inspect output in Excel for the target workflow.

Conditional styles are evaluated once by C# during export. Live Excel conditional
formatting rules, color scales and data bars are not currently supported. Rules
apply complete styles rather than merging individual style properties; include
number formats in a rule when needed. CSV cannot carry spreadsheet styles or
calculated formulas, so its example writes computed ordinary values.

API payloads use the existing DocWeaveExportRequest contract; routes and authentication
are provided by the application hosting DocWeave.