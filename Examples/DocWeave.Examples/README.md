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