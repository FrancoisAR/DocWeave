# Template API Contract V1

This document defines the declarative template alternative to the fluent API. The goal is to let trusted internal services submit export requests over a Web API without exposing the C# fluent surface method by method.

The template format is a strongly shaped JSON contract. It describes the output workbook, sheet templates, regions, bindings, styles, formulas, charts, pivots, protection, and delivery intent. It should compile into the same internal specs used by the fluent API.

```text
Fluent API          -> WorkbookSpec -> Renderer
Template JSON/API   -> WorkbookSpec -> Renderer
Stored Template     -> WorkbookSpec -> Renderer
```

The renderer should not care whether the workbook was defined by fluent C# code or by a JSON template.

## Design Goals

- Use one rendering engine for fluent and template-driven exports.
- Keep templates declarative, versioned, validateable, and auditable.
- Support Web API clients that cannot or should not reference the C# library.
- Separate layout from data.
- Allow stored, approved templates and dynamic trusted templates.
- Avoid arbitrary code execution in formulas, predicates, or data transformations.
- Keep passwords and secrets out of template files.
- Support Excel first, with CSV using a reduced compatible contract.

## Top-Level API Modes

### Dynamic Template Request

Useful for trusted internal services, test tools, and template editors.

```json
{
  "template": {
    "schemaVersion": "1.0",
    "kind": "excelWorkbook",
    "output": {
      "format": "xlsx",
      "fileName": "currency-report.xlsx"
    },
    "dataSources": [],
    "styles": {},
    "sheetTemplates": {},
    "sheets": []
  },
  "data": {},
  "parameters": {},
  "secrets": {},
  "audit": {
    "requestedBy": "finance-service",
    "reason": "Daily currency export"
  }
}
```

### Stored Template Request

Preferred for normal Web API use. The server loads a versioned, approved template by id.

```json
{
  "templateId": "currency-report-v1",
  "data": {},
  "parameters": {},
  "secrets": {},
  "audit": {
    "requestedBy": "finance-service",
    "reason": "Daily currency export"
  }
}
```

Stored templates should record:

- template id
- schema version
- template version
- owning system/team
- allowed output formats
- required sources
- approval status
- created/modified audit details

## Web API Shape

Suggested endpoints:

```text
POST /exports
GET  /exports/{operationId}
POST /exports/{operationId}/cancel
GET  /templates
GET  /templates/{templateId}
POST /templates/validate
POST /templates/{templateId}/render
```

Small exports can return the file directly. Long-running exports should return an operation id.

```json
{
  "operationId": "exp_20260922_000123",
  "status": "running",
  "links": {
    "status": "/exports/exp_20260922_000123",
    "cancel": "/exports/exp_20260922_000123/cancel"
  }
}
```

## Template Structure

```json
{
  "schemaVersion": "1.0",
  "kind": "excelWorkbook",
  "output": {},
  "dataSources": [],
  "parameters": [],
  "styles": {},
  "sheetTemplates": {},
  "sheets": []
}
```

### Output

```json
{
  "format": "xlsx",
  "fileName": "daily-currency-report.xlsx",
  "culture": "en-GB",
  "timeZone": "Europe/London",
  "encryption": {
    "mode": "workbookOpenPassword",
    "passwordRef": "workbookPassword"
  }
}
```

Supported initial formats:

- `xlsx`
- `csv`

Later:

- `xlsm` preserve macros only when source workbook/template is macro-enabled
- `xltx` input template producing `xlsx`
- `xltm` input template producing `xlsm`
- document and presentation outputs in separate renderers

## Data Sources

The template declares required source shape. The request supplies source values.

```json
{
  "dataSources": [
    {
      "name": "currencies",
      "kind": "array",
      "required": true,
      "columns": [
        { "name": "code", "type": "string", "required": true },
        { "name": "country", "type": "string", "required": true },
        { "name": "lastHighRate", "type": "number", "required": true },
        { "name": "lastLowRate", "type": "number", "required": true },
        { "name": "lastUpdated", "type": "date", "required": true }
      ]
    }
  ]
}
```

Request data:

```json
{
  "data": {
    "currencies": [
      {
        "code": "USD",
        "country": "United States",
        "lastHighRate": 1.2875,
        "lastLowRate": 1.242,
        "lastUpdated": "2026-09-18"
      }
    ]
  }
}
```

Possible source kinds:

- `array`
- `table`
- `dataSet`
- `dictionary`
- `scalar`
- `binary`
- `fileReference`
- `queryReference`

Initial implementation should support `array`, `table`, `dictionary`, and `scalar`.

## Styles

Templates should use named style tokens. This avoids repeating the same formatting everywhere.

```json
{
  "styles": {
    "title": {
      "font": { "name": "Aptos Display", "size": 16, "bold": true, "color": "#1F4E78" }
    },
    "card": {
      "fill": { "color": "#F8FBFD" },
      "border": { "style": "thin", "color": "#B7C9D6" }
    },
    "numberGood": {
      "numberFormat": "number",
      "fill": { "color": "#E2F0D9" },
      "border": { "style": "thin", "color": "#70AD47" }
    }
  }
}
```

Style properties:

- `font.name`
- `font.size`
- `font.bold`
- `font.italic`
- `font.underline`
- `font.color`
- `fill.color`
- `fill.pattern`
- `border.style`
- `border.color`
- `alignment.wrapText`
- `alignment.rotation`
- `numberFormat`

Initial number formats:

- `general`
- `number`
- `currency`
- `percent`
- `date`
- `text`

## Workbook Sheets

Output sheets can use inline templates or reference reusable `sheetTemplates`.

```json
{
  "sheets": [
    {
      "name": "Currency Cards",
      "template": "currencyCardsSheet"
    }
  ]
}
```

Reusable sheet template:

```json
{
  "sheetTemplates": {
    "currencyCardsSheet": {
      "regions": []
    }
  }
}
```

## Sheet Regions

Every visible thing on a sheet is a region.

Supported initial region types:

- `cell`
- `table`
- `repeater`
- `formulaCell`
- `rangeStyle`
- `chart`
- `pivotTable`

Later:

- `image`
- `namedRange`
- `validation`
- `conditionalFormatting`
- `templateRange`
- `pageSetup`
- `printArea`

### Cell Region

```json
{
  "type": "cell",
  "cell": "A1",
  "value": "Currency Rate Cards",
  "style": "title"
}
```

Parameter binding:

```json
{
  "type": "cell",
  "cell": "A2",
  "value": "{{parameters.reportDate}}",
  "style": "subtitle"
}
```

### Table Region

```json
{
  "type": "table",
  "source": "currencies",
  "startCell": "A4",
  "headers": true,
  "columns": [
    { "source": "code", "caption": "Code" },
    { "source": "country", "caption": "Country" },
    { "source": "lastHighRate", "caption": "Last High", "style": "numberGood" },
    { "source": "lastLowRate", "caption": "Last Low", "style": "numberWarning" }
  ],
  "styles": {
    "header": "tableHeader",
    "data": "tableData",
    "alternateRow": "tableAlternate"
  },
  "totals": {
    "position": "bottom",
    "label": "Totals",
    "style": "totals"
  }
}
```

### Repeater Region

Repeater maps each source item to a rectangular block, similar to a data repeater or card layout.

```json
{
  "type": "repeater",
  "source": "currencies",
  "startCell": "A4",
  "itemsPerRow": 3,
  "itemSize": { "columns": 4, "rows": 6 },
  "gap": { "columns": 1, "rows": 1 },
  "titleField": "code",
  "fields": [
    { "source": "country", "caption": "Country" },
    { "source": "lastHighRate", "caption": "Last High", "valueStyle": "numberGood" },
    { "source": "lastLowRate", "caption": "Last Low", "valueStyle": "numberWarning" },
    { "source": "lastUpdated", "caption": "Last Updated", "valueStyle": "dateValue" },
    {
      "caption": "High-Low Spread",
      "formula": "VALUE(lastHighRate) - VALUE(lastLowRate)",
      "labelStyle": "spreadLabel",
      "valueStyle": "spreadValue"
    }
  ],
  "styles": {
    "card": "card",
    "title": "cardTitle",
    "label": "cardLabel",
    "value": "cardValue"
  }
}
```

Formula rules:

- Formula expressions must use a safe expression language, not arbitrary C#.
- Expressions can reference current item fields by name.
- Expressions can reference item metadata such as `item.index`, `item.row`, and `item.column`.
- The compiler translates safe expressions into Excel formulas or precomputed values.

For a summary item, the data source can include one explicit summary record, or the template can define a generated summary region.

```json
{
  "type": "repeaterSummary",
  "source": "currencies",
  "startCell": "F25",
  "title": "Summary",
  "itemSize": { "columns": 4, "rows": 6 },
  "fields": [
    { "caption": "Country", "value": "10 currencies" },
    { "caption": "Avg Last High", "formula": "AVG(source.lastHighRate)" },
    { "caption": "Avg Last Low", "formula": "AVG(source.lastLowRate)" },
    { "caption": "High-Low Spread", "formula": "AVG(source.lastHighRate) - AVG(source.lastLowRate)" }
  ],
  "styles": {
    "card": "summaryCard",
    "title": "summaryTitle",
    "label": "cardLabel",
    "value": "spreadValue"
  }
}
```

### Formula Cell Region

```json
{
  "type": "formulaCell",
  "cell": "B12",
  "formula": "SUM(B5:B11)",
  "style": "totals"
}
```

Formula policy:

- `formula` is an Excel formula fragment.
- Leading `=` is optional.
- Template validation should reject formulas that contain external workbook references unless explicitly allowed.
- Web API callers should be able to disable formulas and force precomputed values for high-trust/security scenarios.

### Chart Region

```json
{
  "type": "chart",
  "title": "Income By Month",
  "chartType": "column",
  "dataSource": {
    "sheet": "ChartData",
    "categoryRange": "$A$4:$A$9",
    "valueRange": "$B$4:$B$9"
  },
  "position": {
    "topLeft": "A4",
    "bottomRight": "J22"
  },
  "style": 10,
  "seriesColor": "#4472C4"
}
```

Initial chart types:

- `column`
- `bar`
- `line`
- `area`
- `pie`
- `doughnut`
- `scatter`

Later:

- multiple series
- combo charts
- secondary axes
- pivot charts
- data labels
- legend placement
- chart titles and axis titles as separate options

### Pivot Table Region

```json
{
  "type": "pivotTable",
  "name": "SalesByRegion",
  "source": {
    "sheet": "Sales Data",
    "range": "$A$3:$I$15"
  },
  "startCell": "A3",
  "rows": ["Region", "Product"],
  "columns": ["Quarter"],
  "values": [
    { "field": "Amount", "aggregate": "sum", "caption": "Total Sales" }
  ],
  "grandTotals": {
    "rows": true,
    "columns": true
  },
  "style": "PivotStyleMedium9"
}
```

Pivot support:

- one or more row fields
- one or more column fields
- report filters (`"filters": ["Channel"]`), one row above the table each; the pivot must start at row `filters + 2` or lower
- one or more value fields (`sum`/`count`/`average`/`min`/`max`); with more than one, each column item is repeated per value
- two value fields may share one source field, for example a sum and an average of `Amount`
- source on same or different sheet

Value captions must be unique and must not repeat a source field name (Excel rejects that), and a field can sit on only one axis. Violations fail with a clear error instead of producing a workbook Excel would repair.

Later:

- slicers
- calculated fields
- pivot charts

## CSV Template

CSV uses a smaller template shape because layout, fonts, borders, charts, pivots, and formulas do not apply.

```json
{
  "schemaVersion": "1.0",
  "kind": "csv",
  "output": {
    "format": "csv",
    "fileName": "currencies.csv",
    "delimiter": ",",
    "includeHeaders": true,
    "encoding": "utf-8"
  },
  "source": "currencies",
  "columns": [
    { "source": "code", "caption": "Code" },
    { "source": "country", "caption": "Country" },
    { "source": "lastHighRate", "caption": "Last High" },
    { "source": "lastLowRate", "caption": "Last Low" }
  ],
  "safety": {
    "neutralizeFormulaText": true
  }
}
```

## Validation

Template validation should happen before rendering.

Validation stages:

1. JSON syntax validation.
2. Schema version validation.
3. Contract validation.
4. Source binding validation.
5. Layout validation.
6. Renderer capability validation.
7. Security validation.

Validation result shape:

```json
{
  "valid": false,
  "errors": [
    {
      "path": "$.sheetTemplates.currencyCardsSheet.regions[2].fields[1].source",
      "code": "SourceFieldMissing",
      "message": "Field 'lastHighRate' was not found in source 'currencies'."
    }
  ],
  "warnings": [
    {
      "path": "$.sheets[0].name",
      "code": "SheetNameTruncated",
      "message": "Sheet name will be truncated to Excel's 31-character limit."
    }
  ]
}
```

Important validation rules:

- Sheet names must be valid and unique after sanitization.
- Region start cells must be valid A1 references.
- Region layout must not overlap unless overlap mode is explicitly allowed.
- Referenced styles must exist.
- Referenced sources must exist.
- Referenced fields must exist.
- Formulas must comply with formula policy.
- Password values must be supplied through `secrets`, not inline template text.
- Unsupported features must fail clearly instead of silently dropping content.

## Security Model

Templates submitted over the API are data, not code.

Rules:

- No arbitrary C# predicates or delegates.
- No unrestricted file paths in dynamic templates.
- No raw passwords in stored templates.
- No external workbook links unless explicitly allowed.
- No external data queries unless source kind and caller permission allow it.
- No Web API rendering to arbitrary server paths unless using an approved storage adapter.
- Audit every export request and template id/version used.

Use references for sensitive values:

```json
{
  "output": {
    "encryption": {
      "mode": "workbookOpenPassword",
      "passwordRef": "exportPassword"
    }
  }
}
```

Runtime request:

```json
{
  "secrets": {
    "exportPassword": "supplied-at-runtime"
  }
}
```

## Translation To Internal Specs

The template compiler should be separate from rendering.

Suggested modules:

```text
DocWeave.TemplateContracts
  JSON DTOs / contract records

DocWeave.TemplateValidation
  contract validation
  source validation
  layout validation
  security validation

DocWeave.TemplateCompiler
  template JSON -> WorkbookSpec / CsvSpec

DocWeave.Excel
  WorkbookSpec -> xlsx

DocWeave.Csv
  CsvSpec -> csv

DocWeave.WebApi
  HTTP request/response, operation tracking, auth, audit, progress
```

Compiler output should be the same internal model as the fluent API:

```text
TemplateWorkbook
  -> TemplateCompiler
    -> ExcelWorkbookSpec
      -> ExcelWorkbookRenderer
```

## C# Contract Sketch

```csharp
public sealed record TemplateExportRequest(
    TemplateDocument? Template,
    string? TemplateId,
    IReadOnlyDictionary<string, object?> Data,
    IReadOnlyDictionary<string, object?> Parameters,
    IReadOnlyDictionary<string, string> Secrets,
    ExportAuditMetadata? Audit);

public sealed record TemplateDocument(
    string SchemaVersion,
    string Kind,
    OutputTemplate Output,
    IReadOnlyList<DataSourceDeclaration> DataSources,
    IReadOnlyDictionary<string, StyleTemplate> Styles,
    IReadOnlyDictionary<string, SheetTemplate> SheetTemplates,
    IReadOnlyList<OutputSheetTemplate> Sheets);

public interface ITemplateCompiler
{
    ExcelWorkbookSpec CompileExcel(TemplateDocument template, TemplateBindingContext context);
    CsvExportSpec CompileCsv(TemplateDocument template, TemplateBindingContext context);
}

public interface ITemplateValidator
{
    TemplateValidationResult Validate(TemplateDocument template, TemplateBindingContext context);
}
```

## Versioning

Every template must specify `schemaVersion`.

Versioning rules:

- Minor versions may add optional fields.
- Major versions may change interpretation.
- Stored templates should be pinned to a schema version.
- The Web API should report supported template schema versions.
- The compiler should be able to reject future versions clearly.

Example:

```json
{
  "schemaVersion": "1.0"
}
```

## Initial Implementation Scope

V1 should support:

- Dynamic and stored template request shapes.
- Excel `.xlsx` output.
- CSV output.
- Sheet templates.
- Cell regions.
- Table regions.
- Repeater regions.
- Formula cells.
- Basic style tokens.
- Current chart types.
- Current pivot table capabilities.
- Workbook password reference contract, even if encryption implementation arrives later.
- Validation result model.
- Compilation into current `ExcelWorkbookSpec`.

Out of scope for first implementation:

- Word/PowerPoint renderers.
- Macro creation.
- Arbitrary user code.
- Multiple chart series.
- Template range copying from existing workbook files.
- External data queries.

## Complete Currency Card Example

```json
{
  "schemaVersion": "1.0",
  "kind": "excelWorkbook",
  "output": {
    "format": "xlsx",
    "fileName": "currency-cards.xlsx"
  },
  "dataSources": [
    {
      "name": "currencies",
      "kind": "array",
      "columns": [
        { "name": "code", "type": "string", "required": true },
        { "name": "country", "type": "string", "required": true },
        { "name": "lastHighRate", "type": "number", "required": true },
        { "name": "lastLowRate", "type": "number", "required": true },
        { "name": "lastUpdated", "type": "date", "required": true }
      ]
    }
  ],
  "styles": {
    "title": {
      "font": { "size": 16, "bold": true, "color": "#1F4E78" }
    },
    "card": {
      "fill": { "color": "#F8FBFD" },
      "border": { "style": "thin", "color": "#B7C9D6" }
    },
    "cardTitle": {
      "font": { "size": 13, "bold": true, "color": "#FFFFFF" },
      "fill": { "color": "#1F4E78" },
      "border": { "style": "thin", "color": "#1F4E78" }
    },
    "cardLabel": {
      "font": { "bold": true, "color": "#385723" },
      "fill": { "color": "#E2F0D9" },
      "border": { "style": "thin", "color": "#A9D18E" }
    },
    "cardValue": {
      "border": { "style": "thin", "color": "#D9E2F3" }
    },
    "highRate": {
      "numberFormat": "number",
      "fill": { "color": "#E2F0D9" },
      "border": { "style": "thin", "color": "#70AD47" }
    },
    "lowRate": {
      "numberFormat": "number",
      "fill": { "color": "#FCE4D6" },
      "border": { "style": "thin", "color": "#C65911" }
    },
    "spread": {
      "numberFormat": "number",
      "font": { "bold": true, "color": "#1F4E78" },
      "fill": { "color": "#DDEBF7" },
      "border": { "style": "medium", "color": "#5B9BD5" }
    }
  },
  "sheetTemplates": {
    "currencyCards": {
      "regions": [
        {
          "type": "cell",
          "cell": "A1",
          "value": "Currency Rate Cards",
          "style": "title"
        },
        {
          "type": "repeater",
          "source": "currencies",
          "startCell": "A4",
          "itemsPerRow": 3,
          "itemSize": { "columns": 4, "rows": 6 },
          "gap": { "columns": 1, "rows": 1 },
          "titleField": "code",
          "fields": [
            { "source": "country", "caption": "Country" },
            { "source": "lastHighRate", "caption": "Last High", "valueStyle": "highRate" },
            { "source": "lastLowRate", "caption": "Last Low", "valueStyle": "lowRate" },
            { "source": "lastUpdated", "caption": "Last Updated", "valueStyle": "cardValue" },
            {
              "caption": "High-Low Spread",
              "formula": "VALUE(lastHighRate) - VALUE(lastLowRate)",
              "labelStyle": "spread",
              "valueStyle": "spread"
            }
          ],
          "styles": {
            "card": "card",
            "title": "cardTitle",
            "label": "cardLabel",
            "value": "cardValue"
          }
        },
        {
          "type": "repeaterSummary",
          "source": "currencies",
          "startCell": "F25",
          "title": "Summary",
          "itemSize": { "columns": 4, "rows": 6 },
          "fields": [
            { "caption": "Country", "value": "10 currencies" },
            { "caption": "Avg Last High", "formula": "AVG(source.lastHighRate)", "valueStyle": "highRate" },
            { "caption": "Avg Last Low", "formula": "AVG(source.lastLowRate)", "valueStyle": "lowRate" },
            { "caption": "High-Low Spread", "formula": "AVG(source.lastHighRate) - AVG(source.lastLowRate)", "valueStyle": "spread" }
          ],
          "styles": {
            "card": "card",
            "title": "cardTitle",
            "label": "cardLabel",
            "value": "cardValue"
          }
        }
      ]
    }
  },
  "sheets": [
    {
      "name": "Currency Cards",
      "template": "currencyCards"
    }
  ]
}
```
