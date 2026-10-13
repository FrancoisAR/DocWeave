# Template Generator Tool Design

## Purpose

Some DocWeave users will not be able to reference the `DocWeave` library directly from the application that requests a document. For example, a
service may only be allowed to call an external Web API and send JSON. In that case, creating and maintaining the template JSON by hand is
possible, but it is easy to make mistakes and hard to discover the supported options.

The template generator tool is a developer-only utility for creating, validating, previewing and packaging DocWeave template JSON before that JSON
is copied into, embedded in, or deployed alongside the code that calls the external Web API.

The tool can be used as a one-off on a developer machine. It is not part of the runtime path for systems that call the Web API.

## Design Goals

- Make API template creation easier when the calling application cannot use the fluent `DocWeave` API.
- Let a developer describe import/export options through a small construction surface, then emit the equivalent template JSON.
- Validate generated templates against the same compiler and renderer paths used by the library.
- Produce Web API request JSON that can be sent by non-.NET clients or by .NET applications that do not reference `DocWeave`.
- Keep the tool optional and development-time only.
- Avoid turning the first version into a full visual designer.

## Proposed Shape

Create a separate utility project, for example:

```text
DocWeave.Tool
```

or:

```text
DocWeave.Cli
```

The tool may reference the main `DocWeave` library because it runs on a developer machine. Its output should be plain JSON:

- a template document
- a Web API export/import request payload
- optional preview or validation output
- optional sample rendered `.xlsx` or `.csv` output

The generated JSON is what gets used by the application that calls the external Web API.

## Construction Surface

The original issue asks for "a constructor which has access to the interfaces only." The recommended interpretation is:

- provide a small public authoring surface that models template concepts
- keep it narrower than the full fluent export/import API
- emit declarative JSON rather than rendering directly
- expose interfaces where useful so future tools, tests, or UI designers can build templates without coupling to concrete implementation details

Example conceptual flow:

```csharp
var template = TemplateBuilder.ExcelWorkbook("Monthly Sales")
    .DataSource("sales", source => source
        .Column("Region", "string")
        .Column("Amount", "number"))
    .Sheet("Summary", sheet => sheet
        .Cell("A1", "Monthly Sales")
        .Table("sales", "A4", table => table
            .Column("Region")
            .Column("Amount", caption: "Sales Amount")));

File.WriteAllText("monthly-sales.template.json", template.ToJson());
```

This constructor surface should be treated as an authoring aid, not as a replacement for the existing fluent render API.

## Command-Line MVP

The first version should support file-in/file-out commands:

```text
docweave template validate --template report.template.json --data sample-data.json
docweave template preview --template report.template.json --data sample-data.json
docweave template render --template report.template.json --data sample-data.json --out report.xlsx
docweave request build --template report.template.json --data sample-data.json --format xlsx --out request.json
```

The commands map to existing capabilities:

- `ExcelTemplateWriter.FromJson(...).Validate()`
- `ExcelTemplateWriter.FromJson(...).Preview()`
- `ExcelTemplateWriter.FromJson(...).WithRows(...).ToBytes()`
- `CsvTemplateWriter.FromJson(...).Validate()`
- `CsvTemplateWriter.FromJson(...).ToText()`
- `DocWeaveExportRequest` for the API payload shape

## Inputs

The tool should accept:

- a template JSON file
- a sample data JSON file
- optional scalar parameters
- output format, such as `xlsx` or `csv`
- optional delivery metadata used by the Web API request shape
- optional audit metadata for generated request examples

Sample data should use the same source names as the template:

```json
{
  "sales": [
    { "Region": "North", "Amount": 1200.50 },
    { "Region": "South", "Amount": 980.00 }
  ],
  "createdBy": "finance-service"
}
```

For export templates, JSON arrays should become tabular data sources. JSON objects can be treated as scalar dictionaries or parameters depending on
the command.

## Outputs

The tool should be able to produce:

- formatted template JSON
- validation results
- Excel layout preview JSON
- rendered sample output
- full Web API request JSON

Example generated request:

```json
{
  "format": "xlsx",
  "templateId": null,
  "templateJson": "{ ... }",
  "dataSources": {
    "sales": [
      { "Region": "North", "Amount": 1200.50 },
      { "Region": "South", "Amount": 980.00 }
    ]
  },
  "parameters": {
    "createdBy": "finance-service"
  },
  "delivery": {
    "kind": "download",
    "fileName": "monthly-sales.xlsx"
  },
  "audit": {
    "clientId": "template-generator",
    "requestedBy": "developer"
  }
}
```

## Validation Behaviour

Validation should happen before the tool emits a request payload unless explicitly skipped.

The tool should report:

- invalid JSON
- unsupported template schema version
- missing data sources
- missing source fields
- invalid styles
- invalid cell references
- overlapping regions
- unsupported region types
- security issues such as inline secrets when the contract forbids them

The current library validation returns coarse errors at `$`. The tool can start with that result, but should eventually benefit from richer template
validation with precise JSON paths.

## Relationship To Web API

The tool does not call the Web API in the MVP. It creates JSON that can be submitted to the Web API by another application.

A later version may add:

```text
docweave request send --request request.json --url https://example/api/exports
```

That should remain optional because authentication, networking and deployment-specific policies belong to the host environment.

## Out Of Scope For MVP

- A graphical template designer.
- Stored template management in a server.
- Authentication against a Web API.
- Background export polling.
- Arbitrary code execution in templates.
- A second rendering engine.
- Replacing the fluent API.

## Implementation Notes

- Keep the tool in a separate project so the core library stays focused.
- Reuse the existing template compiler, validation, preview and render paths.
- Add a small adapter from JSON sample data to `IEnumerable<IReadOnlyDictionary<string, object?>>`.
- Prefer stable JSON output with deterministic property ordering where practical, so generated templates are diff-friendly.
- Support both `DocWeave` and, if useful later, `DocWeaveNS` compatibility validation as separate commands or test targets.
- Include example templates and sample data in the repository so the CLI can be tested end to end.

## Suggested First Milestone

1. Add the CLI project.
2. Implement `template validate` for Excel and CSV.
3. Implement JSON sample data binding.
4. Implement `template preview` for Excel.
5. Implement `template render` for local smoke testing.
6. Implement `request build` to emit `DocWeaveExportRequest` JSON.
7. Add tests that verify generated request JSON can be round-tripped into the existing Web API contract records.
