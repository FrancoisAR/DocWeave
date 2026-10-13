# Template Generator Tool Technical Specification

## Overview

The Template Generator Tool is a development-time utility that creates and validates DocWeave template JSON and Web API request payloads.

The tool should use DocWeave's local template construction and validation interfaces where possible. It should not require access to a DocWeave Web
API, but it may optionally call validation-only endpoints when a service URL is configured.

## Project Layout

Proposed project structure:

```text
DocWeave.TemplateTool/
  Commands/
  Authoring/
  Json/
  Validation/
  Remote/
  Program.cs
```

The project should be separate from the core library so production users do not take a dependency on CLI or tool packages.

## Target Framework

The first version should target modern .NET only, for example:

```text
net10.0
```

The generated JSON remains runtime-neutral. Applications that consume the generated JSON do not need to use .NET.

## Dependency Direction

The tool may reference:

- `DocWeave`.
- `System.CommandLine` or another CLI parser if chosen.
- `System.Text.Json`.

The tool should not be referenced by:

- `DocWeave`.
- `DocWeaveNS`.
- Web API contract projects.

## Authoring Model

The tool should have a small authoring model that maps to template JSON.

The model should expose interfaces for the construction surface, for example:

```csharp
public interface ITemplateDocumentBuilder
{
    ITemplateDocumentBuilder SchemaVersion(string version);
    ITemplateDocumentBuilder DataSource(string name, Action<IDataSourceBuilder> configure);
    string ToJson();
}
```

This surface should be narrower than the fluent rendering API. Its purpose is to construct declarative templates, not to render files directly.

The authoring model should support:

- Excel export templates.
- CSV export templates.
- Excel import templates.
- CSV import templates.
- Web API request payloads.

## JSON Handling

The tool should parse and write JSON using structured APIs, not string concatenation.

Recommended rules:

- use `System.Text.Json` DOM or typed records for parsing.
- write indented JSON by default.
- preserve stable property order for generated documents.
- allow loading templates that contain unknown optional fields.
- avoid dropping unknown fields unless the user explicitly normalizes the document.

## Commands

Initial CLI commands:

```text
docweave template new --kind excel-export --out template.json
docweave template open --template template.json
docweave template validate --template template.json --data sample-data.json
docweave template preview --template template.json --data sample-data.json
docweave template save --template template.json --out template.json
docweave request build --template template.json --data sample-data.json --format xlsx --out request.json
docweave template verify --template template.json --data sample-data.json --api-url https://example/api
```

`open` may be a no-op for a pure CLI, but it is useful terminology if the same backend later supports a desktop or web UI.

## Local Validation Pipeline

Local validation should produce a common notification model:

```csharp
public sealed record TemplateNotification(
    TemplateNotificationSeverity Severity,
    string Code,
    string Message,
    string? Path = null,
    string? SuggestedFix = null);
```

Suggested severity enum:

```csharp
public enum TemplateNotificationSeverity
{
    Info,
    Warning,
    Error
}
```

Validation stages:

1. JSON syntax.
2. schema version.
3. template kind.
4. required contract sections.
5. data source binding.
6. field references.
7. style references.
8. layout checks.
9. renderer capability checks.
10. security checks.

The current library validation can be used as the first implementation, but the tool should normalize results into `TemplateNotification`.

## Remote Verification

Remote verification is optional. It should be enabled only when the user supplies a service URL or profile.

Suggested flow:

1. Build a validation request from the current template and sample data.
2. Send it to the configured Web API validation endpoint.
3. Parse returned notifications.
4. Display local and remote results separately.
5. Mark version/capability differences clearly.

Remote verification must not call export/import execution endpoints in the initial version.

## Proposed Web API Contract Additions

Add validation request/response records to the Web API contract layer:

```csharp
public sealed record DocWeaveTemplateValidationRequest(
    string Format,
    string? TemplateId,
    string? TemplateJson,
    IReadOnlyDictionary<string, JsonElement>? DataSources = null,
    IReadOnlyDictionary<string, JsonElement>? Parameters = null);

public sealed record DocWeaveTemplateValidationResponse(
    bool IsValid,
    IReadOnlyList<DocWeaveTemplateNotification> Notifications,
    IReadOnlyList<string>? SupportedSchemaVersions = null);

public sealed record DocWeaveTemplateNotification(
    string Severity,
    string Code,
    string Message,
    string? Path = null,
    string? SuggestedFix = null);
```

Suggested endpoints:

```text
POST /templates/validate
POST /templates/verify
POST /exports/validate
POST /imports/validate
```

Endpoint behaviour:

- return `200 OK` with notifications for valid and invalid templates.
- avoid generating or storing files.
- avoid logging data values.
- include schema/capability information where helpful.

## Data Binding

The tool needs an adapter from sample JSON to DocWeave tabular inputs.

Rules:

- JSON arrays of objects become named tabular data sources.
- An in-process `DataTable` is represented by one named JSON array whose property names match the table columns.
- An in-process `DataSet` is represented by several named JSON arrays, using each `DataTable.TableName` as the JSON data source name.
- Enumerated object collections are represented by one named JSON array whose property names match the public readable properties used by `WithData`.
- JSON object properties can become parameters or named scalar sources.
- primitive values become parameters.
- null values are preserved.
- unsupported nested values should produce validation warnings or errors unless the target template supports them.

Example:

```json
{
  "dataSources": {
    "sales": [
      { "Region": "North", "Amount": 1200.50 }
    ]
  },
  "parameters": {
    "reportDate": "2026-10-08"
  }
}
```

## Request Payload Generation

The tool should emit request JSON matching current Web API contract records.

Export request output should match `DocWeaveExportRequest`.

Import request output should match `DocWeaveImportRequest`.

The tool should support two template modes:

- inline template JSON.
- stored template id.

For inline mode, the tool should embed the current template JSON into the request payload. For stored mode, the tool should write the `templateId`
and omit `templateJson`.

## Persistence

Template files should be plain JSON.

The tool may also support an optional workspace/project file later, but it is not required for the first version.

Possible future project file:

```json
{
  "template": "monthly-sales.template.json",
  "sampleData": "monthly-sales.sample-data.json",
  "format": "xlsx",
  "apiProfiles": {
    "dev": {
      "baseUrl": "https://localhost:5001"
    }
  }
}
```

## Testing Strategy

Tests should cover:

- new template skeleton generation.
- loading existing template JSON.
- saving deterministic JSON.
- local validation success and failure.
- sample data binding.
- Excel preview output.
- request payload generation.
- round-tripping request payloads into Web API contract records.
- remote verification client using a fake HTTP handler.

## Initial Milestone

1. Add the tool project.
2. Add shared notification models.
3. Implement `template new`.
4. Implement `template validate`.
5. Implement sample data binding.
6. Implement `template preview` for Excel.
7. Implement `request build`.
8. Add validation-only Web API contract records.
9. Add tests for request payload round-tripping.

## Risks

- If the tool grows into a visual designer too early, it may slow down the core template work.
- If validation remains exception-based, users will receive vague errors.
- If local and remote validation use different package versions, users may see confusing differences.
- If unknown JSON fields are dropped during save, the tool may damage hand-authored templates.

## Open Decisions

- Whether the first UI is CLI-only or whether a simple desktop UI is also planned.
- Whether validation contract records should live in `DocWeave.WebApi` or a future dedicated contracts package.
- Whether template authoring interfaces should be public library API or internal to the tool until proven stable.
- Whether import and export templates should share one builder hierarchy or separate builders.
