# Template Generator Tool Functional Specification

## Purpose

The Template Generator Tool helps developers create DocWeave import and export template JSON when their application cannot reference the
`DocWeave` library directly and can only call a Web API with JSON payloads.

The tool is primarily a developer-machine utility. It should work without direct Web API access by using DocWeave's local template interfaces and
validation logic. On machines where a DocWeave Web API is reachable, the tool may optionally call validation/verification endpoints to check that
the generated template is accepted by the deployed service.

## Users

- Developers building applications that call a DocWeave Web API.
- Developers creating stored templates for a DocWeave-backed service.
- Testers validating template JSON against sample data.
- Support engineers diagnosing why a template fails before it reaches production.

## Scope

The tool must cover the template functionality exposed through the DocWeave API contract:

- Excel export templates.
- CSV export templates.
- Excel import templates.
- CSV import templates.
- Data source declarations.
- Parameters.
- Styles.
- Sheet templates and regions.
- Import source mappings.
- Web API request payload generation.
- Local validation.
- Optional remote verification through Web API validation endpoints.

The first version does not need to call export/import execution endpoints or download rendered files from the Web API.

## Core Workflows

### Create New Template

The user can create a new template from a starting shape:

- Excel export.
- CSV export.
- Excel import.
- CSV import.

The tool should generate a valid skeleton with:

- `schemaVersion`.
- `kind`.
- required top-level sections for that template kind.
- an initial data source or import source placeholder.
- a minimal sheet or CSV column configuration where applicable.

### Import Existing JSON

The user can open an existing template JSON file and load it into the tool.

The tool should:

- parse the JSON.
- identify the template kind.
- preserve unknown optional fields where possible.
- show validation problems without discarding the document.
- allow the user to save the template back to disk.

### Edit Template

The tool should support editing template concepts rather than only raw JSON:

- output format and file name.
- data sources and fields.
- parameters.
- styles.
- sheets.
- sheet templates.
- regions such as cells, tables, repeaters, charts, pivots and validation regions.
- CSV columns and output options.
- import sheet/source options.
- import column mappings.

A raw JSON view may exist, but the main value is a guided template construction surface.

### Save Template

The user can save the current template as JSON.

Save behaviour should:

- produce formatted, deterministic JSON.
- keep property order stable where practical.
- preserve the selected schema version.
- avoid writing secrets into the template.
- warn if the template has validation errors.

### Export Web API Request Payload

The user can export a full request JSON payload that another application can send to the Web API.

For export requests, this should align with `DocWeaveExportRequest`:

- `format`.
- `templateId` or `templateJson`.
- `dataSources`.
- `parameters`.
- `delivery`.
- `audit`.

For import requests, this should align with `DocWeaveImportRequest`:

- `format`.
- `source`.
- `templateId` or `templateJson`.
- `audit`.

The payload export should support sample data so developers can paste or commit an end-to-end example.

Sample export data should support the same practical source shapes as the template exporter:

- object/enumerated collections represented as JSON arrays of objects.
- dictionary rows represented as JSON arrays of objects with flexible field names.
- `DataTable`-style data represented as one named JSON array.
- `DataSet`-style data represented as multiple named JSON arrays, one per table/source name.

### Validate Locally

The user can validate a template without calling a Web API.

Local validation should check:

- JSON syntax.
- schema version.
- template kind.
- required sections.
- data source references.
- source field references.
- style references.
- region type support.
- cell and range references.
- region overlap.
- CSV delimiter/output options.
- import mapping consistency.
- security rules, such as no inline secrets.

Validation should return specific notifications, not one generic failure.

Each notification should include:

- severity: error, warning or information.
- code.
- message.
- JSON path when known.
- suggested fix when practical.

### Preview

The user can preview non-rendering information before using the template:

- Excel region bounds.
- sheet names.
- required data sources.
- required parameters.
- CSV output column list.
- import output column list.

Preview does not need to render or export a final file.

### Optional Remote Verify

On machines configured with a DocWeave Web API URL, the tool can send a template and optional sample data to a validation/verification endpoint.

Remote verification should:

- not create an export file.
- not persist a template unless the endpoint explicitly says it does.
- check the deployed service's actual supported features.
- return specific notifications for failures.
- report differences between local validation and remote verification where possible.

This is useful when a developer's local package version differs from the deployed Web API version.

## Proposed Web API Validation Endpoints

The Web API should eventually expose validation-only endpoints:

```text
POST /templates/validate
POST /templates/verify
POST /exports/validate
POST /imports/validate
```

Suggested distinction:

- `validate`: checks template structure and rules.
- `verify`: checks template plus supplied sample data and host-specific capabilities.

Neither endpoint should produce a final document.

## Functional Requirements

- The tool can create new template JSON for each supported template kind.
- The tool can open existing template JSON from disk.
- The tool can save edited template JSON to disk.
- The tool can validate templates locally.
- The tool can display validation notifications with actionable detail.
- The tool can export Web API request payload JSON.
- The tool can load sample data JSON for validation and request generation.
- The tool can model `DataTable`, `DataSet`, and enumerable collection inputs as named JSON data sources.
- The tool can preview Excel layout metadata without rendering.
- The tool can preview CSV/import column metadata without rendering.
- The tool can optionally call remote validation endpoints when configured.
- The tool can run entirely without remote API access.

## Non-Functional Requirements

- The tool should be deterministic: repeated saves should not reorder the document unnecessarily.
- The tool should be safe: it must not store passwords or secrets in template files by default.
- The tool should be portable enough for developer machines.
- The tool should not require Excel to be installed.
- The tool should not require the consuming application to reference `DocWeave`.
- The tool should clearly show which DocWeave template schema version it targets.

## Out Of Scope For Initial Version

- Calling Web API export/import execution endpoints.
- Managing stored templates on a remote server.
- Authentication flows beyond passing a configured token/header for optional verification.
- Full visual spreadsheet design.
- Real Excel rendering previews.
- Template approval workflows.
- Background operation polling.
