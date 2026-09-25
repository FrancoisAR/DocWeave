# DocWeave API Export Schema Design

## Purpose

This document defines how an export schema should look when sent through the WebAPI. The schema should not expose the C# fluent API method-by-method. Instead, it should be a declarative report template that describes the workbook, sheets, layout sections, data bindings, formatting, formulas, protection, and output settings.

The schema is effectively a shadow workbook made from one or more shadow sheets:

```text
API request
  -> schema/template definition
  -> named data sources
  -> WorkbookSpec
  -> Excel or CSV renderer
```

## Recommended Shape

The schema should be a versioned JSON document. JSON is the best first format for WebAPI use because it is easy to send, validate, diff, store, and generate.

Top-level shape:

```json
{
  "schemaVersion": "1.0",
  "kind": "excelExport",
  "output": {},
  "workbook": {},
  "sources": {},
  "sheetTemplates": {},
  "sheets": [],
  "styles": {},
  "secrets": {}
}
```

The schema should describe layout and binding intent. The actual data should usually be sent separately in the same request or bound server-side by name.

## API Request Envelope

For dynamic trusted internal exports:

```json
{
  "format": "xlsx",
  "schema": {
    "schemaVersion": "1.0",
    "kind": "excelExport",
    "workbook": {},
    "sheets": [],
    "styles": {}
  },
  "sources": {
    "invoiceRows": [],
    "runInfo": {}
  },
  "parameters": {
    "generatedDate": "2026-09-20T10:30:00Z",
    "requestedBy": "Francois"
  },
  "audit": {
    "operationId": "optional-client-operation-id",
    "correlationId": "upstream-request-id",
    "requestedBy": "Francois",
    "reason": "Monthly finance export"
  },
  "progress": {
    "mode": "polling",
    "callbackUrl": null
  },
  "delivery": {
    "mode": "response"
  },
  "secrets": {
    "exportOpenPassword": "supplied-at-runtime"
  }
}
```

For stored server-side templates:

```json
{
  "format": "xlsx",
  "templateId": "invoice-export-v1",
  "sources": {
    "invoiceRows": []
  },
  "parameters": {
    "requestedBy": "Francois"
  },
  "audit": {
    "correlationId": "upstream-request-id",
    "requestedBy": "Francois",
    "reason": "Scheduled report"
  },
  "progress": {
    "mode": "eventBus"
  },
  "delivery": {
    "mode": "storage",
    "target": "monthly-reports"
  },
  "secrets": {
    "exportOpenPassword": "supplied-at-runtime"
  }
}
```

The stored-template form is preferable for normal use. The dynamic-schema form is useful for trusted internal services and testing.

## Workbook Definition

```json
{
  "workbook": {
    "properties": {
      "title": "Invoice Export",
      "author": "Reporting Service"
    },
    "calculation": {
      "mode": "automatic",
      "fullCalcOnLoad": true
    },
    "protection": {
      "protectStructure": true,
      "passwordRef": "workbookStructurePassword"
    },
    "encryption": {
      "openPasswordRef": "exportOpenPassword",
      "algorithm": "aes256"
    }
  }
}
```

Workbook options should include:

- document properties
- calculation mode
- workbook structure protection
- full workbook encryption
- default theme/style tokens
- culture/locale defaults
- date/number format defaults

## Audit Metadata

The API request can include audit metadata. Audit metadata is not rendered into the workbook unless the schema explicitly binds it into cells. It is used for logging, traceability, and compliance.

```json
{
  "audit": {
    "operationId": "client-generated-id-if-available",
    "correlationId": "upstream-request-id",
    "requestedBy": "Francois",
    "requestedFor": "Finance",
    "reason": "Monthly finance export",
    "tags": {
      "system": "FinancePortal",
      "environment": "Production"
    }
  }
}
```

Audit guidance:

- The API should generate an operation id if the caller does not provide one.
- Correlation id should flow from upstream services where available.
- Audit metadata should not contain passwords, secrets, source rows, or generated file content.
- Audit logs should record source names and row counts, not full source payloads.
- Encryption/protection choices should be logged as flags, not passwords.

## Progress And Cancellation

Long-running exports should be treated as operations with an operation id. The WebAPI can expose cancellation and progress outside the schema itself:

```http
POST /exports
GET /operations/{operationId}
POST /operations/{operationId}/cancel
```

The request can still declare how the caller wants progress delivered:

```json
{
  "progress": {
    "mode": "polling"
  }
}
```

Other progress modes:

```json
{ "progress": { "mode": "signalR", "channel": "exports" } }
{ "progress": { "mode": "serverSentEvents" } }
{ "progress": { "mode": "webhook", "callbackUrl": "https://internal/callbacks/export-progress" } }
{ "progress": { "mode": "eventBus", "topic": "report-export-progress" } }
```

The schema should not define implementation-specific progress mechanics. It should only carry caller intent. The WebAPI/orchestrator decides whether polling, SignalR, server-sent events, webhook, event bus, or audit-only progress is supported.

Cancellation should stop source loading, rendering, encryption, or delivery where safe. Cancelled operations should return a controlled status and audit event.

## Delivery Targets

For v1, returning the generated file in the HTTP response is enough:

```json
{
  "delivery": {
    "mode": "response"
  }
}
```

Later delivery modes can be added without changing the workbook schema:

```json
{
  "delivery": {
    "mode": "email",
    "to": [ "finance@example.com" ],
    "subject": "Monthly finance export",
    "bodyTemplate": "standard-report-email",
    "attachmentName": "finance-export.xlsx"
  }
}
```

```json
{
  "delivery": {
    "mode": "teams",
    "team": "Finance",
    "channel": "Monthly Reports",
    "message": "The latest finance export is ready.",
    "sendAsLink": true
  }
}
```

```json
{
  "delivery": {
    "mode": "storage",
    "target": "monthly-reports",
    "path": "2026/09/finance-export.xlsx",
    "returnLink": true
  }
}
```

Delivery guidance:

- Delivery should happen after rendering and optional encryption.
- Delivery adapters should be separate from Excel/CSV renderers.
- Delivery failures should be audited separately from render failures.
- Email, Teams, SharePoint, storage, and webhooks should be optional integration modules.
- Delivery configuration may contain sensitive routing information and should be redacted in logs.

## Sheet Template Definition

Each sheet to be created should have its own shadow sheet template. A whole export can therefore contain multiple shadow sheets, each with its own layout, data bindings, styles, formulas, protection, and page setup.

There are two valid ways to define a sheet:

- Define reusable sheet templates under `sheetTemplates`, then create output sheets from those templates.
- Define a one-off inline sheet template directly in the `sheets` array.

Reusable templates are better for WebAPI use because they separate the sheet layout from the specific sheet instance.

Reusable sheet template:

```json
{
  "sheetTemplates": {
    "invoiceSheet": {
      "tabColor": "1F4E78",
      "gridlines": false,
      "freeze": {
        "row": 4
      },
      "protection": {
        "protectSheet": true,
        "passwordRef": "invoiceSheetPassword",
        "allowFiltering": true,
        "allowSorting": true
      },
      "pageSetup": {
        "orientation": "landscape",
        "fitToWidth": 1,
        "fitToHeight": 0
      },
      "sections": []
    }
  }
}
```

Output sheet created from a template:

```json
{
  "sheets": [
    {
      "name": "Invoice Export",
      "template": "invoiceSheet",
      "bindings": {
        "invoiceRows": "invoiceRows",
        "runInfo": "runInfo"
      }
    }
  ]
}
```

Inline one-off sheet:

```json
{
  "sheets": [
    {
      "name": "Ad Hoc Export",
      "template": {
        "gridlines": false,
        "sections": []
      }
    }
  ]
}
```

The sheet template is the shadow sheet. It should describe what appears where, but not require the API caller to understand OpenXML internals.

## Section Model

A sheet is made of sections. Each section is an export region.

Supported v1 section types:

- `range`
- `keyValueBlock`
- `table`
- `spacer`

Likely v2 section types:

- `repeatingBlock`
- `pivotTable`
- `chart`
- `image`
- `diagram`
- `validation`

All sections should support:

```json
{
  "type": "table",
  "name": "invoiceTable",
  "startCell": "B4",
  "style": "standardTable"
}
```

Common properties:

- `type`
- `name`
- `source`
- `startCell`
- `style`
- `visible`
- `when`
- `emptySource`
- `protection`

## Data Sources

The schema should declare required sources by name and expected shape.

```json
{
  "sources": {
    "invoiceRows": {
      "required": true,
      "kind": "table",
      "fields": {
        "InvoiceNo": { "type": "string", "required": true },
        "CustomerName": { "type": "string" },
        "InvoiceDate": { "type": "date" },
        "Status": { "type": "string" },
        "Amount": { "type": "decimal" }
      }
    },
    "runInfo": {
      "required": false,
      "kind": "object",
      "fields": {
        "GeneratedDate": { "type": "datetime" },
        "RequestedBy": { "type": "string" }
      }
    }
  }
}
```

The API can use this for validation before rendering.

## Table Section

```json
{
  "type": "table",
  "name": "invoiceTable",
  "source": "invoiceRows",
  "startCell": "B4",
  "excelTableName": "InvoiceTable",
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
      "type": "string",
      "width": 16
    },
    {
      "key": "CustomerName",
      "header": "Customer",
      "type": "string",
      "width": 32
    },
    {
      "key": "InvoiceDate",
      "header": "Date",
      "type": "date",
      "format": "yyyy-MM-dd",
      "width": 14
    },
    {
      "key": "Status",
      "header": "Status",
      "type": "string",
      "width": 16,
      "style": "highlightColumn"
    },
    {
      "key": "Amount",
      "header": "Amount",
      "type": "decimal",
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
```

This section maps a source to a rectangular table region.

## Range Section

Ranges are for titles, labels, metadata, formulas, and summary cells.

```json
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
      "value": "{{parameters.generatedDate}}",
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
      "value": "{{parameters.requestedBy}}",
      "style": "detailValue"
    }
  ]
}
```

## Layout As A Shadow Sheet

For API schemas, each output sheet should be created from a shadow sheet template:

```json
{
  "sheetTemplates": {
    "summarySheet": {
      "sections": [
        { "type": "range", "name": "title", "cells": [] },
        { "type": "keyValueBlock", "name": "runInfo", "source": "runInfo", "startCell": "B2" },
        { "type": "table", "name": "costs", "source": "costs", "startCell": "B6" },
        { "type": "table", "name": "income", "source": "income", "startCell": "H6" }
      ]
    }
  },
  "sheets": [
    {
      "name": "Summary",
      "template": "summarySheet"
    }
  ]
}
```

The schema should not be arbitrary free-form JSON. It should be flexible, but strongly shaped and versioned. That gives you validation, predictable rendering, and backward compatibility.

## Multiple Shadow Sheets

A single export can contain many output sheets, each created from a different shadow sheet template.

```json
{
  "sheetTemplates": {
    "summarySheet": {
      "sections": [
        { "type": "range", "name": "summaryHeader", "cells": [] },
        { "type": "table", "name": "summaryTable", "source": "summaryRows", "startCell": "B4" }
      ]
    },
    "detailSheet": {
      "sections": [
        { "type": "table", "name": "detailTable", "source": "detailRows", "startCell": "A1" }
      ]
    },
    "auditSheet": {
      "sections": [
        { "type": "table", "name": "auditTable", "source": "auditRows", "startCell": "A1" }
      ]
    }
  },
  "sheets": [
    { "name": "Summary", "template": "summarySheet" },
    { "name": "Detail", "template": "detailSheet" },
    { "name": "Audit", "template": "auditSheet" }
  ]
}
```

This supports:

- one template per output sheet
- several output sheets from the same template
- one-off inline sheet templates
- sheet-specific data bindings
- sheet-specific protection/page setup
- dynamic sheet creation for repeated entities later

Example of creating several sheets from one template:

```json
{
  "sheets": [
    {
      "name": "London",
      "template": "regionSheet",
      "bindings": {
        "regionSummary": "regions.london.summary",
        "regionRows": "regions.london.rows"
      }
    },
    {
      "name": "United States",
      "template": "regionSheet",
      "bindings": {
        "regionSummary": "regions.us.summary",
        "regionRows": "regions.us.rows"
      }
    }
  ]
}
```

## Relative Placement

Absolute placement is simple and useful:

```json
{ "startCell": "B4" }
```

But relative placement is useful when sources have variable row counts:

```json
{
  "type": "table",
  "name": "detail",
  "source": "detailRows",
  "placeAfter": {
    "section": "summary",
    "spacingRows": 2
  }
}
```

Supported placement modes:

- `startCell`
- `placeAfter`
- `placeRightOf`
- `stackVertical`
- `stackHorizontal`
- `grid`

For v1, prefer `startCell` and simple `placeAfter`.

## Formula Model

Formulas can be fixed or generated using references.

```json
{
  "formulas": [
    {
      "target": "Profit",
      "formula": "=[@Income]-[@Cost]"
    },
    {
      "target": "Margin",
      "formula": "=IF([@Income]=0,0,[@Profit]/[@Income])",
      "format": "0.00%"
    }
  ]
}
```

For table columns, structured references are preferable. For range cells, normal Excel formulas are fine.

## Style Model

Styles should be named and reusable.

```json
{
  "styles": {
    "tableHeader": {
      "font": { "bold": true, "color": "FFFFFF" },
      "fill": "1F4E78",
      "alignment": { "horizontal": "center", "wrap": true },
      "border": {
        "all": { "style": "thin", "color": "808080" }
      }
    },
    "highlightColumn": {
      "border": {
        "left": { "style": "thick", "color": "000000" },
        "right": { "style": "thick", "color": "000000" }
      }
    }
  }
}
```

Style resolution should be layered:

```text
workbook defaults
  -> section/table style
  -> header/row/alternate/total style
  -> column style
  -> conditional style
  -> explicit cell style
```

## Complete Example

```json
{
  "schemaVersion": "1.0",
  "kind": "excelExport",
  "output": {
    "format": "xlsx",
    "fileName": "invoice-export.xlsx"
  },
  "sources": {
    "invoiceRows": {
      "required": true,
      "kind": "table",
      "fields": {
        "InvoiceNo": { "type": "string", "required": true },
        "CustomerName": { "type": "string" },
        "InvoiceDate": { "type": "date" },
        "Status": { "type": "string" },
        "Amount": { "type": "decimal" }
      }
    }
  },
  "sheetTemplates": {
    "invoiceSheet": {
      "gridlines": false,
      "freeze": { "row": 4 },
      "sections": [
        {
          "type": "range",
          "name": "exportDetails",
          "cells": [
            { "cell": "B1", "value": "Invoice Export", "style": "reportTitle" },
            { "cell": "B2", "value": "Generated:", "style": "detailLabel" },
            { "cell": "C2", "value": "{{parameters.generatedDate}}", "format": "yyyy-MM-dd HH:mm", "style": "detailValue" },
            { "cell": "B3", "value": "Requested By:", "style": "detailLabel" },
            { "cell": "C3", "value": "{{parameters.requestedBy}}", "style": "detailValue" }
          ]
        },
        {
          "type": "table",
          "name": "invoiceTable",
          "source": "invoiceRows",
          "startCell": "B4",
          "excelTableName": "InvoiceTable",
          "style": "standardTable",
          "headers": {
            "show": true,
            "filter": true,
            "freeze": true,
            "style": "tableHeader"
          },
          "columns": [
            { "key": "InvoiceNo", "header": "Invoice No.", "type": "string", "width": 16 },
            { "key": "CustomerName", "header": "Customer", "type": "string", "width": 32 },
            { "key": "InvoiceDate", "header": "Date", "type": "date", "format": "yyyy-MM-dd", "width": 14 },
            { "key": "Status", "header": "Status", "type": "string", "width": 16, "style": "highlightColumn" },
            { "key": "Amount", "header": "Amount", "type": "decimal", "format": "#,##0.00", "width": 16 }
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
          }
        }
      ]
    }
  },
  "workbook": {
    "properties": {
      "title": "Invoice Export"
    },
    "encryption": {
      "openPasswordRef": "exportOpenPassword",
      "algorithm": "aes256"
    }
  },
  "sheets": [
    {
      "name": "Invoice Export",
      "template": "invoiceSheet"
    }
  ],
  "styles": {
    "reportTitle": {
      "font": { "size": 14, "bold": true }
    },
    "detailLabel": {
      "font": { "bold": true }
    },
    "detailValue": {
      "font": { "size": 11 }
    },
    "standardTable": {
      "border": {
        "all": { "style": "thin", "color": "B7B7B7" }
      }
    },
    "tableHeader": {
      "font": { "bold": true, "color": "FFFFFF" },
      "fill": "1F4E78"
    },
    "tableRow": {
      "font": { "size": 11 },
      "fill": "FFFFFF"
    },
    "tableAlternateRow": {
      "font": { "size": 11 },
      "fill": "EAF2F8"
    },
    "highlightColumn": {
      "border": {
        "left": { "style": "thick", "color": "000000" },
        "right": { "style": "thick", "color": "000000" }
      }
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
```

## Why Not Arbitrary JSON?

Avoid arbitrary JSON because it makes validation, security, compatibility, and support much harder.

Use a structured schema because:

- the API can validate requests before rendering
- templates can be versioned
- errors can point to schema paths
- clients can generate schemas safely
- server-side templates can be approved and reused
- new features can be added without breaking old schemas

The schema can still be flexible, but it should be a known contract.

## Versioning

Every schema should include:

```json
{
  "schemaVersion": "1.0",
  "kind": "excelExport"
}
```

Versioning rules:

- Additive fields are allowed in minor versions.
- Breaking changes require a new major version.
- Unknown fields should either produce warnings or fail validation, depending on strict mode.
- Stored templates should record the schema version they were created for.

## Recommended V1 Scope

V1 should support:

- workbook metadata
- one or more sheets
- range sections
- key-value blocks
- table sections
- named sources
- parameters
- column selection/order/header replacement
- data types
- formats
- styles
- borders
- alternating rows
- totals
- simple formulas
- absolute placement
- basic relative placement
- sheet/workbook protection
- workbook encryption by password reference
- validation-only API mode

Later versions can add:

- repeating blocks
- template range copying
- pivots
- charts
- images
- diagrams
- slicers/timelines
- richer relative layout
