# DocWeave Roadmap

What DocWeave does not do yet, and what could come next. Nothing here is scheduled or promised.

- What works today: [getting-started.md](getting-started.md), [pre-release-functionality-guide.md](pre-release-functionality-guide.md), [current-functionality.md](current-functionality.md).
- Where an item comes from is marked **[design]** (already described in a design document) or **[suggestion]** (not in any existing document, so it is a proposal only).
- Size is a rough guess for one developer: **S** = days, **M** = 1-3 weeks, **L** = a month or more.
- "Verified" notes were checked against the code on the day this page was written.

## 1. Where things stand

| Area | State |
|---|---|
| Excel `.xlsx` export (fluent API) | Working: tables, formulas, totals, styles, charts, pivot tables, validation, comments, protection, repeaters |
| Excel `.xlsx` import | Working: dictionaries, `DataTable`, rows, typed objects (classes, structs, positional records; attributes or code mapping), and row-by-row streaming (`EnumerateRows`). Worksheet data only |
| CSV import and export | Working, including streaming reads |
| JSON templates (Excel and CSV) | Working, with the gaps in section 2 |
| Web API | Request and response contracts only. No host project |
| Word, PowerPoint, PDF, HTML | Not started |

## 2. Template system

The template contract is [template-api-contract-v1.md](template-api-contract-v1.md). Sixteen region types are implemented:
`cell`, `formulaCell`, `rangeStyle`, `mergeCells`, `comment`, `column`, `row`, `freezePane`, `autoFilter`, `sheetProtection`,
`dataValidation`, `table`, `repeater`, `summary`, `chart`, `pivotTable`. An in-memory stored-template provider exists.

**Gaps against the contract [design]:**

| # | Item | Size | Notes |
|---|---|---|---|
| T1 | **Schema version checking** | S | Verified: the code never reads `schemaVersion`. The contract says a template must declare it and unknown or future versions must be rejected clearly. Do this before other template changes, because everything else depends on being able to version the format. |
| T2 | **Richer validation results** | M | Verified: `Validate()` reports every problem with the path `$` (the whole template), and the result has no warnings list. The contract wants a precise JSON path per error (`$.sheetTemplates.x.regions[2]...`), a stable error code, and separate **warnings** (for example a sheet name that will be truncated). The seven validation stages in the contract (syntax, schema version, contract, source binding, layout, renderer capability, security) are not reported as stages. |
| T3 | **Secrets by reference** | S | Verified: nothing in the code handles `secrets` or `passwordRef`. Passwords must be supplied at run time and never sit in template text. Needed before encryption (S1) and before protection passwords can be used safely in stored templates. |
| T4 | **Formula policy** | M | Verified: no formula policy exists. The contract says templates submitted over an API must comply with one, since a formula is a place a template can carry active content (for example external links). Needs a design decision: an allow-list of functions, or blocking external references only. |
| T5 | **Persistent, versioned stored templates** | M | Only an in-memory provider exists. The contract asks for versioned, approved templates with allowed output formats recorded per template. Needs a store abstraction (file, database) and pinning of templates to a schema version. |
| T6 | **Template workbook support** | L | Open an existing `.xlsx`/`.xltx`, copy a sheet, fill named cells, ranges and template tables while keeping its formatting (Phase 8 in [excel-implementation-plan.md](excel-implementation-plan.md)). Out of scope for V1 of the contract. The legacy library did this, and many Excel reporting users expect it. |
| T7 | **Multiple chart series** | M | Verified: a chart has one series, named after the chart title. The contract lists multi-series as out of scope for V1. Most real charts need this. |
| T8 | **Nested and grouped regions** | M-L | "Nested mini-tables" and data groups are in the design docs. Not started. |
| T9 | **Template dry run** | S | A preview of region bounds already exists. Extend it to a full dry-run report (row counts, sheets, warnings) without rendering. |
| T10 | **Published JSON Schema for templates** | M | [suggestion] So editors can autocomplete and validate templates. Cheap once T1 and T2 are done. |

## 3. Security and delivery

| # | Item | Size | Notes |
|---|---|---|---|
| S1 | **Workbook encryption (password to open)** | M | [design] Not implemented, and different from protection. Needs a pluggable encryption service, since the Open XML SDK does not encrypt packages itself. Depends on T3. |
| S2 | **Import of encrypted workbooks** | M | [design] The API contract has a password field, but the readers do not use it. |
| S3 | **CSV formula-injection coverage** | S | [suggestion] Formula-like text is neutralised by default. Document exactly which leading characters are covered and add a test for each. |
| S4 | **Delivery adapters** | M each | [design] The design docs list file share, blob storage, email, Teams, SharePoint/OneDrive, webhook, queue. Today output goes to bytes, a file path or a stream. [suggestion] Ship the adapter interface only, and put each adapter in its own package so the core stays free of cloud dependencies. |

## 4. Web API host

Verified: the contract types (`DocWeaveExportRequest`, `DocWeaveOperationAcceptedResponse` and so on) exist, but there is no host.
`CLAUDE.md` plans a separate `DocWeave.WebApi` project.

| # | Item | Size | Notes |
|---|---|---|---|
| W1 | **Host project with the endpoints in the contract** | M | [design] `POST /exports`, `GET /exports/{id}`, `POST /exports/{id}/cancel`, `GET /templates`, `POST /templates/validate`, `POST /templates/{id}/render` |
| W2 | **Background operations** | M | [design] Status, progress and cancellation for long exports. The library already reports progress and accepts a cancellation token, but it has no async API (Verified: no `async` code and no `Task` return types). |
| W3 | **Authentication and per-template authorisation** | M | [suggestion] The contract mentions "caller permission" only. |

## 5. Excel export: features not yet built

Verified absent from the code: conditional formatting, images and logos, hyperlinks, sparklines, subtotals and grouping, page setup and
print areas, defined (named) ranges, row/column outlining, tab colours.

| # | Item | Size | Notes |
|---|---|---|---|
| E1 | **Conditional formatting** | M | [design] Colour scales, data bars, rules such as "cell contains". A common request. Table styling already supports conditional *styles* evaluated in C#, which is different. |
| E2 | **Images and logos** | S-M | [design] |
| E3 | **Hyperlinks** | S | [design] Cell and range links, internal and external. |
| E4 | **Page setup, print areas, tab colours, gridlines** | M | [design] Orientation, print areas, headers and footers, hidden gridlines, tab colours. Needed for reports that get printed or converted to PDF. |
| E5 | **Defined names** | S | [suggestion] Also needed for template workbook support (T6). |
| E6 | **Subtotals and data groups** | M | [design] Row grouping with outline levels. |
| E7 | **Sparklines, slicers, shapes, connectors** | M each | [design] Lowest priority. |
| E8 | **Pivot tables: charts from pivots, calculated fields, grouping** | L | [design] Multiple row, column and value fields and report filters work. Pivot output has only been checked against the Open XML schema. It should be opened and accepted in real Excel for each important report shape (Q1). |
| E9 | **Streaming writer for large exports** | L | [design] Verified: Excel export builds the whole sheet in memory. CSV already streams. Uses the SDK's `OpenXmlWriter` and changes the renderer's structure. Excel's own row limit (1,048,576) is the ceiling either way. |
| E10 | **Cancellation during Excel rendering** | S | Cancellation is checked before rendering starts, not for every row or cell. |
| E11 | **Real Excel dates** | S | Verified: dates are written as ISO text, not as Excel dates. [suggestion] Write real date cells with a date format, with an option to keep text. Likely to surprise users, so treat it as a priority. |

## 6. Import

| # | Item | Size | Notes |
|---|---|---|---|
| I1 | **Typed object mapping** | S remaining | **Done:** `ReadObjects<T>()` and `EnumerateObjects<T>()` for Excel and CSV; matching by header, `[DocWeaveColumn]` name or ordinal, or in code with `ReadObjects<T>(map => ...)`; classes, structs and positional records (built through the constructor, with parameter defaults); the common types; problems collected in an error list. **Still to do:** per-property date and number formats (see I2), and reading nested or list-valued properties. |
| I2 | **Custom converters and conversion options** | M | [design] Per-property converters, and options such as culture and date formats (see "Import Conversion Options" and "Import Error Handling" in [fluent-api-capabilities.md](fluent-api-capabilities.md)). Object mapping already collects errors instead of throwing, but `ReadRows` and `ReadDictionaries` still return text, and a malformed CSV row can still throw. |
| I3 | **Streaming Excel import** | done | `EnumerateRows()` reads a sheet row by row (measured on 200,000 rows: about 8 MB of extra memory against about 180 MB for `ReadRows()`). **Still to do:** the shared string table is still read into memory once; a non-seekable source must be buffered by the caller; progress and audit events are not sent by the streaming methods. |
| I4 | **Multi-sheet and batch import** | M | [design] Several sheets or several files in one call, with per-source passwords. |
| I5 | **Import into a `DataSet`** | S | [design] |
| I6 | **Richer read of an existing workbook** | L | Import ignores styles, comments, charts and protection. [suggestion] Probably not worth doing until someone asks. |
| I7 | **Async API** | M | [suggestion] `Task` and `IAsyncEnumerable` variants for reading and writing, using `ConfigureAwait(false)` as `CLAUDE.md` requires. Pairs with W2. |

## 7. More output formats

Design basis: the shared core (source data, table or document spec, renderer) described in
[fluent-api-capabilities.md](fluent-api-capabilities.md) and [interface-technical-plan.md](interface-technical-plan.md).

| # | Item | Size | Notes |
|---|---|---|---|
| F1 | **Split into packages** | M | [suggestion] A core package plus one per format (`DocWeave.Excel`, `DocWeave.Csv`, later `DocWeave.Word`), so Excel users do not take on Word dependencies. Do this **before** adding a second Office format and before 1.0, because it changes package names. |
| F2 | **Word (`.docx`)** | L | [design] Report sections, summary tables, repeated blocks, from data and templates. Uses the Open XML SDK already in use. |
| F3 | **PowerPoint (`.pptx`)** | L | [design] Slide templates, charts, KPI panels, repeated slide groups. The Open XML SDK supports it, but charts inside slides are harder than in Excel. |
| F4 | **JSON, JSON Lines, XML, TSV** | S each | [design] Cheap once the core is separated from the Excel renderer. |
| F5 | **PDF and HTML** | L | [design] Likely by converting another output, or a dedicated renderer. [suggestion] Decide early whether PDF is rendered directly or converted, since reports are often wanted as PDF. |
| F6 | **OpenDocument (`.ods`, `.odt`)** | M | [suggestion] Only if there is demand. |

## 8. Quality, release and project health

| # | Item | Size | Notes |
|---|---|---|---|
| Q1 | **Acceptance testing in real Excel** | S-M | Only Open XML schema checks exist. Open each report shape, especially pivots and charts, in Excel and record the result. |
| Q2 | **Raise `CS1591` to a warning and clear nullability warnings** | S | The documentation backlog is zero. Nullability warnings remain (CS8602, CS8604, CS8625). |
| Q3 | **Public API review before 1.0** | S | Decide what is public. The `DocWeave*` contract types and the builders are all still preview. |
| Q4 | **Continuous integration** | S | [suggestion] Build and test on `net8.0` and `net10.0`, pack, and publish preview packages. |
| Q5 | **Performance baselines** | S | [suggestion] Volume tests exist but run only when an environment variable is set and record no baseline, so regressions go unnoticed. |
| Q6 | **Guard the documented examples** | S | [suggestion] The getting-started examples compile today, but nothing stops them breaking. Compile them in the test project. |
| Q7 | **Repository URL and a changelog** | S | Package metadata still has a placeholder repository URL. |

## 9. Suggested order

1. **Foundations:** T1 schema version, T3 secrets by reference, E11 real Excel dates, Q1 real-Excel acceptance, Q2 and Q3 cleanup. (I1 typed import and I3 streaming import are done; see section 6.)
2. **Make templates production-ready:** T2 validation results, T4 formula policy, T5 stored templates, T7 multiple chart series.
3. **Round out Excel:** E1 conditional formatting, E2 images, E3 hyperlinks, E4 page setup, then I2 (converters and formats).
4. **Service use:** W1 host, W2 background operations, S1 encryption.
5. **Scale:** E9 streaming export. (I3 streaming import is done.)
6. **New formats:** F1 package split, then F4 (the cheap formats), then F2 Word, F3 PowerPoint, F5 PDF.

T6 (template workbooks) could move up this list if migrating from the legacy library, or users who design reports in Excel first, are the
main use case. That is a product decision, not a technical one.
