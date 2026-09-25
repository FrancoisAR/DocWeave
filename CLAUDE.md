# DocWeave — guidance for Claude Code

This file is read automatically at the start of every session. Keep it short and factual; edit it when the project changes.
Formatting and naming rules are enforced by `.editorconfig`, not repeated here.

## What this repository is

**DocWeave** is a class library (targets .NET 8 and .NET 10) for importing and exporting Excel (`.xlsx`) and CSV data. It offers a fluent C# API and a
declarative JSON template API that both compile to the same internal specs and render through one renderer. Built on
`DocumentFormat.OpenXml` 3.x. No Excel Interop, no Excel installation needed.

| Project | Role | Rules apply? |
|---|---|---|
| `DocWeave/` | The library. **Active.** | Yes |
| `DocWeave.Tests/` | xUnit tests. **Active.** | Yes (relaxed, see `.editorconfig`) |

Design notes and plans are in `DocWeave/docs/` (start with `excel-implementation-plan.md` and `template-api-contract-v1.md`).

## Commands

```bash
dotnet build DocWeave.Tests/DocWeave.Tests.csproj             # builds the library too
dotnet test  DocWeave.Tests/DocWeave.Tests.csproj             # full suite (~3 seconds)
dotnet test  DocWeave.Tests/DocWeave.Tests.csproj --filter "FullyQualifiedName~Pivot"
dotnet build DocWeave/DocWeave.csproj -p:EnforceCodeStyleInBuild=true   # also report style rules (they are suggestions by default)
```

After a clean checkout, or after upgrading the OpenXml package, do a **clean rebuild** (`dotnet clean`). A stale incremental build once
produced 38 bogus `InvalidProgramException` failures.

## How the library is organised

```
DocWeave/
  Csv/                 CSV reader, writer, parser, and JSON-template writer/reader
  Excel/               Fluent builders (ExcelWriter, ExcelReader, *Builder), specs (ExcelSpecs.cs), import
    Rendering/         specs -> .xlsx: ExcelWorkbookRenderer (orchestrator) + Pivot, Chart, Table, SheetFeatures, CellWriter, Names
    Templates/         JSON template -> specs: writer, context, layout, styles, data-source checks
      Regions/         One compiler class per region type ("table", "chart", ...)
  WebApi/              Request/response contracts only (no controllers yet)
  ColumnLetters, ColumnSelection, TabularData, TemplateJsonExtensions   shared by the Excel and CSV paths
```

- **Data flow:** fluent builders or a JSON template produce `ExcelWorkbookSpec`; `ExcelWorkbookRenderer.Render` writes it; the saved
  package is then post-processed by `ExcelPackageNormalizer` (part paths for pivot caches).
- **Add a template region type:** write one `ExcelTemplateRegionCompiler` subclass in `Excel/Templates/Regions/` (compile, bounds, whether
  it `OccupiesCells`) and add it to the list in `ExcelTemplateRegions`. Nothing else changes.
- **Sources** of all kinds (`DataTable`, objects, dictionaries) become `TabularData`. Do not add a second copy of that logic.
- **Deliberate difference:** `ColumnSelection` keeps the *requested* order for CSV and *worksheet* order for Excel. A test pins the CSV
  behaviour. Do not "unify" it without asking.
- The renderers are `static` classes on purpose: they are stateless pure functions. That is the accepted exception to "no static classes".

## Conventions (adapted from the project rules for a library)

**Comments**
- Every public type and member gets `<summary>`. Add `<param>`, `<returns>` and `<exception>` when they tell the reader something
  (a meaningful exception a caller can act on, not every `ArgumentNullException`). If you document one parameter, document all of them
  (the compiler warns otherwise).
- Inline `//` comments explain *why*: non-obvious logic, OpenXML rules, complex LINQ, algorithm steps. Never restate the code.

**Errors**
- Validate arguments with `ArgumentNullException.ThrowIfNull` / `ArgumentException.ThrowIfNullOrWhiteSpace`.
- Exceptions are for exceptional input, not flow control. Use `try/catch` only around code that can realistically fail.
- Return empty collections, never `null`, for collection results.

**Style**
- Prefer LINQ for filtering and projection; split deep queries into named steps and comment the tricky ones. Use a loop where it is
  clearer or where allocations matter (large exports).
- Use switch expressions instead of long `if`/`else` chains. No magic numbers: name Excel limits and OpenXML ids.
- No `#region` (nothing enforces this automatically, so check it yourself). No static mutable state. No credentials in code.
- Keep lines to about 120 characters (advisory; embedded JSON in tests is exempt).

**Async (library rule)**
- All new I/O entry points (file, stream, network) must offer an `Async` variant taking a `CancellationToken`. Do not block on async.
  Use `ConfigureAwait(false)` (enforced by CA2007 in `DocWeave/`). The library has no async code yet; the current API is synchronous.

**Logging:** the library exposes progress and audit callbacks (`OperationTelemetry.cs`) rather than depending on a logging framework.
Do not add a logging package to `DocWeave`.

## Tests

- xUnit, Arrange–Act–Assert. No test may depend on an external service.
- The numbered scenario tests (`_NN_Some_description`) each write `TestOutput/Excel/<test name>.xlsx`. **The number must be unique across
  all Excel test classes**, and renaming a test renames its output file.
- Use `Xlsx.AssertValid(bytes)` for schema validity (it prints what is wrong) and the other `Xlsx.*` helpers to inspect a workbook.
  Shared sample data is in `SampleData`. Do not copy helpers into a test class.
- Import tests read **committed source workbooks** from `TestData/Excel/` (through `ExcelTestData.Load`), like the CSV tests read
  `TestData/Csv/`. They must not generate their input into `TestOutput/`. Describe any new fixture in `TestData/Excel/README.md`.
  A test that exports and then imports its own output (a round trip) may still write to `TestOutput/`.
- `TestOutput/` is git-ignored: the tests regenerate it on every run (with new random ids each time). The files are handy to open in Excel,
  but they are not part of the repository, so do not expect to find them in a fresh checkout until the tests have run.
- Passing `OpenXmlValidator` does **not** prove Excel opens the file cleanly. For pivots, charts and protection, tell the user which
  output file to open in Excel; do not claim it works.

## Refactoring safely

Refactors must not change behaviour. Before starting: run the tests and hash every generated workbook after normalising the random
relationship ids (`R` + 16 hex digits). After: same tests, same hashes. Write characterization tests first for anything the suite does not
cover. Show the user the numbers, not just "tests pass".

## Future Web API project (`DocWeave.WebApi`, not built yet)

These rules apply in full there, and only there:
controllers orchestrate services and return responses, with no business logic; business logic lives in injected services; data access goes
through repositories; use DTOs for input and output and never expose entities; log with `ILogger<T>` at the service layer, not the
controller; integration tests for controllers; no inline SQL; no hard-coded connection strings.

## Working agreements

- Do not commit or push unless asked.
- Report results honestly with the actual numbers, including failures, skipped checks, and anything you could not verify.
- Say so when a request conflicts with a documented decision here or in `DocWeave/docs/`, rather than silently picking one.
