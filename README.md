# DocWeave

DocWeave is a .NET file-processing engine that lets developers import, export, transform, and template spreadsheet-style data (Excel, CSV, and more to come) through a fluent, modern API, or from JSON templates. Excel support is built on the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK).

> **Status: preview.** The public API can still change before 1.0.

## Install
(not yet release as Nuget)
```
dotnet add package DocWeave --prerelease
```

This library is published in two target frameworks—netstandard2.0 and .NET 10—to support two distinct runtime families. The netstandard2.0 build ensures compatibility with .NET Framework applications, while the .NET 10 build is intended for modern .NET (6 and above).

Although the current implementation is simple enough to work across both platforms without friction, future development may introduce features, APIs, or technologies that are not uniformly supported. Maintaining separate builds helps prevent “technology drift” from forcing workarounds, limiting design choices, or constraining the evolution of the library. By keeping the platforms split now, the library can grow naturally without being held back by legacy requirements.

## What it does

- **Excel export:** styled tables, formula columns, totals rows, merged cells, freeze panes, auto filters,
  comments, data validation, sheet and workbook protection, charts, and pivot-table generation
  for tested row, column, value and report-filter scenarios.
- **Excel import:** read a sheet into dictionaries or a `DataTable`, choosing columns by letter or header.
- **CSV export and import:** delimiter, quote, header and column options, streaming reads for large files,
  and typed import through a template. Formula-like text is neutralised on export by default.
- **JSON templates:** describe a layout once, supply the data separately. Useful when a Web API should
  accept data without exposing the whole builder surface.
- **Progress, cancellation and audit hooks** for long exports and imports.

## Quick start

### Export to Excel

```csharp
using DocWeave.Excel;

var invoices = new[]
{
    new { Id = 1, Customer = "Northwind", Net = 100m, Vat = 20m },
    new { Id = 2, Customer = "Contoso", Net = 250m, Vat = 50m }
};

ExcelWriter.Create()
    .AddSheet("Invoices", sheet =>
    {
        sheet.AddTable(invoices)
            .StartAt("A1")
            .HeaderAlias("Id", "Invoice No")
            .FormulaColumn("Gross", ctx => $"C{ctx.WorksheetRow}+D{ctx.WorksheetRow}")
            .HeaderStyle(s => s.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF"))
            .AddTotals.Bottom("Totals");
    })
    .SaveAs("invoices.xlsx");
```

Use `ToBytes()` instead of `SaveAs(path)` to get the file in memory.

### Read from Excel

```csharp
var rows = ExcelReader.FromFile("invoices.xlsx")
    .Sheet("Invoices")
    .StartAt("A1")
    .Columns(c => c.IncludeHeaders("Invoice No", "Customer", "Gross"))
    .ReadDictionaries();
```

### CSV

```csharp
using DocWeave.Csv;

string csv = CsvWriter.FromObjects(invoices)
    .DelimitedBy(";")
    .IncludeHeaders()
    .ToText();

var imported = CsvReader.FromText(csv)
    .DelimitedBy(";")
    .HasHeaderRow()
    .ReadDictionaries();
```

## Documentation

More examples (pivot tables, charts, repeaters, templates, progress and audit, Web API use) are in the
[getting started guide](Documentation/getting-started.md). Design notes are in [`Documentation`](Documentation).

## License

[Apache License 2.0](LICENSE)
