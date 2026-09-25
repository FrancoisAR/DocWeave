using System.Data;
using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

public sealed class ExcelImportExportTests
{
    [Fact]
    public void _01_Export_and_import_datatable_with_header_aliases()
    {
        var table = new DataTable("Invoices");
        table.Columns.Add("InvoiceNo", typeof(string));
        table.Columns.Add("CustomerName", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add("INV-001", "Alice", 125.50m);
        table.Rows.Add("INV-002", "Bob", 300m);

        var bytes = ExcelWriter.Create()
            .AddSheet("Invoices", sheet => sheet
                .AddCell("B2", "Invoice Export")
                .AddTable(table)
                .StartAt("B4")
                .HeaderAlias("InvoiceNo", "Invoice Number")
                .HeaderAlias("CustomerName", "Customer"))
            .ToBytes();
        ExcelTestOutput.SaveWorkbook(nameof(_01_Export_and_import_datatable_with_header_aliases), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Invoices")
            .StartAt("B4")
            .Columns(columns => columns.IncludeHeaders("Invoice Number", "Customer", "Amount"))
            .ReadDataTable();

        Assert.Equal(2, imported.Rows.Count);
        Assert.Equal("Invoice Number", imported.Columns[0].ColumnName);
        Assert.Equal("INV-001", imported.Rows[0]["Invoice Number"]);
        Assert.Equal("Alice", imported.Rows[0]["Customer"]);
        Assert.Equal("125.50", imported.Rows[0]["Amount"]);
    }

    [Fact]
    public void _02_Export_supports_object_collections_and_import_column_letters()
    {
        var rows = new[]
        {
            new RegionExportRow("London", 1200m, 1500m),
            new RegionExportRow("US", 2200m, 2750m)
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Regions", sheet => sheet
                .AddTable(rows)
                .StartAt("A1"))
            .ToBytes();
        ExcelTestOutput.SaveWorkbook(nameof(_02_Export_supports_object_collections_and_import_column_letters), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .SheetAt(0)
            .Columns(columns => columns.IncludeLetters("A", "C"))
            .ReadDictionaries();

        Assert.Equal(2, imported.Count);
        Assert.Equal(["Region", "Income"], imported[0].Keys);
        Assert.Equal("London", imported[0]["Region"]);
        Assert.Equal("1500", imported[0]["Income"]);
    }

    private sealed record RegionExportRow(string Region, decimal Costs, decimal Income);
}
