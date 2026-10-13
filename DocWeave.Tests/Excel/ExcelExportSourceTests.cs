using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Exporting from different kinds of source: object lists, DataTables, DataSets, dictionaries, and mapped columns.
/// </summary>
public sealed class ExcelExportSourceTests
{
    [Fact]
    public void _03_Export_customer_invoice_report()
    {
        var invoices = CreateInvoiceTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Invoices", sheet => sheet
                .AddCell("B2", "Customer Invoice Report")
                .AddCell("B3", $"Generated {DateTime.Today:yyyy-MM-dd}")
                .AddTable(invoices)
                .StartAt("B5")
                .HeaderAlias("InvoiceNo", "Invoice Number")
                .HeaderAlias("CustomerName", "Customer")
                .HeaderAlias("InvoiceDate", "Invoice Date"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_03_Export_customer_invoice_report), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Invoices")
            .StartAt("B5")
            .ReadDataTable();

        Assert.Equal(6, imported.Rows.Count);
        Assert.Equal("Invoice Number", imported.Columns[0].ColumnName);
        Assert.Equal("INV-1001", imported.Rows[0]["Invoice Number"]);
    }

    [Fact]
    public void _04_Export_regional_summary_and_audit_workbook()
    {
        var regions = new[]
        {
            new RegionalSummary("London", 125000m, 210000m, 85000m),
            new RegionalSummary("United States", 190000m, 325000m, 135000m),
            new RegionalSummary("Europe", 155000m, 260000m, 105000m),
            new RegionalSummary("Asia", 175000m, 295000m, 120000m)
        };

        var auditRows = new[]
        {
            new AuditLine("system", "Report", "Created", DateTime.Today.AddMinutes(8)),
            new AuditLine("analyst", "Regions", "Reviewed", DateTime.Today.AddMinutes(15)),
            new AuditLine("approver", "Workbook", "Approved", DateTime.Today.AddMinutes(42))
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Summary", sheet => sheet
                .AddCell("A1", "Regional Summary")
                .AddTable(regions)
                .StartAt("A3"))
            .AddSheet("Audit", sheet => sheet
                .AddCell("A1", "Audit Trail")
                .AddTable(auditRows)
                .StartAt("A3"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_04_Export_regional_summary_and_audit_workbook), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Summary")
            .StartAt("A3")
            .Columns(columns => columns.IncludeHeaders("Region", "Profit"))
            .ReadDictionaries();

        Assert.Equal(4, imported.Count);
        Assert.Equal("London", imported[0]["Region"]);
        Assert.Equal("85000", imported[0]["Profit"]);
    }

    [Fact]
    public void _05_Export_mixed_sources_on_one_sheet()
    {
        var incomeRows = new[]
        {
            new Dictionary<string, object?> { ["Region"] = "London", ["Income"] = 210000m, ["Currency"] = "GBP" },
            new Dictionary<string, object?> { ["Region"] = "United States", ["Income"] = 325000m, ["Currency"] = "USD" },
            new Dictionary<string, object?> { ["Region"] = "Europe", ["Income"] = 260000m, ["Currency"] = "EUR" }
        };

        var costs = new DataTable("Costs");
        costs.Columns.Add("Region", typeof(string));
        costs.Columns.Add("Costs", typeof(decimal));
        costs.Columns.Add("Currency", typeof(string));
        costs.Rows.Add("London", 125000m, "GBP");
        costs.Rows.Add("United States", 190000m, "USD");
        costs.Rows.Add("Europe", 155000m, "EUR");

        var bytes = ExcelWriter.Create()
            .AddSheet("Financials", sheet =>
            {
                sheet.AddCell("A1", "Mixed Source Export");
                sheet.AddCell("A3", "Income");
                sheet.AddRows(incomeRows)
                    .StartAt("A4")
                    .HeaderAlias("Region", "Income Region");

                sheet.AddCell("E3", "Costs");
                sheet.AddTable(costs)
                    .StartAt("E4")
                    .HeaderAlias("Region", "Cost Region");
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_05_Export_mixed_sources_on_one_sheet), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Financials")
            .StartAt("A4")
            .Columns(columns => columns.IncludeHeaders("Income Region", "Income"))
            .ReadDictionaries();

        Assert.Equal(3, imported.Count);
        Assert.Equal("United States", imported[1]["Income Region"]);
    }

    [Fact]
    public void _06_Export_multiple_dataset_tables_and_typed_sources_to_different_sheets()
    {
        var dataSet = CreateBusinessDataSet();
        var activeDirectoryRows = new[]
        {
            new ActiveDirectoryExportRow("AD-1001", "Alice West", "Risk", "London"),
            new ActiveDirectoryExportRow("AD-1002", "Ben Carter", "Finance", "New York"),
            new ActiveDirectoryExportRow("AD-1003", "Claire Ito", "Technology", "Tokyo")
        };
        var dictionaryRows = new[]
        {
            new Dictionary<string, object?> { ["Metric"] = "Generated Sheets", ["Value"] = 4, ["Status"] = "Complete" },
            new Dictionary<string, object?> { ["Metric"] = "Primary Source", ["Value"] = "DataSet", ["Status"] = "Complete" },
            new Dictionary<string, object?> { ["Metric"] = "Additional Types", ["Value"] = "Objects + Dictionaries", ["Status"] = "Complete" }
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Regions", sheet => sheet
                .AddCell("A1", "Regions From DataSet")
                .AddTable(dataSet.Tables["Regions"]!)
                .StartAt("A3"))
            .AddSheet("Invoices", sheet => sheet
                .AddCell("B2", "Invoices From DataSet")
                .AddTable(dataSet.Tables["Invoices"]!)
                .StartAt("B4")
                .HeaderAlias("InvoiceNo", "Invoice Number")
                .HeaderAlias("CustomerName", "Customer"))
            .AddSheet("Directory", sheet => sheet
                .AddCell("A1", "Typed Object Source")
                .AddTable(activeDirectoryRows)
                .StartAt("A3")
                .HeaderAlias("DisplayName", "Display Name"))
            .AddSheet("Run Summary", sheet => sheet
                .AddCell("A1", "Dictionary Source")
                .AddRows(dictionaryRows)
                .StartAt("A3"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_06_Export_multiple_dataset_tables_and_typed_sources_to_different_sheets), bytes);

        var regions = ExcelReader.FromBytes(bytes)
            .Sheet("Regions")
            .StartAt("A3")
            .ReadDataTable();
        var invoices = ExcelReader.FromBytes(bytes)
            .Sheet("Invoices")
            .StartAt("B4")
            .Columns(columns => columns.IncludeHeaders("Invoice Number", "Customer"))
            .ReadDataTable();
        var directory = ExcelReader.FromBytes(bytes)
            .Sheet("Directory")
            .StartAt("A3")
            .Columns(columns => columns.IncludeHeaders("Display Name", "Department"))
            .ReadDictionaries();
        var summary = ExcelReader.FromBytes(bytes)
            .Sheet("Run Summary")
            .StartAt("A3")
            .ReadDictionaries();

        Assert.Equal(4, regions.Rows.Count);
        Assert.Equal(6, invoices.Rows.Count);
        Assert.Equal("INV-1001", invoices.Rows[0]["Invoice Number"]);
        Assert.Equal("Alice West", directory[0]["Display Name"]);
        Assert.Equal("Generated Sheets", summary[0]["Metric"]);
    }

    [Fact]
    public void _07_Export_datatable_with_columns_mapped_to_specific_excel_locations()
    {
        var table = new DataTable("MappedCustomers");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Department", typeof(string));
        table.Columns.Add("Office", typeof(string));
        table.Columns.Add("Status", typeof(string));
        table.Rows.Add(1, "Alice West", "Risk", "London", "Active");
        table.Rows.Add(2, "Ben Carter", "Finance", "New York", "Active");
        table.Rows.Add(3, "Claire Ito", "Technology", "Tokyo", "Pending");

        var bytes = ExcelWriter.Create()
            .AddSheet("Mapped Columns", sheet => sheet
                .AddCell("B2", "Column Location Export")
                .AddTable(table)
                .StartAt("B4")
                .ColumnLocation("ID", "B")
                .ColumnLocation("Name", "C")
                .ColumnLocation("Department", "E")
                .ColumnLocation("Office", "G")
                .ColumnLocation("Status", "H")
                .HeaderAlias("ID", "Employee ID")
                .HeaderAlias("Name", "Employee Name"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_07_Export_datatable_with_columns_mapped_to_specific_excel_locations), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Mapped Columns")
            .StartAt("B4")
            .Columns(columns => columns.IncludeHeaders("Employee ID", "Employee Name", "Department", "Office", "Status"))
            .ReadDataTable();

        Assert.Equal(3, imported.Rows.Count);
        Assert.Equal("Employee ID", imported.Columns[0].ColumnName);
        Assert.Equal("Employee Name", imported.Columns[1].ColumnName);
        Assert.Equal("1", imported.Rows[0]["Employee ID"]);
        Assert.Equal("Alice West", imported.Rows[0]["Employee Name"]);
        Assert.Equal("Risk", imported.Rows[0]["Department"]);
        Assert.Equal("London", imported.Rows[0]["Office"]);
        Assert.Equal("Active", imported.Rows[0]["Status"]);
    }

    private sealed record ActiveDirectoryExportRow(string Id, string DisplayName, string Department, string Office);
}
