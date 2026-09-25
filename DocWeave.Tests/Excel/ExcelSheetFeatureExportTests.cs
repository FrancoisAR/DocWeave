using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Sheet features: data validation, freeze panes, auto filter, protection, native tables and chart labels.
/// </summary>
public sealed class ExcelSheetFeatureExportTests
{
    [Fact]
    public void _42_Export_data_validation_dropdowns_and_numeric_limits()
    {
        var table = new DataTable("ValidationRows");
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Status", typeof(string));
        table.Columns.Add("Score", typeof(int));
        table.Rows.Add("Alice", "Open", 75);
        table.Rows.Add("Bob", "Review", 88);

        var bytes = ExcelWriter.Create()
            .AddSheet("Validation", sheet =>
            {
                sheet.AddCell("A1", "Data Validation Export", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(table)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"));

                sheet.AddDataValidation("B4:B50", validation => validation
                    .List("Open", "Closed", "Review")
                    .InputMessage("Status", "Choose a status.")
                    .ErrorMessage("Invalid status", "Use Open, Closed, or Review."));

                sheet.AddDataValidation("C4:C50", validation => validation
                    .WholeNumberBetween(0, 100)
                    .InputMessage("Score", "Enter a whole number from 0 to 100."));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_42_Export_data_validation_dropdowns_and_numeric_limits), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(2, CountDataValidations(bytes, "Validation"));
        Assert.True(CountStyledCells(bytes, "Validation") > 8);
    }

    [Fact]
    public void _43_Export_freeze_panes_auto_filter_progress_audit_and_import_result()
    {
        var progress = new TestProgress<DocWeaveOperationProgress>();
        var auditEvents = new List<DocWeaveAuditEvent>();
        var rows = new[]
        {
            new { ID = 1, Name = "Alice", Region = "London", Amount = 125.50m },
            new { ID = 2, Name = "Bob", Region = "United States", Amount = 300.00m },
            new { ID = 3, Name = "Claire", Region = "Europe", Amount = 210.25m }
        };

        var bytes = ExcelWriter.Create()
            .WithProgress(progress)
            .Audit(auditEvents.Add)
            .AddSheet("Filtered", sheet =>
            {
                sheet.AddCell("A1", "Filtered export", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(rows)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                    .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000"));
                sheet.FreezePanes(rows: 3, columns: 1);
                sheet.AutoFilter("A3:D6");
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_43_Export_freeze_panes_auto_filter_progress_audit_and_import_result), bytes);

        var pane = GetFreezePane(bytes, "Filtered");
        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Filtered")
            .StartAt("A3")
            .ReadDictionaryResult();

        Xlsx.AssertValid(bytes);
        Assert.Equal(3D, pane.VerticalSplit!.Value);
        Assert.Equal(1D, pane.HorizontalSplit!.Value);
        Assert.Equal("B4", pane.TopLeftCell!.Value);
        Assert.Equal("A3:D6", GetAutoFilterReference(bytes, "Filtered"));
        Assert.Contains(progress.Reports, report => report.Operation == "ExcelExport" && report.Stage == "Completed");
        Assert.Contains(auditEvents, audit => audit.Operation == "ExcelExport" && audit.Message == "Workbook rendered");
        Assert.Equal(["ID", "Name", "Region", "Amount"], imported.Headers);
        Assert.Equal(4, imported.RowCount);
        Assert.Equal("Alice", imported.Rows[0]["Name"]);
    }

    [Fact]
    public void _45_Export_workbook_protection_native_table_and_chart_labels()
    {
        var rows = new[]
        {
            new { Month = "Jan", Income = 125000m, Costs = 72000m },
            new { Month = "Feb", Income = 138000m, Costs = 81000m },
            new { Month = "Mar", Income = 142000m, Costs = 79000m }
        };

        var bytes = ExcelWriter.Create()
            .ProtectStructure("workbook")
            .AddSheet("Protected Report", sheet =>
            {
                sheet.AddCell("A1", "Protected Native Table Export", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(rows)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                    .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000"))
                    .AsNativeTable("MonthlyResults", "TableStyleMedium4");
                sheet.AddChart("Income by Month", chart => chart
                    .Type(ExcelChartType.Column)
                    .DataSource("Protected Report", "$A$4:$A$6", "$B$4:$B$6")
                    .Position("F3", "M18")
                    .ShowDataLabels()
                    .AxisTitles("Month", "Income")
                    .LegendPosition(ExcelChartLegendPosition.Bottom));
                sheet.Protect("sheet", allowAutoFilter: true);
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_45_Export_workbook_protection_native_table_and_chart_labels), bytes);

        Xlsx.AssertValid(bytes);
        Assert.True(HasWorkbookProtection(bytes));
        Assert.True(HasSheetProtection(bytes, "Protected Report"));
        Assert.Equal(1, CountNativeTables(bytes, "Protected Report"));
        Assert.Equal(1, CountChartDataLabels(bytes));
        Assert.Equal(2, CountChartAxisTitles(bytes));
    }
}
