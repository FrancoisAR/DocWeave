using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Charts, from the same sheet or another sheet, in each supported chart type.
/// </summary>
public sealed class ExcelChartExportTests
{
    [Fact]
    public void _15_Export_charts_from_same_sheet_and_another_sheet_datasources()
    {
        var table = new DataTable("ChartData");
        table.Columns.Add("Month", typeof(string));
        table.Columns.Add("Income", typeof(decimal));
        table.Columns.Add("Costs", typeof(decimal));
        table.Rows.Add("Jan", 120000m, 82000m);
        table.Rows.Add("Feb", 135000m, 91000m);
        table.Rows.Add("Mar", 128000m, 97000m);
        table.Rows.Add("Apr", 155000m, 101000m);
        table.Rows.Add("May", 162000m, 112000m);
        table.Rows.Add("Jun", 171000m, 118000m);

        var bytes = ExcelWriter.Create()
            .AddSheet("ChartData", sheet =>
            {
                sheet.AddCell("A1", "Chart Source Data", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(table)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                    .AlternateRowStyle(style => style.BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                    .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000"));

                sheet.AddChart("Costs By Month", chart => chart
                    .Type(ExcelChartType.Line)
                    .DataSource("ChartData", "$A$4:$A$9", "$C$4:$C$9")
                    .Position("E3", "L18")
                    .Style(13)
                    .SeriesColor("#C65911"));
            })
            .AddSheet("Dashboard", sheet => sheet
                .AddCell("A1", "Chart Dashboard", style => style.Bold().FontSize(16).FontColor("#385723"))
                .AddCell("A2", "Chart reads from the ChartData worksheet.", style => style.Italic().FontColor("#666666"))
                .AddChart("Income By Month", chart => chart
                    .Type(ExcelChartType.Column)
                    .DataSource("ChartData", "$A$4:$A$9", "$B$4:$B$9")
                    .Position("A4", "J22")
                    .Style(10)
                    .SeriesColor("#4472C4")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_15_Export_charts_from_same_sheet_and_another_sheet_datasources), bytes);

        Assert.Equal(2, CountChartParts(bytes));
        Assert.Equal(2, CountFormulaCells(bytes, "ChartData"));
    }

    [Fact]
    public void _26_Export_additional_chart_types()
    {
        var table = new DataTable("AdditionalChartData");
        table.Columns.Add("MonthNumber", typeof(int));
        table.Columns.Add("Month", typeof(string));
        table.Columns.Add("Income", typeof(decimal));
        table.Columns.Add("Costs", typeof(decimal));
        table.Columns.Add("MarketShare", typeof(decimal));
        table.Rows.Add(1, "Jan", 120000m, 82000m, 0.18m);
        table.Rows.Add(2, "Feb", 135000m, 91000m, 0.21m);
        table.Rows.Add(3, "Mar", 128000m, 97000m, 0.19m);
        table.Rows.Add(4, "Apr", 155000m, 101000m, 0.23m);
        table.Rows.Add(5, "May", 162000m, 112000m, 0.25m);
        table.Rows.Add(6, "Jun", 171000m, 118000m, 0.28m);

        var bytes = ExcelWriter.Create()
            .AddSheet("Additional Chart Data", sheet => sheet
                .AddCell("A1", "Additional Chart Type Sources", style => style.Bold().FontSize(16).FontColor("#1F4E78"))
                .AddTable(table)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                .ColumnStyle("MarketShare", style => style.NumberFormat(ExcelNumberFormat.Percent).Border(ExcelBorderLineStyle.Thin, "#B4C6E7")))
            .AddSheet("Chart Types", sheet =>
            {
                sheet.AddCell("A1", "Additional Chart Types", style => style.Bold().FontSize(16).FontColor("#385723"));
                sheet.AddChart("Income Area", chart => chart
                    .Type(ExcelChartType.Area)
                    .DataSource("Additional Chart Data", "$B$4:$B$9", "$C$4:$C$9")
                    .Position("A3", "H18")
                    .Style(12)
                    .SeriesColor("#70AD47"));
                sheet.AddChart("Cost Share Pie", chart => chart
                    .Type(ExcelChartType.Pie)
                    .DataSource("Additional Chart Data", "$B$4:$B$9", "$D$4:$D$9")
                    .Position("I3", "P18")
                    .Style(10)
                    .SeriesColor("#C65911"));
                sheet.AddChart("Market Share Doughnut", chart => chart
                    .Type(ExcelChartType.Doughnut)
                    .DataSource("Additional Chart Data", "$B$4:$B$9", "$E$4:$E$9")
                    .Position("A20", "H35")
                    .Style(11)
                    .SeriesColor("#7030A0"));
                sheet.AddChart("Income Scatter", chart => chart
                    .Type(ExcelChartType.Scatter)
                    .DataSource("Additional Chart Data", "$A$4:$A$9", "$C$4:$C$9")
                    .Position("I20", "P35")
                    .Style(13)
                    .SeriesColor("#4472C4"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_26_Export_additional_chart_types), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(4, CountChartParts(bytes));
        Assert.Equal(1, CountChartPartsOfType<C.AreaChart>(bytes));
        Assert.Equal(1, CountChartPartsOfType<C.PieChart>(bytes));
        Assert.Equal(1, CountChartPartsOfType<C.DoughnutChart>(bytes));
        Assert.Equal(1, CountChartPartsOfType<C.ScatterChart>(bytes));
    }
}
