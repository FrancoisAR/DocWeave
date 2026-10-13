using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Totals rows, alternate row fills and summary formulas.
/// </summary>
public sealed class ExcelExportTotalsTests
{
    [Fact]
    public void _13_Export_three_dataset_tables_with_totals_alternate_rows_and_summary_formulas()
    {
        var dataSet = CreateLargeSummaryDataSet();
        var bytes = ExcelWriter.Create()
            .AddSheet("SalesData", sheet => AddLargeDataSheet(
                sheet,
                dataSet.Tables["SalesData"]!,
                "Sales Data",
                "#1F4E78",
                "#DDEBF7",
                "#B4C6E7"))
            .AddSheet("CostData", sheet => AddLargeDataSheet(
                sheet,
                dataSet.Tables["CostData"]!,
                "Cost Data",
                "#385723",
                "#E2F0D9",
                "#A9D18E"))
            .AddSheet("ForecastData", sheet => AddLargeDataSheet(
                sheet,
                dataSet.Tables["ForecastData"]!,
                "Forecast Data",
                "#7030A0",
                "#E4DFEC",
                "#CCC0DA"))
            .AddSheet("Summary", AddLargeDataSummarySheet)
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_13_Export_three_dataset_tables_with_totals_alternate_rows_and_summary_formulas), bytes);

        Assert.Equal(54, CountFormulaCells(bytes, "SalesData") + CountFormulaCells(bytes, "CostData") + CountFormulaCells(bytes, "ForecastData") + CountFormulaCells(bytes, "Summary"));
        Assert.True(CountStyledCells(bytes, "SalesData") > 900);
        Assert.True(CountStyledCells(bytes, "CostData") > 900);
        Assert.True(CountStyledCells(bytes, "ForecastData") > 900);
        Assert.True(CountStyledCells(bytes, "Summary") > 25);
    }

    [Fact]
    public void _14_Export_variable_row_table_with_generated_bottom_totals_row()
    {
        var table = CreateLargeMetricTable("VariableRows", seed: 4000);
        while (table.Rows.Count > 17)
        {
            table.Rows.RemoveAt(table.Rows.Count - 1);
        }

        var bytes = ExcelWriter.Create()
            .AddSheet("Variable Totals", sheet => sheet
                .AddCell("A1", "Variable Row Totals", style => style.Bold().FontSize(16).FontColor("#1F4E78"))
                .AddTable(table)
                .StartAt("B4")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                .AlternateRowStyle(style => style.BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Thin, "#B4C6E7"))
                .AddTotals.Bottom("Dynamic Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_14_Export_variable_row_table_with_generated_bottom_totals_row), bytes);

        Assert.Equal("SUM(C5:C21)", GetCellFormula(bytes, "Variable Totals", "C22"));
        Assert.Equal("SUM(K5:K21)", GetCellFormula(bytes, "Variable Totals", "K22"));
        Assert.Equal("Dynamic Totals", GetCellText(bytes, "Variable Totals", "B22"));
    }
}
