using System.Data;
using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Sample tables and sheets shared by the export tests.
/// </summary>
internal static class SampleData
{
    internal static DataSet CreateBusinessDataSet()
    {
        var dataSet = new DataSet("BusinessExport");
        dataSet.Tables.Add(CreateRegionsTable());
        dataSet.Tables.Add(CreateInvoiceTable());
        return dataSet;
    }

    internal static DataTable CreateRegionsTable()
    {
        var table = new DataTable("Regions");
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Manager", typeof(string));
        table.Columns.Add("OfficeCount", typeof(int));
        table.Columns.Add("Headcount", typeof(int));

        table.Rows.Add("London", "Alice West", 4, 220);
        table.Rows.Add("United States", "Ben Carter", 8, 480);
        table.Rows.Add("Europe", "Claire Ito", 5, 310);
        table.Rows.Add("Asia", "Dev Patel", 6, 360);

        return table;
    }

    internal static DataTable CreatePivotSalesTable()
    {
        var table = new DataTable("PivotSales");
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Product", typeof(string));
        table.Columns.Add("Quarter", typeof(string));
        table.Columns.Add("Channel", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));

        table.Rows.Add("London", "Software", "Q1", "Direct", 125000m);
        table.Rows.Add("London", "Services", "Q1", "Partner", 68000m);
        table.Rows.Add("London", "Software", "Q2", "Direct", 148000m);
        table.Rows.Add("United States", "Software", "Q1", "Direct", 220000m);
        table.Rows.Add("United States", "Services", "Q2", "Partner", 99000m);
        table.Rows.Add("Europe", "Software", "Q1", "Partner", 142000m);
        table.Rows.Add("Europe", "Services", "Q2", "Direct", 76000m);
        table.Rows.Add("Asia", "Software", "Q2", "Direct", 171000m);

        return table;
    }

    internal static DataTable CreateComplexPivotSalesTable()
    {
        var table = new DataTable("ComplexPivotSales");
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Country", typeof(string));
        table.Columns.Add("Product", typeof(string));
        table.Columns.Add("Segment", typeof(string));
        table.Columns.Add("Quarter", typeof(string));
        table.Columns.Add("Channel", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Columns.Add("Units", typeof(int));
        table.Columns.Add("Cost", typeof(decimal));

        table.Rows.Add("Europe", "United Kingdom", "Software", "Enterprise", "Q1", "Direct", 125000m, 24, 69000m);
        table.Rows.Add("Europe", "United Kingdom", "Services", "MidMarket", "Q1", "Partner", 68000m, 14, 41000m);
        table.Rows.Add("Europe", "Germany", "Software", "Enterprise", "Q2", "Direct", 148000m, 27, 76000m);
        table.Rows.Add("Americas", "United States", "Software", "Enterprise", "Q1", "Direct", 220000m, 42, 121000m);
        table.Rows.Add("Americas", "United States", "Services", "MidMarket", "Q2", "Partner", 99000m, 19, 55000m);
        table.Rows.Add("Americas", "Canada", "Hardware", "SMB", "Q1", "Partner", 73000m, 31, 47000m);
        table.Rows.Add("EMEA", "France", "Software", "Enterprise", "Q1", "Partner", 142000m, 28, 78000m);
        table.Rows.Add("EMEA", "France", "Services", "SMB", "Q2", "Direct", 76000m, 16, 43000m);
        table.Rows.Add("APAC", "Japan", "Software", "Enterprise", "Q2", "Direct", 171000m, 34, 93000m);
        table.Rows.Add("APAC", "Singapore", "Hardware", "MidMarket", "Q1", "Partner", 88000m, 38, 56000m);
        table.Rows.Add("APAC", "Australia", "Services", "SMB", "Q2", "Direct", 65000m, 13, 39000m);
        table.Rows.Add("Europe", "United Kingdom", "Hardware", "SMB", "Q3", "Direct", 54000m, 22, 33000m);

        return table;
    }

    internal static IReadOnlyList<CurrencyRateRow> CreateCurrencyRateRows()
    {
        return
        [
            new("USD", "United States", 1.2875m, 1.2420m, new DateTime(2026, 9, 18)),
            new("GBP", "United Kingdom", 1.0000m, 1.0000m, new DateTime(2026, 9, 18)),
            new("EUR", "Eurozone", 1.1764m, 1.1325m, new DateTime(2026, 9, 18)),
            new("JPY", "Japan", 191.2400m, 184.8200m, new DateTime(2026, 9, 17)),
            new("CAD", "Canada", 1.7420m, 1.6810m, new DateTime(2026, 9, 17)),
            new("AUD", "Australia", 1.9485m, 1.8760m, new DateTime(2026, 9, 17)),
            new("CHF", "Switzerland", 1.0340m, 0.9982m, new DateTime(2026, 9, 16)),
            new("ZAR", "South Africa", 22.6500m, 21.9400m, new DateTime(2026, 9, 16)),
            new("INR", "India", 113.8700m, 109.4200m, new DateTime(2026, 9, 15)),
            new("SGD", "Singapore", 1.6640m, 1.6115m, new DateTime(2026, 9, 15)),
            new(
                "Summary",
                "10 currencies",
                new ExcelFormula("AVERAGE(B6,G6,L6,B13,G13,L13,B20,G20,L20,B27)"),
                new ExcelFormula("AVERAGE(B7,G7,L7,B14,G14,L14,B21,G21,L21,B28)"),
                "Formula totals")
        ];
    }

    internal static DataTable CreateInvoiceTable()
    {
        var table = new DataTable("Invoices");
        table.Columns.Add("InvoiceNo", typeof(string));
        table.Columns.Add("CustomerName", typeof(string));
        table.Columns.Add("InvoiceDate", typeof(DateTime));
        table.Columns.Add("Status", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Columns.Add("Currency", typeof(string));

        table.Rows.Add("INV-1001", "Blue River Ltd", new DateTime(2026, 1, 7), "Open", 1250.75m, "GBP");
        table.Rows.Add("INV-1002", "Northwind Finance", new DateTime(2026, 1, 9), "Paid", 4875.00m, "USD");
        table.Rows.Add("INV-1003", "Tokyo Banking", new DateTime(2026, 1, 12), "Pending", 2310.25m, "JPY");
        table.Rows.Add("INV-1004", "Paris Retail", new DateTime(2026, 1, 15), "Open", 995.40m, "EUR");
        table.Rows.Add("INV-1005", "Singapore Ops", new DateTime(2026, 1, 18), "Paid", 7200.00m, "SGD");
        table.Rows.Add("INV-1006", "Dubai Trading", new DateTime(2026, 1, 22), "Closed", 3100.10m, "AED");

        return table;
    }

    internal static DataSet CreateLargeSummaryDataSet()
    {
        var dataSet = new DataSet("LargeSummaryExport");
        dataSet.Tables.Add(CreateLargeMetricTable("SalesData", seed: 1000));
        dataSet.Tables.Add(CreateLargeMetricTable("CostData", seed: 2000));
        dataSet.Tables.Add(CreateLargeMetricTable("ForecastData", seed: 3000));
        return dataSet;
    }

    internal static DataTable CreateLargeMetricTable(string tableName, int seed)
    {
        var table = new DataTable(tableName);
        table.Columns.Add("Item", typeof(string));
        for (var column = 1; column <= 9; column++)
        {
            table.Columns.Add($"Metric{column:00}", typeof(decimal));
        }

        for (var row = 1; row <= 100; row++)
        {
            var values = new object[10];
            values[0] = $"{tableName}-Row-{row:000}";
            for (var column = 1; column <= 9; column++)
            {
                values[column] = seed + row * 10 + column;
            }

            table.Rows.Add(values);
        }

        return table;
    }

    internal static void AddLargeDataSheet(
        ExcelSheetExportBuilder sheet,
        DataTable table,
        string title,
        string headerColor,
        string alternateFillColor,
        string borderColor)
    {
        sheet.AddCell("A1", title, style => style.Bold().FontSize(16).FontColor(headerColor));
        sheet.AddCell("A2", "100 data rows, 10 columns, alternate row styling, and total formulas.", style => style.Italic().FontColor("#666666"));
        sheet.AddTable(table)
            .StartAt("A4")
            .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor(headerColor).Border(ExcelBorderLineStyle.Thin, "#000000"))
            .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, borderColor))
            .AlternateRowStyle(style => style.BackgroundColor(alternateFillColor).Border(ExcelBorderLineStyle.Thin, borderColor))
            .ConditionalStyle("Metric09", value => Convert.ToDecimal(value) > 0m, style => style.Bold().BackgroundColor("#FFF2CC").Border(ExcelBorderLineStyle.Thin, "#BF9000"))
            .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000"));
    }

    internal static void AddLargeDataSummarySheet(ExcelSheetExportBuilder sheet)
    {
        sheet.AddCell("A1", "Dataset Summary", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
        sheet.AddCell("A3", "Source Sheet", style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border());

        for (var column = 2; column <= 10; column++)
        {
            var columnLetter = ColumnLetter(column);
            sheet.AddCell($"{columnLetter}3", $"Metric{column - 1:00}", style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border());
        }

        var sourceSheets = new[] { "SalesData", "CostData", "ForecastData" };
        for (var rowIndex = 0; rowIndex < sourceSheets.Length; rowIndex++)
        {
            var worksheetRow = 4 + rowIndex;
            var sourceSheet = sourceSheets[rowIndex];
            sheet.AddCell($"A{worksheetRow}", sourceSheet, style => style.Bold().BackgroundColor(rowIndex % 2 == 0 ? "#DDEBF7" : "#E2F0D9").Border());

            for (var column = 2; column <= 10; column++)
            {
                var columnLetter = ColumnLetter(column);
                sheet.AddFormulaCell(
                    $"{columnLetter}{worksheetRow}",
                    $"'{sourceSheet}'!{columnLetter}105",
                    style => style.BackgroundColor(rowIndex % 2 == 0 ? "#DDEBF7" : "#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#808080"));
            }
        }
    }

    internal static string ColumnLetter(int column)
    {
        var value = column;
        var letters = string.Empty;
        while (value > 0)
        {
            value--;
            letters = (char)('A' + value % 26) + letters;
            value /= 26;
        }

        return letters;
    }
}

/// <summary>Row type used by shared sample data.</summary>
internal sealed record RegionalSummary(string Region, decimal Costs, decimal Income, decimal Profit);

/// <summary>Row type used by shared sample data.</summary>
internal sealed record AuditLine(string User, string Area, string Action, DateTime Timestamp);

/// <summary>Row type used by shared sample data.</summary>
internal sealed record CurrencyRateRow(string Code, string Country, object LastHighRate, object LastLowRate, object LastUpdated);
