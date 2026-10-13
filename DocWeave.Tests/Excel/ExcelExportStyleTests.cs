using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Cell, column, row and range styling: fonts, fills, borders, number formats, conditional styles, merges, comments and hidden rows.
/// </summary>
public sealed class ExcelExportStyleTests
{
    [Fact]
    public void _08_Export_styled_datatable_with_header_fill_font_and_borders()
    {
        var invoices = CreateInvoiceTable();

        var bytes = ExcelWriter.Create()
            .AddSheet("Styled Invoices", sheet => sheet
                .AddCell("B2", "Styled Invoice Report", style => style
                    .Bold()
                    .FontSize(16)
                    .FontColor("#1F4E78"))
                .AddTable(invoices)
                .StartAt("B4")
                .HeaderStyle(style => style
                    .Bold()
                    .FontColor("#FFFFFF")
                    .BackgroundColor("#1F4E78")
                    .Border(ExcelBorderLineStyle.Thin, "#000000"))
                .DataStyle(style => style
                    .Border(ExcelBorderLineStyle.Thin, "#BFBFBF"))
                .ColumnStyle("Amount", style => style
                    .BackgroundColor("#FFF2CC")
                    .Border(ExcelBorderLineStyle.Thick, "#C65911"))
                .HeaderAlias("InvoiceNo", "Invoice Number")
                .HeaderAlias("CustomerName", "Customer"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_08_Export_styled_datatable_with_header_fill_font_and_borders), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Styled Invoices")
            .StartAt("B4")
            .ReadDataTable();

        Assert.Equal(6, imported.Rows.Count);
        Assert.True(CountStyledCells(bytes, "Styled Invoices") > 20);
    }

    [Fact]
    public void _09_Export_multi_sheet_workbook_with_different_style_sets()
    {
        var regions = new[]
        {
            new RegionalSummary("London", 125000m, 210000m, 85000m),
            new RegionalSummary("United States", 190000m, 325000m, 135000m),
            new RegionalSummary("Europe", 155000m, 260000m, 105000m)
        };
        var auditRows = new[]
        {
            new AuditLine("system", "Report", "Created", DateTime.Today.AddMinutes(8)),
            new AuditLine("approver", "Workbook", "Approved", DateTime.Today.AddMinutes(42))
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Executive", sheet => sheet
                .AddCell("A1", "Executive Summary", style => style.Bold().FontSize(15).FontColor("#7030A0"))
                .AddTable(regions)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#7030A0"))
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#D9D9D9"))
                .ColumnStyle("Profit", style => style.Bold().BackgroundColor("#E2F0D9")))
            .AddSheet("Audit", sheet => sheet
                .AddCell("A1", "Audit Trail", style => style.Bold().FontSize(14).FontColor("#C00000"))
                .AddTable(auditRows)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#C00000"))
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Dotted, "#808080")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_09_Export_multi_sheet_workbook_with_different_style_sets), bytes);

        Assert.True(CountStyledCells(bytes, "Executive") > 10);
        Assert.True(CountStyledCells(bytes, "Audit") > 5);
    }

    [Fact]
    public void _10_Export_mapped_columns_with_column_specific_styles()
    {
        var table = new DataTable("Staff");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Department", typeof(string));
        table.Columns.Add("Office", typeof(string));
        table.Columns.Add("RiskRating", typeof(string));
        table.Rows.Add(1, "Alice West", "Risk", "London", "High");
        table.Rows.Add(2, "Ben Carter", "Finance", "New York", "Medium");
        table.Rows.Add(3, "Claire Ito", "Technology", "Tokyo", "Low");

        var bytes = ExcelWriter.Create()
            .AddSheet("Styled Mapping", sheet => sheet
                .AddCell("B2", "Styled Column Mapping", style => style.Bold().FontSize(16).BackgroundColor("#D9EAF7"))
                .AddTable(table)
                .StartAt("B4")
                .ColumnLocation("ID", "B")
                .ColumnLocation("Name", "D")
                .ColumnLocation("Department", "E")
                .ColumnLocation("Office", "G")
                .ColumnLocation("RiskRating", "I")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#385723").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#A6A6A6"))
                .ColumnStyle("RiskRating", style => style.Bold().FontColor("#9C0006").BackgroundColor("#FFC7CE"))
                .HeaderAlias("RiskRating", "Risk Rating"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_10_Export_mapped_columns_with_column_specific_styles), bytes);

        var imported = ExcelReader.FromBytes(bytes)
            .Sheet("Styled Mapping")
            .StartAt("B4")
            .Columns(columns => columns.IncludeHeaders("ID", "Name", "Department", "Office", "Risk Rating"))
            .ReadDataTable();

        Assert.Equal(3, imported.Rows.Count);
        Assert.Equal("High", imported.Rows[0]["Risk Rating"]);
        Assert.True(CountStyledCells(bytes, "Styled Mapping") > 15);
    }

    [Fact]
    public void _11_Export_financial_table_with_formulas_borders_and_conditional_styles()
    {
        var table = new DataTable("Financials");
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Income", typeof(decimal));
        table.Columns.Add("Costs", typeof(decimal));
        table.Columns.Add("Variance", typeof(decimal));
        table.Rows.Add("London", 210000m, 125000m, 85000m);
        table.Rows.Add("United States", 325000m, 360000m, -35000m);
        table.Rows.Add("Europe", 260000m, 155000m, 105000m);
        table.Rows.Add("Asia", 295000m, 310000m, -15000m);

        var bytes = ExcelWriter.Create()
            .AddSheet("Financial Formulas", sheet =>
            {
                sheet.AddCell("A1", "Financial Formula Export", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(table)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border(ExcelBorderLineStyle.Thin))
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#D9E2F3"))
                    .ColumnStyle("Income", style => style.BackgroundColor("#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#70AD47"))
                    .ColumnStyle("Costs", style => style.BackgroundColor("#FCE4D6").Border(ExcelBorderLineStyle.Thin, "#C65911"))
                    .ConditionalStyle("Variance", value => Convert.ToDecimal(value) < 0m, style => style
                        .Bold()
                        .FontColor("#9C0006")
                        .BackgroundColor("#FFC7CE")
                        .Border(ExcelBorderLineStyle.Thick, "#9C0006"))
                    .FormulaColumn("Margin", context => $"D{context.WorksheetRow}/B{context.WorksheetRow}")
                    .ColumnStyle("Margin", style => style.BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Thin, "#5B9BD5"));

                sheet.AddCell("A9", "Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thin));
                sheet.AddFormulaCell("B9", "SUM(B4:B7)", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thin));
                sheet.AddFormulaCell("C9", "SUM(C4:C7)", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thin));
                sheet.AddFormulaCell("D9", "SUM(D4:D7)", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thin));
                sheet.AddFormulaCell("E9", "D9/B9", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thin));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_11_Export_financial_table_with_formulas_borders_and_conditional_styles), bytes);

        Assert.Equal(8, CountFormulaCells(bytes, "Financial Formulas"));
        Assert.True(CountStyledCells(bytes, "Financial Formulas") > 25);
    }

    [Fact]
    public void _12_Export_multi_sheet_operational_pack_with_formulas_and_conditional_status_styles()
    {
        var operations = new DataTable("Operations");
        operations.Columns.Add("Team", typeof(string));
        operations.Columns.Add("Completed", typeof(int));
        operations.Columns.Add("Failed", typeof(int));
        operations.Columns.Add("Status", typeof(string));
        operations.Rows.Add("Risk", 96, 2, "Green");
        operations.Rows.Add("Finance", 88, 9, "Amber");
        operations.Rows.Add("Technology", 74, 18, "Red");

        var actions = new[]
        {
            new Dictionary<string, object?> { ["Owner"] = "Alice", ["Action"] = "Review Red items", ["DueDate"] = new DateTime(2026, 2, 1), ["Priority"] = "High" },
            new Dictionary<string, object?> { ["Owner"] = "Ben", ["Action"] = "Confirm Amber tolerance", ["DueDate"] = new DateTime(2026, 2, 3), ["Priority"] = "Medium" },
            new Dictionary<string, object?> { ["Owner"] = "Claire", ["Action"] = "Archive Green evidence", ["DueDate"] = new DateTime(2026, 2, 5), ["Priority"] = "Low" }
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Operations", sheet => sheet
                .AddCell("A1", "Operational Pack", style => style.Bold().FontSize(16).FontColor("#385723"))
                .AddTable(operations)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#385723").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#A9D18E"))
                .ConditionalStyle("Status", value => string.Equals(value?.ToString(), "Red", StringComparison.OrdinalIgnoreCase), style => style.Bold().FontColor("#9C0006").BackgroundColor("#FFC7CE").Border(ExcelBorderLineStyle.Thick, "#9C0006"))
                .ConditionalStyle("Status", value => string.Equals(value?.ToString(), "Amber", StringComparison.OrdinalIgnoreCase), style => style.Bold().FontColor("#9C6500").BackgroundColor("#FFEB9C").Border(ExcelBorderLineStyle.Medium, "#BF9000"))
                .FormulaColumn("Total", context => $"B{context.WorksheetRow}+C{context.WorksheetRow}")
                .FormulaColumn("SuccessRate", context => $"B{context.WorksheetRow}/E{context.WorksheetRow}")
                .ColumnStyle("Total", style => style.BackgroundColor("#E2F0D9").Border())
                .ColumnStyle("SuccessRate", style => style.BackgroundColor("#DDEBF7").Border()))
            .AddSheet("Actions", sheet => sheet
                .AddCell("A1", "Action Register", style => style.Bold().FontSize(15).FontColor("#5F497A"))
                .AddRows(actions)
                .StartAt("A3")
                .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#5F497A").Border())
                .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#CCC0DA"))
                .ConditionalStyle("Priority", value => string.Equals(value?.ToString(), "High", StringComparison.OrdinalIgnoreCase), style => style.Bold().FontColor("#9C0006").BackgroundColor("#FFC7CE")))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_12_Export_multi_sheet_operational_pack_with_formulas_and_conditional_status_styles), bytes);

        Assert.Equal(6, CountFormulaCells(bytes, "Operations"));
        Assert.True(CountStyledCells(bytes, "Operations") > 15);
        Assert.True(CountStyledCells(bytes, "Actions") > 10);
    }

    [Fact]
    public void _16_Export_specific_rows_columns_and_cells_with_styles()
    {
        var table = new DataTable("StyledExceptions");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Columns.Add("Status", typeof(string));
        table.Rows.Add(1, "Alice West", "London", 1250m, "Open");
        table.Rows.Add(2, "Ben Carter", "New York", 2750m, "Review");
        table.Rows.Add(3, "Claire Ito", "Tokyo", 950m, "Closed");
        table.Rows.Add(4, "Dev Patel", "London", 4200m, "Escalated");
        table.Rows.Add(5, "Eva Muller", "Europe", 3100m, "Open");

        var bytes = ExcelWriter.Create()
            .AddSheet("Specific Styles", sheet =>
            {
                sheet.Column("G").Style(style => style.BackgroundColor("#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#70AD47"));
                sheet.AddCell("A1", "Specific Row, Column, and Cell Styles", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddTable(table)
                    .StartAt("B4")
                    .HeaderStyle(style => style.Bold().FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border())
                    .DataStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#BFBFBF"))
                    .AlternateRowStyle(style => style.BackgroundColor("#F2F2F2").Border(ExcelBorderLineStyle.Thin, "#BFBFBF"))
                    .ColumnStyle("Amount", style => style.BackgroundColor("#FFF2CC").Border(ExcelBorderLineStyle.Thin, "#BF9000"))
                    .RowStyle(2, style => style.Bold().BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Thick, "#5B9BD5"))
                    .RowStyleWhen(row => string.Equals(row["Status"]?.ToString(), "Escalated", StringComparison.OrdinalIgnoreCase), style => style.Bold().FontColor("#9C0006").BackgroundColor("#FFC7CE").Border(ExcelBorderLineStyle.Thick, "#9C0006"))
                    .CellStyle("Amount", row => Convert.ToDecimal(row["Amount"]) > 3000m, style => style.Bold().FontColor("#9C6500").BackgroundColor("#FFEB9C").Border(ExcelBorderLineStyle.Medium, "#BF9000"))
                    .FormulaColumn("AmountWithTax", context => $"E{context.WorksheetRow}*1.15")
                    .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_16_Export_specific_rows_columns_and_cells_with_styles), bytes);

        Assert.True(CountStyledCells(bytes, "Specific Styles") > 30);
        Assert.Equal("SUM(E5:E9)", GetCellFormula(bytes, "Specific Styles", "E10"));
        Assert.Equal("SUM(G5:G9)", GetCellFormula(bytes, "Specific Styles", "G10"));
    }

    [Fact]
    public void _17_Export_merged_cells_custom_fonts_wrapping_rotation_and_existing_features()
    {
        var table = new DataTable("LayoutFeatures");
        table.Columns.Add("Department", typeof(string));
        table.Columns.Add("Owner", typeof(string));
        table.Columns.Add("Narrative", typeof(string));
        table.Columns.Add("Score", typeof(decimal));
        table.Columns.Add("Status", typeof(string));
        table.Rows.Add("Risk", "Alice West", "Quarterly risk review completed with two follow-up actions for regional leads.", 92.5m, "Green");
        table.Rows.Add("Finance", "Ben Carter", "Month-end reconciliation needs additional evidence before sign-off.", 76.2m, "Amber");
        table.Rows.Add("Technology", "Claire Ito", "Critical patch rollout delayed by vendor dependency and requires executive review.", 58.4m, "Red");
        table.Rows.Add("Operations", "Dev Patel", "Control sample passed and the remaining actions are procedural.", 88.1m, "Green");

        var bytes = ExcelWriter.Create()
            .AddSheet("Layout Features", sheet =>
            {
                sheet.MergeCells("A1:H1");
                sheet.MergeCells("A2:H2");
                sheet.MergeCells("A12:C13");
                sheet.Column("D").Style(style => style.WrapText().FontFamily("Aptos").Border(ExcelBorderLineStyle.Thin, "#BFBFBF"));

                sheet.AddCell("A1", "Layout Feature Export", style => style
                    .Bold()
                    .FontFamily("Aptos Display")
                    .FontSize(18)
                    .FontColor("#FFFFFF")
                    .BackgroundColor("#1F4E78")
                    .Border(ExcelBorderLineStyle.Thick, "#17365D"));

                sheet.AddCell("A2", "Merged title rows, custom fonts, wrapped text, rotated headers, row and cell styles, totals, and charting.", style => style
                    .FontFamily("Aptos")
                    .Italic()
                    .WrapText()
                    .BackgroundColor("#DDEBF7")
                    .Border(ExcelBorderLineStyle.Thin, "#5B9BD5"));

                sheet.AddTable(table)
                    .StartAt("A4")
                    .HeaderStyle(style => style
                        .Bold()
                        .FontFamily("Aptos")
                        .FontColor("#FFFFFF")
                        .BackgroundColor("#385723")
                        .RotateText(45)
                        .WrapText()
                        .Border())
                    .DataStyle(style => style.FontFamily("Aptos").WrapText().Border(ExcelBorderLineStyle.Thin, "#A9D18E"))
                    .AlternateRowStyle(style => style.FontFamily("Aptos").WrapText().BackgroundColor("#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#A9D18E"))
                    .ColumnStyle("Narrative", style => style.FontFamily("Arial").WrapText().BackgroundColor("#F2F2F2").Border(ExcelBorderLineStyle.Thin, "#BFBFBF"))
                    .RowStyleWhen(row => string.Equals(row["Status"]?.ToString(), "Red", StringComparison.OrdinalIgnoreCase), style => style.Bold().FontColor("#9C0006").BackgroundColor("#FFC7CE").WrapText().Border(ExcelBorderLineStyle.Thick, "#9C0006"))
                    .CellStyle("Score", row => Convert.ToDecimal(row["Score"]) < 70m, style => style.Bold().FontColor("#9C0006").BackgroundColor("#FFC7CE").Border(ExcelBorderLineStyle.Medium, "#9C0006"))
                    .AddTotals.Bottom("Totals", style => style.Bold().FontFamily("Aptos").BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Thick, "#000000"));

                sheet.AddCell("A12", "Merged note area with wrapped text for report commentary. This is deliberately long enough to require wrapping when opened in Excel.", style => style
                    .FontFamily("Georgia")
                    .WrapText()
                    .BackgroundColor("#FFF2CC")
                    .Border(ExcelBorderLineStyle.Thick, "#BF9000"));

                sheet.AddChart("Scores By Department", chart => chart
                    .Type(ExcelChartType.Column)
                    .DataSource("Layout Features", "$A$5:$A$8", "$D$5:$D$8")
                    .Position("E12", "H25")
                    .Style(10)
                    .SeriesColor("#70AD47"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_17_Export_merged_cells_custom_fonts_wrapping_rotation_and_existing_features), bytes);

        Assert.Equal(3, CountMergeRanges(bytes, "Layout Features"));
        Assert.True(CountWrappedOrRotatedCells(bytes) > 10);
        Assert.Equal(1, CountChartParts(bytes));
        Assert.Equal("SUM(D5:D8)", GetCellFormula(bytes, "Layout Features", "D9"));
    }

    [Fact]
    public void _18_Export_comments_dimensions_formats_pattern_fills_and_hidden_rows_columns()
    {
        var table = new DataTable("FormatOptions");
        table.Columns.Add("Account", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Columns.Add("Rate", typeof(decimal));
        table.Columns.Add("RunDate", typeof(DateTime));
        table.Columns.Add("Notes", typeof(string));
        table.Rows.Add("Revenue", 125000.50m, 0.125m, new DateTime(2026, 3, 1), "Revenue includes late adjustments and requires wrapped text.");
        table.Rows.Add("Costs", 82500.25m, 0.0825m, new DateTime(2026, 3, 2), "Costs include an accrual review comment.");
        table.Rows.Add("Margin", 42500.25m, 0.0425m, new DateTime(2026, 3, 3), "Margin is preliminary.");

        var bytes = ExcelWriter.Create()
            .AddSheet("Format Options", sheet =>
            {
                sheet.MergeCells("A1:F1");
                sheet.Column("A").Width(18);
                sheet.Column("B").Width(16);
                sheet.Column("C").AutoFitWidth();
                sheet.Column("E").Width(34).Style(style => style.WrapText());
                sheet.Column("F").Hidden();
                sheet.Row(1).Height(28);
                sheet.Row(2).AutoFitHeight();
                sheet.Row(12).Hidden();
                sheet.AddCell("A1", "Formatting Options Export", style => style.Bold().Underline().FontFamily("Aptos Display").FontSize(18).FontColor("#FFFFFF").BackgroundColor("#1F4E78"));
                sheet.AddComment("B5", "Amount reviewed by Finance.", "DocWeave");
                sheet.AddComment("E6", "Long notes should wrap.", "DocWeave");
                sheet.AddTable(table)
                    .StartAt("A4")
                    .HeaderStyle(style => style.Bold().FontFamily("Aptos").FontColor("#FFFFFF").BackgroundColor("#385723").RotateText(45).Border(ExcelBorderLineStyle.Medium))
                    .DataStyle(style => style.FontFamily("Aptos").Border(ExcelBorderLineStyle.Thin, "#BFBFBF"))
                    .AlternateRowStyle(style => style.PatternFill(ExcelFillPattern.LightGrid, "#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#A9D18E"))
                    .ColumnStyle("Notes", style => style.WrapText().PatternFill(ExcelFillPattern.DarkHorizontal, "#FFF2CC").Border(ExcelBorderLineStyle.DashDot, "#BF9000"))
                    .AddTotals.Bottom("Totals", style => style.Bold().BackgroundColor("#D9EAD3").Border(ExcelBorderLineStyle.Double));
                sheet.RangeStyle("B5:B8", style => style.NumberFormat(ExcelNumberFormat.Currency));
                sheet.RangeStyle("C5:C8", style => style.NumberFormat(ExcelNumberFormat.Percent));
                sheet.RangeStyle("D5:D7", style => style.NumberFormat(ExcelNumberFormat.Date));
                sheet.RangeStyle("A10:F10", style => style.PatternFill(ExcelFillPattern.DarkGrid, "#DDEBF7").Border(ExcelBorderLineStyle.SlantDashDot, "#1F4E78"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_18_Export_comments_dimensions_formats_pattern_fills_and_hidden_rows_columns), bytes);

        Assert.Equal(2, CountComments(bytes, "Format Options"));
        Assert.True(CountHiddenColumns(bytes, "Format Options") >= 1);
        Assert.True(CountHiddenRows(bytes, "Format Options") >= 1);
        Assert.True(CountPatternFills(bytes) >= 2);
        Assert.True(CountFormattedCells(bytes, ExcelNumberFormat.Currency) >= 4);
    }

    [Fact]
    public void _19_Export_range_formats_all_border_styles_and_hidden_detail_band()
    {
        var rows = new[]
        {
            new BorderDemoRow("Thin", 10m),
            new BorderDemoRow("Medium", 20m),
            new BorderDemoRow("Thick", 30m),
            new BorderDemoRow("Dashed", 40m),
            new BorderDemoRow("Dotted", 50m),
            new BorderDemoRow("Double", 60m),
            new BorderDemoRow("DashDot", 70m),
            new BorderDemoRow("DashDotDot", 80m),
            new BorderDemoRow("SlantDashDot", 90m)
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Border Styles", sheet =>
            {
                sheet.Column("D").Hidden();
                sheet.Row(15).Hidden();
                sheet.AddCell("A1", "Border and Range Formatting", style => style.Bold().FontSize(16).Underline());
                sheet.AddTable(rows)
                    .StartAt("A3")
                    .HeaderStyle(style => style.Bold().BackgroundColor("#1F4E78").FontColor("#FFFFFF").Border())
                    .DataStyle(style => style.Border());
                sheet.RangeStyle("A4:B4", style => style.Border(ExcelBorderLineStyle.Thin));
                sheet.RangeStyle("A5:B5", style => style.Border(ExcelBorderLineStyle.Medium));
                sheet.RangeStyle("A6:B6", style => style.Border(ExcelBorderLineStyle.Thick));
                sheet.RangeStyle("A7:B7", style => style.Border(ExcelBorderLineStyle.Dashed));
                sheet.RangeStyle("A8:B8", style => style.Border(ExcelBorderLineStyle.Dotted));
                sheet.RangeStyle("A9:B9", style => style.Border(ExcelBorderLineStyle.Double));
                sheet.RangeStyle("A10:B10", style => style.Border(ExcelBorderLineStyle.DashDot));
                sheet.RangeStyle("A11:B11", style => style.Border(ExcelBorderLineStyle.DashDotDot));
                sheet.RangeStyle("A12:B12", style => style.Border(ExcelBorderLineStyle.SlantDashDot));
                sheet.RangeStyle("B4:B12", style => style.NumberFormat(ExcelNumberFormat.Number));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_19_Export_range_formats_all_border_styles_and_hidden_detail_band), bytes);

        Assert.True(CountHiddenColumns(bytes, "Border Styles") >= 1);
        Assert.True(CountHiddenRows(bytes, "Border Styles") >= 1);
        Assert.True(CountFormattedCells(bytes, ExcelNumberFormat.Number) >= 9);
    }

    private sealed record BorderDemoRow(string StyleName, decimal Value);
}
