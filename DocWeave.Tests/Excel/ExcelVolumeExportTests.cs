using System.Data;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// A large workbook with many sheets and a wide table.
/// </summary>
public sealed class ExcelVolumeExportTests
{
    private const string RunLargeExcelVolumeVariable = "DOCWEAVE_RUN_LARGE_EXCEL_VOLUME";

    [Fact]
    public void _41_Export_volume_workbook_with_ten_variable_sheets_and_wide_table()
    {
        var sheetDefinitions = new[]
        {
            (Name: "Volume 01", Rows: 45, NumericColumns: 8, HeaderColor: "#1F4E78", AlternateFill: "#DDEBF7", BorderColor: "#B4C6E7"),
            (Name: "Volume 02", Rows: 72, NumericColumns: 11, HeaderColor: "#385723", AlternateFill: "#E2F0D9", BorderColor: "#A9D18E"),
            (Name: "Volume 03", Rows: 96, NumericColumns: 14, HeaderColor: "#7030A0", AlternateFill: "#E4DFEC", BorderColor: "#CCC0DA"),
            (Name: "Volume 04", Rows: 128, NumericColumns: 17, HeaderColor: "#C65911", AlternateFill: "#FCE4D6", BorderColor: "#F4B183"),
            (Name: "Volume 05", Rows: 155, NumericColumns: 20, HeaderColor: "#5B9BD5", AlternateFill: "#DDEBF7", BorderColor: "#9DC3E6"),
            (Name: "Volume Wide", Rows: 180, NumericColumns: 35, HeaderColor: "#8064A2", AlternateFill: "#E4DFEC", BorderColor: "#B1A0C7"),
            (Name: "Volume 07", Rows: 64, NumericColumns: 9, HeaderColor: "#9E480E", AlternateFill: "#FCE4D6", BorderColor: "#F4B183"),
            (Name: "Volume 08", Rows: 118, NumericColumns: 13, HeaderColor: "#31859B", AlternateFill: "#DAEEF3", BorderColor: "#92CDDC"),
            (Name: "Volume 09", Rows: 142, NumericColumns: 16, HeaderColor: "#60497A", AlternateFill: "#E4DFEC", BorderColor: "#B1A0C7"),
            (Name: "Volume 10", Rows: 210, NumericColumns: 22, HeaderColor: "#4F6228", AlternateFill: "#EBF1DE", BorderColor: "#C4D79B")
        };

        var workbook = ExcelWriter.Create();
        foreach (var definition in sheetDefinitions)
        {
            var table = CreateVolumeExportTable(definition.Name.Replace(" ", string.Empty, StringComparison.Ordinal), definition.Rows, definition.NumericColumns);
            workbook.AddSheet(definition.Name, sheet => AddVolumeExportSheet(
                sheet,
                table,
                definition.Name,
                definition.NumericColumns,
                definition.HeaderColor,
                definition.AlternateFill,
                definition.BorderColor));
        }

        var bytes = workbook.ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_41_Export_volume_workbook_with_ten_variable_sheets_and_wide_table), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(10, CountWorksheets(bytes));
        Assert.True(GetMaxUsedColumn(bytes, "Volume Wide") >= 38);
        Assert.Equal("SUM(C5:C184)", GetCellFormula(bytes, "Volume Wide", "C185"));
        Assert.Equal("SUM(AL5:AL184)", GetCellFormula(bytes, "Volume Wide", "AL185"));
        Assert.True(sheetDefinitions.Sum(definition => CountFormulaCells(bytes, definition.Name)) > 2500);
        Assert.True(CountStyledCells(bytes, "Volume Wide") > 6000);
        Assert.True(CountStyledCells(bytes, "Volume 10") > 4000);
    }

    [Fact]
    public void _50_Export_tall_styled_table_with_thousands_of_rows()
    {
        var rowCount = ShouldRunLargeExcelVolumeTests() ? 20000 : 5000;
        var numericColumns = 40;
        var table = CreateVolumeExportTable("TallStyledRows", rowCount, numericColumns);

        var bytes = ExcelWriter.Create()
            .AddSheet("Tall Styled", sheet => AddVolumeExportSheet(
                sheet,
                table,
                "Tall Styled Rows",
                numericColumns,
                "#1F4E78",
                "#DDEBF7",
                "#9DC3E6"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_50_Export_tall_styled_table_with_thousands_of_rows), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal($"SUM(C5:C{rowCount + 4})", GetCellFormula(bytes, "Tall Styled", $"C{rowCount + 5}"));
        Assert.True(CountStyledCells(bytes, "Tall Styled") > rowCount * 20);
    }

    [Fact]
    public void _51_Export_very_wide_styled_table_with_many_columns()
    {
        var rowCount = ShouldRunLargeExcelVolumeTests() ? 3000 : 900;
        var numericColumns = 120;
        var table = CreateVolumeExportTable("WideStyledColumns", rowCount, numericColumns);

        var bytes = ExcelWriter.Create()
            .AddSheet("Wide Styled", sheet => AddVolumeExportSheet(
                sheet,
                table,
                "Wide Styled Columns",
                numericColumns,
                "#385723",
                "#E2F0D9",
                "#A9D18E"))
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_51_Export_very_wide_styled_table_with_many_columns), bytes);

        Xlsx.AssertValid(bytes);
        Assert.True(GetMaxUsedColumn(bytes, "Wide Styled") >= 123);
        Assert.Equal($"SUM(C5:DR5)", GetCellFormula(bytes, "Wide Styled", "DS5"));
        Assert.True(CountStyledCells(bytes, "Wide Styled") > rowCount * 80);
    }

    [Fact]
    public void _52_Export_multi_sheet_large_styled_workbook()
    {
        var rowCount = ShouldRunLargeExcelVolumeTests() ? 6000 : 1500;
        var sheetDefinitions = new[]
        {
            (Name: "Large North", NumericColumns: 45, HeaderColor: "#7030A0", AlternateFill: "#E4DFEC", BorderColor: "#B1A0C7"),
            (Name: "Large South", NumericColumns: 55, HeaderColor: "#C65911", AlternateFill: "#FCE4D6", BorderColor: "#F4B183"),
            (Name: "Large East", NumericColumns: 65, HeaderColor: "#31859B", AlternateFill: "#DAEEF3", BorderColor: "#92CDDC"),
            (Name: "Large West", NumericColumns: 75, HeaderColor: "#60497A", AlternateFill: "#E4DFEC", BorderColor: "#B1A0C7")
        };

        var workbook = ExcelWriter.Create();
        foreach (var definition in sheetDefinitions)
        {
            var table = CreateVolumeExportTable(definition.Name.Replace(" ", string.Empty, StringComparison.Ordinal), rowCount, definition.NumericColumns);
            workbook.AddSheet(definition.Name, sheet => AddVolumeExportSheet(
                sheet,
                table,
                definition.Name,
                definition.NumericColumns,
                definition.HeaderColor,
                definition.AlternateFill,
                definition.BorderColor));
        }

        var bytes = workbook.ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_52_Export_multi_sheet_large_styled_workbook), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(4, CountWorksheets(bytes));
        Assert.True(sheetDefinitions.Sum(definition => CountFormulaCells(bytes, definition.Name)) > rowCount * 4);
        Assert.True(CountStyledCells(bytes, "Large West") > rowCount * 50);
    }

    private static DataTable CreateVolumeExportTable(string tableName, int rows, int numericColumns)
    {
        var table = new DataTable(tableName);
        table.Columns.Add("Item", typeof(string));
        for (var column = 1; column <= numericColumns; column++)
        {
            table.Columns.Add($"Metric{column:00}", typeof(decimal));
        }

        for (var row = 1; row <= rows; row++)
        {
            var values = new object[numericColumns + 1];
            values[0] = $"{tableName}-Row-{row:0000}";
            for (var column = 1; column <= numericColumns; column++)
            {
                values[column] = row * (column + 3) + column / 10m;
            }

            table.Rows.Add(values);
        }

        return table;
    }

    private static bool ShouldRunLargeExcelVolumeTests()
    {
        return string.Equals(Environment.GetEnvironmentVariable(RunLargeExcelVolumeVariable), "1", StringComparison.Ordinal);
    }

    private static void AddVolumeExportSheet(
        ExcelSheetExportBuilder sheet,
        DataTable table,
        string title,
        int numericColumns,
        string headerColor,
        string alternateFillColor,
        string borderColor)
    {
        var lastNumericColumn = ColumnLetter(2 + numericColumns);
        var rowTotalColumn = ColumnLetter(3 + numericColumns);

        sheet.MergeCells("A1:H1");
        sheet.Column("A").Width(2);
        sheet.Column("B").Width(26);
        sheet.Column("C").Width(14);
        sheet.Column(lastNumericColumn).Width(16);
        sheet.Row(1).Height(26);
        sheet.AddCell("A1", $"{title} Volume Export", style => style
            .Bold()
            .FontFamily("Aptos Display")
            .FontSize(16)
            .FontColor("#FFFFFF")
            .BackgroundColor(headerColor)
            .Border(ExcelBorderLineStyle.Thick, "#000000"));
        sheet.AddCell("B2", $"{table.Rows.Count} rows, {table.Columns.Count} source columns, generated totals, row formulas, alternate row fills, conditional styles, and borders.", style => style
            .Italic()
            .FontFamily("Aptos")
            .WrapText()
            .FontColor("#666666"));

        sheet.AddTable(table)
            .StartAt("B4")
            .HeaderStyle(style => style
                .Bold()
                .FontFamily("Aptos")
                .FontColor("#FFFFFF")
                .BackgroundColor(headerColor)
                .Border(ExcelBorderLineStyle.Thin, "#000000"))
            .DataStyle(style => style
                .FontFamily("Aptos")
                .Border(ExcelBorderLineStyle.Thin, borderColor))
            .AlternateRowStyle(style => style
                .FontFamily("Aptos")
                .BackgroundColor(alternateFillColor)
                .Border(ExcelBorderLineStyle.Thin, borderColor))
            .ColumnStyle("Metric01", style => style
                .NumberFormat(ExcelNumberFormat.Currency)
                .BackgroundColor("#FFF2CC")
                .Border(ExcelBorderLineStyle.Medium, "#BF9000"))
            .ColumnStyle($"Metric{numericColumns:00}", style => style
                .NumberFormat(ExcelNumberFormat.Number)
                .Bold()
                .BackgroundColor("#E2F0D9")
                .Border(ExcelBorderLineStyle.Thick, "#548235"))
            .RowStyle(1, style => style
                .Bold()
                .BackgroundColor("#D9EAD3")
                .Border(ExcelBorderLineStyle.Medium, "#548235"))
            .RowStyleWhen(row => (int)row["__DataRowNumber"]! % 25 == 0, style => style
                .Bold()
                .FontColor("#9C6500")
                .BackgroundColor("#FFEB9C")
                .Border(ExcelBorderLineStyle.Medium, "#BF9000"))
            .FormulaColumn("RowTotal", context => $"SUM(C{context.WorksheetRow}:{lastNumericColumn}{context.WorksheetRow})")
            .FormulaColumn("RowAverage", context => $"{rowTotalColumn}{context.WorksheetRow}/{numericColumns}")
            .ColumnStyle("RowTotal", style => style
                .NumberFormat(ExcelNumberFormat.Number)
                .Bold()
                .BackgroundColor("#DDEBF7")
                .Border(ExcelBorderLineStyle.Thick, "#5B9BD5"))
            .ColumnStyle("RowAverage", style => style
                .NumberFormat(ExcelNumberFormat.Number)
                .BackgroundColor("#FCE4D6")
                .Border(ExcelBorderLineStyle.Medium, "#C65911"))
            .AddTotals.Bottom("Totals", style => style
                .Bold()
                .FontFamily("Aptos")
                .BackgroundColor("#D9EAD3")
                .Border(ExcelBorderLineStyle.Thick, "#000000"));
    }
}
