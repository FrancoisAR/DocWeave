using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;
using static DocWeave.Tests.Excel.SampleData;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Repeating card layouts built from an object collection.
/// </summary>
public sealed class ExcelRepeaterExportTests
{
    [Fact]
    public void _25_Export_object_collection_as_repeating_currency_cards()
    {
        var currencies = CreateCurrencyRateRows();

        var bytes = ExcelWriter.Create()
            .AddSheet("Currency Cards", sheet =>
            {
                sheet.AddCell("A1", "Currency Rate Cards", style => style.Bold().FontSize(16).FontColor("#1F4E78"));
                sheet.AddCell("A2", "Object collection exported as repeating layout blocks with three items per row.", style => style.Italic().FontColor("#666666"));
                sheet.Column("A").Width(16);
                sheet.Column("B").Width(15);
                sheet.Column("F").Width(16);
                sheet.Column("G").Width(15);
                sheet.Column("K").Width(16);
                sheet.Column("L").Width(15);

                sheet.AddRepeater(currencies)
                    .StartAt("A4")
                    .ItemsPerRow(3)
                    .ItemSize(columns: 4, rows: 6)
                    .Gap(columns: 1, rows: 1)
                    .TitleField("Code")
                    .Field("Country")
                    .Field("LastHighRate", "Last High")
                    .Field("LastLowRate", "Last Low")
                    .Field("LastUpdated", "Last Updated")
                    .FormulaField("High-Low Spread", context =>
                    {
                        var valueColumn = ColumnLetter(context.WorksheetColumn + 1);
                        return $"{valueColumn}{context.WorksheetRow + 2}-{valueColumn}{context.WorksheetRow + 3}";
                    })
                    .CardStyle(style => style.BackgroundColor("#F8FBFD").Border(ExcelBorderLineStyle.Thin, "#B7C9D6"))
                    .TitleStyle(style => style.Bold().FontSize(13).FontColor("#FFFFFF").BackgroundColor("#1F4E78").Border(ExcelBorderLineStyle.Thin, "#1F4E78"))
                    .LabelStyle(style => style.Bold().FontColor("#385723").BackgroundColor("#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#A9D18E"))
                    .ValueStyle(style => style.Border(ExcelBorderLineStyle.Thin, "#D9E2F3"))
                    .FieldValueStyle("LastHighRate", style => style.NumberFormat(ExcelNumberFormat.Number).BackgroundColor("#E2F0D9").Border(ExcelBorderLineStyle.Thin, "#70AD47"))
                    .FieldValueStyle("LastLowRate", style => style.NumberFormat(ExcelNumberFormat.Number).BackgroundColor("#FCE4D6").Border(ExcelBorderLineStyle.Thin, "#C65911"))
                    .FieldValueStyle("LastUpdated", style => style.NumberFormat(ExcelNumberFormat.Date).Border(ExcelBorderLineStyle.Thin, "#D9E2F3"))
                    .FieldLabelStyle("High-Low Spread", style => style.Bold().FontColor("#1F4E78").BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Medium, "#5B9BD5"))
                    .FieldValueStyle("High-Low Spread", style => style.NumberFormat(ExcelNumberFormat.Number).Bold().BackgroundColor("#DDEBF7").Border(ExcelBorderLineStyle.Medium, "#5B9BD5"));
            })
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_25_Export_object_collection_as_repeating_currency_cards), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("USD", GetCellText(bytes, "Currency Cards", "A4"));
        Assert.Equal("GBP", GetCellText(bytes, "Currency Cards", "F4"));
        Assert.Equal("EUR", GetCellText(bytes, "Currency Cards", "K4"));
        Assert.Equal("JPY", GetCellText(bytes, "Currency Cards", "A11"));
        Assert.Equal("Summary", GetCellText(bytes, "Currency Cards", "F25"));
        Assert.Equal("B6-B7", GetCellFormula(bytes, "Currency Cards", "B9"));
        Assert.Equal("AVERAGE(B6,G6,L6,B13,G13,L13,B20,G20,L20,B27)", GetCellFormula(bytes, "Currency Cards", "G27"));
        Assert.Equal("G27-G28", GetCellFormula(bytes, "Currency Cards", "G30"));
        Assert.True(CountStyledCells(bytes, "Currency Cards") > 80);
    }
}
