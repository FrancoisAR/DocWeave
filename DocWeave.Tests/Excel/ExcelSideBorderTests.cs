using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

public sealed class ExcelSideBorderTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Financial_report_keeps_first_column_style_and_totals_borders(bool templateMode)
    {
        var rows = Enumerable.Range(1, 20).Select(i => new { Id = $"ID-{i}", Cost = 12m * i, Revenue = 20m * i }).ToArray();
        const string template = """
        {
          "styles": {
            "id": { "font": { "bold": true }, "fill": { "color": "#E2F0D9" } },
            "totals": { "numberFormat": "currency", "border": { "top": { "style": "thin" }, "bottom": { "style": "double", "color": "#0000FF" } } },
            "totalsId": { "font": { "bold": true }, "fill": { "color": "#E2F0D9" }, "border": { "top": { "style": "thin" }, "bottom": { "style": "double", "color": "#0000FF" } } }
          },
          "sheetTemplates": { "t": { "regions": [
            { "type": "table", "source": "rows", "startCell": "A1", "totals": { "label": "Totals", "style": "totals" } },
            { "type": "rangeStyle", "range": "A2:A21", "style": "id" },
            { "type": "rangeStyle", "range": "A22:A22", "style": "totalsId" }
          ] } },
          "sheets": [ { "name": "Data", "template": "t" } ]
        }
        """;
        var bytes = templateMode
            ? ExcelTemplateWriter.FromJson(template).WithData("rows", rows).ToBytes()
            : ExcelWriter.Create().AddSheet("Data", sheet =>
            {
                sheet.AddTable(rows).AddTotals.Bottom("Totals", s => s.NumberFormat(ExcelNumberFormat.Currency)
                    .BorderTop().BorderBottom(ExcelBorderLineStyle.Double, "#0000FF"));
                sheet.RangeStyle("A2:A21", s => s.Bold().BackgroundColor("#E2F0D9"));
                sheet.RangeStyle("A22:A22", s => s.Bold().BackgroundColor("#E2F0D9")
                    .BorderTop().BorderBottom(ExcelBorderLineStyle.Double, "#0000FF"));
            }).ToBytes();
        Xlsx.AssertValid(bytes);
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbook = document.WorkbookPart!;
        var styles = workbook.WorkbookStylesPart!.Stylesheet;
        var cells = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
        CellFormat FormatFor(string address)
        {
            var cell = cells.Single(c => c.CellReference!.Value == address);
            return styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
        }
        for (int row = 2; row <= 22; row++)
        {
            var format = FormatFor($"A{row}");
            var font = styles.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
            var fill = styles.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value);
            Assert.NotNull(font.Bold);
            Assert.Equal("FFE2F0D9", fill.PatternFill!.ForegroundColor!.Rgb!.Value);
        }
        foreach (var address in new[] { "A22", "B22", "C22" })
        {
            var format = FormatFor(address);
            var border = styles.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value);
            Assert.Equal(BorderStyleValues.Thin, border.TopBorder!.Style!.Value);
            Assert.Equal(BorderStyleValues.Double, border.BottomBorder!.Style!.Value);
            Assert.Equal("FF0000FF", border.BottomBorder.Color!.Rgb!.Value);
            Assert.Equal(BorderStyleValues.None, border.LeftBorder!.Style!.Value);
        }
        Assert.Equal("SUM(B2:B21)", cells.Single(c => c.CellReference!.Value == "B22").CellFormula!.Text);
    }
    [Fact]
    public void Totals_support_independent_borders_and_colors()
    {
        var bytes = ExcelWriter.Create()
            .AddSheet("Data", sheet => sheet.AddTable(new[] { new { Id = 1, Cost = 12m } })
                .AddTotals.Bottom("Totals", s => s.Bold()
                    .BorderTop(ExcelBorderLineStyle.Thin, "#FF0000")
                    .BorderBottom(ExcelBorderLineStyle.Double)
                    .BorderLeft(ExcelBorderLineStyle.Medium)
                    .BorderRight(ExcelBorderLineStyle.Dashed)))
            .ToBytes();

        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbook = document.WorkbookPart!;
        var styles = workbook.WorkbookStylesPart!.Stylesheet;
        var cells = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>();
        foreach (var reference in new[] { "A3", "B3" })
        {
            var cell = cells.Single(c => c.CellReference!.Value == reference);
            var format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            var border = styles.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value);
            Assert.True(format.ApplyBorder!.Value);
            Assert.Equal(BorderStyleValues.Thin, border.TopBorder!.Style!.Value);
            Assert.Equal("FFFF0000", border.TopBorder.Color!.Rgb!.Value);
            Assert.Equal(BorderStyleValues.Double, border.BottomBorder!.Style!.Value);
            Assert.Equal(BorderStyleValues.Medium, border.LeftBorder!.Style!.Value);
            Assert.Equal(BorderStyleValues.Dashed, border.RightBorder!.Style!.Value);
        }
    }

    [Fact]
    public void Border_resets_prior_side_overrides_and_None_removes_one_side()
    {
        var bytes = ExcelWriter.Create().AddSheet("Data", sheet => sheet
            .AddCell("A1", "reset", s => s.BorderBottom(ExcelBorderLineStyle.Double).Border())
            .AddCell("A2", "override", s => s.Border().BorderLeft(ExcelBorderLineStyle.None)))
            .ToBytes();
        using var stream = new MemoryStream(bytes);
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbook = document.WorkbookPart!;
        var styles = workbook.WorkbookStylesPart!.Stylesheet;
        var cells = workbook.WorksheetParts.Single().Worksheet.Descendants<Cell>().ToArray();
        Border BorderFor(Cell cell)
        {
            var format = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
            return styles.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value);
        }
        Assert.Equal(BorderStyleValues.Thin, BorderFor(cells[0]).BottomBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.None, BorderFor(cells[1]).LeftBorder!.Style!.Value);
        Assert.Equal(BorderStyleValues.Thin, BorderFor(cells[1]).RightBorder!.Style!.Value);
    }
}