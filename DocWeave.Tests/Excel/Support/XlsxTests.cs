using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

/// <summary>
/// The workbook inspector the export tests rely on. If it stopped noticing problems, every export test would pass by accident.
/// </summary>
public sealed class XlsxTests
{
    [Fact]
    public void A_valid_workbook_passes_and_has_no_validation_errors()
    {
        var bytes = CreateWorkbook();

        Xlsx.AssertValid(bytes);

        Assert.Empty(Xlsx.GetValidationErrors(bytes));
        Assert.Equal(0, Xlsx.CountOpenXmlValidationErrors(bytes));
    }

    [Fact]
    public void A_broken_workbook_fails_and_the_message_says_what_is_wrong()
    {
        var bytes = Break(CreateWorkbook());

        var exception = Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => Xlsx.AssertValid(bytes));

        // The message names the problem and where it is, not just how many there were.
        Assert.Contains("1 OpenXML validation error", exception.Message);
        Assert.Contains("unexpected child element", exception.Message);
        Assert.Contains("sheetData", exception.Message);
        Assert.Contains("/x:worksheet", exception.Message);
        Assert.Equal(1, Xlsx.CountOpenXmlValidationErrors(bytes));
    }

    [Fact]
    public void Inspector_reads_cells_formulas_and_sheets_back()
    {
        var bytes = ExcelWriter.Create()
            .AddSheet("First", sheet => sheet
                .AddCell("A1", "hello")
                .AddFormulaCell("B1", "SUM(1,2)"))
            .AddSheet("Second", sheet => sheet.AddCell("A1", "x"))
            .ToBytes();

        Assert.Equal(2, Xlsx.CountWorksheets(bytes));
        Assert.Equal("hello", Xlsx.GetCellText(bytes, "First", "A1"));
        Assert.Equal("SUM(1,2)", Xlsx.GetCellFormula(bytes, "First", "B1"));
        Assert.Equal(1, Xlsx.CountFormulaCells(bytes, "First"));
        Assert.Equal(0, Xlsx.CountFormulaCells(bytes, "Second"));
    }

    private static byte[] CreateWorkbook()
    {
        return ExcelWriter.Create()
            .AddSheet("Data", sheet => sheet.AddCell("A1", "value"))
            .ToBytes();
    }

    /// <summary>
    /// Adds a second sheetData element to the first worksheet, which the schema does not allow.
    /// </summary>
    private static byte[] Break(byte[] workbookBytes)
    {
        using var stream = new MemoryStream();
        stream.Write(workbookBytes);
        using (var document = SpreadsheetDocument.Open(stream, true))
        {
            var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
            worksheetPart.Worksheet.Append(new SheetData());
            worksheetPart.Worksheet.Save();
        }

        return stream.ToArray();
    }
}
