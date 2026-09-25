using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Builds a one-column workbook by hand whose text lives in the shared string table, the way Excel itself writes it. DocWeave's
/// own exports use inline strings, so they cannot exercise that path.
/// </summary>
internal static class SharedStringWorkbook
{
    /// <param name="sharedStrings">The entries of the shared string table.</param>
    /// <param name="cells">Pairs of (row number, index into <paramref name="sharedStrings"/>), written to column A.</param>
    public static byte[] Create(IReadOnlyList<string> sharedStrings, IEnumerable<(int Row, int StringIndex)> cells)
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var stringsPart = workbookPart.AddNewPart<SharedStringTablePart>();
            stringsPart.SharedStringTable = new SharedStringTable(
                sharedStrings.Select(text => new SharedStringItem(new Text(text))));
            stringsPart.SharedStringTable.Save();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            foreach (var (row, stringIndex) in cells)
            {
                sheetData.Append(new Row(
                    new Cell
                    {
                        CellReference = "A" + row,
                        DataType = CellValues.SharedString,
                        CellValue = new CellValue(stringIndex.ToString())
                    })
                { RowIndex = (uint)row });
            }

            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Save();

            workbookPart.Workbook.AppendChild(new Sheets(new Sheet
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            }));
            workbookPart.Workbook.Save();
        }

        return stream.ToArray();
    }
}
