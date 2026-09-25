using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

internal static class ExcelWorkbookImporter
{
    public static ExcelDocument Import(
        Stream stream,
        string? sheetName,
        int? sheetIndex,
        ExcelCellAddress startCell,
        bool hasHeaderRow,
        ColumnSelection selection,
        string? label)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("The workbook does not contain a workbook part.");
        var worksheetPart = ResolveWorksheetPart(workbookPart, sheetName, sheetIndex);
        var rows = ReadRows(workbookPart, worksheetPart, startCell);

        if (rows.Count == 0)
        {
            return new ExcelDocument([], []);
        }

        var width = rows.Max(row => row.Values.Count == 0 ? 0 : row.Values.Keys.Max()) + 1;
        var sourceHeaders = hasHeaderRow
            ? BuildHeaders(rows[0].Values, width)
            : Enumerable.Range(0, width).Select(index => ExcelCellAddress.ColumnNumberToLetter(startCell.Column + index)).ToList();

        var selectedIndexes = selection.Resolve(sourceHeaders);
        var selectedHeaders = selectedIndexes.Select(index => sourceHeaders[index]).ToList();
        var dataRows = hasHeaderRow ? rows.Skip(1) : rows;
        var outputRows = new List<ExcelRow>();

        foreach (var row in dataRows)
        {
            var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            foreach (var index in selectedIndexes)
            {
                var header = sourceHeaders[index];
                values[header] = row.Values.TryGetValue(index, out var value) ? value : null;
            }

            outputRows.Add(new ExcelRow(label, row.RowNumber, values));
        }

        return new ExcelDocument(selectedHeaders, outputRows);
    }

    private static WorksheetPart ResolveWorksheetPart(WorkbookPart workbookPart, string? sheetName, int? sheetIndex)
    {
        var workbookSheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
            ?? throw new InvalidOperationException("The workbook does not contain any sheets.");

        Sheet? sheet;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            sheet = workbookSheets.FirstOrDefault(item => string.Equals(item.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase));
            if (sheet is null)
            {
                throw new InvalidOperationException($"Worksheet '{sheetName}' was not found.");
            }
        }
        else
        {
            var index = sheetIndex ?? 0;
            if (index >= workbookSheets.Count)
            {
                throw new InvalidOperationException($"Worksheet index {index} was not found.");
            }

            sheet = workbookSheets[index];
        }

        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    }

    private static IReadOnlyList<ImportedWorksheetRow> ReadRows(WorkbookPart workbookPart, WorksheetPart worksheetPart, ExcelCellAddress startCell)
    {
        var result = new List<ImportedWorksheetRow>();
        foreach (var row in worksheetPart.Worksheet.Descendants<Row>().Where(row => row.RowIndex?.Value >= startCell.Row))
        {
            var values = new SortedDictionary<int, string?>();
            foreach (var cell in row.Elements<Cell>())
            {
                var address = ParseCellReference(cell.CellReference?.Value);
                if (address is null || address.Column < startCell.Column)
                {
                    continue;
                }

                values[address.Column - startCell.Column] = ReadCellValue(workbookPart, cell);
            }

            if (values.Count > 0)
            {
                result.Add(new ImportedWorksheetRow((int)(row.RowIndex?.Value ?? 0), values));
            }
        }

        return result;
    }

    private static List<string> BuildHeaders(IReadOnlyDictionary<int, string?> headerValues, int width)
    {
        var headers = new List<string>();
        for (var index = 0; index < width; index++)
        {
            var value = headerValues.TryGetValue(index, out var header) ? header : null;
            headers.Add(string.IsNullOrWhiteSpace(value) ? $"Column{index + 1}" : value);
        }

        return headers;
    }

    private static string? ReadCellValue(WorkbookPart workbookPart, Cell cell)
    {
        var rawValue = cell.CellValue?.InnerText;

        if (cell.DataType is null)
        {
            return rawValue;
        }

        if (cell.DataType.Value == CellValues.SharedString)
        {
            if (int.TryParse(rawValue, out var sharedStringIndex))
            {
                return workbookPart.SharedStringTablePart?.SharedStringTable.ElementAt(sharedStringIndex).InnerText;
            }

            return rawValue;
        }

        if (cell.DataType.Value == CellValues.InlineString)
        {
            return cell.InlineString?.InnerText ?? rawValue;
        }

        if (cell.DataType.Value == CellValues.Boolean)
        {
            return rawValue == "1" ? "TRUE" : "FALSE";
        }

        return rawValue;
    }

    private static ExcelCellAddress? ParseCellReference(string? cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
        {
            return null;
        }

        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(cellReference.SkipWhile(char.IsLetter).TakeWhile(char.IsDigit).ToArray());
        if (letters.Length == 0 || digits.Length == 0)
        {
            return null;
        }

        return new ExcelCellAddress(int.Parse(digits), ExcelCellAddress.ColumnLetterToNumber(letters));
    }

    private sealed record ImportedWorksheetRow(int RowNumber, IReadOnlyDictionary<int, string?> Values);
}
