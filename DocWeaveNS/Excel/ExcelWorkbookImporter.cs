using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

/// <summary>
/// Reads worksheet rows from an .xlsx package. The sheet is read row by row with an <see cref="OpenXmlReader"/>, so only
/// one row of the sheet is held as an object tree at a time.
/// </summary>
internal static class ExcelWorkbookImporter
{
    /// <summary>
    /// Reads the whole sheet into memory. The header row's width is the widest row of the sheet, so a data row that is wider than
    /// the header row still gets a column, named Column N.
    /// </summary>
    public static ExcelDocument Import(
        Stream stream,
        string? sheetName,
        int? sheetIndex,
        ExcelCellAddress startCell,
        bool hasHeaderRow,
        ColumnSelection selection,
        string? label,
        CancellationToken cancellationToken = default)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("The workbook does not contain a workbook part.");
        var worksheetPart = ResolveWorksheetPart(workbookPart, sheetName, sheetIndex);
        var rows = ReadRows(workbookPart, worksheetPart, startCell, cancellationToken).ToList();

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

    /// <summary>
    /// Reads the rows one at a time. The workbook is open while the sequence is enumerated and is closed when enumeration ends
    /// or the enumerator is disposed. Unlike <see cref="Import"/>, the width comes from the first row, because the later rows
    /// have not been read yet.
    /// </summary>
    /// <param name="stream">A seekable stream holding the .xlsx file. It is read, not disposed.</param>
    /// <param name="sheetName">The sheet to read, or null to use <paramref name="sheetIndex"/>.</param>
    /// <param name="sheetIndex">The zero-based sheet position, used when no name is given. Null means the first sheet.</param>
    /// <param name="startCell">The top-left cell of the range to read.</param>
    /// <param name="hasHeaderRow">Whether the first row at or below the start cell holds the column headers.</param>
    /// <param name="selection">Which columns to keep.</param>
    /// <param name="label">A source name copied into each row.</param>
    /// <param name="onHeaders">
    /// Called once, before the first data row is returned, with the headers that are kept. Called with an empty list when
    /// the sheet has no rows.
    /// </param>
    /// <param name="cancellationToken">Checked for each worksheet row.</param>
    public static IEnumerable<ExcelRow> Enumerate(
        Stream stream,
        string? sheetName,
        int? sheetIndex,
        ExcelCellAddress startCell,
        bool hasHeaderRow,
        ColumnSelection selection,
        string? label,
        Action<IReadOnlyList<string>> onHeaders,
        CancellationToken cancellationToken)
    {
        using var document = SpreadsheetDocument.Open(stream, false);
        var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("The workbook does not contain a workbook part.");
        var worksheetPart = ResolveWorksheetPart(workbookPart, sheetName, sheetIndex);

        List<string>? sourceHeaders = null;
        IReadOnlyList<int> selectedIndexes = [];

        // A cell to the right of the header row has no header of its own. It is kept, named Column N, only when no column
        // selection is configured, because a selection says exactly which columns are wanted.
        var keepExtraColumns = !selection.HasRules;

        foreach (var row in ReadRows(workbookPart, worksheetPart, startCell, cancellationToken))
        {
            if (sourceHeaders is null)
            {
                var width = (row.Values.Count == 0 ? 0 : row.Values.Keys.Max()) + 1;
                sourceHeaders = hasHeaderRow
                    ? BuildHeaders(row.Values, width)
                    : Enumerable.Range(0, width).Select(index => ExcelCellAddress.ColumnNumberToLetter(startCell.Column + index)).ToList();
                selectedIndexes = selection.Resolve(sourceHeaders);
                onHeaders(selectedIndexes.Select(index => sourceHeaders[index]).ToList());

                if (hasHeaderRow)
                {
                    continue;
                }
            }

            var values = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
            foreach (var index in selectedIndexes)
            {
                values[sourceHeaders[index]] = row.Values.TryGetValue(index, out var value) ? value : null;
            }

            if (keepExtraColumns)
            {
                foreach (var item in row.Values)
                {
                    var index = item.Key;
                    var value = item.Value;
                    if (index >= sourceHeaders.Count)
                    {
                        var name = hasHeaderRow
                            ? $"Column{index + 1}"
                            : ExcelCellAddress.ColumnNumberToLetter(startCell.Column + index);
                        values[name] = value;
                    }
                }
            }

            yield return new ExcelRow(label, row.RowNumber, values);
        }

        if (sourceHeaders is null)
        {
            onHeaders([]);
        }
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

    /// <summary>
    /// Streams the rows at or below the start row. A row with no cell at or right of the start column is left out.
    /// </summary>
    private static IEnumerable<ImportedWorksheetRow> ReadRows(
        WorkbookPart workbookPart,
        WorksheetPart worksheetPart,
        ExcelCellAddress startCell,
        CancellationToken cancellationToken)
    {
        var sharedStrings = LoadSharedStrings(workbookPart);

        using var reader = OpenXmlReader.Create(worksheetPart);
        while (reader.Read())
        {
            if (reader.ElementType != typeof(Row) || !reader.IsStartElement)
            {
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();

            // Rows above the start row are passed over without building them, which matters for sheets with a long preamble.
            // (Do not call reader.Skip() here: it already moves to the next row, and the next Read() would then drop that row.)
            var rowIndex = ReadRowIndex(reader);
            if (rowIndex is null || rowIndex < startCell.Row)
            {
                continue;
            }

            var row = (Row)reader.LoadCurrentElement()!;
            var values = new SortedDictionary<int, string?>();
            foreach (var cell in row.Elements<Cell>())
            {
                // Cells are placed by their reference, not by their position in the row, because Excel leaves empty cells out.
                var address = ParseCellReference(cell.CellReference?.Value);
                if (address is null || address.Column < startCell.Column)
                {
                    continue;
                }

                values[address.Column - startCell.Column] = ReadCellValue(sharedStrings, cell);
            }

            if (values.Count > 0)
            {
                yield return new ImportedWorksheetRow(rowIndex.Value, values);
            }
        }
    }

    /// <summary>
    /// Reads the shared string table once, into a list, so a cell can be looked up by index. Null when the workbook has none.
    /// </summary>
    private static List<string>? LoadSharedStrings(WorkbookPart workbookPart)
    {
        var part = workbookPart.SharedStringTablePart;
        if (part is null)
        {
            return null;
        }

        var strings = new List<string>();
        using var reader = OpenXmlReader.Create(part);
        while (reader.Read())
        {
            if (reader.ElementType == typeof(SharedStringItem) && reader.IsStartElement)
            {
                strings.Add(reader.LoadCurrentElement()!.InnerText);
            }
        }

        return strings;
    }

    private static int? ReadRowIndex(OpenXmlReader reader)
    {
        foreach (var attribute in reader.Attributes)
        {
            if (attribute.LocalName == "r" && int.TryParse(attribute.Value, out var rowIndex))
            {
                return rowIndex;
            }
        }

        return null;
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

    private static string? ReadCellValue(List<string>? sharedStrings, Cell cell)
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
                return sharedStrings?[sharedStringIndex];
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
