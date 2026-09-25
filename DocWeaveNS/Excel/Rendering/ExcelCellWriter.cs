using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace DocWeave.Excel;

/// <summary>
/// Creates worksheet cells and applies styles to them. Every cell in a workbook goes through here.
/// </summary>
internal static class ExcelCellWriter
{
    internal static void ApplyRangeStyle(
        SortedDictionary<int, SortedDictionary<int, Cell>> rows,
        ExcelSheetSpec sheet,
        ExcelRangeStyleSpec rangeStyle,
        ExcelStyleRegistry styleRegistry)
    {
        // The range may have been given corner to corner in either direction.
        var startRow = Math.Min(rangeStyle.Start.Row, rangeStyle.End.Row);
        var endRow = Math.Max(rangeStyle.Start.Row, rangeStyle.End.Row);
        var startColumn = Math.Min(rangeStyle.Start.Column, rangeStyle.End.Column);
        var endColumn = Math.Max(rangeStyle.Start.Column, rangeStyle.End.Column);

        for (var row = startRow; row <= endRow; row++)
        {
            for (var column = startColumn; column <= endColumn; column++)
            {
                object? value = null;
                if (rows.TryGetValue(row, out var existingRow) && existingRow.TryGetValue(column, out var existingCell))
                {
                    // A cell that already exists keeps its value and only takes the new style. Missing cells are created empty with the style.
                    existingCell.StyleIndex = styleRegistry.GetStyleIndex(rangeStyle.Style);
                    continue;
                }

                SetCell(rows, sheet, row, column, value, rangeStyle.Style, styleRegistry);
            }
        }
    }

    internal static void SetCell(
        SortedDictionary<int, SortedDictionary<int, Cell>> rows,
        ExcelSheetSpec sheet,
        int rowNumber,
        int columnNumber,
        object? value,
        ExcelStyleSpec style,
        ExcelStyleRegistry styleRegistry)
    {
        if (!rows.TryGetValue(rowNumber, out var row))
        {
            row = [];
            rows[rowNumber] = row;
        }

        // A cell with no style of its own takes the sheet-wide style of its column, if one was set.
        var resolvedStyle = style == ExcelStyleSpec.Empty && sheet.ColumnStyles.TryGetValue(columnNumber, out var columnStyle)
            ? columnStyle
            : style;

        row[columnNumber] = CreateCell(rowNumber, columnNumber, value, styleRegistry.GetStyleIndex(resolvedStyle));
    }

    private static Cell CreateCell(int rowNumber, int columnNumber, object? value, uint styleIndex)
    {
        var cell = new Cell
        {
            CellReference = $"{ExcelCellAddress.ColumnNumberToLetter(columnNumber)}{rowNumber}",
            StyleIndex = styleIndex
        };

        if (value is ExcelFormula formula)
        {
            cell.CellFormula = new CellFormula(formula.FormulaText);
            return cell;
        }

        // A null is written as an empty string cell, not left out, so the cell still carries its style.
        if (value is null)
        {
            cell.DataType = CellValues.String;
            cell.CellValue = new CellValue(string.Empty);
            return cell;
        }

        // Numbers use the invariant culture so the decimal separator is always ".".
        if (IsNumeric(value))
        {
            cell.DataType = CellValues.Number;
            cell.CellValue = new CellValue(Convert.ToString(value, CultureInfo.InvariantCulture));
            return cell;
        }

        if (value is bool boolValue)
        {
            cell.DataType = CellValues.Boolean;
            cell.CellValue = new CellValue(boolValue ? "1" : "0");
            return cell;
        }

        // Inline strings avoid building a shared string table in the first implementation.
        // A shared-string strategy can be added later for very large repeated text exports.
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(FormatText(value)));
        return cell;
    }

    private static bool IsNumeric(object value)
    {
        return value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;
    }

    private static string FormatText(object value)
    {
        return value switch
        {
            // Dates and times are written as ISO 8601 text, not as Excel date serial numbers, so the cell shows text unless the reader converts it.
            DateTime dateTime => dateTime.ToString("O", CultureInfo.InvariantCulture),
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty
        };
    }
}
