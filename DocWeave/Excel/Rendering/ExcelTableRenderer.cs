using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

/// <summary>
/// Writes table regions (headers, rows, totals, formula columns) and repeater card layouts into the row map.
/// </summary>
internal static class ExcelTableRenderer
{
    internal static void WriteTable(SortedDictionary<int, SortedDictionary<int, Cell>> rows, ExcelSheetSpec sheet, ExcelTableSpec table, ExcelStyleRegistry styleRegistry)
    {
        var rowOffset = 0;
        // Formula columns come after the data columns.
        var columns = table.Data.Columns.Concat(table.FormulaColumns.Select(column => column.ColumnName)).ToList();

        if (table.IncludeHeaders)
        {
            for (var columnIndex = 0; columnIndex < columns.Count; columnIndex++)
            {
                var sourceName = columns[columnIndex];
                var caption = table.HeaderAliases.TryGetValue(sourceName, out var alias) ? alias : sourceName;
                ExcelCellWriter.SetCell(rows, sheet, table.StartCell.Row, ResolveColumnNumber(table, sourceName, columnIndex), caption, table.HeaderStyle, styleRegistry);
            }

            rowOffset++;
        }

        for (var rowIndex = 0; rowIndex < table.Data.Rows.Count; rowIndex++)
        {
            var values = table.Data.Rows[rowIndex];
            var worksheetRow = table.StartCell.Row + rowOffset + rowIndex;
            // The same row values feed the style predicates and, through the context, the formula factories.
            var rowValues = BuildRowValues(table, values, rowIndex, worksheetRow);

            for (var columnIndex = 0; columnIndex < table.Data.Columns.Count; columnIndex++)
            {
                var sourceName = table.Data.Columns[columnIndex];
                var value = columnIndex < values.Count ? values[columnIndex] : null;
                var style = ResolveDataStyle(table, sourceName, value, rowIndex, rowValues);
                ExcelCellWriter.SetCell(rows, sheet, worksheetRow, ResolveColumnNumber(table, sourceName, columnIndex), value, style, styleRegistry);
            }

            for (var formulaIndex = 0; formulaIndex < table.FormulaColumns.Count; formulaIndex++)
            {
                var formulaColumn = table.FormulaColumns[formulaIndex];
                var columnIndex = table.Data.Columns.Count + formulaIndex;
                var style = ResolveFormulaStyle(table, formulaColumn.ColumnName, rowIndex, rowValues);
                // The factory gets the worksheet row so the formula can refer to cells on the same row.
                var formula = formulaColumn.FormulaFactory(new ExcelFormulaContext(rowIndex, worksheetRow));
                ExcelCellWriter.SetCell(rows, sheet, worksheetRow, ResolveColumnNumber(table, formulaColumn.ColumnName, columnIndex), new ExcelFormula(formula), style, styleRegistry);
            }
        }

        if (table.TotalsRow is not null)
        {
            WriteTotalsRow(rows, sheet, table, columns, rowOffset, styleRegistry);
        }
    }

    internal static void WriteRepeater(
        SortedDictionary<int, SortedDictionary<int, Cell>> rows,
        ExcelSheetSpec sheet,
        ExcelRepeaterSpec repeater,
        ExcelStyleRegistry styleRegistry)
    {
        // With no fields declared, every column except the title column is shown, captioned with its own name.
        var fields = repeater.Fields.Count > 0
            ? repeater.Fields
            : repeater.Data.Columns
                .Where(column => !string.Equals(column, repeater.TitleField, StringComparison.OrdinalIgnoreCase))
                .Select(column => new ExcelRepeaterFieldSpec(column, column))
                .ToList();

        for (var itemIndex = 0; itemIndex < repeater.Data.Rows.Count; itemIndex++)
        {
            // Cards are laid out left to right, wrapping after ItemsPerRow, each card ItemWidth by ItemHeight cells plus the gaps.
            var itemRow = itemIndex / repeater.ItemsPerRow;
            var itemColumn = itemIndex % repeater.ItemsPerRow;
            var top = repeater.StartCell.Row + itemRow * (repeater.ItemHeight + repeater.RowGap);
            var left = repeater.StartCell.Column + itemColumn * (repeater.ItemWidth + repeater.ColumnGap);
            var values = BuildRepeaterValues(repeater, repeater.Data.Rows[itemIndex], itemIndex, top, left);

            ApplyRepeaterCardStyle(rows, sheet, repeater, top, left, styleRegistry);
            WriteRepeaterTitle(rows, sheet, repeater, values, top, left, styleRegistry);

            for (var fieldIndex = 0; fieldIndex < fields.Count; fieldIndex++)
            {
                var worksheetRow = top + 1 + fieldIndex;
                // Fields that do not fit in the card's height are dropped.
                if (worksheetRow >= top + repeater.ItemHeight)
                {
                    break;
                }

                var field = fields[fieldIndex];
                var value = ResolveRepeaterFieldValue(field, values, itemIndex, top, left);
                ExcelCellWriter.SetCell(rows, sheet, worksheetRow, left, field.Caption, field.LabelStyle ?? repeater.LabelStyle, styleRegistry);
                ExcelCellWriter.SetCell(rows, sheet, worksheetRow, left + 1, value, field.ValueStyle ?? repeater.ValueStyle, styleRegistry);
            }
        }
    }

    private static object? ResolveRepeaterFieldValue(
        ExcelRepeaterFieldSpec field,
        IReadOnlyDictionary<string, object?> values,
        int itemIndex,
        int top,
        int left)
    {
        if (field.FormulaFactory is not null)
        {
            var formula = field.FormulaFactory(new ExcelRepeaterFormulaContext(itemIndex, top, left, values));
            return new ExcelFormula(formula);
        }

        values.TryGetValue(field.ColumnName, out var value);
        return value;
    }

    private static IReadOnlyDictionary<string, object?> BuildRepeaterValues(
        ExcelRepeaterSpec repeater,
        IReadOnlyList<object?> sourceValues,
        int itemIndex,
        int top,
        int left)
    {
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            // These double-underscore keys expose position information to the formula and style callbacks. They sit in the same dictionary as the data columns.
            ["__DataIndex"] = itemIndex,
            ["__DataRowNumber"] = itemIndex + 1,
            ["__WorksheetRow"] = top,
            ["__WorksheetColumn"] = left
        };

        for (var columnIndex = 0; columnIndex < repeater.Data.Columns.Count; columnIndex++)
        {
            values[repeater.Data.Columns[columnIndex]] = columnIndex < sourceValues.Count ? sourceValues[columnIndex] : null;
        }

        return values;
    }

    private static void ApplyRepeaterCardStyle(
        SortedDictionary<int, SortedDictionary<int, Cell>> rows,
        ExcelSheetSpec sheet,
        ExcelRepeaterSpec repeater,
        int top,
        int left,
        ExcelStyleRegistry styleRegistry)
    {
        if (repeater.CardStyle == ExcelStyleSpec.Empty)
        {
            return;
        }

        for (var row = top; row < top + repeater.ItemHeight; row++)
        {
            for (var column = left; column < left + repeater.ItemWidth; column++)
            {
                ExcelCellWriter.SetCell(rows, sheet, row, column, null, repeater.CardStyle, styleRegistry);
            }
        }
    }

    private static void WriteRepeaterTitle(
        SortedDictionary<int, SortedDictionary<int, Cell>> rows,
        ExcelSheetSpec sheet,
        ExcelRepeaterSpec repeater,
        IReadOnlyDictionary<string, object?> values,
        int top,
        int left,
        ExcelStyleRegistry styleRegistry)
    {
        if (string.IsNullOrWhiteSpace(repeater.TitleField))
        {
            return;
        }

        values.TryGetValue(repeater.TitleField, out var title);
        ExcelCellWriter.SetCell(rows, sheet, top, left, title, repeater.TitleStyle, styleRegistry);
    }

    private static IReadOnlyDictionary<string, object?> BuildRowValues(ExcelTableSpec table, IReadOnlyList<object?> values, int rowIndex, int worksheetRow)
    {
        var rowValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            // These double-underscore keys expose position information to the style callbacks. They sit in the same dictionary as the data columns.
            ["__DataIndex"] = rowIndex,
            ["__DataRowNumber"] = rowIndex + 1,
            ["__WorksheetRow"] = worksheetRow
        };

        for (var columnIndex = 0; columnIndex < table.Data.Columns.Count; columnIndex++)
        {
            rowValues[table.Data.Columns[columnIndex]] = columnIndex < values.Count ? values[columnIndex] : null;
        }

        return rowValues;
    }

    private static void WriteTotalsRow(
        SortedDictionary<int, SortedDictionary<int, Cell>> rows,
        ExcelSheetSpec sheet,
        ExcelTableSpec table,
        IReadOnlyList<string> columns,
        int dataRowOffset,
        ExcelStyleRegistry styleRegistry)
    {
        var totalRow = table.StartCell.Row + dataRowOffset + table.Data.Rows.Count;
        var firstDataRow = table.StartCell.Row + dataRowOffset;
        // The totals row goes directly below the last data row.
        var lastDataRow = totalRow - 1;
        var totalsStyle = table.TotalsRow!.Style;

        if (columns.Count == 0)
        {
            return;
        }

        // The first column holds the label. Every other column gets a SUM formula if it should be totalled.
        ExcelCellWriter.SetCell(rows, sheet, totalRow, ResolveColumnNumber(table, columns[0], 0), table.TotalsRow.Label, totalsStyle, styleRegistry);

        if (table.Data.Rows.Count == 0)
        {
            return;
        }

        for (var columnIndex = 1; columnIndex < columns.Count; columnIndex++)
        {
            var sourceName = columns[columnIndex];
            if (!ShouldTotalColumn(table, sourceName, columnIndex))
            {
                continue;
            }

            var columnLetter = ExcelCellAddress.ColumnNumberToLetter(ResolveColumnNumber(table, sourceName, columnIndex));
            ExcelCellWriter.SetCell(rows, sheet, totalRow, ResolveColumnNumber(table, sourceName, columnIndex), new ExcelFormula($"SUM({columnLetter}{firstDataRow}:{columnLetter}{lastDataRow})"), totalsStyle, styleRegistry);
        }
    }

    private static bool ShouldTotalColumn(ExcelTableSpec table, string sourceName, int columnIndex)
    {
        // Formula columns are always totalled. Data columns only when at least one value in them is numeric.
        if (columnIndex >= table.Data.Columns.Count)
        {
            return true;
        }

        return table.Data.Rows.Any(row => columnIndex < row.Count && row[columnIndex] is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal);
    }

    private static ExcelStyleSpec ResolveDataStyle(
        ExcelTableSpec table,
        string sourceName,
        object? value,
        int rowIndex,
        IReadOnlyDictionary<string, object?> rowValues)
    {
        // Style precedence, first match wins: cell rule (sees the whole row), value rule for this column, fixed row style,
        // row rule, column style, then the alternate row style on odd rows, then the default data style.
        foreach (var cellStyle in table.ConditionalCellStyles)
        {
            if (string.Equals(cellStyle.ColumnName, sourceName, StringComparison.OrdinalIgnoreCase)
                && cellStyle.Predicate(rowValues))
            {
                return cellStyle.Style;
            }
        }

        foreach (var conditionalStyle in table.ConditionalStyles)
        {
            if (string.Equals(conditionalStyle.ColumnName, sourceName, StringComparison.OrdinalIgnoreCase)
                && conditionalStyle.Predicate(value))
            {
                return conditionalStyle.Style;
            }
        }

        if (table.RowStyles.TryGetValue(rowIndex, out var rowStyle))
        {
            return rowStyle;
        }

        foreach (var conditionalRowStyle in table.ConditionalRowStyles)
        {
            if (conditionalRowStyle.Predicate(rowValues))
            {
                return conditionalRowStyle.Style;
            }
        }

        if (table.ColumnStyles.TryGetValue(sourceName, out var columnStyle))
        {
            return columnStyle;
        }

        return rowIndex % 2 == 1 && table.AlternateRowStyle is not null
            ? table.AlternateRowStyle
            : table.DataStyle;
    }

    private static ExcelStyleSpec ResolveFormulaStyle(ExcelTableSpec table, string sourceName, int rowIndex, IReadOnlyDictionary<string, object?> rowValues)
    {
        // Same order as ResolveDataStyle, minus the value rules: a formula cell has no source value to test.
        foreach (var cellStyle in table.ConditionalCellStyles)
        {
            if (string.Equals(cellStyle.ColumnName, sourceName, StringComparison.OrdinalIgnoreCase)
                && cellStyle.Predicate(rowValues))
            {
                return cellStyle.Style;
            }
        }

        if (table.RowStyles.TryGetValue(rowIndex, out var rowStyle))
        {
            return rowStyle;
        }

        foreach (var conditionalRowStyle in table.ConditionalRowStyles)
        {
            if (conditionalRowStyle.Predicate(rowValues))
            {
                return conditionalRowStyle.Style;
            }
        }

        if (table.ColumnStyles.TryGetValue(sourceName, out var columnStyle))
        {
            return columnStyle;
        }

        return rowIndex % 2 == 1 && table.AlternateRowStyle is not null
            ? table.AlternateRowStyle
            : table.DataStyle;
    }

    internal static int ResolveColumnNumber(ExcelTableSpec table, string sourceName, int sourceIndex)
    {
        // Explicit column placement lets templates behave like a light schema: fields can
        // land in fixed worksheet columns while the source order remains independent.
        return table.ColumnLocations.TryGetValue(sourceName, out var columnNumber)
            ? columnNumber
            : table.StartCell.Column + sourceIndex;
    }
}
