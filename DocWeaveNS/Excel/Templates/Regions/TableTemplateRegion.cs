using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "table", "source": "sales", "startCell": "A3", "columns": [...], "styles": {...}, "totals": {...}, "nativeTable": {...} }
/// </summary>
internal sealed class TableTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "table";

    public override bool OccupiesCells => true;

    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        var start = StartOf(region, startOverride);
        var source = templateContext.GetSource(region.GetRequiredString("source"));
        var columnCount = region.TryGetProperty("columns", out var columns) ? columns.GetArrayLength() : source.Columns.Count;
        var rowCount = source.Rows.Count + (region.GetOptionalBool("headers") ?? true ? 1 : 0) + (region.TryGetProperty("totals", out _) ? 1 : 0);
        return new TemplateRegionBounds(start.Row, start.Column, start.Row + Math.Max(rowCount, 1) - 1, start.Column + Math.Max(columnCount, 1) - 1);
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var sourceName = region.GetRequiredString("source");
        var table = new ExcelTableSpec(templateContext.GetSource(sourceName))
        {
            StartCell = StartOf(region, startCell),
            IncludeHeaders = region.GetOptionalBool("headers") ?? true
        };

        if (region.TryGetProperty("columns", out var columns))
        {
            var mappedData = ProjectColumns(table.Data, columns);
            table = new ExcelTableSpec(mappedData)
            {
                StartCell = table.StartCell,
                IncludeHeaders = table.IncludeHeaders
            };

            foreach (var column in columns.EnumerateArray())
            {
                var source = column.GetRequiredString("source");
                if (column.TryGetProperty("caption", out var caption))
                {
                    table.HeaderAliases[source] = caption.GetString() ?? source;
                }

                if (column.TryGetProperty("style", out var style))
                {
                    table.ColumnStyles[source] = templateContext.ResolveStyle(style.GetString());
                }
            }
        }

        if (region.TryGetProperty("headerAliases", out var aliases))
        {
            foreach (var alias in aliases.EnumerateObject())
            {
                table.HeaderAliases[alias.Name] = alias.Value.GetString() ?? alias.Name;
            }
        }

        if (region.TryGetProperty("columnStyles", out var columnStyles))
        {
            foreach (var columnStyle in columnStyles.EnumerateObject())
            {
                table.ColumnStyles[columnStyle.Name] = templateContext.ResolveStyle(columnStyle.Value.GetString());
            }
        }

        if (region.TryGetProperty("styles", out var tableStyles))
        {
            table.HeaderStyle = templateContext.ResolveStyle(tableStyles.GetOptionalString("header"));
            table.DataStyle = templateContext.ResolveStyle(tableStyles.GetOptionalString("data"));
            var alternateStyle = tableStyles.GetOptionalString("alternateRow");
            if (!string.IsNullOrWhiteSpace(alternateStyle))
            {
                table.AlternateRowStyle = templateContext.ResolveStyle(alternateStyle);
            }
        }

        if (region.TryGetProperty("totals", out var totals))
        {
            table.TotalsRow = new ExcelTotalsRowSpec(
                totals.GetOptionalString("label") ?? "Totals",
                templateContext.ResolveStyle(totals.GetOptionalString("style")));
        }

        if (region.TryGetProperty("nativeTable", out var nativeTable))
        {
            table.NativeTable = new ExcelNativeTableSpec(
                nativeTable.GetRequiredString("name"),
                nativeTable.GetOptionalString("style") ?? "TableStyleMedium2",
                nativeTable.GetOptionalBool("showFilterButtons") ?? true,
                nativeTable.GetOptionalBool("showRowStripes") ?? true);
            table.IncludeHeaders = true;
        }

        sheet.Tables.Add(table);
    }

    /// <summary>
    /// Keeps only the listed source columns, in the order the template lists them.
    /// </summary>
    private static TabularData ProjectColumns(TabularData source, JsonElement columns)
    {
        var selectedColumns = columns
            .EnumerateArray()
            .Select(column => column.GetRequiredString("source"))
            .ToList();

        var indexes = selectedColumns.Select(column => ExcelTemplateValues.RequireColumn(source, column)).ToList();
        var rows = source.Rows
            .Select(row => indexes.Select(index => index < row.Count ? row[index] : null).ToList())
            .Cast<IReadOnlyList<object?>>()
            .ToList();

        return new TabularData(selectedColumns, rows);
    }
}
