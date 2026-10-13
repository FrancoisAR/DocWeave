using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "pivotTable", "name": "...", "source": { "sheet": "...", "range": "..." }, "startCell": "A3",
///   "rows": [...], "columns": [...], "filters": [...], "values": [...] }
/// </summary>
internal sealed class PivotTableTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "pivotTable";

    public override bool OccupiesCells => true;

    /// <summary>
    /// The pivot's size depends on its data, so only its top-left cell is reserved.
    /// </summary>
    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        return TemplateRegionBounds.SingleCell(StartOf(region, startOverride));
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var pivot = new ExcelPivotTableSpec(region.GetRequiredString("name"))
        {
            StartCell = StartOf(region, startCell)
        };

        var source = region.GetProperty("source");
        pivot.Source = new ExcelPivotTableSource(
            source.GetRequiredString("sheet"),
            source.GetRequiredString("range"));

        if (region.TryGetProperty("rows", out var rows))
        {
            pivot.RowFields.AddRange(ReadFieldNames(rows));
        }

        if (region.TryGetProperty("columns", out var columns))
        {
            pivot.ColumnFields.AddRange(ReadFieldNames(columns));
        }

        if (region.TryGetProperty("filters", out var filters))
        {
            pivot.FilterFields.AddRange(ReadFieldNames(filters));
        }

        foreach (var value in region.GetProperty("values").EnumerateArray())
        {
            pivot.ValueFields.Add(new ExcelPivotValueFieldSpec(
                value.GetRequiredString("field"),
                value.GetOptionalString("caption") ?? value.GetRequiredString("field"),
                ParseAggregate(value.GetOptionalString("aggregate") ?? "sum")));
        }

        if (region.TryGetProperty("grandTotals", out var grandTotals))
        {
            pivot.RowGrandTotals = grandTotals.GetOptionalBool("rows") ?? pivot.RowGrandTotals;
            pivot.ColumnGrandTotals = grandTotals.GetOptionalBool("columns") ?? pivot.ColumnGrandTotals;
        }

        pivot.StyleName = region.GetOptionalString("style") ?? pivot.StyleName;
        sheet.PivotTables.Add(pivot);
    }

    private static IEnumerable<string> ReadFieldNames(JsonElement fields)
    {
        return fields.EnumerateArray().Select(item => item.GetString()).Where(value => !string.IsNullOrWhiteSpace(value))!;
    }

    private static ExcelPivotAggregateFunction ParseAggregate(string value)
    {
        return value.Equals("min", StringComparison.OrdinalIgnoreCase)
            ? ExcelPivotAggregateFunction.Minimum
            : value.Equals("max", StringComparison.OrdinalIgnoreCase)
                ? ExcelPivotAggregateFunction.Maximum
                : NetStandardEnum.Parse<ExcelPivotAggregateFunction>(value, ignoreCase: true);
    }
}
