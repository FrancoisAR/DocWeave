using System.Globalization;
using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "summary", "source": "sales", "startCell": "E3", "fields": [ { "caption": "Total", "aggregate": "sum", "source": "Income" } ] }
/// A caption/value list of figures calculated from a data source.
/// </summary>
internal sealed class SummaryTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "summary";

    public override bool OccupiesCells => true;

    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        var start = StartOf(region, startOverride);
        var rows = region.GetProperty("fields").GetArrayLength();
        return new TemplateRegionBounds(start.Row, start.Column, start.Row + Math.Max(rows, 1) - 1, start.Column + 1);
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var source = templateContext.GetSource(region.GetRequiredString("source"));
        var start = StartOf(region, startCell);
        var labelStyle = templateContext.ResolveStyle(region.GetOptionalString("labelStyle"));
        var valueStyle = templateContext.ResolveStyle(region.GetOptionalString("valueStyle"));
        var rowOffset = 0;

        foreach (var field in region.GetProperty("fields").EnumerateArray())
        {
            var caption = field.GetRequiredString("caption");
            object? value;
            if (field.TryGetProperty("value", out var fixedValue))
            {
                value = templateContext.BindValue(ExcelTemplateValues.ReadJsonValue(fixedValue));
            }
            else
            {
                value = CalculateAggregate(source, field.GetRequiredString("aggregate"), field.GetRequiredString("source"));
            }

            sheet.Cells.Add(new ExcelCellSpec(start.Offset(rowOffset, 0), caption, labelStyle));
            sheet.Cells.Add(new ExcelCellSpec(start.Offset(rowOffset, 1), value, valueStyle));
            rowOffset++;
        }
    }

    private static object? CalculateAggregate(TabularData source, string aggregate, string fieldName)
    {
        var index = ExcelTemplateValues.RequireColumn(source, fieldName);
        var values = source.Rows
            .Select(row => index < row.Count ? row[index] : null)
            .Where(value => value is not null)
            .Select(value => Convert.ToDecimal(value, CultureInfo.InvariantCulture))
            .ToList();

        if (aggregate.Equals("count", StringComparison.OrdinalIgnoreCase))
        {
            return values.Count;
        }

        if (values.Count == 0)
        {
            return null;
        }

        return aggregate.ToLowerInvariant() switch
        {
            "sum" => values.Sum(),
            "avg" or "average" => values.Average(),
            "min" => values.Min(),
            "max" => values.Max(),
            _ => throw new InvalidOperationException($"Summary aggregate '{aggregate}' is not supported.")
        };
    }
}
