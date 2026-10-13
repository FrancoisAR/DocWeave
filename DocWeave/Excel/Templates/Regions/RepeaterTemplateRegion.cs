using System.Text.Json;
using System.Text.RegularExpressions;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "repeater", "source": "currencies", "startCell": "A4", "itemsPerRow": 3, "itemSize": {...}, "fields": [...] }
/// One rectangular card per source row.
/// </summary>
internal sealed partial class RepeaterTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "repeater";

    public override bool OccupiesCells => true;

    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        var start = StartOf(region, startOverride);
        var source = templateContext.GetSource(region.GetRequiredString("source"));
        var itemsPerRow = region.GetOptionalInt32("itemsPerRow") ?? 3;
        var itemSize = region.GetProperty("itemSize");
        var itemWidth = itemSize.GetRequiredInt32("columns");
        var itemHeight = itemSize.GetRequiredInt32("rows");
        var columnGap = region.TryGetProperty("gap", out var gap) ? gap.GetOptionalInt32("columns") ?? 1 : 1;
        var rowGap = region.TryGetProperty("gap", out gap) ? gap.GetOptionalInt32("rows") ?? 1 : 1;
        var itemRows = Math.Max(1, (int)Math.Ceiling(source.Rows.Count / (double)itemsPerRow));
        var itemColumns = Math.Min(Math.Max(source.Rows.Count, 1), itemsPerRow);
        var height = itemRows * itemHeight + (itemRows - 1) * rowGap;
        var width = itemColumns * itemWidth + (itemColumns - 1) * columnGap;
        return new TemplateRegionBounds(start.Row, start.Column, start.Row + height - 1, start.Column + width - 1);
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var repeater = new ExcelRepeaterSpec(templateContext.GetSource(region.GetRequiredString("source")))
        {
            StartCell = StartOf(region, startCell),
            TitleField = region.GetOptionalString("titleField")
        };

        if (region.TryGetProperty("itemsPerRow", out var itemsPerRow))
        {
            repeater.ItemsPerRow = itemsPerRow.GetInt32();
        }

        if (region.TryGetProperty("itemSize", out var itemSize))
        {
            repeater.ItemWidth = itemSize.GetRequiredInt32("columns");
            repeater.ItemHeight = itemSize.GetRequiredInt32("rows");
        }

        if (region.TryGetProperty("gap", out var gap))
        {
            repeater.ColumnGap = gap.GetOptionalInt32("columns") ?? repeater.ColumnGap;
            repeater.RowGap = gap.GetOptionalInt32("rows") ?? repeater.RowGap;
        }

        if (region.TryGetProperty("styles", out var repeaterStyles))
        {
            repeater.CardStyle = templateContext.ResolveStyle(repeaterStyles.GetOptionalString("card"));
            repeater.TitleStyle = templateContext.ResolveStyle(repeaterStyles.GetOptionalString("title"));
            repeater.LabelStyle = templateContext.ResolveStyle(repeaterStyles.GetOptionalString("label"));
            repeater.ValueStyle = templateContext.ResolveStyle(repeaterStyles.GetOptionalString("value"));
        }

        if (!region.TryGetProperty("fields", out var fields))
        {
            sheet.Repeaters.Add(repeater);
            return;
        }

        // Formula fields refer to other fields of the same card by name, so remember which row of the card
        // each source field lands on.
        var sourceFieldRows = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var fieldIndex = 0;
        foreach (var fieldElement in fields.EnumerateArray())
        {
            var caption = fieldElement.GetOptionalString("caption") ?? fieldElement.GetOptionalString("source") ?? "Value";
            ExcelRepeaterFieldSpec field;
            if (fieldElement.TryGetProperty("formula", out var formula))
            {
                var formulaText = formula.GetString() ?? string.Empty;
                field = new ExcelRepeaterFieldSpec(caption, caption, context => CompileFormula(formulaText, context, sourceFieldRows));
            }
            else
            {
                var source = fieldElement.GetRequiredString("source");
                field = new ExcelRepeaterFieldSpec(source, caption);
                sourceFieldRows[source] = fieldIndex;
            }

            if (fieldElement.TryGetProperty("labelStyle", out var labelStyle))
            {
                field.LabelStyle = templateContext.ResolveStyle(labelStyle.GetString());
            }

            if (fieldElement.TryGetProperty("valueStyle", out var valueStyle))
            {
                field.ValueStyle = templateContext.ResolveStyle(valueStyle.GetString());
            }

            repeater.Fields.Add(field);
            fieldIndex++;
        }

        sheet.Repeaters.Add(repeater);
    }

    /// <summary>
    /// Turns VALUE(fieldName) references into the address of that field's value cell within the card.
    /// </summary>
    private static string CompileFormula(
        string formula,
        ExcelRepeaterFormulaContext context,
        IReadOnlyDictionary<string, int> sourceFieldRows)
    {
        return ValueExpressionRegex().Replace(formula, match =>
        {
            var fieldName = match.Groups["field"].Value;
            if (!sourceFieldRows.TryGetValue(fieldName, out var fieldIndex))
            {
                throw new InvalidOperationException($"Formula field reference '{fieldName}' was not found in the repeater fields.");
            }

            var valueCell = new ExcelCellAddress(context.WorksheetRow + 1 + fieldIndex, context.WorksheetColumn + 1);
            return valueCell.ToReference();
        });
    }

    [GeneratedRegex("VALUE\\((?<field>[A-Za-z_][A-Za-z0-9_]*)\\)", RegexOptions.IgnoreCase)]
    private static partial Regex ValueExpressionRegex();
}
