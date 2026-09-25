using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "cell", "cell": "A1", "value": "...", "style": "..." }
/// </summary>
internal sealed class CellTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "cell";

    public override bool OccupiesCells => true;

    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        return TemplateRegionBounds.SingleCell(ExcelCellAddress.Parse(region.GetRequiredString("cell")));
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.Cells.Add(new ExcelCellSpec(
            ExcelCellAddress.Parse(region.GetRequiredString("cell")),
            templateContext.BindValue(ExcelTemplateValues.ReadValue(region, "value")),
            templateContext.ResolveStyle(region.GetOptionalString("style"))));
    }
}

/// <summary>
/// { "type": "formulaCell", "cell": "B1", "formula": "SUM(B4:B7)", "style": "..." }
/// </summary>
internal sealed class FormulaCellTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "formulaCell";

    public override bool OccupiesCells => true;

    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        return TemplateRegionBounds.SingleCell(ExcelCellAddress.Parse(region.GetRequiredString("cell")));
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.Cells.Add(new ExcelCellSpec(
            ExcelCellAddress.Parse(region.GetRequiredString("cell")),
            new ExcelFormula(templateContext.BindText(region.GetRequiredString("formula"))),
            templateContext.ResolveStyle(region.GetOptionalString("style"))));
    }
}

/// <summary>
/// { "type": "rangeStyle", "range": "A1:D5", "style": "..." }
/// </summary>
internal sealed class RangeStyleTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "rangeStyle";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var (start, end) = ExcelCellAddress.ParseRange(region.GetRequiredString("range"));
        sheet.RangeStyles.Add(new ExcelRangeStyleSpec(start, end, templateContext.ResolveStyle(region.GetRequiredString("style"))));
    }
}

/// <summary>
/// { "type": "mergeCells", "range": "A1:C1" }
/// </summary>
internal sealed class MergeCellsTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "mergeCells";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.MergeRanges.Add(region.GetRequiredString("range"));
    }
}

/// <summary>
/// { "type": "comment", "cell": "A1", "text": "...", "author": "..." }
/// </summary>
internal sealed class CommentTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "comment";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.Comments.Add(new ExcelCommentSpec(
            ExcelCellAddress.Parse(region.GetRequiredString("cell")),
            templateContext.BindText(region.GetRequiredString("text")),
            region.GetOptionalString("author") ?? "User"));
    }
}

/// <summary>
/// { "type": "column", "column": "B", "width": 20, "autoFit": false, "hidden": false, "style": "..." }
/// </summary>
internal sealed class ColumnTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "column";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var columnNumber = ExcelCellAddress.ColumnLetterToNumber(region.GetRequiredString("column"));
        if (!sheet.ColumnSettings.TryGetValue(columnNumber, out var settings))
        {
            settings = new ExcelColumnSettingsSpec();
            sheet.ColumnSettings[columnNumber] = settings;
        }

        if (region.TryGetProperty("width", out var width))
        {
            settings.Width = width.GetDouble();
        }

        settings.BestFit = region.GetOptionalBool("autoFit") ?? settings.BestFit;
        settings.Hidden = region.GetOptionalBool("hidden") ?? settings.Hidden;

        var styleName = region.GetOptionalString("style");
        if (!string.IsNullOrWhiteSpace(styleName))
        {
            sheet.ColumnStyles[columnNumber] = templateContext.ResolveStyle(styleName);
        }
    }
}

/// <summary>
/// { "type": "row", "row": 1, "height": 30, "autoFit": false, "hidden": false }
/// </summary>
internal sealed class RowTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "row";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var rowNumber = region.GetRequiredInt32("row");
        if (!sheet.RowSettings.TryGetValue(rowNumber, out var settings))
        {
            settings = new ExcelRowSettingsSpec();
            sheet.RowSettings[rowNumber] = settings;
        }

        if (region.TryGetProperty("height", out var height))
        {
            settings.Height = height.GetDouble();
        }

        settings.AutoFit = region.GetOptionalBool("autoFit") ?? settings.AutoFit;
        settings.Hidden = region.GetOptionalBool("hidden") ?? settings.Hidden;
    }
}

/// <summary>
/// { "type": "freezePane", "rows": 1, "columns": 0 }
/// </summary>
internal sealed class FreezePaneTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "freezePane";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.FreezePane = new ExcelFreezePaneSpec(
            region.GetOptionalInt32("rows") ?? 1,
            region.GetOptionalInt32("columns") ?? 0);
    }
}

/// <summary>
/// { "type": "autoFilter", "range": "A3:F100" }
/// </summary>
internal sealed class AutoFilterTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "autoFilter";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.AutoFilterRange = region.GetRequiredString("range");
    }
}

/// <summary>
/// { "type": "sheetProtection", "password": "...", "allowSort": false, "allowAutoFilter": false }
/// </summary>
internal sealed class SheetProtectionTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "sheetProtection";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        sheet.Protection = new ExcelSheetProtectionSpec(
            region.GetOptionalString("password"),
            region.GetOptionalBool("allowSelectLockedCells") ?? true,
            region.GetOptionalBool("allowSort") ?? false,
            region.GetOptionalBool("allowAutoFilter") ?? false);
    }
}
