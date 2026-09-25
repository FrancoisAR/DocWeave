using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// The cells a region takes up, as a rectangle. Used to place regions below one another and to detect overlaps.
/// </summary>
internal sealed record TemplateRegionBounds(int StartRow, int StartColumn, int EndRow, int EndColumn)
{
    /// <summary>
    /// What regions that take no cells of their own report. It is never checked for overlaps.
    /// </summary>
    public static TemplateRegionBounds Default { get; } = new(1, 1, 1, 1);

    public static TemplateRegionBounds SingleCell(ExcelCellAddress address)
    {
        return new TemplateRegionBounds(address.Row, address.Column, address.Row, address.Column);
    }

    /// <summary>
    /// The rectangle between two corners, whichever way round they are given.
    /// </summary>
    public static TemplateRegionBounds Between(string startCell, string endCell)
    {
        var start = ExcelCellAddress.Parse(startCell);
        var end = ExcelCellAddress.Parse(endCell);
        return new TemplateRegionBounds(
            Math.Min(start.Row, end.Row),
            Math.Min(start.Column, end.Column),
            Math.Max(start.Row, end.Row),
            Math.Max(start.Column, end.Column));
    }

    public bool Overlaps(TemplateRegionBounds other)
    {
        return StartRow <= other.EndRow
            && EndRow >= other.StartRow
            && StartColumn <= other.EndColumn
            && EndColumn >= other.StartColumn;
    }
}

/// <summary>
/// Where a region ended up. StartCell is set only when the template asked for relative placement.
/// </summary>
internal sealed record TemplateRegionPlacement(TemplateRegionBounds Bounds, ExcelCellAddress? StartCell);

/// <summary>
/// Works out where regions go on a sheet and checks they do not collide.
/// </summary>
internal static class ExcelTemplateLayout
{
    /// <summary>
    /// Resolves a region's placement. With "placement": { "mode": "belowPrevious" } the region starts
    /// below the previous cell-occupying region, optionally in another column and with a gap of rows.
    /// </summary>
    public static TemplateRegionPlacement ResolvePlacement(
        JsonElement region,
        ExcelTemplateRegionCompiler? compiler,
        ExcelTemplateContext templateContext,
        TemplateRegionBounds? previous)
    {
        ExcelCellAddress? startOverride = null;
        if (region.TryGetProperty("placement", out var placement)
            && string.Equals(placement.GetOptionalString("mode"), "belowPrevious", StringComparison.OrdinalIgnoreCase))
        {
            if (previous is null)
            {
                throw new InvalidOperationException("Template placement mode 'belowPrevious' requires an earlier placed region.");
            }

            var column = placement.GetOptionalString("column");
            var columnNumber = string.IsNullOrWhiteSpace(column)
                ? previous.StartColumn
                : ExcelCellAddress.ColumnLetterToNumber(column);
            // By default one blank row is left between the previous region and this one.
            var gapRows = placement.GetOptionalInt32("gapRows") ?? 1;
            startOverride = new ExcelCellAddress(previous.EndRow + gapRows + 1, columnNumber);
        }

        // A region type with no compiler, or one that takes no cells, gets the default bounds, which are never checked for overlap.
        var bounds = compiler?.EstimateBounds(region, templateContext, startOverride) ?? TemplateRegionBounds.Default;
        return new TemplateRegionPlacement(bounds, startOverride);
    }

    public static void ValidateNoOverlap(List<TemplateRegionBounds> occupied, TemplateRegionBounds next)
    {
        if (occupied.Any(next.Overlaps))
        {
            throw new InvalidOperationException($"Template region starting at {ExcelCellAddress.ColumnNumberToLetter(next.StartColumn)}{next.StartRow} overlaps an earlier region.");
        }

        occupied.Add(next);
    }
}
