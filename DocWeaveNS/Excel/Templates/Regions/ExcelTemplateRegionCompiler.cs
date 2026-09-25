using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// Turns one kind of template region (a "table", a "chart", ...) into worksheet content.
/// A new region type is added by writing one of these and listing it in <see cref="ExcelTemplateRegions"/>.
/// </summary>
internal abstract class ExcelTemplateRegionCompiler
{
    /// <summary>
    /// The value of the region's "type" property that this compiler handles.
    /// </summary>
    public abstract string Type { get; }

    /// <summary>
    /// True for regions that take up cells of their own. Those are checked for overlaps and can be
    /// placed relative to one another. Settings regions (column width, freeze pane, ...) leave this false.
    /// </summary>
    public virtual bool OccupiesCells => false;

    /// <summary>
    /// The cells the region will take up. Only called for regions that occupy cells.
    /// </summary>
    /// <param name="region">The region's JSON object from the template.</param>
    /// <param name="templateContext">The bound data sources, parameters and styles.</param>
    /// <param name="startOverride">Set when the template placed the region below the previous one.</param>
    /// <returns>The rectangle of cells the region will fill.</returns>
    public virtual TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        return TemplateRegionBounds.Default;
    }

    /// <summary>
    /// Adds the region's content to the sheet.
    /// </summary>
    /// <param name="sheet">The worksheet being built.</param>
    /// <param name="region">The region's JSON object from the template.</param>
    /// <param name="templateContext">The bound data sources, parameters and styles.</param>
    /// <param name="startCell">Set when the template placed the region below the previous one; overrides "startCell".</param>
    public abstract void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell);

    /// <summary>
    /// The region's own "startCell", unless relative placement supplied one.
    /// </summary>
    protected static ExcelCellAddress StartOf(JsonElement region, ExcelCellAddress? startOverride)
    {
        return startOverride ?? ExcelCellAddress.Parse(region.GetRequiredString("startCell"));
    }
}

/// <summary>
/// The region types a template can use.
/// </summary>
internal static class ExcelTemplateRegions
{
    private static readonly Dictionary<string, ExcelTemplateRegionCompiler> Compilers = new ExcelTemplateRegionCompiler[]
    {
        new CellTemplateRegion(),
        new FormulaCellTemplateRegion(),
        new TableTemplateRegion(),
        new RepeaterTemplateRegion(),
        new SummaryTemplateRegion(),
        new ChartTemplateRegion(),
        new PivotTableTemplateRegion(),
        new RangeStyleTemplateRegion(),
        new MergeCellsTemplateRegion(),
        new CommentTemplateRegion(),
        new ColumnTemplateRegion(),
        new RowTemplateRegion(),
        new DataValidationTemplateRegion(),
        new FreezePaneTemplateRegion(),
        new AutoFilterTemplateRegion(),
        new SheetProtectionTemplateRegion()
    }.ToDictionary(compiler => compiler.Type, StringComparer.Ordinal);

    /// <summary>
    /// Finds the compiler for a region type, or null when the type is not supported. Type names are case-sensitive.
    /// </summary>
    public static ExcelTemplateRegionCompiler? Find(string type)
    {
        return Compilers.GetValueOrDefault(type);
    }
}
