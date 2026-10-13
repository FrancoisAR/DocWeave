namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for an Excel pivot table placed on a worksheet.
/// </summary>
public sealed class ExcelPivotTableBuilder
{
    private readonly ExcelPivotTableSpec pivotTable;

    internal ExcelPivotTableBuilder(ExcelPivotTableSpec pivotTable)
    {
        this.pivotTable = pivotTable;
    }

    /// <summary>
    /// Sets the worksheet and range that contains the source rows, including headers.
    /// </summary>
    public ExcelPivotTableBuilder Source(string sheetName, string range)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(sheetName, nameof(sheetName));
        NetStandardGuard.ThrowIfNullOrWhiteSpace(range, nameof(range));
        pivotTable.Source = new ExcelPivotTableSource(sheetName, range);
        return this;
    }

    /// <summary>
    /// Sets the top-left cell where the pivot table should be placed on the target worksheet.
    /// </summary>
    public ExcelPivotTableBuilder StartAt(string cellReference)
    {
        pivotTable.StartCell = ExcelCellAddress.Parse(cellReference);
        return this;
    }

    /// <summary>
    /// Adds row fields in display order.
    /// </summary>
    public ExcelPivotTableBuilder Rows(params string[] fieldNames)
    {
        AddFields(pivotTable.RowFields, fieldNames);
        return this;
    }

    /// <summary>
    /// Adds column fields in display order.
    /// </summary>
    public ExcelPivotTableBuilder Columns(params string[] fieldNames)
    {
        AddFields(pivotTable.ColumnFields, fieldNames);
        return this;
    }

    /// <summary>
    /// Adds report filter fields.
    /// </summary>
    public ExcelPivotTableBuilder Filters(params string[] fieldNames)
    {
        AddFields(pivotTable.FilterFields, fieldNames);
        return this;
    }

    /// <summary>
    /// Adds value fields such as sums, counts, and averages.
    /// </summary>
    public ExcelPivotTableBuilder Values(Action<ExcelPivotValueBuilder> configure)
    {
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        configure(new ExcelPivotValueBuilder(pivotTable));
        return this;
    }

    /// <summary>
    /// Sets the built-in Excel pivot table style name.
    /// </summary>
    public ExcelPivotTableBuilder Style(string styleName)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(styleName, nameof(styleName));
        pivotTable.StyleName = styleName;
        return this;
    }

    /// <summary>
    /// Controls row and column grand totals.
    /// </summary>
    public ExcelPivotTableBuilder ShowGrandTotals(bool rows = true, bool columns = true)
    {
        pivotTable.RowGrandTotals = rows;
        pivotTable.ColumnGrandTotals = columns;
        return this;
    }

    /// <summary>
    /// Controls whether Excel should refresh the pivot cache when the workbook is opened.
    /// </summary>
    public ExcelPivotTableBuilder RefreshOnOpen(bool refresh = true)
    {
        pivotTable.RefreshOnOpen = refresh;
        return this;
    }

    private static void AddFields(List<string> target, IEnumerable<string> fieldNames)
    {
        foreach (var fieldName in fieldNames)
        {
            NetStandardGuard.ThrowIfNullOrWhiteSpace(fieldName, nameof(fieldName));
            target.Add(fieldName);
        }
    }
}

/// <summary>
/// Fluent builder for pivot table value fields.
/// </summary>
public sealed class ExcelPivotValueBuilder
{
    private readonly ExcelPivotTableSpec pivotTable;

    internal ExcelPivotValueBuilder(ExcelPivotTableSpec pivotTable)
    {
        this.pivotTable = pivotTable;
    }

    /// <summary>
    /// Adds a value field that adds up the values.
    /// </summary>
    /// <param name="fieldName">The source column to summarise. It must be a header in the pivot's source range.</param>
    /// <param name="caption">The heading shown in the pivot, for example "Total Sales". Defaults to "Sum of" plus the field name. Captions must be unique and must not match a source column name.</param>
    public ExcelPivotValueBuilder Sum(string fieldName, string? caption = null)
    {
        return Add(fieldName, caption, ExcelPivotAggregateFunction.Sum);
    }

    /// <summary>
    /// Adds a value field that counts the entries that have a value.
    /// </summary>
    /// <param name="fieldName">The source column to count. It must be a header in the pivot's source range.</param>
    /// <param name="caption">The heading shown in the pivot. Defaults to "Count of" plus the field name. Captions must be unique and must not match a source column name.</param>
    public ExcelPivotValueBuilder Count(string fieldName, string? caption = null)
    {
        return Add(fieldName, caption, ExcelPivotAggregateFunction.Count);
    }

    /// <summary>
    /// Adds a value field that averages the values.
    /// </summary>
    /// <param name="fieldName">The source column to average. It must be a header in the pivot's source range.</param>
    /// <param name="caption">The heading shown in the pivot. Defaults to "Average of" plus the field name. Captions must be unique and must not match a source column name.</param>
    public ExcelPivotValueBuilder Average(string fieldName, string? caption = null)
    {
        return Add(fieldName, caption, ExcelPivotAggregateFunction.Average);
    }

    /// <summary>
    /// Adds a value field that shows the smallest value.
    /// </summary>
    /// <param name="fieldName">The source column to look at. It must be a header in the pivot's source range.</param>
    /// <param name="caption">The heading shown in the pivot. Defaults to "Minimum of" plus the field name. Captions must be unique and must not match a source column name.</param>
    public ExcelPivotValueBuilder Min(string fieldName, string? caption = null)
    {
        return Add(fieldName, caption, ExcelPivotAggregateFunction.Minimum);
    }

    /// <summary>
    /// Adds a value field that shows the largest value.
    /// </summary>
    /// <param name="fieldName">The source column to look at. It must be a header in the pivot's source range.</param>
    /// <param name="caption">The heading shown in the pivot. Defaults to "Maximum of" plus the field name. Captions must be unique and must not match a source column name.</param>
    public ExcelPivotValueBuilder Max(string fieldName, string? caption = null)
    {
        return Add(fieldName, caption, ExcelPivotAggregateFunction.Maximum);
    }

    private ExcelPivotValueBuilder Add(string fieldName, string? caption, ExcelPivotAggregateFunction function)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(fieldName, nameof(fieldName));
        pivotTable.ValueFields.Add(new ExcelPivotValueFieldSpec(fieldName, caption ?? $"{function} of {fieldName}", function));
        return this;
    }
}

/// <summary>
/// Supported aggregate functions for the initial pivot table export layer.
/// </summary>
public enum ExcelPivotAggregateFunction
{
    /// <summary>
    /// Adds the values.
    /// </summary>
    Sum,
    /// <summary>
    /// Counts the entries that have a value.
    /// </summary>
    Count,
    /// <summary>
    /// The mean of the values.
    /// </summary>
    Average,
    /// <summary>
    /// The smallest value.
    /// </summary>
    Minimum,
    /// <summary>
    /// The largest value.
    /// </summary>
    Maximum
}
