namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for a table-like data region in an Excel worksheet.
/// </summary>
public sealed class ExcelTableExportBuilder
{
    private readonly ExcelTableSpec table;

    internal ExcelTableExportBuilder(ExcelTableSpec table)
    {
        this.table = table;
    }

    /// <summary>
    /// Configures generated total rows for this table.
    /// </summary>
    public ExcelTableTotalsBuilder AddTotals => new(table, this);

    /// <summary>
    /// Sets the top-left cell for this data region.
    /// </summary>
    public ExcelTableExportBuilder StartAt(string cellReference)
    {
        table.StartCell = ExcelCellAddress.Parse(cellReference);
        return this;
    }

    /// <summary>
    /// Controls whether a header row is written before the data rows.
    /// </summary>
    public ExcelTableExportBuilder IncludeHeaders(bool value = true)
    {
        table.IncludeHeaders = value;
        return this;
    }

    /// <summary>
    /// Replaces exported column captions without changing the source column names.
    /// </summary>
    public ExcelTableExportBuilder HeaderAliases(IReadOnlyDictionary<string, string> aliases)
    {
        ArgumentNullException.ThrowIfNull(aliases);
        foreach (var alias in aliases)
        {
            table.HeaderAliases[alias.Key] = alias.Value;
        }

        return this;
    }

    /// <summary>
    /// Replaces one exported column caption without changing the source column name.
    /// </summary>
    public ExcelTableExportBuilder HeaderAlias(string columnName, string caption)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        ArgumentException.ThrowIfNullOrWhiteSpace(caption);

        table.HeaderAliases[columnName] = caption;
        return this;
    }

    /// <summary>
    /// Places a source column in a specific worksheet column, such as B, C, or AA.
    /// </summary>
    public ExcelTableExportBuilder ColumnLocation(string columnName, string columnLetter)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        table.ColumnLocations[columnName] = ExcelCellAddress.ColumnLetterToNumber(columnLetter);
        return this;
    }

    /// <summary>
    /// Places multiple source columns in specific worksheet columns.
    /// </summary>
    public ExcelTableExportBuilder ColumnLocations(IReadOnlyDictionary<string, string> locations)
    {
        ArgumentNullException.ThrowIfNull(locations);
        foreach (var location in locations)
        {
            ColumnLocation(location.Key, location.Value);
        }

        return this;
    }

    /// <summary>
    /// Applies a style to the table header row.
    /// </summary>
    public ExcelTableExportBuilder HeaderStyle(Action<ExcelStyleBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        table.HeaderStyle = BuildStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to all data cells in the table.
    /// </summary>
    public ExcelTableExportBuilder DataStyle(Action<ExcelStyleBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        table.DataStyle = BuildStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to every second data row.
    /// </summary>
    public ExcelTableExportBuilder AlternateRowStyle(Action<ExcelStyleBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        table.AlternateRowStyle = BuildStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to one source column's data cells.
    /// </summary>
    public ExcelTableExportBuilder ColumnStyle(string columnName, Action<ExcelStyleBuilder> configure)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        ArgumentNullException.ThrowIfNull(configure);
        table.ColumnStyles[columnName] = BuildStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to a specific one-based data row within the exported table.
    /// </summary>
    public ExcelTableExportBuilder RowStyle(int dataRowNumber, Action<ExcelStyleBuilder> configure)
    {
        if (dataRowNumber < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(dataRowNumber), "Data row number must be one or greater.");
        }

        ArgumentNullException.ThrowIfNull(configure);
        table.RowStyles[dataRowNumber - 1] = BuildStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to every data cell in rows that match the supplied row predicate.
    /// </summary>
    public ExcelTableExportBuilder RowStyleWhen(Func<IReadOnlyDictionary<string, object?>, bool> predicate, Action<ExcelStyleBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(predicate);
        ArgumentNullException.ThrowIfNull(configure);
        table.ConditionalRowStyles.Add(new ExcelRowStyleSpec(predicate, BuildStyle(configure)));
        return this;
    }

    /// <summary>
    /// Applies a style to one column's cell when the exported row matches the supplied predicate.
    /// </summary>
    public ExcelTableExportBuilder CellStyle(string columnName, Func<IReadOnlyDictionary<string, object?>, bool> predicate, Action<ExcelStyleBuilder> configure)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        ArgumentNullException.ThrowIfNull(predicate);
        ArgumentNullException.ThrowIfNull(configure);
        table.ConditionalCellStyles.Add(new ExcelCellConditionalStyleSpec(columnName, predicate, BuildStyle(configure)));
        return this;
    }

    /// <summary>
    /// Adds a calculated table column whose cells contain formulas.
    /// </summary>
    public ExcelTableExportBuilder FormulaColumn(string columnName, Func<ExcelFormulaContext, string> formulaFactory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        ArgumentNullException.ThrowIfNull(formulaFactory);
        table.FormulaColumns.Add(new ExcelFormulaColumnSpec(columnName, formulaFactory));
        return this;
    }

    /// <summary>
    /// Applies a style when a source column value satisfies the supplied predicate.
    /// </summary>
    public ExcelTableExportBuilder ConditionalStyle(string columnName, Func<object?, bool> predicate, Action<ExcelStyleBuilder> configure)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);
        ArgumentNullException.ThrowIfNull(predicate);
        ArgumentNullException.ThrowIfNull(configure);
        table.ConditionalStyles.Add(new ExcelConditionalStyleSpec(columnName, predicate, BuildStyle(configure)));
        return this;
    }

    /// <summary>
    /// Emits this region as a native Excel table with a TableDefinitionPart.
    /// </summary>
    public ExcelTableExportBuilder AsNativeTable(string name, string styleName = "TableStyleMedium2", bool showFilterButtons = true, bool showRowStripes = true)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        table.NativeTable = new ExcelNativeTableSpec(name, styleName, showFilterButtons, showRowStripes);
        table.IncludeHeaders = true;
        return this;
    }

    private static ExcelStyleSpec BuildStyle(Action<ExcelStyleBuilder> configure)
    {
        var builder = new ExcelStyleBuilder();
        configure(builder);
        return builder.Spec;
    }

    internal static ExcelStyleSpec CreateStyle(Action<ExcelStyleBuilder> configure)
    {
        return BuildStyle(configure);
    }
}

/// <summary>
/// Fluent builder for totals generated relative to a table's exported data rows.
/// </summary>
public sealed class ExcelTableTotalsBuilder
{
    private readonly ExcelTableSpec table;
    private readonly ExcelTableExportBuilder parent;

    internal ExcelTableTotalsBuilder(ExcelTableSpec table, ExcelTableExportBuilder parent)
    {
        this.table = table;
        this.parent = parent;
    }

    /// <summary>
    /// Adds a totals row directly beneath the final exported data row.
    /// </summary>
    public ExcelTableExportBuilder Bottom(string label = "Totals")
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(label);
        table.TotalsRow = new ExcelTotalsRowSpec(label, ExcelStyleSpec.Empty);
        return parent;
    }

    /// <summary>
    /// Adds a styled totals row directly beneath the final exported data row.
    /// </summary>
    public ExcelTableExportBuilder Bottom(string label, Action<ExcelStyleBuilder> configureStyle)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(label);
        ArgumentNullException.ThrowIfNull(configureStyle);
        table.TotalsRow = new ExcelTotalsRowSpec(label, ExcelTableExportBuilder.CreateStyle(configureStyle));
        return parent;
    }
}
