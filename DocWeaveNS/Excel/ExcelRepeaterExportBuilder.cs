namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for repeating object layouts, similar to a report data repeater.
/// </summary>
public sealed class ExcelRepeaterExportBuilder
{
    private readonly ExcelRepeaterSpec repeater;

    internal ExcelRepeaterExportBuilder(ExcelRepeaterSpec repeater)
    {
        this.repeater = repeater;
    }

    /// <summary>
    /// Sets the top-left cell for the first repeated item.
    /// </summary>
    public ExcelRepeaterExportBuilder StartAt(string cellReference)
    {
        repeater.StartCell = ExcelCellAddress.Parse(cellReference);
        return this;
    }

    /// <summary>
    /// Sets how many repeated items should be placed across a worksheet row.
    /// </summary>
    public ExcelRepeaterExportBuilder ItemsPerRow(int count)
    {
        if (count < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(count), "Items per row must be one or greater.");
        }

        repeater.ItemsPerRow = count;
        return this;
    }

    /// <summary>
    /// Sets the repeated item block size, measured in worksheet columns and rows.
    /// </summary>
    public ExcelRepeaterExportBuilder ItemSize(int columns, int rows)
    {
        if (columns < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), "Item width must be one or greater.");
        }

        if (rows < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), "Item height must be one or greater.");
        }

        repeater.ItemWidth = columns;
        repeater.ItemHeight = rows;
        return this;
    }

    /// <summary>
    /// Sets spacing between repeated item blocks, measured in worksheet columns and rows.
    /// </summary>
    public ExcelRepeaterExportBuilder Gap(int columns = 1, int rows = 1)
    {
        if (columns < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), "Column gap cannot be negative.");
        }

        if (rows < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), "Row gap cannot be negative.");
        }

        repeater.ColumnGap = columns;
        repeater.RowGap = rows;
        return this;
    }

    /// <summary>
    /// Uses the selected source field as the item title row.
    /// </summary>
    public ExcelRepeaterExportBuilder TitleField(string columnName)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(columnName, nameof(columnName));
        repeater.TitleField = columnName;
        return this;
    }

    /// <summary>
    /// Adds a field to each repeated item. Fields are written in the order configured.
    /// </summary>
    public ExcelRepeaterExportBuilder Field(string columnName, string? caption = null)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(columnName, nameof(columnName));
        repeater.Fields.Add(new ExcelRepeaterFieldSpec(columnName, caption ?? columnName));
        return this;
    }

    /// <summary>
    /// Adds a calculated field to each repeated item. The formula can be supplied with or without the leading equals sign.
    /// </summary>
    public ExcelRepeaterExportBuilder FormulaField(string caption, Func<ExcelRepeaterFormulaContext, string> formulaFactory)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(caption, nameof(caption));
        NetStandardGuard.ThrowIfNull(formulaFactory, nameof(formulaFactory));
        repeater.Fields.Add(new ExcelRepeaterFieldSpec(caption, caption, formulaFactory));
        return this;
    }

    /// <summary>
    /// Adds multiple source fields to each repeated item.
    /// </summary>
    public ExcelRepeaterExportBuilder Fields(params string[] columnNames)
    {
        foreach (var columnName in columnNames)
        {
            Field(columnName);
        }

        return this;
    }

    /// <summary>
    /// Applies a style to one field caption, overriding the repeater-level label style.
    /// </summary>
    public ExcelRepeaterExportBuilder FieldLabelStyle(string columnNameOrCaption, Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(columnNameOrCaption, nameof(columnNameOrCaption));
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        FindField(columnNameOrCaption).LabelStyle = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to one field value, overriding the repeater-level value style.
    /// </summary>
    public ExcelRepeaterExportBuilder FieldValueStyle(string columnNameOrCaption, Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(columnNameOrCaption, nameof(columnNameOrCaption));
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        FindField(columnNameOrCaption).ValueStyle = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to the whole rectangular item block.
    /// </summary>
    public ExcelRepeaterExportBuilder CardStyle(Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        repeater.CardStyle = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to each repeated item's title cell.
    /// </summary>
    public ExcelRepeaterExportBuilder TitleStyle(Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        repeater.TitleStyle = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to field captions inside each item block.
    /// </summary>
    public ExcelRepeaterExportBuilder LabelStyle(Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        repeater.LabelStyle = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    /// <summary>
    /// Applies a style to field values inside each item block.
    /// </summary>
    public ExcelRepeaterExportBuilder ValueStyle(Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        repeater.ValueStyle = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    private ExcelRepeaterFieldSpec FindField(string columnNameOrCaption)
    {
        var field = repeater.Fields.LastOrDefault(field =>
            string.Equals(field.ColumnName, columnNameOrCaption, StringComparison.OrdinalIgnoreCase)
            || string.Equals(field.Caption, columnNameOrCaption, StringComparison.OrdinalIgnoreCase));

        return field ?? throw new InvalidOperationException($"Repeater field '{columnNameOrCaption}' has not been configured.");
    }
}
