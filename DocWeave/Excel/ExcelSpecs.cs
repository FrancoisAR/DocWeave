namespace DocWeave.Excel;

internal sealed record ExcelWorkbookSpec(IReadOnlyList<ExcelSheetSpec> Sheets)
{
    public ExcelWorkbookProtectionSpec? Protection { get; set; }
}

internal sealed class ExcelSheetSpec
{
    public ExcelSheetSpec(string name)
    {
        Name = name;
    }

    public string Name { get; }
    public List<ExcelCellSpec> Cells { get; } = [];
    public List<ExcelTableSpec> Tables { get; } = [];
    public List<ExcelRepeaterSpec> Repeaters { get; } = [];
    public List<ExcelChartSpec> Charts { get; } = [];
    public List<ExcelPivotTableSpec> PivotTables { get; } = [];
    public Dictionary<int, ExcelStyleSpec> ColumnStyles { get; } = [];
    public List<string> MergeRanges { get; } = [];
    public Dictionary<int, ExcelColumnSettingsSpec> ColumnSettings { get; } = [];
    public Dictionary<int, ExcelRowSettingsSpec> RowSettings { get; } = [];
    public List<ExcelRangeStyleSpec> RangeStyles { get; } = [];
    public List<ExcelCommentSpec> Comments { get; } = [];
    public List<ExcelDataValidationSpec> DataValidations { get; } = [];
    public ExcelFreezePaneSpec? FreezePane { get; set; }
    public string? AutoFilterRange { get; set; }
    public ExcelSheetProtectionSpec? Protection { get; set; }
}

internal sealed record ExcelCellSpec(ExcelCellAddress Address, object? Value, ExcelStyleSpec Style);

internal sealed record ExcelRangeStyleSpec(ExcelCellAddress Start, ExcelCellAddress End, ExcelStyleSpec Style);

internal sealed record ExcelCommentSpec(ExcelCellAddress Address, string Text, string Author);

internal sealed record ExcelFreezePaneSpec(int FrozenRows, int FrozenColumns);

internal sealed record ExcelWorkbookProtectionSpec(bool LockStructure, string? Password);

internal sealed record ExcelSheetProtectionSpec(string? Password, bool AllowSelectLockedCells, bool AllowSort, bool AllowAutoFilter);

internal sealed record ExcelDataValidationSpec(
    string Range,
    ExcelDataValidationKind Kind,
    ExcelDataValidationOperator Operator,
    string? Formula1,
    string? Formula2,
    bool AllowBlank,
    string? InputTitle,
    string? InputMessage,
    string? ErrorTitle,
    string? ErrorMessage);

/// <summary>
/// What kind of value a data validation rule accepts.
/// </summary>
public enum ExcelDataValidationKind
{
    /// <summary>
    /// A value picked from a list, given as fixed values or as a range of cells.
    /// </summary>
    List,
    /// <summary>
    /// A whole number within limits.
    /// </summary>
    WholeNumber,
    /// <summary>
    /// A decimal number within limits.
    /// </summary>
    Decimal,
    /// <summary>
    /// A date within limits.
    /// </summary>
    Date,
    /// <summary>
    /// Text whose length is within limits.
    /// </summary>
    TextLength,
    /// <summary>
    /// Any value for which a custom formula is true.
    /// </summary>
    Custom
}

/// <summary>
/// How a data validation rule compares the entered value with its limit or limits.
/// </summary>
public enum ExcelDataValidationOperator
{
    /// <summary>
    /// The value is at or between the two limits.
    /// </summary>
    Between,
    /// <summary>
    /// The value is outside the two limits.
    /// </summary>
    NotBetween,
    /// <summary>
    /// The value equals the limit.
    /// </summary>
    Equal,
    /// <summary>
    /// The value differs from the limit.
    /// </summary>
    NotEqual,
    /// <summary>
    /// The value is greater than the limit.
    /// </summary>
    GreaterThan,
    /// <summary>
    /// The value is less than the limit.
    /// </summary>
    LessThan,
    /// <summary>
    /// The value is greater than or equal to the limit.
    /// </summary>
    GreaterThanOrEqual,
    /// <summary>
    /// The value is less than or equal to the limit.
    /// </summary>
    LessThanOrEqual
}

/// <summary>
/// What a repeater field's formula can see when it is built for one card.
/// </summary>
/// <param name="DataIndex">The zero-based position of the item in the source data.</param>
/// <param name="WorksheetRow">The worksheet row of the card's top-left cell.</param>
/// <param name="WorksheetColumn">The worksheet column number of the card's top-left cell.</param>
/// <param name="Values">The item's values by column name. Names that start with two underscores are used by the library itself.</param>
public sealed record ExcelRepeaterFormulaContext(
    int DataIndex,
    int WorksheetRow,
    int WorksheetColumn,
    IReadOnlyDictionary<string, object?> Values);

internal sealed class ExcelRepeaterFieldSpec
{
    public ExcelRepeaterFieldSpec(
        string columnName,
        string caption,
        Func<ExcelRepeaterFormulaContext, string>? formulaFactory = null)
    {
        ColumnName = columnName;
        Caption = caption;
        FormulaFactory = formulaFactory;
    }

    public string ColumnName { get; }
    public string Caption { get; }
    public Func<ExcelRepeaterFormulaContext, string>? FormulaFactory { get; }
    public ExcelStyleSpec? LabelStyle { get; set; }
    public ExcelStyleSpec? ValueStyle { get; set; }
}

internal sealed class ExcelColumnSettingsSpec
{
    public double? Width { get; set; }
    public bool BestFit { get; set; }
    public bool Hidden { get; set; }
}

internal sealed class ExcelRowSettingsSpec
{
    public double? Height { get; set; }
    public bool AutoFit { get; set; }
    public bool Hidden { get; set; }
}

internal sealed record ExcelFormulaColumnSpec(string ColumnName, Func<ExcelFormulaContext, string> FormulaFactory);

internal sealed record ExcelConditionalStyleSpec(string ColumnName, Func<object?, bool> Predicate, ExcelStyleSpec Style);

internal sealed record ExcelRowStyleSpec(Func<IReadOnlyDictionary<string, object?>, bool> Predicate, ExcelStyleSpec Style);

internal sealed record ExcelCellConditionalStyleSpec(string ColumnName, Func<IReadOnlyDictionary<string, object?>, bool> Predicate, ExcelStyleSpec Style);

internal sealed record ExcelTotalsRowSpec(string Label, ExcelStyleSpec Style);

internal sealed record ExcelChartDataSource(string SheetName, string CategoryRange, string ValueRange);

internal sealed record ExcelPivotTableSource(string SheetName, string Range);

internal sealed record ExcelPivotValueFieldSpec(string FieldName, string Caption, ExcelPivotAggregateFunction Function);

internal sealed class ExcelPivotTableSpec
{
    public ExcelPivotTableSpec(string name)
    {
        Name = name;
    }

    public string Name { get; }
    public ExcelPivotTableSource? Source { get; set; }
    public ExcelCellAddress StartCell { get; set; } = new(1, 1);
    public List<string> RowFields { get; } = [];
    public List<string> ColumnFields { get; } = [];
    public List<string> FilterFields { get; } = [];
    public List<ExcelPivotValueFieldSpec> ValueFields { get; } = [];
    public string StyleName { get; set; } = "PivotStyleMedium9";
    public bool RowGrandTotals { get; set; } = true;
    public bool ColumnGrandTotals { get; set; } = true;
    public bool RefreshOnOpen { get; set; } = true;
}

internal sealed class ExcelChartSpec
{
    public ExcelChartSpec(string title)
    {
        Title = title;
    }

    public string Title { get; }
    public ExcelChartType Type { get; set; } = ExcelChartType.Column;
    public ExcelChartDataSource? DataSource { get; set; }
    public ExcelCellAddress TopLeft { get; set; } = new(1, 1);
    public ExcelCellAddress BottomRight { get; set; } = new(20, 8);
    public uint Style { get; set; } = 10;
    public string? SeriesColor { get; set; }
    public bool ShowLegend { get; set; } = true;
    public ExcelChartLegendPosition LegendPosition { get; set; } = ExcelChartLegendPosition.Right;
    public bool ShowDataLabels { get; set; }
    public string? CategoryAxisTitle { get; set; }
    public string? ValueAxisTitle { get; set; }
}

internal sealed class ExcelTableSpec
{
    public ExcelTableSpec(TabularData data)
    {
        Data = data;
    }

    public TabularData Data { get; }
    public ExcelCellAddress StartCell { get; set; } = new(1, 1);
    public bool IncludeHeaders { get; set; } = true;
    public Dictionary<string, string> HeaderAliases { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, int> ColumnLocations { get; } = new(StringComparer.OrdinalIgnoreCase);
    public ExcelStyleSpec HeaderStyle { get; set; } = ExcelStyleSpec.Empty;
    public ExcelStyleSpec DataStyle { get; set; } = ExcelStyleSpec.Empty;
    public ExcelStyleSpec? AlternateRowStyle { get; set; }
    public Dictionary<string, ExcelStyleSpec> ColumnStyles { get; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<int, ExcelStyleSpec> RowStyles { get; } = [];
    public List<ExcelRowStyleSpec> ConditionalRowStyles { get; } = [];
    public List<ExcelCellConditionalStyleSpec> ConditionalCellStyles { get; } = [];
    public List<ExcelFormulaColumnSpec> FormulaColumns { get; } = [];
    public List<ExcelConditionalStyleSpec> ConditionalStyles { get; } = [];
    public ExcelTotalsRowSpec? TotalsRow { get; set; }
    public ExcelNativeTableSpec? NativeTable { get; set; }
}

internal sealed record ExcelNativeTableSpec(string Name, string StyleName, bool ShowFilterButtons, bool ShowRowStripes);

internal sealed class ExcelRepeaterSpec
{
    public ExcelRepeaterSpec(TabularData data)
    {
        Data = data;
    }

    public TabularData Data { get; }
    public ExcelCellAddress StartCell { get; set; } = new(1, 1);
    public int ItemsPerRow { get; set; } = 3;
    public int ItemWidth { get; set; } = 4;
    public int ItemHeight { get; set; } = 6;
    public int ColumnGap { get; set; } = 1;
    public int RowGap { get; set; } = 1;
    public string? TitleField { get; set; }
    public List<ExcelRepeaterFieldSpec> Fields { get; } = [];
    public ExcelStyleSpec CardStyle { get; set; } = ExcelStyleSpec.Empty;
    public ExcelStyleSpec TitleStyle { get; set; } = ExcelStyleSpec.Empty;
    public ExcelStyleSpec LabelStyle { get; set; } = ExcelStyleSpec.Empty;
    public ExcelStyleSpec ValueStyle { get; set; } = ExcelStyleSpec.Empty;
}
