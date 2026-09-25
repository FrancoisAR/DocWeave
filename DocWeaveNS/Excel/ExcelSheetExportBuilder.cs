using System.Data;

namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for content placed on a single Excel worksheet.
/// </summary>
public sealed class ExcelSheetExportBuilder
{
    private readonly ExcelSheetSpec sheet;

    internal ExcelSheetExportBuilder(ExcelSheetSpec sheet)
    {
        this.sheet = sheet;
    }

    /// <summary>
    /// Adds a table region from a DataTable.
    /// </summary>
    public ExcelTableExportBuilder AddTable(DataTable dataTable)
    {
        NetStandardGuard.ThrowIfNull(dataTable, nameof(dataTable));
        return AddTabularData(TabularData.FromDataTable(dataTable));
    }

    /// <summary>
    /// Adds a table region from an object collection, using public instance properties as columns.
    /// </summary>
    public ExcelTableExportBuilder AddTable<T>(IEnumerable<T> items)
    {
        NetStandardGuard.ThrowIfNull(items, nameof(items));
        return AddTabularData(TabularData.FromObjects(items));
    }

    /// <summary>
    /// Adds a table region from dictionary rows. Dictionary keys become column names.
    /// </summary>
    public ExcelTableExportBuilder AddRows(IEnumerable<IReadOnlyDictionary<string, object?>> rows)
    {
        NetStandardGuard.ThrowIfNull(rows, nameof(rows));
        return AddTabularData(TabularData.FromDictionaries(rows));
    }

    /// <summary>
    /// Adds a repeating layout region from an object collection, useful for card-like exports.
    /// </summary>
    public ExcelRepeaterExportBuilder AddRepeater<T>(IEnumerable<T> items)
    {
        NetStandardGuard.ThrowIfNull(items, nameof(items));
        return AddRepeaterData(TabularData.FromObjects(items));
    }

    /// <summary>
    /// Adds a repeating layout region from dictionary rows.
    /// </summary>
    public ExcelRepeaterExportBuilder AddRepeaterRows(IEnumerable<IReadOnlyDictionary<string, object?>> rows)
    {
        NetStandardGuard.ThrowIfNull(rows, nameof(rows));
        return AddRepeaterData(TabularData.FromDictionaries(rows));
    }

    /// <summary>
    /// Adds a single static cell. This is useful for report titles and metadata above imported regions.
    /// </summary>
    public ExcelSheetExportBuilder AddCell(string cellReference, object? value)
    {
        sheet.Cells.Add(new ExcelCellSpec(ExcelCellAddress.Parse(cellReference), value, ExcelStyleSpec.Empty));
        return this;
    }

    /// <summary>
    /// Adds a single static cell with style information.
    /// </summary>
    public ExcelSheetExportBuilder AddCell(string cellReference, object? value, Action<ExcelStyleBuilder> configureStyle)
    {
        NetStandardGuard.ThrowIfNull(configureStyle, nameof(configureStyle));
        var style = new ExcelStyleBuilder();
        configureStyle(style);
        sheet.Cells.Add(new ExcelCellSpec(ExcelCellAddress.Parse(cellReference), value, style.Spec));
        return this;
    }

    /// <summary>
    /// Adds a formula cell. The formula can be supplied with or without the leading equals sign.
    /// </summary>
    public ExcelSheetExportBuilder AddFormulaCell(string cellReference, string formula)
    {
        sheet.Cells.Add(new ExcelCellSpec(ExcelCellAddress.Parse(cellReference), new ExcelFormula(formula), ExcelStyleSpec.Empty));
        return this;
    }

    /// <summary>
    /// Adds a styled formula cell. The formula can be supplied with or without the leading equals sign.
    /// </summary>
    public ExcelSheetExportBuilder AddFormulaCell(string cellReference, string formula, Action<ExcelStyleBuilder> configureStyle)
    {
        NetStandardGuard.ThrowIfNull(configureStyle, nameof(configureStyle));
        var style = new ExcelStyleBuilder();
        configureStyle(style);
        sheet.Cells.Add(new ExcelCellSpec(ExcelCellAddress.Parse(cellReference), new ExcelFormula(formula), style.Spec));
        return this;
    }

    /// <summary>
    /// Adds a chart to this worksheet. Its data source may refer to this sheet or another sheet.
    /// </summary>
    public ExcelSheetExportBuilder AddChart(string title, Action<ExcelChartBuilder> configure)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(title, nameof(title));
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));

        var chart = new ExcelChartSpec(title);
        configure(new ExcelChartBuilder(chart));
        sheet.Charts.Add(chart);
        return this;
    }

    /// <summary>
    /// Adds a pivot table to this worksheet. Its source range may be on this sheet or another sheet.
    /// </summary>
    public ExcelSheetExportBuilder AddPivotTable(string name, Action<ExcelPivotTableBuilder> configure)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(name, nameof(name));
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));

        var pivotTable = new ExcelPivotTableSpec(name);
        configure(new ExcelPivotTableBuilder(pivotTable));
        sheet.PivotTables.Add(pivotTable);
        return this;
    }

    /// <summary>
    /// Configures a worksheet column by Excel column letter, such as A, B, or AA.
    /// </summary>
    public ExcelSheetColumnBuilder Column(string columnLetter)
    {
        return new ExcelSheetColumnBuilder(sheet, ExcelCellAddress.ColumnLetterToNumber(columnLetter));
    }

    /// <summary>
    /// Configures a worksheet row by one-based row number.
    /// </summary>
    public ExcelSheetRowBuilder Row(int rowNumber)
    {
        if (rowNumber < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rowNumber), "Row number must be one or greater.");
        }

        return new ExcelSheetRowBuilder(sheet, rowNumber);
    }

    /// <summary>
    /// Applies a style to a rectangular worksheet range such as A1:D5.
    /// </summary>
    public ExcelSheetExportBuilder RangeStyle(string range, Action<ExcelStyleBuilder> configure)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(range, nameof(range));
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        var (start, end) = ExcelCellAddress.ParseRange(range);
        sheet.RangeStyles.Add(new ExcelRangeStyleSpec(start, end, ExcelTableExportBuilder.CreateStyle(configure)));
        return this;
    }

    /// <summary>
    /// Freezes the specified number of top rows and left columns.
    /// </summary>
    public ExcelSheetExportBuilder FreezePanes(int rows = 1, int columns = 0)
    {
        if (rows < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(rows), "Frozen row count cannot be negative.");
        }

        if (columns < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(columns), "Frozen column count cannot be negative.");
        }

        if (rows == 0 && columns == 0)
        {
            sheet.FreezePane = null;
            return this;
        }

        sheet.FreezePane = new ExcelFreezePaneSpec(rows, columns);
        return this;
    }

    /// <summary>
    /// Applies an Excel autofilter to a worksheet range such as A3:F100.
    /// </summary>
    public ExcelSheetExportBuilder AutoFilter(string range)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(range, nameof(range));
        sheet.AutoFilterRange = range;
        return this;
    }

    /// <summary>
    /// Enables worksheet protection. The password uses Excel's legacy sheet/workbook protection hash,
    /// which prevents casual edits but is not file encryption.
    /// </summary>
    public ExcelSheetExportBuilder Protect(string? password = null, bool allowSelectLockedCells = true, bool allowSort = false, bool allowAutoFilter = false)
    {
        sheet.Protection = new ExcelSheetProtectionSpec(password, allowSelectLockedCells, allowSort, allowAutoFilter);
        return this;
    }

    /// <summary>
    /// Adds a cell comment/note.
    /// </summary>
    public ExcelSheetExportBuilder AddComment(string cellReference, string text, string author = "User")
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(text, nameof(text));
        NetStandardGuard.ThrowIfNullOrWhiteSpace(author, nameof(author));
        sheet.Comments.Add(new ExcelCommentSpec(ExcelCellAddress.Parse(cellReference), text, author));
        return this;
    }

    /// <summary>
    /// Adds a data validation rule to a worksheet range, such as a dropdown list or numeric limit.
    /// </summary>
    public ExcelSheetExportBuilder AddDataValidation(string range, Action<ExcelDataValidationBuilder> configure)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(range, nameof(range));
        NetStandardGuard.ThrowIfNull(configure, nameof(configure));
        var builder = new ExcelDataValidationBuilder(range);
        configure(builder);
        sheet.DataValidations.Add(builder.Build());
        return this;
    }

    /// <summary>
    /// Merges a worksheet range such as A1:D1.
    /// </summary>
    public ExcelSheetExportBuilder MergeCells(string range)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(range, nameof(range));
        if (!range.Contains(':', StringComparison.Ordinal))
        {
            throw new ArgumentException("Merge range must contain a colon, for example A1:D1.", nameof(range));
        }

        sheet.MergeRanges.Add(range);
        return this;
    }

    private ExcelTableExportBuilder AddTabularData(TabularData data)
    {
        var table = new ExcelTableSpec(data);
        sheet.Tables.Add(table);
        return new ExcelTableExportBuilder(table);
    }

    private ExcelRepeaterExportBuilder AddRepeaterData(TabularData data)
    {
        var repeater = new ExcelRepeaterSpec(data);
        sheet.Repeaters.Add(repeater);
        return new ExcelRepeaterExportBuilder(repeater);
    }
}

/// <summary>
/// Fluent builder for worksheet data validation rules.
/// </summary>
public sealed class ExcelDataValidationBuilder
{
    private readonly string range;
    private ExcelDataValidationKind kind = ExcelDataValidationKind.List;
    private ExcelDataValidationOperator operatorKind = ExcelDataValidationOperator.Between;
    private string? formula1;
    private string? formula2;
    private bool allowBlank = true;
    private string? inputTitle;
    private string? inputMessage;
    private string? errorTitle;
    private string? errorMessage;

    internal ExcelDataValidationBuilder(string range)
    {
        this.range = range;
    }

    /// <summary>
    /// Limits the cells to a fixed list of values, shown as a drop-down. Excel allows about 255 characters in total for a list typed into the rule; use <see cref="ListFromRange"/> for longer lists.
    /// </summary>
    /// <param name="values">The allowed values. A value may contain commas or quotes.</param>
    public ExcelDataValidationBuilder List(params string[] values)
    {
        kind = ExcelDataValidationKind.List;
        formula1 = $"\"{string.Join(",", values.Select(value => value.Replace("\"", "\"\"", StringComparison.Ordinal)))}\"";
        formula2 = null;
        return this;
    }

    /// <summary>
    /// Limits the cells to the values in a range of cells, shown as a drop-down.
    /// </summary>
    /// <param name="rangeReference">The range that holds the allowed values, for example Lists!$A$2:$A$10, or a defined name.</param>
    public ExcelDataValidationBuilder ListFromRange(string rangeReference)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(rangeReference, nameof(rangeReference));
        kind = ExcelDataValidationKind.List;
        formula1 = rangeReference;
        formula2 = null;
        return this;
    }

    /// <summary>
    /// Limits the cells to whole numbers from <paramref name="minimum"/> to <paramref name="maximum"/>, both allowed.
    /// </summary>
    /// <param name="minimum">The lowest allowed number.</param>
    /// <param name="maximum">The highest allowed number.</param>
    public ExcelDataValidationBuilder WholeNumberBetween(int minimum, int maximum)
    {
        return Comparison(ExcelDataValidationKind.WholeNumber, ExcelDataValidationOperator.Between, minimum.ToString(), maximum.ToString());
    }

    /// <summary>
    /// Limits the cells to decimal numbers from <paramref name="minimum"/> to <paramref name="maximum"/>, both allowed.
    /// </summary>
    /// <param name="minimum">The lowest allowed number.</param>
    /// <param name="maximum">The highest allowed number.</param>
    public ExcelDataValidationBuilder DecimalBetween(decimal minimum, decimal maximum)
    {
        return Comparison(ExcelDataValidationKind.Decimal, ExcelDataValidationOperator.Between, FormattableString.Invariant($"{minimum}"), FormattableString.Invariant($"{maximum}"));
    }

    /// <summary>
    /// Limits the cells to dates from <paramref name="minimum"/> to <paramref name="maximum"/>, both allowed. Only the date part matters.
    /// </summary>
    /// <param name="minimum">The earliest allowed date.</param>
    /// <param name="maximum">The latest allowed date.</param>
    public ExcelDataValidationBuilder DateBetween(DateTime minimum, DateTime maximum)
    {
        return Comparison(ExcelDataValidationKind.Date, ExcelDataValidationOperator.Between, minimum.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), maximum.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Limits the length of the text in the cells.
    /// </summary>
    /// <param name="minimum">The shortest allowed length, in characters.</param>
    /// <param name="maximum">The longest allowed length, in characters.</param>
    public ExcelDataValidationBuilder TextLengthBetween(int minimum, int maximum)
    {
        return Comparison(ExcelDataValidationKind.TextLength, ExcelDataValidationOperator.Between, minimum.ToString(), maximum.ToString());
    }

    /// <summary>
    /// Allows a value only when a formula is true. The formula is written for the first cell of the range and adjusts for the other cells, as in Excel.
    /// </summary>
    /// <param name="formula">The formula, with or without a leading equals sign.</param>
    public ExcelDataValidationBuilder CustomFormula(string formula)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(formula, nameof(formula));
        kind = ExcelDataValidationKind.Custom;
        formula1 = formula.TrimStart('=');
        formula2 = null;
        return this;
    }

    /// <summary>
    /// Says whether a blank cell is accepted. The default is true.
    /// </summary>
    /// <param name="value">True to accept blanks, false to reject them.</param>
    public ExcelDataValidationBuilder AllowBlank(bool value = true)
    {
        allowBlank = value;
        return this;
    }

    /// <summary>
    /// Shows a tip when a cell in the range is selected. Excel limits the title to 32 characters and the message to 255.
    /// </summary>
    /// <param name="title">The tip's heading.</param>
    /// <param name="message">The tip's text.</param>
    public ExcelDataValidationBuilder InputMessage(string title, string message)
    {
        inputTitle = title;
        inputMessage = message;
        return this;
    }

    /// <summary>
    /// Sets the message shown when someone enters a value the rule rejects. Excel limits the title to 32 characters and the message to 255.
    /// </summary>
    /// <param name="title">The message box heading.</param>
    /// <param name="message">The message text.</param>
    public ExcelDataValidationBuilder ErrorMessage(string title, string message)
    {
        errorTitle = title;
        errorMessage = message;
        return this;
    }

    internal ExcelDataValidationSpec Build()
    {
        return new ExcelDataValidationSpec(range, kind, operatorKind, formula1, formula2, allowBlank, inputTitle, inputMessage, errorTitle, errorMessage);
    }

    private ExcelDataValidationBuilder Comparison(ExcelDataValidationKind validationKind, ExcelDataValidationOperator comparisonOperator, string first, string second)
    {
        kind = validationKind;
        operatorKind = comparisonOperator;
        formula1 = first;
        formula2 = second;
        return this;
    }
}
