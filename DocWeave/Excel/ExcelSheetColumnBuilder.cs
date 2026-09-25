namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for worksheet-level column settings.
/// </summary>
public sealed class ExcelSheetColumnBuilder
{
    private readonly ExcelSheetSpec sheet;
    private readonly int columnNumber;

    internal ExcelSheetColumnBuilder(ExcelSheetSpec sheet, int columnNumber)
    {
        this.sheet = sheet;
        this.columnNumber = columnNumber;
    }

    /// <summary>
    /// Sets a default style for generated cells in this worksheet column.
    /// </summary>
    public ExcelSheetColumnBuilder Style(Action<ExcelStyleBuilder> configure)
    {
        ArgumentNullException.ThrowIfNull(configure);
        sheet.ColumnStyles[columnNumber] = ExcelTableExportBuilder.CreateStyle(configure);
        return this;
    }

    /// <summary>
    /// Sets the column width in characters, the unit Excel shows in its own column-width box.
    /// </summary>
    /// <param name="width">The width. It must be greater than zero.</param>
    /// <exception cref="System.ArgumentOutOfRangeException">The width is zero or negative.</exception>
    public ExcelSheetColumnBuilder Width(double width)
    {
        if (width <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(width), "Column width must be greater than zero.");
        }

        Settings.Width = width;
        return this;
    }

    /// <summary>
    /// Sets Excel's best-fit flag on the column. The library does not measure the content, so also set <see cref="Width"/> when the width matters.
    /// </summary>
    /// <param name="value">True to set the flag.</param>
    public ExcelSheetColumnBuilder AutoFitWidth(bool value = true)
    {
        Settings.BestFit = value;
        return this;
    }

    /// <summary>
    /// Hides the column.
    /// </summary>
    /// <param name="value">True to hide the column, false to show it.</param>
    public ExcelSheetColumnBuilder Hidden(bool value = true)
    {
        Settings.Hidden = value;
        return this;
    }

    private ExcelColumnSettingsSpec Settings
    {
        get
        {
            if (!sheet.ColumnSettings.TryGetValue(columnNumber, out var settings))
            {
                settings = new ExcelColumnSettingsSpec();
                sheet.ColumnSettings[columnNumber] = settings;
            }

            return settings;
        }
    }
}
