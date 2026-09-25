namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for worksheet-level row settings.
/// </summary>
public sealed class ExcelSheetRowBuilder
{
    private readonly ExcelSheetSpec sheet;
    private readonly int rowNumber;

    internal ExcelSheetRowBuilder(ExcelSheetSpec sheet, int rowNumber)
    {
        this.sheet = sheet;
        this.rowNumber = rowNumber;
    }

    /// <summary>
    /// Sets the row height in points.
    /// </summary>
    /// <param name="height">The height. It must be greater than zero.</param>
    /// <exception cref="System.ArgumentOutOfRangeException">The height is zero or negative.</exception>
    public ExcelSheetRowBuilder Height(double height)
    {
        if (height <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(height), "Row height must be greater than zero.");
        }

        Settings.Height = height;
        return this;
    }

    /// <summary>
    /// Lets Excel size the row to fit its content instead of using a fixed height.
    /// </summary>
    /// <param name="value">True to fit the content.</param>
    public ExcelSheetRowBuilder AutoFitHeight(bool value = true)
    {
        Settings.AutoFit = value;
        return this;
    }

    /// <summary>
    /// Hides the row.
    /// </summary>
    /// <param name="value">True to hide the row, false to show it.</param>
    public ExcelSheetRowBuilder Hidden(bool value = true)
    {
        Settings.Hidden = value;
        return this;
    }

    private ExcelRowSettingsSpec Settings
    {
        get
        {
            if (!sheet.RowSettings.TryGetValue(rowNumber, out var settings))
            {
                settings = new ExcelRowSettingsSpec();
                sheet.RowSettings[rowNumber] = settings;
            }

            return settings;
        }
    }
}
