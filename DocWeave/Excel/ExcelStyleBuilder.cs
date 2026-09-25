namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for basic Excel cell styling.
/// </summary>
public sealed class ExcelStyleBuilder
{
    internal ExcelStyleSpec Spec { get; private set; } = ExcelStyleSpec.Empty;

    /// <summary>
    /// Sets bold font rendering.
    /// </summary>
    public ExcelStyleBuilder Bold(bool value = true)
    {
        Spec = Spec with { Bold = value };
        return this;
    }

    /// <summary>
    /// Sets italic font rendering.
    /// </summary>
    public ExcelStyleBuilder Italic(bool value = true)
    {
        Spec = Spec with { Italic = value };
        return this;
    }

    /// <summary>
    /// Sets underline font rendering.
    /// </summary>
    public ExcelStyleBuilder Underline(bool value = true)
    {
        Spec = Spec with { Underline = value };
        return this;
    }

    /// <summary>
    /// Sets the font size in points.
    /// </summary>
    public ExcelStyleBuilder FontSize(double value)
    {
        if (value <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(value), "Font size must be greater than zero.");
        }

        Spec = Spec with { FontSize = value };
        return this;
    }

    /// <summary>
    /// Sets the font family name, such as Calibri, Arial, or Aptos.
    /// </summary>
    public ExcelStyleBuilder FontFamily(string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        Spec = Spec with { FontFamily = name };
        return this;
    }

    /// <summary>
    /// Sets the font color from a hex value such as #FFFFFF or FF1F4E78.
    /// </summary>
    public ExcelStyleBuilder FontColor(string hexColor)
    {
        Spec = Spec with { FontColor = ExcelColor.Normalize(hexColor) };
        return this;
    }

    /// <summary>
    /// Sets a solid background fill color from a hex value such as #D9EAF7.
    /// </summary>
    public ExcelStyleBuilder BackgroundColor(string hexColor)
    {
        Spec = Spec with { BackgroundColor = ExcelColor.Normalize(hexColor) };
        return this;
    }

    /// <summary>
    /// Sets a patterned fill using the supplied foreground color.
    /// </summary>
    public ExcelStyleBuilder PatternFill(ExcelFillPattern pattern, string hexColor)
    {
        Spec = Spec with { FillPattern = pattern, BackgroundColor = ExcelColor.Normalize(hexColor) };
        return this;
    }

    /// <summary>
    /// Controls whether cell text wraps within the cell.
    /// </summary>
    public ExcelStyleBuilder WrapText(bool value = true)
    {
        Spec = Spec with { WrapText = value };
        return this;
    }

    /// <summary>
    /// Rotates displayed text. Excel supports 0 to 180 degrees.
    /// </summary>
    public ExcelStyleBuilder RotateText(int degrees)
    {
        if (degrees is < 0 or > 180)
        {
            throw new ArgumentOutOfRangeException(nameof(degrees), "Text rotation must be between 0 and 180 degrees.");
        }

        Spec = Spec with { TextRotation = degrees };
        return this;
    }

    /// <summary>
    /// Sets the display number format for the cell.
    /// </summary>
    public ExcelStyleBuilder NumberFormat(ExcelNumberFormat format)
    {
        Spec = Spec with { NumberFormat = format };
        return this;
    }

    /// <summary>
    /// Sets the same border style and color on all sides of the cell.
    /// </summary>
    public ExcelStyleBuilder Border(ExcelBorderLineStyle style = ExcelBorderLineStyle.Thin, string color = "#000000")
    {
        Spec = Spec with
        {
            BorderStyle = style,
            BorderColor = ExcelColor.Normalize(color)
        };
        return this;
    }
}

/// <summary>
/// Supported border styles for the initial Excel export styling layer.
/// </summary>
public enum ExcelBorderLineStyle
{
    /// <summary>
    /// No border.
    /// </summary>
    None,
    /// <summary>
    /// A thin solid line.
    /// </summary>
    Thin,
    /// <summary>
    /// A medium-weight solid line.
    /// </summary>
    Medium,
    /// <summary>
    /// A thick solid line.
    /// </summary>
    Thick,
    /// <summary>
    /// A dashed line.
    /// </summary>
    Dashed,
    /// <summary>
    /// A dotted line.
    /// </summary>
    Dotted,
    /// <summary>
    /// Two thin parallel lines.
    /// </summary>
    Double,
    /// <summary>
    /// The finest line Excel can draw, shown as a row of tiny dots.
    /// </summary>
    Hair,
    /// <summary>
    /// A line of alternating dashes and dots.
    /// </summary>
    DashDot,
    /// <summary>
    /// A line of a dash followed by two dots, repeated.
    /// </summary>
    DashDotDot,
    /// <summary>
    /// A slanted line of alternating dashes and dots.
    /// </summary>
    SlantDashDot
}

/// <summary>
/// Fill patterns for cell backgrounds. Anything other than <see cref="Solid"/> draws the pattern in the foreground color over the background.
/// </summary>
public enum ExcelFillPattern
{
    /// <summary>
    /// A flat fill in a single color.
    /// </summary>
    Solid,
    /// <summary>
    /// A very light dotted pattern (12.5% gray).
    /// </summary>
    Gray125,
    /// <summary>
    /// A dense pattern (75% gray).
    /// </summary>
    DarkGray,
    /// <summary>
    /// A medium pattern (50% gray).
    /// </summary>
    MediumGray,
    /// <summary>
    /// A light pattern (25% gray).
    /// </summary>
    LightGray,
    /// <summary>
    /// Dark horizontal stripes.
    /// </summary>
    DarkHorizontal,
    /// <summary>
    /// Dark vertical stripes.
    /// </summary>
    DarkVertical,
    /// <summary>
    /// A dark grid.
    /// </summary>
    DarkGrid,
    /// <summary>
    /// A light grid.
    /// </summary>
    LightGrid
}

/// <summary>
/// How a numeric or date value is displayed. Only the display changes, never the stored value.
/// </summary>
public enum ExcelNumberFormat
{
    /// <summary>
    /// No specific format: Excel shows the value as it is.
    /// </summary>
    General,
    /// <summary>
    /// A number with a thousands separator and two decimals (#,##0.00).
    /// </summary>
    Number,
    /// <summary>
    /// A number with a dollar sign, a thousands separator and two decimals ("$"#,##0.00).
    /// </summary>
    Currency,
    /// <summary>
    /// A percentage with two decimals (0.00%). The value 0.25 is shown as 25.00%.
    /// </summary>
    Percent,
    /// <summary>
    /// Excel's built-in short date format. The exact layout follows the reader's regional settings.
    /// </summary>
    Date,
    /// <summary>
    /// Text (@). The cell content is treated as text even when it looks like a number.
    /// </summary>
    Text
}
