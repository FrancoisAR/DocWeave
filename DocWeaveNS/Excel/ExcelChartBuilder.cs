namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for an Excel chart placed on a worksheet.
/// </summary>
public sealed class ExcelChartBuilder
{
    private readonly ExcelChartSpec chart;

    internal ExcelChartBuilder(ExcelChartSpec chart)
    {
        this.chart = chart;
    }

    /// <summary>
    /// Sets the chart type.
    /// </summary>
    public ExcelChartBuilder Type(ExcelChartType type)
    {
        chart.Type = type;
        return this;
    }

    /// <summary>
    /// Sets the source sheet and data ranges for the chart.
    /// </summary>
    public ExcelChartBuilder DataSource(string sheetName, string categoryRange, string valueRange)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(sheetName, nameof(sheetName));
        NetStandardGuard.ThrowIfNullOrWhiteSpace(categoryRange, nameof(categoryRange));
        NetStandardGuard.ThrowIfNullOrWhiteSpace(valueRange, nameof(valueRange));
        chart.DataSource = new ExcelChartDataSource(sheetName, categoryRange, valueRange);
        return this;
    }

    /// <summary>
    /// Sets the top-left and bottom-right chart placement cells on the target worksheet.
    /// </summary>
    public ExcelChartBuilder Position(string topLeftCell, string bottomRightCell)
    {
        chart.TopLeft = ExcelCellAddress.Parse(topLeftCell);
        chart.BottomRight = ExcelCellAddress.Parse(bottomRightCell);
        return this;
    }

    /// <summary>
    /// Sets a built-in Excel chart style number.
    /// </summary>
    public ExcelChartBuilder Style(uint style)
    {
        chart.Style = style;
        return this;
    }

    /// <summary>
    /// Sets the primary series color from a hex value such as #4472C4.
    /// </summary>
    public ExcelChartBuilder SeriesColor(string hexColor)
    {
        chart.SeriesColor = ExcelColor.Normalize(hexColor);
        return this;
    }

    /// <summary>
    /// Controls whether the chart legend is displayed.
    /// </summary>
    public ExcelChartBuilder ShowLegend(bool show = true)
    {
        chart.ShowLegend = show;
        return this;
    }

    /// <summary>
    /// Sets the chart legend position.
    /// </summary>
    public ExcelChartBuilder LegendPosition(ExcelChartLegendPosition position)
    {
        chart.LegendPosition = position;
        return this;
    }

    /// <summary>
    /// Shows value labels on the chart series.
    /// </summary>
    public ExcelChartBuilder ShowDataLabels(bool show = true)
    {
        chart.ShowDataLabels = show;
        return this;
    }

    /// <summary>
    /// Sets category and value axis titles for chart types that use axes.
    /// </summary>
    public ExcelChartBuilder AxisTitles(string? categoryAxisTitle = null, string? valueAxisTitle = null)
    {
        chart.CategoryAxisTitle = categoryAxisTitle;
        chart.ValueAxisTitle = valueAxisTitle;
        return this;
    }
}

/// <summary>
/// Supported chart types for the initial Excel chart export layer.
/// </summary>
public enum ExcelChartType
{
    /// <summary>
    /// Vertical bars, one per category.
    /// </summary>
    Column,
    /// <summary>
    /// Horizontal bars, one per category.
    /// </summary>
    Bar,
    /// <summary>
    /// A line through the values in category order.
    /// </summary>
    Line,
    /// <summary>
    /// Like a line, with the area under it filled.
    /// </summary>
    Area,
    /// <summary>
    /// One slice per category, sized by its share of the total. Has no axes.
    /// </summary>
    Pie,
    /// <summary>
    /// Like a pie with a hole in the middle. Has no axes.
    /// </summary>
    Doughnut,
    /// <summary>
    /// Points placed by a number on each axis. The category range should hold numbers rather than labels.
    /// </summary>
    Scatter
}

/// <summary>
/// Where the chart legend is drawn.
/// </summary>
public enum ExcelChartLegendPosition
{
    /// <summary>
    /// To the right of the chart. This is the default.
    /// </summary>
    Right,
    /// <summary>
    /// To the left of the chart.
    /// </summary>
    Left,
    /// <summary>
    /// Above the chart.
    /// </summary>
    Top,
    /// <summary>
    /// Below the chart.
    /// </summary>
    Bottom
}
