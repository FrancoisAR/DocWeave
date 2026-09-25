using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

/// <summary>
/// Writes chart parts and the drawing anchors that place them on a worksheet.
/// </summary>
internal static class ExcelChartRenderer
{
    internal static void AddCharts(WorksheetPart worksheetPart, ExcelSheetSpec sheetSpec)
    {
        // One drawing part holds every chart on the sheet. Each chart gets its own chart part, referenced from an anchor.
        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        var drawingRelationshipId = worksheetPart.GetIdOfPart(drawingsPart);
        worksheetPart.Worksheet.Append(new Drawing { Id = drawingRelationshipId });

        for (var i = 0; i < sheetSpec.Charts.Count; i++)
        {
            var chart = sheetSpec.Charts[i];
            if (chart.DataSource is null)
            {
                throw new InvalidOperationException($"Chart '{chart.Title}' does not have a data source.");
            }

            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = CreateChartSpace(chart);
            chartPart.ChartSpace.Save();
            // The chart's drawing id in the sheet is i + 2 (so the first chart is 2).
            AddChartAnchor(drawingsPart, chart, drawingsPart.GetIdOfPart(chartPart), (uint)(i + 2));
        }

        drawingsPart.WorksheetDrawing.Save();
    }

    private static C.ChartSpace CreateChartSpace(ExcelChartSpec chart)
    {
        var chartSpace = new C.ChartSpace();
        chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
        chartSpace.Append(new C.Style { Val = (byte)Math.Min(chart.Style, byte.MaxValue) });

        var chartElement = chartSpace.AppendChild(new C.Chart());
        chartElement.Append(CreateChartTitle(chart.Title));

        var plotArea = chartElement.AppendChild(new C.PlotArea());
        plotArea.Append(new C.Layout());

        // Bar and column charts share the bar chart element, so Bar/Column and any unlisted type fall through to it.
        if (chart.Type == ExcelChartType.Line)
        {
            plotArea.Append(CreateLineChart(chart));
        }
        else if (chart.Type == ExcelChartType.Area)
        {
            plotArea.Append(CreateAreaChart(chart));
        }
        else if (chart.Type == ExcelChartType.Pie)
        {
            plotArea.Append(CreatePieChart(chart));
        }
        else if (chart.Type == ExcelChartType.Doughnut)
        {
            plotArea.Append(CreateDoughnutChart(chart));
        }
        else if (chart.Type == ExcelChartType.Scatter)
        {
            plotArea.Append(CreateScatterChart(chart));
        }
        else
        {
            plotArea.Append(CreateBarChart(chart));
        }

        // Pie and doughnut charts have no axes. The others need a pair of axes that refer to the chart's axis ids.
        if (chart.Type == ExcelChartType.Scatter)
        {
            AppendScatterAxes(plotArea, chart);
        }
        else if (chart.Type is not (ExcelChartType.Pie or ExcelChartType.Doughnut))
        {
            AppendAxes(plotArea, chart);
        }

        if (chart.ShowLegend)
        {
            chartElement.Append(new C.Legend(
                new C.LegendPosition { Val = ToOpenXmlLegendPosition(chart.LegendPosition) },
                new C.Layout()));
        }

        chartElement.Append(new C.PlotVisibleOnly { Val = true });

        return chartSpace;
    }

    private static OpenXmlElement CreateBarChart(ExcelChartSpec chart)
    {
        var barChart = new C.BarChart(
            new C.BarDirection { Val = chart.Type == ExcelChartType.Bar ? C.BarDirectionValues.Bar : C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            CreateBarSeries(chart));
        AppendDataLabels(barChart, chart);
        // Every chart type with axes uses the same two ids, 48650112 for the category (or X) axis and 48672768 for the value (or Y) axis.
        // The axes written in AppendAxes / AppendScatterAxes must carry the same ids, and refer to each other through CrossingAxis.
        barChart.Append(new C.AxisId { Val = 48650112U }, new C.AxisId { Val = 48672768U });
        return barChart;
    }

    private static OpenXmlElement CreateLineChart(ExcelChartSpec chart)
    {
        var lineChart = new C.LineChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            CreateLineSeries(chart));
        AppendDataLabels(lineChart, chart);
        lineChart.Append(new C.AxisId { Val = 48650112U }, new C.AxisId { Val = 48672768U });
        return lineChart;
    }

    private static OpenXmlElement CreateAreaChart(ExcelChartSpec chart)
    {
        var areaChart = new C.AreaChart(
            new C.Grouping { Val = C.GroupingValues.Standard },
            CreateAreaSeries(chart));
        AppendDataLabels(areaChart, chart);
        areaChart.Append(new C.AxisId { Val = 48650112U }, new C.AxisId { Val = 48672768U });
        return areaChart;
    }

    private static OpenXmlElement CreatePieChart(ExcelChartSpec chart)
    {
        var pieChart = new C.PieChart(
            new C.VaryColors { Val = true },
            CreatePieSeries(chart));
        AppendDataLabels(pieChart, chart);
        return pieChart;
    }

    private static OpenXmlElement CreateDoughnutChart(ExcelChartSpec chart)
    {
        var doughnutChart = new C.DoughnutChart(
            new C.VaryColors { Val = true },
            CreatePieSeries(chart),
            new C.HoleSize { Val = 55 });
        AppendDataLabels(doughnutChart, chart);
        return doughnutChart;
    }

    private static OpenXmlElement CreateScatterChart(ExcelChartSpec chart)
    {
        var scatterChart = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            CreateScatterSeries(chart));
        AppendDataLabels(scatterChart, chart);
        scatterChart.Append(new C.AxisId { Val = 48650112U }, new C.AxisId { Val = 48672768U });
        return scatterChart;
    }

    private static C.BarChartSeries CreateBarSeries(ExcelChartSpec chart)
    {
        var series = new C.BarChartSeries(
            new C.Index { Val = 0U },
            new C.Order { Val = 0U },
            // Each chart has a single series, named after the chart title.
            new C.SeriesText(new C.NumericValue { Text = chart.Title }),
            CreateSeriesShapeProperties(chart),
            new C.CategoryAxisData(new C.StringReference(new C.Formula(ToSheetFormula(chart.DataSource!.SheetName, chart.DataSource.CategoryRange)))),
            new C.Values(new C.NumberReference(new C.Formula(ToSheetFormula(chart.DataSource.SheetName, chart.DataSource.ValueRange)))));

        return series;
    }

    private static C.LineChartSeries CreateLineSeries(ExcelChartSpec chart)
    {
        return new C.LineChartSeries(
            new C.Index { Val = 0U },
            new C.Order { Val = 0U },
            new C.SeriesText(new C.NumericValue { Text = chart.Title }),
            CreateSeriesShapeProperties(chart),
            new C.CategoryAxisData(new C.StringReference(new C.Formula(ToSheetFormula(chart.DataSource!.SheetName, chart.DataSource.CategoryRange)))),
            new C.Values(new C.NumberReference(new C.Formula(ToSheetFormula(chart.DataSource.SheetName, chart.DataSource.ValueRange)))));
    }

    private static C.AreaChartSeries CreateAreaSeries(ExcelChartSpec chart)
    {
        return new C.AreaChartSeries(
            new C.Index { Val = 0U },
            new C.Order { Val = 0U },
            new C.SeriesText(new C.NumericValue { Text = chart.Title }),
            CreateSeriesShapeProperties(chart),
            new C.CategoryAxisData(new C.StringReference(new C.Formula(ToSheetFormula(chart.DataSource!.SheetName, chart.DataSource.CategoryRange)))),
            new C.Values(new C.NumberReference(new C.Formula(ToSheetFormula(chart.DataSource.SheetName, chart.DataSource.ValueRange)))));
    }

    private static C.PieChartSeries CreatePieSeries(ExcelChartSpec chart)
    {
        return new C.PieChartSeries(
            new C.Index { Val = 0U },
            new C.Order { Val = 0U },
            new C.SeriesText(new C.NumericValue { Text = chart.Title }),
            CreateSeriesShapeProperties(chart),
            new C.CategoryAxisData(new C.StringReference(new C.Formula(ToSheetFormula(chart.DataSource!.SheetName, chart.DataSource.CategoryRange)))),
            new C.Values(new C.NumberReference(new C.Formula(ToSheetFormula(chart.DataSource.SheetName, chart.DataSource.ValueRange)))));
    }

    private static C.ScatterChartSeries CreateScatterSeries(ExcelChartSpec chart)
    {
        return new C.ScatterChartSeries(
            new C.Index { Val = 0U },
            new C.Order { Val = 0U },
            new C.SeriesText(new C.NumericValue { Text = chart.Title }),
            CreateSeriesShapeProperties(chart),
            new C.XValues(new C.NumberReference(new C.Formula(ToSheetFormula(chart.DataSource!.SheetName, chart.DataSource.CategoryRange)))),
            new C.YValues(new C.NumberReference(new C.Formula(ToSheetFormula(chart.DataSource.SheetName, chart.DataSource.ValueRange)))),
            new C.Smooth { Val = false });
    }

    private static C.ChartShapeProperties CreateSeriesShapeProperties(ExcelChartSpec chart)
    {
        var properties = new C.ChartShapeProperties();
        if (chart.SeriesColor is not null)
        {
            // The series color is stored as AARRGGBB. The drawing element takes RRGGBB, so the alpha pair is dropped.
            properties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = chart.SeriesColor[2..] }));
        }

        return properties;
    }

    private static void AppendDataLabels(OpenXmlCompositeElement chartElement, ExcelChartSpec chart)
    {
        if (!chart.ShowDataLabels)
        {
            return;
        }

        chartElement.Append(new C.DataLabels(
            new C.ShowLegendKey { Val = false },
            new C.ShowValue { Val = true },
            new C.ShowCategoryName { Val = false },
            new C.ShowSeriesName { Val = false },
            new C.ShowPercent { Val = false },
            new C.ShowBubbleSize { Val = false }));
    }

    private static void AppendAxes(C.PlotArea plotArea, ExcelChartSpec chart)
    {
        // The axis title goes right after the axis position, which is where the schema puts it.
        var categoryAxis = new C.CategoryAxis(
            new C.AxisId { Val = 48650112U },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = 48672768U },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 });
        if (!string.IsNullOrWhiteSpace(chart.CategoryAxisTitle))
        {
            categoryAxis.InsertAfter(CreateChartTitle(chart.CategoryAxisTitle), categoryAxis.GetFirstChild<C.AxisPosition>());
        }

        plotArea.Append(categoryAxis);

        var valueAxis = new C.ValueAxis(
            new C.AxisId { Val = 48672768U },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.AxisPosition { Val = C.AxisPositionValues.Left },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = 48650112U },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        if (!string.IsNullOrWhiteSpace(chart.ValueAxisTitle))
        {
            valueAxis.InsertAfter(CreateChartTitle(chart.ValueAxisTitle), (OpenXmlElement?)valueAxis.GetFirstChild<C.MajorGridlines>() ?? valueAxis.GetFirstChild<C.AxisPosition>());
        }

        plotArea.Append(valueAxis);
    }

    private static void AppendScatterAxes(C.PlotArea plotArea, ExcelChartSpec chart)
    {
        // A scatter chart has two value axes instead of a category axis and a value axis.
        var xAxis = new C.ValueAxis(
            new C.AxisId { Val = 48650112U },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = 48672768U },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        if (!string.IsNullOrWhiteSpace(chart.CategoryAxisTitle))
        {
            xAxis.InsertAfter(CreateChartTitle(chart.CategoryAxisTitle), (OpenXmlElement?)xAxis.GetFirstChild<C.MajorGridlines>() ?? xAxis.GetFirstChild<C.AxisPosition>());
        }

        plotArea.Append(xAxis);

        var yAxis = new C.ValueAxis(
            new C.AxisId { Val = 48672768U },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.AxisPosition { Val = C.AxisPositionValues.Left },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = 48650112U },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        if (!string.IsNullOrWhiteSpace(chart.ValueAxisTitle))
        {
            yAxis.InsertAfter(CreateChartTitle(chart.ValueAxisTitle), (OpenXmlElement?)yAxis.GetFirstChild<C.MajorGridlines>() ?? yAxis.GetFirstChild<C.AxisPosition>());
        }

        plotArea.Append(yAxis);
    }

    private static C.LegendPositionValues ToOpenXmlLegendPosition(ExcelChartLegendPosition position)
    {
        return position switch
        {
            ExcelChartLegendPosition.Left => C.LegendPositionValues.Left,
            ExcelChartLegendPosition.Top => C.LegendPositionValues.Top,
            ExcelChartLegendPosition.Bottom => C.LegendPositionValues.Bottom,
            _ => C.LegendPositionValues.Right
        };
    }

    private static C.Title CreateChartTitle(string title)
    {
        return new C.Title(
            new C.ChartText(
                new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(
                            new A.RunProperties { Language = "en-US" },
                            new A.Text(title))))),
            new C.Layout(),
            new C.Overlay { Val = false });
    }

    private static void AddChartAnchor(DrawingsPart drawingsPart, ExcelChartSpec chart, string chartRelationshipId, uint drawingId)
    {
        var graphicFrame = new Xdr.GraphicFrame(
            new Xdr.NonVisualGraphicFrameProperties(
                new Xdr.NonVisualDrawingProperties { Id = drawingId, Name = chart.Title },
                new Xdr.NonVisualGraphicFrameDrawingProperties()),
            new Xdr.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 0L, Cy = 0L }),
            new A.Graphic(
                new A.GraphicData(
                    new C.ChartReference { Id = chartRelationshipId })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }));

        drawingsPart.WorksheetDrawing!.Append(new Xdr.TwoCellAnchor(
            CreateFromMarker(chart.TopLeft),
            CreateToMarker(chart.BottomRight),
            graphicFrame,
            new Xdr.ClientData()));
    }

    // Anchor markers count rows and columns from zero, cell addresses from one, hence the - 1. The chart starts at the top-left
    // corner of its first cell and, in the "to" marker, ends at the top-left corner of the bottom-right cell.
    private static Xdr.FromMarker CreateFromMarker(ExcelCellAddress address)
    {
        return new Xdr.FromMarker(
            new Xdr.ColumnId((address.Column - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId((address.Row - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset("0"));
    }

    private static Xdr.ToMarker CreateToMarker(ExcelCellAddress address)
    {
        return new Xdr.ToMarker(
            new Xdr.ColumnId((address.Column - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.ColumnOffset("0"),
            new Xdr.RowId((address.Row - 1).ToString(CultureInfo.InvariantCulture)),
            new Xdr.RowOffset("0"));
    }

    private static string ToSheetFormula(string sheetName, string range)
    {
        var safeSheetName = sheetName.Replace("'", "''", StringComparison.Ordinal);
        return $"'{safeSheetName}'!{range}";
    }
}
