using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "chart", "title": "...", "chartType": "column", "dataSource": {...}, "position": {...} }
/// </summary>
internal sealed class ChartTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "chart";

    public override bool OccupiesCells => true;

    public override TemplateRegionBounds EstimateBounds(JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startOverride)
    {
        return TemplateRegionBounds.Between(region.GetProperty("position").GetRequiredString("topLeft"), region.GetProperty("position").GetRequiredString("bottomRight"));
    }

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var chart = new ExcelChartSpec(templateContext.BindText(region.GetRequiredString("title")))
        {
            Type = NetStandardEnum.Parse<ExcelChartType>(region.GetRequiredString("chartType"), ignoreCase: true)
        };

        var dataSource = region.GetProperty("dataSource");
        chart.DataSource = new ExcelChartDataSource(
            dataSource.GetRequiredString("sheet"),
            dataSource.GetRequiredString("categoryRange"),
            dataSource.GetRequiredString("valueRange"));

        if (region.TryGetProperty("position", out var position))
        {
            chart.TopLeft = ExcelCellAddress.Parse(position.GetRequiredString("topLeft"));
            chart.BottomRight = ExcelCellAddress.Parse(position.GetRequiredString("bottomRight"));
        }

        if (region.TryGetProperty("style", out var style))
        {
            chart.Style = style.GetUInt32();
        }

        var color = region.GetOptionalString("seriesColor");
        if (!string.IsNullOrWhiteSpace(color))
        {
            chart.SeriesColor = ExcelColor.Normalize(color);
        }

        if (region.TryGetProperty("showLegend", out var showLegend))
        {
            chart.ShowLegend = showLegend.GetBoolean();
        }

        var legendPosition = region.GetOptionalString("legendPosition");
        if (!string.IsNullOrWhiteSpace(legendPosition))
        {
            chart.LegendPosition = NetStandardEnum.Parse<ExcelChartLegendPosition>(legendPosition, ignoreCase: true);
        }

        chart.ShowDataLabels = region.GetOptionalBool("showDataLabels") ?? chart.ShowDataLabels;
        if (region.TryGetProperty("axisTitles", out var axisTitles))
        {
            chart.CategoryAxisTitle = axisTitles.GetOptionalString("category");
            chart.ValueAxisTitle = axisTitles.GetOptionalString("value");
        }

        sheet.Charts.Add(chart);
    }
}
