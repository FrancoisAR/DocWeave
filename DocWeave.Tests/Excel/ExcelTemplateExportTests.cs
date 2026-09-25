using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocWeave.Excel;
using static DocWeave.Tests.Excel.Xlsx;

namespace DocWeave.Tests.Excel;

public sealed class ExcelTemplateExportTests
{
    [Fact]
    public void _27_Template_exports_styled_table_with_totals()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#B4C6E7" } },
            "alternate": { "fill": { "color": "#DDEBF7" }, "border": { "style": "thin", "color": "#B4C6E7" } },
            "currency": { "numberFormat": "currency", "border": { "style": "thin", "color": "#70AD47" } },
            "totals": { "font": { "bold": true }, "fill": { "color": "#D9EAD3" }, "border": { "style": "thick", "color": "#000000" } }
          },
          "sheetTemplates": {
            "salesSheet": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Template Sales Export", "style": "title" },
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "A3",
                  "headers": true,
                  "columns": [
                    { "source": "Region", "caption": "Region" },
                    { "source": "Income", "caption": "Income", "style": "currency" },
                    { "source": "Costs", "caption": "Costs", "style": "currency" }
                  ],
                  "styles": { "header": "header", "data": "data", "alternateRow": "alternate" },
                  "totals": { "position": "bottom", "label": "Totals", "style": "totals" }
                }
              ]
            }
          },
          "sheets": [ { "name": "Sales", "template": "salesSheet" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_27_Template_exports_styled_table_with_totals), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("Template Sales Export", GetCellText(bytes, "Sales", "A1"));
        Assert.Equal("Region", GetCellText(bytes, "Sales", "A3"));
        Assert.Equal("SUM(B4:B7)", GetCellFormula(bytes, "Sales", "B8"));
        Assert.Equal("SUM(C4:C7)", GetCellFormula(bytes, "Sales", "C8"));
        Assert.True(CountStyledCells(bytes, "Sales") > 12);
    }

    [Fact]
    public void _28_Template_exports_repeater_cards()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
            "card": { "fill": { "color": "#F8FBFD" }, "border": { "style": "thin", "color": "#B7C9D6" } },
            "cardTitle": { "font": { "size": 13, "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#1F4E78" } },
            "label": { "font": { "bold": true, "color": "#385723" }, "fill": { "color": "#E2F0D9" }, "border": { "style": "thin", "color": "#A9D18E" } },
            "value": { "border": { "style": "thin", "color": "#D9E2F3" } },
            "number": { "numberFormat": "number", "fill": { "color": "#DDEBF7" }, "border": { "style": "medium", "color": "#5B9BD5" } }
          },
          "sheetTemplates": {
            "currencyCards": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Template Currency Cards", "style": "title" },
                {
                  "type": "repeater",
                  "source": "currencies",
                  "startCell": "A4",
                  "itemsPerRow": 2,
                  "itemSize": { "columns": 4, "rows": 5 },
                  "gap": { "columns": 1, "rows": 1 },
                  "titleField": "Code",
                  "fields": [
                    { "source": "Country", "caption": "Country" },
                    { "source": "LastHighRate", "caption": "Last High", "valueStyle": "number" },
                    { "source": "LastLowRate", "caption": "Last Low", "valueStyle": "number" }
                  ],
                  "styles": { "card": "card", "title": "cardTitle", "label": "label", "value": "value" }
                }
              ]
            }
          },
          "sheets": [ { "name": "Currency Cards", "template": "currencyCards" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("currencies", CreateCurrencyRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_28_Template_exports_repeater_cards), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("USD", GetCellText(bytes, "Currency Cards", "A4"));
        Assert.Equal("GBP", GetCellText(bytes, "Currency Cards", "F4"));
        Assert.Equal("EUR", GetCellText(bytes, "Currency Cards", "A10"));
        Assert.Equal("Last High", GetCellText(bytes, "Currency Cards", "A6"));
        Assert.True(CountStyledCells(bytes, "Currency Cards") > 25);
    }

    [Fact]
    public void _29_Template_exports_chart_dashboard()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "title": { "font": { "size": 16, "bold": true, "color": "#385723" } },
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#385723" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#A9D18E" } }
          },
          "sheetTemplates": {
            "dataSheet": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Template Chart Data", "style": "title" },
                {
                  "type": "table",
                  "source": "monthly",
                  "startCell": "A3",
                  "columns": [
                    { "source": "Month", "caption": "Month" },
                    { "source": "Income", "caption": "Income" }
                  ],
                  "styles": { "header": "header", "data": "data" }
                }
              ]
            },
            "dashboard": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Template Chart Dashboard", "style": "title" },
                {
                  "type": "chart",
                  "title": "Income By Month",
                  "chartType": "area",
                  "dataSource": { "sheet": "Template Data", "categoryRange": "$A$4:$A$9", "valueRange": "$B$4:$B$9" },
                  "position": { "topLeft": "A3", "bottomRight": "J20" },
                  "style": 12,
                  "seriesColor": "#70AD47"
                }
              ]
            }
          },
          "sheets": [
            { "name": "Template Data", "template": "dataSheet" },
            { "name": "Template Dashboard", "template": "dashboard" }
          ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("monthly", CreateMonthlyRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_29_Template_exports_chart_dashboard), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountChartParts(bytes));
        Assert.Equal("Template Chart Dashboard", GetCellText(bytes, "Template Dashboard", "A1"));
    }

    [Fact]
    public void _30_Template_table_can_infer_columns_from_datasource()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#B4C6E7" } },
            "currency": { "numberFormat": "currency", "border": { "style": "thin", "color": "#70AD47" } }
          },
          "sheetTemplates": {
            "genericTable": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Generic Source Table", "style": "title" },
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "A3",
                  "headerAliases": {
                    "Income": "Income Total",
                    "Costs": "Cost Total"
                  },
                  "columnStyles": {
                    "Income": "currency",
                    "Costs": "currency"
                  },
                  "styles": { "header": "header", "data": "data" }
                }
              ]
            }
          },
          "sheets": [ { "name": "Generic Table", "template": "genericTable" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_30_Template_table_can_infer_columns_from_datasource), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("Region", GetCellText(bytes, "Generic Table", "A3"));
        Assert.Equal("Income Total", GetCellText(bytes, "Generic Table", "B3"));
        Assert.Equal("Cost Total", GetCellText(bytes, "Generic Table", "C3"));
        Assert.Equal("London", GetCellText(bytes, "Generic Table", "A4"));
        Assert.True(CountStyledCells(bytes, "Generic Table") > 10);
    }

    [Fact]
    public void _31_Template_repeater_can_infer_fields_from_datasource()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
            "card": { "fill": { "color": "#F8FBFD" }, "border": { "style": "thin", "color": "#B7C9D6" } },
            "cardTitle": { "font": { "size": 13, "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#1F4E78" } },
            "label": { "font": { "bold": true, "color": "#385723" }, "fill": { "color": "#E2F0D9" }, "border": { "style": "thin", "color": "#A9D18E" } },
            "value": { "border": { "style": "thin", "color": "#D9E2F3" } }
          },
          "sheetTemplates": {
            "genericRepeater": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Generic Repeater", "style": "title" },
                {
                  "type": "repeater",
                  "source": "currencies",
                  "startCell": "A4",
                  "itemsPerRow": 2,
                  "itemSize": { "columns": 4, "rows": 5 },
                  "gap": { "columns": 1, "rows": 1 },
                  "titleField": "Code",
                  "styles": { "card": "card", "title": "cardTitle", "label": "label", "value": "value" }
                }
              ]
            }
          },
          "sheets": [ { "name": "Generic Cards", "template": "genericRepeater" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("currencies", CreateCurrencyRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_31_Template_repeater_can_infer_fields_from_datasource), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("USD", GetCellText(bytes, "Generic Cards", "A4"));
        Assert.Equal("GBP", GetCellText(bytes, "Generic Cards", "F4"));
        Assert.Equal("Country", GetCellText(bytes, "Generic Cards", "A5"));
        Assert.Equal("LastHighRate", GetCellText(bytes, "Generic Cards", "A6"));
        Assert.Equal("EUR", GetCellText(bytes, "Generic Cards", "A10"));
        Assert.True(CountStyledCells(bytes, "Generic Cards") > 25);
    }

    [Fact]
    public void _32_Template_validation_reports_missing_bound_source()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "sheetTemplates": {
            "missingSourceSheet": {
              "regions": [
                { "type": "table", "source": "missing", "startCell": "A1" }
              ]
            }
          },
          "sheets": [ { "name": "Missing", "template": "missingSourceSheet" } ]
        }
        """;

        var result = ExcelTemplateWriter.FromJson(template).Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Message.Contains("missing", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void _33_Template_repeater_formula_expression_compiles_to_cell_formula()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "card": { "border": { "style": "thin", "color": "#B7C9D6" } },
            "title": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" } },
            "label": { "font": { "bold": true }, "fill": { "color": "#E2F0D9" } },
            "number": { "numberFormat": "number", "border": { "style": "thin", "color": "#5B9BD5" } }
          },
          "sheetTemplates": {
            "formulaCards": {
              "regions": [
                {
                  "type": "repeater",
                  "source": "currencies",
                  "startCell": "A4",
                  "itemsPerRow": 1,
                  "itemSize": { "columns": 4, "rows": 6 },
                  "titleField": "Code",
                  "fields": [
                    { "source": "Country", "caption": "Country" },
                    { "source": "LastHighRate", "caption": "Last High", "valueStyle": "number" },
                    { "source": "LastLowRate", "caption": "Last Low", "valueStyle": "number" },
                    { "caption": "Spread", "formula": "VALUE(LastHighRate)-VALUE(LastLowRate)", "valueStyle": "number" }
                  ],
                  "styles": { "card": "card", "title": "title", "label": "label", "value": "number" }
                }
              ]
            }
          },
          "sheets": [ { "name": "Formula Cards", "template": "formulaCards" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("currencies", CreateCurrencyRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_33_Template_repeater_formula_expression_compiles_to_cell_formula), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("B6-B7", GetCellFormula(bytes, "Formula Cards", "B8"));
    }

    [Fact]
    public void _34_Template_exports_pivot_table_region()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#B4C6E7" } }
          },
          "sheetTemplates": {
            "dataSheet": {
              "regions": [
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "A3",
                  "styles": { "header": "header", "data": "data" }
                }
              ]
            },
            "pivotSheet": {
              "regions": [
                {
                  "type": "pivotTable",
                  "name": "TemplateSalesPivot",
                  "source": { "sheet": "Sales Data", "range": "$A$3:$E$11" },
                  "startCell": "A3",
                  "rows": [ "Region" ],
                  "columns": [ "Quarter" ],
                  "values": [ { "field": "Amount", "aggregate": "sum", "caption": "Total Sales" } ],
                  "grandTotals": { "rows": true, "columns": true },
                  "style": "PivotStyleMedium9"
                }
              ]
            }
          },
          "sheets": [
            { "name": "Sales Data", "template": "dataSheet" },
            { "name": "Pivot", "template": "pivotSheet" }
          ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreatePivotSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_34_Template_exports_pivot_table_region), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(1, CountPivotTableParts(bytes, "Pivot"));
        Assert.Equal(1, CountPivotCacheDefinitionParts(bytes));
    }

    [Fact]
    public void _35_Template_chart_region_supports_legend_options()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#385723" }, "border": { "style": "thin", "color": "#000000" } }
          },
          "sheetTemplates": {
            "dataSheet": {
              "regions": [
                {
                  "type": "table",
                  "source": "monthly",
                  "startCell": "A3",
                  "styles": { "header": "header" }
                }
              ]
            },
            "dashboard": {
              "regions": [
                {
                  "type": "chart",
                  "title": "Income By Month",
                  "chartType": "line",
                  "dataSource": { "sheet": "Legend Data", "categoryRange": "$A$4:$A$9", "valueRange": "$B$4:$B$9" },
                  "position": { "topLeft": "A3", "bottomRight": "J20" },
                  "legendPosition": "bottom",
                  "showLegend": true,
                  "seriesColor": "#4472C4"
                }
              ]
            }
          },
          "sheets": [
            { "name": "Legend Data", "template": "dataSheet" },
            { "name": "Legend Dashboard", "template": "dashboard" }
          ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("monthly", CreateMonthlyRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_35_Template_chart_region_supports_legend_options), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal(C.LegendPositionValues.Bottom, GetFirstChartLegendPosition(bytes));
    }

    [Fact]
    public void _36_Template_validation_reports_overlapping_regions()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "sheetTemplates": {
            "overlap": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "First" },
                { "type": "cell", "cell": "A1", "value": "Second" }
              ]
            }
          },
          "sheets": [ { "name": "Overlap", "template": "overlap" } ]
        }
        """;

        var result = ExcelTemplateWriter.FromJson(template).Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Message.Contains("overlaps", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void _37_Template_summary_region_calculates_aggregates()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styles": {
            "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
            "label": { "font": { "bold": true }, "fill": { "color": "#DDEBF7" }, "border": { "style": "thin", "color": "#5B9BD5" } },
            "value": { "numberFormat": "currency", "border": { "style": "thin", "color": "#5B9BD5" } }
          },
          "sheetTemplates": {
            "summarySheet": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Sales Summary", "style": "title" },
                {
                  "type": "summary",
                  "source": "sales",
                  "startCell": "E3",
                  "labelStyle": "label",
                  "valueStyle": "value",
                  "fields": [
                    { "caption": "Rows", "aggregate": "count", "source": "Income" },
                    { "caption": "Total Income", "aggregate": "sum", "source": "Income" },
                    { "caption": "Average Costs", "aggregate": "avg", "source": "Costs" }
                  ]
                }
              ]
            }
          },
          "sheets": [ { "name": "Summary", "template": "summarySheet" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_37_Template_summary_region_calculates_aggregates), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("Rows", GetCellText(bytes, "Summary", "E3"));
        Assert.Equal("4", GetCellText(bytes, "Summary", "F3"));
        Assert.Equal("1090000", GetCellText(bytes, "Summary", "F4"));
        Assert.Equal("161250", GetCellText(bytes, "Summary", "F5"));
        Assert.True(CountStyledCells(bytes, "Summary") >= 6);
    }

    [Fact]
    public void _38_Template_validation_checks_declared_datasource_shape()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "dataSources": [
            {
              "name": "sales",
              "required": true,
              "columns": [
                { "name": "Region" },
                { "name": "Income" },
                { "name": "Margin" }
              ]
            }
          ],
          "sheetTemplates": {
            "salesSheet": {
              "regions": [
                { "type": "table", "source": "sales", "startCell": "A1" }
              ]
            }
          },
          "sheets": [ { "name": "Sales", "template": "salesSheet" } ]
        }
        """;

        var result = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Message.Contains("Margin", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void _39_Template_parameters_bind_into_cells_formulas_and_chart_titles()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "sheetTemplates": {
            "dataSheet": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Report date {{parameters.reportDate}}" },
                { "type": "formulaCell", "cell": "B1", "formula": "SUM(B4:B7)+{{parameters.adjustment}}" },
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "A3",
                  "columns": [
                    { "source": "Region", "caption": "Region" },
                    { "source": "Income", "caption": "Income" }
                  ]
                },
                {
                  "type": "chart",
                  "title": "Income {{parameters.reportDate}}",
                  "chartType": "bar",
                  "dataSource": { "sheet": "Parameterized", "categoryRange": "$A$4:$A$7", "valueRange": "$B$4:$B$7" },
                  "position": { "topLeft": "D3", "bottomRight": "L18" }
                }
              ]
            }
          },
          "sheets": [ { "name": "Parameterized", "template": "dataSheet" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .WithParameter("reportDate", "2026-09-22")
            .WithParameter("adjustment", 10)
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_39_Template_parameters_bind_into_cells_formulas_and_chart_titles), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("Report date 2026-09-22", GetCellText(bytes, "Parameterized", "A1"));
        Assert.Equal("SUM(B4:B7)+10", GetCellFormula(bytes, "Parameterized", "B1"));
        Assert.Equal(1, CountChartParts(bytes));
    }

    [Fact]
    public void _40_Template_exports_from_stored_template_provider()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "sheetTemplates": {
            "storedSales": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Stored Template Sales" },
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "A3",
                  "columns": [
                    { "source": "Region", "caption": "Region" },
                    { "source": "Income", "caption": "Income" },
                    { "source": "Costs", "caption": "Costs" }
                  ]
                }
              ]
            }
          },
          "sheets": [ { "name": "Stored", "template": "storedSales" } ]
        }
        """;

        var provider = new InMemoryExcelTemplateProvider()
            .Add("stored-sales", template);

        var bytes = ExcelTemplateWriter.FromTemplateId(provider, "stored-sales")
            .WithData("sales", CreateSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_40_Template_exports_from_stored_template_provider), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("Stored Template Sales", GetCellText(bytes, "Stored", "A1"));
        Assert.Equal("London", GetCellText(bytes, "Stored", "A4"));
    }

    [Fact]
    public void _41_Template_exports_layout_regions_and_data_validation()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styleCatalog": {
            "title": { "font": { "size": 16, "bold": true, "color": "#FFFFFF", "name": "Aptos Display" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thick", "color": "#000000" } },
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#385723" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#A9D18E" } }
          },
          "sheetTemplates": {
            "layoutSheet": {
              "regions": [
                { "type": "cell", "cell": "A1", "value": "Layout Template Export", "style": "title" },
                { "type": "mergeCells", "range": "A1:D1" },
                { "type": "rangeStyle", "range": "A1:D1", "style": "title" },
                { "type": "comment", "cell": "A1", "text": "Generated from a template comment.", "author": "Template" },
                { "type": "column", "column": "B", "width": 18 },
                { "type": "row", "row": 1, "height": 24 },
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "A3",
                  "columns": [
                    { "source": "Region", "caption": "Region" },
                    { "source": "Income", "caption": "Income" },
                    { "source": "Costs", "caption": "Costs" }
                  ],
                  "styles": { "header": "header", "data": "data" }
                },
                {
                  "type": "dataValidation",
                  "range": "E4:E20",
                  "validationType": "list",
                  "values": [ "Open", "Closed", "Review" ],
                  "inputMessage": { "title": "Status", "message": "Choose a status." },
                  "errorMessage": { "title": "Invalid status", "message": "Use the dropdown list." }
                }
              ]
            }
          },
          "sheets": [ { "name": "Layout", "template": "layoutSheet" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_41_Template_exports_layout_regions_and_data_validation), bytes);

        Xlsx.AssertValid(bytes);
        Assert.Equal("Layout Template Export", GetCellText(bytes, "Layout", "A1"));
        Assert.Equal(1, CountMergeRanges(bytes, "Layout"));
        Assert.Equal(1, CountComments(bytes, "Layout"));
        Assert.Equal(1, CountDataValidations(bytes, "Layout"));
        Assert.True(CountStyledCells(bytes, "Layout") > 12);
    }

    [Fact]
    public void _42_Template_imports_excel_with_typed_columns_and_derived_values()
    {
        var sourceBytes = ExcelWriter.Create()
            .AddSheet("Import Source", sheet => sheet
                .AddTable(new[]
                {
                    new { ID = 1, Name = "Alice", Amount = 12.50m },
                    new { ID = 2, Name = "Bob", Amount = 30.00m }
                })
                .StartAt("A1"))
            .ToBytes();

        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelImport",
          "input": {
            "sheet": "Import Source",
            "startCell": "A1",
            "hasHeaderRow": true
          },
          "columns": {
            "includeHeaders": [ "ID", "Name", "Amount" ]
          },
          "output": {
            "columns": [
              { "source": "ID", "name": "Identifier", "type": "int", "required": true, "nullable": false },
              { "source": "Name", "name": "Name", "type": "string", "required": true, "nullable": false },
              { "source": "Amount", "name": "Amount", "type": "decimal", "required": true, "nullable": false }
            ],
            "derivedColumns": [
              { "name": "DisplayName", "concat": [ "Identifier", " - ", "Name" ] }
            ]
          }
        }
        """;

        var table = ExcelTemplateReader.FromJson(template)
            .FromBytes(sourceBytes)
            .ReadDataTable("ImportedExcel");

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(typeof(int), table.Columns["Identifier"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Equal(1, table.Rows[0]["Identifier"]);
        Assert.Equal(12.50m, table.Rows[0]["Amount"]);
        Assert.Equal("1 - Alice", table.Rows[0]["DisplayName"]);
    }

    [Fact]
    public void _43_Template_preview_reports_below_previous_region_layout()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "sheetTemplates": {
            "stackedTables": {
              "regions": [
                { "type": "table", "source": "sales", "startCell": "A1" },
                { "type": "table", "source": "sales", "placement": { "mode": "belowPrevious", "column": "A", "gapRows": 1 } }
              ]
            }
          },
          "sheets": [ { "name": "Preview", "template": "stackedTables" } ]
        }
        """;

        var preview = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .Preview();

        Assert.Single(preview.Sheets);
        Assert.Equal("A1", preview.Sheets[0].Regions[0].StartCell);
        Assert.Equal("C5", preview.Sheets[0].Regions[0].EndCell);
        Assert.Equal("A7", preview.Sheets[0].Regions[1].StartCell);
        Assert.Equal("C11", preview.Sheets[0].Regions[1].EndCell);
    }

    [Fact]
    public void _44_Template_exports_freeze_panes_auto_filter_and_stacked_tables()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "styleCatalog": {
            "title": { "font": { "size": 16, "bold": true, "color": "#1F4E78" } },
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#B4C6E7" } }
          },
          "sheetTemplates": {
            "filtered": {
              "regions": [
                { "type": "cell", "cell": "B1", "value": "Template Filtered Export", "style": "title" },
                {
                  "type": "table",
                  "source": "sales",
                  "startCell": "B3",
                  "styles": { "header": "header", "data": "data" }
                },
                {
                  "type": "table",
                  "source": "sales",
                  "placement": { "mode": "belowPrevious", "column": "B", "gapRows": 2 },
                  "styles": { "header": "header", "data": "data" }
                },
                { "type": "freezePane", "rows": 3, "columns": 1 },
                { "type": "autoFilter", "range": "B3:D7" }
              ]
            }
          },
          "sheets": [ { "name": "Filtered Template", "template": "filtered" } ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("sales", CreateSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_44_Template_exports_freeze_panes_auto_filter_and_stacked_tables), bytes);

        var pane = GetFreezePane(bytes, "Filtered Template");

        Xlsx.AssertValid(bytes);
        Assert.Equal("Template Filtered Export", GetCellText(bytes, "Filtered Template", "B1"));
        Assert.Equal("Region", GetCellText(bytes, "Filtered Template", "B3"));
        Assert.Equal("Region", GetCellText(bytes, "Filtered Template", "B10"));
        Assert.Equal(3D, pane.VerticalSplit!.Value);
        Assert.Equal(1D, pane.HorizontalSplit!.Value);
        Assert.Equal("B4", pane.TopLeftCell!.Value);
        Assert.Equal("B3:D7", GetAutoFilterReference(bytes, "Filtered Template"));
        Assert.True(CountStyledCells(bytes, "Filtered Template") > 20);
    }

    [Fact]
    public void _47_Template_exports_protection_native_table_chart_labels_and_pivot_filters()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "excelWorkbook",
          "workbookProtection": { "lockStructure": true, "password": "workbook" },
          "styleCatalog": {
            "header": { "font": { "bold": true, "color": "#FFFFFF" }, "fill": { "color": "#1F4E78" }, "border": { "style": "thin", "color": "#000000" } },
            "data": { "border": { "style": "thin", "color": "#B4C6E7" } }
          },
          "sheetTemplates": {
            "data": {
              "regions": [
                {
                  "type": "table",
                  "source": "pivotSales",
                  "startCell": "A1",
                  "styles": { "header": "header", "data": "data" },
                  "nativeTable": { "name": "TemplatePivotSource", "style": "TableStyleMedium7" }
                },
                { "type": "sheetProtection", "password": "sheet", "allowAutoFilter": true }
              ]
            },
            "dashboard": {
              "regions": [
                {
                  "type": "chart",
                  "title": "Template Sales",
                  "chartType": "column",
                  "dataSource": { "sheet": "Template Data", "categoryRange": "$A$2:$A$9", "valueRange": "$E$2:$E$9" },
                  "position": { "topLeft": "A1", "bottomRight": "H16" },
                  "showDataLabels": true,
                  "axisTitles": { "category": "Region", "value": "Amount" }
                },
                {
                  "type": "pivotTable",
                  "name": "TemplateFilteredPivot",
                  "source": { "sheet": "Template Data", "range": "A1:E9" },
                  "startCell": "J3",
                  "filters": [ "Channel" ],
                  "rows": [ "Region" ],
                  "columns": [ "Quarter" ],
                  "values": [
                    { "field": "Amount", "caption": "Sales", "aggregate": "sum" },
                    { "field": "Amount", "caption": "Average Sales", "aggregate": "average" }
                  ]
                }
              ]
            }
          },
          "sheets": [
            { "name": "Template Data", "template": "data" },
            { "name": "Template Dashboard", "template": "dashboard" }
          ]
        }
        """;

        var bytes = ExcelTemplateWriter.FromJson(template)
            .WithData("pivotSales", CreatePivotSalesRows())
            .ToBytes();

        ExcelTestOutput.SaveWorkbook(nameof(_47_Template_exports_protection_native_table_chart_labels_and_pivot_filters), bytes);

        var definition = GetPivotTableDefinition(bytes, "Template Dashboard");

        Xlsx.AssertValid(bytes);
        Assert.True(HasWorkbookProtection(bytes));
        Assert.True(HasSheetProtection(bytes, "Template Data"));
        Assert.Equal(1, CountNativeTables(bytes, "Template Data"));
        Assert.Equal(1, CountChartDataLabels(bytes));
        Assert.Equal(2, CountChartAxisTitles(bytes));
        Assert.Equal(1, CountPivotTableParts(bytes, "Template Dashboard"));
        Assert.Equal(1U, definition.PageFields!.Count!.Value);

        // Two value fields (a sum and an average of the same source field) both reach the pivot.
        Assert.Equal(2U, definition.DataFields!.Count!.Value);
        Assert.Equal(
            ["Sales", "Average Sales"],
            definition.DataFields.Elements<DataField>().Select(field => field.Name!.Value));
    }

    private static IReadOnlyList<SalesTemplateRow> CreateSalesRows()
    {
        return
        [
            new("London", 210000m, 125000m),
            new("United States", 325000m, 190000m),
            new("Europe", 260000m, 155000m),
            new("Asia", 295000m, 175000m)
        ];
    }

    private static IReadOnlyList<CurrencyTemplateRow> CreateCurrencyRows()
    {
        return
        [
            new("USD", "United States", 1.2875m, 1.2420m),
            new("GBP", "United Kingdom", 1.0000m, 1.0000m),
            new("EUR", "Eurozone", 1.1764m, 1.1325m)
        ];
    }

    private static IReadOnlyList<MonthlyTemplateRow> CreateMonthlyRows()
    {
        return
        [
            new("Jan", 120000m),
            new("Feb", 135000m),
            new("Mar", 128000m),
            new("Apr", 155000m),
            new("May", 162000m),
            new("Jun", 171000m)
        ];
    }

    private static IReadOnlyList<PivotTemplateSalesRow> CreatePivotSalesRows()
    {
        return
        [
            new("London", "Software", "Q1", "Direct", 125000m),
            new("London", "Services", "Q1", "Partner", 68000m),
            new("London", "Software", "Q2", "Direct", 148000m),
            new("United States", "Software", "Q1", "Direct", 220000m),
            new("United States", "Services", "Q2", "Partner", 99000m),
            new("Europe", "Software", "Q1", "Partner", 142000m),
            new("Europe", "Services", "Q2", "Direct", 76000m),
            new("Asia", "Software", "Q2", "Direct", 171000m)
        ];
    }

    private sealed record SalesTemplateRow(string Region, decimal Income, decimal Costs);

    private sealed record CurrencyTemplateRow(string Code, string Country, decimal LastHighRate, decimal LastLowRate);

    private sealed record MonthlyTemplateRow(string Month, decimal Income);

    private sealed record PivotTemplateSalesRow(string Region, string Product, string Quarter, string Channel, decimal Amount);
}
