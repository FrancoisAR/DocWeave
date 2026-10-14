using System.Text.Json;
using DocWeave.Excel;
using DocWeave.WebApi;

internal static class BorderExamples
{
    public static void Run(string outputRoot)
    {
        var directory = Path.Combine(outputRoot, "borders");
        Directory.CreateDirectory(directory);
        var rows = Enumerable.Range(1, 20)
            .Select(i => new { Id = $"ITEM-{1000 + i * 37}", Cost = 12.50m * i, Revenue = 21.75m * i })
            .ToArray();

        // Data starts at row 2; totals follow the twentieth data row at row 22.
        // RangeStyle replaces a whole cell style, so the totals label is configured separately.
        var bytes = ExcelWriter.Create().AddSheet("Financials", sheet =>
        {
            sheet.AddTable(rows)
                .HeaderStyle(s => s.Bold().BorderBottom(ExcelBorderLineStyle.Medium))
                .DataStyle(s => s.Border(ExcelBorderLineStyle.Thin))
                .AlternateRowStyle(s => s.BackgroundColor("#EEF4FA").Border())
                .ColumnStyle("Cost", s => s.NumberFormat(ExcelNumberFormat.Currency).Border())
                .ColumnStyle("Revenue", s => s.NumberFormat(ExcelNumberFormat.Currency).Border())
                .AddTotals.Bottom("Totals", s => s.Bold()
                    .NumberFormat(ExcelNumberFormat.Currency)
                    .BorderTop().BorderBottom(ExcelBorderLineStyle.Double));
            sheet.RangeStyle("A2:A21", s => s.Bold().BackgroundColor("#E2F0D9").BorderLeft(ExcelBorderLineStyle.Medium));
            sheet.RangeStyle("A22:A22", s => s.Bold().BackgroundColor("#E2F0D9")
                .BorderLeft(ExcelBorderLineStyle.Medium).BorderTop().BorderBottom(ExcelBorderLineStyle.Double));
        }).ToBytes();
        File.WriteAllBytes(Path.Combine(directory, "direct-financials.xlsx"), bytes);

        // A small gallery makes border precedence, removal, colors and cell/range styling visible.
        bytes = ExcelWriter.Create().AddSheet("Combinations", sheet =>
        {
            sheet.AddCell("A1", "All thin, bottom double", s => s.Border().BorderBottom(ExcelBorderLineStyle.Double));
            sheet.AddCell("A3", "Four styles and colors", s => s
                .BorderTop(ExcelBorderLineStyle.Thin, "#FF0000")
                .BorderLeft(ExcelBorderLineStyle.Medium, "#008000")
                .BorderBottom(ExcelBorderLineStyle.Double, "#0000FF")
                .BorderRight(ExcelBorderLineStyle.Dashed, "#800080"));
            sheet.AddCell("A5", "Left removed", s => s.Border().BorderLeft(ExcelBorderLineStyle.None));
            sheet.AddCell("A7", "Border last resets sides", s => s.BorderBottom(ExcelBorderLineStyle.Double).Border());
            sheet.AddCell("A9", "Range styling");
            sheet.RangeStyle("A9:C9", s => s.BackgroundColor("#FFF2CC").BorderTop().BorderBottom(ExcelBorderLineStyle.Double));
            sheet.Column("A").Width(36);
        }).ToBytes();
        File.WriteAllBytes(Path.Combine(directory, "border-combinations.xlsx"), bytes);

        // JSON uses the same border model. Missing sides remain unbordered unless a global style is set.
        const string template = """
        {
          "kind": "excelWorkbook",
          "styles": {
            "header": { "font": { "bold": true }, "border": { "bottom": { "style": "medium" } } },
            "money": { "numberFormat": "currency", "border": { "style": "thin" } },
            "alternate": { "fill": { "color": "#EEF4FA" }, "border": { "style": "thin" } },
            "id": { "font": { "bold": true }, "fill": { "color": "#E2F0D9" }, "border": { "left": { "style": "medium" } } },
            "totals": { "font": { "bold": true }, "numberFormat": "currency", "border": { "top": { "style": "thin" }, "bottom": { "style": "double" } } },
            "totalsId": { "font": { "bold": true }, "fill": { "color": "#E2F0D9" }, "border": { "left": { "style": "medium" }, "top": { "style": "thin" }, "bottom": { "style": "double" } } }
          },
          "sheetTemplates": {
            "financials": { "regions": [
              { "type": "table", "source": "rows", "startCell": "A1",
                "styles": { "header": "header", "data": "money", "alternateRow": "alternate" },
                "columnStyles": { "Cost": "money", "Revenue": "money" },
                "totals": { "label": "Totals", "style": "totals" } },
              { "type": "rangeStyle", "range": "A2:A21", "style": "id" },
              { "type": "rangeStyle", "range": "A22:A22", "style": "totalsId" }
            ] }
          },
          "sheets": [ { "name": "Financials", "template": "financials" } ]
        }
        """;
        File.WriteAllText(Path.Combine(directory, "template.json"), template);
        File.WriteAllBytes(Path.Combine(directory, "template-financials.xlsx"),
            ExcelTemplateWriter.FromJson(template).WithData("rows", rows).ToBytes());

        // This is the DTO payload for an API host; endpoint routes are defined by the hosting application.
        var request = new DocWeaveExportRequest("xlsx", null, template,
            new Dictionary<string, JsonElement> { ["rows"] = JsonSerializer.SerializeToElement(rows) },
            Delivery: new DocWeaveExportDeliveryRequest("download", "financials.xlsx"));
        File.WriteAllText(Path.Combine(directory, "api-request.json"),
            JsonSerializer.Serialize(request, new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase }));
        Console.WriteLine("  borders (direct report, border gallery, template report and API payload)");
    }
}