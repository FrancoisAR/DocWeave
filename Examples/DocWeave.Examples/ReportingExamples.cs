using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocWeave.Csv;
using DocWeave.Excel;
using DocWeave.WebApi;

internal static class ReportingExamples
{
    public static void Run(string outputRoot)
    {
        var directory = Path.Combine(outputRoot, "reporting");
        Directory.CreateDirectory(directory);
        var rows = new[]
        {
            new Sale("North", "Q1", "Web", 120m, 200m, "Open"),
            new Sale("North", "Q2", "Store", 160m, 140m, "Review"),
            new Sale("South", "Q1", "Store", 90m, 180m, "Closed"),
            new Sale("South", "Q2", "Web", 210m, 300m, "Open"),
            new Sale("West", "Q1", "Web", 170m, 150m, "Review"),
            new Sale("West", "Q2", "Store", 100m, 240m, "Closed")
        };

        // 1. Formula columns use worksheet coordinates, not source-object calculations.
        // Profit occupies G; the generated totals include its SUM formula.
        Save("01-formulas-totals.xlsx", ExcelWriter.Create().AddSheet("Profit", sheet =>
        {
            sheet.AddTable(rows).StartAt("A3")
                .HeaderStyle(s => s.Bold().BackgroundColor("#DDEBF7"))
                .ColumnStyle("Cost", s => s.NumberFormat(ExcelNumberFormat.Currency))
                .ColumnStyle("Revenue", s => s.NumberFormat(ExcelNumberFormat.Currency))
                .FormulaColumn("Profit", c => $"E{c.WorksheetRow}-D{c.WorksheetRow}")
                .ColumnStyle("Profit", s => s.NumberFormat(ExcelNumberFormat.Currency))
                .AddTotals.Bottom("Totals", s => s.Bold().NumberFormat(ExcelNumberFormat.Currency)
                    .BorderTop().BorderBottom(ExcelBorderLineStyle.Double));
            sheet.AddCell("A1", "Profit and loss", s => s.Bold().FontSize(16));
            // Excel calculates formulas when opened; DocWeave does not evaluate them.
            sheet.AddFormulaCell("I4", "SUM(G4:G9)", s => s.Bold().NumberFormat(ExcelNumberFormat.Currency));
            sheet.FreezePanes(rows: 3, columns: 1);
            sheet.AutoFilter("A3:G9");
        }).ToBytes());

        // 2. Conditional styles are fixed at export time. They do not update after editing in Excel.
        // Row rules, a single-column predicate and value rules are deliberately shown separately.
        Save("02-conditional-styles.xlsx", ExcelWriter.Create().AddSheet("Exceptions", sheet =>
        {
            sheet.AddTable(rows)
                .HeaderStyle(s => s.Bold())
                .AlternateRowStyle(s => s.BackgroundColor("#EEF4FA"))
                .RowStyleWhen(r => Convert.ToDecimal(r["Revenue"]) < Convert.ToDecimal(r["Cost"]),
                    s => s.BackgroundColor("#FFC7CE").FontColor("#9C0006"))
                .CellStyle("Revenue", r => Convert.ToDecimal(r["Revenue"]) >= 2m * Convert.ToDecimal(r["Cost"]),
                    s => s.Bold().BackgroundColor("#C6EFCE").NumberFormat(ExcelNumberFormat.Currency))
                .ConditionalStyle("Status", value => Equals(value, "Review"),
                    s => s.Bold().BackgroundColor("#FFEB9C"))
                .AddTotals.Bottom("Totals", s => s.Bold().BorderBottom(ExcelBorderLineStyle.Double));
        }).ToBytes());

        // 3. The source range includes headers and six records, but excludes totals.
        Save("03-pivot-quarterly.xlsx", ExcelWriter.Create()
            .AddSheet("Data", s => s.AddTable(rows).AsNativeTable("SalesSource"))
            .AddSheet("Pivot", s => s.AddPivotTable("QuarterlyRevenue", p => p
                .Source("Data", "A1:F7").StartAt("A3").Rows("Region").Columns("Quarter")
                .Values(v => v.Sum("Revenue", "Total Revenue"))
                .ShowGrandTotals().Style("PivotStyleMedium9")))
            .ToBytes());

        // 4. Report filters and multiple aggregates; value captions differ from source headers.
        Save("04-pivot-multiple-values.xlsx", ExcelWriter.Create()
            .AddSheet("Data", s => s.AddTable(rows))
            .AddSheet("Pivot", s => s.AddPivotTable("RevenueStatistics", p => p
                .Source("Data", "A1:F7").StartAt("A3").Rows("Region").Filters("Channel")
                .Values(v => v.Sum("Revenue", "Total Revenue").Average("Revenue", "Mean Revenue"))
                .ShowGrandTotals()))
            .ToBytes());

        // 5. Repeating cards show fields and a relative formula in each block.
        Save("05-repeater-formulas.xlsx", ExcelWriter.Create().AddSheet("Cards", s => s
            .AddRepeater(rows).StartAt("A1").ItemsPerRow(2).ItemSize(4, 6).Gap(1, 1)
            .TitleField("Region").Field("Quarter").Field("Cost").Field("Revenue")
            .FormulaField("Profit", c => $"INDIRECT(ADDRESS({c.WorksheetRow + 3},{c.WorksheetColumn + 1}))-INDIRECT(ADDRESS({c.WorksheetRow + 2},{c.WorksheetColumn + 1}))")
            .CardStyle(style => style.Border()).TitleStyle(style => style.Bold().BackgroundColor("#DDEBF7")))
            .ToBytes());

        // 6. CSV has no spreadsheet formulas/styles. Export computed profit as ordinary data.
        var csvRows = rows.Select(r => new { r.Region, r.Quarter, r.Cost, r.Revenue, Profit = r.Revenue - r.Cost }).ToList();
        csvRows.Add(new { Region = "Totals", Quarter = "", Cost = rows.Sum(r => r.Cost),
            Revenue = rows.Sum(r => r.Revenue), Profit = rows.Sum(r => r.Revenue - r.Cost) });
        CsvWriter.FromObjects(csvRows).DelimitedBy(";").HeaderAlias("Profit", "Net Profit")
            .SaveAs(Path.Combine(directory, "06-calculated-csv-totals.csv"));

        // 7. Template equivalent of the quarterly pivot, plus the exact DTO request payload.
        const string template = """
        {
          "schemaVersion": "1.0", "kind": "excelWorkbook",
          "styles": {
            "heading": { "font": { "bold": true }, "fill": { "color": "#DDEBF7" } },
            "totals": { "font": { "bold": true }, "numberFormat": "currency",
              "border": { "top": { "style": "thin" }, "bottom": { "style": "double" } } }
          },
          "sheetTemplates": {
            "data": { "regions": [
              { "type": "table", "source": "sales", "startCell": "A1",
                "styles": { "header": "heading" }, "totals": { "label": "Totals", "style": "totals" } }
            ] },
            "pivot": { "regions": [
              { "type": "cell", "cell": "A1", "value": "Quarterly revenue", "style": "heading" },
              { "type": "pivotTable", "name": "TemplateQuarterly",
                "source": { "sheet": "Data", "range": "A1:F7" }, "startCell": "A3",
                "rows": [ "Region" ], "columns": [ "Quarter" ],
                "values": [ { "field": "Revenue", "aggregate": "sum", "caption": "Total Revenue" } ],
                "grandTotals": { "rows": true, "columns": true }, "style": "PivotStyleMedium9" }
            ] }
          },
          "sheets": [ { "name": "Data", "template": "data" }, { "name": "Pivot", "template": "pivot" } ]
        }
        """;
        Save("07-template-pivot-totals.xlsx", ExcelTemplateWriter.FromJson(template).WithData("sales", rows).ToBytes());
        File.WriteAllText(Path.Combine(directory, "07-template.json"), template);
        var request = new DocWeaveExportRequest("xlsx", null, template,
            new Dictionary<string, JsonElement> { ["sales"] = JsonSerializer.SerializeToElement(rows) },
            Delivery: new DocWeaveExportDeliveryRequest("download", "quarterly-revenue.xlsx"));
        File.WriteAllText(Path.Combine(directory, "07-api-request.json"), JsonSerializer.Serialize(request,
            new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase }));
        Console.WriteLine("  reporting (7 formula, styling, pivot, repeater, CSV and template scenarios)");

        void Save(string name, byte[] bytes)
        {
            // Fail the cookbook run if any example emits an invalid Open XML document.
            using var stream = new MemoryStream(bytes);
            using var workbook = SpreadsheetDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(workbook).Take(5).ToArray();
            if (errors.Length > 0)
                throw new InvalidOperationException(name + ": " + string.Join("; ", errors.Select(e => e.Description)));
            File.WriteAllBytes(Path.Combine(directory, name), bytes);
        }
    }

    private sealed record Sale(string Region, string Quarter, string Channel, decimal Cost, decimal Revenue, string Status);
}