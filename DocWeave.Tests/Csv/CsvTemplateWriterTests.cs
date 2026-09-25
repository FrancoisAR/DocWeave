using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

public sealed class CsvTemplateWriterTests
{
    [Fact]
    public void _06_Csv_template_exports_selected_columns_with_aliases_and_delimiter()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "csv",
          "source": "currencies",
          "output": {
            "delimiter": ";",
            "includeHeaders": true
          },
          "columns": [
            { "source": "Code", "caption": "Currency" },
            { "source": "Country", "caption": "Country" },
            { "source": "LastHighRate", "caption": "High" }
          ]
        }
        """;

        var csv = CsvTemplateWriter.FromJson(template)
            .WithData("currencies", CreateCurrencyRows())
            .ToText();

        CsvTestOutput.SaveCsv(nameof(_06_Csv_template_exports_selected_columns_with_aliases_and_delimiter), csv);

        var lines = csv.Split(Environment.NewLine);
        Assert.Equal("Currency;Country;High", lines[0]);
        Assert.Equal("USD;United States;1.2875", lines[1]);
        Assert.Equal("GBP;United Kingdom;1.0000", lines[2]);
    }

    [Fact]
    public void _07_Csv_template_validation_reports_missing_source()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "csv",
          "source": "missing"
        }
        """;

        var result = CsvTemplateWriter.FromJson(template).Validate();

        Assert.False(result.IsValid);
        Assert.Contains(result.Errors, error => error.Message.Contains("missing", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void _08_Csv_template_import_selects_types_renames_and_derived_columns()
    {
        var template = """
        {
          "schemaVersion": "1.0",
          "kind": "csvImport",
          "input": {
            "delimiter": ",",
            "hasHeaderRow": true,
            "trimFields": true
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
              { "name": "DisplayName", "concat": [ "Identifier", " - ", "Name" ] },
              { "name": "Source", "constant": "CSV" }
            ]
          }
        }
        """;

        var table = CsvTemplateReader.FromJson(template)
            .FromText("ID,Name,Amount,Ignored\n1,Alice,12.50,x\n2,Bob,30.00,y")
            .ReadDataTable("ImportedCsv");

        Assert.Equal(2, table.Rows.Count);
        Assert.Equal(typeof(int), table.Columns["Identifier"]!.DataType);
        Assert.Equal(typeof(decimal), table.Columns["Amount"]!.DataType);
        Assert.Equal(1, table.Rows[0]["Identifier"]);
        Assert.Equal(12.50m, table.Rows[0]["Amount"]);
        Assert.Equal("1 - Alice", table.Rows[0]["DisplayName"]);
        Assert.Equal("CSV", table.Rows[1]["Source"]);
    }

    private static IReadOnlyList<CurrencyCsvRow> CreateCurrencyRows()
    {
        return
        [
            new("USD", "United States", 1.2875m, 1.2420m),
            new("GBP", "United Kingdom", 1.0000m, 1.0000m)
        ];
    }

    private sealed record CurrencyCsvRow(string Code, string Country, decimal LastHighRate, decimal LastLowRate);
}
