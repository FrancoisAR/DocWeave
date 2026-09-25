using DocWeave.Csv;
using DocWeave.Excel;

namespace DocWeave.Tests;

/// <summary>
/// Behaviour of the code the Excel and CSV paths share (tabular data, column selection, ranges, template JSON).
/// These go through the public API, so they hold no matter where the shared code lives.
/// </summary>
public sealed class SharedComponentTests
{
    // ---- tabular data: dictionary keys match ignoring case, in both formats ----

    [Fact]
    public void Excel_export_puts_dictionary_keys_that_differ_only_by_case_in_the_same_column()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Region"] = "London", ["Sales"] = 10 },
            new Dictionary<string, object?> { ["region"] = "Paris", ["SALES"] = 20 }
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Data", sheet => sheet.AddRows(rows).StartAt("A1"))
            .ToBytes();

        var imported = ExcelReader.FromBytes(bytes).Sheet("Data").StartAt("A1").ReadDictionaries();

        Assert.Equal(["Region", "Sales"], imported[0].Keys);
        Assert.Equal("Paris", imported[1]["Region"]);
        Assert.Equal("20", imported[1]["Sales"]);
    }

    [Fact]
    public void Csv_export_puts_dictionary_keys_that_differ_only_by_case_in_the_same_column()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Region"] = "London", ["Sales"] = 10 },
            new Dictionary<string, object?> { ["region"] = "Paris", ["SALES"] = 20 }
        };

        var csv = CsvWriter.FromRows(rows).NewLine("\n").ToText();

        Assert.Equal("Region,Sales\nLondon,10\nParis,20", csv);
    }

    [Fact]
    public void Csv_export_writes_empty_cells_for_columns_a_row_does_not_have()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["A"] = 1, ["B"] = 2 },
            new Dictionary<string, object?> { ["A"] = 3 }
        };

        var csv = CsvWriter.FromRows(rows).NewLine("\n").ToText();

        Assert.Equal("A,B\n1,2\n3,", csv);
    }

    // ---- column selection: each reader keeps its own ordering rule ----

    [Fact]
    public void Excel_import_returns_selected_columns_in_worksheet_order()
    {
        var bytes = CreateRegionProfitWorkbook();

        var rows = ExcelReader.FromBytes(bytes).Sheet("Data").StartAt("A1")
            .Columns(columns => columns.IncludeHeaders("Profit", "Region"))
            .ReadDictionaries();

        Assert.Equal(["Region", "Profit"], rows[0].Keys);
    }

    [Fact]
    public void Csv_import_returns_selected_columns_in_the_requested_order()
    {
        var rows = CsvReader.FromText("Region,Profit\nLondon,5")
            .Columns(columns => columns.IncludeHeaders("Profit", "Region"))
            .ReadDictionaries();

        Assert.Equal(["Profit", "Region"], rows[0].Keys);
    }

    [Fact]
    public void Excel_and_csv_import_name_the_source_when_a_header_is_missing()
    {
        var excel = Assert.Throws<InvalidOperationException>(() =>
            ExcelReader.FromBytes(CreateRegionProfitWorkbook()).Sheet("Data").StartAt("A1")
                .Columns(columns => columns.IncludeHeaders("Nope"))
                .ReadDictionaries());
        var csv = Assert.Throws<InvalidOperationException>(() =>
            CsvReader.FromText("Region,Profit\nLondon,5")
                .Columns(columns => columns.IncludeHeaders("Nope"))
                .ReadDictionaries());

        Assert.Equal("Excel header 'Nope' was not found.", excel.Message);
        Assert.Equal("CSV header 'Nope' was not found.", csv.Message);
    }

    [Fact]
    public void Column_letters_beyond_Z_select_the_right_columns_in_both_formats()
    {
        // 30 columns: C1..C30. Column 27 is AA and column 30 is AD.
        var header = string.Join(",", Enumerable.Range(1, 30).Select(i => $"C{i}"));
        var values = string.Join(",", Enumerable.Range(1, 30).Select(i => $"v{i}"));

        var rows = CsvReader.FromText($"{header}\n{values}")
            .Columns(columns => columns.IncludeLetters("AA", "AD"))
            .ReadDictionaries();

        Assert.Equal(["C27", "C30"], rows[0].Keys);
        Assert.Equal("v27", rows[0]["C27"]);
    }

    // ---- ranges ----

    [Fact]
    public void Range_style_accepts_absolute_references()
    {
        var bytes = ExcelWriter.Create()
            .AddSheet("Data", sheet => sheet
                .AddCell("A1", "x")
                .RangeStyle("$A$1:$B$2", style => style.Bold()))
            .ToBytes();

        Assert.NotEmpty(bytes);
    }

    [Fact]
    public void Range_style_rejects_a_range_without_two_corners()
    {
        Assert.Throws<ArgumentException>(() => ExcelWriter.Create()
            .AddSheet("Data", sheet => sheet.RangeStyle("A1", style => style.Bold())));
    }

    // ---- template JSON helpers ----

    [Fact]
    public void Excel_and_csv_templates_report_a_missing_required_property_the_same_way()
    {
        var excel = ExcelTemplateWriter.FromJson("""{ "sheetTemplates": {}, "sheets": [ { "template": "x" } ] }""").Validate();
        var csv = CsvTemplateWriter.FromJson("""{ "columns": [] }""").Validate();

        Assert.False(excel.IsValid);
        Assert.False(csv.IsValid);
        Assert.Equal("Required template property 'name' was not found.", excel.Errors[0].Message);
        Assert.Equal("Required template property 'source' was not found.", csv.Errors[0].Message);
    }

    [Fact]
    public void Csv_template_reports_a_missing_source_field_even_when_there_are_no_rows()
    {
        const string template = """
            {
              "source": "people",
              "columns": [ { "source": "Name" }, { "source": "Missing" } ]
            }
            """;

        var result = CsvTemplateWriter.FromJson(template)
            .WithData("people", Array.Empty<Person>())
            .Validate();

        Assert.False(result.IsValid);
        Assert.Contains("Missing", result.Errors[0].Message);
    }

    [Fact]
    public void Csv_template_projects_columns_in_the_order_the_template_lists_them()
    {
        const string template = """
            {
              "source": "people",
              "columns": [ { "source": "Age", "caption": "Years" }, { "source": "Name" } ]
            }
            """;

        var csv = CsvTemplateWriter.FromJson(template)
            .WithData("people", new[] { new Person("Ann", 31), new Person("Bo", 42) })
            .ToText();

        Assert.Equal($"Years,Name{Environment.NewLine}31,Ann{Environment.NewLine}42,Bo", csv);
    }

    private static byte[] CreateRegionProfitWorkbook()
    {
        var rows = new[] { new Dictionary<string, object?> { ["Region"] = "London", ["Profit"] = 5 } };
        return ExcelWriter.Create()
            .AddSheet("Data", sheet => sheet.AddRows(rows).StartAt("A1"))
            .ToBytes();
    }

    private sealed record Person(string Name, int Age);
}
