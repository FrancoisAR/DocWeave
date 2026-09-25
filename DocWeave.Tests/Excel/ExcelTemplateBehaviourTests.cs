using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

/// <summary>
/// Pins how the template compiler behaves at its edges: errors, placement and overlap rules.
/// Written before restructuring the compiler so the restructuring can prove it changed nothing.
/// </summary>
public sealed class ExcelTemplateBehaviourTests
{
    // ---- errors ----

    [Fact]
    public void Unknown_region_type_is_reported_as_not_supported()
    {
        var result = Validate("""{ "type": "hologram" }""");

        Assert.False(result.IsValid);
        Assert.Equal("NotSupportedException", result.Errors[0].Code);
        Assert.Contains("'hologram'", result.Errors[0].Message);
    }

    [Fact]
    public void Region_type_names_are_case_sensitive()
    {
        var result = Validate("""{ "type": "Cell", "cell": "A1", "value": "x" }""");

        Assert.False(result.IsValid);
        Assert.Contains("'Cell'", result.Errors[0].Message);
    }

    [Fact]
    public void Region_without_a_type_is_reported()
    {
        var result = Validate("""{ "cell": "A1" }""");

        Assert.False(result.IsValid);
        Assert.Equal("Required template property 'type' was not found.", result.Errors[0].Message);
    }

    [Fact]
    public void Missing_style_is_reported()
    {
        var result = Validate("""{ "type": "cell", "cell": "A1", "value": "x", "style": "nope" }""");

        Assert.Equal("Template style 'nope' was not found.", result.Errors[0].Message);
    }

    [Fact]
    public void Unbound_source_is_reported()
    {
        var result = Validate("""{ "type": "table", "source": "ghost", "startCell": "A1" }""");

        Assert.Equal("Template source 'ghost' has not been bound.", result.Errors[0].Message);
    }

    [Fact]
    public void Unbound_parameter_is_reported_and_a_bound_one_is_substituted()
    {
        const string region = """{ "type": "cell", "cell": "A1", "value": "On {{parameters.when}}" }""";

        var unbound = Validate(region);
        var bytes = Builder(region).WithParameter("when", "Monday").ToBytes();
        var cell = ExcelReader.FromBytes(bytes).Sheet("S").HasHeaderRow(false).ReadDictionaries()[0]["A"];

        Assert.Equal("Template parameter 'when' has not been bound.", unbound.Errors[0].Message);
        Assert.Equal("On Monday", cell);
    }

    [Fact]
    public void Missing_sheet_template_is_reported()
    {
        const string template = """
            { "sheetTemplates": {}, "sheets": [ { "name": "S", "template": "nope" } ] }
            """;

        var result = ExcelTemplateWriter.FromJson(template).Validate();

        Assert.Equal("Sheet template 'nope' was not found.", result.Errors[0].Message);
    }

    [Fact]
    public void Unsupported_data_validation_type_is_reported()
    {
        var result = Validate("""{ "type": "dataValidation", "range": "A1:A5", "validationType": "weird" }""");

        Assert.Equal("Data validation type 'weird' is not supported.", result.Errors[0].Message);
    }

    [Fact]
    public void Unsupported_summary_aggregate_is_reported()
    {
        var result = Validate("""
            { "type": "summary", "source": "sales", "startCell": "A1",
              "fields": [ { "caption": "Middle", "aggregate": "median", "source": "Amount" } ] }
            """);

        Assert.Equal("Summary aggregate 'median' is not supported.", result.Errors[0].Message);
    }

    [Fact]
    public void Summary_over_a_missing_field_is_reported()
    {
        var result = Validate("""
            { "type": "summary", "source": "sales", "startCell": "A1",
              "fields": [ { "caption": "Total", "aggregate": "sum", "source": "Nope" } ] }
            """);

        Assert.Equal("Source field 'Nope' was not found.", result.Errors[0].Message);
    }

    [Fact]
    public void Table_column_that_is_not_in_the_source_is_reported()
    {
        var result = Validate("""
            { "type": "table", "source": "sales", "startCell": "A1", "columns": [ { "source": "Nope" } ] }
            """);

        Assert.Equal("Source field 'Nope' was not found.", result.Errors[0].Message);
    }

    [Fact]
    public void Repeater_formula_naming_an_unknown_field_fails_when_the_workbook_is_built()
    {
        var builder = Builder("""
            { "type": "repeater", "source": "sales", "startCell": "A1", "itemSize": { "columns": 3, "rows": 4 },
              "fields": [ { "source": "Amount" }, { "caption": "Bad", "formula": "VALUE(Nope)" } ] }
            """);

        var exception = Assert.Throws<InvalidOperationException>(() => builder.ToBytes());

        Assert.Equal("Formula field reference 'Nope' was not found in the repeater fields.", exception.Message);
    }

    // ---- data source declarations ----

    [Fact]
    public void Data_source_declared_twice_is_reported()
    {
        var result = ValidateWithDeclarations("""
            [ { "name": "sales" }, { "name": "sales" } ]
            """);

        Assert.Equal("Template source 'sales' is declared more than once.", result.Errors[0].Message);
    }

    [Fact]
    public void Required_declared_source_that_is_not_bound_is_reported_but_optional_is_not()
    {
        var required = ExcelTemplateWriter.FromJson(TemplateWithDeclarations("""[ { "name": "ghost" } ]""")).Validate();
        var optional = ExcelTemplateWriter.FromJson(TemplateWithDeclarations("""[ { "name": "ghost", "required": false } ]""")).Validate();

        Assert.Equal("Required template source 'ghost' has not been bound.", required.Errors[0].Message);
        Assert.True(optional.IsValid);
    }

    [Fact]
    public void Declared_non_nullable_column_rejects_a_null_value()
    {
        var builder = ExcelTemplateWriter.FromJson(TemplateWithDeclarations("""
            [ { "name": "sales", "columns": [ { "name": "Amount", "nullable": false } ] } ]
            """))
            .WithRows("sales", [new Dictionary<string, object?> { ["Region"] = "London", ["Amount"] = null }]);

        Assert.Equal("Source 'sales' field 'Amount' row 1 cannot be null.", builder.Validate().Errors[0].Message);
    }

    [Fact]
    public void Declared_column_type_rejects_a_value_of_another_type()
    {
        var builder = ExcelTemplateWriter.FromJson(TemplateWithDeclarations("""
            [ { "name": "sales", "columns": [ { "name": "Region", "type": "int" } ] } ]
            """))
            .WithRows("sales", [new Dictionary<string, object?> { ["Region"] = "London" }]);

        Assert.Equal("Source 'sales' field 'Region' row 1 must be int.", builder.Validate().Errors[0].Message);
    }

    [Fact]
    public void Declared_column_type_that_is_not_supported_is_reported()
    {
        var builder = ExcelTemplateWriter.FromJson(TemplateWithDeclarations("""
            [ { "name": "sales", "columns": [ { "name": "Region", "type": "guid" } ] } ]
            """))
            .WithRows("sales", [new Dictionary<string, object?> { ["Region"] = "London" }]);

        Assert.Equal("Template data type 'guid' is not supported.", builder.Validate().Errors[0].Message);
    }

    // ---- overlap ----

    [Fact]
    public void Regions_that_overlap_are_reported_with_the_start_of_the_second_one()
    {
        var result = Validate("""
            { "type": "table", "source": "sales", "startCell": "A1" },
            { "type": "cell", "cell": "B2", "value": "in the way" }
            """);

        Assert.Equal("Template region starting at B2 overlaps an earlier region.", result.Errors[0].Message);
    }

    [Fact]
    public void Regions_that_only_touch_do_not_overlap()
    {
        // The table fills A1:B3 (header + 2 rows). D1 and A5 are clear of it.
        var result = Validate("""
            { "type": "table", "source": "sales", "startCell": "A1" },
            { "type": "cell", "cell": "D1", "value": "beside" },
            { "type": "cell", "cell": "A5", "value": "below" }
            """);

        Assert.True(result.IsValid);
    }

    [Fact]
    public void Settings_regions_may_sit_on_top_of_content_without_overlapping()
    {
        // Comments, ranges styles, merges, column/row settings and the like do not take up cells of their own.
        var result = Validate("""
            { "type": "cell", "cell": "A1", "value": "x" },
            { "type": "comment", "cell": "A1", "text": "note" },
            { "type": "mergeCells", "range": "A1:C1" },
            { "type": "column", "column": "A", "width": 20 },
            { "type": "row", "row": 1, "height": 30 },
            { "type": "freezePane", "rows": 1 },
            { "type": "autoFilter", "range": "A1:B3" }
            """);

        Assert.True(result.IsValid);
    }

    // ---- placement ----

    [Fact]
    public void Below_previous_without_an_earlier_region_is_reported()
    {
        var result = Validate("""
            { "type": "table", "source": "sales", "startCell": "A1", "placement": { "mode": "belowPrevious" } }
            """);

        Assert.Equal("Template placement mode 'belowPrevious' requires an earlier placed region.", result.Errors[0].Message);
    }

    [Fact]
    public void Below_previous_starts_under_the_previous_region_in_its_first_column()
    {
        // The title sits in C2. The table has 2 rows + header, and lands one blank row below the title (C4).
        var builder = Builder("""
            { "type": "cell", "cell": "C2", "value": "Title" },
            { "type": "table", "source": "sales", "startCell": "Z99", "placement": { "mode": "belowPrevious" } }
            """);

        var preview = builder.Preview().Sheets[0].Regions;
        var rows = ExcelReader.FromBytes(builder.ToBytes()).Sheet("S").StartAt("C4").ReadDictionaries();

        Assert.Equal("C4", preview[1].StartCell);
        Assert.Equal("London", rows[0]["Region"]);
    }

    [Fact]
    public void Below_previous_honours_the_column_and_row_gap()
    {
        var builder = Builder("""
            { "type": "cell", "cell": "A1", "value": "Title" },
            { "type": "table", "source": "sales", "startCell": "Z99",
              "placement": { "mode": "belowPrevious", "column": "E", "gapRows": 3 } }
            """);

        var preview = builder.Preview().Sheets[0].Regions;

        Assert.Equal("E5", preview[1].StartCell);
    }

    [Fact]
    public void Below_previous_chains_from_the_region_before_it()
    {
        // Table A1:B3, then a table 1 gap below (A5:B7), then another 1 gap below that (A9).
        var builder = Builder("""
            { "type": "table", "source": "sales", "startCell": "A1" },
            { "type": "table", "source": "sales", "startCell": "Z1", "placement": { "mode": "belowPrevious" } },
            { "type": "table", "source": "sales", "startCell": "Z1", "placement": { "mode": "belowPrevious" } }
            """);

        var starts = builder.Preview().Sheets[0].Regions.Select(region => region.StartCell).ToList();

        Assert.Equal(["A1", "A5", "A9"], starts);
    }

    [Fact]
    public void Settings_regions_are_not_carried_into_the_preview()
    {
        var builder = Builder("""
            { "type": "cell", "cell": "A1", "value": "x" },
            { "type": "comment", "cell": "A1", "text": "note" },
            { "type": "freezePane", "rows": 1 }
            """);

        var regions = builder.Preview().Sheets[0].Regions;

        Assert.Single(regions);
        Assert.Equal("cell", regions[0].Type);
    }

    [Fact]
    public void Preview_reports_the_cells_each_region_occupies()
    {
        var builder = Builder("""
            { "type": "table", "source": "sales", "startCell": "B2" },
            { "type": "chart", "title": "T", "chartType": "column",
              "dataSource": { "sheet": "S", "categoryRange": "$B$3:$B$4", "valueRange": "$C$3:$C$4" },
              "position": { "topLeft": "F2", "bottomRight": "M12" } }
            """);

        var regions = builder.Preview().Sheets[0].Regions;

        Assert.Equal(("table", "B2", "C4"), (regions[0].Type, regions[0].StartCell, regions[0].EndCell));
        Assert.Equal(("chart", "F2", "M12"), (regions[1].Type, regions[1].StartCell, regions[1].EndCell));
    }

    // ---- helpers ----

    private static ExcelTemplateValidationResult Validate(string regions) => Builder(regions).Validate();

    private static ExcelTemplateExportBuilder Builder(string regions)
    {
        var template = $$"""
            {
              "schemaVersion": "1.0",
              "kind": "excelWorkbook",
              "styles": { "bold": { "font": { "bold": true } } },
              "sheetTemplates": { "t": { "regions": [ {{regions}} ] } },
              "sheets": [ { "name": "S", "template": "t" } ]
            }
            """;

        return ExcelTemplateWriter.FromJson(template).WithRows("sales", SalesRows());
    }

    private static ExcelTemplateValidationResult ValidateWithDeclarations(string declarations)
    {
        return ExcelTemplateWriter.FromJson(TemplateWithDeclarations(declarations))
            .WithRows("sales", SalesRows())
            .Validate();
    }

    private static string TemplateWithDeclarations(string declarations)
    {
        return $$"""
            {
              "schemaVersion": "1.0",
              "kind": "excelWorkbook",
              "dataSources": {{declarations}},
              "sheetTemplates": { "t": { "regions": [ { "type": "cell", "cell": "A1", "value": "x" } ] } },
              "sheets": [ { "name": "S", "template": "t" } ]
            }
            """;
    }

    private static IReadOnlyList<IReadOnlyDictionary<string, object?>> SalesRows()
    {
        return
        [
            new Dictionary<string, object?> { ["Region"] = "London", ["Amount"] = 100m },
            new Dictionary<string, object?> { ["Region"] = "Paris", ["Amount"] = 200m }
        ];
    }
}
