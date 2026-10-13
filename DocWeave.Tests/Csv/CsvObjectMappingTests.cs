using DocWeave;
using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

public sealed class CsvObjectMappingTests
{
    // ---- Positional records and constructors ----

    [Fact]
    public void ReadObjects_builds_a_positional_record_through_its_constructor_by_header_name()
    {
        var result = CsvReader.FromText("""
            Invoice No,Net Amount,Date
            INV-1,10.5,2026-01-07
            INV-2,20,2026-02-01
            """)
            .ReadObjects<InvoiceRecord>();

        Assert.False(result.HasErrors);
        Assert.Equal(new InvoiceRecord("INV-1", 10.5m, new DateOnly(2026, 1, 7)), result.Items[0]);
        Assert.Equal(new InvoiceRecord("INV-2", 20m, new DateOnly(2026, 2, 1)), result.Items[1]);
    }

    [Fact]
    public void ReadObjects_accepts_an_attribute_on_a_record_parameter_or_on_its_property()
    {
        var reps = CsvReader.FromText("Sales Rep,Age\nAlice,36").ReadObjects<RepRecord>().Items;
        var targets = CsvReader.FromText("Full Name\nBob").ReadObjects<PropertyTargetRecord>().Items;

        Assert.Equal(new RepRecord("Alice", 36), Assert.Single(reps));
        Assert.Equal("Bob", Assert.Single(targets).Name);
    }

    [Fact]
    public void ReadObjects_uses_a_parameter_default_when_the_column_is_missing_or_the_cell_is_blank()
    {
        var missing = CsvReader.FromText("Sales Rep\nAlice").ReadObjects<RepRecord>();
        var blank = CsvReader.FromText("Sales Rep,Age\nBob,").ReadObjects<RepRecord>();

        Assert.Equal(30, Assert.Single(missing.Items).Age);
        Assert.Equal(30, Assert.Single(blank.Items).Age);
    }

    [Fact]
    public void ReadObjects_reports_a_blank_or_missing_required_record_parameter()
    {
        var blank = CsvReader.FromText("Sales Rep,Age\n,5").ReadObjects<RepRecord>();
        var missing = CsvReader.FromText("Age\n5").ReadObjects<RepRecord>();

        Assert.Empty(blank.Items);
        Assert.Equal("Sales Rep", Assert.Single(blank.Errors).Column);
        Assert.Empty(missing.Items);
        Assert.Equal(0, Assert.Single(missing.Errors).RowNumber);
    }

    [Fact]
    public void ReadObjects_needs_a_value_for_a_value_type_parameter_without_a_default()
    {
        var result = CsvReader.FromText("""
            Count,Level
            ,Low
            7,Low
            """)
            .ReadObjects<CountedRecord>();

        Assert.Equal(new CountedRecord(7, Level.Low), Assert.Single(result.Items));
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.RowNumber);
        Assert.Equal("Count", error.Property);
        Assert.Equal("A value is required.", error.Message);
    }

    [Fact]
    public void ReadObjects_fills_both_the_constructor_and_the_extra_init_properties_of_a_record()
    {
        var item = Assert.Single(CsvReader.FromText("Name,Extra\nAlice,9").ReadObjects<MixedRecord>().Items);

        Assert.Equal("Alice", item.Name);
        Assert.Equal(9, item.Extra);
    }

    [Fact]
    public void ReadObjects_gives_an_ignored_constructor_parameter_its_default()
    {
        var item = Assert.Single(CsvReader.FromText("Name,Scratch\nAlice,from the file").ReadObjects<IgnoringRecord>().Items);

        Assert.Equal("kept", item.Scratch);
    }

    [Fact]
    public void ReadObjects_reports_why_a_constructor_rejected_a_row()
    {
        var result = CsvReader.FromText("Value\n3\n-1\n5").ReadObjects<PositiveValue>();

        Assert.Equal([3, 5], result.Items.Select(item => item.Value));
        var error = Assert.Single(result.Errors);
        Assert.Equal(3, error.RowNumber);
        Assert.Contains("must be positive", error.Message);
        Assert.Contains("PositiveValue", error.Message);
    }

    [Fact]
    public void ReadObjects_fills_a_struct_property_by_property()
    {
        var point = Assert.Single(CsvReader.FromText("X,Y\n1,2").ReadObjects<Point>().Items);

        Assert.Equal(1, point.X);
        Assert.Equal(2, point.Y);
    }

    [Fact]
    public void ReadObjects_fails_clearly_for_types_that_cannot_be_built()
    {
        var ambiguous = Assert.Throws<NotSupportedException>(() => CsvReader.FromText("A\n1").ReadObjects<AmbiguousConstructors>());
        var abstractType = Assert.Throws<NotSupportedException>(() => CsvReader.FromText("Sides\n1").ReadObjects<Shape>());

        Assert.Contains("more than one public constructor", ambiguous.Message);
        Assert.Contains("abstract", abstractType.Message);
    }

    // ---- Fluent mapping ----

    [Fact]
    public void ReadObjects_with_a_map_matches_by_header_and_by_position_and_leaves_ignored_properties_alone()
    {
        var item = Assert.Single(CsvReader.FromText("""
            Invoice No,Buyer,Amount
            INV-1,Northwind,12.5
            """)
            .ReadObjects<PlainInvoice>(map => map
                .Column(x => x.InvoiceNumber, "Invoice No")
                .Column(x => x.Customer, 2)
                .Column(x => x.Net, "Amount")
                .Ignore(x => x.Scratch)).Items);

        Assert.Equal("INV-1", item.InvoiceNumber);
        Assert.Equal("Northwind", item.Customer);
        Assert.Equal(12.5m, item.Net);
        Assert.Equal("untouched", item.Scratch);
    }

    [Fact]
    public void ReadObjects_with_a_map_replaces_the_attribute_on_the_same_property()
    {
        const string csv = "Wrong Header,Right Header\nfrom attribute,from map";

        var byAttribute = Assert.Single(CsvReader.FromText(csv).ReadObjects<PlainInvoice>().Items);
        var byMap = Assert.Single(CsvReader.FromText(csv).ReadObjects<PlainInvoice>(map => map.Column(x => x.Attributed, "Right Header")).Items);

        Assert.Equal("from attribute", byAttribute.Attributed);
        Assert.Equal("from map", byMap.Attributed);
    }

    [Fact]
    public void ReadObjects_with_a_map_can_bring_back_or_remove_a_property()
    {
        const string csv = "Hidden,Customer\nshown,Northwind";

        var ignoredByAttribute = Assert.Single(CsvReader.FromText(csv).ReadObjects<PlainInvoice>().Items);
        var brought = Assert.Single(CsvReader.FromText(csv).ReadObjects<PlainInvoice>(map => map.Column(x => x.Hidden, "Hidden")).Items);
        var removed = Assert.Single(CsvReader.FromText(csv).ReadObjects<PlainInvoice>(map => map.Ignore(x => x.Customer)).Items);

        Assert.Null(ignoredByAttribute.Hidden);
        Assert.Equal("shown", brought.Hidden);
        Assert.Null(removed.Customer);
    }

    [Fact]
    public void ReadObjects_with_a_map_reports_a_missing_column_and_a_blank_value_for_required_properties()
    {
        var missing = CsvReader.FromText("Other\nx").ReadObjects<PlainInvoice>(map => map.Required(x => x.Customer));
        var blank = CsvReader.FromText("Customer,Net\n,1").ReadObjects<PlainInvoice>(map => map.Column(x => x.Customer, "Customer", required: true));

        Assert.Empty(missing.Items);
        Assert.Equal("Customer", Assert.Single(missing.Errors).Property);
        Assert.Empty(blank.Items);
        Assert.Equal(2, Assert.Single(blank.Errors).RowNumber);
    }

    [Fact]
    public void ReadObjects_with_a_map_works_on_a_positional_record_by_naming_its_property()
    {
        var item = Assert.Single(CsvReader.FromText("""
            Client,Amount,When
            Alice,5,2026-03-04
            """)
            .ReadObjects<InvoiceRecord>(map => map
                .Column(x => x.InvoiceNo, "Client")
                .Column(x => x.NetAmount, "Amount")
                .Column(x => x.Date, "When")).Items);

        Assert.Equal(new InvoiceRecord("Alice", 5m, new DateOnly(2026, 3, 4)), item);
    }

    [Fact]
    public void A_map_position_must_be_one_or_more()
    {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            CsvReader.FromText("Customer\nx").ReadObjects<PlainInvoice>(map => map.Column(x => x.Net, 0)));

        Assert.Equal("ordinal", exception.ParamName);
    }

    [Fact]
    public void A_map_setting_must_be_a_simple_property_access()
    {
        var exception = Assert.Throws<ArgumentException>(() =>
            CsvReader.FromText("Customer\nx").ReadObjects<PlainInvoice>(map => map.Column(x => x.Customer!.Length, "Customer")));

        Assert.Contains("simple property access", exception.Message);
    }

    [Fact]
    public void EnumerateObjects_with_a_map_streams_objects_and_reports_problems_to_the_callback()
    {
        var errors = new List<DocWeaveImportError>();

        var items = CsvReader.FromText("""
            Client,Amount
            Alice,5
            Bob,x
            """)
            .EnumerateObjects<PlainInvoice>(map => map.Column(x => x.Customer, "Client").Column(x => x.Net, "Amount"), errors.Add)
            .ToList();

        Assert.Equal("Alice", Assert.Single(items).Customer);
        Assert.Equal(3, Assert.Single(errors).RowNumber);
    }

    private enum Level
    {
        Low,
        High
    }

    private sealed record InvoiceRecord(string InvoiceNo, decimal NetAmount, DateOnly Date);

    private sealed record RepRecord([DocWeaveColumn("Sales Rep", Required = true)] string Name, int Age = 30);

    private sealed record PropertyTargetRecord([property: DocWeaveColumn("Full Name")] string Name);

    private sealed record MixedRecord(string Name)
    {
        public int Extra { get; init; }
    }

    private sealed record CountedRecord(int Count, Level Level = Level.High);

    private sealed record IgnoringRecord(string Name, [DocWeaveIgnore] string? Scratch = "kept");

    private sealed class PositiveValue
    {
        public PositiveValue(int value)
        {
            if (value <= 0)
            {
                throw new ArgumentException("must be positive");
            }

            Value = value;
        }

        public int Value { get; }
    }

    private sealed class AmbiguousConstructors
    {
        public AmbiguousConstructors(int a)
        {
        }

        public AmbiguousConstructors(string a)
        {
        }
    }

    private abstract class Shape
    {
        public int Sides { get; set; }
    }

    private struct Point
    {
        public int X { get; set; }
        public int Y { get; set; }
    }

    private sealed class PlainInvoice
    {
        public string? InvoiceNumber { get; set; }
        public string? Customer { get; set; }
        public decimal Net { get; set; }
        public string? Scratch { get; set; } = "untouched";

        [DocWeaveColumn("Wrong Header")]
        public string? Attributed { get; set; }

        [DocWeaveIgnore]
        public string? Hidden { get; set; }
    }
}
