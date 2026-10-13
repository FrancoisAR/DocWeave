using DocWeave;
using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

public sealed class CsvObjectImportTests
{
    [Fact]
    public void ReadObjects_fills_objects_by_header_name()
    {
        var id = Guid.NewGuid();
        var result = CsvReader.FromText($"""
            name,age,balance,joined,active,id
            Alice,36,1234.50,2020-03-01,yes,{id}
            Bob,45,10,2019-12-31,false,{id}
            """)
            .ReadObjects<Person>();

        Assert.False(result.HasErrors);
        Assert.Equal(2, result.RowCount);
        Assert.Equal("Alice", result.Items[0].Name);
        Assert.Equal(36, result.Items[0].Age);
        Assert.Equal(1234.50m, result.Items[0].Balance);
        Assert.Equal(new DateOnly(2020, 3, 1), result.Items[0].Joined);
        Assert.True(result.Items[0].Active);
        Assert.Equal(id, result.Items[0].Id);
        Assert.False(result.Items[1].Active);
    }

    [Fact]
    public void ReadObjects_matches_headers_ignoring_case_spaces_underscores_and_hyphens()
    {
        var result = CsvReader.FromText("""
            NAME,Balance Due,joined-on
            Alice,10.5,2020-03-01
            """)
            .ReadObjects<LooseNames>();

        var item = Assert.Single(result.Items);
        Assert.Equal("Alice", item.Name);
        Assert.Equal(10.5m, item.BalanceDue);
        Assert.Equal(new DateOnly(2020, 3, 1), item.JoinedOn);
    }

    [Fact]
    public void ReadObjects_uses_the_invariant_culture_so_a_decimal_comma_is_an_error_not_a_wrong_value()
    {
        // Under a permissive thousands rule "1,5" would be 15. It must be reported instead.
        var result = CsvReader.FromText("""
            Name;Balance
            Alice;1,5
            Bob;1.5
            """)
            .DelimitedBy(";")
            .ReadObjects<Person>();

        Assert.Equal(["Bob"], result.Items.Select(person => person.Name));
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.RowNumber);
        Assert.Equal("Balance", error.Column);
        Assert.Equal("1,5", error.Value);
    }

    [Fact]
    public void ReadObjects_does_not_read_a_number_as_an_excel_date_in_a_csv_file()
    {
        // The Excel date-number rule applies to Excel cells only. In a CSV file "45123" is not a date.
        var result = CsvReader.FromText("""
            Name,Joined
            Alice,45123
            """)
            .ReadObjects<Person>();

        Assert.Empty(result.Items);
        Assert.Equal("Joined", Assert.Single(result.Errors).Column);
    }

    [Fact]
    public void ReadObjects_reports_the_record_number_of_the_file_for_each_problem()
    {
        // Record 1 is the header, record 3 is a blank line that is skipped, so the bad row is record 4.
        var result = CsvReader.FromText("Name,Age\nAlice,36\n\nBob,old\nCarol,54")
            .ReadObjects<Person>();

        Assert.Equal(["Alice", "Carol"], result.Items.Select(person => person.Name));
        Assert.Equal(4, Assert.Single(result.Errors).RowNumber);
    }

    [Fact]
    public void ReadObjects_lists_every_problem_in_a_row_and_counts_the_row_once()
    {
        var result = CsvReader.FromText("""
            Name,Age,Balance
            Alice,x,y
            """)
            .ReadObjects<Person>();

        Assert.Equal(2, result.Errors.Count);
        Assert.Equal(1, result.RowCount);
        Assert.Empty(result.Items);
    }

    [Fact]
    public void ReadObjects_honours_attribute_names_ordinals_and_required_columns()
    {
        var result = CsvReader.FromText("""
            Full Name,Other,Third
            Alice,x,third value
            ,y,z
            """)
            .ReadObjects<MappedPerson>();

        var item = Assert.Single(result.Items);
        Assert.Equal("Alice", item.Name);
        Assert.Equal("third value", item.Third);
        var error = Assert.Single(result.Errors);
        Assert.Equal(3, error.RowNumber);
        Assert.Equal("Full Name", error.Column);
    }

    [Fact]
    public void ReadObjects_reports_a_missing_required_column_and_reads_nothing()
    {
        var result = CsvReader.FromText("""
            Other
            x
            """)
            .ReadObjects<MappedPerson>();

        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(0, error.RowNumber);
        Assert.Equal("Name", error.Property);
    }

    [Fact]
    public void ReadObjects_counts_ordinals_after_column_selection()
    {
        // Only Full Name and Third are kept, so "Third" is now the second imported column, not the third.
        var result = CsvReader.FromText("""
            Full Name,Other,Third
            Alice,x,t
            """)
            .Columns(columns => columns.IncludeHeaders("Full Name", "Third"))
            .ReadObjects<MappedPerson>();

        var item = Assert.Single(result.Items);
        Assert.Null(item.Third);
        Assert.False(result.HasErrors);
    }

    [Fact]
    public void ReadObjects_treats_a_blank_as_the_default_for_optional_values_and_an_error_for_a_plain_number()
    {
        var result = CsvReader.FromText("""
            Name,Age,Nickname
            Alice,,
            """)
            .ReadObjects<NicknamedPerson>();

        // Age is an int, which cannot hold "no value". Nickname is optional and keeps the value the class gave it.
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal("Age", error.Column);
        Assert.Equal("A value is required.", error.Message);
    }

    [Fact]
    public void ReadObjects_keeps_the_initial_value_of_a_property_for_a_blank_cell()
    {
        var result = CsvReader.FromText("""
            Name,Nickname
            Alice,
            """)
            .ReadObjects<OptionalNickname>();

        Assert.Equal("none", Assert.Single(result.Items).Nickname);
    }

    [Fact]
    public void EnumerateObjects_streams_objects_and_reports_problems_to_the_callback()
    {
        var errors = new List<DocWeaveImportError>();
        var names = new List<string?>();

        foreach (var person in CsvReader.FromText("Name,Age\nAlice,36\nBob,old\nCarol,54").EnumerateObjects<Person>(errors.Add))
        {
            names.Add(person.Name);
        }

        Assert.Equal(["Alice", "Carol"], names);
        Assert.Equal(3, Assert.Single(errors).RowNumber);
    }

    [Fact]
    public void EnumerateObjects_leaves_a_caller_owned_stream_open_when_stopped_early()
    {
        var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("Name,Age\nAlice,1\nBob,2"));

        using (var enumerator = CsvReader.FromStream(stream).EnumerateObjects<Person>().GetEnumerator())
        {
            Assert.True(enumerator.MoveNext());
        }

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void ReadObjects_audit_holds_counts_and_never_values()
    {
        var audits = new List<DocWeaveAuditEvent>();

        CsvReader.FromText("Name,Age\nsecret-name,bad").Audit(audits.Add).ReadObjects<Person>();

        var audit = Assert.Single(audits);
        Assert.Equal(1, audit.Properties["RowCount"]);
        Assert.Equal(1, audit.Properties["ErrorCount"]);
        Assert.DoesNotContain(audit.Properties.Values, value => "secret-name".Equals(value) || "bad".Equals(value));
    }

    [Fact]
    public void ReadObjects_fails_clearly_for_a_property_type_a_text_value_cannot_fill()
    {
        var exception = Assert.Throws<NotSupportedException>(() => CsvReader.FromText("Tags\na").ReadObjects<UnsupportedType>());

        Assert.Contains("Tags", exception.Message);
    }

    private sealed class Person
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public decimal Balance { get; set; }
        public DateOnly Joined { get; set; }
        public bool Active { get; set; }
        public Guid Id { get; set; }
    }

    private sealed class LooseNames
    {
        public string? Name { get; set; }
        public decimal BalanceDue { get; set; }
        public DateOnly JoinedOn { get; set; }
    }

    private sealed class MappedPerson
    {
        [DocWeaveColumn("Full Name", Required = true)]
        public string? Name { get; set; }

        [DocWeaveColumn(Ordinal = 3)]
        public string? Third { get; set; }
    }

    private sealed class NicknamedPerson
    {
        public string? Name { get; set; }
        public int Age { get; set; }
        public string? Nickname { get; set; } = "none";
    }

    private sealed class OptionalNickname
    {
        public string? Name { get; set; }
        public string? Nickname { get; set; } = "none";
    }

    private sealed class UnsupportedType
    {
        public List<string>? Tags { get; set; }
    }
}
