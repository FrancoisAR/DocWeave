using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

public sealed class ExcelObjectImportTests
{
    [Fact]
    public void _53_Import_objects_matches_headers_ignoring_case_spaces_and_underscores()
    {
        var bytes = ExcelTestData.Load("objects-convention-headers.xlsx");

        var result = ExcelReader.FromBytes(bytes)
            .Sheet("Invoices")
            .StartAt("B4")
            .ReadObjects<ConventionInvoice>();

        Assert.False(result.HasErrors);
        Assert.Equal(2, result.RowCount);
        Assert.Equal("INV-001", result.Items[0].InvoiceNo);
        Assert.Equal("Alice", result.Items[0].Customer);
        Assert.Equal(125.50m, result.Items[0].NetAmount);
        Assert.Equal(300m, result.Items[1].NetAmount);
    }

    [Fact]
    public void _54_Import_objects_uses_attributes_for_header_name_ordinal_and_ignored_properties()
    {
        var bytes = ExcelTestData.Load("objects-attribute-mapping.xlsx");

        var invoice = Assert.Single(ExcelReader.FromBytes(bytes).ReadObjects<AttributedInvoice>().Items);

        Assert.Equal("INV-001", invoice.Number);
        Assert.Equal("Alice", invoice.Buyer);
        Assert.Equal(125.50m, invoice.Net);
        Assert.Equal("untouched", invoice.Notes);
    }

    [Fact]
    public void _55_Import_objects_counts_ordinals_from_the_start_cell()
    {
        // Imported from B4, so the second imported column is C. Column A is a note the importer never sees.
        var bytes = ExcelTestData.Load("objects-start-cell-b4.xlsx");

        var invoice = Assert.Single(ExcelReader.FromBytes(bytes).StartAt("B4").ReadObjects<AttributedInvoice>().Items);

        Assert.Equal("Alice", invoice.Buyer);
    }

    [Fact]
    public void _56_Export_typed_objects_and_import_them_back_as_objects()
    {
        var exported = new[]
        {
            new TypedInvoiceExport("INV-001", 125.50m, new DateTime(2026, 1, 7), true, InvoiceStatus.Paid, 4),
            new TypedInvoiceExport("INV-002", 300m, new DateTime(2026, 1, 9), false, InvoiceStatus.Open, null),
            new TypedInvoiceExport("INV-003", 15.25m, new DateTime(2026, 2, 1), false, InvoiceStatus.Closed, 12)
        };

        var bytes = ExcelWriter.Create()
            .AddSheet("Invoices", sheet => sheet.AddTable(exported).StartAt("A1"))
            .ToBytes();
        ExcelTestOutput.SaveWorkbook(nameof(_56_Export_typed_objects_and_import_them_back_as_objects), bytes);

        var imported = ExcelReader.FromBytes(bytes).ReadObjects<TypedInvoiceExport>();

        Assert.False(imported.HasErrors);
        Assert.Equal(exported, imported.Items);
    }

    [Fact]
    public void _57_Import_objects_reads_excel_date_numbers_for_date_properties()
    {
        // Excel stores dates as numbers: 45123 is 16 July 2023, and the fraction is the time of day.
        var bytes = ExcelTestData.Load("objects-excel-date-numbers.xlsx");

        var row = Assert.Single(ExcelReader.FromBytes(bytes).ReadObjects<DatedRow>().Items);

        Assert.Equal(new DateTime(2023, 7, 16, 12, 0, 0), row.When);
        Assert.Equal(new DateOnly(2023, 7, 16), row.Day);
    }

    [Fact]
    public void _58_Import_objects_collects_problems_and_still_returns_the_good_rows()
    {
        var bytes = ExcelTestData.Load("objects-good-and-bad-rows.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<TypedRow>();

        Assert.Equal([1, 4], result.Items.Select(item => item.Count));
        Assert.Equal(4, result.RowCount);
        Assert.True(result.HasErrors);
        Assert.Equal(
            [(3, "Paid"), (4, "Status"), (4, "Date")],
            result.Errors.Select(error => (error.RowNumber, error.Column!)));
        Assert.Equal("maybe", result.Errors[0].Value);
        Assert.Equal("Paid", result.Errors[0].Property);
        Assert.Contains("Open, Paid, Closed", result.Errors[1].Message);
    }

    [Fact]
    public void _59_Import_objects_treats_a_blank_as_an_error_only_for_values_that_cannot_be_empty()
    {
        var bytes = ExcelTestData.Load("objects-blank-values.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<TypedRow>();

        // Count (int) cannot hold "no value", so a blank is an error. OptionalCount (int?) can. Ratio only makes the row exist.
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal("Count", error.Column);
        Assert.Equal("A value is required.", error.Message);
    }

    [Fact]
    public void _60_Import_objects_rejects_fractions_and_out_of_range_values_for_whole_number_properties()
    {
        var bytes = ExcelTestData.Load("objects-whole-number-limits.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<TypedRow>();

        Assert.Equal([3], result.Items.Select(item => item.Count));
        Assert.Equal(2, result.Errors.Count);
        Assert.Contains("whole number", result.Errors[0].Message);
        Assert.Contains("range", result.Errors[1].Message);
    }

    [Fact]
    public void _61_Import_objects_reports_a_missing_required_column_once_and_returns_nothing()
    {
        var bytes = ExcelTestData.Load("objects-missing-required-column.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<RequiredRow>();

        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(0, error.RowNumber);
        Assert.Equal("Code", error.Property);
        Assert.Contains("Required column 'Code' was not found", error.Message);
    }

    [Fact]
    public void _62_Import_objects_reports_a_blank_required_value_and_keeps_defaults_for_blank_optional_values()
    {
        var bytes = ExcelTestData.Load("objects-blank-required-value.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<RequiredRow>();

        var item = Assert.Single(result.Items);
        Assert.Equal("A", item.Code);
        Assert.Equal("default", item.Label);
        var error = Assert.Single(result.Errors);
        Assert.Equal(3, error.RowNumber);
        Assert.Equal("Code", error.Column);
    }

    [Fact]
    public void _63_Import_objects_fails_clearly_for_a_property_type_a_cell_cannot_fill()
    {
        var bytes = ExcelTestData.Load("objects-unsupported-property-type.xlsx");

        var exception = Assert.Throws<NotSupportedException>(() => ExcelReader.FromBytes(bytes).ReadObjects<BadTypeRow>());

        Assert.Contains("Tags", exception.Message);
        Assert.Contains("DocWeaveIgnore", exception.Message);
    }

    [Fact]
    public void _64_Enumerate_objects_reports_problems_to_the_callback_as_it_goes()
    {
        var bytes = ExcelTestData.Load("objects-one-bad-count.xlsx");
        var errors = new List<DocWeaveImportError>();

        var counts = ExcelReader.FromBytes(bytes).EnumerateObjects<TypedRow>(errors.Add).Select(item => item.Count).ToList();

        Assert.Equal([1, 3], counts);
        Assert.Equal(3, Assert.Single(errors).RowNumber);
    }

    [Fact]
    public void _65_Import_objects_reports_progress_and_audit_without_cell_values()
    {
        var bytes = ExcelTestData.Load("objects-progress-and-audit.xlsx");
        var progress = new List<DocWeaveOperationProgress>();
        var audits = new List<DocWeaveAuditEvent>();

        ExcelReader.FromBytes(bytes)
            .WithProgress(new SynchronousProgress<DocWeaveOperationProgress>(progress.Add))
            .Audit(audits.Add)
            .ReadObjects<TypedRow>();

        Assert.Equal(["Starting", "Completed"], progress.Select(item => item.Stage));
        var audit = Assert.Single(audits);
        Assert.Equal(2, audit.Properties["RowCount"]);
        Assert.Equal(1, audit.Properties["ErrorCount"]);
        Assert.DoesNotContain(audit.Properties.Values, value => "secret-value".Equals(value));
    }

    [Fact]
    public void _66_Import_objects_from_an_empty_sheet_still_reports_a_missing_required_column()
    {
        var bytes = ExcelTestData.Load("objects-header-only-sheet.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<RequiredRow>();

        Assert.Empty(result.Items);
        Assert.Single(result.Errors);
    }

    [Fact]
    public void _67_Import_positional_records_through_their_constructor_with_a_fluent_map()
    {
        var bytes = ExcelTestData.Load("records-invoice-columns.xlsx");

        var record = Assert.Single(ExcelReader.FromBytes(bytes).ReadObjects<InvoiceRecord>().Items);
        var mapped = Assert.Single(ExcelReader.FromBytes(bytes).ReadObjects<PlainInvoice>(map => map
            .Column(x => x.InvoiceNumber, "Invoice No")
            .Column(x => x.Net, "Net Amount")).Items);

        Assert.Equal(new InvoiceRecord("INV-001", 12.5m, new DateOnly(2023, 7, 16)), record);
        Assert.Equal("INV-001", mapped.InvoiceNumber);
        Assert.Equal(12.5m, mapped.Net);
    }

    [Fact]
    public void _68_Import_a_record_uses_parameter_defaults_and_reports_a_missing_required_parameter()
    {
        var bytes = ExcelTestData.Load("records-reps-blank-age.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<RepRecord>();

        Assert.Equal([new RepRecord("Ada", 36), new RepRecord("Grace", 30)], result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(4, error.RowNumber);
        Assert.Equal("Sales Rep", error.Column);
    }

    [Fact]
    public void _69_Import_a_class_whose_constructor_rejects_a_row_reports_the_reason()
    {
        var bytes = ExcelTestData.Load("objects-constructor-validation.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<PositiveValue>();

        Assert.Equal([3, 5], result.Items.Select(item => item.Value));
        var error = Assert.Single(result.Errors);
        Assert.Equal(3, error.RowNumber);
        Assert.Contains("must be positive", error.Message);
    }

    [Fact]
    public void _70_Import_with_a_fluent_map_by_header_position_required_and_ignore()
    {
        var bytes = ExcelTestData.Load("objects-fluent-map-invoices.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadObjects<PlainInvoice>(map => map
            .Column(x => x.InvoiceNumber, "Number")
            .Column(x => x.Customer, 2, required: true)
            .Column(x => x.Net, "Amount")
            .Ignore(x => x.Scratch));

        var item = Assert.Single(result.Items);
        Assert.Equal("INV-001", item.InvoiceNumber);
        Assert.Equal("Alice", item.Customer);
        Assert.Equal(125.50m, item.Net);
        Assert.Equal("untouched", item.Scratch);
        var error = Assert.Single(result.Errors);
        Assert.Equal(3, error.RowNumber);
        Assert.Equal("Customer", error.Property);
    }

    [Fact]
    public void _71_Enumerate_objects_accepts_a_fluent_map_and_column_selection()
    {
        var bytes = ExcelTestData.Load("objects-column-selection.xlsx");

        // After selecting Number and Buyer, Buyer is the 2nd imported column, even though it is the 3rd in the sheet.
        var invoices = ExcelReader.FromBytes(bytes)
            .Columns(columns => columns.IncludeHeaders("Number", "Buyer"))
            .EnumerateObjects<PlainInvoice>(map => map
                .Column(x => x.InvoiceNumber, "Number")
                .Column(x => x.Customer, 2), null)
            .ToList();

        Assert.Equal(["Alice", "Bob"], invoices.Select(invoice => invoice.Customer));
    }

    private enum InvoiceStatus
    {
        Open,
        Paid,
        Closed
    }

    private sealed record TypedInvoiceExport(string InvoiceNo, decimal Net, DateTime InvoiceDate, bool Paid, InvoiceStatus Status, int? Quantity);

    private sealed class ConventionInvoice
    {
        public string? InvoiceNo { get; set; }
        public string? Customer { get; set; }
        public decimal NetAmount { get; set; }
    }

    private sealed class AttributedInvoice
    {
        [DocWeaveColumn("Invoice Number")]
        public string? Number { get; set; }

        [DocWeaveColumn(Ordinal = 2)]
        public string? Buyer { get; set; }

        [DocWeaveIgnore]
        public string? Notes { get; set; } = "untouched";

        public decimal Net { get; set; }
    }

    private sealed class TypedRow
    {
        public int Count { get; set; }
        public int? OptionalCount { get; set; }
        public bool Paid { get; set; }
        public InvoiceStatus Status { get; set; }
        public DateTime Date { get; set; }
        public double Ratio { get; set; }
    }

    private sealed class DatedRow
    {
        public DateTime When { get; set; }
        public DateOnly Day { get; set; }
    }

    private sealed class RequiredRow
    {
        [DocWeaveColumn(Required = true)]
        public string? Code { get; set; }

        public string? Label { get; set; } = "default";
    }

    private sealed class BadTypeRow
    {
        public List<string>? Tags { get; set; }
    }

    private sealed record InvoiceRecord(string InvoiceNo, decimal NetAmount, DateOnly Date);

    private sealed record RepRecord([DocWeaveColumn("Sales Rep", Required = true)] string Name, int Age = 30);

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

    private sealed class PlainInvoice
    {
        public string? InvoiceNumber { get; set; }
        public string? Customer { get; set; }
        public decimal Net { get; set; }
        public string? Scratch { get; set; } = "untouched";
    }

    private sealed class SynchronousProgress<T>(Action<T> handler) : IProgress<T>
    {
        public void Report(T value) => handler(value);
    }
}
