using System.Diagnostics;
using DocWeave.Csv;
using Xunit.Abstractions;

namespace DocWeave.Tests.Csv;

public sealed class CsvObjectFixtureTests
{
    private readonly ITestOutputHelper output;

    public CsvObjectFixtureTests(ITestOutputHelper output)
    {
        this.output = output;
    }

    [Theory]
    [InlineData("mock-invoices-100.csv", 100)]
    [InlineData("mock-invoices-1000.csv", 1000)]
    [InlineData("mock-invoices-10000.csv", 10000)]
    [InlineData("mock-invoices-50000.csv", 50000)]
    public void ReadObjects_maps_fixture_rows_into_records_without_problems(string fileName, int expectedRows)
    {
        var path = FixturePath(fileName);
        var stopwatch = Stopwatch.StartNew();

        var result = CsvReader.FromFile(path).ReadObjects<MockInvoice>();

        stopwatch.Stop();
        output.WriteLine($"{fileName}: {result.Items.Count} objects in {stopwatch.ElapsedMilliseconds} ms");

        Assert.False(result.HasErrors);
        Assert.Equal(expectedRows, result.RowCount);
        Assert.Equal(expectedRows, result.Items.Count);
        // The two large fixtures pad their numbers to six digits, the small ones to four.
        var digits = expectedRows >= 10000 ? 6 : 4;
        Assert.Equal($"INV-2026-{1.ToString().PadLeft(digits, '0')}", result.Items[0].InvoiceNo);
        Assert.Equal($"Customer {1.ToString().PadLeft(digits, '0')}", result.Items[0].CustomerName);
        Assert.Equal(new DateOnly(2026, 1, 2), result.Items[0].InvoiceDate);
        Assert.Equal(137.91m, result.Items[0].Amount);
        Assert.Equal("Mock invoice row 1", result.Items[0].Notes);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void EnumerateObjects_streams_a_large_fixture_one_object_at_a_time()
    {
        var path = FixturePath("mock-invoices-50000.csv");
        var count = 0;
        var total = 0m;

        foreach (var invoice in CsvReader.FromFile(path).EnumerateObjects<MockInvoice>())
        {
            count++;
            total += invoice.Amount;
        }

        Assert.Equal(50000, count);
        Assert.True(total > 0m);
    }

    [Fact]
    public void ReadObjects_reads_a_headerless_fixture_by_position_with_a_map()
    {
        // The file has no header row: Id, Name, Region, Amount, Status. Columns are then named A to E.
        var path = FixturePath("mock-no-header-1000.csv");

        var result = CsvReader.FromFile(path)
            .HasHeaderRow(false)
            .ReadObjects<NoHeaderRow>(map => map
                .Column(x => x.Id, 1)
                .Column(x => x.Name, 2)
                .Column(x => x.Region, 3)
                .Column(x => x.Amount, 4)
                .Column(x => x.Status, 5));

        Assert.False(result.HasErrors);
        Assert.Equal(1000, result.Items.Count);
        Assert.Equal(1, result.Items[0].Id);
        Assert.Equal("Name 1", result.Items[0].Name);
        Assert.Equal("London", result.Items[0].Region);
        Assert.Equal(117.17m, result.Items[0].Amount);
        Assert.Equal("Open", result.Items[0].Status);
    }

    [Fact]
    public void ReadObjects_reads_a_pipe_delimited_fixture()
    {
        var path = FixturePath("mock-pipe-1000.csv");

        var result = CsvReader.FromFile(path).DelimitedBy("|").ReadObjects<PipeRow>();

        Assert.False(result.HasErrors);
        Assert.Equal(1000, result.Items.Count);
        Assert.All(result.Items, row => Assert.False(string.IsNullOrEmpty(row.Name)));
    }

    [Fact]
    public void ReadObjects_selects_columns_from_a_wide_fixture_and_maps_only_those()
    {
        var path = FixturePath("mock-wide-1000x75.csv");

        var result = CsvReader.FromFile(path)
            .Columns(columns => columns.IncludeLetters("A", "B"))
            .ReadObjects<WideRow>(map => map
                .Column(x => x.First, 1)
                .Column(x => x.Second, 2));

        Assert.False(result.HasErrors);
        Assert.Equal(1000, result.Items.Count);
        Assert.False(string.IsNullOrEmpty(result.Items[0].First));
    }

    private static string FixturePath(string fileName)
    {
        return Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", fileName);
    }

    private sealed record MockInvoice(
        string InvoiceNo,
        string CustomerName,
        DateOnly InvoiceDate,
        string Region,
        decimal Amount,
        string Status,
        string Notes);

    private sealed class NoHeaderRow
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public string? Region { get; set; }
        public decimal Amount { get; set; }
        public string? Status { get; set; }
    }

    private sealed class PipeRow
    {
        public string? Name { get; set; }
    }

    private sealed class WideRow
    {
        public string? First { get; set; }
        public string? Second { get; set; }
    }
}
