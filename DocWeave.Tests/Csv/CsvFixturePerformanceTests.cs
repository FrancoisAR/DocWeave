using System.Diagnostics;
using DocWeave.Csv;
using Xunit.Abstractions;

namespace DocWeave.Tests.Csv;

public sealed class CsvFixturePerformanceTests
{
    private readonly ITestOutputHelper output;

    public CsvFixturePerformanceTests(ITestOutputHelper output)
    {
        this.output = output;
    }

    [Theory]
    [InlineData("mock-invoices-100.csv", 100)]
    [InlineData("mock-invoices-1000.csv", 1000)]
    [InlineData("mock-invoices-10000.csv", 10000)]
    [InlineData("mock-invoices-50000.csv", 50000)]
    public void ReadDataTable_loads_fixture_with_expected_row_count(string fileName, int expectedRows)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", fileName);
        var stopwatch = Stopwatch.StartNew();

        var table = CsvReader.FromFile(path)
            .Columns(columns => columns.IncludeLetters("A", "B", "E"))
            .ReadDataTable();

        stopwatch.Stop();
        output.WriteLine($"{fileName}: {table.Rows.Count} rows in {stopwatch.ElapsedMilliseconds} ms");

        Assert.Equal(expectedRows, table.Rows.Count);
        Assert.Equal(3, table.Columns.Count);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void ReadDataTable_can_select_sparse_columns_from_wide_fixture()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", "mock-wide-1000x75.csv");
        var stopwatch = Stopwatch.StartNew();

        var table = CsvReader.FromFile(path)
            .Columns(columns => columns.IncludeLetters("A", "B", "Z", "AA", "BW"))
            .ReadDataTable();

        stopwatch.Stop();
        output.WriteLine($"mock-wide-1000x75.csv: {table.Rows.Count} rows, {table.Columns.Count} columns in {stopwatch.ElapsedMilliseconds} ms");

        Assert.Equal(1000, table.Rows.Count);
        Assert.Equal(["Column001", "Column002", "Column026", "Column027", "Column075"], table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName));
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void ReadDataTable_can_load_pipe_fixture()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", "mock-pipe-1000.csv");

        var table = CsvReader.FromFile(path)
            .DelimitedBy("|")
            .Columns(columns => columns.IncludeHeaders("ID", "Region", "Amount"))
            .ReadDataTable();

        Assert.Equal(1000, table.Rows.Count);
        Assert.Equal(3, table.Columns.Count);
    }

    [Fact]
    public void ReadDataTable_can_load_no_header_fixture()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", "mock-no-header-1000.csv");

        var table = CsvReader.FromFile(path)
            .HasHeaderRow(false)
            .Columns(columns => columns.IncludeLetters("A", "C", "E"))
            .ReadDataTable();

        Assert.Equal(1000, table.Rows.Count);
        Assert.Equal(["A", "C", "E"], table.Columns.Cast<System.Data.DataColumn>().Select(column => column.ColumnName));
    }
}
