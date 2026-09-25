using System.Data;
using System.Diagnostics;
using DocWeave.Excel;
using Xunit.Abstractions;

namespace DocWeave.Tests.Excel;

public sealed class ExcelVolumeImportTests
{
    private const string RunLargeExcelVolumeVariable = "DOCWEAVE_RUN_LARGE_EXCEL_VOLUME";
    private readonly ITestOutputHelper output;

    public ExcelVolumeImportTests(ITestOutputHelper output)
    {
        this.output = output;
    }

    [Theory]
    [InlineData(20_000, false)]
    [InlineData(200_000, true)]
    [Trait("Category", "Performance")]
    public void Enumerate_rows_streams_a_large_worksheet(int rowCount, bool requiresVolumeOptIn)
    {
        if (ShouldSkipVolumeTest(requiresVolumeOptIn))
        {
            return;
        }

        var bytes = CreateInvoiceWorkbook(rowCount);
        var stopwatch = Stopwatch.StartNew();

        var count = 0;
        var lastInvoiceNo = string.Empty;
        foreach (var row in ExcelReader.FromBytes(bytes).EnumerateRows())
        {
            count++;
            lastInvoiceNo = row.Values["InvoiceNo"]!;
        }

        stopwatch.Stop();
        output.WriteLine($"{rowCount} rows streamed from a {bytes.Length / 1024} KB workbook in {stopwatch.ElapsedMilliseconds} ms");

        Assert.Equal(rowCount, count);
        Assert.Equal($"INV-{rowCount:000000}", lastInvoiceNo);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(30));
    }

    [Theory]
    [InlineData(20_000, false)]
    [InlineData(200_000, true)]
    [Trait("Category", "Performance")]
    public void Read_objects_maps_a_large_worksheet_into_typed_rows(int rowCount, bool requiresVolumeOptIn)
    {
        if (ShouldSkipVolumeTest(requiresVolumeOptIn))
        {
            return;
        }

        var bytes = CreateInvoiceWorkbook(rowCount);
        var stopwatch = Stopwatch.StartNew();

        var result = ExcelReader.FromBytes(bytes).ReadObjects<VolumeInvoice>();

        stopwatch.Stop();
        output.WriteLine($"{rowCount} rows mapped to objects in {stopwatch.ElapsedMilliseconds} ms");

        Assert.False(result.HasErrors);
        Assert.Equal(rowCount, result.Items.Count);
        Assert.Equal(new VolumeInvoice($"INV-{rowCount:000000}", $"Customer {rowCount}", rowCount * 1.5m, rowCount % 2 == 0), result.Items[^1]);
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(30));
    }

    [Fact]
    public void Read_objects_reports_every_bad_row_in_a_large_worksheet()
    {
        // Every 100th row has a bad amount, so a large file with scattered problems is still read to the end.
        const int rowCount = 10_000;
        var table = new DataTable("Invoices");
        table.Columns.Add("InvoiceNo", typeof(string));
        table.Columns.Add("Customer", typeof(string));
        table.Columns.Add("Net", typeof(string));
        table.Columns.Add("Paid", typeof(string));
        for (var i = 1; i <= rowCount; i++)
        {
            table.Rows.Add($"INV-{i:000000}", $"Customer {i}", i % 100 == 0 ? "n/a" : (i * 1.5m).ToString(System.Globalization.CultureInfo.InvariantCulture), "true");
        }

        var bytes = ExcelWriter.Create().AddSheet("Invoices", sheet => sheet.AddTable(table).StartAt("A1")).ToBytes();

        var result = ExcelReader.FromBytes(bytes).ReadObjects<VolumeInvoice>();

        Assert.Equal(rowCount, result.RowCount);
        Assert.Equal(rowCount - 100, result.Items.Count);
        Assert.Equal(100, result.Errors.Count);
        Assert.All(result.Errors, error => Assert.Equal("Net", error.Column));
        Assert.Equal(101, result.Errors[0].RowNumber);
    }

    private static byte[] CreateInvoiceWorkbook(int rowCount)
    {
        var invoices = Enumerable.Range(1, rowCount)
            .Select(i => new VolumeInvoice($"INV-{i:000000}", $"Customer {i}", i * 1.5m, i % 2 == 0))
            .ToList();

        return ExcelWriter.Create()
            .AddSheet("Invoices", sheet => sheet.AddTable(invoices).StartAt("A1"))
            .ToBytes();
    }

    private static bool ShouldSkipVolumeTest(bool requiresVolumeOptIn)
    {
        return requiresVolumeOptIn
            && !string.Equals(Environment.GetEnvironmentVariable(RunLargeExcelVolumeVariable), "1", StringComparison.Ordinal);
    }

    private sealed record VolumeInvoice(string InvoiceNo, string Customer, decimal Net, bool Paid);
}
