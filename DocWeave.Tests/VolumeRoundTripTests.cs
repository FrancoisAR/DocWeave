using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using DocWeave.Csv;
using DocWeave.Excel;
using Xunit.Abstractions;

namespace DocWeave.Tests;

// Prevent simultaneous multi-million-cell workbooks within one test assembly.
[CollectionDefinition("VolumeMatrix", DisableParallelization = true)]
public sealed class VolumeMatrixCollection { }

public sealed class VolumeMatrixTheoryAttribute : TheoryAttribute
{
    public VolumeMatrixTheoryAttribute()
    {
        if (Environment.GetEnvironmentVariable("DOCWEAVE_RUN_VOLUME_MATRIX") != "1")
            Skip = "Set DOCWEAVE_RUN_VOLUME_MATRIX=1 to run the large volume matrix.";
    }
}

[Collection("VolumeMatrix")]
public sealed class VolumeRoundTripTests(ITestOutputHelper output)
{
    [Theory]
    [InlineData("csv")]
    [InlineData("xlsx")]
    public void Smoke_checks_volume_harness(string format) => Run(100, format);

    [VolumeMatrixTheory]
    [Trait("Category", "VolumeMatrix")]
    [InlineData(10_000, "csv"), InlineData(10_000, "xlsx")]
    [InlineData(50_000, "csv"), InlineData(50_000, "xlsx")]
    [InlineData(100_000, "csv"), InlineData(100_000, "xlsx")]
    [InlineData(150_000, "csv"), InlineData(150_000, "xlsx")]
    [InlineData(200_000, "csv"), InlineData(200_000, "xlsx")]
    [InlineData(250_000, "csv"), InlineData(250_000, "xlsx")]
    [InlineData(300_000, "csv"), InlineData(300_000, "xlsx")]
    [InlineData(400_000, "csv"), InlineData(400_000, "xlsx")]
    [InlineData(500_000, "csv"), InlineData(500_000, "xlsx")]
    public void Export_then_import_mixed_records(int rows, string format) => Run(rows, format);

    private void Run(int rowCount, string format)
    {
        var resultRoot = Environment.GetEnvironmentVariable("DOCWEAVE_VOLUME_RESULTS")
            ?? Path.Combine(AppContext.BaseDirectory, "volume-results");
        Directory.CreateDirectory(resultRoot);
        var runId = Guid.NewGuid().ToString("N");
        var file = Path.Combine(resultRoot, $"{format}-{rowCount}-{runId}.{format}");
        var timings = new Dictionary<string, double>();
        var total = Stopwatch.StartNew();
        var status = "failed";
        long fileBytes = 0;
        try
        {
            var watch = Stopwatch.StartNew();
            using var table = CreateData(rowCount);
            timings["generateMs"] = watch.Elapsed.TotalMilliseconds;
            var expectedCost = table.Rows.Cast<DataRow>().Sum(r => (decimal)r["Cost"]);
            var expectedRevenue = table.Rows.Cast<DataRow>().Sum(r => (decimal)r["Revenue"]);

            // CSV has no formula engine: append a precomputed totals record.
            // XLSX uses the exporter's generated SUM formulas.
            watch.Restart();
            if (format == "csv")
            {
                var totals = table.NewRow();
                totals["Id"] = "Totals";
                foreach (DataColumn column in table.Columns)
                {
                    if (column.DataType == typeof(decimal))
                        totals[column] = table.Rows.Cast<DataRow>().Sum(r => (decimal)r[column]);
                    else if (column.DataType == typeof(double))
                        totals[column] = table.Rows.Cast<DataRow>().Sum(r => (double)r[column]);
                    else if (column.DataType == typeof(int))
                        totals[column] = checked(table.Rows.Cast<DataRow>().Sum(r => (int)r[column]));
                    else if (column.DataType == typeof(long))
                        totals[column] = table.Rows.Cast<DataRow>().Sum(r => (long)r[column]);
                }
                table.Rows.Add(totals);
                CsvWriter.FromDataTable(table).SaveAs(file);
            }
            else
            {
                ExcelWriter.Create().AddSheet("Volume", sheet => sheet.AddTable(table)
                    .HeaderStyle(s => s.Bold())
                    .ColumnStyle("Date", s => s.NumberFormat(ExcelNumberFormat.Date))
                    .ColumnStyle("Cost", s => s.NumberFormat(ExcelNumberFormat.Number))
                    .ColumnStyle("Revenue", s => s.NumberFormat(ExcelNumberFormat.Number))
                    .AddTotals.Bottom("Totals", s => s.Bold().BorderTop().BorderBottom(ExcelBorderLineStyle.Double)))
                    .SaveAs(file);
            }
            timings["exportMs"] = watch.Elapsed.TotalMilliseconds;
            fileBytes = new FileInfo(file).Length;

            // Stream the exported file back; check every record rather than retaining another full table.
            watch.Restart();
            var records = format == "csv"
                ? CsvReader.FromFile(file).EnumerateRows().Select(r => r.Values)
                : ExcelReader.FromFile(file).Sheet("Volume").EnumerateRows().Select(r => r.Values);
            int count = 0;
            decimal cost = 0, revenue = 0;
            foreach (var record in records)
            {
                Assert.Equal(20, record.Count);
                if (count < rowCount)
                {
                    Assert.Equal($"ROW-{count + 1:000000}", record["Id"]);
                    cost += decimal.Parse(record["Cost"]!, CultureInfo.InvariantCulture);
                    revenue += decimal.Parse(record["Revenue"]!, CultureInfo.InvariantCulture);
                    Assert.Equal($"Description {count + 1}, with \"quotes\"", record["Description"]);
                    Assert.False(string.IsNullOrEmpty(record["Date"]));
                }
                else
                {
                    Assert.Equal(rowCount, count);
                    Assert.Equal("Totals", record["Id"]);
                    if (format == "csv")
                    {
                        Assert.Equal(expectedCost, decimal.Parse(record["Cost"]!, CultureInfo.InvariantCulture));
                        Assert.Equal(expectedRevenue, decimal.Parse(record["Revenue"]!, CultureInfo.InvariantCulture));
                    }
                }
                count++;
            }
            timings["importMs"] = watch.Elapsed.TotalMilliseconds;
            Assert.Equal(rowCount + 1, count);
            Assert.Equal(expectedCost, cost);
            Assert.Equal(expectedRevenue, revenue);
            status = "passed";
        }
        finally
        {
            total.Stop();
            timings["totalMs"] = total.Elapsed.TotalMilliseconds;
            var json = JsonSerializer.Serialize(new
            {
                runId, utc = DateTimeOffset.UtcNow, package = typeof(ExcelWriter).Assembly.GetName().Name,
                runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                rowCount, columns = 20, format, status, fileBytes, timings,
                peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64
            });
            File.AppendAllText(Path.Combine(resultRoot, $"timings-{runId}.jsonl"), json + Environment.NewLine);
            output.WriteLine(json);
            if (status == "passed" && Environment.GetEnvironmentVariable("DOCWEAVE_KEEP_VOLUME_FILES") != "1")
                File.Delete(file);
        }
    }

    private static DataTable CreateData(int count)
    {
        var random = new Random(4721);
        var table = new DataTable("Volume");
        var columns = new (string Name, Type Type)[]
        {
            ("Id", typeof(string)), ("Date", typeof(DateTime)), ("Cost", typeof(decimal)),
            ("Revenue", typeof(decimal)), ("Rate", typeof(double)), ("Quantity", typeof(int)),
            ("Description", typeof(string)), ("Active", typeof(bool)), ("Category", typeof(string)),
            ("Reference", typeof(long)), ("Tax", typeof(decimal)), ("Discount", typeof(decimal)),
            ("Weight", typeof(double)), ("Score", typeof(double)), ("City", typeof(string)),
            ("Notes", typeof(string)), ("DueDate", typeof(DateTime)), ("Balance", typeof(decimal)),
            ("Currency", typeof(string)), ("Batch", typeof(int))
        };
        foreach (var column in columns) table.Columns.Add(column.Name, column.Type);
        for (int i = 1; i <= count; i++)
            table.Rows.Add($"ROW-{i:000000}", new DateTime(2020, 1, 1).AddDays(random.Next(3000)),
                random.Next(100_000) / 100m, random.Next(200_000) / 100m, random.NextDouble(),
                random.Next(1, 100), $"Description {i}, with \"quotes\"", i % 2 == 0,
                $"Category {random.Next(8)}", 1_000_000_000L + i, random.Next(1000) / 100m,
                random.Next(1000) / 100m, random.NextDouble() * 100, random.NextDouble() * 10,
                $"City {random.Next(50)}", i % 7 == 0 ? "First line\nSecond line" : "Note",
                new DateTime(2030, 1, 1).AddDays(random.Next(365)), random.Next(100_000) / 100m,
                i % 2 == 0 ? "GBP" : "USD", random.Next(1, 100));
        return table;
    }
}