using System.Diagnostics;
using System.Text;
using DocWeave.Csv;
using Xunit.Abstractions;

namespace DocWeave.Tests.Csv;

public sealed class CsvVolumePerformanceTests
{
    private const string RunVolumeTestsVariable = "DOCWEAVE_RUN_VOLUME_TESTS";
    private readonly ITestOutputHelper output;

    public CsvVolumePerformanceTests(ITestOutputHelper output)
    {
        this.output = output;
    }

    [Theory]
    [InlineData(100_000, false)]
    [InlineData(1_000_000, true)]
    [Trait("Category", "Performance")]
    public void Import_can_stream_large_csv_files(int rowCount, bool requiresVolumeOptIn)
    {
        if (ShouldSkipVolumeTest(requiresVolumeOptIn))
        {
            return;
        }

        var path = Path.Combine(Path.GetTempPath(), $"docweave-import-volume-{rowCount}-{Guid.NewGuid():N}.csv");
        try
        {
            GenerateCsvFile(path, rowCount);

            var stopwatch = Stopwatch.StartNew();
            var count = 0;
            string? firstId = null;
            string? lastId = null;

            foreach (var row in CsvReader.FromFile(path)
                         .Columns(columns => columns.IncludeHeaders("ID", "Amount", "Status"))
                         .EnumerateRows())
            {
                count++;
                firstId ??= row.Values["ID"];
                lastId = row.Values["ID"];
            }

            stopwatch.Stop();
            output.WriteLine($"Imported {count:N0} rows from {Path.GetFileName(path)} in {stopwatch.Elapsed}.");

            Assert.Equal(rowCount, count);
            Assert.Equal("1", firstId);
            Assert.Equal(rowCount.ToString(), lastId);
            Assert.True(stopwatch.Elapsed < TimeoutFor(rowCount), $"Import took {stopwatch.Elapsed}.");
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    [Theory]
    [InlineData(100_000, false)]
    [InlineData(1_000_000, true)]
    [Trait("Category", "Performance")]
    public void Export_can_write_large_csv_files(int rowCount, bool requiresVolumeOptIn)
    {
        if (ShouldSkipVolumeTest(requiresVolumeOptIn))
        {
            return;
        }

        var path = Path.Combine(Path.GetTempPath(), $"docweave-export-volume-{rowCount}-{Guid.NewGuid():N}.csv");
        try
        {
            var stopwatch = Stopwatch.StartNew();

            CsvWriter.FromRows(Enumerable.Range(1, rowCount).Select(index => new VolumeCsvRow(index)))
                .NewLine("\n")
                .SaveAs(path);

            stopwatch.Stop();
            var lineCount = CountLines(path);
            output.WriteLine($"Exported {rowCount:N0} rows to {Path.GetFileName(path)} ({new FileInfo(path).Length:N0} bytes) in {stopwatch.Elapsed}.");

            Assert.Equal(rowCount + 1, lineCount);
            Assert.True(new FileInfo(path).Length > rowCount * 10L);
            Assert.True(stopwatch.Elapsed < TimeoutFor(rowCount), $"Export took {stopwatch.Elapsed}.");
        }
        finally
        {
            DeleteIfExists(path);
        }
    }

    private bool ShouldSkipVolumeTest(bool requiresVolumeOptIn)
    {
        if (!requiresVolumeOptIn || string.Equals(Environment.GetEnvironmentVariable(RunVolumeTestsVariable), "1", StringComparison.Ordinal))
        {
            return false;
        }

        output.WriteLine($"Skipping 1,000,000 row volume case. Set {RunVolumeTestsVariable}=1 to run it.");
        return true;
    }

    private static TimeSpan TimeoutFor(int rowCount)
    {
        return rowCount >= 1_000_000 ? TimeSpan.FromMinutes(3) : TimeSpan.FromSeconds(30);
    }

    private static void GenerateCsvFile(string path, int rowCount)
    {
        using var writer = new StreamWriter(path, append: false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.WriteLine("ID,Region,Amount,Status,Note");

        for (var index = 1; index <= rowCount; index++)
        {
            writer.Write(index);
            writer.Write(',');
            writer.Write(GetRegion(index));
            writer.Write(',');
            writer.Write((index * 1.25m).ToString(System.Globalization.CultureInfo.InvariantCulture));
            writer.Write(',');
            writer.Write(index % 2 == 0 ? "Closed" : "Open");
            writer.Write(',');
            writer.Write("Generated row ");
            writer.WriteLine(index);
        }
    }

    private static int CountLines(string path)
    {
        var count = 0;
        using var reader = File.OpenText(path);
        while (reader.ReadLine() is not null)
        {
            count++;
        }

        return count;
    }

    private static string GetRegion(int index)
    {
        return (index % 4) switch
        {
            0 => "London",
            1 => "United States",
            2 => "Europe",
            _ => "Asia"
        };
    }

    private static void DeleteIfExists(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class VolumeCsvRow : IReadOnlyDictionary<string, object?>
    {
        private static readonly string[] ColumnNames = ["ID", "Region", "Amount", "Status", "Note"];
        private readonly int index;

        public VolumeCsvRow(int index)
        {
            this.index = index;
        }

        public IEnumerable<string> Keys => ColumnNames;

        public IEnumerable<object?> Values => ColumnNames.Select(column => this[column]);

        public int Count => ColumnNames.Length;

        public object? this[string key] => TryGetValue(key, out var value)
            ? value
            : throw new KeyNotFoundException(key);

        public bool ContainsKey(string key)
        {
            return ColumnNames.Contains(key, StringComparer.OrdinalIgnoreCase);
        }

        public bool TryGetValue(string key, out object? value)
        {
            value = key.ToUpperInvariant() switch
            {
                "ID" => index,
                "REGION" => GetRegion(index),
                "AMOUNT" => index * 1.25m,
                "STATUS" => index % 2 == 0 ? "Closed" : "Open",
                "NOTE" => $"Generated row {index}",
                _ => null
            };

            return value is not null;
        }

        public IEnumerator<KeyValuePair<string, object?>> GetEnumerator()
        {
            foreach (var column in ColumnNames)
            {
                yield return new KeyValuePair<string, object?>(column, this[column]);
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
