using System.Data;
using DocWeave;
using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

public sealed class CsvWriterTests
{
    [Fact]
    public void _01_Export_datatable_with_custom_delimiter_column_order_and_header_aliases()
    {
        var table = new DataTable("People");
        table.Columns.Add("ID", typeof(int));
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Region", typeof(string));
        table.Columns.Add("Amount", typeof(decimal));
        table.Rows.Add(1, "Alice", "London", 125.50m);
        table.Rows.Add(2, "Bob", "New York", 300m);

        var csv = CsvWriter.FromDataTable(table)
            .DelimitedBy("|")
            .ColumnOrder("Name", "ID", "Amount")
            .HeaderAlias("Name", "Full Name")
            .HeaderAlias("ID", "Identifier")
            .ToText();

        CsvTestOutput.SaveCsv(nameof(_01_Export_datatable_with_custom_delimiter_column_order_and_header_aliases), csv);

        var imported = CsvReader.FromText(csv)
            .DelimitedBy("|")
            .Columns(columns => columns.IncludeHeaders("Full Name", "Identifier", "Amount"))
            .ReadDictionaries();

        Assert.StartsWith("Full Name|Identifier|Amount|Region", csv);
        Assert.Equal("Alice", imported[0]["Full Name"]);
        Assert.Equal("1", imported[0]["Identifier"]);
        Assert.Equal("125.50", imported[0]["Amount"]);
    }

    [Fact]
    public void _02_Export_object_collection_escapes_delimiters_quotes_and_new_lines()
    {
        var rows = new[]
        {
            new CsvExportRow(1, "Alice", "Contains, comma"),
            new CsvExportRow(2, "Bob \"The Builder\"", "Line 1\nLine 2")
        };

        var csv = CsvWriter.FromObjects(rows)
            .ColumnOrder("ID", "Name", "Notes")
            .ToText();

        CsvTestOutput.SaveCsv(nameof(_02_Export_object_collection_escapes_delimiters_quotes_and_new_lines), csv);

        var imported = CsvReader.FromText(csv).ReadDictionaries();

        Assert.Contains("\"Contains, comma\"", csv);
        Assert.Contains("\"Bob \"\"The Builder\"\"\"", csv);
        Assert.Equal("Line 1\nLine 2", imported[1]["Notes"]);
    }

    [Fact]
    public void _03_Export_dictionary_rows_can_omit_headers()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Region"] = "London", ["Income"] = 210000m, ["Currency"] = "GBP" },
            new Dictionary<string, object?> { ["Region"] = "United States", ["Income"] = 325000m, ["Currency"] = "USD" }
        };

        var csv = CsvWriter.FromRows(rows)
            .DelimitedBy(";")
            .IncludeHeaders(false)
            .ColumnOrder("Currency", "Region", "Income")
            .ToText();

        CsvTestOutput.SaveCsv(nameof(_03_Export_dictionary_rows_can_omit_headers), csv);

        var imported = CsvReader.FromText(csv)
            .DelimitedBy(";")
            .HasHeaderRow(false)
            .Columns(columns => columns.IncludeLetters("A", "B", "C"))
            .ReadDictionaries();

        Assert.StartsWith("GBP;London;210000", csv);
        Assert.Equal("USD", imported[1]["A"]);
        Assert.Equal("United States", imported[1]["B"]);
    }

    [Fact]
    public void _04_Export_tab_delimited_operational_status_with_aliases()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Team"] = "Risk", ["OpenItems"] = 12, ["ClosedItems"] = 88, ["Status"] = "Green", ["Comment"] = "Ready" },
            new Dictionary<string, object?> { ["Team"] = "Finance", ["OpenItems"] = 27, ["ClosedItems"] = 73, ["Status"] = "Amber", ["Comment"] = "Needs review" },
            new Dictionary<string, object?> { ["Team"] = "Technology", ["OpenItems"] = 43, ["ClosedItems"] = 57, ["Status"] = "Red", ["Comment"] = "=follow-up" }
        };

        var csv = CsvWriter.FromRows(rows)
            .DelimitedBy("\t")
            .NewLine("\n")
            .ColumnOrder("Team", "Status", "OpenItems", "ClosedItems", "Comment")
            .HeaderAlias("OpenItems", "Open Items")
            .HeaderAlias("ClosedItems", "Closed Items")
            .ToText();

        CsvTestOutput.SaveCsv(nameof(_04_Export_tab_delimited_operational_status_with_aliases), csv);

        var imported = CsvReader.FromText(csv)
            .DelimitedBy("\t")
            .ReadDictionaries();

        Assert.StartsWith("Team\tStatus\tOpen Items\tClosed Items\tComment", csv);
        Assert.Equal("'=follow-up", imported[2]["Comment"]);
    }

    [Fact]
    public void _05_Export_large_dictionary_sample_with_column_order()
    {
        var rows = Enumerable.Range(1, 250)
            .Select(index => new Dictionary<string, object?>
            {
                ["ID"] = index,
                ["Region"] = (index % 4) switch
                {
                    0 => "London",
                    1 => "United States",
                    2 => "Europe",
                    _ => "Asia"
                },
                ["Amount"] = index * 12.75m,
                ["Status"] = index % 2 == 0 ? "Closed" : "Open",
                ["Notes"] = $"Generated export row {index}"
            });

        var csv = CsvWriter.FromRows(rows)
            .NewLine("\n")
            .ColumnOrder("ID", "Region", "Status", "Amount", "Notes")
            .ToText();

        CsvTestOutput.SaveCsv(nameof(_05_Export_large_dictionary_sample_with_column_order), csv);

        var imported = CsvReader.FromText(csv)
            .Columns(columns => columns.IncludeHeaders("ID", "Region", "Amount"))
            .ReadDictionaries();

        Assert.Equal(250, imported.Count);
        Assert.Equal("1", imported[0]["ID"]);
        Assert.Equal("3187.50", imported[^1]["Amount"]);
    }

    [Fact]
    public void Export_dictionary_rows_matches_keys_case_insensitively()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Name"] = "Alice", ["Amount"] = 10 },
            new Dictionary<string, object?> { ["name"] = "Ben", ["amount"] = 20 }
        };

        var csv = CsvWriter.FromRows(rows)
            .NewLine("\n")
            .ToText();

        Assert.Equal("Name,Amount\nAlice,10\nBen,20", csv);
    }

    [Fact]
    public void Export_reports_progress_and_audit_events()
    {
        var progress = new TestProgress<DocWeaveOperationProgress>();
        var auditEvents = new List<DocWeaveAuditEvent>();
        var rows = Enumerable.Range(1, 12)
            .Select(index => new Dictionary<string, object?>
            {
                ["ID"] = index,
                ["Name"] = $"Name {index}",
                ["Amount"] = index * 10
            });

        var csv = CsvWriter.FromRows(rows)
            .NewLine("\n")
            .WithProgress(progress)
            .Audit(auditEvents.Add)
            .ToText();

        Assert.StartsWith("ID,Name,Amount", csv);
        Assert.Contains(progress.Reports, report => report.Operation == "CsvExport" && report.Stage == "Completed");
        Assert.Contains(auditEvents, audit => audit.Operation == "CsvExport" && audit.Message == "CSV written");
    }

    [Fact]
    public void Export_honors_pre_cancelled_token()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => CsvWriter
            .FromRows([new Dictionary<string, object?> { ["ID"] = 1 }])
            .WithCancellation(cancellation.Token)
            .ToText());
    }

    [Fact]
    public void Export_can_stream_to_text_writer()
    {
        var rows = Enumerable.Range(1, 5)
            .Select(index => new Dictionary<string, object?>
            {
                ["ID"] = index,
                ["Name"] = $"Name {index}"
            });
        using var writer = new StringWriter();

        CsvWriter.FromRows(rows)
            .NewLine("\n")
            .WriteTo(writer);

        Assert.Equal("ID,Name\n1,Name 1\n2,Name 2\n3,Name 3\n4,Name 4\n5,Name 5", writer.ToString());
    }

    private sealed record CsvExportRow(int ID, string Name, string Notes);
}
