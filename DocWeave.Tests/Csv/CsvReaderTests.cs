using DocWeave;
using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

public sealed class CsvReaderTests
{
    [Fact]
    public void ReadDictionaries_can_select_columns_by_letters()
    {
        var rows = CsvReader.FromText("""
            ID,Name,Region,Amount,Status
            1,Alice,London,100.50,Open
            2,Bob,Tokyo,250.00,Closed
            """)
            .Columns(columns => columns.IncludeLetters("A", "B", "E"))
            .ReadDictionaries();

        Assert.Equal(2, rows.Count);
        Assert.Equal(["ID", "Name", "Status"], rows[0].Keys);
        Assert.Equal("1", rows[0]["ID"]);
        Assert.Equal("Alice", rows[0]["Name"]);
        Assert.Equal("Open", rows[0]["Status"]);
    }

    [Fact]
    public void ReadDictionaries_can_select_columns_by_headers()
    {
        var rows = CsvReader.FromText("""
            ID,Name,Region,Amount,Status
            1,Alice,London,100.50,Open
            2,Bob,Tokyo,250.00,Closed
            """)
            .Columns(columns => columns.IncludeHeaders("Name", "Amount"))
            .ReadDictionaries();

        Assert.Equal(2, rows.Count);
        Assert.Equal(["Name", "Amount"], rows[0].Keys);
        Assert.Equal("Alice", rows[0]["Name"]);
        Assert.Equal("100.50", rows[0]["Amount"]);
    }

    [Fact]
    public void ReadDictionaries_preserves_requested_column_order()
    {
        var rows = CsvReader.FromText("""
            ID,Name,Region,Amount,Status
            1,Alice,London,100.50,Open
            """)
            .Columns(columns => columns.IncludeHeaders("Status", "ID", "Name"))
            .ReadDictionaries();

        Assert.Equal(["Status", "ID", "Name"], rows[0].Keys);
        Assert.Equal("Open", rows[0]["Status"]);
    }

    [Fact]
    public void ReadDataTable_loads_large_fixture()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", "mock-invoices-1000.csv");

        var table = CsvReader.FromFile(path)
            .Columns(columns => columns.IncludeHeaders("InvoiceNo", "CustomerName", "Amount"))
            .ReadDataTable();

        Assert.Equal(1000, table.Rows.Count);
        Assert.Equal(3, table.Columns.Count);
    }

    [Fact]
    public void ReadDictionaries_can_exclude_columns_by_letters_and_headers()
    {
        var rows = CsvReader.FromText("""
            ID,Name,Region,Amount,Status
            1,Alice,London,100.50,Open
            """)
            .Columns(columns => columns
                .ExcludeLetters("C")
                .ExcludeHeaders("Status"))
            .ReadDictionaries();

        Assert.Equal(["ID", "Name", "Amount"], rows[0].Keys);
    }

    [Fact]
    public void ReadDictionaries_supports_pipe_delimited_files()
    {
        var rows = CsvReader.FromText("""
            ID|Name|Region|Amount|Status
            1|Alice|London|100.50|Open
            """)
            .DelimitedBy("|")
            .Columns(columns => columns.IncludeHeaders("ID", "Region", "Amount"))
            .ReadDictionaries();

        Assert.Equal("London", rows[0]["Region"]);
        Assert.Equal("100.50", rows[0]["Amount"]);
    }

    [Fact]
    public void ReadDictionaries_supports_files_without_headers()
    {
        var rows = CsvReader.FromText("""
            1,Alice,London,100.50,Open
            2,Bob,Tokyo,250.00,Closed
            """)
            .HasHeaderRow(false)
            .Columns(columns => columns.IncludeLetters("A", "C", "E"))
            .ReadDictionaries();

        Assert.Equal(["A", "C", "E"], rows[0].Keys);
        Assert.Equal("1", rows[0]["A"]);
        Assert.Equal("London", rows[0]["C"]);
        Assert.Equal("Open", rows[0]["E"]);
    }

    [Fact]
    public void ReadDictionaries_handles_quoted_commas_and_quotes()
    {
        var csv = "ID,Name,Notes" + Environment.NewLine +
                  "1,Alice,\"Contains, comma and quoted \"\"value\"\"\"";

        var rows = CsvReader.FromText(csv)
            .ReadDictionaries();

        Assert.Equal("Contains, comma and quoted \"value\"", rows[0]["Notes"]);
    }

    [Fact]
    public void ReadDictionaryResult_returns_headers_rows_progress_and_audit_events()
    {
        var progress = new TestProgress<DocWeaveOperationProgress>();
        var auditEvents = new List<DocWeaveAuditEvent>();

        var result = CsvReader.FromText("""
            ID,Name,Region
            1,Alice,London
            2,Bob,United States
            """)
            .WithProgress(progress)
            .Audit(auditEvents.Add)
            .ReadDictionaryResult();

        Assert.Equal(["ID", "Name", "Region"], result.Headers);
        Assert.Equal(2, result.RowCount);
        Assert.Empty(result.Warnings);
        Assert.Equal("Bob", result.Rows[1]["Name"]);
        Assert.Contains(progress.Reports, report => report.Operation == "CsvImport" && report.Stage == "Completed");
        Assert.Contains(auditEvents, audit => audit.Operation == "CsvImport" && audit.Message == "CSV imported");
    }

    [Fact]
    public void ReadResult_honors_pre_cancelled_token()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => CsvReader
            .FromText("ID,Name\n1,Alice")
            .WithCancellation(cancellation.Token)
            .ReadResult());
    }
}
