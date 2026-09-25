using DocWeave.Excel;

namespace DocWeave.Tests.Excel;

public sealed class ExcelStreamingImportTests
{
    [Fact]
    public void _72_Enumerate_rows_returns_the_same_rows_as_read_rows()
    {
        var bytes = ExcelTestData.Load("stream-invoices-25-rows.xlsx");

        var eager = ExcelReader.FromBytes(bytes).ReadRows();
        var streamed = ExcelReader.FromBytes(bytes).EnumerateRows().ToList();

        Assert.Equal(25, streamed.Count);
        Assert.Equal(eager.Count, streamed.Count);
        for (var i = 0; i < eager.Count; i++)
        {
            Assert.Equal(eager[i].RowNumber, streamed[i].RowNumber);
            Assert.Equal(eager[i].Values.OrderBy(pair => pair.Key), streamed[i].Values.OrderBy(pair => pair.Key));
        }
    }

    [Fact]
    public void _73_Enumerate_rows_starts_at_the_start_cell_and_skips_the_preamble_above_it()
    {
        var bytes = ExcelTestData.Load("stream-invoices-with-preamble.xlsx");

        var rows = ExcelReader.FromBytes(bytes)
            .Sheet("Invoices")
            .StartAt("B5")
            .EnumerateRows()
            .ToList();

        // The header is row 5, so the first data row is row 6.
        Assert.Equal([6, 7, 8], rows.Select(row => row.RowNumber));
        Assert.Equal(["Customer", "InvoiceNo", "Net"], rows[0].Values.Keys.Order());
        Assert.Equal("Customer 1", rows[0].Values["Customer"]);
    }

    [Fact]
    public void _74_Enumerate_rows_keeps_a_cell_right_of_the_header_as_column_n_when_no_columns_are_selected()
    {
        // The header row has 3 columns (A to C). Row 2 also has a value in D, which has no header.
        var bytes = ExcelTestData.Load("stream-cell-right-of-header.xlsx");

        var rows = ExcelReader.FromBytes(bytes).EnumerateRows().ToList();

        Assert.Equal("reviewed", rows[0].Values["Column4"]);
        Assert.False(rows[1].Values.ContainsKey("Column4"));
    }

    [Fact]
    public void _75_Enumerate_rows_leaves_out_a_cell_right_of_the_header_when_columns_are_selected()
    {
        var bytes = ExcelTestData.Load("stream-cell-right-of-header.xlsx");

        var rows = ExcelReader.FromBytes(bytes)
            .Columns(columns => columns.IncludeHeaders("InvoiceNo", "Customer"))
            .EnumerateRows()
            .ToList();

        Assert.Equal(["Customer", "InvoiceNo"], rows[0].Values.Keys.Order());
    }

    [Fact]
    public void _76_Read_rows_still_gives_a_wider_data_row_a_column_of_its_own()
    {
        // The eager read sees the whole sheet first, so its header row is as wide as the widest row. Streaming cannot, and
        // this test pins that the eager behaviour did not change.
        var bytes = ExcelTestData.Load("stream-cell-right-of-header.xlsx");

        var result = ExcelReader.FromBytes(bytes).ReadResult();

        Assert.Equal(["InvoiceNo", "Customer", "Net", "Column4"], result.Headers);
    }

    [Fact]
    public void _77_Enumerate_rows_without_a_header_row_names_columns_by_letter()
    {
        var bytes = ExcelTestData.Load("stream-invoices-2-rows.xlsx");

        var rows = ExcelReader.FromBytes(bytes).HasHeaderRow(false).EnumerateRows().ToList();

        // Row 1 is data now, so the header text comes back as the first row.
        Assert.Equal(3, rows.Count);
        Assert.Equal("InvoiceNo", rows[0].Values["A"]);
        Assert.Equal("Customer 2", rows[2].Values["B"]);
    }

    [Fact]
    public void _78_Enumerate_rows_of_a_sheet_with_only_a_header_returns_nothing()
    {
        var bytes = ExcelTestData.Load("stream-header-only.xlsx");

        Assert.Empty(ExcelReader.FromBytes(bytes).EnumerateRows());
    }

    [Fact]
    public void _79_Enumerate_rows_closes_the_workbook_when_the_caller_stops_early()
    {
        var bytes = ExcelTestData.Load("stream-invoices-50-rows.xlsx");
        var stream = new MemoryStream(bytes);

        using (var enumerator = ExcelReader.FromStream(stream).EnumerateRows().GetEnumerator())
        {
            Assert.True(enumerator.MoveNext());
        }

        Assert.False(stream.CanRead);
    }

    [Fact]
    public void _80_Enumerate_rows_stops_when_cancelled_part_way_through()
    {
        var bytes = ExcelTestData.Load("stream-invoices-100-rows.xlsx");
        using var cancellation = new CancellationTokenSource();
        var rows = ExcelReader.FromBytes(bytes).WithCancellation(cancellation.Token).EnumerateRows();

        var read = 0;
        Assert.Throws<OperationCanceledException>(() =>
        {
            foreach (var row in rows)
            {
                read++;
                if (read == 10)
                {
                    cancellation.Cancel();
                }
            }
        });

        // The token is checked for each worksheet row, so reading stops within a row or two of the cancel.
        Assert.InRange(read, 10, 12);
    }

    [Fact]
    public void _81_Import_reads_text_stored_in_the_shared_string_table_by_index()
    {
        // Excel itself stores text in a shared string table. DocWeave's own exports use inline strings, so this workbook is
        // written by hand to cover that path.
        var bytes = ExcelTestData.Load("stream-shared-string-table.xlsx");

        var streamed = ExcelReader.FromBytes(bytes).HasHeaderRow(false).EnumerateRows().Select(row => row.Values["A"]);
        var eager = ExcelReader.FromBytes(bytes).HasHeaderRow(false).ReadRows().Select(row => row.Values["A"]);

        Assert.Equal(["London", "Europe", "United States", "Europe"], streamed);
        Assert.Equal(["London", "Europe", "United States", "Europe"], eager);
    }

    [Fact]
    public void _82_Import_reads_a_workbook_with_a_sparse_sheet_and_gaps_between_rows()
    {
        // Excel leaves empty rows and cells out of the file. Cells must land in the column their reference says.
        var bytes = ExcelTestData.Load("stream-sparse-sheet.xlsx");

        var rows = ExcelReader.FromBytes(bytes).EnumerateRows().ToList();

        Assert.Equal([2, 6], rows.Select(row => row.RowNumber));
        Assert.Equal("1500", rows[0].Values["Income"]);
        Assert.Equal("Asia", rows[1].Values["Region"]);
        Assert.Equal(rows.Select(row => row.Values["Income"]), ExcelReader.FromBytes(bytes).ReadRows().Select(row => row.Values["Income"]));
    }
}
