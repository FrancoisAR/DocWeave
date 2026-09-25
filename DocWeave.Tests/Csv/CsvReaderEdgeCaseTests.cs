using System.Data;
using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

/// <summary>
/// Regression tests for parser and reader edge cases. Each test names the behaviour it protects.
/// </summary>
public sealed class CsvReaderEdgeCaseTests
{
    // ---- Delimiters ----

    [Fact]
    public void Multi_character_delimiter_partial_match_keeps_all_data()
    {
        // "ab" starts the delimiter "abc" but does not complete it, so it must stay in the value.
        var rows = CsvReader.FromText("H1abcH2\nabxabcy")
            .DelimitedBy("abc")
            .ReadDictionaries();

        Assert.Single(rows);
        Assert.Equal("abx", rows[0]["H1"]);
        Assert.Equal("y", rows[0]["H2"]);
    }

    [Fact]
    public void Multi_character_delimiter_splits_fields()
    {
        var rows = CsvReader.FromText("A||B\n1||2")
            .DelimitedBy("||")
            .ReadDictionaries();

        Assert.Equal("1", rows[0]["A"]);
        Assert.Equal("2", rows[0]["B"]);
    }

    // ---- Quotes ----

    [Fact]
    public void Quote_inside_unquoted_field_is_literal_text()
    {
        var rows = CsvReader.FromText("Size,Name\n5\" pipe,x")
            .ReadDictionaries();

        Assert.Equal("5\" pipe", rows[0]["Size"]);
        Assert.Equal("x", rows[0]["Name"]);
    }

    [Fact]
    public void Quoted_field_may_follow_spaces_after_the_delimiter()
    {
        var rows = CsvReader.FromText("A,B\n1, \"b,c\"")
            .ReadDictionaries();

        Assert.Equal("b,c", rows[0]["B"]);
    }

    [Fact]
    public void Quoted_field_can_contain_new_lines()
    {
        var rows = CsvReader.FromText("A,B\n1,\"line 1\nline 2\"\n2,x")
            .ReadDictionaries();

        Assert.Equal(2, rows.Count);
        Assert.Equal("line 1\nline 2", rows[0]["B"]);
        Assert.Equal("2", rows[1]["A"]);
    }

    [Fact]
    public void Unterminated_quote_throws_with_the_record_number()
    {
        var reader = CsvReader.FromText("ID,Name\n1,\"never closed");

        var exception = Assert.Throws<FormatException>(() => reader.ReadDictionaries());

        Assert.Contains("record 2", exception.Message);
    }

    [Fact]
    public void Quoted_empty_value_at_end_of_file_is_kept()
    {
        var single = CsvReader.FromText("ID\n\"\"").ReadDictionaries();
        var last = CsvReader.FromText("ID,Name\n1,\"\"").ReadDictionaries();

        Assert.Single(single);
        Assert.Equal(string.Empty, single[0]["ID"]);
        Assert.Single(last);
        Assert.Equal(string.Empty, last[0]["Name"]);
    }

    [Fact]
    public void Trailing_delimiter_at_end_of_file_gives_an_empty_last_field()
    {
        var rows = CsvReader.FromText("A,B\n1,").ReadDictionaries();

        Assert.Single(rows);
        Assert.Equal(string.Empty, rows[0]["B"]);
    }

    // ---- Blank lines ----

    [Fact]
    public void Blank_lines_are_skipped_by_default_and_row_numbers_count_them()
    {
        var rows = CsvReader.FromText("ID,Name\n1,A\n\n2,B\n\n")
            .ReadRows();

        Assert.Equal(2, rows.Count);
        Assert.Equal(2, rows[0].RowNumber);
        Assert.Equal(4, rows[1].RowNumber);
    }

    [Fact]
    public void Blank_lines_can_be_kept_as_rows()
    {
        var rows = CsvReader.FromText("ID,Name\n1,A\n\n2,B")
            .SkipEmptyLines(false)
            .ReadDictionaries();

        Assert.Equal(3, rows.Count);
        Assert.Equal(string.Empty, rows[1]["ID"]);
        Assert.Null(rows[1]["Name"]);
    }

    [Fact]
    public void Windows_line_endings_are_handled()
    {
        var rows = CsvReader.FromText("ID,Name\r\n1,A\r\n\r\n2,B\r\n")
            .ReadDictionaries();

        Assert.Equal(2, rows.Count);
        Assert.Equal("B", rows[1]["Name"]);
    }

    // ---- Stream ownership ----

    [Fact]
    public void FromStream_leaves_the_callers_stream_open()
    {
        using var stream = new MemoryStream("ID,Name\n1,A"u8.ToArray());

        var rows = CsvReader.FromStream(stream).ReadDictionaries();

        Assert.Single(rows);
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void FromFile_releases_the_file_after_reading()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docweave-{Guid.NewGuid():N}.csv");
        File.WriteAllText(path, "ID,Name\n1,A");

        try
        {
            CsvReader.FromFile(path).ReadDictionaries();

            // Deleting fails with an IOException if the reader still holds the file open.
            File.Delete(path);
            Assert.False(File.Exists(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void FromFile_releases_the_file_when_enumeration_stops_early()
    {
        var path = Path.Combine(Path.GetTempPath(), $"docweave-{Guid.NewGuid():N}.csv");
        File.WriteAllText(path, "ID,Name\n1,A\n2,B\n3,C");

        try
        {
            var first = CsvReader.FromFile(path).EnumerateRows().First();

            Assert.Equal("1", first.Values["ID"]);
            File.Delete(path);
            Assert.False(File.Exists(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    // ---- Lazy reading ----

    [Fact]
    public void EnumerateRows_reads_lazily_from_a_large_file()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "TestData", "Csv", "mock-invoices-50000.csv");
        using var stream = File.OpenRead(path);

        var first = CsvReader.FromStream(stream).EnumerateRows().First();

        Assert.Equal(2, first.RowNumber);
        // Only a buffer's worth of the file may have been read to produce the first row.
        Assert.True(stream.Position < stream.Length / 10, $"Read {stream.Position} of {stream.Length} bytes.");
    }

    [Fact]
    public void EnumerateRows_returns_the_same_rows_as_ReadRows()
    {
        const string csv = "ID,Name\n1,A\n\n2,B\n3,C";

        var enumerated = CsvReader.FromText(csv).EnumerateRows().ToList();
        var read = CsvReader.FromText(csv).ReadRows();

        Assert.Equal(read.Select(row => row.RowNumber), enumerated.Select(row => row.RowNumber));
        Assert.Equal(read.Select(row => row.Values["Name"]), enumerated.Select(row => row.Values["Name"]));
    }

    // ---- Headers ----

    [Fact]
    public void Duplicate_and_empty_headers_are_renamed_by_default()
    {
        var rows = CsvReader.FromText("Name,name,,Name\n1,2,3,4")
            .ReadDictionaries();

        Assert.Equal(["Name", "name_2", "Column3", "Name_3"], rows[0].Keys);
        Assert.Equal("1", rows[0]["Name"]);
        Assert.Equal("2", rows[0]["name_2"]);
        Assert.Equal("3", rows[0]["Column3"]);
        Assert.Equal("4", rows[0]["Name_3"]);
    }

    [Fact]
    public void Duplicate_headers_can_be_rejected()
    {
        var reader = CsvReader.FromText("Name,NAME\n1,2")
            .DuplicateHeaders(CsvDuplicateHeaderMode.Throw);

        var exception = Assert.Throws<FormatException>(() => reader.ReadDictionaries());

        Assert.Contains("NAME", exception.Message);
    }

    [Fact]
    public void Duplicate_headers_no_longer_break_ReadDataTable()
    {
        var table = CsvReader.FromText("Name,name\nA,B").ReadDataTable();

        Assert.Equal(["Name", "name_2"], table.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
        Assert.Equal("B", table.Rows[0]["name_2"]);
    }

    [Fact]
    public void Selecting_a_renamed_duplicate_header_works()
    {
        var rows = CsvReader.FromText("Name,Name\nA,B")
            .Columns(columns => columns.IncludeHeaders("Name_2"))
            .ReadDictionaries();

        Assert.Equal("B", rows[0]["Name_2"]);
    }

    // ---- Row length ----

    [Fact]
    public void Missing_values_are_null_and_extra_values_are_dropped_by_default()
    {
        var rows = CsvReader.FromText("A,B,C\n1,2\n1,2,3,4")
            .ReadDictionaries();

        Assert.Null(rows[0]["C"]);
        Assert.Equal(["A", "B", "C"], rows[1].Keys);
        Assert.Equal("3", rows[1]["C"]);
    }

    [Fact]
    public void Row_length_mismatch_can_throw_with_source_and_row_number()
    {
        var reader = CsvReader.FromText("A,B\n1,2\n1,2,3", label: "orders.csv")
            .OnRowLengthMismatch(CsvRowLengthMode.Throw);

        var exception = Assert.Throws<FormatException>(() => reader.ReadRows());

        Assert.Contains("orders.csv", exception.Message);
        Assert.Contains("row 3", exception.Message);
        Assert.Contains("3 fields", exception.Message);
    }

    // ---- Nulls in DataTable ----

    [Fact]
    public void ReadDataTable_keeps_missing_values_as_DBNull()
    {
        var table = CsvReader.FromText("ID,Name\n1").ReadDataTable();

        Assert.Equal("1", table.Rows[0]["ID"]);
        Assert.True(table.Rows[0].IsNull("Name"));
    }

    [Fact]
    public void ReadDataTable_keeps_empty_strings_as_empty_strings()
    {
        var table = CsvReader.FromText("ID,Name\n1,\"\"").ReadDataTable();

        Assert.False(table.Rows[0].IsNull("Name"));
        Assert.Equal(string.Empty, table.Rows[0]["Name"]);
    }

    // ---- Empty input ----

    [Fact]
    public void Empty_input_returns_no_rows()
    {
        var table = CsvReader.FromText(string.Empty).ReadDataTable();

        Assert.Empty(table.Rows);
        Assert.Empty(table.Columns);
    }

    [Fact]
    public void Headers_only_input_returns_columns_and_no_rows()
    {
        var table = CsvReader.FromText("ID,Name").ReadDataTable();

        Assert.Empty(table.Rows);
        Assert.Equal(2, table.Columns.Count);
    }

    [Fact]
    public void Files_without_headers_number_rows_from_one()
    {
        var rows = CsvReader.FromText("1,A\n2,B")
            .HasHeaderRow(false)
            .ReadRows();

        Assert.Equal([1, 2], rows.Select(row => row.RowNumber));
    }
}
