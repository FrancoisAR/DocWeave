using DocWeave.Csv;

namespace DocWeave.Tests.Csv;

/// <summary>
/// CSV files are opened in spreadsheets, where a leading = + - @ makes text run as a formula.
/// </summary>
public sealed class CsvWriterSafetyTests
{
    [Fact]
    public void Formula_like_strings_are_prefixed_but_numbers_are_not()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Text"] = "=SUM(A1)", ["Amount"] = -5m },
            new Dictionary<string, object?> { ["Text"] = "+1", ["Amount"] = 7.5m },
            new Dictionary<string, object?> { ["Text"] = "-x", ["Amount"] = 0m },
            new Dictionary<string, object?> { ["Text"] = "@cmd", ["Amount"] = 1m },
            new Dictionary<string, object?> { ["Text"] = "plain", ["Amount"] = 2m }
        };

        var lines = CsvWriter.FromRows(rows).NewLine("\n").ToText().Split('\n');

        Assert.Equal("Text,Amount", lines[0]);
        Assert.Equal("'=SUM(A1),-5", lines[1]);
        Assert.Equal("'+1,7.5", lines[2]);
        Assert.Equal("'-x,0", lines[3]);
        Assert.Equal("'@cmd,1", lines[4]);
        Assert.Equal("plain,2", lines[5]);
    }

    [Fact]
    public void Tab_and_carriage_return_prefixes_are_neutralized_and_still_quoted_when_needed()
    {
        var rows = new[]
        {
            new Dictionary<string, object?> { ["Text"] = "\t=1" },
            new Dictionary<string, object?> { ["Text"] = "\r\n=1" }
        };

        var csv = CsvWriter.FromRows(rows).NewLine("\n").ToText();

        Assert.Contains("'\t=1", csv);
        Assert.Contains("\"'\r\n=1\"", csv);
    }

    [Fact]
    public void Line_feed_prefix_is_neutralized_and_quoted()
    {
        var rows = new[] { new Dictionary<string, object?> { ["Text"] = "\n=1" } };

        var csv = CsvWriter.FromRows(rows).NewLine("\n").ToText();

        Assert.Contains("\"'\n=1\"", csv);
    }

    [Fact]
    public void Formula_like_headers_are_neutralized()
    {
        var rows = new[] { new Dictionary<string, object?> { ["=bad"] = "x" } };

        var csv = CsvWriter.FromRows(rows).NewLine("\n").ToText();

        Assert.StartsWith("'=bad\n", csv);
    }

    [Fact]
    public void Header_aliases_are_neutralized()
    {
        var rows = new[] { new Dictionary<string, object?> { ["Name"] = "x" } };

        var csv = CsvWriter.FromRows(rows).HeaderAlias("Name", "@alias").NewLine("\n").ToText();

        Assert.StartsWith("'@alias\n", csv);
    }

    [Fact]
    public void Formula_text_can_be_written_as_is_when_explicitly_allowed()
    {
        var rows = new[] { new Dictionary<string, object?> { ["Text"] = "=SUM(A1)" } };

        var csv = CsvWriter.FromRows(rows).AllowFormulaText().NewLine("\n").ToText();

        Assert.Equal("Text\n=SUM(A1)", csv);
    }

    [Fact]
    public void Negative_numbers_from_typed_columns_are_never_changed()
    {
        var rows = new[] { new { Delta = -12.5m, Count = -3 } };

        var csv = CsvWriter.FromObjects(rows).NewLine("\n").ToText();

        Assert.Equal("Delta,Count\n-12.5,-3", csv);
    }
}
