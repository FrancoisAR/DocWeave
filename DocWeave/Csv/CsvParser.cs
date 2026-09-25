using System.Text;

namespace DocWeave.Csv;

/// <summary>
/// One parsed CSV record. The number is the one-based position of the record in the source,
/// counting blank lines that were skipped, so it matches what a person sees in a text editor.
/// </summary>
internal readonly record struct CsvParsedRecord(int Number, IReadOnlyList<string> Fields);

/// <summary>
/// Streaming RFC 4180 style parser. Records are produced lazily so large files are never held in memory here.
/// </summary>
internal sealed class CsvParser
{
    private readonly string delimiter;
    private readonly char quote;
    private readonly bool trimFields;
    private readonly bool skipEmptyLines;

    public CsvParser(string delimiter, char quote, bool trimFields, bool skipEmptyLines)
    {
        this.delimiter = delimiter;
        this.quote = quote;
        this.trimFields = trimFields;
        this.skipEmptyLines = skipEmptyLines;
    }

    public IEnumerable<CsvParsedRecord> Parse(TextReader reader)
    {
        var input = new CsvLookaheadReader(reader);
        var fields = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        var fieldWasQuoted = false;
        var recordHasQuotedField = false;
        var recordNumber = 0;

        while (true)
        {
            var current = input.Peek(0);
            if (current < 0)
            {
                break;
            }

            if (inQuotes)
            {
                input.Read();
                if (current == quote)
                {
                    // Inside a quoted field, two quotes in a row stand for one literal quote. A single quote closes the field.
                    if (input.Peek(0) == quote)
                    {
                        input.Read();
                        field.Append(quote);
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    field.Append((char)current);
                }

                continue;
            }

            // A quote only opens a quoted field at the start of the field (ignoring leading spaces).
            // A quote in the middle of an unquoted value, such as 5" pipe, is ordinary text.
            if (current == quote && !fieldWasQuoted && IsBlank(field))
            {
                input.Read();
                field.Clear();
                inQuotes = true;
                fieldWasQuoted = true;
                recordHasQuotedField = true;
                continue;
            }

            // The delimiter is only consumed after it matches in full, so a partial match loses nothing.
            if (input.StartsWith(delimiter))
            {
                input.Skip(delimiter.Length);
                fields.Add(FinishField(field));
                fieldWasQuoted = false;
                continue;
            }

            if (current is '\r' or '\n')
            {
                input.Read();
                if (current == '\r' && input.Peek(0) == '\n')
                {
                    input.Read();
                }

                recordNumber++;

                // A blank line has no delimiters and no quoted empty value. A quoted "" is real data.
                var isBlankLine = fields.Count == 0 && field.Length == 0 && !recordHasQuotedField;
                if (isBlankLine)
                {
                    if (!skipEmptyLines)
                    {
                        yield return new CsvParsedRecord(recordNumber, [string.Empty]);
                    }

                    continue;
                }

                fields.Add(FinishField(field));
                yield return new CsvParsedRecord(recordNumber, fields);
                fields = [];
                fieldWasQuoted = false;
                recordHasQuotedField = false;
                continue;
            }

            input.Read();
            field.Append((char)current);
        }

        if (inQuotes)
        {
            throw new FormatException($"CSV record {recordNumber + 1} has a quoted field that is never closed.");
        }

        // The final record may have no trailing newline. A quoted "" still counts as a value.
        if (fields.Count > 0 || field.Length > 0 || recordHasQuotedField)
        {
            recordNumber++;
            fields.Add(FinishField(field));
            yield return new CsvParsedRecord(recordNumber, fields);
        }
    }

    private static bool IsBlank(StringBuilder value)
    {
        for (var i = 0; i < value.Length; i++)
        {
            if (!char.IsWhiteSpace(value[i]))
            {
                return false;
            }
        }

        return true;
    }

    private string FinishField(StringBuilder field)
    {
        var value = field.ToString();
        field.Clear();
        return trimFields ? value.Trim() : value;
    }
}
