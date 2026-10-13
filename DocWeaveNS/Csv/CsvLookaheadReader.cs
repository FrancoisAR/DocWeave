namespace DocWeave.Csv;

/// <summary>
/// Wraps a TextReader so the parser can look ahead several characters without consuming them.
/// A multi-character delimiter can only be recognised safely if a partial match leaves the input untouched.
/// </summary>
internal sealed class CsvLookaheadReader
{
    private readonly TextReader reader;
    private readonly List<int> buffer = [];

    public CsvLookaheadReader(TextReader reader)
    {
        this.reader = reader;
    }

    /// <summary>
    /// Returns the character at the given offset from the current position, or -1 past the end of the input.
    /// </summary>
    public int Peek(int offset)
    {
        while (buffer.Count <= offset)
        {
            var next = reader.Read();
            if (next < 0)
            {
                return -1;
            }

            buffer.Add(next);
        }

        return buffer[offset];
    }

    public int Read()
    {
        var current = Peek(0);
        if (current >= 0)
        {
            buffer.RemoveAt(0);
        }

        return current;
    }

    public bool StartsWith(string value)
    {
        for (var i = 0; i < value.Length; i++)
        {
            if (Peek(i) != value[i])
            {
                return false;
            }
        }

        return true;
    }

    public void Skip(int count)
    {
        for (var i = 0; i < count; i++)
        {
            Read();
        }
    }
}
