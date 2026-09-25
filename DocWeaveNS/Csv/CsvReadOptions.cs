namespace DocWeave.Csv;

/// <summary>
/// Controls what happens when two header cells resolve to the same name (compared case-insensitively).
/// </summary>
public enum CsvDuplicateHeaderMode
{
    /// <summary>
    /// Keep the first header as is and rename later ones to Name_2, Name_3, and so on.
    /// </summary>
    Rename,

    /// <summary>
    /// Fail with a FormatException that names the duplicate header.
    /// </summary>
    Throw
}

/// <summary>
/// Controls what happens when a data row has a different number of fields than the header row.
/// </summary>
public enum CsvRowLengthMode
{
    /// <summary>
    /// Missing values become null and extra values are dropped.
    /// </summary>
    Ignore,

    /// <summary>
    /// Fail with a FormatException that names the source and row number.
    /// </summary>
    Throw
}
