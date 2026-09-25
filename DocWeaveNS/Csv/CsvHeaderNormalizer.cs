namespace DocWeave.Csv;

/// <summary>
/// Makes header names usable as dictionary keys and DataTable column names.
/// Both compare names case-insensitively, so uniqueness is checked the same way.
/// </summary>
internal static class CsvHeaderNormalizer
{
    public static List<string> Normalize(IReadOnlyList<string> headers, CsvDuplicateHeaderMode mode)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>(headers.Count);

        for (var i = 0; i < headers.Count; i++)
        {
            var name = string.IsNullOrWhiteSpace(headers[i]) ? $"Column{i + 1}" : headers[i];

            if (!seen.Add(name))
            {
                if (mode == CsvDuplicateHeaderMode.Throw)
                {
                    throw new FormatException($"CSV header '{name}' appears more than once (column {ColumnLetters.FromNumber(i + 1)}).");
                }

                // Find the first free name_2, name_3, ... Names are compared ignoring case, like the dictionaries and
                // DataTable columns that will hold them.
                var suffix = 2;
                string candidate;
                do
                {
                    candidate = $"{name}_{suffix++}";
                }
                while (!seen.Add(candidate));

                name = candidate;
            }

            result.Add(name);
        }

        return result;
    }
}
