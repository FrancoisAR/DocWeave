using System.Data;

namespace DocWeave.Csv;

/// <summary>
/// Entry point for creating CSV exports from common structured data sources.
/// </summary>
public static class CsvWriter
{
    /// <summary>
    /// Starts a CSV export from a DataTable.
    /// </summary>
    public static CsvWriteBuilder FromDataTable(DataTable dataTable)
    {
        NetStandardGuard.ThrowIfNull(dataTable, nameof(dataTable));
        return new CsvWriteBuilder(TabularData.FromDataTable(dataTable));
    }

    /// <summary>
    /// Starts a CSV export from an object collection, using public instance properties as columns.
    /// </summary>
    public static CsvWriteBuilder FromObjects<T>(IEnumerable<T> items)
    {
        NetStandardGuard.ThrowIfNull(items, nameof(items));
        return new CsvWriteBuilder(TabularData.FromObjects(items));
    }

    /// <summary>
    /// Starts a CSV export from dictionary rows. Dictionary keys become column names.
    /// </summary>
    public static CsvWriteBuilder FromRows(IEnumerable<IReadOnlyDictionary<string, object?>> rows)
    {
        NetStandardGuard.ThrowIfNull(rows, nameof(rows));
        return new CsvWriteBuilder(TabularData.FromDictionaries(rows));
    }
}
