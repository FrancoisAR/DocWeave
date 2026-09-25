using System.Data;
using System.Reflection;

namespace DocWeave;

/// <summary>
/// Rows and columns read from any supported source (DataTable, objects, dictionaries), shared by the
/// Excel and CSV writers. Each row holds one value per column, in the same order as <see cref="Columns"/>.
/// </summary>
internal sealed record TabularData(IReadOnlyList<string> Columns, IReadOnlyList<IReadOnlyList<object?>> Rows)
{
    /// <summary>
    /// Finds a column by name, ignoring case. Returns -1 when the column does not exist.
    /// </summary>
    public int IndexOf(string column)
    {
        for (var i = 0; i < Columns.Count; i++)
        {
            if (string.Equals(Columns[i], column, StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    public static TabularData FromDataTable(DataTable dataTable)
    {
        // A database null (DBNull) becomes a plain null, so the writers only have one kind of "no value" to handle.
        var columns = dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList();
        var rows = dataTable.Rows
            .Cast<DataRow>()
            .Select(row => columns.Select(column => row[column] is DBNull ? null : row[column]).ToList())
            .Cast<IReadOnlyList<object?>>()
            .ToList();

        return new TabularData(columns, rows);
    }

    /// <summary>
    /// Uses the public readable instance properties of the item type as columns.
    /// For a collection typed as object, the type of the first non-null item is used.
    /// </summary>
    public static TabularData FromObjects<T>(IEnumerable<T> items)
    {
        var materialized = items.Cast<object?>().ToList();
        var itemType = typeof(T) == typeof(object)
            ? materialized.FirstOrDefault(item => item is not null)?.GetType()
            : typeof(T);

        if (itemType is null)
        {
            return new TabularData([], []);
        }

        var properties = itemType
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.GetMethod is not null && property.GetIndexParameters().Length == 0)
            .ToList();

        var columns = properties.Select(property => property.Name).ToList();
        // A null item gives a row of nulls, so it keeps its position among the rows.
        var rows = materialized
            .Select(item => properties.Select(property => item is null ? null : property.GetValue(item)).ToList())
            .Cast<IReadOnlyList<object?>>()
            .ToList();

        return new TabularData(columns, rows);
    }

    /// <summary>
    /// Uses the union of all dictionary keys as columns. Keys are matched ignoring case, so a row that
    /// spells a key differently from the first row still lands in the same column.
    /// </summary>
    public static TabularData FromDictionaries(IEnumerable<IReadOnlyDictionary<string, object?>> rows)
    {
        var materialized = rows.ToList();
        var columns = materialized
            .SelectMany(row => row.Keys)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var outputRows = materialized
            .Select(row => columns.Select(column => TryGetValueIgnoreCase(row, column, out var value) ? value : null).ToList())
            .Cast<IReadOnlyList<object?>>()
            .ToList();

        return new TabularData(columns, outputRows);
    }

    private static bool TryGetValueIgnoreCase(IReadOnlyDictionary<string, object?> row, string key, out object? value)
    {
        if (row.TryGetValue(key, out value))
        {
            return true;
        }

        foreach (var item in row)
        {
            if (string.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase))
            {
                value = item.Value;
                return true;
            }
        }

        value = null;
        return false;
    }
}
