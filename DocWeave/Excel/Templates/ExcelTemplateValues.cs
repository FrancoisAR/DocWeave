using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// Helpers for reading values and data source columns out of a template.
/// </summary>
internal static class ExcelTemplateValues
{
    public static object? ReadValue(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value))
        {
            return null;
        }

        return ReadJsonValue(value);
    }

    public static object? ReadJsonValue(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.Null => null,
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Number when value.TryGetInt64(out var integer) => integer,
            JsonValueKind.Number => value.GetDecimal(),
            _ => value.GetString()
        };
    }

    /// <summary>
    /// Finds a column of a bound data source by name, or fails with a message naming the field.
    /// </summary>
    public static int RequireColumn(TabularData source, string name)
    {
        var index = source.IndexOf(name);
        if (index >= 0)
        {
            return index;
        }

        throw new InvalidOperationException($"Source field '{name}' was not found.");
    }
}
