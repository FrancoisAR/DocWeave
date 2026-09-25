using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// Checks the data sources a caller bound against the sources (and their columns) the template declares.
/// </summary>
internal static class ExcelTemplateDataSources
{
    public static void Validate(JsonElement root, IReadOnlyDictionary<string, TabularData> sources)
    {
        if (!root.TryGetProperty("dataSources", out var declarations))
        {
            return;
        }

        var declaredNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in declarations.EnumerateArray())
        {
            var name = declaration.GetRequiredString("name");
            if (!declaredNames.Add(name))
            {
                throw new InvalidOperationException($"Template source '{name}' is declared more than once.");
            }

            if (!sources.TryGetValue(name, out var source))
            {
                // Declared sources and columns are required unless the template says otherwise, and values may be null unless it sets "nullable": false.
                if (declaration.GetOptionalBool("required") ?? true)
                {
                    throw new InvalidOperationException($"Required template source '{name}' has not been bound.");
                }

                continue;
            }

            if (!declaration.TryGetProperty("columns", out var columns))
            {
                continue;
            }

            foreach (var column in columns.EnumerateArray())
            {
                ValidateColumn(name, source, column);
            }
        }
    }

    private static void ValidateColumn(string sourceName, TabularData source, JsonElement column)
    {
        var columnName = column.GetRequiredString("name");
        var columnIndex = source.IndexOf(columnName);
        if ((column.GetOptionalBool("required") ?? true) && columnIndex < 0)
        {
            throw new InvalidOperationException($"Required field '{columnName}' was not found in source '{sourceName}'.");
        }

        if (columnIndex < 0)
        {
            return;
        }

        var nullable = column.GetOptionalBool("nullable") ?? true;
        var type = column.GetOptionalString("type");
        for (var rowIndex = 0; rowIndex < source.Rows.Count; rowIndex++)
        {
            var value = columnIndex < source.Rows[rowIndex].Count ? source.Rows[rowIndex][columnIndex] : null;
            if (value is null)
            {
                if (!nullable)
                {
                    throw new InvalidOperationException($"Source '{sourceName}' field '{columnName}' row {rowIndex + 1} cannot be null.");
                }

                continue;
            }

            if (!string.IsNullOrWhiteSpace(type) && !IsCompatibleType(value, type))
            {
                throw new InvalidOperationException($"Source '{sourceName}' field '{columnName}' row {rowIndex + 1} must be {type}.");
            }
        }
    }

    private static bool IsCompatibleType(object value, string type)
    {
        return type.ToLowerInvariant() switch
        {
            "string" or "text" => value is string,
            "int" or "integer" => value is byte or sbyte or short or ushort or int or uint or long or ulong,
            "decimal" or "number" => value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal,
            "date" or "datetime" => value is DateTime or DateOnly,
            "bool" or "boolean" => value is bool,
            _ => throw new InvalidOperationException($"Template data type '{type}' is not supported.")
        };
    }
}
