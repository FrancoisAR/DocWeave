using System.Text.Json;

namespace DocWeave;

/// <summary>
/// Readers for properties of a template's JSON, with clear errors for missing required values.
/// Shared by the Excel and CSV template writers.
/// </summary>
internal static class TemplateJsonExtensions
{
    public static string GetRequiredString(this JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && !string.IsNullOrWhiteSpace(property.GetString())
            ? property.GetString()!
            : throw new InvalidOperationException($"Required template property '{propertyName}' was not found.");
    }

    public static string? GetOptionalString(this JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) ? property.GetString() : null;
    }

    public static bool? GetOptionalBool(this JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) ? property.GetBoolean() : null;
    }

    public static int? GetOptionalInt32(this JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) ? property.GetInt32() : null;
    }

    public static int GetRequiredInt32(this JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property)
            ? property.GetInt32()
            : throw new InvalidOperationException($"Required template property '{propertyName}' was not found.");
    }

    public static decimal GetRequiredDecimal(this JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property)
            ? property.GetDecimal()
            : throw new InvalidOperationException($"Required template property '{propertyName}' was not found.");
    }
}
