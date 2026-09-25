using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// Reads the named styles a template declares, from "styleCatalog" and then "styles".
/// </summary>
internal static class ExcelTemplateStyles
{
    public static IReadOnlyDictionary<string, ExcelStyleSpec> Read(JsonElement root)
    {
        var styles = new Dictionary<string, ExcelStyleSpec>(StringComparer.OrdinalIgnoreCase);
        if (root.TryGetProperty("styleCatalog", out var catalogObject))
        {
            foreach (var style in catalogObject.EnumerateObject())
            {
                styles[style.Name] = ReadStyle(style.Value);
            }
        }

        if (!root.TryGetProperty("styles", out var styleObject))
        {
            return styles;
        }

        foreach (var style in styleObject.EnumerateObject())
        {
            styles[style.Name] = ReadStyle(style.Value);
        }

        return styles;
    }

    private static ExcelStyleSpec ReadStyle(JsonElement element)
    {
        var builder = new ExcelStyleBuilder();

        if (element.TryGetProperty("font", out var font))
        {
            if (font.GetOptionalBool("bold") == true)
            {
                builder.Bold();
            }

            if (font.GetOptionalBool("italic") == true)
            {
                builder.Italic();
            }

            if (font.GetOptionalBool("underline") == true)
            {
                builder.Underline();
            }

            var fontName = font.GetOptionalString("name");
            if (!string.IsNullOrWhiteSpace(fontName))
            {
                builder.FontFamily(fontName);
            }

            if (font.TryGetProperty("size", out var size))
            {
                builder.FontSize(size.GetDouble());
            }

            var fontColor = font.GetOptionalString("color");
            if (!string.IsNullOrWhiteSpace(fontColor))
            {
                builder.FontColor(fontColor);
            }
        }

        if (element.TryGetProperty("fill", out var fill))
        {
            var fillColor = fill.GetOptionalString("color");
            if (!string.IsNullOrWhiteSpace(fillColor))
            {
                builder.BackgroundColor(fillColor);
            }
        }

        if (element.TryGetProperty("border", out var border))
        {
            var style = ParseBorderStyle(border.GetOptionalString("style") ?? "thin");
            var color = border.GetOptionalString("color") ?? "#000000";
            builder.Border(style, color);
        }

        if (element.TryGetProperty("alignment", out var alignment))
        {
            if (alignment.TryGetProperty("wrapText", out var wrapText))
            {
                builder.WrapText(wrapText.GetBoolean());
            }

            if (alignment.TryGetProperty("rotation", out var rotation))
            {
                builder.RotateText(rotation.GetInt32());
            }
        }

        var numberFormat = element.GetOptionalString("numberFormat");
        if (!string.IsNullOrWhiteSpace(numberFormat))
        {
            builder.NumberFormat(ParseNumberFormat(numberFormat));
        }

        return builder.Spec;
    }

    private static ExcelBorderLineStyle ParseBorderStyle(string value)
    {
        return NetStandardEnum.Parse<ExcelBorderLineStyle>(value, ignoreCase: true);
    }

    private static ExcelNumberFormat ParseNumberFormat(string value)
    {
        return NetStandardEnum.Parse<ExcelNumberFormat>(value, ignoreCase: true);
    }
}
