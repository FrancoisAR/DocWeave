using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocWeave.Excel;

internal sealed record ExcelStyleSpec(
    bool Bold,
    bool Italic,
    bool Underline,
    double? FontSize,
    string? FontFamily,
    string? FontColor,
    string? BackgroundColor,
    ExcelFillPattern FillPattern,
    bool? WrapText,
    int? TextRotation,
    ExcelNumberFormat NumberFormat,
    ExcelBorderLineStyle BorderStyle,
    string? BorderColor)
{
    public static ExcelStyleSpec Empty { get; } = new(false, false, false, null, null, null, null, ExcelFillPattern.Solid, null, null, ExcelNumberFormat.General, ExcelBorderLineStyle.None, null);
}

internal sealed class ExcelStyleRegistry
{
    // Identical styles share one index. Index 0 is always the empty style, so unstyled cells need no style entry.
    private readonly Dictionary<ExcelStyleSpec, uint> styleIndexes = new() { [ExcelStyleSpec.Empty] = 0 };
    private readonly List<ExcelStyleSpec> styles = [ExcelStyleSpec.Empty];

    public uint GetStyleIndex(ExcelStyleSpec style)
    {
        if (styleIndexes.TryGetValue(style, out var index))
        {
            return index;
        }

        index = (uint)styles.Count;
        styles.Add(style);
        styleIndexes[style] = index;
        return index;
    }

    public Stylesheet CreateStylesheet()
    {
        // Each list starts with a default entry (font, border and cell format 0). The fills list starts with two,
        // None and Gray125, ahead of any custom fill. Registered style i then uses font i, border i and fill i + 1,
        // which is why the loop below appends one font, fill and border per style, even when a style leaves them at their defaults.
        var fonts = new Fonts(new Font()) { Count = 1U };
        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }))
        {
            Count = 2U
        };
        var borders = new Borders(new Border()) { Count = 1U };
        var numberingFormats = new NumberingFormats();
        var cellFormats = new CellFormats(new CellFormat()) { Count = 1U };

        for (var i = 1; i < styles.Count; i++)
        {
            var style = styles[i];
            fonts.Append(CreateFont(style));
            fills.Append(CreateFill(style));
            borders.Append(CreateBorder(style));
            var cellFormat = new CellFormat
            {
                FontId = (uint)i,
                FillId = (uint)(i + 1),
                BorderId = (uint)i,
                NumberFormatId = GetNumberFormatId(style, numberingFormats),
                ApplyFont = true,
                ApplyFill = style.BackgroundColor is not null,
                ApplyBorder = style.BorderStyle != ExcelBorderLineStyle.None,
                ApplyNumberFormat = style.NumberFormat != ExcelNumberFormat.General
            };

            if (style.WrapText is not null || style.TextRotation is not null)
            {
                cellFormat.Alignment = new Alignment
                {
                    WrapText = style.WrapText,
                    TextRotation = style.TextRotation is null ? null : (uint)style.TextRotation.Value
                };
                cellFormat.ApplyAlignment = true;
            }

            cellFormats.Append(cellFormat);
        }

        // Set the Count attributes from what was actually appended.
        fonts.Count = (uint)fonts.ChildElements.Count;
        fills.Count = (uint)fills.ChildElements.Count;
        borders.Count = (uint)borders.ChildElements.Count;
        cellFormats.Count = (uint)cellFormats.ChildElements.Count;

        numberingFormats.Count = (uint)numberingFormats.ChildElements.Count;
        return numberingFormats.Count > 0
            ? new Stylesheet(numberingFormats, fonts, fills, borders, cellFormats)
            : new Stylesheet(fonts, fills, borders, cellFormats);
    }

    private static Font CreateFont(ExcelStyleSpec style)
    {
        var font = new Font();
        if (style.Bold)
        {
            font.Append(new Bold());
        }

        if (style.Italic)
        {
            font.Append(new Italic());
        }

        if (style.Underline)
        {
            font.Append(new Underline());
        }

        if (style.FontSize is not null)
        {
            font.Append(new FontSize { Val = style.FontSize.Value });
        }

        if (style.FontColor is not null)
        {
            font.Append(new Color { Rgb = new HexBinaryValue(style.FontColor) });
        }

        if (style.FontFamily is not null)
        {
            font.Append(new FontName { Val = style.FontFamily });
        }

        return font;
    }

    private static Fill CreateFill(ExcelStyleSpec style)
    {
        if (style.BackgroundColor is null)
        {
            return new Fill(new PatternFill { PatternType = PatternValues.None });
        }

        return new Fill(new PatternFill(
            new ForegroundColor { Rgb = new HexBinaryValue(style.BackgroundColor) },
            new BackgroundColor { Indexed = 64U })
        {
            PatternType = MapPattern(style.FillPattern)
        });
    }

    private static Border CreateBorder(ExcelStyleSpec style)
    {
        if (style.BorderStyle == ExcelBorderLineStyle.None)
        {
            return new Border();
        }

        var borderStyle = MapBorderStyle(style.BorderStyle);
        var color = new Color { Rgb = new HexBinaryValue(style.BorderColor ?? "FF000000") };

        // OpenXML models every side separately, even when the API presents a simple all-side border.
        return new Border(
            new LeftBorder(new Color { Rgb = color.Rgb }) { Style = borderStyle },
            new RightBorder(new Color { Rgb = color.Rgb }) { Style = borderStyle },
            new TopBorder(new Color { Rgb = color.Rgb }) { Style = borderStyle },
            new BottomBorder(new Color { Rgb = color.Rgb }) { Style = borderStyle },
            new DiagonalBorder());
    }

    private static BorderStyleValues MapBorderStyle(ExcelBorderLineStyle style)
    {
        return style switch
        {
            ExcelBorderLineStyle.Thin => BorderStyleValues.Thin,
            ExcelBorderLineStyle.Medium => BorderStyleValues.Medium,
            ExcelBorderLineStyle.Thick => BorderStyleValues.Thick,
            ExcelBorderLineStyle.Dashed => BorderStyleValues.Dashed,
            ExcelBorderLineStyle.Dotted => BorderStyleValues.Dotted,
            ExcelBorderLineStyle.Double => BorderStyleValues.Double,
            ExcelBorderLineStyle.Hair => BorderStyleValues.Hair,
            ExcelBorderLineStyle.DashDot => BorderStyleValues.DashDot,
            ExcelBorderLineStyle.DashDotDot => BorderStyleValues.DashDotDot,
            ExcelBorderLineStyle.SlantDashDot => BorderStyleValues.SlantDashDot,
            _ => BorderStyleValues.None
        };
    }

    private static uint GetNumberFormatId(ExcelStyleSpec style, NumberingFormats numberingFormats)
    {
        return style.NumberFormat switch
        {
            ExcelNumberFormat.General => 0U,
            ExcelNumberFormat.Number => 4U,
            // Ids 4, 10, 14 and 49 below are Excel's built-in formats. Currency has no built-in id, so it is registered as a custom format from id 164.
            // AddNumberingFormat returns 0, which leaves the id at 164.
            ExcelNumberFormat.Currency => 164U + AddNumberingFormat(numberingFormats, 164U, "\"$\"#,##0.00"),
            ExcelNumberFormat.Percent => 10U,
            ExcelNumberFormat.Date => 14U,
            ExcelNumberFormat.Text => 49U,
            _ => 0U
        };
    }

    private static uint AddNumberingFormat(NumberingFormats numberingFormats, uint id, string formatCode)
    {
        if (!numberingFormats.Elements<NumberingFormat>().Any(format => format.NumberFormatId?.Value == id))
        {
            numberingFormats.Append(new NumberingFormat { NumberFormatId = id, FormatCode = formatCode });
        }

        return 0U;
    }

    private static PatternValues MapPattern(ExcelFillPattern pattern)
    {
        return pattern switch
        {
            ExcelFillPattern.Gray125 => PatternValues.Gray125,
            ExcelFillPattern.DarkGray => PatternValues.DarkGray,
            ExcelFillPattern.MediumGray => PatternValues.MediumGray,
            ExcelFillPattern.LightGray => PatternValues.LightGray,
            ExcelFillPattern.DarkHorizontal => PatternValues.DarkHorizontal,
            ExcelFillPattern.DarkVertical => PatternValues.DarkVertical,
            ExcelFillPattern.DarkGrid => PatternValues.DarkGrid,
            ExcelFillPattern.LightGrid => PatternValues.LightGrid,
            _ => PatternValues.Solid
        };
    }
}

internal static class ExcelColor
{
    public static string Normalize(string hexColor)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(hexColor, nameof(hexColor));
        var value = hexColor.Trim().TrimStart('#').ToUpperInvariant();

        // A six digit color has no alpha part, so it is treated as fully opaque.
        if (value.Length == 6)
        {
            value = "FF" + value;
        }

        if (value.Length != 8 || value.Any(character => !Uri.IsHexDigit(character)))
        {
            throw new ArgumentException($"Invalid color '{hexColor}'. Use #RRGGBB or AARRGGBB.", nameof(hexColor));
        }

        return value;
    }
}
