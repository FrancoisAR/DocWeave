using System.Globalization;
using System.Text.Json;

namespace DocWeave.Excel;

/// <summary>
/// { "type": "dataValidation", "range": "B2:B50", "validationType": "list", "values": ["Open", "Closed"], ... }
/// </summary>
internal sealed class DataValidationTemplateRegion : ExcelTemplateRegionCompiler
{
    public override string Type => "dataValidation";

    public override void Compile(ExcelSheetSpec sheet, JsonElement region, ExcelTemplateContext templateContext, ExcelCellAddress? startCell)
    {
        var builder = new ExcelDataValidationBuilder(region.GetRequiredString("range"));
        var validationType = region.GetRequiredString("validationType");
        switch (validationType.ToLowerInvariant())
        {
            case "list":
                if (region.TryGetProperty("values", out var values))
                {
                    builder.List(values.EnumerateArray().Select(value => templateContext.BindText(value.GetString() ?? string.Empty)).ToArray());
                }
                else
                {
                    builder.ListFromRange(templateContext.BindText(region.GetRequiredString("sourceRange")));
                }

                break;
            case "wholenumber":
                builder.WholeNumberBetween(region.GetRequiredInt32("min"), region.GetRequiredInt32("max"));
                break;
            case "decimal":
                builder.DecimalBetween(region.GetRequiredDecimal("min"), region.GetRequiredDecimal("max"));
                break;
            case "date":
                builder.DateBetween(DateTime.Parse(region.GetRequiredString("min"), CultureInfo.InvariantCulture), DateTime.Parse(region.GetRequiredString("max"), CultureInfo.InvariantCulture));
                break;
            case "textlength":
                builder.TextLengthBetween(region.GetRequiredInt32("min"), region.GetRequiredInt32("max"));
                break;
            case "custom":
                builder.CustomFormula(templateContext.BindText(region.GetRequiredString("formula")));
                break;
            default:
                throw new InvalidOperationException($"Data validation type '{validationType}' is not supported.");
        }

        builder.AllowBlank(region.GetOptionalBool("allowBlank") ?? true);
        if (region.TryGetProperty("inputMessage", out var inputMessage))
        {
            builder.InputMessage(templateContext.BindText(inputMessage.GetOptionalString("title") ?? string.Empty), templateContext.BindText(inputMessage.GetOptionalString("message") ?? string.Empty));
        }

        if (region.TryGetProperty("errorMessage", out var errorMessage))
        {
            builder.ErrorMessage(templateContext.BindText(errorMessage.GetOptionalString("title") ?? string.Empty), templateContext.BindText(errorMessage.GetOptionalString("message") ?? string.Empty));
        }

        sheet.DataValidations.Add(builder.Build());
    }
}
