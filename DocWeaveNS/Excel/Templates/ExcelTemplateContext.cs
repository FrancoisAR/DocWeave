using System.Globalization;
using System.Text.RegularExpressions;

namespace DocWeave.Excel;

/// <summary>
/// What a template is compiled against: the data sources and parameters the caller bound, and the styles
/// the template declares. Region compilers use it to look things up and to substitute parameters.
/// </summary>
internal sealed class ExcelTemplateContext
{
    private static readonly Regex ParameterExpression = new("\\{\\{\\s*parameters\\.(?<name>[A-Za-z_][A-Za-z0-9_]*)\\s*\\}\\}", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private readonly IReadOnlyDictionary<string, TabularData> sources;
    private readonly IReadOnlyDictionary<string, object?> parameters;
    private readonly IReadOnlyDictionary<string, ExcelStyleSpec> styles;

    public ExcelTemplateContext(
        IReadOnlyDictionary<string, TabularData> sources,
        IReadOnlyDictionary<string, object?> parameters,
        IReadOnlyDictionary<string, ExcelStyleSpec> styles)
    {
        this.sources = sources;
        this.parameters = parameters;
        this.styles = styles;
    }

    public TabularData GetSource(string sourceName)
    {
        return sources.TryGetValue(sourceName, out var source)
            ? source
            : throw new InvalidOperationException($"Template source '{sourceName}' has not been bound.");
    }

    /// <summary>
    /// Looks up a named style. No name means no styling.
    /// </summary>
    public ExcelStyleSpec ResolveStyle(string? styleName)
    {
        if (string.IsNullOrWhiteSpace(styleName))
        {
            return ExcelStyleSpec.Empty;
        }

        return styles.TryGetValue(styleName, out var style)
            ? style
            : throw new InvalidOperationException($"Template style '{styleName}' was not found.");
    }

    /// <summary>
    /// Replaces {{parameters.name}} placeholders with the bound parameter values.
    /// </summary>
    public string BindText(string text)
    {
        return ParameterExpressionRegex().Replace(text, match =>
        {
            var name = match.Groups["name"].Value;
            return parameters.TryGetValue(name, out var value)
                ? Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty
                : throw new InvalidOperationException($"Template parameter '{name}' has not been bound.");
        });
    }

    /// <summary>
    /// Binds parameters inside a value when it is text. Other values pass through unchanged.
    /// </summary>
    public object? BindValue(object? value)
    {
        return value is string text ? BindText(text) : value;
    }
    private static Regex ParameterExpressionRegex()
    {
        return ParameterExpression;
    }
}
