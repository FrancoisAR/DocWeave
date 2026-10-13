namespace DocWeave.Excel;

/// <summary>
/// Resolves stored Excel template JSON by template id.
/// </summary>
public interface IExcelTemplateProvider
{
    /// <summary>
    /// Returns the JSON of a stored template.
    /// </summary>
    /// <param name="templateId">The id the template was stored under.</param>
    /// <returns>The template document.</returns>
    /// <exception cref="System.InvalidOperationException">No template is stored under that id.</exception>
    string GetTemplateJson(string templateId);
}

/// <summary>
/// Simple in-memory template provider for tests, prototypes, and trusted internal callers.
/// </summary>
public sealed class InMemoryExcelTemplateProvider : IExcelTemplateProvider
{
    private readonly Dictionary<string, string> templates = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Stores a template, replacing any earlier one with the same id. Ids are matched ignoring case.
    /// </summary>
    /// <param name="templateId">The id to store the template under.</param>
    /// <param name="templateJson">The template document.</param>
    /// <returns>This provider, so calls can be chained.</returns>
    public InMemoryExcelTemplateProvider Add(string templateId, string templateJson)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(templateId, nameof(templateId));
        NetStandardGuard.ThrowIfNullOrWhiteSpace(templateJson, nameof(templateJson));
        templates[templateId] = templateJson;
        return this;
    }

    /// <summary>
    /// Returns the JSON of a stored template.
    /// </summary>
    /// <param name="templateId">The id the template was stored under.</param>
    /// <returns>The template document.</returns>
    /// <exception cref="System.InvalidOperationException">No template is stored under that id.</exception>
    public string GetTemplateJson(string templateId)
    {
        NetStandardGuard.ThrowIfNullOrWhiteSpace(templateId, nameof(templateId));
        return templates.TryGetValue(templateId, out var template)
            ? template
            : throw new InvalidOperationException($"Stored Excel template '{templateId}' was not found.");
    }
}
