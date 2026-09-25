namespace DocWeave.Excel;

/// <summary>
/// The cells each region of each sheet would take up.
/// </summary>
public sealed record ExcelTemplatePreviewResult(IReadOnlyList<ExcelTemplateSheetPreview> Sheets);

/// <summary>
/// One output sheet in a template preview.
/// </summary>
/// <param name="Name">The sheet name as written in the template.</param>
/// <param name="Regions">The regions that occupy cells, in template order.</param>
public sealed record ExcelTemplateSheetPreview(string Name, IReadOnlyList<ExcelTemplateRegionPreview> Regions);

/// <summary>
/// The cells one region will take up.
/// </summary>
/// <param name="Type">The region type, for example "table" or "chart".</param>
/// <param name="StartCell">The top-left cell, as an A1 reference.</param>
/// <param name="EndCell">The bottom-right cell, as an A1 reference.</param>
public sealed record ExcelTemplateRegionPreview(string Type, string StartCell, string EndCell);

/// <summary>
/// The outcome of checking a template against the data bound to it.
/// </summary>
public sealed record ExcelTemplateValidationResult(bool IsValid, IReadOnlyList<ExcelTemplateValidationError> Errors);

/// <summary>
/// One problem found in an Excel template.
/// </summary>
/// <param name="Path">Where the problem is. Currently always "$", meaning the whole template.</param>
/// <param name="Code">The name of the kind of problem, taken from the exception type that reported it.</param>
/// <param name="Message">A description of the problem.</param>
public sealed record ExcelTemplateValidationError(string Path, string Code, string Message);
