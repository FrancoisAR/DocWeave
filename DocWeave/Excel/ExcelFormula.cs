namespace DocWeave.Excel;

/// <summary>
/// Represents an Excel formula without the leading equals sign.
/// </summary>
public sealed record ExcelFormula
{
    /// <summary>
    /// Creates a formula.
    /// </summary>
    /// <param name="formulaText">The formula, with or without a leading equals sign.</param>
    /// <exception cref="System.ArgumentException">The text is null or blank.</exception>
    public ExcelFormula(string formulaText)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(formulaText);
        FormulaText = formulaText.TrimStart('=');
    }

    /// <summary>
    /// The formula without its leading equals sign.
    /// </summary>
    public string FormulaText { get; }
}

/// <summary>
/// Context supplied when building row-level table formulas.
/// </summary>
public sealed record ExcelFormulaContext(int DataIndex, int WorksheetRow);
