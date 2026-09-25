namespace DocWeave.Excel;

/// <summary>
/// Entry point for creating Excel workbooks through a fluent export builder.
/// </summary>
public static class ExcelWriter
{
    /// <summary>
    /// Creates a new workbook builder.
    /// </summary>
    public static ExcelWorkbookBuilder Create()
    {
        return new ExcelWorkbookBuilder();
    }
}
