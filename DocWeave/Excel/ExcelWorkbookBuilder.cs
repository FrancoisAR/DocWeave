using DocWeave;

namespace DocWeave.Excel;

/// <summary>
/// Fluent builder for an Excel workbook export.
/// </summary>
public sealed class ExcelWorkbookBuilder
{
    private readonly List<ExcelSheetSpec> sheets = [];
    private readonly DocWeaveOperationContext operationContext = new();
    private ExcelWorkbookProtectionSpec? protection;

    internal ExcelWorkbookBuilder()
    {
    }

    /// <summary>
    /// Adds a worksheet and lets the caller place one or more data regions on it.
    /// </summary>
    public ExcelWorkbookBuilder AddSheet(string name, Action<ExcelSheetExportBuilder> configure)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(configure);

        var sheet = new ExcelSheetSpec(name);
        configure(new ExcelSheetExportBuilder(sheet));
        sheets.Add(sheet);
        return this;
    }

    /// <summary>
    /// Sends progress reports to the caller when rendering starts and finishes. The total is the number of sheets.
    /// </summary>
    public ExcelWorkbookBuilder WithProgress(IProgress<DocWeaveOperationProgress> progress)
    {
        operationContext.Progress = progress ?? throw new ArgumentNullException(nameof(progress));
        return this;
    }

    /// <summary>
    /// Lets the caller cancel the export. The token is checked before rendering starts, not during it, so a workbook that is already rendering runs to the end. A cancelled export throws <see cref="OperationCanceledException"/>.
    /// </summary>
    public ExcelWorkbookBuilder WithCancellation(CancellationToken cancellationToken)
    {
        operationContext.CancellationToken = cancellationToken;
        return this;
    }

    /// <summary>
    /// Sends audit events when rendering starts and finishes. They hold the sheet count, never the cell values.
    /// </summary>
    public ExcelWorkbookBuilder Audit(Action<DocWeaveAuditEvent> audit)
    {
        operationContext.Audit = audit ?? throw new ArgumentNullException(nameof(audit));
        return this;
    }

    /// <summary>
    /// Locks the workbook structure so sheets cannot be added, removed, renamed or reordered. The password uses Excel's legacy protection hash: it stops casual edits but is not encryption.
    /// </summary>
    /// <param name="password">The password needed to unlock the structure, or null to lock it without one.</param>
    public ExcelWorkbookBuilder ProtectStructure(string? password = null)
    {
        protection = new ExcelWorkbookProtectionSpec(LockStructure: true, password);
        return this;
    }

    /// <summary>
    /// Writes the workbook to a file, replacing it if it already exists.
    /// </summary>
    public void SaveAs(string path)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        using var stream = File.Create(path);
        SaveTo(stream);
    }

    /// <summary>
    /// Writes the workbook to a stream owned by the caller.
    /// </summary>
    public void SaveTo(Stream stream)
    {
        ArgumentNullException.ThrowIfNull(stream);
        operationContext.ThrowIfCancellationRequested();
        operationContext.Report("ExcelExport", "Starting", 0, sheets.Count);
        operationContext.AuditEvent("ExcelExport", "Rendering workbook", new Dictionary<string, object?>
        {
            ["SheetCount"] = sheets.Count
        });
        ExcelWorkbookRenderer.Render(new ExcelWorkbookSpec(sheets) { Protection = protection }, stream);
        operationContext.Report("ExcelExport", "Completed", sheets.Count, sheets.Count);
        operationContext.AuditEvent("ExcelExport", "Workbook rendered", new Dictionary<string, object?>
        {
            ["SheetCount"] = sheets.Count
        });
    }

    /// <summary>
    /// Builds the workbook into a byte array suitable for API responses, storage, or email attachments.
    /// </summary>
    public byte[] ToBytes()
    {
        using var stream = new MemoryStream();
        SaveTo(stream);
        return stream.ToArray();
    }
}
