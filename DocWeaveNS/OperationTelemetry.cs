namespace DocWeave;

/// <summary>
/// A progress report from a long-running operation, sent to the caller's <see cref="IProgress{T}"/>.
/// </summary>
/// <param name="Operation">The kind of operation, for example "CsvImport".</param>
/// <param name="Stage">The current step, for example "Starting", "Writing" or "Completed".</param>
/// <param name="Completed">How many rows or units are done so far.</param>
/// <param name="Total">How many there will be in all, or null when that is not known up front.</param>
public sealed record DocWeaveOperationProgress(string Operation, string Stage, int Completed, int? Total);

/// <summary>
/// A record of something that happened during an operation, sent to the caller's audit callback.
/// Row values are never included, only counts and names.
/// </summary>
/// <param name="Operation">The kind of operation, for example "CsvExport".</param>
/// <param name="Message">A short description of the event.</param>
/// <param name="Properties">Extra facts about the event, such as row and column counts.</param>
public sealed record DocWeaveAuditEvent(string Operation, string Message, IReadOnlyDictionary<string, object?> Properties);

/// <summary>
/// The rows read by an import, together with the column headers and any warnings.
/// </summary>
/// <typeparam name="TRow">The type of one row.</typeparam>
/// <param name="Headers">The column headers that were read, after any column selection.</param>
/// <param name="Rows">The rows that were read.</param>
/// <param name="Warnings">Non-fatal problems found while reading.</param>
public sealed record DocWeaveImportResult<TRow>(
    IReadOnlyList<string> Headers,
    IReadOnlyList<TRow> Rows,
    IReadOnlyList<string> Warnings)
{
    /// <summary>
    /// The number of rows that were read.
    /// </summary>
    public int RowCount => Rows.Count;
}

internal sealed class DocWeaveOperationContext
{
    public IProgress<DocWeaveOperationProgress>? Progress { get; set; }
    public Action<DocWeaveAuditEvent>? Audit { get; set; }
    public CancellationToken CancellationToken { get; set; }

    public void ThrowIfCancellationRequested()
    {
        CancellationToken.ThrowIfCancellationRequested();
    }

    public void Report(string operation, string stage, int completed, int? total = null)
    {
        Progress?.Report(new DocWeaveOperationProgress(operation, stage, completed, total));
    }

    public void AuditEvent(string operation, string message, IReadOnlyDictionary<string, object?>? properties = null)
    {
        Audit?.Invoke(new DocWeaveAuditEvent(operation, message, properties ?? new Dictionary<string, object?>()));
    }
}
